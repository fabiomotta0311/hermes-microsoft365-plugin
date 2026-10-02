"""The Bot Framework webhook: the HTTP surface the inbound boundary has always needed.

Why this module exists
----------------------
``teams_inbound`` decides whether an Activity is authentic and which session it belongs to, and
``teams_loopback`` composes the answer. Neither of them is reachable without something that
speaks HTTP, and a Bot Framework bot that cannot be called is a design document, not a product.
This is that surface, and nothing more: it parses a request, hands the parts to the boundary, and
renders the boundary's answer as a status code.

Two decisions are worth stating up front, because they are the ones a reviewer should challenge:

**A 401 never says why.** The boundary can refuse a token for half a dozen distinguishable
reasons -- malformed, wrong algorithm, wrong audience, wrong issuer, expired, bad signature. Every
one of them leaves here as the same anonymous ``unauthorized``. Which check failed is exactly the
feedback an attacker probing with forged tokens needs, and the legitimate caller learns the reason
from their own logs, not from the response.

**A failed reply still answers 200.** The Activity *was* received and correlated; only our answer
to it failed. Returning 5xx would make Bot Framework redeliver, and redelivery cannot help: the
Activity is already marked delivered, so the retry is reported as a duplicate and no second reply
is attempted. A 5xx here would buy a retry storm and no delivery. The failure is reported in the
body for the host to alert on, and the honest limitation is written down below rather than papered
over: **a reply that fails is not retried by this plugin.**

The application never constructs a Graph client and never sends anything on its own. The reply
leaves through the ``reply`` callable the host supplies -- in production, the already-authorised
send operation with its own settings, preflight and permission checks.
"""
from __future__ import annotations

import json
import time
from collections.abc import Callable, Iterable, Mapping
from typing import Any

from .errors import GraphError
from .teams_inbound import (
    BOT_FRAMEWORK_ISSUERS,
    DEFAULT_LEEWAY_SECONDS,
    ConversationRegistry,
    InboundDecision,
    InboundTeamsError,
    VerifySignature,
    admit_activity,
)

#: Where Bot Framework posts an Activity. A bot registration points at this path.
DEFAULT_WEBHOOK_PATH = "/teams/activities"

#: The largest Activity body accepted. An Activity is a small JSON document; anything larger is
#: either a mistake or an attempt to make the endpoint allocate on demand.
MAX_ACTIVITY_BYTES = 256 * 1024

#: The media types accepted for an Activity. Bot Framework sends JSON; the charset parameter is
#: permitted because it is meaningless for JSON and refusing it would reject valid senders.
ALLOWED_MEDIA_TYPES = ("application/json",)

_REASONS = {
    400: "invalid_request",
    401: "unauthorized",
    404: "not_found",
    405: "method_not_allowed",
    411: "length_required",
    413: "payload_too_large",
    415: "unsupported_media_type",
    500: "internal_error",
}

class WebhookError(Exception):
    """A refusal that maps directly onto a response. Never carries caller-supplied data."""

    def __init__(self, status_code: int, *, detail: str | None = None) -> None:
        super().__init__(detail or _REASONS[status_code])
        self.status_code = status_code
        self.status = _REASONS[status_code]


def _json_body(payload: Mapping[str, Any]) -> bytes:
    return json.dumps(payload, separators=(",", ":")).encode("utf-8")


def _reason_phrase(status_code: int, status: str) -> str:
    return status.replace("_", " ").title()


def _media_type(content_type: str) -> str:
    return content_type.split(";", 1)[0].strip().lower()


def _read_declared(stream: Any, declared: int) -> bytes:
    """Read exactly ``declared`` bytes, and never one more.

    ``Content-Length`` is the *only* thing that delimits a request body, and this is the trap
    that makes an otherwise correct-looking endpoint hang in production: under WSGI,
    ``environ["wsgi.input"]`` may be the raw socket stream, so ``read(n)`` for ``n`` larger than
    the body does not return what is available -- it blocks until the peer closes the connection.
    Reading exactly the declared length is what the specification requires, and asking for the
    maximum body size instead is how an endpoint deadlocks against every client that waits for
    its response before hanging up.

    A body shorter than declared is a truncated request, which is refused rather than parsed.
    """
    chunks: list[bytes] = []
    remaining = declared
    while remaining > 0:
        chunk = stream.read(remaining)
        if not chunk:
            break
        chunks.append(chunk)
        remaining -= len(chunk)
    payload = b"".join(chunks)
    if len(payload) != declared:
        raise WebhookError(400, detail="the body was shorter than its declared length")
    return payload


def _http_status_for(exc: GraphError) -> int:
    """Map a boundary refusal onto a status code without disclosing which check refused it."""
    category = getattr(exc, "category", None)
    if category == "authentication_required":
        # Malformed, forged, expired, wrong audience, wrong issuer: all the same answer.
        return 401
    if category == "validation_error":
        return 400
    # A configuration failure is ours to fix, and a retry after fixing it should succeed.
    return 500


def build_application(
    *,
    registry: ConversationRegistry,
    audience: str,
    keys: Iterable[Mapping] | None = None,
    verify_signature: VerifySignature | None = None,
    issuers: Iterable[str] = BOT_FRAMEWORK_ISSUERS,
    leeway_seconds: int = DEFAULT_LEEWAY_SECONDS,
    path: str = DEFAULT_WEBHOOK_PATH,
    on_activity: Callable[[InboundDecision], None] | None = None,
    reply: Callable[[InboundDecision], Any] | None = None,
    max_body_bytes: int = MAX_ACTIVITY_BYTES,
    clock: Callable[[], float] = time.time,
) -> Callable[[dict, Callable], list[bytes]]:
    """Build a WSGI application.

    WSGI, and no new dependency, because the plugin already refuses to require more than it uses:
    the host can mount this behind its own server, and the tests exercise it over a real socket
    with the standard library alone.

    ``on_activity`` receives every admitted, non-duplicate decision -- this is the hand-off point
    where the host puts the text in front of a session. It is called *before* ``reply`` so a
    session that is going to answer has already seen the turn.

    ``reply`` receives the decision and returns anything; a refusal it raises is reported as an
    undelivered answer with a sanitized status, still under a 200. Its response is never inspected
    for authorization -- the callable the host supplies is expected to be the authorised path.
    """
    if not isinstance(audience, str) or not audience:
        raise ValueError("audience is required: an unverifiable token must not be acceptable")
    if not path.startswith("/"):
        raise ValueError("path must be an absolute path")

    allowed_issuers = tuple(issuers)

    def application(environ: dict, start_response: Callable) -> list[bytes]:
        try:
            status_code, status, body = _handle(
                environ,
                registry=registry,
                audience=audience,
                keys=keys,
                verify_signature=verify_signature,
                issuers=allowed_issuers,
                leeway_seconds=leeway_seconds,
                path=path,
                on_activity=on_activity,
                reply=reply,
                max_body_bytes=max_body_bytes,
                clock=clock,
            )
        except WebhookError as refused:
            status_code, status = refused.status_code, refused.status
            body = _json_body({"status": status})
        except Exception:  # noqa: BLE001 - an unexpected failure must not leak its traceback
            status_code, status = 500, _REASONS[500]
            body = _json_body({"status": status})

        headers = {"Content-Type": "application/json", "Content-Length": str(len(body))}
        if status_code == 405:
            headers["Allow"] = "POST"
        start_response(f"{status_code} {_reason_phrase(status_code, status)}", list(headers.items()))
        return [body]

    return application


def _handle(environ: dict, **options: Any) -> tuple[int, str, bytes]:
    method = (environ.get("REQUEST_METHOD") or "").upper()
    request_path = environ.get("PATH_INFO") or ""

    if request_path != options["path"]:
        # An unknown path is not told that the endpoint exists elsewhere.
        raise WebhookError(404)
    if method != "POST":
        raise WebhookError(405)

    content_type = environ.get("CONTENT_TYPE") or ""
    if _media_type(content_type) not in ALLOWED_MEDIA_TYPES:
        raise WebhookError(415)

    length_header = environ.get("CONTENT_LENGTH")
    if not length_header or not str(length_header).strip():
        raise WebhookError(411)
    try:
        declared = int(str(length_header).strip())
    except ValueError:
        raise WebhookError(400) from None
    if declared < 0:
        raise WebhookError(400)
    if declared > options["max_body_bytes"]:
        raise WebhookError(413)

    raw = _read_declared(environ["wsgi.input"], declared)
    if not raw:
        raise WebhookError(400)
    try:
        payload = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        raise WebhookError(400) from None

    try:
        decision = admit_activity(
            authorization=environ.get("HTTP_AUTHORIZATION"),
            payload=payload,
            registry=options["registry"],
            audience=options["audience"],
            keys=options["keys"],
            verify_signature=options["verify_signature"],
            issuers=options["issuers"],
            leeway_seconds=options["leeway_seconds"],
            now=options["clock"](),
        )
    except InboundTeamsError as refused:
        raise WebhookError(_http_status_for(refused)) from None

    if decision.duplicate:
        # Acknowledged so the Bot Framework stops retrying, and nothing is delivered twice.
        return 200, "duplicate", _json_body({"status": "duplicate"})

    if options["on_activity"] is not None:
        # A failure here is a genuine unknown: the host may or may not have consumed the turn.
        # 500 lets the redelivery happen; the duplicate guard keeps it from double-delivering.
        options["on_activity"](decision)

    if options["reply"] is None:
        return 200, "accepted", _json_body({"status": "accepted"})

    try:
        outcome = options["reply"](decision)
    except GraphError as refused:
        # A lifecycle Activity, or a duplicate, legitimately has no reply. That is a 200 with an
        # undelivered answer -- not a failure of the request.
        return 200, "not_delivered", _render_outcome(
            None, getattr(refused, "status", None), "not_delivered"
        )
    except Exception:  # noqa: BLE001
        return 200, "not_delivered", _render_outcome(None, "reply_failed", "not_delivered")

    return 200, "accepted", _render_outcome(outcome, None, "accepted")


def _render_outcome(outcome: Any, error_status: str | None, status: str = "accepted") -> bytes:
    """Render the reply outcome without ever echoing a message id or a chat identifier."""
    payload: dict[str, Any] = {"status": status}
    if outcome is not None and hasattr(outcome, "to_result"):
        result = outcome.to_result()
        payload["reply_delivered"] = bool(result.get("delivered"))
        payload["reply_body_length"] = result.get("body_length")
        if not result.get("delivered"):
            payload["reply_error"] = result.get("error_status") or "send_failed"
    elif error_status is not None:
        payload["reply_delivered"] = False
        payload["reply_error"] = error_status
    body = json.dumps(payload, separators=(",", ":")).encode("utf-8")
    return body
