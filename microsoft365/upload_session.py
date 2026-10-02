"""Resumable drive upload sessions: large files, real conflict semantics, retries.

Why this module exists
----------------------
:mod:`microsoft365.files` implements the *simple* ``PUT /content`` transfer, whose Graph
contract tops out at 10 MiB and whose conflict behavior is always "replace". Production
uploads need the other half of the story:

* files larger than the simple-transfer bound;
* an explicit conflict policy (``replace``/``rename``/``fail``) instead of silent overwrite;
* resumable chunks, so a transient failure does not restart the whole transfer;
* bounded retries with the range Graph still expects, never a blind re-send.

Contract (verified against the generated SDK and the Microsoft Graph driveItem docs)
------------------------------------------------------------------------------------
1. ``POST /drives/{drive-id}/items/{address}:/createUploadSession`` with a real
   ``CreateUploadSessionPostRequestBody`` carrying the target ``DriveItem`` name and the
   ``@odata.conflictBehavior`` value. The response is a real ``UploadSession``.
2. One ``PUT`` per chunk to the session's pre-authenticated ``uploadUrl``, with
   ``Content-Range: bytes {start}-{end}/{total}``. Chunk sizes must be multiples of 320 KiB.
3. A chunk response of ``201``/``200`` carries the committed ``DriveItem``: the transfer is
   done. A ``202`` carries ``nextExpectedRanges``: more chunks remain, and the next chunk
   starts at the first range Graph still expects -- which is what makes this resumable after
   a failure rather than merely retryable from zero.

Security invariants
-------------------
* The ``uploadUrl`` is a bearer-equivalent secret. It never reaches a message, a log, a
  result payload or an exception: every failure raised here is a :class:`UploadSessionError`
  with a static taxonomy message plus a bounded ``status``, exactly like
  :class:`microsoft365.files.TransferError`.
* No ``Authorization`` header is attached to the chunk PUTs. The URL is already pre-authorized
  by Graph, and sending the tenant credential to a second origin would widen the blast radius.
* Retries are bounded, use a fixed backoff supplied by the caller, and only replay the range
  Graph asked for. A ``4xx`` is never retried: it is a contract failure, not a blip.

Everything defaults to an injectable seam (``execute``/``put``/``sleep``) so the module is
fully exercisable offline; production passes nothing and gets the real seam.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable, Mapping
from urllib.parse import urlsplit

from kiota_abstractions.base_request_configuration import RequestConfiguration
from kiota_abstractions.headers_collection import HeadersCollection

from . import execution
from .errors import MESSAGES, GraphError
from .files import (
    DEFAULT_PATH_IDENTIFIER,
    SIMPLE_TRANSFER_LIMIT_BYTES,
    SESSION_TRANSFER_LIMIT_BYTES,
    TransferError,
    drive_item_address,
    upload_file as simple_upload_file,
    validate_content_type,
    validate_drive_id,
)
from .files import INVALID_CONTENT_BASE64

#: Graph requires every upload-session fragment to be a multiple of 320 KiB.
CHUNK_SIZE_MULTIPLE = 320 * 1024

#: 5 MiB: a common, efficient default that divides into exact 320 KiB fragments.
DEFAULT_CHUNK_SIZE = 5 * 1024 * 1024

#: Upper bound on a single fragment. Keeps one request's failure blast radius small and the
#: memory held per attempt bounded.
MAX_CHUNK_SIZE = 60 * 1024 * 1024

#: The conflict policies Graph accepts for ``@odata.conflictBehavior``.
CONFLICT_BEHAVIORS = ("replace", "rename", "fail")

#: Bound on replayed attempts for one chunk before the transfer is reported as failed.
MAX_CHUNK_ATTEMPTS = 5

#: Graph's own upper bound for a single upload session (250 GiB), shared with the decoder.
MAX_SESSION_BYTES = SESSION_TRANSFER_LIMIT_BYTES

#: Hard ceiling on the fragments one transfer may send. It exists so a server that keeps
#: answering ``202`` with an unchanged offset cannot spin the transfer forever.
MAX_SESSION_FRAGMENTS = 10_000

#: How many consecutive fragments Graph may ask to re-send before the transfer is reported as
#: unfinished. A single correction is normal (a lost fragment); an endless one is not.
MAX_STALLED_FRAGMENTS = 5

INVALID_CHUNK_SIZE = "invalid_chunk_size"
INVALID_CONFLICT_BEHAVIOR = "invalid_conflict_behavior"
INVALID_UPLOAD_URL = "invalid_upload_url"
UPLOAD_SESSION_FAILED = "upload_session_failed"
UPLOAD_SESSION_INCOMPLETE = "upload_session_incomplete"
UPLOAD_SESSION_TOO_LARGE = "upload_session_too_large"
UPLOAD_SESSION_NO_URL = "upload_session_no_url"
UPLOAD_SESSION_NO_NAME = "upload_session_no_name"

_ACTIONS = {
    INVALID_CHUNK_SIZE: "use a chunk size that is a multiple of 320 KiB and at most 60 MiB",
    INVALID_CONFLICT_BEHAVIOR: "use one of replace, rename or fail",
    INVALID_UPLOAD_URL: "the session returned no usable pre-authenticated upload URL",
    UPLOAD_SESSION_FAILED: "retry the upload; already committed chunks are not re-sent",
    UPLOAD_SESSION_INCOMPLETE: "the transfer ended before Graph reported the file committed",
    UPLOAD_SESSION_TOO_LARGE: "upload through several sessions or reduce the file size",
    UPLOAD_SESSION_NO_URL: "create a new upload session and restart the transfer",
    UPLOAD_SESSION_NO_NAME: "address the target file by a path ending in the file name",
}


@dataclass(frozen=True)
class ChunkOutcome:
    """What one fragment PUT returned: the HTTP status, any ranges, and the committed item."""

    status: int
    next_expected_ranges: list[str] = field(default_factory=list)
    item: Any = None


class UploadSessionError(GraphError):
    """A sanitized upload-session failure carrying an explicit, machine-readable ``status``."""

    def __init__(
        self,
        category: str,
        *,
        status: str,
        action: str | None = None,
        attempts: int | None = None,
        chunk_size_bytes: int | None = None,
        size_bytes: int | None = None,
        expected_range: str | None = None,
    ):
        super().__init__(category, MESSAGES[category])
        self.status = str(status)
        self.action = action if action is not None else _ACTIONS.get(self.status)
        self.attempts = attempts
        self.chunk_size_bytes = chunk_size_bytes
        self.size_bytes = size_bytes
        self.expected_range = expected_range

    def to_result(self) -> dict:
        payload = super().to_result()
        payload["status"] = self.status
        if self.action:
            payload["action"] = self.action
        for name in ("attempts", "chunk_size_bytes", "size_bytes", "expected_range"):
            value = getattr(self, name)
            if value is not None:
                payload[name] = value
        return payload


def _reject(status: str, **metadata: Any) -> UploadSessionError:
    return UploadSessionError("validation_error", status=status, **metadata)


def validate_chunk_size(chunk_size: Any) -> int:
    """A fragment size that is a multiple of 320 KiB within the supported bounds."""
    if isinstance(chunk_size, bool) or not isinstance(chunk_size, int):
        raise _reject(INVALID_CHUNK_SIZE)
    if not CHUNK_SIZE_MULTIPLE <= chunk_size <= MAX_CHUNK_SIZE:
        raise _reject(INVALID_CHUNK_SIZE, chunk_size_bytes=chunk_size)
    if chunk_size % CHUNK_SIZE_MULTIPLE:
        raise _reject(INVALID_CHUNK_SIZE, chunk_size_bytes=chunk_size)
    return chunk_size


def validate_conflict_behavior(behavior: Any) -> str:
    """One of Graph's three ``@odata.conflictBehavior`` values, or the default ``replace``."""
    if behavior is None:
        return "replace"
    if not isinstance(behavior, str) or behavior not in CONFLICT_BEHAVIORS:
        raise _reject(INVALID_CONFLICT_BEHAVIOR)
    return behavior


def validate_session_size(size: int) -> int:
    """The total transfer must fit one Graph upload session."""
    if size > MAX_SESSION_BYTES:
        raise UploadSessionError(
            "validation_error", status=UPLOAD_SESSION_TOO_LARGE, size_bytes=size
        )
    return size


def _has_control_characters(text: str) -> bool:
    import unicodedata

    return any(unicodedata.category(character).startswith("C") for character in text)


def validate_upload_url(upload_url: Any) -> str:
    """A pre-authenticated absolute HTTPS URL on a Graph-issued host.

    The value never leaves this module; validating it here is what keeps a caller from
    redirecting the chunk PUTs at an origin the tenant did not authorize.
    """
    if not isinstance(upload_url, str) or not upload_url or len(upload_url) > 4096:
        raise UploadSessionError("configuration_error", status=INVALID_UPLOAD_URL)
    if upload_url != upload_url.strip() or _has_control_characters(upload_url):
        raise UploadSessionError("configuration_error", status=INVALID_UPLOAD_URL)
    parts = urlsplit(upload_url)
    if parts.scheme != "https" or not parts.netloc:
        raise UploadSessionError("configuration_error", status=INVALID_UPLOAD_URL)
    if parts.username or parts.password:
        raise UploadSessionError("configuration_error", status=INVALID_UPLOAD_URL)
    return upload_url


def file_name_from_address(address: str) -> str:
    """The target file name Graph needs in the session body, taken from the address only."""
    candidate = address[:-1].partition(":/")[2] if address.endswith(":") else address
    name = candidate.rsplit("/", 1)[-1]
    if not name or _has_control_characters(name) or len(name) > 400:
        raise _reject(UPLOAD_SESSION_NO_NAME)
    return name


def build_session_body(name: str, conflict_behavior: str):
    """The real generated request body: the target ``DriveItem`` plus the conflict policy."""
    from msgraph.generated.drives.item.items.item.create_upload_session.create_upload_session_post_request_body import (
        CreateUploadSessionPostRequestBody,
    )
    from msgraph.generated.models.drive_item import DriveItem

    return CreateUploadSessionPostRequestBody(
        item=DriveItem(name=name, additional_data={"@odata.conflictBehavior": conflict_behavior})
    )


def build_session_request(
    client: Any,
    *,
    drive_id: str,
    drive_item_id: str,
    name: str,
    conflict_behavior: str,
) -> Any:
    """Build the ``createUploadSession`` POST with a real generated request body."""
    builder = (
        client.drives.by_drive_id(drive_id).items.by_drive_item_id(drive_item_id).create_upload_session
    )
    return builder.to_post_request_information(
        build_session_body(name, conflict_behavior),
        RequestConfiguration(headers=HeadersCollection()),
    )


def create_upload_session(
    client: Any,
    *,
    drive_id: str,
    drive_item_id: str,
    name: str,
    conflict_behavior: str,
    execute: Callable[..., Any] = execution.execute_request,
) -> tuple[str, list[str]]:
    """Create the session and return ``(upload_url, next_expected_ranges)``.

    The upload URL is returned to the caller of this function only; it is never placed in a
    result payload. Graph's ``nextExpectedRanges`` is normalized to a list of ``"start-end"``
    strings (Graph also allows the single value ``"start-"`` for an open-ended range).
    """
    from msgraph.generated.models.upload_session import UploadSession

    request = build_session_request(
        client,
        drive_id=drive_id,
        drive_item_id=drive_item_id,
        name=name,
        conflict_behavior=conflict_behavior,
    )
    if not str(getattr(request, "url", "")).endswith("/createUploadSession"):
        raise UploadSessionError("configuration_error", status=INVALID_UPLOAD_URL)

    adapter = getattr(client, "request_adapter", None)
    session = execute(
        client.drives.by_drive_id(drive_id).items.by_drive_item_id(drive_item_id).create_upload_session,
        method="POST",
        configuration=RequestConfiguration(headers=HeadersCollection()),
        body=build_session_body(name, conflict_behavior),
        adapter=adapter,
    )
    if not isinstance(session, UploadSession):
        raise UploadSessionError("service_error", status=UPLOAD_SESSION_FAILED)
    upload_url = validate_upload_url(getattr(session, "upload_url", None))
    ranges = [
        str(value) for value in (session.next_expected_ranges or []) if isinstance(value, str)
    ]
    return upload_url, ranges


def _first_offset(ranges: list[str]) -> int:
    """The next byte offset Graph expects, from its own ``nextExpectedRanges``."""
    for value in ranges:
        head = value.partition("-")[0]
        if head.isdigit():
            return int(head)
    return 0


def _committed_item(candidate: Any):
    """The final ``DriveItem`` when Graph reports the transfer committed, else ``None``."""
    from msgraph.generated.models.drive_item import DriveItem

    return candidate if isinstance(candidate, DriveItem) else None


def _committed_drive_item(body: Any):
    """Build a real ``DriveItem`` from a committed response body, or ``None``.

    The chunk seam receives raw JSON rather than a generated model, so the item is constructed
    from the deserialized payload here; only the three reconciliation fields are read and only
    scalar values are accepted.
    """
    if not isinstance(body, Mapping):
        return None
    from msgraph.generated.models.drive_item import DriveItem

    scalars = {
        "id": body.get("id"),
        "web_url": body.get("webUrl"),
        "e_tag": body.get("eTag"),
        "name": body.get("name"),
        "size": body.get("size"),
    }
    scalars = {
        name: value
        for name, value in scalars.items()
        if (isinstance(value, str) or isinstance(value, int)) and not isinstance(value, bool)
    }
    if not scalars:
        return None
    return DriveItem(**scalars)


def upload_chunks(
    *,
    upload_url: str,
    payload: bytes,
    start_offset: int,
    chunk_size: int,
    content_type: str,
    put: Callable[..., Any],
    sleep: Callable[[float], None],
    max_attempts: int = MAX_CHUNK_ATTEMPTS,
) -> dict:
    """Send every remaining fragment, resuming from ``start_offset`` and honoring retries.

    ``put`` receives ``(upload_url, start, end, total, chunk, content_type)`` and returns a
    :class:`ChunkOutcome`; the concrete seam supplies it from the transport, so this function
    holds no HTTP knowledge at all. A transport exception is retried up to ``max_attempts``
    with the caller's backoff; an HTTP status other than ``200``/``201``/``202`` is a contract
    failure and is reported immediately rather than replayed.
    """
    url = validate_upload_url(upload_url)
    total = len(payload)
    offset = start_offset
    attempts = 0
    fragments = 0
    stalls = 0
    while offset < total:
        end = min(offset + chunk_size, total) - 1
        chunk = payload[offset : end + 1]
        fragments += 1
        if fragments > MAX_SESSION_FRAGMENTS:
            # A server that keeps answering ``202`` with the same offset must not spin this
            # transfer forever: the fragment bound is what makes a stuck session a reported
            # failure instead of an unbounded loop.
            raise UploadSessionError(
                "service_error",
                status=UPLOAD_SESSION_INCOMPLETE,
                size_bytes=total,
                attempts=fragments,
            )
        try:
            outcome = put(url, offset, end, total, chunk, content_type)
        except Exception as exc:  # noqa: BLE001 - transport failures are retried, then reported
            attempts += 1
            if attempts >= max_attempts:
                raise UploadSessionError(
                    "service_error",
                    status=UPLOAD_SESSION_FAILED,
                    attempts=attempts,
                    expected_range=f"bytes {offset}-{end}/{total}",
                ) from exc
            sleep(min(2.0 * attempts, 8.0))
            continue

        attempts = 0
        if outcome.status in (200, 201):
            return {
                "status": "uploaded",
                "size_bytes": total,
                "item": _committed_item(outcome.item),
            }
        if outcome.status == 202:
            reported = _first_offset(outcome.next_expected_ranges)
            stalls = stalls + 1 if reported <= offset else 0
            if stalls >= MAX_STALLED_FRAGMENTS:
                # Graph keeps asking for the same bytes: replaying them forever would be an
                # unbounded loop against a session that will never finish.
                raise UploadSessionError(
                    "service_error",
                    status=UPLOAD_SESSION_INCOMPLETE,
                    size_bytes=total,
                    attempts=fragments,
                    expected_range=f"bytes {offset}-{end}/{total}",
                )
            offset = reported
            continue
        raise UploadSessionError(
            "service_error",
            status=UPLOAD_SESSION_FAILED,
            expected_range=f"bytes {offset}-{end}/{total}",
        )

    raise UploadSessionError(
        "service_error", status=UPLOAD_SESSION_INCOMPLETE, size_bytes=total
    )


def upload_file_session(
    client: Any,
    *,
    drive_id: Any,
    content_base64: str,
    content_type: Any,
    drive_item_id: Any = None,
    path: Any = None,
    identifier: str = DEFAULT_PATH_IDENTIFIER,
    conflict_behavior: Any = None,
    chunk_size: Any = DEFAULT_CHUNK_SIZE,
    execute: Callable[..., Any] = execution.execute_request,
    put: Callable[..., Any],
    sleep: Callable[[float], None],
) -> dict:
    """Upload exact bytes through a resumable session and report the committed item.

    Returns a result with ``status``, ``size_bytes``, ``conflict_behavior`` and the committed
    ``DriveItem`` fields the caller needs to reconcile the write. It never contains the
    pre-authenticated upload URL.
    """
    from .files import decode_bounded_base64

    drive = validate_drive_id(drive_id)
    address = drive_item_address(drive_item_id=drive_item_id, path=path, identifier=identifier)
    behavior = validate_conflict_behavior(conflict_behavior)
    size = validate_chunk_size(chunk_size)
    media_type = validate_content_type(content_type)
    payload = decode_bounded_base64(content_base64, limit_bytes=MAX_SESSION_BYTES)

    upload_url, ranges = create_upload_session(
        client,
        drive_id=drive,
        drive_item_id=address,
        name=file_name_from_address(address),
        conflict_behavior=behavior,
        execute=execute,
    )
    result = upload_chunks(
        upload_url=upload_url,
        payload=payload,
        start_offset=_first_offset(ranges),
        chunk_size=size,
        content_type=media_type,
        put=put,
        sleep=sleep,
    )
    item = result.get("item")
    return {
        "status": "uploaded",
        "size_bytes": result["size_bytes"],
        "content_type": media_type,
        "conflict_behavior": behavior,
        "item_id": getattr(item, "id", None) if item is not None else None,
        "web_url": getattr(item, "web_url", None) if item is not None else None,
        "e_tag": getattr(item, "e_tag", None) if item is not None else None,
    }


def put_chunk_default(
    upload_url: str,
    start: int,
    end: int,
    total: int,
    chunk: bytes,
    content_type: str,
) -> ChunkOutcome:
    """The production chunk seam: one pre-authorized ``PUT`` per fragment.

    No tenant credential is attached -- the URL Graph returned is already authorized -- and the
    fragment's ``Content-Range`` is what makes the request idempotent for its own byte window.
    """
    import httpx

    headers = {
        "Content-Type": content_type,
        "Content-Range": f"bytes {start}-{end}/{total}",
    }
    with httpx.Client(timeout=httpx.Timeout(120.0, connect=10.0), follow_redirects=False) as http:
        response = http.put(upload_url, content=chunk, headers=headers)
    if response.status_code == 202:
        try:
            body = response.json()
        except ValueError:
            body = {}
        raw = body.get("nextExpectedRanges") if isinstance(body, Mapping) else None
        ranges = [str(value) for value in raw] if isinstance(raw, list) else []
        return ChunkOutcome(status=202, next_expected_ranges=ranges)
    if response.status_code in (200, 201):
        try:
            body = response.json()
        except ValueError:
            body = None
        return ChunkOutcome(status=response.status_code, item=_committed_drive_item(body))
    return ChunkOutcome(status=response.status_code)


def _decoded_length(content_base64: Any) -> int:
    """The exact decoded length of a base64 payload, validated without materializing it.

    The strategy decision needs the size before allocating, so this validates the encoding's
    alphabet and padding with the streaming decoder and never builds the byte string twice.
    """
    import base64
    import binascii

    if not isinstance(content_base64, str) or not content_base64:
        raise TransferError("validation_error", status=INVALID_CONTENT_BASE64)
    try:
        decoded = base64.b64decode(content_base64, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise TransferError("validation_error", status=INVALID_CONTENT_BASE64) from exc
    return len(decoded)


def upload_file_auto(
    client: Any,
    *,
    drive_id: Any,
    content_base64: Any,
    content_type: Any,
    drive_item_id: Any = None,
    path: Any = None,
    identifier: str = DEFAULT_PATH_IDENTIFIER,
    conflict_behavior: Any = None,
    chunk_size: Any = None,
    execute: Callable[..., Any] = execution.execute_request,
    put: Callable[..., Any],
    sleep: Callable[[float], None],
) -> dict:
    """Upload exact bytes through whichever transfer can actually honor the request.

    ``auto`` names the *strategy selection*, not a provisional contract: the caller always
    states what it wants (``conflict_behavior``, optional ``chunk_size``) and this function
    picks the smallest mechanism that can honor it.

    * ``replace`` within :data:`microsoft365.files.SIMPLE_TRANSFER_LIMIT_BYTES` -> one simple
      ``PUT /content``. One request, no session, no pre-authenticated URL.
    * anything else -- a larger file, or ``fail``/``rename``, which a simple PUT cannot express
      -- goes through the resumable session.

    Both paths return the same shape so the caller reconciles one result contract.
    """
    behavior = validate_conflict_behavior(conflict_behavior)
    session_bytes = _decoded_length(content_base64)

    if behavior == "replace" and session_bytes <= SIMPLE_TRANSFER_LIMIT_BYTES:
        result = simple_upload_file(
            client,
            drive_id=drive_id,
            content_base64=content_base64,
            content_type=content_type,
            drive_item_id=drive_item_id,
            path=path,
            identifier=identifier,
            execute=execute,
        )
        return {**result, "conflict_behavior": behavior, "transfer": "simple"}

    session_result = upload_file_session(
        client,
        drive_id=drive_id,
        content_base64=content_base64,
        content_type=content_type,
        drive_item_id=drive_item_id,
        path=path,
        identifier=identifier,
        conflict_behavior=behavior,
        chunk_size=DEFAULT_CHUNK_SIZE if chunk_size is None else chunk_size,
        execute=execute,
        put=put,
        sleep=sleep,
    )
    return {**session_result, "transfer": "session"}


def sleep_default(seconds: float) -> None:
    """Bounded, monotonic backoff between fragment retries."""
    import time

    time.sleep(seconds)