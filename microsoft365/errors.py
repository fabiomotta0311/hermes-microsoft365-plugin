"""The canonical sanitized Microsoft Graph error taxonomy (WP-ERR).

One conversion point, one closed set of categories
--------------------------------------------------
Every failure raised by this plugin -- a Kiota ``APIError``/``ODataError`` returned by the
generated builders, an azure-core / azure-identity credential or transport failure, a
``kiota_http`` error, an ``httpx`` transport error, or a plain programming error -- is
converted by :func:`to_graph_error` into exactly one :class:`GraphError`. Handlers, the
execution seam (``microsoft365.execution``) and any future retry loop use this module
instead of inventing their own category or message, so the reported vocabulary can never
drift between call sites.

What may never leave this layer
-------------------------------
Raw SDK/Graph/Kiota/Azure exception text, response bodies, response **header values**,
request objects, token claims and credential material never reach a returned message, an
error string, a tool result or a log payload. Two mechanisms make that structural rather
than aspirational:

* every message is a **static template** from :data:`MESSAGES` -- nothing is interpolated,
  so no attribute value can be spliced into a message even by accident;
* the only exception attributes this module reads are *enumerated signals*: an HTTP status
  integer, a bounded structured error code (used as a lookup key only, never re-emitted),
  a bounded ``Retry-After`` value and a correlation id that must be GUID/hex shaped.

The originating exception is still available to the caller as ``__cause__`` (set by
:func:`to_graph_error`), which is where it belongs: local debugging, never a tool result.

Correlation ids
---------------
A correlation id is emitted only when it matches :data:`_GUID_PATTERN` or
:data:`_HEX_ID_PATTERN`. Both are hex-only shapes, so an id that passes validation cannot
carry words such as ``secret``/``token``/``bearer`` or a base64/JWT payload: the value is
"demonstrably sanitized" by shape, not by trust. Anything else (a bearer token, a session
cookie, an arbitrary debug string) is dropped silently.

Retry policy
------------
The seam itself performs **exactly one attempt**: it never replays a request. The retry
policy lives here, as pure decisions plus one bounded executor (:func:`run_with_retry`):

* only ``GET``/``HEAD``/``OPTIONS`` may be retried -- a ``send``, a ``create_*`` (POST), an
  ``upload`` (PUT) or any PATCH/DELETE is never replayed, even when the failure is a
  throttle that carries a ``Retry-After``;
* retries are bounded by ``max_attempts`` and additionally capped by
  :data:`MAX_ATTEMPTS_LIMIT`;
* ``Retry-After`` is honoured only when it parses safely (:func:`parse_retry_after`) and is
  then clamped to :data:`MAX_RETRY_AFTER_SECONDS`; an unparseable, negative, NaN, infinite,
  binary or ambiguous (multi-valued) value falls back to :data:`DEFAULT_BACKOFF_SECONDS`;
* the sleep callable is injectable, so a retry loop is testable without waiting.
"""
from __future__ import annotations

import math
import re
import time as _time
from dataclasses import dataclass
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from functools import lru_cache
from typing import Any, Callable, Mapping, NamedTuple

#: The closed set of reportable categories. ``tests/test_errors.py`` pins this set exactly.
CATEGORIES = frozenset(
    {
        "configuration_error",
        "validation_error",
        "authentication_required",
        "token_refresh_failed",
        "consent_required",
        "permission_denied",
        "not_found",
        "conflict",
        "precondition_failed",
        "throttled",
        "transport_error",
        "service_error",
        "internal_error",
        "unsupported_auth_mode",
        "operation_not_implemented",
    }
)

#: Static per-category messages. Never interpolated -- this is the leakage barrier.
MESSAGES: Mapping[str, str] = {
    "configuration_error": "the microsoft 365 plugin is not configured for this operation",
    "validation_error": "the operation arguments were rejected by microsoft graph",
    "authentication_required": "microsoft graph credentials are missing or unusable",
    "token_refresh_failed": "the microsoft graph access token could not be refreshed",
    "consent_required": "the tenant administrator has not consented to this operation",
    "permission_denied": "microsoft graph refused the operation for this identity",
    "not_found": "the requested microsoft 365 item does not exist or is not visible",
    "conflict": "the microsoft 365 item conflicts with an existing item",
    "precondition_failed": "the microsoft 365 item changed since it was read",
    "throttled": "microsoft graph throttled the request",
    "transport_error": "the microsoft graph request could not be completed",
    "service_error": "microsoft graph returned an error",
    "internal_error": "the microsoft 365 plugin failed internally",
    "unsupported_auth_mode": "the configured microsoft 365 authentication mode is not supported",
    "operation_not_implemented": "the microsoft 365 operation is not implemented",
}

#: Retrying is allowed for these methods only (idempotent, no side effect on replay).
SAFE_RETRY_METHODS = frozenset({"GET", "HEAD", "OPTIONS"})

#: Upper bound on the number of attempts a caller may request (fail closed above it).
MAX_ATTEMPTS_LIMIT = 5

#: Default number of attempts (the first attempt plus at most two retries).
DEFAULT_MAX_ATTEMPTS = 3

#: Ceiling for any honoured ``Retry-After`` value, in seconds.
MAX_RETRY_AFTER_SECONDS = 60.0

#: Backoff used when the failure is retryable but carries no usable ``Retry-After``.
DEFAULT_BACKOFF_SECONDS = 1.0

_RETRY_REASON_RETRY_AFTER = "retry_after"
_RETRY_REASON_TRANSIENT = "transient"
_RETRY_REASON_UNSAFE_METHOD = "method_not_safe_for_retry"
_RETRY_REASON_NOT_RETRYABLE = "category_not_retryable"
_RETRY_REASON_EXHAUSTED = "attempts_exhausted"

_SECONDS_PATTERN = re.compile(r"^\d{1,9}(?:\.\d{1,3})?$")
_GUID_PATTERN = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$")
_HEX_ID_PATTERN = re.compile(r"^[0-9a-f]{16,64}$")

_STATUS_CATEGORIES = {
    304: "precondition_failed",
    400: "validation_error",
    401: "authentication_required",
    403: "permission_denied",
    404: "not_found",
    409: "conflict",
    410: "not_found",
    412: "precondition_failed",
    413: "validation_error",
    415: "validation_error",
    422: "validation_error",
    428: "precondition_failed",
    429: "throttled",
}

_RETRYABLE_STATUS = frozenset({408, 425, 429, 500, 502, 503, 504})
_RETRYABLE_CATEGORIES = frozenset({"throttled", "transport_error"})

#: Structured error codes (normalized: lowercased, separators removed) that outrank the
#: HTTP status code. They are read from exception attributes only -- never parsed out of a
#: message -- and are used as lookup keys, never re-emitted.
_CODE_CATEGORIES = {
    # authentication
    "invalidauthenticationtoken": "authentication_required",
    "errorinvalidauthenticationtoken": "authentication_required",
    "unauthenticated": "authentication_required",
    "invalidclient": "authentication_required",
    # token lifecycle
    "invalidgrant": "token_refresh_failed",
    "invalidgranted": "token_refresh_failed",
    "tokenexpired": "token_refresh_failed",
    "expiredtoken": "token_refresh_failed",
    "invalidsignature": "token_refresh_failed",
    # consent / admin consent
    "consentrequired": "consent_required",
    "adminconsentrequired": "consent_required",
    "interactionrequired": "consent_required",
    # permissions
    "erroraccessdenied": "permission_denied",
    "authorizationrequestdenied": "permission_denied",
    "accessdenied": "permission_denied",
    "insufficientprivileges": "permission_denied",
    "forbidden": "permission_denied",
    # not found
    "erroritemnotfound": "not_found",
    "itemnotfound": "not_found",
    "errornotfound": "not_found",
    "resourcenotfound": "not_found",
    "notfound": "not_found",
    # conflict
    "conflict": "conflict",
    "namealreadyexists": "conflict",
    "resourcealreadyexists": "conflict",
    "errorresourcealreadyexists": "conflict",
    "concurrencyfailure": "conflict",
    # conditional requests
    "preconditionfailed": "precondition_failed",
    "errorpreconditionfailed": "precondition_failed",
    "etagmismatch": "precondition_failed",
    "resourcenotmodified": "precondition_failed",
    # throttling
    "toomanyrequests": "throttled",
    "activitylimitreached": "throttled",
    "throttled": "throttled",
    # request shape
    "invalidrequest": "validation_error",
    "errorinvalidrequest": "validation_error",
    "badrequest": "validation_error",
    "invalidargument": "validation_error",
    # capability
    "notsupported": "operation_not_implemented",
    "notimplemented": "operation_not_implemented",
    "operationnotsupported": "operation_not_implemented",
    "erroroperationnotsupported": "operation_not_implemented",
}

#: Header names consulted for a correlation id, in priority order.
_CORRELATION_HEADERS = ("client-request-id", "request-id", "x-ms-request-id")

#: Key-name fragments that mark a value as credential material wherever it is nested.
#: Defined once, here, and shared with ``microsoft365.results`` so redaction can never drift
#: between the error path and the result-normalization path.
_SENSITIVE_KEY_PARTS = (
    "secret",
    "token",
    "authorization",
    "cookie",
    "password",
    "credential",
    "bearer",
    "api_key",
    "apikey",
    "private_key",
    "signature",
)

#: Mapping keys that carry a whole request/response object rather than reportable data.
_STRUCTURAL_KEYS = frozenset(
    {"headers", "response_headers", "request_headers", "request_information", "response", "backing_store"}
)


def is_sensitive_key(key: Any) -> bool:
    """True when a mapping key names credential material (case/underscore insensitive)."""
    if isinstance(key, bytes):
        try:
            text = key.decode("ascii", "replace")
        except Exception:  # pragma: no cover - decode with replace cannot fail
            return True
    elif isinstance(key, str):
        text = key
    elif isinstance(key, (int, float)) and not isinstance(key, bool):
        text = str(key)
    else:
        return True
    normalized = text.lower().replace("-", "_")[:200]
    return any(part in normalized for part in _SENSITIVE_KEY_PARTS)


def is_structural_key(key: Any) -> bool:
    """True when a key carries a request/response object instead of reportable data."""
    return isinstance(key, str) and key.strip().lower() in _STRUCTURAL_KEYS


class GraphError(RuntimeError):
    """A sanitized, category-carrying failure.

    ``category`` is stable and safe to report; ``message`` always comes from
    :data:`MESSAGES`, so raw SDK/Graph/Azure text can never be part of it. The original
    exception is chained as ``__cause__`` (see :func:`to_graph_error`) for local debugging
    only -- it must never be returned to a tool caller.
    """

    def __init__(
        self,
        category: str,
        message: str,
        *,
        retryable: bool = False,
        retry_after_seconds: float | None = None,
        correlation_id: str | None = None,
        status_code: int | None = None,
    ):
        if category not in CATEGORIES:
            raise ValueError(f"unknown microsoft 365 error category: {category!r}")
        self.category = category
        self.message = str(message)
        self.retryable = bool(retryable)
        self.retry_after_seconds = retry_after_seconds
        self.correlation_id = correlation_id
        self.status_code = status_code
        super().__init__(f"{self.category}: {self.message}")

    def to_result(self) -> dict:
        """The reportable payload: category plus bounded, sanitized metadata only."""
        payload: dict[str, Any] = {
            "error": self.category,
            "message": self.message,
            "retryable": self.retryable,
        }
        if self.retry_after_seconds is not None:
            payload["retry_after_seconds"] = self.retry_after_seconds
        if self.correlation_id is not None:
            payload["correlation_id"] = self.correlation_id
        if self.status_code is not None:
            payload["status_code"] = self.status_code
        return payload


class RetryDecision(NamedTuple):
    """The explicit retry verdict for one failed attempt. Free of exception text."""

    retry: bool
    delay_seconds: float | None
    reason: str
    category: str


@dataclass(frozen=True)
class _Families:
    authentication: tuple[type, ...]
    transport: tuple[type, ...]
    decoding: tuple[type, ...]
    internal: tuple[type, ...]
    not_found: tuple[type, ...]
    conflict: tuple[type, ...]
    precondition: tuple[type, ...]


@lru_cache(maxsize=1)
def _families() -> _Families:
    """Resolve the installed exception families once, tolerating a partial environment.

    Only classes from the declared dependency set are used; a missing family degrades to an
    empty tuple so classification stays fail-closed (unknown -> ``internal_error``) instead
    of raising at import time.
    """
    authentication: list[type] = []
    transport: list[type] = []
    decoding: list[type] = []
    internal: list[type] = []
    not_found: list[type] = []
    conflict: list[type] = []
    precondition: list[type] = []

    try:
        from azure.core import exceptions as azure_exceptions
    except ImportError:  # pragma: no cover - azure-core ships with azure-identity
        azure_exceptions = None

    if azure_exceptions is not None:
        authentication.append(azure_exceptions.ClientAuthenticationError)
        transport.extend(
            [
                azure_exceptions.ServiceRequestError,
                azure_exceptions.ServiceRequestTimeoutError,
                azure_exceptions.ServiceResponseError,
                azure_exceptions.ServiceResponseTimeoutError,
                azure_exceptions.IncompleteReadError,
                azure_exceptions.StreamConsumedError,
                azure_exceptions.StreamClosedError,
                azure_exceptions.ResponseNotReadError,
            ]
        )
        decoding.extend([azure_exceptions.DecodeError, azure_exceptions.DeserializationError])
        internal.extend([azure_exceptions.SerializationError])
        not_found.append(azure_exceptions.ResourceNotFoundError)
        conflict.append(azure_exceptions.ResourceExistsError)
        precondition.extend(
            [azure_exceptions.ResourceNotModifiedError, azure_exceptions.ResourceModifiedError]
        )

    try:
        import kiota_http._exceptions as kiota_exceptions
    except ImportError:  # pragma: no cover - kiota-http is a declared dependency
        kiota_exceptions = None

    if kiota_exceptions is not None:
        transport.extend([kiota_exceptions.RequestError, kiota_exceptions.ResponseError, kiota_exceptions.RedirectError])
        decoding.append(kiota_exceptions.DeserializationError)
        internal.append(kiota_exceptions.BackingStoreError)

    return _Families(
        authentication=tuple(authentication),
        transport=tuple(transport),
        decoding=tuple(decoding),
        internal=tuple(internal),
        not_found=tuple(not_found),
        conflict=tuple(conflict),
        precondition=tuple(precondition),
    )


def _header_values(source: Any, name: str):
    """Return one header value, tolerating case differences and rejecting ambiguity."""
    if source is None:
        return None
    items = getattr(source, "items", None)
    if callable(items):
        try:
            pairs = list(items())
        except Exception:
            return None
        target = name.lower()
        for key, value in pairs:
            if isinstance(key, str) and key.lower() == target:
                return _single_value(value)
        return None
    get = getattr(source, "get", None)
    if callable(get):
        for key in (name, name.lower()):
            try:
                value = get(key)
            except Exception:
                return None
            if value is not None:
                return _single_value(value)
        return None
    return None


def _single_value(value: Any):
    if isinstance(value, (str, int, float)) and not isinstance(value, bool):
        return value
    if isinstance(value, (list, tuple, set, frozenset)):
        if len(value) == 1:
            return _single_value(next(iter(value)))
    # Multi-valued (ambiguous) or non-scalar values are not trusted.
    return None


def parse_retry_after(value: Any) -> float | None:
    """Parse a single ``Retry-After`` value, safely and within bounds.

    Accepts delta-seconds (int/float/string) and an HTTP-date. Returns ``None`` -- i.e. "do
    not honour this" -- for anything else: negatives, NaN, infinities, booleans, bytes,
    mappings, unparseable text and ambiguous multi-valued inputs. Every accepted value is
    clamped to :data:`MAX_RETRY_AFTER_SECONDS`; an HTTP-date already in the past is 0.
    """
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return _bounded_seconds(float(value))
    if not isinstance(value, str):
        return None

    token = value.strip()
    if not token or len(token) > 128:
        return None
    if _SECONDS_PATTERN.match(token):
        return _bounded_seconds(float(token))

    try:
        parsed = parsedate_to_datetime(token)
    except (TypeError, ValueError, OverflowError, IndexError):
        return None
    if not isinstance(parsed, datetime):
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return _bounded_seconds((parsed - datetime.now(timezone.utc)).total_seconds(), floor_at_zero=True)


def _bounded_seconds(seconds: float, *, floor_at_zero: bool = False) -> float | None:
    if not math.isfinite(seconds):
        return None
    if seconds < 0:
        return 0.0 if floor_at_zero else None
    return min(seconds, MAX_RETRY_AFTER_SECONDS)


def retry_after_seconds(headers: Any) -> float | None:
    """Read ``Retry-After`` from a header collection/dict, case-insensitively."""
    return parse_retry_after(_header_values(headers, "retry-after"))


def sanitize_correlation_id(value: Any) -> str | None:
    """Return the value only when it is a GUID/hex-shaped opaque identifier.

    Hex-only shape means an accepted id cannot embed credential-like words or a token
    payload, so no extra redaction is needed downstream.
    """
    if not isinstance(value, str):
        return None
    token = value.strip().lower()
    if _GUID_PATTERN.match(token) or _HEX_ID_PATTERN.match(token):
        return token
    return None


def _exception_headers(exc: BaseException):
    headers = getattr(exc, "response_headers", None)
    if headers is not None:
        return headers
    response = getattr(exc, "response", None)
    return getattr(response, "headers", None)


def _status_code(exc: BaseException) -> int | None:
    """A real HTTP status from the exception attributes, bounded and never coerced."""
    for attribute in ("response_status_code", "status_code"):
        value = getattr(exc, attribute, None)
        if isinstance(value, bool) or not isinstance(value, int):
            continue
        if 100 <= value <= 599:
            return value
    return None


def _error_code(exc: BaseException) -> str | None:
    """The structured error code, read from attributes only (never from a message)."""
    for attribute in ("error_code", "code"):
        value = getattr(exc, attribute, None)
        if isinstance(value, str) and value:
            return _normalize_code(value)
    nested = getattr(exc, "error", None)
    code = getattr(nested, "code", None)
    if isinstance(code, str) and code:
        return _normalize_code(code)
    return None


def _normalize_code(code: str) -> str:
    normalized = code.strip().lower().replace("_", "").replace("-", "").replace(" ", "")
    return normalized[:64]


def _upstream_category(exc: BaseException) -> str | None:
    """Category from an *upstream* signal, or ``None`` when the failure has none.

    Precedence, highest first: an already-typed :class:`GraphError`; the authentication
    families (``azure-identity`` / azure-core credential errors, because their status code
    can be misleading, e.g. a claims challenge arriving as 400); a bounded structured error
    code; the HTTP status code; azure-core resource subclasses; transport/decoding families;
    ``httpx`` transport errors.
    """
    if isinstance(exc, GraphError):
        return exc.category

    families = _families()
    if families.authentication and isinstance(exc, families.authentication):
        return "authentication_required"

    code_category = _CODE_CATEGORIES.get(_error_code(exc) or "")
    if code_category is not None:
        return code_category

    status = _status_code(exc)
    if status is not None:
        # An unmapped status is still an upstream service condition: report it as such and
        # let _retryable decide whether it may be replayed.
        return _STATUS_CATEGORIES.get(status, "service_error")

    if families.not_found and isinstance(exc, families.not_found):
        return "not_found"
    if families.conflict and isinstance(exc, families.conflict):
        return "conflict"
    if families.precondition and isinstance(exc, families.precondition):
        return "precondition_failed"

    if families.transport and isinstance(exc, families.transport):
        return "transport_error"
    if families.decoding and isinstance(exc, families.decoding):
        return "service_error"
    if families.internal and isinstance(exc, families.internal):
        return "internal_error"
    if type(exc).__module__.split(".")[0] == "httpx":
        return "transport_error"
    return None


def classify(exc: BaseException, *, default: str = "internal_error") -> str:
    """Map any exception to exactly one taxonomy category.

    Upstream signals decide first (:func:`_upstream_category`); when there are none, a local
    programming error is mapped from its type (``ValueError`` -> ``validation_error``,
    ``TypeError`` -> ``configuration_error``) and anything else becomes ``default``.
    """
    if not isinstance(exc, BaseException):
        raise TypeError("classify requires an exception instance")
    if default not in CATEGORIES:
        raise ValueError(f"unknown microsoft 365 error category: {default!r}")

    upstream = _upstream_category(exc)
    if upstream is not None:
        return upstream
    if isinstance(exc, ValueError):
        return "validation_error"
    if isinstance(exc, TypeError):
        return "configuration_error"
    return default


def classify_upstream(exc: BaseException, *, default: str) -> str:
    """Classify a failure that happened while talking to Graph, keeping the caller's default.

    Used by the execution seam: when the failure carries a real upstream signal (status,
    structured error code, SDK/Azure transport or credential family) the taxonomy decides;
    a purely local failure keeps the category the caller declared for that call site, so a
    seam never silently renames its own contract.
    """
    if not isinstance(exc, BaseException):
        raise TypeError("classify_upstream requires an exception instance")
    if default not in CATEGORIES:
        raise ValueError(f"unknown microsoft 365 error category: {default!r}")
    return _upstream_category(exc) or default


def _retryable(category: str, status: int | None) -> bool:
    if status is not None:
        return status in _RETRYABLE_STATUS
    return category in _RETRYABLE_CATEGORIES


def to_graph_error(
    exc: BaseException, *, default_category: str = "internal_error", category: str | None = None
) -> GraphError:
    """The single conversion point: any exception in, one sanitized :class:`GraphError` out.

    ``category`` lets a call site that already decided the category (the seam uses
    :func:`classify_upstream` + its own declaration) pass it in without this function
    re-classifying; ``default_category`` is the fallback used by :func:`classify` otherwise.
    The original exception is attached as ``__cause__`` (never as text), so callers can keep
    ``raise to_graph_error(exc)`` and still have the raw chain for local debugging.
    """
    if isinstance(exc, GraphError):
        return exc
    if not isinstance(exc, BaseException):
        raise TypeError("to_graph_error requires an exception instance")

    if category is None:
        category = classify(exc, default=default_category)
    elif category not in CATEGORIES:
        raise ValueError(f"unknown microsoft 365 error category: {category!r}")

    status = _status_code(exc)
    error = GraphError(
        category,
        MESSAGES[category],
        retryable=_retryable(category, status),
        retry_after_seconds=retry_after_seconds(_exception_headers(exc)),
        correlation_id=_correlation_id(exc),
        status_code=status,
    )
    error.__cause__ = exc
    error.__suppress_context__ = True
    return error


def error_result(exc: BaseException) -> dict:
    """The sanitized payload for a failure, safe to hand to a tool caller."""
    return to_graph_error(exc).to_result()


def _correlation_id(exc: BaseException) -> str | None:
    headers = _exception_headers(exc)
    for name in _CORRELATION_HEADERS:
        candidate = sanitize_correlation_id(_header_values(headers, name))
        if candidate is not None:
            return candidate
    return None


def _normalized_method(method: Any) -> str:
    if not isinstance(method, str):
        raise GraphError("configuration_error", MESSAGES["configuration_error"])
    normalized = method.strip().upper()
    if normalized not in {"GET", "POST", "PUT", "PATCH", "DELETE", "HEAD", "OPTIONS"}:
        raise GraphError("configuration_error", MESSAGES["configuration_error"])
    return normalized


def _validated_attempt(value: Any) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise GraphError("configuration_error", MESSAGES["configuration_error"])
    return value


def _validated_max_attempts(value: Any) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or not 1 <= value <= MAX_ATTEMPTS_LIMIT:
        raise GraphError("configuration_error", MESSAGES["configuration_error"])
    return value


def retry_decision(
    exc: BaseException,
    *,
    method: Any,
    attempt: Any,
    max_attempts: Any = DEFAULT_MAX_ATTEMPTS,
) -> RetryDecision:
    """Decide whether one failed attempt may be replayed, and after how long.

    The method gate is absolute: no ``POST``/``PUT``/``PATCH``/``DELETE`` is ever retried,
    whatever the category or the ``Retry-After`` header says. Invalid configuration (an
    unknown method, a non-integer or out-of-range attempt/limit) fails closed with a
    ``configuration_error`` instead of guessing.
    """
    normalized = _normalized_method(method)
    position = _validated_attempt(attempt)
    limit = _validated_max_attempts(max_attempts)
    error = to_graph_error(exc)

    if normalized not in SAFE_RETRY_METHODS:
        return RetryDecision(False, None, _RETRY_REASON_UNSAFE_METHOD, error.category)
    if not error.retryable:
        return RetryDecision(False, None, _RETRY_REASON_NOT_RETRYABLE, error.category)
    if position + 1 >= limit:
        return RetryDecision(False, None, _RETRY_REASON_EXHAUSTED, error.category)

    if error.retry_after_seconds is not None:
        return RetryDecision(True, error.retry_after_seconds, _RETRY_REASON_RETRY_AFTER, error.category)
    return RetryDecision(True, min(DEFAULT_BACKOFF_SECONDS, MAX_RETRY_AFTER_SECONDS), _RETRY_REASON_TRANSIENT, error.category)


def run_with_retry(
    action: Callable[[], Any],
    *,
    method: Any,
    max_attempts: Any = DEFAULT_MAX_ATTEMPTS,
    sleep: Callable[[float], Any] = _time.sleep,
) -> Any:
    """Run ``action`` at most ``max_attempts`` times, only when the policy allows a replay.

    This is the only retry loop this plugin has: the execution seam performs exactly one
    attempt, and a handler that wants to retry a read routes the attempt through here so the
    method gate and the ``Retry-After`` bound are never re-implemented (or forgotten)
    locally. A write is called once and its sanitized failure is raised immediately.
    """
    normalized = _normalized_method(method)
    limit = _validated_max_attempts(max_attempts)
    if not callable(action):
        raise GraphError("configuration_error", MESSAGES["configuration_error"])
    if not callable(sleep):
        raise GraphError("configuration_error", MESSAGES["configuration_error"])

    failure: GraphError | None = None
    for attempt in range(limit):
        try:
            return action()
        except Exception as exc:  # every failure is classified, never re-raised raw
            failure = to_graph_error(exc)
            decision = retry_decision(failure, method=normalized, attempt=attempt, max_attempts=limit)
            if not decision.retry:
                raise failure
            sleep(decision.delay_seconds)
    raise failure if failure is not None else GraphError("internal_error", MESSAGES["internal_error"])
