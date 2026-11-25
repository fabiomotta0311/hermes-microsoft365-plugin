"""Bounded-streaming small-file download, and experimental multi-range helpers.

Status
------
:mod:`microsoft365.files` implements the one-shot ``GET /content`` transfer, whose Graph
contract is bounded at 10 MiB. A declared-size fast path used to serve files at or under that
bound directly; review found the guard insufficient because the SDK's content builder can
buffer an unbounded response regardless of what the metadata declares, before this module
ever got a chance to enforce a limit, and the SDK does not expose a way to stream that
builder's response incrementally.

This module now implements a real bounded-streaming transport for small files
(``size <= SMALL_DOWNLOAD_BUDGET_BYTES``) that never goes through the SDK's content builder:

1. Read the item's authenticated metadata through the SDK: ``GET /drives/{drive}/items/{addr}``.
   ``size``, ``file.mimeType``, ``name``, ``id`` and ``eTag`` come **only** from this response.
2. Reject anything above :data:`SMALL_DOWNLOAD_BUDGET_BYTES` before any byte is requested: the
   base64 result carrier buffers the whole file in memory, and there is no bounded host artifact
   sink yet to stream a larger file to, so that remains refused with
   ``download_range_not_implemented`` -- honestly incomplete, not silently truncated.
3. Take the item's ``@microsoft.graph.downloadUrl`` (a preauthenticated, time-limited URL;
   Microsoft's docs: a ``Range`` header must be sent to *this* URL, not to ``/content`` --
   and a server that ignores it answers ``200`` with the whole file instead of ``206``).
4. Fetch that URL with a **separate** ``httpx`` client that carries no Graph auth header, no
   cookies, ``trust_env=False`` (no ambient proxy/cert config) and ``follow_redirects=False``
   (a redirect is refused outright rather than silently followed to another host).
5. Send ``Range: bytes=0-{size-1}`` for the whole declared window and require ``206`` with a
   ``Content-Range`` that echoes exactly that window; a ``200`` (range ignored) or a wrong
   window is refused, never silently accepted as the right bytes.
6. Enforce the byte budget twice: before reading, a declared ``Content-Length`` over the
   budget is refused without touching the body; while reading, the stream is read in chunks
   via ``httpx``'s true streaming (``client.stream(...)``, not the eager ``.get()``) and
   aborted the instant accumulated bytes exceed the budget, even with no ``Content-Length``
   at all (a chunked or lying server cannot force an unbounded buffer).
7. Re-read the authenticated metadata (id, size, eTag) **after** the download and refuse the
   result if any of the three changed -- a item replaced or edited mid-download must never be
   returned as if it were the version whose metadata and bytes were both read.

The experimental multi-range helpers below (:func:`download_ranges`, :func:`fetch_range_default`)
remain disconnected from the handler: they still rely on the SDK's content builder for an
individual range, which is not proven to stream, and are not a supported production transport
for anything above the small-file budget.

Security and memory invariants
------------------------------
* Every header this module sets is built from integers or a URL read from authenticated Graph
  metadata; no caller string ever reaches a header, so no CRLF injection is possible.
* The small-file path streams through ``httpx`` with an explicit byte ceiling enforced during
  the read, never after an unbounded buffer has already been filled.
* Every failure is a sanitized taxonomy error with a static message plus a bounded ``status``;
  raw Graph/transport text stays in ``__cause__``.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable
from urllib.parse import urlsplit

import httpx
from kiota_abstractions.base_request_configuration import RequestConfiguration
from kiota_abstractions.headers_collection import HeadersCollection

from . import execution
from .errors import MESSAGES, GraphError
from .files import (
    DEFAULT_PATH_IDENTIFIER,
    METADATA_NOT_A_DRIVE_ITEM,
    SIMPLE_TRANSFER_LIMIT_BYTES,
    TransferError,
    _content_builder,
    _declared_metadata,
    _empty_configuration,
    _item_builder,
    _request_adapter,
    _response_header,
    _stream_response,
    drive_item_address,
    validate_drive_id,
)
from .results import BinaryResult, binary_result

#: Graph's ceiling for a single download of a drive item's contents.
MAX_DOWNLOAD_BYTES = 250 * 1024 * 1024 * 1024

#: Graph recommends aligning ranged download requests to 320 KiB.
RANGE_SIZE_MULTIPLE = 320 * 1024

#: 10 MiB: the same window the simple transfer uses, so the first ranged read of a large file
#: costs exactly what a small file costs.
DEFAULT_RANGE_SIZE = 10 * 1024 * 1024

#: Hard ceiling on the ranges one download may issue, so a server that answers every range
#: with an empty body cannot spin the transfer forever.
MAX_RANGES = 10_000

#: Bound on replayed attempts for one range before the download is reported as failed.
MAX_RANGE_ATTEMPTS = 5

#: The only size at which this plugin can currently return a whole file safely: small enough
#: that a base64 in-memory result (``BinaryResult``) is a reasonable carrier. There is no
#: bounded host artifact sink to stream a larger file to yet, so anything above this remains
#: refused rather than silently truncated or fully buffered against a weaker guarantee.
SMALL_DOWNLOAD_BUDGET_BYTES = 2 * 1024 * 1024

#: The additional, undeclared property Graph attaches to item metadata: a preauthenticated,
#: time-limited URL that itself honors ``Range`` (unlike the SDK's ``/content`` builder).
DOWNLOAD_URL_FIELD = "@microsoft.graph.downloadUrl"

#: A conservative timeout for the small-file transfer: the whole point of the size budget is
#: that this request is small, so it should not need to hang indefinitely.
DOWNLOAD_TIMEOUT_SECONDS = 30.0

INVALID_RANGE_SIZE = "invalid_range_size"
DOWNLOAD_TOO_LARGE = "download_too_large"
RANGE_NOT_CONTENT_PARTIAL = "range_not_content_partial"
RANGE_WINDOW_MISMATCH = "range_window_mismatch"
RANGE_OUT_OF_BOUNDS = "range_out_of_bounds"
RANGE_EMPTY = "range_empty"
RANGE_FAILED = "range_failed"
DOWNLOAD_SIZE_MISMATCH = "download_size_mismatch"
DOWNLOAD_RANGE_NOT_IMPLEMENTED = "download_range_not_implemented"
DOWNLOAD_URL_UNAVAILABLE = "download_url_unavailable"
DOWNLOAD_URL_INVALID = "download_url_invalid"
DOWNLOAD_REDIRECT_REFUSED = "download_redirect_refused"
METADATA_WITHOUT_IDENTITY = "metadata_without_identity"
METADATA_CHANGED_MID_DOWNLOAD = "metadata_changed_mid_download"

_ACTIONS = {
    INVALID_RANGE_SIZE: "use a range size that is a multiple of 320 KiB",
    DOWNLOAD_TOO_LARGE: "download the file in sections, or address a smaller file",
    RANGE_NOT_CONTENT_PARTIAL: "retry the download; the ranged read was not honored",
    RANGE_WINDOW_MISMATCH: "retry the download; the returned bytes are not the requested range",
    RANGE_OUT_OF_BOUNDS: "re-read the item metadata and retry the download",
    RANGE_EMPTY: "retry the download; a range returned no bytes",
    RANGE_FAILED: "retry the download; already read ranges are re-fetched",
    DOWNLOAD_RANGE_NOT_IMPLEMENTED: "download a file at or under the small-file bound",
    DOWNLOAD_URL_UNAVAILABLE: "re-read the item metadata and retry the download",
    DOWNLOAD_URL_INVALID: "re-read the item metadata and retry the download",
    DOWNLOAD_REDIRECT_REFUSED: "retry the download; the preauthenticated URL redirected",
    METADATA_WITHOUT_IDENTITY: "re-read the item metadata and retry the download",
    METADATA_CHANGED_MID_DOWNLOAD: "the file changed during download; retry",
}

FetchSeam = Callable[..., bytes]


class DownloadRangeError(GraphError):
    """A sanitized ranged-download failure carrying an explicit ``status``."""

    def __init__(
        self,
        category: str,
        *,
        status: str,
        action: str | None = None,
        limit_bytes: int | None = None,
        size_bytes: int | None = None,
        range_start: int | None = None,
        range_end: int | None = None,
        attempts: int | None = None,
    ):
        super().__init__(category, MESSAGES[category])
        self.status = str(status)
        self.action = action if action is not None else _ACTIONS.get(self.status)
        self.limit_bytes = limit_bytes
        self.size_bytes = size_bytes
        self.range_start = range_start
        self.range_end = range_end
        self.attempts = attempts

    def to_result(self) -> dict:
        payload = super().to_result()
        payload["status"] = self.status
        if self.action:
            payload["action"] = self.action
        for name in ("limit_bytes", "size_bytes", "range_start", "range_end", "attempts"):
            value = getattr(self, name)
            if value is not None:
                payload[name] = value
        return payload


def validate_range_size(range_size: Any) -> int:
    """A range window that is a multiple of 320 KiB and no larger than the simple bound."""
    if isinstance(range_size, bool) or not isinstance(range_size, int):
        raise DownloadRangeError("validation_error", status=INVALID_RANGE_SIZE)
    if not RANGE_SIZE_MULTIPLE <= range_size <= SIMPLE_TRANSFER_LIMIT_BYTES:
        raise DownloadRangeError("validation_error", status=INVALID_RANGE_SIZE, size_bytes=range_size)
    if range_size % RANGE_SIZE_MULTIPLE:
        raise DownloadRangeError("validation_error", status=INVALID_RANGE_SIZE, size_bytes=range_size)
    return range_size


def validate_declared_size(size: Any) -> int:
    """The declared size must be a real file size this plugin is willing to buffer."""
    if isinstance(size, bool) or not isinstance(size, int) or size < 0:
        raise TransferError("configuration_error", status="metadata_without_size")
    if size > MAX_DOWNLOAD_BYTES:
        raise DownloadRangeError(
            "validation_error", status=DOWNLOAD_TOO_LARGE, size_bytes=size, limit_bytes=MAX_DOWNLOAD_BYTES
        )
    return size


def planned_ranges(size: int, range_size: int) -> tuple[tuple[int, int], ...]:
    """The exact ``(start, end_inclusive)`` windows a download of ``size`` bytes will request."""
    if size == 0:
        return ()
    return tuple(
        (start, min(start + range_size, size) - 1)
        for start in range(0, size, range_size)
    )


@dataclass(frozen=True)
class RangeOutcome:
    """What one ranged read returned: its bytes, the HTTP status and its ``Content-Range``."""

    status: int
    data: bytes
    content_range: str | None = None


def parse_content_range(header: Any) -> tuple[int, int, int] | None:
    """Parse ``bytes {start}-{end}/{total}`` into integers, or ``None`` when unusable.

    The header is the only proof of *which* window Graph returned. A malformed, absent or
    non-integer header returns ``None`` and the caller treats that as a window mismatch rather
    than assuming the request was honored.
    """
    if not isinstance(header, str):
        return None
    text = header.strip()
    if not text.lower().startswith("bytes "):
        return None
    spec = text[6:].strip().partition("/")
    window = spec[0].strip()
    total = spec[2].strip()
    if not total.isdigit():
        return None
    bounds = window.partition("-")
    if not bounds[0].strip().isdigit() or not bounds[2].strip().isdigit():
        return None
    return int(bounds[0]), int(bounds[2]), int(total)


def download_ranges(
    *,
    size: int,
    range_size: int,
    fetch: FetchSeam,
    sleep: Callable[[float], None],
    max_attempts: int = MAX_RANGE_ATTEMPTS,
) -> bytes:
    """Read every planned window and return the exact concatenation, or fail loudly.

    ``fetch(start, end)`` returns a :class:`RangeOutcome`. A transport exception is retried with
    the caller's backoff; a non-``206`` status, a wrong window, an out-of-bounds window or an
    empty window are contract failures and are reported immediately -- the module never pads,
    truncates or reorders bytes to make a broken transfer look complete.
    """
    total = validate_declared_size(size)
    window_size = validate_range_size(range_size)
    windows = planned_ranges(total, window_size)
    if len(windows) > MAX_RANGES:
        raise DownloadRangeError(
            "validation_error", status=DOWNLOAD_TOO_LARGE, size_bytes=total, limit_bytes=MAX_DOWNLOAD_BYTES
        )

    data = bytearray()
    offset = 0
    for start, end in windows:
        chunk = _read_range(
            fetch=fetch,
            sleep=sleep,
            start=start,
            end=end,
            total=total,
            max_attempts=max_attempts,
        )
        expected = end - start + 1
        if len(chunk) != expected:
            raise DownloadRangeError(
                "precondition_failed",
                status=RANGE_WINDOW_MISMATCH,
                range_start=start,
                range_end=end,
                size_bytes=len(chunk),
            )
        data.extend(chunk)
        offset += len(chunk)
        if offset > MAX_DOWNLOAD_BYTES:
            raise DownloadRangeError(
                "validation_error", status=DOWNLOAD_TOO_LARGE, size_bytes=offset, limit_bytes=MAX_DOWNLOAD_BYTES
            )

    if offset != total:
        raise DownloadRangeError(
            "precondition_failed", status=DOWNLOAD_SIZE_MISMATCH, size_bytes=offset, limit_bytes=total
        )
    return bytes(data)


def _read_range(
    *,
    fetch: FetchSeam,
    sleep: Callable[[float], None],
    start: int,
    end: int,
    total: int,
    max_attempts: int,
) -> bytes:
    """Fetch one window, retrying transport failures within a bounded attempt count."""
    attempts = 0
    while True:
        try:
            outcome = fetch(start, end)
        except Exception as exc:  # noqa: BLE001 - transport failures are retried, then reported
            attempts += 1
            if attempts >= max_attempts:
                raise DownloadRangeError(
                    "service_error",
                    status=RANGE_FAILED,
                    attempts=attempts,
                    range_start=start,
                    range_end=end,
                ) from exc
            sleep(min(2.0 * attempts, 8.0))
            continue

        if not isinstance(outcome, RangeOutcome):
            raise DownloadRangeError(
                "service_error", status=RANGE_FAILED, range_start=start, range_end=end
            )
        if outcome.status != 206:
            raise DownloadRangeError(
                "service_error", status=RANGE_NOT_CONTENT_PARTIAL, range_start=start, range_end=end
            )
        if not outcome.data:
            raise DownloadRangeError(
                "precondition_failed", status=RANGE_EMPTY, range_start=start, range_end=end
            )

        window = parse_content_range(outcome.content_range)
        if window is None:
            raise DownloadRangeError(
                "precondition_failed", status=RANGE_WINDOW_MISMATCH, range_start=start, range_end=end
            )
        returned_start, returned_end, returned_total = window
        if returned_start != start or returned_end != end:
            raise DownloadRangeError(
                "precondition_failed", status=RANGE_WINDOW_MISMATCH, range_start=start, range_end=end
            )
        if returned_total != total:
            raise DownloadRangeError(
                "precondition_failed",
                status=RANGE_OUT_OF_BOUNDS,
                size_bytes=returned_total,
                range_start=start,
                range_end=end,
            )
        return bytes(outcome.data)


def _item_identity(item: Any) -> tuple[str, str]:
    """The ``(id, eTag)`` pair that pins the exact version a download must match.

    ``_declared_metadata`` already refuses anything that is not a ``DriveItem``; this reads
    the two identity fields it does not, and refuses a response with no usable id or eTag --
    neither can be used to detect a mid-download change if either is missing or blank.
    """
    from msgraph.generated.models.drive_item import DriveItem

    if not isinstance(item, DriveItem):
        raise TransferError("configuration_error", status=METADATA_NOT_A_DRIVE_ITEM)
    item_id = item.id
    etag = item.e_tag
    if not isinstance(item_id, str) or not item_id.strip():
        raise DownloadRangeError("configuration_error", status=METADATA_WITHOUT_IDENTITY)
    if not isinstance(etag, str) or not etag.strip():
        raise DownloadRangeError("configuration_error", status=METADATA_WITHOUT_IDENTITY)
    return item_id, etag


def _download_url(item: Any) -> str:
    """The preauthenticated download URL Graph attaches to item metadata, validated as https.

    This is read from ``additional_data`` because Kiota does not model an undeclared OData
    property; a caller never supplies this value, so no injected string can reach it.
    """
    data = getattr(item, "additional_data", None)
    url = data.get(DOWNLOAD_URL_FIELD) if isinstance(data, dict) else None
    if not isinstance(url, str) or not url.strip():
        raise DownloadRangeError("configuration_error", status=DOWNLOAD_URL_UNAVAILABLE)
    parsed = urlsplit(url)
    if parsed.scheme != "https" or not parsed.netloc:
        raise DownloadRangeError("configuration_error", status=DOWNLOAD_URL_INVALID)
    return url


def http_client_default() -> httpx.Client:
    """A transport isolated from the Graph SDK client: no auth header, no cookies, no proxy.

    ``trust_env=False`` means no ambient ``HTTP_PROXY``/certificate environment is consulted,
    and ``follow_redirects=False`` means a redirect is reported as a response this module
    inspects and refuses, rather than silently chased to a second, untrusted host.
    """
    return httpx.Client(
        trust_env=False,
        follow_redirects=False,
        timeout=httpx.Timeout(DOWNLOAD_TIMEOUT_SECONDS),
    )


def _bounded_body(response: httpx.Response, *, limit_bytes: int) -> bytes:
    """Read a streamed response's body in chunks, aborting past ``limit_bytes``.

    This is real streaming: ``response`` comes from ``client.stream(...)``, so the body is
    not already fully buffered when this function is called, and a chunk that would push the
    running total past the budget raises before the next chunk is ever requested -- a server
    with no (or a lying) ``Content-Length`` cannot force an unbounded buffer here.
    """
    data = bytearray()
    for chunk in response.iter_bytes():
        data.extend(chunk)
        if len(data) > limit_bytes:
            raise DownloadRangeError(
                "validation_error",
                status=DOWNLOAD_TOO_LARGE,
                size_bytes=len(data),
                limit_bytes=limit_bytes,
            )
    return bytes(data)


def _read_preauth_response(response: httpx.Response, *, size: int) -> bytes:
    """Validate status/``Content-Range``/``Content-Length`` and read the bounded body."""
    status = response.status_code
    if 300 <= status < 400:
        raise DownloadRangeError("configuration_error", status=DOWNLOAD_REDIRECT_REFUSED)

    if size == 0:
        if status not in (200, 204, 206):
            raise DownloadRangeError("service_error", status=RANGE_NOT_CONTENT_PARTIAL)
        data = _bounded_body(response, limit_bytes=SMALL_DOWNLOAD_BUDGET_BYTES)
        if data:
            raise DownloadRangeError(
                "precondition_failed", status=DOWNLOAD_SIZE_MISMATCH, size_bytes=len(data)
            )
        return b""

    if status != 206:
        # A server that ignores Range answers 200 with the whole file. Treated identically
        # to any other non-206 status: never accepted as "close enough".
        raise DownloadRangeError("service_error", status=RANGE_NOT_CONTENT_PARTIAL)

    window = parse_content_range(response.headers.get("Content-Range"))
    if window is None:
        raise DownloadRangeError("precondition_failed", status=RANGE_WINDOW_MISMATCH)
    returned_start, returned_end, returned_total = window
    if returned_start != 0 or returned_end != size - 1 or returned_total != size:
        raise DownloadRangeError("precondition_failed", status=RANGE_WINDOW_MISMATCH)

    declared_length_raw = response.headers.get("Content-Length")
    if declared_length_raw is not None:
        try:
            declared_length = int(declared_length_raw)
        except ValueError:
            declared_length = None
        if declared_length is not None and declared_length > SMALL_DOWNLOAD_BUDGET_BYTES:
            raise DownloadRangeError(
                "validation_error",
                status=DOWNLOAD_TOO_LARGE,
                size_bytes=declared_length,
                limit_bytes=SMALL_DOWNLOAD_BUDGET_BYTES,
            )

    data = _bounded_body(response, limit_bytes=SMALL_DOWNLOAD_BUDGET_BYTES)
    if len(data) != size:
        raise DownloadRangeError(
            "precondition_failed", status=DOWNLOAD_SIZE_MISMATCH, size_bytes=len(data), limit_bytes=size
        )
    return data


def fetch_preauthenticated(url: str, *, size: int, http_client: httpx.Client) -> bytes:
    """GET the preauthenticated download URL for the whole declared window, bounded.

    ``http_client`` is a seam: production uses :func:`http_client_default`, tests use an
    ``httpx.Client`` built on ``httpx.MockTransport`` so the real-seam contract (status,
    headers, chunked body) is exercised without any network access.
    """
    headers = {"Range": f"bytes=0-{size - 1}"} if size > 0 else {}
    try:
        with http_client.stream("GET", url, headers=headers) as response:
            return _read_preauth_response(response, size=size)
    except DownloadRangeError:
        raise
    except httpx.HTTPError as exc:
        raise DownloadRangeError("service_error", status=RANGE_FAILED) from exc


def download_file_auto(
    client: Any,
    *,
    drive_id: Any,
    drive_item_id: Any = None,
    path: Any = None,
    identifier: str = DEFAULT_PATH_IDENTIFIER,
    execute: Callable[..., Any] = execution.execute_request,
    http_client_factory: Callable[[], httpx.Client] = http_client_default,
) -> BinaryResult:
    """Download one small drive file through a bounded-streaming, version-pinned transport.

    Files at or under :data:`SMALL_DOWNLOAD_BUDGET_BYTES` are fetched for real: authenticated
    metadata pins the item's id/eTag, the preauthenticated ``downloadUrl`` is read with a
    separate unauthenticated ``httpx`` client under a strict byte budget enforced during the
    streamed read (not after an unbounded buffer), and the metadata is re-read afterwards to
    refuse a result whose id, eTag or size changed mid-download.

    Anything larger still fails closed with ``download_range_not_implemented``: the result
    carrier (:class:`BinaryResult`) buffers and base64-encodes the whole file, and there is no
    bounded host artifact sink yet to stream a larger file to -- refusing it is honest, not a
    silent truncation or an unbounded in-memory buffer.
    """
    drive = validate_drive_id(drive_id)
    address = drive_item_address(drive_item_id=drive_item_id, path=path, identifier=identifier)
    adapter = _request_adapter(client)

    item = execute(
        _item_builder(client, drive, address),
        method="GET",
        configuration=_empty_configuration(),
        adapter=adapter,
    )
    declared_size, mime_type, name = _declared_metadata(item)

    if declared_size > SMALL_DOWNLOAD_BUDGET_BYTES:
        raise DownloadRangeError(
            "operation_not_implemented", status=DOWNLOAD_RANGE_NOT_IMPLEMENTED,
            size_bytes=declared_size, limit_bytes=SMALL_DOWNLOAD_BUDGET_BYTES,
        )

    item_id, etag = _item_identity(item)
    url = _download_url(item)

    http_client = http_client_factory()
    try:
        data = fetch_preauthenticated(url, size=declared_size, http_client=http_client)
    finally:
        http_client.close()

    refreshed = execute(
        _item_builder(client, drive, address),
        method="GET",
        configuration=_empty_configuration(),
        adapter=adapter,
    )
    refreshed_size, _refreshed_mime, _refreshed_name = _declared_metadata(refreshed)
    refreshed_id, refreshed_etag = _item_identity(refreshed)
    if refreshed_id != item_id or refreshed_etag != etag or refreshed_size != declared_size:
        raise DownloadRangeError(
            "precondition_failed", status=METADATA_CHANGED_MID_DOWNLOAD, size_bytes=refreshed_size
        )

    return binary_result(data, content_type=mime_type, name=name)


def _range_configuration(start: int, end: int) -> RequestConfiguration:
    """``Range`` is built from computed integers only, so no caller text can reach the header."""
    headers = HeadersCollection()
    headers.add("Range", f"bytes={start}-{end}")
    return RequestConfiguration(headers=headers)


def fetch_range_default(
    client: Any,
    *,
    drive_id: str,
    drive_item_id: str,
    start: int,
    end: int,
    execute: Callable[..., Any] = execution.execute_request,
    adapter: Any = None,
) -> RangeOutcome:
    """Experimental seam; not connected to the production download handler."""
    response = execute(
        _content_builder(client, drive_id, drive_item_id),
        method="GET",
        configuration=_range_configuration(start, end),
        adapter=adapter,
    )
    status = getattr(response, "status_code", None)
    if not isinstance(status, int):
        raise DownloadRangeError(
            "configuration_error", status=RANGE_NOT_CONTENT_PARTIAL, range_start=start, range_end=end
        )
    data = _stream_response(response, limit_bytes=end - start + 1)
    return RangeOutcome(
        status=status,
        data=data,
        content_range=_response_header(response, "Content-Range"),
    )


def sleep_default(seconds: float) -> None:
    """Bounded backoff between range retries."""
    import time

    time.sleep(seconds)