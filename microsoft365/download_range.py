"""Ranged drive downloads: files above the simple-transfer bound, read in verified ranges.

Why this module exists
----------------------
:mod:`microsoft365.files` implements the one-shot ``GET /content`` transfer, whose Graph
contract is bounded at 10 MiB. Anything larger is currently refused with
``download_range_not_implemented``. That is the missing half of the large-file story now that
:mod:`microsoft365.upload_session` handles the write side, and a plugin that can upload a 2 GiB
video but not read it back is not production-ready.

Contract (verified against Microsoft Graph driveItem "download contents")
-------------------------------------------------------------------------
1. Read the item's authenticated metadata first: ``GET /drives/{drive-id}/items/{address}``.
   Its ``size``, ``file.mimeType`` and ``name`` are the only sources for the transfer's size
   bound, media type and name -- a caller never supplies any of them.
2. Fetch ``GET /drives/{drive-id}/items/{address}/content`` with
   ``Range: bytes={start}-{end-1}``, one aligned range at a time.
3. A range response carries ``206 Partial Content`` and a ``Content-Range: bytes s-e/total``
   header. This module checks that the returned window is **the window it asked for**, that
   ranges are contiguous and non-overlapping, and that the assembled total equals the declared
   size. A mismatch is a failure, never a silently short or duplicated file.

Security and memory invariants
------------------------------
* Every header this module sets is built from integers it computed; no caller string ever
  reaches a header, so no CRLF injection is possible through a range.
* A declared size above :data:`MAX_DOWNLOAD_BYTES` is refused **before** the first byte request,
  and the accumulated payload is checked against the bound as it grows, so a lying server cannot
  make this buffer without limit.
* Every failure is a sanitized taxonomy error with a static message plus a bounded ``status``;
  raw Graph text stays in ``__cause__``.

Everything defaults to an injectable seam (``fetch``) so the module is fully exercisable
offline; production passes nothing and gets the real seam.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable

from kiota_abstractions.base_request_configuration import RequestConfiguration
from kiota_abstractions.headers_collection import HeadersCollection

from . import execution
from .errors import MESSAGES, GraphError
from .files import (
    DEFAULT_PATH_IDENTIFIER,
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

INVALID_RANGE_SIZE = "invalid_range_size"
DOWNLOAD_TOO_LARGE = "download_too_large"
RANGE_NOT_CONTENT_PARTIAL = "range_not_content_partial"
RANGE_WINDOW_MISMATCH = "range_window_mismatch"
RANGE_OUT_OF_BOUNDS = "range_out_of_bounds"
RANGE_EMPTY = "range_empty"
RANGE_FAILED = "range_failed"
DOWNLOAD_SIZE_MISMATCH = "download_size_mismatch"
RANGE_SIZE_MISMATCH = "download_size_mismatch"
DOWNLOAD_INCOMPLETE = "download_incomplete"

_ACTIONS = {
    INVALID_RANGE_SIZE: "use a range size that is a multiple of 320 KiB",
    DOWNLOAD_TOO_LARGE: "download the file in sections, or address a smaller file",
    RANGE_NOT_CONTENT_PARTIAL: "retry the download; the ranged read was not honored",
    RANGE_WINDOW_MISMATCH: "retry the download; the returned bytes are not the requested range",
    RANGE_OUT_OF_BOUNDS: "re-read the item metadata and retry the download",
    RANGE_EMPTY: "retry the download; a range returned no bytes",
    RANGE_FAILED: "retry the download; already read ranges are re-fetched",
    RANGE_SIZE_MISMATCH: "re-read the item metadata and retry the download",
    DOWNLOAD_INCOMPLETE: "the transfer ended before every declared byte was read",
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


def download_file_auto(
    client: Any,
    *,
    drive_id: Any,
    drive_item_id: Any = None,
    path: Any = None,
    identifier: str = DEFAULT_PATH_IDENTIFIER,
    range_size: Any = None,
    execute: Callable[..., Any] = execution.execute_request,
    fetch: FetchSeam | None = None,
    sleep: Callable[[float], None] | None = None,
) -> BinaryResult:
    """Download one drive file, choosing one-shot or ranged reads by its declared size.

    ``auto`` names the *strategy selection*: the caller's only optional input is
    ``range_size``, and the mechanism is the smallest one that can carry the file. Both paths
    return the same :class:`~microsoft365.results.BinaryResult`, so the caller reconciles one
    result contract either way.

    * declared size within the simple bound -> one ``GET /content`` with no ``Range`` header;
    * anything larger, up to Graph's ceiling -> aligned ranged reads.
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
    if declared_size <= SIMPLE_TRANSFER_LIMIT_BYTES:
        # The metadata read above is the same authenticated read ``download_file`` would do,
        # so the content request is issued here instead of re-reading the item.
        payload = execute(_content_builder(client, drive, address), method="GET", adapter=adapter)
        return binary_result(
            _stream_response(payload, limit_bytes=SIMPLE_TRANSFER_LIMIT_BYTES),
            content_type=mime_type,
            name=name,
        )

    window_size = validate_range_size(DEFAULT_RANGE_SIZE if range_size is None else range_size)
    resolve_fetch = fetch if fetch is not None else fetch_range_default
    resolve_sleep = sleep if sleep is not None else sleep_default
    payload = download_ranges(
        size=declared_size,
        range_size=window_size,
        fetch=lambda start, end: resolve_fetch(
            client,
            drive_id=drive,
            drive_item_id=address,
            start=start,
            end=end,
            execute=execute,
            adapter=adapter,
        ),
        sleep=resolve_sleep,
    )
    return binary_result(payload, content_type=mime_type, name=name)


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
    """The production ranged-read seam: one ``GET`` per window with a ``Range`` header."""
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