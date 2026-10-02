"""Offline tests for ranged drive downloads (``microsoft365.download_range``).

The contract under test is integrity: a ranged download must return *exactly* the declared
bytes, or fail. Each adversarial case here proves the module refuses a transfer that a naive
implementation would have silently accepted as complete.
"""
from __future__ import annotations

from urllib.parse import urlsplit

import pytest
from kiota_abstractions.request_adapter import RequestAdapter

from microsoft365 import execution
from microsoft365.download_range import (
    DEFAULT_RANGE_SIZE,
    MAX_DOWNLOAD_BYTES,
    MAX_RANGES,
    RANGE_SIZE_MULTIPLE,
    SIMPLE_TRANSFER_LIMIT_BYTES,
    DownloadRangeError,
    RangeOutcome,
    download_file_auto,
    download_ranges,
    parse_content_range,
    planned_ranges,
    validate_range_size,
)

MB = 1024 * 1024

#: Every ranged window in production is a multiple of 320 KiB, so the offline corpus uses that
#: same window rather than an artificially small one. Three windows make the corpus ~960 KiB.
WINDOW = RANGE_SIZE_MULTIPLE
RANGE_COUNT = 3
PAYLOAD = bytes(range(256)) * (WINDOW * RANGE_COUNT // 256)
assert len(PAYLOAD) == WINDOW * RANGE_COUNT


def fetcher(payload: bytes, *, calls: list | None = None):
    """A correct ranged reader over ``payload``."""

    def fetch(start: int, end: int) -> RangeOutcome:
        if calls is not None:
            calls.append((start, end))
        window = payload[start : end + 1]
        return RangeOutcome(
            status=206, data=window, content_range=f"bytes {start}-{end}/{len(payload)}"
        )

    return fetch


def no_sleep(seconds: float) -> None:
    return None


# --------------------------------------------------------------------------- range size


def test_a_range_size_must_be_a_multiple_of_320_kib():
    assert validate_range_size(RANGE_SIZE_MULTIPLE) == RANGE_SIZE_MULTIPLE
    assert validate_range_size(DEFAULT_RANGE_SIZE) == DEFAULT_RANGE_SIZE
    for bad in (RANGE_SIZE_MULTIPLE - 1, RANGE_SIZE_MULTIPLE + 1, 1, 0):
        with pytest.raises(DownloadRangeError) as caught:
            validate_range_size(bad)
        assert caught.value.status == "invalid_range_size"


def test_a_range_size_above_the_simple_bound_is_refused():
    with pytest.raises(DownloadRangeError) as caught:
        validate_range_size(SIMPLE_TRANSFER_LIMIT_BYTES + RANGE_SIZE_MULTIPLE)
    assert caught.value.status == "invalid_range_size"


@pytest.mark.parametrize("bad", [True, False, None, "1048576", 10.5, b"x"])
def test_a_non_integer_range_size_is_refused(bad):
    with pytest.raises(DownloadRangeError) as caught:
        validate_range_size(bad)
    assert caught.value.status == "invalid_range_size"


# --------------------------------------------------------------------------- planning


def test_the_last_planned_range_is_short_and_ends_on_the_last_byte():
    windows = planned_ranges(25, 10)
    assert windows == ((0, 9), (10, 19), (20, 24))
    assert windows[-1][1] == 24


def test_an_exact_multiple_plans_no_empty_trailing_range():
    assert planned_ranges(20, 10) == ((0, 9), (10, 19))


def test_an_empty_file_plans_no_range_at_all():
    assert planned_ranges(0, 10) == ()


def test_a_size_that_would_need_too_many_ranges_is_refused_before_any_read():
    # 250 GiB at 320 KiB windows is far past MAX_RANGES, so the download is refused up front
    # rather than after issuing ten thousand requests.
    with pytest.raises(DownloadRangeError) as caught:
        download_ranges(
            size=MAX_DOWNLOAD_BYTES,
            range_size=RANGE_SIZE_MULTIPLE,
            fetch=fetcher(b""),
            sleep=no_sleep,
        )
    assert caught.value.status == "download_too_large"


# --------------------------------------------------------------------------- happy path


def test_every_planned_window_is_requested_once_and_in_order():
    calls: list[tuple[int, int]] = []
    data = download_ranges(
        size=len(PAYLOAD), range_size=WINDOW, fetch=fetcher(PAYLOAD, calls=calls), sleep=no_sleep
    )
    assert data == PAYLOAD
    assert calls == list(planned_ranges(len(PAYLOAD), WINDOW))


def test_contiguous_ranges_reassemble_the_file_exactly():
    data = download_ranges(
        size=len(PAYLOAD), range_size=WINDOW, fetch=fetcher(PAYLOAD), sleep=no_sleep
    )
    assert data == PAYLOAD
    assert len(data) == len(PAYLOAD)


def test_an_empty_file_downloads_to_zero_bytes_without_any_request():
    calls: list[tuple[int, int]] = []
    data = download_ranges(
        size=0,
        range_size=RANGE_SIZE_MULTIPLE,
        fetch=fetcher(b"", calls=calls),
        sleep=no_sleep,
    )
    assert data == b""
    assert calls == []


# --------------------------------------------------------------------------- integrity


def test_a_server_answering_200_instead_of_206_is_refused():
    """Without 206 the server ignored the range and would have re-sent the whole file."""
    def fetch(start, end):
        # A server that ignored the range re-sends the whole file with 200.
        return RangeOutcome(status=200, data=PAYLOAD, content_range=f"bytes {start}-{end}/{len(PAYLOAD)}")

    with pytest.raises(DownloadRangeError) as caught:
        download_ranges(size=len(PAYLOAD), range_size=WINDOW, fetch=fetch, sleep=no_sleep)
    assert caught.value.status == "range_not_content_partial"


def test_a_missing_content_range_header_is_refused():
    def fetch(start, end):
        return RangeOutcome(status=206, data=PAYLOAD[start : end + 1], content_range=None)

    with pytest.raises(DownloadRangeError) as caught:
        download_ranges(size=len(PAYLOAD), range_size=WINDOW, fetch=fetch, sleep=no_sleep)
    assert caught.value.status == "range_window_mismatch"


@pytest.mark.parametrize(
    "header",
    ["bytes abc-def/2048", "bytes 0-9", "items 0-9/2048", "bytes 0-9/x", "", "0-9/2048", None],
)
def test_an_unparseable_content_range_is_refused(header):
    assert parse_content_range(header) is None

    def fetch(start, end):
        return RangeOutcome(status=206, data=PAYLOAD[start : end + 1], content_range=header)

    with pytest.raises(DownloadRangeError) as caught:
        download_ranges(size=len(PAYLOAD), range_size=WINDOW, fetch=fetch, sleep=no_sleep)
    assert caught.value.status == "range_window_mismatch"


def test_a_server_returning_the_wrong_window_is_refused():
    """A short window would otherwise leave a hole in the middle of the file."""
    def fetch(start, end):
        # The server answers a different window than the one requested: a hole would appear
        # in the middle of the file if this were accepted.
        other = (start + WINDOW) % len(PAYLOAD)
        return RangeOutcome(
            status=206,
            data=PAYLOAD[other : other + (end - start + 1)],
            content_range=f"bytes {other}-{other + (end - start)}/{len(PAYLOAD)}",
        )

    with pytest.raises(DownloadRangeError) as caught:
        download_ranges(size=len(PAYLOAD), range_size=WINDOW, fetch=fetch, sleep=no_sleep)
    assert caught.value.status == "range_window_mismatch"


def test_a_range_that_returns_more_bytes_than_it_declares_is_refused():
    def fetch(start, end):
        return RangeOutcome(
            status=206,
            data=PAYLOAD[start : end + WINDOW],
            content_range=f"bytes {start}-{end}/{len(PAYLOAD)}",
        )

    with pytest.raises(DownloadRangeError) as caught:
        download_ranges(size=len(PAYLOAD), range_size=WINDOW, fetch=fetch, sleep=no_sleep)
    assert caught.value.status == "range_window_mismatch"


def test_an_empty_range_response_is_refused_rather_than_read_as_end_of_file():
    def fetch(start, end):
        return RangeOutcome(status=206, data=b"", content_range=f"bytes {start}-{end}/{len(PAYLOAD)}")

    with pytest.raises(DownloadRangeError) as caught:
        download_ranges(size=len(PAYLOAD), range_size=WINDOW, fetch=fetch, sleep=no_sleep)
    assert caught.value.status == "range_empty"


def test_a_total_that_disagrees_with_the_declared_size_is_refused():
    def fetch(start, end):
        return RangeOutcome(
            status=206, data=PAYLOAD[start : end + 1], content_range=f"bytes {start}-{end}/{len(PAYLOAD) + 1}"
        )

    with pytest.raises(DownloadRangeError) as caught:
        download_ranges(size=len(PAYLOAD), range_size=WINDOW, fetch=fetch, sleep=no_sleep)
    assert caught.value.status == "range_out_of_bounds"


def test_a_non_outcome_response_is_refused():
    with pytest.raises(DownloadRangeError) as caught:
        download_ranges(
            size=len(PAYLOAD), range_size=WINDOW, fetch=lambda s, e: b"raw bytes", sleep=no_sleep
        )
    assert caught.value.status == "range_failed"


# --------------------------------------------------------------------------- retries


def test_a_transport_failure_is_retried_and_the_download_still_completes():
    attempts: list[int] = []

    def fetch(start, end):
        attempts.append(start)
        if len(attempts) == 1:
            raise TimeoutError("connection reset")
        return RangeOutcome(
            status=206, data=PAYLOAD[start : end + 1], content_range=f"bytes {start}-{end}/{len(PAYLOAD)}"
        )

    slept: list[float] = []
    data = download_ranges(
        size=len(PAYLOAD), range_size=WINDOW, fetch=fetch, sleep=slept.append
    )
    assert data == PAYLOAD
    assert slept == [2.0]
    # Only the failing range is re-read; the successful ones were not re-fetched.
    assert attempts == [0, 0, WINDOW, 2 * WINDOW]


def test_a_range_that_always_fails_stops_at_the_attempt_bound():
    def fetch(start, end):
        raise TimeoutError("connection reset")

    slept: list[float] = []
    with pytest.raises(DownloadRangeError) as caught:
        download_ranges(size=len(PAYLOAD), range_size=WINDOW, fetch=fetch, sleep=slept.append)
    assert caught.value.status == "range_failed"
    assert caught.value.attempts == 5
    assert len(slept) == 4


def test_the_backoff_is_bounded_and_never_grows_without_limit():
    def fetch(start, end):
        raise TimeoutError("reset")

    slept: list[float] = []
    with pytest.raises(DownloadRangeError):
        download_ranges(
            size=len(PAYLOAD), range_size=WINDOW, fetch=fetch, sleep=slept.append, max_attempts=50
        )
    assert max(slept) <= 8.0


# --------------------------------------------------------------------------- sanitization


def test_a_ranged_failure_never_carries_raw_graph_text():
    def fetch(start, end):
        raise RuntimeError("secret-token=abc123 at https://graph.microsoft.com/v1.0/drives/x")

    with pytest.raises(DownloadRangeError) as caught:
        download_ranges(size=len(PAYLOAD), range_size=WINDOW, fetch=fetch, sleep=no_sleep)
    result = caught.value.to_result()
    assert result["status"] == "range_failed"
    assert "secret-token" not in str(result)
    assert "graph.microsoft.com" not in str(result)
    # The raw detail stays available for a log, never in the rendered result.
    assert isinstance(caught.value.__cause__, RuntimeError)


def test_a_ranged_failure_carries_the_action_a_caller_can_take():
    with pytest.raises(DownloadRangeError) as caught:
        validate_range_size(1)
    assert caught.value.to_result()["action"]


# --------------------------------------------------------------------------- strategy


class _Adapter(RequestAdapter):
    """An offline transport that answers every send from a responder instead of a network."""

    def __init__(self, responder):
        self.base_url = "https://graph.microsoft.com/v1.0"
        self._responder = responder
        self.requests: list = []

    def enable_backing_store(self, backing_store_factory=None):
        return None

    def get_serialization_writer_factory(self):
        from kiota_abstractions.base_serialization_writer_factory import (
            JsonSerializationWriterFactory,
        )

        return JsonSerializationWriterFactory()

    def _answer(self, request_information):
        self.requests.append(request_information)
        return self._responder(request_information)

    async def send_async(self, request_information, parsable_factory=None, error_map=None, **kwargs):
        return self._answer(request_information)

    async def send_collection_async(self, request_information, parsable_factory=None, error_map=None, **kwargs):
        return self._answer(request_information)

    async def send_collection_of_primitive_async(self, request_information, *args, **kwargs):
        return self._answer(request_information)

    async def send_primitive_async(self, request_information, *args, **kwargs):
        return self._answer(request_information)

    async def send_no_response_content_async(self, request_information, error_map=None, **kwargs):
        return self._answer(request_information)

    async def convert_to_native_async(self, request_information, *args, **kwargs):
        return self._answer(request_information)


def _header(headers, name: str):
    """Read one request header case-insensitively, unwrapping Kiota's multi-value container.

    Kiota normalizes header names to lowercase and stores values in a ``set``, so a single
    header arrives as ``{"bytes=0-9"}``. Any container this cannot read is "no header", which
    is what the responder treats as the unranged path.
    """
    if headers is None:
        return None
    values = headers.get(name)
    if values is None:
        return None
    if isinstance(values, (set, frozenset, list, tuple)):
        return next(iter(values)) if values else None
    return str(values)


class _RawResponse:
    """A minimal ``ApiResponse`` carrying bytes, a status and its ``Content-Range``."""

    def __init__(self, status: int, data: bytes, content_range: str | None):
        self.status_code = status
        self._data = data
        self.headers = {"Content-Range": content_range} if content_range else {}

    def read(self, size: int) -> bytes:
        chunk, self._data = self._data[:size], self._data[size:]
        return chunk


def drive_payload(size: int) -> bytes:
    body, remainder = divmod(size, 256)
    return bytes(range(256)) * body + bytes(remainder)


def ranged_adapter(size: int, *, capture: list):
    """Serve one metadata ``DriveItem`` plus one correctly-windowed range per ``Range`` request."""
    from msgraph.generated.models.drive_item import DriveItem
    from msgraph.generated.models.file import File as DriveFile

    payload = drive_payload(size)
    item = DriveItem(
        id="item-1", name="big.bin", size=size, file=DriveFile(mime_type="application/octet-stream")
    )

    def responder(request_information):
        headers = getattr(request_information, "headers", None)
        window = _header(headers, "range")
        path = urlsplit(str(request_information.url)).path
        if not path.endswith("/content"):
            capture.append(("metadata", None))
            return item
        if not window:
            # The small-file path: one whole-file GET with no Range header.
            capture.append(("whole", None))
            return _RawResponse(200, payload, None)
        start, end = (int(part) for part in window[6:].split("-"))
        capture.append((start, end))
        return _RawResponse(206, payload[start : end + 1], f"bytes {start}-{end}/{size}")

    return _Adapter(responder), payload


def _drive_client(adapter: RequestAdapter):
    from msgraph import GraphServiceClient

    return GraphServiceClient(request_adapter=adapter)


def _no_wait_execute(adapter: RequestAdapter):
    """``execute_request`` with the retry backoff removed, so offline tests do not wait."""
    import functools

    return functools.partial(execution.execute_request, adapter=adapter, sleep=lambda seconds: None)


def test_a_small_file_is_downloaded_in_one_request_with_no_range_header():
    import base64

    capture: list = []
    adapter, payload = ranged_adapter(2048, capture=capture)
    client = _drive_client(adapter)
    result = download_file_auto(
        client, drive_id="drive-1", path="small.bin"
    )
    assert result.size == 2048
    assert result.name == "big.bin"
    assert base64.b64decode(result.content_base64) == payload
    # metadata + one whole-file content GET, and no ranged read was needed
    assert [window for window in capture if window[0] != "metadata"] == [("whole", None)]


def test_a_large_file_is_downloaded_through_aligned_ranges():
    import base64

    size = 25 * MB
    capture: list = []
    adapter, payload = ranged_adapter(size, capture=capture)
    client = _drive_client(adapter)
    result = download_file_auto(
        client, drive_id="drive-1", path="big.bin"
    )
    assert result.size == size
    assert result.content_type == "application/octet-stream"
    assert base64.b64decode(result.content_base64) == payload
    assert [window for window in capture if window[0] not in {"metadata", "whole"}] == list(
        planned_ranges(size, DEFAULT_RANGE_SIZE)
    )


def test_an_explicit_range_size_is_honored():
    size = 12 * MB
    capture: list = []
    adapter, _ = ranged_adapter(size, capture=capture)
    client = _drive_client(adapter)
    download_file_auto(
        client,
        drive_id="drive-1",
        path="big.bin",
        range_size=RANGE_SIZE_MULTIPLE,
    )
    assert [window for window in capture if window[0] not in {"metadata", "whole"}] == list(
        planned_ranges(size, RANGE_SIZE_MULTIPLE)
    )


def test_the_ranged_path_sets_only_a_range_header():
    """Credentials travel in the client's transport; this code path adds no Authorization header."""
    seen: list = []
    # Above the 10 MiB simple bound, so the ranged path is the one under test.
    size = 25 * MB
    adapter, _ = ranged_adapter(size, capture=[])
    inner = adapter._responder

    def spy(request_information):
        headers = getattr(request_information, "headers", None)
        seen.append({str(key).lower() for key in (headers.keys() if headers else ())})
        return inner(request_information)

    adapter._responder = spy
    client = _drive_client(adapter)
    download_file_auto(client, drive_id="drive-1", path="big.bin")
    assert seen, "the ranged path issued no request"
    assert all("Authorization" not in names for names in seen)
    assert any("range" in names for names in seen)
