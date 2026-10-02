"""Upload sessions: large files, explicit conflicts, resumable fragments (PRO).

Every assertion here is offline. The session POST is exercised against a recording request
adapter through the generated SDK builder, the fragment PUTs against an injected seam, and the
production ``httpx`` seam is proven not to attach any tenant credential.
"""
from __future__ import annotations

import base64

import pytest
from msgraph import GraphServiceClient
from msgraph.generated.models.upload_session import UploadSession

from microsoft365 import upload_session
from microsoft365.errors import CATEGORIES, GraphError

#: The addressing literal the SDK accepts; the generated builder percent-encodes it whole.
ADDRESS = "root:/reports/q3.txt:"
#: ...and how it appears in a built URL, exactly as for the already-verified ``/content`` rows.
ENCODED_ADDRESS_PATH = "/drives/drive-1/items/root%3A%2Freports%2Fq3.txt%3A"
DRIVE = "drive-1"
SENTINEL = "upload-url-secret-sentinel"  # pragma: allowlist secret
SESSION_URL = "https://upload.contoso.invalid/session"
#: A pre-authenticated Graph session URL is itself a bearer-equivalent secret; it is used as a
#: sentinel here to prove it never escapes the module.
SESSION_URL_SENTINEL = f"{SESSION_URL}?tempauth={SENTINEL}"  # pragma: allowlist secret


class ChunkRecorder:
    """A recording fragment seam: no HTTP, deterministic outcomes.

    The first fragment answers ``202`` (``nextExpectedRanges``), the last answers ``201`` with a
    committed item, so the resumable path is what gets exercised by default.
    """

    def __init__(self, *, committed_on_last: bool = True, fail_times: int = 0):
        self.calls: list[dict] = []
        self.committed_on_last = committed_on_last
        self.fail_times = fail_times

    def __call__(self, upload_url, start, end, total, chunk, content_type):
        self.calls.append(
            {
                "upload_url": upload_url,
                "start": start,
                "end": end,
                "total": total,
                "bytes": chunk,
                "content_type": content_type,
            }
        )
        if self.fail_times > 0:
            self.fail_times -= 1
            raise RuntimeError("transport reset")
        if self.committed_on_last and end + 1 < total:
            return upload_session.ChunkOutcome(status=202, next_expected_ranges=[f"{end + 1}-"])
        item = upload_session._committed_drive_item(
            {"id": "item-1", "webUrl": "https://contoso.invalid/q3.txt", "size": total}
        )
        return upload_session.ChunkOutcome(status=201, item=item)


def _session_client(session) -> GraphServiceClient:
    from tests.test_handlers_files import HandlerGraphAdapter

    adapter = HandlerGraphAdapter(
        {("POST", f"{ENCODED_ADDRESS_PATH}/createUploadSession"): session}
    )
    return GraphServiceClient(request_adapter=adapter)


def _session_response(url: str = SESSION_URL) -> UploadSession:
    return UploadSession(
        upload_url=url,
        expiration_date_time=None,
        next_expected_ranges=["0-"],
    )


# ---------------------------------------------------------------- validation


def test_chunk_size_must_be_a_multiple_of_320_kib():
    assert upload_session.validate_chunk_size(320 * 1024) == 320 * 1024
    for bad in (0, 1024, 320 * 1024 - 1, upload_session.MAX_CHUNK_SIZE + 320 * 1024, True, "5"):
        with pytest.raises(upload_session.UploadSessionError) as caught:
            upload_session.validate_chunk_size(bad)
        assert caught.value.status == upload_session.INVALID_CHUNK_SIZE
        assert caught.value.to_result()["status"] == upload_session.INVALID_CHUNK_SIZE


def test_conflict_behavior_defaults_to_replace_and_accepts_only_graph_values():
    assert upload_session.validate_conflict_behavior(None) == "replace"
    for value in upload_session.CONFLICT_BEHAVIORS:
        assert upload_session.validate_conflict_behavior(value) == value
    for bad in ("overwrite", "REPLACE", True, 1, ""):
        with pytest.raises(upload_session.UploadSessionError) as caught:
            upload_session.validate_conflict_behavior(bad)
        assert caught.value.status == upload_session.INVALID_CONFLICT_BEHAVIOR


def test_upload_url_must_be_an_absolute_https_url_without_credentials():
    assert upload_session.validate_upload_url("https://upload.contoso.invalid/s").endswith("/s")
    for bad in (
        "http://upload.contoso.invalid/s",
        "https://user:pw@upload.contoso.invalid/s",
        "//upload.contoso.invalid/s",
        "upload.contoso.invalid",
        " https://upload.contoso.invalid/s",
        "https://upload.contoso.invalid/s\n",
        None,
        "",
    ):
        with pytest.raises(upload_session.UploadSessionError) as caught:
            upload_session.validate_upload_url(bad)
        assert caught.value.status == upload_session.INVALID_UPLOAD_URL


def test_session_size_above_the_graph_bound_is_refused_before_any_request():
    with pytest.raises(upload_session.UploadSessionError) as caught:
        upload_session.validate_session_size(upload_session.MAX_SESSION_BYTES + 1)
    assert caught.value.status == upload_session.UPLOAD_SESSION_TOO_LARGE
    assert caught.value.size_bytes == upload_session.MAX_SESSION_BYTES + 1


def test_file_name_is_taken_from_the_address_and_never_from_free_input():
    assert upload_session.file_name_from_address(ADDRESS) == "q3.txt"
    assert upload_session.file_name_from_address("root:/a/b/c.bin") == "c.bin"
    with pytest.raises(upload_session.UploadSessionError) as caught:
        upload_session.file_name_from_address("root:/")
    assert caught.value.status == upload_session.UPLOAD_SESSION_NO_NAME


# ---------------------------------------------------------------- session creation


def test_create_upload_session_posts_to_the_generated_builder_and_returns_the_url():
    client = _session_client(_session_response())
    url, ranges = upload_session.create_upload_session(
        client,
        drive_id=DRIVE,
        drive_item_id=ADDRESS,
        name="q3.txt",
        conflict_behavior="rename",
    )

    assert url == SESSION_URL
    assert ranges == ["0-"]
    request = client.request_adapter.requests[-1]
    assert request.http_method.value == "POST"
    assert request.url.endswith(f"{ENCODED_ADDRESS_PATH}/createUploadSession")


def test_create_upload_session_refuses_a_non_upload_session_response():
    client = _session_client({"id": "not-a-session"})

    with pytest.raises(upload_session.UploadSessionError) as caught:
        upload_session.create_upload_session(
            client, drive_id=DRIVE, drive_item_id=ADDRESS, name="q3.txt", conflict_behavior="replace"
        )

    assert caught.value.status == upload_session.UPLOAD_SESSION_FAILED
    assert caught.value.category in CATEGORIES


def test_a_failed_transfer_never_leaks_the_pre_authenticated_upload_url():
    def put(upload_url, start, end, total, chunk, content_type):
        raise RuntimeError("boom")  # pragma: allowlist secret

    with pytest.raises(upload_session.UploadSessionError) as caught:
        upload_session.upload_chunks(
            upload_url=SESSION_URL_SENTINEL,
            payload=b"j" * 8,
            start_offset=0,
            chunk_size=320 * 1024,
            content_type="text/plain",
            put=put,
            sleep=lambda _seconds: None,
            max_attempts=2,
        )

    assert SENTINEL not in str(caught.value)
    assert SENTINEL not in str(caught.value.to_result())


# ---------------------------------------------------------------- chunked transfer


def test_upload_chunks_sends_every_fragment_with_its_content_range():
    payload = b"a" * (320 * 1024 * 2 + 5)
    recorder = ChunkRecorder()

    result = upload_session.upload_chunks(
        upload_url="https://upload.contoso.invalid/s",
        payload=payload,
        start_offset=0,
        chunk_size=320 * 1024,
        content_type="text/plain",
        put=recorder,
        sleep=lambda _seconds: None,
    )

    assert result["status"] == "uploaded"
    assert result["size_bytes"] == len(payload)
    assert [call["start"] for call in recorder.calls] == [0, 320 * 1024, 640 * 1024]
    assert [call["end"] for call in recorder.calls] == [320 * 1024 - 1, 640 * 1024 - 1, len(payload) - 1]
    assert all(call["total"] == len(payload) for call in recorder.calls)
    assert all(call["content_type"] == "text/plain" for call in recorder.calls)
    assert b"".join(call["bytes"] for call in recorder.calls) == payload


def test_upload_chunks_resumes_from_the_range_graph_still_expects():
    payload = b"b" * (320 * 1024 * 3)
    recorder = ChunkRecorder()

    result = upload_session.upload_chunks(
        upload_url="https://upload.contoso.invalid/s",
        payload=payload,
        start_offset=320 * 1024,
        chunk_size=320 * 1024,
        content_type="text/plain",
        put=recorder,
        sleep=lambda _seconds: None,
    )

    assert result["status"] == "uploaded"
    assert [call["start"] for call in recorder.calls] == [320 * 1024, 640 * 1024]


def test_upload_chunks_follows_a_mid_transfer_range_correction():
    payload = b"c" * (320 * 1024 * 4)
    calls: list[tuple[int, int]] = []

    def put(upload_url, start, end, total, chunk, content_type):
        calls.append((start, end))
        if start == 0:
            # Graph accepts the fragment but reports a different remaining range than the
            # sequential one, which is what makes the resume logic necessary.
            return upload_session.ChunkOutcome(status=202, next_expected_ranges=[f"{end + 1}-"])
        return upload_session.ChunkOutcome(status=202, next_expected_ranges=[f"{end + 1}-"])

    with pytest.raises(upload_session.UploadSessionError) as caught:
        upload_session.upload_chunks(
            upload_url="https://upload.contoso.invalid/s",
            payload=payload,
            start_offset=0,
            chunk_size=320 * 1024,
            content_type="text/plain",
            put=put,
            sleep=lambda _seconds: None,
        )

    assert caught.value.status == upload_session.UPLOAD_SESSION_INCOMPLETE
    assert calls == [(320 * 1024 * i, 320 * 1024 * (i + 1) - 1) for i in range(4)]


def test_upload_chunks_honors_a_single_jump_back_requested_by_graph():
    payload = b"k" * (320 * 1024 * 3)
    calls: list[int] = []

    def put(upload_url, start, end, total, chunk, content_type):
        calls.append(start)
        if start == 0:
            # Graph reports it never received fragment 0: the resume must honor that once and
            # then continue forward.
            return upload_session.ChunkOutcome(status=202, next_expected_ranges=["0-"])
        return upload_session.ChunkOutcome(status=202, next_expected_ranges=[f"{end + 1}-"])

    with pytest.raises(upload_session.UploadSessionError) as caught:
        upload_session.upload_chunks(
            upload_url="https://upload.contoso.invalid/s",
            payload=payload,
            start_offset=0,
            chunk_size=320 * 1024,
            content_type="text/plain",
            put=put,
            sleep=lambda _seconds: None,
        )

    assert caught.value.status == upload_session.UPLOAD_SESSION_INCOMPLETE
    assert caught.value.expected_range == "bytes 0-327679/983040"
    assert calls == [0] * upload_session.MAX_STALLED_FRAGMENTS


def test_upload_chunks_resumes_forward_after_a_single_replayed_fragment():
    payload = b"m" * (320 * 1024 * 3)
    calls: list[int] = []
    replays = {"left": 1}

    def put(upload_url, start, end, total, chunk, content_type):
        calls.append(start)
        if replays["left"] > 0:
            replays["left"] -= 1
            return upload_session.ChunkOutcome(status=202, next_expected_ranges=["0-"])
        return upload_session.ChunkOutcome(status=202, next_expected_ranges=[f"{end + 1}-"])

    with pytest.raises(upload_session.UploadSessionError) as caught:
        upload_session.upload_chunks(
            upload_url="https://upload.contoso.invalid/s",
            payload=payload,
            start_offset=0,
            chunk_size=320 * 1024,
            content_type="text/plain",
            put=put,
            sleep=lambda _seconds: None,
        )

    # One replayed fragment is tolerated and the transfer moves forward afterwards.
    assert calls[:4] == [0, 0, 320 * 1024, 640 * 1024]
    assert caught.value.status == upload_session.UPLOAD_SESSION_INCOMPLETE


def test_upload_chunks_retries_a_transport_failure_with_bounded_backoff():
    payload = b"d" * 16
    delays: list[float] = []
    attempts = {"count": 0}

    def put(upload_url, start, end, total, chunk, content_type):
        attempts["count"] += 1
        if attempts["count"] <= 2:
            raise RuntimeError("connection reset")
        return upload_session.ChunkOutcome(status=201, item=None)

    result = upload_session.upload_chunks(
        upload_url="https://upload.contoso.invalid/s",
        payload=payload,
        start_offset=0,
        chunk_size=320 * 1024,
        content_type="text/plain",
        put=put,
        sleep=delays.append,
    )

    assert result["status"] == "uploaded"
    assert attempts["count"] == 3
    assert delays == [2.0, 4.0]


def test_upload_chunks_gives_up_after_the_attempt_bound_and_names_the_range():
    def put(upload_url, start, end, total, chunk, content_type):
        raise RuntimeError("connection reset")

    with pytest.raises(upload_session.UploadSessionError) as caught:
        upload_session.upload_chunks(
            upload_url="https://upload.contoso.invalid/s",
            payload=b"e" * 8,
            start_offset=0,
            chunk_size=320 * 1024,
            content_type="text/plain",
            put=put,
            sleep=lambda _seconds: None,
        )

    assert caught.value.status == upload_session.UPLOAD_SESSION_FAILED
    assert caught.value.attempts == upload_session.MAX_CHUNK_ATTEMPTS
    assert caught.value.expected_range == "bytes 0-7/8"
    assert "connection reset" not in str(caught.value)


def test_upload_chunks_never_replays_a_permanent_http_failure():
    calls = []

    def put(upload_url, start, end, total, chunk, content_type):
        calls.append(start)
        return upload_session.ChunkOutcome(status=409)

    with pytest.raises(upload_session.UploadSessionError) as caught:
        upload_session.upload_chunks(
            upload_url="https://upload.contoso.invalid/s",
            payload=b"f" * 8,
            start_offset=0,
            chunk_size=320 * 1024,
            content_type="text/plain",
            put=put,
            sleep=lambda _seconds: None,
        )

    assert calls == [0]
    assert caught.value.status == upload_session.UPLOAD_SESSION_FAILED


def test_upload_chunks_reports_an_incomplete_transfer_instead_of_claiming_success():
    def put(upload_url, start, end, total, chunk, content_type):
        # Always asks for offset 0 again: an unfinished transfer must never look uploaded.
        return upload_session.ChunkOutcome(status=202, next_expected_ranges=["0-"])

    with pytest.raises(upload_session.UploadSessionError) as caught:
        upload_session.upload_chunks(
            upload_url="https://upload.contoso.invalid/s",
            payload=b"g" * 8,
            start_offset=0,
            chunk_size=320 * 1024,
            content_type="text/plain",
            put=put,
            sleep=lambda _seconds: None,
            max_attempts=2,
        )

    assert caught.value.status in {
        upload_session.UPLOAD_SESSION_INCOMPLETE,
        upload_session.UPLOAD_SESSION_FAILED,
    }


# ---------------------------------------------------------------- orchestration


def test_upload_file_session_creates_sends_and_reports_the_committed_item():
    payload = b"h" * (320 * 1024 + 7)
    client = _session_client(_session_response())
    recorder = ChunkRecorder()

    result = upload_session.upload_file_session(
        client,
        drive_id=DRIVE,
        content_base64=base64.b64encode(payload).decode(),
        content_type="text/plain",
        path="reports/q3.txt",
        conflict_behavior="fail",
        chunk_size=320 * 1024,
        put=recorder,
        sleep=lambda _seconds: None,
    )

    assert result["status"] == "uploaded"
    assert result["size_bytes"] == len(payload)
    assert result["conflict_behavior"] == "fail"
    assert result["item_id"] == "item-1"
    assert result["web_url"] == "https://contoso.invalid/q3.txt"
    assert b"".join(call["bytes"] for call in recorder.calls) == payload
    assert client.request_adapter.requests[-1].url.endswith(
        f"{ENCODED_ADDRESS_PATH}/createUploadSession"
    )


def test_upload_file_session_never_returns_the_pre_authenticated_upload_url():
    payload = b"i" * 8
    client = _session_client(_session_response(SESSION_URL_SENTINEL))
    recorder = ChunkRecorder()

    result = upload_session.upload_file_session(
        client,
        drive_id=DRIVE,
        content_base64=base64.b64encode(payload).decode(),
        content_type="application/octet-stream",
        path="reports/q3.txt",
        put=recorder,
        sleep=lambda _seconds: None,
    )

    assert SENTINEL not in repr(result)
    assert all(call["upload_url"] == SESSION_URL_SENTINEL for call in recorder.calls)


def test_upload_file_session_validates_before_it_creates_a_session():
    class FailIfTouched:
        @property
        def request_adapter(self):
            raise AssertionError("no client work may happen before validation")

    for arguments in (
        {"drive_id": "", "content_base64": "", "content_type": "text/plain"},
        {"drive_id": DRIVE, "content_base64": "!!!", "content_type": "text/plain"},
        {"drive_id": DRIVE, "content_base64": "", "content_type": "no-slash"},
        {"drive_id": DRIVE, "content_base64": "", "content_type": "text/plain", "chunk_size": 1000},
        {
            "drive_id": DRIVE,
            "content_base64": "",
            "content_type": "text/plain",
            "conflict_behavior": "clobber",
        },
    ):
        with pytest.raises(GraphError):
            upload_session.upload_file_session(
                FailIfTouched(),
                put=ChunkRecorder(),
                sleep=lambda _seconds: None,
                **arguments,
            )


# ---------------------------------------------------------------- production seam


def test_production_chunk_seam_sends_only_range_and_content_type_headers(monkeypatch):
    captured: dict = {}

    class Response:
        status_code = 202

        @staticmethod
        def json():
            return {"nextExpectedRanges": ["10-"]}

    class Client:
        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

        def put(self, url, content=None, headers=None):
            captured["url"] = url
            captured["content"] = content
            captured["headers"] = headers
            return Response()

    import httpx

    monkeypatch.setattr(httpx, "Client", lambda **kwargs: Client())

    outcome = upload_session.put_chunk_default(
        "https://upload.contoso.invalid/s", 0, 9, 40, b"z" * 10, "text/plain"
    )

    assert outcome.status == 202
    assert outcome.next_expected_ranges == ["10-"]
    assert captured["headers"] == {
        "Content-Type": "text/plain",
        "Content-Range": "bytes 0-9/40",
    }
    assert "Authorization" not in captured["headers"]
    assert captured["content"] == b"z" * 10


def test_validate_upload_url_accepts_a_query_bearing_graph_session_url():
    assert (
        upload_session.validate_upload_url(SESSION_URL_SENTINEL) == SESSION_URL_SENTINEL
    )


def test_production_chunk_seam_reports_a_committed_item(monkeypatch):
    class Response:
        status_code = 201

        @staticmethod
        def json():
            return {"id": "item-9", "webUrl": "https://contoso.invalid/x", "size": 40}

    class Client:
        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

        def put(self, url, content=None, headers=None):
            return Response()

    import httpx

    monkeypatch.setattr(httpx, "Client", lambda **kwargs: Client())

    outcome = upload_session.put_chunk_default(
        "https://upload.contoso.invalid/s", 0, 39, 40, b"z" * 40, "text/plain"
    )

    assert outcome.status == 201
    assert outcome.item.id == "item-9"
    assert outcome.item.web_url == "https://contoso.invalid/x"


def test_production_chunk_seam_tolerates_a_non_json_response(monkeypatch):
    class Response:
        status_code = 202

        @staticmethod
        def json():
            raise ValueError("not json")

    class Client:
        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

        def put(self, url, content=None, headers=None):
            return Response()

    import httpx

    monkeypatch.setattr(httpx, "Client", lambda **kwargs: Client())

    outcome = upload_session.put_chunk_default(
        "https://upload.contoso.invalid/s", 0, 0, 1, b"z", "text/plain"
    )

    assert outcome.status == 202
    assert outcome.next_expected_ranges == []


def test_sleep_default_is_bounded_and_delegates_to_time():
    import time

    recorded = []
    original = time.sleep
    time.sleep = lambda seconds: recorded.append(seconds)
    try:
        upload_session.sleep_default(2.5)
    finally:
        time.sleep = original

    assert recorded == [2.5]