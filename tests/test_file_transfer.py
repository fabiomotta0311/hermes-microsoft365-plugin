"""WP4 -- bounded, exact drive ``/content`` transfer.

Every request in this module is built by the **generated** msgraph-sdk builders and bound to
:class:`StrictTransportAdapter` (``tests/test_sdk_contract.py``), whose six ``send_*`` entry
points raise, so nothing here can reach a network. The metadata read and the byte read are
observed through :class:`RecordingTransferSeam`, a callable with the exact signature of
:func:`microsoft365.execution.execute_request`; production always uses that seam (pinned by
test), and every recorded artifact is the SDK's own output for a real builder.
"""
from __future__ import annotations

import asyncio
import base64
import inspect
from typing import Any
from urllib.parse import urlsplit

import pytest
from msgraph import GraphServiceClient

from microsoft365 import execution, files, results
from microsoft365.errors import CATEGORIES, MESSAGES, GraphError
from tests.test_sdk_contract import StrictTransportAdapter

MI_BYTE = 1024 * 1024
LIMIT = 10 * MI_BYTE
DRIVE = "drive"
MIME = "application/pdf"
PUZZLE_PATH = "relatórios/2026 #1 (final)?.txt"
PUZZLE_ENCODED_PATH = (
    "/drives/drive/items/root%3A%2Frelat%C3%B3rios%2F2026%20%231%20%28final%29%3F.txt%3A"
)
ITEM_ADDRESS = "root:/a/b.txt:"
ITEM_ENCODED_PATH = "/drives/drive/items/root%3A%2Fa%2Fb.txt%3A"


@pytest.fixture
def graph_client():
    return GraphServiceClient(request_adapter=StrictTransportAdapter())


class HostileStreamingResponse:
    """A response whose eager read would defeat the transfer bound."""

    def __init__(self, chunks, *, content_length=None):
        self.headers = {} if content_length is None else {"Content-Length": str(content_length)}
        self._chunks = iter(chunks)
        self.read_called = False

    def read(self, *args, **kwargs):
        self.read_called = True
        raise AssertionError("bounded download must not eagerly read the response")

    def iter_bytes(self, *args, **kwargs):
        yield from self._chunks


class RecordingTransferSeam:
    """Records real request artifacts and returns canned authenticated responses.

    Deliberately a function-like seam (never an object fake, a ``SimpleNamespace`` or a
    permissive ``__getattr__``) whose keyword signature is exactly
    :func:`microsoft365.execution.execute_request`'s.
    """

    def __init__(self, *, metadata, content):
        self.metadata = metadata
        self.content = content
        self.calls: list[dict[str, Any]] = []

    def __call__(self, builder, *, method, configuration=None, body=None, adapter=None):
        request = execution.request_information_sender(
            builder, method=method, configuration=configuration, body=body
        )
        path = urlsplit(request.url).path
        self.calls.append(
            {
                "method": method,
                "url": request.url,
                "path": path,
                "content": request.content,
                "content_type": request.headers.get("content-type"),
                "adapter": adapter,
            }
        )
        return self.content if path.endswith("/content") else self.metadata

    @property
    def paths(self) -> list[str]:
        return [call["path"] for call in self.calls]

    @property
    def sent_bodies(self) -> list[Any]:
        return [call["content"] for call in self.calls]


def drive_item(*, size, mime_type=MIME, name="report.pdf", item_id="item"):
    """A real generated ``DriveItem`` metadata response (with a real ``File`` facet)."""
    from msgraph.generated.models.drive_item import DriveItem
    from msgraph.generated.models.file import File

    return DriveItem(
        id=item_id,
        name=name,
        size=size,
        file=None if mime_type is None else File(mime_type=mime_type),
    )


def folder_item():
    from msgraph.generated.models.drive_item import DriveItem
    from msgraph.generated.models.folder import Folder

    return DriveItem(id="folder", name="reports", size=0, folder=Folder(child_count=0))


def assert_transfer_error(excinfo, *, category, status):
    error = excinfo.value
    assert isinstance(error, files.TransferError)
    assert error.category in CATEGORIES
    assert error.category == category
    assert error.status == status
    payload = error.to_result()
    assert payload["error"] == category
    assert payload["status"] == status
    assert payload["message"] == MESSAGES[category]
    return payload


def seam_for(*, size, content, mime_type=MIME, name="report.pdf"):
    return RecordingTransferSeam(
        metadata=drive_item(size=size, mime_type=mime_type, name=name), content=content
    )


# ---------------------------------------------------------------------------------------
# The 10 MiB simple-transfer contract and the encoded-length bound
# ---------------------------------------------------------------------------------------


def test_simple_transfer_contract_is_exactly_ten_mib():
    assert files.SIMPLE_TRANSFER_LIMIT_BYTES == 10 * 1024 * 1024 == LIMIT


def test_encoded_length_bound_is_the_exact_base64_length_of_the_limit():
    limit = LIMIT
    assert files.encoded_length_bound(limit) == len(base64.b64encode(b"x" * limit)) == 13981016
    # The bound is a *bound*, not a proof: one extra byte encodes to the same length, which is
    # why the decoded-size check exists as a second gate.
    assert len(base64.b64encode(b"x" * (limit + 1))) == files.encoded_length_bound(limit)


@pytest.mark.parametrize("value", [LIMIT + 1, 0, -1, True, False, "10485760", 1.5, None])
def test_the_product_bound_can_only_be_lowered(graph_client, value):
    with pytest.raises(files.TransferError) as excinfo:
        files.upload_file(
            graph_client,
            drive_id=DRIVE,
            path="a/b.txt",
            content_base64=base64.b64encode(b"x").decode("ascii"),
            content_type="text/plain",
            limit_bytes=value,
        )
    assert_transfer_error(excinfo, category="configuration_error", status="invalid_transfer_limit")


def test_a_lowered_limit_is_honoured(graph_client):
    seam = seam_for(size=5, content=b"12345", name="a.txt")
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(
            graph_client, drive_id=DRIVE, path="a/b.txt", execute=seam, limit_bytes=4
        )

    result = assert_transfer_error(
        excinfo,
        category="operation_not_implemented",
        status=files.DOWNLOAD_RANGE_NOT_IMPLEMENTED,
    )
    assert result["limit_bytes"] == 4
    assert seam.paths == [ITEM_ENCODED_PATH]


# ---------------------------------------------------------------------------------------
# Path addressing: validated before anything is built
# ---------------------------------------------------------------------------------------


def test_path_addressing_uses_the_documented_relative_form():
    assert files.drive_item_address(path="a/b.txt") == ITEM_ADDRESS
    assert files.drive_item_address(path="a/b.txt", identifier="01ABCDEF!x") == "01ABCDEF!x:/a/b.txt:"


def test_an_opaque_item_id_is_passed_through_unchanged():
    assert files.drive_item_address(drive_item_id="01ABCDEF!x") == "01ABCDEF!x"


def test_a_prebuilt_address_is_validated_as_an_address():
    assert files.drive_item_address(drive_item_id=ITEM_ADDRESS) == ITEM_ADDRESS
    with pytest.raises(files.TransferError) as excinfo:
        files.drive_item_address(drive_item_id="root:/../etc/passwd:")
    assert_transfer_error(excinfo, category="validation_error", status="invalid_drive_path")


@pytest.mark.parametrize("drive_item_id", [None, "", "   ", 7, "root:", "root://a.txt", "a:b"])
def test_addressing_rejects_a_missing_or_malformed_item_id(drive_item_id):
    with pytest.raises(files.TransferError) as excinfo:
        files.drive_item_address(drive_item_id=drive_item_id)
    assert excinfo.value.category == "validation_error"


def test_addressing_requires_exactly_one_of_item_id_or_path():
    with pytest.raises(files.TransferError) as excinfo:
        files.drive_item_address()
    assert_transfer_error(excinfo, category="validation_error", status="addressing_required")

    with pytest.raises(files.TransferError) as excinfo:
        files.drive_item_address(drive_item_id="item", path="a/b.txt")
    assert_transfer_error(excinfo, category="validation_error", status="addressing_conflict")


@pytest.mark.parametrize(
    ("path", "label"),
    [
        ("", "empty"),
        ("   ", "blank"),
        (" a.txt", "leading space"),
        ("a.txt ", "trailing space"),
        ("/a.txt", "absolute"),
        ("/drives/d/items/x/content", "absolute graph route"),
        ("a/", "trailing separator"),
        ("a//b.txt", "empty segment"),
        ("a/./b.txt", "dot segment"),
        ("a/../b.txt", "traversal"),
        ("../a.txt", "leading traversal"),
        ("a/..", "trailing traversal"),
        (".", "dot"),
        ("..", "dot dot"),
        (ITEM_ADDRESS, "already addressed"),
        ("root:/a.txt", "drive-absolute prefix"),
        ("C:/a/b.txt", "drive letter"),
        ("c:/windows/system32", "drive letter lowercase"),
        ("https://graph.microsoft.com/v1.0/drives/d/root:/a.txt", "url"),
        ("a\\b.txt", "backslash"),
        ("\\a.txt", "leading backslash"),
        ("a\tb.txt", "tab"),
        ("a\nb.txt", "newline"),
        ("a\x00b.txt", "nul"),
        ("a\x7fb.txt", "delete"),
        ("a\u202eb.txt", "bidi override"),
    ],
)
def test_unsafe_paths_are_rejected_before_any_request_is_built(path, label):
    with pytest.raises(files.TransferError) as excinfo:
        files.validate_relative_drive_path(path)
    assert_transfer_error(excinfo, category="validation_error", status="invalid_drive_path")


@pytest.mark.parametrize("path", ["a b.txt", "a#b.txt", "a?b.txt", "#", "?"])
def test_reserved_and_unicode_characters_are_accepted_not_mangled(path):
    # The path is handed to the generated builder verbatim; encoding is the SDK's job.
    assert files.validate_relative_drive_path(path) == path
    assert files.drive_item_address(path=path) == f"root:/{path}:"


def test_the_builder_encodes_unicode_and_reserved_characters_not_string_mangling(graph_client):
    seam = seam_for(size=3, content=b"abc", name="a.txt")
    result = files.download_file(
        graph_client, drive_id=DRIVE, path=PUZZLE_PATH, execute=seam
    )

    assert seam.paths == [PUZZLE_ENCODED_PATH, f"{PUZZLE_ENCODED_PATH}/content"]
    # Nothing is escaped by hand: the raw characters never appear in the rendered URL.
    for raw in (" ", "#", "?", "ó"):
        assert raw not in seam.calls[0]["url"]
    assert result.name == "a.txt"


@pytest.mark.parametrize(
    ("value", "name"),
    [
        ("a" * (files.MAX_PATH_LENGTH + 1), "one long segment"),
        ("/".join(["a"] * 2100), "many segments"),
    ],
)
def test_path_length_is_bounded(value, name):
    with pytest.raises(files.TransferError):
        files.drive_item_address(path=value)


@pytest.mark.parametrize("identifier", ["a" * (files.MAX_IDENTIFIER_LENGTH + 1), "ro ot", "ro/ot", "ro:ot", "ro\not"])
def test_path_identifier_is_validated(identifier):
    with pytest.raises(files.TransferError) as excinfo:
        files.drive_item_address(path="a.txt", identifier=identifier)
    assert excinfo.value.category == "validation_error"


@pytest.mark.parametrize("drive_id", ["", "   ", None, 7, "a/b", "a:b", "a\nb", "a" * 401])
def test_drive_id_is_validated(graph_client, drive_id):
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=drive_id, path="a.txt")
    assert excinfo.value.category == "validation_error"


# ---------------------------------------------------------------------------------------
# Download: exact bytes, /content target, content type from authenticated metadata
# ---------------------------------------------------------------------------------------


def test_download_targets_content_and_returns_exact_bytes(graph_client):
    payload = bytes(range(256)) * 4
    seam = seam_for(size=len(payload), content=payload, name="b.txt")

    result = files.download_file(graph_client, drive_id=DRIVE, path="a/b.txt", execute=seam)

    assert seam.calls[0]["method"] == "GET"
    assert seam.calls[0]["path"] == ITEM_ENCODED_PATH
    assert seam.calls[1]["method"] == "GET"
    assert seam.calls[1]["path"] == f"{ITEM_ENCODED_PATH}/content"
    assert seam.calls[1]["url"].endswith("/content")
    assert all(call["adapter"] is graph_client.request_adapter for call in seam.calls)
    assert isinstance(result, results.BinaryResult)
    assert result.content_type == MIME
    assert result.size == len(payload)
    assert result.name == "b.txt"
    assert base64.b64decode(result.content_base64, validate=True) == payload


def test_download_metadata_read_is_not_the_content_request(graph_client):
    payload = b"abc"
    seam = seam_for(size=len(payload), content=payload)
    files.download_file(graph_client, drive_id=DRIVE, drive_item_id=ITEM_ADDRESS, execute=seam)

    assert seam.paths[0] == ITEM_ENCODED_PATH
    assert not seam.paths[0].endswith("/content")
    assert seam.calls[0]["content"] is None


def test_download_content_type_comes_from_authenticated_metadata_not_the_caller(graph_client):
    seam = seam_for(size=4, content=b"<h1>", name="report.html", mime_type="application/pdf")

    result = files.download_file(graph_client, drive_id=DRIVE, path="report.html", execute=seam)

    assert result.content_type == "application/pdf"
    assert result.name == "report.html"
    with pytest.raises(TypeError):
        files.download_file(
            graph_client, drive_id=DRIVE, path="report.html", content_type="text/html"
        )


def test_download_refuses_metadata_that_is_not_a_drive_item(graph_client):
    from msgraph.generated.models.message import Message

    seam = RecordingTransferSeam(metadata=Message(id="m"), content=b"abc")
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="a/b.txt", execute=seam)
    assert_transfer_error(
        excinfo, category="configuration_error", status="metadata_not_a_drive_item"
    )


@pytest.mark.parametrize("size", [None, "12", 1.5, True, -1])
def test_download_refuses_a_metadata_without_a_usable_size(graph_client, size):
    seam = RecordingTransferSeam(metadata=drive_item(size=size), content=b"abc")
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="a/b.txt", execute=seam)
    assert_transfer_error(excinfo, category="configuration_error", status="metadata_without_size")


@pytest.mark.parametrize("mime_type", [None, "", "   "])
def test_download_refuses_an_item_that_is_not_a_file(graph_client, mime_type):
    seam = RecordingTransferSeam(
        metadata=drive_item(size=4, mime_type=mime_type), content=b"abcd"
    )
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="a/b.txt", execute=seam)
    assert_transfer_error(excinfo, category="validation_error", status="not_a_file")


def test_download_refuses_a_folder(graph_client):
    seam = RecordingTransferSeam(metadata=folder_item(), content=b"")
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="reports", execute=seam)
    assert_transfer_error(excinfo, category="validation_error", status="not_a_file")


def test_download_refuses_a_declared_size_over_the_bound_without_reading_the_content(graph_client):
    seam = RecordingTransferSeam(metadata=drive_item(size=LIMIT + 1), content=b"x")
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="a/b.txt", execute=seam)

    payload = assert_transfer_error(
        excinfo,
        category="operation_not_implemented",
        status=files.DOWNLOAD_RANGE_NOT_IMPLEMENTED,
    )
    assert payload["limit_bytes"] == LIMIT
    assert payload["size_bytes"] == LIMIT + 1
    assert payload["action"]
    # Bounded reading: the byte request was never issued.
    assert seam.paths == [ITEM_ENCODED_PATH]


def test_download_accepts_a_declared_size_exactly_at_the_bound(graph_client):
    payload = b"x" * 8
    seam = seam_for(size=8, content=payload)
    result = files.download_file(
        graph_client, drive_id=DRIVE, path="a.txt", execute=seam, limit_bytes=8
    )
    assert result.size == 8
    assert len(seam.calls) == 2


def test_download_refuses_a_response_that_is_not_bytes(graph_client):
    from msgraph.generated.models.message import Message

    seam = RecordingTransferSeam(metadata=drive_item(size=3), content=Message(id="m"))
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="a/b.txt", execute=seam)
    assert_transfer_error(
        excinfo, category="configuration_error", status="content_response_not_bytes"
    )
    assert len(seam.calls) == 2


def test_download_refuses_a_payload_over_the_bound_even_when_metadata_understated_it(graph_client):
    seam = seam_for(size=4, content=b"12345678")
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(
            graph_client, drive_id=DRIVE, path="a.txt", execute=seam, limit_bytes=4
        )
    payload = assert_transfer_error(
        excinfo,
        category="operation_not_implemented",
        status=files.DOWNLOAD_RANGE_NOT_IMPLEMENTED,
    )
    assert payload["limit_bytes"] == 4
    assert payload["size_bytes"] == 8


def test_download_interrupts_a_hostile_stream_before_materializing_it(graph_client):
    response = HostileStreamingResponse([b"1", b"2", b"3", b"4", b"5", b"6"])
    seam = seam_for(size=4, content=response)
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="a.txt", execute=seam, limit_bytes=4)
    payload = assert_transfer_error(excinfo, category="operation_not_implemented", status=files.DOWNLOAD_RANGE_NOT_IMPLEMENTED)
    assert payload["size_bytes"] == 5
    assert response.read_called is False
    assert list(response._chunks) == [b"6"]


def test_download_rejects_a_content_length_over_the_bound_before_streaming(graph_client):
    response = HostileStreamingResponse([b"never-read"], content_length=5)
    seam = seam_for(size=4, content=response)
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="a.txt", execute=seam, limit_bytes=4)
    payload = assert_transfer_error(excinfo, category="operation_not_implemented", status=files.DOWNLOAD_RANGE_NOT_IMPLEMENTED)
    assert payload["size_bytes"] == 5
    assert response.read_called is False
    assert list(response._chunks) == [b"never-read"]


def test_download_reports_a_declared_size_mismatch(graph_client):
    seam = seam_for(size=5, content=b"four")
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="a/b.txt", execute=seam)
    payload = assert_transfer_error(
        excinfo, category="precondition_failed", status=files.DECLARED_SIZE_MISMATCH
    )
    assert payload["size_bytes"] == 4


def test_download_refuses_a_metadata_target_that_is_the_content_request(graph_client, monkeypatch):
    from microsoft365 import sdk_contract

    def content_target(client, case, **kwargs):
        del case, kwargs
        return sdk_contract.build_content_request_information(
            client, "download_files", drive_id=DRIVE, drive_item_id=ITEM_ADDRESS
        )

    monkeypatch.setattr(files, "build_item_request_information", content_target)
    seam = seam_for(size=3, content=b"abc")
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="a/b.txt", execute=seam)
    assert_transfer_error(
        excinfo, category="configuration_error", status="metadata_target_mismatch"
    )
    assert seam.calls == []


def test_download_refuses_a_content_target_that_is_not_the_content_request(
    graph_client, monkeypatch
):
    from microsoft365 import sdk_contract

    def wrong_target(client, case, **kwargs):
        del case, kwargs
        return sdk_contract.build_item_request_information(
            client, "drive_item_read", drive_id=DRIVE, drive_item_id="item"
        )

    monkeypatch.setattr(files, "build_content_request_information", wrong_target)
    seam = seam_for(size=3, content=b"abc")
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="a/b.txt", execute=seam)
    assert_transfer_error(excinfo, category="configuration_error", status="content_target_mismatch")
    assert seam.calls == []


def test_download_requires_a_request_adapter_bound_client(graph_client):
    graph_client.request_adapter = None
    with pytest.raises(files.TransferError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="a/b.txt")
    assert_transfer_error(
        excinfo, category="configuration_error", status="invalid_request_adapter"
    )


# ---------------------------------------------------------------------------------------
# Upload: encoded-length gate, decoded-size check, exact bytes and Content-Type
# ---------------------------------------------------------------------------------------


def b64(payload: bytes) -> str:
    return base64.b64encode(payload).decode("ascii")


def test_upload_puts_exact_bytes_with_content_type_to_the_content_url(graph_client):
    payload = bytes(range(256)) * 16
    seam = RecordingTransferSeam(metadata=None, content=None)

    result = files.upload_file(
        graph_client,
        drive_id=DRIVE,
        path="a/b.txt",
        content_base64=b64(payload),
        content_type="text/csv",
        execute=seam,
    )

    assert seam.paths == [f"{ITEM_ENCODED_PATH}/content"]
    call = seam.calls[0]
    assert call["method"] == "PUT"
    assert call["url"].endswith("/content")
    assert call["content"] == payload
    # The explicit media type is the one sent: the SDK's own default for a byte body is
    # application/octet-stream, so this value can only come from the request configuration.
    assert call["content_type"] == {"text/csv"}
    assert call["content_type"] != {"application/octet-stream"}
    assert call["adapter"] is graph_client.request_adapter
    assert result == {
        "status": "uploaded",
        "size_bytes": len(payload),
        "content_type": "text/csv",
    }


def test_upload_rejects_an_oversized_encoded_payload_before_decoding(graph_client, monkeypatch):
    oversized = b64(b"x" * (LIMIT + 4096))
    assert len(oversized) > files.encoded_length_bound(LIMIT)

    def boom(*args, **kwargs):
        raise AssertionError("base64 decode must not run for an oversized payload")

    monkeypatch.setattr(files.base64, "b64decode", boom)
    seam = RecordingTransferSeam(metadata=None, content=None)
    with pytest.raises(files.TransferError) as excinfo:
        files.upload_file(
            graph_client,
            drive_id=DRIVE,
            path="a/b.txt",
            content_base64=oversized,
            content_type="text/plain",
            execute=seam,
        )

    failure = assert_transfer_error(
        excinfo,
        category="operation_not_implemented",
        status=files.UPLOAD_SESSION_NOT_IMPLEMENTED,
    )
    assert failure["limit_bytes"] == LIMIT
    assert failure["encoded_length"] == len(oversized)
    assert failure["action"]
    assert oversized[:40] not in str(failure)
    assert seam.calls == []


def test_upload_over_the_limit_but_within_the_encoded_bound_is_caught_after_decoding(graph_client):
    payload = b"x" * (LIMIT + 1)
    encoded = b64(payload)
    # The encoded gate cannot see this one: the bound is not a proof of the decoded size.
    assert len(encoded) <= files.encoded_length_bound(LIMIT)
    assert len(base64.b64decode(encoded, validate=True)) == LIMIT + 1

    seam = RecordingTransferSeam(metadata=None, content=None)
    with pytest.raises(files.TransferError) as excinfo:
        files.upload_file(
            graph_client,
            drive_id=DRIVE,
            path="a/b.txt",
            content_base64=encoded,
            content_type="text/plain",
            execute=seam,
        )

    payload = assert_transfer_error(
        excinfo,
        category="operation_not_implemented",
        status=files.UPLOAD_SESSION_NOT_IMPLEMENTED,
    )
    assert payload["size_bytes"] == LIMIT + 1
    assert payload["limit_bytes"] == LIMIT
    assert seam.calls == []


def test_upload_accepts_a_payload_exactly_at_the_bound(graph_client):
    payload = b"x" * LIMIT
    seam = RecordingTransferSeam(metadata=None, content=None)
    result = files.upload_file(
        graph_client,
        drive_id=DRIVE,
        path="a/b.txt",
        content_base64=b64(payload),
        content_type="application/octet-stream",
        execute=seam,
    )
    assert result["size_bytes"] == LIMIT
    assert seam.sent_bodies == [payload]


def test_upload_decoded_size_check_fires_within_the_encoded_bound(graph_client):
    # 4-byte limit -> encoded bound of 8 characters, which decodes to 6 bytes.
    assert files.encoded_length_bound(4) == 8
    seam = RecordingTransferSeam(metadata=None, content=None)
    with pytest.raises(files.TransferError) as excinfo:
        files.upload_file(
            graph_client,
            drive_id=DRIVE,
            path="a/b.txt",
            content_base64="AAAAAAAA",
            content_type="text/plain",
            execute=seam,
            limit_bytes=4,
        )
    payload = assert_transfer_error(
        excinfo,
        category="operation_not_implemented",
        status=files.UPLOAD_SESSION_NOT_IMPLEMENTED,
    )
    assert payload["size_bytes"] == 6
    assert payload["limit_bytes"] == 4


@pytest.mark.parametrize(
    "content_base64",
    ["not base64!!", "AAA", "AAAAA", "AA=A", "====", "YWJj$", "YWJj\n", "YWJjé", "\ufeffYWJj", "  ", 7, None, b"YWJj"],
)
def test_upload_refuses_a_payload_that_is_not_base64_text(graph_client, content_base64):
    seam = RecordingTransferSeam(metadata=None, content=None)
    with pytest.raises(files.TransferError) as excinfo:
        files.upload_file(
            graph_client,
            drive_id=DRIVE,
            path="a/b.txt",
            content_base64=content_base64,
            content_type="text/plain",
            execute=seam,
        )
    assert_transfer_error(
        excinfo, category="validation_error", status="invalid_content_base64"
    )
    assert seam.calls == []


@pytest.mark.parametrize(
    "content_type",
    [
        "",
        "   ",
        "text",
        "text/",
        "/plain",
        " text/plain",
        "text/plain ",
        "text/plain\r\nX-Injected: 1",
        "text/plain\n",
        "text plain",
        7,
        None,
        "a" * (files.MAX_CONTENT_TYPE_LENGTH + 1),
    ],
)
def test_upload_requires_a_plain_media_type_and_rejects_header_injection(
    graph_client, content_type
):
    seam = RecordingTransferSeam(metadata=None, content=None)
    with pytest.raises(files.TransferError) as excinfo:
        files.upload_file(
            graph_client,
            drive_id=DRIVE,
            path="a/b.txt",
            content_base64=b64(b"payload"),
            content_type=content_type,
            execute=seam,
        )
    assert_transfer_error(
        excinfo, category="validation_error", status="invalid_content_type"
    )
    assert seam.calls == []


def test_upload_refuses_a_content_target_that_is_not_the_content_request(graph_client, monkeypatch):
    from microsoft365 import sdk_contract

    def wrong_target(client, case, **kwargs):
        del case, kwargs
        return sdk_contract.build_item_request_information(
            client, "drive_item_read", drive_id=DRIVE, drive_item_id="item"
        )

    monkeypatch.setattr(files, "build_content_request_information", wrong_target)
    seam = RecordingTransferSeam(metadata=None, content=None)
    with pytest.raises(files.TransferError) as excinfo:
        files.upload_file(
            graph_client,
            drive_id=DRIVE,
            path="a/b.txt",
            content_base64=b64(b"payload"),
            content_type="text/plain",
            execute=seam,
        )
    assert_transfer_error(excinfo, category="configuration_error", status="content_target_mismatch")
    assert seam.calls == []


# ---------------------------------------------------------------------------------------
# Round trips: empty, several KiB, and exactly at the configured bound
# ---------------------------------------------------------------------------------------


def test_empty_file_round_trips_both_ways(graph_client):
    download = seam_for(size=0, content=b"")
    result = files.download_file(graph_client, drive_id=DRIVE, path="empty.txt", execute=download)

    assert result.size == 0
    assert result.content_base64 == ""
    assert results.normalize_result(result)["content_base64"] == ""

    upload = RecordingTransferSeam(metadata=None, content=None)
    ack = files.upload_file(
        graph_client,
        drive_id=DRIVE,
        path="empty.txt",
        content_base64=result.content_base64,
        content_type="text/plain",
        execute=upload,
    )
    assert ack["size_bytes"] == 0
    assert upload.sent_bodies == [b""]
    assert upload.calls[0]["method"] == "PUT"
    assert upload.calls[0]["content"] == b""
    assert upload.calls[0]["path"] == "/drives/drive/items/root%3A%2Fempty.txt%3A/content"


def test_exact_byte_round_trip_at_several_kib(graph_client):
    payload = bytes((index * 7) % 256 for index in range(5 * 1024))

    download = seam_for(size=len(payload), content=payload, name="q1.bin")
    result = files.download_file(graph_client, drive_id=DRIVE, path="a/q1.bin", execute=download)

    upload = RecordingTransferSeam(metadata=None, content=None)
    files.upload_file(
        graph_client,
        drive_id=DRIVE,
        path="a/q1.bin",
        content_base64=result.content_base64,
        content_type=result.content_type,
        execute=upload,
    )

    assert upload.sent_bodies == [payload]
    assert base64.b64decode(result.content_base64, validate=True) == payload
    assert download.paths[1] == upload.paths[0]


def test_exact_byte_round_trip_at_the_configured_limit(graph_client):
    payload = bytes((index * 13) % 256 for index in range(LIMIT))

    download = seam_for(size=len(payload), content=payload, name="limit.bin")
    result = files.download_file(
        graph_client, drive_id=DRIVE, path="a/limit.bin", execute=download
    )

    assert result.size == LIMIT
    assert len(result.content_base64) == files.encoded_length_bound(LIMIT)
    assert base64.b64decode(result.content_base64, validate=True) == payload
    normalized = results.normalize_result(result)
    assert normalized["content_base64"] == result.content_base64
    assert "TRUNCATED" not in normalized["content_base64"]

    upload = RecordingTransferSeam(metadata=None, content=None)
    files.upload_file(
        graph_client,
        drive_id=DRIVE,
        path="a/limit.bin",
        content_base64=result.content_base64,
        content_type=result.content_type,
        execute=upload,
    )
    assert upload.sent_bodies == [payload]


# ---------------------------------------------------------------------------------------
# results.py: the binary carrier
# ---------------------------------------------------------------------------------------


def test_binary_result_helper_preserves_base64_outside_string_truncation():
    payload = b"y" * 8192
    result = results.binary_result(payload, content_type=MIME, name="b.txt")

    assert result.size == 8192
    encoded = base64.b64encode(payload).decode("ascii")
    assert result.content_base64 == encoded
    normalized = results.normalize_result(result, max_string=4000)
    assert normalized["content_base64"] == encoded
    assert normalized["content_type"] == MIME
    assert normalized["name"] == "b.txt"


@pytest.mark.parametrize("payload", ["not bytes", 5, [1, 2, 3], 3.5, None])
def test_binary_result_helper_requires_bytes_and_a_content_type(payload):
    # An integer is the decisive case: ``bytes(5)`` would silently produce five NUL bytes.
    with pytest.raises(TypeError):
        results.binary_result(payload, content_type=MIME)
    with pytest.raises(ValueError):
        results.binary_result(b"payload", content_type="   ")
    with pytest.raises(ValueError):
        results.binary_result(b"payload", content_type=None)


# ---------------------------------------------------------------------------------------
# No network, single seam
# ---------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "method",
    [
        "send_async",
        "send_collection_async",
        "send_collection_of_primitive_async",
        "send_primitive_async",
        "send_no_response_content_async",
        "convert_to_native_async",
    ],
)
def test_every_transport_entry_point_raises_in_this_suite(method):
    adapter = StrictTransportAdapter()
    with pytest.raises(AssertionError, match="must not run"):
        asyncio.run(getattr(adapter, method)())


def test_transfer_production_default_is_the_execution_seam():
    assert (
        inspect.signature(files.download_file).parameters["execute"].default
        is execution.execute_request
    )
    assert (
        inspect.signature(files.upload_file).parameters["execute"].default
        is execution.execute_request
    )


def test_production_download_never_reaches_the_network(graph_client):
    with pytest.raises(GraphError) as excinfo:
        files.download_file(graph_client, drive_id=DRIVE, path="a/b.txt")
    assert isinstance(excinfo.value.__cause__, AssertionError)
    assert "network transport must not run" in str(excinfo.value.__cause__)
    assert excinfo.value.category in CATEGORIES


def test_production_upload_never_reaches_the_network(graph_client):
    with pytest.raises(GraphError) as excinfo:
        files.upload_file(
            graph_client,
            drive_id=DRIVE,
            path="a/b.txt",
            content_base64=b64(b"payload"),
            content_type="text/plain",
        )
    assert isinstance(excinfo.value.__cause__, AssertionError)
    assert "network transport must not run" in str(excinfo.value.__cause__)
    assert excinfo.value.category in CATEGORIES


def test_every_transfer_status_is_a_static_sanitized_payload():
    assert files.UPLOAD_SESSION_NOT_IMPLEMENTED == "upload_session_not_implemented"
    assert files.DOWNLOAD_RANGE_NOT_IMPLEMENTED == "download_range_not_implemented"
    assert files.DECLARED_SIZE_MISMATCH == "declared_size_mismatch"
    error = files.TransferError(
        "operation_not_implemented",
        status=files.UPLOAD_SESSION_NOT_IMPLEMENTED,
        limit_bytes=LIMIT,
        size_bytes=LIMIT + 1,
    )
    payload = error.to_result()
    assert payload["message"] == MESSAGES["operation_not_implemented"]
    assert set(payload) == {"error", "message", "retryable", "status", "action", "limit_bytes", "size_bytes"}
