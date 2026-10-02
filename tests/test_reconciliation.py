"""Offline tests for post-write reconciliation (``microsoft365.reconciliation``).

The rule under test: a write is a claim, and this module either confirms it against the
server's own state or says plainly that it could not. Each adversarial case proves a broken
write is never reported as a good one.
"""
from __future__ import annotations

import pytest
from kiota_abstractions.request_adapter import RequestAdapter
from msgraph import GraphServiceClient
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.file import File

from microsoft365.reconciliation import (
    CONFIRMED,
    MISMATCHED,
    UNVERIFIED,
    Reconciliation,
    compare,
    reconcile_drive_item,
)

MB = 1024 * 1024
PAYLOAD_SIZE = 3 * MB


def item(*, size=PAYLOAD_SIZE, name="report.bin", identity="item-1", etag=None):
    model = DriveItem(
        id=identity, name=name, size=size, file=File(mime_type="application/octet-stream")
    )
    if etag:
        model.additional_data = {"@microsoft.graph.cTag": etag}
    return model


# ------------------------------------------------------------------ compare: confirmed


def test_a_matching_size_and_name_is_confirmed():
    verdict = compare(item(), expected_size=PAYLOAD_SIZE, expected_name="report.bin")
    assert verdict.status == CONFIRMED
    assert verdict.confirmed
    assert verdict.to_result()["expected"]["size"] == PAYLOAD_SIZE
    assert verdict.to_result()["observed"]["size"] == PAYLOAD_SIZE


def test_the_confirmation_carries_how_many_reads_it_took():
    verdict = compare(item(), expected_size=PAYLOAD_SIZE)
    assert verdict.to_result()["attempts"] == 1


# ------------------------------------------------------------------ compare: mismatched


def test_a_size_that_differs_is_mismatched():
    """A truncated transfer that still answered 2xx must not read as a good write."""
    verdict = compare(item(size=PAYLOAD_SIZE - 1), expected_size=PAYLOAD_SIZE)
    assert verdict.status == MISMATCHED
    assert verdict.reason == "size_differs_from_what_was_written"
    assert not verdict.confirmed


def test_a_larger_size_is_also_mismatched():
    verdict = compare(item(size=PAYLOAD_SIZE + 1), expected_size=PAYLOAD_SIZE)
    assert verdict.status == MISMATCHED


def test_a_renamed_item_is_mismatched():
    """A ``rename`` conflict policy that produced ``report (2).bin`` did not honour the request."""
    verdict = compare(item(name="report (2).bin"), expected_size=PAYLOAD_SIZE, expected_name="report.bin")
    assert verdict.status == MISMATCHED
    assert verdict.reason == "name_differs_from_what_was_written"


def test_reading_back_a_different_item_is_mismatched():
    verdict = compare(
        item(identity="item-2"), expected_size=PAYLOAD_SIZE, expected_id="item-1"
    )
    assert verdict.status == MISMATCHED
    assert verdict.reason == "the_item_written_is_not_the_item_read"


# ------------------------------------------------------------------ compare: unverified


def test_an_item_without_a_size_is_unverified_not_confirmed():
    """No size means this module cannot tell a good write from a bad one."""
    verdict = compare(item(size=None), expected_size=PAYLOAD_SIZE)
    assert verdict.status == UNVERIFIED
    assert verdict.reason == "item_declared_no_size"
    assert not verdict.confirmed


def test_an_item_without_a_name_is_unverified_not_confirmed():
    verdict = compare(item(name=""), expected_size=PAYLOAD_SIZE, expected_name="report.bin")
    assert verdict.status == UNVERIFIED


def test_nothing_declared_to_check_against_is_unverified():
    """Re-reading proves nothing when there is nothing to prove."""
    verdict = compare(item())
    assert verdict.status == UNVERIFIED
    assert verdict.reason == "nothing_was_declared_to_check_against"


def test_a_non_integer_expected_size_is_unverified():
    verdict = compare(item(), expected_size="three megabytes")
    assert verdict.status == UNVERIFIED
    assert verdict.reason == "expected_size_is_not_an_integer"


# ------------------------------------------------------------------ etag handling


def test_an_etag_is_carried_but_never_interpreted():
    verdict = compare(item(etag='"c1a2b3",1'), expected_size=PAYLOAD_SIZE)
    assert verdict.status == CONFIRMED
    assert verdict.to_result()["observed"]["etag"] == {"@microsoft.graph.cTag": '"c1a2b3",1'}


# ------------------------------------------------------------------ read-back seam


class _Adapter(RequestAdapter):
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
        return self._responder(len(self.requests))

    async def send_async(self, request_information, parsable_factory=None, error_map=None, **kwargs):
        return self._answer(request_information)

    async def send_collection_async(self, request_information, *args, **kwargs):
        return self._answer(request_information)

    async def send_collection_of_primitive_async(self, request_information, *args, **kwargs):
        return self._answer(request_information)

    async def send_primitive_async(self, request_information, *args, **kwargs):
        return self._answer(request_information)

    async def send_no_response_content_async(self, request_information, error_map=None, **kwargs):
        return self._answer(request_information)

    async def convert_to_native_async(self, request_information, *args, **kwargs):
        return self._answer(request_information)


def no_wait(seconds: float) -> None:
    return None


def client_for(responder) -> GraphServiceClient:
    return GraphServiceClient(request_adapter=_Adapter(responder))


def test_a_straight_read_back_confirms_the_write():
    client = client_for(lambda attempt: item())
    verdict = reconcile_drive_item(
        client,
        drive_id="drive-1",
        path="report.bin",
        expected_size=PAYLOAD_SIZE,
        expected_name="report.bin",
        sleep=no_wait,
    )
    assert verdict.status == CONFIRMED
    assert verdict.to_result()["attempts"] == 1


def test_an_item_that_is_not_visible_yet_is_retried_and_then_confirmed():
    def responder(attempt):
        return None if attempt == 1 else item()

    client = client_for(responder)
    verdict = reconcile_drive_item(
        client,
        drive_id="drive-1",
        path="report.bin",
        expected_size=PAYLOAD_SIZE,
        sleep=no_wait,
    )
    assert verdict.status == CONFIRMED
    assert verdict.to_result()["attempts"] == 2


def test_an_item_that_never_appears_is_unverified_at_the_attempt_bound():
    client = client_for(lambda attempt: None)
    verdict = reconcile_drive_item(
        client,
        drive_id="drive-1",
        path="report.bin",
        expected_size=PAYLOAD_SIZE,
        sleep=no_wait,
    )
    assert verdict.status == UNVERIFIED
    assert verdict.reason == "the_item_did_not_appear"
    assert verdict.to_result()["attempts"] == 3


def test_a_failed_read_back_is_unverified_and_never_a_failed_write():
    def responder(attempt):
        raise TimeoutError("connection reset after the write")

    client = client_for(responder)
    verdict = reconcile_drive_item(
        client,
        drive_id="drive-1",
        path="report.bin",
        expected_size=PAYLOAD_SIZE,
        sleep=no_wait,
    )
    assert verdict.status == UNVERIFIED
    assert verdict.reason == "the_item_could_not_be_read_back"


def test_a_mismatch_is_not_retried():
    """The server answered clearly and contradicted the claim: re-reading only delays the report."""
    client = client_for(lambda attempt: item(size=PAYLOAD_SIZE - 1))
    verdict = reconcile_drive_item(
        client,
        drive_id="drive-1",
        path="report.bin",
        expected_size=PAYLOAD_SIZE,
        sleep=no_wait,
    )
    assert verdict.status == MISMATCHED
    assert verdict.to_result()["attempts"] == 1


def test_the_backoff_between_attempts_is_bounded():
    delays: list[float] = []
    client = client_for(lambda attempt: None)
    reconcile_drive_item(
        client,
        drive_id="drive-1",
        path="report.bin",
        expected_size=PAYLOAD_SIZE,
        sleep=delays.append,
    )
    assert delays == [0.2, 0.5]
    assert max(delays) <= 0.5


def test_the_read_back_targets_the_same_item_the_write_addressed():
    adapter = _Adapter(lambda attempt: item())
    client = GraphServiceClient(request_adapter=adapter)
    reconcile_drive_item(
        client,
        drive_id="drive-1",
        path="reports/q3.bin",
        expected_size=PAYLOAD_SIZE,
        sleep=no_wait,
    )
    from urllib.parse import unquote, urlsplit

    path = urlsplit(str(adapter.requests[0].url)).path
    # The generated builder renders the documented ``<root>:/<path>:`` address form.
    assert unquote(path) == "/drives/drive-1/items/root:/reports/q3.bin:"


# ------------------------------------------------------------------ envelope


def test_the_envelope_never_carries_an_exception_or_a_url():
    """Only the verdict, the compared numbers and the attempt count are rendered."""
    verdict = Reconciliation(
        MISMATCHED,
        reason="size_differs_from_what_was_written",
        expected={"size": 10},
        observed={"size": 9},
        attempts=1,
    )
    rendered = verdict.to_result()
    assert set(rendered) == {"status", "reason", "expected", "observed", "attempts"}
    assert isinstance(rendered["status"], str)


def test_the_verdict_is_immutable():
    verdict = compare(item(), expected_size=PAYLOAD_SIZE)
    with pytest.raises(Exception):
        verdict.status = CONFIRMED