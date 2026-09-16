"""WP-ERR — the canonical sanitized Microsoft Graph error taxonomy.

Every test here uses **real** exception classes from the installed stack (Kiota
``APIError``, ``msgraph`` ``ODataError``/``MainError``, azure-core/azure-identity errors,
``kiota_http`` errors, ``httpx`` transport errors). Nothing is mocked and no request is
ever sent: the taxonomy is a pure conversion layer, so it can be exercised offline.

The leakage tests plant ``SENTINEL`` inside real exception attributes (message, error
body, nested ``additional_data``, mixed-case response headers) and assert the literal
never reaches a returned error, an error string or a retry decision.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from microsoft365 import errors as taxonomy

#: A synthetic probe value with no credential shape (see the plan's detect-secrets notes).
SENTINEL = "SENTINEL-REDACTION-PROBE-8F3A2B"

EXPECTED_CATEGORIES = {
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


def _kiota_api_error(message=SENTINEL, status=429, headers=None):
    from kiota_abstractions.api_error import APIError

    return APIError(message, status, headers if headers is not None else {})


def _odata_error(*, code, status, message=SENTINEL, headers=None):
    from msgraph.generated.models.o_data_errors.main_error import MainError
    from msgraph.generated.models.o_data_errors.o_data_error import ODataError

    error = ODataError(
        message=message,
        response_status_code=status,
        response_headers=headers if headers is not None else {},
    )
    error.error = MainError(code=code, message=message)
    return error


class _AzureResponse:
    """A concrete azure-core response double: only the documented protocol members.

    azure-core's ``HttpResponseError`` takes its status and headers from the response
    object, never from a ``status_code`` keyword (verified in the installed source), so the
    status-code path of the taxonomy is only reachable with a real response-shaped object.
    """

    def __init__(self, status_code, headers=None):
        self.status_code = status_code
        self.headers = headers or {}
        self.reason = "test-reason"

    def text(self):
        return "not json on purpose"

    def body(self):  # pragma: no cover - azure-core only calls text() in this path
        return self.text().encode()

    def read(self):  # pragma: no cover - azure-core only calls text() in this path
        return self.text().encode()


def _azure_status_error(status, headers=None):
    from azure.core.exceptions import HttpResponseError

    return HttpResponseError(
        message=f"raw azure text {SENTINEL}",
        response=_AzureResponse(status, headers),
    )


# ---------------------------------------------------------------------------------------
# Taxonomy shape
# ---------------------------------------------------------------------------------------


def test_taxonomy_owns_exactly_the_planned_categories():
    assert set(taxonomy.CATEGORIES) == EXPECTED_CATEGORIES


def test_every_category_has_a_static_message_template():
    assert set(taxonomy.MESSAGES) == set(taxonomy.CATEGORIES)
    for category, template in taxonomy.MESSAGES.items():
        assert isinstance(template, str) and template
        # Static templates: no interpolation means no raw text can be spliced in.
        assert "{" not in template and "}" not in template, category
        assert "\n" not in template, category


def test_unknown_category_is_rejected_instead_of_silently_accepted():
    with pytest.raises(ValueError):
        taxonomy.GraphError("not_a_category", "message")


def test_graph_error_keeps_a_safe_string_form_and_never_the_cause_text():
    from kiota_abstractions.api_error import APIError

    cause = APIError(f"raw graph body {SENTINEL}", 500, {})
    converted = taxonomy.to_graph_error(cause)

    assert isinstance(converted, taxonomy.GraphError)
    assert converted.category == "service_error"
    assert converted.__cause__ is cause
    assert SENTINEL not in str(converted)
    assert SENTINEL not in str(converted.to_result())
    assert converted.to_result()["error"] == "service_error"


def test_conversion_is_idempotent_and_preserves_the_original_error():
    original = taxonomy.GraphError("throttled", taxonomy.MESSAGES["throttled"], retryable=True)

    assert taxonomy.to_graph_error(original) is original
    assert taxonomy.classify(original) == "throttled"


def test_non_exception_input_is_rejected():
    with pytest.raises(TypeError):
        taxonomy.to_graph_error("not an exception")


# ---------------------------------------------------------------------------------------
# Mapping boundaries — driven by real exception attributes, never by parsing a message
# ---------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "status, category",
    [
        (400, "validation_error"),
        (413, "validation_error"),
        (415, "validation_error"),
        (422, "validation_error"),
        (401, "authentication_required"),
        (403, "permission_denied"),
        (404, "not_found"),
        (410, "not_found"),
        (409, "conflict"),
        (412, "precondition_failed"),
        (428, "precondition_failed"),
        (429, "throttled"),
        (500, "service_error"),
        (503, "service_error"),
        (501, "service_error"),
    ],
)
def test_kiota_status_code_maps_to_exactly_one_category(status, category):
    assert taxonomy.classify(_kiota_api_error(status=status)) == category


@pytest.mark.parametrize(
    "code, status, category",
    [
        ("ErrorAccessDenied", 403, "permission_denied"),
        ("Authorization_RequestDenied", 403, "permission_denied"),
        ("ConsentRequired", 403, "consent_required"),
        ("AdminConsentRequired", 403, "consent_required"),
        ("interaction_required", 401, "consent_required"),
        ("InvalidAuthenticationToken", 401, "authentication_required"),
        ("invalid_grant", 401, "token_refresh_failed"),
        ("token_expired", 401, "token_refresh_failed"),
        ("ErrorItemNotFound", 404, "not_found"),
        ("nameAlreadyExists", 409, "conflict"),
        ("preconditionFailed", 412, "precondition_failed"),
        ("TooManyRequests", 429, "throttled"),
        ("ActivityLimitReached", 429, "throttled"),
        ("NotSupported", 501, "operation_not_implemented"),
        ("InvalidRequest", 400, "validation_error"),
    ],
)
def test_odata_error_code_decides_before_the_status_code(code, status, category):
    assert taxonomy.classify(_odata_error(code=code, status=status)) == category


def test_odata_code_is_used_only_as_a_signal_and_never_re_emitted():
    error = taxonomy.to_graph_error(_odata_error(code="TooManyRequests", status=429))
    payload = json.dumps(error.to_result())

    assert error.category == "throttled"
    assert "TooManyRequests" not in payload
    assert SENTINEL not in payload


@pytest.mark.parametrize(
    "status, category",
    [(400, "validation_error"), (404, "not_found"), (412, "precondition_failed"), (429, "throttled")],
)
def test_azure_core_status_code_is_also_an_attribute_signal(status, category):
    assert taxonomy.classify(_azure_status_error(status)) == category


def test_azure_response_headers_are_read_without_echoing_them():
    error = taxonomy.to_graph_error(
        _azure_status_error(
            429,
            headers={
                "Retry-After": "5",
                "Authorization": f"Bearer {SENTINEL}",
                "x-ms-request-id": "3F2504E0-4F89-41D3-9A0C-0305E82C3301",
            },
        )
    )
    payload = error.to_result()

    assert error.category == "throttled"
    assert error.retry_after_seconds == 5.0
    assert payload["correlation_id"] == "3f2504e0-4f89-41d3-9a0c-0305e82c3301"
    assert SENTINEL not in json.dumps(payload)
    assert "Bearer" not in json.dumps(payload)


def test_azure_core_resource_subclasses_map_without_a_status_code():
    from azure.core.exceptions import (
        ResourceExistsError,
        ResourceModifiedError,
        ResourceNotFoundError,
        ResourceNotModifiedError,
    )

    assert taxonomy.classify(ResourceNotFoundError(message=SENTINEL)) == "not_found"
    assert taxonomy.classify(ResourceExistsError(message=SENTINEL)) == "conflict"
    assert taxonomy.classify(ResourceNotModifiedError(message=SENTINEL)) == "precondition_failed"
    assert taxonomy.classify(ResourceModifiedError(message=SENTINEL)) == "precondition_failed"


def test_azure_credential_families_map_to_authentication_categories():
    from azure.core.exceptions import ClientAuthenticationError
    from azure.identity import AuthenticationRequiredError, CredentialUnavailableError

    assert taxonomy.classify(ClientAuthenticationError(message=SENTINEL)) == "authentication_required"
    assert (
        taxonomy.classify(AuthenticationRequiredError(["User.Read"], message=SENTINEL))
        == "authentication_required"
    )
    assert taxonomy.classify(CredentialUnavailableError(message=SENTINEL)) == "authentication_required"


def test_authentication_classes_win_over_a_misleading_status_code():
    from azure.identity import CredentialUnavailableError

    error = CredentialUnavailableError(message=SENTINEL)
    error.status_code = 400

    assert taxonomy.classify(error) == "authentication_required"


@pytest.mark.parametrize(
    "status, category, retryable",
    [
        (429, "throttled", True),
        (408, "service_error", True),
        (500, "service_error", True),
        (502, "service_error", True),
        (503, "service_error", True),
        (504, "service_error", True),
        (501, "service_error", False),
        (505, "service_error", False),
        (404, "not_found", False),
        (401, "authentication_required", False),
        (403, "permission_denied", False),
    ],
)
def test_retryable_flag_follows_the_status_set(status, category, retryable):
    error = taxonomy.to_graph_error(_kiota_api_error(status=status))

    assert error.category == category
    assert error.retryable is retryable


@pytest.mark.parametrize(
    "module, name",
    [("kiota_http._exceptions", "RequestError"), ("kiota_http._exceptions", "ResponseError")],
)
def test_kiota_transport_errors_are_transport_failures(module, name):
    import importlib

    error_class = getattr(importlib.import_module(module), name)

    assert taxonomy.classify(error_class(SENTINEL)) == "transport_error"


def test_azure_transport_errors_are_transport_failures():
    from azure.core.exceptions import (
        IncompleteReadError,
        ServiceRequestError,
        ServiceResponseError,
        ServiceResponseTimeoutError,
    )

    for error_class in (ServiceRequestError, ServiceResponseError, ServiceResponseTimeoutError, IncompleteReadError):
        assert taxonomy.classify(error_class(SENTINEL)) == "transport_error"


def test_httpx_transport_errors_are_transport_failures_without_parsing_text():
    httpx = pytest.importorskip("httpx")

    for error_class in (httpx.ConnectError, httpx.ReadTimeout, httpx.RemoteProtocolError):
        assert taxonomy.classify(error_class(SENTINEL)) == "transport_error"


def test_kiota_and_azure_decoding_failures_are_service_failures():
    import kiota_http._exceptions as kiota_errors
    from azure.core.exceptions import DecodeError

    assert taxonomy.classify(kiota_errors.BackingStoreError(SENTINEL)) == "internal_error"
    assert taxonomy.classify(kiota_errors.DeserializationError(SENTINEL)) == "service_error"
    assert taxonomy.classify(DecodeError(message=SENTINEL)) == "service_error"


def test_builtin_programming_errors_map_to_stable_categories():
    assert taxonomy.classify(ValueError(SENTINEL)) == "validation_error"
    assert taxonomy.classify(TypeError(SENTINEL)) == "configuration_error"
    assert taxonomy.classify(KeyError(SENTINEL)) == "internal_error"
    assert taxonomy.classify(AttributeError(SENTINEL)) == "internal_error"
    assert taxonomy.classify(RuntimeError(SENTINEL)) == "internal_error"
    assert taxonomy.classify(Exception(SENTINEL)) == "internal_error"


def test_classification_default_is_caller_supplied_and_never_leaks():
    assert taxonomy.classify(AssertionError(SENTINEL), default="transport_error") == "transport_error"
    assert taxonomy.classify(AssertionError(SENTINEL)) == "internal_error"
    with pytest.raises(ValueError):
        taxonomy.classify(AssertionError(SENTINEL), default="not_a_category")


# ---------------------------------------------------------------------------------------
# Sentinel-secret leakage
# ---------------------------------------------------------------------------------------


def test_raw_kiota_text_and_headers_never_reach_a_returned_error():
    headers = {
        "Retry-After": "7",
        "AUTHORIZATION": f"Bearer {SENTINEL}",
        "Set-Cookie": f"session={SENTINEL}",
        "X-Custom-Debug": SENTINEL,
    }
    error = taxonomy.to_graph_error(_kiota_api_error(headers=headers))
    payload = error.to_result()

    assert error.category == "throttled"
    assert error.retry_after_seconds == 7.0
    assert payload["retry_after_seconds"] == 7.0
    for rendered in (str(error), repr(payload), json.dumps(payload)):
        assert SENTINEL not in rendered
        assert "Bearer" not in rendered
        assert "session=" not in rendered


def test_odata_body_and_nested_additional_data_never_reach_a_returned_error():
    error = _odata_error(code="TooManyRequests", status=429, headers={"Retry-After": "1"})
    error.additional_data = {
        "client_secret": SENTINEL,  # pragma: allowlist secret
        "nested": {
            "authorization": f"Bearer {SENTINEL}",
            "deeper": {"access_token": SENTINEL, "refreshToken": SENTINEL},
        },
    }

    converted = taxonomy.to_graph_error(error)
    payload = converted.to_result()

    assert converted.category == "throttled"
    assert SENTINEL not in json.dumps(payload)
    assert SENTINEL not in str(converted)
    assert "client_secret" not in json.dumps(payload)


def test_azure_credential_error_text_never_reaches_a_returned_error_or_a_retry_decision():
    from azure.identity import CredentialUnavailableError

    cause = CredentialUnavailableError(message=f"client_secret={SENTINEL}")  # pragma: allowlist secret
    error = taxonomy.to_graph_error(cause)

    assert error.category == "authentication_required"
    assert SENTINEL not in str(error)
    assert SENTINEL not in json.dumps(error.to_result())
    decision = taxonomy.retry_decision(error, method="GET", attempt=0, max_attempts=3)
    assert decision.retry is False
    assert SENTINEL not in repr(decision)


def test_mixed_case_headers_are_read_case_insensitively(tmp_path):
    from kiota_abstractions.headers_collection import HeadersCollection

    collection = HeadersCollection()
    collection.add("rEtRy-AfTeR", "3")

    assert taxonomy.retry_after_seconds({"RETRY-AFTER": "3"}) == 3.0
    assert taxonomy.retry_after_seconds({"retry-after": "3"}) == 3.0
    assert taxonomy.retry_after_seconds(collection) == 3.0
    assert taxonomy.retry_after_seconds({"Retry-After": [SENTINEL]}) is None
    assert taxonomy.retry_after_seconds({SENTINEL: "3"}) is None
    assert taxonomy.retry_after_seconds(None) is None
    assert taxonomy.retry_after_seconds(tmp_path) is None


def test_retry_after_accepts_only_safely_parsed_values():
    parse = taxonomy.parse_retry_after

    assert parse("12") == 12.0
    assert parse("0") == 0.0
    assert parse("1.5") == 1.5
    assert parse(12) == 12.0
    assert parse("  4  ") == 4.0
    assert parse("-1") is None
    assert parse(-1) is None
    assert parse("abc") is None
    assert parse("NaN") is None
    assert parse(float("nan")) is None
    assert parse("inf") is None
    assert parse(float("inf")) is None
    assert parse(True) is None
    assert parse(None) is None
    assert parse(b"7") is None
    assert parse(SENTINEL) is None
    assert parse({}) is None
    # Bounded: an absurd header can never pin the caller for hours.
    assert parse("999999") == taxonomy.MAX_RETRY_AFTER_SECONDS
    assert parse(999999) == taxonomy.MAX_RETRY_AFTER_SECONDS


def test_retry_after_http_date_is_bounded_and_never_negative():
    from email.utils import format_datetime
    from datetime import datetime, timedelta, timezone

    now = datetime.now(timezone.utc)
    future = format_datetime(now + timedelta(hours=3))
    past = format_datetime(now - timedelta(hours=3))

    assert taxonomy.parse_retry_after(future) == taxonomy.MAX_RETRY_AFTER_SECONDS
    assert taxonomy.parse_retry_after(past) == 0.0
    assert taxonomy.parse_retry_after("Wed, 99 Zzz 2026 99:99:99 GMT") is None


def test_correlation_ids_are_emitted_only_when_they_are_opaque_identifiers():
    guid = "0F1E2D3C-4B5A-6978-8192-A3B4C5D6E7F8"
    other = "deadbeefdeadbeefdeadbeefdeadbeef"

    assert taxonomy.sanitize_correlation_id(guid) == guid.lower()
    assert taxonomy.sanitize_correlation_id(other) == other
    assert taxonomy.sanitize_correlation_id(SENTINEL) is None
    assert taxonomy.sanitize_correlation_id(f"Bearer {SENTINEL}") is None
    assert taxonomy.sanitize_correlation_id("aaaaaaaa") is None
    assert taxonomy.sanitize_correlation_id("a" * 200) is None
    assert taxonomy.sanitize_correlation_id(b"abc") is None
    assert taxonomy.sanitize_correlation_id(None) is None

    error = taxonomy.to_graph_error(
        _kiota_api_error(status=500, headers={"client-request-id": guid, "Authorization": f"Bearer {SENTINEL}"})
    )
    payload = error.to_result()

    assert payload["correlation_id"] == guid.lower()
    assert SENTINEL not in json.dumps(payload)

    hostile = taxonomy.to_graph_error(_kiota_api_error(status=500, headers={"client-request-id": SENTINEL}))
    assert "correlation_id" not in hostile.to_result()


def test_oversized_and_binary_attribute_values_are_bounded_and_never_copied():
    from kiota_abstractions.api_error import APIError

    huge = taxonomy.to_graph_error(APIError("x" * 500_000, 500, {"Retry-After": "7"}))
    payload = huge.to_result()

    assert huge.category == "service_error"
    assert len(json.dumps(payload)) < 500
    assert "x" * 100 not in json.dumps(payload)

    binary = taxonomy.to_graph_error(APIError(SENTINEL.encode(), 500, {b"retry-after": b"7"}))
    assert binary.category == "service_error"
    assert binary.retry_after_seconds is None
    assert SENTINEL not in str(binary)
    assert SENTINEL not in json.dumps(binary.to_result())


def test_results_never_render_a_raw_exception_as_text():
    from microsoft365.results import normalize_result

    payload = normalize_result(_kiota_api_error(status=404, headers={"Authorization": f"Bearer {SENTINEL}"}))

    assert payload["error"] == "not_found"
    assert SENTINEL not in json.dumps(payload)
    assert "raw" not in json.dumps(payload).lower()


def test_results_redact_credential_like_keys_at_any_depth():
    from microsoft365.results import normalize_result

    page = {
        "value": [
            {
                "subject": "keep me",
                "client_secret": SENTINEL,  # pragma: allowlist secret
                "access_token": SENTINEL,
                "refreshToken": SENTINEL,
                "Authorization": f"Bearer {SENTINEL}",
                "X-API-KEY": SENTINEL,
                "password": SENTINEL,
                "Set-Cookie": SENTINEL,
                "private_key": SENTINEL,
                "credential": SENTINEL,
                "signature": SENTINEL,
                "nested": {"deeper": {"clientSecret": SENTINEL, "authorization": SENTINEL}},
            }
        ]
    }
    rendered = json.dumps(normalize_result(page))

    assert SENTINEL not in rendered
    assert "Bearer" not in rendered
    assert "keep me" in rendered
    assert "[REDACTED]" in rendered


def test_results_bound_strings_break_cycles_and_handle_bytes():
    from microsoft365.results import normalize_result

    cyclic: dict = {"name": "root"}
    cyclic["self"] = cyclic
    payload = normalize_result({"cycle": cyclic, "blob": b"\x00\x01", "long": "y" * 9000})

    assert payload["cycle"]["self"] == "[CYCLE]"
    assert payload["blob"] == {"type": "bytes", "size": 2}
    assert len(payload["long"]) < 9000
    assert payload["long"].endswith("[TRUNCATED]")


def test_results_recurse_through_model_additional_data_without_leaking():
    from msgraph.generated.models.message import Message
    from microsoft365.results import normalize_result

    model = Message(subject="hello", additional_data={"outer": {"inner": {"access_token": SENTINEL}}})
    rendered = json.dumps(normalize_result(model))

    assert "hello" in rendered
    assert SENTINEL not in rendered


def test_exception_with_a_cyclic_additional_data_terminates(tmp_path):
    error = _kiota_api_error(status=429, headers={})
    error.additional_data = {"self": error, "deep": {"deeper": {"secret_thing": SENTINEL}}}  # pragma: allowlist secret

    converted = taxonomy.to_graph_error(error)

    assert converted.category == "throttled"
    assert SENTINEL not in json.dumps(converted.to_result())
    assert not (tmp_path / "noop").exists()


def test_error_result_helper_is_the_single_conversion_point():
    result = taxonomy.error_result(_kiota_api_error(status=429, headers={"Retry-After": "2"}))

    assert result == {
        "error": "throttled",
        "message": taxonomy.MESSAGES["throttled"],
        "retryable": True,
        "retry_after_seconds": 2.0,
        "status_code": 429,
    }
    assert SENTINEL not in json.dumps(result)


def test_error_result_omits_absent_optional_fields():
    result = taxonomy.error_result(Exception(SENTINEL))

    assert result == {
        "error": "internal_error",
        "message": taxonomy.MESSAGES["internal_error"],
        "retryable": False,
    }


def test_status_code_is_bounded_to_real_http_codes():
    error = _kiota_api_error(status=999999)
    assert taxonomy.to_graph_error(error).status_code is None

    error = _kiota_api_error(status=True)
    assert taxonomy.to_graph_error(error).status_code is None


def test_registration_unavailable_payload_uses_a_taxonomy_category():
    from microsoft365.registration import service_tool_handler

    payload = json.loads(service_tool_handler("outlook", {"action": "search"}))

    assert payload["error"] in taxonomy.CATEGORIES


def test_taxonomy_module_keeps_no_credential_client_or_secret_access():
    source = Path(taxonomy.__file__).read_text(encoding="utf-8")

    for forbidden in ("get_secret", "secret_scope", "ClientSecretCredential", "GraphServiceClient"):
        assert forbidden not in source
    # The only header values this module may read are the bounded retry/correlation ones.
    assert re.search(r"response_headers|\.headers\b", source)


# ---------------------------------------------------------------------------------------
# Retry policy
# ---------------------------------------------------------------------------------------


def test_retry_decision_honours_retry_after_for_a_safe_read():
    decision = taxonomy.retry_decision(_kiota_api_error(status=429, headers={"Retry-After": "7"}), method="GET", attempt=0)

    assert decision.retry is True
    assert decision.delay_seconds == 7.0
    assert decision.category == "throttled"
    assert SENTINEL not in repr(decision)


def test_retry_decision_falls_back_to_a_bounded_default_delay():
    decision = taxonomy.retry_decision(_kiota_api_error(status=503, headers={}), method="GET", attempt=0)

    assert decision.retry is True
    assert decision.delay_seconds == taxonomy.DEFAULT_BACKOFF_SECONDS
    assert decision.delay_seconds <= taxonomy.MAX_RETRY_AFTER_SECONDS


@pytest.mark.parametrize("method", ["POST", "PUT", "PATCH", "DELETE", "post", "put"])
def test_retry_decision_refuses_to_replay_a_write_even_when_throttled(method):
    decision = taxonomy.retry_decision(
        _kiota_api_error(status=429, headers={"Retry-After": "1"}), method=method, attempt=0
    )

    assert decision.retry is False
    assert decision.delay_seconds is None
    assert decision.reason == "method_not_safe_for_retry"


@pytest.mark.parametrize("method", ["GET", "HEAD", "OPTIONS", "get"])
def test_retry_decision_allows_only_safe_methods(method):
    assert taxonomy.retry_decision(_kiota_api_error(status=429, headers={}), method=method, attempt=0).retry is True


def test_retry_decision_stops_when_attempts_are_exhausted():
    decision = taxonomy.retry_decision(_kiota_api_error(status=429, headers={}), method="GET", attempt=2, max_attempts=3)

    assert decision.retry is False
    assert decision.reason == "attempts_exhausted"
    assert decision.delay_seconds is None


def test_retry_decision_refuses_a_non_retryable_category():
    decision = taxonomy.retry_decision(_kiota_api_error(status=404, headers={}), method="GET", attempt=0)

    assert decision.retry is False
    assert decision.reason == "category_not_retryable"
    assert decision.category == "not_found"


def test_retry_decision_reason_codes_are_a_closed_set():
    decisions = [
        taxonomy.retry_decision(_kiota_api_error(status=429, headers={"Retry-After": "1"}), method="GET", attempt=0),
        taxonomy.retry_decision(_kiota_api_error(status=503, headers={}), method="GET", attempt=0),
        taxonomy.retry_decision(_kiota_api_error(status=429, headers={}), method="POST", attempt=0),
        taxonomy.retry_decision(_kiota_api_error(status=404, headers={}), method="GET", attempt=0),
        taxonomy.retry_decision(_kiota_api_error(status=429, headers={}), method="GET", attempt=9, max_attempts=3),
    ]
    reasons = {decision.reason for decision in decisions}

    assert reasons == {
        "retry_after",
        "transient",
        "method_not_safe_for_retry",
        "category_not_retryable",
        "attempts_exhausted",
    }


@pytest.mark.parametrize(
    "kwargs",
    [
        {"method": "GET", "attempt": 0, "max_attempts": 0},
        {"method": "GET", "attempt": 0, "max_attempts": 99},
        {"method": "GET", "attempt": 0, "max_attempts": "3"},
        {"method": "GET", "attempt": 0, "max_attempts": True},
        {"method": "GET", "attempt": -1},
        {"method": "GET", "attempt": "0"},
        {"method": "TRACE", "attempt": 0},
        {"method": None, "attempt": 0},
    ],
)
def test_invalid_retry_configuration_fails_closed(kwargs):
    with pytest.raises(taxonomy.GraphError) as caught:
        taxonomy.retry_decision(_kiota_api_error(status=429, headers={}), **kwargs)

    assert caught.value.category == "configuration_error"


def test_run_with_retry_retries_a_throttled_read_within_the_bound():
    calls = []
    sleeps = []
    error = _kiota_api_error(status=429, headers={"Retry-After": "3"})

    def action():
        calls.append(1)
        raise error

    with pytest.raises(taxonomy.GraphError) as caught:
        taxonomy.run_with_retry(action, method="GET", max_attempts=3, sleep=sleeps.append)

    assert len(calls) == 3
    assert sleeps == [3.0, 3.0]
    assert caught.value.category == "throttled"
    assert isinstance(caught.value.__cause__, Exception)


def test_run_with_retry_returns_the_first_success_without_sleeping():
    calls = []
    sleeps = []

    def action():
        calls.append(1)
        if len(calls) < 2:
            raise _kiota_api_error(status=503, headers={})
        return {"value": ["ok"]}

    assert taxonomy.run_with_retry(action, method="GET", max_attempts=3, sleep=sleeps.append) == {"value": ["ok"]}
    assert len(calls) == 2
    assert sleeps == [taxonomy.DEFAULT_BACKOFF_SECONDS]


@pytest.mark.parametrize("method", ["POST", "PUT", "PATCH", "DELETE"])
def test_run_with_retry_never_replays_a_write(method):
    calls = []
    sleeps = []

    def action():
        calls.append(1)
        raise _kiota_api_error(status=429, headers={"Retry-After": "1"})

    with pytest.raises(taxonomy.GraphError) as caught:
        taxonomy.run_with_retry(action, method=method, max_attempts=5, sleep=sleeps.append)

    assert len(calls) == 1
    assert sleeps == []
    assert caught.value.category == "throttled"


def test_run_with_retry_never_replays_a_send_or_an_upload():
    sends = []

    def send():
        sends.append("send")
        raise _kiota_api_error(status=503, headers={"Retry-After": "1"})

    def upload():
        sends.append("upload")
        raise _kiota_api_error(status=429, headers={"Retry-After": "1"})

    with pytest.raises(taxonomy.GraphError):
        taxonomy.run_with_retry(send, method="POST", sleep=lambda seconds: None)
    with pytest.raises(taxonomy.GraphError):
        taxonomy.run_with_retry(upload, method="PUT", sleep=lambda seconds: None)

    assert sends == ["send", "upload"]


def test_run_with_retry_stops_on_a_non_retryable_failure():
    calls = []

    def action():
        calls.append(1)
        raise _kiota_api_error(status=404, headers={})

    with pytest.raises(taxonomy.GraphError) as caught:
        taxonomy.run_with_retry(action, method="GET", max_attempts=3, sleep=lambda seconds: None)

    assert len(calls) == 1
    assert caught.value.category == "not_found"


def test_run_with_retry_ignores_an_unparseable_retry_after():
    sleeps = []
    error = _kiota_api_error(status=503, headers={"Retry-After": SENTINEL})

    calls = []

    def action():
        calls.append(1)
        raise error

    with pytest.raises(taxonomy.GraphError):
        taxonomy.run_with_retry(action, method="GET", max_attempts=2, sleep=sleeps.append)

    assert calls == [1, 1]
    assert sleeps == [taxonomy.DEFAULT_BACKOFF_SECONDS]


def test_run_with_retry_never_sleeps_longer_than_the_bound():
    sleeps = []

    def action():
        raise _kiota_api_error(status=429, headers={"Retry-After": "999999"})

    with pytest.raises(taxonomy.GraphError):
        taxonomy.run_with_retry(action, method="GET", max_attempts=2, sleep=sleeps.append)

    assert sleeps == [taxonomy.MAX_RETRY_AFTER_SECONDS]


@pytest.mark.parametrize("kwargs", [{"method": "GET", "max_attempts": 0}, {"method": "GET", "sleep": None}, {"method": "TRACE"}])
def test_run_with_retry_fails_closed_on_invalid_configuration(kwargs):
    def action():  # pragma: no cover - must never run
        raise AssertionError("action must not run with invalid retry configuration")

    with pytest.raises(taxonomy.GraphError) as caught:
        taxonomy.run_with_retry(action, **kwargs)

    assert caught.value.category == "configuration_error"


def test_run_with_retry_sanitizes_a_non_graph_failure():
    def action():
        raise RuntimeError(f"client_secret={SENTINEL}")  # pragma: allowlist secret

    with pytest.raises(taxonomy.GraphError) as caught:
        taxonomy.run_with_retry(action, method="GET", max_attempts=1)

    assert caught.value.category == "internal_error"
    assert SENTINEL not in str(caught.value)


# ---------------------------------------------------------------------------------------
# The execution seam consumes the canonical taxonomy
# ---------------------------------------------------------------------------------------


def test_seam_error_is_a_taxonomy_error():
    from microsoft365.execution import ExecutionError

    assert issubclass(ExecutionError, taxonomy.GraphError)
    assert taxonomy.classify(ExecutionError("throttled", taxonomy.MESSAGES["throttled"])) == "throttled"


def test_seam_error_categories_must_be_taxonomy_categories():
    from microsoft365.execution import ExecutionError

    with pytest.raises(ValueError):
        ExecutionError("definitely_not_a_category", "message")


def test_seam_classifies_a_real_graph_failure_by_its_attributes():
    import asyncio

    from microsoft365.execution import run_async

    async def failing():
        raise _odata_error(code="TooManyRequests", status=429, headers={"Retry-After": "7"})

    with pytest.raises(taxonomy.GraphError) as caught:
        run_async(failing(), category="transport_error")

    assert caught.value.category == "throttled"
    assert caught.value.retry_after_seconds == 7.0
    assert caught.value.status_code == 429
    assert SENTINEL not in str(caught.value)


def test_seam_keeps_its_supplied_default_for_unclassifiable_failures():
    from microsoft365.execution import run_async

    async def failing():
        raise AssertionError("network transport must not run")

    with pytest.raises(taxonomy.GraphError) as caught:
        run_async(failing(), category="transport_error")

    assert caught.value.category == "transport_error"


def test_seam_preserves_a_taxonomy_category_from_the_client_factory():
    from microsoft365.execution import resolve_request_adapter

    def factory():
        raise taxonomy.GraphError("unsupported_auth_mode", taxonomy.MESSAGES["unsupported_auth_mode"])

    with pytest.raises(taxonomy.GraphError) as caught:
        resolve_request_adapter(factory)

    assert caught.value.category == "unsupported_auth_mode"


def test_seam_still_wraps_an_unknown_factory_failure_as_authentication_required():
    import asyncio  # noqa: F401 - kept for symmetry with the module under test

    from microsoft365.execution import resolve_request_adapter

    def factory():
        raise RuntimeError(f"client_secret={SENTINEL}")  # pragma: allowlist secret

    with pytest.raises(taxonomy.GraphError) as caught:
        resolve_request_adapter(factory)

    assert caught.value.category == "authentication_required"
    assert SENTINEL not in str(caught.value)


# ---------------------------------------------------------------------------------------
# Client construction uses the taxonomy and never echoes credential material
# ---------------------------------------------------------------------------------------


def test_client_rejects_delegated_mode_as_unsupported_auth_mode(monkeypatch):
    import agent.secret_scope

    from microsoft365.client import create_graph_client
    from microsoft365.contract import Settings

    monkeypatch.setattr(
        agent.secret_scope, "get_secret", lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("secret read"))
    )

    with pytest.raises(taxonomy.GraphError) as caught:
        create_graph_client(Settings(authentication_mode="delegated", tenant_id="t", client_id="c"))

    assert caught.value.category == "unsupported_auth_mode"


def test_client_requires_identity_configuration_before_reading_the_secret(monkeypatch):
    import agent.secret_scope

    from microsoft365.client import create_graph_client
    from microsoft365.contract import Settings

    calls = []
    monkeypatch.setattr(agent.secret_scope, "get_secret", lambda *args, **kwargs: calls.append(1) or SENTINEL)

    with pytest.raises(taxonomy.GraphError) as caught:
        create_graph_client(Settings(tenant_id="", client_id="client"))

    assert caught.value.category == "configuration_error"
    assert calls == []


def test_client_missing_secret_is_authentication_required_without_echoing_anything(monkeypatch):
    import agent.secret_scope

    from microsoft365.client import create_graph_client
    from microsoft365.contract import Settings

    monkeypatch.setattr(agent.secret_scope, "get_secret", lambda name, default=None: "")

    with pytest.raises(taxonomy.GraphError) as caught:
        create_graph_client(Settings(tenant_id="tenant", client_id="client"))

    assert caught.value.category == "authentication_required"
    assert SENTINEL not in str(caught.value)
    assert "client_secret" not in str(caught.value).lower()


def test_client_sanitizes_a_credential_construction_failure(monkeypatch):
    import agent.secret_scope
    import azure.identity

    from microsoft365.client import create_graph_client
    from microsoft365.contract import Settings

    monkeypatch.setattr(agent.secret_scope, "get_secret", lambda name, default=None: SENTINEL)

    def exploding(**kwargs):
        raise RuntimeError(f"invalid client secret {kwargs.get('client_secret')}")

    monkeypatch.setattr(azure.identity, "ClientSecretCredential", exploding)

    with pytest.raises(taxonomy.GraphError) as caught:
        create_graph_client(Settings(tenant_id="tenant", client_id="client"))

    assert caught.value.category == "authentication_required"
    assert SENTINEL not in str(caught.value)
    assert SENTINEL not in json.dumps(caught.value.to_result())
    assert isinstance(caught.value.__cause__, RuntimeError)


def test_client_never_puts_the_secret_in_its_reported_error(monkeypatch):
    import agent.secret_scope
    import azure.identity
    import msgraph

    from microsoft365.client import create_graph_client
    from microsoft365.contract import Settings

    captured = []
    monkeypatch.setattr(agent.secret_scope, "get_secret", lambda name, default=None: SENTINEL)
    monkeypatch.setattr(azure.identity, "ClientSecretCredential", lambda **kwargs: captured.append(kwargs) or "cred")
    monkeypatch.setattr(msgraph, "GraphServiceClient", lambda **kwargs: captured.append(kwargs) or "client")

    assert create_graph_client(Settings(tenant_id="tenant", client_id="client")) == "client"
    assert captured[0]["client_secret"] == SENTINEL
    # The value is consumed by the SDK and never rendered by the plugin.
    assert SENTINEL not in str(taxonomy.MESSAGES["authentication_required"])
