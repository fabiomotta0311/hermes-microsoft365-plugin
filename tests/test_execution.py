"""WP0 — the single Microsoft Graph execution seam, proven offline.

The transport double is ``StrictTransportAdapter`` from ``tests/test_sdk_contract.py``
(imported, not re-implemented): every ``send_*`` raises ``AssertionError``, so any
accidental network attempt fails the test that made it.
"""
from __future__ import annotations

import asyncio
import inspect
import json
from pathlib import Path
from urllib.parse import urlsplit

import pytest
from kiota_abstractions.base_request_configuration import RequestConfiguration
from kiota_abstractions.request_adapter import RequestAdapter

from tests.test_sdk_contract import StrictTransportAdapter, graph_client


def _query_configuration(builder, **values):
    query_type = getattr(type(builder), f"{type(builder).__name__}GetQueryParameters")
    return RequestConfiguration(query_parameters=query_type(**values))


def test_execution_seam_never_touches_network(graph_client):
    from microsoft365.execution import request_information_sender

    builder = graph_client.users.by_user_id("user").messages
    request = request_information_sender(
        builder,
        method="GET",
        configuration=_query_configuration(builder, top=3, filter="contains(subject,'hello')"),
    )

    recorded = {
        "method": request.http_method.value,
        "path": urlsplit(request.url).path,
        "query": request.query_parameters,
    }

    assert recorded["method"] == "GET"
    assert recorded["path"] == "/users/user/messages"
    assert recorded["query"] == {"%24filter": "contains(subject,'hello')", "%24top": 3}
    # The injected double raises on every send: building must not reach it.
    assert isinstance(graph_client.request_adapter, StrictTransportAdapter)


def test_execution_seam_sends_only_through_the_injected_adapter(graph_client):
    from microsoft365.execution import ExecutionError, execute_request

    builder = graph_client.users.by_user_id("user").messages

    with pytest.raises(ExecutionError) as caught:
        execute_request(builder, method="GET", adapter=graph_client.request_adapter)

    assert caught.value.category == "transport_error"
    assert isinstance(caught.value.__cause__, AssertionError)


def test_execution_seam_routes_reads_through_retry_policy(graph_client, monkeypatch):
    from microsoft365 import execution

    seen = []

    def policy(action, **kwargs):
        seen.append(kwargs["method"])
        return action()

    monkeypatch.setattr(execution, "run_with_retry", policy)
    builder = graph_client.users.by_user_id("user").messages

    with pytest.raises(execution.ExecutionError):
        execution.execute_request(builder, method="GET", adapter=graph_client.request_adapter)

    assert seen == ["GET"]


@pytest.mark.parametrize("method", ["POST", "PUT", "PATCH", "DELETE"])
def test_execution_seam_routes_writes_through_single_attempt_policy(graph_client, monkeypatch, method):
    from microsoft365 import execution

    calls = []

    def policy(action, **kwargs):
        calls.append(kwargs["method"])
        return action()

    monkeypatch.setattr(execution, "run_with_retry", policy)
    if method == "POST":
        builder = graph_client.users.by_user_id("user").messages.by_message_id("message").send
        body = None
    elif method == "PUT":
        builder = graph_client.drives.by_drive_id("drive").items.by_drive_item_id("item").content
        body = b"payload"
    elif method == "PATCH":
        from msgraph.generated.models.event import Event

        builder = graph_client.users.by_user_id("user").calendar.events.by_event_id("event")
        body = Event(subject="subject")
        monkeypatch.setattr(execution, "request_information_sender", lambda *args, **kwargs: None)
    else:
        builder = graph_client.users.by_user_id("user").messages.by_message_id("message")
        body = None

    with pytest.raises(execution.ExecutionError):
        execution.execute_request(builder, method=method, body=body, adapter=graph_client.request_adapter)

    assert calls == [method]


def test_execution_seam_fails_closed_without_an_injected_adapter(graph_client):
    from microsoft365.execution import ExecutionError, execute_request

    builder = graph_client.users.by_user_id("user").messages

    with pytest.raises(ExecutionError) as caught:
        execute_request(builder, method="GET", adapter=None)

    assert caught.value.category == "configuration_error"


def test_execution_seam_rejects_a_builder_not_bound_to_the_injected_adapter(graph_client):
    from microsoft365.execution import ExecutionError, execute_request

    builder = graph_client.users.by_user_id("user").messages

    with pytest.raises(ExecutionError) as caught:
        execute_request(builder, method="GET", adapter=StrictTransportAdapter())

    assert caught.value.category == "configuration_error"


def test_execution_seam_exposes_content_url_headers_and_exact_bytes(graph_client):
    from microsoft365.execution import request_information_sender

    builder = graph_client.drives.by_drive_id("drive").items.by_drive_item_id("root:/a/b.txt:").content
    payload = b"\x00binary\xff"

    request = request_information_sender(builder, method="PUT", body=payload)

    assert request.http_method.value == "PUT"
    assert urlsplit(request.url).path == "/drives/drive/items/root%3A%2Fa%2Fb.txt%3A/content"
    assert "application/octet-stream" in request.headers.get("Content-Type")
    assert request.content == payload


def test_execution_seam_rejects_an_unsupported_method(graph_client):
    from microsoft365.execution import ExecutionError, request_information_sender

    builder = graph_client.users.by_user_id("user").messages

    with pytest.raises(ExecutionError) as caught:
        request_information_sender(builder, method="TRACE")

    assert caught.value.category == "configuration_error"


def test_execution_seam_wraps_a_missing_body_in_a_typed_error(graph_client):
    from microsoft365.execution import ExecutionError, request_information_sender

    builder = graph_client.users.by_user_id("user").messages

    with pytest.raises(ExecutionError) as caught:
        request_information_sender(builder, method="POST")

    assert caught.value.category == "configuration_error"
    assert isinstance(caught.value.__cause__, TypeError)


def test_execution_seam_handles_a_post_without_a_body(graph_client):
    from microsoft365.execution import request_information_sender

    builder = graph_client.users.by_user_id("user").messages.by_message_id("message").send
    request = request_information_sender(builder, method="POST")

    assert request.http_method.value == "POST"
    assert urlsplit(request.url).path == "/users/user/messages/message/send"


def test_run_async_returns_the_value_of_the_awaitable():
    from microsoft365.execution import run_async

    async def answer():
        return 42

    assert run_async(answer()) == 42


def test_run_async_wraps_sdk_failures_in_a_typed_error_without_raw_text():
    from microsoft365.execution import ExecutionError, run_async

    async def boom():
        raise ValueError("client_secret=SENTINEL-raw-sdk-text")

    with pytest.raises(ExecutionError) as caught:
        run_async(boom(), category="transport_error")

    assert caught.value.category == "transport_error"
    assert "SENTINEL-raw-sdk-text" not in str(caught.value)
    assert isinstance(caught.value.__cause__, ValueError)


def test_run_async_rejects_a_non_awaitable():
    from microsoft365.execution import ExecutionError, run_async

    with pytest.raises(ExecutionError) as caught:
        run_async(object())

    assert caught.value.category == "configuration_error"


def test_run_async_fails_closed_inside_a_running_loop():
    from microsoft365.execution import ExecutionError, run_async

    async def nested():
        run_async(asyncio.sleep(0))

    with pytest.raises(ExecutionError) as caught:
        asyncio.run(nested())

    assert caught.value.category == "internal_error"


def test_resolve_request_adapter_returns_the_adapter_of_the_injected_client(graph_client):
    from microsoft365.execution import resolve_request_adapter

    assert resolve_request_adapter(lambda: graph_client) is graph_client.request_adapter


def test_resolve_request_adapter_wraps_credential_failure_in_a_typed_error():
    from microsoft365.execution import ExecutionError, resolve_request_adapter

    sentinel = "SENTINEL-raw-azure-text"

    def failing_factory():
        raise RuntimeError(f"MICROSOFT365_CLIENT_SECRET is not available in the active secret scope: {sentinel}")

    with pytest.raises(ExecutionError) as caught:
        resolve_request_adapter(failing_factory)

    assert caught.value.category == "authentication_required"
    assert sentinel not in str(caught.value)
    assert "MICROSOFT365_CLIENT_SECRET" not in str(caught.value)
    assert isinstance(caught.value.__cause__, RuntimeError)


def test_resolve_request_adapter_fails_closed_without_a_client_factory():
    from microsoft365.execution import ExecutionError, resolve_request_adapter

    with pytest.raises(ExecutionError) as caught:
        resolve_request_adapter(None)

    assert caught.value.category == "configuration_error"


def test_execution_module_keeps_no_credential_or_client_in_module_state():
    from microsoft365 import execution

    module_values = [value for name, value in vars(execution).items() if not name.startswith("_")]
    assert not any(isinstance(value, RequestAdapter) for value in module_values)
    assert not any(type(value).__name__ == "GraphServiceClient" for value in module_values)

    source = Path(execution.__file__).read_text(encoding="utf-8")
    for forbidden in ("get_secret", "secret_scope", "ClientSecretCredential", "azure.identity"):
        assert forbidden not in source


def test_generated_builder_methods_are_async_so_one_seam_is_required(graph_client):
    from microsoft365.execution import run_async

    builder = graph_client.users.by_user_id("user").messages

    assert inspect.iscoroutinefunction(builder.get)
    assert not inspect.iscoroutinefunction(run_async)


def test_registration_invokes_async_handlers_through_the_seam():
    from microsoft365 import registration

    async def handler(args):
        return {"action": args["action"], "executed_through_seam": True}

    assert registration.invoke_handler(handler, {"action": "teams.list_teams"}) == {
        "action": "teams.list_teams",
        "executed_through_seam": True,
    }


def test_service_tool_handler_keeps_unavailable_default_for_unhandled_operations():
    """No handler, no dispatch: the payload is unchanged for an operation without one."""
    from microsoft365 import registration

    assert json.loads(registration.service_tool_handler("planner", {"action": "list_plans"}))["error"] == "configuration_error"
    assert json.loads(registration.service_tool_handler("planner", None)) == {
        "error": "operation_not_implemented",
        "service": "planner",
    }
    # no operation named at all: nothing to dispatch, for any service
    assert json.loads(registration.service_tool_handler("outlook", None)) == {
        "error": "operation_not_implemented",
        "service": "outlook",
    }


def test_an_executable_read_dispatches_and_reports_the_real_failure(monkeypatch):
    """``outlook.search`` is executable now, so dispatch reaches the client factory.

    With no secret available the honest outcome is the taxonomy's ``authentication_required``,
    not ``operation_not_implemented``: the operation is implemented and exposed, and what is
    missing is the credential. The secret lookup is recorded, which proves the read really
    dispatched instead of being answered by a stub.
    """
    import agent.secret_scope

    from microsoft365 import registration
    from microsoft365.contract import Settings
    from microsoft365.errors import CATEGORIES
    from microsoft365.preflight import SECRET_ENV_NAME

    lookups: list = []
    monkeypatch.setattr(
        agent.secret_scope,
        "get_secret",
        lambda name, default=None: lookups.append(name) or None,
    )

    payload = json.loads(
        registration.service_tool_handler(
            "outlook",
            {"action": "search", "user_id": "user"},
            settings=Settings.from_mapping(
                {
                    "tenant_id": "tenant",
                    "client_id": "client",
                    "capabilities": {"outlook": {"search": True}},
                }
            ),
        )
    )

    assert payload["error"] == "authentication_required"
    assert payload["error"] in CATEGORIES
    assert lookups == [SECRET_ENV_NAME]


def test_service_tool_handler_returns_a_typed_error_and_never_raw_sdk_text(monkeypatch, graph_client):
    from microsoft365 import execution, registration

    # Supported handler shape: synchronous, using the seam for the graph call.
    def handler(args):
        del args
        builder = graph_client.users.by_user_id("user").messages
        return execution.execute_request(builder, method="GET", adapter=graph_client.request_adapter)

    monkeypatch.setitem(registration.HANDLER_TABLE, "outlook.search", handler)

    payload = json.loads(registration.service_tool_handler("outlook", {"action": "search"}))

    assert payload["error"] == "transport_error"
    assert "network transport must not run" not in json.dumps(payload)


@pytest.mark.parametrize(
    "category",
    ("authentication_required", "permission_denied", "throttled"),
)
def test_registered_dispatch_returns_canonical_graph_error_envelope(monkeypatch, category):
    from microsoft365 import register, registration
    from microsoft365.errors import GraphError, MESSAGES
    from tests.test_registration_secrets import RecordingContext

    def handler(args):
        del args
        raise GraphError(
            category,
            MESSAGES[category],
            retryable=category == "throttled",
            retry_after_seconds=2 if category == "throttled" else None,
        )

    monkeypatch.setitem(registration.HANDLER_TABLE, "outlook.search", handler)
    monkeypatch.setattr(
        registration,
        "active_actions",
        lambda settings, service: ("search",) if service == "outlook" else (),
    )
    ctx = RecordingContext({
        "tenant_id": "tenant",
        "client_id": "client",
        "user_id": "user",
        "capabilities": {"outlook": {"search": True}},
    })
    register(ctx)

    payload = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({"action": "search", "user_id": "user"}))

    assert payload == {
        "error": category,
        "message": MESSAGES[category],
        "retryable": category == "throttled",
        **({"retry_after_seconds": 2} if category == "throttled" else {}),
    }
    assert "traceback" not in json.dumps(payload).lower()


def test_registered_dispatch_preserves_execution_error_envelope(monkeypatch):
    from microsoft365 import register, registration
    from microsoft365.execution import ExecutionError
    from microsoft365.errors import MESSAGES
    from tests.test_registration_secrets import RecordingContext

    def handler(args):
        del args
        raise ExecutionError("service_error", MESSAGES["service_error"])

    monkeypatch.setitem(registration.HANDLER_TABLE, "outlook.search", handler)
    monkeypatch.setattr(
        registration,
        "active_actions",
        lambda settings, service: ("search",) if service == "outlook" else (),
    )
    ctx = RecordingContext({
        "tenant_id": "tenant",
        "client_id": "client",
        "user_id": "user",
        "capabilities": {"outlook": {"search": True}},
    })
    register(ctx)

    payload = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({"action": "search", "user_id": "user"}))

    assert payload == {
        "error": "service_error",
        "message": str(ExecutionError("service_error", MESSAGES["service_error"])),
    }


def test_async_handler_must_not_nest_the_seam_and_fails_closed(monkeypatch, graph_client):
    from microsoft365 import execution, registration

    async def nested_handler(args):
        del args
        builder = graph_client.users.by_user_id("user").messages
        return execution.execute_request(builder, method="GET", adapter=graph_client.request_adapter)

    monkeypatch.setitem(registration.HANDLER_TABLE, "outlook.read", nested_handler)

    payload = json.loads(registration.service_tool_handler("outlook", {"action": "read"}))

    assert payload["error"] == "internal_error"
