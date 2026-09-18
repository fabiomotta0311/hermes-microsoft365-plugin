"""WP12 adversarial coverage: approval, dispatch, retry, and registry invariants."""
from __future__ import annotations

import json

import pytest

from tests.test_validation import (
    EXECUTABLE_READS,
    SERVICES,
    USER,
    VALID_PAYLOADS,
    WITHHELD_WRITES,
    RecordingContext,
    _settings,
)


def _registered(capabilities=None):
    from microsoft365 import register

    ctx = RecordingContext(
        {
            "tenant_id": "tenant",
            "client_id": "client",
            "authentication_mode": "application",
            "capabilities": capabilities or {service: True for service in SERVICES},
        }
    )
    register(ctx)
    return ctx


def test_final_payload_mutation_from_read_to_write_is_blocked_before_runtime(monkeypatch):
    """Approval of the first payload cannot authorize a later action mutation."""
    import agent.secret_scope
    import azure.identity
    import msgraph
    from microsoft365 import registration

    reached = []

    def forbidden(name):
        def fail(*args, **kwargs):
            reached.append(name)
            raise AssertionError(name)

        return fail

    monkeypatch.setattr(agent.secret_scope, "get_secret", forbidden("secret"))
    monkeypatch.setattr(azure.identity, "ClientSecretCredential", forbidden("credential"))
    monkeypatch.setattr(msgraph, "GraphServiceClient", forbidden("client"))

    ctx = _registered()
    payload = {"action": "search", "user_id": USER}
    assert ctx.hooks["pre_tool_call"](tool_name="microsoft365_outlook", args=payload) is None
    payload["action"] = "send"

    result = json.loads(
        registration.service_tool_handler("outlook", payload, settings=_settings())
    )
    assert result["error"] == "operation_not_implemented"
    assert reached == []


def test_disabled_valid_operation_is_refused_before_secret_credential_or_client(monkeypatch):
    import agent.secret_scope
    import azure.identity
    import msgraph
    from microsoft365 import registration

    reached = []

    def forbidden(name):
        def fail(*args, **kwargs):
            reached.append(name)
            raise AssertionError(name)

        return fail

    monkeypatch.setattr(agent.secret_scope, "get_secret", forbidden("secret"))
    monkeypatch.setattr(azure.identity, "ClientSecretCredential", forbidden("credential"))
    monkeypatch.setattr(msgraph, "GraphServiceClient", forbidden("client"))

    settings = _settings("disabled_read")
    result = json.loads(
        registration.service_tool_handler(
            "outlook", {"action": "read", "user_id": USER, "message_id": "message"}, settings=settings
        )
    )
    assert result["error"] == "operation_not_implemented"
    assert reached == []


def test_malformed_write_is_blocked_before_approval_even_if_host_marks_write_executable(monkeypatch):
    from microsoft365 import registration

    ctx = _registered({"outlook": {"send": True}})
    monkeypatch.setattr(registration, "active_actions", lambda settings, service: ("send",))
    directive = ctx.hooks["pre_tool_call"](
        tool_name="microsoft365_outlook",
        args={"action": "send", "user_id": USER, "message_id": "message", "save_to_sent_items": "false"},
    )
    assert directive["action"] == "block"
    assert "save_to_sent_items" in directive["message"]
    assert "rule_key" not in directive


def test_approval_directive_shape_is_exact_when_host_marks_a_write_executable(monkeypatch):
    """Plugin output is explicit; host CORE-1/CORE-2 remains outside this plugin."""
    from microsoft365 import registration

    ctx = _registered({"outlook": {"send": True}})
    monkeypatch.setattr(registration, "active_actions", lambda settings, service: ("send",))
    directive = ctx.hooks["pre_tool_call"](
        tool_name="microsoft365_outlook",
        args={"action": "send", "user_id": USER, "message_id": "message"},
    )
    assert directive == {
        "action": "approve",
        "message": "Microsoft 365 send: external side effect",
        "rule_key": "microsoft365.outlook.send",
    }


@pytest.mark.parametrize("key", sorted(WITHHELD_WRITES))
def test_every_write_remains_non_executable_and_has_a_handler(key):
    from microsoft365.contract import OPERATION_REGISTRY
    from microsoft365.registration import HANDLER_TABLE

    assert OPERATION_REGISTRY[key].write is True
    assert OPERATION_REGISTRY[key].executable is False
    assert key in HANDLER_TABLE


def test_registry_and_handler_table_cannot_drift():
    from microsoft365.contract import OPERATION_REGISTRY
    from microsoft365.registration import HANDLER_TABLE

    assert set(HANDLER_TABLE) == set(OPERATION_REGISTRY)
    assert {key for key, item in OPERATION_REGISTRY.items() if item.executable} == set(EXECUTABLE_READS)


def test_retry_policy_allows_only_idempotent_methods_and_never_replays_writes():
    from microsoft365.errors import GraphError, retry_decision

    transient = GraphError("throttled", "throttled", retryable=True)
    assert retry_decision(transient, method="GET", attempt=0).retry is True
    for method in ("POST", "PUT", "PATCH", "DELETE", "post"):
        assert retry_decision(transient, method=method, attempt=0).retry is False


def test_retry_policy_fails_closed_for_malformed_configuration():
    from microsoft365.errors import GraphError, retry_decision

    transient = GraphError("throttled", "throttled", retryable=True)
    with pytest.raises(GraphError) as caught:
        retry_decision(transient, method="GET", attempt=-1)
    assert caught.value.category == "configuration_error"


def test_payload_validation_precedes_authentication_for_unknown_and_malformed_calls(monkeypatch):
    import agent.secret_scope
    from microsoft365 import registration

    calls = []
    monkeypatch.setattr(agent.secret_scope, "get_secret", lambda *args, **kwargs: calls.append(args))
    ctx = _registered({"outlook": {"search": True}})
    tool = ctx.tools["microsoft365_outlook"]["handler"]

    for payload in ({"action": "search", "user_id": 7}, {"action": "unknown"}):
        result = json.loads(tool(payload))
        assert result["error"] == "validation_error"
    assert calls == []
