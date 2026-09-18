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


_READ_TO_WRITE_CASES = (
    ("sharepoint", "download_files", "upload_files"),
    ("onedrive", "download_files", "upload_files"),
    ("outlook", "read", "create_draft"),
    ("outlook", "read", "send"),
    ("calendar", "search", "create_events"),
    ("calendar", "search", "update_events"),
    ("todo", "read", "create_tasks"),
    ("todo", "read", "update_tasks"),
    ("planner", "read", "create_tasks"),
    ("planner", "read", "update_tasks"),
)


def _runtime_probes(monkeypatch):
    import agent.secret_scope
    import azure.identity
    import msgraph

    reached = []

    def forbidden(name):
        def fail(*args, **kwargs):
            reached.append(name)
            raise AssertionError(name)

        return fail

    monkeypatch.setattr(agent.secret_scope, "get_secret", forbidden("secret"))
    monkeypatch.setattr(azure.identity, "ClientSecretCredential", forbidden("credential"))
    monkeypatch.setattr(msgraph, "GraphServiceClient", forbidden("client"))
    return reached


@pytest.mark.parametrize("service,read_action,write_action", _READ_TO_WRITE_CASES)
def test_registered_read_to_write_mutations_remain_non_executable(
    monkeypatch, service, read_action, write_action
):
    """A final-payload mutation cannot turn any registered read into a write."""
    from microsoft365 import registration

    reached = _runtime_probes(monkeypatch)
    ctx = _registered()
    hook = ctx.hooks["pre_tool_call"]
    read_payload = dict(VALID_PAYLOADS[f"{service}.{read_action}"])
    write_payload = dict(VALID_PAYLOADS[f"{service}.{write_action}"])

    assert hook(tool_name=f"microsoft365_{service}", args=read_payload) is None
    # Model the host-owned CORE-1 mutation: retain the approved object identity, but replace
    # every operation-specific property with a valid destination/write payload.
    read_payload.clear()
    read_payload.update(write_payload)

    result = json.loads(ctx.tools[f"microsoft365_{service}"]["handler"](read_payload))
    assert result["error"] == "operation_not_implemented"
    assert result["operation"] == write_action
    assert reached == []

    # The plugin-local dispatch seam independently reaches the same closed result.
    assert json.loads(
        registration.service_tool_handler(service, read_payload, settings=_settings())
    ) == result


def test_all_write_destinations_and_identifiers_are_rechecked_after_multiple_mutations(monkeypatch):
    """Repeated action, destination, and user-id rewrites never reach client construction."""
    from microsoft365 import registration

    reached = _runtime_probes(monkeypatch)
    ctx = _registered()
    hook = ctx.hooks["pre_tool_call"]
    payload = dict(VALID_PAYLOADS["outlook.search"])
    assert hook(tool_name="microsoft365_outlook", args=payload) is None

    # Several modifiers in a loop model chained host middleware. Each final payload is a
    # complete, validator-approved write, while identifiers deliberately come from a different
    # destination/user than the originally approved read.
    mutations = (
        VALID_PAYLOADS["outlook.send"],
        VALID_PAYLOADS["calendar.create_events"],
        VALID_PAYLOADS["todo.update_tasks"],
        VALID_PAYLOADS["planner.create_tasks"],
        VALID_PAYLOADS["onedrive.upload_files"],
    )
    for candidate in mutations:
        # The destination is selected from the payload itself; dispatch must not trust the
        # tool name or any stale approval associated with the original Outlook read.
        service_name = {
            "send": "outlook", "create_events": "calendar", "update_tasks": "todo",
            "create_tasks": "planner", "upload_files": "onedrive",
        }[candidate["action"]]
        payload.clear()
        payload.update(candidate)
        result = json.loads(ctx.tools[f"microsoft365_{service_name}"]["handler"](payload))
        assert result["error"] == "operation_not_implemented"
        assert result["operation"] == candidate["action"]
    assert reached == []


def test_malformed_or_host_owned_approval_directives_are_never_emitted_by_plugin(monkeypatch):
    """Plugin directives are either exact approvals or exact blocks; CORE-1/CORE-2 are host-owned."""
    from microsoft365 import registration

    def simulated_active_actions(settings, service):
        return tuple(settings.selected(service))

    monkeypatch.setattr(registration, "active_actions", simulated_active_actions)
    ctx = _registered(capabilities={"outlook": {"send": True}})
    hook = ctx.hooks["pre_tool_call"]
    cases = (
        {"action": "send", "user_id": USER, "message_id": "message", "save_to_sent_items": "false"},
        {"action": "send", "user_id": USER, "message_id": "message", "unexpected": True},
        {"action": "send", "user_id": USER},
        {"action": "send", "user_id": USER, "message_id": "message"},
    )
    for payload in cases:
        directive = hook(tool_name="microsoft365_outlook", args=payload)
        assert directive is not None
        if directive["action"] == "approve":
            assert directive == {
                "action": "approve",
                "message": "Microsoft 365 send: external side effect",
                "rule_key": "microsoft365.outlook.send",
            }
        else:
            assert directive["action"] == "block"
            assert set(directive) == {"action", "message"}
            assert "rule_key" not in directive

    # A host that rewrites arguments after an approval is outside plugin control; the plugin
    # still emits no malformed directive and dispatch remains fail-closed.
    assert hook(tool_name="microsoft365_outlook", args=VALID_PAYLOADS["outlook.send"]) == {
        "action": "approve",
        "message": "Microsoft 365 send: external side effect",
        "rule_key": "microsoft365.outlook.send",
    }
