from __future__ import annotations

import json


def _settings(*, user_id="user-a", mode="application", capabilities=None):
    from microsoft365.contract import Settings

    return Settings.from_mapping(
        {
            "tenant_id": "tenant-a",
            "client_id": "client-a",
            "user_id": user_id,
            "authentication_mode": mode,
            "capabilities": capabilities or {"outlook": {"search": True}},
        }
    )


def test_preflight_derives_identity_requirements_and_redacts_identifiers(monkeypatch):
    import agent.secret_scope
    from microsoft365.preflight import build_preflight

    calls = []
    monkeypatch.setattr(agent.secret_scope, "get_secret", lambda *args: calls.append(args) or "secret")
    result = build_preflight(_settings(), sdk_available=True)

    assert calls == [("MICROSOFT365_CLIENT_SECRET", None)]
    assert result["requirements"] == {"tenant_id": True, "client_id": True, "user_id": True}
    assert result["configuration"] == {
        "tenant_id_configured": True,
        "client_id_configured": True,
        "user_id_configured": True,
        "authentication_mode": "application",
        "capabilities": {"outlook": {"search": True}},
    }
    assert "tenant-a" not in json.dumps(result)
    assert "client-a" not in json.dumps(result)
    assert "user-a" not in json.dumps(result)


def test_planner_only_preflight_does_not_require_user_or_secret_in_delegated_mode(monkeypatch):
    import agent.secret_scope
    from microsoft365.preflight import build_preflight
    from microsoft365.contract import Settings

    monkeypatch.setattr(agent.secret_scope, "get_secret", lambda *args: (_ for _ in ()).throw(AssertionError("secret")))
    result = build_preflight(
        Settings.from_mapping({"authentication_mode": "delegated", "capabilities": {"planner": True}}),
        sdk_available=True,
    )

    assert result["requirements"] == {"tenant_id": False, "client_id": False, "user_id": False}
    assert result["secret_present"] is None
    assert result["missing"] == []
    assert result["configuration"] == {
        "tenant_id_configured": False,
        "client_id_configured": False,
        "user_id_configured": False,
        "authentication_mode": "delegated",
        "capabilities": {"planner": {key: True for key in (
            "list_plans", "list_buckets", "list_tasks", "read", "create_tasks", "update_tasks"
        )}},
    }


def test_preflight_user_requirement_is_derived_from_selected_operation_metadata():
    from microsoft365.contract import Settings
    from microsoft365.preflight import build_preflight

    result = build_preflight(
        Settings.from_mapping({"tenant_id": "t", "client_id": "c", "capabilities": {"planner": True}}),
        sdk_available=True,
    )
    assert result["requirements"]["user_id"] is False
    assert "user_id" not in result["missing"]


def test_divergent_user_is_rejected_before_secret_credential_and_client(monkeypatch):
    import agent.secret_scope
    import azure.identity
    import msgraph
    from microsoft365 import register
    from tests.test_validation import RecordingContext

    reached = []
    for module, name in ((agent.secret_scope, "secret"), (azure.identity, "credential"), (msgraph, "client")):
        monkeypatch.setattr(module, "get_secret" if name == "secret" else (
            "ClientSecretCredential" if name == "credential" else "GraphServiceClient"
        ), lambda *args, _name=name, **kwargs: reached.append(_name))

    ctx = RecordingContext({
        "tenant_id": "tenant-a", "client_id": "client-a", "user_id": "user-a",
        "capabilities": {"outlook": {"search": True}},
    })
    register(ctx)
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"](
        {"action": "search", "user_id": "user-b"}
    ))
    assert result["error"] == "validation_error"
    assert "user-a" not in json.dumps(result) and "user-b" not in json.dumps(result)
    assert reached == []


def test_cross_profile_a_b_a_binding_is_fail_closed_without_secret_or_client(monkeypatch):
    from microsoft365.validation import check

    profile_a = _settings(user_id="user-a")
    profile_b = _settings(user_id="user-b")
    for settings, supplied, allowed in (
        (profile_a, "user-a", True),
        (profile_b, "user-a", False),
        (profile_a, "user-a", True),
    ):
        rejection = check(settings, service="outlook", arguments={"action": "search", "user_id": supplied})
        assert (rejection is None) is allowed


def test_writes_remain_non_executable_after_boundary_is_applied():
    from microsoft365.contract import OPERATION_REGISTRY

    assert all(not definition.executable for definition in OPERATION_REGISTRY.values() if definition.write)
