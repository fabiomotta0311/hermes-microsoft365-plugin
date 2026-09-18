from __future__ import annotations

import json


class RecordingContext:
    def __init__(self, config=None):
        self.config = config or {}
        self.tools = {}
        self.hooks = {}

    def get_config(self, key, default=None):
        return self.config.get(key, default)

    def register_tool(self, name, **kwargs):
        self.tools[name] = kwargs

    def register_hook(self, name, callback):
        self.hooks[name] = callback


def test_dispatch_tables_keep_profile_context_across_a_b_a(monkeypatch):
    from dataclasses import replace

    from microsoft365 import handlers, registration
    from microsoft365.contract import Settings

    seen = []

    def record_profile(args, context):
        del args
        seen.append(context.settings.client_id)
        return {"profile": context.settings.client_id}

    monkeypatch.setitem(
        handlers._REGISTRY,
        "outlook.search",
        replace(handlers._REGISTRY["outlook.search"], function=record_profile),
    )
    profile_a = Settings(tenant_id="tenant", client_id="profile-a")
    profile_b = Settings(tenant_id="tenant", client_id="profile-b")
    table_a = registration._dispatch_table(profile_a)
    table_b = registration._dispatch_table(profile_b)

    assert table_a["outlook.search"]({"action": "search"}) == {"profile": "profile-a"}
    assert table_b["outlook.search"]({"action": "search"}) == {"profile": "profile-b"}
    assert table_a["outlook.search"]({"action": "search"}) == {"profile": "profile-a"}
    assert seen == ["profile-a", "profile-b", "profile-a"]


def test_registration_does_not_resolve_or_capture_secret(monkeypatch):
    import agent.secret_scope
    from microsoft365 import register

    monkeypatch.setattr(
        agent.secret_scope,
        "get_secret",
        lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("secret read during registration")),
    )
    ctx = RecordingContext({"capabilities": {"outlook": {"search": True}}})

    register(ctx)

    assert "microsoft365_preflight" in ctx.tools
    assert "pre_tool_call" in ctx.hooks


def test_preflight_reads_secret_presence_through_secret_scope_only(monkeypatch):
    import agent.secret_scope
    from microsoft365.preflight import build_preflight
    from microsoft365.contract import Settings

    calls = []
    monkeypatch.setattr(agent.secret_scope, "get_secret", lambda name, default=None: calls.append(name) or "sentinel")
    result = build_preflight(
        Settings.from_mapping({
            "tenant_id": "tenant",
            "client_id": "client",
            "user_id": "user",
            "capabilities": {"outlook": {"search": True}},
        }),
        sdk_available=True,
    )

    assert calls == ["MICROSOFT365_CLIENT_SECRET"]
    assert result["secret_present"] is True
    assert "sentinel" not in repr(result)
    assert "client_secret" not in result["configuration"]
    assert result["remote_verification"] == "not_tested"


def test_client_creation_resolves_secret_only_at_execution(monkeypatch):
    import agent.secret_scope
    import azure.identity
    import msgraph
    from microsoft365.client import create_graph_client
    from microsoft365.contract import Settings

    calls = []
    credentials = []
    clients = []
    monkeypatch.setattr(agent.secret_scope, "get_secret", lambda name, default=None: calls.append(name) or "sentinel")
    monkeypatch.setattr(azure.identity, "ClientSecretCredential", lambda **kwargs: credentials.append(kwargs) or "credential")
    monkeypatch.setattr(msgraph, "GraphServiceClient", lambda **kwargs: clients.append(kwargs) or "client")

    result = create_graph_client(Settings(tenant_id="tenant", client_id="client"))

    assert result == "client"
    assert calls == ["MICROSOFT365_CLIENT_SECRET"]
    assert credentials == [{"tenant_id": "tenant", "client_id": "client", "client_secret": "sentinel"}]
    assert clients == [{"credentials": "credential", "scopes": ["https://graph.microsoft.com/.default"]}]


def test_unsupported_only_selection_is_retained_but_not_registered():
    from microsoft365 import register

    ctx = RecordingContext({
        "authentication_mode": "application",
        "capabilities": {"teams": {"send_messages": True}},
    })
    register(ctx)

    assert "microsoft365_teams" not in ctx.tools
    payload = json.loads(ctx.tools["microsoft365_preflight"]["handler"]({}))
    assert payload["operation_status"]["teams.send_messages"]["auth_status"] == "unsupported_auth_mode"
    assert payload["selected_operations"] == ["teams.send_messages"]


def test_malformed_configuration_registers_no_service_tools_and_preflight_reports_paths():
    from microsoft365 import register

    ctx = RecordingContext({"capabilities": {"outlook": {"send": "false"}}})
    register(ctx)

    assert set(ctx.tools) == {"microsoft365_preflight"}
    payload = json.loads(ctx.tools["microsoft365_preflight"]["handler"]({}))
    assert payload["locally_ready"] is False
    assert payload["configuration_errors"] == [
        "capabilities.outlook.send must be a boolean (got str)"
    ]
