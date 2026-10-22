"""The host contract at the plugin's dispatch boundary.

A tool handler's result crosses into the host, and the host accepts only a
``str`` or the multimodal envelope (``tools/registry.py::_normalize_handler_result``).
Anything else becomes a ``tool_result_contract`` error -- so a plugin that returns a raw
``dict`` on its success path fails *at the host boundary*, after the work already
succeeded.

These tests dispatch through the **live host registry**, using the plugin's **own
registered tool handler**. That is the point: the plugin's offline tests call the
handler directly, so a result the host would reject was never handed to the host. A
test that only inspects the plugin's own return value cannot see this class of defect.
"""
from __future__ import annotations

import json


def _host_registry():
    """The host's live tool registry -- the object that decides what a result may be."""
    from tools.registry import registry

    return registry


class _CapturingContext:
    """A plugin context that keeps the tools the plugin registers.

    This is the plugin's real registration path (``microsoft365.registration.register_plugin``),
    so the handler the host calls is the plugin's handler, not a hand-written stand-in.
    """

    def __init__(self, config: dict | None = None):
        self.config = config or {}
        self.tools: dict[str, dict] = {}

    def get_config(self, key, default=None):
        return self.config.get(key, default)

    def register_tool(self, name, **kwargs):
        self.tools[name] = kwargs

    def register_hook(self, name, callback):
        pass


def _plugin_service_tool(settings_config: dict, service: str):
    """Register the plugin and return the host-facing handler of one of its tools."""
    from microsoft365.registration import register_plugin

    context = _CapturingContext(settings_config)
    register_plugin(context)

    tool = context.tools[f"microsoft365_{service}"]
    return tool["handler"], tool


def _dispatch_through_host(name: str, handler, args: dict):
    """Register a handler on the live host registry, dispatch it, and clean up."""
    registry = _host_registry()
    registry.register(
        name,
        toolset="microsoft365",
        schema={
            "name": name,
            "description": "probe",
            "parameters": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        handler=handler,
        check_fn=lambda: True,
        is_async=False,
    )
    try:
        return registry.dispatch(name, args)
    finally:
        registry.deregister(name)


def test_the_host_rejects_a_raw_dict_result_so_the_plugin_must_serialize():
    """RED control: the host really does reject a bare ``dict``.

    Without this, the fix below could pass for the wrong reason -- e.g. if the host
    accepted dicts, serializing would be cosmetic rather than required. This pins the
    host's contract from the host's own code, so the plugin's boundary fix is provably
    load-bearing.
    """
    def raw_dict_handler(args):
        del args
        return {"status": "succeeded", "result": [], "continuation": {}}

    result = _dispatch_through_host(
        "microsoft365_probe_raw_dict", raw_dict_handler, {}
    )

    decoded = json.loads(result) if isinstance(result, str) else result
    assert decoded.get("error_type") == "tool_result_contract"
    assert decoded.get("result_type") == "dict"


def test_a_successful_service_call_reaches_the_host_as_a_string(monkeypatch):
    """A successful ``outlook.search`` must reach the host as a string.

    This is the regression the live report describes: every *error* path worked, and the
    first *successful* call returned ``tool_result_contract`` -- because the plugin
    serialized its error paths and returned the success payload as a raw ``dict``.

    The handler here is the plugin's own registered tool handler, dispatched through the
    host's registry, so the whole chain a live call takes is exercised.
    """
    from microsoft365 import registration

    payload = {
        "operation": "outlook.search",
        "status": "succeeded",
        "result": {"items": []},
        "continuation": {
            "truncated": False,
            "next_link_present": False,
            "pages_fetched": 1,
            "items_fetched": 0,
            "stop_reason": "complete",
        },
    }

    def handler(args):
        del args
        return payload

    monkeypatch.setitem(registration.HANDLER_TABLE, "outlook.search", handler)

    plugin_handler, _tool = _plugin_service_tool(
        {
            "tenant_id": "tenant",
            "client_id": "client",
            "user_id": "user@example.com",
            "capabilities": {"outlook": {"search": True}},
        },
        "outlook",
    )

    result = _dispatch_through_host(
        "microsoft365_outlook_probe",
        plugin_handler,
        {"action": "search", "user_id": "user@example.com"},
    )

    assert isinstance(result, str), (
        "the host accepts only a string or the multimodal envelope; a raw dict becomes "
        "tool_result_contract even when the call succeeded"
    )
    decoded = json.loads(result)
    assert decoded["status"] == "succeeded"
    assert decoded["result"] == {"items": []}
    assert "continuation" in decoded


def test_the_plugin_dispatch_boundary_serializes_a_non_string_success_payload(monkeypatch):
    """The plugin's own boundary returns a string for every dispatch outcome.

    ``service_tool_handler`` is the single place every service tool returns through, so
    the serialization lives there: no individual handler can reintroduce the mismatch,
    and a handler added later is covered by construction.
    """
    from microsoft365 import registration

    def handler(args):
        del args
        return {"status": "succeeded", "result": {"items": []}}

    monkeypatch.setitem(registration.HANDLER_TABLE, "outlook.search", handler)

    outcome = registration.service_tool_handler("outlook", {"action": "search"})

    assert isinstance(outcome, str)
    assert json.loads(outcome)["status"] == "succeeded"


def test_the_plugin_dispatch_boundary_keeps_a_string_payload_untouched(monkeypatch):
    """A handler that already returns a string is not double-encoded.

    The boundary serializes only what is not already a string; a payload that is already
    JSON text must pass through byte for byte.
    """
    from microsoft365 import registration

    body = '{"status": "succeeded"}'

    def handler(args):
        del args
        return body

    monkeypatch.setitem(registration.HANDLER_TABLE, "outlook.search", handler)

    outcome = registration.service_tool_handler("outlook", {"action": "search"})

    assert outcome == body


def test_the_plugin_dispatch_boundary_survives_a_payload_json_cannot_represent(monkeypatch):
    """A value the JSON encoder cannot represent becomes a readable result, not a crash.

    ``default=str`` is what makes this hold: a payload can legitimately carry a
    ``datetime``, a ``Path`` or an SDK object, and the failure mode we want at the
    boundary is a readable result rather than a ``TypeError`` escaping into the host.
    """
    from microsoft365 import registration

    class Opaque:
        def __str__(self):
            return "opaque-value"

    def handler(args):
        del args
        return {"status": "succeeded", "value": Opaque()}

    monkeypatch.setitem(registration.HANDLER_TABLE, "outlook.search", handler)

    outcome = registration.service_tool_handler("outlook", {"action": "search"})

    assert isinstance(outcome, str)
    assert json.loads(outcome)["value"] == "opaque-value"


def test_the_plugin_dispatch_boundary_preserves_the_multimodal_envelope(monkeypatch):
    """The multimodal envelope is passed through, not re-encoded as JSON.

    The host reads the envelope structurally (``_multimodal`` plus a ``content`` list).
    Serializing it would turn a structured result into an opaque string and the host
    would lose the content it was meant to render.
    """
    from microsoft365 import registration

    envelope = {
        "_multimodal": True,
        "content": [{"type": "text", "text": "hello"}],
    }

    def handler(args):
        del args
        return envelope

    monkeypatch.setitem(registration.HANDLER_TABLE, "outlook.search", handler)

    outcome = registration.service_tool_handler("outlook", {"action": "search"})

    assert outcome is envelope
