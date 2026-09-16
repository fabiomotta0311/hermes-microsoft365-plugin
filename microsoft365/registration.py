"""Conditional schemas and native Hermes registration."""
from __future__ import annotations

import json

from .contract import ConfigurationError, OPERATIONS, Settings, WRITE_OPERATIONS, operation_status
from .preflight import build_preflight


def active_actions(settings: Settings, service: str) -> tuple[str, ...]:
    return tuple(
        operation
        for operation in OPERATIONS[service]
        if operation in settings.selected(service)
        and operation_status(settings.authentication_mode, service, operation).executable
    )


def schema_for(service: str, actions: tuple[str, ...]):
    if not actions:
        return None
    return {
        "name": f"microsoft365_{service}",
        "description": f"Microsoft 365 {service} operations with explicit host approval for writes.",
        "parameters": {
            "type": "object",
            "properties": {"action": {"type": "string", "enum": list(actions)}},
            "required": ["action"],
            "additionalProperties": True,
        },
    }


def _read_settings(ctx):
    raw = {
        key: ctx.get_config(key, default)
        for key, default in (
            ("tenant_id", ""),
            ("client_id", ""),
            ("user_id", ""),
            ("authentication_mode", "application"),
            ("capabilities", {}),
        )
    }
    try:
        return Settings.from_mapping(raw), []
    except ConfigurationError as exc:
        return None, [str(exc)]


def register_plugin(ctx) -> None:
    settings, errors = _read_settings(ctx)

    def preflight_handler(args, **kwargs):
        del args, kwargs
        return json.dumps(build_preflight(settings, configuration_errors=errors), default=str)

    ctx.register_tool(
        name="microsoft365_preflight",
        toolset="microsoft365",
        schema={
            "name": "microsoft365_preflight",
            "description": "Report local Microsoft 365 configuration, operation, and SDK status without contacting Graph.",
            "parameters": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        handler=preflight_handler,
        check_fn=lambda: True,
        is_async=False,
        emoji="🧩",
    )

    def pre_tool_call(*, tool_name="", args=None, **kwargs):
        del kwargs
        prefix = "microsoft365_"
        if settings is None or not isinstance(tool_name, str) or not tool_name.startswith(prefix):
            return None
        service = tool_name[len(prefix):]
        if service not in OPERATIONS:
            return None
        if not isinstance(args, dict):
            return {"action": "block", "message": "Microsoft 365 arguments must be an object"}
        operation = args.get("action")
        if operation not in active_actions(settings, service):
            return {"action": "block", "message": f"Microsoft 365 operation is not executable: {service}.{operation}"}
        if operation in WRITE_OPERATIONS:
            return {
                "action": "approve",
                "message": f"Microsoft 365 {operation}: external side effect",
                "rule_key": f"microsoft365.{service}.{operation}",
            }
        return None

    ctx.register_hook("pre_tool_call", pre_tool_call)
    if settings is None:
        return
    for service in OPERATIONS:
        actions = active_actions(settings, service)
        schema = schema_for(service, actions)
        if schema is None:
            continue

        def unavailable_handler(args, _service=service, **kwargs):
            del args, kwargs
            return json.dumps({"error": "operation_not_implemented", "service": _service})

        ctx.register_tool(
            name=f"microsoft365_{service}",
            toolset="microsoft365",
            schema=schema,
            handler=unavailable_handler,
            check_fn=lambda: True,
            is_async=False,
            emoji="📎",
        )
