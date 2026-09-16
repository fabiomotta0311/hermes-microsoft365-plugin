"""Conditional schemas and native Hermes registration."""
from __future__ import annotations

import inspect
import json
from typing import Any, Callable

from .contract import ConfigurationError, OPERATIONS, Settings, WRITE_OPERATIONS, operation_status
from .execution import ExecutionError, run_async
from .preflight import build_preflight
from .validation import check as check_operation_arguments

#: Operation handlers keyed ``"<service>.<operation>"``, populated by the handler work
#: packages (WP6-WP9). Every entry is executed through :func:`invoke_handler`, which is
#: the only place this plugin crosses the sync/async execution seam.
HANDLER_TABLE: dict[str, Callable[[dict], Any]] = {}


def invoke_handler(handler, args):
    """Run one registered handler, awaiting async handlers through the execution seam."""
    result = handler(args)
    if inspect.isawaitable(result):
        return run_async(result)
    return result


def unavailable_handler(args, *, service: str) -> str:
    """Fallback for every operation that has no handler yet."""
    del args
    return json.dumps({"error": "operation_not_implemented", "service": service})


def service_tool_handler(service: str, args):
    """Dispatch one service tool call: registered handler, else unavailable."""
    operation = args.get("action") if isinstance(args, dict) else None
    handler = HANDLER_TABLE.get(f"{service}.{operation}") if isinstance(operation, str) else None
    if handler is None:
        return unavailable_handler(args, service=service)
    try:
        return invoke_handler(handler, args)
    except ExecutionError as exc:
        return json.dumps({"error": exc.category, "message": str(exc)})


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
        # WP13: the operation-specific validator runs before anything else this hook can do --
        # before the approval directive, and before any handler could be reached. It is the
        # same function and the same payload the dispatch path re-checks (defence in depth),
        # so an invalid, disabled or unsupported call can never reach a secret, a credential
        # or a Graph client, and can never be approved.
        rejection = check_operation_arguments(settings, service=service, arguments=args)
        if rejection is not None:
            return {"action": "block", "message": rejection.message}
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

        def service_handler(args, _service=service, **kwargs):
            del kwargs
            # WP13 defence in depth: the dispatch path validates the very payload it is about
            # to execute with the same validator the hook used, so a call that did not pass the
            # hook (another dispatch path, a host that skips hooks, a payload changed after
            # approval) is refused here instead of reaching a handler, a secret or a client.
            rejection = check_operation_arguments(settings, service=_service, arguments=args)
            if rejection is not None:
                return json.dumps(rejection.to_payload())
            return service_tool_handler(_service, args)

        ctx.register_tool(
            name=f"microsoft365_{service}",
            toolset="microsoft365",
            schema=schema,
            handler=service_handler,
            check_fn=lambda: True,
            is_async=False,
            emoji="📎",
        )
