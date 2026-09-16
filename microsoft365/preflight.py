"""Side-effect-free local preflight reporting."""
from __future__ import annotations

from dataclasses import asdict
from typing import Iterable

from .contract import OPERATIONS, OPERATION_REGISTRY, Settings, operation_status

SECRET_ENV_NAME = "MICROSOFT365_CLIENT_SECRET"  # pragma: allowlist secret


def _secret_presence() -> tuple[bool, str | None]:
    try:
        from agent.secret_scope import get_secret

        return bool(get_secret(SECRET_ENV_NAME, None)), None
    except Exception as exc:
        # Unscoped access and backend failures fail closed; never return details
        # that may contain provider-specific secret material.
        return False, type(exc).__name__


def build_preflight(
    settings: Settings | None,
    *,
    sdk_available: bool | None = None,
    configuration_errors: Iterable[str] = (),
) -> dict:
    errors = list(configuration_errors)
    if sdk_available is None:
        try:
            import azure.identity  # noqa: F401
            import msgraph  # noqa: F401
            sdk_available = True
        except ImportError:
            sdk_available = False
    secret_present, secret_error = _secret_presence()
    if settings is None:
        return {
            "locally_ready": False,
            "configuration_errors": errors,
            "secret_present": secret_present,
            "secret_scope_error": secret_error,
            "remote_verification": "not_tested",
            "selected_operations": [],
            "operation_status": {},
            "configuration": {},
        }

    selected = sorted(
        f"{service}.{operation}"
        for service in OPERATIONS
        for operation in settings.selected(service)
    )
    statuses = {}
    for key in selected:
        service, operation = key.split(".", 1)
        status = operation_status(settings.authentication_mode, service, operation)
        statuses[key] = {
            **asdict(status),
            "permissions": list(OPERATION_REGISTRY[key].permissions),
            "write": OPERATION_REGISTRY[key].write,
        }
    missing = [
        name
        for name, present in (
            ("tenant_id", bool(settings.tenant_id.strip())),
            ("client_id", bool(settings.client_id.strip())),
            (SECRET_ENV_NAME, secret_present),
            ("capabilities", bool(selected)),
        )
        if not present
    ]
    if any(service != "sharepoint" for service, _ in (item.split(".", 1) for item in selected)) and not settings.user_id.strip():
        missing.append("user_id")
    all_executable = bool(selected) and all(item["executable"] for item in statuses.values())
    return {
        "locally_ready": bool(sdk_available and not errors and not missing and all_executable),
        "sdk_available": bool(sdk_available),
        "configuration_errors": errors,
        "missing": sorted(set(missing)),
        "secret_present": secret_present,
        "secret_scope_error": secret_error,
        "authentication_mode": settings.authentication_mode,
        "remote_verification": "not_tested",
        "selected_operations": selected,
        "operation_status": statuses,
        "configuration": {
            "tenant_id": settings.tenant_id,
            "client_id": settings.client_id,
            "user_id": settings.user_id,
            "authentication_mode": settings.authentication_mode,
            "capabilities": {key: dict(value) for key, value in settings.capabilities.items()},
        },
    }
