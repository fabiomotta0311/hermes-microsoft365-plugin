"""Complete administrative operation and configuration contract."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Mapping

OPERATIONS = {
    "outlook": ("search", "read", "create_draft", "send"),
    "sharepoint": ("search", "read", "download_files", "upload_files"),
    "onedrive": ("search", "read", "download_files", "upload_files"),
    "calendar": ("search", "create_events", "update_events"),
    "teams": ("list_teams", "list_channels", "search_messages", "send_messages"),
    "todo": ("list_task_lists", "search", "read", "create_tasks", "update_tasks"),
    "planner": ("list_plans", "list_buckets", "list_tasks", "read", "create_tasks", "update_tasks"),
}
WRITE_OPERATIONS = frozenset(
    {
        "create_draft", "send", "upload_files", "create_events", "update_events",
        "send_messages", "create_tasks", "update_tasks",
    }
)

_PERMISSION_MAP = {
    "outlook.search": ("Mail.Read",), "outlook.read": ("Mail.Read",),
    "outlook.create_draft": ("Mail.ReadWrite",), "outlook.send": ("Mail.Send",),
    "sharepoint.search": ("Sites.Read.All",), "sharepoint.read": ("Sites.Read.All",),
    "sharepoint.download_files": ("Files.Read.All",), "sharepoint.upload_files": ("Files.ReadWrite.All",),
    "onedrive.search": ("Files.Read.All",), "onedrive.read": ("Files.Read.All",),
    "onedrive.download_files": ("Files.Read.All",), "onedrive.upload_files": ("Files.ReadWrite.All",),
    "calendar.search": ("Calendars.Read",), "calendar.create_events": ("Calendars.ReadWrite",),
    "calendar.update_events": ("Calendars.ReadWrite",),
    "teams.list_teams": ("Team.ReadBasic.All",), "teams.list_channels": ("Channel.ReadBasic.All",),
    "teams.search_messages": (), "teams.send_messages": (),
    "todo.list_task_lists": ("Tasks.Read.All",), "todo.search": ("Tasks.Read.All",),
    "todo.read": ("Tasks.Read.All",), "todo.create_tasks": ("Tasks.ReadWrite.All",),
    "todo.update_tasks": ("Tasks.ReadWrite.All",),
    "planner.list_plans": ("Tasks.Read.All",), "planner.list_buckets": ("Tasks.Read.All",),
    "planner.list_tasks": ("Tasks.Read.All",), "planner.read": ("Tasks.Read.All",),
    "planner.create_tasks": ("Tasks.ReadWrite.All",), "planner.update_tasks": ("Tasks.ReadWrite.All",),
}
_UNSUPPORTED_APPLICATION = frozenset({"teams.search_messages", "teams.send_messages"})
_NOT_VERIFIED_APPLICATION = frozenset({"todo.create_tasks", "todo.update_tasks"})


@dataclass(frozen=True)
class OperationDefinition:
    service: str
    operation: str
    permissions: tuple[str, ...]
    write: bool
    implementation_status: str = "contract_foundation"
    executable: bool = False


OPERATION_REGISTRY = {
    key: OperationDefinition(
        service=key.split(".", 1)[0],
        operation=key.split(".", 1)[1],
        permissions=permissions,
        write=key.split(".", 1)[1] in WRITE_OPERATIONS,
    )
    for key, permissions in _PERMISSION_MAP.items()
}


@dataclass(frozen=True)
class OperationStatus:
    auth_status: str
    implementation_status: str
    executable: bool
    reason: str


class ConfigurationError(ValueError):
    pass


def operation_status(auth_mode: str, service: str, operation: str) -> OperationStatus:
    key = f"{service}.{operation}"
    definition = OPERATION_REGISTRY.get(key)
    if definition is None:
        return OperationStatus("unknown_operation", "unknown", False, f"Unknown operation: {key}")
    if auth_mode != "application":
        return OperationStatus("not_implemented", definition.implementation_status, False, "Delegated authentication is not implemented")
    if key in _UNSUPPORTED_APPLICATION:
        return OperationStatus("unsupported_auth_mode", definition.implementation_status, False, "Operation requires delegated authentication")
    if key in _NOT_VERIFIED_APPLICATION:
        return OperationStatus("not_verified", definition.implementation_status, False, "Application permission support is not yet verified")
    return OperationStatus(
        "supported",
        definition.implementation_status,
        definition.executable,
        "Authentication mode is supported; executable contract is reported separately",
    )


@dataclass(frozen=True)
class Settings:
    tenant_id: str = ""
    client_id: str = ""
    user_id: str = ""
    authentication_mode: str = "application"
    capabilities: Mapping[str, Mapping[str, bool]] = field(default_factory=dict)

    @classmethod
    def from_mapping(cls, raw: Mapping[str, Any] | None) -> "Settings":
        raw = raw or {}
        capabilities = raw.get("capabilities", {})
        if not isinstance(capabilities, Mapping):
            raise ConfigurationError(f"capabilities must be a mapping (got {type(capabilities).__name__})")
        unknown_services = sorted(set(capabilities) - set(OPERATIONS))
        if unknown_services:
            raise ConfigurationError(f"capabilities.{unknown_services[0]} is an unknown capability")
        normalized: dict[str, dict[str, bool]] = {}
        for service, service_value in capabilities.items():
            if service_value is True:
                normalized[service] = {operation: True for operation in OPERATIONS[service]}
                continue
            if not isinstance(service_value, Mapping):
                raise ConfigurationError(
                    f"capabilities.{service} must be a boolean or operation mapping "
                    f"(got {type(service_value).__name__})"
                )
            unknown_operations = sorted(set(service_value) - set(OPERATIONS[service]))
            if unknown_operations:
                raise ConfigurationError(
                    f"capabilities.{service}.{unknown_operations[0]} is an unknown operation"
                )
            normalized[service] = {}
            for operation, enabled in service_value.items():
                if type(enabled) is not bool:
                    raise ConfigurationError(
                        f"capabilities.{service}.{operation} must be a boolean "
                        f"(got {type(enabled).__name__})"
                    )
                normalized[service][operation] = enabled
        mode = raw.get("authentication_mode", "application")
        if mode not in {"application", "delegated"}:
            raise ConfigurationError("authentication_mode must be 'application' or 'delegated'")
        return cls(
            tenant_id=str(raw.get("tenant_id") or ""),
            client_id=str(raw.get("client_id") or ""),
            user_id=str(raw.get("user_id") or ""),
            authentication_mode=str(mode),
            capabilities=normalized,
        )

    def selected(self, service: str) -> set[str]:
        return {name for name, enabled in self.capabilities.get(service, {}).items() if enabled}
