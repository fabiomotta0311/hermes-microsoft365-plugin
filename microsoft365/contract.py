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

# --------------------------------------------------------------------------------------
# Planner container semantics (WP2)
# --------------------------------------------------------------------------------------


class PlannerContractError(ValueError):
    """A Planner call would have been sent with an incomplete or malformed container."""


#: Claim states this repository is allowed to attach to an endpoint permission.
#:
#: ``documented_not_verified`` is the only state reachable offline: the role is the one
#: recorded for that endpoint (inherited from the pre-existing permission matrix, or from
#: the page named in ``documentation_page``), and no endpoint page was re-read and no
#: tenant was contacted while this milestone was built (rule R10; remote verification is
#: WP16). ``verified`` is reserved for a claim recomputed against the endpoint reference
#: **and** confirmed on a test tenant.
PERMISSION_CLAIM_STATUSES = frozenset({"documented_not_verified", "verified"})


@dataclass(frozen=True)
class PlannerEndpoint:
    """One verified request target of a Planner operation, with the container it indexes.

    Container semantics, recorded decision (WP2): a Planner operation indexes by exactly
    one kind of container, and the identifiers that identify that container must be
    non-blank before anything is sent. ``path_identifiers`` are rendered into
    ``path_template``; ``body_identifiers`` are identifiers the endpoint reads from the
    request body instead -- Planner *creation* carries ``plan_id`` and ``bucket_id`` in
    the body (``planId``/``bucketId``), never in the path, while listing buckets and
    tasks and reading or updating a task carry them in the path.
    """

    operation: str
    method: str
    path_template: str
    container: str
    path_identifiers: tuple[str, ...] = ()
    body_identifiers: tuple[str, ...] = ()
    application_permissions: tuple[str, ...] = ()
    claim_status: str = "documented_not_verified"
    contract_call: str = ""
    contract_case: str = ""
    required_headers: tuple[str, ...] = ()
    documentation_page: str = ""

    @property
    def endpoint(self) -> str:
        """The endpoint this row claims its application role from."""
        return f"{self.method} {self.path_template}"

    @property
    def required_identifiers(self) -> tuple[str, ...]:
        """Identifiers this endpoint cannot be called without."""
        return self.path_identifiers + self.body_identifiers


#: The complete Planner endpoint table. ``planner.list_plans`` is the only operation here
#: the SDK exposes twice: without ``group_id`` it reads the tenant-wide route, with it the
#: group-owned route.
PLANNER_ENDPOINTS: tuple[PlannerEndpoint, ...] = (
    PlannerEndpoint(
        operation="list_plans",
        method="GET",
        path_template="/planner/plans",
        container="tenant_or_group",
        application_permissions=("Tasks.Read.All",),
        contract_call="build_collection_request_information",
        contract_case="planner_plans",
    ),
    PlannerEndpoint(
        operation="list_plans",
        method="GET",
        path_template="/groups/{group_id}/planner/plans",
        container="tenant_or_group",
        path_identifiers=("group_id",),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_collection_request_information",
        contract_case="planner_plans",
    ),
    PlannerEndpoint(
        operation="list_buckets",
        method="GET",
        path_template="/planner/plans/{plan_id}/buckets",
        container="plan",
        path_identifiers=("plan_id",),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_collection_request_information",
        contract_case="planner_buckets",
    ),
    PlannerEndpoint(
        operation="list_tasks",
        method="GET",
        path_template="/planner/plans/{plan_id}/tasks",
        container="plan",
        path_identifiers=("plan_id",),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_collection_request_information",
        contract_case="planner_tasks",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks",
    ),
    PlannerEndpoint(
        operation="read",
        method="GET",
        path_template="/planner/tasks/{planner_task_id}",
        container="task",
        path_identifiers=("planner_task_id",),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_item_request_information",
        contract_case="planner_read",
    ),
    PlannerEndpoint(
        operation="create_tasks",
        method="POST",
        path_template="/planner/tasks",
        container="plan_and_bucket",
        body_identifiers=("plan_id", "bucket_id"),
        application_permissions=("Tasks.ReadWrite.All",),
        contract_call="build_write_request_information",
        contract_case="planner_create_tasks",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/planner-post-tasks",
    ),
    PlannerEndpoint(
        operation="update_tasks",
        method="PATCH",
        path_template="/planner/tasks/{planner_task_id}",
        container="task",
        path_identifiers=("planner_task_id",),
        application_permissions=("Tasks.ReadWrite.All",),
        contract_call="build_write_request_information",
        contract_case="planner_update_tasks",
        required_headers=("If-Match",),
        documentation_page="https://learn.microsoft.com/en-us/graph/api/plannertask-update",
    ),
)

PLANNER_OPERATIONS: tuple[str, ...] = (
    "planner.list_plans",
    "planner.list_buckets",
    "planner.list_tasks",
    "planner.read",
    "planner.create_tasks",
    "planner.update_tasks",
)


def planner_endpoints(key: str) -> tuple[PlannerEndpoint, ...]:
    """Every verified endpoint of a Planner operation, in declaration order."""
    if key not in PLANNER_OPERATIONS:
        raise PlannerContractError(f"{key} is not a declared Planner operation")
    return tuple(
        endpoint for endpoint in PLANNER_ENDPOINTS if f"planner.{endpoint.operation}" == key
    )


def _planner_container(endpoints: tuple[PlannerEndpoint, ...]) -> str:
    containers = tuple(dict.fromkeys(endpoint.container for endpoint in endpoints))
    return containers[0] if len(containers) == 1 else "+".join(containers)


def required_planner_identifiers(key: str) -> tuple[str, ...]:
    """Identifiers a call must supply whichever endpoint of the operation it uses."""
    endpoints = planner_endpoints(key)
    per_endpoint = [set(endpoint.required_identifiers) for endpoint in endpoints]
    shared = set.intersection(*per_endpoint) if per_endpoint else set()
    ordered = dict.fromkeys(
        name for endpoint in endpoints for name in endpoint.required_identifiers
    )
    return tuple(name for name in ordered if name in shared)


def optional_planner_identifiers(key: str) -> tuple[str, ...]:
    """Identifiers an endpoint accepts or requires without the operation requiring them.

    ``group_id`` is the only one today: supplying it selects the group-owned plan route,
    omitting it (or passing ``None``) selects the tenant-wide route.
    """
    endpoints = planner_endpoints(key)
    required = set(required_planner_identifiers(key))
    ordered = dict.fromkeys(
        name for endpoint in endpoints for name in endpoint.required_identifiers
    )
    return tuple(name for name in ordered if name not in required)


def _planner_identifier_is_missing(
    key: str, name: str, endpoints: tuple[PlannerEndpoint, ...]
) -> str:
    where = "; ".join(sorted({endpoint.endpoint for endpoint in endpoints}))
    return f"{name} is required for {key} ({where})"


def _planner_identifier_value(
    key: str,
    name: str,
    arguments: Mapping[str, Any],
    *,
    required: bool,
    endpoints: tuple[PlannerEndpoint, ...],
) -> str | None:
    value = arguments.get(name)
    if value is None:
        if required:
            raise PlannerContractError(_planner_identifier_is_missing(key, name, endpoints))
        return None
    if not isinstance(value, str):
        raise PlannerContractError(f"{name} for {key} must be a non-blank string")
    normalized = value.strip()
    if normalized:
        return normalized
    if required:
        raise PlannerContractError(_planner_identifier_is_missing(key, name, endpoints))
    raise PlannerContractError(f"{name} for {key} must be a non-blank string or be omitted")


def validate_planner_identifiers(key: str, arguments: Mapping[str, Any]) -> dict[str, str]:
    """Normalize and validate the container identifiers a Planner call must carry.

    Returns only the identifiers that will be sent, so a blank or absent required
    identifier -- and a blank optional one -- is refused here, with the endpoint it
    belongs to, instead of being rendered into a request against the wrong container.
    ``None`` for an identifier the operation does not require means "not supplied": for
    ``planner.list_plans`` that selects the tenant-wide route, while a blank string is an
    error and never a silent fallback to it.
    """
    endpoints = planner_endpoints(key)
    if not isinstance(arguments, Mapping):
        raise PlannerContractError(f"{key} arguments must be a mapping")
    resolved: dict[str, str] = {}
    for name in required_planner_identifiers(key):
        resolved[name] = _planner_identifier_value(
            key, name, arguments, required=True, endpoints=endpoints
        )
    for name in optional_planner_identifiers(key):
        value = _planner_identifier_value(
            key, name, arguments, required=False, endpoints=endpoints
        )
        if value is not None:
            resolved[name] = value
    return resolved


def _planner_permissions(key: str) -> tuple[str, ...]:
    """Least-privileged application roles claimed by the operation's own endpoints."""
    return tuple(
        dict.fromkeys(
            role for endpoint in planner_endpoints(key) for role in endpoint.application_permissions
        )
    )


#: Permission claims for the services whose endpoint contract the registry does not declare
#: yet. Planner rows are deliberately absent: each Planner operation derives its claim from
#: the endpoint table above, so a blanket per-service claim cannot drift from the endpoint it
#: belongs to.
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
}
_UNSUPPORTED_APPLICATION = frozenset({"teams.search_messages", "teams.send_messages"})
_NOT_VERIFIED_APPLICATION = frozenset({"todo.create_tasks", "todo.update_tasks"})

#: What each implementation status is allowed to mean. ``contract_verified`` requires the
#: registry to declare, for every endpoint of the operation, the ``sdk_contract`` dispatcher
#: and case that a strict offline test pins (``tests/test_planner_contracts.py`` does this for
#: Planner). ``implemented`` additionally requires a registered handler.
IMPLEMENTATION_STATUS_LABELS = {
    "contract_foundation":
        "Listed with a permission claim. The registry does not declare this operation's "
        "verified endpoint, container or identifier contract yet, so nothing here asserts "
        "that its request shape is correct.",
    "contract_verified":
        "Every endpoint the operation uses is pinned by a strict offline request-contract "
        "test. No handler exists, so the operation is not executable.",
    "implemented":
        "A handler and a strict offline request-contract test both exist and the operation "
        "is executable.",
}


@dataclass(frozen=True)
class OperationDefinition:
    service: str
    operation: str
    permissions: tuple[str, ...]
    write: bool
    endpoints: tuple[str, ...] = ()
    container: str = ""
    required_identifiers: tuple[str, ...] = ()
    optional_identifiers: tuple[str, ...] = ()
    required_headers: tuple[str, ...] = ()
    contract_cases: tuple[tuple[str, str], ...] = ()
    permission_claim_statuses: tuple[str, ...] = ()
    documentation_pages: tuple[str, ...] = ()
    implementation_status: str = "contract_foundation"
    executable: bool = False


def _planner_definition(key: str) -> OperationDefinition:
    operation = key.split(".", 1)[1]
    endpoints = planner_endpoints(key)
    pinned = all(endpoint.contract_call and endpoint.contract_case for endpoint in endpoints)
    return OperationDefinition(
        service="planner",
        operation=operation,
        permissions=_planner_permissions(key),
        write=operation in WRITE_OPERATIONS,
        endpoints=tuple(endpoint.endpoint for endpoint in endpoints),
        container=_planner_container(endpoints),
        required_identifiers=required_planner_identifiers(key),
        optional_identifiers=optional_planner_identifiers(key),
        required_headers=tuple(
            dict.fromkeys(name for endpoint in endpoints for name in endpoint.required_headers)
        ),
        contract_cases=tuple(
            (endpoint.contract_call, endpoint.contract_case) for endpoint in endpoints
        ),
        permission_claim_statuses=tuple(
            dict.fromkeys(endpoint.claim_status for endpoint in endpoints)
        ),
        documentation_pages=tuple(
            dict.fromkeys(
                endpoint.documentation_page for endpoint in endpoints if endpoint.documentation_page
            )
        ),
        implementation_status="contract_verified" if pinned else "contract_foundation",
    )


OPERATION_REGISTRY = {
    **{
        key: OperationDefinition(
            service=key.split(".", 1)[0],
            operation=key.split(".", 1)[1],
            permissions=permissions,
            write=key.split(".", 1)[1] in WRITE_OPERATIONS,
        )
        for key, permissions in _PERMISSION_MAP.items()
    },
    **{key: _planner_definition(key) for key in PLANNER_OPERATIONS},
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
            if service_value is False:
                normalized[service] = {operation: False for operation in OPERATIONS[service]}
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
