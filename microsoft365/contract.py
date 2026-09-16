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


# --------------------------------------------------------------------------------------
# Shared endpoint-table helpers
# --------------------------------------------------------------------------------------
#
# Every service whose endpoint contract is pinned declares a table of typed endpoint rows
# and derives its registry entry (permissions, container, identifiers, required headers,
# contract cases, claim statuses) from that table. The derivation is shared here so a
# service cannot quietly invent a different notion of "container" or "claim".


def _endpoint_container(endpoints) -> str:
    """The single container of an endpoint set, or the joined union when it has several."""
    containers = tuple(dict.fromkeys(endpoint.container for endpoint in endpoints))
    return containers[0] if len(containers) == 1 else "+".join(containers)


def _endpoint_permissions(endpoints) -> tuple[str, ...]:
    """Least-privileged application roles claimed by the endpoints themselves."""
    return tuple(
        dict.fromkeys(role for endpoint in endpoints for role in endpoint.application_permissions)
    )


def _required_endpoint_identifiers(endpoints) -> tuple[str, ...]:
    """Identifiers a call must supply whichever endpoint of the operation it uses."""
    per_endpoint = [set(endpoint.required_identifiers) for endpoint in endpoints]
    shared = set.intersection(*per_endpoint) if per_endpoint else set()
    ordered = dict.fromkeys(
        name for endpoint in endpoints for name in endpoint.required_identifiers
    )
    return tuple(name for name in ordered if name in shared)


def _optional_endpoint_identifiers(endpoints) -> tuple[str, ...]:
    """Identifiers an endpoint accepts or requires without the operation requiring them."""
    required = set(_required_endpoint_identifiers(endpoints))
    ordered = dict.fromkeys(
        name for endpoint in endpoints for name in endpoint.required_identifiers
    )
    return tuple(name for name in ordered if name not in required)


def _missing_identifier_message(key: str, name: str, endpoints) -> str:
    where = "; ".join(sorted({endpoint.endpoint for endpoint in endpoints}))
    return f"{name} is required for {key} ({where})"


def _resolve_endpoint_identifiers(*, error, key, arguments, endpoints) -> dict[str, str]:
    """Normalize and validate the identifiers one endpoint set indexes by.

    Returns only the identifiers that will be sent, so a blank or absent required
    identifier -- and a blank optional one -- is refused with the endpoint it belongs to
    instead of being rendered into a request against the wrong container. ``None`` for an
    identifier the operation does not require means "not supplied" and selects the
    endpoint that has no such identifier; a blank string is an error and never a silent
    fallback to it. The caller supplies the typed error so each service keeps its own.
    """
    if not isinstance(arguments, Mapping):
        raise error(f"{key} arguments must be a mapping")
    resolved: dict[str, str] = {}

    def value(name: str, *, required: bool) -> str | None:
        raw = arguments.get(name)
        if raw is None:
            if required:
                raise error(_missing_identifier_message(key, name, endpoints))
            return None
        if not isinstance(raw, str):
            raise error(f"{name} for {key} must be a non-blank string")
        normalized = raw.strip()
        if normalized:
            return normalized
        if required:
            raise error(_missing_identifier_message(key, name, endpoints))
        raise error(f"{name} for {key} must be a non-blank string or be omitted")

    for name in _required_endpoint_identifiers(endpoints):
        resolved[name] = value(name, required=True)
    for name in _optional_endpoint_identifiers(endpoints):
        optional_value = value(name, required=False)
        if optional_value is not None:
            resolved[name] = optional_value
    return resolved


def _endpoint_backed_definition(*, service: str, key: str, endpoints) -> OperationDefinition:
    """Registry entry derived from a service's own endpoint table.

    ``permissions``, ``container``, the identifier sets, the required headers, the contract
    cases and the claim statuses are all read from the endpoints, so a claim cannot drift
    from the endpoint it belongs to. ``contract_verified`` requires every endpoint to name
    the ``sdk_contract`` dispatcher and case a strict offline test pins.
    """
    operation = key.split(".", 1)[1]
    pinned = all(endpoint.contract_call and endpoint.contract_case for endpoint in endpoints)
    return OperationDefinition(
        service=service,
        operation=operation,
        permissions=_endpoint_permissions(endpoints),
        write=operation in WRITE_OPERATIONS,
        endpoints=tuple(endpoint.endpoint for endpoint in endpoints),
        container=_endpoint_container(endpoints),
        required_identifiers=_required_endpoint_identifiers(endpoints),
        optional_identifiers=_optional_endpoint_identifiers(endpoints),
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


def required_planner_identifiers(key: str) -> tuple[str, ...]:
    """Identifiers a call must supply whichever endpoint of the operation it uses."""
    return _required_endpoint_identifiers(planner_endpoints(key))


def optional_planner_identifiers(key: str) -> tuple[str, ...]:
    """Identifiers an endpoint accepts or requires without the operation requiring them.

    ``group_id`` is the only one today: supplying it selects the group-owned plan route,
    omitting it (or passing ``None``) selects the tenant-wide route.
    """
    return _optional_endpoint_identifiers(planner_endpoints(key))


def validate_planner_identifiers(key: str, arguments: Mapping[str, Any]) -> dict[str, str]:
    """Normalize and validate the container identifiers a Planner call must carry.

    Returns only the identifiers that will be sent, so a blank or absent required
    identifier -- and a blank optional one -- is refused here, with the endpoint it
    belongs to, instead of being rendered into a request against the wrong container.
    ``None`` for an identifier the operation does not require means "not supplied": for
    ``planner.list_plans`` that selects the tenant-wide route, while a blank string is an
    error and never a silent fallback to it.
    """
    return _resolve_endpoint_identifiers(
        error=PlannerContractError,
        key=key,
        arguments=arguments,
        endpoints=planner_endpoints(key),
    )


# --------------------------------------------------------------------------------------
# Microsoft To Do semantics (WP3)
# --------------------------------------------------------------------------------------


class TodoContractError(ValueError):
    """A To Do call would have carried a field or filter the endpoint does not accept."""


@dataclass(frozen=True)
class TodoEndpoint:
    """One verified request target of a To Do operation, with the container it indexes.

    To Do indexes by **list** and by **task**: ``todo.list_task_lists`` reads the task lists
    of one user, and every other operation reads or writes a task inside one list. To Do
    addresses its container exclusively by path -- there is no body identifier -- so
    ``path_identifiers`` is the complete identifier set and every one of them must be
    non-blank before anything is sent.

    A claim may only call itself ``verified`` with recorded evidence: ``verified`` means the
    endpoint reference was recomputed **and** confirmed on a test tenant (WP16), so an
    offline milestone can only record ``documented_not_verified``.
    """

    operation: str
    method: str
    path_template: str
    container: str
    path_identifiers: tuple[str, ...] = ()
    application_permissions: tuple[str, ...] = ()
    claim_status: str = "documented_not_verified"
    claim_evidence: tuple[str, ...] = ()
    contract_call: str = ""
    contract_case: str = ""
    required_headers: tuple[str, ...] = ()
    documentation_page: str = ""

    def __post_init__(self) -> None:
        if self.claim_status not in PERMISSION_CLAIM_STATUSES:
            raise TodoContractError(
                f"{self.endpoint} declares unknown permission claim status {self.claim_status!r}"
            )
        if self.claim_status == "verified" and not self.claim_evidence:
            raise TodoContractError(
                f"{self.endpoint} claims a verified application role with no recorded evidence"
            )

    @property
    def endpoint(self) -> str:
        """The endpoint this row claims its application role from."""
        return f"{self.method} {self.path_template}"

    @property
    def required_identifiers(self) -> tuple[str, ...]:
        """Identifiers this endpoint cannot be called without."""
        return self.path_identifiers


#: The complete To Do endpoint table. Every role below is attributed to the endpoint it
#: belongs to instead of to a blanket "To Do" family, and no row carries a recorded endpoint
#: reference unless the repository already recorded one: only the task listing has one, so
#: the other four report ``not recorded`` rather than an invented URL.
TODO_ENDPOINTS: tuple[TodoEndpoint, ...] = (
    TodoEndpoint(
        operation="list_task_lists",
        method="GET",
        path_template="/users/{user_id}/todo/lists",
        container="user",
        path_identifiers=("user_id",),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_collection_request_information",
        contract_case="todo_lists",
    ),
    TodoEndpoint(
        operation="search",
        method="GET",
        path_template="/users/{user_id}/todo/lists/{todo_list_id}/tasks",
        container="list",
        path_identifiers=("user_id", "todo_list_id"),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_collection_request_information",
        contract_case="todo_tasks",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks",
    ),
    TodoEndpoint(
        operation="read",
        method="GET",
        path_template="/users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}",
        container="task",
        path_identifiers=("user_id", "todo_list_id", "todo_task_id"),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_item_request_information",
        contract_case="todo_read",
    ),
    TodoEndpoint(
        operation="create_tasks",
        method="POST",
        path_template="/users/{user_id}/todo/lists/{todo_list_id}/tasks",
        container="list",
        path_identifiers=("user_id", "todo_list_id"),
        application_permissions=("Tasks.ReadWrite.All",),
        contract_call="build_write_request_information",
        contract_case="todo_create_tasks",
    ),
    TodoEndpoint(
        operation="update_tasks",
        method="PATCH",
        path_template="/users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}",
        container="task",
        path_identifiers=("user_id", "todo_list_id", "todo_task_id"),
        application_permissions=("Tasks.ReadWrite.All",),
        contract_call="build_write_request_information",
        contract_case="todo_update_tasks",
    ),
)

TODO_OPERATIONS: tuple[str, ...] = (
    "todo.list_task_lists",
    "todo.search",
    "todo.read",
    "todo.create_tasks",
    "todo.update_tasks",
)

#: The To Do operations that write. Their application-mode support is derived from the
#: endpoint table above by :func:`todo_write_support`, never asserted here.
TODO_WRITE_OPERATIONS: frozenset[str] = frozenset({"todo.create_tasks", "todo.update_tasks"})

#: The ``sdk_contract`` case names of the To Do writes, read from the endpoint table so the
#: field policy below is keyed by the same name the request builders use.
TODO_TASK_WRITE_CASES: frozenset[str] = frozenset(
    endpoint.contract_case
    for endpoint in TODO_ENDPOINTS
    if f"todo.{endpoint.operation}" in TODO_WRITE_OPERATIONS
)


def todo_endpoints(key: str) -> tuple[TodoEndpoint, ...]:
    """Every verified endpoint of a To Do operation, in declaration order."""
    if key not in TODO_OPERATIONS:
        raise TodoContractError(f"{key} is not a declared To Do operation")
    return tuple(endpoint for endpoint in TODO_ENDPOINTS if f"todo.{endpoint.operation}" == key)


def required_todo_identifiers(key: str) -> tuple[str, ...]:
    """Identifiers a call must supply whichever endpoint of the operation it uses."""
    return _required_endpoint_identifiers(todo_endpoints(key))


def optional_todo_identifiers(key: str) -> tuple[str, ...]:
    """Identifiers a To Do endpoint accepts without the operation requiring them.

    Empty today: unlike ``planner.list_plans``, no To Do operation has a second endpoint
    selectable by supplying an extra identifier.
    """
    return _optional_endpoint_identifiers(todo_endpoints(key))


def validate_todo_identifiers(key: str, arguments: Mapping[str, Any]) -> dict[str, str]:
    """Normalize and validate the container identifiers a To Do call must carry.

    Refuses a blank or absent required identifier -- naming the endpoint it belongs to --
    before any approval, client or token work happens, so a To Do request can never be
    rendered against the wrong list or task.
    """
    return _resolve_endpoint_identifiers(
        error=TodoContractError,
        key=key,
        arguments=arguments,
        endpoints=todo_endpoints(key),
    )


#: The only To Do task fields a write may carry, and where each accepted field comes from.
#:
#: The generated ``TodoTask`` model declares 22 Graph properties; this milestone cannot
#: re-read the endpoint reference offline (R10), so the writable set stays deliberately
#: minimal and fail-closed. A field outside this mapping is refused with the case name --
#: including ``subject``, which no To Do model has, ``id`` and the ``*DateTime`` stamps the
#: service maintains itself, and every collection property -- instead of being sent and
#: silently dropped by Graph. This is a strict subset of the model on purpose: refusing a
#: field the endpoint might accept is honest, accepting a field it ignores is not.
TODO_TASK_WRITE_FIELDS: dict[str, str] = {
    "title": "the task text: `TodoTask.title`, the field the To Do request contract pins",
    "body": "the generated `TodoTask.body` (`ItemBody`)",
    "due_date_time": "the generated `TodoTask.due_date_time` (`DateTimeTimeZone`)",
    "status": "the generated `TodoTask.status` (`TaskStatus`), serialized only as an enum member",
}


def validate_todo_task_fields(case: str, fields: Mapping[str, Any]) -> dict[str, Any]:
    """Refuse any To Do task field the endpoint contract does not accept.

    ``case`` is the ``sdk_contract`` case name of the write (``todo_update_tasks``), the same
    spelling the request builder uses. Returns the accepted mapping; the SDK layer maps the
    enum-valued entry to its generated member and type-checks the model-valued ones. ``None``
    is refused for the same reason a field outside the policy is: it carries no value, so it
    would vanish from the request body instead of changing anything.
    """
    if case not in TODO_TASK_WRITE_CASES:
        raise TodoContractError(f"{case} is not a declared To Do write case")
    if not isinstance(fields, Mapping) or not fields:
        raise TodoContractError(f"{case} requires a non-empty mapping of task fields")
    accepted: dict[str, Any] = {}
    for name, value in fields.items():
        if name not in TODO_TASK_WRITE_FIELDS:
            raise TodoContractError(
                f"unknown field for {case}: {name}; writable To Do task fields are "
                f"{', '.join(TODO_TASK_WRITE_FIELDS)}"
            )
        if value is None:
            raise TodoContractError(f"{name} of {case} must carry a value, not None")
        accepted[name] = value
    return accepted


#: What an application-mode write support status is allowed to mean.
WRITE_SUPPORT_STATUSES = frozenset({"not_verified", "supported"})


@dataclass(frozen=True)
class ApplicationWriteSupport:
    """How much is known about one write endpoint in application mode.

    ``supported`` requires recorded evidence -- at least one endpoint of the write whose
    permission claim is ``verified``, which needs the endpoint reference recomputed *and*
    confirmed on a test tenant (WP16). Without that evidence the status is ``not_verified``,
    the operation stays non-executable, and no other field of this record can change that.
    """

    operation: str
    status: str
    reason: str
    evidence: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        if self.status not in WRITE_SUPPORT_STATUSES:
            raise TodoContractError(f"unknown application write support status: {self.status!r}")
        if self.status == "supported" and not self.evidence:
            raise TodoContractError(
                f"{self.operation} cannot be reported as supported in application mode "
                "without recorded endpoint evidence"
            )

    @property
    def supported(self) -> bool:
        return self.status == "supported"


#: The one reason a To Do write is reported ``not_verified``. Quoted verbatim in
#: ``references/graph-permissions.md`` so the document and the registry cannot disagree.
TODO_WRITE_NOT_VERIFIED_REASON = (
    "the application permission of this write is unverified: no endpoint reference was "
    "re-read offline and no test tenant was contacted (R10; remote verification is WP16)"
)


def todo_write_support(key: str) -> ApplicationWriteSupport:
    """Application-mode support of a To Do write, derived from its own endpoints.

    Derived, never asserted: the status flips to ``supported`` only when every endpoint of
    the write carries a ``verified`` application claim with recorded evidence, and stays
    ``not_verified`` until then. ``TodoEndpoint`` refuses a ``verified`` claim without
    evidence, so a promotion cannot be faked by editing one string.
    """
    if key not in TODO_WRITE_OPERATIONS:
        raise TodoContractError(f"{key} is not a declared To Do write operation")
    endpoints = todo_endpoints(key)
    verified = tuple(
        endpoint.endpoint
        for endpoint in endpoints
        if endpoint.claim_status == "verified" and endpoint.claim_evidence
    )
    if endpoints and len(verified) == len(endpoints):
        return ApplicationWriteSupport(
            operation=key,
            status="supported",
            reason="every endpoint of this write carries a verified application claim",
            evidence=verified,
        )
    return ApplicationWriteSupport(
        operation=key,
        status="not_verified",
        reason=TODO_WRITE_NOT_VERIFIED_REASON,
    )


#: Permission claims for the services whose endpoint contract the registry does not declare
#: yet. Planner and To Do rows are deliberately absent: each of their operations derives its
#: claim from its own endpoint table, so a blanket per-service claim cannot drift from the
#: endpoint it belongs to.
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
}
_UNSUPPORTED_APPLICATION = frozenset({"teams.search_messages", "teams.send_messages"})

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
    return _endpoint_backed_definition(
        service="planner", key=key, endpoints=planner_endpoints(key)
    )


def _todo_definition(key: str) -> OperationDefinition:
    return _endpoint_backed_definition(service="todo", key=key, endpoints=todo_endpoints(key))


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
    **{key: _todo_definition(key) for key in TODO_OPERATIONS},
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
    if key in TODO_WRITE_OPERATIONS:
        # Application-mode support of a To Do write is derived from the endpoint table, not
        # asserted: it is ``not_verified`` until every endpoint of the write carries a
        # verified application claim with recorded evidence (WP16). ``executable`` stays
        # whatever the registry says, so support never promotes execution by itself.
        support = todo_write_support(key)
        if not support.supported:
            return OperationStatus("not_verified", definition.implementation_status, False, support.reason)
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
