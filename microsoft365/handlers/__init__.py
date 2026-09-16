"""Operation handlers and the self-populating handler registry (WP6).

How a handler is registered
---------------------------
Each ``handlers/<service>.py`` module registers its own operation handlers at import time
with the :func:`handler` decorator. :func:`load_handlers` walks this package with
``pkgutil.iter_modules`` and imports **every** module it finds, so a later work package adds a
service by dropping in ``handlers/<service>.py`` and never edits ``registration.py``. There is
no per-service import list to keep in step. A module that cannot be imported raises, an
unknown operation is refused and a duplicate registration is refused
(:class:`HandlerRegistrationError`): a malformed handler module is never skipped silently.

What a handler is
-----------------
A handler is a **synchronous** callable ``handler(arguments, context) -> dict`` returning the
tool payload. It never builds a URL, never reads a secret and never creates a client: the
generated SDK builders come from ``context``'s client, the credential is resolved by the
host's scoped secret lookup inside ``client_factory``, and the single sync/async boundary is
``microsoft365.execution``.

Why a handler rebuilds its request instead of handing the contract's request information over
----------------------------------------------------------------------------------------------
``microsoft365/sdk_contract.py`` (WP1) is the verified request contract of every operation and
the oracle this package is pinned against: ``tests/test_handlers_outlook_calendar.py`` asserts
that the request a handler executes is identical to it -- method, path, typed query, serialized
body and conditional headers. It cannot be *handed over*, though: Kiota keeps a typed
configuration's query values in the request's ``query_parameters`` instead of the rendered URL,
while the execution seam builds its request from a builder plus a configuration. A handler
therefore rebuilds the same request from the same declared arguments, with the same generated
builder and the same generated typed query class -- and the pin is what keeps the two in
lockstep.

Rules a handler obeys
---------------------
* every action-specific check runs **before** the client (and therefore before any secret,
  credential or token work), and re-checks in the handler what the shared argument contract
  already checks -- defence in depth, never a replacement;
* an argument the verified request contract cannot express, and a field that would not change
  the serialized request, are **refused** (:func:`refuse`), never silently dropped;
* a failure is a :class:`microsoft365.errors.GraphError` from the canonical taxonomy; no raw
  SDK/Graph/Azure text is ever interpolated into a message.
"""
from __future__ import annotations

import importlib
import json
import pkgutil
from dataclasses import dataclass
from typing import Any, Callable, Mapping

from kiota_abstractions.base_request_configuration import RequestConfiguration
from kiota_abstractions.headers_collection import HeadersCollection

from ..contract import OPERATIONS, WRITE_OPERATIONS, Settings
from ..errors import MESSAGES, GraphError
from ..execution import ExecutionError, resolve_request_adapter
from ..results import normalize_result

#: Item budget of a read that did not ask for one (the contract's ``$top`` range is 1-100).
DEFAULT_LIMIT = 50
#: Highest item budget a read may ask for; the paginator enforces the same ceiling.
MAX_LIMIT = 100
#: Longest argument name echoed in a refusal message (never a value, never unbounded).
_LABEL_LENGTH = 48

#: Every registered handler, keyed ``"<service>.<operation>"``. Populated at import time by
#: :func:`handler` from the modules :func:`load_handlers` imports.
_REGISTRY: dict[str, "HandlerRecord"] = {}


class HandlerRegistrationError(ValueError):
    """A handler module tried to register something that could never be dispatched."""


@dataclass(frozen=True)
class HandlerRecord:
    """One registered operation handler, with the operation it belongs to."""

    service: str
    operation: str
    function: Callable[[Mapping, "HandlerContext"], dict]
    write: bool

    @property
    def key(self) -> str:
        return f"{self.service}.{self.operation}"


@dataclass(frozen=True)
class HandlerContext:
    """What a handler may use, injected by the caller that dispatched it.

    ``client_factory`` produces the graph client -- and therefore the credential, resolved
    inside the host's scoped secret lookup -- only when a handler asks for it, which is after
    every action-specific check has passed. No handler reads a secret or builds a credential
    itself.
    """

    settings: Settings
    client_factory: Callable[[], Any]

    def __post_init__(self) -> None:
        if not isinstance(self.settings, Settings):
            raise ExecutionError("configuration_error", MESSAGES["configuration_error"])
        if not callable(self.client_factory):
            raise ExecutionError("configuration_error", MESSAGES["configuration_error"])


def handler(service: str, operation: str) -> Callable:
    """Register one operation handler at import time (see the module docstring).

    The service and the operation are checked against the administrative contract here, so a
    typo fails while the module is imported instead of silently producing a handler nothing
    can dispatch.
    """
    service = str(service)
    if service not in OPERATIONS:
        raise HandlerRegistrationError(f"{service!r} is not a declared microsoft 365 service")
    if not isinstance(operation, str) or operation not in OPERATIONS[service]:
        raise HandlerRegistrationError(f"{service}.{operation} is not a declared operation")
    key = f"{service}.{operation}"

    def decorate(function: Callable) -> Callable:
        if not callable(function):
            raise HandlerRegistrationError(f"{key} must be registered with a callable")
        if key in _REGISTRY:
            raise HandlerRegistrationError(
                f"{key} is already registered by {_REGISTRY[key].function.__name__}"
            )
        _REGISTRY[key] = HandlerRecord(
            service=service,
            operation=operation,
            function=function,
            write=operation in WRITE_OPERATIONS,
        )
        return function

    return decorate


def load_handlers() -> dict[str, HandlerRecord]:
    """Import every module in this package and return the populated mapping.

    Discovery is the whole registration mechanism: a new ``handlers/<service>.py`` module is
    found and imported without a single edit outside it. An import failure is never caught --
    a malformed module must fail loudly and visibly.
    """
    for module in sorted(pkgutil.iter_modules(__path__), key=lambda found: found.name):
        importlib.import_module(f"{__name__}.{module.name}")
    return dict(_REGISTRY)


def open_graph_client(context: HandlerContext) -> tuple[Any, Any]:
    """Create the graph client once and return it with its adapter, fail-closed.

    ``resolve_request_adapter`` is the plugin's single entry point for authentication: a
    client factory that raises becomes a typed ``authentication_required`` (or the category
    the factory itself decided) instead of a raw credential error, and the returned adapter is
    the one injected into the client that produced every builder a handler uses.
    """
    if not isinstance(context, HandlerContext):
        raise ExecutionError("configuration_error", MESSAGES["configuration_error"])
    created: list = []

    def factory():
        client = context.client_factory()
        created.append(client)
        return client

    adapter = resolve_request_adapter(factory)
    return created[0], adapter


def typed_configuration(builder, **values) -> RequestConfiguration:
    """Build the generated builder's own typed query configuration.

    The query class is the one the SDK generates for this builder, exactly as
    ``microsoft365/sdk_contract.py`` does it, so an undeclared query name raises instead of
    reaching the wire. Values that are ``None`` are omitted: a caller that asked for no
    ``$select`` gets no ``$select``.
    """
    query_type = getattr(type(builder), f"{type(builder).__name__}GetQueryParameters")
    supplied = {name: value for name, value in values.items() if value is not None}
    return RequestConfiguration(query_parameters=query_type(**supplied))


def headers_configuration(**headers) -> RequestConfiguration:
    """Build a configuration carrying only the conditional headers that have a value.

    A real ``HeadersCollection`` is required: a plain ``dict`` raises inside the generated
    ``configure()``.
    """
    collection = HeadersCollection()
    for name, value in headers.items():
        if value:
            collection.add(name, value)
    return RequestConfiguration(headers=collection)


def success_payload(key: str, value: Any = None) -> dict:
    """The tool payload of a completed operation, with the result normalized when there is one."""
    payload: dict = {"operation": key, "status": "succeeded"}
    if value is not None:
        payload["result"] = normalize_result(value)
    return payload


def collection_payload(key: str, paged) -> dict:
    """The payload of a bounded collection read: the normalized page plus its continuation.

    The continuation is reported as the paginator decided it -- truncated or not, whether a
    next link existed, how many requests and items were fetched and why the traversal stopped.
    """
    continuation = paged.continuation
    return {
        "operation": key,
        "status": "succeeded",
        "result": paged.payload,
        "continuation": {
            "truncated": continuation.truncated,
            "next_link_present": continuation.next_link_present,
            "pages_fetched": continuation.pages_fetched,
            "items_fetched": continuation.items_fetched,
            "stop_reason": continuation.stop_reason,
        },
    }


def refuse(category: str, message: str) -> None:
    """Raise a sanitized taxonomy failure. A message never carries a value from the payload."""
    raise GraphError(category, message)


# --------------------------------------------------------------------------------------
# Shared argument surface
# --------------------------------------------------------------------------------------
# The argument contract (microsoft365/validation.py) already bounds every declared argument;
# these helpers are what a handler needs on top of it: the *typed* value the generated models
# and query classes actually take, and a refusal (never a silent drop) for a value the
# contract cannot describe. They are shared by the service modules instead of being copied
# into each one.


def identifier(arguments: Mapping, name: str, *, operation: str) -> str:
    """A non-blank identifier this call cannot run without."""
    value = arguments.get(name)
    if not isinstance(value, str) or not value.strip():
        refuse("validation_error", f"{operation}: {name} must be a non-blank string identifier")
    return value


def text(arguments: Mapping, name: str, *, operation: str) -> str | None:
    """An optional free-text argument, passed through exactly as the caller supplied it."""
    value = arguments.get(name)
    if value is None:
        return None
    if not isinstance(value, str) or not value.strip():
        refuse("validation_error", f"{operation}: {name} must be a non-blank string")
    return value


def select(arguments: Mapping, *, operation: str) -> list | None:
    """The caller's ``$select``: a list of field names, or nothing at all."""
    value = arguments.get("select")
    if value is None:
        return None
    if not isinstance(value, list) or not value:
        refuse("validation_error", f"{operation}: select must be a non-empty list of field names")
    return value


def recipients(arguments: Mapping, name: str, *, operation: str, required: bool = False) -> list | None:
    """Real ``Recipient`` models for a declared recipient list.

    An entry is either an address string or an ``{address, name}`` object, as the argument
    contract declares; anything else is refused instead of being dropped from the request.
    """
    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.recipient import Recipient

    value = arguments.get(name)
    if value is None:
        if required:
            refuse("validation_error", f"{operation}: {name} is required")
        return None
    if not isinstance(value, list) or not value:
        refuse(
            "validation_error",
            f"{operation}: {name} must be a non-empty list of email addresses",
        )
    models = []
    for index, entry in enumerate(value):
        if isinstance(entry, str):
            address, display = entry, None
        elif isinstance(entry, Mapping):
            address, display = entry.get("address"), entry.get("name")
        else:
            address, display = None, None
        if not isinstance(address, str) or not address.strip():
            refuse(
                "validation_error",
                f"{operation}: {name}[{index}] must be an email address or an address/name object",
            )
        models.append(
            Recipient(
                email_address=EmailAddress(
                    address=address, name=display if isinstance(display, str) and display else None
                )
            )
        )
    return models


def field_label(value: Any) -> str:
    """A bounded, printable argument name for a refusal message: never the value itself."""
    if isinstance(value, str):
        cleaned = "".join(character if character.isprintable() else "?" for character in value)
        return cleaned[:_LABEL_LENGTH]
    return type(value).__name__


def item_budget(arguments: Mapping, *, operation: str) -> int:
    """The caller's item budget: the number of items a bounded read may return."""
    top = arguments.get("top")
    if top is None:
        return DEFAULT_LIMIT
    if type(top) is not int or not 1 <= top <= MAX_LIMIT:
        refuse(
            "validation_error",
            f"{operation}: top must be an integer between 1 and {MAX_LIMIT}",
        )
    return top


def serialized_keys(request_information) -> set:
    """Every JSON key a built request body carries, nested keys included."""
    content = getattr(request_information, "content", None)
    if not content:
        return set()
    try:
        payload = json.loads(content.decode("utf-8"))
    except (ValueError, AttributeError, UnicodeDecodeError):
        return set()
    return _payload_keys(payload)


def _payload_keys(payload: Any) -> set:
    found: set = set()
    if isinstance(payload, Mapping):
        for key, value in payload.items():
            found.add(key)
            found |= _payload_keys(value)
    elif isinstance(payload, (list, tuple)):
        for item in payload:
            found |= _payload_keys(item)
    return found


def require_changed_fields(
    request_information, *, operation: str, supplied: Mapping[str, tuple]
) -> None:
    """Refuse a field that would not change the request it was meant to change.

    ``supplied`` maps each declared field the caller asked to change to the wire keys it must
    produce. Kiota's writer drops a value it cannot serialize (a plain string where an enum
    member is expected, for instance), so a request that does not actually carry the change is
    refused here instead of being sent as a silent no-op.
    """
    present = serialized_keys(request_information)
    ignored = sorted(name for name, keys in supplied.items() if not present.intersection(keys))
    if ignored:
        refuse(
            "validation_error",
            f"{operation}: {', '.join(ignored)} would not change the request, so it is refused "
            "instead of being sent unchanged",
        )
