"""The single deterministic sync/async seam for Microsoft Graph requests (WP0).

Why this module exists
----------------------
msgraph-sdk 1.62.0 generates *coroutine* methods (``MessagesRequestBuilder.get``,
``ContentRequestBuilder.put``, ``SendRequestBuilder.post`` ...) while Hermes registers
plugin tools with ``is_async=False``. Handlers therefore need exactly one place that
crosses that boundary, instead of every handler inventing its own.

Strategy chosen: (A) real typed request information + injected adapter
----------------------------------------------------------------------
* :func:`request_information_sender` builds the **real** ``RequestInformation`` with the
  generated builder's own ``to_<method>_request_information``. It never sends, so HTTP
  method, URL, query parameters, headers and the serialized body stay assertable offline
  (by contract tests and by handlers before they execute).
* :func:`execute_request` binds that same builder/configuration to an **injected**
  ``RequestAdapter`` -- the one the client was constructed with -- and performs the send
  through the SDK's own generated caller (``builder.get()``/``post()``/``put()``/
  ``patch()``/``delete()``) inside :func:`run_async`.

Why the send is delegated to the generated caller instead of calling
``adapter.send_async`` here: in the installed SDK the send variant is a property of the
endpoint, not of the HTTP method (verified against the generated sources).
``MessagesRequestBuilder.get`` uses ``send_async(request, MessageCollectionResponse)``,
``ContentRequestBuilder.get`` uses ``send_primitive_async(request, "bytes")``, and
``SendMailRequestBuilder.post``/``ContentRequestBuilder.delete`` use
``send_no_response_content_async``. Reproducing that dispatch inside the plugin would
duplicate generated code and drift from it, so the seam owns the *request artifact* and
the *sync/async boundary*, while the SDK keeps owning its dispatch -- always through the
adapter injected into the client, never through a transport created here.

Request-body convention
-----------------------
Generated ``to_*_request_information`` and ``*()`` callables are invoked by inspecting
their real signature: an endpoint with a request body leads with a ``body`` parameter
(``to_post_request_information(body, request_configuration=None)``), an endpoint without
one does not (``SendRequestBuilder.post(request_configuration=None)``). No per-operation
table is hardcoded here.

Guarantees
----------
* No credential, token, client or adapter lives in module state: the adapter is always a
  parameter, and this module never imports a credential library nor the host's secret
  scope.
* Failures are typed: :class:`ExecutionError` carries a stable ``category``. Raw
  SDK/Azure exception text is never copied into the message; the original exception is
  only chained as ``__cause__`` for local debugging.
* The seam is deliberately not async: it is the one boundary a synchronous tool handler
  calls. Handlers are synchronous and call ``execute_request``; ``run_async`` refuses to
  nest inside a running event loop (typed ``internal_error``) instead of deadlocking, so
  an ``async`` handler must never call the seam from inside its own coroutine.
"""
from __future__ import annotations

import asyncio
import inspect

from kiota_abstractions.request_adapter import RequestAdapter

from .errors import MESSAGES, GraphError, classify_upstream, to_graph_error

_HTTP_METHODS = ("GET", "POST", "PUT", "PATCH", "DELETE")
_CALLER_NAMES = {method: method.lower() for method in _HTTP_METHODS}


class ExecutionError(GraphError):
    """Typed, sanitized failure raised at the execution seam.

    The category comes from the canonical taxonomy in :mod:`microsoft365.errors` (this class
    *is* a taxonomy error, so a category invented here is rejected). The message is a static
    taxonomy template: raw SDK, Kiota, Graph or Azure text never appears in it. The
    originating exception, when there is one, is available as ``__cause__`` and is for local
    debugging only -- handlers must not return it.
    """


def _normalized_method(method) -> str:
    normalized = str(method or "").strip().upper()
    if normalized not in _HTTP_METHODS:
        raise ExecutionError("configuration_error", "unsupported HTTP method for a graph request")
    return normalized


def _takes_body(callable_object) -> bool:
    """True when the generated callable leads with its request-body parameter."""
    parameters = list(inspect.signature(callable_object).parameters)
    return bool(parameters) and parameters[0] == "body"


def _call(callable_object, *, body, configuration):
    if _takes_body(callable_object):
        return callable_object(body, configuration)
    return callable_object(configuration)


def _sanitized(category: str, exc: BaseException) -> ExecutionError:
    """Build a seam error from the taxonomy: same category, static message, safe metadata."""
    return ExecutionError(
        category,
        MESSAGES[category],
        retryable=exc.retryable,
        retry_after_seconds=exc.retry_after_seconds,
        correlation_id=exc.correlation_id,
        status_code=exc.status_code,
    )


def run_async(awaitable, *, category: str = "service_error"):
    """Run ``awaitable`` to completion on a private event loop.

    This is the only place the plugin crosses from a synchronous tool handler into the
    SDK's coroutines. The event loop is created and torn down per call, so no loop,
    client or credential is shared between calls.

    A failure that carries a real upstream signal (an HTTP status, a structured Graph error code,
    an Azure/Kiota/transport family) is classified by :mod:`microsoft365.errors`; a purely
    local failure keeps the ``category`` this call site declares. Either way the reported
    message is a static taxonomy template -- raw text stays in ``__cause__`` only.
    """
    if not inspect.isawaitable(awaitable):
        raise ExecutionError("configuration_error", MESSAGES["configuration_error"])

    try:
        asyncio.get_running_loop()
    except RuntimeError:
        pass
    else:
        close = getattr(awaitable, "close", None)
        if callable(close):
            close()
        raise ExecutionError(
            "internal_error", "cannot block on an async graph call inside a running event loop"
        )

    try:
        return asyncio.run(awaitable)
    except ExecutionError:
        raise
    except Exception as exc:
        converted = to_graph_error(exc, category=classify_upstream(exc, default=category))
        raise _sanitized(converted.category, converted) from exc


def request_information_sender(builder, *, method, configuration=None, body=None):
    """Build the generated builder's typed ``RequestInformation``. Never sends.

    The returned object exposes everything a later work package or handler needs to
    assert before executing: ``http_method``, ``url``, ``query_parameters``, ``headers``
    and the serialized ``content``.
    """
    normalized = _normalized_method(method)
    builder_method = f"to_{normalized.lower()}_request_information"
    build = getattr(builder, builder_method, None)
    if not callable(build):
        raise ExecutionError("configuration_error", f"builder does not expose {builder_method}")
    try:
        return _call(build, body=body, configuration=configuration)
    except ExecutionError:
        raise
    except (TypeError, ValueError) as exc:
        raise ExecutionError(
            "configuration_error", f"{builder_method} rejected the supplied request"
        ) from exc


def execute_request(builder, *, method, configuration=None, body=None, adapter):
    """Execute one graph request and return the SDK's typed response.

    ``adapter`` must be the request adapter injected into the client that produced
    ``builder`` (``client.request_adapter``); anything else fails closed before any work.
    """
    normalized = _normalized_method(method)
    if not isinstance(adapter, RequestAdapter):
        raise ExecutionError(
            "configuration_error", "a request adapter must be injected before executing a graph request"
        )
    if getattr(builder, "request_adapter", None) is not adapter:
        raise ExecutionError(
            "configuration_error", "builder is not bound to the injected request adapter"
        )

    request_information_sender(builder, method=normalized, configuration=configuration, body=body)

    caller = getattr(builder, _CALLER_NAMES[normalized], None)
    if not callable(caller):
        raise ExecutionError(
            "configuration_error", f"builder does not expose {_CALLER_NAMES[normalized]}()"
        )
    awaitable = _call(caller, body=body, configuration=configuration)
    return run_async(awaitable, category="transport_error")


def resolve_request_adapter(client_factory):
    """Resolve the injected adapter fail-closed, wrapping auth/configuration failures.

    ``client_factory`` is supplied by the caller and creates the graph client (and so
    reads the credential inside the host's secret scope). Nothing is cached here.
    """
    if not callable(client_factory):
        raise ExecutionError(
            "configuration_error", "a graph client factory is required to resolve a request adapter"
        )
    try:
        client = client_factory()
    except ExecutionError:
        raise
    except GraphError as exc:
        # A category the taxonomy already decided (e.g. unsupported_auth_mode) is preserved
        # instead of being flattened into authentication_required.
        raise _sanitized(exc.category, exc) from exc
    except Exception as exc:
        raise ExecutionError(
            "authentication_required", MESSAGES["authentication_required"]
        ) from exc
    adapter = getattr(client, "request_adapter", None)
    if not isinstance(adapter, RequestAdapter):
        raise ExecutionError("configuration_error", "graph client does not expose a request adapter")
    return adapter
