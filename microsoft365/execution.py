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

_HTTP_METHODS = ("GET", "POST", "PUT", "PATCH", "DELETE")
_CALLER_NAMES = {method: method.lower() for method in _HTTP_METHODS}


class ExecutionError(RuntimeError):
    """Typed, sanitized failure raised at the execution seam.

    ``category`` is stable and safe to report; ``message`` never carries raw SDK, Kiota,
    Graph or Azure text. The originating exception, when there is one, is available as
    ``__cause__`` and is for local debugging only -- handlers must not return it.
    """

    def __init__(self, category: str, message: str):
        self.category = str(category)
        self.message = str(message)
        super().__init__(f"{self.category}: {self.message}")


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


def run_async(awaitable, *, category: str = "service_error"):
    """Run ``awaitable`` to completion on a private event loop.

    This is the only place the plugin crosses from a synchronous tool handler into the
    SDK's coroutines. The event loop is created and torn down per call, so no loop,
    client or credential is shared between calls.
    """
    if not inspect.isawaitable(awaitable):
        raise ExecutionError("configuration_error", "run_async requires an awaitable")

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
        raise ExecutionError(category, "graph call failed") from exc


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
    except Exception as exc:
        raise ExecutionError("authentication_required", "graph client could not be created") from exc
    adapter = getattr(client, "request_adapter", None)
    if not isinstance(adapter, RequestAdapter):
        raise ExecutionError("configuration_error", "graph client does not expose a request adapter")
    return adapter
