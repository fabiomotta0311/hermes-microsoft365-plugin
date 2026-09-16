"""Graph client construction with execution-time scoped secrets.

The credential value never leaves this module: it is read from the host's scoped secret
lookup inside the call, handed to the SDK, and never stored, logged or echoed. Every
failure raised here is a sanitized taxonomy error -- raw Azure text (which can quote the
credential) is chained as ``__cause__`` and is not part of the message.
"""
from __future__ import annotations

from .contract import Settings
from .errors import MESSAGES, GraphError
from .preflight import SECRET_ENV_NAME

GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default"


def _failure(category: str, cause: BaseException | None = None) -> GraphError:
    error = GraphError(category, MESSAGES[category])
    if cause is not None:
        error.__cause__ = cause
        error.__suppress_context__ = True
    return error


def create_graph_client(settings: Settings):
    """Build the Graph client for the application auth mode, fail-closed.

    Categories: ``unsupported_auth_mode`` when the configured mode is not implemented,
    ``configuration_error`` when the identity settings are incomplete,
    ``authentication_required`` when no secret is available or the credential cannot be
    constructed.
    """
    if settings.authentication_mode != "application":
        raise _failure("unsupported_auth_mode")
    if not settings.tenant_id or not settings.client_id:
        raise _failure("configuration_error")

    from agent.secret_scope import get_secret

    secret = get_secret(SECRET_ENV_NAME, None)
    if not secret:
        raise _failure("authentication_required")

    from azure.identity import ClientSecretCredential
    from msgraph import GraphServiceClient

    try:
        credential = ClientSecretCredential(
            tenant_id=settings.tenant_id,
            client_id=settings.client_id,
            client_secret=secret,
        )
    except GraphError:
        raise
    except Exception as exc:
        # A raw credential failure may quote the secret; it is chained, never reported.
        raise _failure("authentication_required", exc)

    try:
        return GraphServiceClient(credentials=credential, scopes=[GRAPH_DEFAULT_SCOPE])
    except GraphError:
        raise
    except Exception as exc:
        raise _failure("configuration_error", exc)
