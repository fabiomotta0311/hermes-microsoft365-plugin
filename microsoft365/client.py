"""Graph client construction with execution-time scoped secrets."""
from __future__ import annotations

from .contract import Settings
from .preflight import SECRET_ENV_NAME

GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default"


def create_graph_client(settings: Settings):
    if settings.authentication_mode != "application":
        raise RuntimeError("delegated authentication is not implemented")
    if not settings.tenant_id or not settings.client_id:
        raise RuntimeError("tenant_id and client_id are required")

    from agent.secret_scope import get_secret

    secret = get_secret(SECRET_ENV_NAME, None)
    if not secret:
        raise RuntimeError(f"{SECRET_ENV_NAME} is not available in the active secret scope")

    from azure.identity import ClientSecretCredential
    from msgraph import GraphServiceClient

    credential = ClientSecretCredential(
        tenant_id=settings.tenant_id,
        client_id=settings.client_id,
        client_secret=secret,
    )
    return GraphServiceClient(credentials=credential, scopes=[GRAPH_DEFAULT_SCOPE])
