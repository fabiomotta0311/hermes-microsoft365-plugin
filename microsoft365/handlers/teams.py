"""Teams Graph handlers with explicit application/delegated boundaries (WP8).

``list_teams`` and ``list_channels`` are the only application-mode executable operations.
The two other handlers are retained and contract-tested for delegated authentication, but the
registry and dispatch gate refuse them before any client, credential, secret, or network work.
Bot Framework integrations are intentionally outside this module.
"""
from __future__ import annotations

from typing import Mapping

from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.entity_type import EntityType
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.search_query import SearchQuery
from msgraph.generated.models.search_request import SearchRequest
from msgraph.generated.models.chat_message import ChatMessage
from msgraph.generated.search.query.query_post_request_body import QueryPostRequestBody

from ..execution import execute_request
from ..paging import paginate
from . import (
    HandlerContext,
    collection_payload,
    handler,
    identifier,
    item_budget,
    open_graph_client,
    refuse,
    select,
    success_payload,
    text,
    typed_configuration,
)

LIST_TEAMS = "teams.list_teams"
LIST_CHANNELS = "teams.list_channels"
SEARCH_MESSAGES = "teams.search_messages"
SEND_MESSAGES = "teams.send_messages"


def _body(arguments: Mapping, *, operation: str) -> ItemBody:
    content = text(arguments, "body", operation=operation)
    if content is None:
        refuse("validation_error", f"{operation}: body is required")
    content_type = arguments.get("content_type", "text")
    if content_type == "text":
        kind = BodyType.Text
    elif content_type == "html":
        kind = BodyType.Html
    else:
        refuse("validation_error", f"{operation}: content_type must be text or html")
    return ItemBody(content_type=kind, content=content)


@handler("teams", "list_teams")
def list_teams(arguments: Mapping, context: HandlerContext) -> dict:
    """List joined teams without sending unsupported OData parameters."""
    user_id = identifier(arguments, "user_id", operation=LIST_TEAMS)
    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).joined_teams
    paged = paginate(builder, adapter=adapter, limit=item_budget(arguments, operation=LIST_TEAMS))
    return collection_payload(LIST_TEAMS, paged)


@handler("teams", "list_channels")
def list_channels(arguments: Mapping, context: HandlerContext) -> dict:
    """List a team's channels; ``$select`` is supported but ``$top`` is not."""
    team_id = identifier(arguments, "team_id", operation=LIST_CHANNELS)
    fields = select(arguments, operation=LIST_CHANNELS)
    # ``top`` is not a declared argument for this endpoint. Keep this defense in depth for
    # direct handler calls that bypass the shared validator.
    if "top" in arguments:
        refuse("validation_error", f"{LIST_CHANNELS}: top is not supported by the Graph endpoint")
    client, adapter = open_graph_client(context)
    builder = client.teams.by_team_id(team_id).channels
    configuration = typed_configuration(builder, select=fields)
    paged = paginate(builder, adapter=adapter, limit=item_budget(arguments, operation=LIST_CHANNELS), configuration=configuration)
    return collection_payload(LIST_CHANNELS, paged)


@handler("teams", "search_messages")
def search_messages(arguments: Mapping, context: HandlerContext) -> dict:
    """Search chat messages using the real nested ``search/query`` request shape."""
    query = text(arguments, "query", operation=SEARCH_MESSAGES)
    from_ = arguments.get("from_", 0)
    size = arguments.get("size", 25)
    if type(from_) is not int or from_ < 0:
        refuse("validation_error", f"{SEARCH_MESSAGES}: from_ must be a non-negative integer")
    if type(size) is not int or not 1 <= size <= 100:
        refuse("validation_error", f"{SEARCH_MESSAGES}: size must be an integer between 1 and 100")
    body = QueryPostRequestBody(
        requests=[
            SearchRequest(
                entity_types=[EntityType.ChatMessage],
                query=SearchQuery(query_string=query),
                from_=from_,
                size=size,
            )
        ]
    )
    client, adapter = open_graph_client(context)
    response = execute_request(client.search.query, method="POST", body=body, adapter=adapter)
    return success_payload(SEARCH_MESSAGES, response)


@handler("teams", "send_messages")
def send_messages(arguments: Mapping, context: HandlerContext) -> dict:
    """Post a real ``ChatMessage`` to a channel's messages collection."""
    team_id = identifier(arguments, "team_id", operation=SEND_MESSAGES)
    channel_id = identifier(arguments, "channel_id", operation=SEND_MESSAGES)
    body = _body(arguments, operation=SEND_MESSAGES)
    client, adapter = open_graph_client(context)
    builder = client.teams.by_team_id(team_id).channels.by_channel_id(channel_id).messages
    response = execute_request(builder, method="POST", body=ChatMessage(body=body), adapter=adapter)
    return success_payload(SEND_MESSAGES, response)
