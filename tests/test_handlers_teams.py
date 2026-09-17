"""Strict offline contracts and auth-mode gates for Teams Graph handlers."""
from __future__ import annotations

import json
from urllib.parse import urlsplit

import pytest
from msgraph.generated.models.channel import Channel
from msgraph.generated.models.channel_collection_response import ChannelCollectionResponse
from msgraph.generated.models.entity_type import EntityType
from msgraph.generated.models.team import Team
from msgraph.generated.models.team_collection_response import TeamCollectionResponse

from microsoft365.contract import OPERATIONS, OPERATION_REGISTRY, Settings, operation_status
from microsoft365.registration import HANDLER_TABLE, active_actions, service_tool_handler
from tests.test_handlers_outlook_calendar import HandlerGraphAdapter, invocation, open_context, request_identity, search_query, serialized


TEAMS = "teams"


def all_selected_settings(**overrides) -> Settings:
    capabilities = {service: True for service in OPERATIONS}
    capabilities[TEAMS] = {"list_teams": True, "list_channels": True, "search_messages": True, "send_messages": True}
    return Settings.from_mapping({"tenant_id": "tenant", "client_id": "client", "capabilities": capabilities, **overrides})


def test_all_teams_handlers_are_registered_but_only_application_reads_are_active():
    assert {key for key in HANDLER_TABLE if key.startswith("teams.")} == {
        "teams.list_teams", "teams.list_channels", "teams.search_messages", "teams.send_messages"
    }
    assert active_actions(all_selected_settings(), TEAMS) == ("list_teams", "list_channels")


def test_application_and_delegated_statuses_are_explicit_and_honest():
    for operation in ("search_messages", "send_messages"):
        app = operation_status("application", TEAMS, operation)
        delegated = operation_status("delegated", TEAMS, operation)
        assert app.auth_status == "unsupported_auth_mode"
        assert not app.executable
        assert delegated.auth_status == "not_implemented"
        assert delegated.reason
    assert OPERATION_REGISTRY["teams.search_messages"].delegated.permissions == (
        "Chat.Read", "ChannelMessage.Read.All"
    )
    assert OPERATION_REGISTRY["teams.send_messages"].delegated.permissions == ("ChannelMessage.Send",)


def test_list_teams_uses_joined_teams_without_odata_parameters():
    response = TeamCollectionResponse(value=[Team(id="team-1")])
    adapter = HandlerGraphAdapter({("GET", "/users/user/joinedTeams"): response})
    payload = invocation("teams.list_teams")({"user_id": "user"}, open_context(adapter))
    request = adapter.requests[0]
    assert urlsplit(request.url).path == "/users/user/joinedTeams"
    assert search_query(request) == {}
    assert payload["result"]["value"][0]["id"] == "team-1"


def test_list_channels_allows_select_but_not_top_and_limits_processing():
    response = ChannelCollectionResponse(value=[Channel(id="channel-1", display_name="General")])
    adapter = HandlerGraphAdapter({("GET", "/teams/team-1/channels"): response})
    payload = invocation("teams.list_channels")(
        {"team_id": "team-1", "select": ["id", "display_name"]}, open_context(adapter)
    )
    request = adapter.requests[0]
    assert urlsplit(request.url).path == "/teams/team-1/channels"
    assert search_query(request) == {"select": ["id", "display_name"]}
    assert payload["result"]["value"][0]["display_name"] == "General"


def test_list_channels_rejects_top_before_client():
    called = False

    def client_factory():
        nonlocal called
        called = True
        raise AssertionError("client must not be touched")

    from microsoft365.handlers import HandlerContext

    with pytest.raises(Exception) as caught:
        invocation("teams.list_channels")({"team_id": "team-1", "top": 1}, HandlerContext(all_selected_settings(), client_factory))
    assert getattr(caught.value, "category", None) == "validation_error"
    assert not called


def test_search_messages_builds_real_nested_search_query_body():
    adapter = HandlerGraphAdapter({("POST", "/search/query"): {}})
    invocation("teams.search_messages")(
        {"query": "incident", "from_": 10, "size": 20}, open_context(adapter)
    )
    body = serialized(adapter.requests[0])
    assert body["requests"][0]["entityTypes"] == ["chatMessage"]
    assert body["requests"][0]["query"] == {"queryString": "incident"}
    assert body["requests"][0]["from"] == 10
    assert body["requests"][0]["size"] == 20


def test_send_messages_builds_chat_message_on_channel_messages():
    adapter = HandlerGraphAdapter({("POST", "/teams/team-1/channels/channel-1/messages"): {}})
    invocation("teams.send_messages")(
        {"team_id": "team-1", "channel_id": "channel-1", "body": "hello"}, open_context(adapter)
    )
    request = adapter.requests[0]
    assert urlsplit(request.url).path == "/teams/team-1/channels/channel-1/messages"
    assert serialized(request) == {"body": {"content": "hello", "contentType": "text"}}


@pytest.mark.parametrize("operation, arguments", [
    ("search_messages", {"query": "incident"}),
    ("send_messages", {"team_id": "team-1", "channel_id": "channel-1", "body": "hello"}),
])
def test_application_dispatch_refuses_delegated_only_teams_operations_before_client(operation, arguments):
    called = False

    def boom():
        nonlocal called
        called = True
        raise AssertionError("client must not be created")

    settings = all_selected_settings()
    settings = Settings.from_mapping({
        "tenant_id": settings.tenant_id,
        "client_id": settings.client_id,
        "capabilities": settings.capabilities,
        "authentication_mode": "application",
    })
    payload = json.loads(service_tool_handler(TEAMS, {"action": operation, **arguments}, settings=settings))
    assert payload["error"] == "operation_not_implemented"
    assert payload["operation"] == operation
    assert not called
