"""Strict read-side msgraph-sdk contracts: item reads, searches, teams, planner.

Same rules as ``tests/test_sdk_contract.py``: real generated builders, a concrete
``RequestAdapter`` double whose ``send_*`` methods all raise, no ``SimpleNamespace``,
no ``MagicMock`` and no permissive ``__getattr__``.
"""
from __future__ import annotations

from typing import Any
from urllib.parse import urlsplit

import pytest
from kiota_abstractions.request_adapter import RequestAdapter


class StrictTransportAdapter(RequestAdapter):
    """Concrete Kiota boundary double: no generated builder/member fakery.

    ``get_serialization_writer_factory`` raises here on purpose -- these contracts are
    all GETs, so no body may ever be serialized. Body-serializing rows live in
    ``tests/test_sdk_contract_writes.py``, which supplies a real writer factory.
    """

    def __init__(self):
        self.base_url = "https://graph.microsoft.com/v1.0"

    def enable_backing_store(self, backing_store_factory=None):
        return None

    def get_serialization_writer_factory(self):
        raise AssertionError("no GET contract may serialize a request body")

    async def send_async(self, *args: Any, **kwargs: Any):
        raise AssertionError("network transport must not run")

    async def send_collection_async(self, *args: Any, **kwargs: Any):
        raise AssertionError("network transport must not run")

    async def send_collection_of_primitive_async(self, *args: Any, **kwargs: Any):
        raise AssertionError("network transport must not run")

    async def send_primitive_async(self, *args: Any, **kwargs: Any):
        raise AssertionError("network transport must not run")

    async def send_no_response_content_async(self, *args: Any, **kwargs: Any):
        raise AssertionError("network transport must not run")

    async def convert_to_native_async(self, *args: Any, **kwargs: Any):
        raise AssertionError("native conversion must not run")


@pytest.fixture
def graph_client():
    from msgraph import GraphServiceClient

    return GraphServiceClient(request_adapter=StrictTransportAdapter())


ITEM_READ_ARGUMENTS = {
    "user_id": "user",
    "message_id": "msg",
    "event_id": "evt",
    "todo_list_id": "list",
    "todo_task_id": "task",
    "planner_task_id": "ptask",
    "drive_id": "drive",
    "drive_item_id": "root:/a/b.txt:",
}


@pytest.mark.parametrize(
    ("case", "expected_path", "select"),
    [
        ("outlook_read", "/users/user/messages/msg", ["id", "subject"]),
        ("calendar_read", "/users/user/calendar/events/evt", ["id", "subject"]),
        ("todo_read", "/users/user/todo/lists/list/tasks/task", ["id", "title"]),
        ("planner_read", "/planner/tasks/ptask", ["id", "title"]),
        ("drive_item_read", "/drives/drive/items/root%3A%2Fa%2Fb.txt%3A", ["id", "name"]),
    ],
)
def test_real_generated_item_builder_accepts_typed_get_configuration(
    graph_client, case, expected_path, select
):
    from microsoft365.sdk_contract import build_item_request_information

    request = build_item_request_information(
        graph_client, case, select=select, **ITEM_READ_ARGUMENTS
    )

    assert request.http_method.value == "GET"
    assert request.query_parameters == {"%24select": select}
    assert urlsplit(request.url).path == expected_path
    assert type(request).__name__ == "RequestInformation"

    # the per-read select is what the caller asked for, not a hardcoded list
    renamed = build_item_request_information(
        graph_client, case, select=["id"], **ITEM_READ_ARGUMENTS
    )
    assert renamed.query_parameters == {"%24select": ["id"]}

    # and omitting it sends no $select at all
    bare = build_item_request_information(graph_client, case, **ITEM_READ_ARGUMENTS)
    assert bare.query_parameters == {}


@pytest.mark.parametrize("case", ["outlook_read", "calendar_read", "todo_read"])
def test_user_scoped_item_read_rejects_an_empty_user_id(graph_client, case):
    from microsoft365.sdk_contract import build_item_request_information

    arguments = dict(ITEM_READ_ARGUMENTS)
    arguments["user_id"] = "   "

    with pytest.raises(ValueError, match="user_id is required"):
        build_item_request_information(graph_client, case, select=["id"], **arguments)


@pytest.mark.parametrize(
    ("case", "missing", "arguments"),
    [
        ("outlook_read", "message_id", {"user_id": "user", "message_id": ""}),
        ("calendar_read", "event_id", {"user_id": "user", "event_id": ""}),
        ("todo_read", "todo_task_id", {"user_id": "user", "todo_list_id": "list", "todo_task_id": ""}),
        ("planner_read", "planner_task_id", {"planner_task_id": ""}),
        ("drive_item_read", "drive_item_id", {"drive_id": "drive", "drive_item_id": ""}),
    ],
)
def test_item_read_rejects_a_missing_identifier_before_building_anything(graph_client, case, missing, arguments):
    from microsoft365.sdk_contract import build_item_request_information

    with pytest.raises(ValueError, match=f"{missing} is required"):
        build_item_request_information(graph_client, case, **arguments)


def test_item_read_rejects_an_unknown_case(graph_client):
    from microsoft365.sdk_contract import build_item_request_information

    with pytest.raises(ValueError, match="unknown item contract"):
        build_item_request_information(graph_client, "planner_delete", **ITEM_READ_ARGUMENTS)


def test_planner_plan_listing_covers_global_and_group_scoped_routes(graph_client):
    from microsoft365.sdk_contract import build_collection_request_information

    group_request = build_collection_request_information(
        graph_client, "planner_plans", group_id="group", owner_id="group", limit=7
    )

    assert group_request.http_method.value == "GET"
    assert urlsplit(group_request.url).path == "/groups/group/planner/plans"
    assert group_request.query_parameters == {"%24filter": "owner eq 'group'", "%24top": 7}

    global_request = build_collection_request_information(
        graph_client, "planner_plans", owner_id="group", limit=7
    )

    assert urlsplit(global_request.url).path == "/planner/plans"
    assert global_request.query_parameters == {"%24filter": "owner eq 'group'", "%24top": 7}


def test_planner_group_scoped_listing_rejects_a_blank_group_id(graph_client):
    from microsoft365.sdk_contract import build_collection_request_information

    with pytest.raises(ValueError, match="group_id is required"):
        build_collection_request_information(graph_client, "planner_plans", group_id="   ")


def test_teams_joined_teams_sends_no_odata_parameters_even_with_a_limit(graph_client):
    """joinedTeams supports no OData parameters -- $top included (WP8 rule)."""
    from microsoft365.sdk_contract import build_collection_request_information

    request = build_collection_request_information(
        graph_client, "teams_joined", user_id="user", limit=7
    )

    assert request.http_method.value == "GET"
    assert urlsplit(request.url).path == "/users/user/joinedTeams"
    assert request.query_parameters == {}
    assert request.url.endswith("/users/user/joinedTeams")


def test_teams_channels_allows_select_and_never_sets_top(graph_client):
    """channels allows $select but not $top (WP8 rule)."""
    from microsoft365.sdk_contract import build_collection_request_information

    request = build_collection_request_information(
        graph_client, "teams_channels", team_id="team", limit=7
    )

    assert urlsplit(request.url).path == "/teams/team/channels"
    assert request.query_parameters == {"%24select": ["id", "displayName"]}
    assert "%24top" not in request.query_parameters


SEARCH_CASES = ("sites_search", "drive_search", "drive_item_search")


def test_search_contract_exposes_exactly_the_verified_cases():
    import microsoft365.sdk_contract as sdk_contract

    assert sdk_contract.SEARCH_CASES == SEARCH_CASES


def test_sites_search_hits_root_sites_with_a_typed_search_parameter(graph_client):
    from microsoft365.sdk_contract import build_search_request_information

    request = build_search_request_information(graph_client, "sites_search", query="projeto")

    assert request.http_method.value == "GET"
    assert urlsplit(request.url).path == "/sites"
    assert request.query_parameters == {"%24search": "projeto"}
    assert request.url == "/sites?%24search=projeto"


def test_sites_search_leaves_url_encoding_to_the_generated_builder(graph_client):
    """$search is a plain query value, not an OData literal: no quote doubling here."""
    from microsoft365.sdk_contract import build_search_request_information

    request = build_search_request_information(graph_client, "sites_search", query="it's")

    assert request.query_parameters == {"%24search": "it's"}
    assert request.url == "/sites?%24search=it%27s"


def test_drive_search_uses_the_generated_search_method(graph_client):
    from microsoft365.sdk_contract import build_search_request_information

    request = build_search_request_information(
        graph_client, "drive_search", drive_id="drive", query="hello"
    )

    assert request.http_method.value == "GET"
    assert urlsplit(request.url).path == "/drives/drive/search(q='hello')"
    assert request.query_parameters == {}


def test_drive_item_search_uses_the_generated_item_search_method(graph_client):
    from microsoft365.sdk_contract import build_search_request_information

    request = build_search_request_information(
        graph_client, "drive_item_search", drive_id="drive", drive_item_id="item", query="hello"
    )

    assert urlsplit(request.url).path == "/drives/drive/items/item/search(q='hello')"


def test_drive_search_doubles_single_quotes_as_an_odata_literal(graph_client):
    """The generated builder URL-encodes `q` but does not escape OData quotes."""
    from microsoft365.sdk_contract import build_search_request_information

    request = build_search_request_information(
        graph_client, "drive_search", drive_id="drive", query="it's"
    )

    # doubled quote survives as `%27%27`; a bare escape would render `%27`
    assert urlsplit(request.url).path == "/drives/drive/search(q='it%27%27s')"


@pytest.mark.parametrize("case", SEARCH_CASES)
def test_search_contract_rejects_a_blank_query(graph_client, case):
    from microsoft365.sdk_contract import build_search_request_information

    with pytest.raises(ValueError, match="query is required"):
        build_search_request_information(
            graph_client, case, drive_id="drive", drive_item_id="item", query="   "
        )


def test_drive_search_requires_a_drive_id(graph_client):
    from microsoft365.sdk_contract import build_search_request_information

    with pytest.raises(ValueError, match="drive_id is required"):
        build_search_request_information(graph_client, "drive_search", query="hello")


def test_drive_item_search_requires_a_drive_item_id(graph_client):
    from microsoft365.sdk_contract import build_search_request_information

    with pytest.raises(ValueError, match="drive_item_id is required"):
        build_search_request_information(
            graph_client, "drive_item_search", drive_id="drive", query="hello"
        )


def test_search_contract_rejects_an_unknown_case(graph_client):
    from microsoft365.sdk_contract import build_search_request_information

    with pytest.raises(ValueError, match="unknown search contract"):
        build_search_request_information(graph_client, "site_delete", query="hello")
