"""WP5 — safe, bounded Microsoft Graph pagination, proven offline.

The transport double here is deliberately *not* ``StrictTransportAdapter`` from
``tests/test_sdk_contract.py``: pagination must actually receive real generated collection
models. ``PagedGraphAdapter`` therefore returns the exact pages the test supplied, records
every ``RequestInformation`` it was handed, and raises for any request the test did not
supply — so a paginator that follows a link the test never expected, or that rewrites a
query it should have kept verbatim, cannot pass by accident.

Nothing in this file constructs a URL by hand for the plugin to consume: request
expectations are keyed by the request the generated builder produced (path + typed query
parameters for the first page, path + query string for followed links).
"""
from __future__ import annotations

import base64
import json
from typing import Any
from urllib.parse import parse_qsl, unquote, urlsplit

import pytest
from kiota_abstractions.base_request_configuration import RequestConfiguration
from kiota_abstractions.request_adapter import RequestAdapter
from kiota_serialization_json.json_serialization_writer_factory import JsonSerializationWriterFactory
from msgraph import GraphServiceClient
from msgraph.generated.models.message import Message
from msgraph.generated.models.message_collection_response import MessageCollectionResponse
from msgraph.generated.models.todo_task import TodoTask
from msgraph.generated.models.todo_task_collection_response import TodoTaskCollectionResponse
from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder

from microsoft365.execution import ExecutionError
from microsoft365.paging import (
    GLOBAL_MAX_PAGES,
    NextLinkRejected,
    paginate,
    validate_next_link,
)
from tests.test_sdk_contract import StrictTransportAdapter

FIRST_PATH = "/users/user/messages"
RESOURCE_PATH = "/v1.0/users/user/messages"
FIRST_CONFIG = {"top": 2, "filter": "contains(subject,'hello')"}
# The first request carries the caller's query in typed ``query_parameters`` (Kiota keeps
# the URL relative and puts OData values in the dict), the followed link carries its own
# query inside the raw URL. Both forms are keyed the same way below.
FIRST_QUERY = {"$filter": "contains(subject,'hello')", "$top": "2"}


def request_key(request_info: Any) -> tuple:
    """Canonical, transport-independent identity of one request.

    The generated builder puts a typed configuration into both the URL string and the
    ``query_parameters`` mapping, so the pairs are unioned (a set) instead of concatenated.
    """
    parts = urlsplit(request_info.url)
    query = list(parse_qsl(parts.query, keep_blank_values=True))
    for name, value in (getattr(request_info, "query_parameters", None) or {}).items():
        query.append((unquote(str(name)), str(value)))
    return (parts.path, tuple(sorted(set(query))))


def page(items: list, *, next_link: str | None = None, count: int | None = None):
    """A real generated collection response."""
    return MessageCollectionResponse(
        value=[Message(id=item) for item in items],
        odata_count=count,
        odata_next_link=next_link,
    )


def link(token: str, *, path: str = RESOURCE_PATH, query: str = "") -> str:
    suffix = query or f"%24skiptoken={token}"
    return f"https://graph.microsoft.com{path}?{suffix}"


class PagedGraphAdapter(RequestAdapter):
    """Concrete Kiota boundary double returning pre-built real generated models.

    Real methods only: no ``__getattr__``, no ``MagicMock``, no ``SimpleNamespace``. It
    answers exactly the pages the test supplied (keyed by request identity) and raises for
    every URL the test did not supply.
    """

    def __init__(self, pages: dict, *, base_url: str = "https://graph.microsoft.com/v1.0"):
        self.base_url = base_url
        self._pages = dict(pages)
        self.requests: list = []
        self._serialization = JsonSerializationWriterFactory()

    def enable_backing_store(self, backing_store_factory=None) -> None:
        return None

    def get_serialization_writer_factory(self):
        return self._serialization

    async def send_async(self, request_info: Any, parsable_factory: Any, error_map=None, **kwargs: Any):
        key = request_key(request_info)
        self.requests.append(request_info)
        if key not in self._pages:
            raise AssertionError(f"unexpected request: {key}")
        return self._pages[key]

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

    @property
    def requested_keys(self) -> list:
        return [request_key(request) for request in self.requests]


def messages_builder(adapter: RequestAdapter):
    return GraphServiceClient(request_adapter=adapter).users.by_user_id("user").messages


def first_page_configuration(builder) -> RequestConfiguration:
    query_type = getattr(type(builder), f"{type(builder).__name__}GetQueryParameters")
    return RequestConfiguration(query_parameters=query_type(**FIRST_CONFIG))


def run_pagination(pages: dict, *, limit: int, **kwargs):
    adapter = PagedGraphAdapter(pages)
    builder = messages_builder(adapter)
    result = paginate(
        builder,
        adapter=adapter,
        limit=limit,
        configuration=first_page_configuration(builder),
        **kwargs,
    )
    return adapter, result


# ---------------------------------------------------------------- next-link validation


@pytest.mark.parametrize(
    ("candidate", "reason"),
    [
        ("http://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x", "scheme"),
        ("ftp://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x", "scheme"),
        ("https://graph.microsoft.com.evil.tld/v1.0/users/user/messages?%24skiptoken=x", "host"),
        ("https://evilgraph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x", "host"),
        ("https://evil.tld/v1.0/users/user/messages?%24skiptoken=x", "host"),
        ("https://graph.microsoft.com.:443/v1.0/users/user/messages?%24skiptoken=x", "host"),
        ("https://graph.microsoft.com.xn--tld/v1.0/users/user/messages?%24skiptoken=x", "host"),
        ("https://graph.microsoft.com:8443/v1.0/users/user/messages?%24skiptoken=x", "host"),
        ("https://graph.microsoft.com@evil.tld/v1.0/users/user/messages?%24skiptoken=x", "host"),
        ("https://user:pass@graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x", "host"),
        ("https://graph.microsoft.com/beta/users/user/messages?%24skiptoken=x", "version_path"),
        ("https://graph.microsoft.com/users/user/messages?%24skiptoken=x", "version_path"),
        ("https://graph.microsoft.com/v1.0?%24skiptoken=x", "version_path"),
        ("https://graph.microsoft.com/v1.0/users/other/messages?%24skiptoken=x", "expected_path"),
        ("https://graph.microsoft.com/v1.0/users/user/messages/send?%24skiptoken=x", "expected_path"),
        ("https://graph.microsoft.com/v1.0/users/user/messages", "query_missing"),
        ("https://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x#fragment", "fragment"),
        ("/v1.0/users/user/messages?%24skiptoken=x", "not_a_url"),
        ("https://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x ", "not_a_url"),
        (" https://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x", "not_a_url"),
        ("", "not_a_url"),
        (None, "not_a_url"),
        (1234, "not_a_url"),
    ],
)
def test_next_link_must_be_https_graph_origin_and_expected_path(candidate, reason):
    """Origin, version path, expected resource path and fragment are all enforced."""
    with pytest.raises(NextLinkRejected) as failure:
        validate_next_link(candidate, expected_paths=("users/user/messages",))

    assert failure.value.reason == reason


@pytest.mark.parametrize(
    "candidate",
    [
        "https://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x",
        "https://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x&%24top=3",
        "https://GRAPH.MICROSOFT.COM/v1.0/users/user/messages?%24skiptoken=x",
        "https://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x&%24filter=contains(subject%2C'hello')",
    ],
)
def test_valid_next_link_is_returned_verbatim(candidate):
    """Validation never rewrites the link: the caller follows exactly what Graph sent."""
    assert validate_next_link(candidate, expected_paths=("users/user/messages",)) == candidate


def test_next_link_accepts_an_expected_path_prefixed_with_the_graph_version():
    candidate = "https://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x"
    assert validate_next_link(candidate, expected_paths=("/v1.0/users/user/messages",)) == candidate


def test_next_link_validation_requires_at_least_one_expected_path():
    with pytest.raises(ValueError):
        validate_next_link(
            "https://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x", expected_paths=()
        )


def todo_tasks_builder(adapter: RequestAdapter):
    return (
        GraphServiceClient(request_adapter=adapter)
        .users.by_user_id("user")
        .todo.lists.by_todo_task_list_id("list")
        .tasks
    )


def test_pagination_derives_the_expected_path_from_the_callers_own_request():
    """The tasks collection may only be continued by a link for the tasks collection."""
    tasks_path = "/users/user/todo/lists/list/tasks"
    pages = {
        (tasks_path, ()): TodoTaskCollectionResponse(
            value=[TodoTask(id="t1")],
            odata_next_link="https://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x",
        )
    }
    adapter = PagedGraphAdapter(pages)
    builder = todo_tasks_builder(adapter)

    result = paginate(builder, adapter=adapter, limit=5)

    assert len(adapter.requests) == 1
    assert result.continuation.stop_reason == "rejected_next_link"
    assert result.continuation.rejection_reason == "expected_path"


def test_pagination_follows_a_link_that_stays_on_the_callers_collection():
    tasks_path = "/users/user/todo/lists/list/tasks"
    link_on_tasks = f"https://graph.microsoft.com/v1.0{tasks_path}?%24skiptoken=p2"
    pages = {
        (tasks_path, ()): TodoTaskCollectionResponse(
            value=[TodoTask(id="t1")], odata_next_link=link_on_tasks
        ),
        (f"/v1.0{tasks_path}", (("$skiptoken", "p2"),)): TodoTaskCollectionResponse(
            value=[TodoTask(id="t2")]
        ),
    }
    adapter = PagedGraphAdapter(pages)
    builder = todo_tasks_builder(adapter)

    result = paginate(builder, adapter=adapter, limit=5)

    assert len(adapter.requests) == 2
    assert [task.id for task in result.items] == ["t1", "t2"]
    assert result.continuation.stop_reason == "complete"


def test_pagination_rejects_a_related_collection_link_the_caller_did_not_allow():
    pages = {
        (FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(
            ["m1"], next_link=link("p2", path="/v1.0/users/user/mailFolders/inbox/messages")
        )
    }

    adapter, result = run_pagination(pages, limit=10)

    assert len(adapter.requests) == 1
    assert result.continuation.stop_reason == "rejected_next_link"
    assert result.continuation.rejection_reason == "expected_path"


def test_pagination_continues_a_related_collection_when_the_caller_allows_it():
    inbox = "/v1.0/users/user/mailFolders/inbox/messages"
    pages = {
        (FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(
            ["m1"], next_link=link("p2", path=inbox)
        ),
        (inbox, (("$skiptoken", "p2"),)): page(["m2"]),
    }

    adapter, result = run_pagination(
        pages,
        limit=10,
        expected_paths=("users/user/messages", "users/user/mailFolders/inbox/messages"),
    )

    assert len(adapter.requests) == 2
    assert [item.id for item in result.items] == ["m1", "m2"]
    assert result.continuation.stop_reason == "complete"


# ------------------------------------------------------------------ bounds and limits


def test_pagination_stops_at_caller_limit_and_never_exceeds_it():
    pages = {
        (FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(["m1", "m2"], next_link=link("p2")),
        (RESOURCE_PATH, (("$skiptoken", "p2"),)): page(
            ["m3", "m4", "m5"], next_link=link("p3")
        ),
        (RESOURCE_PATH, (("$skiptoken", "p3"),)): page(["m6"]),
    }

    adapter, result = run_pagination(pages, limit=3)

    assert [item.id for item in result.items] == ["m1", "m2", "m3"]
    assert result.continuation.items_fetched == 3
    assert result.continuation.pages_fetched == 2
    assert result.continuation.truncated is True
    assert result.continuation.next_link_present is True
    assert result.continuation.stop_reason == "caller_limit"
    # the third page was never requested: the caller limit bounds the traversal
    assert len(adapter.requests) == 2
    assert result.continuation.next_link == link("p2")


def test_pagination_reports_partial_page_at_the_caller_limit_without_a_next_link():
    pages = {
        (FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(["m1", "m2", "m3", "m4", "m5"])
    }

    adapter, result = run_pagination(pages, limit=3)

    assert [item.id for item in result.items] == ["m1", "m2", "m3"]
    assert len(adapter.requests) == 1
    assert result.continuation.stop_reason == "caller_limit"
    assert result.continuation.truncated is True
    assert result.continuation.next_link_present is False
    assert result.continuation.next_link is None
    assert "@odata.nextLink" not in result.payload


def test_pagination_stops_at_global_page_cap_and_never_exceeds_it():
    chain = 15
    pages = {(FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(["m0"], next_link=link("p2"))}
    for index in range(2, chain + 1):
        token = f"p{index}"
        pages[(RESOURCE_PATH, (("$skiptoken", token),))] = page(
            [f"m{index}"], next_link=link(f"p{index + 1}") if index < chain else None
        )

    adapter, result = run_pagination(pages, limit=100, max_pages=999)

    assert result.continuation.pages_fetched == GLOBAL_MAX_PAGES
    assert len(adapter.requests) == GLOBAL_MAX_PAGES
    assert len(result.items) == GLOBAL_MAX_PAGES
    assert result.continuation.stop_reason == "page_cap"
    assert result.continuation.truncated is True
    assert result.continuation.next_link_present is True
    assert result.continuation.next_link == link("p11")


def test_pagination_clamps_caller_page_cap_to_the_global_cap():
    pages = {(FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(["m0"], next_link=link("p2"))}
    for index in range(2, 8):
        pages[(RESOURCE_PATH, (("$skiptoken", f"p{index}"),))] = page(
            [f"m{index}"], next_link=link(f"p{index + 1}")
        )

    adapter, result = run_pagination(pages, limit=100, max_pages=3)

    assert result.continuation.pages_fetched == 3
    assert len(adapter.requests) == 3
    assert result.continuation.stop_reason == "page_cap"
    assert result.continuation.next_link == link("p4")


def test_pagination_detects_a_link_that_points_back_at_the_first_request():
    pages = {
        (FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(
            ["m1"],
            next_link=(
                "https://graph.microsoft.com" + RESOURCE_PATH
                + "?%24filter=contains(subject%2C'hello')&%24top=2"
            ),
        )
    }

    adapter, result = run_pagination(pages, limit=100)

    assert len(adapter.requests) == 1
    assert result.continuation.stop_reason == "cycle"
    assert result.continuation.pages_fetched == 1
    assert result.continuation.truncated is True
    assert result.continuation.next_link_present is True
    assert result.continuation.next_link is None
    assert "@odata.nextLink" not in result.payload


def test_pagination_detects_a_cycle_between_two_links():
    pages = {
        (FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(["m1"], next_link=link("loop")),
        (RESOURCE_PATH, (("$skiptoken", "loop"),)): page(["m2"], next_link=link("loop")),
    }

    adapter, result = run_pagination(pages, limit=100)

    assert len(adapter.requests) == 2
    assert [item.id for item in result.items] == ["m1", "m2"]
    assert result.continuation.stop_reason == "cycle"
    assert result.continuation.truncated is True
    assert result.continuation.next_link is None


def test_pagination_detects_a_cycle_when_only_the_query_order_changes():
    pages = {
        (FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(
            ["m1"], next_link=link("loop", query="%24skiptoken=loop&%24top=3")
        ),
        (RESOURCE_PATH, (("$skiptoken", "loop"), ("$top", "3"))): page(
            ["m2"], next_link=link("loop", query="%24top=3&%24skiptoken=loop")
        ),
    }

    adapter, result = run_pagination(pages, limit=100)

    assert len(adapter.requests) == 2
    assert result.continuation.stop_reason == "cycle"


# ------------------------------------------------------------- truthful continuation


def test_pagination_reports_completion_truthfully():
    pages = {
        (FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(
            ["m1", "m2"], next_link=link("p2"), count=4
        ),
        (RESOURCE_PATH, (("$skiptoken", "p2"),)): page(["m3", "m4"], count=4),
    }

    adapter, result = run_pagination(pages, limit=10)

    assert [item.id for item in result.items] == ["m1", "m2", "m3", "m4"]
    assert len(adapter.requests) == 2
    continuation = result.continuation
    assert (continuation.truncated, continuation.next_link_present) == (False, False)
    assert (continuation.pages_fetched, continuation.items_fetched) == (2, 4)
    assert continuation.stop_reason == "complete"
    assert continuation.next_link is None
    assert continuation.rejection_reason is None
    assert result.payload["@odata.count"] == 4
    assert "@odata.nextLink" not in result.payload
    assert [item["id"] for item in result.payload["value"]] == ["m1", "m2", "m3", "m4"]


def test_pagination_reports_exact_fit_as_complete():
    pages = {
        (FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(["m1", "m2"], next_link=link("p2")),
        (RESOURCE_PATH, (("$skiptoken", "p2"),)): page(["m3", "m4"]),
    }

    _, result = run_pagination(pages, limit=4)

    assert result.continuation.truncated is False
    assert result.continuation.stop_reason == "complete"
    assert result.continuation.items_fetched == 4


def test_pagination_reports_an_empty_collection_as_complete():
    pages = {(FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page([])}

    adapter, result = run_pagination(pages, limit=10)

    assert len(adapter.requests) == 1
    assert result.items == ()
    assert result.continuation.stop_reason == "complete"
    assert result.continuation.truncated is False
    assert result.payload["value"] == []


@pytest.mark.parametrize(
    ("foreign", "reason"),
    [
        ("https://evil.tld/v1.0/users/user/messages?%24skiptoken=x", "host"),
        ("https://graph.microsoft.com.evil.tld/v1.0/users/user/messages?%24skiptoken=x", "host"),
        ("http://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x", "scheme"),
        ("https://graph.microsoft.com/beta/users/user/messages?%24skiptoken=x", "version_path"),
        ("https://graph.microsoft.com/v1.0/users/other/messages?%24skiptoken=x", "expected_path"),
        ("https://graph.microsoft.com/v1.0/users/user/messages", "query_missing"),
    ],
)
def test_pagination_never_follows_a_rejected_next_link(foreign, reason):
    pages = {(FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(["m1"], next_link=foreign)}

    adapter, result = run_pagination(pages, limit=100)

    assert len(adapter.requests) == 1
    continuation = result.continuation
    assert continuation.stop_reason == "rejected_next_link"
    assert continuation.rejection_reason == reason
    assert continuation.truncated is True
    assert continuation.next_link_present is True
    assert continuation.next_link is None
    # the refused origin must not be echoed back to the model either
    assert "@odata.nextLink" not in result.payload
    assert "evil" not in json.dumps(result.payload)
    assert "beta" not in json.dumps(result.payload)


def test_pagination_stays_bound_when_the_caller_limit_hits_a_foreign_link():
    pages = {
        (FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(
            ["m1", "m2"], next_link="https://evil.tld/v1.0/users/user/messages?%24skiptoken=x"
        )
    }

    adapter, result = run_pagination(pages, limit=2)

    assert len(adapter.requests) == 1
    assert result.continuation.stop_reason == "caller_limit"
    assert result.continuation.rejection_reason == "host"
    assert result.continuation.next_link is None
    assert "@odata.nextLink" not in result.payload


# ------------------------------------------------------- follow mechanics and safety


def test_pagination_follows_links_through_the_generated_builder_without_rewriting_the_query(
    monkeypatch,
):
    pages = {
        (FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(["m1"], next_link=link("p2")),
        (RESOURCE_PATH, (("$skiptoken", "p2"),)): page(
            ["m2"], next_link=link("p3", query="%24skiptoken=p3&%24top=3")
        ),
        (RESOURCE_PATH, (("$skiptoken", "p3"), ("$top", "3"))): page(["m3"]),
    }
    followed: list = []
    original = MessagesRequestBuilder.with_url

    def spy(self, raw_url):
        followed.append(raw_url)
        return original(self, raw_url)

    monkeypatch.setattr(MessagesRequestBuilder, "with_url", spy)

    adapter, result = run_pagination(pages, limit=10)

    assert followed == [link("p2"), link("p3", query="%24skiptoken=p3&%24top=3")]
    assert [item.id for item in result.items] == ["m1", "m2", "m3"]
    assert len(adapter.requests[1:]) == len(followed)
    for request, expected in zip(adapter.requests[1:], followed):
        assert request.http_method.value == "GET"
        # the followed request is exactly the link, and the caller's own query is gone
        assert dict(request.query_parameters) == {}
        assert request.url == expected
        assert "$filter" not in request.url
    assert adapter.requests[0].url.startswith(FIRST_PATH)
    assert request_key(adapter.requests[0]) == (FIRST_PATH, tuple(sorted(FIRST_QUERY.items())))


def test_pagination_refuses_a_link_the_builder_does_not_reproduce(monkeypatch):
    """A builder that mangles the link must stop the traversal, not fetch a guessed URL."""
    pages = {(FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(["m1"], next_link=link("p2"))}

    def mangled(self, raw_url):
        return MessagesRequestBuilder(
            self.request_adapter, "/v1.0/users/user/messages?%24skiptoken=other"
        )

    monkeypatch.setattr(MessagesRequestBuilder, "with_url", mangled)

    adapter, result = run_pagination(pages, limit=10)

    assert len(adapter.requests) == 1
    assert result.continuation.stop_reason == "rejected_next_link"
    assert result.continuation.rejection_reason == "not_preserved"
    assert result.continuation.next_link is None


def test_pagination_rejects_invalid_limit_and_page_cap():
    pages = {(FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): page(["m1"])}
    for limit in (0, -1, 101, 1000, True, "3", 3.0, None):
        adapter = PagedGraphAdapter(pages)
        builder = messages_builder(adapter)
        with pytest.raises(ValueError):
            paginate(
                builder,
                adapter=adapter,
                limit=limit,
                configuration=first_page_configuration(builder),
            )
        assert adapter.requests == []
    for max_pages in (0, -1, True, "2", 2.5):
        adapter = PagedGraphAdapter(pages)
        builder = messages_builder(adapter)
        with pytest.raises(ValueError):
            paginate(
                builder,
                adapter=adapter,
                limit=3,
                configuration=first_page_configuration(builder),
                max_pages=max_pages,
            )
        assert adapter.requests == []


def test_pagination_fails_closed_when_the_builder_is_not_bound_to_the_injected_adapter():
    stray = PagedGraphAdapter({})
    builder = messages_builder(stray)

    with pytest.raises(ExecutionError) as failure:
        paginate(builder, adapter=StrictTransportAdapter(), limit=3)

    assert failure.value.category == "configuration_error"
    assert stray.requests == []


def test_pagination_rejects_a_response_that_is_not_a_collection():
    pages = {(FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): Message(id="m1")}

    with pytest.raises(ExecutionError) as failure:
        run_pagination(pages, limit=3)

    assert failure.value.category == "service_error"


# ----------------------------------------------------------------- normalization reuse


def test_pagination_normalization_preserves_redaction_and_bounds():
    secret = "sentinel-client-secret-value"
    long_subject = "x" * 6000
    pages = {
        (FIRST_PATH, tuple(sorted(FIRST_QUERY.items()))): MessageCollectionResponse(
            value=[
                Message(id="m1", additional_data={"client_secret": secret}),
                Message(id="m2", subject=long_subject),
            ]
        )
    }

    _, result = run_pagination(pages, limit=10)

    payload = json.dumps(result.payload)
    assert secret not in payload
    assert result.payload["value"][0]["additional_data"]["client_secret"] == "[REDACTED]"
    assert result.payload["value"][1]["subject"].endswith("...[TRUNCATED]")
    # the raw generated models handed to the caller are untouched by normalization
    assert result.items[1].subject == long_subject


def test_collection_page_envelope_keeps_value_count_and_next_link():
    from microsoft365.results import normalize_collection_page

    envelope = normalize_collection_page(
        [Message(id="m1"), Message(id="m2")],
        odata_count=7,
        next_link="https://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=x",
    )

    assert [item["id"] for item in envelope["value"]] == ["m1", "m2"]
    assert envelope["@odata.count"] == 7
    assert envelope["@odata.nextLink"].endswith("%24skiptoken=x")

    assert normalize_collection_page(None) == {"value": []}
    assert normalize_collection_page([]) == {"value": []}
    bounded = normalize_collection_page(
        [Message(id=str(index)) for index in range(5)], max_items=2
    )
    assert [item["id"] for item in bounded["value"]] == ["0", "1"]


def test_collection_page_envelope_bounds_continuation_metadata():
    """Count and link are bounded like every other scalar that reaches the model."""
    from microsoft365.results import normalize_collection_page

    oversized_count = normalize_collection_page([], odata_count="9" * 5000)
    assert oversized_count["@odata.count"].endswith("...[TRUNCATED]")

    oversized_link = normalize_collection_page(
        [], next_link="https://graph.microsoft.com/v1.0/users/user/messages?%24skiptoken=" + "x" * 5000
    )
    assert oversized_link["@odata.nextLink"].endswith("...[TRUNCATED]")


def test_collection_page_envelope_keeps_binary_artifacts_out_of_string_truncation():
    from microsoft365.results import BinaryResult, normalize_collection_page

    payload = b"y" * 8192
    encoded = base64.b64encode(payload).decode("ascii")
    envelope = normalize_collection_page(
        [BinaryResult(content_base64=encoded, content_type="application/octet-stream", size=len(payload))]
    )

    assert envelope["value"][0]["content_base64"] == encoded
    assert base64.b64decode(envelope["value"][0]["content_base64"], validate=True) == payload
