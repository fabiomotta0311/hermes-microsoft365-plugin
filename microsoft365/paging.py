"""Bounded, origin-validated Microsoft Graph pagination (WP5).

Why this module exists
----------------------
When a graph collection does not fit into one response, the service returns an
``@odata.nextLink``. That URL is *service-controlled input*: following it blindly sends the
caller's bearer token to whichever origin the link names, and following it without a bound
turns a single request into an unbounded traversal. Both are refused here.

Rules enforced (each pinned by a test in ``tests/test_paging.py``)

1. **Validate before use.** A next link must be an absolute ``https`` URL whose host is
   exactly ``graph.microsoft.com`` (no subdomain, no user info, no port, no look-alike
   suffix), whose path is under ``/v1.0`` (never ``/beta``) and is exactly one of the
   expected resource paths, and which carries a query -- a link without one cannot advance
   the traversal. Fragments are refused. Nothing is rewritten: the validated string is the
   string that gets used.
2. **Follow through the generated builder.** The link is applied with the generated
   builder's own ``with_url()`` -- never string concatenation -- and the resulting request
   must reproduce the link verbatim before it is sent. The caller's first-page query
   configuration is deliberately *not* reapplied, so a followed request can only carry the
   query the service supplied.
3. **Never loop.** The identity (resource path + parsed query) of every request already
   fetched is remembered; a repeated identity stops the traversal instead of looping.
4. **Bound the traversal.** The item count never exceeds the caller's ``limit`` and the
   number of requests never exceeds :data:`GLOBAL_MAX_PAGES`; a caller-supplied ``max_pages``
   may lower that cap, never raise it.
5. **Report truthfully.** :class:`Continuation` states whether the result was truncated,
   whether a next link existed, why traversal stopped, and how many pages/items were
   fetched. A link that was refused is never echoed back to the caller or to the model.

Nothing here holds a credential, client or adapter in module state: the adapter is always a
parameter, and it must be the adapter injected into the client that produced ``builder``.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Sequence
from urllib.parse import parse_qsl, unquote, urlsplit

from microsoft365.execution import ExecutionError, execute_request, request_information_sender
from microsoft365.results import normalize_collection_page

GRAPH_HOST = "graph.microsoft.com"
GRAPH_VERSION_PATH = "/v1.0"
#: Hard ceiling on requests per traversal; a caller's ``max_pages`` can only lower it.
GLOBAL_MAX_PAGES = 10
#: Highest item count a single traversal may be asked for (the contract's ``$top`` range).
MAX_CALLER_LIMIT = 100

STOP_COMPLETE = "complete"
STOP_CALLER_LIMIT = "caller_limit"
STOP_PAGE_CAP = "page_cap"
STOP_CYCLE = "cycle"
STOP_REJECTED = "rejected_next_link"


class NextLinkRejected(ValueError):
    """A ``@odata.nextLink`` was refused before any request was sent.

    ``reason`` is a stable code that is safe to surface; the message never contains the
    refused URL, so a hostile link cannot leak back through an error string.
    """

    def __init__(self, reason: str, message: str):
        self.reason = str(reason)
        self.message = str(message)
        super().__init__(f"next_link_rejected({self.reason}): {self.message}")


@dataclass(frozen=True)
class Continuation:
    """Truthful report of how a traversal ended."""

    truncated: bool
    next_link_present: bool
    pages_fetched: int
    items_fetched: int
    stop_reason: str
    next_link: str | None = None
    rejection_reason: str | None = None


@dataclass(frozen=True)
class PagedResult:
    """One bounded traversal: the raw generated items plus a normalized envelope."""

    items: tuple
    continuation: Continuation
    payload: dict


def _resource_path(url: Any) -> str:
    """Resource path of a graph URL: no ``/v1.0`` prefix, no surrounding slashes, lowercase."""
    path = urlsplit(str(url or "")).path
    if path.startswith(GRAPH_VERSION_PATH + "/"):
        path = path[len(GRAPH_VERSION_PATH) + 1:]
    elif path == GRAPH_VERSION_PATH:
        path = ""
    return path.strip("/").lower()


def _expected_paths(expected_paths: Sequence[str] | None) -> tuple[str, ...]:
    if expected_paths is None:
        raise ValueError("expected_paths is required to validate a next link")
    normalized = []
    for candidate in expected_paths:
        path = _resource_path(candidate)
        if not path:
            raise ValueError("expected_paths entries must be non-empty graph resource paths")
        normalized.append(path)
    if not normalized:
        raise ValueError("expected_paths must contain at least one graph resource path")
    return tuple(dict.fromkeys(normalized))


def _query_identity(pairs) -> tuple:
    """Order-insensitive, duplicate-insensitive identity of a query.

    Kiota renders a typed configuration into **both** ``RequestInformation.url`` and
    ``RequestInformation.query_parameters``, so the two sources are unioned rather than
    concatenated -- otherwise the same request would have two different identities.
    """
    return tuple(sorted({(str(name), str(value)) for name, value in pairs}))


def _request_identity(request) -> tuple:
    """Identity of a request this plugin built: resource path plus parsed query."""
    parts = urlsplit(str(getattr(request, "url", "") or ""))
    query = list(parse_qsl(parts.query, keep_blank_values=True))
    for name, value in (getattr(request, "query_parameters", None) or {}).items():
        query.append((unquote(str(name)), str(value)))
    return (_resource_path(getattr(request, "url", "")), _query_identity(query))


def _link_identity(link: str) -> tuple:
    """Identity of a server-supplied link: resource path plus parsed query, order-insensitive."""
    parts = urlsplit(link)
    return (_resource_path(link), _query_identity(parse_qsl(parts.query, keep_blank_values=True)))


def validate_next_link(url: Any, *, expected_paths: Sequence[str]) -> str:
    """Return ``url`` unchanged when it is a safe graph continuation, else raise.

    ``expected_paths`` are resource paths relative to the graph version (``users/u/messages``)
    and are compared case-insensitively with any ``/v1.0`` prefix removed.
    """
    if not isinstance(url, str) or not url.strip():
        raise NextLinkRejected("not_a_url", "a next link must be a non-empty absolute URL")
    if url != url.strip():
        # Refused rather than trimmed: validation must never hand back a rewritten link.
        raise NextLinkRejected("not_a_url", "a next link must not be padded with whitespace")
    parts = urlsplit(url)
    if not parts.scheme or not parts.netloc:
        raise NextLinkRejected("not_a_url", "a next link must be an absolute URL")
    if parts.scheme.lower() != "https":
        raise NextLinkRejected("scheme", "a next link must use https")
    if parts.netloc.lower() != GRAPH_HOST:
        # One exact comparison covers subdomains, look-alike suffixes, user info and ports.
        raise NextLinkRejected("host", f"a next link must target {GRAPH_HOST} exactly")
    if not parts.path.startswith(GRAPH_VERSION_PATH + "/"):
        raise NextLinkRejected(
            "version_path", f"a next link must stay under {GRAPH_VERSION_PATH}"
        )
    if parts.fragment:
        raise NextLinkRejected("fragment", "a next link must not carry a fragment")
    if not parts.query:
        raise NextLinkRejected(
            "query_missing", "a next link without a query cannot advance the traversal"
        )
    if _resource_path(url) not in _expected_paths(expected_paths):
        raise NextLinkRejected("expected_path", "a next link must stay on an expected resource path")
    return url


def _limit(limit: Any) -> int:
    if type(limit) is not int or not 1 <= limit <= MAX_CALLER_LIMIT:
        raise ValueError(f"limit must be an integer between 1 and {MAX_CALLER_LIMIT}")
    return limit


def _page_cap(max_pages: Any) -> int:
    if max_pages is None:
        return GLOBAL_MAX_PAGES
    if type(max_pages) is not int or max_pages < 1:
        raise ValueError("max_pages must be a positive integer")
    return min(max_pages, GLOBAL_MAX_PAGES)


def _page_items(response) -> list:
    """The page's items, or a typed failure when the response is not a collection at all."""
    if not hasattr(response, "value") and not hasattr(response, "odata_next_link"):
        raise ExecutionError("service_error", "graph collection response has no page shape")
    return list(getattr(response, "value", None) or [])


def _follow(builder, link: str):
    """Bind the generated builder to a validated link. Builds the request, never sends it."""
    followed = builder.with_url(link)
    request = request_information_sender(followed, method="GET")
    if request.url != link:
        raise NextLinkRejected(
            "not_preserved", "the generated builder did not reproduce the next link verbatim"
        )
    return followed


def paginate(
    builder,
    *,
    adapter,
    limit: int,
    configuration=None,
    expected_paths: Sequence[str] | None = None,
    max_pages: int | None = None,
    max_string: int = 4000,
) -> PagedResult:
    """Follow ``@odata.nextLink`` safely, bounded by ``limit`` and the global page cap.

    ``configuration`` applies to the **first** request only -- the caller's own ``$top``,
    ``$filter`` or ``$select``. Followed links are never rewritten with it: a continuation
    carries exactly the query the service put in it.

    ``expected_paths`` defaults to the first request's own resource path, so a link may only
    continue the collection the caller asked for. ``max_pages`` may lower, never raise, the
    global cap.
    """
    budget = _limit(limit)
    page_cap = _page_cap(max_pages)
    first = request_information_sender(builder, method="GET", configuration=configuration)
    allowed = _expected_paths((first.url,) if expected_paths is None else expected_paths)
    seen = {_request_identity(first)}

    items: list = []
    observed_count = None
    pages = 0
    followed_link = None
    response = execute_request(builder, method="GET", configuration=configuration, adapter=adapter)

    while True:
        pages += 1
        if observed_count is None and getattr(response, "odata_count", None) is not None:
            observed_count = response.odata_count
        page_values = _page_items(response)
        remaining = budget - len(items)
        taken = page_values[:remaining] if remaining > 0 else []
        items.extend(taken)
        leftover = len(page_values) > len(taken)
        raw_link = getattr(response, "odata_next_link", None)
        next_link_present = isinstance(raw_link, str) and bool(raw_link.strip())
        raw_link = raw_link if next_link_present else None

        validated = None
        rejected = None
        if raw_link is not None:
            try:
                validated = validate_next_link(raw_link, expected_paths=allowed)
            except NextLinkRejected as failure:
                rejected = failure.reason

        if raw_link is None and not leftover:
            stop_reason = STOP_COMPLETE
        elif len(items) >= budget:
            stop_reason = STOP_CALLER_LIMIT
        elif rejected is not None:
            stop_reason = STOP_REJECTED
        elif _link_identity(validated) in seen:
            stop_reason = STOP_CYCLE
        elif pages >= page_cap:
            stop_reason = STOP_PAGE_CAP
        else:
            try:
                followed = _follow(builder, validated)
            except NextLinkRejected as failure:
                stop_reason, rejected, validated = STOP_REJECTED, failure.reason, None
            else:
                seen.add(_link_identity(validated))
                followed_link = validated
                response = execute_request(followed, method="GET", adapter=adapter)
                continue

        # Only a link that passed validation and was not followed is handed back. When the
        # caller's limit cut a page short, that page's *own* link would skip the items that
        # were dropped, so the continuation handed back is the URL of the partial page.
        if stop_reason == STOP_CALLER_LIMIT and leftover:
            next_link = followed_link
        elif stop_reason in (STOP_CALLER_LIMIT, STOP_PAGE_CAP) and validated is not None:
            next_link = validated
        else:
            next_link = None
        break

    continuation = Continuation(
        truncated=stop_reason != STOP_COMPLETE,
        next_link_present=next_link_present,
        pages_fetched=pages,
        items_fetched=len(items),
        stop_reason=stop_reason,
        next_link=next_link,
        rejection_reason=rejected,
    )
    return PagedResult(
        items=tuple(items),
        continuation=continuation,
        payload=normalize_collection_page(
            items,
            odata_count=observed_count,
            next_link=next_link,
            max_items=budget,
            max_string=max_string,
        ),
    )
