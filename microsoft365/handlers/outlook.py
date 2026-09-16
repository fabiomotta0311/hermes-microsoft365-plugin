"""Outlook handlers: search, read, create_draft and send (WP6).

Each handler builds the request of its operation from the arguments the administrative
argument contract declares, on the generated SDK builders, and is pinned byte-for-byte
against the verified request-information contract of the same operation
(``microsoft365/sdk_contract.py``; see ``tests/test_handlers_outlook_calendar.py``).

Action-specific rules this module owns (what the shared argument contract cannot express)
----------------------------------------------------------------------------------------
* ``outlook.search``: ``query`` and ``filter`` are two different filters, so supplying both is
  refused instead of silently dropping one; ``select`` is typed, so it reaches the request as
  ``$select`` and never as a hand-built query string.
* ``outlook.create_draft``: a draft is **composed**, so it always needs a body and recipients;
  ``message_id`` is the alternate path of ``outlook.send`` and cannot stand in for either.
* ``outlook.send``: exactly one form -- an existing draft (``message_id``) or a composed
  message (``subject`` plus ``to_recipients``) -- because mixing them would send one silently.
  ``save_to_sent_items`` is sent only when the caller supplied it, so the service default is
  never asserted on the caller's behalf.
"""
from __future__ import annotations

from typing import Mapping

from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.message import Message
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
    SendMailPostRequestBody,
)

from ..execution import execute_request
from ..paging import paginate
from . import (
    HandlerContext,
    collection_payload,
    handler,
    identifier,
    item_budget,
    open_graph_client,
    recipients,
    refuse,
    select,
    success_payload,
    text,
    typed_configuration,
)

SEARCH = "outlook.search"
READ = "outlook.read"
CREATE_DRAFT = "outlook.create_draft"
SEND = "outlook.send"


def search_filter(arguments: Mapping, *, operation: str) -> str | None:
    """The OData ``$filter`` of a message search, or nothing when the caller sent no search.

    ``query`` is the containment filter the verified collection contract defines
    (``contains(subject,'…')``, with the quote-doubling OData requires); ``filter`` is the
    caller's own expression and is passed through untouched.
    """
    query = arguments.get("query")
    supplied = arguments.get("filter")
    if query is not None and supplied is not None:
        refuse(
            "validation_error",
            f"{operation}: query and filter are two different filters; supplying both would "
            "silently drop one",
        )
    if supplied is not None:
        if not isinstance(supplied, str) or not supplied.strip():
            refuse(
                "validation_error",
                f"{operation}: filter must be a non-blank OData filter expression",
            )
        return supplied
    if query is not None:
        if not isinstance(query, str) or not query.strip():
            refuse("validation_error", f"{operation}: query must be a non-blank search string")
        return "contains(subject,'" + query.replace("'", "''") + "')"
    return None


def message_model(arguments: Mapping, *, operation: str, require_body: bool) -> Message:
    """The real generated ``Message`` a composed call sends."""
    subject = text(arguments, "subject", operation=operation)
    if subject is None:
        refuse("validation_error", f"{operation}: subject is required")
    content = text(arguments, "body", operation=operation)
    if require_body and content is None:
        refuse(
            "validation_error",
            f"{operation}: a draft needs a body; use message_id to send content that already "
            "exists",
        )
    return Message(
        subject=subject,
        body=ItemBody(content_type=BodyType.Text, content=content) if content is not None else None,
        to_recipients=recipients(arguments, "to_recipients", operation=operation, required=True),
        cc_recipients=recipients(arguments, "cc_recipients", operation=operation),
        bcc_recipients=recipients(arguments, "bcc_recipients", operation=operation),
    )


# ---------------------------------------------------------------- reads


@handler("outlook", "search")
def search(arguments: Mapping, context: HandlerContext) -> dict:
    """List one mailbox's messages, bounded by the caller's item budget."""
    user_id = identifier(arguments, "user_id", operation=SEARCH)
    filters = search_filter(arguments, operation=SEARCH)
    fields = select(arguments, operation=SEARCH)
    limit = item_budget(arguments, operation=SEARCH)

    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).messages
    configuration = typed_configuration(builder, top=limit, filter=filters, select=fields)
    paged = paginate(builder, adapter=adapter, limit=limit, configuration=configuration)
    return collection_payload(SEARCH, paged)


@handler("outlook", "read")
def read(arguments: Mapping, context: HandlerContext) -> dict:
    """Read one message, with the caller's ``$select`` and nothing else."""
    user_id = identifier(arguments, "user_id", operation=READ)
    message_id = identifier(arguments, "message_id", operation=READ)
    fields = select(arguments, operation=READ)

    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).messages.by_message_id(message_id)
    configuration = typed_configuration(builder, select=fields)
    response = execute_request(builder, method="GET", configuration=configuration, adapter=adapter)
    return success_payload(READ, response)


# ---------------------------------------------------------------- writes
#
# The four writes below are implemented, exercised offline through the seam (method, URL,
# body and If-Match are pinned against the verified contract) and deliberately **not
# executable**: while the generic Hermes approval-order defect (CORE-1/CORE-2) can let a hook
# change arguments after approval was evaluated against different ones, a write must not be
# reachable from the model surface at all (R5). They are not dead code -- the registry flip of
# WP9 is the only thing between them and execution.


@handler("outlook", "create_draft")
def create_draft(arguments: Mapping, context: HandlerContext) -> dict:
    """Create a draft message, which always carries a body and recipients."""
    user_id = identifier(arguments, "user_id", operation=CREATE_DRAFT)
    if arguments.get("message_id") is not None:
        refuse(
            "validation_error",
            f"{CREATE_DRAFT}: a draft is composed from its own body and recipients, so "
            "message_id cannot stand in for either of them",
        )
    model = message_model(arguments, operation=CREATE_DRAFT, require_body=True)

    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).messages
    response = execute_request(builder, method="POST", body=model, adapter=adapter)
    return success_payload(CREATE_DRAFT, response)


@handler("outlook", "send")
def send(arguments: Mapping, context: HandlerContext) -> dict:
    """Send a composed message, or send a draft that already exists."""
    user_id = identifier(arguments, "user_id", operation=SEND)
    message_id = arguments.get("message_id")
    composed = arguments.get("subject") is not None or arguments.get("to_recipients") is not None
    if message_id is not None and composed:
        refuse(
            "validation_error",
            f"{SEND}: send either an existing draft (message_id) or a composed message "
            "(subject and to_recipients); supplying both would send one of them silently",
        )
    if message_id is None and arguments.get("subject") is None:
        refuse(
            "validation_error",
            f"{SEND}: a send needs either message_id or a composed message with subject and "
            "to_recipients",
        )

    if message_id is not None:
        draft_id = identifier(arguments, "message_id", operation=SEND)
        client, adapter = open_graph_client(context)
        builder = client.users.by_user_id(user_id).messages.by_message_id(draft_id).send
        execute_request(builder, method="POST", adapter=adapter)
        return success_payload(SEND, {"accepted": True, "form": "existing"})

    save_to_sent_items = arguments.get("save_to_sent_items")
    if save_to_sent_items is not None and type(save_to_sent_items) is not bool:
        refuse(
            "validation_error",
            f"{SEND}: save_to_sent_items must be a real boolean",
        )
    model = message_model(arguments, operation=SEND, require_body=False)

    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).send_mail
    request_body = SendMailPostRequestBody(message=model, save_to_sent_items=save_to_sent_items)
    execute_request(builder, method="POST", body=request_body, adapter=adapter)
    return success_payload(SEND, {"accepted": True, "form": "composed"})
