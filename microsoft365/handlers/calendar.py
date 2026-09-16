"""Calendar handlers: search, create_events and update_events (WP6).

The same rules as the Outlook module apply: each request is built on the generated SDK
builders from the declared arguments and pinned byte-for-byte against the verified
request-information contract of the operation (``microsoft365/sdk_contract.py``; see
``tests/test_handlers_outlook_calendar.py``).

Action-specific rules this module owns
--------------------------------------
* ``calendar.search``: a date range is **not** part of the verified request contract of this
  endpoint -- the endpoint's acceptance of a date filter is not recorded offline -- so
  ``start_date_time``/``end_date_time``/``time_zone`` are refused instead of being turned into a
  guessed OData expression. ``filter`` is the caller's own expression and is passed through.
* ``calendar.create_events``/``update_events``: an offsetless date and time must be
  accompanied by ``time_zone`` (the argument contract refuses one without it; the handler
  re-checks, because the request it builds is what would reach Graph), and a value that
  carries an offset keeps it.
* ``calendar.update_events``: the ETag is mandatory (a conditional update), ``time_zone`` on
  its own changes nothing and is refused, and **every** field the caller asked to change is
  proven to have changed the serialized request before anything is sent.

``create_events`` and ``update_events`` are implemented, exercised offline through the seam
and deliberately **not executable** until the generic host approval fix (CORE-1/CORE-2)
lands (R5).
"""
from __future__ import annotations

from datetime import datetime
from typing import Any, Mapping

from msgraph.generated.models.attendee import Attendee
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from msgraph.generated.models.event import Event
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.location import Location

from ..execution import execute_request, request_information_sender
from ..paging import paginate
from ..validation import EVENT_PATCH_FIELDS
from . import (
    HandlerContext,
    collection_payload,
    field_label,
    handler,
    headers_configuration,
    identifier,
    item_budget,
    open_graph_client,
    recipients,
    refuse,
    require_changed_fields,
    success_payload,
    text,
    typed_configuration,
)

SEARCH = "calendar.search"
CREATE_EVENTS = "calendar.create_events"
UPDATE_EVENTS = "calendar.update_events"

#: The arguments of ``calendar.search`` that have no verified request form on this endpoint.
UNVERIFIED_SEARCH_ARGUMENTS = ("start_date_time", "end_date_time", "time_zone")

#: The event fields this plugin can change, each with the JSON keys it must produce in the
#: serialized request. The *allowlist* is the shared argument contract
#: (:data:`microsoft365.validation.EVENT_PATCH_FIELDS`); this map is what lets the handler
#: prove that an accepted field really did change the request it was meant to change, and a
#: test pins that the two cover exactly the same fields.
EVENT_WIRE_KEYS: Mapping[str, tuple[str, ...]] = {
    "subject": ("subject",),
    "body": ("body",),
    "start_date_time": ("start",),
    "end_date_time": ("end",),
    "time_zone": ("timeZone",),
    "location": ("location",),
    "attendees": ("attendees",),
    "is_all_day": ("isAllDay",),
}

#: The fields a ``calendar.create_events`` call may carry, as the argument contract declares.
CREATE_EVENT_FIELDS = (
    "subject",
    "body",
    "start_date_time",
    "end_date_time",
    "time_zone",
    "location",
    "attendees",
    "is_all_day",
)


# ---------------------------------------------------------------- argument helpers


def date_time_time_zone(
    arguments: Mapping, name: str, *, operation: str, time_zone: str | None = None, required: bool = False
) -> DateTimeTimeZone | None:
    """The real ``DateTimeTimeZone`` one declared date and time becomes.

    An offsetless value must be accompanied by ``time_zone``: without one it would be applied
    in the wrong zone, so it is refused. A value that carries an offset keeps it.
    """
    value = arguments.get(name)
    if value is None:
        if required:
            refuse("validation_error", f"{operation}: {name} is required")
        return None
    if not isinstance(value, str) or not value.strip():
        refuse("validation_error", f"{operation}: {name} must be a date and time")
    if offsetless(value) and not time_zone:
        refuse(
            "validation_error",
            f"{operation}: {name} carries no UTC offset, so time_zone is required",
        )
    return DateTimeTimeZone(date_time=value, time_zone=time_zone)


def offsetless(value: str) -> bool:
    """True when an ISO-8601 date and time carries no UTC offset.

    A value the standard library cannot parse is treated as offsetless, which is the stricter
    reading (the argument contract already refuses a value that is not a date and time).
    """
    try:
        return datetime.fromisoformat(value).tzinfo is None
    except ValueError:
        return True


def non_blank(value: Any, *, name: str, operation: str) -> str:
    if not isinstance(value, str) or not value.strip():
        refuse("validation_error", f"{operation}: {name} must be a non-blank string")
    return value


def attendee_models(arguments: Mapping, name: str, *, operation: str) -> list | None:
    """Real ``Attendee`` models for a declared recipient list."""
    models = recipients(arguments, name, operation=operation)
    if not models:
        return None
    return [Attendee(email_address=recipient.email_address) for recipient in models]


def changed_fields(arguments: Mapping, names, *, operation: str) -> dict:
    """The declarations of every field the caller actually supplied, by wire key."""
    supplied = {}
    for name in names:
        if arguments.get(name) is None:
            continue
        keys = EVENT_WIRE_KEYS.get(name)
        if keys is None:  # pragma: no cover - a declared field without a wire key is a bug
            refuse(
                "validation_error",
                f"{operation}: {field_label(name)} cannot be expressed in the request this "
                "plugin sends",
            )
        supplied[name] = keys
    return supplied


# ---------------------------------------------------------------- reads


@handler("calendar", "search")
def search(arguments: Mapping, context: HandlerContext) -> dict:
    """List one mailbox's calendar events, bounded by the caller's item budget."""
    user_id = identifier(arguments, "user_id", operation=SEARCH)
    for name in UNVERIFIED_SEARCH_ARGUMENTS:
        if arguments.get(name) is not None:
            refuse(
                "validation_error",
                f"{SEARCH}: {name} is not part of the verified request contract of this "
                "operation, so the call is refused instead of guessing a date filter",
            )
    filters = text(arguments, "filter", operation=SEARCH)
    limit = item_budget(arguments, operation=SEARCH)

    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).calendar.events
    configuration = typed_configuration(builder, top=limit, filter=filters)
    paged = paginate(builder, adapter=adapter, limit=limit, configuration=configuration)
    return collection_payload(SEARCH, paged)


# ---------------------------------------------------------------- writes


@handler("calendar", "create_events")
def create_events(arguments: Mapping, context: HandlerContext) -> dict:
    """Create one calendar event."""
    user_id = identifier(arguments, "user_id", operation=CREATE_EVENTS)
    subject = text(arguments, "subject", operation=CREATE_EVENTS)
    if subject is None:
        refuse("validation_error", f"{CREATE_EVENTS}: subject is required")
    time_zone = text(arguments, "time_zone", operation=CREATE_EVENTS)
    content = text(arguments, "body", operation=CREATE_EVENTS)
    location = text(arguments, "location", operation=CREATE_EVENTS)
    is_all_day = arguments.get("is_all_day")
    if is_all_day is not None and type(is_all_day) is not bool:
        refuse("validation_error", f"{CREATE_EVENTS}: is_all_day must be a real boolean")

    start = date_time_time_zone(
        arguments, "start_date_time", operation=CREATE_EVENTS, time_zone=time_zone, required=True
    )
    end = date_time_time_zone(
        arguments, "end_date_time", operation=CREATE_EVENTS, time_zone=time_zone, required=True
    )
    model = Event(
        subject=subject,
        start=start,
        end=end,
        body=ItemBody(content_type=BodyType.Text, content=content) if content else None,
        location=Location(display_name=location) if location else None,
        attendees=attendee_models(arguments, "attendees", operation=CREATE_EVENTS),
        is_all_day=is_all_day,
    )

    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).calendar.events
    request_information = request_information_sender(builder, method="POST", body=model)
    require_changed_fields(
        request_information,
        operation=CREATE_EVENTS,
        supplied=changed_fields(arguments, CREATE_EVENT_FIELDS, operation=CREATE_EVENTS),
    )
    response = execute_request(builder, method="POST", body=model, adapter=adapter)
    return success_payload(CREATE_EVENTS, response)


@handler("calendar", "update_events")
def update_events(arguments: Mapping, context: HandlerContext) -> dict:
    """Change the declared fields of one calendar event, conditionally on its ETag."""
    user_id = identifier(arguments, "user_id", operation=UPDATE_EVENTS)
    event_id = identifier(arguments, "event_id", operation=UPDATE_EVENTS)
    etag = arguments.get("etag")
    if not isinstance(etag, str) or not etag.strip():
        refuse(
            "validation_error",
            f"{UPDATE_EVENTS}: etag is required, because without it the service cannot tell "
            "whether the event changed since it was read",
        )
    fields = update_fields(arguments, operation=UPDATE_EVENTS)
    model = event_update_model(fields, operation=UPDATE_EVENTS)

    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).calendar.events.by_event_id(event_id)
    configuration = headers_configuration(**{"If-Match": etag})
    request_information = request_information_sender(
        builder, method="PATCH", configuration=configuration, body=model
    )
    require_changed_fields(
        request_information,
        operation=UPDATE_EVENTS,
        supplied={name: EVENT_WIRE_KEYS[name] for name in fields},
    )
    response = execute_request(
        builder, method="PATCH", configuration=configuration, body=model, adapter=adapter
    )
    return success_payload(UPDATE_EVENTS, response)


def update_fields(arguments: Mapping, *, operation: str) -> dict:
    """The declared fields of a PATCH, refused when one of them could change nothing."""
    fields = arguments.get("fields")
    if not isinstance(fields, Mapping) or not fields:
        refuse(
            "validation_error",
            f"{operation}: fields must be a non-empty object of the event fields to change",
        )
    unknown = sorted(name for name in fields if name not in EVENT_PATCH_FIELDS)
    if unknown:
        refuse(
            "validation_error",
            f"{operation}: {field_label(unknown[0])} is not a field this call may change; the "
            f"fields it accepts are {', '.join(EVENT_PATCH_FIELDS)}",
        )
    uncovered = sorted(name for name in fields if name not in EVENT_WIRE_KEYS)
    if uncovered:  # pragma: no cover - pinned by test against the declared allowlist
        refuse(
            "validation_error",
            f"{operation}: {field_label(uncovered[0])} cannot be expressed in the request this "
            "plugin sends",
        )
    if "time_zone" in fields and not ({"start_date_time", "end_date_time"} & set(fields)):
        refuse(
            "validation_error",
            f"{operation}: time_zone only applies to a start or an end date and time, so on "
            "its own it would change nothing",
        )
    return dict(fields)


def event_update_model(fields: Mapping, *, operation: str) -> Event:
    """The real (partial) generated ``Event`` a PATCH sends.

    Only the fields the caller declared become attributes of the model, so nothing the caller
    did not ask to change can reach the request.
    """
    time_zone = fields.get("time_zone")
    if "time_zone" in fields and (not isinstance(time_zone, str) or not time_zone.strip()):
        refuse("validation_error", f"{operation}: time_zone must be a non-blank time zone name")
    attributes: dict = {}
    for name, value in fields.items():
        if name == "subject":
            attributes["subject"] = non_blank(value, name=name, operation=operation)
        elif name == "body":
            attributes["body"] = ItemBody(
                content_type=BodyType.Text,
                content=non_blank(value, name=name, operation=operation),
            )
        elif name in ("start_date_time", "end_date_time"):
            attributes["start" if name == "start_date_time" else "end"] = date_time_time_zone(
                {name: value}, name, operation=operation, time_zone=time_zone, required=True
            )
        elif name == "location":
            attributes["location"] = Location(
                display_name=non_blank(value, name=name, operation=operation)
            )
        elif name == "attendees":
            attributes["attendees"] = attendee_models({name: value}, name, operation=operation)
        elif name == "is_all_day":
            if type(value) is not bool:
                refuse("validation_error", f"{operation}: is_all_day must be a real boolean")
            attributes["is_all_day"] = value
        # ``time_zone`` is carried inside the start/end values it belongs to.
    return Event(**attributes)
