"""Microsoft Planner handlers using verified generated Graph builders (WP9)."""
from __future__ import annotations

from typing import Mapping

from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from msgraph.generated.models.planner_task import PlannerTask

from ..execution import execute_request, request_information_sender
from ..paging import paginate
from ..sdk_contract import _partial_model
from . import HandlerContext, collection_payload, handler, identifier, item_budget, open_graph_client, refuse, success_payload, typed_configuration, require_changed_fields

LIST_PLANS = "planner.list_plans"
LIST_BUCKETS = "planner.list_buckets"
LIST_TASKS = "planner.list_tasks"
READ = "planner.read"
CREATE = "planner.create_tasks"
UPDATE = "planner.update_tasks"

_CREATE_FIELDS = {
    "plan_id": "identifier", "bucket_id": "identifier", "title": "text",
    "due_date_time": "timestamp", "start_date_time": "timestamp",
    "percent_complete": "integer", "priority": "integer",
}
_UPDATE_WIRE = {
    "title": ("title",), "bucket_id": ("bucketId",),
    "due_date_time": ("dueDateTime",), "start_date_time": ("startDateTime",),
    "percent_complete": ("percentComplete",), "priority": ("priority",),
}


def _planner_create_model(fields: Mapping) -> PlannerTask:
    if not isinstance(fields, Mapping):
        raise ValueError("planner.create_tasks fields must be a mapping")
    fields = {name: value for name, value in fields.items() if name != "action"}
    required = ("plan_id", "bucket_id", "title")
    if any(name not in fields or not isinstance(fields[name], str) or not fields[name].strip() for name in required):
        raise ValueError("planner.create_tasks requires non-blank plan_id, bucket_id and title")
    unknown = set(fields) - set(_CREATE_FIELDS)
    if unknown:
        raise ValueError(f"unknown field for planner.create_tasks: {sorted(unknown)[0]}")
    normalized = dict(fields)
    for name in ("due_date_time", "start_date_time"):
        if isinstance(normalized.get(name), str):
            normalized[name] = DateTimeTimeZone(date_time=normalized[name])
        if name in normalized and not isinstance(normalized[name], DateTimeTimeZone):
            raise ValueError(f"{name} must be a DateTimeTimeZone")
    for name in ("percent_complete", "priority"):
        if name in normalized and type(normalized[name]) is not int:
            raise ValueError(f"{name} must be an integer")
    return PlannerTask(**normalized)


def _planner_update_model(fields: Mapping) -> PlannerTask:
    if not isinstance(fields, Mapping) or not fields:
        raise ValueError("planner.update_tasks requires a non-empty fields mapping")
    unknown = set(fields) - set(_UPDATE_WIRE)
    if unknown:
        raise ValueError(f"unknown field for planner.update_tasks: {sorted(unknown)[0]}")
    normalized = dict(fields)
    for name, value in list(normalized.items()):
        if value is None:
            raise ValueError(f"{name} must not be None")
        if name in {"due_date_time", "start_date_time"} and isinstance(value, str):
            normalized[name] = DateTimeTimeZone(date_time=value)
        if name in {"due_date_time", "start_date_time"} and not isinstance(normalized[name], DateTimeTimeZone):
            raise ValueError(f"{name} must be a DateTimeTimeZone")
        if name in {"percent_complete", "priority"} and type(value) is not int:
            raise ValueError(f"{name} must be an integer")
        if name in {"title", "bucket_id"} and (not isinstance(value, str) or not value.strip()):
            raise ValueError(f"{name} must be a non-blank string")
    return PlannerTask(**normalized)


def _plans_builder(client, group_id):
    return client.groups.by_group_id(group_id).planner.plans if group_id is not None else client.planner.plans


@handler("planner", "list_plans")
def list_plans(arguments: Mapping, context: HandlerContext) -> dict:
    group_id = arguments.get("group_id")
    if group_id is not None and (not isinstance(group_id, str) or not group_id.strip()):
        refuse("validation_error", f"{LIST_PLANS}: group_id must be a non-blank string")
    limit = item_budget(arguments, operation=LIST_PLANS)
    client, adapter = open_graph_client(context)
    builder = _plans_builder(client, group_id)
    return collection_payload(LIST_PLANS, paginate(builder, adapter=adapter, limit=limit, configuration=typed_configuration(builder, top=limit, filter=None)))


@handler("planner", "list_buckets")
def list_buckets(arguments: Mapping, context: HandlerContext) -> dict:
    plan_id = identifier(arguments, "plan_id", operation=LIST_BUCKETS)
    client, adapter = open_graph_client(context)
    builder = client.planner.plans.by_planner_plan_id(plan_id).buckets
    limit = item_budget(arguments, operation=LIST_BUCKETS)
    return collection_payload(LIST_BUCKETS, paginate(builder, adapter=adapter, limit=limit, configuration=typed_configuration(builder, top=limit)))


@handler("planner", "list_tasks")
def list_tasks(arguments: Mapping, context: HandlerContext) -> dict:
    plan_id = identifier(arguments, "plan_id", operation=LIST_TASKS)
    client, adapter = open_graph_client(context)
    builder = client.planner.plans.by_planner_plan_id(plan_id).tasks
    limit = item_budget(arguments, operation=LIST_TASKS)
    return collection_payload(LIST_TASKS, paginate(builder, adapter=adapter, limit=limit, configuration=typed_configuration(builder, top=limit)))


@handler("planner", "read")
def read(arguments: Mapping, context: HandlerContext) -> dict:
    task_id = identifier(arguments, "planner_task_id", operation=READ)
    client, adapter = open_graph_client(context)
    builder = client.planner.tasks.by_planner_task_id(task_id)
    return success_payload(READ, execute_request(builder, method="GET", adapter=adapter))


@handler("planner", "create_tasks")
def create_tasks(arguments: Mapping, context: HandlerContext) -> dict:
    model = _planner_create_model(arguments)
    client, adapter = open_graph_client(context)
    builder = client.planner.tasks
    request_information_sender(builder, method="POST", body=model)
    return success_payload(CREATE, execute_request(builder, method="POST", body=model, adapter=adapter))


@handler("planner", "update_tasks")
def update_tasks(arguments: Mapping, context: HandlerContext) -> dict:
    task_id = identifier(arguments, "planner_task_id", operation=UPDATE)
    etag = arguments.get("etag")
    if not isinstance(etag, str) or not etag.strip():
        refuse("validation_error", f"{UPDATE}: etag is required")
    fields = arguments.get("fields")
    model = _planner_update_model(fields)
    client, adapter = open_graph_client(context)
    builder = client.planner.tasks.by_planner_task_id(task_id)
    from . import headers_configuration
    configuration = headers_configuration(**{"If-Match": etag})
    request_information = request_information_sender(builder, method="PATCH", configuration=configuration, body=model)
    require_changed_fields(request_information, operation=UPDATE, supplied={name: keys for name, keys in _UPDATE_WIRE.items() if name in fields})
    return success_payload(UPDATE, execute_request(builder, method="PATCH", configuration=configuration, body=model, adapter=adapter))
