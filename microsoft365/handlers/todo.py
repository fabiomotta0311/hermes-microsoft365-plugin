"""Microsoft To Do handlers using the generated Graph builders (WP9)."""
from __future__ import annotations

from typing import Mapping

from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.todo_task import TodoTask

from ..execution import execute_request, request_information_sender
from ..sdk_contract import _todo_task_write_model
from . import (
    HandlerContext, collection_payload, handler, identifier, item_budget,
    open_graph_client, refuse, select, success_payload, typed_configuration,
)
from ..paging import paginate

LIST = "todo.list_task_lists"
SEARCH = "todo.search"
READ = "todo.read"
CREATE = "todo.create_tasks"
UPDATE = "todo.update_tasks"


def _fields(arguments: Mapping, operation: str) -> dict:
    fields = arguments.get("fields")
    if not isinstance(fields, Mapping) or not fields:
        refuse("validation_error", f"{operation}: fields must be a non-empty object")
    return dict(fields)


def _todo_model(case: str, fields: Mapping) -> TodoTask:
    normalized = dict(fields)
    if isinstance(normalized.get("body"), str):
        normalized["body"] = ItemBody(content_type=BodyType.Text, content=normalized["body"])
    if isinstance(normalized.get("due_date_time"), str):
        normalized["due_date_time"] = DateTimeTimeZone(date_time=normalized["due_date_time"])
    return _todo_task_write_model(case, normalized)


def _list_builder(client, arguments: Mapping, operation: str):
    user_id = identifier(arguments, "user_id", operation=operation)
    list_id = identifier(arguments, "todo_list_id", operation=operation)
    return client.users.by_user_id(user_id).todo.lists.by_todo_task_list_id(list_id)


@handler("todo", "list_task_lists")
def list_task_lists(arguments: Mapping, context: HandlerContext) -> dict:
    user_id = identifier(arguments, "user_id", operation=LIST)
    limit = item_budget(arguments, operation=LIST)
    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).todo.lists
    paged = paginate(builder, adapter=adapter, limit=limit, configuration=typed_configuration(builder, top=limit))
    return collection_payload(LIST, paged)


@handler("todo", "search")
def search(arguments: Mapping, context: HandlerContext) -> dict:
    user_id = identifier(arguments, "user_id", operation=SEARCH)
    list_id = identifier(arguments, "todo_list_id", operation=SEARCH)
    title = arguments.get("title")
    if title is not None and (not isinstance(title, str) or not title.strip()):
        refuse("validation_error", f"{SEARCH}: title must be a non-blank string")
    limit = item_budget(arguments, operation=SEARCH)
    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).todo.lists.by_todo_task_list_id(list_id).tasks
    configuration = typed_configuration(builder, top=limit, filter=(f"contains(title,'{title.replace(chr(39), chr(39)*2)}')" if title is not None else None), select=select(arguments, operation=SEARCH))
    return collection_payload(SEARCH, paginate(builder, adapter=adapter, limit=limit, configuration=configuration))


@handler("todo", "read")
def read(arguments: Mapping, context: HandlerContext) -> dict:
    user_id = identifier(arguments, "user_id", operation=READ)
    list_id = identifier(arguments, "todo_list_id", operation=READ)
    task_id = identifier(arguments, "todo_task_id", operation=READ)
    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).todo.lists.by_todo_task_list_id(list_id).tasks.by_todo_task_id(task_id)
    return success_payload(READ, execute_request(builder, method="GET", configuration=typed_configuration(builder, select=select(arguments, operation=READ)), adapter=adapter))


@handler("todo", "create_tasks")
def create_tasks(arguments: Mapping, context: HandlerContext) -> dict:
    user_id = identifier(arguments, "user_id", operation=CREATE)
    list_id = identifier(arguments, "todo_list_id", operation=CREATE)
    fields = _fields(arguments, CREATE)
    if "title" not in fields:
        refuse("validation_error", f"{CREATE}: fields.title is required")
    model = _todo_model("todo_create_tasks", fields)
    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).todo.lists.by_todo_task_list_id(list_id).tasks
    request_information_sender(builder, method="POST", body=model)
    return success_payload(CREATE, execute_request(builder, method="POST", body=model, adapter=adapter))


@handler("todo", "update_tasks")
def update_tasks(arguments: Mapping, context: HandlerContext) -> dict:
    user_id = identifier(arguments, "user_id", operation=UPDATE)
    list_id = identifier(arguments, "todo_list_id", operation=UPDATE)
    task_id = identifier(arguments, "todo_task_id", operation=UPDATE)
    fields = _fields(arguments, UPDATE)
    model = _todo_model("todo_update_tasks", fields)
    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).todo.lists.by_todo_task_list_id(list_id).tasks.by_todo_task_id(task_id)
    request_information_sender(builder, method="PATCH", body=model)
    return success_payload(UPDATE, execute_request(builder, method="PATCH", body=model, adapter=adapter))
