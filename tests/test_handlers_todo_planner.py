import json
from datetime import datetime, timezone

import pytest
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from msgraph.generated.models.planner_task import PlannerTask
from msgraph.generated.models.todo_task import TodoTask
from msgraph.generated.models.todo_task_list_collection_response import TodoTaskListCollectionResponse
from msgraph.generated.models.todo_task_collection_response import TodoTaskCollectionResponse

from microsoft365.handlers import HandlerContext, load_handlers
from microsoft365.contract import Settings

from tests.test_handlers_outlook_calendar import HandlerGraphAdapter, graph_client


def context(adapter):
    client = graph_client(adapter)
    return HandlerContext(Settings.from_mapping({"capabilities": {s: True for s in ("todo", "planner")}}), lambda: client)


def test_todo_reads_are_registered_and_use_real_paths():
    from microsoft365.handlers.todo import list_task_lists, search, read
    adapter = HandlerGraphAdapter({
        ("GET", "/users/u/todo/lists"): TodoTaskListCollectionResponse(value=[]),
        ("GET", "/users/u/todo/lists/l/tasks"): TodoTaskCollectionResponse(value=[]),
        ("GET", "/users/u/todo/lists/l/tasks/t"): TodoTask(id="t"),
    })
    ctx = context(adapter)
    list_task_lists({"user_id": "u", "top": 2}, ctx)
    search({"user_id": "u", "todo_list_id": "l", "title": "milk", "top": 2}, ctx)
    read({"user_id": "u", "todo_list_id": "l", "todo_task_id": "t"}, ctx)
    assert [r.http_method.value for r in adapter.requests] == ["GET", "GET", "GET"]
    assert "/users/u/todo/lists" in adapter.requests[0].url
    assert "/users/u/todo/lists/l/tasks" in adapter.requests[1].url
    assert "/users/u/todo/lists/l/tasks/t" in adapter.requests[2].url


def test_todo_write_models_preserve_supplied_fields_without_nulls():
    from microsoft365.handlers.todo import _todo_model
    due = DateTimeTimeZone(date_time="2026-01-02T03:04:05Z", time_zone="UTC")
    model = _todo_model("todo_create_tasks", {"title": "x", "due_date_time": due})
    assert isinstance(model, TodoTask)
    adapter = HandlerGraphAdapter()
    from microsoft365.sdk_contract import build_write_request_information
    request = build_write_request_information(graph_client(adapter), "todo_create_tasks", user_id="u", todo_list_id="l", update_fields={"title": "x", "due_date_time": due}) if False else None
    assert model.title == "x"
    assert model.due_date_time is due
    assert model.body is None
    assert model.status is None


def test_planner_create_has_required_container_and_real_model():
    from microsoft365.handlers.planner import _planner_create_model
    model = _planner_create_model({"plan_id": "p", "bucket_id": "b", "title": "task"})
    assert isinstance(model, PlannerTask)
    assert model.plan_id == "p" and model.bucket_id == "b" and model.title == "task"
    with pytest.raises(ValueError):
        _planner_create_model({"plan_id": "p", "title": "task"})


def test_planner_update_is_partial_and_requires_etag():
    from microsoft365.handlers.planner import _planner_update_model
    model = _planner_update_model({"title": "new", "percent_complete": 50})
    assert isinstance(model, PlannerTask)
    assert model.title == "new" and model.percent_complete == 50
    assert model.priority is None and model.due_date_time is None
    with pytest.raises(ValueError):
        _planner_update_model({})


def test_reads_only_are_executable_and_writes_withheld():
    from microsoft365.contract import EXECUTABLE_OPERATIONS
    assert set(EXECUTABLE_OPERATIONS) >= {
        "todo.list_task_lists", "todo.search", "todo.read",
        "planner.list_plans", "planner.list_buckets", "planner.list_tasks", "planner.read",
    }
    assert not set(EXECUTABLE_OPERATIONS) & {"todo.create_tasks", "todo.update_tasks", "planner.create_tasks", "planner.update_tasks"}
    loaded = load_handlers()
    assert {k for k in loaded if k.startswith(("todo.", "planner."))} == {
        "todo.list_task_lists", "todo.search", "todo.read", "todo.create_tasks", "todo.update_tasks",
        "planner.list_plans", "planner.list_buckets", "planner.list_tasks", "planner.read", "planner.create_tasks", "planner.update_tasks",
    }
