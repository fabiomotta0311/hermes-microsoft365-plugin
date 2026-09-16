"""Typed msgraph-sdk 1.62.0 request-information contracts."""
from __future__ import annotations

from kiota_abstractions.base_request_configuration import RequestConfiguration


def _required(value: str, name: str) -> str:
    value = str(value or "").strip()
    if not value:
        raise ValueError(f"{name} is required")
    return value


def _limit(value: int) -> int:
    if type(value) is not int or not 1 <= value <= 100:
        raise ValueError("limit must be an integer between 1 and 100")
    return value


def _odata_literal(value: str) -> str:
    return _required(value, "query").replace("'", "''")


def _request(builder, **query_values):
    query_type = getattr(type(builder), f"{type(builder).__name__}GetQueryParameters")
    query = query_type(**query_values)
    configuration = RequestConfiguration(query_parameters=query)
    return builder.to_get_request_information(configuration)


def build_collection_request_information(
    client,
    case: str,
    *,
    user_id: str = "",
    todo_list_id: str = "",
    drive_id: str = "",
    drive_item_id: str = "",
    team_id: str = "",
    plan_id: str = "",
    owner_id: str = "",
    query: str = "",
    limit: int = 50,
):
    """Build request info with each generated collection's exact query type.

    The function stops at Kiota's transport boundary. It neither authenticates
    nor sends a request, making exact SDK shape testable offline.
    """
    bounded = _limit(limit)
    if case == "outlook_messages":
        builder = client.users.by_user_id(_required(user_id, "user_id")).messages
        return _request(builder, top=bounded, filter=f"contains(subject,'{_odata_literal(query)}')")
    if case == "calendar_events":
        builder = client.users.by_user_id(_required(user_id, "user_id")).calendar.events
        return _request(builder, top=bounded, filter=f"contains(subject,'{_odata_literal(query)}')")
    if case == "todo_lists":
        builder = client.users.by_user_id(_required(user_id, "user_id")).todo.lists
        return _request(builder, top=bounded)
    if case == "todo_tasks":
        builder = client.users.by_user_id(_required(user_id, "user_id")).todo.lists.by_todo_task_list_id(
            _required(todo_list_id, "todo_list_id")
        ).tasks
        return _request(builder, top=bounded, filter=f"contains(title,'{_odata_literal(query)}')")
    if case == "drive_children":
        builder = client.drives.by_drive_id(_required(drive_id, "drive_id")).items.by_drive_item_id(
            _required(drive_item_id, "drive_item_id")
        ).children
        return _request(builder, top=bounded)
    if case == "teams_joined":
        builder = client.users.by_user_id(_required(user_id, "user_id")).joined_teams
        return _request(builder)
    if case == "teams_channels":
        builder = client.teams.by_team_id(_required(team_id, "team_id")).channels
        return _request(builder, select=["id", "displayName"])
    if case == "planner_plans":
        builder = client.planner.plans
        return _request(builder, top=bounded, filter=f"owner eq '{_odata_literal(owner_id)}'")
    if case == "planner_buckets":
        builder = client.planner.plans.by_planner_plan_id(_required(plan_id, "plan_id")).buckets
        return _request(builder, top=bounded)
    if case == "planner_tasks":
        builder = client.planner.plans.by_planner_plan_id(_required(plan_id, "plan_id")).tasks
        return _request(builder, top=bounded)
    raise ValueError(f"unknown collection contract: {case}")
