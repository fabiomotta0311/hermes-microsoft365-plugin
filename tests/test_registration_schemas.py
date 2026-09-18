import pytest


SERVICES = ("outlook", "sharepoint", "onedrive", "calendar", "teams", "todo", "planner")


def _maximum_settings():
    from microsoft365.contract import OPERATIONS, Settings

    return Settings.from_mapping({"capabilities": {service: True for service in OPERATIONS}})


def test_registered_service_schema_is_a_discriminated_union_of_action_contracts():
    from microsoft365.registration import active_actions, schema_for

    settings = _maximum_settings()
    for service in SERVICES:
        actions = active_actions(settings, service)
        schema = schema_for(service, actions)
        assert schema is not None
        parameters = schema["parameters"]
        assert parameters["type"] == "object"
        branches = parameters["oneOf"]
        assert [branch["properties"]["action"]["enum"] for branch in branches] == [
            [action] for action in actions
        ]
        assert all(branch["additionalProperties"] is False for branch in branches)


def test_registered_schema_branches_are_derived_from_runtime_operation_contracts():
    from microsoft365.registration import active_actions, schema_for
    from microsoft365.validation import ARGUMENT_CONTRACTS, operation_schema

    settings = _maximum_settings()
    for service in SERVICES:
        actions = active_actions(settings, service)
        schema = schema_for(service, actions)
        branches = schema["parameters"]["oneOf"] if schema else []
        expected = [operation_schema(f"{service}.{action}") for action in actions]
        assert branches == expected
        for action in actions:
            assert f"{service}.{action}" in ARGUMENT_CONTRACTS


def test_schema_for_rejects_unknown_or_uncontracted_actions():
    from microsoft365.registration import schema_for

    with pytest.raises(ValueError, match="declared operation"):
        schema_for("outlook", ("not_an_operation",))


def test_operation_schemas_carry_runtime_bounds_for_specialized_values():
    from microsoft365.validation import operation_schema

    calendar = operation_schema("calendar.create_events")
    assert calendar["properties"]["start_date_time"]["maxLength"] == 64
    select = operation_schema("outlook.read")["properties"]["select"]
    assert select["items"]["maxLength"] == 64

    planner = operation_schema("planner.create_tasks")
    assert planner["properties"]["percent_complete"] == {
        "type": "integer",
        "minimum": 0,
        "maximum": 100,
    }
    planner_update = operation_schema("planner.update_tasks")
    assert planner_update["properties"]["fields"]["properties"]["priority"] == {
        "type": "integer",
        "minimum": 0,
        "maximum": 10,
    }


def test_registered_tool_surface_uses_the_strict_action_schema():
    from microsoft365 import register
    from microsoft365.contract import OPERATIONS

    class Context:
        def __init__(self):
            self.config = {"capabilities": {service: True for service in OPERATIONS}}
            self.schemas = {}

        def get_config(self, key, default=None):
            return self.config.get(key, default)

        def register_tool(self, name, **kwargs):
            self.schemas[name] = kwargs["schema"]

        def register_hook(self, name, callback):
            pass

    context = Context()
    register(context)
    for service in SERVICES:
        schema = context.schemas[f"microsoft365_{service}"]
        assert "oneOf" in schema["parameters"]
        assert schema["parameters"]["additionalProperties"] is False
        assert all(branch["additionalProperties"] is False for branch in schema["parameters"]["oneOf"])


def test_outer_schema_strictness_preserves_valid_action_specific_properties():
    from microsoft365.registration import schema_for
    from microsoft365.validation import operation_schema

    schema = schema_for("calendar", ("create_events", "search"))
    parameters = schema["parameters"]

    assert parameters["additionalProperties"] is False
    create_branch = next(
        branch for branch in parameters["oneOf"] if branch["properties"]["action"]["enum"] == ["create_events"]
    )
    expected = operation_schema("calendar.create_events")
    assert create_branch["properties"]["subject"] == expected["properties"]["subject"]
    assert "subject" in create_branch["required"]
