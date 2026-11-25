import json

import pytest

from microsoft365.contract import OPERATION_REGISTRY, Settings, operation_status
from microsoft365.registration import HANDLER_TABLE, active_actions, service_tool_handler


@pytest.mark.parametrize("mode", ["application", "delegated"])
def test_all_writes_are_retained_at_dispatch_and_in_status(mode):
    settings = Settings.from_mapping({"authentication_mode": mode, "capabilities": {
        service: True for service in ("outlook", "calendar", "sharepoint", "onedrive", "teams", "todo", "planner")
    }})
    writes = {key: definition for key, definition in OPERATION_REGISTRY.items() if definition.write}
    assert writes
    for key, definition in writes.items():
        service, action = key.split(".")
        assert not definition.executable, key
        assert not operation_status(mode, service, action).executable, key
        assert action not in active_actions(settings, service), key
        if key in HANDLER_TABLE:
            called = []
            result = service_tool_handler(service, {"action": action}, settings=settings,
                                          dispatch_table={key: lambda args: called.append(args)})
            assert json.loads(result)["error"] == "operation_not_implemented", key
            assert not called, key
