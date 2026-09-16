from __future__ import annotations

import pytest

EXPECTED_OPERATIONS = {
    "outlook": ("search", "read", "create_draft", "send"),
    "sharepoint": ("search", "read", "download_files", "upload_files"),
    "onedrive": ("search", "read", "download_files", "upload_files"),
    "calendar": ("search", "create_events", "update_events"),
    "teams": ("list_teams", "list_channels", "search_messages", "send_messages"),
    "todo": ("list_task_lists", "search", "read", "create_tasks", "update_tasks"),
    "planner": ("list_plans", "list_buckets", "list_tasks", "read", "create_tasks", "update_tasks"),
}


def test_operation_registry_preserves_complete_product_contract():
    from microsoft365.contract import OPERATIONS, OPERATION_REGISTRY

    assert OPERATIONS == EXPECTED_OPERATIONS
    assert set(OPERATION_REGISTRY) == {
        f"{service}.{operation}"
        for service, operations in EXPECTED_OPERATIONS.items()
        for operation in operations
    }


def test_application_auth_matrix_keeps_unsupported_and_unverified_explicit():
    from microsoft365.contract import operation_status

    assert operation_status("application", "teams", "search_messages").auth_status == "unsupported_auth_mode"
    assert operation_status("application", "teams", "send_messages").auth_status == "unsupported_auth_mode"
    assert operation_status("application", "todo", "create_tasks").auth_status == "not_verified"
    assert operation_status("application", "outlook", "search").auth_status == "supported"
    assert operation_status("delegated", "outlook", "search").auth_status == "not_implemented"


def test_capabilities_accept_only_real_booleans_and_known_names():
    from microsoft365.contract import ConfigurationError, Settings

    settings = Settings.from_mapping({"capabilities": {"outlook": {"search": True, "send": False}}})
    assert settings.selected("outlook") == {"search"}

    for malformed in ("false", 0, 1, None, []):
        with pytest.raises(ConfigurationError, match="must be a boolean or operation mapping"):
            Settings.from_mapping({"capabilities": {"outlook": malformed}})
    with pytest.raises(ConfigurationError, match="must be a boolean"):
        Settings.from_mapping({"capabilities": {"outlook": {"search": "false"}}})
    with pytest.raises(ConfigurationError, match="unknown capability"):
        Settings.from_mapping({"capabilities": {"mail": {"search": True}}})
    with pytest.raises(ConfigurationError, match="unknown operation"):
        Settings.from_mapping({"capabilities": {"outlook": {"delete": True}}})


def test_service_true_enables_every_administrative_operation_without_hiding_status():
    from microsoft365.contract import Settings

    settings = Settings.from_mapping({"capabilities": {"teams": True}})
    assert settings.selected("teams") == set(EXPECTED_OPERATIONS["teams"])


def test_supported_action_enum_never_becomes_empty():
    from microsoft365.registration import active_actions, schema_for
    from microsoft365.contract import Settings

    unsupported_only = Settings.from_mapping({"capabilities": {"teams": {"send_messages": True}}})
    assert active_actions(unsupported_only, "teams") == ()
    assert schema_for("teams", ()) is None
