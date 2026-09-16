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
    assert Settings.from_mapping({"capabilities": {"outlook": False}}).selected("outlook") == set()

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


def test_every_operation_records_its_endpoint_write_class_and_statuses():
    """WP13: the per-mode matrix adds fields; nothing may be left implicit."""
    from microsoft365.contract import (
        IMPLEMENTATION_STATUS_LABELS,
        OPERATION_REGISTRY,
        PLANNER_OPERATIONS,
        REMOTE_VERIFICATION_STATUSES,
        TODO_OPERATIONS,
        WRITE_OPERATIONS,
    )

    assert REMOTE_VERIFICATION_STATUSES == frozenset({"not_tested", "verified"})
    verified = {
        key
        for key, definition in OPERATION_REGISTRY.items()
        if definition.implementation_status == "contract_verified"
    }
    assert verified == set(PLANNER_OPERATIONS) | set(TODO_OPERATIONS)

    for key, definition in OPERATION_REGISTRY.items():
        service, operation = key.split(".", 1)
        assert (definition.service, definition.operation) == (service, operation), key
        assert definition.write is (operation in WRITE_OPERATIONS), key
        assert definition.implementation_status in IMPLEMENTATION_STATUS_LABELS, key
        assert definition.remote_verification == "not_tested", key
        assert definition.executable is False, key
        assert definition.app.mode == "application", key
        assert definition.delegated.mode == "delegated", key
        assert definition.endpoint == "; ".join(definition.endpoints), key
        if key in verified:
            assert definition.endpoints, key
        else:
            # No endpoint table is declared for this service yet (WP1/WP6-WP9 add them):
            # the endpoint is reported as unrecorded rather than invented.
            assert definition.endpoints == (), key


def test_application_support_status_is_recorded_per_operation():
    """The application support status now lives in the registry, per operation."""
    from microsoft365.contract import OPERATION_REGISTRY, TODO_WRITE_OPERATIONS

    assert OPERATION_REGISTRY["outlook.search"].app.status == "supported"
    assert OPERATION_REGISTRY["outlook.search"].app.permissions == ("Mail.Read",)
    assert OPERATION_REGISTRY["teams.send_messages"].app.status == "unsupported_auth_mode"
    assert OPERATION_REGISTRY["teams.send_messages"].app.permissions == ()

    unsupported = {
        key
        for key, definition in OPERATION_REGISTRY.items()
        if definition.app.status == "unsupported_auth_mode"
    }
    assert unsupported == {"teams.search_messages", "teams.send_messages"}
    not_verified = {
        key for key, definition in OPERATION_REGISTRY.items() if definition.app.status == "not_verified"
    }
    assert not_verified == set(TODO_WRITE_OPERATIONS)

    for key, definition in OPERATION_REGISTRY.items():
        if definition.app.status not in {"unsupported_auth_mode", "not_verified"}:
            assert definition.app.status == "supported", key
            assert definition.app.permissions, key
        assert definition.delegated.status == "not_implemented", key
