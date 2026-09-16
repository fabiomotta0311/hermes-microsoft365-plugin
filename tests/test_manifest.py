from pathlib import Path

import yaml


def test_manifest_is_standalone_complete_and_secret_safe():
    manifest = yaml.safe_load(Path("microsoft365/plugin.yaml").read_text(encoding="utf-8"))

    assert manifest["kind"] == "standalone"
    assert manifest["provides_tools"] == [
        "microsoft365_preflight",
        "microsoft365_outlook",
        "microsoft365_sharepoint",
        "microsoft365_onedrive",
        "microsoft365_calendar",
        "microsoft365_teams",
        "microsoft365_todo",
        "microsoft365_planner",
    ]
    assert manifest["provides_hooks"] == ["pre_tool_call"]
    assert manifest["requires_env"] == [{
        "name": "MICROSOFT365_CLIENT_SECRET",
        "description": "Microsoft Entra application client secret",
        "url": "https://entra.microsoft.com/",
        "secret": True,
    }]
    assert "client_secret" not in manifest["config_schema"]
    assert manifest["python_dependencies"][0] == "msgraph-sdk==1.62.0"
