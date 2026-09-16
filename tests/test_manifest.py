import importlib.metadata as md
import tomllib
from pathlib import Path

import pytest
import yaml
from packaging.version import Version

MANIFEST_DISTRIBUTION = "hermes-microsoft365-plugin"


def _manifest_version() -> str:
    manifest = yaml.safe_load(Path("microsoft365/plugin.yaml").read_text(encoding="utf-8"))
    return manifest["version"]


def _declared_package_version() -> str:
    project = tomllib.loads(Path("pyproject.toml").read_text(encoding="utf-8"))["project"]
    return project["version"]


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


def test_declared_versions_agree_with_package_metadata():
    declared = _declared_package_version()
    wheel = str(Version(declared))

    # The wheel records the PEP 440 canonical spelling of the project version. Equal-looking
    # spellings (``0.1.0-alpha.1`` vs ``0.1.0a1``) would still leave the installed metadata and
    # the manifest disagreeing, so the declaration itself must already be canonical.
    assert wheel == declared, (
        f"pyproject version {declared!r} is not the PEP 440 canonical form ({wheel!r}) "
        "that the built wheel records"
    )
    assert _manifest_version() == wheel


def test_manifest_version_matches_installed_distribution_metadata():
    try:
        installed = md.version(MANIFEST_DISTRIBUTION)
    except md.PackageNotFoundError:
        pytest.skip(f"{MANIFEST_DISTRIBUTION} is not installed in this environment")

    assert _manifest_version() == installed
