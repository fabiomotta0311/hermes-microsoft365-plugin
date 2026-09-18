from pathlib import Path
import subprocess


PLUGIN = Path(__file__).parents[1] / "microsoft365" / "desktop" / "plugin.js"


def test_desktop_shell_uses_sanitized_dashboard_queries_and_is_registered():
    source = PLUGIN.read_text(encoding="utf-8")
    assert "ROUTES_AREA" in source
    assert "SIDEBAR_NAV_AREA" in source
    assert "PALETTE_AREA" in source
    assert "'/microsoft365'" in source or '"/microsoft365"' in source
    for text in ("Microsoft 365", "Alpha", "não contata a Microsoft", "Delegated", "Application"):
        assert text in source
    assert "useQuery" in source
    for endpoint in ("/configuration", "/capabilities", "/preflight"):
        assert f'"{endpoint}"' in source
    assert "ctx.rest" in source
    assert "isCapabilitiesResponse" in source
    assert "isConfigurationResponse" in source
    assert "isPreflightResponse" in source
    assert "ctx.host" not in source
    assert "host.navigate" in source
    assert "access_token" not in source
    assert "client_secret" not in source
    assert "fake" not in source.lower()
    assert "CAPABILITIES =" not in source
    assert "access_token" not in source
    assert "client_secret" not in source


def test_desktop_plugin_parses_without_node_dependencies():
    result = subprocess.run(
        ["node", "--check", str(PLUGIN)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr
