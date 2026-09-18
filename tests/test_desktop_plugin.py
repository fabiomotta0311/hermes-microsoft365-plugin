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


def test_permission_helper_renders_truthful_sanitized_capability_details():
    source = PLUGIN.read_text(encoding="utf-8")
    for text in (
        "sanitizeCapabilityRow", "CapabilityHelper", 'jsx("details",', 'jsx("summary",',
        "Permissões exatas", "Modo configurado", "Executável", "Status de autenticação",
        "Status de implementação", "capabilityChecklist", "readOnly: true", "textarea",
    ):
        assert text in source
    assert "escrita retida; não habilitada" in source
    assert "não confirma consentimento administrativo" in source
    assert "row.permissions" in source
    assert "row.service" in source
    assert "row.operation" in source
    assert "row.write" in source
    assert "row.status" in source


def test_permission_helper_does_not_render_identifiers_or_claim_approval():
    source = PLUGIN.read_text(encoding="utf-8")
    helper = source[source.index("function sanitizeCapabilityRow"):source.index("function Microsoft365Page")]
    for forbidden in (
        "tenant_id", "client_id", "user_id", "access_token", "client_secret",
        "admin-consent", "admin consent", "aprovação concluída", "consentimento concedido",
        "OAuth", "Graph call",
    ):
        assert forbidden.lower() not in helper.lower()
    assert "item.key" not in helper
    assert "status.remote_verification" not in helper


def test_permission_helper_preserves_honest_query_states():
    source = PLUGIN.read_text(encoding="utf-8")
    assert "Carregando capacidades…" in source
    assert "Não foi possível carregar as capacidades." in source
    assert "Nenhuma capacidade foi publicada pelo dashboard." in source
    assert "query.data.operations.length === 0" in source


def test_desktop_plugin_parses_without_node_dependencies():
    result = subprocess.run(
        ["node", "--check", str(PLUGIN)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr
