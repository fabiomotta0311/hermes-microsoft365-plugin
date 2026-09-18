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


def test_wizard_has_explicit_progression_gates_for_each_step():
    source = PLUGIN.read_text(encoding="utf-8")
    assert "function canAdvanceFromStep" in source
    assert "capabilities.isLoading" in source
    assert "capabilities.isError" in source
    assert "capabilities.data.operations.length > 0" in source
    assert 'mode === "application"' in source
    assert "configuration.isLoading" in source
    assert "configuration.isError" in source
    assert "configuration.data" in source
    assert "preflight.isLoading" in source
    assert "preflight.isError" in source
    assert "preflight.data.locally_ready === true" in source
    assert "canAdvanceFromStep(draft.step" in source


def test_wizard_navigation_is_honest_and_completion_is_local_only():
    source = PLUGIN.read_text(encoding="utf-8")
    assert "disabled: index > draft.step" in source
    assert "aria-live" in source
    assert "Concluir revisão local" in source
    assert "tenant consent" not in source.lower()
    assert "consentimento do tenant" not in source.lower()
    assert "sucesso remoto" not in source.lower()
    assert "success" not in source.lower()


def test_permission_helper_preserves_honest_query_states():
    source = PLUGIN.read_text(encoding="utf-8")
    assert "Carregando capacidades…" in source
    assert "Não foi possível carregar as capacidades." in source
    assert "Nenhuma capacidade foi publicada pelo dashboard." in source
    assert "query.data.operations.length === 0" in source


def test_preflight_renders_bounded_local_diagnostic_details_and_ready_state():
    source = PLUGIN.read_text(encoding="utf-8")
    for text in (
        "sanitizePreflightItems", "PreflightDetails", "configuration_errors", "missing",
        "Requisitos locais informados pelo dashboard", "Requisitos ausentes",
        "Erros de configuração local", "não comprovam acesso ao tenant Microsoft",
        "Nenhum requisito local pendente foi informado pelo dashboard",
    ):
        assert text in source
    assert "slice(0, PREFLIGHT_ITEM_LIMIT)" in source
    assert "data.configuration_errors" in source
    assert "data.missing" in source


def test_preflight_details_do_not_leak_identifiers_exceptions_or_claim_remote_access():
    source = PLUGIN.read_text(encoding="utf-8")
    details = source[source.index("function sanitizePreflightItems"):source.index("function Microsoft365Page")]
    for forbidden in (
        "tenant_id", "client_id", "user_id", "access_token", "client_secret",
        "exception", "stack", "OAuth", "Graph", "tenant access",
    ):
        assert forbidden.lower() not in details.lower()
    assert "remote_verification" not in details
    assert "não verificado" in source


def test_desktop_plugin_parses_without_node_dependencies():
    result = subprocess.run(
        ["node", "--check", str(PLUGIN)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr
