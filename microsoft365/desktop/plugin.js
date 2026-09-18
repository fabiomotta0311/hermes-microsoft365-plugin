import { Badge, Button, ROUTES_AREA, SIDEBAR_NAV_AREA, PALETTE_AREA, host, useQuery } from "@hermes/plugin-sdk";
import { useEffect, useState } from "react";
import { jsx } from "react/jsx-runtime";

const STORAGE_KEY = "microsoft365.desktop-draft";
const STEPS = ["Funcionalidades", "Modo de acesso", "Identidade", "Permissões", "Revisão"];
const INITIAL_DRAFT = { step: 0, mode: "application", selected: [] };

function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function isConfigurationResponse(value) {
  return isRecord(value)
    && (value.authentication_mode === "application" || value.authentication_mode === "delegated")
    && isRecord(value.capabilities)
    && typeof value.tenant_id_configured === "boolean"
    && typeof value.client_id_configured === "boolean"
    && typeof value.user_id_configured === "boolean";
}

function isCapabilitiesResponse(value) {
  return isRecord(value)
    && (value.authentication_mode === "application" || value.authentication_mode === "delegated")
    && Array.isArray(value.operations)
    && value.operations.every((item) => isRecord(item)
      && typeof item.service === "string"
      && typeof item.operation === "string"
      && Array.isArray(item.permissions)
      && item.permissions.every((permission) => typeof permission === "string")
      && typeof item.write === "boolean"
      && isRecord(item.status)
      && typeof item.status.auth_status === "string"
      && typeof item.status.implementation_status === "string"
      && typeof item.status.executable === "boolean"
      && typeof item.status.reason === "string");
}

function isPreflightResponse(value) {
  return isRecord(value)
    && typeof value.locally_ready === "boolean"
    && typeof value.remote_verification === "string"
    && (!Array.isArray(value.configuration_errors)
      || value.configuration_errors.every((item) => typeof item === "string"))
    && (!Array.isArray(value.missing) || value.missing.every((item) => typeof item === "string"));
}

function useDashboardQuery(ctx, path, guard) {
  return useQuery({
    queryKey: ["microsoft365", path],
    queryFn: async () => {
      const payload = await ctx.rest(path, { method: "GET" });
      if (!guard(payload)) throw new Error("Resposta inválida do dashboard");
      return payload;
    },
    retry: false,
  });
}

function storageName() {
  const profile = host && host.state && host.state.profile;
  const value = profile && typeof profile.get === "function" ? profile.get() : "default";
  const safe = String(value || "default").replace(/[^a-zA-Z0-9._-]/g, "_");
  return `${STORAGE_KEY}.${safe}`;
}

function textStyle(size, color) {
  return { fontSize: size, color: color || "var(--ui-text-primary)" };
}

function Card({ title, description, disabled, selected, onClick, children }) {
  return jsx("button", {
    type: "button", disabled, onClick,
    style: {
      display: "block", width: "100%", textAlign: "left", padding: "16px", marginBottom: "12px",
      borderRadius: "10px", border: `1px solid ${selected ? "var(--ui-accent)" : "var(--ui-stroke-secondary)"}`,
      background: "var(--ui-background-secondary)", color: "var(--ui-text-primary)", opacity: disabled ? 0.55 : 1,
      cursor: disabled ? "not-allowed" : "pointer",
    },
    children: [
      jsx("div", { style: { display: "flex", justifyContent: "space-between", gap: "12px" }, children: [
        jsx("strong", { children: title }),
        disabled ? jsx(Badge, { children: "Ainda não implementado" }) : null,
      ] }),
      jsx("p", { style: textStyle("13px", "var(--ui-text-secondary)"), children: description }),
      children,
    ],
  });
}

function Section({ title, children }) {
  return jsx("section", { style: { marginTop: "24px" }, children: [
    jsx("h2", { style: textStyle("16px"), children: title }), children,
  ] });
}

function QueryMessage({ query, loading, empty, error }) {
  if (query.isLoading) return jsx("p", { style: textStyle("13px", "var(--ui-text-secondary)"), children: loading });
  if (query.isError) return jsx("p", { role: "alert", style: textStyle("13px", "var(--ui-text-secondary)"), children: error });
  if (empty) return jsx("p", { style: textStyle("13px", "var(--ui-text-secondary)"), children: empty });
  return null;
}

function groupCapabilities(operations) {
  return operations.reduce((groups, item) => {
    const group = groups.find((candidate) => candidate.service === item.service);
    if (group) group.items.push(item);
    else groups.push({ service: item.service, items: [item] });
    return groups;
  }, []);
}

function Microsoft365Page({ ctx }) {
  const [draft, setDraft] = useState(INITIAL_DRAFT);
  const [loaded, setLoaded] = useState(false);
  const key = storageName();
  const configuration = useDashboardQuery(ctx, "/configuration", isConfigurationResponse);
  const capabilities = useDashboardQuery(ctx, "/capabilities", isCapabilitiesResponse);
  const preflight = useDashboardQuery(ctx, "/preflight", isPreflightResponse);

  useEffect(() => {
    let active = true;
    Promise.resolve(ctx.storage.get(key)).then((saved) => {
      if (active && saved && typeof saved === "object") {
        setDraft({ ...INITIAL_DRAFT, ...saved, step: Math.max(0, Math.min(STEPS.length - 1, Number(saved.step) || 0)) });
      }
      if (active) setLoaded(true);
    }).catch(() => { if (active) setLoaded(true); });
    return () => { active = false; };
  }, [ctx, key]);

  useEffect(() => {
    if (loaded) ctx.storage.set(key, { step: draft.step, mode: draft.mode, selected: draft.selected });
  }, [ctx, draft, key, loaded]);

  const update = (patch) => setDraft((current) => ({ ...current, ...patch }));
  const next = () => update({ step: Math.min(STEPS.length - 1, draft.step + 1) });
  const previous = () => update({ step: Math.max(0, draft.step - 1) });
  const capabilityRows = capabilities.data ? capabilities.data.operations : [];
  const capabilityGroups = groupCapabilities(capabilityRows);
  const configuredMode = configuration.data && configuration.data.authentication_mode;
  const preflightData = preflight.data;

  return jsx("main", { style: { maxWidth: "940px", padding: "32px", margin: "0 auto" }, children: [
    jsx("header", { children: [
      jsx("div", { style: { display: "flex", alignItems: "center", gap: "10px" }, children: [
        jsx("h1", { style: { ...textStyle("26px"), margin: 0 }, children: "Conecte o Microsoft 365 ao Hermes" }),
        jsx(Badge, { children: "Alpha" }),
      ] }),
      jsx("p", { style: textStyle("14px", "var(--ui-text-secondary)"), children: "Configuração local e segura para revisar capacidades antes de qualquer integração." }),
    ] }),
    jsx("nav", { "aria-label": "Etapas da configuração", style: { display: "flex", gap: "8px", marginTop: "28px", flexWrap: "wrap" }, children: STEPS.map((label, index) => jsx("button", {
      key: label, type: "button", onClick: () => update({ step: index }),
      style: { padding: "8px 10px", border: 0, borderBottom: `2px solid ${index === draft.step ? "var(--ui-accent)" : "var(--ui-stroke-secondary)"}`, background: "transparent", color: index === draft.step ? "var(--ui-text-primary)" : "var(--ui-text-secondary)" },
      children: `${index + 1}. ${label}`,
    })) }),
    jsx(Section, { title: "Visão geral", children: jsx("div", { style: { padding: "18px", border: "1px solid var(--ui-stroke-secondary)", borderRadius: "10px" }, children: [
      jsx("strong", { children: "Primeiro uso" }),
      jsx("p", { style: textStyle("14px", "var(--ui-text-secondary)"), children: "Escolha um modo e revise o catálogo. Os controles abaixo alteram somente este rascunho local." }),
      jsx(QueryMessage, { query: configuration, loading: "Carregando configuração…", error: "Não foi possível carregar a configuração." }),
      configuredMode ? jsx(Badge, { children: `Modo configurado: ${configuredMode}` }) : null,
    ] }) }),
    jsx(Section, { title: "Modo de acesso", children: [
      jsx(Card, { title: "Aplicativo administrado (Application)", description: "O aplicativo age conforme uma configuração administrada pelo tenant.", selected: draft.mode === "application", onClick: () => update({ mode: "application" }), children: jsx("span", { style: textStyle("12px", "var(--ui-text-tertiary)"), children: "Disponível para revisão local" }) }),
      jsx(Card, { title: "Conta Microsoft (Delegated)", description: "O Hermes agiria em nome da pessoa conectada, respeitando seus próprios acessos.", disabled: true, selected: false, children: jsx("span", { style: textStyle("12px", "var(--ui-text-tertiary)"), children: "Indisponível nesta primeira versão" }) }),
    ] }),
    jsx(Section, { title: "Capacidades disponíveis", children: [
      jsx(QueryMessage, { query: capabilities, loading: "Carregando capacidades…", error: "Não foi possível carregar as capacidades.", empty: "Nenhuma capacidade foi publicada pelo dashboard." }),
      capabilityGroups.length ? jsx("div", { style: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(230px, 1fr))", gap: "12px" }, children: capabilityGroups.map((group) => jsx("div", { key: group.service, style: { padding: "14px", border: "1px solid var(--ui-stroke-secondary)", borderRadius: "8px" }, children: [
        jsx("strong", { children: group.service }),
        jsx("ul", { style: { paddingLeft: "18px", marginBottom: 0 }, children: group.items.map((item) => jsx("li", { key: `${item.service}.${item.operation}`, style: textStyle("13px", "var(--ui-text-secondary)"), children: `${item.operation}${item.write ? " (escrita)" : ""} — ${item.status.reason}` })) }),
      ] })) }) : null,
    ] }),
    jsx(Section, { title: "Pré-verificação local", children: jsx("div", { style: { padding: "16px", border: "1px solid var(--ui-stroke-secondary)", borderRadius: "10px" }, children: [
      jsx("strong", { children: "Ambiente local" }),
      jsx("p", { style: textStyle("13px", "var(--ui-text-secondary)"), children: "Esta verificação não contata a Microsoft e não confirma um tenant." }),
      jsx(QueryMessage, { query: preflight, loading: "Executando pré-verificação…", error: "Não foi possível executar a pré-verificação." }),
      preflightData ? jsx("div", { children: [
        jsx("p", { style: textStyle("13px", "var(--ui-text-secondary)"), children: preflightData.locally_ready ? "Pronto para revisão local." : "Revisão local bloqueada até corrigir os requisitos informados pelo dashboard." }),
        jsx(Badge, { children: `Verificação remota: ${preflightData.remote_verification}` }),
      ] }) : null,
    ] }) }),
    jsx("footer", { style: { display: "flex", justifyContent: "space-between", marginTop: "28px" }, children: [
      jsx(Button, { onClick: previous, disabled: draft.step === 0, children: "Voltar" }),
      jsx(Button, { onClick: next, children: draft.step === STEPS.length - 1 ? "Concluir revisão local" : "Continuar" }),
    ] }),
  ] });
}

export default {
  id: "microsoft365",
  name: "Microsoft 365",
  defaultEnabled: false,
  register(ctx) {
    ctx.register({ id: "microsoft365-route", area: ROUTES_AREA, data: { path: "/microsoft365" }, render: () => jsx(Microsoft365Page, { ctx }) });
    ctx.register({ id: "microsoft365-sidebar", area: SIDEBAR_NAV_AREA, data: { path: "/microsoft365", label: "Microsoft 365", codicon: "cloud" } });
    ctx.register({ id: "microsoft365-command", area: PALETTE_AREA, data: { label: "Configurar Microsoft 365", keywords: ["Microsoft", "365", "configuração"] }, run: () => host.navigate("/microsoft365") });
  },
};
