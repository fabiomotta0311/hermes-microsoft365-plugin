import { Button, Badge, ROUTES_AREA, SIDEBAR_NAV_AREA, PALETTE_AREA } from "@hermes/plugin-sdk";
import { useEffect, useState } from "react";
import { jsx } from "react/jsx-runtime";

// This slice intentionally stays local: the backend contract is expected to expose
// sanitized GET /capabilities and GET /preflight responses, plus a later validated
// configuration write. TODO: connect that contract through ctx.rest in a later slice.
// Never put credentials or remote authority in ctx.storage; only this visual draft lives there.

const STORAGE_KEY = "microsoft365.desktop-draft";
const STEPS = ["Funcionalidades", "Modo de acesso", "Identidade", "Permissões", "Revisão"];
const CAPABILITIES = [
  { service: "Outlook", items: ["Pesquisar e-mails", "Ler mensagens"] },
  { service: "Calendário", items: ["Consultar eventos", "Criar eventos (indisponível)"] },
  { service: "SharePoint e OneDrive", items: ["Pesquisar arquivos", "Baixar arquivos"] },
  { service: "Teams", items: ["Listar equipes e canais", "Pesquisar mensagens"] },
  { service: "To Do e Planner", items: ["Consultar tarefas", "Criar tarefas (indisponível)"] },
];
const INITIAL_DRAFT = { step: 0, mode: "application", selected: [] };

function storageName(host) {
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
    type: "button",
    disabled,
    onClick,
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
    jsx("h2", { style: textStyle("16px"), children: title }),
    children,
  ] });
}

function Microsoft365Page({ ctx, host }) {
  const [draft, setDraft] = useState(INITIAL_DRAFT);
  const [loaded, setLoaded] = useState(false);
  const key = storageName(host);

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
      jsx(Badge, { children: "Status: não verificado em um tenant" }),
    ] }) }),
    jsx(Section, { title: "Modo de acesso", children: [
      jsx(Card, { title: "Aplicativo administrado (Application)", description: "O aplicativo age conforme uma configuração administrada pelo tenant.", selected: draft.mode === "application", onClick: () => update({ mode: "application" }), children: jsx("span", { style: textStyle("12px", "var(--ui-text-tertiary)"), children: "Disponível para revisão local" }) }),
      jsx(Card, { title: "Conta Microsoft (Delegated)", description: "O Hermes agiria em nome da pessoa conectada, respeitando seus próprios acessos.", disabled: true, selected: false, children: jsx("span", { style: textStyle("12px", "var(--ui-text-tertiary)"), children: "Indisponível nesta primeira versão" }) }),
    ] }),
    jsx(Section, { title: "Capacidades disponíveis", children: jsx("div", { style: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(230px, 1fr))", gap: "12px" }, children: CAPABILITIES.map((group) => jsx("div", { key: group.service, style: { padding: "14px", border: "1px solid var(--ui-stroke-secondary)", borderRadius: "8px" }, children: [
      jsx("strong", { children: group.service }),
      jsx("ul", { style: { paddingLeft: "18px", marginBottom: 0 }, children: group.items.map((item) => jsx("li", { key: item, style: textStyle("13px", "var(--ui-text-secondary)"), children: item })) }),
    ] })) }) }),
    jsx(Section, { title: "Pré-verificação local", children: jsx("div", { style: { padding: "16px", border: "1px solid var(--ui-stroke-secondary)", borderRadius: "10px" }, children: [
      jsx("strong", { children: "Ambiente local" }),
      jsx("p", { style: textStyle("13px", "var(--ui-text-secondary)"), children: "Pronto para revisar o rascunho. Esta verificação não contata a Microsoft." }),
      jsx(Badge, { children: "Somente local" }),
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
    const host = ctx.host;
    ctx.register({ id: "microsoft365-route", area: ROUTES_AREA, data: { path: "/microsoft365" }, render: () => jsx(Microsoft365Page, { ctx, host }) });
    ctx.register({ id: "microsoft365-sidebar", area: SIDEBAR_NAV_AREA, data: { path: "/microsoft365", label: "Microsoft 365", codicon: "cloud" } });
    ctx.register({ id: "microsoft365-command", area: PALETTE_AREA, data: { label: "Configurar Microsoft 365", keywords: ["Microsoft", "365", "configuração"] }, run: () => host.navigate("/microsoft365") });
  },
};
