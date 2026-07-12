// src/components/AppSwitcher.jsx
//
// Barra fina no topo do Cuide para alternar entre os módulos do ecossistema
// Apoiare, e para trocar o tema (escuro/médio/claro) — fica aqui porque é
// a barra fixa visível em toda navegação logada, diferente da tela de login.

import { useTemaCuide, ROTULO_TEMA } from "../theme/cuideTheme";

export function AppSwitcher({ session, onLogout }) {
  const { themeKey, tema: t, ciclarTema } = useTemaCuide();

  const bar = {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "10px 16px",
    background: t.card,
    borderBottom: `1px solid ${t.border}`,
    fontFamily: "'Inter', system-ui, sans-serif",
  };
  const icon = {
    width: 32,
    height: 32,
    borderRadius: 8,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    textDecoration: "none",
    fontSize: 16,
    background: t.inp,
    border: `1px solid ${t.border}`,
  };
  const active = { background: t.accent, borderColor: t.accent };
  const themeBtn = {
    background: t.inp,
    border: `1px solid ${t.border}`,
    color: t.text2,
    borderRadius: 20,
    padding: "5px 12px",
    cursor: "pointer",
    fontSize: 12,
  };
  const logoutBtn = {
    background: "transparent",
    border: `1px solid ${t.border}`,
    color: t.text2,
    borderRadius: 8,
    padding: "6px 12px",
    cursor: "pointer",
    fontSize: 13,
  };

  return (
    <div style={bar}>
      <div style={{ display: "flex", gap: 8 }}>
        <a
          href="https://cuide.apoiare.com.br"
          style={{ ...icon, ...active }}
          title="Cuide (Clínica)"
        >
          🩺
        </a>
        <a
          href="https://micontas.com.br"
          style={icon}
          title="MiContas (Financeiro)"
        >
          📊
        </a>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <button onClick={ciclarTema} title="Alternar tema" style={themeBtn}>
          {ROTULO_TEMA[themeKey]}
        </button>
        <span style={{ color: t.text2, fontSize: 13 }}>{session.nome}</span>
        <button onClick={onLogout} style={logoutBtn}>
          Sair
        </button>
      </div>
    </div>
  );
}
