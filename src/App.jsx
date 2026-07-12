// src/App.jsx
//
// Ponto de entrada do Cuide. Não altera nada do MiContas — apenas consome
// o mesmo Firebase. O CuideSistema importado abaixo é o app clínico que
// você já tem pronto (Cuide.jsx, ~6900 linhas) — nenhuma linha dele muda
// aqui, ele só passa a rodar "dentro" dessa casca de sessão/SSO.

import { useState, useEffect, lazy, Suspense } from "react";
import { RequireCuideAccess } from "./auth/RequireCuideAccess";
import { LoginSSO } from "./auth/LoginSSO";
import { AppSwitcher } from "./components/AppSwitcher";

// Carregado sob demanda: o bundle do Cuide (~900KB) só é baixado
// depois que a sessão existir, então a tela de login fica leve.
const CuideSistema = lazy(() => import("./Cuide"));

const SESSION_KEY = "cuide_session_v1";

function saveSession(s) {
  localStorage.setItem(SESSION_KEY, JSON.stringify(s));
}
function loadSession() {
  try {
    const s = localStorage.getItem(SESSION_KEY);
    return s ? JSON.parse(s) : null;
  } catch {
    return null;
  }
}
function clearSession() {
  localStorage.removeItem(SESSION_KEY);
}

export default function App() {
  const [session, setSession] = useState(null);
  const [checked, setChecked] = useState(false);

  useEffect(() => {
    const saved = loadSession();
    if (saved) setSession(saved);
    setChecked(true);
  }, []);

  if (!checked) return null; // evita flash de tela de login

  if (!session) {
    return (
      <LoginSSO
        onLogin={(s) => {
          saveSession(s);
          setSession(s);
        }}
      />
    );
  }

  function handleLogout() {
    clearSession();
    setSession(null);
  }

  return (
    <RequireCuideAccess session={session}>
      <AppSwitcher session={session} onLogout={handleLogout} />
      <Suspense fallback={<div style={{padding:40,textAlign:"center"}}>Carregando…</div>}>
        <CuideSistema onLogout={handleLogout} initialAuth={session} />
      </Suspense>
    </RequireCuideAccess>
  );
}
