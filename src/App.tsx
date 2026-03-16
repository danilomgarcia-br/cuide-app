import { useState, useEffect } from "react";
import CuideSistema from "./Cuide";

// ── tipos ──────────────────────────────────────────────────────────────────
interface SessionUser {
  usuario: string;
  nome: string;
  role: string;
}

// ── helpers localStorage ───────────────────────────────────────────────────
const SESSION_KEY = "cuide_session_v1";
function saveSession(u: SessionUser) {
  localStorage.setItem(SESSION_KEY, JSON.stringify(u));
}
function loadSession(): SessionUser | null {
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

// ── tela de login ──────────────────────────────────────────────────────────
function LoginScreen({ onLogin }: { onLogin: (u: SessionUser) => void }) {
  const [usuario, setUsuario] = useState("");
  const [senha, setSenha] = useState("");
  const [erro, setErro] = useState("");
  const [loading, setLoading] = useState(false);
  const [showSenha, setShowSenha] = useState(false);

  // profissionais seed — espelho do Cuide.jsx
  const SEED_PROF = [
    { id: 1, nome: "Dra. Ana Souza",      usuario: "11111111111", senha: "280409Dm@", role: "profissional" },
    { id: 2, nome: "Dr. Carlos Lima",     usuario: "22222222222", senha: "280409Dm@", role: "profissional" },
    { id: 3, nome: "Coordenador Silva",   usuario: "33333333333", senha: "280409Dm@", role: "coordenador" },
    { id: 4, nome: "Admin Sistema",       usuario: "31028313896", senha: "280409Dm@", role: "administrador" },
    { id: 5, nome: "Faturamento User",    usuario: "55555555555", senha: "280409Dm@", role: "faturamento_supervisor" },
  ];

  function handleLogin(e: React.FormEvent) {
    e.preventDefault();
    setErro("");
    setLoading(true);
    setTimeout(() => {
      const u = usuario.replace(/\D/g, "");
      const prof = SEED_PROF.find(
        (p) => p.usuario === u && p.senha === senha
      );
      if (prof) {
        const session: SessionUser = { usuario: prof.usuario, nome: prof.nome, role: prof.role };
        saveSession(session);
        onLogin(session);
      } else {
        setErro("Usuário ou senha incorretos.");
      }
      setLoading(false);
    }, 600);
  }

  const maskCPF = (v: string) => {
    const d = v.replace(/\D/g, "").slice(0, 11);
    return d
      .replace(/(\d{3})(\d)/, "$1.$2")
      .replace(/(\d{3})(\d)/, "$1.$2")
      .replace(/(\d{3})(\d{1,2})$/, "$1-$2");
  };

  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(135deg, #07090f 0%, #0c1525 50%, #07090f 100%)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontFamily: "'Inter', system-ui, sans-serif",
      padding: "16px",
    }}>
      {/* card */}
      <div style={{
        background: "#0c1017",
        border: "1px solid #151f30",
        borderRadius: "20px",
        padding: "48px 40px",
        width: "100%",
        maxWidth: "420px",
        boxShadow: "0 25px 60px rgba(0,0,0,.6)",
      }}>
        {/* logo */}
        <div style={{ textAlign: "center", marginBottom: "36px" }}>
          <div style={{
            width: 64, height: 64, borderRadius: "16px",
            background: "linear-gradient(135deg, #3b82f6, #6366f1)",
            display: "flex", alignItems: "center", justifyContent: "center",
            margin: "0 auto 16px",
            fontSize: 28,
            boxShadow: "0 8px 24px rgba(99,102,241,.35)",
          }}>🩺</div>
          <h1 style={{ color: "#e2e8f0", fontSize: 26, fontWeight: 700, margin: 0 }}>Cuide</h1>
          <p style={{ color: "#64748b", fontSize: 13, marginTop: 6 }}>
            Sistema de Gestão Clínica
          </p>
        </div>

        {/* form */}
        <form onSubmit={handleLogin}>
          {/* usuário (CPF) */}
          <div style={{ marginBottom: "18px" }}>
            <label style={{ display: "block", color: "#94a3b8", fontSize: 12, fontWeight: 600, marginBottom: 8, textTransform: "uppercase", letterSpacing: ".05em" }}>
              CPF (usuário)
            </label>
            <input
              type="text"
              inputMode="numeric"
              placeholder="000.000.000-00"
              value={maskCPF(usuario)}
              onChange={e => setUsuario(e.target.value.replace(/\D/g, ""))}
              required
              style={{
                width: "100%", boxSizing: "border-box",
                background: "#0a0f1c", border: "1px solid #182035",
                borderRadius: "10px", padding: "13px 16px",
                color: "#e2e8f0", fontSize: 15, outline: "none",
                transition: "border .2s",
              }}
              onFocus={e => e.target.style.borderColor = "#3b82f6"}
              onBlur={e => e.target.style.borderColor = "#182035"}
            />
          </div>

          {/* senha */}
          <div style={{ marginBottom: "24px" }}>
            <label style={{ display: "block", color: "#94a3b8", fontSize: 12, fontWeight: 600, marginBottom: 8, textTransform: "uppercase", letterSpacing: ".05em" }}>
              Senha
            </label>
            <div style={{ position: "relative" }}>
              <input
                type={showSenha ? "text" : "password"}
                placeholder="••••••••"
                value={senha}
                onChange={e => setSenha(e.target.value)}
                required
                style={{
                  width: "100%", boxSizing: "border-box",
                  background: "#0a0f1c", border: "1px solid #182035",
                  borderRadius: "10px", padding: "13px 44px 13px 16px",
                  color: "#e2e8f0", fontSize: 15, outline: "none",
                  transition: "border .2s",
                }}
                onFocus={e => e.target.style.borderColor = "#3b82f6"}
                onBlur={e => e.target.style.borderColor = "#182035"}
              />
              <button
                type="button"
                onClick={() => setShowSenha(s => !s)}
                style={{
                  position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)",
                  background: "none", border: "none", cursor: "pointer",
                  color: "#64748b", fontSize: 18, padding: 0,
                }}
              >
                {showSenha ? "🙈" : "👁️"}
              </button>
            </div>
          </div>

          {/* erro */}
          {erro && (
            <div style={{
              background: "#3f0a0a", border: "1px solid #f8717130",
              borderRadius: "8px", padding: "10px 14px",
              color: "#f87171", fontSize: 13, marginBottom: "18px",
              display: "flex", alignItems: "center", gap: 8,
            }}>
              ⚠️ {erro}
            </div>
          )}

          {/* botão */}
          <button
            type="submit"
            disabled={loading}
            style={{
              width: "100%",
              background: loading ? "#1e3a5f" : "linear-gradient(135deg, #3b82f6, #6366f1)",
              border: "none", borderRadius: "10px",
              padding: "14px", color: "#fff",
              fontSize: 15, fontWeight: 600, cursor: loading ? "not-allowed" : "pointer",
              transition: "opacity .2s, transform .1s",
              boxShadow: "0 4px 16px rgba(99,102,241,.3)",
            }}
          >
            {loading ? "Verificando..." : "Entrar"}
          </button>
        </form>

        <p style={{ color: "#2d3f5a", fontSize: 11, textAlign: "center", marginTop: 28, marginBottom: 0 }}>
          © 2025 Cuide — Todos os direitos reservados
        </p>
      </div>
    </div>
  );
}

// ── App principal ──────────────────────────────────────────────────────────
export default function App() {
  const [session, setSession] = useState<SessionUser | null>(null);
  const [checked, setChecked] = useState(false);

  useEffect(() => {
    const saved = loadSession();
    if (saved) setSession(saved);
    setChecked(true);
  }, []);

  function handleLogout() {
    clearSession();
    setSession(null);
  }

  if (!checked) return null; // evita flash

  if (!session) {
    return <LoginScreen onLogin={setSession} />;
  }

  // injeta logout no sistema via prop (o CuideSistema já chama onLogout se receber)
  return <CuideSistema onLogout={handleLogout} initialAuth={session} />;
}
