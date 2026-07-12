// src/auth/LoginSSO.jsx
//
// Tela de login do Cuide, usando o MESMO Firebase Authentication do
// MiContas (não mais comparação de hash no navegador). O fluxo espelha
// o caminho PRINCIPAL do MiContas (signIn do Firebase Auth) — ver
// auth.js -> fazerLogin(). Continua com a MESMA base de e-mail (o
// MiContas indexa por e-mail, não por CPF).
//
// O QUE MUDOU EM RELAÇÃO À VERSÃO ANTERIOR:
//  - Antes: buscava o membro em micontas_v15 e comparava hash SHA-256
//    no navegador. Isso reimplementava (mais fraco) só o caminho de
//    fallback do MiContas, sem nunca passar pelo Firebase Auth de verdade.
//  - Agora: chama signInWithEmailAndPassword do Firebase Auth primeiro
//    (mesma validação de servidor que o MiContas usa). A leitura de
//    micontas_v15 continua existindo, mas só para pegar nome/role — não
//    para validar senha.
//
// CASO DE BORDA IMPORTANTE: membros do MiContas que nunca fizeram login
// desde a migração de segurança (meados de junho/2026) podem ainda não
// ter uma conta no Firebase Authentication — só senha em hash dentro do
// Realtime Database (fluxo "legado" do MiContas). Para esses, o login
// aqui vai falhar com uma mensagem pedindo para entrar uma vez no
// MiContas primeiro (isso migra a conta automaticamente lá). Não
// replico aqui o fallback de hash do MiContas de propósito: criar
// contas no Firebase Auth "pelas costas" do MiContas, a partir do
// Cuide, é uma decisão de arquitetura maior que prefiro não tomar sem
// sua confirmação explícita — se quiser esse comportamento, eu adiciono.

import { useState } from "react";
import { signInWithEmailAndPassword } from "firebase/auth";
import { auth } from "../services/firebase";
import { buscarContaPorEmail } from "../services/micontasApi";
import { useTemaCuide, ROTULO_TEMA } from "../theme/cuideTheme";

export function LoginSSO({ onLogin }) {
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [erro, setErro] = useState("");
  const [loading, setLoading] = useState(false);
  const [showSenha, setShowSenha] = useState(false);
  const { themeKey, tema: t, ciclarTema } = useTemaCuide();

  async function handleLogin(e) {
    e.preventDefault();
    setErro("");
    setLoading(true);
    try {
      const emailLimpo = email.trim().toLowerCase();

      // Mesmo caminho principal do MiContas: valida no Firebase Auth.
      await signInWithEmailAndPassword(auth, emailLimpo, senha);

      // Senha OK — agora só busca nome/role em micontas_v15 (leitura).
      const resultado = await buscarContaPorEmail(emailLimpo);
      if (!resultado) {
        throw { code: "cuide/sem-cadastro" };
      }
      const { emailKey, membro, membroIndex } = resultado;

      onLogin({
        email: emailLimpo,
        emailKey,
        membroIndex,
        nome: membro.nome,
        role: membro.acesso || "membro",
      });
    } catch (err) {
      if (err && err.code === "cuide/sem-cadastro") {
        setErro("Login válido, mas este e-mail não tem cadastro no MiContas.");
      } else if (
        err && (err.code === "auth/wrong-password" ||
        err.code === "auth/user-not-found" ||
        err.code === "auth/invalid-credential")
      ) {
        setErro(
          "Usuário ou senha incorretos. Se é seu primeiro acesso desde a atualização de segurança, entre uma vez no MiContas antes de tentar aqui."
        );
      } else {
        setErro((err && err.message) || "Falha ao entrar.");
      }
    } finally {
      setLoading(false);
    }
  }

  return (
    <div style={{
      minHeight: "100vh",
      background: t.bg,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontFamily: "'Inter', system-ui, sans-serif",
      padding: 16,
      transition: "background .2s",
    }}>
      <button
        type="button"
        onClick={ciclarTema}
        title="Alternar tema"
        style={{
          position: "fixed", top: 16, right: 16,
          background: t.card, border: `1px solid ${t.border}`,
          borderRadius: 20, padding: "6px 14px",
          color: t.text2, fontSize: 13, cursor: "pointer",
        }}
      >
        {ROTULO_TEMA[themeKey]}
      </button>

      <div style={{
        background: t.card,
        border: `1px solid ${t.border}`,
        borderRadius: 20,
        padding: "48px 40px",
        width: "100%",
        maxWidth: 420,
        boxShadow: "0 25px 60px rgba(0,0,0,.4)",
        transition: "background .2s, border-color .2s",
      }}>
        <div style={{ textAlign: "center", marginBottom: 36 }}>
          <div style={{
            width: 64, height: 64, borderRadius: 16,
            background: t.accent,
            display: "flex", alignItems: "center", justifyContent: "center",
            margin: "0 auto 16px",
            fontSize: 28,
            boxShadow: `0 8px 24px rgba(${t.accentRgb},.35)`,
          }}>🩺</div>
          <h1 style={{ color: t.text, fontSize: 26, fontWeight: 700, margin: 0 }}>
            Cuide
          </h1>
          <p style={{ color: t.text2, fontSize: 13, marginTop: 6 }}>
            Módulo de Gestão Clínica — MiContas
          </p>
        </div>

        <form onSubmit={handleLogin}>
          <label style={{
            display: "block", color: t.text3, fontSize: 12, fontWeight: 600,
            marginBottom: 8, textTransform: "uppercase", letterSpacing: ".05em",
          }}>
            E-mail
          </label>
          <input
            type="email"
            placeholder="voce@exemplo.com"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            required
            style={{
              width: "100%", boxSizing: "border-box",
              background: t.inp, border: `1px solid ${t.border}`,
              borderRadius: 10, padding: "13px 16px",
              color: t.text, fontSize: 15, outline: "none",
            }}
          />

          <label style={{
            display: "block", color: t.text3, fontSize: 12, fontWeight: 600,
            marginTop: 18, marginBottom: 8, textTransform: "uppercase", letterSpacing: ".05em",
          }}>
            Senha
          </label>
          <div style={{ position: "relative" }}>
            <input
              type={showSenha ? "text" : "password"}
              placeholder="••••••••"
              value={senha}
              onChange={(e) => setSenha(e.target.value)}
              required
              style={{
                width: "100%", boxSizing: "border-box",
                background: t.inp, border: `1px solid ${t.border}`,
                borderRadius: 10, padding: "13px 44px 13px 16px",
                color: t.text, fontSize: 15, outline: "none",
              }}
            />
            <button
              type="button"
              onClick={() => setShowSenha((s) => !s)}
              style={{
                position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)",
                background: "none", border: "none", cursor: "pointer",
                color: t.text2, fontSize: 18, padding: 0,
              }}
            >
              {showSenha ? "🙈" : "👁️"}
            </button>
          </div>

          {erro && (
            <div style={{
              background: "#3f0a0a", border: "1px solid #f8717130",
              borderRadius: 8, padding: "10px 14px",
              color: "#f87171", fontSize: 13, margin: "18px 0",
              lineHeight: 1.4,
            }}>
              ⚠️ {erro}
            </div>
          )}

          <button
            type="submit"
            disabled={loading}
            style={{
              width: "100%",
              background: loading ? t.border : t.accent,
              border: "none", borderRadius: 10,
              padding: 14, color: "#fff",
              fontSize: 15, fontWeight: 600,
              cursor: loading ? "not-allowed" : "pointer",
              marginTop: 24,
              boxShadow: loading ? "none" : `0 4px 16px rgba(${t.accentRgb},.3)`,
            }}
          >
            {loading ? "Verificando..." : "Entrar"}
          </button>
        </form>

        <p style={{ color: t.text3, fontSize: 11, textAlign: "center", marginTop: 28 }}>
          Mesmo login usado no MiContas
        </p>
      </div>
    </div>
  );
}
