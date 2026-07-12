// src/auth/RequireCuideAccess.jsx
//
// Bloqueia o acesso ao Cuide caso o membro ainda não tenha o campo
// `cuideAtivo: true` dentro do seu próprio objeto em micontas_v15/{emailKey}/membros[].
//
// Repare que este arquivo NÃO sabe nada sobre a estrutura do JSON — ele só
// chama verificarAcessoCuide(email), que já faz a busca no nó certo e o
// find() dentro do array de membros. Se a estrutura do banco mudar de novo,
// só o micontasApi.js precisa ser ajustado, este arquivo continua igual.

import { useEffect, useState } from "react";
import { verificarAcessoCuide } from "../services/micontasApi";
import { useTemaCuide } from "../theme/cuideTheme";

export function RequireCuideAccess({ session, children }) {
  const [status, setStatus] = useState("checking"); // checking | allowed | blocked
  const { tema: t } = useTemaCuide();

  useEffect(() => {
    let vivo = true;
    verificarAcessoCuide(session.email)
      .then((liberado) => {
        if (vivo) setStatus(liberado ? "allowed" : "blocked");
      })
      .catch(() => {
        if (vivo) setStatus("blocked");
      });
    return () => {
      vivo = false;
    };
  }, [session.email]);

  const telaBase = {
    minHeight: "100vh",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    background: t.bg,
    color: t.text2,
    fontFamily: "'Inter', system-ui, sans-serif",
    padding: 16,
  };

  if (status === "checking") {
    return <div style={telaBase}>Verificando acesso...</div>;
  }

  if (status === "blocked") {
    return (
      <div style={telaBase}>
        <div style={{ textAlign: "center", maxWidth: 380 }}>
          <div style={{ fontSize: 40, marginBottom: 16 }}>🔒</div>
          <h2 style={{ color: t.text, margin: "0 0 8px" }}>
            Módulo Cuide não ativado
          </h2>
          <p style={{ color: t.text2, fontSize: 14, lineHeight: 1.5 }}>
            Fale com o administrador da sua conta no MiContas para liberar o
            acesso ao módulo de gestão clínica.
          </p>
        </div>
      </div>
    );
  }

  return children;
}
