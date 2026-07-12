// src/services/micontasApi.js
//
// Leitura (somente leitura) da base de contas do MiContas, para o Cuide
// reconhecer o mesmo usuário/senha. Nunca escreve em micontas_v15 — quem
// é dono desses dados é o MiContas.
//
// Espelha a mesma resolução de "links" que o auth.js do MiContas usa
// (micontas_v15_links/{chave} -> chave real), para o caso de contas
// vinculadas a um titular diferente.

import { ref, get, child } from "firebase/database";
import { db } from "./firebase";

const NODE = "micontas_v15";
const LINKS_NODE = "micontas_v15_links";

function emailParaChave(email) {
  return email.trim().toLowerCase().replace(/\./g, "_");
}

async function resolverChave(email) {
  let chave = emailParaChave(email);
  const linkSnap = await get(child(ref(db), `${LINKS_NODE}/${chave}`));
  if (linkSnap.exists()) {
    chave = String(linkSnap.val()).replace(/\./g, "_");
  }
  return chave;
}

async function getConta(chave) {
  const snap = await get(child(ref(db), `${NODE}/${chave}`));
  return snap.exists() ? snap.val() : null;
}

function encontrarMembro(conta, email) {
  if (!conta || !Array.isArray(conta.membros)) return null;
  const emailLower = email.trim().toLowerCase();
  const idx = conta.membros.findIndex(
    (m) => (m.email || "").toLowerCase().trim() === emailLower
  );
  return idx === -1 ? null : { membro: conta.membros[idx], index: idx };
}

/**
 * Busca a conta MiContas pelo e-mail (já resolvendo links) e retorna
 * o membro correspondente. Não valida senha — isso é feito pelo
 * Firebase Authentication antes de chamar esta função (ver LoginSSO.jsx).
 *
 * @returns {Promise<{emailKey:string, membro:object, membroIndex:number}|null>}
 */
export async function buscarContaPorEmail(email) {
  const emailLimpo = email.trim().toLowerCase();
  const chave = await resolverChave(emailLimpo);
  const conta = await getConta(chave);
  const achado = encontrarMembro(conta, emailLimpo);
  if (!achado) return null;
  return { emailKey: chave, membro: achado.membro, membroIndex: achado.index };
}

/**
 * true somente se o membro tiver o campo `cuideAtivo: true` liberado
 * pelo administrador da conta no MiContas.
 */
export async function verificarAcessoCuide(email) {
  const resultado = await buscarContaPorEmail(email);
  return !!(resultado && resultado.membro && resultado.membro.cuideAtivo);
}
