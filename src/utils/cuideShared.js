// src/utils/cuideShared.js
//
// Cópia dos helpers/constantes do Cuide.jsx usados pelas páginas que agora
// são carregadas sob demanda (React.lazy). Mantidos idênticos ao original
// para não alterar nenhum comportamento — apenas isolados em módulo próprio
// para permitir o code-splitting.

export const ESP_CORES={"Nutrição":"#FF8C00","Terapia Ocupacional":"#DAA520","Supervisão ABA":"#C2185B","Fisioterapia":"#6A0572","Natação Terapêutica":"#008080","Psicopedagogia":"#A0522D","Psicomotricidade":"#556B2F","Musicoterapia":"#BDB76B","Arteterapia":"#FFB6C1","Aplicador ABA (A.T.)":"#FF00FF","Terapia Assistida por Animais":"#9932CC","Psicologia Convencional/Psicanálise":"#8B4513","Fonoaudiologia":"#9370DB","Avaliação Neuropsicológica":"#FFFACD","Avaliação Neuropsicopedagógica":"#FFFACD","Coordenação ABA":"#3b82f6","Neurofeedback":"#06b6d4","Psicoterapia":"#7c3aed","Snoezelen":"#84cc16","Agendamento":"#22c55e","Faturamento":"#f59e0b","Atendimento":"#60a5fa","Supervisão ADM":"#e879f9","Administrador":"#f87171","Financeiro":"#34d399","Gestão de Pessoas":"#a78bfa","Outro":"#94a3b8"};
export const espCor=e=>ESP_CORES[e]||"#94a3b8";

export const STATUS_CHAMADO=["aberto","andamento","encerrado","novo_paciente","novo_agendamento","devolvido"];
export const COR_CHAMADO={aberto:"#f59e0b",andamento:"#3b82f6",encerrado:"#10b981",novo_paciente:"#a78bfa",novo_agendamento:"#38bdf8",devolvido:"#f87171"};
export const LABEL_CHAMADO={aberto:"📂 Aberto",andamento:"🔄 Andamento",encerrado:"✅ Encerrado",novo_paciente:"👤 Novo Paciente",novo_agendamento:"📅 Novo Agendamento",devolvido:"↩️ Devolvido"};
export const PERFIL_LABEL={profissional:"Profissional",atendimento:"Atendimento",secretaria:"Secretária",supervisor_adm:"Supervisor ADM",coordenador:"Coordenador",administrador:"Administrador",faturamento_supervisor:"Sup. Faturamento",gestao_pessoas:"Gestão de Pessoas",agendamento:"Agendamento"};

export const STATUS_AG={
  agendado:      {label:"Agendado",          color:"#94a3b8", icon:"📋", credConv:false, repasse:false, descr:"Agendado"},
  atendido:      {label:"Atendido",          color:"#adff2f", icon:"✔️", credConv:true,  repasse:true,  numAut:true,  descr:"Realizado — pede Nº autorização"},
  faltou:        {label:"Faltou",            color:"#ff4444", icon:"🚫", credConv:false, repasse:false, descr:"Paciente faltou sem aviso"},
  faltou_pacote: {label:"Faltou (Pacote)",   color:"#8b0000", icon:"📦", credConv:true,  repasse:true,  descr:"Paciente pacote faltou"},
  n_registrado:  {label:"Não Registrado",    color:"#d1d5db", icon:"⚪", credConv:false, repasse:false, autoAlert:true, descr:"Passou horário sem status"},
  desmarcou_pac: {label:"Paciente Desmarcou",color:"#111827", icon:"⬛", credConv:false, repasse:false, descr:"Falta justificada c/ antecedência"},
  desmarcou_prof:{label:"Prof. Desmarcou",   color:"#4b5563", icon:"🔘", credConv:false, repasse:false, descr:"Desmarcado pelo profissional"},
  faturado:      {label:"Faturado",    color:"#22c55e", icon:"💵", credConv:true, repasse:false, fatRole:true, descr:"Crédito de convênio confirmado"},
  cancelado:     {label:"Cancelado",         color:"#475569", icon:"❌", credConv:false, repasse:false, descr:"Cancelado"},
};

export const DIAS_SEMANA=["Domingo","Segunda","Terça","Quarta","Quinta","Sexta","Sábado"];
export const TURNOS_H=["07:00","07:30","08:00","08:30","09:00","09:30","10:00","10:30","11:00","11:30","12:00","12:30","13:00","13:30","14:00","14:30","15:00","15:30","16:00","16:30","17:00","17:30","18:00","18:30","19:00","19:30","20:00"];

export const hoje=new Date();
export const ymd=d=>{const dd=new Date(d);return dd.getFullYear()+"-"+String(dd.getMonth()+1).padStart(2,"0")+"-"+String(dd.getDate()).padStart(2,"0");};
export const brDate=s=>{if(!s)return"";const[y,m,d]=s.split("-");return d+"/"+m+"/"+y;};
export const toMin=h=>{const[a,b]=String(h||"00:00").split(":").map(Number);return a*60+b;};
export const toTime=t=>String(Math.floor(t/60)%24).padStart(2,"0")+":"+String(t%60).padStart(2,"0");
export const addMin=(h,m)=>toTime(toMin(h)+Number(m||0));
export const overlaps=(a1,a2,b1,b2)=>toMin(a1)<toMin(b2)&&toMin(b1)<toMin(a2);
export const brl=v=>new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"}).format(Number(v||0));
export const profShort=n=>{if(!n)return"";const p=n.split(" ");return p.length<=2?n:p[0]+" "+p[p.length-1];};
export const rawD=v=>(v||"").replace(/\D/g,"");
export const maskCPF=v=>rawD(v).slice(0,11).replace(/(\d{3})(\d)/,"$1.$2").replace(/(\d{3})(\d)/,"$1.$2").replace(/(\d{3})(\d{1,2})$/,"$1-$2");
export const maskPhone=v=>{const d=rawD(v).slice(0,11);return d.length<=10?d.replace(/(\d{2})(\d{4})(\d{0,4})/,"($1) $2-$3"):d.replace(/(\d{2})(\d{5})(\d{0,4})/,"($1) $2-$3");};
export const diffDays=(d1,d2)=>Math.ceil((new Date(d2)-new Date(d1))/(1000*86400));
export const hoje_str=ymd(hoje);

export function exportCSV(nome, cabecalho, linhas) {
  const esc = v => '"' + String(v ?? "").replace(/"/g, '""') + '"';
  const csv = [cabecalho, ...linhas].map(r => r.map(esc).join(",")).join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = nome + ".csv"; a.click();
  URL.revokeObjectURL(url);
}
