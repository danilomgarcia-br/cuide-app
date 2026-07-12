// src/pages/GestaoPage.jsx
//
// Extraído de Cuide.jsx para permitir code-splitting: esta página
// (Gestão + Relatórios/Atividades/Importar/Exportar) só é baixada
// quando o usuário realmente acessa a aba "gestao".
// Nenhuma lógica interna foi alterada — apenas os helpers que antes
// vinham do escopo do módulo agora são importados de cuideShared.js.

import { useState } from "react";
import * as XLSX from "xlsx";
import { ymd, brDate, hoje_str, STATUS_AG, STATUS_CHAMADO, COR_CHAMADO, LABEL_CHAMADO, PERFIL_LABEL, DIAS_SEMANA, TURNOS_H, maskCPF, maskPhone, exportCSV, espCor, profShort, brl, diffDays, rawD, toMin } from "../utils/cuideShared";

function RelatoriosTab({agenda,pacientes,profissionais,procedimentos,showToast}){
  const [perRel_rt,setPerRel_rt]=useState("semana");
  const [fProfRel_rt,setFProfRel_rt]=useState("");
  const diasBackR={hoje:1,semana:7,mes:30,trimestre:90}[perRel_rt]||7;
  const dtLimR=new Date();dtLimR.setDate(dtLimR.getDate()-diasBackR);
  const agRel=agenda.filter(a=>{
  const dt=new Date(a.data);if(dt<dtLimR)return false;
  if(fProfRel_rt&&String(a.profissionalId)!==String(fProfRel_rt))return false;
  return true;
  });
  const statusCountR=Object.keys(STATUS_AG).reduce((acc,k)=>({...acc,[k]:agRel.filter(a=>a.status===k).length}),{});
  const espMapR={};
  agRel.forEach(a=>{const p=profissionais.find(x=>x.id===Number(a.profissionalId));const esp=(p?.especialidades||[p?.especialidade])[0]||"Outros";espMapR[esp]=(espMapR[esp]||0)+1;});
  const espArrR=Object.entries(espMapR).sort((a,b)=>b[1]-a[1]).slice(0,8);
  const maxEspR=Math.max(1,...espArrR.map(e=>e[1]));
  const fatCountR=agRel.filter(a=>a.status==="faturado").length;
  const receitaR=agRel.filter(a=>a.status==="faturado").reduce((s,a)=>{const p=procedimentos.find(x=>x.id===Number(a.procedimentoId));return s+(p?.valor||0);},0);
  const dayL_rt=[],dayA_rt=[],dayF_rt=[];
  for(let i=6;i>=0;i--){const d=new Date();d.setDate(d.getDate()-i);const ds=ymd(d);dayL_rt.push(DIAS_SEMANA[d.getDay()].slice(0,3));dayA_rt.push(agenda.filter(a=>a.data===ds&&["atendido","faturado"].includes(a.status)).length);dayF_rt.push(agenda.filter(a=>a.data===ds&&["faltou","faltou_pacote"].includes(a.status)).length);}
  const maxDR=Math.max(1,...dayA_rt,...dayF_rt);
  const exportAgRel=()=>{
  const cab=["Data","Horário","Paciente","Profissional","Especialidade","Convênio","Status","Nº Autorização"];
  const linhas=agRel.map(a=>{const pac=pacientes.find(p=>p.id===Number(a.pacienteId));const pr=profissionais.find(p=>p.id===Number(a.profissionalId));return[brDate(a.data),a.horarioSessao,pac?.nome||"",pr?.nome||"",(pr?.especialidades||[""])[0]||"",a.convenio||"",STATUS_AG[a.status]?.label||a.status,a.numAutorizacao||""];});
  exportCSV("agenda_"+perRel_rt,cab,linhas);
  if(showToast)showToast("📥 CSV exportado","ok");
  };
  const exportExcelRel=()=>{
  if(typeof XLSX==="undefined")return;
  const wb=XLSX.utils.book_new();
  const rows=agRel.map(a=>{const pac=pacientes.find(p=>p.id===Number(a.pacienteId));const pr=profissionais.find(p=>p.id===Number(a.profissionalId));return{Data:brDate(a.data),Horário:a.horarioSessao,Paciente:pac?.nome||"—",Profissional:pr?.nome||"—",Status:STATUS_AG[a.status]?.label||a.status,Convênio:a.convenio||""};});
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Agenda");
  const pacRows=pacientes.map(p=>({Nome:p.nome,CPF:maskCPF(p.cpf||""),Nascimento:brDate(p.nascimento||""),Convênio:p.convenio||"",Plano:p.plano||""}));
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(pacRows),"Pacientes");
  XLSX.writeFile(wb,"Cuide_Relatorio_"+hoje_str+".xlsx");
  if(showToast)showToast("✅ Excel gerado","ok");
  };
  return(<div>
  <div style={{display:"flex",gap:8,marginBottom:16,alignItems:"center",flexWrap:"wrap"}}>
    <select value={perRel_rt} onChange={e=>setPerRel_rt(e.target.value)} style={{fontSize:12}}>
      <option value="hoje">Hoje</option><option value="semana">Últimos 7 dias</option>
      <option value="mes">Último mês</option><option value="trimestre">Último trimestre</option>
    </select>
    <select value={fProfRel_rt} onChange={e=>setFProfRel_rt(e.target.value)} style={{fontSize:12}}>
      <option value="">Todos profissionais</option>
      {profissionais.map(p=><option key={p.id} value={p.id}>{profShort(p.nome)}</option>)}
    </select>
    <div style={{marginLeft:"auto",display:"flex",gap:6}}>
      <button className="btn secondary" style={{fontSize:11}} onClick={exportAgRel}>📥 CSV Agenda</button>
      <button className="btn primary" style={{fontSize:11,background:"#166534",borderColor:"#166534"}} onClick={exportExcelRel}>📊 Exportar Excel</button>
    </div>
  </div>
  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:14}}>
    <div className="card" style={{padding:16,gridColumn:"span 2"}}>
      <div style={{fontWeight:900,fontSize:13,marginBottom:10}}>📈 Atendidos vs Faltas — Últimos 7 dias</div>
      <div style={{display:"flex",alignItems:"flex-end",gap:6,height:90}}>
        {dayL_rt.map((d,i)=>(
          <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:2}}>
            <div style={{display:"flex",gap:1,alignItems:"flex-end",height:70}}>
              <div style={{width:10,background:"#34d399",borderRadius:"3px 3px 0 0",height:Math.max(2,Math.round((dayA_rt[i]/maxDR)*60))+"px"}} title={"Atend: "+dayA_rt[i]}/>
              <div style={{width:10,background:"#f87171",borderRadius:"3px 3px 0 0",height:Math.max(2,Math.round((dayF_rt[i]/maxDR)*60))+"px"}} title={"Faltas: "+dayF_rt[i]}/>
            </div>
            <div style={{fontSize:9,color:"var(--mt)",fontWeight:700}}>{d}</div>
          </div>
        ))}
      </div>
      <div style={{display:"flex",gap:12,marginTop:6,fontSize:10,color:"var(--mt)"}}>
        <span><span style={{display:"inline-block",width:8,height:8,background:"#34d399",borderRadius:2,marginRight:3}}/>Atendidos</span>
        <span><span style={{display:"inline-block",width:8,height:8,background:"#f87171",borderRadius:2,marginRight:3}}/>Faltas</span>
      </div>
    </div>
    <div className="card" style={{padding:16}}>
      <div style={{fontWeight:900,fontSize:13,marginBottom:10}}>🔵 Distribuição de Status</div>
      <div style={{display:"flex",alignItems:"center",gap:12}}>
        <DonutChart size={90} data={Object.entries(STATUS_AG).filter(([k])=>statusCountR[k]>0).map(([k,v])=>({v:statusCountR[k],c:v.color}))}/>
        <div style={{flex:1}}>
          {Object.entries(STATUS_AG).filter(([k])=>statusCountR[k]>0).slice(0,6).map(([k,v])=>(
            <div key={k} style={{display:"flex",alignItems:"center",gap:5,marginBottom:4,fontSize:11}}>
              <span style={{width:8,height:8,borderRadius:"50%",background:v.color,flexShrink:0,display:"inline-block"}}/>
              <span style={{color:"var(--mt)",flex:1}}>{v.label}</span>
              <b>{statusCountR[k]}</b>
            </div>
          ))}
        </div>
      </div>
    </div>
    <div className="card" style={{padding:16}}>
      <div style={{fontWeight:900,fontSize:13,marginBottom:10}}>💰 Faturamento do Período</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
        <div style={{background:"#052e1c30",borderRadius:8,padding:"10px 12px"}}>
          <div style={{fontSize:10,color:"var(--mt)",fontWeight:800}}>FATURADAS</div>
          <div style={{fontSize:22,fontWeight:900,color:"#34d399"}}>{fatCountR}</div>
        </div>
        <div style={{background:"#1e3a5f30",borderRadius:8,padding:"10px 12px"}}>
          <div style={{fontSize:10,color:"var(--mt)",fontWeight:800}}>RECEITA EST.</div>
          <div style={{fontSize:16,fontWeight:900,color:"#a78bfa"}}>{brl(receitaR)}</div>
        </div>
      </div>
    </div>
    <div className="card" style={{padding:16}}>
      <div style={{fontWeight:900,fontSize:13,marginBottom:10}}>👩‍⚕️ Sessões por Especialidade</div>
      {espArrR.length===0&&<div className="muted" style={{fontSize:12}}>Sem dados no período.</div>}
      {espArrR.map(([esp,cnt])=>(
        <div key={esp} style={{marginBottom:7}}>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:2}}>
            <span style={{color:"var(--mt)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:160}}>{esp}</span>
            <span style={{fontWeight:900,color:espCor(esp)}}>{cnt}</span>
          </div>
          <div style={{height:5,borderRadius:3,background:"var(--sx)",overflow:"hidden"}}>
            <div style={{width:Math.round((cnt/maxEspR)*100)+"%",height:"100%",background:espCor(esp),borderRadius:3}}/>
          </div>
        </div>
      ))}
    </div>
  </div>
  </div>);
  }

function AtividadesTab({atividades,profissionais}){
  const [fUser,setFUser]=useState("");
  const [fAcao,setFAcao]=useState("");
  const [fData,setFData]=useState("");
  const lista=(atividades||[]).slice().reverse();
  const users=[...new Set(lista.map(a=>a.usuario).filter(Boolean))];
  const acoes=[...new Set(lista.map(a=>a.acao).filter(Boolean))];
  const filtradas=lista.filter(a=>
  (!fUser||a.usuario===fUser)&&
  (!fAcao||a.acao===fAcao)&&
  (!fData||a.data===fData)
  );
  return(<div>
  <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:14,alignItems:"center"}}>
    <input type="date" value={fData} onChange={e=>setFData(e.target.value)} style={{width:140,fontSize:12}}/>
    <select value={fUser} onChange={e=>setFUser(e.target.value)} style={{minWidth:160,fontSize:12}}>
      <option value="">Todos os usuários</option>
      {users.map(u=><option key={u}>{u}</option>)}
    </select>
    <select value={fAcao} onChange={e=>setFAcao(e.target.value)} style={{minWidth:140,fontSize:12}}>
      <option value="">Todas as ações</option>
      {acoes.map(a=><option key={a}>{a}</option>)}
    </select>
    {(fUser||fAcao||fData)&&<button className="btn secondary small" onClick={()=>{setFUser("");setFAcao("");setFData("");}}>✕ Limpar</button>}
    <span style={{marginLeft:"auto",fontSize:11,color:"var(--mt)"}}>{filtradas.length} registro(s)</span>
  </div>
  <div className="card" style={{overflow:"hidden"}}>
    <div className="grid-header" style={{gridTemplateColumns:"95px 58px 150px 120px 1fr"}}>
      <div>Data</div><div>Hora</div><div>Usuário</div><div>Ação</div><div>Detalhe</div>
    </div>
    {filtradas.length===0&&<div style={{padding:18,textAlign:"center"}} className="muted">Nenhuma atividade registrada no sistema.</div>}
    {filtradas.map((a,i)=>(
      <div key={a.id||i} className="grid-row" style={{gridTemplateColumns:"95px 58px 150px 120px 1fr",background:i%2?"var(--gr)":""}}>
        <div style={{fontSize:11}}>{brDate(a.data)}</div>
        <div style={{fontSize:11,color:"#a78bfa",fontWeight:700}}>{a.hora}</div>
        <div style={{fontSize:11,fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.usuario}</div>
        <div><span style={{padding:"2px 7px",borderRadius:20,fontSize:10,fontWeight:800,background:"#7c6af720",color:"#a78bfa"}}>{a.acao}</span></div>
        <div style={{fontSize:11,color:"var(--mt)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.detalhe}</div>
      </div>
    ))}
  </div>
  </div>);
  }

function ImportarTab({agenda,setAgenda,pacientes,setPacientes,profissionais,setProfissionais,auth,showToast}){
  if(auth.role!=="administrador")return(<div style={{padding:48,textAlign:"center",color:"var(--mt)",fontSize:14}}>🔒 Acesso restrito ao Administrador.</div>);
  const [tipo,setTipo]=useState("pacientes");
  const [arquivo,setArquivo]=useState(null);
  const [linhas,setLinhas]=useState([]);
  const [cabecalho,setCabecalho]=useState([]);
  const [mapa,setMapa]=useState({});
  const [preview,setPreview]=useState([]);
  const CAMPOS_PAC=["nome","nascimento","cpf","sexo","celular","convenio","plano","cep","logradouro","numero","bairro","cidade","estado","resp1Nome","resp1Cpf","resp1Whatsapp","infoImportantes"];
  const CAMPOS_PROF=["nome","usuario","senha","role","especialidade"];
  const CAMPOS_AG=["pacienteNome","profissionalNome","data","horarioSessao","tempoSessao","convenio","status"];
  const campos=tipo==="pacientes"?CAMPOS_PAC:tipo==="profissionais"?CAMPOS_PROF:CAMPOS_AG;
  const resetar=()=>{setLinhas([]);setCabecalho([]);setPreview([]);setMapa({});setArquivo(null);};
  const lerArquivo=e=>{
    const f=e.target.files[0];if(!f)return;setArquivo(f.name);
    const r=new FileReader();
    r.onload=ev=>{
      const txt=ev.target.result;
      const rows=txt.split(/\r?\n/).filter(l=>l.trim());
      if(rows.length<2)return;
      const sep=rows[0].includes(";")?";":",";
      const cab=rows[0].split(sep).map(x=>x.trim().replace(/^["']|["']$/g,""));
      const ls=rows.slice(1).map(row=>row.split(sep).map(x=>x.trim().replace(/^["']|["']$/g,"")));
      setCabecalho(cab);setLinhas(ls);setPreview(ls.slice(0,3));
      const autoMapa={};
      cab.forEach((h,i)=>{const hL=h.toLowerCase();const campo=campos.find(c=>c.toLowerCase()===hL||hL.includes(c.toLowerCase()));if(campo)autoMapa[campo]=i;});
      setMapa(autoMapa);
    };
    r.readAsText(f,"UTF-8");
  };
  const importarDados=()=>{
    if(!linhas.length)return;
    if(tipo==="pacientes"){
      const novos=linhas.map((l,idx)=>{
        const get=c=>mapa[c]!==undefined?l[mapa[c]]||"":"";
        return{id:Date.now()+idx,nome:get("nome"),nascimento:get("nascimento"),cpf:rawD(get("cpf")),sexo:get("sexo"),celular:rawD(get("celular")),convenio:get("convenio"),plano:get("plano"),cep:get("cep"),logradouro:get("logradouro"),numero:get("numero"),bairro:get("bairro"),cidade:get("cidade"),estado:get("estado"),resp1Nome:get("resp1Nome"),resp1Cpf:rawD(get("resp1Cpf")),resp1Whatsapp:rawD(get("resp1Whatsapp")),infoImportantes:get("infoImportantes"),arquivos:[]};
      }).filter(p=>p.nome);
      setPacientes(prev=>[...prev,...novos]);
      if(showToast)showToast("✅ "+novos.length+" pacientes importados com sucesso","ok");resetar();
    } else if(tipo==="profissionais"){
      const novos=linhas.map((l,idx)=>{
        const get=c=>mapa[c]!==undefined?l[mapa[c]]||"":"";
        return{id:Date.now()+idx,nome:get("nome"),usuario:rawD(get("usuario")),senha:get("senha"),role:get("role")||"profissional",especialidades:[get("especialidade")].filter(Boolean),especialidade:get("especialidade"),temposAtendimento:[50],escala:{},filiaisAtendimento:[],filialAcesso:[],carimbo:null};
      }).filter(p=>p.nome);
      setProfissionais(prev=>[...prev,...novos]);
      if(showToast)showToast("✅ "+novos.length+" profissionais importados","ok");resetar();
    } else {
      const novos=linhas.map((l,idx)=>{
        const get=c=>mapa[c]!==undefined?l[mapa[c]]||"":"";
        const pac=pacientes.find(p=>p.nome.toLowerCase()===get("pacienteNome").toLowerCase());
        const prof=profissionais.find(p=>p.nome.toLowerCase()===get("profissionalNome").toLowerCase());
        return{id:Date.now()+idx,pacienteId:pac?.id||null,profissionalId:prof?.id||null,data:get("data"),horarioSessao:get("horarioSessao")||"08:00",horarioFimSessao:"",tempoSessao:Number(get("tempoSessao"))||50,status:get("status")||"agendado",convenio:get("convenio")||"",tipo:"sessao"};
      }).filter(a=>a.data);
      setAgenda(prev=>[...prev,...novos]);
      if(showToast)showToast("✅ "+novos.length+" agendamentos importados","ok");resetar();
    }
  };
  return(<div>
    <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
      {["pacientes","profissionais","agendamentos"].map(t=>(
        <button key={t} className={"btn "+(tipo===t?"primary":"secondary")} onClick={()=>{setTipo(t);resetar();}} style={{fontSize:11,textTransform:"capitalize"}}>{t}</button>
      ))}
    </div>
    <div className="section-box" style={{marginBottom:14}}>
      <div className="section-title">📂 Arquivo CSV</div>
      <div style={{fontSize:11,color:"var(--mt)",marginBottom:8}}>Selecione um arquivo .csv com cabeçalho na primeira linha. Separador: vírgula ou ponto-e-vírgula.</div>
      <label className="btn secondary" style={{cursor:"pointer",fontSize:11}}>
        📂 Selecionar arquivo {arquivo&&<span style={{marginLeft:6,color:"#34d399"}}>✓ {arquivo}</span>}
        <input type="file" accept=".csv,.txt" onChange={lerArquivo} style={{display:"none"}}/>
      </label>
      {arquivo&&<button className="btn secondary small" onClick={resetar} style={{marginLeft:8}}>✕ Limpar</button>}
    </div>
    {cabecalho.length>0&&<>
      <div className="section-box" style={{marginBottom:14}}>
        <div className="section-title">🗂️ Mapeamento de Colunas</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:8}}>
          {campos.map(campo=>(
            <div key={campo}>
              <label>{campo}</label>
              <select value={mapa[campo]!==undefined?mapa[campo]:""} onChange={e=>setMapa(m=>({...m,[campo]:e.target.value===""?undefined:Number(e.target.value)}))}>
                <option value="">— ignorar —</option>
                {cabecalho.map((h,i)=><option key={i} value={i}>{h}</option>)}
              </select>
            </div>
          ))}
        </div>
      </div>
      {preview.length>0&&<div className="section-box" style={{marginBottom:14}}>
        <div className="section-title">👁️ Pré-visualização (3 primeiras linhas)</div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",fontSize:11,borderCollapse:"collapse"}}>
            <thead><tr>{cabecalho.map((h,i)=><th key={i} style={{padding:"4px 8px",background:"var(--sx)",borderBottom:"1px solid var(--cb)",textAlign:"left",fontWeight:700}}>{h}</th>)}</tr></thead>
            <tbody>{preview.map((row,i)=><tr key={i}>{row.map((cell,j)=><td key={j} style={{padding:"4px 8px",borderBottom:"1px solid var(--cb)",color:"var(--mt)"}}>{cell}</td>)}</tr>)}</tbody>
          </table>
        </div>
      </div>}
      <div style={{display:"flex",gap:8}}>
        <button className="btn secondary" onClick={resetar}>✕ Cancelar</button>
        <button className="btn primary" onClick={importarDados}>📥 Importar {linhas.length} registro(s)</button>
      </div>
    </>}
  </div>);
}

function ExportarTab({agenda,pacientes,profissionais,procedimentos,salas,filiais,chamados,showToast}){
  const [formato,setFormato]=useState("xlsx");
  const [modulos,setModulos]=useState({agenda:true,pacientes:true,profissionais:false,financeiro:false,chamados:false,salas:false});
  const [periodoEx,setPeriodoEx]=useState("mes");
  const toggleMod=k=>setModulos(m=>({...m,[k]:!m[k]}));
  const diasBack={hoje:1,semana:7,mes:30,trimestre:90,tudo:3650}[periodoEx]||30;
  const dtLim=new Date();dtLim.setDate(dtLim.getDate()-diasBack);
  const agEx=periodoEx==="tudo"?agenda:agenda.filter(a=>new Date(a.data)>=dtLim);
  const exportar=()=>{
    if(typeof XLSX==="undefined"){if(showToast)showToast("❌ Biblioteca XLSX não carregada","err");return;}
    const wb=XLSX.utils.book_new();
    if(modulos.agenda){
      const rows=agEx.map(a=>{
        const pac=pacientes.find(p=>p.id===Number(a.pacienteId));
        const pr=profissionais.find(p=>p.id===Number(a.profissionalId));
        const proc=procedimentos.find(p=>p.id===Number(a.procedimentoId));
        return{Data:brDate(a.data),Horario:a.horarioSessao,"Hora Fim":a.horarioFimSessao||"",Paciente:pac?.nome||"—",Profissional:pr?.nome||"—",Especialidade:(pr?.especialidades||[""])[0]||"",Procedimento:proc?.nome||"",Convenio:a.convenio||"",Status:STATUS_AG[a.status]?.label||a.status,"Nr Autorizacao":a.numAutorizacao||"",Observacoes:a.observacoes||""};
      });
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Agenda");
    }
    if(modulos.pacientes){
      const rows=pacientes.map(p=>({Nome:p.nome,CPF:maskCPF(p.cpf||""),Nascimento:brDate(p.nascimento||""),Sexo:p.sexo||"",Celular:maskPhone(p.celular||""),Email:p.email||"",Convenio:p.convenio||"",Plano:p.plano||"",Logradouro:p.logradouro||"",Numero:p.numero||"",Bairro:p.bairro||"",Cidade:p.cidade||"",Estado:p.estado||"",Responsavel:p.resp1Nome||"","Cel Responsavel":p.resp1Whatsapp||""}));
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Pacientes");
    }
    if(modulos.profissionais){
      const rows=profissionais.map(p=>({Nome:p.nome,CPF:maskCPF(p.usuario||""),Especialidades:(p.especialidades||[]).join(", "),Perfil:PERFIL_LABEL[p.role]||p.role,"Nivel Repasse":p.nivelRepasse||"","Forma Pgto":p.formaPagamento||"",Email:p.email||"",Celular:maskPhone(p.celular||"")}));
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Profissionais");
    }
    if(modulos.financeiro){
      const fat=agEx.filter(a=>a.status==="faturado");
      const rows=fat.map(a=>{
        const pac=pacientes.find(p=>p.id===Number(a.pacienteId));
        const pr=profissionais.find(p=>p.id===Number(a.profissionalId));
        const proc=procedimentos.find(p=>p.id===Number(a.procedimentoId));
        const valor=proc?.valor||0;
        const nivelKey=(pr?.nivelRepasse||"Pleno").replace("ú","u");
        const pct=(proc?.niveisRepasse||{})[nivelKey]||45;
        const repasse=Math.round(valor*(pct/100)*100)/100;
        return{Data:brDate(a.data),Paciente:pac?.nome||"—",Profissional:pr?.nome||"—",Procedimento:proc?.nome||"","Valor (R$)":valor,"Repasse %":pct,"Repasse (R$)":repasse,"Valor Clinica (R$)":Math.round((valor-repasse)*100)/100,Convenio:a.convenio||""};
      });
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Financeiro");
    }
    if(modulos.chamados){
      const rows=(chamados||[]).map(ch=>({Titulo:ch.titulo,Setor:ch.setor,Status:ch.status,Prioridade:ch.prioridade||"Normal",Descricao:ch.descricao||"",Abertura:brDate(ch.data||"")}));
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Chamados");
    }
    if(modulos.salas){
      const rows=(salas||[]).map(s=>{const f=filiais.find(x=>x.id===s.filialId);return{Sala:s.nome,Filial:f?.nome||"",Especialidade:s.especialidade||"","Custo Mensal":s.custoMensal||0,Ativa:s.ativa===false?"Não":"Sim"};});
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Salas");
    }
    if(wb.SheetNames.length===0){if(showToast)showToast("⚠️ Selecione ao menos um módulo","warn");return;}
    const ts=hoje_str.replace(/-/g,"");
    XLSX.writeFile(wb,"Cuide_Exportacao_"+ts+".xlsx");
    if(showToast)showToast("✅ Arquivo exportado com sucesso","ok");
  };
  const MODS=[
    {k:"agenda",label:"📅 Agenda",desc:"Agendamentos do período selecionado"},
    {k:"pacientes",label:"👤 Pacientes",desc:"Cadastro completo de pacientes"},
    {k:"profissionais",label:"🩺 Profissionais",desc:"Cadastro de profissionais"},
    {k:"financeiro",label:"💰 Financeiro",desc:"Sessões faturadas com repasse"},
    {k:"chamados",label:"📨 Chamados",desc:"Todos os chamados do sistema"},
    {k:"salas",label:"🏢 Salas",desc:"Salas e filiais cadastradas"},
  ];
  return(<div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:20}}>
      <div className="card" style={{padding:16}}>
        <div style={{fontWeight:900,fontSize:13,marginBottom:12,color:"var(--tx)"}}>📦 Módulos a exportar</div>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {MODS.map(m=>(<label key={m.k} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",borderRadius:8,background:modulos[m.k]?"#7c6af710":"transparent",border:"1px solid "+(modulos[m.k]?"#7c6af740":"var(--cpb)"),cursor:"pointer",transition:".15s"}}>
            <input type="checkbox" checked={!!modulos[m.k]} onChange={()=>toggleMod(m.k)} style={{width:15,height:15,accentColor:"#7c6af7"}}/>
            <div style={{flex:1}}>
              <div style={{fontWeight:700,fontSize:12,color:modulos[m.k]?"#a78bfa":"var(--tx)"}}>{m.label}</div>
              <div style={{fontSize:10,color:"var(--mt)",marginTop:1}}>{m.desc}</div>
            </div>
            {modulos[m.k]&&<span style={{fontSize:9,fontWeight:900,color:"#7c6af7",background:"#7c6af720",padding:"1px 6px",borderRadius:10}}>✓</span>}
          </label>))}
        </div>
      </div>
      <div className="card" style={{padding:16}}>
        <div style={{fontWeight:900,fontSize:13,marginBottom:12}}>⚙️ Configurações</div>
        <div style={{marginBottom:14}}>
          <label style={{fontSize:11,fontWeight:700,color:"var(--mt)",display:"block",marginBottom:6}}>PERÍODO (para Agenda e Financeiro)</label>
          <div style={{display:"flex",flexDirection:"column",gap:5}}>
            {[["hoje","Hoje"],["semana","Últimos 7 dias"],["mes","Último mês"],["trimestre","Último trimestre"],["tudo","Tudo (sem filtro de data)"]].map(([v,l])=>(
              <label key={v} style={{display:"flex",alignItems:"center",gap:7,cursor:"pointer",fontSize:12}}>
                <input type="radio" name="periodoEx" value={v} checked={periodoEx===v} onChange={()=>setPeriodoEx(v)} style={{accentColor:"#7c6af7"}}/>
                <span style={{color:periodoEx===v?"var(--tx)":"var(--mt)",fontWeight:periodoEx===v?700:400}}>{l}</span>
              </label>
            ))}
          </div>
        </div>
        <div style={{background:"var(--sx)",borderRadius:8,padding:"10px 12px",fontSize:11,color:"var(--mt)"}}>
          <div style={{fontWeight:700,color:"var(--tx)",marginBottom:4}}>📋 Resumo</div>
          <div>Agenda: <b style={{color:"#a78bfa"}}>{agEx.filter(()=>modulos.agenda).length || (modulos.agenda?agEx.length:0)}</b> registros</div>
          <div>Pacientes: <b style={{color:"#34d399"}}>{modulos.pacientes?pacientes.length:0}</b> registros</div>
          <div>Profissionais: <b style={{color:"#a78bfa"}}>{modulos.profissionais?profissionais.length:0}</b> registros</div>
          <div style={{marginTop:6,paddingTop:6,borderTop:"1px solid var(--sc)"}}>
            Módulos selecionados: <b style={{color:"var(--tx)"}}>{Object.values(modulos).filter(Boolean).length}</b>
          </div>
        </div>
      </div>
    </div>
    <div style={{display:"flex",justifyContent:"flex-end"}}>
      <button className="btn primary" style={{fontSize:13,padding:"10px 28px",background:"#166534",borderColor:"#166534",fontWeight:800}} onClick={exportar}>
        📤 Exportar Excel
      </button>
    </div>
  </div>);
}


function GestaoPage({agenda,pacientes,profissionais,procedimentos,convenios,salas,filiais,chamados,fila,pedidos,auth,showToast,atividades,setPacientes,setProfissionais,setAgenda}){
  const [tab,setTab]=useState("overview");
  const [periodo,setPeriodo]=useState("mes");
  const [selProf,setSelProf]=useState("");

  // ── Ranges ────────────────────────────────────────────────────────────
  const getRange=()=>{
    const d=new Date();
    if(periodo==="hoje")return{ini:hoje_str,fim:hoje_str};
    if(periodo==="semana"){const x=new Date(d);x.setDate(x.getDate()-6);return{ini:ymd(x),fim:hoje_str};}
    if(periodo==="mes"){const x=new Date(d);x.setDate(1);return{ini:ymd(x),fim:hoje_str};}
    if(periodo==="trimestre"){const x=new Date(d);x.setMonth(x.getMonth()-2);x.setDate(1);return{ini:ymd(x),fim:hoje_str};}
    return{ini:hoje_str,fim:hoje_str};
  };
  const {ini,fim}=getRange();
  const agP=agenda.filter(a=>a.data>=ini&&a.data<=fim);
  const agHoje=agenda.filter(a=>a.data===hoje_str);

  // ── KPIs principais ────────────────────────────────────────────────────
  const total=agP.length;
  const atend=agP.filter(a=>["atendido","faturado"].includes(a.status)).length;
  const faltas=agP.filter(a=>["faltou","faltou_pacote"].includes(a.status)).length;
  const canc=agP.filter(a=>["cancelado","desmarcou_pac","desmarcou_prof"].includes(a.status)).length;
  const fat=agP.filter(a=>a.status==="faturado").length;
  const receitaEst=agP.filter(a=>a.status==="faturado").reduce((s,a)=>{const p=procedimentos.find(x=>x.id===Number(a.procedimentoId));return s+(p?.valor||0);},0);
  const naoFat=agP.filter(a=>a.status==="nao_faturado").length;
  const receitaPerdas=agP.filter(a=>a.status==="nao_faturado").reduce((s,a)=>{const p=procedimentos.find(x=>x.id===Number(a.procedimentoId));return s+(p?.valor||0);},0);
  const txPresenca=total>0?Math.round((atend/total)*100):0;
  const txFalta=total>0?Math.round((faltas/total)*100):0;
  const txCanc=total>0?Math.round((canc/total)*100):0;
  const txFat=atend>0?Math.round((fat/atend)*100):0;

  // ── Hoje ──────────────────────────────────────────────────────────────
  const hojeTot=agHoje.length;
  const hojePend=agHoje.filter(a=>["agendado","confirmado"].includes(a.status)).length;
  const hojeAtend=agHoje.filter(a=>["atendido","faturado"].includes(a.status)).length;
  const hojeFalta=agHoje.filter(a=>["faltou","faltou_pacote"].includes(a.status)).length;

  // ── Evolução 14 dias ──────────────────────────────────────────────────
  const dias14=Array.from({length:14},(_,i)=>{const d=new Date();d.setDate(d.getDate()-13+i);const ds=ymd(d);const ags=agenda.filter(a=>a.data===ds);return{ds,total:ags.length,atend:ags.filter(a=>["atendido","faturado"].includes(a.status)).length,falta:ags.filter(a=>["faltou","faltou_pacote"].includes(a.status)).length};});
  const maxDia=Math.max(...dias14.map(d=>d.total),1);

  // ── Profissionais stats ───────────────────────────────────────────────
  const profStats=profissionais.filter(p=>["profissional","coordenador","coordenador_aba"].includes(p.role)).map(prof=>{
    const aps=agP.filter(a=>Number(a.profissionalId)===Number(prof.id));
    const at=aps.filter(a=>["atendido","faturado"].includes(a.status)).length;
    const fa=aps.filter(a=>["faltou","faltou_pacote"].includes(a.status)).length;
    const ft=aps.filter(a=>a.status==="faturado").length;
    const rec=aps.filter(a=>a.status==="faturado").reduce((s,a)=>{const p=procedimentos.find(x=>x.id===Number(a.procedimentoId));return s+(p?.valor||0);},0);
    const tx=aps.length>0?Math.round((at/aps.length)*100):0;
    const perf=tx>=85?"A":tx>=70?"B":tx>=50?"C":"D";
    return{...prof,tot:aps.length,at,fa,ft,rec,tx,perf};
  }).sort((a,b)=>b.at-a.at);

  // ── Salas ─────────────────────────────────────────────────────────────
  const salaStats=salas.filter(s=>s.ativa!==false).map(sala=>{
    const aps=agP.filter(a=>Number(a.salaId)===Number(sala.id));
    const dias=Math.max(1,Math.round((new Date(fim)-new Date(ini))/(86400000)));
    const diasUteis=Math.max(1,Math.round(dias*5/7));
    const cap=diasUteis*8;
    const pct=Math.min(100,Math.round((aps.length/cap)*100));
    const filial=filiais.find(f=>f.id===sala.filialId);
    const receita=aps.filter(a=>a.status==="faturado").reduce((s,a)=>{const p=procedimentos.find(x=>x.id===Number(a.procedimentoId));return s+(p?.valor||0);},0);
    const custo=Number(sala.custoMensal||0);
    const roi=custo>0?Math.round(((receita-custo)/custo)*100):null;
    // Ocupação por hora (TURNOS_H)
    const porHora=TURNOS_H.map(h=>{
      const qtd=aps.filter(a=>toMin(a.horarioSessao)<=toMin(h)&&toMin(a.horarioFimSessao||a.horarioSessao)>toMin(h)).length;
      const diasU=diasUteis;
      const pctH=Math.min(100,Math.round((qtd/Math.max(1,diasU))*100));
      return{hora:h,qtd,pct:pctH};
    });
    return{...sala,total:aps.length,pct,cap,filial:filial?.nome||"",receita,custo,roi,porHora};
  }).sort((a,b)=>b.pct-a.pct);

  // ── Convênios ─────────────────────────────────────────────────────────
  const convStats=agP.reduce((acc,a)=>{
    const cv=a.convenio||"Particular";
    if(!acc[cv])acc[cv]={nome:cv,total:0,atend:0,fat:0,rec:0};
    acc[cv].total++;
    if(["atendido","faturado"].includes(a.status))acc[cv].atend++;
    if(a.status==="faturado"){acc[cv].fat++;const p=procedimentos.find(x=>x.id===Number(a.procedimentoId));acc[cv].rec+=(p?.valor||0);}
    return acc;
  },{});
  const convArr=Object.values(convStats).sort((a,b)=>b.total-a.total);
  const maxConv=Math.max(...convArr.map(c=>c.total),1);

  // ── Chamados ──────────────────────────────────────────────────────────
  const chamAbertos=chamados.filter(c=>["aberto","andamento"].includes(c.status)).length;
  const chamEnc=chamados.filter(c=>c.status==="encerrado").length;
  const chamTotal=Math.max(chamados.length,1);
  const txResol=Math.round((chamEnc/chamTotal)*100);
  const chamPorSetor=chamados.reduce((acc,ch)=>{
    const s=ch.setor||"Outros";
    if(!acc[s])acc[s]={setor:s,total:0,abertos:0,enc:0};
    acc[s].total++;
    if(["aberto","andamento"].includes(ch.status))acc[s].abertos++;
    if(ch.status==="encerrado")acc[s].enc++;
    return acc;
  },{});

  // ── Fila espera ────────────────────────────────────────────────────────
  const filaEsp=fila.reduce((acc,f)=>{
    const e=f.especialidade||"Outros";
    if(!acc[e])acc[e]={esp:e,total:0};
    acc[e].total++;return acc;
  },{});

  // ── Pedidos vencendo ──────────────────────────────────────────────────
  const pedVenc7=pedidos.filter(p=>p.dataValidade&&diffDays(hoje_str,p.dataValidade)>=0&&diffDays(hoje_str,p.dataValidade)<=7);
  const pedVenc30=pedidos.filter(p=>p.dataValidade&&diffDays(hoje_str,p.dataValidade)>=0&&diffDays(hoje_str,p.dataValidade)<=30);
  const pedVencidos=pedidos.filter(p=>p.dataValidade&&diffDays(hoje_str,p.dataValidade)<0);

  // ── Indicadores de processo ───────────────────────────────────────────
  const ocupMedia=salaStats.length>0?Math.round(salaStats.reduce((s,x)=>s+x.pct,0)/salaStats.length):0;
  const custoSalasTotal=salas.reduce((s,sala)=>s+Number(sala.custoMensal||0),0);
  const receitaSalasTotal=salaStats.reduce((s,x)=>s+(x.receita||0),0);
  const roiSalas=custoSalasTotal>0?Math.round(((receitaSalasTotal-custoSalasTotal)/custoSalasTotal)*100):null;
  const indicadores=[
    {label:"Presença",val:txPresenca,meta:85,unit:"%",cor:"#34d399",rev:false,icon:"✅"},
    {label:"Faturamento",val:txFat,meta:70,unit:"%",cor:"#22c55e",rev:false,icon:"💰"},
    {label:"Resolução chamados",val:txResol,meta:80,unit:"%",cor:"#a78bfa",rev:false,icon:"📨"},
    {label:"Taxa de falta",val:txFalta,meta:15,unit:"%",cor:"#f87171",rev:true,icon:"🚫"},
    {label:"Cancelamentos",val:txCanc,meta:10,unit:"%",cor:"#f59e0b",rev:true,icon:"❌"},
    {label:"Ocupação salas",val:ocupMedia,meta:60,unit:"%",cor:"#a78bfa",rev:false,icon:"🏢"},
  ];

  // ── Alertas ───────────────────────────────────────────────────────────
  const alertas=[];
  if(txFalta>20)alertas.push({n:"alto",msg:"Taxa de faltas crítica: "+txFalta+"%",area:"Agenda"});
  if(chamAbertos>10)alertas.push({n:"alto",msg:chamAbertos+" chamados sem resolução",area:"Chamados"});
  if(pedVenc7.length>0)alertas.push({n:"alto",msg:pedVenc7.length+" pedido(s) vencendo em 7 dias",area:"Pedidos"});
  if(fila.length>15)alertas.push({n:"medio",msg:fila.length+" pacientes na fila de espera",area:"Fila"});
  profStats.filter(p=>p.tx<60&&p.tot>3).forEach(p=>alertas.push({n:"medio",msg:"Presença baixa: "+profShort(p.nome)+" ("+p.tx+"%)",area:"Profissional"}));
  if(alertas.length===0)alertas.push({n:"ok",msg:"Todos os indicadores dentro do esperado",area:"Geral"});

  const TB=({id,icon,label})=>(
    <button className="btn tab-btn" onClick={()=>setTab(id)}
      style={{background:tab===id?"var(--na)":"transparent",color:tab===id?"#7c6af7":"var(--mt)",fontWeight:tab===id?800:500}}>
      {icon} {label}
    </button>
  );

  const PillPeriodo=({v,l})=>(
    <button onClick={()=>setPeriodo(v)} style={{padding:"4px 12px",borderRadius:20,fontSize:11,fontWeight:800,border:"1.5px solid "+(periodo===v?"#7c6af7":"var(--cpb)"),background:periodo===v?"#7c6af715":"transparent",color:periodo===v?"#7c6af7":"var(--mt)",cursor:"pointer"}}>{l}</button>
  );

  return(<div className="page-wrap">
    {/* ── Header ── */}
    <div className="page-head">
      <h1>📊 Gestão — Processos & Indicadores</h1>
      <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
        <PillPeriodo v="hoje" l="Hoje"/>
        <PillPeriodo v="semana" l="7 dias"/>
        <PillPeriodo v="mes" l="Mês"/>
        <PillPeriodo v="trimestre" l="Trimestre"/>
      </div>
    </div>

    {/* ── Alertas operacionais ── */}
    <div style={{display:"flex",gap:7,flexWrap:"wrap",marginBottom:16}}>
      {alertas.map((a,i)=>(
        <div key={i} style={{display:"flex",alignItems:"center",gap:6,padding:"5px 13px",borderRadius:20,fontSize:11,fontWeight:700,
          background:a.n==="alto"?"#450a0a25":a.n==="medio"?"#451a0325":"#052e1c25",
          border:"1px solid "+(a.n==="alto"?"#f8717150":a.n==="medio"?"#f59e0b50":"#34d39950"),
          color:a.n==="alto"?"#f87171":a.n==="medio"?"#f59e0b":"#34d399"}}>
          {a.n==="alto"?"🚨":a.n==="medio"?"⚠️":"✅"}
          <span style={{fontWeight:900}}>{a.area}:</span> {a.msg}
        </div>
      ))}
    </div>

    {/* ── Snapshot de hoje ── */}
    <div style={{background:"var(--card)",border:"1px solid var(--cb)",borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:20,alignItems:"center",flexWrap:"wrap"}}>
      <div style={{fontWeight:900,fontSize:12,color:"var(--mt)",letterSpacing:".5px",textTransform:"uppercase",minWidth:60}}>Hoje</div>
      {[{l:"Total",v:hojeTot,c:"#a78bfa"},{l:"Atendidos",v:hojeAtend,c:"#34d399"},{l:"Pendentes",v:hojePend,c:"#f59e0b"},{l:"Faltas",v:hojeFalta,c:"#f87171"}].map(s=>(
        <div key={s.l} style={{textAlign:"center",padding:"0 10px",borderLeft:"1px solid var(--sc)"}}>
          <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:24,color:s.c,lineHeight:1}}>{s.v}</div>
          <div style={{fontSize:10,color:"var(--mt)",fontWeight:700,marginTop:2}}>{s.l}</div>
        </div>
      ))}
      <div style={{marginLeft:"auto",fontSize:11,color:"var(--mt)"}}>
        Receita est. período: <span style={{fontWeight:900,color:"#34d399",fontSize:14}}>{brl(receitaEst)}</span>
      </div>
    </div>

    {/* ── KPIs do período ── */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(145px,1fr))",gap:10,marginBottom:18}}>
      {[
        {icon:"📅",label:"Sessões",val:total,color:"#a78bfa",sub:"agendadas"},
        {icon:"✅",label:"Atendidos",val:atend,color:"#34d399",sub:txPresenca+"% presença"},
        {icon:"🚫",label:"Faltas",val:faltas,color:txFalta>15?"#f87171":"#f59e0b",sub:txFalta+"% do total",alert:txFalta>20},
        {icon:"❌",label:"Cancelados",val:canc,color:"#64748b",sub:txCanc+"% do total"},
        {icon:"💰",label:"Faturados",val:fat,color:"#22c55e",sub:txFat+"% dos atend."},
        {icon:"💵",label:"Receita Est.",val:brl(receitaEst),color:"#34d399",sub:"período",small:true},
        {icon:"📉",label:"Perdas",val:naoFat>0?brl(receitaPerdas):"—",color:naoFat>0?"#ef4444":"#64748b",sub:naoFat+" sessão(ões)",alert:naoFat>3,small:true},
        {icon:"📨",label:"Chamados",val:chamAbertos,color:chamAbertos>5?"#f87171":"#fb923c",sub:"em aberto",alert:chamAbertos>10},
        {icon:"⏳",label:"Fila Espera",val:fila.length,color:"#a78bfa",sub:"pacientes"},
      ].map(k=>(
        <div key={k.label} style={{background:"var(--card)",border:"1px solid var(--cb)",borderTop:"3px solid "+k.color,borderRadius:14,padding:"14px 16px",position:"relative",overflow:"hidden"}}>
          {k.alert&&<div style={{position:"absolute",top:6,right:8,fontSize:9,fontWeight:900,color:"#f87171",background:"#f8717120",padding:"1px 6px",borderRadius:20}}>ALERTA</div>}
          <div style={{fontSize:18,marginBottom:6}}>{k.icon}</div>
          <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:k.small?18:26,color:k.color,lineHeight:1,letterSpacing:"-1px"}}>{k.val}</div>
          <div style={{fontSize:10,fontWeight:800,letterSpacing:".6px",color:"var(--mt)",textTransform:"uppercase",marginTop:3}}>{k.label}</div>
          <div style={{fontSize:10,color:"var(--mt)",marginTop:2,fontStyle:"italic"}}>{k.sub}</div>
        </div>
      ))}
    </div>

    {/* ── Tabs ── */}
    <div className="tab-bar">
      <TB id="overview"      icon="📈" label="Visão Geral"/>
      <TB id="produtividade" icon="🩺" label="Produtividade"/>
      <TB id="processos"     icon="⚙️" label="Processos"/>
      <TB id="indicadores"   icon="🎯" label="Indicadores"/>
      <TB id="ocupacao"      icon="🏢" label="Salas & Filiais"/>
      <TB id="convenios"     icon="🏥" label="Convênios"/>
      <TB id="dashboard"    icon="📈" label="Dashboard"/>
      <TB id="relatorios"   icon="📊" label="Relatórios"/>
      <TB id="atividades"   icon="📋" label="Atividades"/>
      <TB id="importar"     icon="📥" label="Importar"/>
      <TB id="exportar"     icon="📤" label="Exportar"/>
    </div>

    {/* ══ VISÃO GERAL ══ */}
    {tab==="overview"&&<div style={{display:"grid",gridTemplateColumns:"1.4fr 1fr",gap:14}}>

      {/* Gráfico evolução 14 dias */}
      <div className="card" style={{padding:18}}>
        <div style={{fontWeight:900,fontSize:13,marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span>📈 Evolução diária — 14 dias</span>
          <div style={{display:"flex",gap:10,fontSize:10,color:"var(--mt)"}}>
            <span style={{display:"flex",alignItems:"center",gap:3}}><span style={{width:9,height:9,background:"#34d399",borderRadius:2,display:"inline-block"}}/> Atend.</span>
            <span style={{display:"flex",alignItems:"center",gap:3}}><span style={{width:9,height:9,background:"#f87171",borderRadius:2,display:"inline-block"}}/> Faltas</span>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"flex-end",gap:2,height:100,marginBottom:4}}>
          {dias14.map((d,i)=>{
            const hA=maxDia>0?Math.max(2,Math.round((d.atend/maxDia)*88)):2;
            const hF=maxDia>0?Math.max(2,Math.round((d.falta/maxDia)*88)):2;
            const isT=d.ds===hoje_str;
            return(<div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:1}}>
              {d.total>0&&<div style={{fontSize:8,color:"var(--mt)",fontWeight:700,marginBottom:1}}>{d.total}</div>}
              <div style={{display:"flex",gap:1,alignItems:"flex-end",height:88}}>
                <div style={{width:7,background:isT?"#a78bfa":"#34d399",borderRadius:"2px 2px 0 0",height:hA+"px",opacity:isT?1:.8}} title={"Atend: "+d.atend}/>
                <div style={{width:7,background:"#f87171",borderRadius:"2px 2px 0 0",height:hF+"px",opacity:.8}} title={"Faltas: "+d.falta}/>
              </div>
              {i%3===0&&<div style={{fontSize:7,color:"var(--mt)",fontWeight:700,marginTop:2}}>{d.ds.slice(8)+"/"+d.ds.slice(5,7)}</div>}
            </div>);
          })}
        </div>
      </div>

      {/* Status doughnut */}
      <div className="card" style={{padding:18}}>
        <div style={{fontWeight:900,fontSize:13,marginBottom:12}}>🔵 Distribuição de Status</div>
        <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:14}}>
          <DonutChart size={90} data={Object.entries(STATUS_AG).filter(([k])=>agP.filter(a=>a.status===k).length>0).map(([k,v])=>({v:agP.filter(a=>a.status===k).length,c:v.color}))}/>
          <div style={{flex:1}}>
            {Object.entries(STATUS_AG).filter(([k])=>agP.filter(a=>a.status===k).length>0).map(([k,v])=>{
              const cnt=agP.filter(a=>a.status===k).length;
              const pct=total>0?Math.round((cnt/total)*100):0;
              return(<div key={k} style={{display:"flex",alignItems:"center",gap:5,marginBottom:4}}>
                <span style={{width:7,height:7,borderRadius:"50%",background:v.color,flexShrink:0,display:"inline-block"}}/>
                <span style={{fontSize:10,color:"var(--mt)",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{v.label}</span>
                <span style={{fontSize:11,fontWeight:900,color:v.color,minWidth:20,textAlign:"right"}}>{cnt}</span>
                <span style={{fontSize:9,color:"var(--mt)",width:26,textAlign:"right"}}>{pct}%</span>
              </div>);
            })}
          </div>
        </div>
        <div style={{background:"var(--sx)",borderRadius:9,padding:"8px 12px"}}>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:4}}>
            <span style={{fontWeight:700}}>Taxa de Presença</span>
            <span style={{fontWeight:900,color:txPresenca>=80?"#34d399":txPresenca>=60?"#f59e0b":"#f87171",fontSize:13}}>{txPresenca}%</span>
          </div>
          <div style={{height:6,borderRadius:3,background:"var(--cb)",overflow:"hidden"}}>
            <div style={{width:txPresenca+"%",height:"100%",background:txPresenca>=80?"#34d399":txPresenca>=60?"#f59e0b":"#f87171",borderRadius:3,transition:"width .5s"}}/>
          </div>
        </div>
      </div>

      {/* Fila espera */}
      <div className="card" style={{padding:18}}>
        <div style={{fontWeight:900,fontSize:13,marginBottom:12,display:"flex",justifyContent:"space-between"}}>
          <span>⏳ Fila de Espera</span>
          <span style={{fontSize:11,color:"#a78bfa",fontWeight:800}}>{fila.length} pacientes</span>
        </div>
        {fila.length===0&&<div className="muted" style={{fontSize:12}}>Fila vazia.</div>}
        {Object.values(filaEsp).sort((a,b)=>b.total-a.total).map(f=>{
          const pct=Math.round((f.total/(fila.length||1))*100);
          const cor=espCor(f.esp);
          return(<div key={f.esp} style={{marginBottom:9}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}>
              <span style={{color:"var(--mt)",fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:160}}>{f.esp}</span>
              <span style={{fontWeight:900,color:cor,flexShrink:0}}>{f.total}</span>
            </div>
            <div style={{height:5,borderRadius:3,background:"var(--sx)",overflow:"hidden"}}>
              <div style={{width:pct+"%",height:"100%",background:cor,borderRadius:3}}/>
            </div>
          </div>);
        })}
      </div>

      {/* Chamados por setor */}
      <div className="card" style={{padding:18}}>
        <div style={{fontWeight:900,fontSize:13,marginBottom:12,display:"flex",justifyContent:"space-between"}}>
          <span>📨 Chamados por Setor</span>
          <span style={{fontSize:11,color:chamAbertos>0?"#f87171":"#34d399",fontWeight:800}}>{chamAbertos} abertos</span>
        </div>
        {Object.values(chamPorSetor).sort((a,b)=>b.abertos-a.abertos).map(s=>{
          const pctEnc=s.total>0?Math.round((s.enc/s.total)*100):0;
          return(<div key={s.setor} style={{marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}>
              <span style={{fontWeight:700}}>{s.setor}</span>
              <div style={{display:"flex",gap:8}}>
                <span style={{color:"#f87171",fontWeight:800}}>{s.abertos} ab.</span>
                <span style={{color:"#34d399",fontWeight:800}}>{s.enc} enc.</span>
              </div>
            </div>
            <div style={{height:5,borderRadius:3,background:"var(--sx)",overflow:"hidden"}}>
              <div style={{width:pctEnc+"%",height:"100%",background:"#34d399",borderRadius:3}}/>
            </div>
          </div>);
        })}
        {Object.keys(chamPorSetor).length===0&&<div className="muted" style={{fontSize:12}}>Nenhum chamado.</div>}
      </div>
    </div>}

    {/* ══ PRODUTIVIDADE ══ */}
    {tab==="produtividade"&&<div>
      <div className="card" style={{overflow:"hidden"}}>
        <div className="grid-header" style={{gridTemplateColumns:"2fr 60px 60px 60px 60px 80px 90px 70px"}}>
          <div>Profissional</div>
          <div style={{textAlign:"center"}}>Total</div>
          <div style={{textAlign:"center"}}>Atend.</div>
          <div style={{textAlign:"center"}}>Faltas</div>
          <div style={{textAlign:"center"}}>Fat.</div>
          <div style={{textAlign:"center"}}>Presença</div>
          <div style={{textAlign:"right"}}>Receita</div>
          <div style={{textAlign:"center"}}>Nota</div>
        </div>
        {profStats.length===0&&<div style={{padding:16}} className="muted">Nenhum dado no período.</div>}
        {profStats.map((prof,i)=>{
          const cor=espCor((prof.especialidades||[""])[0]);
          const notaCor={"A":"#34d399","B":"#a78bfa","C":"#f59e0b","D":"#f87171"}[prof.perf]||"#64748b";
          return(<div key={prof.id} className="grid-row" style={{gridTemplateColumns:"2fr 60px 60px 60px 60px 80px 90px 70px",background:i%2?"var(--gr)":""}}>
            <div>
              <div style={{fontWeight:800,fontSize:12,color:cor}}>{prof.nome}</div>
              <div style={{fontSize:10,color:"var(--mt)"}}>{(prof.especialidades||[""])[0]}</div>
            </div>
            <div style={{textAlign:"center",fontWeight:800,fontSize:13}}>{prof.tot}</div>
            <div style={{textAlign:"center",fontWeight:800,fontSize:13,color:"#34d399"}}>{prof.at}</div>
            <div style={{textAlign:"center",fontWeight:800,fontSize:13,color:prof.fa>3?"#f87171":"var(--mt)"}}>{prof.fa}</div>
            <div style={{textAlign:"center",fontWeight:800,fontSize:13,color:"#22c55e"}}>{prof.ft}</div>
            <div style={{textAlign:"center"}}>
              <div style={{fontWeight:900,fontSize:13,color:prof.tx>=80?"#34d399":prof.tx>=60?"#f59e0b":"#f87171"}}>{prof.tx}%</div>
              <div style={{height:3,borderRadius:2,background:"var(--sx)",marginTop:2,overflow:"hidden"}}>
                <div style={{width:prof.tx+"%",height:"100%",background:prof.tx>=80?"#34d399":prof.tx>=60?"#f59e0b":"#f87171"}}/>
              </div>
            </div>
            <div style={{textAlign:"right",fontWeight:800,fontSize:12,color:"#34d399"}}>{brl(prof.rec)}</div>
            <div style={{textAlign:"center"}}>
              <span style={{display:"inline-flex",alignItems:"center",justifyContent:"center",width:28,height:28,borderRadius:"50%",fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:14,background:notaCor+"20",color:notaCor,border:"2px solid "+notaCor}}>{prof.perf}</span>
            </div>
          </div>);
        })}
      </div>
      {/* Legenda nota */}
      <div style={{marginTop:10,display:"flex",gap:12,fontSize:11,color:"var(--mt)"}}>
        {[["A","≥85% presença","#34d399"],["B","70–84%","#a78bfa"],["C","50–69%","#f59e0b"],["D","<50%","#f87171"]].map(([n,d,c])=>(
          <span key={n} style={{display:"flex",alignItems:"center",gap:5}}>
            <span style={{display:"inline-flex",alignItems:"center",justifyContent:"center",width:18,height:18,borderRadius:"50%",fontWeight:900,fontSize:10,background:c+"20",color:c,border:"1.5px solid "+c}}>{n}</span> {d}
          </span>
        ))}
      </div>
    </div>}

    {/* ══ PROCESSOS ══ */}
    {tab==="processos"&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      {/* Chamados detalhado */}
      <div className="card" style={{padding:18}}>
        <div style={{fontWeight:900,fontSize:13,marginBottom:14}}>📨 Chamados — Detalhamento</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
          {[{l:"Total",v:chamados.length,c:"#a78bfa"},{l:"Abertos",v:chamados.filter(c=>c.status==="aberto").length,c:"#f59e0b"},{l:"Andamento",v:chamados.filter(c=>c.status==="andamento").length,c:"#818cf8"},{l:"Encerrados",v:chamEnc,c:"#34d399"}].map(k=>(
            <div key={k.l} style={{background:"var(--sx)",borderRadius:9,padding:"10px 12px"}}>
              <div style={{fontSize:10,color:"var(--mt)",fontWeight:800,marginBottom:2}}>{k.l}</div>
              <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:22,color:k.c}}>{k.v}</div>
            </div>
          ))}
        </div>
        <div style={{fontWeight:800,fontSize:11,marginBottom:8,color:"var(--lb)",letterSpacing:".5px"}}>POR SETOR</div>
        {Object.values(chamPorSetor).sort((a,b)=>b.total-a.total).map(s=>(
          <div key={s.setor} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:"1px solid var(--sx)"}}>
            <span style={{fontSize:12,fontWeight:700}}>{s.setor}</span>
            <div style={{display:"flex",gap:10}}>
              <span style={{fontSize:11,color:"#f87171",fontWeight:800}}>{s.abertos} ab.</span>
              <span style={{fontSize:11,color:"#34d399",fontWeight:800}}>{s.enc} enc.</span>
              <span style={{fontSize:10,color:"var(--mt)"}}>{s.total} tot.</span>
            </div>
          </div>
        ))}
        {Object.keys(chamPorSetor).length===0&&<div className="muted" style={{fontSize:12}}>Nenhum chamado.</div>}
      </div>

      {/* Pedidos médicos */}
      <div className="card" style={{padding:18}}>
        <div style={{fontWeight:900,fontSize:13,marginBottom:14}}>🩻 Pedidos Médicos</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
          {[{l:"Total",v:pedidos.length,c:"#a78bfa"},{l:"Venc. 30d",v:pedVenc30.length,c:"#f59e0b"},{l:"Urgente 7d",v:pedVenc7.length,c:"#f87171"},{l:"Vencidos",v:pedVencidos.length,c:"#64748b"}].map(k=>(
            <div key={k.l} style={{background:"var(--sx)",borderRadius:9,padding:"10px 12px"}}>
              <div style={{fontSize:10,color:"var(--mt)",fontWeight:800,marginBottom:2}}>{k.l}</div>
              <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:22,color:k.c}}>{k.v}</div>
            </div>
          ))}
        </div>
        <div style={{fontWeight:800,fontSize:11,marginBottom:8,color:"var(--lb)",letterSpacing:".5px"}}>VENCENDO EM 7 DIAS</div>
        {pedVenc7.length===0&&<div style={{fontSize:12,color:"#34d399"}}>✅ Nenhum vencendo esta semana</div>}
        {pedVenc7.map(p=>{
          const pac=pacientes.find(x=>x.id===Number(p.pacienteId));
          const dias=diffDays(hoje_str,p.dataValidade);
          return(<div key={p.id} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid var(--sx)",fontSize:12}}>
            <div>
              <div style={{fontWeight:700}}>{pac?.nome?.split(" ")[0]||"—"} {pac?.nome?.split(" ").slice(-1)[0]||""}</div>
              <div style={{fontSize:10,color:"var(--mt)"}}>{p.especialidade||p.tipo}</div>
            </div>
            <span style={{fontWeight:900,color:dias<=3?"#f87171":"#f59e0b",fontSize:14}}>{dias}d</span>
          </div>);
        })}
      </div>

      {/* Distribuição de especialidades */}
      <div className="card" style={{padding:18,gridColumn:"1/-1"}}>
        <div style={{fontWeight:900,fontSize:13,marginBottom:14}}>🩺 Sessões por Especialidade — período</div>
        {(()=>{
          const espStats=agP.reduce((acc,a)=>{
            const prof=profissionais.find(p=>p.id===Number(a.profissionalId));
            const esp=(prof?.especialidades||[prof?.especialidade||"Outro"])[0]||"Outro";
            if(!acc[esp])acc[esp]={esp,total:0,atend:0};
            acc[esp].total++;
            if(["atendido","faturado"].includes(a.status))acc[esp].atend++;
            return acc;
          },{});
          const arr=Object.values(espStats).sort((a,b)=>b.total-a.total);
          const mx=Math.max(...arr.map(e=>e.total),1);
          return(<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:10}}>
            {arr.map(e=>{
              const cor=espCor(e.esp);
              const pct=Math.round((e.total/mx)*100);
              const txP=e.total>0?Math.round((e.atend/e.total)*100):0;
              return(<div key={e.esp} style={{background:"var(--sx)",borderRadius:10,padding:"12px 14px",borderLeft:"3px solid "+cor}}>
                <div style={{fontSize:11,fontWeight:800,color:cor,marginBottom:6,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{e.esp}</div>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",marginBottom:5}}>
                  <span style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:22,color:"var(--tx)"}}>{e.total}</span>
                  <span style={{fontSize:11,fontWeight:800,color:txP>=80?"#34d399":txP>=60?"#f59e0b":"#f87171"}}>{txP}% pres.</span>
                </div>
                <div style={{height:4,borderRadius:2,background:"var(--cb)",overflow:"hidden"}}>
                  <div style={{width:pct+"%",height:"100%",background:cor,borderRadius:2}}/>
                </div>
              </div>);
            })}
          </div>);
        })()}
      </div>
    </div>}

    {/* ══ INDICADORES ══ */}
    {tab==="indicadores"&&<div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:12,marginBottom:18}}>
        {indicadores.map(ind=>{
          const ok=ind.rev?(ind.val<=ind.meta):(ind.val>=ind.meta);
          const pctBar=ind.rev?Math.min(100,Math.round((ind.meta/Math.max(ind.val,1))*100)):Math.min(100,ind.val);
          return(<div key={ind.label} style={{background:"var(--card)",border:"1.5px solid "+(ok?ind.cor+"40":"#f8717140"),borderRadius:13,padding:"16px 18px",position:"relative",overflow:"hidden"}}>
            <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:ok?ind.cor:"#f87171",borderRadius:"13px 13px 0 0"}}/>
            <div style={{fontSize:20,marginBottom:8}}>{ind.icon}</div>
            <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",letterSpacing:".5px",textTransform:"uppercase",marginBottom:6}}>{ind.label}</div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",marginBottom:8}}>
              <span style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:30,color:ok?ind.cor:"#f87171",lineHeight:1,letterSpacing:"-1.5px"}}>{ind.val}{ind.unit}</span>
              <div style={{textAlign:"right"}}>
                <div style={{fontSize:10,color:"var(--mt)"}}>Meta</div>
                <div style={{fontSize:14,fontWeight:900,color:"var(--mt)"}}>{ind.meta}{ind.unit}</div>
              </div>
            </div>
            <div style={{height:6,borderRadius:3,background:"var(--sx)",overflow:"hidden",marginBottom:8}}>
              <div style={{width:pctBar+"%",height:"100%",background:ok?ind.cor:"#f87171",borderRadius:3,transition:"width .5s"}}/>
            </div>
            <div style={{fontSize:11,fontWeight:800,color:ok?"#34d399":"#f87171",display:"flex",alignItems:"center",gap:5}}>
              {ok?"✅ Meta atingida":"⚠️ Abaixo da meta"}
              <span style={{marginLeft:"auto",fontSize:10,color:"var(--mt)"}}>Dif: {ind.rev?ind.val-ind.meta:ind.val-ind.meta > 0?"+":""}{ ind.val-ind.meta}{ind.unit}</span>
            </div>
          </div>);
        })}
      </div>

      {/* Histórico mensal simplificado */}
      <div className="card" style={{padding:18}}>
        <div style={{fontWeight:900,fontSize:13,marginBottom:14}}>📅 Distribuição semanal — últimas 4 semanas</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10}}>
          {Array.from({length:4},(_,w)=>{
            const fim4=new Date();fim4.setDate(fim4.getDate()-(3-w)*7);
            const ini4=new Date(fim4);ini4.setDate(fim4.getDate()-6);
            const inS=ymd(ini4),fiS=ymd(fim4);
            const ags=agenda.filter(a=>a.data>=inS&&a.data<=fiS);
            const at=ags.filter(a=>["atendido","faturado"].includes(a.status)).length;
            const fa=ags.filter(a=>["faltou","faltou_pacote"].includes(a.status)).length;
            const tx=ags.length>0?Math.round((at/ags.length)*100):0;
            const label=w===3?"Esta sem.":"Sem. -"+(3-w);
            return(<div key={w} style={{background:"var(--sx)",borderRadius:10,padding:"14px",textAlign:"center"}}>
              <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",marginBottom:8}}>{label}</div>
              <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:22,color:tx>=80?"#34d399":tx>=60?"#f59e0b":"#f87171"}}>{tx}%</div>
              <div style={{fontSize:10,color:"var(--mt)",marginTop:2}}>presença</div>
              <div style={{marginTop:8,display:"flex",justifyContent:"space-around",fontSize:10}}>
                <span style={{color:"var(--tx)",fontWeight:700}}>{ags.length} sess.</span>
                <span style={{color:"#f87171",fontWeight:700}}>{fa} falta</span>
              </div>
            </div>);
          })}
        </div>
      </div>
    </div>}

    {/* ══ OCUPAÇÃO SALAS ══ */}
    {tab==="ocupacao"&&<div>
      {/* ── KPIs financeiros de salas ── */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(170px,1fr))",gap:8,marginBottom:14}}>
        <div className="card" style={{padding:"12px 14px",borderTop:"3px solid #a78bfa"}}>
          <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:22,color:"#a78bfa"}}>{ocupMedia}%</div>
          <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",marginTop:3}}>Ocupação média</div>
          <div style={{fontSize:10,color:"var(--mt)"}}>{salaStats.length} sala(s) ativas</div>
        </div>
        <div className="card" style={{padding:"12px 14px",borderTop:"3px solid #ef4444"}}>
          <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:22,color:"#f87171"}}>{brl(custoSalasTotal)}</div>
          <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",marginTop:3}}>Custo mensal salas</div>
          <div style={{fontSize:10,color:"var(--mt)"}}>soma de todas as salas</div>
        </div>
        <div className="card" style={{padding:"12px 14px",borderTop:"3px solid #34d399"}}>
          <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:22,color:"#34d399"}}>{brl(receitaSalasTotal)}</div>
          <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",marginTop:3}}>Receita nas salas</div>
          <div style={{fontSize:10,color:"var(--mt)"}}>sessões faturadas no período</div>
        </div>
        {roiSalas!==null&&<div className="card" style={{padding:"12px 14px",borderTop:"3px solid "+(roiSalas>=0?"#22c55e":"#ef4444")}}>
          <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:22,color:roiSalas>=0?"#34d399":"#f87171"}}>{roiSalas>=0?"+":""}{roiSalas}%</div>
          <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",marginTop:3}}>ROI salas</div>
          <div style={{fontSize:10,color:"var(--mt)"}}>(receita − custo) / custo</div>
        </div>}
      </div>

      {/* ── Cards por sala ── */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12,marginBottom:16}}>
        {salaStats.map(sala=>{
          const cor=sala.pct>=80?"#f87171":sala.pct>=50?"#f59e0b":"#34d399";
          const corBarra=sala.cor||cor;
          // pico de horário
          const pico=sala.porHora?.reduce((a,b)=>b.qtd>a.qtd?b:a,{hora:"—",qtd:0,pct:0});
          return(<div key={sala.id} className="card" style={{padding:16,borderTop:"4px solid "+corBarra}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
              <div>
                <div style={{fontWeight:800,fontSize:13}}>{sala.nome}</div>
                <div style={{fontSize:10,color:"var(--mt)"}}>{sala.filial}</div>
                {sala.especialidade&&<div style={{fontSize:10,color:espCor(sala.especialidade),fontWeight:700,marginTop:1}}>{sala.especialidade}</div>}
              </div>
              <span style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:26,color:corBarra,lineHeight:1}}>{sala.pct}%</span>
            </div>
            {/* Barra principal */}
            <div style={{height:7,borderRadius:4,background:"var(--sx)",overflow:"hidden",marginBottom:8}}>
              <div style={{width:sala.pct+"%",height:"100%",background:corBarra,borderRadius:4,transition:"width .5s"}}/>
            </div>
            {/* Info grid */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:4,marginBottom:8,fontSize:11}}>
              <div><span style={{color:"var(--mt)"}}>Sessões: </span><b>{sala.total}</b></div>
              <div><span style={{color:"var(--mt)"}}>Cap.: </span><b>{sala.cap}</b></div>
              {sala.custo>0&&<div><span style={{color:"var(--mt)"}}>Custo: </span><b style={{color:"#f87171"}}>{brl(sala.custo)}</b></div>}
              {sala.custo>0&&<div><span style={{color:"var(--mt)"}}>Receita: </span><b style={{color:"#34d399"}}>{brl(sala.receita)}</b></div>}
              {sala.roi!==null&&<div style={{gridColumn:"1/3"}}><span style={{color:"var(--mt)"}}>ROI: </span><b style={{color:sala.roi>=0?"#34d399":"#f87171"}}>{sala.roi>=0?"+":""}{sala.roi}%</b></div>}
              {pico?.qtd>0&&<div style={{gridColumn:"1/3"}}><span style={{color:"var(--mt)"}}>🔥 Pico: </span><b>{pico.hora}</b><span style={{color:"var(--mt)"}}> ({pico.qtd} ag./dia em média)</span></div>}
            </div>
            {/* Mini heat-map de horários */}
            <div style={{marginTop:6}}>
              <div style={{fontSize:9,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",marginBottom:4}}>Ocupação por horário</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:2}}>
                {(sala.porHora||[]).filter((_,i)=>i%2===0).map(h=>{
                  const bg=h.qtd===0?"var(--sx)":h.pct>=80?"#ef4444":h.pct>=50?"#f59e0b":h.pct>=20?"#34d399":"#14b8a620";
                  return(<div key={h.hora} title={h.hora+": "+h.qtd+" ag. ("+h.pct+"%)"} style={{width:18,height:18,borderRadius:3,background:bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:7,fontWeight:800,color:h.qtd>0?"#fff":"var(--lb)",cursor:"default"}}>
                    {h.hora.slice(0,2)}
                  </div>);
                })}
              </div>
              <div style={{display:"flex",gap:8,marginTop:5,fontSize:9,color:"var(--mt)"}}>
                <span style={{display:"flex",alignItems:"center",gap:2}}><span style={{width:8,height:8,borderRadius:2,background:"#14b8a620",display:"inline-block"}}/>Livre</span>
                <span style={{display:"flex",alignItems:"center",gap:2}}><span style={{width:8,height:8,borderRadius:2,background:"#34d399",display:"inline-block"}}/>Baixo</span>
                <span style={{display:"flex",alignItems:"center",gap:2}}><span style={{width:8,height:8,borderRadius:2,background:"#f59e0b",display:"inline-block"}}/>Médio</span>
                <span style={{display:"flex",alignItems:"center",gap:2}}><span style={{width:8,height:8,borderRadius:2,background:"#ef4444",display:"inline-block"}}/>Alto</span>
              </div>
            </div>
            <div style={{fontSize:10,padding:"4px 8px",borderRadius:6,background:cor+"15",color:cor,fontWeight:800,textAlign:"center",marginTop:8}}>
              {sala.pct>=80?"🔴 Alta ocupação":sala.pct>=50?"🟡 Ocupação moderada":"🟢 Disponível"}
            </div>
          </div>);
        })}
        {salaStats.length===0&&<div className="muted" style={{fontSize:12,padding:16}}>Nenhuma sala cadastrada.</div>}
      </div>

      {/* ── Por filial ── */}
      {filiais.length>0&&<div className="card" style={{padding:18}}>
        <div style={{fontWeight:900,fontSize:13,marginBottom:12}}>🏢 Ocupação por Filial</div>
        {filiais.map(fil=>{
          const ss=salaStats.filter(s=>s.filialId===fil.id);
          if(ss.length===0)return null;
          const med=Math.round(ss.reduce((s,x)=>s+x.pct,0)/ss.length);
          const custoFil=ss.reduce((s,x)=>s+(x.custo||0),0);
          const recFil=ss.reduce((s,x)=>s+(x.receita||0),0);
          const cor=med>=80?"#f87171":med>=50?"#f59e0b":"#34d399";
          return(<div key={fil.id} style={{marginBottom:14,paddingBottom:14,borderBottom:"1px solid var(--sc)"}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:5,flexWrap:"wrap",gap:4}}>
              <span style={{fontWeight:800}}>{fil.nome}</span>
              <div style={{display:"flex",gap:12,fontSize:11}}>
                {custoFil>0&&<span style={{color:"#f87171"}}>Custo: <b>{brl(custoFil)}</b></span>}
                {recFil>0&&<span style={{color:"#34d399"}}>Receita: <b>{brl(recFil)}</b></span>}
                <span style={{fontWeight:900,color:cor}}>{med}% · {ss.length} salas</span>
              </div>
            </div>
            <div style={{height:7,borderRadius:4,background:"var(--sx)",overflow:"hidden"}}>
              <div style={{width:med+"%",height:"100%",background:cor,borderRadius:4,transition:"width .5s"}}/>
            </div>
          </div>);
        })}
      </div>}
    </div>}

    {/* ══ CONVÊNIOS ══ */}
    {tab==="convenios"&&<div>
      <div className="card" style={{overflow:"hidden",marginBottom:14}}>
        <div className="grid-header" style={{gridTemplateColumns:"1.6fr 70px 70px 80px 70px 100px"}}>
          <div>Convênio</div>
          <div style={{textAlign:"center"}}>Sessões</div>
          <div style={{textAlign:"center"}}>Atend.</div>
          <div style={{textAlign:"center"}}>Presença</div>
          <div style={{textAlign:"center"}}>Fat.</div>
          <div style={{textAlign:"right"}}>Receita Est.</div>
        </div>
        {convArr.length===0&&<div style={{padding:16}} className="muted">Nenhum dado no período.</div>}
        {convArr.map((cv,i)=>{
          const tx=cv.total>0?Math.round((cv.atend/cv.total)*100):0;
          return(<div key={cv.nome} className="grid-row" style={{gridTemplateColumns:"1.6fr 70px 70px 80px 70px 100px",background:i%2?"var(--gr)":""}}>
            <div>
              <div style={{fontWeight:800,fontSize:12}}>{cv.nome}</div>
              <div style={{height:3,borderRadius:2,background:"var(--sx)",marginTop:4,overflow:"hidden",maxWidth:100}}>
                <div style={{width:Math.round((cv.total/maxConv)*100)+"%",height:"100%",background:"#a78bfa"}}/>
              </div>
            </div>
            <div style={{textAlign:"center",fontWeight:800}}>{cv.total}</div>
            <div style={{textAlign:"center",fontWeight:800,color:"#34d399"}}>{cv.atend}</div>
            <div style={{textAlign:"center"}}>
              <span style={{fontWeight:900,color:tx>=80?"#34d399":tx>=60?"#f59e0b":"#f87171"}}>{tx}%</span>
            </div>
            <div style={{textAlign:"center",fontWeight:800,color:"#22c55e"}}>{cv.fat}</div>
            <div style={{textAlign:"right",fontWeight:800,color:"#34d399",fontSize:12}}>{brl(cv.rec)}</div>
          </div>);
        })}
      </div>
      {/* Barras comparativas */}
      <div className="card" style={{padding:18}}>
        <div style={{fontWeight:900,fontSize:13,marginBottom:14}}>📊 Volume por Convênio</div>
        {convArr.map(cv=>{
          const pct=Math.round((cv.total/maxConv)*100);
          const tx=cv.total>0?Math.round((cv.atend/cv.total)*100):0;
          return(<div key={cv.nome} style={{marginBottom:12}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:4}}>
              <span style={{fontWeight:800}}>{cv.nome}</span>
              <div style={{display:"flex",gap:12,fontSize:11}}>
                <span style={{color:"var(--mt)"}}>{cv.total} sessões</span>
                <span style={{fontWeight:800,color:tx>=80?"#34d399":tx>=60?"#f59e0b":"#f87171"}}>{tx}% presença</span>
                <span style={{fontWeight:800,color:"#34d399"}}>{brl(cv.rec)}</span>
              </div>
            </div>
            <div style={{height:8,borderRadius:4,background:"var(--sx)",overflow:"hidden"}}>
              <div style={{width:pct+"%",height:"100%",background:"#a78bfa",borderRadius:4,transition:"width .5s"}}/>
            </div>
          </div>);
        })}
        {convArr.length===0&&<div className="muted" style={{fontSize:12}}>Nenhum dado.</div>}
      </div>
    </div>}

    {/* ══ DASHBOARD ══ */}
    {tab==="dashboard"&&(()=>{
      const agHoje2=agenda.filter(a=>a.data===hoje_str);
      const ocupSalas=new Set(agHoje2.map(a=>a.salaId));
      const totalSalas2=salas.filter(s=>s.ativa!==false).length||1;
      const pctOcup=Math.round((ocupSalas.size/totalSalas2)*100);
      const pedVenc30=pedidos.filter(p=>p.dataValidade&&diffDays(hoje_str,p.dataValidade)>=0&&diffDays(hoje_str,p.dataValidade)<=30);
      const chamAbDash=chamados.filter(c=>["aberto","andamento"].includes(c.status)).length;
      return(<div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:12,marginBottom:18}}>
          {[
            {icon:"📅",label:"Agendamentos Hoje",val:agHoje2.length,color:"#a78bfa"},
            {icon:"✅",label:"Atendidos Hoje",val:agHoje2.filter(a=>["atendido","faturado"].includes(a.status)).length,color:"#34d399"},
            {icon:"🏢",label:"Ocupação Salas",val:pctOcup+"%",color:"#f59e0b"},
            {icon:"⏳",label:"Fila de Espera",val:fila.length,color:"#a78bfa"},
            {icon:"📨",label:"Chamados Abertos",val:chamAbDash,color:chamAbDash>5?"#f87171":"#fb923c"},
            {icon:"🩻",label:"Pedidos Vencendo",val:pedVenc30.length,color:"#fb923c"},
            {icon:"👤",label:"Pacientes",val:pacientes.length,color:"#38bdf8"},
            {icon:"🩺",label:"Profissionais",val:profissionais.filter(p=>["profissional","coordenador","coordenador_aba"].includes(p.role)).length,color:"#a3e635"},
          ].map(k=>(
            <div key={k.label} className="card" style={{padding:"14px 16px",borderTop:"3px solid "+k.color}}>
              <div style={{fontSize:20,marginBottom:6}}>{k.icon}</div>
              <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:24,color:k.color,lineHeight:1,letterSpacing:"-1px"}}>{k.val}</div>
              <div style={{fontSize:10,fontWeight:800,letterSpacing:".6px",color:"var(--mt)",textTransform:"uppercase",marginTop:3}}>{k.label}</div>
            </div>
          ))}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          <div className="card" style={{padding:18}}>
            <div style={{fontWeight:900,fontSize:13,marginBottom:12}}>🏢 Ocupação das Salas — Hoje</div>
            {salas.filter(s=>s.ativa!==false).slice(0,8).map(s=>{
              const agS=agHoje2.filter(a=>a.salaId===s.id);
              const pct=Math.min(100,Math.round((agS.length/8)*100));
              return(<div key={s.id} style={{marginBottom:8}}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}>
                  <span style={{color:"var(--mt)",fontWeight:700}}>{s.nome}</span>
                  <span style={{fontWeight:900,color:s.cor||"#a78bfa"}}>{agS.length} sess.</span>
                </div>
                <div style={{height:6,borderRadius:3,background:"var(--sx)",overflow:"hidden"}}>
                  <div style={{width:pct+"%",height:"100%",background:s.cor||"#a78bfa",borderRadius:3}}/>
                </div>
              </div>);
            })}
            {salas.filter(s=>s.ativa!==false).length===0&&<div className="muted" style={{fontSize:12}}>Nenhuma sala ativa.</div>}
          </div>
          <div className="card" style={{padding:18}}>
            <div style={{fontWeight:900,fontSize:13,marginBottom:12}}>📨 Chamados por Status</div>
            {STATUS_CHAMADO.map(st=>{
              const cnt=chamados.filter(c=>c.status===st).length;
              return(<div key={st} style={{marginBottom:8}}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}>
                  <span style={{color:"var(--mt)",fontWeight:700}}>{LABEL_CHAMADO[st]}</span>
                  <span style={{fontWeight:900,color:COR_CHAMADO[st]}}>{cnt}</span>
                </div>
                <div style={{height:6,borderRadius:3,background:"var(--sx)",overflow:"hidden"}}>
                  <div style={{width:Math.round((cnt/(chamados.length||1))*100)+"%",height:"100%",background:COR_CHAMADO[st],borderRadius:3}}/>
                </div>
              </div>);
            })}
          </div>
          <div className="card" style={{padding:18}}>
            <div style={{fontWeight:900,fontSize:13,marginBottom:12}}>🩻 Pedidos Médicos — Alertas</div>
            {pedVenc30.length===0&&<div style={{fontSize:12,color:"#34d399"}}>Nenhum vencendo nos próximos 30 dias</div>}
            {pedVenc30.map(p=>{
              const pac=pacientes.find(x=>x.id===Number(p.pacienteId));
              const dias=diffDays(hoje_str,p.dataValidade);
              return(<div key={p.id} style={{background:dias<=7?"#450a0a15":"#451a0315",border:"1px solid "+(dias<=7?"#f8717130":"#f59e0b30"),borderRadius:7,padding:"7px 9px",marginBottom:5,fontSize:11}}>
                <b style={{color:dias<=7?"#f87171":"#f59e0b"}}>{dias<=7?"🚨":"⚠️"} {dias}d — {pac?.nome||"—"}</b>
                <div style={{color:"var(--mt)",fontSize:10}}>{p.especialidade||p.tipo}</div>
              </div>);
            })}
          </div>
          <div className="card" style={{padding:18}}>
            <div style={{fontWeight:900,fontSize:13,marginBottom:12}}>⏳ Fila de Espera por Especialidade</div>
            {fila.length===0&&<div className="muted" style={{fontSize:12}}>Fila vazia.</div>}
            {[...new Set(fila.map(f=>f.especialidade))].map(esp=>{
              const cnt=fila.filter(f=>f.especialidade===esp).length;
              const pct=Math.round((cnt/(fila.length||1))*100);
              const cor=espCor(esp);
              return(<div key={esp} style={{marginBottom:8}}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}>
                  <span style={{color:"var(--mt)",fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:160}}>{esp}</span>
                  <span style={{fontWeight:900,color:cor}}>{cnt}</span>
                </div>
                <div style={{height:5,borderRadius:3,background:"var(--sx)",overflow:"hidden"}}>
                  <div style={{width:pct+"%",height:"100%",background:cor,borderRadius:3}}/>
                </div>
              </div>);
            })}
          </div>
        </div>
      </div>);
    })()}

    {/* ══ RELATÓRIOS ══ */}
    {tab==="relatorios"&&<RelatoriosTab agenda={agenda} pacientes={pacientes} profissionais={profissionais} procedimentos={procedimentos} showToast={showToast}/>}

    {/* ══ ATIVIDADES DO SISTEMA ══ */}
    {tab==="atividades"&&<AtividadesTab atividades={atividades} profissionais={profissionais}/>}

    {/* ══ IMPORTAR DADOS ══ */}
    {tab==="importar"&&<ImportarTab agenda={agenda} setAgenda={setAgenda} pacientes={pacientes} setPacientes={setPacientes} profissionais={profissionais} setProfissionais={setProfissionais} auth={auth} showToast={showToast}/>}

    {/* ══ EXPORTAR ══ */}
    {tab==="exportar"&&<ExportarTab agenda={agenda} pacientes={pacientes} profissionais={profissionais} procedimentos={procedimentos} salas={salas} filiais={filiais} chamados={chamados} showToast={showToast}/>}

  </div>);
}

// ═══════════════════════════════════════════════════════════════════════════════
// APP ROOT
// ═══════════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════════
// ATIVIDADES (Log de sistema)
// ═══════════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════════
// RELATÓRIOS GERENCIAIS
// ═══════════════════════════════════════════════════════════════════════════════
function MiniBar({label,value,max,color}){
  const pct=max>0?Math.round((value/max)*100):0;
  return(<div style={{marginBottom:6}}>
    <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:2}}><span style={{color:"var(--mt)"}}>{label}</span><b style={{color}}>{value}</b></div>
    <div style={{height:8,borderRadius:4,background:"var(--sx)",overflow:"hidden"}}>
      <div style={{width:pct+"%",height:"100%",background:color,borderRadius:4,transition:"width .4s"}}/>
    </div>
  </div>);
}
function DonutChart({data,size=100}){
  const total=data.reduce((s,d)=>s+d.v,0)||1;
  let cum=0;
  const r=40,cx=50,cy=50,circ=2*Math.PI*r;
  return(<svg width={size} height={size} viewBox="0 0 100 100">
    {data.map((d,i)=>{
      const pct=d.v/total;
      const off=circ*(1-pct);
      const rot=-90+(cum/total)*360;
      cum+=d.v;
      return(<circle key={i} r={r} cx={cx} cy={cy} fill="none" stroke={d.c} strokeWidth={14} strokeDasharray={circ} strokeDashoffset={off} transform={"rotate("+rot+" 50 50)"} style={{transition:"stroke-dashoffset .4s"}}/>);
    })}
    <text x="50" y="55" textAnchor="middle" style={{fontSize:14,fontWeight:900,fill:"var(--tx)"}}>{total}</text>
  </svg>);
}
// ═══════════════════════════════════════════════════════════════════════════════
// RELATÓRIOS + DASHBOARD — página unificada com abas + exportação Excel
// ═══════════════════════════════════════════════════════════════════════════════

// Helper: gerar CSV e forçar download

export default GestaoPage;
