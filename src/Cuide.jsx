import { useMemo, useState, useEffect, lazy, Suspense } from "react";
import * as XLSX from "xlsx";
import { ErrorBoundary } from "./components/ErrorBoundary";

// ═══════════════════════════════════════════════════════════════════════════════
// TEMA
// ═══════════════════════════════════════════════════════════════════════════════
const DARK={bg:"#0f0f14",sidebar:"#16161f",card:"#16161f",cardBorder:"#2a2a3d",inp:"#1e1e2a",inpBorder:"#2a2a3d",text:"#e8e8f0",muted:"#606080",label:"#606080",secBox:"#1e1e2a",secBorder:"#2a2a3d",filterBar:"#16161f",filterBorder:"#2a2a3d",calHead:"#16161f",day:"#1e1e2a",dayBorder:"#2a2a3d",dayHover:"#252535",navBtn:"#9090b0",navHover:"#1e1e2a",navActive:"#7c6af71f",gridRow:"#1e1e2a",modal:"#16161f",modalBg:"rgba(0,0,0,.85)",upload:"#2a2a3d",scrollTrack:"#16161f",scrollThumb:"#2a2a3d",chipBorder:"#2a2a3d",sideBorder:"#2a2a3d"};
const LIGHT={bg:"#f2f2f7",sidebar:"#ffffff",card:"#ffffff",cardBorder:"#d1d1d6",inp:"#f4f4f8",inpBorder:"#d1d1d6",text:"#1c1c1e",muted:"#8e8e93",label:"#8e8e93",secBox:"#f4f4f8",secBorder:"#d1d1d6",filterBar:"#ffffff",filterBorder:"#d1d1d6",calHead:"#ffffff",day:"#f4f4f8",dayBorder:"#d1d1d6",dayHover:"#e5e5ea",navBtn:"#3a3a3c",navHover:"#f4f4f8",navActive:"#7c6af71a",gridRow:"#f4f4f8",modal:"#ffffff",modalBg:"rgba(15,23,42,.5)",upload:"#d1d1d6",scrollTrack:"#f2f2f7",scrollThumb:"#d1d1d6",chipBorder:"#d1d1d6",sideBorder:"#d1d1d6"};
// "Médio" — mesmos valores de html.mid-mode do style.css do MiContas
// (--bg:#363640 --surface:#40404c --surface2:#4a4a58 --surface3:#565664
// --border:#5c5c6c --accent:#a394fb --text:#f5f5fa --text2:#c7c7d6 --text3:#9797a8),
// seguindo o mesmo mapeamento de campos usado em DARK/LIGHT acima.
const MID={bg:"#363640",sidebar:"#40404c",card:"#40404c",cardBorder:"#5c5c6c",inp:"#4a4a58",inpBorder:"#5c5c6c",text:"#f5f5fa",muted:"#9797a8",label:"#9797a8",secBox:"#4a4a58",secBorder:"#5c5c6c",filterBar:"#40404c",filterBorder:"#5c5c6c",calHead:"#40404c",day:"#4a4a58",dayBorder:"#5c5c6c",dayHover:"#565664",navBtn:"#c7c7d6",navHover:"#4a4a58",navActive:"#a394fb1f",gridRow:"#4a4a58",modal:"#40404c",modalBg:"rgba(0,0,0,.7)",upload:"#5c5c6c",scrollTrack:"#40404c",scrollThumb:"#5c5c6c",chipBorder:"#5c5c6c",sideBorder:"#5c5c6c"};
const THEMES_CUIDE={dark:DARK,mid:MID,light:LIGHT};

// ═══════════════════════════════════════════════════════════════════════════════
// CONSTANTES — todas definidas ANTES de qualquer função ou seed
// ═══════════════════════════════════════════════════════════════════════════════
const ESP_CORES={"Nutrição":"#FF8C00","Terapia Ocupacional":"#DAA520","Supervisão ABA":"#C2185B","Fisioterapia":"#6A0572","Natação Terapêutica":"#008080","Psicopedagogia":"#A0522D","Psicomotricidade":"#556B2F","Musicoterapia":"#BDB76B","Arteterapia":"#FFB6C1","Aplicador ABA (A.T.)":"#FF00FF","Terapia Assistida por Animais":"#9932CC","Psicologia Convencional/Psicanálise":"#8B4513","Fonoaudiologia":"#9370DB","Avaliação Neuropsicológica":"#FFFACD","Avaliação Neuropsicopedagógica":"#FFFACD","Coordenação ABA":"#3b82f6","Neurofeedback":"#06b6d4","Psicoterapia":"#7c3aed","Snoezelen":"#84cc16","Agendamento":"#22c55e","Faturamento":"#f59e0b","Atendimento":"#60a5fa","Supervisão ADM":"#e879f9","Administrador":"#f87171","Financeiro":"#34d399","Gestão de Pessoas":"#a78bfa","Outro":"#94a3b8"};
const ESPECIALIDADES_LIST=Object.keys(ESP_CORES);
const espCor=e=>ESP_CORES[e]||"#94a3b8";
const MONTHS_PT=["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
const DIAS_SEMANA=["Domingo","Segunda","Terça","Quarta","Quinta","Sexta","Sábado"];
const TEMPOS_SESSAO=[30,40,45,50,60];
const TURNOS_H=["07:00","07:30","08:00","08:30","09:00","09:30","10:00","10:30","11:00","11:30","12:00","12:30","13:00","13:30","14:00","14:30","15:00","15:30","16:00","16:30","17:00","17:30","18:00","18:30","19:00","19:30","20:00"];
const CONVENIOS_LIST=["Particular","ABM Saúde","Alice (ABA)","Cassi","Hospital Cruz Azul (ABA)","Porto Seguro (ABA)","Saúde Caixa","SulAmerica (ABA)","UniVida","Saúde Abas","Vivest","Unimed CNU","Metrus","Blue Saúde","Bradesco Saúde (ABA)","MediService","Cruz Azul Saúde (ABA)","OAB / ACMC / CAASP","Outro"];
const PADRAO_CHAMADO=[
  {id:"01",label:"01 — Novo Paciente",campos:["Nome Completo","Convênio","Coordenador","Especialidade","Profissional","Início","Horário","Local"]},
  {id:"02",label:"02 — Novo Agendamento",campos:["Nome Completo","Convênio","Coordenador","Especialidade","Profissional","Início","Horário","Local"]},
  {id:"03",label:"03 — Troca de Horário",campos:["Nome Completo","Convênio","Coordenador","Especialidade","Horário Anterior","Novo Horário","Início","Local"]},
  {id:"04",label:"04 — Troca de Profissional",campos:["Nome Completo","Convênio","Coordenador","Especialidade","Profissional Anterior","Novo Profissional","Início","Horário","Local"]},
  {id:"05",label:"05 — Aumento de Carga Horária",campos:["Nome Completo","Convênio","Coordenador","Especialidade","Profissional","Início","Horário","Local"]},
  {id:"06",label:"06 — Redução de Carga Horária",campos:["Nome Completo","Convênio","Coordenador","Especialidade","Profissional","Último Atendimento","Horário","Local"]},
  {id:"07",label:"07 — Cancelamento de Atendimento",campos:["Nome Completo","Convênio","Coordenador","Especialidade Cancelada","Profissional","Último Atendimento","Motivo","Local"]},
];
const ESTADOS_CIVIS=["Solteiro(a)","Casado(a)","Divorciado(a)","Viúvo(a)","Separado(a)","União Estável"];
const UFS=["AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO"];
const CONSELHOS_OPT=["CRM","CRO","CRP","CREFITO","CRN","CRBM","CRF","Outro"];
const GRAUS=["Próprio(a)","Mãe","Pai","Irmão(ã)","Avó/Avô","Cônjuge","Tutor(a)","Outros"];
const FORMA_PAGAMENTO=["PIX","TED","Boleto","Dinheiro","Cartão","Outro"];
const NIVEL_REPASSE=["Trainee","Júnior","Pleno","Sênior"];
// ─── MODELOS DE EVOLUÇÃO POR ESPECIALIDADE ───────────────────────────────────
const MODELOS_EVOLUCAO={
  "Psicologia Convencional/Psicanálise":[
    {id:"psi1",titulo:"Mando — Modelo 1",texto:"Paciente apresentou tentativas de mando espontâneo durante a sessão. Foram trabalhadas estratégias de solicitação funcional com suporte visual e verbal. Paciente respondeu positivamente aos prompts, com redução gradual da dependência de dica. Conduta: manter reforçamento diferencial e ampliar repertório de mandos."},
    {id:"psi2",titulo:"Tato — Modelo 5",texto:"Sessão focada em tatos de objetos do cotidiano. Paciente nomeou corretamente X de Y itens apresentados. Observou-se generalização para novos estímulos. Conduta: introduzir tatos de ações e atributos na próxima sessão."},
    {id:"psi3",titulo:"Habilidades Sociais",texto:"Trabalhadas habilidades de esperar a vez, contato visual e saudação. Paciente demonstrou engajamento satisfatório nas atividades em pares. Dificuldades observadas em turnos de conversação. Conduta: reforçar comportamentos pró-sociais e ampliar grupo."},
    {id:"psi4",titulo:"Reestruturação Cognitiva",texto:"Sessão de psicoeducação sobre pensamentos automáticos e distorções cognitivas. Paciente identificou situações-gatilho e crenças disfuncionais associadas. Foram trabalhadas estratégias de questionamento socrático. Conduta: tarefa de registro de pensamentos durante a semana."},
    {id:"psi5",titulo:"Psicoeducação Ansiedade",texto:"Abordadas psicoeducação sobre ciclo de ansiedade e resposta fisiológica ao estresse. Paciente demonstrou boa compreensão do modelo cognitivo-comportamental. Foram introduzidas técnicas de respiração diafragmática. Conduta: prática diária de 5 minutos de respiração guiada."},
  ],
  "Psicomotricidade":[
    {id:"pm1",titulo:"Circuito Motor",texto:"Paciente realizou circuito motor com obstáculos envolvendo equilíbrio, coordenação e planejamento motor. Observou-se dificuldade em sequências motoras complexas, com melhora progressiva após repetição. Tônus adequado para a atividade. Conduta: aumentar complexidade do circuito gradualmente."},
    {id:"pm2",titulo:"Atividades com Bola",texto:"Trabalhadas habilidades de arremesso, recepção e chute com bola de tamanhos variados. Paciente demonstrou coordenação olho-mão em desenvolvimento. Dissociação de cinturas ainda em processo. Conduta: inserir atividades de mira e precisão."},
    {id:"pm3",titulo:"Lateralidade e Esquema Corporal",texto:"Sessão com foco em reconhecimento de lateralidade e partes do corpo. Paciente demonstrou confusão lateral em atividades de espelho. Esquema corporal em construção. Conduta: utilizar recursos visuais de apoio e jogos de identificação."},
    {id:"pm4",titulo:"Equilíbrio Estático e Dinâmico",texto:"Atividades de equilíbrio sobre superfícies instáveis e trilhas de equilíbrio. Paciente apresentou tremor distal ao manter posição unipodal. Equilíbrio dinâmico mais preservado. Conduta: progredir para superfícies com maior instabilidade."},
  ],
  "Psicopedagogia":[
    {id:"pp1",titulo:"Jogo da Memória",texto:"Sessão com jogo da memória para estimulação de atenção sustentada e memória visual. Paciente acertou X pares em Y tentativas. Observou-se impulsividade na virada das peças. Estratégias de autorregulação foram trabalhadas. Conduta: aumentar número de pares e trabalhar planejamento."},
    {id:"pp2",titulo:"Quebra-Cabeça",texto:"Atividade com quebra-cabeça de grau crescente de complexidade. Paciente demonstrou dificuldade no planejamento visual e na busca de pistas contextuais. Tolerância à frustração em desenvolvimento. Conduta: trabalhar estratégias de resolução de problemas por etapas."},
    {id:"pp3",titulo:"Leitura e Interpretação",texto:"Atividade de leitura silenciosa seguida de questionamentos interpretativos. Paciente demonstrou decodificação adequada com dificuldade na compreensão inferencial. Vocabulário receptivo compatível com a faixa etária. Conduta: inserir atividades de predição e monitoramento de leitura."},
    {id:"pp4",titulo:"Escrita e Ortografia",texto:"Trabalhadas atividades de ditado e produção textual com foco em ortografia e coesão. Paciente apresentou trocas fonológicas e omissões. Texto com ideia central identificável. Conduta: trabalhar consciência fonológica e regras ortográficas de forma lúdica."},
  ],
  "Fisioterapia":[
    {id:"ft1",titulo:"Avaliação e ADM",texto:"Paciente apresentou-se com queixa de dor em repouso EVA X/10 e EVA Y/10 em movimento. Avaliação de amplitude de movimento (ADM): flexão X°, extensão Y°, rotação interna Z°. Força muscular grau W. Realizados exercícios de mobilização articular e alongamento. Conduta: progressão de carga na próxima sessão."},
    {id:"ft2",titulo:"Reabilitação da Marcha",texto:"Sessão com foco em treino de marcha, equilíbrio e propriocepção. Paciente realizou deambulação com auxílio de X metros com supervisão. Observou-se padrão em trendelenburg à direita. Conduta: fortalecer abdutores de quadril e progredir distância percorrida."},
    {id:"ft3",titulo:"Fortalecimento Muscular",texto:"Protocolo de fortalecimento muscular de quadríceps e isquiotibiais com carga progressiva. Paciente realizou X séries de Y repetições com carga de Z kg. Sem relato de dor durante execução. Conduta: aumentar carga em 10% na próxima sessão."},
    {id:"ft4",titulo:"Fisioterapia Respiratória",texto:"Sessão de fisioterapia respiratória com técnicas de higiene brônquica e expansão pulmonar. SpO2 pré X% e pós Y%. Ausculta com redução de ruídos adventícios. Conduta: manter protocolo e orientar exercícios respiratórios domiciliares."},
    {id:"ft5",titulo:"RPG / Postura",texto:"Sessão de RPG com cadeias musculares posteriores e anterolateral. Paciente relatou melhora do desconforto cervical ao final da sessão. Observados padrões posturais em correção. Conduta: trabalhar consciência postural nas AVDs."},
  ],
  "Fonoaudiologia":[
    {id:"fono1",titulo:"Trabalho de Fonemas",texto:"Sessão com foco em produção dos fonemas /X/ e /Y/ em posição inicial, medial e final de palavras. Paciente demonstrou produção adequada em nível de sílaba com dificuldade de generalização para palavras. Conduta: ampliar nível de complexidade linguística e inserir atividades contextualizadas."},
    {id:"fono2",titulo:"Nomeação e Vocabulário",texto:"Trabalhadas atividades de nomeação de figuras e categorização semântica. Paciente evocou X de Y itens com pistas fonológicas e semânticas. Vocabulário expressivo abaixo do esperado para a faixa etária. Conduta: ampliar repertório semântico com temas funcionais."},
    {id:"fono3",titulo:"Respiração e Fonação",texto:"Sessão de treino respiratório e de voz com exercícios de apoio diafragmático, projeção vocal e ressonância. Paciente demonstrou padrão clavicular predominante. Qualidade vocal melhorou ao longo dos exercícios. Conduta: treino domiciliar de respiração guiada."},
    {id:"fono4",titulo:"Consciência Fonológica",texto:"Atividades de rima, aliteração e manipulação fonêmica. Paciente demonstrou habilidade preservada em rima e dificuldade em segmentação fonêmica. Conduta: progredir para atividades de transposição e análise fonêmica."},
  ],
  "Terapia Ocupacional":[
    {id:"to1",titulo:"AVD — Vestuário",texto:"Trabalhadas atividades de vida diária com foco em vestuário: abotoamento, zíper e calçar sapatos. Paciente demonstrou dificuldade em coordenação bimanual fina. Realizou tarefa com supervisão e assistência mínima. Conduta: graduar dificuldade e trabalhar pinça e oponência."},
    {id:"to2",titulo:"Integração Sensorial",texto:"Sessão de estimulação sensorial com ênfase em propriocepção e processamento tátil. Paciente apresentou hipersensibilidade tátil nas mãos com tolerância aumentada ao longo da sessão. Conduta: manter protocolo de dessensibilização e orientar família sobre estratégias domiciliares."},
    {id:"to3",titulo:"Coordenação Motora Fina",texto:"Atividades de encaixe, colagem, recorte e uso de instrumentos. Paciente demonstrou preensão digital adequada com fadiga muscular precoce. Traçado gráfico em desenvolvimento. Conduta: introduzir atividades pré-grafismo e progressão de resistência."},
    {id:"to4",titulo:"Culinária Terapêutica",texto:"Atividade de culinária terapêutica para estimulação de sequenciamento, atenção e coordenação. Paciente seguiu recipe com apoio visual em X de Y etapas de forma independente. Conduta: aumentar complexidade da receita e reduzir suporte gradualmente."},
  ],
  "Coordenação ABA":[
    {id:"aba1",titulo:"Supervisão — Mando",texto:"Supervisionada sessão de aplicação de mando com paciente. Profissional demonstrou bom manejo de prompts. Foram discutidos critérios de transferência de dica e planejamento de generalização. Conduta: revisar programa de mando e atualizar critérios de domínio."},
    {id:"aba2",titulo:"Supervisão — Pareamento",texto:"Sessão de supervisão com foco em pareamento e construção de vínculo terapêutico. Discutidas estratégias de preferência e análise funcional. Profissional demonstrou boa sensibilidade ao ritmo do paciente. Conduta: implementar pesquisa de preferência estruturada."},
    {id:"aba3",titulo:"Supervisão — Redução de Comportamentos",texto:"Supervisão de protocolo de redução de comportamentos disruptivos. Analisados dados de linha de base e função do comportamento. Plano de intervenção revisado com foco em reforçamento diferencial. Conduta: treinar profissional em extinção com reforçamento alternativo."},
    {id:"aba4",titulo:"Reunião de Equipe",texto:"Reunião de equipe multiprofissional para alinhamento de objetivos terapêuticos. Discutidos resultados de avaliação VB-MAPP/ABLLS-R. Estabelecidos novos alvos de ensino e metas funcionais. Conduta: atualizar programação e comunicar família."},
  ],
  "Musicoterapia":[
    {id:"mt1",titulo:"Improvisação Rítmica",texto:"Sessão de improvisação rítmica com instrumentos de percussão. Paciente demonstrou engajamento e coordenação rítmica em desenvolvimento. Respostas emocionais observadas através da variação de intensidade e tempo. Conduta: explorar instrumentos melódicos e ampliar repertório expressivo."},
    {id:"mt2",titulo:"Canção Terapêutica",texto:"Utilização de canções conhecidas pelo paciente como recurso de expressão e comunicação. Paciente verbalizou trechos e demonstrou prazer durante a atividade. Observada melhora no engajamento e contato visual. Conduta: compor canção personalizada com temáticas do cotidiano."},
    {id:"mt3",titulo:"Percussão e Expressão",texto:"Sessão com ênfase em percussão corporal e instrumental para estimulação sensório-motora e expressão emocional. Paciente regulou intensidade do toque conforme consigna. Conduta: ampliar para instrumentos de sopro e exploração tímbrica."},
  ],
  "Nutrição":[
    {id:"nut1",titulo:"Exploração Sensorial Alimentar",texto:"Sessão de exploração sensorial com alimentos de diferentes texturas, cores e odores. Paciente aceitou contato visual com X itens e manipulação de Y itens. Sem contato oral nesta sessão. Conduta: ampliar exploração tátil e aproximação progressiva ao contato oral."},
    {id:"nut2",titulo:"Food Chaining",texto:"Aplicada técnica de food chaining para ampliar repertório alimentar. Partindo de alimento preferido, introduzido alimento similar em textura. Paciente aceitou a variação com mínima resistência. Conduta: progredir para alimentos com diferença de textura mais marcante."},
    {id:"nut3",titulo:"Culinária Terapêutica",texto:"Atividade de culinária terapêutica com participação ativa do paciente no preparo. Observado aumento de tolerância ao contato com novos alimentos durante o preparo. Paciente provou X alimentos de forma voluntária. Conduta: envolver família na replicação da atividade em domicílio."},
    {id:"nut4",titulo:"Orientação Familiar",texto:"Sessão de orientação com responsáveis sobre seletividade alimentar, estratégias comportamentais e ambiente alimentar. Discutidos erros e acertos das refeições em casa. Conduta: elaborar cardápio estruturado e combinados para as refeições."},
  ],
  "Psicoterapia":[
    {id:"pt1",titulo:"Psicoeducação Ansiedade",texto:"Sessão de psicoeducação sobre o transtorno de ansiedade: sintomas, manutenção e ciclo de evitação. Paciente demonstrou boa assimilação e identificou situações pessoais de ativação ansiosa. Conduta: iniciar registro de pensamentos e situações ansiogênicas."},
    {id:"pt2",titulo:"Reestruturação Cognitiva",texto:"Identificadas crenças nucleares e pensamentos automáticos negativos. Paciente realizou questionamento socrático com mediação do terapeuta. Foram trabalhados registros de evidências. Conduta: consolidar técnica de disputa cognitiva e ampliar para outros contextos."},
    {id:"pt3",titulo:"Regulação Emocional",texto:"Trabalhadas estratégias de regulação emocional: identificação, nomeação e manejo de emoções. Paciente identificou X emoções e relatou dificuldade em tolerar frustração. Conduta: exercícios de mindfulness e técnicas de aterramento."},
    {id:"pt4",titulo:"Exposição Gradual",texto:"Iniciado protocolo de exposição gradual hierárquica. Paciente construiu hierarquia de situações temidas. Realizada exposição imaginária ao item X com SUDs inicial de Y e final de Z. Conduta: progredir hierarquia e introduzir exposição in vivo."},
    {id:"pt5",titulo:"Luto e Perdas",texto:"Sessão dedicada ao processamento de luto. Paciente narrou aspectos da perda com carga emocional significativa. Trabalhados estágios do luto e ressignificação. Conduta: trabalhar memórias positivas e construção de narrativa adaptativa."},
  ],
};
const SETORES_CHAMADO=["Agendamento","Recepção","Faturamento","Coordenação","TI","Gestão de Pessoas","Administração"];
const STATUS_CHAMADO=["aberto","andamento","encerrado","novo_paciente","novo_agendamento","devolvido"];
const COR_CHAMADO={aberto:"#f59e0b",andamento:"#3b82f6",encerrado:"#10b981",novo_paciente:"#a78bfa",novo_agendamento:"#38bdf8",devolvido:"#f87171"};
const LABEL_CHAMADO={aberto:"📂 Aberto",andamento:"🔄 Andamento",encerrado:"✅ Encerrado",novo_paciente:"👤 Novo Paciente",novo_agendamento:"📅 Novo Agendamento",devolvido:"↩️ Devolvido"};
const PRIORIDADE_CHAMADO=["Baixa","Normal","Alta","Urgente"];
const COR_PRIORIDADE={Baixa:"#94a3b8",Normal:"#a78bfa",Alta:"#f59e0b",Urgente:"#ef4444"};
const STATUS_CONTRATO=["pendente","enviado","assinado","concluido","cancelado"];
const COR_CONTRATO={pendente:"#94a3b8",enviado:"#f59e0b",assinado:"#a78bfa",concluido:"#22c55e",cancelado:"#ef4444"};
const LABEL_CONTRATO={pendente:"⏳ Pendente",enviado:"📤 Enviado",assinado:"✍️ Assinado",concluido:"✅ Concluído",cancelado:"❌ Cancelado"};
const PERFIS=["profissional","atendimento","secretaria","supervisor_adm","coordenador","administrador","faturamento_supervisor","gestao_pessoas","agendamento"];
const PERFIL_LABEL={profissional:"Profissional",atendimento:"Atendimento",secretaria:"Secretária",supervisor_adm:"Supervisor ADM",coordenador:"Coordenador",administrador:"Administrador",faturamento_supervisor:"Sup. Faturamento",gestao_pessoas:"Gestão de Pessoas",agendamento:"Agendamento"};
// STATUS conforme especificação — cores, crédito convênio, repasse, permissões
const STATUS_AG={
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
// Status exclusivos do Mapeamento/Espelho (não aparecem no AgendaModal)
const STATUS_MAPA_ONLY={
  reposicao: {label:"Reposição (CF)", color:"#1e90ff", icon:"🔄"},
  fechado:   {label:"Fechado",        color:"#1e293b", icon:"🔒"},
};
const PODE_STATUS=r=>["atendimento","supervisor_adm","administrador","agendamento"].includes(r);
const PODE_FATURAR=r=>["faturamento_supervisor","supervisor_adm","administrador"].includes(r);
const PODE_PACIENTE=r=>["agendamento","administrador"].includes(r);
const PODE_PROFISSIONAL=r=>["gestao_pessoas","administrador"].includes(r);
const PODE_EVOLUIR=r=>["profissional","coordenador","administrador"].includes(r);
const REPETICAO_OPTS=["Nenhuma","Semanal","Quinzenal","Mensal","Toda Segunda","Toda Terça","Toda Quarta","Toda Quinta","Toda Sexta","Todo Sábado"];

// Log de atividades
function criarLog(atividades,setAtividades,usuario,acao,detalhe){
  if(!setAtividades)return;
  setAtividades(a=>[{id:Date.now(),data:hoje_str,hora:new Date().toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"}),usuario,acao,detalhe},...a]);
}

// ═══════════════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════════════
const hoje=new Date();
const ymd=d=>{const dd=new Date(d);return dd.getFullYear()+"-"+String(dd.getMonth()+1).padStart(2,"0")+"-"+String(dd.getDate()).padStart(2,"0");};
const brDate=s=>{if(!s)return"";const[y,m,d]=s.split("-");return d+"/"+m+"/"+y;};
const getDIM=(y,m)=>new Date(y,m+1,0).getDate();
const getFD=(y,m)=>new Date(y,m,1).getDay();
const toMin=h=>{const[a,b]=String(h||"00:00").split(":").map(Number);return a*60+b;};
const toTime=t=>String(Math.floor(t/60)%24).padStart(2,"0")+":"+String(t%60).padStart(2,"0");
const addMin=(h,m)=>toTime(toMin(h)+Number(m||0));
const overlaps=(a1,a2,b1,b2)=>toMin(a1)<toMin(b2)&&toMin(b1)<toMin(a2);
const brl=v=>new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"}).format(Number(v||0));
const profShort=n=>{if(!n)return"";const p=n.split(" ");return p.length<=2?n:p[0]+" "+p[p.length-1];};
const rawD=v=>(v||"").replace(/\D/g,"");
const maskCPF=v=>rawD(v).slice(0,11).replace(/(\d{3})(\d)/,"$1.$2").replace(/(\d{3})(\d)/,"$1.$2").replace(/(\d{3})(\d{1,2})$/,"$1-$2");
const maskPhone=v=>{const d=rawD(v).slice(0,11);return d.length<=10?d.replace(/(\d{2})(\d{4})(\d{0,4})/,"($1) $2-$3"):d.replace(/(\d{2})(\d{5})(\d{0,4})/,"($1) $2-$3");};
const maskCEP=v=>rawD(v).slice(0,8).replace(/(\d{5})(\d{0,3})/,"$1-$2");
const diffDays=(d1,d2)=>Math.ceil((new Date(d2)-new Date(d1))/(1000*86400));
const hoje_str=ymd(hoje);
async function buscaCEP(cep){try{const r=await fetch("https://viacep.com.br/ws/"+rawD(cep).slice(0,8)+"/json/");const d=await r.json();if(d.erro)return null;return{logradouro:d.logradouro,bairro:d.bairro,cidade:d.localidade,estado:d.uf};}catch{return null;}}

// Gera datas de repetição
function gerarRepeticoes(dataBase,tipo,qtd,indefinida){
  if(!tipo||tipo==="Nenhuma")return[dataBase];
  const real=indefinida?104:qtd;
  if(!real||real<=1)return[dataBase];
  const datas=[dataBase];
  const base=new Date(dataBase+"T12:00:00");
  const diaSemMap={"Toda Segunda":1,"Toda Terça":2,"Toda Quarta":3,"Toda Quinta":4,"Toda Sexta":5,"Todo Sábado":6};
  for(let i=1;i<real;i++){
    const d=new Date(base);
    if(tipo==="Semanal"){d.setDate(base.getDate()+7*i);}
    else if(tipo==="Quinzenal"){d.setDate(base.getDate()+14*i);}
    else if(tipo==="Mensal"){d.setMonth(base.getMonth()+i);}
    else if(diaSemMap[tipo]!==undefined){d.setDate(base.getDate()+7*i);}
    datas.push(ymd(d));
  }
  return datas;
}

// ═══════════════════════════════════════════════════════════════════════════════
// SEED DATA
// ═══════════════════════════════════════════════════════════════════════════════
const escalaDefault=DIAS_SEMANA.reduce((a,d)=>({...a,[d]:{ativo:!["Domingo","Sábado"].includes(d),inicio:"08:00",fim:"18:00",pausaInicio:"12:00",pausaFim:"13:00",usarPausa:false}}),{});

const seedFiliais=[
  {id:1,nome:"Unidade 1 — MC 1118",codigo:"MC1118",logradouro:"Rua Coronel Cardoso de Siqueira",numero:"1118",bairro:"Vila Oliveira",cidade:"Mogi das Cruzes",estado:"SP",telefone:"",ativa:true},
  {id:2,nome:"Unidade 2 — MC 1180",codigo:"MC1180",logradouro:"Rua Coronel Cardoso de Siqueira",numero:"1180",bairro:"Vila Oliveira",cidade:"Mogi das Cruzes",estado:"SP",telefone:"",ativa:true},
  {id:3,nome:"Unidade 3 — Cruzeiro",codigo:"MC701",logradouro:"Rua Cruzeiro do Sul",numero:"701",bairro:"Vila Oliveira",cidade:"Mogi das Cruzes",estado:"SP",telefone:"",ativa:true},
  {id:4,nome:"Unidade 4 — Acqua",codigo:"ACQUA",logradouro:"Rua Maurício Schwartzmann",numero:"46",bairro:"Vila Oliveira",cidade:"Mogi das Cruzes",estado:"SP",telefone:"",ativa:true},
  {id:5,nome:"Unidade Suzano",codigo:"SUZ461",logradouro:"Rua Felício de Camargo",numero:"461",bairro:"Centro",cidade:"Suzano",estado:"SP",telefone:"",ativa:true},
  {id:6,nome:"Unidade Petrolina",codigo:"PETRO178",logradouro:"Rua Aureliano Francisco Neto",numero:"178",bairro:"Atrás da Banca",cidade:"Petrolina",estado:"PE",telefone:"",ativa:true},
];
const seedSalas=[
  {id:1,filialId:1,nome:"Sala 01",especialidade:"Psicologia Convencional/Psicanálise",cor:"#8B4513",ativa:true,custoMensal:1500},
  {id:2,filialId:1,nome:"Sala 02",especialidade:"Terapia Ocupacional",cor:"#DAA520",ativa:true,custoMensal:1200},
  {id:3,filialId:1,nome:"Sala 03",especialidade:"Fonoaudiologia",cor:"#9370DB",ativa:true,custoMensal:1200},
  {id:4,filialId:2,nome:"Sala 01",especialidade:"Fisioterapia",cor:"#6A0572",ativa:true,custoMensal:1000},
  {id:5,filialId:2,nome:"Sala 02",especialidade:"",cor:"#38bdf8",ativa:true,custoMensal:800},
];
const seedProfissionais=[
  {id:1,nome:"Dra. Ana Souza",usuario:"11111111111",senha:"1111",role:"profissional",especialidades:["Psicologia Convencional/Psicanálise"],temposAtendimento:[50],escala:{...escalaDefault},temConselho:true,conselho:"CRP",registroConselho:"06/12345",conveniosAtendidos:["Particular","Unimed"],filiaisAtendimento:[1,2],filialAcesso:[1,2]},
  {id:2,nome:"Dr. Carlos Lima",usuario:"22222222222",senha:"2222",role:"profissional",especialidades:["Fisioterapia","Natação Terapêutica"],temposAtendimento:[40,50],escala:{...escalaDefault},temConselho:true,conselho:"CREFITO",registroConselho:"3/12345",conveniosAtendidos:["Particular","Bradesco Saúde"],filiaisAtendimento:[1],filialAcesso:[1]},
  {id:3,nome:"Coordenador Silva",usuario:"33333333333",senha:"3333",role:"coordenador",especialidades:["Coordenação ABA"],temposAtendimento:[50],escala:{...escalaDefault},temConselho:false,conveniosAtendidos:[],filiaisAtendimento:[1,2,3],filialAcesso:[1,2,3]},
  {id:4,nome:"Admin Sistema",usuario:"31028313896",senha:"122",role:"administrador",especialidades:["Administrador"],temposAtendimento:[50],escala:{...escalaDefault},temConselho:false,conveniosAtendidos:[],filiaisAtendimento:seedFiliais.map(f=>f.id),filialAcesso:seedFiliais.map(f=>f.id)},
  {id:5,nome:"Faturamento User",usuario:"55555555555",senha:"5555",role:"faturamento_supervisor",especialidades:["Faturamento"],temposAtendimento:[50],escala:{...escalaDefault},temConselho:false,conveniosAtendidos:[],filiaisAtendimento:[1],filialAcesso:[1]},
];
const seedPacientes=[
  {id:1,nome:"João Pedro Alves",nascimento:"2010-03-15",cpf:"12345678901",sexo:"M",celular:"11999990001",convenio:"Unimed",plano:"Nacional Plus",cep:"08773100",logradouro:"Rua das Flores",numero:"100",bairro:"Centro",cidade:"Mogi das Cruzes",estado:"SP",resp1Nome:"Maria Alves",resp1Cpf:"98765432100",resp1Whatsapp:"11999990000",infoImportantes:"TEA leve"},
  {id:2,nome:"Beatriz Costa",nascimento:"2015-07-20",cpf:"23456789012",sexo:"F",celular:"11999990002",convenio:"Bradesco Saúde",plano:"Top Nacional",cep:"08773200",logradouro:"Av. Principal",numero:"200",bairro:"Vila Nova",cidade:"Mogi das Cruzes",estado:"SP",resp1Nome:"Paula Costa",resp1Cpf:"87654321009",resp1Whatsapp:"11999990003",infoImportantes:""},
];
const seedConvenios=[
  {id:1,nome:"Particular",codigo:"PART",contato:"—",manual:""},
  {id:2,nome:"ABM Saúde",codigo:"ABM",contato:"",manual:""},
  {id:3,nome:"Alice (ABA)",codigo:"ALICE",contato:"",manual:""},
  {id:4,nome:"Cassi",codigo:"CASSI",contato:"",manual:""},
  {id:5,nome:"Hospital Cruz Azul (ABA)",codigo:"HCA",contato:"",manual:""},
  {id:6,nome:"Porto Seguro (ABA)",codigo:"PS",contato:"",manual:""},
  {id:7,nome:"Saúde Caixa",codigo:"SCAIXA",contato:"",manual:""},
  {id:8,nome:"SulAmerica (ABA)",codigo:"SULA",contato:"",manual:""},
  {id:9,nome:"UniVida",codigo:"UNIVIDA",contato:"",manual:""},
  {id:10,nome:"Saúde Abas",codigo:"SABAS",contato:"",manual:""},
  {id:11,nome:"Vivest",codigo:"VIVEST",contato:"",manual:""},
  {id:12,nome:"Unimed CNU",codigo:"UNICNU",contato:"",manual:""},
  {id:13,nome:"Metrus",codigo:"METRUS",contato:"",manual:""},
  {id:14,nome:"Blue Saúde",codigo:"BLUE",contato:"",manual:""},
  {id:15,nome:"Bradesco Saúde (ABA)",codigo:"BRAD",contato:"0800 722 0101",manual:""},
  {id:16,nome:"MediService",codigo:"MEDIS",contato:"",manual:""},
  {id:17,nome:"Cruz Azul Saúde (ABA)",codigo:"CAS",contato:"",manual:""},
  {id:18,nome:"OAB / ACMC / CAASP",codigo:"PARC",contato:"",manual:""},
];
const seedProcedimentos=[
  {id:1,nome:"Psicologia",categoria:"Terapia",duracao:50,valor:200,pctRepasse:45,convenio:"Particular"},
  {id:2,nome:"Terapia Ocupacional",categoria:"Terapia",duracao:50,valor:200,pctRepasse:45,convenio:"Particular"},
  {id:3,nome:"Fonoaudiologia",categoria:"Terapia",duracao:50,valor:200,pctRepasse:45,convenio:"Particular"},
  {id:4,nome:"Fisioterapia",categoria:"Reabilitação",duracao:50,valor:180,pctRepasse:45,convenio:"Particular"},
  {id:5,nome:"Avaliação Neuropsicológica (infantil)",categoria:"Avaliação",duracao:50,valor:2600,pctRepasse:50,convenio:"Particular"},
  {id:6,nome:"Avaliação Neuropsicológica (adulto)",categoria:"Avaliação",duracao:50,valor:3600,pctRepasse:50,convenio:"Particular"},
  {id:7,nome:"Nutrição",categoria:"Terapia",duracao:50,valor:180,pctRepasse:45,convenio:"Particular"},
  {id:8,nome:"Terapia Assistida por Animais",categoria:"Terapia",duracao:50,valor:170,pctRepasse:40,convenio:"Particular"},
  {id:9,nome:"Psicomotricidade",categoria:"Terapia",duracao:50,valor:180,pctRepasse:45,convenio:"Particular"},
  {id:10,nome:"Coordenação ABA",categoria:"ABA",duracao:50,valor:260,pctRepasse:45,convenio:"Particular"},
  {id:11,nome:"Aplicador ABA",categoria:"ABA",duracao:50,valor:120,pctRepasse:40,convenio:"Particular"},
  {id:12,nome:"Psicopedagogia",categoria:"Terapia",duracao:50,valor:180,pctRepasse:45,convenio:"Particular"},
  {id:13,nome:"Musicoterapia",categoria:"Terapia",duracao:50,valor:170,pctRepasse:40,convenio:"Particular"},
  {id:14,nome:"Arteterapia",categoria:"Terapia",duracao:50,valor:170,pctRepasse:40,convenio:"Particular"},
  {id:15,nome:"Neurofeedback (sessão)",categoria:"Neurofeedback",duracao:50,valor:220,pctRepasse:45,convenio:"Particular"},
  {id:16,nome:"Mapeamento Neurofeedback",categoria:"Neurofeedback",duracao:120,valor:550,pctRepasse:50,convenio:"Particular"},
  {id:17,nome:"Snoezelen",categoria:"Terapia",duracao:50,valor:120,pctRepasse:40,convenio:"Particular"},
];
const seedManuais=[
  {id:1,titulo:"Manual Unimed",descricao:"Guia de autorização e faturamento",url:"https://www.unimed.coop.br",arquivo:null,data:hoje_str},
  {id:2,titulo:"Manual Bradesco",descricao:"Procedimentos e tabela",url:"https://www.bradescosaude.com.br",arquivo:null,data:hoje_str},
];
const seedAlertas=[
  {id:1,titulo:"Agenda de hoje",descricao:"3 pacientes confirmados",lido:false},
  {id:2,titulo:"Pedido vencendo",descricao:"João Pedro — autorização vence em 5 dias",lido:false},
];
const seedRepasses=[{id:1,profissionalId:1,percentual:60},{id:2,profissionalId:2,percentual:55}];
const seedAgenda=(()=>{
  const base=[];
  const hoje=new Date();
  const ymd=d=>d.toISOString().slice(0,10);
  // Gera sessões semanais para demonstração
  const sessoes=[
    {pacienteId:1,profissionalId:1,salaId:1,filialId:1,procedimentoId:1,horarioSessao:"08:00",horarioFimSessao:"08:50",tempoSessao:50,convenio:"Unimed",plano:"Nacional Plus",status:"agendado"},
    {pacienteId:2,profissionalId:2,salaId:4,filialId:2,procedimentoId:2,horarioSessao:"09:00",horarioFimSessao:"09:40",tempoSessao:40,convenio:"Bradesco Saúde",plano:"Top Nacional",status:"agendado"},
    {pacienteId:1,profissionalId:2,salaId:4,filialId:2,procedimentoId:2,horarioSessao:"10:00",horarioFimSessao:"10:40",tempoSessao:40,convenio:"Unimed",plano:"Nacional Plus",status:"agendado"},
  ];
  let id=1000;
  for(let w=0;w<4;w++){
    sessoes.forEach((s,si)=>{
      const d=new Date(hoje);
      d.setDate(d.getDate()+(w*7)+si);
      if(d.getDay()===0||d.getDay()===6)d.setDate(d.getDate()+1);
      base.push({...s,id:id++,data:ymd(d),repeticaoGrupo:900+si});
    });
  }
  // Alguns com status variados para demonstrar cores
  const ontem=new Date(hoje);ontem.setDate(ontem.getDate()-1);
  base.push({id:id++,pacienteId:1,profissionalId:1,salaId:1,filialId:1,procedimentoId:1,data:ymd(ontem),horarioSessao:"08:00",horarioFimSessao:"08:50",tempoSessao:50,convenio:"Unimed",plano:"Nacional Plus",status:"atendido"});
  base.push({id:id++,pacienteId:2,profissionalId:2,salaId:4,filialId:2,procedimentoId:2,data:ymd(ontem),horarioSessao:"09:00",horarioFimSessao:"09:40",tempoSessao:40,convenio:"Bradesco Saúde",plano:"Top Nacional",status:"faltou"});
  return base;
})();

// ═══════════════════════════════════════════════════════════════════════════════
// CSS (string concatenation — sem template literals para evitar parse errors)
// ═══════════════════════════════════════════════════════════════════════════════
function getThemeVars(t){
  return ":root{--bg:"+t.bg+";--sb:"+t.sidebar+";--card:"+t.card+";--cb:"+t.cardBorder+
    ";--inp:"+t.inp+";--ib:"+t.inpBorder+";--tx:"+t.text+";--mt:"+t.muted+
    ";--lb:"+t.label+";--sx:"+t.secBox+";--sc:"+t.secBorder+
    ";--fb:"+t.filterBar+";--fc:"+t.filterBorder+";--ch:"+t.calHead+
    ";--dy:"+t.day+";--db:"+t.dayBorder+";--dh:"+t.dayHover+
    ";--nv:"+t.navBtn+";--nh:"+t.navHover+";--na:"+t.navActive+
    ";--gr:"+t.gridRow+";--md:"+t.modal+";--mb:"+t.modalBg+
    ";--up:"+t.upload+";--st:"+t.scrollTrack+";--sm:"+t.scrollThumb+
    ";--cpb:"+t.chipBorder+";--sbd:"+t.sideBorder+"}";}

const CSS="*{box-sizing:border-box}body{margin:0}"+
".app-shell{font-family:'DM Sans','Segoe UI',sans-serif;min-height:100vh;background:var(--bg);color:var(--tx);display:flex;overflow:hidden;height:100vh}"+
/* ── SIDEBAR ─────────────────────────────────────────────────────── */
".sidebar{background:var(--sb);border-right:1px solid var(--sbd);display:flex;flex-direction:column;flex-shrink:0;transition:width .22s cubic-bezier(.4,0,.2,1);overflow:hidden;height:100vh;position:relative;z-index:30}"+
".sidebar-inner{display:flex;flex-direction:column;height:100%;min-width:0}"+
".sidebar-brand{display:flex;align-items:center;gap:10px;padding:14px 13px 10px;border-bottom:1px solid var(--sbd);flex-shrink:0}"+
".brand-logo{width:40px;height:40px;border-radius:14px;background:#7c6af7;display:flex;align-items:center;justify-content:center;flex-shrink:0;color:#fff;box-shadow:0 6px 18px #7c6af770;position:relative;overflow:hidden}"+
".brand-text{font-family:'DM Serif Display',serif;font-weight:900;font-size:18px;color:#7c6af7;white-space:nowrap;overflow:hidden;line-height:1.1}"+
".sidebar-user{padding:8px 12px;display:flex;align-items:center;gap:9px;border-bottom:1px solid var(--sbd);flex-shrink:0}"+
".user-avatar{width:28px;height:28px;border-radius:8px;background:#7c6af720;border:1px solid #7c6af730;display:flex;align-items:center;justify-content:center;font-size:12px;flex-shrink:0;font-weight:800;color:#a78bfa}"+
".user-info{overflow:hidden;flex:1}"+
".user-name{font-size:11px;font-weight:800;color:var(--tx);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}"+
".user-role{font-size:9px;font-weight:700;letter-spacing:.4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}"+
".sidebar-nav{flex:1;overflow-y:auto;overflow-x:hidden;padding:6px 6px;display:flex;flex-direction:column;gap:1px}"+
".sidebar-nav::-webkit-scrollbar{width:3px}.sidebar-nav::-webkit-scrollbar-track{background:transparent}.sidebar-nav::-webkit-scrollbar-thumb{background:var(--sm);border-radius:3px}"+
".nav-btn{display:flex;align-items:center;gap:9px;width:100%;padding:7px 9px;border-radius:8px;border:none;background:transparent;color:var(--nv);font-weight:600;font-size:11.5px;cursor:pointer;text-align:left;transition:.12s;white-space:nowrap;position:relative;min-height:34px}"+
".nav-btn:hover{background:var(--nh);color:#7c6af7}"+
".nav-btn.active{background:#7c6af7;color:#fff;font-weight:800}"+
".nav-icon{font-size:15px;flex-shrink:0;width:20px;text-align:center;line-height:1}"+
".nav-label{overflow:hidden;text-overflow:ellipsis;flex:1}"+
".nav-badge{background:#ef4444;color:#fff;border-radius:10px;font-size:9px;font-weight:900;padding:1px 5px;min-width:14px;text-align:center;flex-shrink:0}"+
".nav-tooltip{position:absolute;left:calc(100% + 8px);top:50%;transform:translateY(-50%);background:#1e293b;color:#e2e8f0;padding:4px 9px;border-radius:6px;font-size:11px;font-weight:700;white-space:nowrap;pointer-events:none;opacity:0;transition:opacity .15s;z-index:999;box-shadow:0 4px 12px rgba(0,0,0,.3)}"+
".nav-btn:hover .nav-tooltip{opacity:1}"+
".sidebar-footer{padding:8px 6px;border-top:1px solid var(--sbd);display:flex;flex-direction:column;gap:4px;flex-shrink:0}"+
".sidebar-toggle{position:absolute;right:-13px;top:22px;width:26px;height:26px;border-radius:50%;background:var(--sb);border:1.5px solid var(--sbd);display:flex;align-items:center;justify-content:center;cursor:pointer;color:var(--mt);font-size:11px;z-index:40;transition:all .15s;box-shadow:0 2px 8px rgba(0,0,0,.12)}"+
".sidebar-toggle:hover{background:var(--na);color:#7c6af7;border-color:#7c6af7}"+
/* ── ALERT PANEL desliza do sidebar ──────────────────────────────── */
".alert-panel{position:fixed;left:0;top:0;bottom:0;width:300px;background:var(--sb);border-right:1px solid var(--sbd);z-index:60;transform:translateX(-100%);transition:transform .25s cubic-bezier(.4,0,.2,1);box-shadow:4px 0 24px rgba(0,0,0,.18);display:flex;flex-direction:column}"+
".alert-panel.open{transform:translateX(0)}"+
".alert-overlay{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:59;opacity:0;pointer-events:none;transition:opacity .25s}"+
".alert-overlay.open{opacity:1;pointer-events:all}"+
/* ── MAIN ────────────────────────────────────────────────────────── */
".main{flex:1;display:flex;flex-direction:column;overflow:hidden;position:relative;min-width:0}"+
/* ── CARDS ───────────────────────────────────────────────────────── */
".card{background:var(--card);border:1px solid var(--cb);border-radius:14px;transition:box-shadow .15s}"+
".rh{cursor:pointer}.rh:hover{box-shadow:0 2px 12px rgba(124,106,247,.10);border-color:#7c6af740!important}"+
/* ── INPUTS ──────────────────────────────────────────────────────── */
"input,select,textarea{width:100%;background:var(--inp);border:1.5px solid var(--ib);color:var(--tx);border-radius:9px;padding:8px 12px;font-family:inherit;font-size:13px;transition:border-color .15s,box-shadow .15s}"+
"input:focus,select:focus,textarea:focus{outline:none;border-color:#7c6af7;box-shadow:0 0 0 3px #7c6af715}"+
"select option{background:var(--card)}"+
"label{font-size:10px;font-weight:900;color:var(--lb);letter-spacing:.7px;display:block;margin-bottom:3px;text-transform:uppercase}"+
/* ── BUTTONS ─────────────────────────────────────────────────────── */
".btn{cursor:pointer;border:none;transition:all .15s;font-family:inherit;padding:8px 14px;border-radius:9px;font-weight:800;font-size:12px;letter-spacing:.2px}"+
".btn:hover{filter:brightness(1.08);transform:translateY(-1px)}.btn:active{transform:translateY(0)}.btn:disabled{opacity:.4;cursor:default;transform:none}"+
".btn.primary{background:#7c6af7;color:#fff;box-shadow:0 2px 8px #7c6af740}"+
".btn.secondary{background:var(--gr);color:var(--mt);border:1.5px solid var(--cb)}"+
".btn.ok{background:linear-gradient(135deg,#065f46,#047857);color:#34d399}"+
".btn.danger{background:linear-gradient(135deg,#450a0a,#7f1d1d);color:#f87171}"+
".btn.warn{background:linear-gradient(135deg,#451a03,#78350f);color:#fbbf24}"+
".btn.small{padding:4px 9px;font-size:11px;border-radius:7px}"+
/* ── FILTER BAR ──────────────────────────────────────────────────── */
".filter-bar{display:flex;gap:7px;padding:8px 12px;background:var(--fb);border-bottom:1.5px solid var(--fc);align-items:flex-end;flex-wrap:wrap;flex-shrink:0}"+
".filter-bar select,.filter-bar input{min-width:100px;width:auto}"+
/* ── MODAL ───────────────────────────────────────────────────────── */
".modal-bg{position:fixed;inset:0;background:var(--mb);backdrop-filter:blur(8px);z-index:100;display:flex;align-items:flex-start;justify-content:center;padding:14px;overflow-y:auto}"+
".modal{background:var(--md);border:1px solid var(--sc);border-radius:18px;padding:22px;width:540px;max-width:100%;margin:auto;animation:popIn .18s ease}"+
"@keyframes popIn{from{opacity:0;transform:scale(.95) translateY(10px)}to{opacity:1;transform:none}}"+
".modal-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px}"+
".modal-head h2{margin:0;font-family:'DM Serif Display',serif;font-weight:900;font-size:18px}"+
".icon-btn{background:transparent;color:var(--mt);border:none;font-size:22px;cursor:pointer;line-height:1;border-radius:7px;padding:2px 5px;transition:.12s}"+
".icon-btn:hover{background:var(--sx);color:var(--tx)}"+
/* ── LAYOUT HELPERS ──────────────────────────────────────────────── */
".stack{display:flex;flex-direction:column;gap:11px}"+
".g2{display:grid;grid-template-columns:1fr 1fr;gap:11px}"+
".g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px}"+
".g4{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:9px}"+
".section-box{background:var(--sx);border:1.5px solid var(--sc);border-radius:10px;padding:12px}"+
".section-title{font-size:10px;font-weight:900;color:var(--lb);letter-spacing:.7px;text-transform:uppercase;margin-bottom:9px}"+
".helper{font-size:10px;color:var(--mt);margin-top:2px}"+
".actions{display:flex;gap:7px;margin-top:3px;justify-content:flex-end}"+
".page-wrap{padding:16px;overflow:auto;flex:1;min-height:0}"+
".page-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:7px}"+
".page-head h1{margin:0;font-family:'DM Serif Display',serif;font-weight:900;font-size:20px}"+
".muted{color:var(--mt)}"+
/* ── CALENDAR ────────────────────────────────────────────────────── */
".calendar-head{padding:9px 13px;background:var(--ch);border-bottom:1px solid var(--fc);display:flex;align-items:center;gap:7px;flex-shrink:0}"+
".calendar-title{font-family:'DM Serif Display',serif;font-weight:900;font-size:16px;min-width:190px}"+
".dow-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;margin-bottom:2px}"+
".dow{text-align:center;font-size:9px;font-weight:900;color:var(--lb);letter-spacing:.8px;padding:2px 0}"+
".month-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:3px}"+
".day-cell{border-radius:9px;background:var(--dy);border:1px solid var(--db);transition:all .1s;cursor:pointer;min-height:82px;padding:4px}"+
".day-cell:hover{border-color:#7c6af755;background:var(--dh);box-shadow:0 2px 8px rgba(124,106,247,.08)}"+
".day-cell.today{border-color:#7c6af7;box-shadow:0 0 0 2px #7c6af720}"+
".day-cell.selected{border-color:#7c6af7;background:var(--na)}"+
".day-top{display:flex;justify-content:space-between;align-items:center;margin-bottom:2px}"+
".day-num{font-size:11px;font-weight:800;color:#7c6af7}"+
/* ── GRID TABLES ─────────────────────────────────────────────────── */
".grid-header,.grid-row{display:grid;gap:7px;padding:9px 14px;align-items:center}"+
".grid-header{border-bottom:1px solid var(--fc);font-size:10px;font-weight:900;color:var(--lb);letter-spacing:.7px;background:var(--sx)}"+
".grid-row{border-bottom:1px solid var(--db);transition:background .1s}.grid-row:hover{background:var(--dh)!important}"+
/* ── MISC ────────────────────────────────────────────────────────── */
".chips{display:flex;flex-wrap:wrap;gap:5px;margin-top:3px}"+
".chip-btn{padding:4px 9px;border-radius:999px;font-size:11px;font-weight:900;border:1.5px solid var(--cpb);cursor:pointer;transition:all .12s;background:transparent;color:var(--mt)}"+
".tab-bar{display:flex;gap:3px;margin-bottom:12px;border-bottom:1px solid var(--fc);padding-bottom:7px;flex-wrap:wrap}"+
".tab-btn{padding:5px 11px;border-radius:6px;border:none;font-family:inherit;font-size:11px;cursor:pointer;transition:all .12s;font-weight:700}"+
".upload-zone{border:2px dashed var(--up);border-radius:10px;padding:14px;text-align:center;cursor:pointer;color:var(--mt);transition:.15s;font-size:12px}"+
".upload-zone:hover{border-color:#7c6af7;color:#a78bfa;background:#7c6af708}"+
".file-chip{display:inline-flex;align-items:center;gap:3px;background:#7c6af720;color:#a78bfa;padding:2px 7px;border-radius:20px;font-size:11px;font-weight:700;margin:2px}"+
/* ── TOAST ───────────────────────────────────────────────────────── */
".toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%);padding:10px 20px;border-radius:10px;font-weight:800;font-size:13px;z-index:999;animation:slideUp .25s ease;white-space:nowrap;pointer-events:none;box-shadow:0 4px 20px rgba(0,0,0,.25);backdrop-filter:blur(8px)}"+
"@keyframes slideUp{from{opacity:0;transform:translate(-50%,10px)}to{opacity:1;transform:translate(-50%,0)}}"+
/* ── FAB CHAMADO ─────────────────────────────────────────────────── */
".fab-chamado{position:fixed;right:20px;bottom:20px;z-index:50;width:52px;height:52px;border-radius:50%;background:#7c6af7;border:none;color:#fff;cursor:pointer;box-shadow:0 4px 20px #7c6af760;display:flex;align-items:center;justify-content:center;font-size:22px;transition:all .2s;animation:fabPop .3s ease}"+
".fab-chamado:hover{transform:scale(1.1) translateY(-2px);box-shadow:0 8px 30px #7c6af780}"+
".fab-chamado:active{transform:scale(.96)}"+
"@keyframes fabPop{from{opacity:0;transform:scale(.7)}to{opacity:1;transform:scale(1)}}"+
"@keyframes micPulse{0%,100%{box-shadow:0 0 0 0 #ef444450}50%{box-shadow:0 0 0 7px #ef444415}}"+
/* ── SCROLLBARS ──────────────────────────────────────────────────── */
"::-webkit-scrollbar{width:5px;height:5px}"+
"::-webkit-scrollbar-track{background:transparent}"+
"::-webkit-scrollbar-thumb{background:var(--sm);border-radius:4px}"+
"::-webkit-scrollbar-thumb:hover{background:var(--mt)}"+
"input[type=checkbox]{width:auto}input[type=number]{width:auto}input:read-only{color:var(--mt)}"+
".info-box{padding:9px 12px;border-radius:9px;font-size:11px;border:1px solid}";

// ═══════════════════════════════════════════════════════════════════════════════
// BASE COMPONENTS
// ═══════════════════════════════════════════════════════════════════════════════
function AlertPanel({alertas,setAlertas,pacientes,profissionais,agenda,auth,open,onClose}){
  const pend=alertas.filter(a=>!a.lido);
  const [editando,setEditando]=useState(null);
  const [novoModal,setNovoModal]=useState(false);
  const [form,setForm]=useState({titulo:"",descricao:""});

  // ─── Aniversariantes ───────────────────────────────────────────────────────
  const hoje=new Date();
  const mm=String(hoje.getMonth()+1).padStart(2,"0");
  const dd=String(hoje.getDate()).padStart(2,"0");
  const amanha=new Date(hoje); amanha.setDate(amanha.getDate()+1);
  const mm2=String(amanha.getMonth()+1).padStart(2,"0");
  const dd2=String(amanha.getDate()).padStart(2,"0");
  const proxDias=7;

  const calcAniversariantes=(lista,tipo)=>(lista||[]).filter(p=>p.nascimento).map(p=>{
    const [y,m,d]=p.nascimento.split("-");
    const diff=(()=>{
      const aniv=new Date(hoje.getFullYear(),Number(m)-1,Number(d));
      if(aniv<new Date(hoje.getFullYear(),hoje.getMonth(),hoje.getDate())) aniv.setFullYear(aniv.getFullYear()+1);
      return Math.round((aniv-new Date(hoje.getFullYear(),hoje.getMonth(),hoje.getDate()))/(1000*60*60*24));
    })();
    const anos=hoje.getFullYear()-Number(y)-(diff===0?0:(diff>0&&diff<=proxDias?0:1));
    return({...p,_tipo:tipo,_diff:diff,_anos:diff===0?anos:hoje.getFullYear()-Number(y)-(diff>0?1:0)+1,_m:m,_d:d});
  }).filter(p=>p._diff<=proxDias).sort((a,b)=>a._diff-b._diff);

  const anivPacientes=calcAniversariantes(pacientes,"paciente");
  const anivProfs=calcAniversariantes(profissionais,"profissional");
  const anivTodos=[...anivPacientes,...anivProfs].sort((a,b)=>a._diff-b._diff);
  // ─── Reuniões próximas para o profissional logado ──────────────────────────
  const profLogadoId=profissionais.find(p=>p.usuario===auth?.usuario)?.id;
  const hoje7=new Date(); hoje7.setDate(hoje7.getDate()+7);
  const hj=ymd(new Date());
  const hj7=ymd(hoje7);
  const reunioesPendentes=(agenda||[]).filter(ag=>
    ag.tipo==="reuniao"&&
    ag.data>=hj&&ag.data<=hj7&&
    (profLogadoId==null||(ag.participantesIds||[]).includes(profLogadoId)||(ag.profissionalId&&Number(ag.profissionalId)===Number(profLogadoId)))
  ).sort((a,b)=>a.data.localeCompare(b.data)||a.horarioSessao.localeCompare(b.horarioSessao));

  const salvarAlerta=()=>{
    if(!form.titulo.trim())return;
    if(editando){setAlertas(a=>a.map(x=>x.id===editando.id?{...x,...form}:x));}
    else{setAlertas(a=>[...a,{id:Date.now(),titulo:form.titulo,descricao:form.descricao,lido:false}]);}
    setEditando(null);setNovoModal(false);setForm({titulo:"",descricao:""});
  };
  const abrirEditar=a=>{setForm({titulo:a.titulo,descricao:a.descricao});setEditando(a);setNovoModal(true);};
  const deletar=id=>setAlertas(a=>a.filter(x=>x.id!==id));
  return(<>
    <div className={"alert-overlay"+(open?" open":"")} onClick={()=>{if(!novoModal)onClose();}}/>
    <div className={"alert-panel"+(open?" open":"")}>
      <div style={{padding:"14px 14px 10px",borderBottom:"1px solid var(--sbd)",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
        <div>
          <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:15}}>🔔 Alertas</div>
          <div style={{fontSize:10,color:"var(--mt)",marginTop:1}}>{pend.length} pendente(s) de {alertas.length}</div>
        </div>
        <div style={{display:"flex",gap:5,alignItems:"center"}}>
          <button className="btn primary small" onClick={()=>{setEditando(null);setForm({titulo:"",descricao:""});setNovoModal(true);}}>+ Novo</button>
          {pend.length>0&&<button className="btn secondary small" onClick={()=>setAlertas(a=>a.map(x=>({...x,lido:true})))}>✓ Todos</button>}
          <button className="icon-btn" onClick={onClose}>×</button>
        </div>
      </div>
      <div style={{flex:1,overflowY:"auto",padding:"10px 12px",display:"flex",flexDirection:"column",gap:7}}>
        {/* ── Aniversariantes ── */}
        {anivTodos.length>0&&<div style={{marginBottom:8}}>
          <div style={{fontSize:10,fontWeight:900,color:"#f59e0b",textTransform:"uppercase",letterSpacing:.6,marginBottom:6,display:"flex",alignItems:"center",gap:6}}>
            🎂 Aniversariantes <span style={{background:"#f59e0b",color:"#000",borderRadius:20,padding:"0 6px",fontSize:9,fontWeight:900}}>{anivTodos.length}</span>
          </div>
          {anivTodos.map(p=>{
            const isHoje=p._diff===0;
            const isAmanha=p._diff===1;
            const cor=isHoje?"#f59e0b":isHoje||isAmanha?"#fb923c":"#94a3b8";
            const bgCor=isHoje?"#f59e0b18":isAmanha?"#fb923c10":"var(--sx)";
            const label=isHoje?"🎉 Hoje!":isAmanha?"Amanhã":"Em "+p._diff+" dias";
            const emoji=p._tipo==="paciente"?"👤":"🩺";
            return(<div key={p.id+"_"+p._tipo} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",borderRadius:9,background:bgCor,border:"1px solid "+(isHoje?"#f59e0b40":isAmanha?"#fb923c30":"var(--sc)"),marginBottom:5}}>
              <span style={{fontSize:18}}>{isHoje?"🎂":"📅"}</span>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontWeight:800,fontSize:12,color:isHoje?"#fbbf24":"var(--tx)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                  {p.nome}
                </div>
                <div style={{fontSize:10,color:"var(--mt)",marginTop:1}}>
                  {emoji} {p._tipo==="paciente"?"Paciente":"Profissional"} · {p._d}/{p._m} · {p._anos} anos
                </div>
              </div>
              <span style={{fontSize:10,fontWeight:900,color:cor,whiteSpace:"nowrap",background:bgCor,padding:"2px 7px",borderRadius:20,border:"1px solid "+cor+"50"}}>{label}</span>
            </div>);
          })}
          <div style={{height:1,background:"var(--sc)",margin:"8px 0"}}/>
        </div>}

        {/* ── Reuniões próximas ── */}
        {reunioesPendentes.length>0&&<div style={{marginBottom:8}}>
          <div style={{fontSize:10,fontWeight:900,color:"#7c6af7",textTransform:"uppercase",letterSpacing:.6,marginBottom:6,display:"flex",alignItems:"center",gap:6}}>
            🤝 Reuniões <span style={{background:"#7c6af7",color:"#fff",borderRadius:20,padding:"0 6px",fontSize:9,fontWeight:900}}>{reunioesPendentes.length}</span>
          </div>
          {reunioesPendentes.map(ag=>{
            const diffD=Math.round((new Date(ag.data+"T12:00:00")-new Date(hj+"T12:00:00"))/(1000*60*60*24));
            const isHoje=diffD===0;
            const isAmanha=diffD===1;
            const cor=isHoje?"#7c6af7":isAmanha?"#a78bfa":"#94a3b8";
            const label=isHoje?"Hoje":isAmanha?"Amanhã":"Em "+diffD+" dias";
            const nPart=(ag.participantesIds||[]).length;
            return(<div key={ag.id} style={{display:"flex",alignItems:"flex-start",gap:8,padding:"7px 10px",borderRadius:9,background:isHoje?"#7c6af715":"var(--sx)",border:"1px solid "+(isHoje?"#7c6af740":"var(--sc)"),marginBottom:5}}>
              <span style={{fontSize:16,marginTop:1}}>🤝</span>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontWeight:800,fontSize:12,color:isHoje?"#a78bfa":"var(--tx)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{ag.tituloReuniao||"Reunião"}</div>
                <div style={{fontSize:10,color:"var(--mt)",marginTop:1}}>{brDate(ag.data)} às {ag.horarioSessao}{ag.horarioFimSessao?" – "+ag.horarioFimSessao:""}{nPart>0?" · "+nPart+" participante(s)":""}</div>
                {ag.pautaReuniao&&<div style={{fontSize:9,color:"var(--mt)",marginTop:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>📋 {ag.pautaReuniao}</div>}
              </div>
              <span style={{fontSize:10,fontWeight:900,color:cor,whiteSpace:"nowrap",padding:"2px 7px",borderRadius:20,border:"1px solid "+cor+"50",background:cor+"15"}}>{label}</span>
            </div>);
          })}
          <div style={{height:1,background:"var(--sc)",margin:"8px 0"}}/>
        </div>}

        {alertas.length===0&&anivTodos.length===0&&reunioesPendentes.length===0&&<div style={{textAlign:"center",color:"var(--mt)",padding:"30px 0",fontSize:12}}>Sem alertas.</div>}
        {alertas.length===0&&anivTodos.length>0&&null}
        {alertas.map(a=>(
          <div key={a.id} style={{padding:"9px 12px",borderRadius:10,background:a.lido?"var(--sx)":"var(--na)",border:"1px solid "+(a.lido?"var(--sc)":"#7c6af740"),transition:".15s"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:6}}>
              <div style={{fontWeight:800,fontSize:12,color:a.lido?"var(--mt)":"var(--tx)",flex:1}}>{a.titulo}</div>
              <div style={{display:"flex",gap:3,flexShrink:0}}>
                {!a.lido&&<span style={{padding:"1px 6px",borderRadius:20,fontSize:9,fontWeight:900,background:"#7c6af7",color:"#fff"}}>NOVO</span>}
                <button onClick={()=>abrirEditar(a)} style={{background:"transparent",border:"none",cursor:"pointer",color:"var(--mt)",fontSize:12,padding:"0 2px"}} title="Editar">✏️</button>
                <button onClick={()=>deletar(a.id)} style={{background:"transparent",border:"none",cursor:"pointer",color:"#f87171",fontSize:12,padding:"0 2px"}} title="Excluir">🗑️</button>
              </div>
            </div>
            <div style={{fontSize:11,color:"var(--mt)",marginTop:3,lineHeight:1.4}}>{a.descricao}</div>
            {!a.lido&&<button onClick={()=>setAlertas(arr=>arr.map(x=>x.id===a.id?{...x,lido:true}:x))} style={{marginTop:5,fontSize:10,color:"#a78bfa",background:"transparent",border:"none",cursor:"pointer",fontWeight:700,padding:0}}>Marcar como lido</button>}
          </div>
        ))}
      </div>
    </div>
    {novoModal&&<div className="modal-bg" style={{zIndex:70}} onClick={e=>e.target===e.currentTarget&&(setNovoModal(false),setEditando(null))}>
      <div className="modal" style={{width:420}}>
        <div className="modal-head"><h2>{editando?"✏️ Editar":"+ Novo"} Alerta</h2><button className="icon-btn" onClick={()=>{setNovoModal(false);setEditando(null);}}>×</button></div>
        <div className="stack">
          <div><label>Título *</label><input value={form.titulo} onChange={e=>setForm(f=>({...f,titulo:e.target.value}))} autoFocus placeholder="Título do alerta"/></div>
          <div><label>Descrição</label><textarea rows={3} value={form.descricao} onChange={e=>setForm(f=>({...f,descricao:e.target.value}))} placeholder="Detalhes..."/></div>
          <div className="actions">
            <button className="btn secondary" onClick={()=>{setNovoModal(false);setEditando(null);}}>Cancelar</button>
            <button className="btn primary" onClick={salvarAlerta}>{editando?"💾 Salvar":"+ Criar Alerta"}</button>
          </div>
        </div>
      </div>
    </div>}
  </>);
}

function UploadField({label,value,onChange}){
  const id="upl"+(label||"").replace(/\W/g,"");
  return(<div><label>{label||"Arquivo"}</label>
    <label htmlFor={id} className="upload-zone" style={{display:"block"}}>
      {value&&value.length?<span style={{color:"#a78bfa"}}>✅ {value.length} arquivo{value.length>1?"s":""}</span>:<span>📎 Clique para selecionar</span>}
    </label>
    <input id={id} type="file" multiple style={{display:"none"}} onChange={e=>{const f=Array.from(e.target.files);onChange&&onChange(f.map(x=>({nome:x.name,tamanho:x.size,url:URL.createObjectURL(x)})));}}/>
    {value&&value.length>0&&<div style={{marginTop:4}}>{value.map((f,i)=><span key={i} className="file-chip">📄 {f.nome}</span>)}</div>}
  </div>);
}

function CampoEndereco({form,setForm}){
  const [loadingCep,setLoadingCep]=useState(false);
  const handleCep=async v=>{
    const c=rawD(v).slice(0,8);
    setForm(f=>({...f,cep:c}));
    if(c.length===8){setLoadingCep(true);const r=await buscaCEP(c);setLoadingCep(false);if(r)setForm(f=>({...f,...r}));}
  };
  return(<>
    <div><label>CEP *</label><input value={maskCEP(form.cep||"")} onChange={e=>handleCep(e.target.value)} placeholder="00000-000" maxLength={9}/>{loadingCep&&<div className="helper">🔍 Buscando...</div>}</div>
    <div className="g2">
      <div><label>Logradouro *</label><input value={form.logradouro||""} onChange={e=>setForm(f=>({...f,logradouro:e.target.value}))}/></div>
      <div><label>Número *</label><input value={form.numero||""} onChange={e=>setForm(f=>({...f,numero:e.target.value}))}/></div>
    </div>
    <div className="g2">
      <div><label>Complemento</label><input value={form.complemento||""} onChange={e=>setForm(f=>({...f,complemento:e.target.value}))}/></div>
      <div><label>Bairro *</label><input value={form.bairro||""} onChange={e=>setForm(f=>({...f,bairro:e.target.value}))}/></div>
    </div>
    <div className="g2">
      <div><label>Cidade *</label><input value={form.cidade||""} onChange={e=>setForm(f=>({...f,cidade:e.target.value}))}/></div>
      <div><label>Estado</label><select value={form.estado||""} onChange={e=>setForm(f=>({...f,estado:e.target.value}))}><option value="">UF</option>{UFS.map(u=><option key={u}>{u}</option>)}</select></div>
    </div>
  </>);
}

function LoginModal({profissionais,onClose,onLogin}){
  const [cpf,setCpf]=useState("");
  const [senha,setSenha]=useState("");
  const [err,setErr]=useState("");
  const tentar=()=>{
    const c=rawD(cpf).slice(0,11);
    const p=profissionais.find(x=>x.usuario===c&&x.senha===senha);
    if(!p){setErr("CPF ou senha inválidos");return;}
    onLogin(p);
  };
  return(<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{width:380,marginTop:80}}>
      <div className="modal-head"><h2>🔐 Acesso ao Sistema</h2><button className="icon-btn" onClick={onClose}>×</button></div>
      <div className="stack">
        <div><label>CPF (somente números)</label>
          <input value={maskCPF(cpf)} onChange={e=>setCpf(rawD(e.target.value).slice(0,11))} placeholder="000.000.000-00" maxLength={14} autoFocus onKeyDown={e=>e.key==="Enter"&&tentar()}/>
        </div>
        <div><label>Senha</label><input type="password" value={senha} onChange={e=>setSenha(e.target.value)} onKeyDown={e=>e.key==="Enter"&&tentar()}/></div>
        {err&&<div style={{color:"#f87171",fontSize:12}}>{err}</div>}
        <div className="actions"><button className="btn primary" style={{width:"100%"}} onClick={tentar}>Entrar</button></div>
        <div className="helper" style={{textAlign:"center"}}>Admin: 31028313896 / 122 · Fat: 55555555555 / 5555</div>
      </div>
    </div>
  </div>);
}

function ChamadoButton({setChamados,showToast}){
  const [open,setOpen]=useState(false);
  const [form,setForm]=useState({setor:"Agendamento",tipo:"manual",padraoId:"01",camposValues:{},assunto:"",descricao:""});
  const snf=p=>setForm(f=>({...f,...p}));
  const isPadrao=form.tipo==="padrao";
  const padraoSel=PADRAO_CHAMADO.find(p=>p.id===form.padraoId)||PADRAO_CHAMADO[0];
  const buildDesc=()=>isPadrao?padraoSel.campos.map(c=>c+": "+(form.camposValues[c]||"")).join("\n"):form.descricao;
  const salvar=()=>{
    if(!isPadrao&&!form.assunto.trim())return alert("Assunto obrigatório");
    const num=String(Math.floor(100000+Math.random()*900000));
    const label=isPadrao?(padraoSel.label):(form.assunto);
    setChamados(c=>[...c,{id:Date.now(),numero:num,setor:form.setor,tipo:isPadrao?"padrao_"+form.padraoId:"manual",nome:label,descricao:buildDesc(),data:hoje_str,status:"aberto",resp:""}]);
    showToast("📨 Chamado #"+num+" aberto","ok");
    setOpen(false);setForm({setor:"Agendamento",tipo:"manual",padraoId:"01",camposValues:{},assunto:"",descricao:""});
  };
  return(<>
    <button className="fab-chamado" onClick={()=>setOpen(true)} title="Abrir chamado rápido">📨</button>
    {open&&<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&setOpen(false)}>
      <div className="modal" style={{width:500}}>
        <div className="modal-head">
          <div>
            <h2 style={{marginBottom:2}}>📨 Chamado Rápido</h2>
            <div style={{fontSize:11,color:"var(--mt)",fontWeight:400}}>Abre um chamado sem sair da tela atual</div>
          </div>
          <button className="icon-btn" onClick={()=>setOpen(false)}>×</button>
        </div>
        <div className="stack">
          <div className="g2">
            <div><label>Setor</label><select value={form.setor} onChange={e=>snf({setor:e.target.value})}>{SETORES_CHAMADO.map(s=><option key={s}>{s}</option>)}</select></div>
            <div><label>Tipo</label>
              <select value={form.tipo} onChange={e=>snf({tipo:e.target.value,camposValues:{}})}>
                <option value="manual">✏️ Texto livre</option>
                {form.setor==="Agendamento"&&<option value="padrao">📋 Padronizado (Anexo VII)</option>}
              </select>
            </div>
          </div>
          {isPadrao&&<div><label>Modelo</label>
            <select value={form.padraoId} onChange={e=>snf({padraoId:e.target.value,camposValues:{}})}>
              {PADRAO_CHAMADO.map(p=><option key={p.id} value={p.id}>{p.label}</option>)}
            </select>
          </div>}
          {isPadrao&&<div className="section-box" style={{background:"#7c6af708",borderColor:"#7c6af725"}}>
            <div className="section-title" style={{color:"#a78bfa"}}>📋 {padraoSel.label}</div>
            <div style={{display:"flex",flexDirection:"column",gap:6}}>
              {padraoSel.campos.map(campo=>(
                <div key={campo} style={{display:"grid",gridTemplateColumns:"140px 1fr",alignItems:"center",gap:7}}>
                  <label style={{textTransform:"none",fontSize:11,fontWeight:700,color:"var(--mt)",margin:0}}>{campo}:</label>
                  <input value={form.camposValues[campo]||""} onChange={e=>snf({camposValues:{...form.camposValues,[campo]:e.target.value}})} style={{fontSize:11,padding:"5px 8px"}}/>
                </div>
              ))}
            </div>
          </div>}
          {!isPadrao&&<>
            <div><label>Assunto *</label><input value={form.assunto} onChange={e=>snf({assunto:e.target.value})} autoFocus placeholder="Resumo do chamado..."/></div>
            <div><label>Descrição</label><textarea rows={3} value={form.descricao} onChange={e=>snf({descricao:e.target.value})} placeholder="Detalhes adicionais..."/></div>
          </>}
          <div className="actions">
            <button className="btn secondary" onClick={()=>setOpen(false)}>Cancelar</button>
            <button className="btn primary" onClick={salvar}>📨 Enviar Chamado</button>
          </div>
        </div>
      </div>
    </div>}
  </>);
}

// ═══════════════════════════════════════════════════════════════════════════════
// AGENDA MODAL — com repetições, filial, convênio editável, botão manual
// ═══════════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════════
// SESSÃO MODAL — abre ao clicar num agendamento existente, atualiza status
// ═══════════════════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════════════════
// REUNIÃO ATA MODAL
// ═══════════════════════════════════════════════════════════════════════════════
function ReuniaoAtaModal({ag,profissionais,chamados,setChamados,showToast,auth,onClose,onEditar}){
  const participantes=(ag.participantesIds||[]).map(id=>profissionais.find(p=>p.id===id)).filter(Boolean);
  const [ata,setAta]=useState(ag.ata||"");
  const [encerrada,setEncerrada]=useState(ag.ataEncerrada||false);
  const [analisando,setAnalisando]=useState(false);
  const [analise,setAnalise]=useState(ag.atAnalise||null);
  const [chamadosSugeridos,setChamadosSugeridos]=useState([]);
  const [chamadosSel,setChamadosSel]=useState([]);
  const [salvando,setSalvando]=useState(false);

  // Salva ata no agendamento pai via callback
  const salvarAta=(novoEncerramento)=>{
    setSalvando(true);
    onEditar({...ag,ata,ataEncerrada:novoEncerramento!==undefined?novoEncerramento:encerrada,atAnalise:analise});
    setTimeout(()=>{setSalvando(false);showToast("✅ Ata salva","ok");},300);
  };

  const analisarComIA=async()=>{
    if(!ata.trim())return alert("Escreva o conteúdo da ata antes de analisar.");
    setAnalisando(true);
    setAnalise(null);
    setChamadosSugeridos([]);
    try{
      const partNomes=participantes.map(p=>p.nome).join(", ")||"Não informados";
      const prompt="Você é um assistente de gestão clínica especializado em clínicas de saúde. Analise a ata de reunião abaixo e retorne APENAS JSON, sem markdown.\n\n"+"ATA DA REUNIÃO:\n"+"- Título: "+(ag.tituloReuniao||"Reunião")+"\n"+"- Data: "+brDate(ag.data)+" às "+ag.horarioSessao+"\n"+"- Participantes: "+partNomes+"\n"+"- Pauta original: "+(ag.pautaReuniao||"Não definida")+"\n"+"- Conteúdo da ata: \"\"\""+ata+"\"\"\"\n\n"+"Retorne SOMENTE este JSON:\n"+'{"resumo":"resumo em 2-3 linhas","decisoes":["lista de decisões"],"pendencias":["lista de pendências"],"qualidade":"ok ou alerta ou incompleta","observacoes":"observações sobre qualidade","chamadosSugeridos":[{"titulo":"título","setor":"Agendamento ou Recepção ou Faturamento ou Coordenação ou TI ou Gestão de Pessoas ou Administração","descricao":"descrição","prioridade":"Urgente ou Alta ou Normal ou Baixa","justificativa":"justificativa"}]}'
      const resp=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1500,messages:[{role:"user",content:prompt}]})
      });
      const data=await resp.json();
      const txt=(data.content||[]).filter(b=>b.type==="text").map(b=>b.text).join("");
      let parsed;
      try{parsed=JSON.parse(txt.replace(/```json|```/g,"").trim());}
      catch{parsed={resumo:"Não foi possível analisar.",decisoes:[],pendencias:[],qualidade:"alerta",observacoes:"Erro ao processar resposta da IA.",chamadosSugeridos:[]};}
      setAnalise(parsed);
      setChamadosSugeridos(parsed.chamadosSugeridos||[]);
      setChamadosSel((parsed.chamadosSugeridos||[]).map((_,i)=>i));
    }catch(e){
      showToast("❌ Erro IA: "+e.message,"err");
    }
    setAnalisando(false);
  };

  const abrirChamadosSelecionados=()=>{
    const selecionados=chamadosSugeridos.filter((_,i)=>chamadosSel.includes(i));
    if(!selecionados.length)return alert("Selecione ao menos um chamado.");
    selecionados.forEach(ch=>{
      const num=String(Math.floor(100000+Math.random()*900000));
      setChamados(a=>[...a,{
        id:Date.now()+Math.random(),numero:num,
        setor:ch.setor,tipo:"reuniao",
        nome:ch.titulo,
        descricao:ch.descricao+" [Originado da reunião: "+ag.tituloReuniao+" — "+brDate(ag.data)+"]",
        data:hoje_str,status:"aberto",prioridade:ch.prioridade||"Normal",resp:""
      }]);
    });
    showToast("📨 "+selecionados.length+" chamado(s) aberto(s)","ok");
    salvarAta();
  };

  const toggleChamado=(i)=>setChamadosSel(s=>s.includes(i)?s.filter(x=>x!==i):[...s,i]);

  const corQualidade=analise?.qualidade==="ok"?"#22c55e":analise?.qualidade==="alerta"?"#f59e0b":"#ef4444";
  const icQualidade=analise?.qualidade==="ok"?"✅":analise?.qualidade==="alerta"?"⚠️":"❌";

  return(<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{width:900,maxHeight:"95vh",overflowY:"auto"}}>

      {/* Header */}
      <div className="modal-head" style={{background:"linear-gradient(135deg,#1e3a5f,#1e4d3a)",borderRadius:"12px 12px 0 0",margin:"-20px -20px 0",padding:"16px 20px"}}>
        <div style={{flex:1}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <span style={{fontSize:22}}>🤝</span>
            <div>
              <h2 style={{margin:0,fontSize:16,color:"#fff"}}>{ag.tituloReuniao||"Reunião"}</h2>
              <div style={{fontSize:11,color:"#94a3b8",marginTop:2}}>{brDate(ag.data)} às {ag.horarioSessao} · {participantes.length} participante(s)</div>
            </div>
          </div>
        </div>
        <div style={{display:"flex",gap:6,alignItems:"center"}}>
          {encerrada&&<span style={{padding:"3px 10px",borderRadius:20,background:"#22c55e20",color:"#22c55e",fontSize:11,fontWeight:800}}>✅ Encerrada</span>}
          <button className="icon-btn" onClick={onClose} style={{color:"#94a3b8"}}>×</button>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 340px",gap:16,marginTop:16}}>

        {/* ── COLUNA ESQUERDA: ATA ── */}
        <div className="stack">
          {/* Info participantes */}
          {participantes.length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:5}}>
            {participantes.map(p=><span key={p.id} style={{padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,background:"#7c6af715",color:"#a78bfa",border:"1px solid #7c6af730"}}>{p.nome.split(" ")[0]}</span>)}
          </div>}
          {ag.pautaReuniao&&<div style={{background:"#f59e0b08",border:"1px solid #f59e0b20",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#fcd34d"}}>
            <span style={{fontWeight:800,marginRight:6}}>📋 Pauta:</span>{ag.pautaReuniao}
          </div>}

          {/* Editor de ata */}
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
              <label style={{margin:0,fontWeight:800,fontSize:13}}>📝 Ata da Reunião</label>
              <div style={{display:"flex",alignItems:"center",gap:6}}>
                {!encerrada&&<MicBtn onResult={t=>setAta(a=>a+(a?" ":"")+t)} title="Gravar áudio para transcrição da ata"/>}
                {encerrada&&<span style={{fontSize:10,color:"#ef4444",fontWeight:700}}>🔒 Encerrada — somente leitura</span>}
              </div>
            </div>
            {!encerrada
              ?<textarea rows={14} value={ata} onChange={e=>setAta(e.target.value)}
                  placeholder={"Registre aqui os tópicos discutidos, decisões tomadas, encaminhamentos e responsáveis...\n\n• Abertura da reunião\n• Pontos discutidos\n• Decisões tomadas\n• Próximos passos"}
                  style={{fontFamily:"inherit",lineHeight:1.7,fontSize:13}}/>
              :<div style={{background:"var(--sx)",border:"1px solid var(--cb)",borderRadius:10,padding:14,fontSize:13,lineHeight:1.7,minHeight:200,whiteSpace:"pre-wrap",color:"var(--tx)"}}>
                {ata||<span style={{color:"var(--mt)",fontStyle:"italic"}}>Sem ata registrada.</span>}
              </div>}
          </div>

          {/* Ações da ata */}
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {!encerrada&&<button className="btn secondary" onClick={()=>salvarAta()} disabled={salvando}>
              {salvando?"⏳ Salvando...":"💾 Salvar Ata"}
            </button>}
            {!encerrada&&<button className="btn primary" onClick={analisarComIA} disabled={analisando}>
              {analisando?"⏳ Analisando...":"🤖 Analisar com IA"}
            </button>}
            {!encerrada&&analise&&<button onClick={()=>{if(confirm("Encerrar a ata? Ela ficará somente leitura.")){setEncerrada(true);salvarAta(true);}}}
              style={{padding:"7px 14px",borderRadius:8,border:"1.5px solid #22c55e50",background:"#22c55e10",color:"#22c55e",fontWeight:800,cursor:"pointer",fontSize:12}}>
              ✅ Encerrar Ata
            </button>}
            {encerrada&&<button onClick={()=>{if(confirm("Reabrir ata para edição?")){setEncerrada(false);}}}
              style={{padding:"7px 14px",borderRadius:8,border:"1.5px solid #f59e0b50",background:"#f59e0b10",color:"#f59e0b",fontWeight:800,cursor:"pointer",fontSize:12}}>
              🔓 Reabrir Ata
            </button>}
          </div>
        </div>

        {/* ── COLUNA DIREITA: Análise IA ── */}
        <div className="stack" style={{minWidth:0}}>
          <div style={{background:"var(--sx)",borderRadius:10,padding:12,border:"1px solid var(--cb)"}}>
            <div style={{fontWeight:900,fontSize:12,color:"var(--mt)",textTransform:"uppercase",letterSpacing:.5,marginBottom:8}}>🤖 Análise por IA</div>
            {!analise&&!analisando&&<div style={{color:"var(--mt)",fontSize:12,textAlign:"center",padding:"20px 0"}}>
              Escreva a ata e clique em<br/><b style={{color:"#a78bfa"}}>🤖 Analisar com IA</b><br/>para obter sugestões de chamados.
            </div>}
            {analisando&&<div style={{textAlign:"center",padding:"30px 0",color:"#a78bfa",fontSize:13}}>
              ⏳ Processando ata com IA...
            </div>}
            {analise&&!analisando&&<>
              {/* Qualidade */}
              <div style={{background:corQualidade+"12",border:"1px solid "+corQualidade+"30",borderRadius:8,padding:"8px 10px",marginBottom:10}}>
                <div style={{fontWeight:800,fontSize:12,color:corQualidade,marginBottom:4}}>{icQualidade} {analise.qualidade==="ok"?"Ata completa":analise.qualidade==="alerta"?"Verificar ata":"Ata incompleta"}</div>
                <div style={{fontSize:11,color:"var(--mt)"}}>{analise.observacoes}</div>
              </div>

              {/* Resumo */}
              {analise.resumo&&<div style={{marginBottom:10}}>
                <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",letterSpacing:.5,marginBottom:4}}>Resumo</div>
                <div style={{fontSize:12,color:"var(--tx)",lineHeight:1.5}}>{analise.resumo}</div>
              </div>}

              {/* Decisões */}
              {analise.decisoes?.length>0&&<div style={{marginBottom:10}}>
                <div style={{fontSize:10,fontWeight:800,color:"#22c55e",textTransform:"uppercase",letterSpacing:.5,marginBottom:4}}>✅ Decisões</div>
                {analise.decisoes.map((d,i)=><div key={i} style={{fontSize:11,color:"var(--tx)",padding:"3px 0",borderBottom:"1px solid var(--db)"}}>• {d}</div>)}
              </div>}

              {/* Pendências */}
              {analise.pendencias?.length>0&&<div style={{marginBottom:10}}>
                <div style={{fontSize:10,fontWeight:800,color:"#f59e0b",textTransform:"uppercase",letterSpacing:.5,marginBottom:4}}>⏳ Pendências</div>
                {analise.pendencias.map((p,i)=><div key={i} style={{fontSize:11,color:"var(--tx)",padding:"3px 0",borderBottom:"1px solid var(--db)"}}>• {p}</div>)}
              </div>}

              {/* Chamados sugeridos */}
              {chamadosSugeridos.length>0&&<div>
                <div style={{fontSize:10,fontWeight:800,color:"#a78bfa",textTransform:"uppercase",letterSpacing:.5,marginBottom:6}}>📨 Chamados sugeridos ({chamadosSugeridos.length})</div>
                {chamadosSugeridos.map((ch,i)=>{
                  const sel=chamadosSel.includes(i);
                  const corPrio=ch.prioridade==="Urgente"?"#ef4444":ch.prioridade==="Alta"?"#f59e0b":ch.prioridade==="Normal"?"#a78bfa":"#94a3b8";
                  return(<div key={i} onClick={()=>toggleChamado(i)}
                    style={{marginBottom:6,padding:"8px 10px",borderRadius:8,border:"1.5px solid "+(sel?"#a78bfa":"var(--cpb)"),
                      background:sel?"#a78bfa12":"transparent",cursor:"pointer",transition:".15s"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:6}}>
                      <div style={{flex:1}}>
                        <div style={{fontWeight:800,fontSize:11,color:sel?"#c4b5fd":"var(--tx)"}}>{sel?"☑ ":"☐ "}{ch.titulo}</div>
                        <div style={{fontSize:10,color:"var(--mt)",marginTop:2}}>{ch.setor} · <span style={{color:corPrio,fontWeight:700}}>{ch.prioridade}</span></div>
                        <div style={{fontSize:10,color:"var(--mt)",marginTop:2,fontStyle:"italic"}}>{ch.justificativa}</div>
                      </div>
                    </div>
                  </div>);
                })}
                <div style={{display:"flex",gap:6,marginTop:8}}>
                  <button onClick={()=>setChamadosSel(chamadosSugeridos.map((_,i)=>i))} style={{flex:1,padding:"5px",borderRadius:7,border:"1px solid var(--cpb)",background:"transparent",color:"var(--mt)",fontSize:10,cursor:"pointer"}}>✓ Todos</button>
                  <button onClick={()=>setChamadosSel([])} style={{flex:1,padding:"5px",borderRadius:7,border:"1px solid var(--cpb)",background:"transparent",color:"var(--mt)",fontSize:10,cursor:"pointer"}}>✗ Nenhum</button>
                </div>
                {chamadosSel.length>0&&<button className="btn primary" style={{width:"100%",marginTop:8,fontSize:12}}
                  onClick={abrirChamadosSelecionados}>
                  📨 Abrir {chamadosSel.length} chamado(s)
                </button>}
              </div>}
            </>}
          </div>
        </div>
      </div>
    </div>
  </div>);
}

function SessaoModal({ag,pacientes,profissionais,procedimentos,salas,filiais,convenios,auth,agenda,prontuarios,setProntuarios,modelosEv,caixa,setCaixa,onClose,onSave,onEditar}){
  const [status,setStatus]=useState(ag.status||"agendado");
  const [numAut,setNumAut]=useState(ag.numAutorizacao||"");
  const [confirmacaoEnviada,setConfirmacaoEnviada]=useState(ag.confirmacaoEnviada||false);
  const [evModal,setEvModal]=useState(false);

  const pac=pacientes.find(p=>p.id===Number(ag.pacienteId));
  const prof=profissionais.find(p=>p.id===Number(ag.profissionalId));
  const sala=salas.find(s=>s.id===Number(ag.salaId));
  const filial=filiais.find(f=>f.id===Number(ag.filialId)||(sala&&f.id===sala.filialId));
  const proc=procedimentos.find(p=>p.id===Number(ag.procedimentoId));
  const isParticular=(ag.convenio||"Particular")==="Particular";
  const jaRecebido=(caixa||[]).some(e=>e.agId===ag.id);
  const [showCaixa,setShowCaixa]=useState(false);
  const [formaPag,setFormaPag]=useState("PIX");
  const [valorCaixa,setValorCaixa]=useState(proc?.valor||0);
  const stInfo=STATUS_AG[status]||{color:"#64748b",icon:"📋",label:status};
  const podeStatus=PODE_STATUS(auth.role);
  const podeFaturar=PODE_FATURAR(auth.role);
  const podeEvolucao=PODE_EVOLUIR(auth.role);
  const precisaAut=STATUS_AG[status]?.numAut&&!numAut;
  const profLogado=profissionais.find(p=>p.usuario===auth.usuario);
  const ehProfDaSessao=profLogado&&Number(profLogado.id)===Number(ag.profissionalId);
  const jaEvoluido=(prontuarios||[]).some(p=>p.agId===ag.id);

  // Outras sessões do mesmo paciente no mesmo dia
  const outrasSessoes=(agenda||[]).filter(a=>a.id!==ag.id&&Number(a.pacienteId)===Number(ag.pacienteId)&&a.data===ag.data);

  const salvar=(novoStatus,novoConf)=>{
    const s=novoStatus||status;
    if(STATUS_AG[s]?.numAut&&!numAut){alert("Preencha o número de autorização do convênio");return;}
    onSave({...ag,status:s,numAutorizacao:numAut,confirmacaoEnviada:novoConf!==undefined?novoConf:confirmacaoEnviada});
    onClose();
  };

  // WhatsApp confirmação
  const NOME_CLINICA="Espaço Terapêutico Cinthia França";
  const buildWAMsg=()=>{
    const endFilial=filial?filial.nome+" — "+[filial.logradouro,filial.numero,filial.bairro,filial.cidade].filter(Boolean).join(", "):"";
    let msg="*CONFIRMAÇÃO DE ATENDIMENTO*\n";
    msg+="*"+NOME_CLINICA+"*\n\n";
    msg+="👤 *Paciente:* "+(pac?.nome||"—")+"\n";
    msg+="📅 *Data:* "+brDate(ag.data)+"\n";
    msg+="🕐 *Horário:* "+ag.horarioSessao+(ag.horarioFimSessao?" às "+ag.horarioFimSessao:"")+"\n";
    msg+="🩺 *Especialidade:* "+(prof?.especialidades||[prof?.especialidade])[0]+"\n";
    msg+="👩‍⚕️ *Profissional:* "+(prof?.nome||"—")+"\n";
    if(outrasSessoes.length>0){
      outrasSessoes.forEach(s=>{
        const sp=profissionais.find(p=>p.id===Number(s.profissionalId));
        const sf=filiais.find(f=>f.id===Number(s.filialId));
        msg+="\n━━━━━━━━━━━━━━━━\n";
        msg+="🕐 *Horário:* "+s.horarioSessao+(s.horarioFimSessao?" às "+s.horarioFimSessao:"")+"\n";
        msg+="🩺 *Especialidade:* "+(sp?.especialidades||[sp?.especialidade])[0]+"\n";
        if(sf)msg+="📍 "+sf.nome+"\n";
      });
    }
    msg+="\n━━━━━━━━━━━━━━━━\n";
    msg+="📍 "+endFilial+"\n";
    msg+="\n_Por favor confirme sua presença._";
    return msg;
  };


  useEffect(()=>{setValorCaixa(proc?.valor||0);},[proc?.id]);

  const registrarCaixa=()=>{
    if(!setCaixa)return;
    const entrada={
      id:Date.now(),agId:ag.id,data:ag.data,hora:ag.horarioSessao,
      pacienteId:ag.pacienteId,profissionalId:ag.profissionalId,procedimentoId:ag.procedimentoId,
      valor:Number(valorCaixa),formaPagamento:formaPag,
      descr:(proc?.nome||"Sessão")+" — "+(pac?.nome||"")+" — "+brDate(ag.data),
      tipo:"entrada",origem:"particular",
    };
    setCaixa(c=>[...c,entrada]);
    onSave({...ag,caixaRecebido:true,formaPagamento:formaPag,valorRecebido:Number(valorCaixa),numAutorizacao:numAut,confirmacaoEnviada});
    setShowCaixa(false);
  };

  const waNum=pac?.resp1Whatsapp||pac?.celular||"";
  const waConfUrl=waNum?"https://wa.me/55"+rawD(waNum)+"?text="+encodeURIComponent(buildWAMsg()):"";

  return(<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{width:640,maxHeight:"94vh",overflowY:"auto"}}>
      <div className="modal-head">
        <h2 style={{fontSize:16}}>📋 Sessão — <span style={{color:espCor((prof?.especialidades||[""])[0])}}>{pac?.nome||"—"}</span></h2>
        <div style={{display:"flex",gap:6,alignItems:"center"}}>
          {/* ✅ Confirmação enviada — ícone puro */}
          <div style={{position:"relative",flexShrink:0}}>
            <button onClick={()=>{const newVal=!confirmacaoEnviada;setConfirmacaoEnviada(newVal);onSave({...ag,status,confirmacaoEnviada:newVal,numAutorizacao:numAut});}}
            title={confirmacaoEnviada?"Confirmação enviada — clique para desfazer":"Confirmar: presença confirmada"}
            style={{background:confirmacaoEnviada?"#38bdf820":"transparent",border:"2px solid "+(confirmacaoEnviada?"#38bdf8":"var(--cpb)"),borderRadius:8,width:34,height:34,cursor:"pointer",fontSize:17,color:confirmacaoEnviada?"#38bdf8":"var(--mt)",display:"flex",alignItems:"center",justifyContent:"center",transition:".15s"}}>
              📨
            </button>
            {confirmacaoEnviada&&<span style={{position:"absolute",top:-4,right:-4,width:10,height:10,background:"#38bdf8",borderRadius:"50%",border:"2px solid var(--bg)",pointerEvents:"none"}}/>}
          </div>
          {/* 📝 Evolução — somente profissional da sessão */}
          {(podeEvolucao&&(ehProfDaSessao||auth.role==="administrador"||auth.role==="coordenador"))&&(
            <div style={{position:"relative",flexShrink:0}}>
              <button onClick={()=>setEvModal(true)}
              title={jaEvoluido?"Sessão evoluída — clique para adicionar":"Registrar evolução desta sessão"}
              style={{background:jaEvoluido?"#90ee9020":"transparent",border:"2px solid "+(jaEvoluido?"#90ee90":"var(--cpb)"),borderRadius:8,width:34,height:34,cursor:"pointer",fontSize:17,color:jaEvoluido?"#90ee90":"var(--mt)",display:"flex",alignItems:"center",justifyContent:"center",transition:".15s"}}>
                📝
              </button>
              {jaEvoluido&&<span style={{position:"absolute",top:-4,right:-4,width:10,height:10,background:"#90ee90",borderRadius:"50%",border:"2px solid var(--bg)",pointerEvents:"none"}}/>}
            </div>
          )}
          {/* ── Ação financeira: Faturado — somente para convênio ── */}
          {podeFaturar&&!isParticular&&(<>
            <div style={{width:1,height:28,background:"var(--cpb)",flexShrink:0,alignSelf:"center",margin:"0 2px"}}/>
            <div style={{position:"relative",flexShrink:0}}>
              <button onClick={()=>{onSave({...ag,fatStatus:ag.fatStatus==="faturado"?null:"faturado",numAutorizacao:numAut,confirmacaoEnviada});}}
              title={ag.fatStatus==="faturado"?"💵 Faturado — clique para desfazer":"Marcar como faturado — gera crédito do convênio"}
              style={{background:ag.fatStatus==="faturado"?"#22c55e20":"transparent",border:"2px solid "+(ag.fatStatus==="faturado"?"#22c55e":"var(--cpb)"),borderRadius:8,width:34,height:34,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",transition:".15s",flexDirection:"column",gap:0}}>
                <span style={{fontSize:13,lineHeight:1}}>💵</span>
                <span style={{fontSize:7,fontWeight:900,color:ag.fatStatus==="faturado"?"#22c55e":"var(--mt)",lineHeight:1,marginTop:1}}>FAT.</span>
              </button>
              {ag.fatStatus==="faturado"&&<span style={{position:"absolute",top:-4,right:-4,width:10,height:10,background:"#22c55e",borderRadius:"50%",border:"2px solid var(--bg)",pointerEvents:"none"}}/>}
            </div>
          </>)}
          {/* ✏️ Editar agendamento */}
          {(auth.role==="administrador"||auth.role==="agendamento"||auth.role==="secretaria"||auth.role==="atendimento"||auth.role==="supervisor_adm")&&onEditar&&(
            <button onClick={()=>{onClose();onEditar(ag);}}
              title="Editar dados do agendamento"
              style={{background:"transparent",border:"2px solid var(--cpb)",borderRadius:8,width:34,height:34,cursor:"pointer",fontSize:16,color:"var(--mt)",display:"flex",alignItems:"center",justifyContent:"center",transition:".15s"}}>
              ✏️
            </button>
          )}
          <button className="icon-btn" onClick={onClose}>×</button>
        </div>
      </div>
      <div className="stack">
        {/* Indicadores no topo */}
        <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
          <span style={{padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:800,background:stInfo.color+"25",color:stInfo.color,border:"1px solid "+stInfo.color+"50"}}>{stInfo.icon} {stInfo.label}</span>
          {confirmacaoEnviada&&<span title="Confirmação de presença enviada" style={{padding:"3px 8px",borderRadius:20,fontSize:11,fontWeight:800,background:"#38bdf820",color:"#38bdf8",border:"1px solid #38bdf840"}}>📨</span>}
          {jaEvoluido&&<span style={{padding:"3px 8px",borderRadius:20,fontSize:11,fontWeight:800,background:"#90ee9025",color:"#90ee90",border:"1px solid #90ee9040"}}>📝</span>}
          {ag.fatStatus==="faturado"&&<span style={{padding:"3px 8px",borderRadius:20,fontSize:11,fontWeight:800,background:"#22c55e25",color:"#22c55e",border:"1px solid #22c55e40"}}>💵 Faturado</span>}
        </div>

        {/* Info */}
        <div className="section-box">
          <div className="section-title">📅 Informações</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,fontSize:12}}>
            <div><span className="muted" style={{fontSize:10}}>DATA</span><br/><b>{brDate(ag.data)}</b></div>
            <div><span className="muted" style={{fontSize:10}}>HORÁRIO</span><br/><b>{ag.horarioSessao} → {ag.horarioFimSessao}</b></div>
            <div><span className="muted" style={{fontSize:10}}>PROFISSIONAL</span><br/><b style={{color:espCor((prof?.especialidades||[""])[0])}}>{prof?.nome||"—"}</b></div>
            <div><span className="muted" style={{fontSize:10}}>ESPECIALIDADE</span><br/><b>{(prof?.especialidades||[""])[0]}</b></div>
            <div><span className="muted" style={{fontSize:10}}>SALA / FILIAL</span><br/><b>{sala?.nome||"—"} {filial?"· "+filial.nome:""}</b></div>
            <div><span className="muted" style={{fontSize:10}}>PROCEDIMENTO</span><br/><b>{proc?.nome||"—"}</b></div>
            {ag.convenio&&<div><span className="muted" style={{fontSize:10}}>CONVÊNIO</span><br/><b>{ag.convenio}{ag.plano?" · "+ag.plano:""}</b></div>}
            {ag.numAutorizacao&&<div><span className="muted" style={{fontSize:10}}>Nº AUTORIZAÇÃO</span><br/><b style={{color:"#34d399"}}>{ag.numAutorizacao}</b></div>}
          </div>
          {pac?.infoImportantes&&<div style={{marginTop:8,fontSize:11,color:"#f59e0b",background:"#451a0320",padding:"5px 9px",borderRadius:6}}>⚠️ {pac.infoImportantes}</div>}
        </div>

        {/* Status */}
        <div className="section-box">
          <div className="section-title">📊 Status da Sessão</div>
          {!podeStatus&&<div className="helper" style={{marginBottom:6}}>⚠️ Somente Atendimento / Supervisor / Admin / Agendamento</div>}
          <div className="chips">
            {Object.entries(STATUS_AG).filter(([k])=>!["faturado","cancelado"].includes(k)).map(([k,v])=>{
              const active=status===k;
              const ok=podeStatus;
              return(<button key={k} className="chip-btn" disabled={!ok} onClick={()=>ok&&setStatus(k)} style={{background:active?v.color+"30":"transparent",color:active?v.color:"var(--mt)",borderColor:active?v.color:"var(--cpb)",opacity:ok?1:.35,fontSize:10}}>{v.icon} {v.label}</button>);
            })}
          </div>
          {STATUS_AG[status]?.numAut&&<div style={{marginTop:8}}>
            <label>Nº Autorização do Convênio *</label>
            <input value={numAut} onChange={e=>setNumAut(e.target.value)} placeholder="Informe o número de autorização"/>
          </div>}
        </div>

        {/* WhatsApp */}
        {waNum&&<div className="section-box" style={{background:"#052e1c15",borderColor:"#25d36630"}}>
          <div className="section-title" style={{color:"#25d366"}}>📱 WhatsApp</div>
          <div style={{fontSize:11,color:"var(--mt)",marginBottom:8}}>Enviar para: <b style={{color:"var(--tx)"}}>{maskPhone(waNum)}</b> ({pac?.resp1Nome||pac?.nome})</div>
          <a href={waConfUrl} target="_blank" rel="noreferrer"
              onClick={()=>{setConfirmacaoEnviada(true);onSave({...ag,status,confirmacaoEnviada:true,numAutorizacao:numAut});}}
              className="btn" style={{background:"#25d366",color:"#fff",textDecoration:"none",display:"inline-flex",alignItems:"center",gap:6,fontSize:12}}>
              📨 Enviar Confirmação
            </a>
          {outrasSessoes.length>0&&<div style={{fontSize:10,color:"#f59e0b",marginTop:5}}>⚡ {outrasSessoes.length} sessão(ões) adicional incluída automaticamente</div>}
        </div>}

        {ag.observacoes&&<div><label>Observações</label><div style={{fontSize:12,padding:"8px 11px",background:"var(--sx)",borderRadius:8,color:"var(--mt)"}}>{ag.observacoes}</div></div>}
        <div className="actions">
          <button className="btn secondary" onClick={onClose}>Fechar</button>
          {podeStatus&&<button className="btn primary" disabled={precisaAut} onClick={()=>salvar()}>💾 Salvar Status</button>}
        </div>
      </div>
    </div>
    {/* Evolução inline */}
    {evModal&&<EvolucaoRapidaModal ag={ag} pac={pac} prof={prof} profLogado={profLogado} prontuarios={prontuarios} setProntuarios={setProntuarios} modelosEv={modelosEv} onClose={()=>setEvModal(false)} onSaved={()=>{setEvModal(false);onSave({...ag,status,numAutorizacao:numAut,confirmacaoEnviada});onClose();}}/>}
  </div>);
}

// ── Evolução Rápida — abre sobre a SessaoModal ──────────────────────────────
// ── Hook: ditado por voz (Web Speech API) ────────────────────────────────────
function useSpeechDictation(onResult, onInterim){
  const [ativo,setAtivo]=useState(false);
  const [suporte]=useState(()=>!!(window.SpeechRecognition||window.webkitSpeechRecognition));
  const recRef=React.useRef(null);

  const iniciar=()=>{
    const SR=window.SpeechRecognition||window.webkitSpeechRecognition;
    if(!SR)return alert("Seu navegador não suporta reconhecimento de voz. Use Chrome ou Edge.");
    const rec=new SR();
    rec.lang="pt-BR";
    rec.continuous=true;
    rec.interimResults=true;
    rec.onresult=e=>{
      let final="",interim="";
      for(let i=e.resultIndex;i<e.results.length;i++){
        if(e.results[i].isFinal)final+=e.results[i][0].transcript+" ";
        else interim+=e.results[i][0].transcript;
      }
      if(final)onResult(final);
      if(onInterim)onInterim(interim);
    };
    rec.onerror=e=>{
      if(e.error==="not-allowed")alert("Permissão de microfone negada. Autorize nas configurações do navegador.");
      setAtivo(false);
    };
    rec.onend=()=>setAtivo(false);
    recRef.current=rec;
    rec.start();
    setAtivo(true);
  };

  const parar=()=>{
    recRef.current?.stop();
    setAtivo(false);
  };

  const toggle=()=>ativo?parar():iniciar();
  return{ativo,suporte,toggle,parar};
}

// ── Botão de microfone reutilizável ───────────────────────────────────────────
function MicBtn({onResult,title,style}){
  const [interim,setInterim]=useState("");
  const {ativo,suporte,toggle}=useSpeechDictation(onResult,setInterim);
  if(!suporte)return null;
  return(<div style={{display:"inline-flex",alignItems:"center",gap:6,...style}}>
    <button onClick={toggle} title={ativo?"Parar ditado (clique)":title||"Ditar por voz"}
      style={{background:ativo?"#ef444415":"transparent",border:"2px solid "+(ativo?"#ef4444":"#7c6af7"),
        borderRadius:20,padding:"4px 12px",cursor:"pointer",fontSize:12,fontWeight:800,
        color:ativo?"#ef4444":"#7c6af7",display:"flex",alignItems:"center",gap:5,transition:".15s",
        animation:ativo?"micPulse 1.2s infinite":undefined}}>
      {ativo?"🔴 Parar":"🎙️ Ditar"}
    </button>
    {ativo&&interim&&<span style={{fontSize:11,color:"var(--mt)",fontStyle:"italic",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>…{interim}</span>}
  </div>);
}

function EvolucaoRapidaModal({ag,pac,prof,profLogado,prontuarios,setProntuarios,modelosEv,onClose,onSaved}){
  const esp=(prof?.especialidades||[""])[0]||"";
  // Hora oficial: respeita usarEvolucaoDiferente
  const horaOficial=ag.usarEvolucaoDiferente&&ag.horaEvolucao?ag.horaEvolucao:ag.horarioSessao;
  const [texto,setTexto]=useState("");
  const [conduta,setConduta]=useState("");
  const [modeloAberto,setModeloAberto]=useState(false);
  const modelosDisp=(modelosEv||MODELOS_EVOLUCAO)[esp]||[];
  const jaEvoluido=(prontuarios||[]).some(p=>p.agId===ag.id);

  const registrar=()=>{
    if(!texto.trim())return alert("Texto de evolução obrigatório");
    const ev={id:Date.now(),agId:ag.id,pacienteId:ag.pacienteId,profId:profLogado?.id||ag.profissionalId,
      especialidade:esp,horaEvolucao:horaOficial,texto,conduta,proxData:"",dataEvolucao:ag.data||hoje_str};
    setProntuarios(a=>[...a,ev]);
    onSaved();
  };

  return(<div className="modal-bg" style={{zIndex:110}} onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{width:660,maxHeight:"90vh",overflowY:"auto"}}>
      <div className="modal-head">
        <div>
          <h2 style={{marginBottom:2}}>📝 Evolução — {pac?.nome}</h2>
          <div style={{fontSize:11,color:"var(--mt)"}}>Sessão: {brDate(ag.data)} {ag.horarioSessao} · <span style={{color:espCor(esp)}}>{esp}</span></div>
        </div>
        <button className="icon-btn" onClick={onClose}>×</button>
      </div>
      {jaEvoluido&&<div style={{padding:"8px 12px",background:"#f59e0b15",borderRadius:8,border:"1px solid #f59e0b30",fontSize:11,color:"#f59e0b",marginBottom:10}}>⚠️ Esta sessão já possui evolução registrada. Um novo registro será adicionado.</div>}
      <div className="stack">
        {/* Data e hora travadas */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
          <div>
            <label>📅 Data da Sessão</label>
            <div style={{padding:"7px 10px",background:"var(--sx)",borderRadius:8,fontSize:13,fontWeight:800,color:"#a78bfa",border:"1px solid #7c6af740"}}>
              {brDate(ag.data)}
            </div>
            <div style={{fontSize:9,color:"var(--mt)",marginTop:2}}>Data da sessão agendada</div>
          </div>
          <div>
            <label>🕐 Hora Oficial</label>
            <div style={{padding:"7px 10px",background:"var(--sx)",borderRadius:8,fontSize:16,fontWeight:900,
              color:ag.usarEvolucaoDiferente?"#f59e0b":"#a78bfa",
              border:"1px solid "+(ag.usarEvolucaoDiferente?"#f59e0b40":"#7c6af740")}}>
              {horaOficial}
            </div>
            <div style={{fontSize:9,marginTop:2}}>
              {ag.usarEvolucaoDiferente
                ?<span style={{color:"#f59e0b",fontWeight:700}}>⚡ Horário alternativo de evolução</span>
                :<span style={{color:"var(--mt)"}}>🕐 Horário da sessão agendada</span>}
            </div>
          </div>
          <div>
            <label>🩺 Especialidade</label>
            <div style={{padding:"7px 10px",background:"var(--sx)",borderRadius:8,fontSize:12,fontWeight:700,color:espCor(esp)||"var(--mt)",border:"1px solid var(--cb)"}}>
              {esp||"—"}
            </div>
          </div>
        </div>
        {modelosDisp.length>0&&<div className="section-box" style={{borderColor:"#7c6af730"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:7}}>
            <div className="section-title" style={{color:"#a78bfa",margin:0}}>🧩 Modelos — {esp}</div>
            <button className="btn secondary small" onClick={()=>setModeloAberto(o=>!o)}>{modeloAberto?"▲ Ocultar":"▼ Ver modelos"}</button>
          </div>
          {modeloAberto&&<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(190px,1fr))",gap:5}}>
            {modelosDisp.map(m=>(
              <button key={m.id} className="btn secondary" style={{textAlign:"left",fontSize:11,padding:"7px 9px"}}
                onClick={()=>{setTexto(m.texto);setModeloAberto(false);}}>
                <div style={{fontWeight:800,color:"#93c5fd",marginBottom:2}}>{m.titulo}</div>
                <div style={{color:"var(--mt)",fontSize:9,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.texto.slice(0,70)}...</div>
              </button>
            ))}
          </div>}
        </div>}
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
            <label style={{margin:0}}>Evolução * <span style={{fontWeight:400,textTransform:"none",fontSize:10,color:"var(--mt)"}}>— descreva o que foi trabalhado na sessão</span></label>
            <MicBtn onResult={t=>setTexto(p=>p+t)} title="Ditar evolução por voz"/>
          </div>
          <textarea rows={7} value={texto} onChange={e=>setTexto(e.target.value)} autoFocus placeholder="Descreva a evolução da sessão... ou clique em 🎙️ Ditar para falar" style={{fontFamily:"inherit",lineHeight:1.6}}/>
        </div>
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
            <label style={{margin:0}}>Conduta</label>
            <MicBtn onResult={t=>setConduta(p=>p+t)} title="Ditar conduta por voz"/>
          </div>
          <textarea rows={2} value={conduta} onChange={e=>setConduta(e.target.value)} placeholder="Próximas condutas... ou 🎙️ Ditar"/>
        </div>
        <div className="actions">
          <button className="btn secondary" onClick={onClose}>Cancelar</button>
          <button className="btn primary" onClick={registrar}>📝 Registrar Evolução</button>
        </div>
      </div>
    </div>
  </div>);
}


function AgendaModal({auth,profissionais,pacientes,procedimentos,salas,agendamentos,convenios,manuais,filiais,onClose,onSave,editing,defaultData,defaultProfId,defaultSalaId}){
  const blank={
    tipo:"sessao",
    pacienteId:"",profissionalId:defaultProfId||"",profResponsavelId:"",
    salaId:defaultSalaId||"",filialId:"",procedimentoId:"",
    data:defaultData||hoje_str,horarioSessao:"08:00",horarioFimSessao:"08:50",tempoSessao:50,
    status:"agendado",convenio:"",plano:"",convenioEditado:false,
    usarEvolucaoDiferente:false,horaEvolucao:"",dataEvolucao:"",horaFimEvolucao:"",tempoEvolucao:50,
    observacoes:"",repeticao:"Nenhuma",qtdRepeticoes:1,repeticaoIndefinida:false,
    // reuniao
    tituloReuniao:"",participantesIds:[],pautaReuniao:""
  };
  const [form,setForm]=useState(editing?{...blank,...editing,participantesIds:editing.participantesIds||[]}:blank);
  const s=p=>setForm(f=>({...f,...p}));
  const isReuniao=form.tipo==="reuniao";
  const DIAS_ABBR=["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"];
  const prof=profissionais.find(p=>p.id===Number(form.profissionalId));
  const pac=pacientes.find(p=>p.id===Number(form.pacienteId));
  const conv=convenios.find(c=>c.nome===(form.convenio||pac?.convenio));
  const manual=conv?.manual?(manuais.find(m=>m.url===conv.manual)||{url:conv.manual}):null;
  const podeStatus=["atendimento","supervisor_adm","administrador","coordenador","coordenador_aba"].includes(auth.role);
  const podeFaturar=["faturamento_supervisor","administrador"].includes(auth.role);

  useEffect(()=>{
    if(pac&&!form.convenioEditado){s({convenio:pac.convenio||"",plano:pac.plano||""});}
  },[form.pacienteId]);
  useEffect(()=>{
    if(form.horarioSessao&&form.tempoSessao){s({horarioFimSessao:addMin(form.horarioSessao,form.tempoSessao)});}
  },[form.horarioSessao,form.tempoSessao]);

  const toggleParticipante=(id)=>{
    const cur=form.participantesIds||[];
    s({participantesIds:cur.includes(id)?cur.filter(x=>x!==id):[...cur,id]});
  };
  const [buscaPartic,setBuscaPartic]=useState("");
  const [filtroEspPartic,setFiltroEspPartic]=useState("");

  const salvar=()=>{
    if(isReuniao&&!form.tituloReuniao.trim())return alert("Informe o título da reunião.");
    if(!isReuniao&&!form.pacienteId)return alert("Selecione o paciente.");
    if(!form.data||!form.horarioSessao)return alert("Data e horário obrigatórios.");
    const base={...form,id:editing?.id||Date.now()};
    if(form.repeticao==="Nenhuma"||isReuniao){
      onSave(base);
    }else{
      const datas=[];
      let cur=new Date(form.data+"T12:00:00");
      const lim=form.repeticaoIndefinida?104:form.qtdRepeticoes;
      for(let i=0;i<lim;i++){
        datas.push(ymd(cur));
        if(form.repeticao==="Diária")cur.setDate(cur.getDate()+1);
        else if(form.repeticao==="Semanal")cur.setDate(cur.getDate()+7);
        else if(form.repeticao==="Quinzenal")cur.setDate(cur.getDate()+14);
        else if(form.repeticao==="Mensal")cur.setMonth(cur.getMonth()+1);
      }
      const grupo=Date.now();
      datas.forEach((d,i)=>onSave({...base,id:grupo+i,data:d,grupoRepeticao:grupo}));
    }
  };

  return(<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{width:700,maxHeight:"95vh",overflowY:"auto"}}>

      {/* Header */}
      <div className="modal-head">
        <h2>{editing?"✏️ Editar":isReuniao?"🤝 Nova":"📅 Novo"} {isReuniao?"Reunião":"Agendamento"}</h2>
        <button className="icon-btn" onClick={onClose}>×</button>
      </div>

      <div className="stack">
        {/* Toggle tipo */}
        <div style={{display:"flex",gap:8}}>
          {[["sessao","📅 Sessão Clínica"],["reuniao","🤝 Reunião"]].map(([v,l])=>(
            <button key={v} onClick={()=>s({tipo:v})}
              style={{flex:1,padding:"10px",borderRadius:10,border:"2px solid "+(form.tipo===v?"#7c6af7":"var(--cpb)"),
                background:form.tipo===v?"#7c6af718":"transparent",color:form.tipo===v?"#a78bfa":"var(--mt)",
                fontWeight:form.tipo===v?900:500,cursor:"pointer",fontSize:13,transition:".15s"}}>
              {l}
            </button>
          ))}
        </div>

        {/* ─── SESSÃO CLÍNICA ─── */}
        {!isReuniao&&<>
          <div className="section-box">
            <div className="section-title">👤 Paciente & Convênio</div>
            <div className="g2">
              <div><label>Paciente *</label>
                <select value={form.pacienteId} onChange={e=>s({pacienteId:e.target.value,convenioEditado:false})}>
                  <option value="">Selecione...</option>{pacientes.map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}
                </select>
              </div>
              <div>
                <label style={{display:"flex",gap:6,alignItems:"center"}}>
                  Convênio
                  {manual&&<a href={manual.url} target="_blank" rel="noreferrer" style={{fontSize:16,textDecoration:"none"}} title="Manual">💡</a>}
                </label>
                <select value={form.convenio||""} onChange={e=>s({convenio:e.target.value,plano:"",convenioEditado:true})}>
                  <option value="">Selecione...</option>
                  {(()=>{
                    const convs=pac?.convenios?.length?pac.convenios:[{nome:pac?.convenio||"",plano:pac?.plano||""}];
                    return convs.filter(cv=>cv.nome).map((cv,i)=><option key={i} value={cv.nome}>{cv.nome}{cv.plano?" — "+cv.plano:""}</option>);
                  })()}
                  {!pac&&CONVENIOS_LIST.map(cv=><option key={cv} value={cv}>{cv}</option>)}
                </select>
              </div>
            </div>
            {(()=>{
              const convs=pac?.convenios?.length?pac.convenios:[{nome:pac?.convenio||"",plano:pac?.plano||""}];
              const sel=convs.find(cv=>cv.nome===form.convenio);
              const planos=sel?.planos||[sel?.plano].filter(Boolean);
              if(!sel||!planos.length)return null;
              return(<div style={{marginTop:8}}>
                <label>Plano</label>
                <select value={form.plano||""} onChange={e=>s({plano:e.target.value})}>
                  <option value="">Selecione...</option>
                  {planos.map((pl,i)=><option key={i} value={pl}>{pl}</option>)}
                </select>
              </div>);
            })()}
            {!pac&&<div style={{marginTop:8}}><label>Plano</label><input value={form.plano||""} onChange={e=>s({plano:e.target.value})}/></div>}
          </div>

          <div className="section-box">
            <div className="section-title">🩺 Profissional & Sala</div>
            <div className="g2">
              <div><label>Profissional *</label>
                <select value={form.profissionalId} onChange={e=>s({profissionalId:e.target.value})}>
                  <option value="">Selecione...</option>{profissionais.filter(p=>p.role==="profissional").map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}
                </select>
              </div>
              <div><label>Prof. Responsável (evolução)</label>
                <select value={form.profResponsavelId} onChange={e=>s({profResponsavelId:e.target.value})}>
                  <option value="">Mesmo profissional</option>{profissionais.filter(p=>p.role==="profissional"||p.role==="coordenador").map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}
                </select>
              </div>
            </div>
            <div className="g3" style={{marginTop:8}}>
              <div><label>Filial</label>
                <select value={form.filialId} onChange={e=>s({filialId:e.target.value,salaId:""})}>
                  <option value="">Todas</option>{filiais.map(f=><option key={f.id} value={f.id}>{f.nome}</option>)}
                </select>
              </div>
              <div><label>Sala</label>
                <select value={form.salaId} onChange={e=>s({salaId:e.target.value})}>
                  <option value="">Sem sala</option>{salas.filter(s=>s.ativa&&(!form.filialId||s.filialId===Number(form.filialId))).map(s=><option key={s.id} value={s.id}>{s.nome}{s.especialidade?" ("+s.especialidade+")":""}</option>)}
                </select>
              </div>
              <div><label>Procedimento</label>
                <select value={form.procedimentoId} onChange={e=>s({procedimentoId:e.target.value})}>
                  <option value="">Selecione</option>{procedimentos.map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}
                </select>
              </div>
            </div>
          </div>
        </>}

        {/* ─── REUNIÃO ─── */}
        {isReuniao&&<>
          <div className="section-box" style={{borderColor:"#0d948840"}}>
            <div className="section-title" style={{color:"#14b8a6"}}>🤝 Dados da Reunião</div>
            <div><label>Título / Assunto *</label>
              <input value={form.tituloReuniao||""} onChange={e=>s({tituloReuniao:e.target.value})} autoFocus placeholder="Ex: Reunião de equipe multiprofissional, Supervisão ABA..."/>
            </div>
            <div style={{marginTop:8}}><label>Pauta / Objetivos</label>
              <textarea rows={2} value={form.pautaReuniao||""} onChange={e=>s({pautaReuniao:e.target.value})} placeholder="Pontos a serem discutidos na reunião..."/>
            </div>
            <div className="g2" style={{marginTop:8}}>
              <div><label>🏢 Filial</label>
                <select value={form.filialId||""} onChange={e=>s({filialId:e.target.value,salaId:""})}>
                  <option value="">Todas as filiais</option>
                  {filiais.map(f=><option key={f.id} value={f.id}>{f.nome}</option>)}
                </select>
              </div>
              <div><label>🚪 Sala</label>
                <select value={form.salaId||""} onChange={e=>s({salaId:e.target.value})}>
                  <option value="">Sem sala definida</option>
                  {salas.filter(s=>s.ativa&&(!form.filialId||s.filialId===Number(form.filialId))).map(s=><option key={s.id} value={s.id}>{s.nome}{s.especialidade?" ("+s.especialidade+")":""}</option>)}
                </select>
              </div>
            </div>
          </div>

          <div className="section-box" style={{borderColor:"#0d948840"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
              <div className="section-title" style={{color:"#14b8a6",margin:0}}>👥 Participantes
                {(form.participantesIds||[]).length>0&&<span style={{marginLeft:8,fontSize:11,fontWeight:700,color:"#2dd4bf",background:"#14b8a620",padding:"1px 8px",borderRadius:20,border:"1px solid #14b8a640"}}>{(form.participantesIds||[]).length}</span>}
              </div>
              {(form.participantesIds||[]).length>0&&<button onClick={()=>s({participantesIds:[]})} style={{background:"none",border:"none",cursor:"pointer",fontSize:11,color:"#f87171",fontWeight:700}}>✕ Limpar todos</button>}
            </div>

            {/* Chips dos selecionados */}
            {(form.participantesIds||[]).length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:4,marginBottom:10,padding:"8px 10px",background:"var(--sx)",borderRadius:8,border:"1px solid var(--cb)"}}>
              {(form.participantesIds||[]).map(id=>{
                const p=profissionais.find(x=>x.id===id);
                if(!p)return null;
                const esp=(p.especialidades||[])[0]||"";
                return(<span key={id} style={{display:"inline-flex",alignItems:"center",gap:4,padding:"3px 10px 3px 8px",borderRadius:20,background:"#14b8a620",color:"#2dd4bf",fontSize:11,fontWeight:700,border:"1px solid #14b8a640"}}>
                  <span style={{width:6,height:6,borderRadius:"50%",background:espCor(esp)||"#14b8a6",flexShrink:0}}/>
                  {p.nome.split(" ")[0]} {p.nome.split(" ").slice(-1)[0]!==p.nome.split(" ")[0]?p.nome.split(" ").slice(-1)[0]:""}
                  <button onClick={()=>toggleParticipante(p.id)} style={{background:"none",border:"none",cursor:"pointer",color:"#5eead4",fontSize:13,lineHeight:1,padding:0,marginLeft:2}}>×</button>
                </span>);
              })}
            </div>}

            {/* Barra de busca + filtro */}
            <div style={{display:"flex",gap:6,marginBottom:8}}>
              <div style={{position:"relative",flex:1}}>
                <span style={{position:"absolute",left:9,top:"50%",transform:"translateY(-50%)",fontSize:13,color:"var(--mt)",pointerEvents:"none"}}>🔍</span>
                <input value={buscaPartic} onChange={e=>setBuscaPartic(e.target.value)}
                  placeholder="Buscar por nome..." autoComplete="off"
                  style={{paddingLeft:30,width:"100%",fontSize:12}}/>
                {buscaPartic&&<button onClick={()=>setBuscaPartic("")} style={{position:"absolute",right:8,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"var(--mt)",fontSize:14,lineHeight:1}}>×</button>}
              </div>
              <select value={filtroEspPartic} onChange={e=>setFiltroEspPartic(e.target.value)} style={{minWidth:140,fontSize:11}}>
                <option value="">Todas especialidades</option>
                {[...new Set(profissionais.flatMap(p=>p.especialidades||[p.especialidade]).filter(Boolean))].sort().map(e=><option key={e} value={e}>{e}</option>)}
              </select>
            </div>

            {/* Seleção rápida por especialidade */}
            {!buscaPartic&&!filtroEspPartic&&<div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:8}}>
              {[...new Set(profissionais.flatMap(p=>p.especialidades||[p.especialidade]).filter(Boolean))].slice(0,8).map(esp=>{
                const ids=profissionais.filter(p=>(p.especialidades||[p.especialidade]).includes(esp)).map(p=>p.id);
                const todos=ids.every(id=>(form.participantesIds||[]).includes(id));
                return(<button key={esp} onClick={()=>{
                  const cur=form.participantesIds||[];
                  s({participantesIds:todos?cur.filter(id=>!ids.includes(id)):[...new Set([...cur,...ids])]});
                }} style={{fontSize:10,padding:"2px 8px",borderRadius:20,border:"1.5px solid "+(todos?"#14b8a6":"var(--cpb)"),background:todos?"#14b8a620":"transparent",color:todos?"#2dd4bf":"var(--mt)",cursor:"pointer",fontWeight:todos?800:500}}>
                  {todos?"✓ ":""}{esp.split(" ")[0]}
                </button>);
              })}
            </div>}

            {/* Lista filtrada — virtualizada com max-height + scroll */}
            {(()=>{
              const termo=buscaPartic.toLowerCase().trim();
              const lista=profissionais.filter(p=>{
                if(filtroEspPartic&&!(p.especialidades||[p.especialidade]).includes(filtroEspPartic))return false;
                if(!termo)return true;
                return p.nome.toLowerCase().includes(termo)||(p.especialidades||[]).some(e=>e.toLowerCase().includes(termo));
              });
              const total=lista.length;
              if(total===0)return(<div style={{textAlign:"center",color:"var(--mt)",padding:"16px 0",fontSize:12}}>Nenhum profissional encontrado.</div>);
              return(<>
                <div style={{fontSize:10,color:"var(--mt)",marginBottom:5}}>{total} profissional(is) encontrado(s){(form.participantesIds||[]).length>0?" · "+((form.participantesIds||[]).filter(id=>lista.find(p=>p.id===id)).length)+" selecionado(s) nesta lista":""}</div>
                <div style={{maxHeight:220,overflowY:"auto",border:"1px solid var(--cb)",borderRadius:8,background:"var(--bg)"}}>
                  {lista.map((p,i)=>{
                    const sel=(form.participantesIds||[]).includes(p.id);
                    const esp=(p.especialidades||[])[0]||"";
                    const cor=espCor(esp)||"#64748b";
                    return(<div key={p.id} onClick={()=>toggleParticipante(p.id)}
                      style={{display:"flex",alignItems:"center",gap:9,padding:"7px 11px",cursor:"pointer",borderBottom:i<lista.length-1?"1px solid var(--db)":"none",background:sel?"#14b8a60a":"transparent",transition:".1s"}}>
                      <div style={{width:16,height:16,borderRadius:4,border:"2px solid "+(sel?"#14b8a6":"var(--cb)"),background:sel?"#14b8a6":"transparent",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
                        {sel&&<span style={{color:"#fff",fontSize:10,fontWeight:900,lineHeight:1}}>✓</span>}
                      </div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:12,fontWeight:sel?800:500,color:sel?"#2dd4bf":"var(--tx)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.nome}</div>
                        {esp&&<div style={{fontSize:10,color:cor,fontWeight:600,marginTop:1}}>{esp}</div>}
                      </div>
                      {p.filiaisAtendimento?.length>0&&<div style={{fontSize:9,color:"var(--mt)",flexShrink:0}}>{p.filiaisAtendimento.length} filial(is)</div>}
                    </div>);
                  })}
                </div>
              </>);
            })()}
          </div>

          <div className="section-box" style={{borderColor:"#7c3aed30"}}>
            <div className="section-title" style={{color:"#a78bfa"}}>📝 Ata da Reunião <span style={{fontSize:10,fontWeight:500,color:"var(--mt)"}}>(opcional — pode ser registrada após a reunião)</span></div>
            <textarea rows={4} value={form.ata||""} onChange={e=>s({ata:e.target.value})}
              placeholder={"Registre aqui os pontos discutidos, decisões tomadas e encaminhamentos...\n• Abertura\n• Discussão\n• Decisões\n• Próximos passos"}
              style={{fontFamily:"inherit",lineHeight:1.6}}/>
            {form.ata&&form.ata.trim().length>30&&<div style={{marginTop:6,fontSize:11,color:"#a78bfa",background:"#7c3aed10",padding:"5px 10px",borderRadius:7,border:"1px solid #7c3aed20"}}>
              💡 Após salvar, abra a reunião na agenda para analisar a ata com IA e sugerir chamados.
            </div>}
          </div>
        </>}

        {/* ─── DATA & HORÁRIO (ambos) ─── */}
        <div className="section-box">
          <div className="section-title">🕐 Data & Horário</div>
          <div className="g4">
            <div><label>Data *</label><input type="date" value={form.data} onChange={e=>s({data:e.target.value})}/></div>
            <div><label>Início</label><input type="time" value={form.horarioSessao} onChange={e=>s({horarioSessao:e.target.value})}/></div>
            <div><label>Duração</label>
              <select value={form.tempoSessao} onChange={e=>s({tempoSessao:Number(e.target.value)})}>
                {(isReuniao?[15,20,30,45,60,75,90,120,150,180,240]:(prof?.temposAtendimento||TEMPOS_SESSAO)).map(t=><option key={t} value={t}>{t} min</option>)}
              </select>
            </div>
            <div><label>Fim</label><input type="time" value={form.horarioFimSessao} readOnly style={{opacity:.7}}/></div>
          </div>
          {!isReuniao&&<div style={{marginTop:8}}>
            <label style={{display:"flex",alignItems:"center",gap:7,cursor:"pointer",textTransform:"none",fontSize:12}}>
              <input type="checkbox" checked={form.usarEvolucaoDiferente} onChange={e=>s({usarEvolucaoDiferente:e.target.checked})}/>
              <span style={{fontWeight:700}}>Horário de evolução diferente</span>
            </label>
            {form.usarEvolucaoDiferente&&<div className="g4" style={{marginTop:8}}>
              <div><label>Data Evolução</label><input type="date" value={form.dataEvolucao||form.data} onChange={e=>s({dataEvolucao:e.target.value})}/></div>
              <div><label>Início Evolução</label><input type="time" value={form.horaEvolucao||""} onChange={e=>{
                const min=Number(e.target.value.split(":")[0])*60+Number(e.target.value.split(":")[1]);
                s({horaEvolucao:e.target.value,horaFimEvolucao:toTime(min+(form.tempoEvolucao||form.tempoSessao||50))});
              }}/></div>
              <div><label>Duração</label>
                <select value={form.tempoEvolucao||form.tempoSessao} onChange={e=>{
                  const t=Number(e.target.value);
                  if(form.horaEvolucao){const[h,m]=form.horaEvolucao.split(":").map(Number);s({tempoEvolucao:t,horaFimEvolucao:toTime(h*60+m+t)});}
                  else s({tempoEvolucao:t});
                }}>
                  {(prof?.temposAtendimento||TEMPOS_SESSAO).map(t=><option key={t} value={t}>{t} min</option>)}
                </select>
              </div>
              <div><label>Fim Evolução</label><input type="time" value={form.horaFimEvolucao||""} readOnly style={{opacity:.7}}/></div>
            </div>}
          </div>}
        </div>

        {/* ─── REPETIÇÃO (só sessão) ─── */}
        {!isReuniao&&<div className="section-box">
          <div className="section-title">🔁 Repetição</div>
          <div className="g2">
            <div><label>Tipo</label>
              <select value={form.repeticao} onChange={e=>s({repeticao:e.target.value,qtdRepeticoes:e.target.value==="Nenhuma"?1:form.qtdRepeticoes||8,repeticaoIndefinida:false})}>
                {REPETICAO_OPTS.map(o=><option key={o}>{o}</option>)}
              </select>
            </div>
            {form.repeticao!=="Nenhuma"&&<div>
              <label style={{display:"flex",alignItems:"center",gap:7,cursor:"pointer",textTransform:"none",fontSize:12,marginBottom:6}}>
                <input type="checkbox" checked={!!form.repeticaoIndefinida} onChange={e=>s({repeticaoIndefinida:e.target.checked,qtdRepeticoes:e.target.checked?104:form.qtdRepeticoes||8})}/>
                <span style={{fontWeight:800,color:"#a78bfa"}}>∞ Para sempre (até cancelar)</span>
              </label>
              {!form.repeticaoIndefinida&&<div>
                <label>Quantidade (1–50)</label>
                <input type="number" min={1} max={50} value={form.qtdRepeticoes} onChange={e=>s({qtdRepeticoes:Math.min(50,Math.max(1,Number(e.target.value)))})}/>
                <div className="helper">{form.qtdRepeticoes} sessão(ões) serão criadas</div>
              </div>}
              {form.repeticaoIndefinida&&<div className="helper" style={{color:"#a78bfa",marginTop:2}}>104 sessões (~2 anos) serão criadas.</div>}
            </div>}
          </div>
        </div>}

        {/* ─── STATUS (só sessão) ─── */}
        {!isReuniao&&<div className="section-box">
          <div className="section-title">📊 Status</div>
          {!podeStatus&&<div className="helper" style={{marginBottom:6}}>⚠️ Apenas Atendimento / Supervisor ADM / Admin alteram status</div>}
          <div className="chips">
            {Object.entries(STATUS_AG).map(([k,v])=>{
              const active=form.status===k;
              const ok=k==="faturado"?podeFaturar:k==="agendado"||k==="confirmado"||k==="cancelado"?true:podeStatus;
              return(<button key={k} className="chip-btn" disabled={!ok} onClick={()=>ok&&s({status:k})} style={{background:active?v.color+"30":"transparent",color:active?v.color:"var(--mt)",borderColor:active?v.color:"var(--cpb)",opacity:ok?1:.35}}>{v.icon} {v.label}</button>);
            })}
          </div>
        </div>}

        {/* ── Cobrança Particular ── */}
        {isParticular&&!isReuniao&&<div className="section-box" style={{borderColor:jaRecebido?"#22c55e40":"#f59e0b40",background:jaRecebido?"#22c55e05":"#f59e0b05"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div className="section-title" style={{color:jaRecebido?"#22c55e":"#f59e0b",margin:0}}>
              💵 Cobrança Particular
              {jaRecebido&&<span style={{marginLeft:8,fontSize:10,fontWeight:700,color:"#22c55e",background:"#22c55e20",padding:"1px 8px",borderRadius:20}}>✅ Recebido</span>}
            </div>
            {!jaRecebido&&<button className="btn primary" style={{fontSize:11,background:"#f59e0b",borderColor:"#f59e0b",padding:"4px 12px"}} onClick={()=>setShowCaixa(v=>!v)}>
              {showCaixa?"✕ Cancelar":"💵 Registrar Pagamento"}
            </button>}
            {jaRecebido&&<button className="btn secondary" style={{fontSize:10}} onClick={()=>setShowCaixa(v=>!v)}>ver</button>}
          </div>
          {/* Valor do procedimento */}
          {!showCaixa&&<div style={{marginTop:8,display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
            <div style={{fontSize:13,fontWeight:900,color:jaRecebido?"#22c55e":"var(--tx)"}}>{brl(proc?.valor||0)}</div>
            <div style={{fontSize:11,color:"var(--mt)"}}>{proc?.nome||"Sessão"}</div>
            {ag.formaPagamento&&jaRecebido&&<span style={{fontSize:11,padding:"2px 8px",borderRadius:20,background:"#22c55e20",color:"#22c55e",fontWeight:800}}>{ag.formaPagamento}</span>}
          </div>}
          {showCaixa&&<div style={{marginTop:10,display:"flex",flexDirection:"column",gap:8}}>
            <div className="g2">
              <div>
                <label>Valor cobrado (R$)</label>
                <input type="number" value={valorCaixa} onChange={e=>setValorCaixa(e.target.value)} step="0.01" min="0" style={{fontSize:14,fontWeight:700}}/>
              </div>
              <div>
                <label>Forma de pagamento</label>
                <div style={{display:"flex",gap:5,flexWrap:"wrap",marginTop:4}}>
                  {["Dinheiro","Cartão","PIX"].map(fp=>(
                    <button key={fp} onClick={()=>setFormaPag(fp)}
                      style={{padding:"5px 12px",borderRadius:8,border:"2px solid "+(formaPag===fp?"#22c55e":"var(--cpb)"),background:formaPag===fp?"#22c55e20":"transparent",color:formaPag===fp?"#22c55e":"var(--mt)",fontWeight:formaPag===fp?900:500,cursor:"pointer",fontSize:12}}>
                      {fp==="Dinheiro"?"💵":fp==="Cartão"?"💳":"📱"} {fp}
                    </button>
                  ))}
                </div>
              </div>
            </div>
            <button className="btn primary" onClick={registrarCaixa} style={{background:"#22c55e",borderColor:"#22c55e",fontWeight:900}}>
              ✅ Confirmar Recebimento de {brl(Number(valorCaixa))} via {formaPag}
            </button>
          </div>}
        </div>}

        <div><label>Observações</label><textarea rows={2} value={form.observacoes||""} onChange={e=>s({observacoes:e.target.value})}/></div>
        <div className="actions">
          <button className="btn secondary" onClick={onClose}>Cancelar</button>
          <button className="btn primary" onClick={salvar}>
            {isReuniao?"🤝 Salvar Reunião":form.repeticao!=="Nenhuma"&&form.repeticaoIndefinida?"∞ Criar ~2 anos":form.repeticao!=="Nenhuma"&&form.qtdRepeticoes>1?"🔁 Criar "+form.qtdRepeticoes+" sessões":"Salvar"}
          </button>
        </div>
      </div>
    </div>
  </div>);
}


// ═══════════════════════════════════════════════════════════════════════════════
// CADASTRO PACIENTE MODAL
// ═══════════════════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════════════════
// CONTRATO PDF — modal com template editável + dados preenchidos + impressão
// ═══════════════════════════════════════════════════════════════════════════════
const TEMPLATE_CONTRATO_PACIENTE = `CONTRATO DE PRESTAÇÃO DE SERVIÇOS TERAPÊUTICOS

CONTRATADA: Espaço Terapêutico Cinthia França
CNPJ: [CNPJ da Clínica]
Endereço: [Endereço da Clínica]

CONTRATANTE (Responsável Legal): {{resp1Nome}}
CPF: {{resp1Cpf}}
WhatsApp: {{resp1Whatsapp}}

PACIENTE: {{nome}}
CPF: {{cpf}}
Data de Nascimento: {{nascimento}}
Convênio: {{convenio}} — Plano: {{plano}}

1. OBJETO DO CONTRATO
O presente contrato tem por objeto a prestação de serviços terapêuticos especializados ao paciente identificado acima, conforme especialidade(s) indicada(s) e avaliação inicial.

2. PERIODICIDADE E HORÁRIOS
As sessões serão realizadas conforme agenda estabelecida entre as partes, com duração e frequência definidas pelo profissional responsável.

3. CANCELAMENTO E FALTAS
3.1. Cancelamentos com menos de 24h de antecedência poderão ser cobrados.
3.2. Faltas sem aviso prévio poderão ser descontadas do pacote ou cobradas separadamente.

4. AUTORIZAÇÃO DE IMAGEM E DADOS
O Contratante autoriza o uso de imagem e dados do paciente exclusivamente para fins terapêuticos e de acompanhamento clínico, em conformidade com a LGPD.

5. CIÊNCIA DAS POLÍTICAS
Declaro estar ciente das políticas de atendimento, cancelamento e faltas da clínica.

Local e data: {{cidade}}, ___/___/______

_________________________          _________________________
Assinatura do Contratante           Assinatura do Responsável
{{resp1Nome}}`;

const TEMPLATE_CONTRATO_PROFISSIONAL = `CONTRATO DE PRESTAÇÃO DE SERVIÇOS PROFISSIONAIS

CONTRATANTE: Espaço Terapêutico Cinthia França
CNPJ: [CNPJ da Clínica]

CONTRATADO(A): {{nome}}
CPF: [CPF do Profissional]
Registro Profissional: {{conselho}} {{numConselho}}
Especialidade(s): {{especialidades}}

1. OBJETO
Prestação de serviços de atendimento terapêutico nas especialidades acima, nas unidades da Contratante.

2. REMUNERAÇÃO E REPASSE
O repasse será calculado conforme tabela de procedimentos vigente, com periodicidade definida em acordo complementar.

3. JORNADA
Dias e horários conforme escala de trabalho acordada, podendo ser revisada mediante comunicação prévia.

4. RESPONSABILIDADES
O(a) Contratado(a) se compromete a manter sigilo profissional, registrar evoluções no sistema, e cumprir os protocolos da clínica.

5. PRAZO
Contrato por prazo indeterminado, podendo ser rescindido por qualquer das partes com aviso prévio de 30 dias.

Local e data: [Cidade], ___/___/______

_________________________          _________________________
Espaço Terapêutico                  Profissional Contratado(a)
Cinthia França                       {{nome}}`;

function ContratoPdfModal({pessoa, tipo, templateOverride, onClose}) {
  const NOME_CLINICA = "Espaço Terapêutico Cinthia França";
  const templateBase = templateOverride || (tipo === "paciente" ? TEMPLATE_CONTRATO_PACIENTE : TEMPLATE_CONTRATO_PROFISSIONAL);
  const [statusContrato, setStatusContrato] = useState(pessoa.statusContrato||"pendente");
  const [editandoTmpl, setEditandoTmpl] = useState(false);
  const [tmplBase, setTmplBase] = useState(templateBase);

  const preencherTemplate = (tmpl, dados) => {
    return tmpl.replace(/\{\{(\w+)\}\}/g, (_, k) => {
      const v = dados[k];
      if (!v) return "[" + k + "]";
      if (k === "cpf" || k === "resp1Cpf") return maskCPF(v);
      if (k === "resp1Whatsapp" || k === "celular") return maskPhone(v);
      if (k === "nascimento") return brDate(v);
      if (k === "especialidades") return Array.isArray(v) ? v.join(", ") : v;
      return v;
    });
  };

  const [template, setTemplate] = useState(preencherTemplate(tmplBase, pessoa));

  const imprimir = () => {
    const win = window.open("", "_blank");
    if(!win)return alert("Popup bloqueado. Permita popups para imprimir.");
    const esc=template.replace(/</g,"&lt;").replace(/>/g,"&gt;");
    const html="<!DOCTYPE html><html><head><meta charset='utf-8'/>"+
      "<title>Contrato "+pessoa.nome+"</title>"+
      "<style>body{font-family:Arial,sans-serif;font-size:13px;line-height:1.8;color:#111;max-width:720px;margin:40px auto;padding:24px;}"+
      "pre{white-space:pre-wrap;font-family:inherit;font-size:13px;line-height:1.8;}"+
      "@media print{body{margin:0;padding:16px;}}</style></head>"+
      "<body><pre>"+esc+"</pre>"+
      "<scr"+"ipt>window.onload=function(){window.print();}</"+"script></body></html>";
    win.document.write(html);
    win.document.close();
  };

  const corSt = COR_CONTRATO[statusContrato]||"#94a3b8";
  const labelSt = LABEL_CONTRATO[statusContrato]||statusContrato;

  return(<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{width:720,maxHeight:"94vh",overflowY:"auto"}}>
      <div className="modal-head">
        <div>
          <h2 style={{marginBottom:4}}>📄 Contrato — {pessoa.nome}</h2>
          <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
            <span style={{fontSize:11,color:"var(--mt)"}}>{tipo==="paciente"?"Prestação de Serviços Terapêuticos":"Prestação de Serviços Profissionais"}</span>
            <span style={{padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:800,background:corSt+"25",color:corSt,border:"1px solid "+corSt+"50"}}>{labelSt}</span>
          </div>
        </div>
        <div style={{display:"flex",gap:6,alignItems:"center"}}>
          <button className="btn secondary" style={{fontSize:11}} onClick={()=>setEditandoTmpl(o=>!o)}>{editandoTmpl?"👁️ Visualizar":"✏️ Editar Modelo"}</button>
          <button className="btn primary" style={{fontSize:11}} onClick={imprimir}>🖨️ Imprimir / PDF</button>
          <button className="icon-btn" onClick={onClose}>×</button>
        </div>
      </div>

      {/* Status do contrato */}
      <div style={{padding:"10px 16px",borderBottom:"1px solid var(--sbd)",display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
        <span style={{fontSize:11,fontWeight:700,color:"var(--mt)",marginRight:4}}>Status:</span>
        {STATUS_CONTRATO.map(s=>{
          const cor=COR_CONTRATO[s]; const ativo=statusContrato===s;
          return(<button key={s} onClick={()=>setStatusContrato(s)}
            style={{padding:"3px 11px",borderRadius:20,fontSize:11,fontWeight:800,cursor:"pointer",
              background:ativo?cor+"25":"transparent",color:ativo?cor:"var(--mt)",
              border:"1.5px solid "+(ativo?cor:"var(--cpb)"),transition:".12s"}}>
            {LABEL_CONTRATO[s]}
          </button>);
        })}
      </div>

      {editandoTmpl ? (
        <div className="stack" style={{padding:"0 4px"}}>
          <div style={{padding:"8px 12px",background:"#7c6af715",borderRadius:8,fontSize:11,color:"#a78bfa",border:"1px solid #7c6af730"}}>
            💡 Use <code style={{background:"var(--sx)",padding:"1px 4px",borderRadius:3}}>{"{{"+"campo"+"}}"}  </code> para dados automáticos. Ex: <code style={{background:"var(--sx)",padding:"1px 4px",borderRadius:3}}>{"{{"+"nome"+"}}"}</code> <code style={{background:"var(--sx)",padding:"1px 4px",borderRadius:3}}>{"{{"+"convenio"+"}}"}</code> <code style={{background:"var(--sx)",padding:"1px 4px",borderRadius:3}}>{"{{"+"resp1Nome"+"}}"}</code>
          </div>
          <div><label>Modelo de Contrato (template base)</label>
            <textarea rows={24} value={tmplBase} onChange={e=>setTmplBase(e.target.value)} style={{fontFamily:"monospace",fontSize:11,lineHeight:1.6}}/>
          </div>
          <div className="actions">
            <button className="btn secondary" onClick={()=>setEditandoTmpl(false)}>Cancelar</button>
            <button className="btn primary" onClick={()=>{setTemplate(preencherTemplate(tmplBase,pessoa));setEditandoTmpl(false);}}>✅ Aplicar Modelo</button>
          </div>
        </div>
      ) : (
        <div className="stack" style={{padding:"0 4px"}}>
          <div style={{background:"var(--sx)",border:"1px solid var(--sc)",borderRadius:10,padding:"24px 28px"}}>
            <pre style={{whiteSpace:"pre-wrap",fontFamily:"'Georgia',serif",fontSize:13,lineHeight:1.9,color:"var(--tx)",margin:0}}>{template}</pre>
          </div>
          <div><label>Ajuste fino do texto (edição direta)</label>
            <textarea rows={8} value={template} onChange={e=>setTemplate(e.target.value)} style={{fontFamily:"monospace",fontSize:11,lineHeight:1.5}}/>
          </div>
          <div className="actions">
            <button className="btn secondary" onClick={()=>setTemplate(preencherTemplate(tmplBase,pessoa))}>🔄 Restaurar dados</button>
            <button className="btn primary" onClick={imprimir}>🖨️ Imprimir / Salvar PDF</button>
          </div>
        </div>
      )}
    </div>
  </div>);
}


function CadastroPacienteModal({editing,pacientes,onClose,onSave}){
  const blank={nome:"",nascimento:"",cpf:"",sexo:"",estadoCivil:"",profissao:"",email:"",celular:"",telFixo:"",convenio:"Particular",plano:"",convenios:[],cep:"",logradouro:"",numero:"",complemento:"",bairro:"",cidade:"",estado:"SP",resp1Nome:"",resp1Cpf:"",resp1Whatsapp:"",resp1TelFixo:"",resp1Grau:"",resp2Nome:"",resp2Celular:"",resp2TelFixo:"",resp2Grau:"",infoImportantes:"",arquivos:[]};
  const [form,setForm]=useState(editing?{...blank,...editing}:blank);
  const [tab,setTab]=useState("dados");
  const sf=p=>setForm(f=>({...f,...p}));
  const validar=()=>{
    const req=[["nome","Nome"],["nascimento","Nascimento"],["cpf","CPF"],["sexo","Sexo"],["celular","Celular"],["convenio","Convênio"],["cep","CEP"],["logradouro","Logradouro"],["bairro","Bairro"],["cidade","Cidade"],["numero","Número"],["resp1Nome","Responsável 1 Nome"],["resp1Cpf","Responsável 1 CPF"],["resp1Whatsapp","Responsável 1 WhatsApp"]];
    for(const[k,l]of req){if(!form[k]||rawD(form[k]).length<(k.includes("cpf")||k.includes("Cpf")?11:k==="celular"||k.includes("Whatsapp")?10:1)){alert("Preencha: "+l);return false;}}
    return true;
  };
  const salvar=()=>{if(!validar())return;onSave({...form,id:editing?.id||Date.now(),cpf:rawD(form.cpf).slice(0,11)});};
  const TABS=[["dados","👤 Dados"],["endereco","📍 Endereço"],["responsavel","👨‍👩‍👧 Responsável"],["docs","📎 Documentos"]];
  return(<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{width:660,maxHeight:"90vh",overflowY:"auto"}}>
      <div className="modal-head"><h2>{editing?"✏️ Editar":"👤 Novo"} Paciente</h2><button className="icon-btn" onClick={onClose}>×</button></div>
      <div className="tab-bar">{TABS.map(([k,l])=><button key={k} className="btn tab-btn" onClick={()=>setTab(k)} style={{background:tab===k?"var(--na)":"transparent",color:tab===k?"#7c6af7":"var(--mt)",fontWeight:tab===k?800:500}}>{l}</button>)}</div>
      <div className="stack">
        {tab==="dados"&&<>
          <div><label>Nome Completo *</label><input value={form.nome} onChange={e=>sf({nome:e.target.value})} autoFocus/></div>
          <div className="g3">
            <div><label>Nascimento *</label><input type="date" value={form.nascimento} onChange={e=>sf({nascimento:e.target.value})}/></div>
            <div><label>CPF *</label><input value={maskCPF(form.cpf)} onChange={e=>sf({cpf:rawD(e.target.value).slice(0,11)})} maxLength={14}/></div>
            <div><label>Sexo *</label><select value={form.sexo} onChange={e=>sf({sexo:e.target.value})}><option value="">Selecione</option><option value="M">Masculino</option><option value="F">Feminino</option><option value="O">Outro</option></select></div>
          </div>
          <div className="g2">
            <div><label>Estado Civil</label><select value={form.estadoCivil||""} onChange={e=>sf({estadoCivil:e.target.value})}><option value="">Selecione</option>{ESTADOS_CIVIS.map(c=><option key={c}>{c}</option>)}</select></div>
            <div><label>Profissão</label><input value={form.profissao||""} onChange={e=>sf({profissao:e.target.value})}/></div>
          </div>
          <div><label>E-mail</label><input type="email" value={form.email||""} onChange={e=>sf({email:e.target.value})} placeholder="email@exemplo.com"/></div>
          <div className="g2">
            <div><label>Celular * (WhatsApp)</label><input value={maskPhone(form.celular)} onChange={e=>sf({celular:rawD(e.target.value).slice(0,11)})} maxLength={15}/></div>
            <div><label>Tel. Fixo (opcional)</label><input value={maskPhone(form.telFixo||"")} onChange={e=>sf({telFixo:rawD(e.target.value).slice(0,11)})} maxLength={15}/></div>
          </div>
          <div>
            <label>Convênios</label>
            {(form.convenios||[]).map((cv,i)=>(
              <div key={i} style={{display:"flex",gap:6,marginBottom:5,alignItems:"center"}}>
                <select value={cv.nome||""} onChange={e=>{const a=[...(form.convenios||[])];a[i]={...a[i],nome:e.target.value};sf({convenios:a,convenio:i===0?e.target.value:form.convenio});}} style={{flex:2}}>
                  <option value="">Selecione...</option>{CONVENIOS_LIST.map(c=><option key={c}>{c}</option>)}
                </select>
                <input value={cv.plano||""} onChange={e=>{const a=[...(form.convenios||[])];a[i]={...a[i],plano:e.target.value};sf({convenios:a,plano:i===0?e.target.value:form.plano});}} placeholder="Plano" style={{flex:2}}/>
                <button onClick={()=>{const a=(form.convenios||[]).filter((_,j)=>j!==i);sf({convenios:a,convenio:a[0]?.nome||"Particular",plano:a[0]?.plano||""});}} style={{background:"none",border:"none",color:"#f87171",cursor:"pointer",fontSize:16,flexShrink:0}}>×</button>
              </div>
            ))}
            <button className="btn secondary small" onClick={()=>sf({convenios:[...(form.convenios||[]),{nome:"",plano:""}]})} style={{fontSize:11,marginTop:2}}>+ Adicionar convênio</button>
          </div>
          <div><label>Informações Importantes</label><textarea rows={2} value={form.infoImportantes||""} onChange={e=>sf({infoImportantes:e.target.value})} placeholder="Tipagem sanguínea, alergias, medicamentos, diagnósticos..."/></div>
        </>}
        {tab==="endereco"&&<CampoEndereco form={form} setForm={setForm}/>}
        {tab==="responsavel"&&<>
          <div className="section-box">
            <div className="section-title">👤 Responsável Principal (emissão NF)</div>
            <div className="stack">
              <div><label>Nome *</label><input value={form.resp1Nome} onChange={e=>sf({resp1Nome:e.target.value})}/></div>
              <div className="g2">
                <div><label>CPF *</label><input value={maskCPF(form.resp1Cpf||"")} onChange={e=>sf({resp1Cpf:rawD(e.target.value).slice(0,11)})} maxLength={14}/></div>
                <div><label>Parentesco</label><select value={form.resp1Grau||""} onChange={e=>sf({resp1Grau:e.target.value})}><option value="">Selecione</option>{GRAUS.map(g=><option key={g}>{g}</option>)}</select></div>
              </div>
              <div className="g2">
                <div><label>WhatsApp * (notificações)</label><input value={maskPhone(form.resp1Whatsapp||"")} onChange={e=>sf({resp1Whatsapp:rawD(e.target.value).slice(0,11)})} maxLength={15}/></div>
                <div><label>Tel. Fixo</label><input value={maskPhone(form.resp1TelFixo||"")} onChange={e=>sf({resp1TelFixo:rawD(e.target.value).slice(0,11)})} maxLength={15}/></div>
              </div>
            </div>
          </div>
          <div className="section-box">
            <div className="section-title">👤 Responsável Secundário (opcional)</div>
            <div className="stack">
              <div><label>Nome</label><input value={form.resp2Nome||""} onChange={e=>sf({resp2Nome:e.target.value})}/></div>
              <div className="g2">
                <div><label>Parentesco</label><select value={form.resp2Grau||""} onChange={e=>sf({resp2Grau:e.target.value})}><option value="">Selecione</option>{GRAUS.map(g=><option key={g}>{g}</option>)}</select></div>
                <div><label>Celular</label><input value={maskPhone(form.resp2Celular||"")} onChange={e=>sf({resp2Celular:rawD(e.target.value).slice(0,11)})} maxLength={15}/></div>
              </div>
              <div><label>Tel. Fixo</label><input value={maskPhone(form.resp2TelFixo||"")} onChange={e=>sf({resp2TelFixo:rawD(e.target.value).slice(0,11)})} maxLength={15}/></div>
            </div>
          </div>
        </>}
        {tab==="docs"&&<>
          <div className="section-box" style={{borderColor:"#f59e0b40",background:"#f59e0b08"}}>
            <div className="section-title" style={{color:"#f59e0b"}}>📎 Documentos Obrigatórios</div>
            <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:10}}>
              {[["RG / CNH do Responsável","🪪"],["Carteirinha do Convênio","💳"],["Laudo / Diagnóstico (se houver)","📋"]].map(([nome,icon])=>{
                const anexado=(form.arquivos||[]).some(f=>f.nome?.toLowerCase().includes(nome.toLowerCase().split(" ")[0].toLowerCase()));
                return(<div key={nome} style={{display:"flex",alignItems:"center",gap:8,fontSize:12}}>
                  <span style={{fontSize:14,flexShrink:0}}>{icon}</span>
                  <span style={{flex:1}}>{nome}</span>
                  <span style={{fontSize:10,fontWeight:800,padding:"1px 7px",borderRadius:20,background:anexado?"#10b98120":"#ef444420",color:anexado?"#34d399":"#f87171"}}>{anexado?"✓ Anexado":"Pendente"}</span>
                </div>);
              })}
            </div>
            <UploadField label="Anexar Documentos" value={form.arquivos||[]} onChange={fs=>sf({arquivos:fs})}/>
          </div>
          <div className="section-box">
            <div className="section-title">📤 Importar Ficha do Paciente</div>
            <div className="helper" style={{marginBottom:8}}>Aceita ficha digitalizada (PDF/imagem) de outros sistemas ou clínicas.</div>
            <UploadField label="Importar Ficha (PDF / Imagem)" value={form.fichaImportada||[]} onChange={fs=>sf({fichaImportada:fs})}/>
            {(form.fichaImportada||[]).length>0&&<div style={{marginTop:6,fontSize:11,color:"#34d399"}}>✅ Ficha importada — verifique os dados nas outras abas.</div>}
          </div>
          {(form.arquivos||[]).length>0&&<div className="section-box">
            <div className="section-title">Arquivos anexados</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
              {(form.arquivos||[]).map((f,i)=><a key={i} href={f.url} target="_blank" rel="noreferrer" className="file-chip">📄 {f.nome}</a>)}
            </div>
          </div>}
        </>}
        <div className="actions"><button className="btn secondary" onClick={onClose}>Cancelar</button><button className="btn primary" onClick={salvar}>Salvar Paciente</button></div>
      </div>
    </div>
  </div>);
}


// ═══════════════════════════════════════════════════════════════════════════════
// CADASTRO PROFISSIONAL MODAL — com múltiplas especialidades, tempos, filiais
// ═══════════════════════════════════════════════════════════════════════════════
function CadastroProfissionalModal({editing,profissionais,filiais,onClose,onSave}){
  const blank={nome:"",usuario:"",senha:"",role:"profissional",especialidades:[],temposAtendimento:[50],escala:{...escalaDefault},temConselho:false,conselho:"CRP",registroConselho:"",conveniosAtendidos:[],filiaisAtendimento:[],filialAcesso:[],email:"",nascimento:"",sexo:"",estadoCivil:"",profissao:"",celular:"",formaPagamento:"PIX",dadosPagamento:"",nivelRepasse:"Pleno",arquivos:[],carimbo:null};
  const [form,setForm]=useState(editing?{...blank,...editing,especialidades:editing.especialidades||[editing.especialidade||""],temposAtendimento:editing.temposAtendimento||[editing.tempoAtendimento||50]}:blank);
  const [tab,setTab]=useState("dados");
  const sf=p=>setForm(f=>({...f,...p}));
  const toggleEsp=e=>sf({especialidades:form.especialidades.includes(e)?form.especialidades.filter(x=>x!==e):[...form.especialidades,e]});
  const toggleTempo=t=>sf({temposAtendimento:form.temposAtendimento.includes(t)?form.temposAtendimento.filter(x=>x!==t):[...form.temposAtendimento,t]});
  const toggleFilialAt=id=>sf({filiaisAtendimento:form.filiaisAtendimento.includes(id)?form.filiaisAtendimento.filter(x=>x!==id):[...form.filiaisAtendimento,id]});
  const toggleFilialAc=id=>sf({filialAcesso:form.filialAcesso.includes(id)?form.filialAcesso.filter(x=>x!==id):[...form.filialAcesso,id]});
  const salvar=()=>{
    if(!form.nome.trim())return alert("Nome obrigatório");
    if(!form.usuario.trim())return alert("CPF/usuário obrigatório");
    if(!form.especialidades.length)return alert("Selecione pelo menos uma especialidade");
    if(!form.temposAtendimento.length)return alert("Selecione pelo menos um tempo de sessão");
    onSave({...form,especialidade:form.especialidades[0],tempoAtendimento:form.temposAtendimento[0]});
  };
  const TABS=[["dados","👤 Dados"],["usuario","🔑 Usuário"],["escala","📅 Escala"],["filiais","🏢 Filiais"],["pagamento","💰 Pagamento"],["docs","📎 Documentos"]];
  return(<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{width:760,maxHeight:"90vh",overflowY:"auto"}}>
      <div className="modal-head"><h2>{editing?"✏️ Editar":"🩺 Novo"} Profissional</h2><button className="icon-btn" onClick={onClose}>×</button></div>
      <div className="tab-bar">{TABS.map(([k,l])=><button key={k} className="btn tab-btn" onClick={()=>setTab(k)} style={{background:tab===k?"var(--na)":"transparent",color:tab===k?"#7c6af7":"var(--mt)",fontWeight:tab===k?800:500}}>{l}</button>)}</div>
      <div className="stack">
        {tab==="dados"&&<>
          <div><label>Nome *</label><input value={form.nome} onChange={e=>sf({nome:e.target.value})} autoFocus/></div>
          <div className="g3">
            <div><label>Nascimento</label><input type="date" value={form.nascimento||""} onChange={e=>sf({nascimento:e.target.value})}/></div>
            <div><label>Sexo</label><select value={form.sexo||""} onChange={e=>sf({sexo:e.target.value})}><option value="">Selecione</option><option value="M">Masculino</option><option value="F">Feminino</option><option value="O">Outro</option></select></div>
            <div><label>Estado Civil</label><select value={form.estadoCivil||""} onChange={e=>sf({estadoCivil:e.target.value})}><option value="">Selecione</option>{ESTADOS_CIVIS.map(c=><option key={c}>{c}</option>)}</select></div>
          </div>
          <div className="g2">
            <div><label>Profissão/Formação</label><input value={form.profissao||""} onChange={e=>sf({profissao:e.target.value})}/></div>
            <div><label>Celular</label><input value={maskPhone(form.celular||"")} onChange={e=>sf({celular:rawD(e.target.value).slice(0,11)})} maxLength={15}/></div>
          </div>
          <div><label>E-mail</label><input type="email" value={form.email||""} onChange={e=>sf({email:e.target.value})} placeholder="email@exemplo.com"/></div>

          <div><label>Especialidades (selecione uma ou mais)</label>
            <div className="chips">{ESPECIALIDADES_LIST.map(e=>{const sel=form.especialidades.includes(e);const cor=espCor(e);return(<button key={e} className="chip-btn" onClick={()=>toggleEsp(e)} style={{background:sel?cor+"22":"transparent",color:sel?cor:"var(--mt)",borderColor:sel?cor:"var(--cpb)",fontWeight:sel?800:400}}>{sel?"✓ ":""}{e}</button>);})}</div>
          </div>
          <div><label>Tempos de Sessão</label>
            <div className="chips">{TEMPOS_SESSAO.map(t=>{const sel=form.temposAtendimento.includes(t);return(<button key={t} className="chip-btn" onClick={()=>toggleTempo(t)} style={{background:sel?"#7c6af722":"transparent",color:sel?"#a78bfa":"var(--mt)",borderColor:sel?"#7c6af7":"var(--cpb)",fontWeight:sel?800:400}}>{sel?"✓ ":""}{t} min</button>);})}</div>
          </div>
          <div><label>Convênios atendidos</label>
            <div className="chips">{CONVENIOS_LIST.map(c=>{const sel=(form.conveniosAtendidos||[]).includes(c);return(<button key={c} className="chip-btn" onClick={()=>sf({conveniosAtendidos:sel?form.conveniosAtendidos.filter(x=>x!==c):[...(form.conveniosAtendidos||[]),c]})} style={{background:sel?"#10b98122":"transparent",color:sel?"#34d399":"var(--mt)",borderColor:sel?"#10b981":"var(--cpb)",fontWeight:sel?800:400}}>{sel?"✓ ":""}{c}</button>);})}</div>
          </div>
          <label style={{display:"flex",gap:6,alignItems:"center",cursor:"pointer",textTransform:"none",fontSize:12,fontWeight:700}}>
            <input type="checkbox" checked={form.temConselho} onChange={e=>sf({temConselho:e.target.checked})}/>
            Possui registro em conselho
          </label>
          {form.temConselho&&<div className="g2">
            <div><label>Conselho</label><select value={form.conselho||""} onChange={e=>sf({conselho:e.target.value})}>{CONSELHOS_OPT.map(c=><option key={c}>{c}</option>)}</select></div>
            <div><label>Registro</label><input value={form.registroConselho||""} onChange={e=>sf({registroConselho:e.target.value})}/></div>
          </div>}
          <div><label>Informações Importantes</label><textarea rows={2} value={form.infoImportantes||""} onChange={e=>sf({infoImportantes:e.target.value})} placeholder="Tipagem sanguínea, alergias, medicamentos..."/></div>
        </>}
        {tab==="usuario"&&<div className="stack">
          <div className="section-box" style={{borderColor:"#7c6af730"}}>
            <div className="section-title" style={{color:"#a78bfa"}}>Credenciais de Acesso</div>
            <div className="g2">
              <div>
                <label>CPF (usuario) *</label>
                <input value={maskCPF(form.usuario)} onChange={e=>sf({usuario:rawD(e.target.value).slice(0,11)})} maxLength={14} placeholder="000.000.000-00"/>
                <div style={{fontSize:10,color:"var(--mt)",marginTop:3}}>Usado para login no sistema</div>
              </div>
              <div>
                <label>Perfil de acesso</label>
                <select value={form.role} onChange={e=>sf({role:e.target.value})}>
                  {PERFIS.map(p=><option key={p} value={p}>{PERFIL_LABEL[p]||p}</option>)}
                </select>
              </div>
            </div>
            <div className="g2" style={{marginTop:8}}>
              <div>
                <label>{editing?"Nova senha (deixe em branco para manter)":"Senha *"}</label>
                <input type="password" value={form.senha} onChange={e=>sf({senha:e.target.value})} autoComplete="new-password"/>
              </div>
              <div>
                <label>Confirmar senha</label>
                <input type="password" value={form.senhaConfirm||""} onChange={e=>sf({senhaConfirm:e.target.value})}/>
                {form.senhaConfirm&&form.senha!==form.senhaConfirm&&<div style={{fontSize:10,color:"#f87171",marginTop:3}}>Senhas nao coincidem</div>}
                {form.senhaConfirm&&form.senha===form.senhaConfirm&&form.senha&&<div style={{fontSize:10,color:"#34d399",marginTop:3}}>Senhas coincidem</div>}
              </div>
            </div>
          </div>
          <div className="section-box" style={{borderColor:"#f59e0b30"}}>
            <div className="section-title" style={{color:"#f59e0b"}}>Permissoes do Perfil</div>
            {(()=>{
              const pOk={administrador:["Acesso total","Gerenciar usuarios","Financeiro","Relatorios"],profissional:["Agenda propria","Evolucoes","Prontuarios vinculados"],coordenador:["Todas as agendas","Evolucoes","Comissao de prontuarios"],secretaria:["Agendamento","Pacientes","Fila de espera"],faturamento:["Faturamento","Convenios","Relatorios"],agendamento:["Agenda completa","Cadastros","Chamados"]}[form.role]||[];
              const pNo={profissional:["Financeiro","Outros usuarios"],coordenador:["Faturamento"],secretaria:["Evolucoes clinicas","Financeiro"],faturamento:["Evolucoes clinicas"],agendamento:["Evolucoes","Financeiro"],administrador:[]}[form.role]||[];
              return(<div style={{display:"flex",gap:16,flexWrap:"wrap",marginTop:6}}>
                <div>{pOk.map((p,i)=>(<div key={i} style={{fontSize:12,color:"#34d399",padding:"2px 0"}}>{"+ "+p}</div>))}</div>
                {pNo.length>0&&<div>{pNo.map((p,i)=>(<div key={i} style={{fontSize:12,color:"#f87171",padding:"2px 0"}}>{"- "+p}</div>))}</div>}
              </div>);
            })()}
          </div>
          {editing&&<div style={{padding:"10px 14px",background:"#ef444408",borderRadius:9,border:"1px solid #ef444430",fontSize:11,color:"#f87171"}}>
            Atencao: Alteracoes nas credenciais tem efeito no proximo login.
          </div>}
        </div>}
        {tab==="escala"&&<>
          <div className="section-box">
            <div className="section-title">📅 Escala Semanal + Pausa</div>
            <div style={{display:"grid",gridTemplateColumns:"120px 1fr 1fr 80px 80px auto",gap:6,marginBottom:6,padding:"0 4px"}}>
              <div className="helper">DIA</div><div className="helper">INÍCIO</div><div className="helper">FIM</div>
              <div className="helper" style={{fontSize:9}}>PAUSA↓</div><div className="helper" style={{fontSize:9}}>PAUSA↑</div>
              <div className="helper" style={{fontSize:9}}>PAUSA</div>
            </div>
            {DIAS_SEMANA.map(dia=>{
              const esc=form.escala?.[dia]||{ativo:false,inicio:"08:00",fim:"18:00",pausaInicio:"12:00",pausaFim:"13:00",usarPausa:false};
              const upd=patch=>sf({escala:{...form.escala,[dia]:{...esc,...patch}}});
              return(<div key={dia} style={{display:"grid",gridTemplateColumns:"120px 1fr 1fr 80px 80px auto",gap:6,alignItems:"center",marginBottom:4}}>
                <label style={{display:"flex",alignItems:"center",gap:5,cursor:"pointer",textTransform:"none",fontSize:12}}>
                  <input type="checkbox" checked={!!esc.ativo} onChange={()=>upd({ativo:!esc.ativo})}/>
                  <span style={{color:esc.ativo?"#a78bfa":"var(--mt)",fontWeight:esc.ativo?700:400}}>{dia}</span>
                </label>
                <input type="time" disabled={!esc.ativo} value={esc.inicio} onChange={e=>upd({inicio:e.target.value})} style={{opacity:esc.ativo?1:0.3}}/>
                <input type="time" disabled={!esc.ativo} value={esc.fim} onChange={e=>upd({fim:e.target.value})} style={{opacity:esc.ativo?1:0.3}}/>
                <input type="time" disabled={!esc.ativo||!esc.usarPausa} value={esc.pausaInicio||"12:00"} onChange={e=>upd({pausaInicio:e.target.value})} style={{opacity:esc.ativo&&esc.usarPausa?1:0.3}}/>
                <input type="time" disabled={!esc.ativo||!esc.usarPausa} value={esc.pausaFim||"13:00"} onChange={e=>upd({pausaFim:e.target.value})} style={{opacity:esc.ativo&&esc.usarPausa?1:0.3}}/>
                <input type="checkbox" checked={!!esc.usarPausa} disabled={!esc.ativo} onChange={()=>upd({usarPausa:!esc.usarPausa})} title="Ativar pausa"/>
              </div>);
            })}
          </div>
        </>}
        {tab==="filiais"&&<>
          <div className="section-box">
            <div className="section-title">🏢 Filiais de Atendimento</div>
            <div className="chips">{filiais.map(f=>{const sel=(form.filiaisAtendimento||[]).includes(f.id);return(<button key={f.id} className="chip-btn" onClick={()=>toggleFilialAt(f.id)} style={{background:sel?"#7c6af722":"transparent",color:sel?"#a78bfa":"var(--mt)",borderColor:sel?"#7c6af7":"var(--cpb)",fontWeight:sel?800:400}}>{sel?"✓ ":""}{f.nome}</button>);})}</div>
          </div>
          <div className="section-box">
            <div className="section-title">🔑 Filiais de Acesso</div>
            <div className="helper" style={{marginBottom:6}}>Filiais que o profissional pode visualizar no sistema</div>
            <div className="chips">{filiais.map(f=>{const sel=(form.filialAcesso||[]).includes(f.id);return(<button key={f.id} className="chip-btn" onClick={()=>toggleFilialAc(f.id)} style={{background:sel?"#10b98122":"transparent",color:sel?"#34d399":"var(--mt)",borderColor:sel?"#10b981":"var(--cpb)",fontWeight:sel?800:400}}>{sel?"✓ ":""}{f.nome}</button>);})}</div>
          </div>
        </>}
        {tab==="pagamento"&&<>
          <div className="section-box">
            <div className="section-title">💰 Forma de Pagamento e Repasse</div>
            <div className="stack">
              <div className="g2">
                <div><label>Forma de Pagamento</label><select value={form.formaPagamento||"PIX"} onChange={e=>sf({formaPagamento:e.target.value})}>{FORMA_PAGAMENTO.map(f=><option key={f}>{f}</option>)}</select></div>
                <div><label>Nível de Repasse</label><select value={form.nivelRepasse||"Pleno"} onChange={e=>sf({nivelRepasse:e.target.value})}>{NIVEL_REPASSE.map(n=><option key={n}>{n}</option>)}</select></div>
              </div>
              <div><label>Dados de Pagamento (chave PIX, banco, agência, conta)</label><textarea rows={2} value={form.dadosPagamento||""} onChange={e=>sf({dadosPagamento:e.target.value})} placeholder="Ex: chave PIX 000.000.000-00 ou banco/agência/conta"/></div>
            </div>
          </div>
        </>}
        {tab==="docs"&&<>
          {/* ── Carimbo do Profissional ── */}
          <div className="section-box" style={{borderColor:"#7c6af730"}}>
            <div className="section-title" style={{color:"#a78bfa"}}>🔏 Carimbo do Profissional</div>
            <div className="helper" style={{marginBottom:10}}>
              Imagem do carimbo que aparecerá no prontuário ao registrar evoluções. O sistema removerá automaticamente o fundo branco, deixando transparente.
            </div>
            {form.carimbo
              ? <div style={{display:"flex",flexDirection:"column",gap:10,alignItems:"flex-start"}}>
                  <div style={{background:"repeating-conic-gradient(#1e293b 0% 25%,#0f172a 0% 50%) 0 0/16px 16px",borderRadius:10,padding:12,display:"inline-block",border:"1px solid var(--cb)"}}>
                    <img src={form.carimbo} alt="carimbo" style={{maxWidth:280,maxHeight:120,display:"block"}}/>
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <label className="btn secondary" style={{cursor:"pointer",fontSize:11}}>
                      🔄 Trocar imagem
                      <input type="file" accept="image/*" style={{display:"none"}} onChange={e=>{
                        const file=e.target.files[0];if(!file)return;
                        const reader=new FileReader();
                        reader.onload=ev=>{
                          const img=new Image();
                          img.onload=()=>{
                            const canvas=document.createElement("canvas");
                            canvas.width=img.width;canvas.height=img.height;
                            const ctx=canvas.getContext("2d");
                            ctx.drawImage(img,0,0);
                            const id=ctx.getImageData(0,0,canvas.width,canvas.height);
                            const d=id.data;
                            for(let i=0;i<d.length;i+=4){
                              const r=d[i],g=d[i+1],b=d[i+2];
                              const brightness=(r+g+b)/3;
                              if(brightness>210){d[i+3]=0;}
                              else if(brightness>170){d[i+3]=Math.round((210-brightness)/40*255);}
                            }
                            ctx.putImageData(id,0,0);
                            sf({carimbo:canvas.toDataURL("image/png")});
                          };
                          img.src=ev.target.result;
                        };
                        reader.readAsDataURL(file);
                        e.target.value="";
                      }}/>
                    </label>
                    <button className="btn secondary" style={{fontSize:11,color:"#f87171",borderColor:"#f87171"}} onClick={()=>sf({carimbo:null})}>🗑️ Remover</button>
                  </div>
                </div>
              : <label style={{display:"inline-flex",alignItems:"center",gap:8,cursor:"pointer",padding:"10px 16px",background:"var(--sx)",border:"2px dashed #7c6af760",borderRadius:10,fontSize:12,color:"#a78bfa",fontWeight:700}}>
                  📁 Selecionar imagem do carimbo
                  <input type="file" accept="image/*" style={{display:"none"}} onChange={e=>{
                    const file=e.target.files[0];if(!file)return;
                    const reader=new FileReader();
                    reader.onload=ev=>{
                      const img=new Image();
                      img.onload=()=>{
                        const canvas=document.createElement("canvas");
                        canvas.width=img.width;canvas.height=img.height;
                        const ctx=canvas.getContext("2d");
                        ctx.drawImage(img,0,0);
                        const id=ctx.getImageData(0,0,canvas.width,canvas.height);
                        const d=id.data;
                        for(let i=0;i<d.length;i+=4){
                          const r=d[i],g=d[i+1],b=d[i+2];
                          const brightness=(r+g+b)/3;
                          if(brightness>210){d[i+3]=0;}
                          else if(brightness>170){d[i+3]=Math.round((210-brightness)/40*255);}
                        }
                        ctx.putImageData(id,0,0);
                        sf({carimbo:canvas.toDataURL("image/png")});
                      };
                      img.src=ev.target.result;
                    };
                    reader.readAsDataURL(file);
                    e.target.value="";
                  }}/>
                </label>}
          </div>
          {/* ── Outros documentos ── */}
          <div className="section-box">
            <div className="section-title">📎 Documentos do Profissional</div>
            <div className="helper" style={{marginBottom:8}}>Foto do conselho, documentos para referência.</div>
            <UploadField label="Conselho + Documentos" value={form.arquivos||[]} onChange={fs=>sf({arquivos:fs})}/>
          </div>
          {(form.arquivos||[]).length>0&&<div className="section-box">
            <div className="section-title">Arquivos anexados</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
              {(form.arquivos||[]).map((f,i)=><a key={i} href={f.url} target="_blank" rel="noreferrer" className="file-chip">📄 {f.nome}</a>)}
            </div>
          </div>}
        </>}
        <div className="actions"><button className="btn secondary" onClick={onClose}>Cancelar</button><button className="btn primary" onClick={salvar}>Salvar</button></div>
      </div>
    </div>
  </div>);
}


// ═══════════════════════════════════════════════════════════════════════════════
// PRONTUÁRIO MODAL
// ═══════════════════════════════════════════════════════════════════════════════
function HistoricoProntuario({evs,agenda,paciente,profissionais}){
  const [periodoIni,setPeriodoIni]=useState("");
  const [periodoFim,setPeriodoFim]=useState("");
  const hoje=new Date();
  const subMes=n=>{const d=new Date();d.setMonth(d.getMonth()-n);return d.toISOString().slice(0,10);};
  const evsFiltr=evs.filter(ev=>{
    if(periodoIni&&ev.dataEvolucao<periodoIni)return false;
    if(periodoFim&&ev.dataEvolucao>periodoFim)return false;
    return true;
  });
  const gerarPDF=()=>{
    const CLINICA="Espaço Terapêutico Cinthia França";
    const periodoStr=(periodoIni||periodoFim)
      ?(" — Período: "+(periodoIni?brDate(periodoIni):"início")+" a "+(periodoFim?brDate(periodoFim):"hoje")):"";
    const rows=evsFiltr.map(ev=>{
      const agRef=(agenda||[]).find(a=>a.id===ev.agId);
      const evProf=profissionais.find(p=>p.id===Number(ev.profId));
      const usouAlt=agRef?.usarEvolucaoDiferente&&agRef?.horaEvolucao;
      const horaIni=usouAlt?agRef.horaEvolucao:(agRef?.horarioSessao||ev.horaEvolucao||"—");
      const horaFim=agRef?.horarioFimSessao||"";
      return(
        "<div style='border:1px solid #ddd;border-radius:8px;padding:14px;margin-bottom:14px;border-left:4px solid #7c6af7'>"+
        "<div style='display:flex;justify-content:space-between;margin-bottom:8px;flex-wrap:wrap;gap:6px'>"+
        "<div>"+
        "<span style='font-weight:900;color:#1e40af;font-size:15px'>"+brDate(ev.dataEvolucao)+"</span>"+
        "<span style='margin:0 10px;color:#374151;font-size:13px;font-weight:700'>"+horaIni+(horaFim?" &ndash; "+horaFim:"")+"</span>"+
        (usouAlt?"<span style='font-size:10px;color:#b45309;background:#fef3c7;padding:2px 7px;border-radius:12px;font-weight:700'>&#9889; Horário alternativo</span>":"")+
        "</div>"+
        "<div style='text-align:right'>"+
        (ev.especialidade?"<div style='font-size:10px;background:#eff6ff;color:#1d4ed8;padding:2px 8px;border-radius:12px;font-weight:700;display:inline-block'>"+ev.especialidade+"</div>":"") +
        "<div style='font-size:11px;color:#555;margin-top:4px'>Prof: <b>"+(evProf?.nome||"—")+"</b></div>"+
        (agRef?"<div style='font-size:10px;color:#6b7280'>Sessão agendada: "+agRef.horarioSessao+(agRef.horarioFimSessao?" – "+agRef.horarioFimSessao:"")+"</div>":"")+
        "</div>"+
        "</div>"+
        "<div style='font-size:13px;line-height:1.7;white-space:pre-wrap;color:#111;padding:6px 0'>"+ev.texto.replace(/</g,"&lt;").replace(/>/g,"&gt;")+"</div>"+
        (ev.conduta?"<div style='margin-top:8px;padding:6px 10px;background:#f8fafc;border-radius:6px;font-size:12px;color:#475569;border:1px solid #e2e8f0'>&#128204; Conduta: "+ev.conduta.replace(/</g,"&lt;")+"</div>":"")+
        (ev.proxData?"<div style='font-size:11px;color:#b45309;margin-top:6px;font-weight:700'>&#128197; Próxima sessão: "+brDate(ev.proxData)+"</div>":"")+
        "</div>"
      );
    }).join("");
    const html="<!DOCTYPE html><html><head><meta charset='UTF-8'>"+
      "<title>Prontuário — "+paciente.nome+"</title>"+
      "<style>body{font-family:Arial,sans-serif;padding:30px;color:#111;max-width:820px;margin:0 auto}"+
      "h1{color:#1e3a8a;font-size:20px;margin-bottom:4px}"+
      ".info{color:#555;font-size:12px;margin-bottom:4px}"+
      ".header-block{display:flex;justify-content:space-between;border-bottom:2px solid #1e3a8a;padding-bottom:14px;margin-bottom:22px;align-items:flex-start}"+
      ".clinica{text-align:right;font-size:11px;color:#555}"+
      "@media print{body{padding:10px}@page{margin:15mm}}</style></head><body>"+
      "<div class='header-block'>"+
      "<div><h1>📋 Histórico de Prontuário</h1>"+
      "<div class='info'><b>"+paciente.nome+"</b></div>"+
      "<div class='info'>Nasc: "+brDate(paciente.nascimento||"")+" · "+paciente.convenio+(paciente.plano?" / "+paciente.plano:"")+"</div>"+
      "<div class='info'>"+periodoStr.replace("—","")+"</div></div>"+
      "<div class='clinica'>"+CLINICA+"<br/>Emitido em: "+brDate(hoje_str)+"<br/>"+evsFiltr.length+" evolução(ões)</div>"+
      "</div>"+rows+
      "<div style='margin-top:30px;padding-top:14px;border-top:1px solid #ccc;font-size:10px;color:#999;text-align:center'>"+
      CLINICA+" — FOCOE · Cuide</div>"+
      "<scr"+"ipt>window.onload=function(){window.print();}</"+"script></body></html>";
    const w=window.open("","_blank","width=900,height=700");
    w.document.write(html);w.document.close();
  };

  return(<div className="stack">
    {/* ── Barra de período + PDF ── */}
    <div style={{display:"flex",gap:8,alignItems:"flex-end",flexWrap:"wrap",padding:"10px 12px",background:"var(--na)",borderRadius:10,border:"1px solid var(--sc)"}}>
      <div>
        <label style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",letterSpacing:.5,display:"block",marginBottom:3}}>De</label>
        <input type="date" value={periodoIni} onChange={e=>setPeriodoIni(e.target.value)} style={{fontSize:12,width:135}}/>
      </div>
      <div>
        <label style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",letterSpacing:.5,display:"block",marginBottom:3}}>Até</label>
        <input type="date" value={periodoFim} onChange={e=>setPeriodoFim(e.target.value)} style={{fontSize:12,width:135}}/>
      </div>
      <div style={{display:"flex",gap:4,flexWrap:"wrap",alignItems:"flex-end",paddingBottom:1}}>
        {[["30d","30 dias",subMes(1),hoje_str],["3m","3 meses",subMes(3),hoje_str],["6m","6 meses",subMes(6),hoje_str],["1a","1 ano",subMes(12),hoje_str],["tudo","Tudo","",""]].map(([k,l,ini,fim])=>(
          <button key={k} className="btn secondary small"
            style={{fontSize:10,background:(periodoIni===ini&&periodoFim===fim)?"#7c6af722":"transparent",color:(periodoIni===ini&&periodoFim===fim)?"#a78bfa":"var(--mt)",border:"1px solid var(--cpb)"}}
            onClick={()=>{setPeriodoIni(ini);setPeriodoFim(fim);}}>{l}</button>
        ))}
      </div>
      <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center"}}>
        <span style={{fontSize:11,color:"var(--mt)"}}>{evsFiltr.length}/{evs.length}</span>
        <button className="btn secondary" style={{fontSize:11,gap:4,display:"flex",alignItems:"center"}}
          onClick={gerarPDF} disabled={evsFiltr.length===0}>🖨️ PDF</button>
      </div>
    </div>

    {/* ── Evoluções ── */}
    {evsFiltr.length===0
      ?<div className="muted" style={{textAlign:"center",padding:20}}>Nenhuma evolução no período.</div>
      :evsFiltr.map(ev=>{
        const evProf=profissionais.find(p=>p.id===Number(ev.profId));
        const cor=espCor(ev.especialidade||"")||"#7c6af7";
        const agRef=(agenda||[]).find(a=>a.id===ev.agId);
        const usouAlt=agRef?.usarEvolucaoDiferente&&agRef?.horaEvolucao;
        const horaIni=usouAlt?agRef.horaEvolucao:(agRef?.horarioSessao||ev.horaEvolucao||"—");
        const horaFim=agRef?.horarioFimSessao||"";
        return(<div key={ev.id} className="card" style={{padding:14,borderLeft:"4px solid "+cor}}>
          {/* Cabeçalho */}
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:8,flexWrap:"wrap",gap:6}}>
            <div style={{display:"flex",flexDirection:"column",gap:4}}>
              <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontWeight:900,fontSize:15,color:"#a78bfa"}}>{brDate(ev.dataEvolucao)}</span>
                <div style={{display:"flex",alignItems:"center",gap:5,background:"var(--sx)",padding:"3px 10px",borderRadius:20,border:"1px solid var(--sc)"}}>
                  <span style={{fontSize:12,fontWeight:800,color:usouAlt?"#f59e0b":"#e2e8f0"}}>🕐 {horaIni}</span>
                  {horaFim&&<>
                    <span style={{fontSize:11,color:"var(--mt)"}}>–</span>
                    <span style={{fontSize:12,fontWeight:700,color:"var(--mt)"}}>{horaFim}</span>
                  </>}
                </div>
                {usouAlt&&<span style={{fontSize:9,padding:"2px 7px",borderRadius:20,background:"#f59e0b20",color:"#f59e0b",fontWeight:800,border:"1px solid #f59e0b40"}}>⚡ Horário alternativo</span>}
              </div>
              {agRef&&<div style={{fontSize:10,color:"var(--mt)"}}>
                Sessão agendada: <b>{agRef.horarioSessao}{agRef.horarioFimSessao?" – "+agRef.horarioFimSessao:""}</b>
                {usouAlt&&<span style={{color:"#f59e0b",marginLeft:6}}>(evolução em horário diferente)</span>}
              </div>}
            </div>
            <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:4}}>
              {ev.especialidade&&<span style={{fontSize:9,padding:"2px 8px",borderRadius:20,background:cor+"22",color:cor,fontWeight:800,border:"1px solid "+cor+"44"}}>{ev.especialidade}</span>}
              <div style={{display:"flex",alignItems:"center",gap:6}}>
                {evProf?.carimbo&&<img src={evProf.carimbo} alt="carimbo" style={{height:24,opacity:.85}}/>}
                <div style={{textAlign:"right"}}>
                  <div style={{fontSize:12,fontWeight:800,color:"var(--tx)"}}>{evProf?.nome||"—"}</div>
                  {evProf&&<div style={{fontSize:9,color:"var(--mt)"}}>{(evProf.especialidades||[])[0]}</div>}
                </div>
              </div>
            </div>
          </div>
          {/* Texto */}
          <div style={{fontSize:13,lineHeight:1.7,whiteSpace:"pre-wrap",color:"var(--tx)",padding:"6px 0",borderTop:"1px solid var(--db)"}}>{ev.texto}</div>
          {ev.conduta&&<div style={{fontSize:11,color:"var(--mt)",marginTop:6,padding:"6px 10px",background:"var(--sx)",borderRadius:7,border:"1px solid var(--sc)"}}>📌 <b>Conduta:</b> {ev.conduta}</div>}
          {ev.proxData&&<div style={{fontSize:11,color:"#f59e0b",marginTop:5,fontWeight:700}}>📅 Próxima sessão: {brDate(ev.proxData)}</div>}
        </div>);
    })}
  </div>);
}

function ProntuarioModal({paciente,agenda,profissionais,procedimentos,prontuarios,setProntuarios,modelosEv,auth,onClose}){
  const [tab,setTab]=useState("historico");
  const [nova,setNova]=useState({horaEvolucao:"",dataEvolucao:"",texto:"",conduta:"",proxData:"",especialidade:"",agId:""});
  const [modeloAberto,setModeloAberto]=useState(false);

  const agsPac=(agenda||[]).filter(a=>Number(a.pacienteId)===Number(paciente.id)).sort((a,b)=>b.data.localeCompare(a.data));
  const evs=(prontuarios||[]).filter(p=>Number(p.pacienteId)===Number(paciente.id)).sort((a,b)=>b.dataEvolucao.localeCompare(a.dataEvolucao));

  const profLogado=profissionais.find(p=>p.usuario===auth.usuario);

  // ── Regra: profissional que pode evoluir ─────────────────────────────────
  // Se nova.agId preenchido, verifica se o agendamento tem profResponsavelId
  const agSelecionado=nova.agId?(agenda||[]).find(a=>String(a.id)===String(nova.agId)):null;
  const profRespId=agSelecionado?.profResponsavelId||agSelecionado?.profissionalId;
  const profResp=profRespId?profissionais.find(p=>p.id===Number(profRespId)):null;
  // podeEv base (role) E deve ser o profResponsavel do agendamento selecionado (se houver)
  const podeEvRole=PODE_EVOLUIR(auth.role);
  const podeEv=podeEvRole&&(!agSelecionado||!profRespId||(profLogado&&Number(profLogado.id)===Number(profRespId)));
  const bloqueadoPorProf=podeEvRole&&agSelecionado&&profRespId&&profLogado&&Number(profLogado.id)!==Number(profRespId);

  const espProf=(profLogado?.especialidades||[""])[0]||"";
  const espEvolucao=nova.especialidade||espProf;
  const modelosDisp=(modelosEv||MODELOS_EVOLUCAO)[espEvolucao]||[];

  // ── Ao selecionar sessão: pré-preenche data/hora oficial ─────────────────
  const selecionarSessao=(agId)=>{
    const ag=(agenda||[]).find(a=>String(a.id)===String(agId));
    if(!ag){setNova(n=>({...n,agId:"",dataEvolucao:"",horaEvolucao:""}));return;}
    // Se usarEvolucaoDiferente → usa horaEvolucao do agendamento; data permanece a da sessão
    const dataOf=ag.data;
    const horaOf=ag.usarEvolucaoDiferente&&ag.horaEvolucao?ag.horaEvolucao:ag.horarioSessao;
    setNova(n=>({...n,agId,dataEvolucao:dataOf,horaEvolucao:horaOf}));
  };

  const registrar=()=>{
    if(!nova.texto.trim())return alert("Texto obrigatório");
    if(bloqueadoPorProf)return alert("Somente "+profResp?.nome+" pode registrar a evolução desta sessão.");
    // Prof oficial: se agendamento tem profResponsavelId diferente, usa esse
    const profOfId=agSelecionado?.profResponsavelId?Number(agSelecionado.profResponsavelId):profLogado?.id;
    const ev={
      id:Date.now(),
      pacienteId:paciente.id,
      profId:profOfId,
      agId:nova.agId||null,
      especialidade:espEvolucao,
      dataEvolucao:nova.dataEvolucao||hoje_str,
      horaEvolucao:nova.horaEvolucao||toTime(new Date().getHours()*60+new Date().getMinutes()),
      texto:nova.texto,
      conduta:nova.conduta,
      proxData:nova.proxData,
    };
    setProntuarios(a=>[...a,ev]);
    setNova({horaEvolucao:"",dataEvolucao:"",texto:"",conduta:"",proxData:"",especialidade:"",agId:""});
    setTab("historico");
  };

  const tabBtn=(key,label)=>(
    <button className="btn tab-btn" onClick={()=>setTab(key)} style={{background:tab===key?"var(--na)":"transparent",color:tab===key?"#7c6af7":"var(--mt)",fontWeight:tab===key?800:500}}>{label}</button>
  );

  return(<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{width:780,maxHeight:"93vh",overflowY:"auto"}}>
      <div className="modal-head">
        <div>
          <h2 style={{marginBottom:2}}>📋 Prontuário — {paciente.nome}</h2>
          <div style={{fontSize:11,color:"var(--mt)"}}>{paciente.convenio||"Particular"} · {paciente.plano||""} · Nasc: {brDate(paciente.nascimento||"")}</div>
        </div>
        <button className="icon-btn" onClick={onClose}>×</button>
      </div>
      <div className="tab-bar">
        {tabBtn("historico","📋 Histórico ("+evs.length+")")}
        {podeEv&&tabBtn("nova","✍️ Nova Evolução")}
        {tabBtn("agendamentos","📅 Sessões ("+agsPac.length+")")}
        {tabBtn("ficha","📄 Ficha")}
      </div>

      {/* HISTÓRICO */}
      {tab==="historico"&&<HistoricoProntuario evs={evs} agenda={agenda} paciente={paciente} profissionais={profissionais} />}

      {/* NOVA EVOLUÇÃO */}
      {tab==="nova"&&podeEvRole&&<div className="stack">

        {/* ── Vincular à sessão agendada ── */}
        <div className="section-box" style={{borderColor:"#7c6af730",background:"var(--na)"}}>
          <div className="section-title" style={{color:"#a78bfa",marginBottom:6}}>📅 Sessão de Referência</div>
          <select value={nova.agId} onChange={e=>selecionarSessao(e.target.value)}>
            <option value="">— Selecione a sessão (opcional) —</option>
            {agsPac.slice(0,20).map(ag=>{
              const stInfo=STATUS_AG[ag.status]||{icon:"📋",label:ag.status};
              const prAg=profissionais.find(p=>p.id===Number(ag.profissionalId));
              const hasResp=ag.profResponsavelId&&ag.profResponsavelId!==ag.profissionalId;
              const respAg=hasResp?profissionais.find(p=>p.id===Number(ag.profResponsavelId)):null;
              return(<option key={ag.id} value={ag.id}>{brDate(ag.data)} {ag.horarioSessao} — {prAg?.nome||"—"}{respAg?" (resp: "+respAg.nome+")":""} {stInfo.icon}</option>);
            })}
          </select>

          {bloqueadoPorProf&&<div style={{marginTop:8,padding:"8px 12px",background:"#450a0a",border:"1px solid #ef444440",borderRadius:8,fontSize:12,color:"#fca5a5"}}>
            🔒 Esta sessão só pode ser evoluída por <b>{profResp?.nome}</b>. Você está logado como {profLogado?.nome}.
          </div>}

          {nova.agId&&!bloqueadoPorProf&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginTop:8}}>
            <div>
              <label>📅 Data Oficial</label>
              <div style={{padding:"7px 10px",background:"var(--sx)",borderRadius:8,fontSize:13,fontWeight:800,color:"#a78bfa",border:"1px solid #7c6af740",letterSpacing:.3}}>
                {brDate(nova.dataEvolucao)||"—"}
              </div>
              <div style={{fontSize:9,color:"var(--mt)",marginTop:3}}>
                {agSelecionado?.usarEvolucaoDiferente?"⚡ Data do horário alternativo":"📅 Data da sessão agendada"}
              </div>
            </div>
            <div>
              <label>🕐 Hora Oficial</label>
              <div style={{padding:"7px 10px",background:"var(--sx)",borderRadius:8,fontSize:16,fontWeight:900,color:agSelecionado?.usarEvolucaoDiferente?"#f59e0b":"#a78bfa",border:"1px solid "+(agSelecionado?.usarEvolucaoDiferente?"#f59e0b40":"#7c6af740"),letterSpacing:.5}}>
                {nova.horaEvolucao||"—"}
              </div>
              <div style={{fontSize:9,color:"var(--mt)",marginTop:3}}>
                {agSelecionado?.usarEvolucaoDiferente
                  ?<span style={{color:"#f59e0b",fontWeight:700}}>⚡ Horário de evolução diferente ativo</span>
                  :<span>🕐 Horário da sessão agendada</span>}
              </div>
            </div>
            <div>
              <label>👤 Prof. Responsável</label>
              <div style={{padding:"7px 10px",background:"var(--sx)",borderRadius:8,fontSize:12,fontWeight:700,color:profResp?"#a78bfa":"var(--mt)",border:"1px solid var(--cb)"}}>
                {profResp?profResp.nome:(profissionais.find(p=>p.id===Number(agSelecionado?.profissionalId))?.nome||"—")}
              </div>
            </div>
          </div>}

          {!nova.agId&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginTop:8}}>
            <div><label>📅 Data da Evolução</label><input type="date" value={nova.dataEvolucao||hoje_str} onChange={e=>setNova(n=>({...n,dataEvolucao:e.target.value}))}/></div>
            <div><label>🕐 Hora da Evolução</label><input type="time" value={nova.horaEvolucao||toTime(new Date().getHours()*60+new Date().getMinutes())} onChange={e=>setNova(n=>({...n,horaEvolucao:e.target.value}))}/></div>
          </div>}
        </div>

        {!bloqueadoPorProf&&<div className="stack">
          <div className="g2">
            <div><label>Próx. Sessão</label><input type="date" value={nova.proxData} onChange={e=>setNova(n=>({...n,proxData:e.target.value}))}/></div>
            <div><label>Especialidade da Evolução</label>
              <select value={nova.especialidade||espProf} onChange={e=>setNova(n=>({...n,especialidade:e.target.value}))}>
                <option value="">-- {espProf||"Selecione"} --</option>
                {ESPECIALIDADES_LIST.map(e=><option key={e} value={e}>{e}</option>)}
              </select>
            </div>
          </div>
          {modelosDisp.length>0&&<div className="section-box" style={{borderColor:"#7c6af730"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:7}}>
              <div className="section-title" style={{color:"#a78bfa",margin:0}}>🧩 Modelos — {espEvolucao}</div>
              <button className="btn secondary small" onClick={()=>setModeloAberto(o=>!o)}>{modeloAberto?"▲ Ocultar":"▼ Ver modelos"}</button>
            </div>
            {modeloAberto&&<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(195px,1fr))",gap:6}}>
              {modelosDisp.map(m=>(
                <button key={m.id} className="btn secondary" style={{textAlign:"left",fontSize:11,padding:"7px 9px"}}
                  onClick={()=>{setNova(n=>({...n,texto:m.texto}));setModeloAberto(false);}}>
                  <div style={{fontWeight:800,color:"#93c5fd",marginBottom:2}}>{m.titulo}</div>
                  <div style={{color:"var(--mt)",fontSize:9,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.texto.slice(0,80)}...</div>
                </button>
              ))}
            </div>}
          </div>}
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
              <label style={{margin:0}}>Evolução *</label>
              <MicBtn onResult={t=>setNova(n=>({...n,texto:n.texto+t}))} title="Ditar evolução por voz"/>
            </div>
            <textarea rows={7} value={nova.texto} onChange={e=>setNova(n=>({...n,texto:e.target.value}))} autoFocus placeholder="Descreva a evolução da sessão... ou clique em 🎙️ Ditar para falar" style={{fontFamily:"inherit",lineHeight:1.6}}/>
          </div>
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
              <label style={{margin:0}}>Conduta</label>
              <MicBtn onResult={t=>setNova(n=>({...n,conduta:n.conduta+t}))} title="Ditar conduta por voz"/>
            </div>
            <textarea rows={2} value={nova.conduta} onChange={e=>setNova(n=>({...n,conduta:e.target.value}))} placeholder="Próximas condutas... ou 🎙️ Ditar"/>
          </div>
          <div className="actions"><button className="btn primary" onClick={registrar}>📝 Registrar Evolução</button></div>
        </div>}
      </div>}
      {/* AGENDAMENTOS */}
      {tab==="agendamentos"&&<div className="stack">
        {agsPac.slice(0,30).map(ag=>{
          const ap=profissionais.find(p=>p.id===Number(ag.profissionalId));
          const st=STATUS_AG[ag.status]||{label:ag.status,color:"#64748b",icon:"📋"};
          const evAg=(prontuarios||[]).filter(p=>p.agId===ag.id);
          return(<div key={ag.id} className="card" style={{padding:11,borderLeft:"3px solid "+st.color}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:12}}>
              <span style={{fontWeight:800}}>{brDate(ag.data)} {ag.horarioSessao}</span>
              <div style={{display:"flex",gap:5,alignItems:"center"}}>
                {evAg.length>0&&<span style={{fontSize:9,padding:"1px 5px",borderRadius:20,background:"#90ee9025",color:"#90ee90",fontWeight:800}}>📝 Evoluído</span>}
                <span style={{padding:"1px 7px",borderRadius:20,fontSize:10,fontWeight:800,background:st.color+"20",color:st.color}}>{st.icon} {st.label}</span>
              </div>
            </div>
            <div style={{fontSize:11,color:"var(--mt)",marginTop:3}}>{ap?.nome} {ag.convenio?"— "+ag.convenio:""}</div>
          </div>);
        })}
      </div>}

      {/* FICHA DO PACIENTE */}
      {tab==="ficha"&&<FichaPacienteView paciente={paciente}/>}
    </div>
  </div>);
}

// ── Ficha do paciente com dados + termo de ciência ─────────────────────────
function FichaPacienteView({paciente}){
  const p=paciente;
  const NOME_CLINICA="Espaço Terapêutico Cinthia França";
  const info=(label,val)=>val?(<div style={{marginBottom:8}}>
    <span style={{fontSize:10,fontWeight:900,color:"var(--mt)",textTransform:"uppercase",letterSpacing:".5px"}}>{label}</span>
    <div style={{fontSize:13,color:"var(--tx)",marginTop:2,fontWeight:500}}>{val}</div>
  </div>):null;
  return(<div style={{display:"flex",flexDirection:"column",gap:14}}>
    <div className="section-box">
      <div className="section-title">👤 Dados Pessoais</div>
      <div className="g2">
        {info("Nome Completo",p.nome)}
        {info("CPF",maskCPF(p.cpf||""))}
        {info("Data de Nascimento",brDate(p.nascimento))}
        {info("Sexo",p.sexo==="M"?"Masculino":p.sexo==="F"?"Feminino":p.sexo)}
        {info("Estado Civil",p.estadoCivil)}
        {info("Celular",maskPhone(p.celular||""))}
        {info("E-mail",p.email)}
        {info("Profissão",p.profissao)}
      </div>
    </div>
    <div className="section-box">
      <div className="section-title">📍 Endereço</div>
      <div className="g2">
        {info("Logradouro",p.logradouro+(p.numero?" nº "+p.numero:""))}
        {info("Bairro",p.bairro)}
        {info("Cidade / UF",[p.cidade,p.estado].filter(Boolean).join(" / "))}
        {info("CEP",maskCEP(p.cep||""))}
      </div>
    </div>
    <div className="section-box">
      <div className="section-title">💊 Convênio</div>
      <div className="g2">
        {info("Convênio",p.convenio)}
        {info("Plano",p.plano)}
        {info("Nº Carteirinha",p.numCarteirinha)}
        {info("Validade Carteirinha",brDate(p.validadeCarteirinha||""))}
      </div>
    </div>
    {(p.resp1Nome||p.resp2Nome)&&<div className="section-box">
      <div className="section-title">👨‍👩‍👧 Responsáveis</div>
      <div className="g2">
        {info("Responsável 1",p.resp1Nome)}
        {info("CPF Resp. 1",maskCPF(p.resp1Cpf||""))}
        {info("WhatsApp Resp. 1",maskPhone(p.resp1Whatsapp||""))}
        {info("Grau (Resp. 1)",p.resp1Grau)}
        {p.resp2Nome&&info("Responsável 2",p.resp2Nome)}
        {p.resp2Nome&&info("WhatsApp Resp. 2",maskPhone(p.resp2Whatsapp||""))}
      </div>
    </div>}
    {p.infoImportantes&&<div style={{padding:"10px 13px",background:"#f59e0b15",borderRadius:10,border:"1px solid #f59e0b30"}}>
      <div style={{fontWeight:800,fontSize:11,color:"#f59e0b",marginBottom:4}}>⚠️ INFORMAÇÕES IMPORTANTES</div>
      <div style={{fontSize:12,color:"var(--tx)"}}>{p.infoImportantes}</div>
    </div>}
    {/* Termo de ciência */}
    <div className="section-box" style={{borderColor:"#7c6af740"}}>
      <div className="section-title" style={{color:"#a78bfa"}}>📃 Termo de Ciência</div>
      <div style={{fontSize:12,lineHeight:1.7,color:"var(--tx)"}}>
        Eu, <b>{p.resp1Nome||p.nome}</b>, responsável pelo(a) paciente <b>{p.nome}</b>, declaro que fui devidamente informado(a) sobre o tratamento proposto pelo <b>{NOME_CLINICA}</b>, incluindo objetivos, metodologias e periodicidade dos atendimentos. Estou ciente das políticas de cancelamento, reagendamento e faltas da clínica, bem como das condições de cobertura do meu convênio (<b>{p.convenio||"Particular"}</b>).
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginTop:16}}>
        <div style={{borderTop:"1px solid var(--cpb)",paddingTop:6,textAlign:"center"}}>
          <div style={{fontSize:10,color:"var(--mt)"}}>Assinatura do Responsável</div>
          <div style={{fontSize:11,color:"var(--tx)",marginTop:2}}>{p.resp1Nome||p.nome}</div>
        </div>
        <div style={{borderTop:"1px solid var(--cpb)",paddingTop:6,textAlign:"center"}}>
          <div style={{fontSize:10,color:"var(--mt)"}}>Data / Local</div>
          <div style={{fontSize:11,color:"var(--tx)",marginTop:2}}>___/___/______ · ___________</div>
        </div>
      </div>
    </div>
    {/* Documentos anexados */}
    {p.documentos?.length>0&&<div className="section-box">
      <div className="section-title">📎 Documentos Anexados</div>
      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
        {p.documentos.map((d,i)=><span key={i} className="file-chip">📄 {d.nome}</span>)}
      </div>
    </div>}
  </div>);
}



// ═══════════════════════════════════════════════════════════════════════════════
// PEDIDOS MÉDICOS
// ═══════════════════════════════════════════════════════════════════════════════
function PedidosMedicosPage({pedidos,setPedidos,pacientes,setChamados,showToast}){
  const [modal,setModal]=useState(false);const[editing,setEditing]=useState(null);
  const [filtro,setFiltro]=useState("");
  const blank={pacienteId:"",especialidade:"",dataEmissao:hoje_str,dataValidade:"",descricao:"",arquivos:[]};
  const [form,setForm]=useState(blank);
  const sf=p=>setForm(f=>({...f,...p}));

  useEffect(()=>{
    pedidos.forEach(p=>{
      if(!p.dataValidade)return;
      const d=diffDays(hoje_str,p.dataValidade);
      if(d<=7&&d>=0&&!p.chamadoAberto){
        const num=String(Math.floor(100000+Math.random()*900000));
        const pac=pacientes.find(x=>x.id===p.pacienteId);
        setChamados(c=>[...c,{id:Date.now()+p.id,numero:num,setor:"Faturamento",tipo:"pedido_medico",nome:pac?.nome||"—",descricao:"Pedido médico vencendo em "+d+" dias — "+p.especialidade,data:hoje_str,status:"aberto",resp:""}]);
        setPedidos(a=>a.map(x=>x.id===p.id?{...x,chamadoAberto:true}:x));
      }
    });
  },[pedidos]);

  const salvar=()=>{
    if(!form.pacienteId||!form.especialidade||!form.dataValidade)return alert("Preencha paciente, especialidade e validade");
    if(editing)setPedidos(a=>a.map(x=>x.id===editing.id?{...form,id:editing.id}:x));
    else setPedidos(a=>[...a,{...form,id:Date.now(),chamadoAberto:false}]);
    showToast("✅ Pedido salvo","ok");setModal(false);
  };
  const vencColor=d=>d===null?"#64748b":d<0?"#f87171":d<=7?"#f87171":d<=30?"#f59e0b":"#34d399";
  const vencLabel=d=>d===null?"—":d<0?"⚠️ VENCIDO há "+Math.abs(d)+" dias":d===0?"⚠️ VENCE HOJE":d<=7?"⚠️ Vence em "+d+" dias":d<=30?"⏳ Vence em "+d+" dias":"✅ Válido — "+d+" dias";
  const filtrados=pedidos.filter(p=>{const pac=pacientes.find(x=>x.id===p.pacienteId);return!filtro||(pac?.nome||"").toLowerCase().includes(filtro.toLowerCase());});
  const alertCount=pedidos.filter(p=>{if(!p.dataValidade)return false;return diffDays(hoje_str,p.dataValidade)<=30;}).length;
  return(<div className="page-wrap">
    <div className="page-head">
      <h1>🩻 Pedidos Médicos</h1>
      <div style={{display:"flex",gap:8}}>
        {alertCount>0&&<span style={{padding:"5px 10px",borderRadius:8,background:"#451a0340",color:"#f59e0b",fontSize:12,fontWeight:800}}>⚠️ {alertCount} vencendo</span>}
        <input value={filtro} onChange={e=>setFiltro(e.target.value)} placeholder="🔍 Paciente..." style={{width:180,fontSize:12}}/>
        <button className="btn primary" onClick={()=>{setEditing(null);setForm(blank);setModal(true);}}>+ Novo Pedido</button>
      </div>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:10}}>
      {filtrados.map(p=>{const pac=pacientes.find(x=>x.id===p.pacienteId);const d=p.dataValidade?diffDays(hoje_str,p.dataValidade):null;const vc=vencColor(d);
        return(<div key={p.id} className="card" style={{padding:14,borderLeft:"3px solid "+vc}}>
          <div style={{display:"flex",justifyContent:"space-between"}}>
            <div><div style={{fontWeight:800}}>{pac?.nome||"—"}</div><div style={{fontSize:11,color:espCor(p.especialidade),fontWeight:700,marginTop:2}}>{p.especialidade}</div></div>
            <div style={{display:"flex",gap:4,alignSelf:"flex-start"}}>
              <button className="btn secondary small" onClick={()=>{setEditing(p);setForm({...p});setModal(true);}}>✏️</button>
              <button className="btn danger small" onClick={()=>{if(confirm("Excluir?"))setPedidos(a=>a.filter(x=>x.id!==p.id));}}>🗑️</button>
            </div>
          </div>
          <div style={{marginTop:8,fontSize:11,display:"flex",gap:10}}>
            <span style={{color:"var(--mt)"}}>📅 Emissão: {brDate(p.dataEmissao)}</span>
            <span style={{color:"var(--mt)"}}>⏰ Validade: {brDate(p.dataValidade)}</span>
          </div>
          {d!==null&&<div style={{marginTop:5,fontSize:11,fontWeight:800,color:vc}}>{vencLabel(d)}</div>}
          {p.descricao&&<div style={{fontSize:11,color:"var(--mt)",marginTop:4}}>{p.descricao}</div>}
          {p.arquivos?.length>0&&<div style={{marginTop:5}}>{p.arquivos.map((f,i)=><span key={i} className="file-chip">📄 {f.nome}</span>)}</div>}
          {/* WA Pedido Médico */}
          {(pac?.resp1Whatsapp||pac?.celular)&&(()=>{
            const waNum=pac.resp1Whatsapp||pac.celular;
            const esp=p.especialidade||"";
            let msg="*PEDIDO MÉDICO — "+esp.toUpperCase()+"*\n*Espaço Terapêutico Cinthia França*\n\n";
            msg+="👤 *Paciente:* "+(pac?.nome||"—")+"\n🩺 *Especialidade:* "+esp+"\n";
            msg+="📅 *Emissão:* "+brDate(p.dataEmissao)+" · Validade: "+brDate(p.dataValidade)+"\n\n";
            if(d!==null&&d<=30)msg+="⚠️ *ATENÇÃO:* Pedido "+(d<0?"VENCIDO":"vencendo em "+d+" dias")+" — solicite renovação.\n\n";
            msg+="Por favor, providencie a renovação do pedido médico com:\n• CID-10 correspondente\n• Quantidade de sessões\n• Assinatura e CRM/CRO/CRP do médico.";
            const url="https://wa.me/55"+rawD(waNum)+"?text="+encodeURIComponent(msg);
            return <a key="wa" href={url} target="_blank" rel="noreferrer" className="btn" style={{marginTop:8,background:"#25d366",color:"#fff",textDecoration:"none",display:"inline-flex",alignItems:"center",gap:5,fontSize:11}}>🩺 WhatsApp Pedido Médico</a>;
          })()}
        </div>);
      })}
      {filtrados.length===0&&<div style={{gridColumn:"1/-1",textAlign:"center",padding:"40px",color:"var(--mt)"}}>Nenhum pedido.</div>}
    </div>
    {modal&&<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&setModal(false)}>
      <div className="modal" style={{width:560}}>
        <div className="modal-head"><h2>{editing?"✏️ Editar":"+ Novo"} Pedido Médico</h2><button className="icon-btn" onClick={()=>setModal(false)}>×</button></div>
        <div className="stack">
          <div><label>Paciente *</label><select value={form.pacienteId} onChange={e=>sf({pacienteId:Number(e.target.value)})}><option value="">Selecione...</option>{pacientes.map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}</select></div>
          <div><label>Especialidade *</label><select value={form.especialidade} onChange={e=>sf({especialidade:e.target.value})}><option value="">Selecione</option>{ESPECIALIDADES_LIST.map(e=><option key={e}>{e}</option>)}</select></div>
          <div className="g2"><div><label>Data de Emissão</label><input type="date" value={form.dataEmissao} onChange={e=>sf({dataEmissao:e.target.value})}/></div><div><label>Data de Validade *</label><input type="date" value={form.dataValidade} onChange={e=>sf({dataValidade:e.target.value})}/></div></div>
          <div><label>Descrição</label><textarea rows={2} value={form.descricao} onChange={e=>sf({descricao:e.target.value})}/></div>
          <UploadField label="Arquivos do Pedido" value={form.arquivos||[]} onChange={fs=>sf({arquivos:fs})}/>
          <div className="actions"><button className="btn secondary" onClick={()=>setModal(false)}>Cancelar</button><button className="btn primary" onClick={salvar}>Salvar</button></div>
        </div>
      </div>
    </div>}
  </div>);
}

// ═══════════════════════════════════════════════════════════════════════════════
// PACIENTES PAGE
// ═══════════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════════
// COMISSÃO DE PRONTUÁRIOS
// ═══════════════════════════════════════════════════════════════════════════════
function ComissaoProntuariosModal({prontuarios,setProntuarios,pacientes,profissionais,agenda,onClose}){
  const [filaIdx,setFilaIdx]=useState(0);
  const [analisando,setAnalisando]=useState(false);
  const [resultados,setResultados]=useState({});  // {evId: {status,erros,sugestoes,loading}}
  const [filtroStatus,setFiltroStatus]=useState("todos");
  const [filtroProf,setFiltroProf]=useState("");
  const [busca,setBusca]=useState("");

  // Fila: todas as evoluções ordenadas por data desc
  const fila=prontuarios
    .filter(ev=>{
      const pac=pacientes.find(p=>p.id===Number(ev.pacienteId));
      const prof=profissionais.find(p=>p.id===Number(ev.profId));
      const matchBusca=!busca||(pac?.nome||"").toLowerCase().includes(busca.toLowerCase())||(prof?.nome||"").toLowerCase().includes(busca.toLowerCase());
      const matchProf=!filtroProf||String(ev.profId)===filtroProf;
      const r=resultados[ev.id];
      const matchStatus=filtroStatus==="todos"||
        (filtroStatus==="pendente"&&!r)||
        (filtroStatus==="ok"&&r?.status==="ok")||
        (filtroStatus==="alerta"&&r?.status==="alerta")||
        (filtroStatus==="erro"&&r?.status==="erro");
      return matchBusca&&matchProf&&matchStatus;
    })
    .sort((a,b)=>b.dataEvolucao?.localeCompare(a.dataEvolucao||"")||0);

  const evAtual=fila[filaIdx];
  const pacAtual=evAtual?pacientes.find(p=>p.id===Number(evAtual.pacienteId)):null;
  const profAtual=evAtual?profissionais.find(p=>p.id===Number(evAtual.profId)):null;
  const agRef=evAtual?.agId?(agenda||[]).find(a=>a.id===evAtual.agId):null;

  const analisarEvolucao=async(ev)=>{
    if(!ev)return;
    setResultados(r=>({...r,[ev.id]:{loading:true}}));
    try{
      const pac=pacientes.find(p=>p.id===Number(ev.pacienteId));
      const prof=profissionais.find(p=>p.id===Number(ev.profId));
      const ag=ev.agId?(agenda||[]).find(a=>a.id===ev.agId):null;
      const prompt="Você é um auditor especializado em prontuários clínicos de saúde. Analise a evolução abaixo e retorne APENAS um JSON no formato especificado.\n\n"+"DADOS DA EVOLUÇÃO:\n"+"- Paciente: "+(pac?.nome||"Não informado")+"\n"+"- Profissional: "+(prof?.nome||"Não informado")+" ("+(ev.especialidade||"especialidade não informada")+")\n"+"- Data: "+brDate(ev.dataEvolucao||"")+" às "+(ev.horaEvolucao||"horário não informado")+"\n"+"- Sessão vinculada: "+(ag?brDate(ag.data)+" "+ag.horarioSessao:"Não vinculada")+"\n"+"- Texto da evolução: \""+( ev.texto||"")+"\"\n"+"- Conduta registrada: \""+(ev.conduta||"")+"\"\n\n"+"Verifique e retorne SOMENTE este JSON (sem markdown, sem texto fora do JSON):\n"+'{"status":"ok ou alerta ou erro","pontuacao":0,"erros":["lista de erros críticos"],"alertas":["lista de pontos de atenção"],"sugestoes":["sugestões de melhoria"],"resumo":"resumo em 1 frase da qualidade"}'+"\n\nCritérios: verifique se o texto é clinicamente adequado para a especialidade, se tem objetividade, se registra o que foi trabalhado, se conduta está coerente, se há ausência de dados importantes, linguagem inadequada, informações inconsistentes ou incompletas."

      const resp=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,messages:[{role:"user",content:prompt}]})
      });
      const data=await resp.json();
      const txt=(data.content||[]).filter(b=>b.type==="text").map(b=>b.text).join("");
      let parsed;
      try{
        const clean=txt.replace(/```json|```/g,"").trim();
        parsed=JSON.parse(clean);
      }catch{
        parsed={status:"alerta",pontuacao:50,erros:[],alertas:["Não foi possível analisar automaticamente."],sugestoes:[],resumo:"Análise indisponível."};
      }
      setResultados(r=>({...r,[ev.id]:{...parsed,loading:false,analisadoEm:new Date().toLocaleString("pt-BR")}}));
    }catch(e){
      setResultados(r=>({...r,[ev.id]:{status:"alerta",pontuacao:0,erros:[],alertas:["Erro ao conectar com IA: "+e.message],sugestoes:[],resumo:"Falha na análise.",loading:false}}));
    }
  };

  const analisarTodos=async()=>{
    setAnalisando(true);
    for(const ev of fila.slice(0,20)){
      if(!resultados[ev.id]||resultados[ev.id].loading===false&&!resultados[ev.id].status){
        await analisarEvolucao(ev);
        await new Promise(res=>setTimeout(res,400));
      }
    }
    setAnalisando(false);
  };

  const statusBadge=(s)=>{
    if(!s)return<span style={{padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:800,background:"var(--sx)",color:"var(--mt)"}}>⏳ Pendente</span>;
    if(s==="ok")return<span style={{padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:800,background:"#22c55e20",color:"#22c55e"}}>✅ OK</span>;
    if(s==="alerta")return<span style={{padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:800,background:"#f59e0b20",color:"#f59e0b"}}>⚠️ Alerta</span>;
    if(s==="erro")return<span style={{padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:800,background:"#ef444420",color:"#ef4444"}}>❌ Erro</span>;
    return null;
  };

  const statsCounts={ok:0,alerta:0,erro:0,pendente:0};
  prontuarios.forEach(ev=>{
    const r=resultados[ev.id];
    if(!r||r.loading)statsCounts.pendente++;
    else if(r.status==="ok")statsCounts.ok++;
    else if(r.status==="alerta")statsCounts.alerta++;
    else if(r.status==="erro")statsCounts.erro++;
  });

  const profsList=[...new Map(prontuarios.map(ev=>[ev.profId,profissionais.find(p=>p.id===Number(ev.profId))])).values()].filter(Boolean);

  return(<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&onClose()}>
    <div className="modal" style={{width:1100,maxHeight:"95vh",overflowY:"auto",display:"grid",gridTemplateColumns:"340px 1fr",gap:0,padding:0}}>

      {/* ── PAINEL ESQUERDO: Fila ── */}
      <div style={{borderRight:"1px solid var(--sc)",display:"flex",flexDirection:"column",height:"95vh",maxHeight:"95vh"}}>
        <div style={{padding:"14px 16px",borderBottom:"1px solid var(--sc)",background:"var(--sx)"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <h2 style={{margin:0,fontSize:15,fontWeight:900}}>🏛️ Comissão de Prontuários</h2>
            <button className="icon-btn" onClick={onClose}>×</button>
          </div>
          {/* KPIs */}
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:4,marginBottom:10}}>
            {[["⏳",statsCounts.pendente,"#64748b","Pendente"],["✅",statsCounts.ok,"#22c55e","OK"],["⚠️",statsCounts.alerta,"#f59e0b","Alerta"],["❌",statsCounts.erro,"#ef4444","Erro"]].map(([ic,n,cor,lb])=>(
              <div key={lb} style={{background:cor+"15",border:"1px solid "+cor+"30",borderRadius:8,padding:"5px 6px",textAlign:"center"}}>
                <div style={{fontSize:15,fontWeight:900,color:cor}}>{n}</div>
                <div style={{fontSize:8,color:"var(--mt)",fontWeight:700}}>{lb.toUpperCase()}</div>
              </div>
            ))}
          </div>
          {/* Filtros */}
          <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="🔍 Paciente ou profissional..." style={{width:"100%",fontSize:11,marginBottom:6}}/>
          <div style={{display:"flex",gap:4,marginBottom:6,flexWrap:"wrap"}}>
            {["todos","pendente","ok","alerta","erro"].map(s=>(
              <button key={s} onClick={()=>{setFiltroStatus(s);setFilaIdx(0);}} style={{padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:800,border:"1.5px solid "+(filtroStatus===s?"#7c6af7":"var(--cpb)"),background:filtroStatus===s?"#7c6af720":"transparent",color:filtroStatus===s?"#a78bfa":"var(--mt)",cursor:"pointer",textTransform:"capitalize"}}>{s}</button>
            ))}
          </div>
          <select value={filtroProf} onChange={e=>{setFiltroProf(e.target.value);setFilaIdx(0);}} style={{width:"100%",fontSize:11,marginBottom:8}}>
            <option value="">Todos os profissionais</option>
            {profsList.map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}
          </select>
          <button className="btn primary" style={{width:"100%",fontSize:11}} onClick={analisarTodos} disabled={analisando}>
            {analisando?"⏳ Analisando...":"🤖 Analisar todos com IA (até 20)"}
          </button>
        </div>
        {/* Lista de evoluções */}
        <div style={{overflowY:"auto",flex:1}}>
          {fila.length===0&&<div style={{padding:24,textAlign:"center",color:"var(--mt)",fontSize:12}}>Nenhuma evolução encontrada.</div>}
          {fila.map((ev,i)=>{
            const pac=pacientes.find(p=>p.id===Number(ev.pacienteId));
            const prof=profissionais.find(p=>p.id===Number(ev.profId));
            const r=resultados[ev.id];
            const isAtiva=i===filaIdx;
            return(<div key={ev.id} onClick={()=>setFilaIdx(i)}
              style={{padding:"10px 14px",borderBottom:"1px solid var(--sc)",cursor:"pointer",background:isAtiva?"var(--na)":"",borderLeft:isAtiva?"3px solid #7c6af7":"3px solid transparent",transition:".1s"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:6}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:800,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{pac?.nome||"—"}</div>
                  <div style={{fontSize:10,color:"var(--mt)",marginTop:1}}>{prof?.nome||"—"} · {brDate(ev.dataEvolucao||"")} {ev.horaEvolucao}</div>
                  <div style={{fontSize:10,color:espCor(ev.especialidade||""),fontWeight:700,marginTop:1}}>{ev.especialidade||"—"}</div>
                </div>
                <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:3,flexShrink:0}}>
                  {r?.loading?<span style={{fontSize:10,color:"#a78bfa"}}>⏳</span>:statusBadge(r?.status)}
                  {r?.pontuacao!=null&&!r.loading&&<span style={{fontSize:9,color:"var(--mt)",fontWeight:700}}>{r.pontuacao}/100</span>}
                </div>
              </div>
              {r?.resumo&&!r.loading&&<div style={{fontSize:9,color:"var(--mt)",marginTop:4,fontStyle:"italic",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.resumo}</div>}
            </div>);
          })}
        </div>
      </div>

      {/* ── PAINEL DIREITO: Detalhe + IA ── */}
      <div style={{display:"flex",flexDirection:"column",height:"95vh",maxHeight:"95vh",overflow:"hidden"}}>
        {!evAtual
          ?<div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100%",color:"var(--mt)",fontSize:13}}>Selecione uma evolução na fila</div>
          :<div style={{overflowY:"auto",flex:1,padding:20}}>
            {/* Cabeçalho da evolução */}
            <div style={{background:"var(--sx)",borderRadius:12,padding:14,marginBottom:14,border:"1px solid var(--cb)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:8}}>
                <div>
                  <div style={{fontWeight:900,fontSize:16}}>{pacAtual?.nome||"—"}</div>
                  <div style={{fontSize:12,color:"var(--mt)",marginTop:2}}>{pacAtual?.convenio||"Particular"} · {pacAtual?.nascimento?brDate(pacAtual.nascimento):""}</div>
                </div>
                <div style={{textAlign:"right"}}>
                  <div style={{fontWeight:800,color:"#a78bfa",fontSize:13}}>{brDate(evAtual.dataEvolucao||"")} às {evAtual.horaEvolucao||"—"}</div>
                  <div style={{fontSize:11,color:espCor(evAtual.especialidade||""),fontWeight:700,marginTop:2}}>{profAtual?.nome||"—"} · {evAtual.especialidade||"—"}</div>
                  {agRef&&<div style={{fontSize:10,color:"var(--mt)",marginTop:1}}>📅 Sessão: {brDate(agRef.data)} {agRef.horarioSessao}</div>}
                </div>
              </div>
            </div>

            {/* Texto da evolução */}
            <div style={{marginBottom:14}}>
              <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",letterSpacing:.5,marginBottom:6}}>Texto da Evolução</div>
              <div style={{background:"var(--card)",border:"1px solid var(--cb)",borderRadius:10,padding:14,fontSize:13,lineHeight:1.7,whiteSpace:"pre-wrap",color:"var(--tx)"}}>{evAtual.texto||<span style={{color:"var(--mt)",fontStyle:"italic"}}>Sem texto registrado.</span>}</div>
            </div>
            {evAtual.conduta&&<div style={{marginBottom:14}}>
              <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",letterSpacing:.5,marginBottom:6}}>Conduta</div>
              <div style={{background:"var(--card)",border:"1px solid var(--cb)",borderRadius:10,padding:12,fontSize:12,color:"var(--tx)"}}>{evAtual.conduta}</div>
            </div>}

            {/* Botão analisar */}
            {!resultados[evAtual.id]&&<button className="btn primary" style={{width:"100%",marginBottom:14}} onClick={()=>analisarEvolucao(evAtual)}>
              🤖 Analisar esta evolução com IA
            </button>}
            {resultados[evAtual.id]?.loading&&<div style={{textAlign:"center",padding:"20px",color:"#a78bfa",fontSize:13}}>
              ⏳ Analisando com IA...
            </div>}

            {/* Resultado da análise */}
            {resultados[evAtual.id]&&!resultados[evAtual.id].loading&&(()=>{
              const r=resultados[evAtual.id];
              const corStatus=r.status==="ok"?"#22c55e":r.status==="alerta"?"#f59e0b":"#ef4444";
              return(<div>
                {/* Score */}
                <div style={{background:corStatus+"10",border:"1px solid "+corStatus+"30",borderRadius:12,padding:16,marginBottom:12}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                    <div style={{fontWeight:900,fontSize:14,color:corStatus}}>{statusBadge(r.status)} Pontuação: {r.pontuacao}/100</div>
                    <div style={{fontSize:10,color:"var(--mt)"}}>Analisado: {r.analisadoEm}</div>
                  </div>
                  <div style={{height:8,borderRadius:4,background:"var(--sx)",overflow:"hidden"}}>
                    <div style={{width:r.pontuacao+"%",height:"100%",background:corStatus,borderRadius:4,transition:"width .5s"}}/>
                  </div>
                  {r.resumo&&<div style={{marginTop:8,fontSize:12,fontStyle:"italic",color:"var(--mt)"}}>{r.resumo}</div>}
                </div>

                {/* Erros */}
                {r.erros?.length>0&&<div style={{marginBottom:10}}>
                  <div style={{fontSize:11,fontWeight:900,color:"#ef4444",marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>❌ Erros críticos</div>
                  {r.erros.map((e,i)=><div key={i} style={{padding:"6px 10px",background:"#450a0a",border:"1px solid #ef444430",borderRadius:7,fontSize:12,color:"#fca5a5",marginBottom:4}}>• {e}</div>)}
                </div>}

                {/* Alertas */}
                {r.alertas?.length>0&&<div style={{marginBottom:10}}>
                  <div style={{fontSize:11,fontWeight:900,color:"#f59e0b",marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>⚠️ Pontos de atenção</div>
                  {r.alertas.map((a,i)=><div key={i} style={{padding:"6px 10px",background:"#451a03",border:"1px solid #f59e0b30",borderRadius:7,fontSize:12,color:"#fcd34d",marginBottom:4}}>• {a}</div>)}
                </div>}

                {/* Sugestões */}
                {r.sugestoes?.length>0&&<div style={{marginBottom:10}}>
                  <div style={{fontSize:11,fontWeight:900,color:"#a78bfa",marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>💡 Sugestões</div>
                  {r.sugestoes.map((s,i)=><div key={i} style={{padding:"6px 10px",background:"#0c1a2e",border:"1px solid #7c6af730",borderRadius:7,fontSize:12,color:"#93c5fd",marginBottom:4}}>• {s}</div>)}
                </div>}

                <button className="btn secondary" style={{width:"100%",marginTop:4,fontSize:11}} onClick={()=>analisarEvolucao(evAtual)}>🔄 Reanalisar</button>
              </div>);
            })()}

            {/* Navegação */}
            <div style={{display:"flex",justifyContent:"space-between",marginTop:16,paddingTop:12,borderTop:"1px solid var(--sc)"}}>
              <button className="btn secondary" disabled={filaIdx===0} onClick={()=>setFilaIdx(i=>i-1)}>← Anterior</button>
              <span style={{fontSize:11,color:"var(--mt)",alignSelf:"center"}}>{filaIdx+1} / {fila.length}</span>
              <button className="btn secondary" disabled={filaIdx>=fila.length-1} onClick={()=>setFilaIdx(i=>i+1)}>Próxima →</button>
            </div>
          </div>}
      </div>
    </div>
  </div>);
}

function PacientesPage({pacientes,setPacientes,setChamados,showToast,auth,agenda,profissionais,procedimentos,prontuarios,setProntuarios,templatePaciente}){
  const [search,setSearch]=useState("");const[modal,setModal]=useState(false);const[editing,setEditing]=useState(null);const[prontuarioPac,setProntuarioPac]=useState(null);const[contratoPac,setContratoPac]=useState(null);const[viewPac,setViewPac]=useState("grid");const[comissaoOpen,setComissaoOpen]=useState(false);
  const podeEditar=PODE_PACIENTE(auth.role);
  const salvar=p=>{if(editing){setPacientes(a=>a.map(x=>x.id===editing.id?p:x));showToast("✅ Atualizado","ok");}else{setPacientes(a=>[...a,p]);const num=String(Math.floor(100000+Math.random()*900000));setChamados(c=>[...c,{id:Date.now(),numero:num,setor:"Faturamento",tipo:"novo_paciente",nome:p.nome,descricao:"Novo cadastro: "+p.nome+" — "+p.convenio,data:hoje_str,status:"aberto",resp:""}]);showToast("✅ Salvo · 📨 #"+num+" → Faturamento","ok");}setModal(false);};
  const del=id=>{if(!podeEditar)return showToast("❌ Sem permissão","err");if(!confirm("Excluir paciente?"))return;setPacientes(a=>a.filter(x=>x.id!==id));showToast("🗑️ Removido","err");};
  const filtered=pacientes.filter(p=>p.nome.toLowerCase().includes(search.toLowerCase())||(p.cpf||"").includes(rawD(search)));
  return(<div className="page-wrap">
    <div className="page-head"><h1>👤 Pacientes</h1><div style={{display:"flex",gap:6,alignItems:"center"}}>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Nome ou CPF..." style={{width:200,fontSize:12}}/>
      <button onClick={()=>setViewPac(v=>v==="grid"?"list":"grid")} style={{background:"transparent",border:"1.5px solid var(--cpb)",borderRadius:8,padding:"5px 10px",cursor:"pointer",fontSize:12,color:"var(--mt)",whiteSpace:"nowrap"}}>{viewPac==="grid"?"☰ Lista":"⊞ Grade"}</button>
      <button className="btn secondary" onClick={()=>setComissaoOpen(true)} style={{borderColor:"#a78bfa50",color:"#a78bfa",background:"#a78bfa08",whiteSpace:"nowrap"}}>🏛️ Comissão de Prontuários{(prontuarios||[]).length>0&&<span style={{marginLeft:5,background:"#a78bfa",color:"#fff",borderRadius:20,padding:"0 6px",fontSize:10,fontWeight:800}}>{(prontuarios||[]).length}</span>}</button>
      {podeEditar&&<button className="btn primary" onClick={()=>{setEditing(null);setModal(true);}}>+ Novo Paciente</button>}
    </div></div>
    {viewPac==="list"&&<div style={{borderRadius:10,overflow:"hidden",border:"1px solid var(--sc)",marginBottom:10}}>
      <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 90px",background:"var(--sx)",padding:"7px 12px",fontSize:10,fontWeight:900,color:"var(--mt)",letterSpacing:.5,gap:8}}><div>PACIENTE</div><div>CONVÊNIO</div><div>CELULAR</div><div>CIDADE</div><div></div></div>
      {filtered.map(p=>{const nEv=(prontuarios||[]).filter(pr=>pr.pacienteId===p.id).length;const cor=p.convenio==="Particular"?"#4ade80":"#a78bfa";
        return(<div key={p.id} style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 90px",padding:"9px 12px",borderTop:"1px solid var(--sc)",gap:8,alignItems:"center",cursor:podeEditar?"pointer":"default"}} onClick={()=>podeEditar&&(setEditing(p),setModal(true))}>
          <div><div style={{fontWeight:800,fontSize:12}}>{p.nome}</div><div style={{fontSize:10,color:"var(--mt)"}}>{p.nascimento&&brDate(p.nascimento)}{p.cpf&&" · "+maskCPF(p.cpf)}</div></div>
          <div><span style={{padding:"1px 6px",borderRadius:20,fontSize:10,fontWeight:800,background:cor+"20",color:cor}}>{p.convenio||"Particular"}</span></div>
          <div style={{fontSize:11,color:"var(--mt)"}}>{maskPhone(p.celular||"")}</div>
          <div style={{fontSize:11,color:"var(--mt)"}}>{p.cidade||"—"}</div>
          <div style={{display:"flex",gap:4}} onClick={e=>e.stopPropagation()}>
            <button className="btn secondary small" title={"Prontuário"+(nEv>0?" ("+nEv+")":"")} onClick={()=>setProntuarioPac(p)}>📋</button>
            <button className="btn secondary small" title="Contrato" style={{color:"#f59e0b",borderColor:"#f59e0b50"}} onClick={()=>setContratoPac(p)}>📄</button>
          </div>
        </div>);})}
      {filtered.length===0&&<div style={{padding:"24px",textAlign:"center",color:"var(--mt)"}}>Nenhum paciente encontrado.</div>}
    </div>}
    <div style={{display:viewPac==="grid"?"grid":"none",gridTemplateColumns:"repeat(auto-fill,minmax(290px,1fr))",gap:10}}>
      {filtered.map(p=>{const nEv=(prontuarios||[]).filter(pr=>pr.pacienteId===p.id).length;const cor=p.convenio==="Particular"?"#4ade80":"#a78bfa";
        return(<div key={p.id} className="card" style={{padding:13,borderLeft:"3px solid "+cor}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontWeight:800,fontSize:14}}>{p.nome}</div>
              <div className="muted" style={{fontSize:11,marginTop:2}}>{maskCPF(p.cpf||"")} · {p.nascimento?brDate(p.nascimento):"—"}</div>
            </div>
            <div style={{display:"flex",gap:4,flexShrink:0,alignItems:"center"}}>
              <span style={{padding:"2px 7px",borderRadius:20,fontSize:10,fontWeight:800,background:"var(--sx)",color:cor}}>{p.convenio||"Particular"}</span>
              {podeEditar&&<button onClick={e=>{e.stopPropagation();setEditing(p);setModal(true);}} style={{background:"none",border:"1px solid var(--cpb)",borderRadius:6,color:"var(--mt)",cursor:"pointer",fontSize:11,padding:"2px 6px"}} title="Editar">✏️</button>}
              {podeEditar&&<button onClick={e=>{e.stopPropagation();del(p.id);}} style={{background:"none",border:"none",color:"var(--mt)",cursor:"pointer",fontSize:13}}>✕</button>}
            </div>
          </div>
          <div style={{marginTop:7,display:"flex",gap:10,fontSize:11,color:"var(--mt)",flexWrap:"wrap"}}>
            {p.celular&&<span>📞 {maskPhone(p.celular)}</span>}{p.cidade&&<span>📍 {p.cidade}</span>}
          </div>
          {p.infoImportantes&&<div style={{marginTop:5,fontSize:11,color:"#f59e0b",background:"#451a0310",padding:"3px 8px",borderRadius:5}}>⚠️ {p.infoImportantes.slice(0,70)}</div>}
          <div style={{marginTop:10,paddingTop:8,borderTop:"1px solid var(--db)",display:"flex",gap:5}}>
            <button className="btn secondary small" style={{flex:1,fontSize:11}} onClick={e=>{e.stopPropagation();setProntuarioPac(p);}}>📋 Prontuário {nEv>0&&<span style={{color:"#34d399"}}>({nEv})</span>}</button>
            <button className="btn secondary small" style={{fontSize:11,color:"#f59e0b",borderColor:"#f59e0b50",background:"#f59e0b08"}} onClick={e=>{e.stopPropagation();setContratoPac(p);}}>📄 Contrato</button>
          </div>
        </div>);
      })}
      {filtered.length===0&&<div style={{gridColumn:"1/-1",textAlign:"center",padding:"40px",color:"var(--mt)"}}>Nenhum paciente encontrado.</div>}
    </div>
    {modal&&podeEditar&&<CadastroPacienteModal editing={editing} pacientes={pacientes} onClose={()=>setModal(false)} onSave={salvar}/>}
    {prontuarioPac&&<ProntuarioModal paciente={prontuarioPac} agenda={agenda||[]} profissionais={profissionais||[]} procedimentos={procedimentos||[]} prontuarios={prontuarios||[]} setProntuarios={setProntuarios} modelosEv={MODELOS_EVOLUCAO} auth={auth} onClose={()=>setProntuarioPac(null)}/>}
    {comissaoOpen&&<ComissaoProntuariosModal prontuarios={prontuarios||[]} setProntuarios={setProntuarios} pacientes={pacientes||[]} profissionais={profissionais||[]} agenda={agenda||[]} onClose={()=>setComissaoOpen(false)}/>}
    {contratoPac&&<ContratoPdfModal pessoa={contratoPac} tipo="paciente" templateOverride={templatePaciente} onClose={()=>setContratoPac(null)}/>}
  </div>);
}

// ═══════════════════════════════════════════════════════════════════════════════
// PROFISSIONAIS PAGE
// ═══════════════════════════════════════════════════════════════════════════════

// ── Tipos de Usuário inline (dentro de Profissionais) ───────────────────────
function TiposUsuarioInline({auth}){
  const blankTipo={role:"",label:"",descr:"",permissoes:[]};
  const [editando,setEditando]=useState(null);
  const [tipos,setTipos]=useState(TIPOS_USUARIO_DEF);
  if(auth.role!=="administrador")return(<div style={{padding:24,textAlign:"center",color:"var(--mt)",background:"var(--sx)",borderRadius:10}}>🔒 Restrito ao Administrador.</div>);
  return(<div>
    <div style={{display:"flex",justifyContent:"flex-end",marginBottom:10}}>
      <button className="btn primary" style={{fontSize:12}} onClick={()=>setEditando({...blankTipo,_new:true})}>+ Novo Tipo</button>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:10}}>
      {tipos.map(t=>{const tCor=espCor(t.role==="administrador"?"Administrador":t.role==="faturamento_supervisor"?"Faturamento":t.role==="agendamento"?"Agendamento":t.role==="coordenador"||t.role==="coordenador_aba"?"Coordenação ABA":"Outro");return(
        <div key={t.role} className="card" style={{padding:14,borderLeft:"4px solid "+tCor}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div><div style={{fontWeight:900,fontSize:13}}>{t.label}</div>
              <code style={{fontSize:10,color:"#a78bfa",background:"var(--sx)",padding:"1px 5px",borderRadius:4}}>{t.role}</code>
            </div>
            <div style={{display:"flex",gap:5}}>
              <button className="btn secondary small" onClick={()=>setEditando({...t,_origRole:t.role})}>✏️</button>
              {t.role!=="administrador"&&<button className="btn danger small" onClick={()=>{if(confirm("Excluir "+t.label+"?"))setTipos(a=>a.filter(x=>x.role!==t.role));}}>🗑️</button>}
            </div>
          </div>
          <div style={{marginTop:8,fontSize:12,color:"var(--mt)",lineHeight:1.5}}>{t.descr}</div>
          <div style={{marginTop:7,display:"flex",flexWrap:"wrap",gap:4}}>
            {(t.permissoes||[]).map(p=><span key={p} style={{padding:"2px 7px",borderRadius:20,fontSize:10,fontWeight:800,background:"#7c6af720",color:"#a78bfa"}}>{p}</span>)}
          </div>
        </div>
      );})}
    </div>
    {editando&&<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&setEditando(null)}>
      <div className="modal" style={{width:500}}>
        <div className="modal-head"><h2>{editando._new?"+ Novo Tipo":"✏️ "+editando.label}</h2><button className="icon-btn" onClick={()=>setEditando(null)}>x</button></div>
        <div className="stack">
          <div className="g2">
            <div><label>Nome *</label><input value={editando.label||""} onChange={e=>setEditando(x=>({...x,label:e.target.value}))} autoFocus/></div>
            <div><label>Código (role) *</label><input value={editando.role||""} onChange={e=>setEditando(x=>({...x,role:e.target.value.toLowerCase().replace(/\s+/g,"_")}))} placeholder="ex: recepcionista"/></div>
          </div>
          <div><label>Descrição</label><textarea rows={3} value={editando.descr||""} onChange={e=>setEditando(x=>({...x,descr:e.target.value}))}/></div>
          <div className="actions">
            <button className="btn secondary" onClick={()=>setEditando(null)}>Cancelar</button>
            <button className="btn primary" onClick={()=>{
              if(!editando.label||!editando.role)return alert("Preencha nome e código");
              setTipos(a=>{
                const origRole=editando._origRole;
                if(origRole)return a.map(x=>x.role===origRole?{label:editando.label,role:editando.role,descr:editando.descr,permissoes:editando.permissoes||[]}:x);
                return[...a,{label:editando.label,role:editando.role,descr:editando.descr,permissoes:[]}];
              });
              setEditando(null);
            }}>Salvar</button>
          </div>
        </div>
      </div>
    </div>}
  </div>);
}


function TabelaRepassePanel({profissionais,setProfissionais,convenios,procedimentos,showToast}){
  const profsAtivos=profissionais.filter(p=>["profissional","coordenador","coordenador_aba"].includes(p.role));
  const [selPId,setSelPId]=useState(profsAtivos[0]?.id||"");
  const [showForm,setShowForm]=useState(false);
  const [editRow,setEditRow]=useState(null);
  const [subTab,setSubTab]=useState("tabela");
  const blankRow={convenio:"Particular",procedimentoId:"",valorTabela:0,pctRepasse:45};
  const [form,setForm]=useState({...blankRow});
  const sf=p=>setForm(f=>({...f,...p}));
  const profSel=profissionais.find(p=>Number(p.id)===Number(selPId));
  const tabela=profSel?.tabelaRepasse||[];
  const convsList=[...(convenios||[]).map(cv=>cv.nome||cv),"Particular"].filter((v,i,a)=>v&&a.indexOf(v)===i);
  const conveniosAceitos=profSel?.conveniosAtendidos||[];
  const saveRow=()=>{
    if(!form.convenio||!form.procedimentoId){showToast("Preencha convênio e procedimento","err");return;}
    const row={...form,id:editRow?.id||Date.now(),procedimentoId:Number(form.procedimentoId),valorTabela:Number(form.valorTabela),pctRepasse:Number(form.pctRepasse)};
    const novaTab=editRow?tabela.map(r=>r.id===editRow.id?row:r):[...tabela,row];
    setProfissionais(a=>a.map(p=>Number(p.id)===Number(selPId)?{...p,tabelaRepasse:novaTab}:p));
    showToast(editRow?"✅ Linha atualizada":"✅ Linha adicionada","ok");
    setEditRow(null);setForm({...blankRow});setShowForm(false);
  };
  const delRow=id=>{
    if(!confirm("Excluir esta linha?"))return;
    setProfissionais(a=>a.map(p=>Number(p.id)===Number(selPId)?{...p,tabelaRepasse:tabela.filter(r=>r.id!==id)}:p));
    showToast("Linha removida","ok");
  };
  const startEdit=row=>{setEditRow(row);setForm({...row});setShowForm(true);};
  const cancelForm=()=>{setEditRow(null);setForm({...blankRow});setShowForm(false);};
  const toggleConv=nome=>{
    const cur=profSel?.conveniosAtendidos||[];
    const novo=cur.includes(nome)?cur.filter(x=>x!==nome):[...cur,nome];
    setProfissionais(a=>a.map(p=>Number(p.id)===Number(selPId)?{...p,conveniosAtendidos:novo}:p));
  };
  return(<div>
    <div style={{background:"var(--card)",border:"1px solid var(--cb)",borderRadius:12,padding:"12px 16px",marginBottom:14,display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
      <span style={{fontSize:20}}>🩺</span>
      <select value={selPId} onChange={e=>{setSelPId(e.target.value);cancelForm();setSubTab("tabela");}} style={{flex:1,minWidth:220,fontSize:13,fontWeight:700}}>
        <option value="">— Selecione o profissional —</option>
        {profsAtivos.map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}
      </select>
      {profSel&&<div style={{fontSize:11,color:"var(--mt)"}}>
        {(profSel.especialidades||[]).join(" · ")}
        <span style={{color:"#a78bfa",fontWeight:700,marginLeft:8}}>Nível: {profSel.nivelRepasse||"Pleno"}</span>
      </div>}
      {profSel&&<div style={{marginLeft:"auto",display:"flex",gap:6,flexWrap:"wrap"}}>
        {conveniosAceitos.length>0?conveniosAceitos.map(cv=>(
          <span key={cv} style={{padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:800,background:"#0d9488",color:"#fff"}}>{cv}</span>
        )):<span style={{fontSize:11,color:"var(--mt)",fontStyle:"italic"}}>Nenhum convênio vinculado</span>}
      </div>}
    </div>
    {!profSel&&<div style={{padding:"48px",textAlign:"center",color:"var(--mt)",fontSize:13}}>👆 Selecione um profissional para configurar sua tabela de repasse.</div>}
    {profSel&&<>
      <div style={{display:"flex",gap:6,marginBottom:14}}>
        {[["tabela","📋 Tabela de Valores"],["convenios","🏥 Convênios Aceitos"]].map(([k,l])=>(
          <button key={k} className="btn tab-btn" onClick={()=>setSubTab(k)} style={{background:subTab===k?"var(--na)":"transparent",color:subTab===k?"#7c6af7":"var(--mt)",fontWeight:subTab===k?800:500,fontSize:12}}>
            {l}
            {k==="tabela"&&tabela.length>0&&<span style={{marginLeft:5,background:"#7c6af7",color:"#fff",borderRadius:20,padding:"0 5px",fontSize:10,fontWeight:800}}>{tabela.length}</span>}
            {k==="convenios"&&conveniosAceitos.length>0&&<span style={{marginLeft:5,background:"#0d9488",color:"#fff",borderRadius:20,padding:"0 5px",fontSize:10,fontWeight:800}}>{conveniosAceitos.length}</span>}
          </button>
        ))}
      </div>
      {subTab==="convenios"&&<div>
        <div style={{marginBottom:10,fontSize:12,color:"var(--mt)"}}>Selecione os convênios atendidos por <b style={{color:"var(--tx)"}}>{profSel.nome}</b>.</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:8}}>
          {["Particular",...(convenios||[]).map(cv=>cv.nome||cv)].filter((v,i,a)=>v&&a.indexOf(v)===i).map(cvNome=>{
            const on=conveniosAceitos.includes(cvNome);
            const cvObj=(convenios||[]).find(x=>(x.nome||x)===cvNome);
            const totalProcs=(cvObj?.procedimentos||[]).length;
            return(<div key={cvNome} onClick={()=>toggleConv(cvNome)} style={{padding:"10px 14px",borderRadius:10,cursor:"pointer",border:"2px solid "+(on?"#0d9488":"var(--cb)"),background:on?"#0d948818":"var(--card)",display:"flex",alignItems:"center",gap:8}}>
              <span style={{fontSize:18}}>{on?"✅":"⬜"}</span>
              <div>
                <div style={{fontWeight:800,fontSize:12,color:on?"#0d9488":"var(--tx)"}}>{cvNome}</div>
                <div style={{fontSize:10,color:"var(--mt)"}}>{cvNome==="Particular"?"Pagamento direto":totalProcs+" procedimento(s)"}</div>
              </div>
            </div>);
          })}
        </div>
        {convsList.length<=1&&<div className="muted" style={{padding:20,textAlign:"center"}}>Nenhum convênio cadastrado além de Particular.</div>}
      </div>}
      {subTab==="tabela"&&<>
        <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:12}}>
          <button className="btn primary" onClick={()=>{cancelForm();setShowForm(true);}}>+ Adicionar Linha</button>
          {tabela.length>0&&<span style={{fontSize:11,color:"var(--mt)"}}>{tabela.length} linha(s) · clique ✏️ para editar</span>}
        </div>
        {showForm&&<div className="card" style={{padding:16,marginBottom:14,borderLeft:"3px solid #7c6af7"}}>
          <div style={{fontWeight:800,fontSize:13,color:"#7c6af7",marginBottom:12}}>{editRow?"✏️ Editar linha":"+ Nova linha"} — {profSel.nome}</div>
          <div style={{display:"grid",gridTemplateColumns:"1.3fr 2fr 130px 105px auto",gap:10,alignItems:"end"}}>
            <div><label>Convênio *</label>
              <select value={form.convenio} onChange={e=>sf({convenio:e.target.value})}>
                <option value="Particular">Particular</option>
                {(convenios||[]).map(cv=><option key={cv.nome||cv}>{cv.nome||cv}</option>)}
              </select>
            </div>
            <div><label>Procedimento *</label>
              <select value={form.procedimentoId} onChange={e=>{
                const proc=(procedimentos||[]).find(p=>p.id===Number(e.target.value));
                sf({procedimentoId:Number(e.target.value),valorTabela:proc?.valor||0,pctRepasse:proc?.pctRepasse||45});
              }}>
                <option value="">— selecione —</option>
                {(procedimentos||[]).map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}
              </select>
            </div>
            <div><label>Valor Tabela (R$)</label>
              <input type="number" min={0} step={0.01} value={form.valorTabela} onChange={e=>sf({valorTabela:Number(e.target.value)})}/>
            </div>
            <div><label>% Repasse</label>
              <input type="number" min={0} max={100} step={1} value={form.pctRepasse} onChange={e=>sf({pctRepasse:Number(e.target.value)})}/>
            </div>
            <div style={{display:"flex",gap:5,paddingBottom:2}}>
              <button className="btn primary" onClick={saveRow}>💾</button>
              <button className="btn secondary" onClick={cancelForm}>✕</button>
            </div>
          </div>
          {form.procedimentoId&&<div style={{marginTop:10,padding:"8px 12px",background:"var(--sx)",borderRadius:8,fontSize:11,display:"flex",gap:20,flexWrap:"wrap"}}>
            <span>Repasse: <b style={{color:"#34d399"}}>R$ {((form.valorTabela||0)*(form.pctRepasse||0)/100).toFixed(2)}</b></span>
            <span>Clínica retém: <b style={{color:"#f59e0b"}}>R$ {((form.valorTabela||0)*(1-(form.pctRepasse||0)/100)).toFixed(2)}</b></span>
          </div>}
        </div>}
        {tabela.length===0&&!showForm&&<div style={{padding:"32px",textAlign:"center",color:"var(--mt)",border:"2px dashed var(--cb)",borderRadius:12,fontSize:13}}>
          Nenhuma linha cadastrada.<br/><span style={{fontSize:11}}>Clique em + Adicionar Linha para configurar os valores de repasse por convênio e procedimento.</span>
        </div>}
        {tabela.length>0&&(()=>{
          const grupos=[...new Set(tabela.map(r=>r.convenio))];
          return(<div style={{display:"flex",flexDirection:"column",gap:12}}>
            {grupos.map(grp=>{
              const rows=tabela.filter(r=>r.convenio===grp);
              const totVal=rows.reduce((s,r)=>s+(r.valorTabela||0),0);
              const totRep=rows.reduce((s,r)=>s+(r.valorTabela||0)*(r.pctRepasse||0)/100,0);
              return(<div key={grp} className="card" style={{overflow:"hidden"}}>
                <div style={{padding:"8px 14px",background:"var(--sx)",borderBottom:"1px solid var(--sc)",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div style={{display:"flex",gap:8,alignItems:"center"}}>
                    <span style={{padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:800,background:"#0d948820",color:"#0d9488"}}>🏥 {grp}</span>
                    <span style={{fontSize:11,color:"var(--mt)"}}>{rows.length} procedimento(s)</span>
                  </div>
                  <div style={{fontSize:11,fontWeight:800}}>
                    Tabela: <span style={{color:"#f59e0b"}}>{brl(totVal)}</span>
                    <span style={{margin:"0 8px",color:"var(--mt)"}}>/</span>
                    Repasse: <span style={{color:"#34d399"}}>{brl(totRep)}</span>
                  </div>
                </div>
                <div className="grid-header" style={{gridTemplateColumns:"2fr 110px 80px 110px 100px 68px"}}>
                  <div>Procedimento</div>
                  <div style={{textAlign:"right"}}>Valor Tabela</div>
                  <div style={{textAlign:"center"}}>% Rep.</div>
                  <div style={{textAlign:"right"}}>Repasse/Sessão</div>
                  <div style={{textAlign:"right"}}>Clínica Retém</div>
                  <div></div>
                </div>
                {rows.map((row,i)=>{
                  const proc=(procedimentos||[]).find(p=>p.id===Number(row.procedimentoId));
                  const rep=(row.valorTabela||0)*(row.pctRepasse||0)/100;
                  const cli=(row.valorTabela||0)-rep;
                  return(<div key={row.id} className="grid-row" style={{gridTemplateColumns:"2fr 110px 80px 110px 100px 68px",background:i%2?"var(--gr)":""}}>
                    <div>
                      <div style={{fontWeight:700,fontSize:12}}>{proc?.nome||"—"}</div>
                      {proc?.categoria&&<div style={{fontSize:9,color:"var(--mt)"}}>{proc.categoria}</div>}
                    </div>
                    <div style={{textAlign:"right",fontWeight:800,color:"#f59e0b",fontSize:12}}>{brl(row.valorTabela||0)}</div>
                    <div style={{textAlign:"center",fontWeight:900,color:"#a78bfa",fontSize:13}}>{row.pctRepasse||45}%</div>
                    <div style={{textAlign:"right",fontWeight:900,color:"#34d399",fontSize:12}}>{brl(rep)}</div>
                    <div style={{textAlign:"right",color:"var(--mt)",fontSize:11,fontWeight:700}}>{brl(cli)}</div>
                    <div style={{display:"flex",gap:3}}>
                      <button className="btn secondary small" onClick={()=>startEdit(row)}>✏️</button>
                      <button className="btn danger small" onClick={()=>delRow(row.id)}>🗑️</button>
                    </div>
                  </div>);
                })}
              </div>);
            })}
            <div style={{padding:"10px 14px",background:"var(--sx)",borderRadius:10,border:"1px solid var(--cb)",display:"flex",gap:24,fontWeight:900,fontSize:12,flexWrap:"wrap"}}>
              <span style={{color:"var(--mt)"}}>TOTAIS ({tabela.length} linhas)</span>
              <span>Tabela: <b style={{color:"#f59e0b"}}>{brl(tabela.reduce((s,r)=>s+(r.valorTabela||0),0))}</b></span>
              <span>Repasse: <b style={{color:"#34d399"}}>{brl(tabela.reduce((s,r)=>s+(r.valorTabela||0)*(r.pctRepasse||0)/100,0))}</b></span>
              <span>Clínica: <b style={{color:"#a78bfa"}}>{brl(tabela.reduce((s,r)=>s+(r.valorTabela||0)*(1-(r.pctRepasse||0)/100),0))}</b></span>
            </div>
          </div>);
        })()}
      </>}
      <div style={{marginTop:12,padding:"10px 14px",background:"#1e3a5f18",borderRadius:10,border:"1px solid #7c6af728",fontSize:11,color:"var(--mt)",lineHeight:1.7}}>
        💡 <b style={{color:"var(--tx)"}}>Base dos cálculos:</b> ao marcar sessão como <b>Atendido</b>, o sistema consulta esta tabela (convênio + procedimento do profissional) para gerar automaticamente a conta a pagar (CP). Se não houver linha específica, usa o % padrão do procedimento global.
      </div>
    </>}
  </div>);
}


function RecebimentosPanel({profissionais,procedimentos,convenios,lancamentos,agenda,pacientes,showToast}){
  const hoje_str=ymd(new Date());
  const profsAtivos=profissionais.filter(p=>["profissional","coordenador","coordenador_aba"].includes(p.role));
  const [selProfId,setSelProfId]=useState("");
  const [recebProf,setRecebProf]=useState(null);
  const [recebPeriodo,setRecebPeriodo]=useState("mes");
  const [recebIni,setRecebIni]=useState(()=>{const d=new Date();d.setDate(1);return ymd(d);});
  const [recebFim,setRecebFim]=useState(()=>ymd(new Date()));
  const aplicarRecebPeriodo=p=>{
    setRecebPeriodo(p);const d=new Date();
    if(p==="hoje"){setRecebIni(hoje_str);setRecebFim(hoje_str);}
    else if(p==="semana"){const x=new Date(d);x.setDate(d.getDate()-6);setRecebIni(ymd(x));setRecebFim(hoje_str);}
    else if(p==="mes"){const x=new Date(d);x.setDate(1);setRecebIni(ymd(x));setRecebFim(hoje_str);}
    else if(p==="trimestre"){const x=new Date(d);x.setMonth(d.getMonth()-3);setRecebIni(ymd(x));setRecebFim(hoje_str);}
  };
  const agReceb=(agenda||[]).filter(a=>{
    const repOk=STATUS_AG[a.status]?.repasse===true;
    const profOk=!selProfId||String(a.profissionalId)===String(selProfId);
    return repOk&&a.data>=recebIni&&a.data<=recebFim&&profOk;
  });
  const Pill=({v,l})=>(<button onClick={()=>aplicarRecebPeriodo(v)} style={{padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:800,border:"1.5px solid "+(recebPeriodo===v?"#7c6af7":"var(--cpb)"),background:recebPeriodo===v?"#7c6af715":"transparent",color:recebPeriodo===v?"#7c6af7":"var(--mt)",cursor:"pointer"}}>{l}</button>);
  const profsVisiveis=profsAtivos.filter(p=>!selProfId||String(p.id)===String(selProfId));

  const exportPDF=()=>{
    let html='<html><head><style>body{font-family:Arial,sans-serif;font-size:12px;margin:24px}h1{font-size:16px}h2{font-size:13px;color:#2563eb;border-bottom:1px solid #e2e8f0;padding-bottom:4px;margin-top:22px}p.sub{font-size:10px;color:#64748b;margin:0 0 14px}table{width:100%;border-collapse:collapse;margin-bottom:8px}th{background:#f1f5f9;text-align:left;padding:5px 8px;font-size:10px;font-weight:700;border:1px solid #e2e8f0}td{padding:5px 8px;border:1px solid #e2e8f0;font-size:11px}.tr2{background:#f8fafc}.tf td{font-weight:700;background:#f0fdf4}.r{text-align:right}</style></head><body>';
    html+='<h1>Repasse de Profissionais</h1><p class="sub">Período: '+brDate(recebIni)+' a '+brDate(recebFim)+'</p>';
    let gTot=0,gSess=0;
    profsVisiveis.forEach(prof=>{
      const aps=agReceb.filter(a=>Number(a.profissionalId)===Number(prof.id));
      if(!aps.length)return;
      const tB=aps.reduce((s,a)=>{const p=(procedimentos||[]).find(x=>x.id===Number(a.procedimentoId));return s+(p?.valor||0);},0);
      const tR=aps.reduce((s,a)=>{const p=(procedimentos||[]).find(x=>x.id===Number(a.procedimentoId));return s+(p?.valor||0)*((p?.pctRepasse||45)/100);},0);
      gTot+=tR;gSess+=aps.length;
      html+='<h2>'+prof.nome+' — '+(prof.especialidades||[''])[0]+'</h2>';
      html+='<table><thead><tr><th>Data</th><th>Horário</th><th>Paciente</th><th>Procedimento</th><th>Status</th><th class="r">Bruto</th><th class="r">%</th><th class="r">Repasse</th></tr></thead><tbody>';
      aps.sort((a,b)=>a.data.localeCompare(b.data)).forEach((ag,i)=>{
        const pN=(pacientes||[]).find(p=>p.id===Number(ag.pacienteId));
        const pr=(procedimentos||[]).find(p=>p.id===Number(ag.procedimentoId));
        const pct=(pr?.pctRepasse||45)/100;const vb=pr?.valor||0;
        html+='<tr'+(i%2?' class="tr2"':'')+'><td>'+brDate(ag.data)+'</td><td>'+ag.horarioSessao+'</td><td>'+(pN?.nome||'—')+'</td><td>'+(pr?.nome||'—')+'</td><td>'+(STATUS_AG[ag.status]?.label||ag.status)+'</td><td class="r">R$ '+vb.toFixed(2)+'</td><td class="r">'+Math.round(pct*100)+'%</td><td class="r">R$ '+(vb*pct).toFixed(2)+'</td></tr>';
      });
      html+='<tr class="tf"><td colspan="5">TOTAL ('+aps.length+' sessões)</td><td class="r">R$ '+tB.toFixed(2)+'</td><td></td><td class="r">R$ '+tR.toFixed(2)+'</td></tr></tbody></table>';
    });
    html+='<p style="margin-top:20px;font-size:12px;font-weight:700">TOTAL GERAL — '+gSess+' sessões — R$ '+gTot.toFixed(2)+'</p></body></html>';
    const w=window.open('','_blank');w.document.write(html);w.document.close();w.print();
  };

  const exportExcel=()=>{
    if(typeof XLSX==='undefined'){alert('SheetJS não carregado');return;}
    const wb=XLSX.utils.book_new();
    profsVisiveis.forEach(prof=>{
      const aps=agReceb.filter(a=>Number(a.profissionalId)===Number(prof.id));
      if(!aps.length)return;
      const rows=aps.sort((a,b)=>a.data.localeCompare(b.data)).map(ag=>{
        const pN=(pacientes||[]).find(p=>p.id===Number(ag.pacienteId));
        const pr=(procedimentos||[]).find(p=>p.id===Number(ag.procedimentoId));
        const pct=(pr?.pctRepasse||45)/100;const vb=pr?.valor||0;
        return{"Data":brDate(ag.data),"Horário":ag.horarioSessao,"Paciente":pN?.nome||"—","Procedimento":pr?.nome||"—","Status":STATUS_AG[ag.status]?.label||ag.status,"Valor Bruto":vb,"% Repasse":Math.round(pct*100),"Valor Repasse":Math.round(vb*pct*100)/100};
      });
      const sn=(profShort(prof.nome)||"Prof").replace(/[^a-zA-Z0-9 ]/g,"").slice(0,28);
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),sn);
    });
    XLSX.writeFile(wb,"Repasse_"+recebIni+"_"+recebFim+".xlsx");
  };

  const totSess=profsVisiveis.reduce((s,p)=>s+agReceb.filter(a=>Number(a.profissionalId)===Number(p.id)).length,0);
  const totRep=profsVisiveis.reduce((s,p)=>s+agReceb.filter(a=>Number(a.profissionalId)===Number(p.id)).reduce((sr,a)=>{const pr=(procedimentos||[]).find(x=>x.id===Number(a.procedimentoId));return sr+(pr?.valor||0)*((pr?.pctRepasse||45)/100);},0),0);
  const totBruto=profsVisiveis.reduce((s,p)=>s+agReceb.filter(a=>Number(a.profissionalId)===Number(p.id)).reduce((sr,a)=>{const pr=(procedimentos||[]).find(x=>x.id===Number(a.procedimentoId));return sr+(pr?.valor||0);},0),0);

  return(<div>
    {/* Filtros */}
    <div style={{background:"var(--card)",border:"1px solid var(--cb)",borderRadius:13,padding:"14px 16px",marginBottom:14,display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
      <Pill v="hoje" l="Hoje"/><Pill v="semana" l="7 dias"/><Pill v="mes" l="Mês"/><Pill v="trimestre" l="Trimestre"/>
      <div style={{width:1,height:24,background:"var(--sc)",flexShrink:0}}/>
      <div style={{display:"flex",alignItems:"center",gap:6}}>
        <label style={{color:"var(--mt)",fontWeight:700,fontSize:11}}>De</label>
        <input type="date" value={recebIni} onChange={e=>{setRecebIni(e.target.value);setRecebPeriodo("custom");}} style={{width:135,fontSize:12}}/>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:6}}>
        <label style={{color:"var(--mt)",fontWeight:700,fontSize:11}}>Até</label>
        <input type="date" value={recebFim} onChange={e=>{setRecebFim(e.target.value);setRecebPeriodo("custom");}} style={{width:135,fontSize:12}}/>
      </div>
      <div style={{width:1,height:24,background:"var(--sc)",flexShrink:0}}/>
      <select value={selProfId} onChange={e=>setSelProfId(e.target.value)} style={{minWidth:190,fontSize:12}}>
        <option value="">Todos os profissionais</option>
        {profsAtivos.map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}
      </select>
      <div style={{marginLeft:"auto",display:"flex",gap:6}}>
        <button className="btn secondary" style={{fontSize:11}} onClick={exportPDF}>🖨️ PDF</button>
        <button className="btn" style={{fontSize:11,background:"#166534",color:"#fff",border:"1px solid #166534"}} onClick={exportExcel}>📊 Excel</button>
      </div>
    </div>
    {/* KPIs */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:14}}>
      {[{label:"Sessões com repasse",val:totSess,color:"#a78bfa"},{label:"Receita bruta",val:brl(totBruto),color:"#f59e0b"},{label:"Total a repassar",val:brl(totRep),color:"#34d399"}].map(k=>(
        <div key={k.label} className="card" style={{padding:"12px 16px",borderTop:"3px solid "+k.color}}>
          <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:22,color:k.color,lineHeight:1}}>{k.val}</div>
          <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",letterSpacing:".5px",marginTop:4}}>{k.label}</div>
        </div>
      ))}
    </div>
    {/* Cards */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(260px,1fr))",gap:12,marginBottom:16}}>
      {profsVisiveis.map(prof=>{
        const aps=agReceb.filter(a=>Number(a.profissionalId)===Number(prof.id));
        const tRep=aps.reduce((s,a)=>{const pr=(procedimentos||[]).find(p=>p.id===Number(a.procedimentoId));return s+(pr?.valor||0)*((pr?.pctRepasse||45)/100);},0);
        const tBruto=aps.reduce((s,a)=>{const pr=(procedimentos||[]).find(p=>p.id===Number(a.procedimentoId));return s+(pr?.valor||0);},0);
        const pctMed=aps.length?Math.round(aps.reduce((s,a)=>{const p=(procedimentos||[]).find(x=>x.id===Number(a.procedimentoId));return s+(p?.pctRepasse||45);},0)/aps.length):45;
        const cor=espCor((prof.especialidades||[""])[0]);
        const lancProf=(lancamentos||[]).filter(l=>l.tipo==="CP"&&Number(l.profissionalId)===Number(prof.id)&&l.data>=recebIni&&l.data<=recebFim);
        const pago=lancProf.filter(l=>l.status==="pago").reduce((s,l)=>s+l.valor,0);
        const pendente=lancProf.filter(l=>l.status!=="pago").reduce((s,l)=>s+l.valor,0);
        return(<div key={prof.id} className="card" style={{padding:16,borderTop:"3px solid "+cor,cursor:"pointer"}} onClick={()=>setRecebProf(recebProf?.id===prof.id?null:prof)}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
            <div>
              <div style={{fontWeight:800,fontSize:13,color:cor}}>{prof.nome}</div>
              <div style={{fontSize:10,color:"var(--mt)",marginTop:2}}>{(prof.especialidades||[""])[0]} · {prof.nivelRepasse||"Pleno"}</div>
            </div>
            <div style={{textAlign:"right"}}>
              <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:22,color:"#34d399",lineHeight:1}}>{brl(tRep)}</div>
              <div style={{fontSize:9,color:"var(--mt)",marginTop:1}}>a repassar</div>
            </div>
          </div>
          <div style={{display:"flex",gap:6,marginBottom:10}}>
            {[{v:aps.length,l:"sessões",c:"#a78bfa"},{v:brl(tBruto),l:"bruto",c:"var(--mt)"},{v:pctMed+"%",l:"% med.",c:"#a78bfa"}].map(x=>(
              <div key={x.l} style={{flex:1,background:"var(--sx)",borderRadius:7,padding:"6px 0",textAlign:"center"}}>
                <div style={{fontWeight:900,color:x.c,fontSize:14}}>{x.v}</div>
                <div style={{fontSize:9,color:"var(--mt)"}}>{x.l}</div>
              </div>
            ))}
          </div>
          {lancProf.length>0&&<div style={{display:"flex",gap:5,flexWrap:"wrap",fontSize:10}}>
            {pago>0&&<span style={{padding:"2px 8px",borderRadius:20,background:"#34d39920",color:"#34d399",fontWeight:800}}>✅ Pago: {brl(pago)}</span>}
            {pendente>0&&<span style={{padding:"2px 8px",borderRadius:20,background:"#f59e0b20",color:"#f59e0b",fontWeight:800}}>⏳ Pend.: {brl(pendente)}</span>}
          </div>}
          <div style={{fontSize:10,color:"#a78bfa",marginTop:8,fontWeight:700}}>{recebProf?.id===prof.id?"▼ Ocultar sessões":"▶ Ver sessões"}</div>
        </div>);
      })}
      {profsVisiveis.length===0&&<div className="muted" style={{padding:20,gridColumn:"1/-1",textAlign:"center"}}>Nenhum profissional no filtro.</div>}
    </div>
    {/* Detalhe */}
    {recebProf&&(()=>{
      const aps=agReceb.filter(a=>Number(a.profissionalId)===Number(recebProf.id));
      const cor=espCor((recebProf.especialidades||[""])[0]);
      return(<div className="card" style={{overflow:"hidden",borderTop:"3px solid "+cor,marginBottom:8}}>
        <div style={{padding:"12px 16px",borderBottom:"1px solid var(--sc)",display:"flex",justifyContent:"space-between",alignItems:"center",background:"var(--sx)"}}>
          <b style={{fontSize:13,color:cor}}>{recebProf.nome} — Detalhe do período</b>
          <button className="btn secondary small" onClick={()=>setRecebProf(null)}>✕ Fechar</button>
        </div>
        <div className="grid-header" style={{gridTemplateColumns:"90px 58px 1fr 130px 78px 90px 65px 90px"}}>
          <div>Data</div><div>Hora</div><div>Paciente</div><div>Procedimento</div><div>Status</div>
          <div style={{textAlign:"right"}}>Bruto</div><div style={{textAlign:"center"}}>%</div><div style={{textAlign:"right"}}>Repasse</div>
        </div>
        {aps.length===0&&<div style={{padding:14}} className="muted">Nenhuma sessão com repasse no período.</div>}
        {aps.sort((a,b)=>a.data.localeCompare(b.data)).map((ag,i)=>{
          const pN=(pacientes||[]).find(p=>p.id===Number(ag.pacienteId));
          const pr=(procedimentos||[]).find(p=>p.id===Number(ag.procedimentoId));
          const pct=(pr?.pctRepasse||45)/100;const vb=pr?.valor||0;const vr=Math.round(vb*pct*100)/100;
          const st=STATUS_AG[ag.status]||{color:"#64748b",label:ag.status,icon:"📋"};
          return(<div key={ag.id} className="grid-row" style={{gridTemplateColumns:"90px 58px 1fr 130px 78px 90px 65px 90px",background:i%2?"var(--gr)":""}}>
            <div style={{fontSize:11,fontWeight:700}}>{brDate(ag.data)}</div>
            <div style={{fontWeight:800,color:"#a78bfa",fontSize:11}}>{ag.horarioSessao}</div>
            <div style={{fontWeight:700,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{pN?.nome||"—"}</div>
            <div style={{fontSize:11,color:"var(--mt)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{pr?.nome||"—"}</div>
            <div><span style={{padding:"1px 5px",borderRadius:20,fontSize:9,fontWeight:800,background:st.color+"25",color:st.color}}>{st.icon} {st.label}</span></div>
            <div style={{textAlign:"right",fontWeight:700,fontSize:12}}>{brl(vb)}</div>
            <div style={{textAlign:"center",fontWeight:800,color:"#a78bfa",fontSize:12}}>{Math.round(pct*100)}%</div>
            <div style={{textAlign:"right",fontWeight:900,color:"#34d399",fontSize:12}}>{brl(vr)}</div>
          </div>);
        })}
        {aps.length>0&&(()=>{
          const tot=aps.reduce((s,a)=>{const p=(procedimentos||[]).find(x=>x.id===Number(a.procedimentoId));return s+(p?.valor||0)*((p?.pctRepasse||45)/100);},0);
          const bruto=aps.reduce((s,a)=>{const p=(procedimentos||[]).find(x=>x.id===Number(a.procedimentoId));return s+(p?.valor||0);},0);
          return(<div className="grid-row" style={{gridTemplateColumns:"90px 58px 1fr 130px 78px 90px 65px 90px",background:"var(--sx)",fontWeight:900,fontSize:12}}>
            <div style={{gridColumn:"1/6"}}>TOTAL — {aps.length} sessões</div>
            <div style={{textAlign:"right",color:"var(--mt)"}}>{brl(bruto)}</div><div/>
            <div style={{textAlign:"right",color:"#34d399",fontSize:13}}>{brl(tot)}</div>
          </div>);
        })()}
      </div>);
    })()}
  </div>);
}

function ProfissionaisPage({profissionais,setProfissionais,setChamados,showToast,auth,filiais,agenda,procedimentos,lancamentos,setLancamentos,convenios,templateProfissional,pacientes}){
  const [modal,setModal]=useState(false);const[editing,setEditing]=useState(null);const[contratoProf,setContratoProf]=useState(null);const[viewProf,setViewProf]=useState("grid");const[tabProfP,setTabProfP]=useState("lista");  const podeEditar=PODE_PROFISSIONAL(auth.role);
  const salvar=p=>{if(editing){setProfissionais(a=>a.map(x=>x.id===editing.id?p:x));showToast("✅ Atualizado","ok");}else{setProfissionais(a=>[...a,p]);const num=String(Math.floor(100000+Math.random()*900000));setChamados(c=>[...c,{id:Date.now(),numero:num,setor:"Faturamento",tipo:"novo_profissional",nome:p.nome,descricao:"Novo profissional: "+p.nome+" — "+(p.especialidades||[]).join(", "),data:hoje_str,status:"aberto",resp:""}]);showToast("✅ Salvo · 📨 #"+num+" → Faturamento","ok");}setModal(false);};
  const del=id=>{if(!podeEditar)return showToast("❌ Sem permissão","err");if(!confirm("Excluir?"))return;setProfissionais(a=>a.filter(x=>x.id!==id));showToast("🗑️ Removido","err");};
  return(<div className="page-wrap">
    <div className="page-head"><h1>🩺 Profissionais</h1>
      <div style={{display:"flex",gap:6,alignItems:"center"}}>
        <button onClick={()=>setViewProf(v=>v==="grid"?"list":"grid")} style={{background:"transparent",border:"1.5px solid var(--cpb)",borderRadius:8,padding:"5px 10px",cursor:"pointer",fontSize:12,color:"var(--mt)",whiteSpace:"nowrap"}}>{viewProf==="grid"?"☰ Lista":"⊞ Grade"}</button>
        {podeEditar&&<button className="btn primary" onClick={()=>{setEditing(null);setModal(true);}}>+ Novo Profissional</button>}
      </div>
    </div>
    <div className="tab-bar" style={{marginBottom:14}}>
      {[["lista","🩺 Profissionais"],["recebimentos","💸 Repasse"],["tabRepasse","📋 Tabela de Repasse"],["tipos","👥 Tipos de Usuário"]].map(([k,lbl])=>(
        <button key={k} className="btn tab-btn" onClick={()=>setTabProfP(k)} style={{background:tabProfP===k?"var(--na)":"transparent",color:tabProfP===k?"#7c6af7":"var(--mt)",fontWeight:tabProfP===k?800:500}}>{lbl}</button>
      ))}
    </div>
    {tabProfP==="tipos"&&<TiposUsuarioInline auth={auth}/>}

    {tabProfP==="tabRepasse"&&<TabelaRepassePanel profissionais={profissionais} setProfissionais={setProfissionais} convenios={convenios} procedimentos={procedimentos} showToast={showToast}/>}
    {tabProfP==="recebimentos"&&<RecebimentosPanel profissionais={profissionais} procedimentos={procedimentos} convenios={convenios} lancamentos={lancamentos} agenda={agenda} pacientes={pacientes} showToast={showToast}/>}
    {tabProfP==="lista"&&<>
    {viewProf==="list"&&<div style={{borderRadius:10,overflow:"hidden",border:"1px solid var(--sc)",marginBottom:10}}>
      <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 80px",background:"var(--sx)",padding:"7px 12px",fontSize:10,fontWeight:900,color:"var(--mt)",letterSpacing:.5,gap:8}}><div>PROFISSIONAL</div><div>ESPECIALIDADE</div><div>DIAS</div><div></div></div>
      {profissionais.map(p=>{const esps=p.especialidades||[p.especialidade||""];const cor=espCor(esps[0]);
        return(<div key={p.id} style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 80px",padding:"9px 12px",borderTop:"1px solid var(--sc)",gap:8,alignItems:"center",cursor:podeEditar?"pointer":"default"}} onClick={()=>podeEditar&&(setEditing(p),setModal(true))}>
          <div><div style={{fontWeight:800,fontSize:12}}>{p.nome}</div><div style={{fontSize:10,color:"var(--mt)"}}>{PERFIL_LABEL[p.role]||p.role}</div></div>
          <div style={{display:"flex",flexWrap:"wrap",gap:3}}>{esps.slice(0,2).map(e=><span key={e} style={{fontSize:9,padding:"1px 5px",borderRadius:10,background:espCor(e)+"22",color:espCor(e),fontWeight:800}}>{e}</span>)}</div>
          <div style={{fontSize:11,color:"#10b981",fontWeight:700}}>{DIAS_SEMANA.filter(d=>p.escala?.[d]?.ativo).map(d=>d.slice(0,3)).join("·")||"—"}</div>
          <div style={{display:"flex",gap:4}} onClick={e=>e.stopPropagation()}>
            <button className="btn secondary small" onClick={()=>setContratoProf(p)}>📄</button>
            {podeEditar&&<button onClick={e=>{e.stopPropagation();if(confirm("Excluir?"))setProfissionais(a=>a.filter(x=>x.id!==p.id));}} style={{background:"none",border:"none",color:"var(--mt)",cursor:"pointer",fontSize:13}}>✕</button>}
          </div>
        </div>);})}
    </div>}
    <div style={{display:viewProf==="grid"?"grid":"none",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:10}}>
      {profissionais.map(p=>{const esps=p.especialidades||[p.especialidade||""];const cor=espCor(esps[0]);
        return(<div key={p.id} className="card rh" onClick={()=>{if(podeEditar){setEditing(p);setModal(true);}}} style={{padding:13,borderTop:"3px solid "+cor}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div><div style={{fontWeight:800,fontSize:14}}>{p.nome}</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:3,marginTop:3}}>{esps.map(e=><span key={e} style={{fontSize:9,padding:"1px 5px",borderRadius:10,background:espCor(e)+"22",color:espCor(e),fontWeight:800}}>{e}</span>)}</div>
            </div>
            <div style={{display:"flex",gap:4,alignItems:"flex-start"}}>
              <span style={{padding:"2px 7px",borderRadius:20,fontSize:9,fontWeight:800,background:"var(--sx)",color:"var(--mt)"}}>{PERFIL_LABEL[p.role]||p.role}</span>
              {podeEditar&&<button onClick={e=>{e.stopPropagation();del(p.id);}} style={{background:"none",border:"none",color:"var(--mt)",cursor:"pointer",fontSize:13}}>✕</button>}
            </div>
          </div>
          <div style={{marginTop:8,display:"grid",gridTemplateColumns:"1fr 1fr",gap:5}}>
            <div style={{background:"var(--sx)",borderRadius:6,padding:"5px 7px"}}><div style={{fontSize:9,color:"var(--lb)",fontWeight:800,letterSpacing:.5}}>DIAS</div><div style={{fontSize:11,fontWeight:700,color:"#10b981"}}>{DIAS_SEMANA.filter(d=>p.escala?.[d]?.ativo).map(d=>d.slice(0,3)).join("·")||"—"}</div></div>
            <div style={{background:"var(--sx)",borderRadius:6,padding:"5px 7px"}}><div style={{fontSize:9,color:"var(--lb)",fontWeight:800,letterSpacing:.5}}>FILIAIS</div><div style={{fontSize:10,fontWeight:700,color:"#a78bfa"}}>{(p.filiaisAtendimento||[]).length>0?filiais.filter(f=>(p.filiaisAtendimento||[]).includes(f.id)).map(f=>f.codigo).join("·"):"—"}</div></div>
          </div>
          <div style={{marginTop:8,paddingTop:7,borderTop:"1px solid var(--db)"}} onClick={e=>e.stopPropagation()}>
            <button className="btn secondary small" style={{width:"100%",fontSize:11}} onClick={()=>setContratoProf(p)}>📄 Gerar Contrato PDF</button>
          </div>
        </div>);
      })}
    </div>
    {modal&&podeEditar&&<CadastroProfissionalModal editing={editing} profissionais={profissionais} filiais={filiais} onClose={()=>setModal(false)} onSave={salvar}/>}
    {contratoProf&&<ContratoPdfModal pessoa={contratoProf} tipo="profissional" templateOverride={templateProfissional} onClose={()=>setContratoProf(null)}/>}
    </>}
  </div>);
}

// ═══════════════════════════════════════════════════════════════════════════════
// SALAS & FILIAIS
// ═══════════════════════════════════════════════════════════════════════════════
function SalasPage({salas,setSalas,filiais,setFiliais,showToast}){
  const [tabSF,setTabSF]=useState("salas");
  const [modoLista,setModoLista]=useState(false);const[modalS,setModalS]=useState(false);const[editingS,setEditingS]=useState(null);const[modalF,setModalF]=useState(false);const[editingF,setEditingF]=useState(null);
  const CORES=["#38bdf8","#4ade80","#f472b6","#fb923c","#c084fc","#fbbf24","#f87171","#34d399","#a78bfa","#f59e0b","#6A0572","#008080","#DAA520","#9370DB"];
  const blankS={filialId:filiais[0]?.id||1,nome:"",especialidade:"",cor:"#38bdf8",ativa:true,custoMensal:0};
  const blankF={nome:"",codigo:"",logradouro:"",numero:"",bairro:"",cidade:"",estado:"SP",cep:"",telefone:"",ativa:true};
  const [fs,setFs]=useState(blankS);const[ff,setFf]=useState(blankF);
  const saveS=()=>{if(!fs.nome.trim())return alert("Nome obrigatório");if(editingS)setSalas(a=>a.map(x=>x.id===editingS.id?{...fs,id:editingS.id}:x));else setSalas(a=>[...a,{...fs,id:Date.now()}]);showToast("✅ Sala salva","ok");setModalS(false);};
  const saveF=()=>{if(!ff.nome.trim())return alert("Nome obrigatório");if(editingF)setFiliais(a=>a.map(x=>x.id===editingF.id?{...ff,id:editingF.id}:x));else setFiliais(a=>[...a,{...ff,id:Date.now()}]);showToast("✅ Filial salva","ok");setModalF(false);};
  const salasPorFilial=fid=>salas.filter(s=>s.filialId===fid);
  return(<div className="page-wrap">
    <div className="page-head"><h1>🏢 Salas & Filiais</h1>
      <div style={{display:"flex",gap:7,alignItems:"center"}}>
        <button className="btn secondary small" onClick={()=>setModoLista(v=>!v)} title={modoLista?"Modo cards":"Modo lista"} style={{fontSize:13}}>{modoLista?"▦":"☰"}</button>
        {tabSF==="salas"?<button className="btn primary" onClick={()=>{setEditingS(null);setFs(blankS);setModalS(true);}}>+ Nova Sala</button>:<button className="btn primary" onClick={()=>{setEditingF(null);setFf(blankF);setModalF(true);}}>+ Nova Filial</button>}
      </div>
    </div>
    <div className="tab-bar"><button className="btn tab-btn" onClick={()=>setTabSF("salas")} style={{background:tabSF==="salas"?"var(--na)":"transparent",color:tabSF==="salas"?"#7c6af7":"var(--mt)",fontWeight:tabSF==="salas"?800:500}}>🚪 Salas ({salas.length})</button><button className="btn tab-btn" onClick={()=>setTabSF("filiais")} style={{background:tabSF==="filiais"?"var(--na)":"transparent",color:tabSF==="filiais"?"#7c6af7":"var(--mt)",fontWeight:tabSF==="filiais"?800:500}}>🏢 Filiais ({filiais.length})</button></div>
    {tabSF==="salas"&&<div>{modoLista
      ?<div className="card" style={{overflow:"hidden"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 120px 80px 60px",background:"var(--sx)",padding:"6px 12px",fontSize:10,fontWeight:900,color:"var(--mt)",borderBottom:"1px solid var(--cb)"}}>
            <div>SALA</div><div>FILIAL</div><div>ESPECIALIDADE</div><div>CUSTO</div><div></div>
          </div>
          {salas.map(s=>{const fil=filiais.find(f=>f.id===s.filialId);return(<div key={s.id} style={{display:"grid",gridTemplateColumns:"1fr 1fr 120px 80px 60px",padding:"8px 12px",borderBottom:"1px solid var(--db)",alignItems:"center",cursor:"pointer"}} onClick={()=>{setEditingS(s);setFs({...s});setModalS(true);}}>
            <div style={{display:"flex",alignItems:"center",gap:7}}><div style={{width:10,height:10,borderRadius:"50%",background:s.cor,flexShrink:0}}/><span style={{fontWeight:700,fontSize:12}}>{s.nome}</span>{!s.ativa&&<span style={{fontSize:9,color:"#f87171",fontWeight:800}}>INATIVA</span>}</div>
            <div style={{fontSize:11,color:"var(--mt)"}}>{fil?.nome||"—"}</div>
            <div style={{fontSize:11,color:s.especialidade?espCor(s.especialidade):"var(--mt)"}}>{s.especialidade||"—"}</div>
            <div style={{fontSize:11,color:"#34d399"}}>{s.custoMensal?brl(s.custoMensal):"—"}</div>
            <div><button className="btn secondary small" style={{fontSize:10}} onClick={e=>{e.stopPropagation();setEditingS(s);setFs({...s});setModalS(true);}}>✏️</button></div>
          </div>);})}
        </div>
      :<div>{filiais.filter(f=>f.ativa).map(fil=>{const sl=salasPorFilial(fil.id);return(<div key={fil.id} style={{marginBottom:18}}><div style={{fontWeight:900,color:"#a78bfa",fontSize:12,marginBottom:7,letterSpacing:.5}}>🏢 {fil.nome} <span style={{color:"var(--mt)",fontWeight:400,fontSize:11}}>({fil.codigo})</span></div><div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(210px,1fr))",gap:7}}>{sl.map(s=><div key={s.id} className="card rh" onClick={()=>{setEditingS(s);setFs({...s});setModalS(true);}} style={{padding:11,borderLeft:"4px solid "+s.cor}}><div style={{fontWeight:800,color:s.cor,fontSize:12}}>{s.nome}</div>{s.especialidade&&<div style={{fontSize:11,color:espCor(s.especialidade),marginTop:1,fontWeight:700}}>{s.especialidade}</div>}<div className="muted" style={{fontSize:10,marginTop:2}}>{s.ativa?"✅ Ativa":"❌ Inativa"}</div></div>)}{sl.length===0&&<div className="muted" style={{fontSize:11}}>Nenhuma sala.</div>}</div></div>);})}
      </div>
    }</div>}
    {tabSF==="filiais"&&<div style={{display:modoLista?"block":"grid",gridTemplateColumns:"repeat(auto-fill,minmax(260px,1fr))",gap:10}}>{filiais.map(f=><div key={f.id} className="card rh" style={{padding:13,borderLeft:"3px solid #7c6af7"}} onClick={()=>{setEditingF(f);setFf({...blankF,...f});setModalF(true);}}><div style={{display:"flex",justifyContent:"space-between"}}><div><div style={{fontWeight:800,fontSize:13}}>{f.nome}</div><div className="muted" style={{fontSize:11}}>{f.codigo} · {salasPorFilial(f.id).length} sala(s)</div>{f.logradouro&&<div style={{fontSize:10,color:"var(--mt)",marginTop:2}}>📍 {f.logradouro}{f.numero?", "+f.numero:""}{f.bairro?" — "+f.bairro:""}</div>}{f.cidade&&<div style={{fontSize:10,color:"var(--mt)"}}>{f.cidade}{f.estado?" / "+f.estado:""}</div>}{f.telefone&&<div style={{fontSize:10,color:"var(--mt)"}}>📞 {maskPhone(f.telefone)}</div>}</div><span style={{fontSize:11,fontWeight:800,color:f.ativa?"#34d399":"#f87171"}}>{f.ativa?"Ativa":"Inativa"}</span></div></div>)}</div>}
    {modalS&&<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&setModalS(false)}><div className="modal" style={{width:520}}><div className="modal-head"><h2>{editingS?"✏️ Editar":"+ Nova"} Sala</h2><button className="icon-btn" onClick={()=>setModalS(false)}>×</button></div><div className="stack"><div className="g2"><div><label>Filial</label><select value={fs.filialId} onChange={e=>setFs(f=>({...f,filialId:Number(e.target.value)}))}>{filiais.map(f=><option key={f.id} value={f.id}>{f.nome}</option>)}</select></div><div><label>Nome *</label><input value={fs.nome} onChange={e=>setFs(f=>({...f,nome:e.target.value}))} autoFocus/></div></div><div><label>Especialidade</label><select value={fs.especialidade||""} onChange={e=>setFs(f=>({...f,especialidade:e.target.value}))}><option value="">Sem esp.</option>{ESPECIALIDADES_LIST.map(e=><option key={e}>{e}</option>)}</select></div><div><label>💵 Custo Mensal (R$)</label><input type="number" min="0" step="50" value={fs.custoMensal||0} onChange={e=>setFs(f=>({...f,custoMensal:Number(e.target.value)}))} placeholder="0,00"/></div><div><label>Cor</label><div style={{display:"flex",gap:6,marginTop:4,flexWrap:"wrap"}}>{CORES.map(c=><div key={c} onClick={()=>setFs(f=>({...f,cor:c}))} style={{width:24,height:24,borderRadius:"50%",background:c,cursor:"pointer",border:"3px solid "+(fs.cor===c?"#fff":"transparent"),transform:fs.cor===c?"scale(1.2)":"scale(1)",transition:".1s"}}/>)}</div></div><label style={{display:"flex",alignItems:"center",gap:7,cursor:"pointer",textTransform:"none",fontSize:12,fontWeight:700}}><input type="checkbox" checked={fs.ativa} onChange={e=>setFs(f=>({...f,ativa:e.target.checked}))}/>Sala ativa</label><div className="actions"><button className="btn secondary" onClick={()=>setModalS(false)}>Cancelar</button><button className="btn primary" onClick={saveS}>Salvar</button></div></div></div></div>}
    {modalF&&<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&setModalF(false)}><div className="modal" style={{width:520}}><div className="modal-head"><h2>{editingF?"✏️ Editar":"+ Nova"} Filial</h2><button className="icon-btn" onClick={()=>setModalF(false)}>×</button></div><div className="stack">
      <div className="g2"><div><label>Nome *</label><input value={ff.nome} onChange={e=>setFf(f=>({...f,nome:e.target.value}))} autoFocus/></div><div><label>Código</label><input value={ff.codigo} onChange={e=>setFf(f=>({...f,codigo:e.target.value}))}/></div></div>
      <div className="section-box"><div className="section-title">📍 Endereço</div>
        <div className="g2"><div><label>CEP</label><input value={ff.cep||""} onChange={e=>setFf(f=>({...f,cep:e.target.value}))} placeholder="00000-000" maxLength={9}/></div><div><label>Telefone</label><input value={ff.telefone||""} onChange={e=>setFf(f=>({...f,telefone:rawD(e.target.value).slice(0,11)}))}/></div></div>
        <div><label>Logradouro</label><input value={ff.logradouro||""} onChange={e=>setFf(f=>({...f,logradouro:e.target.value}))}/></div>
        <div className="g2"><div><label>Número</label><input value={ff.numero||""} onChange={e=>setFf(f=>({...f,numero:e.target.value}))}/></div><div><label>Bairro</label><input value={ff.bairro||""} onChange={e=>setFf(f=>({...f,bairro:e.target.value}))}/></div></div>
        <div className="g2"><div><label>Cidade</label><input value={ff.cidade||""} onChange={e=>setFf(f=>({...f,cidade:e.target.value}))}/></div><div><label>Estado</label><select value={ff.estado||"SP"} onChange={e=>setFf(f=>({...f,estado:e.target.value}))}>{UFS.map(u=><option key={u}>{u}</option>)}</select></div></div>
      </div>
      <label style={{display:"flex",alignItems:"center",gap:7,cursor:"pointer",textTransform:"none",fontSize:12,fontWeight:700}}><input type="checkbox" checked={ff.ativa} onChange={e=>setFf(f=>({...f,ativa:e.target.checked}))}/>Filial ativa</label>
      <div className="actions"><button className="btn secondary" onClick={()=>setModalF(false)}>Cancelar</button><button className="btn primary" onClick={saveF}>Salvar</button></div>
    </div></div></div>}
  </div>);
}

// ═══════════════════════════════════════════════════════════════════════════════
// ESPELHO DE AGENDA
// ═══════════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════════
// ESPELHO DE AGENDA — grade com divs explícitos por linha (sem fragments)
// ═══════════════════════════════════════════════════════════════════════════════
function EspelhoAgendaPage({data,agenda,pacientes,profissionais,salas,filiais,convenios,procedimentos,onAbrirAgendamento,onFecharHorario,horariosF,onAbrirSessao,onBuscaInteligente}){
  const [filtroEsp,setFiltroEsp]=useState("");
  const [filtroProf,setFiltroProf]=useState("");
  const [profSel,setProfSel]=useState("");
  const filtroDia=data;
  const dias=useMemo(()=>{
    const d=new Date(filtroDia+"T12:00:00");
    d.setDate(d.getDate()-d.getDay());
    return Array.from({length:7},(_,i)=>{const x=new Date(d);x.setDate(d.getDate()+i);return ymd(x);});
  },[filtroDia]);
  const NOMES=["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"];
  const profs=profissionais.filter(p=>p.role==="profissional"
    &&(!filtroEsp||(p.especialidades||[p.especialidade]).includes(filtroEsp))
    &&(!filtroProf||p.nome.toLowerCase().includes(filtroProf.toLowerCase()))
    &&(!profSel||String(p.id)===String(profSel)));
  const COL="75px repeat(7,1fr)";

  // ── Indicadores contextuais ──────────────────────────────────
  const agBase=agenda.filter(a=>dias.includes(a.data)&&a.tipo!=="reuniao");
  // Filtro por profissional ou especialidade selecionados
  const agFiltrada=agBase.filter(a=>{
    if(profSel&&String(a.profissionalId)!==String(profSel))return false;
    if(filtroEsp){const p=profissionais.find(x=>x.id===Number(a.profissionalId));if(!(p?.especialidades||[p?.especialidade]).includes(filtroEsp))return false;}
    return true;
  });
  const totalSem=agFiltrada.length;
  const atendSem=agFiltrada.filter(a=>["atendido","faturado"].includes(a.status)).length;
  const faltaSem=agFiltrada.filter(a=>["faltou","faltou_pacote"].includes(a.status)).length;
  const agendSem=agFiltrada.filter(a=>["agendado","confirmado"].includes(a.status)).length;
  const txPres=totalSem>0?Math.round(atendSem/totalSem*100):0;

  // Taxa de ocupação: slots disponíveis vs agendados
  const calcOcupacao=()=>{
    let slotsDisp=0,slotsOcup=0;
    profs.forEach(prof=>{
      const diasAtivos=DIAS_SEMANA.filter(d=>prof.escala?.[d]?.ativo);
      dias.forEach(dia=>{
        const diaNome=DIAS_SEMANA[new Date(dia+"T12:00:00").getDay()];
        if(!diasAtivos.includes(diaNome))return;
        const esc=prof.escala?.[diaNome];
        if(!esc)return;
        const ini=toMin(esc.inicio||"08:00");
        const fim=toMin(esc.fim||"18:00");
        const tempos=prof.temposAtendimento||[50];
        const slot=tempos[0]||50;
        // slots teóricos no dia
        const nSlots=Math.floor((fim-ini)/slot);
        slotsDisp+=nSlots;
        // slots ocupados
        const agProf=agFiltrada.filter(a=>a.data===dia&&String(a.profissionalId)===String(prof.id));
        slotsOcup+=agProf.length;
      });
    });
    return slotsDisp>0?Math.round(slotsOcup/slotsDisp*100):0;
  };
  const txOcup=calcOcupacao();
  const profSelObj=profSel?profissionais.find(p=>String(p.id)===String(profSel)):null;
  return(<div className="page-wrap" style={{paddingBottom:30}}>
    <div className="page-head"><h1>🪞 Espelho de Agenda</h1>
      <div style={{display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>
        <select value={filtroEsp} onChange={e=>{setFiltroEsp(e.target.value);setProfSel("");setFiltroProf("");}} style={{width:"auto",minWidth:150,fontSize:12}}>
          <option value="">Todas Especialidades</option>
          {[...new Set(profissionais.filter(p=>p.role==="profissional").flatMap(p=>p.especialidades||[p.especialidade]))].map(e=><option key={e} value={e}>{e}</option>)}
        </select>
        <select value={profSel} onChange={e=>{setProfSel(e.target.value);setFiltroProf("");if(e.target.value){const p=profissionais.find(x=>String(x.id)===e.target.value);if(p)setFiltroEsp("");}}} style={{width:"auto",minWidth:170,fontSize:12}}>
          <option value="">Todos os Profissionais</option>
          {profissionais.filter(p=>p.role==="profissional"&&(!filtroEsp||(p.especialidades||[p.especialidade]).includes(filtroEsp))).map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}
        </select>
        <div style={{position:"relative",display:"flex",alignItems:"center"}}>
          <span style={{position:"absolute",left:9,fontSize:12,color:"var(--mt)",pointerEvents:"none"}}>🔍</span>
          <input value={filtroProf} onChange={e=>{setFiltroProf(e.target.value);setProfSel("");}} placeholder="Buscar..." style={{paddingLeft:28,width:140,fontSize:12}}/>
          {filtroProf&&<button onClick={()=>setFiltroProf("")} style={{position:"absolute",right:7,background:"none",border:"none",cursor:"pointer",color:"var(--mt)",fontSize:13,lineHeight:1}}>×</button>}
        </div>
        {onBuscaInteligente&&<button className="btn primary" style={{fontSize:11,whiteSpace:"nowrap"}} onClick={onBuscaInteligente}>🤖 Encaixe Inteligente</button>}
      </div>
    </div>
    {/* ── Contexto do filtro ── */}
    {(profSelObj||filtroEsp)&&<div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,padding:"7px 12px",background:"var(--sx)",borderRadius:9,border:"1px solid var(--cb)"}}>
      {profSelObj&&<div style={{display:"flex",alignItems:"center",gap:7}}>
        <div style={{width:8,height:8,borderRadius:"50%",background:espCor((profSelObj.especialidades||[])[0]||"")}}/>
        <span style={{fontWeight:800,fontSize:12}}>{profSelObj.nome}</span>
        <span style={{fontSize:11,color:"var(--mt)"}}>{(profSelObj.especialidades||[]).join(", ")}</span>
      </div>}
      {filtroEsp&&!profSelObj&&<span style={{fontWeight:800,fontSize:12,color:espCor(filtroEsp)}}>{filtroEsp}</span>}
      <span style={{fontSize:10,color:"var(--mt)",marginLeft:"auto"}}>{profs.length} profissional(is) · semana {brDate(dias[0]).slice(0,5)}–{brDate(dias[6]).slice(0,5)}</span>
      <button onClick={()=>{setProfSel("");setFiltroEsp("");setFiltroProf("");}} style={{background:"none",border:"none",cursor:"pointer",color:"#f87171",fontSize:11,fontWeight:700}}>✕ Limpar</button>
    </div>}
    <div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:6,marginBottom:12}}>
      {[
        ["SESSÕES","📅",totalSem,"#a78bfa"],
        ["ATENDIDOS","✅",atendSem,"#34d399"],
        ["FALTAS","🚫",faltaSem,"#f87171"],
        ["AGUARDANDO","📋",agendSem,"#f59e0b"],
        ["PRESENÇA","📊",txPres+"%",txPres>=80?"#34d399":txPres>=60?"#f59e0b":"#f87171"],
        ["OCUPAÇÃO","🏥",txOcup+"%",txOcup>=80?"#34d399":txOcup>=60?"#f59e0b":"#f87171"],
      ].map(([t,ic,v,cor])=>(
        <div key={t} className="card" style={{padding:"7px 8px",textAlign:"center",borderTop:"3px solid "+cor}}>
          <div style={{fontSize:8,fontWeight:900,color:"var(--mt)",letterSpacing:.4}}>{ic} {t}</div>
          <div style={{fontSize:18,fontWeight:900,color:cor,marginTop:2}}>{v}</div>
        </div>
      ))}
    </div>
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      {profs.map(prof=>{
        const esps=prof.especialidades||[prof.especialidade||""];
        const cor=espCor(esps[0]);
        return(<div key={prof.id} className="card" style={{overflow:"hidden",borderTop:"4px solid "+cor}}>
          <div style={{background:"var(--sx)",padding:"9px 13px",borderBottom:"1px solid var(--fc)",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:5}}>
            <b style={{color:cor}}>{prof.nome.toUpperCase()}</b>
            <span className="muted" style={{fontSize:11}}>({esps.join(", ")}) · {(prof.temposAtendimento||[prof.tempoAtendimento]).join("/")}min</span>
          </div>
          <div style={{overflowX:"auto"}}>
            {/* HEADER ROW */}
            <div style={{display:"grid",gridTemplateColumns:COL,minWidth:620}}>
              <div style={{padding:"4px",textAlign:"center",fontSize:9,fontWeight:900,color:"var(--mt)",borderBottom:"1px solid var(--fc)",borderRight:"1px solid var(--fc)"}}>HORA</div>
              {NOMES.map((n,i)=>{
                const isH=dias[i]===hoje_str;
                return(<div key={i} style={{padding:"3px 2px",textAlign:"center",fontSize:9,fontWeight:900,color:isH?"#a78bfa":"var(--mt)",borderBottom:"1px solid var(--fc)",borderRight:"1px solid var(--fc)",background:isH?"var(--na)":""}}>{n}<br/><small style={{fontWeight:400,fontSize:8}}>{brDate(dias[i]).slice(0,5)}</small></div>);
              })}
            </div>
            {/* DATA ROWS — one div per turno */}
            {TURNOS_H.map((hora,idx)=>(
              <div key={hora} style={{display:"grid",gridTemplateColumns:COL,minWidth:620}}>
                {/* Hora label */}
                <div style={{display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,color:"#7c6af7",borderBottom:"1px solid var(--fc)",borderRight:"1px solid var(--fc)",background:idx%2?"var(--gr)":"",minHeight:40}}>{hora}</div>
                {/* 7 day cells */}
                {dias.map((dia,di)=>{
                  const diaNome=DIAS_SEMANA[di];
                  const esc=prof.escala?.[diaNome];
                  const emPausa=esc?.usarPausa&&toMin(hora)>=toMin(esc.pausaInicio||"12:00")&&toMin(hora)<toMin(esc.pausaFim||"13:00");
                  const noExp=esc?.ativo&&toMin(hora)>=toMin(esc.inicio||"08:00")&&toMin(hora)<toMin(esc.fim||"18:00")&&!emPausa;
                  const chave=prof.id+"_"+dia+"_"+hora;
                  const fechado=horariosF?.includes(chave);
                  const ag=agenda.find(a=>a.profissionalId===prof.id&&a.data===dia&&toMin(hora)>=toMin(a.horarioSessao)&&toMin(hora)<toMin(a.horarioFimSessao));
                  const isS=ag?.horarioSessao===hora;
                  const base={borderBottom:"1px solid var(--fc)",borderRight:"1px solid var(--fc)",background:idx%2?"var(--gr)":"",minHeight:40,position:"relative"};
                  if(!noExp)return(<div key={dia} style={{...base,display:"flex",alignItems:"center",justifyContent:"center",fontSize:8,fontWeight:900,color:"var(--lb)",background:"var(--sx)"}}>{emPausa?"☕":"—"}</div>);
                  if(fechado)return(<div key={dia} style={{...base,display:"flex",alignItems:"center",justifyContent:"center",fontSize:8,fontWeight:900,color:"var(--lb)",cursor:"pointer"}} onClick={()=>onFecharHorario&&onFecharHorario(chave,false)}>🔒 FECHADO<br/><span style={{fontSize:7,fontWeight:400}}>clique p/ liberar</span></div>);
                  const reservado=horariosF?.includes(chave+"_reservado");
                  if(reservado)return(<div key={dia} style={{...base,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontSize:8,fontWeight:900,color:"#f59e0b",cursor:"pointer",background:"#f59e0b10"}} onClick={()=>onFecharHorario&&onFecharHorario(chave+"_reservado",false)}>🟡 RESERVADO<br/><span style={{fontSize:7,fontWeight:400}}>clique p/ liberar</span></div>);
                  if(ag){
                    const pac=pacientes.find(p=>p.id===Number(ag.pacienteId));
                    const stInfo=STATUS_AG[ag.status]||{color:"#64748b",icon:"📋"};
                    return(<div key={dia} style={{...base,background:cor+"20",borderLeft:isS?"2px solid "+cor:"none",padding:"2px 3px",overflow:"hidden",cursor:"pointer"}} onClick={()=>onAbrirSessao&&onAbrirSessao(ag)}>
                      {isS&&<><div style={{fontWeight:800,fontSize:9,color:"var(--tx)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{pac?.nome?.split(" ")[0]}</div><div style={{fontSize:8,color:stInfo.color}}>{stInfo.icon} {ag.horarioSessao}</div></>}
                    </div>);
                  }
                  return(<div key={dia} style={{...base,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",cursor:"pointer",gap:2,padding:2}}>
                    <span onClick={()=>onAbrirAgendamento({data:dia,profId:prof.id,hora})} style={{fontSize:9,fontWeight:800,color:"#10b981"}}>✅ LIVRE</span>
                    <div style={{display:"flex",gap:2}}>
                      <span onClick={e=>{e.stopPropagation();onFecharHorario&&onFecharHorario(chave,true);}} style={{fontSize:7,color:"var(--mt)",cursor:"pointer",padding:"1px 3px",borderRadius:3,background:"var(--sx)"}} title="Fechar">🔒</span>
                      <span onClick={e=>{e.stopPropagation();onFecharHorario&&onFecharHorario(chave+"_reservado",true);}} style={{fontSize:7,color:"#f59e0b",cursor:"pointer",padding:"1px 3px",borderRadius:3,background:"var(--sx)"}} title="Reservar">🟡</span>
                    </div>
                  </div>);
                })}
              </div>
            ))}
          </div>
        </div>);
      })}
      {profs.length===0&&<div style={{textAlign:"center",color:"var(--mt)",padding:"40px"}}>Nenhum profissional.</div>}
    </div>
  </div>);
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAPEAMENTO DE SALAS — grade com divs explícitos por linha
// ═══════════════════════════════════════════════════════════════════════════════
function MapeamentoSalasPage({data,agenda,salas,filiais,pacientes,profissionais,horariosF,onAbrirAgendamento,onFecharHorario,onAbrirSessao,onBuscaInteligente}){
  const filtroDia=data;
  const [filtroFilial,setFiltroFilial]=useState("");
  const [filtroSala,setFiltroSala]=useState("");
  const eventos=agenda.filter(a=>a.data===filtroDia);
  const getAg=(salaId,hora)=>eventos.find(ag=>ag.salaId===salaId&&toMin(hora)>=toMin(ag.horarioSessao)&&toMin(hora)<toMin(ag.horarioFimSessao));
  const agora=toTime(new Date().getHours()*60+new Date().getMinutes());
  const salasEx=salas.filter(s=>s.ativa&&(!filtroSala||s.id===Number(filtroSala))&&(!filtroFilial||s.filialId===Number(filtroFilial)));
  const emUso=salas.filter(s=>s.ativa&&getAg(s.id,agora)).length;
  const totalAtivas=salas.filter(s=>s.ativa).length;
  const indFiliais=filiais.map(f=>{const sl=salas.filter(s=>s.ativa&&s.filialId===f.id);const occ=sl.filter(s=>getAg(s.id,agora)).length;return{...f,total:sl.length,occ};});
  // Dynamic grid template
  const COLS="80px "+salasEx.map(()=>"1fr").join(" ");
  return(<div className="page-wrap" style={{paddingBottom:30}}>
    <div className="page-head"><h1>🗺️ Mapeamento de Salas</h1>
      <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
        <select value={filtroFilial} onChange={e=>{setFiltroFilial(e.target.value);setFiltroSala("");}} style={{width:"auto",minWidth:120}}><option value="">Todas Filiais</option>{filiais.map(f=><option key={f.id} value={f.id}>{f.nome}</option>)}</select>
        <select value={filtroSala} onChange={e=>setFiltroSala(e.target.value)} style={{width:"auto",minWidth:130}}><option value="">Todas Salas</option>{salas.filter(s=>!filtroFilial||s.filialId===Number(filtroFilial)).map(s=><option key={s.id} value={s.id}>{s.nome}</option>)}</select>
        {onBuscaInteligente&&<button className="btn primary" style={{fontSize:11,whiteSpace:"nowrap"}} onClick={onBuscaInteligente}>🤖 Encaixe Inteligente</button>}
      </div>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:10}}>
      {[["OCUPAÇÃO",totalAtivas?Math.round(emUso/totalAtivas*100)+"%":"—","#7c6af7"],["EM USO",emUso+"/"+totalAtivas,"#10b981"],["HOJE",eventos.length,"#f59e0b"]].map(([t,v,c])=>(
        <div key={t} className="card" style={{padding:9,textAlign:"center",borderTop:"3px solid "+c}}><div className="muted" style={{fontSize:9,fontWeight:900,letterSpacing:.5}}>{t}</div><div style={{fontSize:18,fontWeight:900,color:c,marginTop:1}}>{v}</div></div>
      ))}
    </div>
    <div style={{display:"flex",gap:7,flexWrap:"wrap",marginBottom:12}}>
      {indFiliais.map(f=>(
        <div key={f.id} className="card" style={{padding:"6px 10px",borderLeft:"3px solid #7c6af7",minWidth:140}}>
          <div style={{fontWeight:800,fontSize:11,color:"#a78bfa"}}>{f.nome}</div>
          <div style={{fontSize:11,marginTop:2,display:"flex",gap:7}}><span style={{color:"#10b981"}}>🟢 {f.total-f.occ} livres</span><span style={{color:"#f59e0b"}}>🟡 {f.occ} em uso</span></div>
        </div>
      ))}
    </div>
    {salasEx.length===0?<div style={{textAlign:"center",color:"var(--mt)",padding:"40px"}}>Nenhuma sala para exibir.</div>:<div className="card" style={{overflowX:"auto",borderRadius:10}}>
      {/* HEADER */}
      <div style={{display:"grid",gridTemplateColumns:COLS,minWidth:400,background:"var(--sx)"}}>
        <div style={{padding:"6px",textAlign:"center",fontSize:9,fontWeight:900,color:"var(--mt)",border:"1px solid var(--fc)"}}>HORA</div>
        {salasEx.map(s=>{
          const fil=filiais.find(f=>f.id===s.filialId);
          return(<div key={s.id} style={{padding:"4px",textAlign:"center",color:s.cor,fontWeight:900,fontSize:10,border:"1px solid var(--fc)"}}>
            {s.nome.toUpperCase()}
            {s.especialidade&&<><br/><span style={{color:espCor(s.especialidade),fontSize:8,fontWeight:700}}>{s.especialidade}</span></>}
            <br/><small style={{fontWeight:400,fontSize:8,color:"var(--mt)"}}>{fil?.codigo||""}</small>
          </div>);
        })}
      </div>
      {/* DATA ROWS */}
      {TURNOS_H.map((hora,idx)=>{
        const fora=toMin(hora)<toMin("07:00")||toMin(hora)>=toMin("20:00");
        return(<div key={hora} style={{display:"grid",gridTemplateColumns:COLS,minWidth:400,background:fora?"var(--sx)":idx%2?"var(--gr)":"",borderBottom:"1px solid var(--db)"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,color:fora?"var(--mt)":"#93c5fd",borderRight:"1px solid var(--fc)",fontSize:11,minHeight:42}}>{hora}</div>
          {salasEx.map(sala=>{
            const ag=getAg(sala.id,hora);
            if(fora)return(<div key={sala.id} style={{display:"flex",alignItems:"center",justifyContent:"center",fontSize:8,color:"var(--mt)",borderRight:"1px solid var(--fc)"}}>—</div>);
            const chaveSala="sala_"+sala.id+"_"+filtroDia+"_"+hora;
            const fechadoSala=horariosF?.includes(chaveSala);
            const reservadoSala=horariosF?.includes(chaveSala+"_reservado");
            if(fechadoSala)return(<div key={sala.id} style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",borderRight:"1px solid var(--fc)",minHeight:42,fontSize:7,fontWeight:900,color:"var(--lb)",cursor:"pointer",background:"var(--sx)"}} onClick={()=>onFecharHorario&&onFecharHorario(chaveSala,false)}>🔒<br/>FECHADO</div>);
            if(reservadoSala)return(<div key={sala.id} style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",borderRight:"1px solid var(--fc)",minHeight:42,fontSize:7,fontWeight:900,color:"#f59e0b",cursor:"pointer",background:"#f59e0b10"}} onClick={()=>onFecharHorario&&onFecharHorario(chaveSala+"_reservado",false)}>🟡<br/>RESERVADO</div>);
            if(!ag)return(<div key={sala.id} style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",borderRight:"1px solid var(--fc)",cursor:"pointer",minHeight:42,gap:2}}>
              <span onClick={()=>onAbrirAgendamento({data:filtroDia,salaId:sala.id,hora})} style={{fontSize:9,fontWeight:800,color:"#10b981"}}>✅ LIVRE</span>
              <div style={{display:"flex",gap:2}}>
                <span onClick={e=>{e.stopPropagation();onFecharHorario&&onFecharHorario(chaveSala,true);}} style={{fontSize:7,color:"var(--mt)",cursor:"pointer",padding:"1px 3px",borderRadius:3,background:"var(--sx)"}} title="Fechar">🔒</span>
                <span onClick={e=>{e.stopPropagation();onFecharHorario&&onFecharHorario(chaveSala+"_reservado",true);}} style={{fontSize:7,color:"#f59e0b",cursor:"pointer",padding:"1px 3px",borderRadius:3,background:"var(--sx)"}} title="Reservar">🟡</span>
              </div>
            </div>);
            const pac=pacientes.find(p=>p.id===Number(ag.pacienteId));
            const prof=profissionais.find(p=>p.id===Number(ag.profissionalId));
            const stInfo=STATUS_AG[ag.status]||{color:"#64748b",icon:"📋"};
            const isS=ag.horarioSessao===hora;
            return(<div key={sala.id} style={{padding:"2px 4px",borderRight:"1px solid var(--fc)",background:sala.cor+"20",borderLeft:isS?"2px solid "+sala.cor:"none",cursor:"pointer",minHeight:42}} onClick={()=>onAbrirSessao&&onAbrirSessao(ag)}>
              {isS&&<><div style={{fontWeight:800,fontSize:9,color:"var(--tx)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{pac?.nome?.split(" ")[0]}</div><div style={{fontSize:8,color:stInfo.color}}>{stInfo.icon} {prof?.nome?.split(" ").pop()}</div></>}
            </div>);
          })}
        </div>);
      })}
    </div>}
  </div>);
}

// ═══════════════════════════════════════════════════════════════════════════════
function ConveniosProcedimentosPage({convenios,setConvenios,procedimentos,setProcedimentos,showToast}){
  const [tab,setTab]=useState("convenios");
  const [busca,setBusca]=useState("");
  // Convênio modal
  const [mC,setMC]=useState(false);
  const [eC,setEC]=useState(null);
  const blankC={nome:"",codigo:"",registro:"",cnpj:"",tipo:"Plano de Saúde",
    contato:"",email:"",telefone:"",site:"",manual:"",
    vigenciaIni:"",vigenciaFim:"",observacoes:"",ativo:true};
  const [fc,setFc]=useState(blankC);
  const sC=p=>setFc(f=>({...f,...p}));

  // Procedimento modal
  const [mP,setMP]=useState(false);
  const [eP,setEP]=useState(null);
  const blankP={nome:"",codigoTUSS:"",categoria:"Sessão Terapêutica",
    especialidade:"",duracao:50,valor:0,
    niveisRepasse:{Trainee:30,Junior:35,Pleno:45,Senior:55},
    conveniosValores:[],observacoes:""};
  const [fp,setFp]=useState(blankP);
  const sP=p=>setFp(f=>({...f,...p}));
  // Tabela de valores por convênio dentro do procedimento
  const [convValLine,setConvValLine]=useState({convenioId:"",valor:0});

  const TIPOS_CONVENIO=["Plano de Saúde","Particular","Municipal","Estadual","Federal","Internacional","Outro"];
  const CATS_PROC=["Sessão Terapêutica","Avaliação","Consulta","Procedimento","Grupo","Supervisão","Outro"];

  const saveC=()=>{
    if(!fc.nome.trim())return alert("Nome obrigatório");
    const obj={...fc,id:eC?.id||Date.now()};
    if(eC)setConvenios(a=>a.map(x=>x.id===eC.id?obj:x));
    else setConvenios(a=>[...a,obj]);
    setMC(false);showToast("✅ Convênio salvo","ok");
  };
  const delC=id=>{if(confirm("Excluir convênio?"))setConvenios(a=>a.filter(x=>x.id!==id));};

  const saveP=()=>{
    if(!fp.nome.trim())return alert("Nome obrigatório");
    const obj={...fp,id:eP?.id||Date.now()};
    if(eP)setProcedimentos(a=>a.map(x=>x.id===eP.id?obj:x));
    else setProcedimentos(a=>[...a,obj]);
    setMP(false);showToast("✅ Procedimento salvo","ok");
  };
  const delP=id=>{if(confirm("Excluir procedimento?"))setProcedimentos(a=>a.filter(x=>x.id!==id));};

  const addConvValor=()=>{
    if(!convValLine.convenioId)return;
    const exists=(fp.conveniosValores||[]).find(x=>x.convenioId===convValLine.convenioId);
    if(exists){sP({conveniosValores:(fp.conveniosValores||[]).map(x=>x.convenioId===convValLine.convenioId?{...convValLine}:x)});}
    else{sP({conveniosValores:[...(fp.conveniosValores||[]),{...convValLine}]});}
    setConvValLine({convenioId:"",valor:0});
  };
  const remConvValor=id=>sP({conveniosValores:(fp.conveniosValores||[]).filter(x=>x.convenioId!==id)});

  const filtC=convenios.filter(x=>!busca||x.nome.toLowerCase().includes(busca.toLowerCase())||( x.codigo||"").includes(busca));
  const filtP=procedimentos.filter(x=>!busca||x.nome.toLowerCase().includes(busca.toLowerCase())||(x.codigoTUSS||"").includes(busca));

  const tabBtn=(k,l,n)=>(
    <button key={k} className="btn tab-btn" onClick={()=>{setTab(k);setBusca("");}}
      style={{background:tab===k?"var(--na)":"transparent",color:tab===k?"#7c6af7":"var(--mt)",fontWeight:tab===k?800:500}}>
      {l}{n>0&&<span style={{marginLeft:4,background:"#7c6af7",color:"#fff",borderRadius:20,padding:"0 6px",fontSize:9,fontWeight:900}}>{n}</span>}
    </button>
  );

  return(<div className="page-wrap">
    <div className="page-head">
      <h1>🏥 Convênios & Procedimentos</h1>
      <div style={{display:"flex",gap:8}}>
        <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="🔍 Buscar..." style={{width:180,fontSize:12}}/>
        {tab==="convenios"
          ?<button className="btn primary" onClick={()=>{setEC(null);setFc(blankC);setMC(true);}}>+ Convênio</button>
          :<button className="btn primary" onClick={()=>{setEP(null);setFp(blankP);setMP(true);}}>+ Procedimento</button>}
      </div>
    </div>

    <div className="tab-bar">
      {tabBtn("convenios","🏥 Convênios",convenios.length)}
      {tabBtn("procedimentos","💊 Procedimentos",procedimentos.length)}
    </div>

    {/* ── CONVÊNIOS ── */}
    {tab==="convenios"&&<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:10}}>
      {filtC.map(c=>{
        const cor=c.ativo===false?"#64748b":"#7c6af7";
        const nProc=procedimentos.filter(p=>(p.conveniosValores||[]).some(v=>v.convenioId===String(c.id))).length;
        return(<div key={c.id} className="card" style={{padding:14,borderTop:"3px solid "+cor,opacity:c.ativo===false?.6:1}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontWeight:900,fontSize:14,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.nome}</div>
              <div style={{fontSize:10,color:"var(--mt)",marginTop:2,display:"flex",gap:8,flexWrap:"wrap"}}>
                {c.tipo&&<span style={{background:cor+"20",color:cor,padding:"1px 7px",borderRadius:20,fontWeight:700}}>{c.tipo}</span>}
                {c.codigo&&<span>Cód: {c.codigo}</span>}
                {c.registro&&<span>Reg: {c.registro}</span>}
              </div>
            </div>
            <div style={{display:"flex",gap:4,flexShrink:0}}>
              <button className="btn secondary small" onClick={()=>{setEC(c);setFc({...blankC,...c});setMC(true);}}>✏️</button>
              <button className="btn danger small" onClick={()=>delC(c.id)}>🗑️</button>
            </div>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:3,fontSize:11,color:"var(--mt)"}}>
            {c.cnpj&&<span>🏢 CNPJ: {c.cnpj}</span>}
            {c.telefone&&<span>📞 {c.telefone}</span>}
            {c.email&&<span>✉️ {c.email}</span>}
            {c.contato&&<span>👤 {c.contato}</span>}
          </div>
          <div style={{marginTop:8,paddingTop:7,borderTop:"1px solid var(--db)",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div style={{display:"flex",gap:6}}>
              {c.manual&&<a href={c.manual} target="_blank" rel="noreferrer" style={{fontSize:11,color:"#a78bfa",fontWeight:700}}>💡 Manual</a>}
              {c.site&&<a href={c.site} target="_blank" rel="noreferrer" style={{fontSize:11,color:"#34d399",fontWeight:700}}>🌐 Site</a>}
            </div>
            {nProc>0&&<span style={{fontSize:10,color:"var(--mt)"}}>{nProc} procedimento(s)</span>}
          </div>
          {c.vigenciaFim&&<div style={{marginTop:5,fontSize:10,color:"#f59e0b"}}>⏰ Vigência até {brDate(c.vigenciaFim)}</div>}
        </div>);
      })}
      {filtC.length===0&&<div style={{gridColumn:"1/-1",textAlign:"center",color:"var(--mt)",padding:40}}>Nenhum convênio encontrado.</div>}
    </div>}

    {/* ── PROCEDIMENTOS ── */}
    {tab==="procedimentos"&&<div style={{borderRadius:10,overflow:"hidden",border:"1px solid var(--sc)"}}>
      <div style={{display:"grid",gridTemplateColumns:"2fr 110px 70px 90px 80px 80px 72px",background:"var(--sx)",padding:"8px 12px",fontSize:10,fontWeight:900,color:"var(--mt)",letterSpacing:.5,gap:8}}>
        <div>PROCEDIMENTO</div><div>CATEGORIA</div><div>DUR.</div><div>VALOR BASE</div><div>HABILITADOS</div><div>TUSS</div><div></div>
      </div>
      {filtP.map((p,i)=>{
        const nConv=(p.conveniosValores||[]).length;
        return(<div key={p.id} style={{display:"grid",gridTemplateColumns:"2fr 110px 70px 90px 80px 80px 72px",padding:"10px 12px",borderTop:"1px solid var(--sc)",gap:8,alignItems:"center",background:i%2?"var(--gr)":""}}>
          <div>
            <div style={{fontWeight:800,fontSize:13}}>{p.nome}</div>
            <div style={{fontSize:10,color:"var(--mt)",marginTop:1}}>
              {p.especialidade&&<span style={{color:espCor(p.especialidade),fontWeight:700,marginRight:6}}>{p.especialidade}</span>}
              {nConv>0&&<span style={{color:"#a78bfa"}}>💊 {nConv} convênio(s)</span>}
            </div>
          </div>
          <div style={{fontSize:11,color:"var(--mt)"}}>{p.categoria}</div>
          <div style={{fontSize:11,color:"var(--mt)"}}>{p.duracao}min</div>
          <div style={{color:"#34d399",fontWeight:800,fontSize:13}}>{brl(p.valor||0)}</div>
          <div style={{display:"flex",gap:2,flexWrap:"wrap"}}>{NIVEL_REPASSE.map(n=>{const k=n.replace("ú","u");const pct=(p.niveisRepasse||{})[k];const ok=pct>0;return ok?<span key={n} style={{fontSize:9,padding:"1px 5px",borderRadius:20,background:"#10b98120",color:"#34d399",fontWeight:800,border:"1px solid #10b98140"}} title={n+": "+pct+"%"}>{n.slice(0,2)}</span>:null;})}</div>
          <div style={{fontSize:10,color:"var(--mt)",fontFamily:"monospace"}}>{p.codigoTUSS||"—"}</div>
          <div style={{display:"flex",gap:4}}>
            <button className="btn secondary small" onClick={()=>{setEP(p);setFp({...blankP,...p,niveisRepasse:p.niveisRepasse||{Trainee:30,Junior:35,Pleno:45,Senior:55},conveniosValores:p.conveniosValores||[]});setConvValLine({convenioId:"",valor:0});setMP(true);}}>✏️</button>
            <button className="btn danger small" onClick={()=>delP(p.id)}>🗑️</button>
          </div>
        </div>);
      })}
      {filtP.length===0&&<div style={{padding:32,textAlign:"center",color:"var(--mt)"}}>Nenhum procedimento encontrado.</div>}
    </div>}

    {/* ══ MODAL CONVÊNIO ══ */}
    {mC&&<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&setMC(false)}>
      <div className="modal" style={{width:640,maxHeight:"90vh",overflowY:"auto"}}>
        <div className="modal-head">
          <h2>{eC?"✏️ Editar":"+ Novo"} Convênio</h2>
          <button className="icon-btn" onClick={()=>setMC(false)}>×</button>
        </div>
        <div className="stack">
          <div className="section-box">
            <div className="section-title">🏥 Identificação</div>
            <div><label>Nome *</label><input value={fc.nome} onChange={e=>sC({nome:e.target.value})} autoFocus placeholder="Nome do convênio ou operadora"/></div>
            <div className="g3" style={{marginTop:8}}>
              <div><label>Tipo</label>
                <select value={fc.tipo} onChange={e=>sC({tipo:e.target.value})}>
                  {TIPOS_CONVENIO.map(t=><option key={t}>{t}</option>)}
                </select>
              </div>
              <div><label>Código ANS / Interno</label><input value={fc.codigo||""} onChange={e=>sC({codigo:e.target.value})} placeholder="Ex: 123456"/></div>
              <div><label>Registro</label><input value={fc.registro||""} onChange={e=>sC({registro:e.target.value})} placeholder="Nº registro"/></div>
            </div>
            <div style={{marginTop:8}}><label>CNPJ</label><input value={fc.cnpj||""} onChange={e=>sC({cnpj:e.target.value})} placeholder="00.000.000/0000-00"/></div>
          </div>

          <div className="section-box">
            <div className="section-title">📞 Contato</div>
            <div className="g2">
              <div><label>Responsável / Contato</label><input value={fc.contato||""} onChange={e=>sC({contato:e.target.value})}/></div>
              <div><label>Telefone</label><input value={fc.telefone||""} onChange={e=>sC({telefone:e.target.value})} placeholder="(00) 0000-0000"/></div>
            </div>
            <div className="g2" style={{marginTop:8}}>
              <div><label>E-mail</label><input type="email" value={fc.email||""} onChange={e=>sC({email:e.target.value})} placeholder="contato@convenio.com.br"/></div>
              <div><label>Site</label><input value={fc.site||""} onChange={e=>sC({site:e.target.value})} placeholder="https://..."/></div>
            </div>
          </div>

          <div className="section-box">
            <div className="section-title">📋 Contrato & Vigência</div>
            <div><label>URL do Manual de Autorização</label><input value={fc.manual||""} onChange={e=>sC({manual:e.target.value})} placeholder="https://..."/></div>
            <div className="g2" style={{marginTop:8}}>
              <div><label>Vigência Início</label><input type="date" value={fc.vigenciaIni||""} onChange={e=>sC({vigenciaIni:e.target.value})}/></div>
              <div><label>Vigência Fim</label><input type="date" value={fc.vigenciaFim||""} onChange={e=>sC({vigenciaFim:e.target.value})}/></div>
            </div>
            <div style={{marginTop:8}}>
              <label style={{display:"flex",alignItems:"center",gap:7,cursor:"pointer",textTransform:"none",fontSize:12}}>
                <input type="checkbox" checked={fc.ativo!==false} onChange={e=>sC({ativo:e.target.checked})}/>
                <span style={{fontWeight:700}}>Convênio ativo</span>
              </label>
            </div>
          </div>

          <div><label>Observações</label><textarea rows={2} value={fc.observacoes||""} onChange={e=>sC({observacoes:e.target.value})}/></div>
          <div className="actions">
            <button className="btn secondary" onClick={()=>setMC(false)}>Cancelar</button>
            <button className="btn primary" onClick={saveC}>💾 Salvar Convênio</button>
          </div>
        </div>
      </div>
    </div>}

    {/* ══ MODAL PROCEDIMENTO ══ */}
    {mP&&<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&setMP(false)}>
      <div className="modal" style={{width:660,maxHeight:"90vh",overflowY:"auto"}}>
        <div className="modal-head">
          <h2>{eP?"✏️ Editar":"+ Novo"} Procedimento</h2>
          <button className="icon-btn" onClick={()=>setMP(false)}>×</button>
        </div>
        <div className="stack">
          <div className="section-box">
            <div className="section-title">💊 Identificação</div>
            <div><label>Nome *</label><input value={fp.nome} onChange={e=>sP({nome:e.target.value})} autoFocus placeholder="Nome do procedimento"/></div>
            <div className="g3" style={{marginTop:8}}>
              <div><label>Categoria</label>
                <select value={fp.categoria} onChange={e=>sP({categoria:e.target.value})}>
                  {CATS_PROC.map(c=><option key={c}>{c}</option>)}
                </select>
              </div>
              <div><label>Especialidade</label>
                <select value={fp.especialidade||""} onChange={e=>sP({especialidade:e.target.value})}>
                  <option value="">Todas</option>
                  {ESPECIALIDADES_LIST.map(e=><option key={e}>{e}</option>)}
                </select>
              </div>
              <div><label>Código TUSS</label><input value={fp.codigoTUSS||""} onChange={e=>sP({codigoTUSS:e.target.value})} placeholder="Ex: 50000470"/></div>
            </div>
          </div>

          <div className="section-box">
            <div className="section-title">💰 Valores & Habilitação por Nível</div>
            <div className="g2" style={{marginBottom:8}}>
              <div><label>Duração padrão</label>
                <select value={fp.duracao} onChange={e=>sP({duracao:Number(e.target.value)})}>
                  {TEMPOS_SESSAO.map(t=><option key={t} value={t}>{t} min</option>)}
                </select>
              </div>
              <div><label>Valor base (R$)</label>
                <input type="number" min={0} step={0.01} value={fp.valor} onChange={e=>sP({valor:Number(e.target.value)})}/>
              </div>
            </div>
            <div style={{fontSize:11,color:"var(--mt)",marginBottom:8}}>
              Define quais níveis de repasse podem executar este procedimento e qual % cada nível recebe.
              Deixe 0% para <b>desabilitar</b> o nível.
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8}}>
              {NIVEL_REPASSE.map(nivel=>{
                const key=nivel.replace("ú","u");
                const pct=(fp.niveisRepasse||{})[key]??0;
                const habilitado=pct>0;
                return(<div key={nivel} style={{background:habilitado?"#10b98115":"var(--sx)",border:"1.5px solid "+(habilitado?"#10b98150":"var(--cpb)"),borderRadius:10,padding:"10px 8px",textAlign:"center"}}>
                  <div style={{fontWeight:900,fontSize:12,color:habilitado?"#34d399":"var(--mt)",marginBottom:6}}>{nivel}</div>
                  <input type="number" min={0} max={100} step={1} value={pct}
                    onChange={e=>{const v=Number(e.target.value);sP({niveisRepasse:{...(fp.niveisRepasse||{}), [key]:v}});}}
                    style={{width:"100%",textAlign:"center",fontWeight:800,fontSize:14,color:habilitado?"#34d399":"var(--mt)",background:"transparent",border:"none",outline:"none",padding:0}}/>
                  <div style={{fontSize:9,color:"var(--mt)",marginTop:2}}>
                    {habilitado?<><span style={{color:"#34d399",fontWeight:700}}>✓ Habilitado</span><br/>{brl((fp.valor||0)*pct/100)}/sessão</>:<span style={{color:"#ef4444",fontWeight:700}}>✗ Bloqueado</span>}
                  </div>
                </div>);
              })}
            </div>
          </div>

          <div className="section-box" style={{borderColor:"#7c6af730"}}>
            <div className="section-title" style={{color:"#a78bfa"}}>🏥 Valores por Convênio</div>
            <div style={{fontSize:11,color:"var(--mt)",marginBottom:8}}>Defina valores específicos por convênio. Se não informado, usa o valor base acima.</div>
            <div style={{display:"flex",gap:8,alignItems:"flex-end",marginBottom:10}}>
              <div style={{flex:2}}>
                <label>Convênio</label>
                <select value={convValLine.convenioId} onChange={e=>setConvValLine(l=>({...l,convenioId:e.target.value}))}>
                  <option value="">Selecione...</option>
                  {convenios.map(c=><option key={c.id} value={String(c.id)}>{c.nome}</option>)}
                </select>
              </div>
              <div style={{flex:1}}>
                <label>Valor (R$)</label>
                <input type="number" min={0} step={0.01} value={convValLine.valor} onChange={e=>setConvValLine(l=>({...l,valor:Number(e.target.value)}))}/>
              </div>
              <button className="btn secondary" style={{flexShrink:0}} onClick={addConvValor}>+ Adicionar</button>
            </div>
            {(fp.conveniosValores||[]).length>0&&<div style={{display:"flex",flexDirection:"column",gap:4}}>
              {(fp.conveniosValores||[]).map(v=>{
                const conv=convenios.find(c=>String(c.id)===String(v.convenioId));
                return(<div key={v.convenioId} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"5px 10px",background:"var(--sx)",borderRadius:7,border:"1px solid var(--cb)"}}>
                  <span style={{fontSize:12,fontWeight:700}}>{conv?.nome||v.convenioId}</span>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <span style={{color:"#34d399",fontWeight:800,fontSize:13}}>{brl(v.valor)}</span>
                    <button onClick={()=>remConvValor(v.convenioId)} style={{background:"none",border:"none",color:"#f87171",cursor:"pointer",fontSize:14}}>✕</button>
                  </div>
                </div>);
              })}
            </div>}
          </div>

          <div><label>Observações</label><textarea rows={2} value={fp.observacoes||""} onChange={e=>sP({observacoes:e.target.value})}/></div>
          <div className="actions">
            <button className="btn secondary" onClick={()=>setMP(false)}>Cancelar</button>
            <button className="btn primary" onClick={saveP}>💾 Salvar Procedimento</button>
          </div>
        </div>
      </div>
    </div>}
  </div>);
}


// ═══════════════════════════════════════════════════════════════════════════════
// MANUAIS — com opção de arquivo + importação
// ═══════════════════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════════════════
// MANUAIS — Convênios + Contratos + Modelos de Evolução
// ═══════════════════════════════════════════════════════════════════════════════
function ManuaisPage({manuais,setManuais,showToast,templatePaciente,setTemplatePaciente,templateProfissional,setTemplateProfissional,modelosEvolucao,setModelosEvolucao}){
  const [tab,setTab]=useState("convenios");
  const [modal,setModal]=useState(false);const[editing,setEditing]=useState(null);
  const [form,setForm]=useState({titulo:"",descricao:"",url:"",arquivo:null,data:hoje_str});
  const save=()=>{if(!form.titulo.trim())return alert("Título obrigatório");if(editing)setManuais(a=>a.map(x=>x.id===editing.id?{...form,id:editing.id}:x));else setManuais(a=>[...a,{...form,id:Date.now()}]);showToast("✅ Salvo","ok");setModal(false);};
  const importarArq=e=>{const file=e.target.files[0];if(!file)return;const url=URL.createObjectURL(file);setForm(f=>({...f,arquivo:{nome:file.name,url},titulo:f.titulo||file.name.replace(/\.[^.]+$/,""),url:f.url||url}));};

  // Modelos de evolução editáveis
  const [espSel,setEspSel]=useState(ESPECIALIDADES_LIST[0]);
  const [modeloEdit,setModeloEdit]=useState(null);
  const [modeloForm,setModeloForm]=useState({titulo:"",texto:""});
  const modEsp=(modelosEvolucao||MODELOS_EVOLUCAO)[espSel]||[];
  const salvarModelo=()=>{
    if(!modeloForm.titulo.trim()||!modeloForm.texto.trim())return alert("Preencha título e texto");
    const base={...(modelosEvolucao||MODELOS_EVOLUCAO)};
    if(!base[espSel])base[espSel]=[];
    if(modeloEdit){base[espSel]=base[espSel].map(m=>m.id===modeloEdit.id?{...m,...modeloForm}:m);}
    else{base[espSel]=[...base[espSel],{id:"custom_"+Date.now(),titulo:modeloForm.titulo,texto:modeloForm.texto}];}
    setModelosEvolucao(base);setModeloEdit(null);setModeloForm({titulo:"",texto:""});showToast("✅ Modelo salvo","ok");
  };
  const excluirModelo=id=>{
    if(!confirm("Excluir modelo?"))return;
    const base={...(modelosEvolucao||MODELOS_EVOLUCAO)};
    base[espSel]=base[espSel].filter(m=>m.id!==id);
    setModelosEvolucao(base);showToast("🗑️ Removido","ok");
  };

  const tabBtn=(k,l,badge)=>(
    <button key={k} className="btn tab-btn" onClick={()=>setTab(k)}
      style={{background:tab===k?"var(--na)":"transparent",color:tab===k?"#7c6af7":"var(--mt)",fontWeight:tab===k?800:500,display:"flex",gap:5,alignItems:"center"}}>
      {l}{badge&&<span style={{padding:"1px 6px",borderRadius:20,fontSize:9,background:"#7c6af7",color:"#fff",fontWeight:900}}>{badge}</span>}
    </button>
  );

  return(<div className="page-wrap">
    <div className="page-head">
      <h1>📚 Biblioteca de Modelos</h1>
      {tab==="convenios"&&<button className="btn primary" onClick={()=>{setEditing(null);setForm({titulo:"",descricao:"",url:"",arquivo:null,data:hoje_str});setModal(true);}}>+ Novo Manual</button>}
    </div>

    <div className="tab-bar" style={{marginBottom:16}}>
      {tabBtn("convenios","📋 Manuais de Convênio",manuais.length||null)}
      {tabBtn("contrato_pac","👤 Contrato Paciente")}
      {tabBtn("contrato_prof","🩺 Contrato Profissional")}
      {tabBtn("modelos_ev","📝 Modelos de Evolução")}
      {tabBtn("govbr","🏛️ Assinatura Gov.br")}
    </div>

    {/* ── MANUAIS DE CONVÊNIO ── */}
    {tab==="convenios"&&<>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(250px,1fr))",gap:11}}>
        {manuais.map(m=><div key={m.id} className="card" style={{padding:14,borderTop:"3px solid #7c6af7"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div><div style={{fontWeight:800}}>📋 {m.titulo}</div><div className="muted" style={{fontSize:11,marginTop:2}}>{m.descricao}</div><div className="muted" style={{fontSize:10,marginTop:2}}>📅 {brDate(m.data)}</div></div>
            <div style={{display:"flex",gap:4}}><button className="btn secondary small" onClick={()=>{setEditing(m);setForm({...m});setModal(true);}}>✏️</button><button className="btn danger small" onClick={()=>{if(confirm("Excluir?"))setManuais(a=>a.filter(x=>x.id!==m.id));}}>🗑️</button></div>
          </div>
          <div style={{marginTop:8,display:"flex",gap:6}}>
            {m.url&&<a href={m.url} target="_blank" rel="noreferrer" style={{fontSize:12,color:"#a78bfa",fontWeight:700}}>🔗 Abrir URL</a>}
            {m.arquivo&&<a href={m.arquivo.url} target="_blank" rel="noreferrer" style={{fontSize:12,color:"#34d399",fontWeight:700}}>📄 {m.arquivo.nome}</a>}
          </div>
        </div>)}
        {manuais.length===0&&<div style={{gridColumn:"1/-1",textAlign:"center",color:"var(--mt)",padding:"40px"}}>Nenhum manual cadastrado.</div>}
      </div>
    </>}

    {/* ── CONTRATO PACIENTE ── */}
    {(tab==="contrato_pac"||tab==="contrato_prof")&&(()=>{
      const isPac=tab==="contrato_pac";
      const tmpl=isPac?templatePaciente:templateProfissional;
      const setTmpl=isPac?setTemplatePaciente:setTemplateProfissional;
      const variaveis=isPac
        ?["nome","cpf","nascimento","celular","convenio","plano","cidade","resp1Nome","resp1Cpf","resp1Whatsapp"]
        :["nome","especialidades","conselho","numConselho","nivelRepasse","cidade"];
      return(<div className="stack">
        <div style={{padding:"10px 14px",background:"#7c6af715",borderRadius:10,border:"1px solid #7c6af730"}}>
          <div style={{fontWeight:800,fontSize:12,color:"#a78bfa",marginBottom:6}}>💡 Variáveis disponíveis</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
            {variaveis.map(v=>(
              <code key={v} style={{background:"var(--sx)",padding:"2px 7px",borderRadius:5,fontSize:11,color:"#93c5fd",cursor:"copy",border:"1px solid var(--sc)"}}
                title={"Clique para copiar"}
                onClick={()=>{navigator.clipboard?.writeText("{{"+v+"}}");showToast("📋 Copiado: {{"+v+"}}","ok");}}>
                {"{{"}{v}{"}}"}
              </code>
            ))}
          </div>
          <div style={{fontSize:10,color:"var(--mt)",marginTop:6}}>Clique na variável para copiar. Cole no modelo onde quiser que o dado do {isPac?"paciente":"profissional"} apareça automaticamente.</div>
        </div>
        <div>
          <label style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span>{isPac?"📄 Modelo do Contrato de Paciente":"📄 Modelo do Contrato de Profissional"}</span>
            <button className="btn secondary small" onClick={()=>setTmpl(isPac?TEMPLATE_CONTRATO_PACIENTE:TEMPLATE_CONTRATO_PROFISSIONAL)}>🔄 Restaurar padrão</button>
          </label>
          <textarea rows={28} value={tmpl} onChange={e=>setTmpl(e.target.value)}
            style={{fontFamily:"monospace",fontSize:12,lineHeight:1.7,marginTop:6}}/>
        </div>
        <div style={{padding:"8px 12px",background:"#10b98115",borderRadius:8,border:"1px solid #10b98130",fontSize:11,color:"#34d399"}}>
          ✅ As alterações são salvas automaticamente e usadas em todos os contratos gerados para {isPac?"pacientes":"profissionais"}.
        </div>
      </div>);
    })()}

    {/* ── MODELOS DE EVOLUÇÃO ── */}
    {tab==="modelos_ev"&&<div className="stack">
      <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
        <label style={{fontSize:12,fontWeight:700,margin:0}}>Especialidade:</label>
        <select value={espSel} onChange={e=>{setEspSel(e.target.value);setModeloEdit(null);setModeloForm({titulo:"",texto:""});}} style={{width:220,fontSize:12}}>
          {ESPECIALIDADES_LIST.map(e=><option key={e}>{e}</option>)}
        </select>
        <span style={{fontSize:11,color:"var(--mt)"}}>{modEsp.length} modelo(s)</span>
      </div>

      {/* Editor de modelo */}
      <div className="section-box" style={{borderColor:"#7c6af740"}}>
        <div className="section-title" style={{color:"#a78bfa"}}>{modeloEdit?"✏️ Editando: "+modeloEdit.titulo:"➕ Novo Modelo — "+espSel}</div>
        <div className="stack" style={{gap:8}}>
          <div><label>Título do Modelo</label>
            <input value={modeloForm.titulo} onChange={e=>setModeloForm(f=>({...f,titulo:e.target.value}))} placeholder="Ex: Avaliação Inicial, Sessão de Mando..."/>
          </div>
          <div><label>Texto do Modelo</label>
            <textarea rows={6} value={modeloForm.texto} onChange={e=>setModeloForm(f=>({...f,texto:e.target.value}))}
              placeholder={"Texto base da evolução para "+espSel+". Use X, Y, Z como placeholders para valores a preencher..."}
              style={{fontFamily:"inherit",lineHeight:1.6}}/>
          </div>
          <div style={{display:"flex",gap:6}}>
            {modeloEdit&&<button className="btn secondary" onClick={()=>{setModeloEdit(null);setModeloForm({titulo:"",texto:""});}}>Cancelar</button>}
            <button className="btn primary" onClick={salvarModelo}>{modeloEdit?"💾 Salvar Alteração":"➕ Adicionar Modelo"}</button>
          </div>
        </div>
      </div>

      {/* Lista de modelos */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:8}}>
        {modEsp.length===0&&<div style={{gridColumn:"1/-1",textAlign:"center",color:"var(--mt)",padding:24}}>Nenhum modelo para {espSel}.</div>}
        {modEsp.map(m=>{
          const isCustom=String(m.id).startsWith("custom_");
          return(<div key={m.id} className="card" style={{padding:12,borderLeft:"3px solid "+espCor(espSel)}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:6,marginBottom:6}}>
              <div style={{fontWeight:800,fontSize:12,color:espCor(espSel),flex:1}}>{m.titulo}</div>
              <div style={{display:"flex",gap:3,flexShrink:0}}>
                {isCustom&&<span style={{fontSize:9,padding:"1px 5px",borderRadius:20,background:"#f59e0b20",color:"#f59e0b",fontWeight:800}}>custom</span>}
                <button onClick={()=>{setModeloEdit(m);setModeloForm({titulo:m.titulo,texto:m.texto});}} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,color:"var(--mt)"}}>✏️</button>
                <button onClick={()=>excluirModelo(m.id)} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,color:"#f87171"}}>🗑️</button>
              </div>
            </div>
            <div style={{fontSize:11,color:"var(--mt)",lineHeight:1.5,overflow:"hidden",display:"-webkit-box",WebkitLineClamp:3,WebkitBoxOrient:"vertical"}}>{m.texto}</div>
          </div>);
        })}
      </div>
    </div>}

    {/* ── ASSINATURA GOV.BR ── */}
    {tab==="govbr"&&<div className="stack">
      <div style={{background:"linear-gradient(135deg,#1351B4,#2670E8)",borderRadius:14,padding:"20px 24px",color:"#fff",marginBottom:4}}>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:8}}>
          <div style={{width:48,height:48,borderRadius:12,background:"rgba(255,255,255,0.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:24}}>🏛️</div>
          <div>
            <div style={{fontWeight:900,fontSize:18,fontFamily:"'DM Serif Display',serif"}}>Assinatura Digital Gov.br</div>
            <div style={{fontSize:12,opacity:.85,marginTop:2}}>Plataforma oficial do Governo Federal para assinatura eletrônica com validade jurídica</div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {["ICP-Brasil","MP 2.200-2","Lei 14.063/2020","Validade Jurídica"].map(b=>(
            <span key={b} style={{padding:"2px 10px",borderRadius:20,background:"rgba(255,255,255,0.2)",fontSize:10,fontWeight:700}}>{b}</span>
          ))}
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <div className="section-box" style={{borderColor:"#1351B430"}}>
          <div className="section-title" style={{color:"#2670E8"}}>📋 O que é</div>
          <div style={{fontSize:12,lineHeight:1.7,color:"var(--tx)"}}>
            {"O Gov.br é a plataforma de assinatura eletrônica do Governo Federal Brasileiro. Permite assinar documentos digitalmente com validade jurídica equivalente à assinatura manuscrita, conforme a Lei nº 14.063/2020 e a MP 2.200-2/2001."}
          </div>
          <div style={{fontSize:12,lineHeight:1.7,color:"var(--tx)",marginTop:8}}>
            {"Para clínicas, é ideal para assinar contratos de pacientes, contratos de profissionais, laudos, relatórios médicos e documentos institucionais."}
          </div>
        </div>

        <div className="section-box" style={{borderColor:"#22c55e30"}}>
          <div className="section-title" style={{color:"#22c55e"}}>✅ Níveis de Conta</div>
          {[
            ["Bronze","Cadastro básico","Documentos simples","#cd7f32"],
            ["Prata","Validação por banco ou biometria","Contratos e prontuários","#94a3b8"],
            ["Ouro","Certificado ICP-Brasil ou reconhecimento facial","Documentos com máxima validade","#f59e0b"],
          ].map(([nivel,como,uso,cor])=>(
            <div key={nivel} style={{display:"flex",gap:8,alignItems:"flex-start",padding:"7px 0",borderBottom:"1px solid var(--db)"}}>
              <span style={{fontWeight:900,fontSize:11,color:cor,minWidth:48}}>{nivel}</span>
              <div>
                <div style={{fontSize:11,fontWeight:700,color:"var(--tx)"}}>{como}</div>
                <div style={{fontSize:10,color:"var(--mt)"}}>{uso}</div>
              </div>
            </div>
          ))}
        </div>
      </div>

      <div className="section-box" style={{borderColor:"#1351B430"}}>
        <div className="section-title" style={{color:"#2670E8"}}>🪜 Passo a passo — Como assinar um documento</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:10,marginTop:4}}>
          {[
            ["1","Gerar o PDF","No Cuide, gere o contrato do paciente ou profissional clicando em 📄 Contrato","#7c6af7"],
            ["2","Acessar o portal","Acesse assinador.iti.br ou o app Gov.br no celular","#6366f1"],
            ["3","Fazer login","Entre com seu CPF e senha Gov.br (conta Prata ou Ouro recomendado)","#8b5cf6"],
            ["4","Enviar o documento","Clique em Assinar documento e faça upload do PDF gerado","#a78bfa"],
            ["5","Posicionar assinatura","Escolha onde a assinatura aparecerá no documento","#c084fc"],
            ["6","Confirmar","Confirme com senha, biometria ou token e baixe o PDF assinado","#22c55e"],
          ].map(([n,titulo,desc,cor])=>(
            <div key={n} style={{padding:"10px 12px",borderRadius:9,background:"var(--sx)",border:"1px solid var(--cb)"}}>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:5}}>
                <div style={{width:22,height:22,borderRadius:"50%",background:cor,color:"#fff",fontWeight:900,fontSize:11,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{n}</div>
                <span style={{fontWeight:800,fontSize:12,color:"var(--tx)"}}>{titulo}</span>
              </div>
              <div style={{fontSize:11,color:"var(--mt)",lineHeight:1.5}}>{desc}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <div className="section-box" style={{borderColor:"#f59e0b30"}}>
          <div className="section-title" style={{color:"#f59e0b"}}>⚖️ Validade Jurídica</div>
          <div style={{fontSize:12,lineHeight:1.7,color:"var(--tx)"}}>{
            ["Documentos assinados pelo Gov.br têm validade jurídica plena conforme:",
             "• Lei 14.063/2020 — assinaturas eletrônicas em atos jurídicos",
             "• MP 2.200-2/2001 — infraestrutura ICP-Brasil",
             "• CFM/CFP/CREFITO — reconhecem assinatura eletrônica qualificada",
             "Recomenda-se nível Prata ou Ouro para contratos terapêuticos."
            ].map((t,i)=><div key={i} style={{marginBottom:3}}>{t}</div>)
          }</div>
        </div>

        <div className="section-box" style={{borderColor:"#34d39930"}}>
          <div className="section-title" style={{color:"#34d399"}}>🔗 Links Úteis</div>
          <div style={{display:"flex",flexDirection:"column",gap:7}}>
            {[
              ["🏛️ Portal Gov.br","https://www.gov.br/pt-br/servicos/assinatura-eletronica","Página oficial do serviço"],
              ["✍️ Assinador ITI","https://assinador.iti.br","Ferramenta de assinatura online"],
              ["📱 App Gov.br","https://www.gov.br/governodigital/pt-br/conta-gov-br/conta-gov-br","Download do aplicativo"],
              ["📖 Lei 14.063/2020","https://www.planalto.gov.br/ccivil_03/_ato2019-2022/2020/lei/l14063.htm","Legislação de assinaturas eletrônicas"],
              ["🎓 Tutorial oficial","https://www.iti.br/icp-brasil/assinatura-digital","Guia completo do ITI"],
            ].map(([label,url,desc])=>(
              <a key={label} href={url} target="_blank" rel="noreferrer"
                style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",borderRadius:8,background:"var(--sx)",border:"1px solid var(--cb)",textDecoration:"none",color:"var(--tx)"}}>
                <span style={{fontSize:13}}>{label.split(" ")[0]}</span>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:11,color:"#a78bfa"}}>{label.split(" ").slice(1).join(" ")}</div>
                  <div style={{fontSize:10,color:"var(--mt)"}}>{desc}</div>
                </div>
                <span style={{fontSize:10,color:"var(--mt)"}}>↗</span>
              </a>
            ))}
          </div>
        </div>
      </div>

      <div style={{padding:"12px 16px",background:"#1351B410",borderRadius:10,border:"1px solid #1351B430",fontSize:11,color:"var(--mt)",lineHeight:1.6}}>
        💡 <b style={{color:"var(--tx)"}}>Dica para clínicas:</b> Crie uma conta Gov.br nível Prata para todos os profissionais. Use o Cuide para gerar os PDFs de contrato, envie pelo assinador.iti.br, e arquive os documentos assinados na aba de Documentos do cadastro do profissional ou paciente.
      </div>
    </div>}

    {/* Modal manuais convênio */}
    {modal&&<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&setModal(false)}>
      <div className="modal"><div className="modal-head"><h2>{editing?"✏️ Editar":"+ Novo"} Manual</h2><button className="icon-btn" onClick={()=>setModal(false)}>×</button></div>
        <div className="stack">
          <div><label>Título *</label><input value={form.titulo} onChange={e=>setForm(f=>({...f,titulo:e.target.value}))} autoFocus/></div>
          <div><label>Descrição</label><textarea rows={2} value={form.descricao||""} onChange={e=>setForm(f=>({...f,descricao:e.target.value}))}/></div>
          <div><label>URL (link externo)</label><input value={form.url||""} onChange={e=>setForm(f=>({...f,url:e.target.value}))} placeholder="https://..."/></div>
          <div>
            <label>📎 Importar Arquivo (PDF, imagem...)</label>
            <label className="upload-zone" style={{display:"block",cursor:"pointer"}}>
              {form.arquivo?<span style={{color:"#34d399"}}>✅ {form.arquivo.nome}</span>:<span>Clique para selecionar arquivo</span>}
              <input type="file" style={{display:"none"}} accept=".pdf,.doc,.docx,.png,.jpg,.jpeg" onChange={importarArq}/>
            </label>
          </div>
          <div><label>Data</label><input type="date" value={form.data||hoje_str} onChange={e=>setForm(f=>({...f,data:e.target.value}))}/></div>
          <div className="actions"><button className="btn secondary" onClick={()=>setModal(false)}>Cancelar</button><button className="btn primary" onClick={save}>Salvar</button></div>
        </div>
      </div>
    </div>}
  </div>);
}


// ═══════════════════════════════════════════════════════════════════════════════
// CHAMADOS — com Anexo VII, status e prioridade
// ═══════════════════════════════════════════════════════════════════════════════
function ChamadosPage({chamados,setChamados,showToast,fila,setFila,pacientes}){
  const [editando,setEditando]=useState(null);
  const [filtroSt,setFiltroSt]=useState("todos");
  const [filtroSet,setFiltroSet]=useState("");
  const [filtroPrior,setFiltroPrior]=useState("");
  const [novoModal,setNovoModal]=useState(false);
  const blankNovo={setor:"Agendamento",tipo:"manual",padraoId:"01",camposValues:{},assunto:"",descricao:"",pacienteId:"",especialidade:"",preferencia:"Manhã",prioridade:"Normal"};
  const [novoForm,setNovoForm]=useState(blankNovo);
  const snf=p=>setNovoForm(f=>({...f,...p}));

  const filtered=chamados.filter(c=>{
    if(filtroSt!=="todos"&&c.status!==filtroSt)return false;
    if(filtroSet&&c.setor!==filtroSet)return false;
    if(filtroPrior&&c.prioridade!==filtroPrior)return false;
    return true;
  });

  const buildDescricao=(padraoId,camposValues)=>{
    const tmpl=PADRAO_CHAMADO.find(p=>p.id===padraoId);
    if(!tmpl)return "";
    return tmpl.campos.map(c=>c+": "+(camposValues[c]||"")).join("\n");
  };

  const abrirNovo=()=>{
    const num=String(Math.floor(100000+Math.random()*900000));
    const isPadrao=novoForm.tipo==="padrao";
    const descr=isPadrao?buildDescricao(novoForm.padraoId,novoForm.camposValues):novoForm.descricao;
    const tmplLabel=isPadrao?(PADRAO_CHAMADO.find(p=>p.id===novoForm.padraoId)?.label||"Chamado"):(novoForm.assunto||"Chamado");
    const tipoCh=isPadrao?"padrao_"+novoForm.padraoId:"manual";
    setChamados(a=>[...a,{id:Date.now(),numero:num,setor:novoForm.setor,tipo:tipoCh,nome:tmplLabel,descricao:descr,data:hoje_str,status:"aberto",prioridade:novoForm.prioridade||"Normal",resp:""}]);
    if(novoForm.setor==="Agendamento"&&novoForm.pacienteId&&novoForm.especialidade){
      setFila&&setFila(f=>[...f,{id:Date.now()+1,pacienteId:Number(novoForm.pacienteId),especialidade:novoForm.especialidade,preferencia:novoForm.preferencia||"Manhã",observacao:descr}]);
      showToast("📨 Chamado #"+num+" · ⏳ Adicionado à Fila","ok");
    } else {
      showToast("📨 Chamado #"+num+" aberto","ok");
    }
    setNovoModal(false);setNovoForm(blankNovo);
  };

  const isPadrao=novoForm.tipo==="padrao";
  const padraoSel=PADRAO_CHAMADO.find(p=>p.id===novoForm.padraoId)||PADRAO_CHAMADO[0];

  // Counts
  const countSt=st=>chamados.filter(c=>c.status===st).length;

  return(<div className="page-wrap">
    <div className="page-head"><h1>📨 Chamados</h1>
      <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
        <span className="muted" style={{fontSize:12}}>{chamados.filter(c=>c.status==="aberto"||c.status==="andamento").length} em aberto</span>
        <button className="btn secondary" style={{fontSize:11,borderColor:"#34d399",color:"#34d399"}}
          onClick={()=>{const num=String(Math.floor(100000+Math.random()*900000));setChamados(a=>[...a,{id:Date.now(),numero:num,setor:"Agendamento",tipo:"novo_paciente",nome:"Novo Paciente",descricao:"Solicitação de cadastro de novo paciente",data:hoje_str,status:"aberto",prioridade:"Normal",resp:""}]);showToast("📨 Chamado #"+num+" — Novo Paciente","ok");}}>
          👤 Novo Paciente
        </button>
        <button className="btn secondary" style={{fontSize:11,borderColor:"#a78bfa",color:"#a78bfa"}}
          onClick={()=>{const num=String(Math.floor(100000+Math.random()*900000));setChamados(a=>[...a,{id:Date.now(),numero:num,setor:"Agendamento",tipo:"novo_agendamento",nome:"Novo Agendamento",descricao:"Solicitação de agendamento",data:hoje_str,status:"aberto",prioridade:"Normal",resp:""}]);showToast("📨 Chamado #"+num+" — Novo Agendamento","ok");}}>
          📅 Novo Agendamento
        </button>
        <button className="btn primary" onClick={()=>setNovoModal(true)}>+ Novo Chamado</button>
      </div>
    </div>

    {/* Status pills summary */}
    <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap"}}>
      {STATUS_CHAMADO.map(s=>(
        <button key={s} onClick={()=>setFiltroSt(filtroSt===s?"todos":s)}
          style={{padding:"4px 11px",borderRadius:20,fontSize:11,fontWeight:800,cursor:"pointer",border:"1.5px solid "+(filtroSt===s?COR_CHAMADO[s]:"var(--cpb)"),background:filtroSt===s?COR_CHAMADO[s]+"20":"transparent",color:filtroSt===s?COR_CHAMADO[s]:"var(--mt)",transition:".12s"}}>
          {LABEL_CHAMADO[s]} <span style={{fontWeight:900}}>({countSt(s)})</span>
        </button>
      ))}
      {filtroSt!=="todos"&&<button onClick={()=>setFiltroSt("todos")} style={{padding:"4px 9px",borderRadius:20,fontSize:11,fontWeight:700,cursor:"pointer",border:"1.5px solid var(--cpb)",background:"transparent",color:"var(--mt)"}}>✕ Limpar</button>}
    </div>

    {/* Filtros */}
    <div style={{display:"flex",gap:7,marginBottom:12,flexWrap:"wrap"}}>
      <select value={filtroSet} onChange={e=>setFiltroSet(e.target.value)} style={{width:"auto",minWidth:140}}>
        <option value="">Todos setores</option>{SETORES_CHAMADO.map(s=><option key={s}>{s}</option>)}
      </select>
      <select value={filtroPrior} onChange={e=>setFiltroPrior(e.target.value)} style={{width:"auto",minWidth:120}}>
        <option value="">Todas prioridades</option>{PRIORIDADE_CHAMADO.map(p=><option key={p}>{p}</option>)}
      </select>
    </div>

    <div className="stack">
      {filtered.map(ch=>{
        const cor=COR_CHAMADO[ch.status]||"#475569";
        const pCor=COR_PRIORIDADE[ch.prioridade]||"#94a3b8";
        return(<div key={ch.id} className="card" style={{padding:13,borderLeft:"3px solid "+cor}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div style={{flex:1}}>
              <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap",marginBottom:4}}>
                <span style={{fontWeight:900,color:"var(--tx)"}}>#{ch.numero}</span>
                <span style={{padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:800,background:cor+"25",color:cor}}>{LABEL_CHAMADO[ch.status]||ch.status}</span>
                <span style={{padding:"2px 7px",borderRadius:20,fontSize:10,fontWeight:800,background:pCor+"25",color:pCor,border:"1px solid "+pCor+"40"}}>⚑ {ch.prioridade||"Normal"}</span>
                <span style={{padding:"2px 7px",borderRadius:20,fontSize:10,fontWeight:800,background:"var(--na)",color:"#a78bfa"}}>{ch.setor}</span>
                <span className="muted" style={{fontSize:10}}>{brDate(ch.data)}</span>
                {ch.tipo&&ch.tipo.startsWith("padrao_")&&<span style={{padding:"2px 6px",borderRadius:20,fontSize:9,fontWeight:800,background:"#7c6af720",color:"#a78bfa"}}>📋 Padronizado</span>}
              </div>
              <div style={{fontWeight:700,color:"var(--tx)",marginBottom:ch.descricao?4:0}}>{ch.nome}</div>
              {ch.descricao&&<pre style={{fontSize:10,color:"var(--mt)",whiteSpace:"pre-wrap",fontFamily:"inherit",background:"var(--sx)",padding:"5px 8px",borderRadius:6,maxHeight:80,overflow:"auto",margin:0}}>{ch.descricao}</pre>}
              {ch.resp&&<div style={{fontSize:10,color:"var(--mt)",marginTop:3}}>👤 Resp: {ch.resp}</div>}
            </div>
            <div style={{display:"flex",gap:5,marginLeft:9,flexShrink:0}}>
              <button className="btn secondary small" onClick={()=>setEditando({...ch})}>✏️</button>
              {ch.status!=="encerrado"&&<button className="btn ok small" onClick={()=>{setChamados(a=>a.map(x=>x.id===ch.id?{...x,status:"encerrado"}:x));showToast("✅ Encerrado","ok");}}>✅</button>}
            </div>
          </div>
        </div>);
      })}
      {filtered.length===0&&<div style={{textAlign:"center",color:"var(--mt)",padding:"40px"}}>Nenhum chamado encontrado.</div>}
    </div>

    {/* NOVO CHAMADO */}
    {novoModal&&<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&setNovoModal(false)}>
      <div className="modal" style={{width:520}}>
        <div className="modal-head"><h2>📨 Novo Chamado</h2><button className="icon-btn" onClick={()=>setNovoModal(false)}>×</button></div>
        <div className="stack">
          <div className="g3">
            <div><label>Setor</label><select value={novoForm.setor} onChange={e=>snf({setor:e.target.value})}>{SETORES_CHAMADO.map(s=><option key={s}>{s}</option>)}</select></div>
            <div><label>Tipo</label>
              <select value={novoForm.tipo} onChange={e=>snf({tipo:e.target.value,camposValues:{}})}>
                <option value="manual">✏️ Texto livre</option>
                {novoForm.setor==="Agendamento"&&<option value="padrao">📋 Padronizado</option>}
              </select>
            </div>
            <div><label>Prioridade</label>
              <select value={novoForm.prioridade} onChange={e=>snf({prioridade:e.target.value})}>
                {PRIORIDADE_CHAMADO.map(p=><option key={p} value={p}>{p}</option>)}
              </select>
            </div>
          </div>
          {isPadrao&&<div><label>Modelo (Anexo VII)</label>
            <select value={novoForm.padraoId} onChange={e=>snf({padraoId:e.target.value,camposValues:{}})}>
              {PADRAO_CHAMADO.map(p=><option key={p.id} value={p.id}>{p.label}</option>)}
            </select>
          </div>}
          {isPadrao&&<div className="section-box" style={{background:"#7c6af708",borderColor:"#7c6af725"}}>
            <div className="section-title" style={{color:"#a78bfa"}}>📋 {padraoSel.label}</div>
            <div style={{display:"flex",flexDirection:"column",gap:6}}>
              {padraoSel.campos.map(campo=>(
                <div key={campo} style={{display:"grid",gridTemplateColumns:"140px 1fr",alignItems:"center",gap:7}}>
                  <label style={{textTransform:"none",fontSize:11,fontWeight:700,color:"var(--mt)",margin:0}}>{campo}:</label>
                  <input value={novoForm.camposValues[campo]||""} onChange={e=>snf({camposValues:{...novoForm.camposValues,[campo]:e.target.value}})} style={{fontSize:11,padding:"5px 8px"}}/>
                </div>
              ))}
            </div>
          </div>}
          {!isPadrao&&<>
            <div><label>Assunto *</label><input value={novoForm.assunto} onChange={e=>snf({assunto:e.target.value})} autoFocus placeholder="Resumo do chamado..."/></div>
            <div><label>Descrição</label><textarea rows={3} value={novoForm.descricao} onChange={e=>snf({descricao:e.target.value})}/></div>
          </>}
          {novoForm.setor==="Agendamento"&&<div style={{padding:"7px 11px",borderRadius:8,background:"#38bdf808",border:"1px solid #38bdf820",fontSize:11,color:"#38bdf8"}}>
            ⏳ Chamados do setor Agendamento entram automaticamente na Fila de Espera.
          </div>}
          <div className="actions">
            <button className="btn secondary" onClick={()=>setNovoModal(false)}>Cancelar</button>
            <button className="btn primary" onClick={()=>{
              if(!isPadrao&&!novoForm.assunto.trim())return alert("Assunto obrigatório");
              abrirNovo();
            }}>📨 Enviar</button>
          </div>
        </div>
      </div>
    </div>}

    {/* EDITAR */}
    {editando&&<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&setEditando(null)}>
      <div className="modal" style={{width:520}}>
        <div className="modal-head"><h2>✏️ Editar #{editando.numero}</h2><button className="icon-btn" onClick={()=>setEditando(null)}>×</button></div>
        <div className="stack">
          <div className="g2">
            <div><label>Status</label>
              <select value={editando.status} onChange={e=>setEditando(x=>({...x,status:e.target.value}))}>
                {STATUS_CHAMADO.map(s=><option key={s} value={s}>{LABEL_CHAMADO[s]||s}</option>)}
              </select>
            </div>
            <div><label>Prioridade</label>
              <select value={editando.prioridade||"Normal"} onChange={e=>setEditando(x=>({...x,prioridade:e.target.value}))}>
                {PRIORIDADE_CHAMADO.map(p=><option key={p}>{p}</option>)}
              </select>
            </div>
          </div>
          <div className="g2">
            <div><label>Setor</label><select value={editando.setor} onChange={e=>setEditando(x=>({...x,setor:e.target.value}))}>{SETORES_CHAMADO.map(s=><option key={s}>{s}</option>)}</select></div>
            <div><label>Responsável</label><input value={editando.resp||""} onChange={e=>setEditando(x=>({...x,resp:e.target.value}))}/></div>
          </div>
          <div><label>Assunto</label><input value={editando.nome} onChange={e=>setEditando(x=>({...x,nome:e.target.value}))}/></div>
          <div><label>Descrição / Detalhes</label><textarea rows={5} value={editando.descricao} onChange={e=>setEditando(x=>({...x,descricao:e.target.value}))}/></div>
          <div className="actions">
            <button className="btn secondary" onClick={()=>setEditando(null)}>Cancelar</button>
            <button className="btn danger small" onClick={()=>{setChamados(a=>a.filter(x=>x.id!==editando.id));showToast("🗑️ Excluído","err");setEditando(null);}}>🗑️</button>
            <button className="btn primary" onClick={()=>{setChamados(a=>a.map(x=>x.id===editando.id?editando:x));showToast("✅ Atualizado","ok");setEditando(null);}}>Salvar</button>
          </div>
        </div>
      </div>
    </div>}
  </div>);
}

function FinanceiroPage({agenda,pacientes,profissionais,procedimentos,convenios,setAgenda,showToast,naoRecebiveis,setNaoRecebiveis}){
  const [tab,setTab]=useState("faturados");
  const [periodo,setPeriodo]=useState("mes");
  const [dataIni,setDataIni]=useState(()=>{const d=new Date();d.setDate(1);return ymd(d);});
  const [dataFim,setDataFim]=useState(hoje_str);
  const [filtroConv,setFiltroConv]=useState("");
  const [filtroProf,setFiltroProf]=useState("");

  const aplicarPeriodo=(p)=>{
    setPeriodo(p);const d=new Date();
    if(p==="hoje"){setDataIni(hoje_str);setDataFim(hoje_str);}
    else if(p==="semana"){const x=new Date(d);x.setDate(d.getDate()-6);setDataIni(ymd(x));setDataFim(hoje_str);}
    else if(p==="mes"){const x=new Date(d);x.setDate(1);setDataIni(ymd(x));setDataFim(hoje_str);}
    else if(p==="trimestre"){const x=new Date(d);x.setMonth(d.getMonth()-2);x.setDate(1);setDataIni(ymd(x));setDataFim(hoje_str);}
  };

  const Pill=({v,l})=>(<button onClick={()=>aplicarPeriodo(v)} style={{padding:"4px 12px",borderRadius:20,fontSize:11,fontWeight:800,border:"1.5px solid "+(periodo===v?"#7c6af7":"var(--cpb)"),background:periodo===v?"#7c6af715":"transparent",color:periodo===v?"#7c6af7":"var(--mt)",cursor:"pointer",whiteSpace:"nowrap"}}>{l}</button>);

  const faturados=agenda.filter(ag=>{
    const convOk=!filtroConv||(ag.convenio||"")===filtroConv;
    const profOk=!filtroProf||String(ag.profissionalId)===String(filtroProf);
    return (ag.fatStatus==="faturado"||ag.status==="faturado")&&ag.data>=dataIni&&ag.data<=dataFim&&convOk&&profOk;
  });
  // Pendências e Glosas: registros manuais, independentes das sessões
  const nrBase=(naoRecebiveis||[]).filter(nr=>{
    const convOk=!filtroConv||nr.convenio===filtroConv;
    const profOk=!filtroProf||String(nr.profissionalId)===String(filtroProf);
    return nr.data>=dataIni&&nr.data<=dataFim&&convOk&&profOk;
  });
  const pendencias=nrBase.filter(nr=>nr.tipo==="pendencia");
  const recursos=nrBase.filter(nr=>nr.tipo==="glosa");
  const perdas=[...pendencias,...recursos];
  const [modalNR,setModalNR]=useState(false);
  const [formNR,setFormNR]=useState({tipo:"pendencia",convenio:"",profissionalId:"",procedimentoId:"",data:hoje_str,valor:0,descricao:""});

  const porConv=faturados.reduce((acc,ag)=>{
    const conv=ag.convenio||"Particular";
    const proc=procedimentos.find(p=>p.id===Number(ag.procedimentoId));
    const val=proc?.valor||0;
    if(!acc[conv])acc[conv]={nome:conv,qtd:0,total:0};
    acc[conv].qtd++;acc[conv].total+=val;
    return acc;
  },{});
  const totalGeral=Object.values(porConv).reduce((s,cv)=>s+cv.total,0);
  const totalPerdas=perdas.reduce((s,ag)=>{const p=procedimentos.find(x=>x.id===Number(ag.procedimentoId));return s+(p?.valor||0);},0);

  const convs=[...new Set([...agenda.filter(a=>a.fatStatus==="faturado").map(a=>a.convenio||"Particular"),...(naoRecebiveis||[]).map(nr=>nr.convenio)].filter(Boolean))];
  const profsFat=[...new Set([...agenda.filter(a=>a.fatStatus==="faturado").map(a=>a.profissionalId),...(naoRecebiveis||[]).map(nr=>nr.profissionalId)].filter(Boolean))];

  const porConvPerda=perdas.reduce((acc,ag)=>{
    const conv=ag.convenio||"Particular";
    const proc=procedimentos.find(p=>p.id===Number(ag.procedimentoId));
    const val=proc?.valor||0;
    if(!acc[conv])acc[conv]={nome:conv,qtd:0,total:0};
    acc[conv].qtd++;acc[conv].total+=val;
    return acc;
  },{});

  const FiltroBar=()=>(<div style={{background:"var(--card)",border:"1px solid var(--cb)",borderRadius:13,padding:"14px 16px",marginBottom:16,display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
    <div style={{display:"flex",gap:6,flexWrap:"wrap"}}><Pill v="hoje" l="Hoje"/><Pill v="semana" l="7 dias"/><Pill v="mes" l="Mês"/><Pill v="trimestre" l="Trimestre"/></div>
    <div style={{width:1,height:24,background:"var(--sc)",flexShrink:0}}/>
    <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
      <div style={{display:"flex",alignItems:"center",gap:6,fontSize:12}}>
        <label style={{color:"var(--mt)",fontWeight:700,fontSize:11}}>De</label>
        <input type="date" value={dataIni} onChange={e=>{setDataIni(e.target.value);setPeriodo("custom");}} style={{width:140,fontSize:12}}/>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:6,fontSize:12}}>
        <label style={{color:"var(--mt)",fontWeight:700,fontSize:11}}>Até</label>
        <input type="date" value={dataFim} onChange={e=>{setDataFim(e.target.value);setPeriodo("custom");}} style={{width:140,fontSize:12}}/>
      </div>
    </div>
    <div style={{width:1,height:24,background:"var(--sc)",flexShrink:0}}/>
    <select value={filtroConv} onChange={e=>setFiltroConv(e.target.value)} style={{minWidth:140,fontSize:12}}>
      <option value="">Todos convênios</option>{convs.map(cv=><option key={cv}>{cv}</option>)}
    </select>
    <select value={filtroProf} onChange={e=>setFiltroProf(e.target.value)} style={{minWidth:150,fontSize:12}}>
      <option value="">Todos profissionais</option>{profsFat.map(id=>{const p=profissionais.find(x=>x.id===Number(id));return<option key={id} value={id}>{p?.nome||id}</option>;})}
    </select>
    <span style={{marginLeft:"auto",fontSize:11,color:"var(--mt)",fontStyle:"italic"}}>{brDate(dataIni)} → {brDate(dataFim)}</span>
  </div>);

  return(<div className="page-wrap">
    <div className="page-head"><h1>💼 Faturamento</h1></div>

    {/* KPI sumário topo */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:8,marginBottom:16}}>
      <div className="card" style={{padding:"12px 14px",borderTop:"3px solid #22c55e"}}>
        <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:22,color:"#34d399"}}>{brl(totalGeral)}</div>
        <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",marginTop:3}}>Faturado no período</div>
        <div style={{fontSize:10,color:"var(--mt)"}}>{faturados.length} sessões</div>
      </div>
      <div className="card" style={{padding:"12px 14px",borderTop:"3px solid #f97316"}}>
        <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:20,color:"#f97316"}}>{pendencias.length}</div>
        <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",marginTop:3}}>📭 Pendências</div>
        <div style={{fontSize:10,color:"var(--mt)"}}>não faturadas</div>
      </div>
      <div className="card" style={{padding:"12px 14px",borderTop:"3px solid #dc2626"}}>
        <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:20,color:"#dc2626"}}>{recursos.length}</div>
        <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",marginTop:3}}>🚫 Recursos</div>
        <div style={{fontSize:10,color:"var(--mt)"}}>glosas a recorrer</div>
      </div>
      {totalGeral+totalPerdas>0&&<div className="card" style={{padding:"12px 14px",borderTop:"3px solid #a78bfa"}}>
        <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:22,color:"#a78bfa"}}>{Math.round(totalPerdas/(totalGeral+totalPerdas)*100)}%</div>
        <div style={{fontSize:10,fontWeight:800,color:"var(--mt)",textTransform:"uppercase",marginTop:3}}>Taxa de perda</div>
        <div style={{fontSize:10,color:"var(--mt)"}}>sobre receita potencial</div>
      </div>}
    </div>

    {/* Tabs */}
    <div className="tab-bar" style={{marginBottom:14}}>
      {[["faturados","💵 Créditos","#22c55e"],["pendencias","📭 Pendências","#f97316"],["recursos","🚫 Recursos de Glosa","#dc2626"]].map(([k,l,cor])=>{
        const cnt=k==="faturados"?faturados.length:k==="pendencias"?pendencias.length:recursos.length;
        return(<button key={k} className="btn tab-btn" onClick={()=>setTab(k)} style={{background:tab===k?"var(--na)":"transparent",color:tab===k?cor:"var(--mt)",fontWeight:tab===k?800:500,border:tab===k?"1.5px solid "+cor+"40":"none"}}>
          {l}{cnt>0&&<span style={{marginLeft:5,background:cor,color:"#fff",borderRadius:20,padding:"0 5px",fontSize:10,fontWeight:800}}>{cnt}</span>}
        </button>);})}
    </div>

    <FiltroBar/>

    {/* ── TAB FATURADOS ── */}
    {tab==="faturados"&&<>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(190px,1fr))",gap:9,marginBottom:14}}>
        {Object.values(porConv).map(cv=>(
          <div key={cv.nome} className="card" style={{padding:13,borderLeft:"3px solid #22c55e"}}>
            <div style={{fontWeight:800,fontSize:13}}>{cv.nome}</div>
            <div style={{marginTop:4,fontSize:20,fontWeight:900,color:"#34d399"}}>{brl(cv.total)}</div>
            <div className="muted" style={{fontSize:11,marginTop:2}}>{cv.qtd} sessão(ões) faturada(s)</div>
          </div>
        ))}
        {Object.keys(porConv).length>0&&<div className="card" style={{padding:13,borderLeft:"3px solid #f59e0b"}}>
          <div style={{fontWeight:800,fontSize:12,color:"#f59e0b"}}>TOTAL GERAL</div>
          <div style={{fontSize:22,fontWeight:900,color:"#fbbf24",marginTop:4}}>{brl(totalGeral)}</div>
          <div className="muted" style={{fontSize:11,marginTop:2}}>{faturados.length} sessões</div>
        </div>}
      </div>
      <div className="card" style={{overflow:"hidden"}}>
        <div className="grid-header" style={{gridTemplateColumns:"90px 60px 1fr 130px 110px 130px 90px"}}>
          <div>Data</div><div>Hora</div><div>Paciente</div><div>Profissional</div><div>Convênio</div><div>Procedimento</div><div style={{textAlign:"right"}}>Valor</div>
        </div>
        {faturados.length===0&&<div style={{padding:16}} className="muted">Nenhum lançamento faturado no período.</div>}
        {faturados.sort((a,b)=>a.data.localeCompare(b.data)||a.horarioSessao.localeCompare(b.horarioSessao)).map((ag,i)=>{
          const pac=pacientes.find(p=>p.id===Number(ag.pacienteId));
          const prof=profissionais.find(p=>p.id===Number(ag.profissionalId));
          const proc=procedimentos.find(p=>p.id===Number(ag.procedimentoId));
          return(<div key={ag.id} className="grid-row" style={{gridTemplateColumns:"90px 60px 1fr 130px 110px 130px 90px",background:i%2?"var(--gr)":""}}>
            <div style={{fontWeight:700,fontSize:11}}>{brDate(ag.data)}</div>
            <div style={{fontWeight:800,color:"#a78bfa",fontSize:11}}>{ag.horarioSessao}</div>
            <div style={{fontWeight:700,fontSize:12}}>{pac?.nome||"—"}</div>
            <div style={{fontSize:11,color:espCor((prof?.especialidades||[""])[0])}}>{profShort(prof?.nome||"—")}</div>
            <div><span style={{padding:"1px 6px",borderRadius:20,fontSize:10,fontWeight:800,background:"#22c55e20",color:"#22c55e"}}>{ag.convenio||"Particular"}</span></div>
            <div className="muted" style={{fontSize:11}}>{proc?.nome||"—"}</div>
            <div style={{textAlign:"right",fontWeight:800,color:"#34d399",fontSize:12}}>{brl(proc?.valor||0)}</div>
          </div>);
        })}
        {faturados.length>0&&<div className="grid-row" style={{gridTemplateColumns:"90px 60px 1fr 130px 110px 130px 90px",background:"var(--sx)",fontWeight:900}}>
          <div style={{gridColumn:"1/7",fontSize:12}}>TOTAL ({faturados.length} sessões)</div>
          <div style={{textAlign:"right",color:"#34d399",fontSize:13}}>{brl(totalGeral)}</div>
        </div>}
      </div>
    </>}

    {/* ── TAB PERDAS ── */}
    {tab==="pendencias"&&<div>
      <div style={{padding:"10px 14px",background:"#431a0310",border:"1px solid #f9731630",borderRadius:10,marginBottom:14,fontSize:12,color:"#fdba74",lineHeight:1.6}}>
        📭 <b>Pendências de faturamento</b> — sessões realizadas cujo faturamento ainda não foi enviado ou confirmado pelo convênio. Regularize o envio para garantir o recebimento.
      </div>
      <div className="card" style={{overflow:"hidden"}}>
        <div className="grid-header" style={{gridTemplateColumns:"90px 60px 1fr 130px 110px 130px"}}><div>Data</div><div>Hora</div><div>Paciente</div><div>Profissional</div><div>Convênio</div><div>Valor</div></div>
        {pendencias.length===0&&<div style={{padding:24,textAlign:"center"}} className="muted">✅ Nenhuma pendência no período.</div>}
        {pendencias.map((ag,i)=>{
          const pac=pacientes.find(p=>p.id===Number(ag.pacienteId));
          const prof=profissionais.find(p=>p.id===Number(ag.profissionalId));
          const proc=procedimentos.find(p=>p.id===Number(ag.procedimentoId));
          const val=proc?.valor||0;
          return(<div key={ag.id} className="grid-row" style={{gridTemplateColumns:"90px 60px 1fr 130px 110px 130px",background:i%2?"var(--gr)":""}}>
            <div style={{fontSize:11,fontWeight:700}}>{brDate(ag.data)}</div>
            <div style={{fontSize:11,color:"#a78bfa",fontWeight:800}}>{ag.horarioSessao}</div>
            <div style={{fontSize:11,fontWeight:700}}>{pac?.nome||"—"}</div>
            <div style={{fontSize:11,color:"var(--mt)"}}>{prof?.nome||"—"}</div>
            <div style={{fontSize:11}}>{ag.convenio||"—"}</div>
            <div style={{fontSize:11,fontWeight:800,color:"#f97316"}}>{brl(val)}</div>
          </div>);
        })}
        {pendencias.length>0&&<div className="grid-row" style={{gridTemplateColumns:"90px 60px 1fr 130px 110px 130px",background:"#f9731615",fontWeight:900}}>
          <div style={{gridColumn:"1/6",fontSize:12,color:"#f97316"}}>TOTAL PENDÊNCIAS ({pendencias.length} sessões)</div>
          <div style={{fontSize:12,color:"#f97316"}}>{brl(pendencias.reduce((s,ag)=>{const p=procedimentos.find(x=>x.id===Number(ag.procedimentoId));return s+(p?.valor||0);},0))}</div>
        </div>}
      </div>
    </div>}

    {tab==="recursos"&&<div>
      <div style={{padding:"10px 14px",background:"#450a0a10",border:"1px solid #dc262630",borderRadius:10,marginBottom:14,fontSize:12,color:"#fca5a5",lineHeight:1.6}}>
        🚫 <b>Recursos de glosa</b> — sessões cujo faturamento foi recusado/glosado pelo convênio. Analise o motivo e prepare o recurso dentro do prazo contratual.
      </div>
      <div className="card" style={{overflow:"hidden"}}>
        <div className="grid-header" style={{gridTemplateColumns:"90px 60px 1fr 130px 110px 130px"}}><div>Data</div><div>Hora</div><div>Paciente</div><div>Profissional</div><div>Convênio</div><div>Valor</div></div>
        {recursos.length===0&&<div style={{padding:24,textAlign:"center"}} className="muted">✅ Nenhuma glosa registrada no período.</div>}
        {recursos.map((ag,i)=>{
          const pac=pacientes.find(p=>p.id===Number(ag.pacienteId));
          const prof=profissionais.find(p=>p.id===Number(ag.profissionalId));
          const proc=procedimentos.find(p=>p.id===Number(ag.procedimentoId));
          const val=proc?.valor||0;
          return(<div key={ag.id} className="grid-row" style={{gridTemplateColumns:"90px 60px 1fr 130px 110px 130px",background:i%2?"var(--gr)":""}}>
            <div style={{fontSize:11,fontWeight:700}}>{brDate(ag.data)}</div>
            <div style={{fontSize:11,color:"#a78bfa",fontWeight:800}}>{ag.horarioSessao}</div>
            <div style={{fontSize:11,fontWeight:700}}>{pac?.nome||"—"}</div>
            <div style={{fontSize:11,color:"var(--mt)"}}>{prof?.nome||"—"}</div>
            <div style={{fontSize:11}}>{ag.convenio||"—"}</div>
            <div style={{fontSize:11,fontWeight:800,color:"#dc2626"}}>{brl(val)}</div>
          </div>);
        })}
        {recursos.length>0&&<div className="grid-row" style={{gridTemplateColumns:"90px 60px 1fr 130px 110px 130px",background:"#dc262615",fontWeight:900}}>
          <div style={{gridColumn:"1/6",fontSize:12,color:"#dc2626"}}>TOTAL GLOSAS ({recursos.length} sessões)</div>
          <div style={{fontSize:12,color:"#dc2626"}}>{brl(recursos.reduce((s,ag)=>{const p=procedimentos.find(x=>x.id===Number(ag.procedimentoId));return s+(p?.valor||0);},0))}</div>
        </div>}
      </div>
    </div>}
  </div>);
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN AGENDA — com filtro de filial
// ═══════════════════════════════════════════════════════════════════════════════
function MainAgenda({auth,agenda,setAgenda,pacientes,profissionais,salas,filiais,alertas,setAlertas,onAbrirNovo,onAbrirSessao}){
  const [year,setYear]=useState(hoje.getFullYear());const[month,setMonth]=useState(hoje.getMonth());const[selDate,setSelDate]=useState(hoje_str);
  const [fSala,setFSala]=useState("");const[fEsp,setFEsp]=useState("");const[fPac,setFPac]=useState("");const[fFilial,setFFilial]=useState("");const[fProf,setFProf]=useState("");
  const [view,setView]=useState("mensal");
  const [showAlertas,setShowAlertas]=useState(false);
  const pend=alertas.filter(a=>!a.lido);
  const profShortName=p=>{if(!p?.nome)return"—";const pts=p.nome.trim().split(" ").filter(Boolean);return pts.length===1?pts[0]:pts[0]+" "+pts[pts.length-1];};
  const filtered=useMemo(()=>agenda.filter(a=>{
    const prof=profissionais.find(p=>p.id===Number(a.profissionalId));const pac=pacientes.find(p=>p.id===Number(a.pacienteId));
    return(!fSala||String(a.salaId)===String(fSala))&&(!fEsp||(prof?.especialidades||[prof?.especialidade]).includes(fEsp))&&(!fPac||String(a.pacienteId)===String(fPac))&&(!fFilial||String(a.filialId)===String(fFilial)||salas.find(s=>s.id===a.salaId&&s.filialId===Number(fFilial)))&&(!fProf||String(a.profissionalId)===String(fProf));
  }),[agenda,profissionais,pacientes,fSala,fEsp,fPac,fFilial,fProf,salas]);
  const dIM=getDIM(year,month);const fD=getFD(year,month);
  const cells=[...Array(fD).fill(null),...Array.from({length:dIM},(_,i)=>i+1)];
  const evOn=d=>filtered.filter(a=>a.data===d);
  const weekDays=useMemo(()=>{const d=new Date(selDate+"T12:00:00");d.setDate(d.getDate()-d.getDay());return Array.from({length:7},(_,i)=>{const x=new Date(d);x.setDate(d.getDate()+i);return ymd(x);});},[selDate]);
  const CardSessao=({ag})=>{
    const prof=profissionais.find(p=>p.id===Number(ag.profissionalId));
    const pac=pacientes.find(p=>p.id===Number(ag.pacienteId));
    const stColor=STATUS_AG[ag.status]?.color||espCor((prof?.especialidades||[prof?.especialidade])[0]);
    if(ag.tipo==="reuniao"){
      const nPart=(ag.participantesIds||[]).length;
      return(<div onClick={e=>{e.stopPropagation();onAbrirSessao&&onAbrirSessao(ag);}}
        style={{background:"#7c6af718",borderLeft:"3px solid #7c6af7",padding:"2px 5px",borderRadius:3,cursor:"pointer",lineHeight:1.3,marginBottom:1}}
        title={"🤝 "+ag.tituloReuniao+" · "+ag.horarioSessao+(nPart?" · "+nPart+" participante(s)":"")}>
        <div style={{fontSize:10,fontWeight:900,color:"#a78bfa",display:"flex",alignItems:"center",gap:3}}>
          <span>🤝</span><span>{ag.horarioSessao}</span>
        </div>
        <div style={{fontSize:9,color:"#93c5fd",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",fontWeight:700}}>{ag.tituloReuniao||"Reunião"}</div>
        {nPart>0&&<div style={{fontSize:8,color:"#64748b"}}>{nPart} participante(s)</div>}
      </div>);
    }
    return(<div onClick={e=>{e.stopPropagation();onAbrirSessao&&onAbrirSessao(ag);}} style={{background:stColor+"20",borderLeft:"3px solid "+stColor,padding:"2px 4px",borderRadius:3,cursor:"pointer",lineHeight:1.3,marginBottom:1}} title={profShortName(prof)+" · "+(pac?.nome||"")+" · "+(STATUS_AG[ag.status]?.label||ag.status)}>
      <div style={{fontSize:10,fontWeight:900,color:stColor}}>{ag.horarioSessao} <span style={{fontWeight:500,color:"var(--tx)",fontSize:9}}>{pac?.nome?.split(" ")[0]||"—"}</span></div>
      <div style={{fontSize:9,color:"var(--mt)",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{profShortName(prof)}</div>
    </div>);
  };
  return(<div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",position:"relative"}}>
    {showAlertas&&<div style={{position:"absolute",right:14,top:6,width:290,zIndex:50}}>
      <div className="card" style={{padding:11,border:"1px solid var(--sc)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:7}}><b style={{color:"#93c5fd",fontSize:12}}>🔔 Alertas</b><button className="btn small secondary" onClick={()=>{setAlertas(a=>a.map(x=>({...x,lido:true})));setShowAlertas(false);}}>Limpar</button></div>
        {pend.map(a=><div key={a.id} style={{padding:"6px 9px",borderRadius:6,background:"var(--sx)",border:"1px solid var(--sc)",marginBottom:4}}><div style={{fontWeight:800,fontSize:12}}>{a.titulo}</div><div style={{fontSize:11,color:"var(--mt)",marginTop:1}}>{a.descricao}</div></div>)}
        {pend.length===0&&<div className="muted" style={{fontSize:12}}>Sem alertas.</div>}
      </div>
    </div>}
    <div className="filter-bar" style={{flexWrap:"wrap",gap:6}}>
      <div><label>Filial</label><select value={fFilial} onChange={e=>setFFilial(e.target.value)} style={{minWidth:90}}><option value="">Todas</option>{filiais.map(f=><option key={f.id} value={f.id}>{f.nome}</option>)}</select></div>
      <div><label>Profissional</label><select value={fProf} onChange={e=>setFProf(e.target.value)} style={{minWidth:140}}><option value="">Todos</option>{profissionais.filter(p=>p.role==="profissional"||p.role==="coordenador"||p.role==="coordenador_aba").map(p=><option key={p.id} value={p.id}>{profShortName(p)}</option>)}</select></div>
      <div><label>Paciente</label><select value={fPac} onChange={e=>setFPac(e.target.value)} style={{minWidth:140}}><option value="">Todos</option>{pacientes.map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}</select></div>
      <div><label>Especialidade</label><select value={fEsp} onChange={e=>setFEsp(e.target.value)} style={{minWidth:120}}><option value="">Todas</option>{[...new Set(profissionais.flatMap(p=>p.especialidades||[p.especialidade]))].map(e=><option key={e} value={e}>{e}</option>)}</select></div>
      <div><label>Sala</label><select value={fSala} onChange={e=>setFSala(e.target.value)} style={{minWidth:80}}><option value="">Todas</option>{salas.filter(s=>!fFilial||s.filialId===Number(fFilial)).map(s=><option key={s.id} value={s.id}>{s.nome}</option>)}</select></div>
      <div style={{marginLeft:"auto",display:"flex",gap:5,alignSelf:"flex-end",flexWrap:"wrap"}}>
        <div style={{display:"flex",borderRadius:6,overflow:"hidden",border:"1px solid var(--sc)"}}>
          {[["mensal","📅 Mês"],["semanal","📊 Sem."],["diaria","📋 Dia"],["recepcao","🏥 Recepção"],["lista","☰ Lista"]].map(([v,l])=>(<button key={v} onClick={()=>setView(v)} style={{padding:"4px 8px",fontSize:10,fontWeight:800,background:view===v?"#7c6af7":"transparent",color:view===v?"#fff":"var(--mt)",border:"none",cursor:"pointer"}}>{l}</button>))}
        </div>
        <button className="btn secondary" style={{position:"relative"}} onClick={()=>setShowAlertas(v=>!v)}>
          🔔{pend.length>0&&<span style={{position:"absolute",top:-4,right:-4,background:"#ef4444",color:"#fff",borderRadius:"50%",fontSize:9,width:15,height:15,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900}}>{pend.length}</span>}
        </button>
        <button className="btn primary" onClick={()=>onAbrirNovo(selDate)}>+ Agendar</button>
      </div>
    </div>
    {view==="mensal"&&<div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <div className="calendar-head">
        <button className="btn secondary small" onClick={()=>month===0?(setYear(y=>y-1),setMonth(11)):setMonth(m=>m-1)}>‹</button>
        <div className="calendar-title">{MONTHS_PT[month]} {year}</div>
        <button className="btn secondary small" onClick={()=>month===11?(setYear(y=>y+1),setMonth(0)):setMonth(m=>m+1)}>›</button>
        <button className="btn secondary small" onClick={()=>{setYear(hoje.getFullYear());setMonth(hoje.getMonth());setSelDate(hoje_str);}}>Hoje</button>
        <span className="muted" style={{marginLeft:"auto",fontSize:11}}>{filtered.length} agend.</span>
      </div>
      <div style={{flex:1,overflow:"auto",padding:7}}>
        <div className="dow-grid">{["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"].map(d=><div key={d} className="dow">{d}</div>)}</div>
        <div className="month-grid">
          {cells.map((day,i)=>{if(!day)return<div key={"e"+i}/>;const ds=year+"-"+String(month+1).padStart(2,"0")+"-"+String(day).padStart(2,"0");const isToday=ds===hoje_str,isSel=ds===selDate,evs=evOn(ds);
            return(<div key={ds} className={"day-cell"+(isToday?" today":"")+(isSel?" selected":"")} onClick={()=>setSelDate(ds)} onDoubleClick={()=>onAbrirNovo(ds)}>
              <div className="day-top"><span className="day-num">{day}</span>{evs.length>0&&<span style={{fontSize:9,color:"var(--mt)"}}>{evs.length}</span>}</div>
              <div>{evs.slice(0,3).map(ag=><CardSessao key={ag.id} ag={ag}/>)}{evs.length>3&&<div style={{fontSize:9,color:"var(--mt)",textAlign:"center",fontWeight:700}}>+{evs.length-3} mais</div>}</div>
            </div>);
          })}
        </div>
      </div>
    </div>}
    {view==="semanal"&&<div style={{flex:1,overflow:"auto",padding:8}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
        <button className="btn secondary small" onClick={()=>{const d=new Date(selDate+"T12:00:00");d.setDate(d.getDate()-7);setSelDate(ymd(d));}}>‹ Ant.</button>
        <b style={{fontSize:12}}>{brDate(weekDays[0]).slice(0,5)} – {brDate(weekDays[6])}</b>
        <button className="btn secondary small" onClick={()=>{const d=new Date(selDate+"T12:00:00");d.setDate(d.getDate()+7);setSelDate(ymd(d));}}>Próx. ›</button>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:6}}>
        {weekDays.map(ds=>{const evs=evOn(ds);const isT=ds===hoje_str;return(<div key={ds} style={{background:"var(--card)",border:"1px solid var(--sc)",borderRadius:8,padding:7,borderTop:"3px solid "+(isT?"#7c6af7":"var(--sc)")}}>
          <div style={{fontWeight:800,fontSize:10,color:isT?"#7c6af7":"var(--mt)",marginBottom:5}}>{["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"][new Date(ds+"T12:00:00").getDay()]} {brDate(ds).slice(0,5)}</div>
          {evs.length===0&&<div style={{fontSize:9,color:"var(--lb)",textAlign:"center",padding:"8px 0"}}>—</div>}
          {evs.map(ag=><CardSessao key={ag.id} ag={ag}/>)}
        </div>);})}
      </div>
    </div>}
    {view==="diaria"&&<div style={{flex:1,overflow:"auto",padding:8}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
        <button className="btn secondary small" onClick={()=>{const d=new Date(selDate+"T12:00:00");d.setDate(d.getDate()-1);setSelDate(ymd(d));}}>‹</button>
        <div style={{display:"flex",gap:8,alignItems:"center"}}><b style={{fontSize:13}}>{brDate(selDate)}</b><input type="date" value={selDate} onChange={e=>setSelDate(e.target.value)} style={{fontSize:11,width:130}}/></div>
        <button className="btn secondary small" onClick={()=>{const d=new Date(selDate+"T12:00:00");d.setDate(d.getDate()+1);setSelDate(ymd(d));}}>›</button>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:5}}>
        {TURNOS_H.map(hora=>{const evs=evOn(selDate).filter(a=>a.horarioSessao===hora);return(<div key={hora} style={{display:"flex",gap:8,alignItems:"flex-start",borderBottom:"1px solid var(--sc)",paddingBottom:4}}>
          <div style={{width:45,fontWeight:800,fontSize:11,color:"#a78bfa",flexShrink:0,paddingTop:2}}>{hora}</div>
          <div style={{flex:1,display:"flex",flexWrap:"wrap",gap:5}}>
            {evs.length===0?<span style={{fontSize:10,color:"var(--lb)"}}>—</span>:evs.map(ag=>{
              const prof=profissionais.find(p=>p.id===Number(ag.profissionalId));const pac=pacientes.find(p=>p.id===Number(ag.pacienteId));
              const stColor=STATUS_AG[ag.status]?.color||espCor((prof?.especialidades||[prof?.especialidade])[0]);
              return(<div key={ag.id} onClick={()=>onAbrirSessao&&onAbrirSessao(ag)} style={{background:stColor+"18",border:"1px solid "+stColor+"60",borderLeft:"3px solid "+stColor,borderRadius:6,padding:"4px 8px",cursor:"pointer",minWidth:160}}>
                <div style={{fontWeight:800,fontSize:11,color:stColor}}>{STATUS_AG[ag.status]?.icon} {ag.horarioSessao}–{ag.horarioFimSessao}</div>
                <div style={{fontWeight:700,fontSize:11,color:"var(--tx)"}}>{pac?.nome||"—"}</div>
                <div style={{fontSize:10,color:"var(--mt)"}}>{profShortName(prof)}</div>
              </div>);
            })}
          </div>
        </div>);})}
      </div>
    </div>}
    {view==="recepcao"&&(()=>{
      const evsDia=evOn(selDate).filter(a=>a.tipo!=="reuniao");
      const reunioesDia=evOn(selDate).filter(a=>a.tipo==="reuniao");
      const porStatus=Object.keys(STATUS_AG).reduce((acc,k)=>({...acc,[k]:evsDia.filter(a=>a.status===k).length}),{});
      const agendPend=evsDia.filter(a=>["agendado","confirmado"].includes(a.status)).length;
      const atend=evsDia.filter(a=>["atendido","faturado"].includes(a.status)).length;
      const faltas=evsDia.filter(a=>["faltou","faltou_pacote"].includes(a.status)).length;
      const emCurso=evsDia.filter(a=>{
        const agora=toTime(new Date().getHours()*60+new Date().getMinutes());
        return a.status==="agendado"&&toMin(a.horarioSessao)<=toMin(agora)&&toMin(a.horarioFimSessao||a.horarioSessao)>toMin(agora);
      }).length;
      return(<div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
        {/* Nav + KPIs */}
        <div style={{padding:"6px 10px",borderBottom:"1px solid var(--sc)",flexShrink:0}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:7}}>
            <button className="btn secondary small" onClick={()=>{const d=new Date(selDate+"T12:00:00");d.setDate(d.getDate()-1);setSelDate(ymd(d));}}>‹</button>
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              <b style={{fontSize:13}}>{brDate(selDate)}</b>
              <input type="date" value={selDate} onChange={e=>setSelDate(e.target.value)} style={{fontSize:11,width:130}}/>
            </div>
            <button className="btn secondary small" onClick={()=>{const d=new Date(selDate+"T12:00:00");d.setDate(d.getDate()+1);setSelDate(ymd(d));}}>›</button>
          </div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {[[evsDia.length,"TOTAL","#a78bfa"],[agendPend,"AGUARDANDO","#f59e0b"],[emCurso,"EM ATEND.","#a78bfa"],[atend,"ATENDIDOS","#34d399"],[faltas,"FALTAS","#f87171"]].map(([v,t,cor])=>(
              <div key={t} style={{padding:"4px 10px",borderRadius:8,background:cor+"15",border:"1px solid "+cor+"40",textAlign:"center",minWidth:70}}>
                <div style={{fontSize:16,fontWeight:900,color:cor,lineHeight:1}}>{v}</div>
                <div style={{fontSize:8,fontWeight:900,color:"var(--mt)",letterSpacing:.5,marginTop:1}}>{t}</div>
              </div>
            ))}
            {reunioesDia.length>0&&<div style={{padding:"4px 10px",borderRadius:8,background:"#7c6af715",border:"1px solid #7c6af740",textAlign:"center",minWidth:70}}>
              <div style={{fontSize:16,fontWeight:900,color:"#a78bfa",lineHeight:1}}>{reunioesDia.length}</div>
              <div style={{fontSize:8,fontWeight:900,color:"var(--mt)",letterSpacing:.5,marginTop:1}}>REUNIÕES</div>
            </div>}
          </div>
        </div>

        {/* Grade por hora */}
        <div style={{flex:1,overflowY:"auto"}}>
          {TURNOS_H.map(hora=>{
            const sessoes=evsDia.filter(a=>a.horarioSessao===hora);
            const reuns=reunioesDia.filter(a=>a.horarioSessao===hora);
            const total=sessoes.length+reuns.length;
            if(total===0)return null;
            return(<div key={hora} style={{display:"flex",gap:0,borderBottom:"1px solid var(--db)"}}>
              {/* Coluna hora */}
              <div style={{width:52,flexShrink:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"flex-start",paddingTop:8,background:"var(--sx)",borderRight:"2px solid var(--sc)"}}>
                <div style={{fontWeight:900,fontSize:13,color:"#a78bfa"}}>{hora}</div>
                <div style={{fontSize:9,color:"var(--mt)",marginTop:2,fontWeight:700}}>{total} sess.</div>
              </div>
              {/* Cards */}
              <div style={{flex:1,display:"flex",flexWrap:"wrap",gap:5,padding:"6px 8px"}}>
                {reuns.map(ag=>(
                  <div key={ag.id} onClick={()=>onAbrirSessao&&onAbrirSessao(ag)}
                    style={{display:"flex",alignItems:"stretch",borderRadius:7,overflow:"hidden",border:"1px solid #7c6af740",cursor:"pointer",minWidth:160,maxWidth:220,flex:"1 1 160px",background:"var(--card)"}}>
                    <div style={{width:5,background:"#7c6af7",flexShrink:0}}/>
                    <div style={{padding:"5px 8px",flex:1,minWidth:0}}>
                      <div style={{fontSize:10,fontWeight:900,color:"#a78bfa"}}>🤝 Reunião</div>
                      <div style={{fontSize:11,fontWeight:700,color:"var(--tx)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{ag.tituloReuniao||"Reunião"}</div>
                      <div style={{fontSize:9,color:"var(--mt)",marginTop:1}}>{(ag.participantesIds||[]).length} participante(s)</div>
                    </div>
                  </div>
                ))}
                {sessoes.map(ag=>{
                  const prof=profissionais.find(p=>p.id===Number(ag.profissionalId));
                  const pac=pacientes.find(p=>p.id===Number(ag.pacienteId));
                  const esp=(prof?.especialidades||[prof?.especialidade])[0]||"";
                  const espColor=espCor(esp)||"#64748b";
                  const st=STATUS_AG[ag.status]||{label:ag.status,color:"#64748b",icon:"📋"};
                  const stColor=st.color;
                  const sala=salas.find(s=>s.id===Number(ag.salaId));
                  const nomePac=pac?.nome||"—";
                  const nomeProf=profShortName(prof);
                  return(<div key={ag.id} onClick={()=>onAbrirSessao&&onAbrirSessao(ag)}
                    style={{display:"flex",alignItems:"stretch",borderRadius:7,overflow:"hidden",border:"1px solid "+espColor+"40",cursor:"pointer",minWidth:170,maxWidth:240,flex:"1 1 170px",background:"var(--card)",boxShadow:"0 1px 4px #0002"}}>
                    {/* Faixa lateral — cor da terapia */}
                    <div style={{width:5,background:espColor,flexShrink:0}}/>
                    <div style={{padding:"5px 8px",flex:1,minWidth:0,display:"flex",flexDirection:"column",gap:2}}>
                      {/* Status badge + hora fim */}
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                        <span style={{fontSize:9,fontWeight:900,background:stColor+"22",color:stColor,padding:"1px 5px",borderRadius:20,border:"1px solid "+stColor+"40"}}>{st.icon} {st.label}</span>
                        <span style={{fontSize:9,color:"var(--mt)",fontWeight:600}}>{ag.horarioFimSessao||""}</span>
                      </div>
                      {/* Paciente — nome destaque */}
                      <div style={{fontSize:12,fontWeight:900,color:"var(--tx)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",lineHeight:1.2}}>{nomePac}</div>
                      {/* Profissional */}
                      <div style={{fontSize:10,fontWeight:700,color:espColor,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{nomeProf}</div>
                      {/* Rodapé: especialidade + sala */}
                      <div style={{display:"flex",gap:5,marginTop:1,flexWrap:"wrap"}}>
                        {esp&&<span style={{fontSize:8,color:espColor,fontWeight:800,background:espColor+"15",padding:"1px 5px",borderRadius:20}}>{esp.split(" ")[0]}</span>}
                        {sala&&<span style={{fontSize:8,color:"var(--mt)",fontWeight:600}}>{sala.nome}</span>}
                        {ag.convenio&&<span style={{fontSize:8,color:"#22c55e",fontWeight:700}}>{ag.convenio.split(" ")[0]}</span>}
                      </div>
                    </div>
                  </div>);
                })}
              </div>
            </div>);
          })}
          {evsDia.length===0&&<div style={{textAlign:"center",padding:"48px 0",color:"var(--mt)",fontSize:13}}>Nenhuma sessão agendada para este dia.</div>}
        </div>
      </div>);
    })()}

    {view==="lista"&&<div style={{flex:1,overflow:"auto",padding:8}}>
      <div style={{display:"flex",gap:8,marginBottom:8,alignItems:"center"}}>
        <input type="date" value={selDate} onChange={e=>setSelDate(e.target.value)} style={{fontSize:11,width:140}}/>
        <span className="muted" style={{fontSize:11}}>{filtered.filter(a=>a.data>=selDate).length} sessão(ões) a partir desta data</span>
      </div>
      <div className="card" style={{overflow:"hidden"}}>
        <div className="grid-header" style={{gridTemplateColumns:"90px 70px 1fr 1fr 90px 100px"}}><div>Data</div><div>Hora</div><div>Paciente</div><div>Profissional</div><div>Status</div><div>Convênio</div></div>
        {filtered.filter(a=>a.data>=selDate).sort((a,b)=>a.data.localeCompare(b.data)||a.horarioSessao.localeCompare(b.horarioSessao)).slice(0,100).map((ag,i)=>{
          const prof=profissionais.find(p=>p.id===Number(ag.profissionalId));const pac=pacientes.find(p=>p.id===Number(ag.pacienteId));const st=STATUS_AG[ag.status]||{label:ag.status,color:"#64748b",icon:"📋"};
          return(<div key={ag.id} className="grid-row" style={{gridTemplateColumns:"90px 70px 1fr 1fr 90px 100px",background:i%2?"var(--gr)":"",cursor:"pointer"}} onClick={()=>onAbrirSessao&&onAbrirSessao(ag)}>
            <div style={{fontWeight:700,fontSize:11}}>{brDate(ag.data)}</div><div style={{fontWeight:800,color:"#a78bfa",fontSize:11}}>{ag.horarioSessao}</div>
            <div style={{fontWeight:700,fontSize:11}}>{pac?.nome||"—"}</div><div style={{fontSize:11,color:"var(--mt)"}}>{profShortName(prof)}</div>
            <div><span style={{padding:"1px 5px",borderRadius:20,fontSize:9,fontWeight:800,background:st.color+"25",color:st.color}}>{st.icon} {st.label}</span></div>
            <div style={{fontSize:11,color:"var(--mt)"}}>{ag.convenio||"—"}</div>
          </div>);
        })}
        {filtered.filter(a=>a.data>=selDate).length===0&&<div style={{padding:14}} className="muted">Nenhum agendamento.</div>}
      </div>
    </div>}
  </div>);
}

// ═══════════════════════════════════════════════════════════════════════════════
// GESTÃO — Processos & Indicadores
// ═══════════════════════════════════════════════════════════════════════════════
function WaitingListPage({fila,setFila,pacientes,auth,showToast,agenda,profissionais}){
  const podeEditar=auth?.role==="agendamento"||auth?.role==="administrador"||auth?.role==="coordenador"||auth?.role==="coordenador_aba";
  const blank={pacienteId:"",especialidade:"",preferencia:"Manhã",observacao:"",convenio:"",profissionalDesejado:""};
  const [modal,setModal]=useState(false);
  const [editing,setEditing]=useState(null);
  const [form,setForm]=useState(blank);
  const sf=p=>setForm(f=>({...f,...p}));
  const [iaSugestoes,setIaSugestoes]=useState([]);
  const [iaLoading,setIaLoading]=useState(false);
  const sugerirEncaixes=()=>{
    if(!fila.length)return showToast("⚠️ Fila vazia","warn");
    setIaLoading(true);
    setTimeout(()=>{
      const sugs=[];
      fila.forEach(fp=>{
        const pac=pacientes.find(p=>p.id===fp.pacienteId);
        if(!pac)return;
        (profissionais||[]).filter(p=>(p.especialidades||[]).includes(fp.especialidade)).forEach(prof=>{
          for(let i=1;i<=7;i++){
            const d=new Date();d.setDate(d.getDate()+i);
            if(d.getDay()===0||d.getDay()===6)continue;
            const ds=d.toISOString().slice(0,10);
            const diaNome=DIAS_SEMANA[d.getDay()];
            if(!prof.escala?.[diaNome]?.ativo)continue;
            const agProf=(agenda||[]).filter(a=>Number(a.profissionalId)===Number(prof.id)&&a.data===ds);
            const ocup=new Set(agProf.map(a=>a.horarioSessao));
            const inicio=prof.escala[diaNome].inicio||"08:00";
            const fim=prof.escala[diaNome].fim||"18:00";
            const slots=TURNOS_H.filter(h=>h>=inicio&&h<fim&&!ocup.has(h));
            if(slots.length>0&&sugs.length<8)
              sugs.push({pac:pac.nome,prof:prof.nome,esp:fp.especialidade,data:brDate(ds),ds,horario:slots[0],preferencia:fp.preferencia,convenio:fp.convenio||""});
          }
        });
      });
      setIaSugestoes(sugs.slice(0,8));
      setIaLoading(false);
      if(!sugs.length)showToast("Nenhum horário disponível encontrado","warn");
    },700);
  };
  const save=()=>{
    if(!form.pacienteId)return alert("Selecione o paciente");
    if(!form.especialidade)return alert("Especialidade obrigatória");
    if(editing) setFila(a=>a.map(x=>x.id===editing.id?{...form,id:editing.id,pacienteId:Number(form.pacienteId)}:x));
    else setFila(a=>[...a,{...form,id:Date.now(),pacienteId:Number(form.pacienteId)}]);
    showToast("✅ Fila atualizada","ok");
    setModal(false);
  };
  const abrir=(item)=>{
    setEditing(item||null);
    setForm(item?{...item,pacienteId:String(item.pacienteId)}:blank);
    setModal(true);
  };
  const remover=(id)=>{if(confirm("Remover da fila?"))setFila(a=>a.filter(x=>x.id!==id));};
  return(<div className="page-wrap">
    <div className="page-head">
      <h1>⏳ Fila de Espera</h1>
      <div style={{display:"flex",gap:8,alignItems:"center"}}>
        <span className="muted">{fila.length} registro(s)</span>
        {podeEditar&&<button className="btn primary" onClick={()=>abrir(null)}>+ Adicionar</button>}
      </div>
    </div>
    {fila.length===0
      ?<div style={{textAlign:"center",color:"var(--mt)",padding:"40px"}}>Fila vazia.</div>
      :<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:10}}>
        {fila.map(item=>{
          const pac=pacientes.find(p=>p.id===item.pacienteId);
          const cor=espCor(item.especialidade);
          return(<div key={item.id} className="card" style={{padding:14,borderLeft:"4px solid "+cor}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div style={{flex:1}}>
                <div style={{fontWeight:800,fontSize:13,color:"var(--tx)"}}>{pac?.nome||"—"}</div>
                <div style={{marginTop:4,display:"flex",flexWrap:"wrap",gap:4}}>
                  <span style={{padding:"2px 7px",borderRadius:20,fontSize:10,fontWeight:800,background:cor+"20",color:cor}}>{item.especialidade}</span>
                  <span style={{padding:"2px 7px",borderRadius:20,fontSize:10,fontWeight:700,background:"var(--na)",color:"var(--mt)"}}>{item.preferencia}</span>
                  {item.convenio&&<span style={{padding:"2px 7px",borderRadius:20,fontSize:10,fontWeight:700,background:"#22c55e20",color:"#22c55e"}}>{item.convenio}</span>}
                </div>
                {item.profissionalDesejado&&<div style={{fontSize:11,color:"var(--mt)",marginTop:4}}>Prof. desejado: {item.profissionalDesejado}</div>}
                {item.observacao&&<div style={{fontSize:11,color:"var(--mt)",marginTop:3,fontStyle:"italic"}}>{item.observacao}</div>}
              </div>
              {podeEditar&&<div style={{display:"flex",gap:4,marginLeft:6,flexShrink:0}}>
                <button className="btn secondary small" onClick={()=>abrir(item)}>✏️</button>
                <button className="btn danger small" onClick={()=>remover(item.id)}>🗑️</button>
              </div>}
            </div>
          </div>);
        })}
      </div>}

    {/* ── IA Encaixes ── */}
    <div className="card" style={{marginTop:16,padding:20,border:"1px solid #a78bfa30",borderRadius:14}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8}}>
        <div>
          <div style={{fontWeight:900,fontSize:14,color:"#a78bfa"}}>🧠 IA — Sugestões de Encaixe</div>
          <div style={{fontSize:11,color:"var(--mt)",marginTop:2}}>Analisa a fila e a agenda para encontrar horários livres disponíveis.</div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          {iaSugestoes.length>0&&<span style={{fontSize:11,color:"#34d399",fontWeight:800}}>✅ {iaSugestoes.length} sugestão(ões)</span>}
          <button className="btn primary" onClick={sugerirEncaixes} disabled={iaLoading}
            style={{background:"#7c3aed",borderColor:"#7c3aed",fontSize:12}}>
            {iaLoading?"⏳ Analisando…":"🧠 Gerar Sugestões"}
          </button>
          {iaSugestoes.length>0&&<button className="btn secondary" style={{fontSize:11}} onClick={()=>setIaSugestoes([])}>✕ Limpar</button>}
        </div>
      </div>
      {fila.length===0&&<div style={{fontSize:12,color:"var(--mt)",background:"var(--sx)",padding:"10px 14px",borderRadius:8}}>
        ℹ️ Adicione pacientes na fila acima para gerar sugestões de encaixe.
      </div>}
      {fila.length>0&&iaSugestoes.length===0&&!iaLoading&&<div style={{fontSize:12,color:"var(--mt)",background:"var(--sx)",padding:"10px 14px",borderRadius:8}}>
        Clique em "Gerar Sugestões" para analisar {fila.length} paciente(s) na fila.
      </div>}
      {iaSugestoes.length>0&&<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:8,marginTop:4}}>
        {iaSugestoes.map((s,i)=>(
          <div key={i} style={{background:"var(--sx)",border:"1px solid #a78bfa30",borderRadius:10,padding:"11px 14px",display:"flex",flexDirection:"column",gap:5}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
              <div>
                <div style={{fontWeight:800,fontSize:12,color:"var(--tx)"}}>{s.pac}</div>
                <div style={{fontSize:11,color:espCor(s.esp),fontWeight:700,marginTop:1}}>{s.esp}</div>
              </div>
              <span style={{fontSize:9,color:"#a78bfa",background:"#a78bfa20",padding:"2px 8px",borderRadius:20,fontWeight:800,flexShrink:0}}>{s.preferencia}</span>
            </div>
            <div style={{fontSize:11,color:"var(--mt)"}}>👩‍⚕️ {profShort(s.prof)}</div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:3}}>
              <span style={{fontWeight:900,color:"#a78bfa",fontSize:12}}>📅 {s.data}</span>
              <span style={{fontSize:12,color:"#34d399",fontWeight:800}}>🕐 {s.horario}</span>
            </div>
            {s.convenio&&<div style={{fontSize:10,color:"#22c55e"}}>💳 {s.convenio}</div>}
          </div>
        ))}
      </div>}
    </div>

    {modal&&<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&setModal(false)}>
      <div className="modal" style={{width:480}}>
        <div className="modal-head"><h2>{editing?"✏️ Editar":"+ Adicionar"} — Fila de Espera</h2><button className="icon-btn" onClick={()=>setModal(false)}>×</button></div>
        <div className="stack">
          <div><label>Paciente *</label><select value={form.pacienteId} onChange={e=>sf({pacienteId:e.target.value})}><option value="">Selecione...</option>{pacientes.map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}</select></div>
          <div className="g2">
            <div><label>Especialidade *</label><select value={form.especialidade} onChange={e=>sf({especialidade:e.target.value})}><option value="">Selecione...</option>{[...new Set(Object.keys(ESP_CORES).filter(k=>!["Agendamento","Faturamento","Atendimento","Supervisão ADM","Administrador","Financeiro","Gestão de Pessoas","Outro"].includes(k)))].map(e=><option key={e}>{e}</option>)}</select></div>
            <div><label>Convênio</label><select value={form.convenio||""} onChange={e=>sf({convenio:e.target.value})}><option value="">Selecione...</option>{CONVENIOS_LIST.map(c=><option key={c}>{c}</option>)}</select></div>
          </div>
          <div className="g2">
            <div><label>Preferência de Turno</label><select value={form.preferencia} onChange={e=>sf({preferencia:e.target.value})}>{["Manhã","Tarde","Integral","Qualquer"].map(p=><option key={p}>{p}</option>)}</select></div>
            <div><label>Profissional desejado</label><input value={form.profissionalDesejado||""} onChange={e=>sf({profissionalDesejado:e.target.value})} placeholder="Opcional"/></div>
          </div>
          <div><label>Observações</label><textarea rows={2} value={form.observacao||""} onChange={e=>sf({observacao:e.target.value})}/></div>
          <div className="actions"><button className="btn secondary" onClick={()=>setModal(false)}>Cancelar</button><button className="btn primary" onClick={save}>Salvar</button></div>
        </div>
      </div>
    </div>}
    {modalNR&&<div className="modal-bg" onClick={e=>e.target===e.currentTarget&&setModalNR(false)}>
      <div className="modal" style={{width:500}}>
        <div className="modal-head">
          <h2>{formNR.tipo==="glosa"?"🚫 Registrar Glosa":"📭 Registrar Pendência"}</h2>
          <button className="icon-btn" onClick={()=>setModalNR(false)}>×</button>
        </div>
        <div className="stack">
          <div className="g2">
            <div><label>Data</label><input type="date" value={formNR.data} onChange={e=>setFormNR(f=>({...f,data:e.target.value}))}/></div>
            <div><label>Convênio *</label>
              <select value={formNR.convenio} onChange={e=>setFormNR(f=>({...f,convenio:e.target.value}))}>
                <option value="">Selecione...</option>
                {[...new Set(agenda.map(a=>a.convenio||"").filter(cv=>cv&&cv!=="Particular"))].map(cv=><option key={cv}>{cv}</option>)}
              </select>
            </div>
          </div>
          <div className="g2">
            <div><label>Profissional</label>
              <select value={formNR.profissionalId} onChange={e=>setFormNR(f=>({...f,profissionalId:e.target.value}))}>
                <option value="">Todos</option>
                {profissionais.filter(p=>p.role==="profissional").map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}
              </select>
            </div>
            <div><label>Procedimento</label>
              <select value={formNR.procedimentoId} onChange={e=>{const proc=procedimentos.find(p=>p.id===Number(e.target.value));setFormNR(f=>({...f,procedimentoId:e.target.value,valor:proc?.valor||f.valor}));}}>
                <option value="">Selecione...</option>
                {procedimentos.map(p=><option key={p.id} value={p.id}>{p.nome}</option>)}
              </select>
            </div>
          </div>
          <div className="g2">
            <div><label>Valor (R$)</label><input type="number" min="0" step="0.01" value={formNR.valor} onChange={e=>setFormNR(f=>({...f,valor:Number(e.target.value)}))}/></div>
            <div><label>Tipo</label>
              <select value={formNR.tipo} onChange={e=>setFormNR(f=>({...f,tipo:e.target.value}))}>
                <option value="pendencia">Pendência — não confirmado</option>
                <option value="glosa">Glosa — recusado pelo convênio</option>
              </select>
            </div>
          </div>
          <div><label>Descrição / Motivo</label>
            <textarea rows={2} value={formNR.descricao} onChange={e=>setFormNR(f=>({...f,descricao:e.target.value}))} placeholder={formNR.tipo==="glosa"?"Ex: Guia vencida, procedimento não coberto...":"Ex: Lote enviado em 10/03, aguardando retorno..."}/>
          </div>
          <div className="actions">
            <button className="btn secondary" onClick={()=>setModalNR(false)}>Cancelar</button>
            <button className="btn primary" style={{background:formNR.tipo==="glosa"?"#dc2626":"#f97316",borderColor:formNR.tipo==="glosa"?"#dc2626":"#f97316"}}
              onClick={()=>{
                if(!formNR.convenio)return alert("Selecione o convênio");
                setNaoRecebiveis(nr=>[...nr,{...formNR,id:Date.now(),valor:Number(formNR.valor)}]);
                showToast((formNR.tipo==="glosa"?"🚫 Glosa":"📭 Pendência")+" registrada","ok");
                setModalNR(false);
              }}>
              Salvar
            </button>
          </div>
        </div>
      </div>
    </div>}

  </div>);
}


// ═══════════════════════════════════════════════════════════════════════════════
// CAIXINHA — Controle de recebimentos particulares
// ═══════════════════════════════════════════════════════════════════════════════
function CaixinhaPage({caixa,setCaixa,pacientes,profissionais,procedimentos,agenda,showToast}){
  const [tabC,setTabC]=useState("hoje");
  const [formaPagFiltro,setFormaPagFiltro]=useState("");
  const [dataFiltro,setDataFiltro]=useState(hoje_str);

  // Período
  const agora=new Date();
  const iniMes=ymd(new Date(agora.getFullYear(),agora.getMonth(),1));

  const entradas=(caixa||[]).filter(e=>{
    if(tabC==="hoje")return e.data===hoje_str;
    if(tabC==="mes")return e.data>=iniMes&&e.data<=hoje_str;
    if(tabC==="data")return e.data===dataFiltro;
    return true;
  }).filter(e=>!formaPagFiltro||e.formaPagamento===formaPagFiltro)
    .sort((a,b)=>b.data.localeCompare(a.data)||b.hora.localeCompare(a.hora));

  const totalGeral=entradas.reduce((s,e)=>s+Number(e.valor||0),0);
  const porForma=["Dinheiro","Cartão","PIX"].map(fp=>({
    fp,
    total:entradas.filter(e=>e.formaPagamento===fp).reduce((s,e)=>s+Number(e.valor||0),0),
    qtd:entradas.filter(e=>e.formaPagamento===fp).length,
  }));

  const remover=(id)=>{
    if(!confirm("Remover este registro do caixa?"))return;
    setCaixa(c=>c.filter(e=>e.id!==id));
    showToast("🗑️ Registro removido","warn");
  };

  const exportarCSV=()=>{
    const cab=["Data","Hora","Paciente","Profissional","Procedimento","Valor","Forma","Descrição"];
    const linhas=entradas.map(e=>{
      const pac=pacientes.find(p=>p.id===Number(e.pacienteId));
      const prof=profissionais.find(p=>p.id===Number(e.profissionalId));
      const proc=procedimentos.find(p=>p.id===Number(e.procedimentoId));
      return[brDate(e.data),e.hora,pac?.nome||"",prof?.nome||"",proc?.nome||"",e.valor,e.formaPagamento,e.descr||""];
    });
    exportCSV("caixinha_"+hoje_str,cab,linhas);
    showToast("📥 CSV exportado","ok");
  };

  const COR_FP={Dinheiro:"#4ade80",Cartão:"#a78bfa",PIX:"#a78bfa"};
  const ICON_FP={Dinheiro:"💵",Cartão:"💳",PIX:"📱"};

  return(<div className="page-wrap">
    <div className="page-head">
      <h1>💵 Caixinha — Particulares</h1>
      <div style={{display:"flex",gap:7}}>
        {entradas.length>0&&<button className="btn secondary" style={{fontSize:11}} onClick={exportarCSV}>📥 Exportar CSV</button>}
      </div>
    </div>

    {/* ── KPIs ── */}
    <div style={{display:"grid",gridTemplateColumns:"1fr repeat(3,1fr)",gap:8,marginBottom:16}}>
      <div className="card" style={{padding:"14px 16px",borderTop:"3px solid #4ade80"}}>
        <div style={{fontFamily:"'DM Serif Display',serif",fontWeight:900,fontSize:24,color:"#4ade80"}}>{brl(totalGeral)}</div>
        <div style={{fontSize:10,fontWeight:900,color:"var(--mt)",textTransform:"uppercase",marginTop:3}}>Total do período</div>
        <div style={{fontSize:10,color:"var(--mt)"}}>{entradas.length} recebimento(s)</div>
      </div>
      {porForma.map(({fp,total,qtd})=>(
        <div key={fp} className="card" style={{padding:"12px 14px",borderTop:"3px solid "+COR_FP[fp]}}>
          <div style={{fontSize:18,fontWeight:900,color:COR_FP[fp]}}>{brl(total)}</div>
          <div style={{fontSize:10,fontWeight:900,color:"var(--mt)",textTransform:"uppercase",marginTop:2}}>{ICON_FP[fp]} {fp}</div>
          <div style={{fontSize:10,color:"var(--mt)"}}>{qtd} recebimento(s)</div>
        </div>
      ))}
    </div>

    {/* ── Filtros ── */}
    <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
      {[["hoje","Hoje"],["mes","Este mês"],["data","Data específica"],["tudo","Tudo"]].map(([v,l])=>(
        <button key={v} onClick={()=>setTabC(v)}
          style={{padding:"4px 14px",borderRadius:20,fontSize:11,fontWeight:800,border:"1.5px solid "+(tabC===v?"#4ade80":"var(--cpb)"),background:tabC===v?"#4ade8015":"transparent",color:tabC===v?"#4ade80":"var(--mt)",cursor:"pointer"}}>{l}</button>
      ))}
      {tabC==="data"&&<input type="date" value={dataFiltro} onChange={e=>setDataFiltro(e.target.value)} style={{fontSize:12,width:150}}/>}
      <div style={{width:1,height:20,background:"var(--sc)"}}/>
      {["","Dinheiro","Cartão","PIX"].map(fp=>(
        <button key={fp} onClick={()=>setFormaPagFiltro(fp)}
          style={{padding:"3px 10px",borderRadius:20,fontSize:11,border:"1.5px solid "+(formaPagFiltro===fp?(COR_FP[fp]||"#a78bfa"):"var(--cpb)"),background:formaPagFiltro===fp?(COR_FP[fp]||"#a78bfa")+"15":"transparent",color:formaPagFiltro===fp?(COR_FP[fp]||"#a78bfa"):"var(--mt)",cursor:"pointer",fontWeight:formaPagFiltro===fp?800:500}}>
          {fp?ICON_FP[fp]+" "+fp:"Todas formas"}
        </button>
      ))}
    </div>

    {/* ── Tabela ── */}
    {entradas.length===0
      ?<div style={{textAlign:"center",padding:"48px 0",color:"var(--mt)",fontSize:13}}>
          Nenhum recebimento registrado{tabC==="hoje"?" hoje":""}.
          <div style={{fontSize:11,marginTop:6}}>Registre pagamentos particulares na janela da sessão.</div>
        </div>
      :<div className="card" style={{overflow:"hidden"}}>
        <div className="grid-header" style={{gridTemplateColumns:"90px 55px 1fr 130px 130px 100px 70px 40px"}}>
          <div>Data</div><div>Hora</div><div>Paciente</div><div>Profissional</div><div>Procedimento</div><div>Valor</div><div>Forma</div><div></div>
        </div>
        {entradas.map((e,i)=>{
          const pac=pacientes.find(p=>p.id===Number(e.pacienteId));
          const prof=profissionais.find(p=>p.id===Number(e.profissionalId));
          const proc=procedimentos.find(p=>p.id===Number(e.procedimentoId));
          const cor=COR_FP[e.formaPagamento]||"#64748b";
          return(<div key={e.id} className="grid-row" style={{gridTemplateColumns:"90px 55px 1fr 130px 130px 100px 70px 40px",background:i%2?"var(--gr)":""}}>
            <div style={{fontSize:11,fontWeight:700}}>{brDate(e.data)}</div>
            <div style={{fontSize:11,color:"#a78bfa",fontWeight:800}}>{e.hora}</div>
            <div style={{fontSize:11,fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{pac?.nome||"—"}</div>
            <div style={{fontSize:10,color:"var(--mt)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{prof?.nome||"—"}</div>
            <div style={{fontSize:10,color:"var(--mt)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{proc?.nome||"—"}</div>
            <div style={{fontSize:12,fontWeight:900,color:"#4ade80"}}>{brl(Number(e.valor||0))}</div>
            <div><span style={{fontSize:10,padding:"2px 7px",borderRadius:20,background:cor+"20",color:cor,fontWeight:800,border:"1px solid "+cor+"40"}}>{ICON_FP[e.formaPagamento]||""} {e.formaPagamento}</span></div>
            <div><button onClick={()=>remover(e.id)} style={{background:"none",border:"none",cursor:"pointer",color:"var(--mt)",fontSize:13}} title="Remover">✕</button></div>
          </div>);
        })}
        <div className="grid-row" style={{gridTemplateColumns:"90px 55px 1fr 130px 130px 100px 70px 40px",background:"#4ade8010",fontWeight:900}}>
          <div style={{gridColumn:"1/6",fontSize:12,color:"#4ade80"}}>TOTAL ({entradas.length} recebimento(s))</div>
          <div style={{fontSize:13,fontWeight:900,color:"#4ade80"}}>{brl(totalGeral)}</div>
        </div>
      </div>}
  </div>);
}


function exportCSV(nome, cabecalho, linhas) {
  const esc = v => '"' + String(v ?? "").replace(/"/g, '""') + '"';
  const csv = [cabecalho, ...linhas].map(r => r.map(esc).join(",")).join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = nome + ".csv"; a.click();
  URL.revokeObjectURL(url);
}


const TIPOS_USUARIO_DEF=[
  {role:"agendamento",label:"Supervisor/Agente Agendamento",descr:"Acesso à agenda para edição completa, edição completa do cadastro de pacientes",permissoes:["agenda_edit","paciente_edit","chamados"]},
  {role:"faturamento_supervisor",label:"Supervisor de Faturamento",descr:"Acesso à agenda para consulta e confirmação no botão Faturado, página faturamento com análise de valores, abertura de faturas e repasse de profissionais",permissoes:["agenda_view","faturamento","financeiro"]},
  {role:"coordenador",label:"Coordenador",descr:"Acesso à agenda completo sem edição, acesso ao prontuário (exceto psicologia/psicanálise) com permissão de edição",permissoes:["agenda_view","prontuario_edit_parcial"]},
  {role:"coordenador_aba",label:"Coordenador ABA",descr:"Acesso à agenda completo sem edição, acesso a todos os prontuários com permissão de edição",permissoes:["agenda_view","prontuario_edit_total"]},
  {role:"profissional",label:"Profissional",descr:"Acesso à agenda própria, registro de evoluções no prontuário",permissoes:["agenda_view_own","prontuario_own"]},
  {role:"administrador",label:"Administrador",descr:"Acesso completo ao sistema",permissoes:["all"]},
];


// Carregado sob demanda: a página de Gestão (relatórios, atividades,
// importar/exportar) só é baixada quando o usuário abre essa aba.
const GestaoPage = lazy(() => import("./pages/GestaoPage"));

export default function App({themeKey:themeKeyExterno,onToggleTheme,themeLabel}={}){
  // Se a casca de SSO (App.jsx / theme/cuideTheme.js) fornecer o tema,
  // usamos ela como fonte única de verdade — assim o botão de tema do
  // AppSwitcher (barra de cima), da tela de login e daqui de dentro do
  // Cuide ficam sincronizados nos MESMOS 3 modos do MiContas (escuro/
  // médio/claro), em vez de serem sistemas desconectados como antes.
  // Se ninguém fornecer nada (uso avulso do Cuide fora da casca), cai
  // num ciclo local de 3 modos equivalente.
  const [themeKeyLocal,setThemeKeyLocal]=useState("dark");
  const themeKey=themeKeyExterno||themeKeyLocal;
  const CICLO=["dark","mid","light"];
  const toggleTheme=onToggleTheme||(()=>setThemeKeyLocal(k=>CICLO[(CICLO.indexOf(k)+1)%CICLO.length]));
  const ROTULO_LOCAL={dark:"🌙 Escuro",mid:"🌓 Médio",light:"☀️ Claro"};
  const [sidebarExpanded,setSidebarExpanded]=useState(true);
  const [alertPanelOpen,setAlertPanelOpen]=useState(false);
  const T=THEMES_CUIDE[themeKey]||DARK;
  const themeVars=getThemeVars(T);

  const [page,setPage]=useState("agenda");
  const [profissionais,setProfissionais]=useState(seedProfissionais);
  const [pacientes,setPacientes]=useState(seedPacientes);
  const [procedimentos,setProcedimentos]=useState(seedProcedimentos);
  const [convenios,setConvenios]=useState(seedConvenios);
  const [salas,setSalas]=useState(seedSalas);
  const [filiais,setFiliais]=useState(seedFiliais);
  const [agenda,setAgenda]=useState(seedAgenda);
  const [alertas,setAlertas]=useState(seedAlertas);
  const [fila,setFila]=useState([{id:1,pacienteId:1,especialidade:"Psicologia Convencional/Psicanálise",preferencia:"Manhã",observacao:"Pode encaixe"},{id:2,pacienteId:2,especialidade:"Fisioterapia",preferencia:"Tarde",observacao:"3x/semana"}]);
  const [manuais,setManuais]=useState(seedManuais);
  const [templatePaciente,setTemplatePaciente]=useState(TEMPLATE_CONTRATO_PACIENTE);
  const [templateProfissional,setTemplateProfissional]=useState(TEMPLATE_CONTRATO_PROFISSIONAL);
  const [modelosEvolucaoEdit,setModelosEvolucaoEdit]=useState(MODELOS_EVOLUCAO);
  const [prontuarios,setProntuarios]=useState([]);
  const [pedidos,setPedidos]=useState([]);
  const [horariosF,setHorariosF]=useState([]);
  const [lancamentos,setLancamentos]=useState([]);
  const [naoRecebiveis,setNaoRecebiveis]=useState([]);
  const [caixa,setCaixa]=useState([]);
  const [chamados,setChamados]=useState([
    {id:1,numero:"100001",setor:"Faturamento",tipo:"novo_paciente",nome:"João Pedro Alves",descricao:"Novo cadastro",data:hoje_str,status:"aberto",resp:""},
    {id:2,numero:"100002",setor:"Faturamento",tipo:"novo_profissional",nome:"Dra. Ana Souza",descricao:"Novo profissional",data:hoje_str,status:"aberto",resp:""},
  ]);
  const [auth,setAuth]=useState({usuario:"31028313896",role:"administrador",nome:"Admin Sistema"});
  const [loginOpen,setLoginOpen]=useState(false);
  const [agModal,setAgModal]=useState(false);
  const [sessaoModal,setSessaoModal]=useState(null);
  const [reuniaoModal,setReuniaoModal]=useState(null);
  const [atividades,setAtividades]=useState([
    {id:1,data:hoje_str,hora:"08:00",usuario:"Admin Sistema",acao:"Login",detalhe:"Acesso ao sistema"},
  ]);
  const [agEditing,setAgEditing]=useState(null);
  const [agDefault,setAgDefault]=useState({});
  const [selDate,setSelDate]=useState(hoje_str);
  const [toast,setToast]=useState(null);
  const showToast=(msg,type="ok")=>{setToast({msg,type});setTimeout(()=>setToast(null),4000);};
  const abrirNovo=(date,extra={})=>{setSelDate(date||hoje_str);setAgEditing(null);setAgDefault(extra||{});setAgModal(true);};
  const toggleFecharHorario=(chave,fechar)=>setHorariosF(a=>fechar?[...a,chave]:a.filter(x=>x!==chave));
  const logActivity=(acao,detalhe)=>{setAtividades(a=>[{id:Date.now(),data:hoje_str,hora:new Date().toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"}),usuario:auth.nome||auth.usuario,acao,detalhe},...a]);};
  const abrirSessao=(ag)=>{if(ag.tipo==="reuniao")setReuniaoModal(ag);else setSessaoModal(ag);};
  const chamadosAbertos=chamados.filter(c=>c.status==="aberto").length;
  const pedidosAlerta=pedidos.filter(p=>{if(!p.dataValidade)return false;return diffDays(hoje_str,p.dataValidade)<=30;}).length;

  const NAV=[
    ["agenda","📅","Agenda"],
    ["espelho","🪞","Espelho"],
    ["salasMapa","🗺️","Mapeamento"],
    ["fila","⏳","Fila de Espera"],
    ["pacientes","👤","Pacientes"],
    ["profissionais","🩺","Profissionais"],
    ["convenios","🏥","Convênios"],
    ["salas","🏢","Salas & Filiais"],
    ["pedidos","🩻","Pedidos Médicos",pedidosAlerta],
    ["caixa","💵","Caixinha"],
    ["financeiro","💼","Faturamento"],
    ["chamados","📨","Chamados",chamadosAbertos],

    ["gestao","📊","Gestão"],
    ["manuais","📚","Manuais"],
  ];


  const pendAlertas=alertas.filter(a=>!a.lido).length;
  const userProf=profissionais.find(p=>p.usuario===auth.usuario);
  const SBW=sidebarExpanded?220:58;

  return(<div className="app-shell">
    <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;700;800;900&family=DM+Serif+Display:ital@0;1&display=swap"/>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"/>
    <style>{themeVars+CSS}</style>

    <aside className="sidebar" style={{width:SBW}}>
      <div className="sidebar-inner">
        <button className="sidebar-toggle" onClick={()=>setSidebarExpanded(v=>!v)} title={sidebarExpanded?"Recolher menu":"Expandir menu"}>{sidebarExpanded?"◀":"▶"}</button>
        <div className="sidebar-brand">
          <div className="brand-logo">
            <svg width="26" height="26" viewBox="0 0 26 26" fill="none">
              {/* Petal top */}
              <ellipse cx="13" cy="6.5" rx="2.4" ry="4.5" fill="white" opacity="0.95"/>
              {/* Petal bottom */}
              <ellipse cx="13" cy="19.5" rx="2.4" ry="4.5" fill="white" opacity="0.6"/>
              {/* Petal left */}
              <ellipse cx="6.5" cy="13" rx="4.5" ry="2.4" fill="white" opacity="0.75"/>
              {/* Petal right */}
              <ellipse cx="19.5" cy="13" rx="4.5" ry="2.4" fill="white" opacity="0.75"/>
              {/* Petal top-right */}
              <ellipse cx="18.5" cy="7.5" rx="2.4" ry="4.2" fill="#99f6e4" opacity="0.85" transform="rotate(45 18.5 7.5)"/>
              {/* Petal bottom-left */}
              <ellipse cx="7.5" cy="18.5" rx="2.4" ry="4.2" fill="#99f6e4" opacity="0.65" transform="rotate(45 7.5 18.5)"/>
              {/* Center circle */}
              <circle cx="13" cy="13" r="3.2" fill="white"/>
              <circle cx="13" cy="13" r="1.5" fill="#0f766e"/>
            </svg>
          </div>
          {sidebarExpanded&&<div style={{display:"flex",flexDirection:"column",justifyContent:"center",gap:0,lineHeight:1}}><div style={{fontFamily:"'DM Serif Display',serif",fontWeight:500,fontSize:9,letterSpacing:"3.5px",textTransform:"uppercase",color:"#a78bfa",opacity:.7,marginBottom:1}}>FOCOE</div><div style={{fontFamily:"'DM Serif Display',serif",fontWeight:800,fontSize:20,letterSpacing:"-0.5px",color:"#7c6af7",lineHeight:1}}>Cuide</div></div>}
        </div>
        <div className="sidebar-user">
          <div className="user-avatar">{(userProf?.nome||auth.usuario||"?")[0].toUpperCase()}</div>
          {sidebarExpanded&&<div className="user-info">
            <div className="user-name">{userProf?.nome?.split(" ")[0]||auth.usuario}</div>
            <div className="user-role" style={{color:espCor((userProf?.especialidades||["Outro"])[0])}}>{PERFIL_LABEL[auth.role]||auth.role}</div>
          </div>}
        </div>
        <nav className="sidebar-nav">
          <button className="nav-btn" onClick={()=>setAlertPanelOpen(true)} style={{position:"relative"}}>
            <span className="nav-icon" style={{position:"relative"}}>
              🔔
              {pendAlertas>0&&<span style={{position:"absolute",top:-4,right:-4,width:14,height:14,borderRadius:"50%",background:"#ef4444",color:"#fff",fontSize:8,fontWeight:900,display:"flex",alignItems:"center",justifyContent:"center"}}>{pendAlertas}</span>}
            </span>
            {sidebarExpanded?<span className="nav-label">Alertas {pendAlertas>0&&<span className="nav-badge">{pendAlertas}</span>}</span>:<span className="nav-tooltip">Alertas{pendAlertas>0?" ("+pendAlertas+")":""}</span>}
          </button>
          {NAV.map(([key,icon,label,badge])=>(
            <button key={key} className={"nav-btn"+(page===key?" active":"")} onClick={()=>setPage(key)}>
              <span className="nav-icon">{icon}</span>
              {sidebarExpanded?<><span className="nav-label">{label}</span>{badge>0&&<span className="nav-badge">{badge}</span>}</>:<span className="nav-tooltip">{label}{badge>0?" ("+badge+")":""}</span>}
            </button>
          ))}
        </nav>
        <div className="sidebar-footer">
          {sidebarExpanded?<>
            <button className="btn secondary" style={{width:"100%",fontSize:11,display:"flex",alignItems:"center",gap:5,justifyContent:"center"}} onClick={toggleTheme}>{themeLabel||ROTULO_LOCAL[themeKey]}</button>
            <button className="btn secondary" style={{width:"100%",fontSize:11}} onClick={()=>setLoginOpen(true)}>🔐 Trocar usuário</button>
          </>:<>
            <button className="nav-btn" onClick={toggleTheme} style={{justifyContent:"center"}}><span className="nav-icon">{{dark:"🌙",mid:"🌓",light:"☀️"}[themeKey]}</span><span className="nav-tooltip">{themeLabel||ROTULO_LOCAL[themeKey]}</span></button>
            <button className="nav-btn" onClick={()=>setLoginOpen(true)} style={{justifyContent:"center"}}><span className="nav-icon">🔐</span><span className="nav-tooltip">Trocar usuário</span></button>
          </>}
        </div>
      </div>
    </aside>

    <AlertPanel alertas={alertas} setAlertas={setAlertas} pacientes={pacientes} profissionais={profissionais} agenda={agenda} auth={auth} open={alertPanelOpen} onClose={()=>setAlertPanelOpen(false)}/>
    <ChamadoButton setChamados={setChamados} showToast={showToast}/>

    <main className="main">
      <ErrorBoundary key={page}>
      {page==="agenda"&&<MainAgenda auth={auth} agenda={agenda} setAgenda={setAgenda} pacientes={pacientes} profissionais={profissionais} salas={salas} filiais={filiais} alertas={alertas} setAlertas={setAlertas} onAbrirNovo={abrirNovo} onAbrirSessao={abrirSessao}/>}
      {page==="espelho"&&<EspelhoAgendaPage data={selDate} agenda={agenda} pacientes={pacientes} profissionais={profissionais} salas={salas} filiais={filiais} convenios={convenios} procedimentos={procedimentos} horariosF={horariosF} onFecharHorario={toggleFecharHorario} onAbrirAgendamento={d=>abrirNovo(d.data,{profissionalId:d.profId})} onAbrirSessao={abrirSessao} onBuscaInteligente={()=>setPage("fila")}/>}
      {page==="salasMapa"&&<MapeamentoSalasPage data={selDate} agenda={agenda} salas={salas} filiais={filiais} pacientes={pacientes} profissionais={profissionais} horariosF={horariosF} onFecharHorario={toggleFecharHorario} onAbrirAgendamento={d=>abrirNovo(d.data,{salaId:d.salaId})} onAbrirSessao={abrirSessao} onBuscaInteligente={()=>setPage("fila")}/>}
      {page==="fila"&&<WaitingListPage fila={fila} setFila={setFila} pacientes={pacientes} auth={auth} showToast={showToast} agenda={agenda} profissionais={profissionais}/>}
      {page==="pacientes"&&<PacientesPage pacientes={pacientes} setPacientes={setPacientes} setChamados={setChamados} showToast={showToast} auth={auth} agenda={agenda} profissionais={profissionais} procedimentos={procedimentos} prontuarios={prontuarios} setProntuarios={setProntuarios} logActivity={logActivity} templatePaciente={templatePaciente}/>}
      {page==="profissionais"&&<ProfissionaisPage profissionais={profissionais} setProfissionais={setProfissionais} setChamados={setChamados} showToast={showToast} auth={auth} filiais={filiais} logActivity={logActivity} agenda={agenda} procedimentos={procedimentos} lancamentos={lancamentos} setLancamentos={setLancamentos} convenios={convenios} templateProfissional={templateProfissional} pacientes={pacientes}/>}
      {page==="convenios"&&<ConveniosProcedimentosPage convenios={convenios} setConvenios={setConvenios} procedimentos={procedimentos} setProcedimentos={setProcedimentos} showToast={showToast}/>}
      {page==="salas"&&<SalasPage salas={salas} setSalas={setSalas} filiais={filiais} setFiliais={setFiliais} showToast={showToast}/>}
      {page==="pedidos"&&<PedidosMedicosPage pedidos={pedidos} setPedidos={setPedidos} pacientes={pacientes} setChamados={setChamados} showToast={showToast}/>}
      {page==="caixa"&&<CaixinhaPage caixa={caixa} setCaixa={setCaixa} pacientes={pacientes} profissionais={profissionais} procedimentos={procedimentos} agenda={agenda} showToast={showToast}/>}
      {page==="financeiro"&&<FinanceiroPage agenda={agenda} setAgenda={setAgenda} pacientes={pacientes} profissionais={profissionais} procedimentos={procedimentos} convenios={convenios} showToast={showToast} naoRecebiveis={naoRecebiveis} setNaoRecebiveis={setNaoRecebiveis}/>}
      {page==="chamados"&&<ChamadosPage chamados={chamados} setChamados={setChamados} showToast={showToast} fila={fila} setFila={setFila} pacientes={pacientes}/>}

      {page==="gestao"&&<Suspense fallback={<div style={{padding:40,textAlign:"center"}}>Carregando…</div>}><GestaoPage agenda={agenda} pacientes={pacientes} profissionais={profissionais} procedimentos={procedimentos} convenios={convenios} salas={salas} filiais={filiais} chamados={chamados} fila={fila} pedidos={pedidos} auth={auth} showToast={showToast} atividades={atividades} setPacientes={setPacientes} setProfissionais={setProfissionais} setAgenda={setAgenda}/></Suspense>}
            {page==="manuais"&&<ManuaisPage manuais={manuais} setManuais={setManuais} showToast={showToast} templatePaciente={templatePaciente} setTemplatePaciente={setTemplatePaciente} templateProfissional={templateProfissional} setTemplateProfissional={setTemplateProfissional} modelosEvolucao={modelosEvolucaoEdit} setModelosEvolucao={setModelosEvolucaoEdit}/>}
      </ErrorBoundary>
    </main>
    {agModal&&<AgendaModal auth={auth} profissionais={profissionais} pacientes={pacientes} procedimentos={procedimentos} salas={salas} agendamentos={agenda} convenios={convenios} manuais={manuais} filiais={filiais} editing={agEditing} defaultData={selDate} defaultProfId={agDefault.profissionalId} defaultSalaId={agDefault.salaId} onClose={()=>setAgModal(false)}
      onSave={(novo,multiplo)=>{
        if(multiplo&&Array.isArray(novo)){setAgenda(a=>[...a,...novo]);logActivity("Agendamento","Criadas "+novo.length+" sessions");showToast("✅ "+novo.length+" sessões criadas","ok");}
        else if(agEditing){setAgenda(a=>a.map(x=>x.id===agEditing.id?{...novo,id:agEditing.id}:x));showToast("✅ Atualizado","ok");}
        else{setAgenda(a=>[...a,{...novo,id:Date.now()}]);showToast("✅ Agendamento salvo","ok");}
        setAgModal(false);
      }}/>}
    {reuniaoModal&&<ReuniaoAtaModal ag={reuniaoModal} profissionais={profissionais} chamados={chamados} setChamados={setChamados} showToast={showToast} auth={auth}
      onClose={()=>setReuniaoModal(null)}
      onEditar={ag=>{setAgenda(a=>a.map(x=>x.id===ag.id?ag:x));setReuniaoModal(ag);}}/>}
        {sessaoModal&&<SessaoModal ag={sessaoModal} pacientes={pacientes} profissionais={profissionais} procedimentos={procedimentos} salas={salas} filiais={filiais} convenios={convenios} auth={auth} agenda={agenda} prontuarios={prontuarios} setProntuarios={setProntuarios} modelosEv={modelosEvolucaoEdit} caixa={caixa} setCaixa={setCaixa} onClose={()=>setSessaoModal(null)}
      onEditar={ag=>{setAgEditing(ag);setAgDefault({});setAgModal(true);setPage("agenda");}}
      onSave={ag=>{
        const prev=agenda.find(x=>x.id===ag.id);
        setAgenda(a=>a.map(x=>x.id===ag.id?ag:x));
        const proc=procedimentos.find(p=>p.id===Number(ag.procedimentoId));
        const prof=profissionais.find(p=>p.id===Number(ag.profissionalId));
        const conv=ag.convenio||"Particular";

        // ── ATENDIDO → Gera repasse ao profissional (independente de faturamento) ──
        const virarAtendido=(ag.status==="atendido"||ag.status==="faltou_pacote")&&(prev?.status!=="atendido"&&prev?.status!=="faltou_pacote");
        if(virarAtendido){
          const linhaTabela=(prof?.tabelaRepasse||[]).find(r=>
            Number(r.procedimentoId)===Number(ag.procedimentoId)&&r.convenio===conv
          )||(prof?.tabelaRepasse||[]).find(r=>
            Number(r.procedimentoId)===Number(ag.procedimentoId)
          );
          const valBruto=linhaTabela?.valorTabela??proc?.valor??0;
          const pctRepasse=linhaTabela?.pctRepasse??proc?.pctRepasse??45;
          const valRep=Math.round(valBruto*(pctRepasse/100)*100)/100;
          const fonte=linhaTabela?"tabela profissional":"tabela global";
          setLancamentos(l=>[...l,{
            id:Date.now(),tipo:"CP",agId:ag.id,data:ag.data,convenio:conv,
            pacienteId:ag.pacienteId,profissionalId:ag.profissionalId,procedimentoId:ag.procedimentoId,
            valor:valRep,pctRepasse,status:"pendente",
            descr:"CP — "+(proc?.nome||"Sessão")+" — "+brDate(ag.data)+" — "+(prof?.nome||"")+" ["+fonte+"]"
          }]);
          showToast("💸 Repasse "+brl(valRep)+" ("+pctRepasse+"%) → "+profShort(prof?.nome||""),"ok");
          logActivity("Repasse","CP R$"+valRep+" → "+(prof?.nome||"")+" | "+(proc?.nome||"")+" | "+fonte);
        }

        // ── FATURADO (fatStatus) → Gera crédito do convênio (independente do repasse) ──
        if(ag.fatStatus==="faturado"&&prev?.fatStatus!=="faturado"){
          const valBruto=proc?.valor??0;
          setLancamentos(l=>[...l,{
            id:Date.now()+1,tipo:"CR",agId:ag.id,data:ag.data,convenio:conv,
            pacienteId:ag.pacienteId,profissionalId:ag.profissionalId,procedimentoId:ag.procedimentoId,
            valor:valBruto,status:"recebido",
            descr:"CR — "+(proc?.nome||"Sessão")+" — "+brDate(ag.data)+" — "+conv
          }]);
          showToast("✅ Crédito "+brl(valBruto)+" registrado — "+conv,"ok");
          logActivity("Faturamento","CR R$"+valBruto+" — "+(proc?.nome||"")+" — "+conv);
        }



        // ── Outros status — log simples ──
        if(!virarAtendido&&ag.status!==prev?.status){
          logActivity("Status","Sessão #"+ag.id+" → "+(STATUS_AG[ag.status]?.label||ag.status));
          if(ag.status!=="atendido")showToast("✅ "+(STATUS_AG[ag.status]?.label||ag.status),"ok");
        }
        setSessaoModal(null);
      }}/>}
    {loginOpen&&<LoginModal profissionais={profissionais} onClose={()=>setLoginOpen(false)} onLogin={user=>{setAuth({usuario:user.usuario,role:user.role,nome:user.nome});setLoginOpen(false);showToast("✅ Bem-vindo, "+profShort(user.nome),"ok");}}/>}
    {toast&&<div className="toast" style={{background:{ok:"#022c22",err:"#3f0a0a",warn:"#3f1a03"}[toast.type]||"#022c22",color:{ok:"#34d399",err:"#f87171",warn:"#fbbf24"}[toast.type]||"#34d399",border:"1px solid "+(toast.type==="ok"?"#34d39930":toast.type==="err"?"#f8717130":"#fbbf2430")}}>{toast.msg}</div>}
  </div>);
}
