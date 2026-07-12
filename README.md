# Cuide × MiContas — camada de integração

Este pacote **não altera nada do MiContas**. Ele é a casca de sessão/SSO que
envolve o seu `Cuide.jsx` já existente, conectando-o ao mesmo Firebase do
MiContas.

## Por que isso é diferente do plano original que você colou

O plano anterior (aquele com `fetch("https://api.micontas.com.br/...")` e
JWT) assume que o MiContas tem uma API REST própria. Pelo que você já me
contou sobre o MiContas em conversas anteriores — e pelo que você mesmo
descreveu (Firebase Realtime Database, sem backend próprio) — **esse
endpoint não existe**. O MiContas lê e escreve direto no Firebase a partir
do navegador.

Por isso a integração real acontece assim:

```
Cuide (cuide.apoiare.com.br)  ──┐
                                 ├──►  MESMO projeto Firebase (Realtime Database)
MiContas (seu domínio atual)  ──┘
```

Nada de API intermediária: o Cuide lê/escreve direto nos mesmos dados que
o MiContas já usa. É mais simples e mais barato que manter uma API própria.

## O que você precisa fazer para isso funcionar de verdade

1. **`src/config/firebase.js`** — cole a mesma `firebaseConfig` que está no
   `index.html` do MiContas (mesmo projeto, mesmas credenciais).
2. **Crie a flag `cuideAtivo`** dentro de cada objeto de membro em
   `micontas_v15/{emailKey}/membros[i].cuideAtivo`, e adicione no painel
   admin do MiContas um botão "Liberar Cuide" que grava `true` ali. Essa é
   a única peça que efetivamente entra no MiContas — e é deliberada, já
   que foi você quem pediu esse botão.

## Estrutura real do banco (confirmada)

```
micontas_v15/
  {emailComoChave}/          ← e-mail com pontos trocados por "_"
    membros: [
      { nome, email, acesso, senha (hash SHA-256), cuideAtivo, ... }
    ]
    movimentacoes: [...]     ← AINDA NÃO CONFIRMADO — assumido no mesmo nível
    contasAPagar: [...]      ← AINDA NÃO CONFIRMADO
    contasAReceber: [...]    ← AINDA NÃO CONFIRMADO
```

Tudo isso já está refletido em `src/services/micontasApi.js` — é o único
arquivo que sabe desses nomes. `LoginSSO.jsx` e `RequireCuideAccess.jsx` só
chamam funções desse serviço (`buscarContaPorEmail`, `verificarAcessoCuide`,
`buscarPerfil`) e não sabem nada sobre a forma do JSON.

### Duas pendências que só você resolve

1. **Hash da senha**: implementei SHA-256 puro (sem sal) via Web Crypto,
   tentando reproduzir a função `hashSenha` do MiContas. Se a função real
   tiver sal, normalização de caixa, ou qualquer passo a mais, o hash vai
   bater diferente e o login vai "dar senha errada" mesmo com a senha
   certa. Cole a função `hashSenha` real do MiContas (ou o `tentarLoginLegado`)
   e eu ajusto para bater byte a byte.
2. **Login por e-mail, não por CPF**: como a chave da conta é o e-mail, o
   campo de login do Cuide mudou de CPF para e-mail. Se seus clientes
   preferem logar com CPF no dia a dia, precisa de um índice extra
   (`cpf -> email`) no banco — hoje isso não existe no pacote.
3. **`movimentacoes` / `contasAPagar` / `contasAReceber`**: assumi que
   ficam no mesmo nível de `membros`, dentro da conta. Confirme os nomes
   reais (e se são array ou nó com push-keys) antes de ligar
   `enviarEntradaCaixa` de verdade — se estiver errado, o Cuide vai
   escrever num nó que o MiContas nunca lê.


## Erro "Failed to load module script... MIME type text/plain" (Hostinger)

Se você viu esse erro depois do deploy: a causa era um `index.html` de
origem que na verdade era um **arquivo de build antigo** (apontava direto
para `./assets/index-cwinjmsm.js`, um nome de arquivo com hash que só
existia numa build anterior). O Vite gera nomes com hash novos a cada
build (`npm run build`), então esse caminho fixo ficava desatualizado, o
`.js` dava 404, o Hostinger devolvia uma página de erro em texto puro, e o
navegador recusava carregar isso como módulo — daí o erro de MIME type.

O que corrigi neste pacote:

- **`index.html`** (raiz do projeto) agora é o arquivo de *origem* correto
  do Vite: `<script type="module" src="/src/main.jsx">`. O Vite injeta o
  nome com hash certo automaticamente na build — você nunca mais edita
  esse caminho à mão.
- **`src/main.jsx`** — faltava esse ponto de entrada real (o que faz o
  `ReactDOM.createRoot(...).render(<App />)`). Sem ele, o build nem tinha
  como funcionar direito.
- **`vite.config.js`** — configuração mínima com o plugin do React.
- **`public/favicon.svg`** e **`public/icons.svg`** — o Vite copia tudo
  que está em `public/` direto para a raiz da build, então o favicon
  continua funcionando em `/favicon.svg`.
- **`public/.htaccess`** — força o `Content-Type` certo para `.js`/`.css`
  no Apache/LiteSpeed da Hostinger (proteção extra, caso o problema volte
  por outro motivo) e adiciona fallback de SPA.

### Como aplicar no seu repositório existente

1. Substitua o `index.html` da raiz do seu repo pelo daqui.
2. Adicione `src/main.jsx` e `vite.config.js` (arquivos novos).
3. Copie `favicon.svg` e `icons.svg` para dentro de uma pasta `public/` na
   raiz do repo (crie a pasta se não existir).
4. Adicione `public/.htaccess`.
5. Apague o antigo `index.html` compilado se ele ainda estiver por aí solto.
6. Dê commit e push na branch `main` — o GitHub Action que você já tem
   (`gemini-code-1783734757199.yaml`) vai rodar `npm install && npm run
   build` e mandar o `dist/` novo, correto, para o Hostinger via FTP.

Também limpei o `.gitignore`: a linha `assets/` que estava lá ignorava
qualquer pasta chamada `assets` em qualquer lugar do repo, o que não faz
sentido agora que os estáticos vivem em `public/` — removi essa linha
(mantive `dist`, `node_modules`, `.env`, `.DS_Store`).


Como Cuide e MiContas ficam em domínios diferentes, o navegador não
compartilha sessão entre eles sozinho. Duas opções:

- **Opção A (implementada aqui, zero mudanças no MiContas):** o usuário
  digita CPF/senha de novo ao entrar no Cuide, mas contra a mesma base de
  dados do MiContas — então é a mesma senha, só não é "automático".
- **Opção B (login silencioso, precisa de 1 pequena adição no MiContas):**
  ao fazer login no MiContas, gravar um cookie de sessão de curta duração
  com `domain=.apoiare.com.br`. O Cuide leria esse cookie e pularia a tela
  de login. Não implementei isso porque você pediu para não mexer no
  MiContas — mas é uma mudança pequena e localizada, se topar fazer depois.

## Próximo passo: ligar o financeiro do Cuide de verdade

O `Cuide.jsx` que você me mandou tem a função `registrarCaixa` dentro de
`SessaoModal` (por volta da linha 976) que hoje só grava no estado local
(`setCaixa`). Quando você quiser, eu adapto essa função para também chamar
`enviarEntradaCaixa(contaId, dadosSessao)` — aí o pagamento aparece tanto
no Cuide quanto nas Movimentações do MiContas, no mesmo lançamento. O
mesmo vale para o `RecebimentosPanel` (repasses) e o `FinanceiroPage`
(faturamento de convênio). Prefiro fazer essa parte só quando você
confirmar os nomes reais dos campos no banco do MiContas, pra não
escrever dado no lugar errado.

## Estrutura

```
src/
  config/firebase.js         → conexão com o Firebase do MiContas
  services/micontasApi.js    → única porta de entrada/saída de dados do MiContas
  auth/LoginSSO.jsx          → tela de login (mesma base de usuários)
  auth/RequireCuideAccess.jsx→ bloqueia quem não tem o módulo liberado
  components/AppSwitcher.jsx → barra para trocar entre Cuide e MiContas
  App.jsx                    → junta tudo e renderiza o Cuide.jsx original
```
