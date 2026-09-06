# Gringo V1 — automação de lead simples + RPA + dashboard

Este pacote entrega a V1 combinada nesta conversa:

- 1 script único do Apps Script (substitui os 4 scripts antigos) rodando dentro da
  planilha **Gringo_Automação** que você criou.
- 1 robô de RPA (Playwright) que abre o anúncio certo e preenche Nome/E-mail/Telefone,
  rodando de graça no GitHub Actions, sem servidor pra manter.
- 1 dashboard estático pra GitHub Pages, lendo a mesma planilha.

Existe um ponto em aberto de propósito: leads **sem** código continuam só indo pro
status `Sem_Codigo` (não entram no RPA). Combinamos deixar isso hoje reservado pra
depois decidirmos se compensa custo colocar IA tentando achar o anúncio mais parecido.

## Passo a passo pra colocar no ar

### 1. Apps Script (dentro da planilha Gringo_Automação)

1. Abra a planilha > Extensions > Apps Script.
2. Cole o conteúdo de `apps_script/Code.gs` (substitua o `Code.gs` padrão).
3. Rode a função `configurarPlanilhaV1` uma vez (autorize as permissões pedidas pelo Google).
   Isso cria a aba **Entrada** e a aba **Config** (com um token de segurança gerado
   automaticamente).
4. Confirme que a aba **Estoque_base_Dados** que você já criou tem exatamente estes
   cabeçalhos na linha 1 (nessa ordem):
   `Codigo_Anuncio | Marca | Modelo | Ano | Link_Anuncio`
5. Rode a função `instalarGatilho` uma vez. Isso cria o gatilho que processa
   automaticamente cada linha nova (usei `onChange` instalável, não `onEdit` simples,
   porque o SendPulse insere linha via API e o `onEdit` não pega isso).
6. Deploy > Nova implantação > tipo "Aplicativo da Web". Executar como **Eu**, acesso
   **Qualquer pessoa**. Copie a URL gerada.
7. Na aba **Config**, copie o valor de `TOKEN`.

### 2. Apontar o SendPulse pra planilha nova

No fluxo do SendPulse, troque o destino dos nós **"Inserir linha do Google Planilhas"**
pra planilha `Gringo_Automação`, aba `Entrada`, mantendo a mesma ordem de campos que
já existe hoje: Nome, Telefone, Data, Data/Hora, Mensagem. Não precisa mexer no resto
do fluxo (o Agente de IA, os filtros, etc. continuam como estão).

### 3. GitHub — repositório do robô + dashboard

1. Crie um repositório novo no GitHub (pode ser privado) e suba todo o conteúdo desta
   pasta (`git init`, `git add .`, `git commit`, `git remote add origin ...`, `git push`).
2. Em Settings > Secrets and variables > Actions, crie dois secrets:
   - `SHEET_WEBAPP_URL` — a URL do Aplicativo da Web (passo 1.6 acima)
   - `SHEET_TOKEN` — o token da aba Config (passo 1.7 acima)
3. Em Settings > Pages, ative o GitHub Pages apontando pra pasta `/docs` na branch
   principal. Isso publica o dashboard.
4. O dashboard **não lê a planilha publicada como CSV** — isso exporia nome/e-mail/
   telefone dos clientes pra qualquer um com o link, pra sempre, mesmo escondendo a
   aba depois. Em vez disso, ele chama um endpoint novo do próprio Web App
   (`action=getStats`) que devolve só números agregados, usando um token separado
   (`DASH_TOKEN`, de baixo privilégio — só lê estatística, não dá acesso a dado de
   cliente nem às ações do robô). Depois de colar a versão atualizada do
   `apps_script/Code.gs` (que já inclui `getStats`) e rodar `configurarPlanilhaV1`
   de novo, pegue o `DASH_TOKEN` na aba Config e cole em `docs/index.html`, na
   constante `DASH_TOKEN`. Redeploy do Web App (mesma implantação, "Nova versão").

### 4. Validar o robô antes de ligar de vez

O formulário de contato do anúncio eu só consegui inspecionar visualmente pelas suas
telas (campos `#whatsapp-name`, `#whatsapp-email`, `#whatsapp-telefone` e um botão
"Chamar no Whatsapp") — não testei o envio de verdade. Antes de deixar rodando sozinho:

1. Vá em Actions > "Gringo RPA - preencher leads" > Run workflow > marque
   `dry_run = true`. Isso preenche o formulário mas não clica em enviar — só confirma
   que os seletores batem com o site.
2. Depois de validar visualmente (pode rodar localmente com `HEADLESS=false` se quiser
   ver a tela), rode de novo com `dry_run = false` numa linha de teste primeiro.
3. Só depois disso, deixe o agendamento automático (`cron`, a cada 15 min) correndo sozinho.

## O que ficou simplificado em relação ao script antigo

- Acabou o "banco de veículos" de 20 marcas fixas tentando adivinhar o veículo por
  regex — agora o código do anúncio (quando existe) vai direto pro PROCV na aba de
  estoque, que é 100% preciso e não precisa de manutenção de lista.
- Uma função só processa cada linha (antes eram 3 versões de `processarLinha`
  competindo, além do `puxarClientesComCodigo2` isolado).
- Deduplicação num lugar só (antes checava a planilha local E uma planilha CRM externa;
  agora o protocolo mora só na aba Entrada).
- O robô de preenchimento no site da loja parceira, que antes era manual, passa a
  rodar sozinho — isso é o que fecha o requisito "atendimento 100% IA / sem intervenção
  manual" da V1.

## Pendências que ficaram fora do escopo de hoje (V2/V3, conforme combinamos)

- Simulação de pré-aprovação C6 (V2).
- Integração da API Credicarro no WhatsApp (V3).
- IA tentando achar o anúncio certo pros leads sem código (decisão futura, considerando custo).
