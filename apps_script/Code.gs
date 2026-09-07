/**
 * GRINGO V1 — script único da planilha "Gringo_Automação"
 * Substitui os 4 scripts antigos (puxarClientesComCodigo2, processarLinha x3, onEdit/onChange soltos).
 *
 * O que ele faz:
 *  1. Cria a estrutura da planilha (aba Entrada + aba Config) na primeira execução.
 *  2. A cada linha nova em "Entrada" (inserida pelo SendPulse), extrai Nome/E-mail/CPF/Código
 *     do texto bruto da mensagem.
 *  3. Se veio "código: NNNNNN" na mensagem, procura esse código na aba "Estoque_base_Dados"
 *     (equivalente a um PROCV) e traz Marca/Modelo/Ano/Link/Loja/Cidade/Valor etc. do anúncio.
 *  4. Marca o status da linha: Pendente_RPA | Codigo_Sem_Match | Sem_Codigo.
 *  5. Expõe um Web App (doGet/doPost) para o robô de RPA (rodando no GitHub Actions)
 *     buscar os leads pendentes e, depois de preencher o formulário do anúncio,
 *     avisar de volta que foi enviado — e também expõe estatísticas agregadas
 *     (sem PII) pro dashboard público.
 *
 * INSTALAÇÃO (uma vez só):
 *  a) Extensions > Apps Script, na planilha Gringo_Automação. Cole este arquivo substituindo o Code.gs padrão.
 *  b) Rode a função `configurarPlanilhaV1` uma vez (autorize as permissões pedidas).
 *     Isso cria a aba "Entrada" e a aba "Config" com um token de segurança gerado automaticamente.
 *  c) Se a aba "Estoque_base_Dados" ou "Entrada" já existiam de antes (cabeçalho antigo, com
 *     menos colunas), rode `atualizarCabecalhoEstoque` e `atualizarCabecalhoEntrada` uma vez —
 *     elas só reescrevem a linha 1 (cabeçalho), nunca tocam nas linhas de dados já existentes.
 *  d) Rode `instalarGatilho` uma vez (cria o gatilho onChange instalável — necessário porque
 *     o SendPulse insere linhas via API, e o onEdit simples não dispara nesse caso).
 *  e) Deploy > Nova implantação > Aplicativo da Web. Executar como "Eu", Quem pode acessar "Qualquer pessoa".
 *     Copie a URL gerada — ela vai para o GitHub Actions (segredo SHEET_WEBAPP_URL).
 *  f) Na aba "Config", copie o valor de TOKEN — vai para o GitHub Actions (segredo SHEET_TOKEN).
 *  g) No SendPulse, aponte o(s) nó(s) "Inserir linha do Google Planilhas" para esta planilha,
 *     aba "Entrada", preenchendo nas colunas: Nome, Telefone, Data, Data_Hora, Mensagem
 *     (nessa ordem — é a mesma ordem que o fluxo atual já usa, só troca o destino).
 */

const ABA_ENTRADA = "Entrada";
const ABA_ESTOQUE = "Estoque_base_Dados";
const ABA_CONFIG = "Config";

// As colunas novas (17-24) foram sempre ACRESCENTADAS no final, nunca inseridas
// no meio — isso preserva as posições de Status/Protocolo/Data_Envio_RPA que já
// existiam antes, então nenhum dado antigo precisa ser movido de coluna.
const COL = {
  TIMESTAMP: 1,
  NOME: 2,
  TELEFONE: 3,
  DATA: 4,
  DATA_HORA: 5,
  MENSAGEM: 6,
  EMAIL: 7,
  CPF: 8,
  CODIGO_ANUNCIO: 9,
  MARCA: 10,
  MODELO: 11,
  ANO: 12,
  LINK_ANUNCIO: 13,
  STATUS: 14,
  PROTOCOLO: 15,
  DATA_ENVIO_RPA: 16,
  TIPO_BASE: 17,
  VERSAO: 18,
  TYPE: 19,
  ID_ADMIX: 20,
  NOME_REVENDA: 21,
  REV_DDD: 22,
  VALOR_ANUNCIO: 23,
  CIDADE: 24,
};

const CABECALHO_ENTRADA = [
  "Timestamp", "Nome", "Telefone", "Data", "Data_Hora", "Mensagem",
  "Email", "CPF", "Codigo_Anuncio", "Marca", "Modelo", "Ano",
  "Link_Anuncio", "Status", "Protocolo", "Data_Envio_RPA",
  "Tipo_Base", "Versao", "Type", "Id_Admix", "Nome_Revenda", "Rev_DDD",
  "Valor_Anuncio", "Cidade",
];

// Mesma lógica: Codigo_Anuncio, Marca, Modelo, Ano e Link_Anuncio já existiam
// e ficam intactos; os campos novos pedidos pro relatório (loja, cidade, valor,
// etc.) foram só acrescentados à direita.
const CABECALHO_ESTOQUE = [
  "Codigo_Anuncio", "Marca", "Modelo", "Ano", "Link_Anuncio",
  "Tipo_Base", "Versao", "Type", "Id_Admix", "Nome_Revenda", "Rev_DDD",
  "Valor_Anuncio", "Cidade",
];

const STATUS_PENDENTE = "Pendente_RPA";
const STATUS_SEM_MATCH = "Codigo_Sem_Match";
const STATUS_SEM_CODIGO = "Sem_Codigo";
const STATUS_ENVIADO = "Enviado";
const STATUS_ERRO = "Erro_RPA";

// ---------------------------------------------------------------------------
// SETUP — rodar uma vez manualmente
// ---------------------------------------------------------------------------

function configurarPlanilhaV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let entrada = ss.getSheetByName(ABA_ENTRADA);
  if (!entrada) {
    entrada = ss.insertSheet(ABA_ENTRADA);
  }
  if (entrada.getLastRow() === 0) {
    entrada.getRange(1, 1, 1, CABECALHO_ENTRADA.length).setValues([CABECALHO_ENTRADA]);
    entrada.setFrozenRows(1);
  }

  const estoque = ss.getSheetByName(ABA_ESTOQUE);
  if (!estoque) {
    const criado = ss.insertSheet(ABA_ESTOQUE);
    criado.getRange(1, 1, 1, CABECALHO_ESTOQUE.length).setValues([CABECALHO_ESTOQUE]);
    criado.setFrozenRows(1);
  }

  let config = ss.getSheetByName(ABA_CONFIG);
  if (!config) {
    config = ss.insertSheet(ABA_CONFIG);
  }
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty("TOKEN");
  if (!token) {
    token = Utilities.getUuid();
    props.setProperty("TOKEN", token);
  }
  // Token separado, só de leitura de estatísticas agregadas (sem PII), pra usar
  // no dashboard público do GitHub Pages sem expor dado de cliente nem dar
  // acesso às ações do robô de RPA.
  let dashToken = props.getProperty("DASH_TOKEN");
  if (!dashToken) {
    dashToken = Utilities.getUuid();
    props.setProperty("DASH_TOKEN", dashToken);
  }
  config.clear();
  config.getRange(1, 1, 4, 2).setValues([
    ["Chave", "Valor"],
    ["TOKEN", token],
    ["DASH_TOKEN", dashToken],
    ["Atualizado em", new Date()],
  ]);

  SpreadsheetApp.getUi().alert(
    "Planilha configurada.\n\nAba Entrada e Config prontas.\n" +
    "Confirme os cabeçalhos da aba Estoque_base_Dados e rode instalarGatilho() em seguida."
  );
}

// Reescreve SÓ a linha 1 (cabeçalho) da aba Entrada com as colunas atuais —
// não toca em nenhuma linha de dado já existente. Rode uma vez depois de subir
// esta versão do Code.gs, se a aba Entrada já existia com o cabeçalho antigo
// (16 colunas, sem os campos de loja/cidade/valor etc.).
function atualizarCabecalhoEntrada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entrada = ss.getSheetByName(ABA_ENTRADA);
  entrada.getRange(1, 1, 1, CABECALHO_ENTRADA.length).setValues([CABECALHO_ENTRADA]);
  SpreadsheetApp.getUi().alert(
    "Cabeçalho da aba Entrada atualizado (" + CABECALHO_ENTRADA.length + " colunas). " +
    "Nenhuma linha de dado foi alterada."
  );
}

// Mesma ideia, mas pra aba Estoque_base_Dados.
function atualizarCabecalhoEstoque() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const estoque = ss.getSheetByName(ABA_ESTOQUE);
  estoque.getRange(1, 1, 1, CABECALHO_ESTOQUE.length).setValues([CABECALHO_ESTOQUE]);
  SpreadsheetApp.getUi().alert(
    "Cabeçalho da aba Estoque_base_Dados atualizado (" + CABECALHO_ESTOQUE.length + " colunas). " +
    "Nenhuma linha de dado foi alterada."
  );
}

function instalarGatilho() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "aoMudarPlanilha") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("aoMudarPlanilha").forSpreadsheet(ss).onChange().create();
  SpreadsheetApp.getUi().alert("Gatilho onChange instalado.");
}

// ---------------------------------------------------------------------------
// PROCESSAMENTO
// ---------------------------------------------------------------------------

function aoMudarPlanilha(e) {
  processarPendentes();
}

// Roda manualmente ou via gatilho: processa todas as linhas da aba Entrada
// que ainda não têm Status preenchido.
function processarPendentes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entrada = ss.getSheetByName(ABA_ENTRADA);
  if (!entrada) return;

  const lastRow = entrada.getLastRow();
  if (lastRow < 2) return;

  const dados = entrada.getRange(2, 1, lastRow - 1, CABECALHO_ENTRADA.length).getValues();
  const estoqueMapa = carregarEstoque();

  const protocolosExistentes = {};
  dados.forEach(function (linha) {
    const p = linha[COL.PROTOCOLO - 1];
    if (p) protocolosExistentes[p] = true;
  });

  dados.forEach(function (linha, idx) {
    const numeroLinha = idx + 2;
    const statusAtual = linha[COL.STATUS - 1];
    if (statusAtual) return; // já processado

    processarLinha(entrada, numeroLinha, linha, estoqueMapa, protocolosExistentes);
  });
}

function carregarEstoque() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const estoque = ss.getSheetByName(ABA_ESTOQUE);
  const mapa = {};
  if (!estoque) return mapa;
  const lastRow = estoque.getLastRow();
  if (lastRow < 2) return mapa;
  const dados = estoque.getRange(2, 1, lastRow - 1, CABECALHO_ESTOQUE.length).getValues();
  dados.forEach(function (linha) {
    const codigo = String(linha[0]).trim();
    if (!codigo) return;
    mapa[codigo] = {
      marca: linha[1],
      modelo: linha[2],
      ano: linha[3],
      link: linha[4],
      tipoBase: linha[5],
      versao: linha[6],
      type: linha[7],
      idAdmix: linha[8],
      nomeRevenda: linha[9],
      revDDD: linha[10],
      valorAnuncio: linha[11],
      cidade: linha[12],
    };
  });
  return mapa;
}

function processarLinha(entrada, numeroLinha, linha, estoqueMapa, protocolosExistentes) {
  const mensagem = String(linha[COL.MENSAGEM - 1] || "");
  const telefone = String(linha[COL.TELEFONE - 1] || "");
  let nome = String(linha[COL.NOME - 1] || "");
  let email = String(linha[COL.EMAIL - 1] || "");
  let cpf = String(linha[COL.CPF - 1] || "");
  let codigo = String(linha[COL.CODIGO_ANUNCIO - 1] || "").trim();

  if (mensagem) {
    if (!nome) {
      const nomeMatch = mensagem.match(/Sou\s+([^,\.]+)/i);
      if (nomeMatch) {
        nome = nomeMatch[1].trim();
        if (nome.toLowerCase().indexOf(" de ") !== -1) {
          nome = nome.split(" de ")[0].trim();
        }
      }
    }
    if (!email) {
      const emailMatch = mensagem.match(/email:\s*([^\s,]+)/i);
      if (emailMatch) email = emailMatch[1].trim();
    }
    if (!cpf) {
      const cpfMatch = mensagem.match(/CPF:\s*([0-9]+)/i);
      if (cpfMatch) cpf = cpfMatch[1].trim();
    }
    if (!codigo) {
      const codigoMatch = mensagem.match(/c[oó]digo:?\s*(\d+)/i);
      if (codigoMatch) codigo = codigoMatch[1].trim();
    }
  }

  let marca = "", modelo = "", ano = "", link = "";
  let tipoBase = "", versao = "", type = "", idAdmix = "";
  let nomeRevenda = "", revDDD = "", valorAnuncio = "", cidade = "";
  let status;

  if (codigo && estoqueMapa[codigo]) {
    const info = estoqueMapa[codigo];
    marca = info.marca; modelo = info.modelo; ano = info.ano; link = info.link;
    tipoBase = info.tipoBase; versao = info.versao; type = info.type; idAdmix = info.idAdmix;
    nomeRevenda = info.nomeRevenda; revDDD = info.revDDD; valorAnuncio = info.valorAnuncio; cidade = info.cidade;
    status = STATUS_PENDENTE;
  } else if (codigo) {
    status = STATUS_SEM_MATCH; // veio com código mas não achamos no estoque
  } else {
    status = STATUS_SEM_CODIGO; // sem código na mensagem — vai só pro atendimento IA / dashboard
  }

  const telLimpo = telefone.replace(/\D/g, "");
  const tel8 = telLimpo.slice(-8);
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const protocolo = tel8 + (codigo || "SC") + hoje;

  if (protocolosExistentes[protocolo]) {
    status = STATUS_ENVIADO; // já existe um igual processado antes — evita reenviar
  } else {
    protocolosExistentes[protocolo] = true;
  }

  entrada.getRange(numeroLinha, COL.NOME).setValue(nome);
  entrada.getRange(numeroLinha, COL.EMAIL).setValue(email);
  entrada.getRange(numeroLinha, COL.CPF).setValue(cpf);
  entrada.getRange(numeroLinha, COL.CODIGO_ANUNCIO).setValue(codigo);
  entrada.getRange(numeroLinha, COL.MARCA).setValue(marca);
  entrada.getRange(numeroLinha, COL.MODELO).setValue(modelo);
  entrada.getRange(numeroLinha, COL.ANO).setValue(ano);
  entrada.getRange(numeroLinha, COL.LINK_ANUNCIO).setValue(link);
  entrada.getRange(numeroLinha, COL.STATUS).setValue(status);
  entrada.getRange(numeroLinha, COL.PROTOCOLO).setValue(protocolo);
  entrada.getRange(numeroLinha, COL.TIPO_BASE).setValue(tipoBase);
  entrada.getRange(numeroLinha, COL.VERSAO).setValue(versao);
  entrada.getRange(numeroLinha, COL.TYPE).setValue(type);
  entrada.getRange(numeroLinha, COL.ID_ADMIX).setValue(idAdmix);
  entrada.getRange(numeroLinha, COL.NOME_REVENDA).setValue(nomeRevenda);
  entrada.getRange(numeroLinha, COL.REV_DDD).setValue(revDDD);
  entrada.getRange(numeroLinha, COL.VALOR_ANUNCIO).setValue(valorAnuncio);
  entrada.getRange(numeroLinha, COL.CIDADE).setValue(cidade);
}

// ---------------------------------------------------------------------------
// WEB APP — usado pelo robô de RPA (GitHub Actions) e pelo dashboard
// ---------------------------------------------------------------------------

function checarToken(e) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("TOKEN");
  return e.parameter.token && e.parameter.token === token;
}

function checarTokenDashboard(e) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("DASH_TOKEN");
  return e.parameter.token && e.parameter.token === token;
}

function doGet(e) {
  const acao = e.parameter.action;

  // getStats usa um token separado (DASH_TOKEN), de baixo privilégio: só
  // números agregados e uma lista de leads SEM PII (sem nome/e-mail/telefone/
  // protocolo). É o único endpoint seguro pra chamar de dentro do dashboard
  // público (o código-fonte da página é público).
  if (acao === "getStats") {
    if (!checarTokenDashboard(e)) {
      return ContentService.createTextOutput(JSON.stringify({ erro: "token inválido" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify(getStats()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (!checarToken(e)) {
    return ContentService.createTextOutput(JSON.stringify({ erro: "token inválido" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (acao === "getPendentes") {
    return ContentService.createTextOutput(JSON.stringify(getPendentes()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ erro: "ação desconhecida" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Agrupa o valor do anúncio numa faixa de preço, pra montar o "perfil de valor"
// pedido no dashboard sem precisar expor o valor exato lead a lead (embora o
// valor exato também vá no array `detalhe`, que já não tem nenhum dado pessoal
// do cliente — só do anúncio/loja).
function faixaDeValor(valor) {
  const v = Number(valor) || 0;
  if (!v) return "Sem valor";
  if (v < 30000) return "Até 30 mil";
  if (v < 60000) return "30 a 60 mil";
  if (v < 100000) return "60 a 100 mil";
  if (v < 150000) return "100 a 150 mil";
  return "Acima de 150 mil";
}

// Estatísticas agregadas pro dashboard público — nunca inclui nome, e-mail,
// telefone, CPF ou protocolo. O array `detalhe` traz um registro por lead só
// com dados do ANÚNCIO/LOJA (marca, modelo, loja, cidade, valor, status, dia),
// pra alimentar os filtros e a exportação de CSV direto no navegador de quem
// abre o dashboard, sem precisar expor dado de cliente.
function getStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entrada = ss.getSheetByName(ABA_ENTRADA);
  const lastRow = entrada.getLastRow();
  if (lastRow < 2) {
    return {
      total: 0, comCodigo: 0, porStatus: {}, porDia: {},
      porLoja: {}, porCidade: {}, porFaixaValor: {}, detalhe: [],
    };
  }

  const dados = entrada.getRange(2, 1, lastRow - 1, CABECALHO_ENTRADA.length).getValues();
  const porStatus = {};
  const porDia = {};
  const porLoja = {};
  const porCidade = {};
  const porFaixaValor = {};
  const detalhe = [];
  let comCodigo = 0;

  dados.forEach(function (linha) {
    const codigo = linha[COL.CODIGO_ANUNCIO - 1];
    if (codigo) comCodigo++;

    const status = linha[COL.STATUS - 1] || "Sem status";
    porStatus[status] = (porStatus[status] || 0) + 1;

    const data = linha[COL.DATA - 1];
    let diaChave = "";
    if (data instanceof Date) {
      diaChave = Utilities.formatDate(data, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (data) {
      diaChave = String(data).slice(0, 10);
    }
    if (diaChave) porDia[diaChave] = (porDia[diaChave] || 0) + 1;

    const loja = String(linha[COL.NOME_REVENDA - 1] || "").trim();
    if (loja) porLoja[loja] = (porLoja[loja] || 0) + 1;

    const cidade = String(linha[COL.CIDADE - 1] || "").trim();
    if (cidade) porCidade[cidade] = (porCidade[cidade] || 0) + 1;

    const valor = linha[COL.VALOR_ANUNCIO - 1];
    const faixa = faixaDeValor(valor);
    porFaixaValor[faixa] = (porFaixaValor[faixa] || 0) + 1;

    detalhe.push({
      data: diaChave,
      status: status,
      temCodigo: !!codigo,
      marca: linha[COL.MARCA - 1] || "",
      modelo: linha[COL.MODELO - 1] || "",
      ano: linha[COL.ANO - 1] || "",
      tipoBase: linha[COL.TIPO_BASE - 1] || "",
      versao: linha[COL.VERSAO - 1] || "",
      type: linha[COL.TYPE - 1] || "",
      nomeRevenda: loja,
      revDDD: linha[COL.REV_DDD - 1] || "",
      valorAnuncio: valor || "",
      cidade: cidade,
    });
  });

  return {
    total: dados.length,
    comCodigo: comCodigo,
    porStatus: porStatus,
    porDia: porDia,
    porLoja: porLoja,
    porCidade: porCidade,
    porFaixaValor: porFaixaValor,
    detalhe: detalhe,
    atualizadoEm: new Date().toISOString(),
  };
}

function doPost(e) {
  if (!checarToken(e)) {
    return ContentService.createTextOutput(JSON.stringify({ erro: "token inválido" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const acao = e.parameter.action;
  const protocolo = e.parameter.protocolo;

  if (acao === "marcarEnviado") {
    marcarStatus(protocolo, STATUS_ENVIADO);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (acao === "marcarErro") {
    marcarStatus(protocolo, STATUS_ERRO);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ erro: "ação desconhecida" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getPendentes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entrada = ss.getSheetByName(ABA_ENTRADA);
  const lastRow = entrada.getLastRow();
  if (lastRow < 2) return [];

  const dados = entrada.getRange(2, 1, lastRow - 1, CABECALHO_ENTRADA.length).getValues();
  const resultado = [];
  dados.forEach(function (linha) {
    if (linha[COL.STATUS - 1] !== STATUS_PENDENTE) return;
    resultado.push({
      protocolo: linha[COL.PROTOCOLO - 1],
      nome: linha[COL.NOME - 1],
      email: linha[COL.EMAIL - 1],
      telefone: linha[COL.TELEFONE - 1],
      link: linha[COL.LINK_ANUNCIO - 1],
    });
  });
  return resultado;
}

function marcarStatus(protocolo, status) {
  if (!protocolo) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entrada = ss.getSheetByName(ABA_ENTRADA);
  const lastRow = entrada.getLastRow();
  if (lastRow < 2) return;
  const protocolos = entrada.getRange(2, COL.PROTOCOLO, lastRow - 1, 1).getValues();
  for (let i = 0; i < protocolos.length; i++) {
    if (String(protocolos[i][0]).trim() === String(protocolo).trim()) {
      const numeroLinha = i + 2;
      entrada.getRange(numeroLinha, COL.STATUS).setValue(status);
      entrada.getRange(numeroLinha, COL.DATA_ENVIO_RPA).setValue(new Date());
      return;
    }
  }
}

// ---------------------------------------------------------------------------
// UTILITÁRIO DE TESTE — rodar manualmente quando quiser reprocessar uma linha
// da aba Entrada do zero (limpa Status/Protocolo/Data_Envio_RPA e corrige o
// e-mail, depois chama processarPendentes de novo). Existe só pra não
// depender de editar célula por célula na mão, o que já causou linha
// desalinhada/e-mail trocado por telefone em testes anteriores.
// Ajuste NUMERO_LINHA e EMAIL_CORRETO antes de rodar, se precisar.
// ---------------------------------------------------------------------------
function repararLinhaTeste() {
  const NUMERO_LINHA = 3; // linha 3 da aba Entrada (a que tem o protocolo 48308713TESTE00120260906)
  const EMAIL_CORRETO = "daniel.silva@usadosbr.com";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entrada = ss.getSheetByName(ABA_ENTRADA);
  entrada.getRange(NUMERO_LINHA, COL.EMAIL).setValue(EMAIL_CORRETO);
  entrada.getRange(NUMERO_LINHA, COL.STATUS).clearContent();
  entrada.getRange(NUMERO_LINHA, COL.PROTOCOLO).clearContent();
  entrada.getRange(NUMERO_LINHA, COL.DATA_ENVIO_RPA).clearContent();

  processarPendentes();

  // O alert() só funciona quando a Sheets tem uma UI ativa na hora da chamada
  // (varia conforme a forma que você roda a função pelo editor). Envolvemos
  // em try/catch pra nunca aparecer como "erro" no registro — o que importa
  // (corrigir e reprocessar a linha) já rodou nas linhas acima.
  try {
    SpreadsheetApp.getUi().alert(
      "Linha " + NUMERO_LINHA + " corrigida (e-mail = " + EMAIL_CORRETO + ") e reprocessada.\n" +
      "Confira o novo Status na coluna N."
    );
  } catch (err) {
    Logger.log("Linha " + NUMERO_LINHA + " corrigida e reprocessada (alerta de UI indisponível neste contexto).");
  }
}
