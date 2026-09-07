/**
 * Relatório diário — Gringo V1
 *
 * Roda periodicamente (a cada 15 min) no GitHub Actions, mas só REALMENTE
 * gera e envia o relatório quando:
 *   1. já passou do horário configurado na aba Config (chave Horario_Relatorio,
 *      ex.: "08:00" — você pode mudar isso a qualquer momento direto na
 *      planilha, sem precisar mexer em código nem no GitHub);
 *   2. ainda não foi enviado hoje (evita mandar duas vezes seguidas, mesmo
 *      rodando de 15 em 15 min);
 *   3. existe pelo menos um e-mail com Ativo=TRUE na aba Relatorio_Destinatarios.
 *
 * O relatório traz dois blocos, sempre a partir do mesmo `detalhe` (sem PII)
 * que já alimenta o dashboard: um resumo do MÊS ATUAL (até ontem) e uma visão
 * só do DIA ANTERIOR.
 *
 * Variáveis de ambiente esperadas (Secrets do GitHub Actions):
 *   SHEET_WEBAPP_URL    -> URL do deploy "Aplicativo da Web" do Apps Script
 *   SHEET_TOKEN         -> Token de privilégio do robô (aba Config, TOKEN)
 *   DASH_TOKEN          -> Token de estatísticas agregadas (aba Config, DASH_TOKEN)
 *   GMAIL_USER          -> e-mail que envia (ex.: daniel.silva@usadosbr.com)
 *   GMAIL_APP_PASSWORD  -> senha de app gerada na conta Google (não é a senha normal)
 */

const { chromium } = require("playwright");
const nodemailer = require("nodemailer");

const WEBAPP_URL = process.env.SHEET_WEBAPP_URL;
const SHEET_TOKEN = process.env.SHEET_TOKEN;
const DASH_TOKEN = process.env.DASH_TOKEN;
const GMAIL_USER = process.env.GMAIL_USER;
const GMAIL_APP_PASSWORD = process.env.GMAIL_APP_PASSWORD;

const TIMEZONE = "America/Sao_Paulo";

if (!WEBAPP_URL || !SHEET_TOKEN || !DASH_TOKEN || !GMAIL_USER || !GMAIL_APP_PASSWORD) {
  console.error(
    "Faltam variáveis de ambiente: confirme SHEET_WEBAPP_URL, SHEET_TOKEN, DASH_TOKEN, " +
    "GMAIL_USER e GMAIL_APP_PASSWORD nos Secrets do GitHub Actions."
  );
  process.exit(1);
}

function hojeNoFuso() {
  // yyyy-MM-dd no fuso de Brasília, sem depender do fuso do runner do GitHub (UTC).
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: TIMEZONE, year: "numeric", month: "2-digit", day: "2-digit",
  }).format(new Date());
}

function agoraHHMMNoFuso() {
  return new Intl.DateTimeFormat("en-GB", {
    timeZone: TIMEZONE, hour: "2-digit", minute: "2-digit", hour12: false,
  }).format(new Date());
}

function ontemNoFuso() {
  const hoje = hojeNoFuso();
  const d = new Date(hoje + "T12:00:00Z"); // meio-dia UTC evita virar de dia por causa do fuso
  d.setUTCDate(d.getUTCDate() - 1);
  return d.toISOString().slice(0, 10);
}

function minutosDoDia(hhmm) {
  const partes = String(hhmm || "00:00").split(":").map(Number);
  return (partes[0] || 0) * 60 + (partes[1] || 0);
}

async function chamarWebApp(params) {
  const url = `${WEBAPP_URL}?${new URLSearchParams(params).toString()}`;
  const resp = await fetch(url);
  return resp.json();
}

async function postWebApp(params) {
  const url = `${WEBAPP_URL}?${new URLSearchParams(params).toString()}`;
  await fetch(url, { method: "POST" });
}

function formatarMoeda(v) {
  return (Number(v) || 0).toLocaleString("pt-BR", { maximumFractionDigits: 0 });
}

function topN(obj, n) {
  return Object.entries(obj).sort((a, b) => b[1] - a[1]).slice(0, n);
}

function agregarLinhas(linhas) {
  const total = linhas.length;
  const comCodigo = linhas.filter((r) => r.temCodigo).length;
  const valorTotal = linhas.reduce((s, r) => s + (Number(r.valorAnuncio) || 0), 0);
  const porLoja = {};
  const porCidade = {};
  linhas.forEach((r) => {
    if (r.nomeRevenda) porLoja[r.nomeRevenda] = (porLoja[r.nomeRevenda] || 0) + 1;
    if (r.cidade) porCidade[r.cidade] = (porCidade[r.cidade] || 0) + 1;
  });
  return {
    total,
    comCodigo,
    pctMatch: total ? Math.round((comCodigo / total) * 100) : 0,
    valorTotal,
    topLojas: topN(porLoja, 5),
    topCidades: topN(porCidade, 5),
  };
}

function linhasTabela(pares, rotuloVazio) {
  if (!pares.length) {
    return `<tr><td colspan="2" style="padding:5px 8px;color:#7f9db4;">${rotuloVazio}</td></tr>`;
  }
  return pares
    .map(([nome, contagem]) => `
      <tr>
        <td style="padding:5px 8px;border-bottom:1px solid rgba(255,255,255,.08);">${nome}</td>
        <td style="padding:5px 8px;text-align:right;border-bottom:1px solid rgba(255,255,255,.08);">${contagem}</td>
      </tr>`)
    .join("");
}

function blocoResumo(titulo, r) {
  return `
    <h2 style="margin:26px 0 10px;font-size:16px;color:#eaf3fb;">${titulo}</h2>
    <table style="border-collapse:collapse;margin-bottom:14px;">
      <tr><td style="padding:4px 10px 4px 0;color:#7f9db4;font-size:13px;">Total de leads</td>
          <td style="padding:4px 0;font-weight:700;font-size:15px;">${r.total}</td></tr>
      <tr><td style="padding:4px 10px 4px 0;color:#7f9db4;font-size:13px;">Com match (código bateu)</td>
          <td style="padding:4px 0;font-weight:700;font-size:15px;color:#39d97a;">${r.comCodigo} (${r.pctMatch}%)</td></tr>
      <tr><td style="padding:4px 10px 4px 0;color:#7f9db4;font-size:13px;">Valor total em anúncios</td>
          <td style="padding:4px 0;font-weight:700;font-size:15px;color:#ffd500;">R$ ${formatarMoeda(r.valorTotal)}</td></tr>
    </table>
    <div style="display:flex;gap:28px;flex-wrap:wrap;">
      <table style="border-collapse:collapse;min-width:220px;">
        <tr><th style="text-align:left;padding:4px 8px;color:#7f9db4;font-size:11.5px;text-transform:uppercase;">Top lojas</th><th></th></tr>
        ${linhasTabela(r.topLojas, "Sem dados de loja ainda")}
      </table>
      <table style="border-collapse:collapse;min-width:220px;">
        <tr><th style="text-align:left;padding:4px 8px;color:#7f9db4;font-size:11.5px;text-transform:uppercase;">Top cidades</th><th></th></tr>
        ${linhasTabela(r.topCidades, "Sem dados de cidade ainda")}
      </table>
    </div>`;
}

function montarHtmlRelatorio({ dataRelatorio, mesLabel, diaLabel, mes, dia }) {
  return `<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><style>
  body { background:#071726; color:#eaf3fb; font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif; margin:0; padding:30px; }
  .topo { background:linear-gradient(90deg,#004e87,#39d97a,#ffd500); height:6px; margin:-30px -30px 22px; }
  h1 { font-size:21px; margin:0 0 2px; }
  .marca { color:#ffd500; font-weight:800; } .marca2 { color:#39d97a; font-weight:800; }
  .sub { color:#7f9db4; font-size:12.5px; margin-bottom:6px; }
</style></head>
<body>
  <div class="topo"></div>
  <h1><span class="marca">GRINGO</span> × <span class="marca2">usadosbr</span> — Relatório diário</h1>
  <div class="sub">Gerado em ${dataRelatorio}</div>
  ${blocoResumo("Resumo do mês atual (" + mesLabel + ")", mes)}
  ${blocoResumo("Visão do dia anterior (" + diaLabel + ")", dia)}
</body></html>`;
}

async function gerarPdf(html) {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: "networkidle" });
  const pdf = await page.pdf({
    format: "A4",
    printBackground: true,
    margin: { top: "18px", bottom: "18px", left: "18px", right: "18px" },
  });
  await browser.close();
  return pdf;
}

async function enviarEmail(destinatarios, pdfBuffer, dataRotulo) {
  const transporter = nodemailer.createTransport({
    host: "smtp.gmail.com",
    port: 465,
    secure: true,
    auth: { user: GMAIL_USER, pass: GMAIL_APP_PASSWORD },
  });
  await transporter.sendMail({
    from: `"Gringo × usadosbr" <${GMAIL_USER}>`,
    to: destinatarios.join(","),
    subject: `Relatório diário Gringo × usadosbr — ${dataRotulo}`,
    text: "Segue em anexo o relatório diário (resumo do mês atual e visão do dia anterior). Este e-mail é gerado automaticamente.",
    attachments: [{ filename: `relatorio-gringo-${dataRotulo}.pdf`, content: pdfBuffer }],
  });
}

async function main() {
  const config = await chamarWebApp({ action: "getConfigRelatorio", token: SHEET_TOKEN });
  if (config.erro) throw new Error("Web App (getConfigRelatorio): " + config.erro);

  if (!config.destinatarios || config.destinatarios.length === 0) {
    console.log("Nenhum destinatário ativo em Relatorio_Destinatarios. Nada a fazer.");
    return;
  }

  const hoje = hojeNoFuso();
  if (config.ultimoEnvio === hoje) {
    console.log(`Relatório de hoje (${hoje}) já foi enviado. Nada a fazer.`);
    return;
  }

  const agora = agoraHHMMNoFuso();
  if (minutosDoDia(agora) < minutosDoDia(config.horario)) {
    console.log(`Ainda não chegou o horário configurado (${config.horario}); agora são ${agora} (horário de Brasília).`);
    return;
  }

  console.log(`Horário batido (configurado: ${config.horario}, agora: ${agora}). Gerando relatório...`);

  const stats = await chamarWebApp({ action: "getStats", token: DASH_TOKEN });
  if (stats.erro) throw new Error("Web App (getStats): " + stats.erro);

  const detalhe = stats.detalhe || [];
  const ontem = ontemNoFuso();
  const mesAtual = hoje.slice(0, 7); // yyyy-MM

  const doMes = detalhe.filter((r) => String(r.data || "").slice(0, 7) === mesAtual);
  const doDia = detalhe.filter((r) => r.data === ontem);

  const html = montarHtmlRelatorio({
    dataRelatorio: new Date().toLocaleString("pt-BR", { timeZone: TIMEZONE }),
    mesLabel: mesAtual,
    diaLabel: ontem,
    mes: agregarLinhas(doMes),
    dia: agregarLinhas(doDia),
  });

  const pdf = await gerarPdf(html);
  await enviarEmail(config.destinatarios, pdf, hoje);
  await postWebApp({ action: "marcarRelatorioEnviado", token: SHEET_TOKEN });

  console.log(`Relatório enviado para: ${config.destinatarios.join(", ")}`);
}

main().catch((err) => {
  console.error("Falha ao gerar/enviar o relatório:", err);
  process.exit(1);
});
