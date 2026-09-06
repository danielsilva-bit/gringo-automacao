/**
 * Robô de RPA — Gringo V1
 *
 * Busca os leads "Pendente_RPA" na planilha (via Web App do Apps Script),
 * abre o link do anúncio no usadosbr.com, preenche Nome/E-mail/Telefone
 * no formulário de contato e envia — registrando o lead direto na loja
 * parceira, sem intervenção manual.
 *
 * Variáveis de ambiente esperadas (configuradas como Secrets no GitHub Actions):
 *   SHEET_WEBAPP_URL  -> URL do deploy "Aplicativo da Web" do Apps Script
 *   SHEET_TOKEN       -> Token gerado por configurarPlanilhaV1() (aba Config)
 *   DRY_RUN           -> "true" para só simular (não clica em enviar) — use isso
 *                        na primeira rodada pra validar os seletores do formulário
 *                        antes de ligar de vez a automação.
 */

const { chromium } = require("playwright");

const WEBAPP_URL = process.env.SHEET_WEBAPP_URL;
const TOKEN = process.env.SHEET_TOKEN;
const DRY_RUN = String(process.env.DRY_RUN || "false").toLowerCase() === "true";

if (!WEBAPP_URL || !TOKEN) {
  console.error("Faltam SHEET_WEBAPP_URL e/ou SHEET_TOKEN nas variáveis de ambiente.");
  process.exit(1);
}

async function buscarPendentes() {
  const url = `${WEBAPP_URL}?action=getPendentes&token=${encodeURIComponent(TOKEN)}`;
  const resp = await fetch(url);
  const json = await resp.json();
  if (json.erro) throw new Error(`Web App retornou erro: ${json.erro}`);
  return json;
}

async function marcarStatus(acao, protocolo) {
  const url = `${WEBAPP_URL}?action=${acao}&protocolo=${encodeURIComponent(protocolo)}&token=${encodeURIComponent(TOKEN)}`;
  await fetch(url, { method: "POST" });
}

// O campo do site espera DDD+número (ex.: 11948308713, 11 dígitos), sem o
// código do país. Os leads chegam do WhatsApp com "55" na frente
// (ex.: 5511948308713) — tira esse prefixo antes de preencher.
function formatarTelefoneParaSite(telefone) {
  let numeros = String(telefone || "").replace(/\D/g, "");
  if (numeros.length >= 12 && numeros.startsWith("55")) {
    numeros = numeros.slice(2);
  }
  return numeros;
}

async function preencherFormulario(page, lead) {
  // Seletores conforme inspecionados na página de anúncio do usadosbr.com.
  // Se o site mudar o HTML, ajuste aqui.
  await page.goto(lead.link, { waitUntil: "domcontentloaded", timeout: 30000 });

  await page.waitForSelector("#whatsapp-name", { timeout: 15000 });
  await page.fill("#whatsapp-name", lead.nome || "");
  await page.fill("#whatsapp-email", lead.email || "");
  await page.fill("#whatsapp-telefone", formatarTelefoneParaSite(lead.telefone));

  if (DRY_RUN) {
    console.log(`[DRY_RUN] Formulário preenchido para protocolo ${lead.protocolo}, não enviei.`);
    return;
  }

  // Botão de envio observado no formulário ("Chamar no Whatsapp").
  // Validar manualmente na primeira execução se este é de fato o botão que
  // registra o lead do lado da loja (e não só um link direto pro WhatsApp).
  const botao = page.locator("button:has-text('Chamar no Whatsapp')").first();
  await botao.click();
  await page.waitForTimeout(2000);
}

async function main() {
  const pendentes = await buscarPendentes();
  console.log(`Leads pendentes: ${pendentes.length}`);
  if (pendentes.length === 0) return;

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext();

  for (const lead of pendentes) {
    const page = await context.newPage();
    try {
      console.log(`Processando protocolo ${lead.protocolo} -> ${lead.link}`);
      await preencherFormulario(page, lead);
      if (!DRY_RUN) {
        await marcarStatus("marcarEnviado", lead.protocolo);
        console.log(`Protocolo ${lead.protocolo} marcado como Enviado.`);
      }
    } catch (err) {
      console.error(`Erro no protocolo ${lead.protocolo}:`, err.message);
      if (!DRY_RUN) {
        await marcarStatus("marcarErro", lead.protocolo);
      }
    } finally {
      await page.close();
      // Pausa curta entre envios pra não parecer tráfego automatizado agressivo.
      await new Promise((r) => setTimeout(r, 3000));
    }
  }

  await browser.close();
}

main().catch((err) => {
  console.error("Falha geral no robô:", err);
  process.exit(1);
});
