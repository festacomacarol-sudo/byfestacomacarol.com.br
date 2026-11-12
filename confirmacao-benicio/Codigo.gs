/**
 * Confirmação de presença — Benício 3 anos
 * Cole este código em Extensões > Apps Script da sua planilha Google.
 * Depois: Implantar > Nova implantação > App da Web
 *   - Executar como: Eu
 *   - Quem pode acessar: Qualquer pessoa
 * Copie a URL que termina em /exec e cole em CONFIG.SCRIPT_URL no index.html.
 */

const NOME_ABA = 'Confirmações';
const CABECALHO = ['Data/hora', 'Nome', 'WhatsApp', 'Vem?', 'Adultos', 'Crianças', 'Total', 'Acompanhantes', 'Recado'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = e.parameter || {};
    const aba = pegarAba_();
    aba.appendRow([
      new Date(),
      limpar_(p.nome),
      limpar_(p.whatsapp),
      limpar_(p.presenca),
      Number(p.adultos) || 0,
      Number(p.criancas) || 0,
      Number(p.total) || 0,
      limpar_(p.acompanhantes),
      limpar_(p.recado)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Permite testar a URL no navegador
function doGet() {
  return ContentService.createTextOutput('Confirmação do Benício funcionando ✔');
}

function pegarAba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(NOME_ABA);
  if (!aba) aba = ss.insertSheet(NOME_ABA);
  if (aba.getLastRow() === 0) {
    aba.appendRow(CABECALHO);
    aba.getRange(1, 1, 1, CABECALHO.length)
      .setFontWeight('bold').setBackground('#f6e27a').setFontColor('#141414');
    aba.setFrozenRows(1);
    aba.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
    // linha de totais de convidados confirmados (fica na coluna K)
    aba.getRange('K1').setValue('Total confirmados').setFontWeight('bold');
    aba.getRange('K2').setFormula('=SUMIF(D2:D;"Sim";G2:G)');
  }
  return aba;
}

// Evita que textos começando com = + - @ virem fórmula na planilha
function limpar_(v) {
  v = String(v || '').trim().slice(0, 1000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
