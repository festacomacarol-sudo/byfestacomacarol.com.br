/**
 * Confirmações de presença — 15 anos da Katherine
 * Cole este código em: Planilha Google > Extensões > Apps Script
 * Depois: Implantar > Nova implantação > Tipo "App da Web"
 *   Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 * Copie a URL que termina em /exec e envie para colocar no site.
 */

const COLUNAS = ['Data/hora', 'Nome completo', 'Telefone', 'Vai?', 'Pessoas', 'Acompanhantes', 'Crianças (nome e idade)'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // usa a primeira aba da planilha (a que você vê ao abrir)
    const aba = ss.getSheets()[0];
    if (aba.getLastRow() === 0) {
      aba.appendRow(COLUNAS);
      aba.getRange(1, 1, 1, COLUNAS.length)
        .setFontWeight('bold').setBackground('#DC7889').setFontColor('#FFFFFF');
      aba.setFrozenRows(1);
    }
    // o apóstrofo mantém o telefone como texto
    aba.appendRow([
      d.data || new Date(), String(d.nome || '').slice(0, 120), "'" + String(d.telefone || '').slice(0, 20),
      d.presenca || '', Number(d.pessoas) || 0,
      String(d.acompanhantes || '').slice(0, 500), String(d.criancas || '').slice(0, 500)
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

// Teste: abra a URL /exec no navegador. Se aparecer "Planilha conectada", está tudo certo.
function doGet() {
  return ContentService.createTextOutput('Planilha conectada ✔ — ' + SpreadsheetApp.getActiveSpreadsheet().getName());
}
