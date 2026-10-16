/**
 * Lista de Oração 24h · ICM Grajaú · 17/10/2026
 *
 * Como publicar (uma vez só):
 *  1. Crie uma planilha nova no Google Sheets.
 *  2. Menu Extensões > Apps Script. Apague o conteúdo e cole este arquivo (Code.gs).
 *  3. Clique em "+" ao lado de Arquivos > HTML e crie um arquivo chamado Index.
 *     Cole nele o conteúdo do arquivo Index.html.
 *  4. Clique em Implantar > Nova implantação > tipo "Aplicativo da Web".
 *     - Executar como: Eu
 *     - Quem pode acessar: Qualquer pessoa
 *  5. Copie o link gerado e envie no grupo.
 */

var NOME_ABA = 'Horarios';
var SEMENTE = { '05:30': ['Borges'] };

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Lista de Oração 24h · 17/10/2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function _aba() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName(NOME_ABA);
  if (!aba) {
    aba = planilha.insertSheet(NOME_ABA);
    aba.appendRow(['hora', 'nomes']);
    aba.getRange('A:A').setNumberFormat('@');
    Object.keys(SEMENTE).forEach(function (h) {
      aba.appendRow([h, SEMENTE[h].join(';')]);
    });
  }
  return aba;
}

/** Devolve { "HH:MM": ["Nome", ...] } para a página. */
function getSlots() {
  var valores = _aba().getDataRange().getValues();
  var resultado = {};
  for (var i = 1; i < valores.length; i++) {
    var hora = String(valores[i][0]);
    var nomes = String(valores[i][1] || '');
    resultado[hora] = nomes ? nomes.split(';').filter(function (n) { return n; }) : [];
  }
  return resultado;
}

function _limpar(texto) {
  return String(texto || '').replace(/[;\n\r]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
}

function _linhaDaHora(aba, hora) {
  var valores = aba.getDataRange().getValues();
  for (var i = 1; i < valores.length; i++) {
    if (String(valores[i][0]) === hora) return i + 1;
  }
  return -1;
}

/** Adiciona um nome ao horário. Retorna { ok, erro?, slots }. */
function addName(hora, nome) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var n = _limpar(nome);
    if (!n) return { ok: false, erro: 'Digite seu nome.' };
    var aba = _aba();
    var linha = _linhaDaHora(aba, hora);
    var atuais = linha > 0 ? String(aba.getRange(linha, 2).getValue() || '').split(';').filter(function (x) { return x; }) : [];
    var repetido = atuais.some(function (x) { return x.toLowerCase() === n.toLowerCase(); });
    if (repetido) return { ok: false, erro: 'Esse nome já está neste horário.' };
    atuais.push(n);
    if (linha > 0) {
      aba.getRange(linha, 2).setValue(atuais.join(';'));
    } else {
      aba.appendRow([hora, atuais.join(';')]);
    }
    return { ok: true, slots: getSlots() };
  } finally {
    lock.releaseLock();
  }
}

/** Remove um nome do horário. */
function removeName(hora, nome) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var aba = _aba();
    var linha = _linhaDaHora(aba, hora);
    if (linha < 0) return { ok: false, erro: 'Horário não encontrado.' };
    var atuais = String(aba.getRange(linha, 2).getValue() || '').split(';').filter(function (x) { return x; });
    var nova = atuais.filter(function (x) { return x !== nome; });
    aba.getRange(linha, 2).setValue(nova.join(';'));
    return { ok: true, slots: getSlots() };
  } finally {
    lock.releaseLock();
  }
}
