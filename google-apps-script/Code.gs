/**
 * Meu Caixa — sincronização com Planilha Google.
 *
 * Cole este código em Extensões → Apps Script da sua planilha,
 * rode a função "configurar" uma vez e implante como App da Web.
 * Cada aba (gastos, proventos, pagamentos, cartoes) guarda um tipo de lançamento.
 * A última coluna, "extra", guarda campos novos do app (por exemplo, as faturas
 * marcadas como pagas), para que versões futuras do app não exijam trocar este código.
 */

const VERSAO = 2;
const COLUNAS = {
  gastos: ["id", "data", "mes", "marca", "categoria", "valor", "parcelas", "cartao", "pessoa", "criado"],
  proventos: ["id", "data", "descricao", "tipo", "valor", "criado"],
  pagamentos: ["id", "data", "pessoa", "valor", "obs", "criado"],
  cartoes: ["id", "nome", "bandeira", "venc", "cor", "criado"],
};
const NUMERICAS = ["valor", "parcelas", "venc", "criado"];
const EXTRA = "extra";

function colunasDe(nome) { return COLUNAS[nome].concat([EXTRA]); }

/** Rode uma vez: cria as abas e gera a sua chave (aparece no registro de execução). */
function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(COLUNAS).forEach(function (nome) {
    const sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
    garantirCabecalho(sh, nome);
    sh.setFrozenRows(1);
  });
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("CHAVE")) props.setProperty("CHAVE", Utilities.getUuid().replace(/-/g, "").slice(0, 24));
  Logger.log("Sua chave do Meu Caixa: " + props.getProperty("CHAVE"));
}

/** Escreve o cabeçalho (inclusive a coluna "extra" em planilhas antigas) e deixa as colunas de texto como texto. */
function garantirCabecalho(sh, nome) {
  const cols = colunasDe(nome);
  const atual = sh.getRange(1, 1, 1, cols.length).getValues()[0];
  if (atual.join("|") === cols.join("|")) return;
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight("bold");
  cols.forEach(function (c, i) {
    if (NUMERICAS.indexOf(c) < 0) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat("@");
  });
}

function doPost(e) {
  let corpo = {};
  try { corpo = JSON.parse(e.postData.contents); } catch (err) { return saida({ ok: false, erro: "pedido", versao: VERSAO }); }
  const chave = PropertiesService.getScriptProperties().getProperty("CHAVE");
  if (!chave || corpo.chave !== chave) return saida({ ok: false, erro: "chave", versao: VERSAO });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (Array.isArray(corpo.ops) && corpo.ops.length) aplicar(corpo.ops);
    return saida({ ok: true, versao: VERSAO, dados: lerTudo() });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return saida({ ok: true, app: "meu-caixa", versao: VERSAO });
}

function saida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Aplica as alterações: o estado final de cada id vence (gravar ou excluir). */
function aplicar(ops) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const porAba = {};
  ops.forEach(function (o) {
    if (!o || !COLUNAS[o.col]) return;
    const id = o.op === "del" ? o.id : o.doc && o.doc.id;
    if (typeof id !== "string" || !id) return;
    (porAba[o.col] = porAba[o.col] || {})[id] = o.op === "del" ? null : o.doc;
  });
  Object.keys(porAba).forEach(function (nome) {
    const sh = ss.getSheetByName(nome);
    if (!sh) return;
    garantirCabecalho(sh, nome);
    const largura = colunasDe(nome).length;
    const final = porAba[nome];
    const ultima = sh.getLastRow();
    const ids = ultima > 1 ? sh.getRange(2, 1, ultima - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
    const linhaDe = {};
    ids.forEach(function (id, i) { if (id) linhaDe[id] = i + 2; });
    const novas = [], apagar = [];
    Object.keys(final).forEach(function (id) {
      const doc = final[id];
      if (doc === null) { if (linhaDe[id]) apagar.push(linhaDe[id]); return; }
      const linha = montarLinha(nome, doc);
      if (linhaDe[id]) sh.getRange(linhaDe[id], 1, 1, largura).setValues([linha]);
      else novas.push(linha);
    });
    if (novas.length) {
      const falta = sh.getLastRow() + novas.length - sh.getMaxRows();
      if (falta > 0) sh.insertRowsAfter(sh.getMaxRows(), falta);
      sh.getRange(sh.getLastRow() + 1, 1, novas.length, largura).setValues(novas);
    }
    apagar.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  });
}

function montarLinha(nome, doc) {
  const cols = COLUNAS[nome];
  const linha = cols.map(function (c) { return valorCelula(c, doc[c]); });
  const extra = {};
  Object.keys(doc).forEach(function (k) { if (k !== "id" && cols.indexOf(k) < 0) extra[k] = doc[k]; });
  linha.push(Object.keys(extra).length ? JSON.stringify(extra) : "");
  return linha;
}

function valorCelula(col, v) {
  if (v === undefined || v === null) return "";
  if (NUMERICAS.indexOf(col) >= 0) return typeof v === "number" ? v : (Number(v) || "");
  const s = String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function lerTudo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const dados = {};
  Object.keys(COLUNAS).forEach(function (nome) {
    const cols = COLUNAS[nome];
    const sh = ss.getSheetByName(nome);
    dados[nome] = [];
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, cols.length + 1).getValues().forEach(function (r) {
      if (!r[0]) return;
      let doc = {};
      const bruto = r[cols.length];
      if (bruto) { try { const ex = JSON.parse(bruto); if (ex && typeof ex === "object") doc = ex; } catch (err) {} }
      cols.forEach(function (c, i) {
        let v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, c === "mes" ? "yyyy-MM" : "yyyy-MM-dd");
        if (NUMERICAS.indexOf(c) >= 0) v = v === "" ? (c === "venc" ? null : 0) : Number(v);
        else v = String(v);
        doc[c] = v;
      });
      dados[nome].push(doc);
    });
  });
  return dados;
}
