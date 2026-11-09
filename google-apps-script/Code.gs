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

/* ======================= Robô do Telegram ======================= */
// 1) No Telegram, fale com @BotFather, crie um robô e copie o token.
// 2) Cole o token entre as aspas abaixo, salve e rode a função "configurarTelegram".
// 3) No Telegram, abra o seu robô e mande /start. Pronto.
// Depois de configurar, você pode apagar o token daqui: ele fica guardado no projeto.
const TELEGRAM_TOKEN = "";

function configurarTelegram() {
  const props = PropertiesService.getScriptProperties();
  const token = TELEGRAM_TOKEN.trim() || props.getProperty("TG_TOKEN");
  if (!token) throw new Error("Cole o token do robô em TELEGRAM_TOKEN, na parte do Telegram, e rode de novo.");
  props.setProperty("TG_TOKEN", token);
  const eu = tg("getMe", {});
  if (!eu.ok) throw new Error("O Telegram recusou o token. Confira se copiou o token inteiro do @BotFather.");
  tg("deleteWebhook", {});
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (["lerTelegram", "enviarResumoMensal"].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("lerTelegram").timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger("enviarResumoMensal").timeBased().onMonthDay(1).atHour(8).create();
  Logger.log("Robô @" + eu.result.username + " ligado. Abra o Telegram, procure @" + eu.result.username + " e mande /start.");
}

function tg(metodo, dados) {
  const token = PropertiesService.getScriptProperties().getProperty("TG_TOKEN");
  if (!token) return { ok: false };
  const r = UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/" + metodo, {
    method: "post", contentType: "application/json", payload: JSON.stringify(dados), muteHttpExceptions: true,
  });
  try { return JSON.parse(r.getContentText()); } catch (e) { return { ok: false }; }
}

function responder(chat, texto) { tg("sendMessage", { chat_id: chat, text: texto }); }

/** Roda a cada minuto: lê as mensagens novas e lança na planilha. */
function lerTelegram() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("TG_TOKEN")) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    const r = tg("getUpdates", { offset: Number(props.getProperty("TG_OFFSET") || 0), timeout: 0, allowed_updates: ["message"] });
    if (!r.ok || !r.result.length) return;
    r.result.forEach(function (u) {
      props.setProperty("TG_OFFSET", String(u.update_id + 1));
      const m = u.message;
      if (!m || typeof m.text !== "string") return;
      const chat = String(m.chat.id);
      const dono = props.getProperty("TG_CHAT");
      if (!dono) {
        // A primeira conta que mandar /start vira a dona do robô; as outras são ignoradas.
        if (/^\/start/.test(m.text.trim())) { props.setProperty("TG_CHAT", chat); responder(chat, "Pronto! Este robô agora é seu.\n\n" + AJUDA); }
        return;
      }
      if (chat !== dono) return;
      tratarMensagem(chat, m.text.trim());
    });
  } finally {
    lock.releaseLock();
  }
}

const AJUDA = "Mande o lançamento do jeito que falaria:\n" +
  "• uber 25,90 nubank\n• renner 300 3x inter p/ ana\n• mercado 120 ontem\n• recebi 2918,23 salario\n• ana pagou 50\n" +
  "• antecipei 50 nubank (ou: antecipei 50 nubank dezembro)\n• apliquei 200 reserva\n• resgatei 100\n\n" +
  "Comandos:\n/desfazer – apaga o último lançamento feito por aqui\n/resumo – resumo do mês atual\n/poupanca – saldo da poupança\n/ajuda – esta mensagem";

const MESES_TG = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
const SEM_CARTAO_ID = "sys-sem-cartao";

function hojeISO() { return Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd"); }
function fmtBR(v) { return "R$ " + (Math.round(v * 100) / 100).toFixed(2).replace(".", ",").replace(/\B(?=(\d{3})+(?!\d))/g, "."); }
function idxMes(s) { return +s.slice(0, 4) * 12 + (+s.slice(5, 7) - 1); }
function ymDe(i) { return Math.floor(i / 12) + "-" + String(i % 12 + 1).padStart(2, "0"); }
function semAcentoTG(s) { return String(s || "").normalize("NFD").replace(/\p{M}/gu, "").toLowerCase().trim(); }
function saldoPoupanca(pagamentos, ateYm) {
  return pagamentos.filter(function (p) { return (p.tipoMov === "aporte" || p.tipoMov === "resgate") && (!ateYm || p.data.slice(0, 7) <= ateYm); })
    .reduce(function (s, p) { return s + (p.tipoMov === "aporte" ? p.valor : -p.valor); }, 0);
}

function lancarDoRobo(chat, col, doc, resumo) {
  doc.id = "tg-" + Utilities.getUuid().slice(0, 13);
  doc.criado = Date.now();
  aplicar([{ col: col, op: "set", doc: doc }]);
  PropertiesService.getScriptProperties().setProperty("TG_ULTIMO", JSON.stringify({ col: col, id: doc.id, resumo: resumo }));
  responder(chat, "Lançado: " + resumo + "\nSe ficou errado, mande /desfazer.");
}

function tratarMensagem(chat, texto) {
  const props = PropertiesService.getScriptProperties();
  if (/^\/(start|ajuda|help)/i.test(texto)) return responder(chat, AJUDA);
  if (/^\/desfazer/i.test(texto)) {
    const ult = JSON.parse(props.getProperty("TG_ULTIMO") || "null");
    if (!ult) return responder(chat, "Não há lançamento recente para desfazer.");
    aplicar([{ col: ult.col, op: "del", id: ult.id }]);
    props.deleteProperty("TG_ULTIMO");
    return responder(chat, "Desfeito: " + ult.resumo);
  }
  if (/^\/resumo/i.test(texto)) return responder(chat, resumoDoMes(hojeISO().slice(0, 7)));
  const dados = lerTudo();
  if (/^\/poupan/i.test(texto)) {
    const ym = hojeISO().slice(0, 7);
    const doMes = dados.pagamentos.filter(function (p) { return (p.tipoMov === "aporte" || p.tipoMov === "resgate") && p.data.slice(0, 7) === ym; });
    const ap = doMes.filter(function (p) { return p.tipoMov === "aporte"; }).reduce(function (s, p) { return s + p.valor; }, 0);
    const rs = doMes.filter(function (p) { return p.tipoMov === "resgate"; }).reduce(function (s, p) { return s + p.valor; }, 0);
    return responder(chat, "Poupança: saldo " + fmtBR(saldoPoupanca(dados.pagamentos)) + "\nEm " + MESES_TG[idxMes(ym) % 12] + ": aplicou " + fmtBR(ap) + " · resgatou " + fmtBR(rs));
  }
  const nomes = {};
  dados.gastos.concat(dados.pagamentos).forEach(function (x) { if (x.pessoa && x.pessoa.trim()) nomes[x.pessoa.trim().toLowerCase()] = x.pessoa.trim(); });
  const cartoes = dados.cartoes.filter(function (c) { return c.id !== SEM_CARTAO_ID; });
  const ctx = {
    cartoes: cartoes.map(function (c) { return { id: c.id, nome: c.nome, prox: c.venceProxMes !== false }; }),
    pessoas: Object.keys(nomes).map(function (k) { return nomes[k]; }),
    hoje: hojeISO(),
  };

  // Poupança: "apliquei 200 reserva", "guardei 100", "resgatei 50"
  const mPoup = /^(apliquei|guardei|investi|resgatei|saquei|retirei)\b\s*/i.exec(texto);
  if (mPoup) {
    const tipo = /^(resgatei|saquei|retirei)$/i.test(mPoup[1]) ? "resgate" : "aporte";
    const r = interpretarLancamento(texto.slice(mPoup[0].length) || "x", ctx);
    if (!r || r.erro || !(r.doc.valor > 0)) return responder(chat, "Não achei o valor. Exemplo: apliquei 200 reserva");
    const desc = r.col === "gastos" && r.doc.marca !== "Gasto" ? r.doc.marca.replace(/^(na|no|em)\s+/i, "") : "";
    const doc = { pessoa: "", data: r.doc.data, valor: r.doc.valor, obs: desc, tipoMov: tipo };
    return lancarDoRobo(chat, "pagamentos", doc, (tipo === "aporte" ? "Aplicação na poupança: " : "Resgate da poupança: ") + fmtBR(doc.valor) + (desc ? " · " + desc : ""));
  }

  // Antecipação de fatura: "antecipei 50 nubank", "antecipei 50 nubank dezembro"
  const mAnt = /^(antecipei|adiantei|abati)\b\s*/i.exec(texto);
  if (mAnt) {
    const resto = texto.slice(mAnt[0].length);
    const r = interpretarLancamento(resto || "x", ctx);
    if (!r || r.erro || !(r.doc.valor > 0)) return responder(chat, "Não achei o valor. Exemplo: antecipei 50 nubank");
    const sa = semAcentoTG(resto);
    const semCartao = /sem cart|pix|dinheiro/.test(sa);
    const cartaoId = semCartao ? "_sem" : r.doc.cartao;
    if (!cartaoId) return responder(chat, "Diga de qual cartão é a fatura. Exemplo: antecipei 50 nubank");
    const cartao = cartoes.filter(function (c) { return c.id === cartaoId; })[0];
    const hoje = idxMes(hojeISO());
    let alvo = cartaoId !== "_sem" && cartao && cartao.venceProxMes !== false ? hoje + 1 : hoje;
    MESES_TG.forEach(function (nomeMes, i) {
      if (new RegExp("\\b" + semAcentoTG(nomeMes) + "\\b").test(sa)) { alvo = Math.floor(hoje / 12) * 12 + i; if (alvo < hoje) alvo += 12; }
    });
    const doc = { pessoa: "", data: r.doc.data, valor: r.doc.valor, obs: "via Telegram", tipoMov: "abate", cartao: cartaoId, mesFatura: ymDe(alvo), gasto: "" };
    return lancarDoRobo(chat, "pagamentos", doc, "Antecipação de " + fmtBR(doc.valor) + " na fatura de " + MESES_TG[alvo % 12] + " · " + (cartaoId === "_sem" ? "sem cartão" : cartao.nome));
  }

  const r = interpretarLancamento(texto, ctx);
  if (!r) return;
  if (r.erro) return responder(chat, r.erro);
  lancarDoRobo(chat, r.col, r.doc, r.resumo);
}

/** Dia 1º de cada mês, às 8h: manda o resumo das faturas do mês. */
function enviarResumoMensal() {
  const chat = PropertiesService.getScriptProperties().getProperty("TG_CHAT");
  if (chat) responder(chat, resumoDoMes(hojeISO().slice(0, 7)));
}

function resumoDoMes(ym) {
  const fmt = fmtBR, mi = idxMes;
  const m = mi(ym);
  const d = lerTudo();
  const cartoes = {};
  d.cartoes.forEach(function (c) { cartoes[c.id] = c; });
  const pagoNoMes = function (c) { return !!(c && Array.isArray(c.pagos) && c.pagos.indexOf(ym) >= 0); };
  const movs = d.pagamentos.filter(function (p) { return p.tipoMov === "abate" || p.tipoMov === "aporte" || p.tipoMov === "resgate"; });
  const abatidoPor = {};
  movs.forEach(function (a) { if (a.tipoMov === "abate" && a.mesFatura === ym) abatidoPor[a.cartao] = (abatidoPor[a.cartao] || 0) + a.valor; });

  const porCartao = {}, porPessoa = {};
  let meu = 0, outros = 0;
  d.gastos.forEach(function (g) {
    const n = Math.max(1, g.parcelas | 0);
    const ini = /^\d{4}-\d{2}$/.test(g.mes) ? mi(g.mes) : mi(g.data);
    const k = m - ini;
    if (k < 0 || k >= n) return;
    const tot = Math.round(g.valor * 100), base = Math.floor(tot / n);
    const v = (k === 0 ? tot - base * (n - 1) : base) / 100;
    const c = cartoes[g.cartao] && g.cartao !== SEM_CARTAO_ID ? cartoes[g.cartao] : null;
    const id = c ? c.id : "_sem";
    const pc = porCartao[id] = porCartao[id] || { nome: c ? c.nome : "Sem cartão", meu: 0, outros: 0, paga: c ? pagoNoMes(c) : pagoNoMes(cartoes[SEM_CARTAO_ID]), ab: abatidoPor[id] || 0 };
    if (g.pessoa && g.pessoa.trim()) {
      pc.outros += v; outros += v;
      const p = g.pessoa.trim(), kp = p.toLowerCase();
      porPessoa[kp] = porPessoa[kp] || { nome: p, devido: 0, pago: 0 };
      porPessoa[kp].devido += v;
    } else { pc.meu += v; meu += v; }
  });
  d.pagamentos.forEach(function (p) {
    if (!p.pessoa || p.tipoMov || p.data.slice(0, 7) !== ym) return;
    const kp = p.pessoa.trim().toLowerCase();
    if (porPessoa[kp]) porPessoa[kp].pago += p.valor;
  });

  // Proventos: os lançados no mês + os fixos (um lançado com o mesmo nome vale no lugar do fixo).
  const mesP = function (p) { return /^\d{4}-\d{2}$/.test(p.mes) ? p.mes : p.data.slice(0, 7); };
  const lancados = d.proventos.filter(function (p) { return p.fixo !== true && mesP(p) === ym; });
  const fixos = d.proventos.filter(function (f) {
    if (f.fixo !== true || (f.desde && f.desde > ym)) return false;
    const n = semAcentoTG(f.descricao);
    return !lancados.some(function (l) { const a = semAcentoTG(l.descricao); return a === n || a.indexOf(n) >= 0; });
  });
  const somaL = lancados.reduce(function (s, p) { return s + p.valor; }, 0), somaF = fixos.reduce(function (s, p) { return s + p.valor; }, 0);
  const entrou = somaL + somaF;

  // Caixa do mês: antecipação sai no mês em que foi paga; aplicação na poupança sai, resgate volta.
  let ajuste = 0, aplicou = 0, resgatou = 0;
  movs.forEach(function (p) {
    if (p.tipoMov === "abate") {
      if (!/^\d{4}-\d{2}$/.test(p.mesFatura || "")) return;
      const pm = p.data.slice(0, 7);
      if (pm === p.mesFatura) return;
      if (pm === ym) ajuste += p.valor;
      if (p.mesFatura === ym) ajuste -= p.valor;
    } else if (p.data.slice(0, 7) === ym) {
      if (p.tipoMov === "aporte") { ajuste += p.valor; aplicou += p.valor; } else { ajuste -= p.valor; resgatou += p.valor; }
    }
  });

  const ids = Object.keys(porCartao).sort(function (a, b) { return porCartao[a].nome.localeCompare(porCartao[b].nome); });
  let faltaPagar = 0, antecipado = 0;
  ids.forEach(function (id) { const pc = porCartao[id]; antecipado += pc.ab; if (!pc.paga) faltaPagar += Math.max(0, pc.meu + pc.outros - pc.ab); });

  const linhas = ["Meu Caixa – " + MESES_TG[m % 12] + " de " + Math.floor(m / 12), "O que você paga e recebe neste mês.", ""];
  linhas.push("Seus gastos que vencem no mês: " + fmt(meu));
  linhas.push("De outras pessoas: " + fmt(outros));
  linhas.push("Total das faturas do mês: " + fmt(meu + outros));
  if (antecipado > 0.004) linhas.push("Já antecipado: " + fmt(antecipado));
  linhas.push("Falta pagar: " + fmt(faltaPagar));
  linhas.push("");
  linhas.push("Por cartão:");
  if (!ids.length) linhas.push("• nenhum gasto neste mês");
  ids.forEach(function (id) {
    const pc = porCartao[id], tot = pc.meu + pc.outros;
    const situacao = pc.paga ? " – paga" : pc.ab > 0.004 ? " – antecipado " + fmt(pc.ab) + " · falta " + fmt(Math.max(0, tot - pc.ab)) : " – a pagar";
    linhas.push("• " + pc.nome + ": " + fmt(tot) + " (seu " + fmt(pc.meu) + (pc.outros ? " · outros " + fmt(pc.outros) : "") + ")" + situacao);
  });
  const pessoas = Object.keys(porPessoa).sort();
  if (pessoas.length) {
    linhas.push("");
    linhas.push("Gastos de outras pessoas:");
    pessoas.forEach(function (k) {
      const p = porPessoa[k], falta = Math.max(0, p.devido - p.pago);
      linhas.push("• " + p.nome + ": " + fmt(p.devido) + (p.pago ? " · pagou " + fmt(p.pago) : "") + " · falta " + fmt(falta));
    });
  }
  linhas.push("");
  linhas.push("Entra no caixa: " + fmt(entrou) + (somaF ? " (fixos " + fmt(somaF) + (somaL ? " + lançados " + fmt(somaL) : "") + ")" : ""));
  const saldoP = saldoPoupanca(d.pagamentos, ym);
  if (aplicou || resgatou || saldoP) linhas.push("Poupança: saldo " + fmt(saldoP) + (aplicou || resgatou ? " (no mês: aplicou " + fmt(aplicou) + " · resgatou " + fmt(resgatou) + ")" : ""));
  linhas.push("Sobra: " + fmt(entrou - Math.max(0, meu + ajuste)));
  return linhas.join("\n");
}

/* ============ Interpretador de frases (o mesmo do app) ============ */
// Interpreta uma frase curta ("uber 25,90 nubank", "renner 300 3x inter p/ ana",
// "recebi 2918,23 salario", "ana pagou 50") e devolve o lançamento pronto.
// ctx: { cartoes: [{id, nome, prox}], pessoas: [nomes], hoje: "AAAA-MM-DD" }
// prox = a fatura do cartão vence no mês seguinte ao da compra.
function interpretarLancamento(texto, ctx) {
  var sem = function (s) { return String(s || "").normalize("NFD").replace(/\p{M}/gu, "").toLowerCase(); };
  var orig = String(texto || "").replace(/\s+/g, " ").trim();
  if (!orig) return null;
  var W = orig.split(" ");
  var P = W.map(function (w) { return sem(w).replace(/[.,;:!?]+$/, function (m) { return /\d$/.test(w.slice(0, -m.length)) ? m : ""; }); });
  var usado = W.map(function () { return false; });
  var hoje = ctx.hoje;
  var pad = function (n) { return String(n).padStart(2, "0"); };
  var ymd = function (d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); };
  var dHoje = new Date(+hoje.slice(0, 4), +hoje.slice(5, 7) - 1, +hoje.slice(8, 10));

  // data
  var data = hoje;
  P.forEach(function (w, i) {
    if (usado[i]) return;
    if (w === "hoje") { usado[i] = true; }
    else if (w === "ontem" || w === "anteontem") { var d = new Date(dHoje); d.setDate(d.getDate() - (w === "ontem" ? 1 : 2)); data = ymd(d); usado[i] = true; }
    else {
      var m = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/.exec(w);
      if (m && +m[1] >= 1 && +m[1] <= 31 && +m[2] >= 1 && +m[2] <= 12) {
        var y = m[3] ? (+m[3] < 100 ? 2000 + +m[3] : +m[3]) : dHoje.getFullYear();
        data = y + "-" + pad(m[2]) + "-" + pad(m[1]); usado[i] = true;
      }
    }
  });

  // parcelas: "3x", "x3", "3 vezes"
  var parcelas = 1;
  P.forEach(function (w, i) {
    if (usado[i]) return;
    var m = /^(\d{1,2})x$/.exec(w) || /^x(\d{1,2})$/.exec(w);
    if (m && +m[1] >= 1) { parcelas = +m[1]; usado[i] = true; }
    else if (/^\d{1,2}$/.test(w) && (P[i + 1] === "vezes" || P[i + 1] === "parcelas") && !usado[i + 1]) { parcelas = +w; usado[i] = usado[i + 1] = true; }
  });

  // valor: primeiro número (aceita 25,90 · 25.90 · 1.500,00 · R$ 25,90)
  var valor = null;
  P.forEach(function (w, i) {
    if (usado[i] || valor !== null) return;
    if (w === "r$") { usado[i] = true; return; }
    var v = w.replace(/^r\$/, "");
    if (!/^(\d{1,3}(\.\d{3})+|\d+)([.,]\d{1,2})?$/.test(v)) return;
    var n = v.indexOf(",") >= 0 ? parseFloat(v.replace(/\./g, "").replace(",", "."))
      : /\.\d{3}$/.test(v) ? parseFloat(v.replace(/\./g, "")) : parseFloat(v);
    if (n > 0) { valor = Math.round(n * 100) / 100; usado[i] = true; if (P[i - 1] === "r$") usado[i - 1] = true; }
  });
  P.forEach(function (w, i) { if (w === "reais" || w === "real") usado[i] = true; });

  // sequência de palavras (ignorando as já usadas) que bate com um nome
  var acharSeq = function (nome) {
    var alvo = sem(nome).split(/\s+/).filter(Boolean);
    if (!alvo.length) return -1;
    for (var i = 0; i + alvo.length <= P.length; i++) {
      var ok = true;
      for (var j = 0; j < alvo.length; j++) if (usado[i + j] || P[i + j] !== alvo[j]) { ok = false; break; }
      if (ok) return i;
    }
    return -1;
  };
  var marcar = function (i, n) { for (var j = 0; j < n; j++) usado[i + j] = true; };

  // cartão (o nome mais longo primeiro)
  var cartao = "", cartaoNome = "", cartaoProx = false;
  (ctx.cartoes || []).slice().sort(function (a, b) { return b.nome.length - a.nome.length; }).some(function (c) {
    var i = acharSeq(c.nome);
    if (i < 0) return false;
    marcar(i, sem(c.nome).split(/\s+/).filter(Boolean).length);
    cartao = c.id; cartaoNome = c.nome; cartaoProx = !!c.prox; return true;
  });

  // tipo
  var tipo = "gasto";
  var iPagou = P.findIndex(function (w, i) { return !usado[i] && (w === "pagou" || w === "devolveu" || w === "pagamento" || w === "acertou"); });
  if (["recebi", "salario", "provento", "entrou", "recebimento"].indexOf(P[0]) >= 0) tipo = "provento";
  else if (iPagou >= 0) tipo = "pagamento";

  // "este mês" / "neste mês" / "mesmo mês": o provento entra no próprio mês
  var mesmoMes = false;
  P.forEach(function (w, i) {
    if ((w === "este" || w === "neste" || w === "mesmo") && P[i + 1] === "mes" && !usado[i] && !usado[i + 1]) { mesmoMes = true; usado[i] = usado[i + 1] = true; }
  });

  // pessoa: nomes já usados no app, ou "p/ Ana", "pra Ana", "para Ana"
  var pessoa = "";
  (ctx.pessoas || []).slice().sort(function (a, b) { return b.length - a.length; }).some(function (n) {
    var i = acharSeq(n);
    if (i < 0) return false;
    marcar(i, sem(n).split(/\s+/).filter(Boolean).length);
    if (i > 0 && ["p/", "pra", "para", "pro", "da", "do"].indexOf(P[i - 1]) >= 0) usado[i - 1] = true;
    pessoa = n; return true;
  });
  if (!pessoa) P.forEach(function (w, i) {
    if (pessoa || usado[i] || ["p/", "pra", "para", "pro"].indexOf(w) < 0 || i + 1 >= P.length || usado[i + 1]) return;
    if (tipo === "gasto" && !/^[A-ZÀ-Ý]/.test(W[i + 1]) && w !== "p/") return;
    pessoa = W[i + 1].charAt(0).toUpperCase() + W[i + 1].slice(1).toLowerCase(); usado[i] = usado[i + 1] = true;
  });
  if (tipo === "pagamento") {
    usado[iPagou] = true;
    if (!pessoa && iPagou > 0 && !usado[iPagou - 1]) { pessoa = W[iPagou - 1].charAt(0).toUpperCase() + W[iPagou - 1].slice(1).toLowerCase(); usado[iPagou - 1] = true; }
    P.forEach(function (w, i) { if (w === "me") usado[i] = true; });
  }

  // categoria
  var CAT = {
    "Despesa fixa": ["aluguel", "condominio", "luz", "energia", "agua", "internet", "claro", "vivo", "tim", "oi", "netflix", "youtube", "spotify", "academia", "crossfit", "plano", "faculdade", "escola", "emprestimo", "seguro", "iptu", "ipva", "financiamento", "fixa", "fixo", "assinatura", "icloud", "apple", "multa"],
    "Lazer": ["lazer", "sauna", "spa", "bar", "cinema", "show", "viagem", "passeio", "festa", "balada", "hotel", "praia", "jogo", "shopping", "ingresso"],
    "Alimentação": ["alimentacao", "comida", "mercado", "supermercado", "ifood", "padaria", "restaurante", "lanche", "pizza", "hamburguer", "feira", "hortifrut", "acougue", "almoco", "jantar", "cafe", "mercadinho", "panificadora", "frutas", "sorvete", "suplemento"],
  };
  var categoria = "Outros";
  var explicito = { fixa: "Despesa fixa", fixo: "Despesa fixa", lazer: "Lazer", alimentacao: "Alimentação", comida: "Alimentação", outros: "Outros" };
  P.forEach(function (w, i) { if (!usado[i] && explicito[w]) { categoria = explicito[w]; usado[i] = true; } });
  if (categoria === "Outros") Object.keys(CAT).some(function (k) {
    return P.some(function (w, i) { return !usado[i] && CAT[k].some(function (c) { return w === c || (c.length > 4 && w.indexOf(c) === 0); }); }) && (categoria = k);
  });

  // o que sobrou vira a loja / descrição
  var FILLER = ["no", "na", "do", "da", "de", "com", "cartao", "em", "pelo", "pela", "a", "o", "vista", "avista", "credito", "debito", "via", "e", "gastei", "comprei", "paguei"];
  var resto = W.filter(function (w, i) { return !usado[i] && FILLER.indexOf(P[i]) < 0; }).join(" ").trim();
  var cap = function (s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; };

  if (valor === null) return { erro: "Não achei o valor. Escreva, por exemplo: uber 25,90 nubank" };
  var MES3 = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
  var somaMes = function (ym, n) { var t = +ym.slice(0, 4) * 12 + (+ym.slice(5, 7) - 1) + n; return Math.floor(t / 12) + "-" + pad(t % 12 + 1); };
  var mesTxt = function (ym) { return MES3[+ym.slice(5, 7) - 1] + "/" + ym.slice(2, 4); };
  var mes = somaMes(data.slice(0, 7), cartaoProx ? 1 : 0);
  var fmt = function (v) { return "R$ " + v.toFixed(2).replace(".", ",").replace(/\B(?=(\d{3})+(?!\d))/g, "."); };
  var dataBR = data.slice(8, 10) + "/" + data.slice(5, 7);
  if (tipo === "provento") {
    var ehSal = P.indexOf("salario") >= 0;
    var desc = cap(resto.replace(/^(recebi|provento|entrou|recebimento)\s*/i, "")) || (ehSal ? "Salário" : "Provento");
    if (sem(desc) === "salario") desc = "Salário";
    var tipoP = ehSal ? "Salário" : /freela|extra|bico/.test(sem(resto)) ? "Extra / freela" : "Outros";
    var mesP = somaMes(data.slice(0, 7), mesmoMes ? 0 : 1);
    return { col: "proventos", doc: { descricao: desc, tipo: tipoP, data: data, mes: mesP, valor: valor },
      resumo: "Provento: " + desc + " · " + fmt(valor) + " · " + dataBR + " · entra em " + mesTxt(mesP) };
  }
  if (tipo === "pagamento") {
    if (!pessoa) return { erro: "Diga quem pagou. Exemplo: ana pagou 50" };
    return { col: "pagamentos", doc: { pessoa: pessoa, data: data, valor: valor, obs: cap(resto) },
      resumo: "Pagamento de " + pessoa + " · " + fmt(valor) + " · " + dataBR };
  }
  var marca = cap(resto) || "Gasto";
  return { col: "gastos", doc: { marca: marca, data: data, mes: mes, categoria: categoria, valor: valor, parcelas: parcelas, cartao: cartao, pessoa: pessoa },
    resumo: "Gasto: " + marca + " · " + fmt(valor) + (parcelas > 1 ? " em " + parcelas + "x" : "") + " · " + (cartaoNome || "sem cartão") + " · " + categoria + (pessoa ? " · de " + pessoa : "") + " · " + dataBR + " · fatura de " + mesTxt(mes) };
}
