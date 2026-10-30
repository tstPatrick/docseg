/* =====================================================================
   DocSeg Digital — Google Apps Script (versão com login individual)
   ---------------------------------------------------------------------
   O que este script faz:
   - Login individual por técnico (perfis: admin, gestor, tecnico)
   - Recebe os documentos (PDF) do app e guarda no Google Drive desta conta
   - Registra cada envio na planilha "DocSeg Digital - Registros"
   - Painel do gestor: quem enviou, quando ("há 5 min") e baixar o PDF
   - E-mail automático: OPCIONAL, ligado por empresa (padrão: desligado)
   - Guarda de PDFs por 24 meses (limpeza automática semanal)
   - Aceite dos Termos de Uso (data e hora registradas)
   - Numeração da PT por empresa e lembrete diário (push)
   ===================================================================== */

var TERMS_VERSION = "1.0";      // mude para "1.1" etc. para pedir novo aceite a todos
var RETENCAO_MESES = 24;        // PDFs mais antigos que isso são apagados do Drive
var PASTA_RAIZ = "DocSeg Digital - PDFs";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var a = data.action;

    if (a === "login") return loginUser(data.username || "", data.passwordHash || "");
    if (a === "changePassword") return changePassword(data);
    if (a === "acceptTerms") return acceptTerms(data);

    if (a === "listUsers") return listUsers(data);
    if (a === "saveUser") return saveUser(data);
    if (a === "deleteUser") return deleteUser(data);
    if (a === "saveFoto") return saveFoto(data);
    if (a === "saveFcmToken") return saveFcmToken(data);
    if (a === "changeAdminPassword") return changeAdminPassword(data);

    if (a === "listEmpresas") return listEmpresas(data);
    if (a === "saveEmpresa") return saveEmpresa(data);
    if (a === "saveAviso") return saveAviso(data);
    if (a === "savePerguntas") return savePerguntas(data);
    if (a === "deleteEmpresa") return deleteEmpresa(data);

    if (a === "enviarMensagem") return enviarMensagem(data);
    if (a === "listMensagens") return listMensagens(data);
    if (a === "marcarLidas") return marcarLidas(data);
    if (a === "docAberto") return docAberto(data);
    if (a === "docFechado") return docFechado(data);
    if (a === "listAbertos") return listAbertos(data);
    if (a === "lembrarColaborador") return lembrarColaborador(data);
    if (a === "responderChamado") return responderChamado(data);
    if (a === "relatorioAgora") return relatorioAgora(data);
    if (a === "backupNow") return backupNow(data);
    if (a === "listBackups") return listBackups(data);
    if (a === "restoreBackup") return restoreBackup(data);

    if (a === "saveChamado") return saveChamado(data);
    if (a === "listChamados") return listChamados(data);
    if (a === "updateChamado") return updateChamado(data);

    if (a === "listPainel") return listPainel(data);
    if (a === "getPdf") return getPdf(data);

    if (a === "sendBatch") return sendBatchDocuments(data);
    if (a === "listAuditoria") return listAuditoria(data);
    if (a === "deleteAuditoria") return deleteAuditoria(data);

    if (a === "listLiberacoes") return listLiberacoes(data);
    if (a === "minhasLiberacoes") return minhasLiberacoes(data);
    if (a === "decidirLiberacao") return decidirLiberacao(data);
    if (a === "encerrarLiberacao") return encerrarLiberacao(data);
    if (a === "getCarimbo") return getCarimbo(data);
    if (a === "saveCarimbo") return saveCarimbo(data);

    if (a === "saveEmailEmpresa") return saveEmailEmpresa(data);
    if (a === "saveSons") return saveSons(data);
    if (a === "apagarChamado") return apagarChamado(data);
    if (a === "limparChamados") return limparChamados(data);

    if (a === "nextNumber" || a === "peekNumber" || a === "setNumber" || a === "sheetUrl") return numeracao(data);

    // Envio de um documento só (formato antigo do app, sem "action")
    return receiveSingle(data);
  } catch (err) {
    return ContentService
      .createTextOutput("erro: " + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

/* ============ RECEBIMENTO DE DOCUMENTOS ============ */
// Quem envia é SEMPRE identificado pelo login e senha (nada de e-mail
// "solto"). A empresa vem do cadastro do usuário no servidor, não do app.
function receiveSingle(data) {
  var caller = authCaller(data);
  if (!caller) return textOut("erro: não autorizado");
  var r = saveDocuments(caller, [data], data.empresa);
  if (r.error) return textOut("erro: " + r.error);
  return textOut("sucesso: documento recebido");
}

function sendBatchDocuments(data) {
  var caller = authCaller(data);
  if (!caller) return jsonOut({ error: "não autorizado" });
  var items = data.items || [];
  if (!items.length) return jsonOut({ error: "nenhum documento para enviar" });
  var r = saveDocuments(caller, items, data.empresa);
  if (r.error) return jsonOut({ error: r.error });
  return jsonOut({ ok: true, sent: r.saved, emailed: r.emailed });
}

/* ============ DOCUMENTOS CONTRATADOS (opcionais por empresa) ============ */
// Documento novo: incluir a chave em DOC_TIPOS, o nome em DOC_NOMES e o
// reconhecimento em tipoDocKey. A liberação (Apto / Não apto), as pastas
// "Não Aptos" e o carimbo passam a valer para ele sem mais nada aqui.
var DOC_TIPOS = ["pt", "veiculo", "escada"];
var DOC_NOMES = { pt: "Permissão de Trabalho", veiculo: "Checklist de Veículo", escada: "Checklist de Escada" };
function tipoDocKey(t) {
  t = (t || "").toString().toLowerCase();
  if (t.indexOf("permiss") !== -1) return "pt";
  if (t.indexOf("culo") !== -1) return "veiculo";
  if (t.indexOf("escada") !== -1) return "escada";
  return "outro";
}
// Empresas antigas (sem o campo) continuam com todos os documentos.
function docsDaEmpresa(rec) {
  var d = rec && rec.documentos;
  if (!d || !d.length) return DOC_TIPOS.slice();
  return DOC_TIPOS.filter(function (k) { return d.indexOf(k) !== -1; });
}

function saveDocuments(caller, items, empresaFallback) {
  var empresaNome = caller.role === "admin" ? (empresaFallback || "") : (caller.empresa || "");
  if (!empresaNome) return { error: "usuário sem empresa cadastrada" };
  var contratados = getEmpresaExtras(empresaNome).documentos;
  for (var k = 0; k < items.length; k++) {
    var chave = tipoDocKey(items[k].tipoDocumento);
    if (chave !== "outro" && contratados.indexOf(chave) === -1) {
      return { error: "o documento \"" + DOC_NOMES[chave] + "\" não está contratado por esta empresa" };
    }
  }
  var nomeTecnico = caller.nome || caller.username;

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var saved = 0, blobs = [], linhas = [], pendentes = [];
  try {
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      if (!it.pdfBase64) continue;
      if (it.docId && (jaRecebido(it.docId) || acharLiberacao_(it.docId))) { saved++; continue; } // reenvio: já está salvo
      var blob = Utilities.newBlob(Utilities.base64Decode(it.pdfBase64), "application/pdf",
        it.pdfNome || ("documento_" + Date.now() + ".pdf"));
      // Empresa com liberação ligada para este documento: o PDF espera na
      // pasta "_Pendentes" (fora da pasta da empresa) até o gestor decidir.
      if (it.docId && liberacaoExigida_(empresaNome, tipoDocKey(it.tipoDocumento))) {
        registrarPendente_(caller, empresaNome, it, pastaPendentes_().createFile(blob));
        saved++;
        pendentes.push(it);
        continue;
      }
      var file = pastaDoMes(empresaNome).createFile(blob);
      registrarNaPlanilha({
        tipoDocumento: it.tipoDocumento, ptNumero: it.ptNumero, ptData: it.ptData, ptArea: it.ptArea,
        empresa: empresaNome, usuario: nomeTecnico, login: caller.username, resumo: it.resumo,
        pdfNome: blob.getName(), fileId: file.getId(), docId: it.docId || "",
        horaIni: String(it.ptHoraIni || "").slice(0, 5), horaFim: String(it.ptHoraFim || "").slice(0, 5), motivo: String(it.ptMotivo || "").slice(0, 60),
        naoConformes: Array.isArray(it.naoConformes) ? it.naoConformes.slice(0, 40).map(function (x) { return String(x).slice(0, 160); }) : []
      });
      saved++;
      blobs.push(blob);
      linhas.push((saved) + ". " + (it.tipoDocumento || "Documento") + " — Nº " + (it.ptNumero || "S/N") +
        (it.ptArea ? " — " + it.ptArea : "") + " (" + (it.ptData || "") + ")");
    }
  } finally {
    lock.releaseLock();
  }

  // E-mail: só se a empresa ligou a opção. Falha de e-mail nunca desfaz o envio.
  var emailed = false;
  try {
    var ext = getEmpresaExtras(empresaNome);
    if (blobs.length && ext.emailAtivo && ext.emailDestino) {
      GmailApp.sendEmail(ext.emailDestino,
        "[DocSeg Digital] " + blobs.length + " documento" + (blobs.length > 1 ? "s" : "") + " - " + empresaNome,
        "Enviado por " + nomeTecnico + " pelo DocSeg Digital:\n\n" + linhas.join("\n"),
        { attachments: blobs, name: "DocSeg Digital" });
      emailed = true;
    }
  } catch (mailErr) {
    Logger.log("Falha ao enviar e-mail: " + mailErr.message);
  }
  try { items.forEach(function (it) { if (it.abertoId) removerAberto_(it.abertoId); }); } catch (e) {}
  if (blobs.length) pushGestores_(empresaNome, "Novo documento", nomeTecnico + " enviou " + blobs.length + " documento" + (blobs.length > 1 ? "s" : "") + ".");
  pendentes.forEach(function (it) {
    var nc = Array.isArray(it.naoConformes) ? it.naoConformes.length : 0;
    pushLiberadores_(empresaNome, "📋 " + tituloLib_(it.tipoDocumento, it.ptNumero) + " aguardando sua avaliação",
      nomeTecnico + " enviou" + (it.ptArea ? " · " + it.ptArea : "") + (nc ? ". " + nc + (nc > 1 ? " itens não conformes." : " item não conforme.") : "."));
  });
  return { saved: saved, emailed: emailed, pendentes: pendentes.length };
}

function jaRecebido(docId) {
  var sheet = getOrCreateAuditSheet();
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var from = Math.max(2, last - 300);
  var col = sheet.getRange(from, 13, last - from + 1, 1).getValues();
  for (var i = 0; i < col.length; i++) if (String(col[i][0]) === String(docId)) return true;
  return false;
}

/* ============ GOOGLE DRIVE (PDFs) ============ */
// Estrutura: DocSeg Digital - PDFs / <Empresa> / <AAAA-MM> / arquivo.pdf
function getRaiz() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("root_folder_id");
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* recria abaixo */ }
  }
  var pasta = DriveApp.createFolder(PASTA_RAIZ);
  props.setProperty("root_folder_id", pasta.getId());
  return pasta;
}

function subpasta(pai, nome) {
  var it = pai.getFoldersByName(nome);
  return it.hasNext() ? it.next() : pai.createFolder(nome);
}

function pastaDoMes(empresaNome) {
  var tz = Session.getScriptTimeZone();
  var mes = Utilities.formatDate(new Date(), tz, "yyyy-MM");
  return subpasta(subpasta(getRaiz(), slugSheetTabName(empresaNome)), mes);
}

// Pasta da empresa (nível acima dos meses). Quem recebe acesso aqui enxerga
// também todas as subpastas e PDFs, inclusive os que ainda vão chegar.
function pastaDaEmpresa(empresaNome) {
  return subpasta(getRaiz(), slugSheetTabName(empresaNome));
}

function parseEmails(txt) {
  var vistos = {}, lista = [];
  String(txt || "").split(/[\s,;]+/).forEach(function (e) {
    e = e.trim().toLowerCase();
    if (e && !vistos[e]) { vistos[e] = true; lista.push(e); }
  });
  return lista;
}

// Compartilha a pasta da empresa (somente leitura) com os e-mails do cadastro
// e retira o acesso de quem saiu da lista. Devolve quem ficou com acesso e
// quem não pôde ser incluído (e-mail sem conta Google, por exemplo).
function sincronizarAcessoPasta(empresaNome, emails, jaCompartilhados) {
  var pasta = pastaDaEmpresa(empresaNome);
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var ficam = [], falhas = [];
  emails.forEach(function (e) {
    if (!re.test(e)) { falhas.push(e + " (e-mail inválido)"); return; }
    if (jaCompartilhados.indexOf(e) !== -1) { ficam.push(e); return; }
    try { pasta.addViewer(e); ficam.push(e); }
    catch (err) { falhas.push(e + " (" + err.message + ")"); }
  });
  jaCompartilhados.forEach(function (e) {
    if (emails.indexOf(e) === -1) {
      try { pasta.removeViewer(e); } catch (err) { Logger.log("removeViewer " + e + ": " + err.message); }
    }
  });
  return { shared: ficam, falhas: falhas, url: pasta.getUrl() };
}

/* ============ PLANILHA DE REGISTROS (auditoria) ============ */
var AUDIT_HEADERS = [
  "Data/Hora do envio", "Tipo de documento", "Número/ID", "Data do documento",
  "Área/Local", "Empresa", "Técnico", "E-mail(s) de destino", "Resumo", "Nome do arquivo",
  "ID do arquivo (Drive)", "Login", "ID do documento (app)", "Não conformidades", "Início", "Término", "Motivo do encerramento",
  "Liberação", "Liberado por", "Liberado em", "Motivo (não apto)"
];
var AUDIT_ULTIMA = "Motivo (não apto)";

function getOrCreateSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty("audit_sheet_id");
  if (sheetId) {
    try { return SpreadsheetApp.openById(sheetId); } catch (e) { /* recria abaixo */ }
  }
  var ss = SpreadsheetApp.create("DocSeg Digital - Registros");
  props.setProperty("audit_sheet_id", ss.getId());
  return ss;
}

function ensureAuditHeaders(sheet) {
  var lastCol = sheet.getLastColumn();
  var headerRow = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (headerRow.indexOf(AUDIT_ULTIMA) === -1) {
    sheet.getRange(1, 1, 1, AUDIT_HEADERS.length).setValues([AUDIT_HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function getOrCreateAuditSheet() {
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName("Registros") || ss.getSheets()[0];
  if (sheet.getName() !== "Registros") sheet.setName("Registros");
  ensureAuditHeaders(sheet);
  return sheet;
}

// Nome de aba/pasta válido: até 95 caracteres e sem [ ] * / \ ? :
function slugSheetTabName(nome) {
  var limpo = String(nome || "").replace(/[\[\]\*\/\\\?:]/g, "-").trim();
  if (!limpo) limpo = "Sem empresa";
  return limpo.substring(0, 95);
}

function getOrCreateEmpresaSheet(ss, empresaNome) {
  var tabName = slugSheetTabName(empresaNome);
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.insertSheet(tabName);
  ensureAuditHeaders(sheet);
  return sheet;
}

function getOrCreateSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function registrarNaPlanilha(d) {
  var ss = getOrCreateSpreadsheet();
  var row = [
    new Date(), d.tipoDocumento || "", d.ptNumero || "", d.ptData || "", d.ptArea || "",
    d.empresa || "", d.usuario || "", "", d.resumo || "", d.pdfNome || "",
    d.fileId || "", d.login || "", d.docId || "", (d.naoConformes || []).join(" | "),
    d.horaIni || "", d.horaFim || "", d.motivo || "",
    d.liberacao || "", d.liberadoPor || "", d.liberadoEm || "", d.motivoNaoApto || ""
  ];
  getOrCreateAuditSheet().appendRow(row);
  if (d.empresa) getOrCreateEmpresaSheet(ss, d.empresa).appendRow(row);
}

function hhmm_(x) { // Sheets pode converter "08:12" em data/hora; devolve sempre "HH:mm"
  if (x instanceof Date) return Utilities.formatDate(x, TZ_CONTRATO, "HH:mm");
  var t = String(x || "").trim();
  var m = /^(\d{1,2}):(\d{2})/.exec(t);
  return m ? (m[1].length < 2 ? "0" + m[1] : m[1]) + ":" + m[2] : "";
}
function rowToObj(v, tz) {
  var ts = v[0] instanceof Date ? v[0].getTime() : new Date(v[0]).getTime();
  var fid = String(v[10] || "");
  return {
    id: String(ts), ts: ts,
    dataEnvio: Utilities.formatDate(new Date(ts), tz, "dd/MM/yyyy HH:mm"),
    tipoDocumento: v[1], numero: v[2], dataDocumento: v[3], area: v[4],
    empresa: v[5], usuario: v[6], destEmail: v[7], resumo: v[8], pdfNome: v[9],
    fileId: fid === "EXPIRADO" ? "" : fid, expirado: fid === "EXPIRADO", login: v[11] || "",
    naoConformes: v[13] ? String(v[13]).split(" | ") : [],
    horaIni: hhmm_(v[14]), horaFim: hhmm_(v[15]), motivo: String(v[16] || ""),
    liberacao: String(v[17] || ""), liberadoPor: String(v[18] || ""), motivoNaoApto: String(v[20] || "")
  };
}

// Histórico de UMA empresa (tela de auditoria do Administrador)
function listAuditoria(data) {
  if (!requireAdmin(data)) return jsonOut({ error: "não autorizado" });
  var empresaNome = (data.empresa || "").toString().trim();
  if (!empresaNome) return jsonOut({ error: "empresa não informada" });
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName(slugSheetTabName(empresaNome));
  if (!sheet) return jsonOut({ ok: true, rows: [] });
  ensureAuditHeaders(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonOut({ ok: true, rows: [] });
  var values = sheet.getRange(2, 1, lastRow - 1, AUDIT_HEADERS.length).getValues();
  var fromTs = data.from ? new Date(data.from + "T00:00:00").getTime() : null;
  var toTs = data.to ? new Date(data.to + "T23:59:59").getTime() : null;
  var tz = Session.getScriptTimeZone();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var o = rowToObj(values[i], tz);
    if (fromTs !== null && o.ts < fromTs) continue;
    if (toTs !== null && o.ts > toTs) continue;
    rows.push(o);
  }
  rows.sort(function (a, b) { return b.ts - a.ts; });
  return jsonOut({ ok: true, rows: rows });
}

// "Apagar" nunca remove de vez: move a linha para a aba "Excluídos".
function moveRowsToExcluidos(ss, sheet, rowIndexes) {
  if (!sheet || !rowIndexes.length) return 0;
  var excl = ss.getSheetByName("Excluídos");
  if (!excl) {
    excl = ss.insertSheet("Excluídos");
    excl.getRange(1, 1, 1, AUDIT_HEADERS.length + 2).setValues([AUDIT_HEADERS.concat(["Apagado em", "Apagado da aba"])]);
    excl.setFrozenRows(1);
  } else if (excl.getRange(1, AUDIT_HEADERS.length).getValue() !== AUDIT_ULTIMA) {
    excl.getRange(1, 1, 1, AUDIT_HEADERS.length + 2).setValues([AUDIT_HEADERS.concat(["Apagado em", "Apagado da aba"])]);
  }
  rowIndexes.sort(function (a, b) { return b - a; });
  var moved = 0;
  for (var i = 0; i < rowIndexes.length; i++) {
    var r = rowIndexes[i];
    var vals = sheet.getRange(r, 1, 1, AUDIT_HEADERS.length).getValues()[0];
    excl.appendRow(vals.concat([new Date(), sheet.getName()]));
    sheet.deleteRow(r);
    moved++;
  }
  return moved;
}

function deleteAuditoria(data) {
  if (!requireAdmin(data)) return jsonOut({ error: "não autorizado" });
  var empresaNome = (data.empresa || "").toString().trim();
  if (!empresaNome) return jsonOut({ error: "empresa não informada" });
  if (data.mode !== "rows" && data.mode !== "period") return jsonOut({ error: "modo de exclusão inválido" });

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ss = getOrCreateSpreadsheet();
    var empSheet = ss.getSheetByName(slugSheetTabName(empresaNome));
    var genSheet = ss.getSheetByName("Registros");

    var idSet = null;
    if (data.mode === "rows") {
      idSet = {};
      (data.ids || []).forEach(function (id) { idSet[String(id)] = true; });
      if (!Object.keys(idSet).length) return jsonOut({ error: "nenhum documento selecionado" });
    }
    var fromTs = data.from ? new Date(data.from + "T00:00:00").getTime() : null;
    var toTs = data.to ? new Date(data.to + "T23:59:59").getTime() : null;
    if (data.mode === "period" && fromTs === null && toTs === null) {
      return jsonOut({ error: "informe ao menos uma data" });
    }

    function collectMatches(sheet) {
      if (!sheet) return [];
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];
      var values = sheet.getRange(2, 1, lastRow - 1, AUDIT_HEADERS.length).getValues();
      var matches = [];
      for (var i = 0; i < values.length; i++) {
        var v = values[i];
        if (String(v[5] || "").trim() !== empresaNome) continue;
        var ts = v[0] instanceof Date ? v[0].getTime() : new Date(v[0]).getTime();
        var isMatch = data.mode === "rows"
          ? !!idSet[String(ts)]
          : (fromTs === null || ts >= fromTs) && (toTs === null || ts <= toTs);
        if (isMatch) matches.push(i + 2);
      }
      return matches;
    }

    var deletedFromEmpresa = moveRowsToExcluidos(ss, empSheet, collectMatches(empSheet));
    moveRowsToExcluidos(ss, genSheet, collectMatches(genSheet));
    return jsonOut({ ok: true, deleted: deletedFromEmpresa });
  } catch (err) {
    return jsonOut({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}

/* ============ PAINEL DO GESTOR ============ */
// Admin: informa "empresa". Gestor: sempre a empresa dele.
function listPainel(data) {
  var caller = authCaller(data);
  if (!caller || (caller.role !== "admin" && caller.role !== "gestor")) return jsonOut({ error: "não autorizado" });
  var empresaNome = caller.role === "gestor" ? caller.empresa : (data.empresa || "").toString().trim();
  if (!empresaNome) return jsonOut({ error: "empresa não informada" });
  var slug = slugEmpresa(empresaNome);
  var tz = Session.getScriptTimeZone();
  var agora = new Date();
  var dias = Math.min(Math.max(parseInt(data.dias, 10) || 30, 1), 365);
  var desde = agora.getTime() - dias * 86400000;
  var inicioHoje = new Date(Utilities.formatDate(agora, tz, "yyyy/MM/dd") + " 00:00:00").getTime();

  // Documentos da empresa no período
  var docs = [];
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName(slugSheetTabName(empresaNome));
  if (sheet) {
    ensureAuditHeaders(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var from = Math.max(2, lastRow - 1500);
      var values = sheet.getRange(from, 1, lastRow - from + 1, AUDIT_HEADERS.length).getValues();
      for (var i = values.length - 1; i >= 0; i--) {
        var o = rowToObj(values[i], tz);
        if (o.ts < desde) continue;
        docs.push(o);
      }
    }
  }
  docs.sort(function (a, b) { return b.ts - a.ts; });
  if (docs.length > 600) docs = docs.slice(0, 600);

  // Técnicos da empresa + resumo de atividade
  var fotos = getFotosMap();
  var props = PropertiesService.getScriptProperties().getProperties();
  var tecnicos = [];
  for (var k in props) {
    if (k.indexOf("user_") !== 0) continue;
    try {
      var rec = JSON.parse(props[k]);
      if (rec.role !== "tecnico" || slugEmpresa(rec.empresa || "") !== slug) continue;
      var uname = k.substring(5);
      var ultimo = null, hoje = 0, total = 0;
      for (var j = 0; j < docs.length; j++) {
        if (docs[j].login !== uname) continue;
        total++;
        if (!ultimo) ultimo = docs[j];
        if (docs[j].ts >= inicioHoje) hoje++;
      }
      tecnicos.push({
        username: uname, nome: rec.nome || uname, foto: fotos[uname] || "",
        ultimoTs: ultimo ? ultimo.ts : 0, ultimoTipo: ultimo ? ultimo.tipoDocumento : "",
        ultimoNumero: ultimo ? ultimo.numero : "", hoje: hoje, total: total
      });
    } catch (e) { /* ignora registro corrompido */ }
  }
  tecnicos.sort(function (a, b) { return (b.ultimoTs || 0) - (a.ultimoTs || 0) || String(a.nome).localeCompare(String(b.nome)); });

  return jsonOut({
    ok: true, empresa: empresaNome, serverNow: agora.getTime(), inicioHoje: inicioHoje,
    tecnicos: tecnicos, docs: docs
  });
}

// Entrega o PDF (base64) — só para quem tem direito: admin, gestor da
// empresa do documento ou o próprio técnico que enviou.
function getPdf(data) {
  var caller = authCaller(data);
  if (!caller) return jsonOut({ error: "não autorizado" });
  var fileId = (data.fileId || "").toString();
  if (!fileId || fileId === "EXPIRADO") return jsonOut({ error: "PDF não disponível (já passou do prazo de guarda)" });
  var sheet = getOrCreateAuditSheet();
  var last = sheet.getLastRow();
  var cell = last < 2 ? null : sheet.getRange(2, 11, last - 1, 1).createTextFinder(fileId).matchEntireCell(true).findNext();
  var empresaDoc, loginDoc;
  if (cell) {
    var v = sheet.getRange(cell.getRow(), 1, 1, AUDIT_HEADERS.length).getValues()[0];
    empresaDoc = String(v[5] || ""); loginDoc = String(v[11] || "");
  } else {
    // PDF que ainda espera liberação (ou já foi decidido): está na aba "Liberações".
    var lib = acharLiberacaoPorArquivo_(fileId);
    if (!lib) return jsonOut({ error: "documento não encontrado" });
    empresaDoc = lib.empresa; loginDoc = lib.login;
  }
  var ok = caller.role === "admin" ||
    ((caller.role === "gestor" || caller.role === "tst") && slugEmpresa(caller.empresa) === slugEmpresa(empresaDoc)) ||
    (caller.role === "tecnico" && caller.username === loginDoc);
  if (!ok) return jsonOut({ error: "não autorizado" });
  try {
    var blob = DriveApp.getFileById(fileId).getBlob();
    return jsonOut({ ok: true, nome: blob.getName(), base64: Utilities.base64Encode(blob.getBytes()) });
  } catch (e) {
    return jsonOut({ error: "arquivo não encontrado no Drive" });
  }
}

/* ============ GUARDA DE PDFs (24 meses) ============ */
// Roda sozinha toda semana (ver instalarGatilhos). Apaga do Drive os PDFs
// mais antigos que RETENCAO_MESES e marca "EXPIRADO" na planilha. O registro
// (quem, quando, qual documento) continua na planilha.
function aplicarRetencao() {
  var ss = getOrCreateSpreadsheet();
  var corte = new Date();
  corte.setMonth(corte.getMonth() - RETENCAO_MESES);
  var corteTs = corte.getTime();
  var sheets = ss.getSheets();
  var apagados = 0;
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var nome = sh.getName();
    if (nome === "Fotos" || nome === "Aceites" || nome === "Excluídos" || nome === LIB_ABA || nome === "Carimbos") continue;
    var last = sh.getLastRow();
    if (last < 2 || sh.getLastColumn() < 11) continue;
    var vals = sh.getRange(2, 1, last - 1, 11).getValues();
    for (var i = 0; i < vals.length; i++) {
      var ts = vals[i][0] instanceof Date ? vals[i][0].getTime() : new Date(vals[i][0]).getTime();
      var fid = String(vals[i][10] || "");
      if (!fid || fid === "EXPIRADO" || isNaN(ts) || ts >= corteTs) continue;
      try { DriveApp.getFileById(fid).setTrashed(true); apagados++; } catch (e) { /* já não existe */ }
      sh.getRange(i + 2, 11).setValue("EXPIRADO");
    }
  }
  Logger.log("Retenção: " + apagados + " PDF(s) removido(s) do Drive.");
}

// Rode UMA vez no editor (botão ▶) para criar os gatilhos automáticos
// (limpeza semanal dos PDFs antigos e, se configurado, o lembrete das 7h).
function instalarGatilhos() {
  var existentes = ScriptApp.getProjectTriggers();
  existentes.forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === "aplicarRetencao" || f === "enviarLembretesDiarios" || f === "fazerBackupDiario" || f === "enviarRelatoriosMensais" || f === "lembrarLiberacoesPendentes") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("aplicarRetencao").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
  ScriptApp.newTrigger("fazerBackupDiario").timeBased().everyDays(1).atHour(3).create();
  ScriptApp.newTrigger("enviarRelatoriosMensais").timeBased().onMonthDay(1).atHour(7).create();
  // O lembrete diário das 7h só é ligado se as chaves do Firebase estiverem
  // nas Propriedades do script (fcm_client_email e fcm_private_key).
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("fcm_client_email") && props.getProperty("fcm_private_key")) {
    ScriptApp.newTrigger("enviarLembretesDiarios").timeBased().everyDays(1).atHour(7).create();
    // Lembra os gestores dos documentos que esperam liberação há muito tempo.
    ScriptApp.newTrigger("lembrarLiberacoesPendentes").timeBased().everyMinutes(5).create();
  }
}

// Rode UMA vez no editor para o Google pedir as permissões (Drive, Planilhas, Gmail).
function autorizar() {
  getRaiz();
  getOrCreateAuditSheet();
  GmailApp.getAliases();
  Logger.log("Permissões OK. Pasta dos PDFs: " + getRaiz().getUrl());
}

/* ============ USUÁRIOS E LOGIN (admin / gestor / tecnico) ============ */
// Cada pessoa tem o seu login. O servidor guarda, por usuário: perfil, nome,
// empresa, hash da senha, e-mail de recuperação (opcional), se precisa trocar
// a senha no primeiro acesso e qual versão dos Termos já aceitou.
//
// Perfis:
//   admin   — acesso total (você)
//   gestor  — vê o painel da PRÓPRIA empresa, baixa PDFs, cadastra técnicos
//   tecnico — preenche e envia os documentos
//   tst     — Técnico de Segurança: avalia e libera os documentos da empresa
//             (só com a liberação ligada; cadastrado pelo gestor)
//   (consulta = login antigo da empresa, DESATIVADO)

function sha256Hex(str) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return digest.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

function textOut(s) {
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function normUser(s) { return (s || "").toString().trim().toLowerCase(); }

function getUserRec(uname) {
  var raw = PropertiesService.getScriptProperties().getProperty("user_" + uname);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

// A conta "admin" tem senha própria (primeiro acesso: admin0110 — TROQUE!).
function senhaAdminEhPadrao() {
  return getAdminPasswordHash() === sha256Hex("docseg-user:admin0110");
}
function getAdminPasswordHash() {
  var stored = PropertiesService.getScriptProperties().getProperty("admin_password_hash");
  if (stored) return stored;
  return sha256Hex("docseg-user:admin0110");
}

// Identifica quem está chamando. Aceita username/passwordHash (usuários) ou
// adminUsername/adminPasswordHash (telas de administração do app).
// Devolve { username, role, nome, empresa, rec } ou null.
function authCaller(params) {
  var uname = normUser(params.adminUsername || params.username);
  var hash = params.adminPasswordHash || params.passwordHash || "";
  if (!uname || !hash) return null;
  if (uname === "admin") {
    if (hash !== getAdminPasswordHash()) return null;
    return { username: "admin", role: "admin", nome: "Administrador", empresa: "", rec: null };
  }
  var rec = getUserRec(uname);
  if (!rec || rec.passwordHash !== hash) return null;
  if (rec.role !== "admin" && rec.role !== "gestor" && rec.role !== "tecnico" && rec.role !== "tst") return null; // "consulta" = desativado
  if (empresaSuspensa(rec.empresa)) return null; // empresa inadimplente: ninguém dela usa o sistema
  return { username: uname, role: rec.role, nome: rec.nome || uname, empresa: rec.empresa || "", rec: rec };
}

// Empresa suspensa (inadimplência): login e todas as ações dos usuários dela são recusados.
function empresaSuspensa(nome) {
  if (!nome) return false;
  var ex = getEmpresaExtras(nome);
  return !!(ex.suspenso || ex.contratoVencido);
}

// Contrato (licença de uso). Datas em "AAAA-MM-DD", no fuso de Salvador.
// Vale até o fim do dia "contratoFim" (+ dias de carência). Sem data = sem vencimento.
var TZ_CONTRATO = "America/Bahia";
function hojeStr() { return Utilities.formatDate(new Date(), TZ_CONTRATO, "yyyy-MM-dd"); }
function diasEntre(aStr, bStr) { // b - a, em dias
  var pa = aStr.split("-"), pb = bStr.split("-");
  return Math.round((Date.UTC(+pb[0], +pb[1] - 1, +pb[2]) - Date.UTC(+pa[0], +pa[1] - 1, +pa[2])) / 86400000);
}
function dataValida(str) { return /^\d{4}-\d{2}-\d{2}$/.test(str || "") && !isNaN(Date.parse(str + "T00:00:00Z")); }
// Devolve { vencido, diasRestantes (null se sem data; negativo = já passou), limite }
function contratoInfo(rec) {
  var fim = rec && rec.contratoFim;
  if (!dataValida(fim)) return { vencido: false, diasRestantes: null };
  var dias = diasEntre(hojeStr(), fim);
  var carencia = Math.max(0, parseInt(rec.carenciaDias, 10) || 0);
  return { vencido: dias + carencia < 0, diasRestantes: dias };
}
var MSG_VENCIDO_PADRAO = "Acesso suspenso temporariamente. Entre em contato com o DocSeg Digital para renovar.";
var MSG_SUSPENSO_PADRAO = "Acesso suspenso. Entre em contato com o DocSeg Digital para regularizar.";

function requireAdmin(params) {
  var c = authCaller(params);
  return !!c && c.role === "admin";
}

function changeAdminPassword(data) {
  var ok = authCaller({ username: "admin", passwordHash: data.currentPasswordHash || "" });
  if (!ok) return jsonOut({ error: "senha atual incorreta" });
  var newHash = (data.newPasswordHash || "").toString();
  if (!newHash) return jsonOut({ error: "informe a nova senha" });
  if (newHash === sha256Hex("docseg-user:admin0110")) return jsonOut({ error: "escolha uma senha diferente da padrão" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    PropertiesService.getScriptProperties().setProperty("admin_password_hash", newHash);
    return jsonOut({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function loginUser(username, passwordHash) {
  var uname = normUser(username);
  if (!uname) return jsonOut({ error: "informe o usuário" });
  if (uname === "admin") {
    if (passwordHash !== getAdminPasswordHash()) return jsonOut({ error: "senha incorreta" });
    return jsonOut({
      ok: true, role: "admin", username: "admin", nome: "Administrador", empresa: "", cfg: {},
      mustChange: false, termsOk: true, termsVersion: TERMS_VERSION, foto: "",
      perguntas: {}, empresaAviso: "", empresaLogoDataUrl: "", empresaLogoDataUrl2: "", empresaPtPrefix: "",
      senhaPadrao: senhaAdminEhPadrao(), sons: getSons_(),
      serverNow: Date.now()
    });
  }
  var rec = getUserRec(uname);
  if (!rec) return jsonOut({ error: "usuario nao encontrado" });
  if (rec.role === "consulta") {
    return jsonOut({ error: "Este acesso da empresa foi desativado. Peça ao responsável o seu login individual." });
  }
  if (rec.passwordHash !== passwordHash) return jsonOut({ error: "senha incorreta" });
  var extras = getEmpresaExtras(rec.empresa || "");
  if (extras.suspenso) {
    return jsonOut({ error: "acesso suspenso", suspenso: true, motivo: "suspenso", suspensoMsg: extras.suspensoMsg || MSG_SUSPENSO_PADRAO });
  }
  if (extras.contratoVencido) {
    return jsonOut({ error: "licença vencida", suspenso: true, motivo: "vencido", vencimento: extras.contratoFim,
      suspensoMsg: extras.suspensoMsg || MSG_VENCIDO_PADRAO });
  }
  return jsonOut({
    ok: true, role: rec.role, username: uname, nome: rec.nome || uname, empresa: rec.empresa || "",
    cfg: rec.cfg || {}, mustChange: !!rec.mustChange,
    termsOk: rec.termsVersion === TERMS_VERSION, termsVersion: TERMS_VERSION,
    foto: getFoto(uname),
    perguntas: extras.perguntas, empresaDocumentos: extras.documentos, empresaAviso: extras.aviso, empresaAvisoEm: extras.avisoEm,
    empresaAvisoGestor: rec.role === "gestor" ? extras.avisoGestor : "",
    contratoFim: rec.role === "gestor" ? extras.contratoFim : "",
    empresaTeste: rec.role === "gestor" ? !!extras.teste : false,
    alertaHoras: extras.alertaHoras,
    contratoAvisoDias: rec.role === "gestor" ? extras.avisoDias : 0,
    contratoCarencia: rec.role === "gestor" ? extras.carenciaDias : 0,
    empresaLogoDataUrl: extras.logoDataUrl,
    empresaLogoDataUrl2: extras.logoDataUrl2,
    empresaThemeColor: extras.themeColor, empresaThemeColor2: extras.themeColor2,
    empresaPtPrefix: extras.ptPrefix,
    empresaLiberacao: extras.liberacao,
    carimbo: (rec.role === "gestor" || rec.role === "tst") ? carimboResumo_(uname, rec) : null,
    empresaTemTst: rec.role === "gestor" ? empresaTemTst_(rec.empresa) : false,
    sons: getSons_(),
    serverNow: Date.now()
  });
}

// Troca de senha pelo próprio usuário (também usada no primeiro acesso).
function changePassword(data) {
  var uname = normUser(data.username);
  if (uname === "admin") {
    return changeAdminPassword({ currentPasswordHash: data.currentPasswordHash, newPasswordHash: data.newPasswordHash });
  }
  var newHash = (data.newPasswordHash || "").toString();
  if (!newHash) return jsonOut({ error: "informe a nova senha" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var rec = getUserRec(uname);
    if (!rec || rec.passwordHash !== (data.currentPasswordHash || "")) return jsonOut({ error: "senha atual incorreta" });
    rec.passwordHash = newHash;
    rec.mustChange = false;
    rec.updatedAt = new Date().toISOString();
    PropertiesService.getScriptProperties().setProperty("user_" + uname, JSON.stringify(rec));
    return jsonOut({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Aceite dos Termos de Uso — guarda versão, data e hora no cadastro e numa aba "Aceites".
function acceptTerms(data) {
  var caller = authCaller(data);
  if (!caller || caller.role === "admin") return jsonOut({ error: "não autorizado" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var rec = getUserRec(caller.username);
    rec.termsVersion = TERMS_VERSION;
    rec.termsAt = new Date().toISOString();
    PropertiesService.getScriptProperties().setProperty("user_" + caller.username, JSON.stringify(rec));
    var sh = getOrCreateSheet(getOrCreateSpreadsheet(), "Aceites",
      ["Data/Hora", "Login", "Nome", "Empresa", "Versão dos termos"]);
    sh.appendRow([new Date(), caller.username, caller.nome, caller.empresa, TERMS_VERSION]);
    return jsonOut({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function sameEmpresa(a, b) {
  var sa = slugEmpresa(a), sb = slugEmpresa(b);
  return !!sa && sa === sb;
}

function listUsers(params) {
  var caller = authCaller(params);
  if (!caller || (caller.role !== "admin" && caller.role !== "gestor" && caller.role !== "tst")) return jsonOut({ error: "não autorizado" });
  var all = PropertiesService.getScriptProperties().getProperties();
  var fotos = getFotosMap();
  var users = [];
  for (var k in all) {
    if (k.indexOf("user_") !== 0) continue;
    try {
      var rec = JSON.parse(all[k]);
      if (caller.role === "gestor" && !(equipeDoGestor_(rec) && sameEmpresa(rec.empresa, caller.empresa))) continue;
      if (caller.role === "tst" && !(rec.role === "tecnico" && sameEmpresa(rec.empresa, caller.empresa))) continue; // só para as mensagens
      var uname = k.substring(5);
      users.push({
        username: uname, role: rec.role, nome: rec.nome || "", empresa: rec.empresa || "",
        email: rec.email || "", cfg: rec.cfg || {}, mustChange: !!rec.mustChange,
        termsAt: rec.termsAt || "", hasFoto: !!fotos[uname], legado: rec.role === "consulta"
      });
    } catch (e) { /* ignora registro corrompido */ }
  }
  users.sort(function (a, b) { return String(a.empresa + a.nome).localeCompare(String(b.empresa + b.nome)); });
  return jsonOut({ ok: true, users: users });
}

function saveUser(data) {
  var caller = authCaller(data);
  if (!caller || (caller.role !== "admin" && caller.role !== "gestor")) return jsonOut({ error: "não autorizado" });
  var username = normUser(data.username);
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) {
    return jsonOut({ error: "login inválido: use de 3 a 30 letras minúsculas, números, ponto, hífen ou _ (sem espaços)." });
  }
  if (username === "admin") return jsonOut({ error: "o login \"admin\" é reservado." });
  var role = data.role;
  if (role !== "admin" && role !== "gestor" && role !== "tecnico" && role !== "tst") return jsonOut({ error: "perfil inválido" });
  var nome = (data.nome || "").toString().trim();
  var empresa = (data.empresa || "").toString().trim();
  var email = (data.email || "").toString().trim();
  if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return jsonOut({ error: "e-mail inválido" });

  if (caller.role === "gestor") {
    if (role !== "tecnico" && role !== "tst") return jsonOut({ error: "o gestor só cadastra colaboradores e técnicos de segurança" });
    if (role === "tst" && !getEmpresaExtras(caller.empresa).liberacao.ativo) return jsonOut({ error: "o Técnico de Segurança só existe com a liberação de documentos ligada" });
    empresa = caller.empresa; // sempre a empresa dele
  }
  if (role !== "admin") {
    if (!nome) return jsonOut({ error: "informe o nome completo" });
    if (!empresa) return jsonOut({ error: "selecione a empresa" });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var key = "user_" + username;
    var existing = getUserRec(username);
    if (existing && caller.role === "gestor" &&
        !(equipeDoGestor_(existing) && sameEmpresa(existing.empresa, caller.empresa))) {
      return jsonOut({ error: "este login pertence a outra empresa" });
    }
    var rec = existing || {};
    rec.role = role;
    rec.nome = nome;
    rec.empresa = role === "admin" ? "" : empresa;
    rec.email = email;
    rec.cfg = { empresaPadrao: rec.empresa, tecNome: nome };
    if (data.passwordHash) {
      rec.passwordHash = data.passwordHash;
      rec.mustChange = true;            // senha provisória: troca no primeiro acesso
    }
    if (!rec.passwordHash) return jsonOut({ error: "defina uma senha provisória para este usuário" });
    rec.updatedAt = new Date().toISOString();
    props.setProperty(key, JSON.stringify(rec));
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function deleteUser(data) {
  var caller = authCaller(data);
  if (!caller || (caller.role !== "admin" && caller.role !== "gestor")) return jsonOut({ error: "não autorizado" });
  var username = normUser(data.username);
  if (!username || username === "admin") return jsonOut({ error: "usuário inválido" });
  var rec = getUserRec(username);
  if (!rec) return jsonOut({ ok: true });
  if (caller.role === "gestor" && !(equipeDoGestor_(rec) && sameEmpresa(rec.empresa, caller.empresa))) {
    return jsonOut({ error: "não autorizado" });
  }
  PropertiesService.getScriptProperties().deleteProperty("user_" + username);
  setFoto(username, "");
  return jsonOut({ ok: true });
}

/* ============ FOTO DO TÉCNICO ============ */
// Fica numa aba "Fotos" da planilha (miniatura pequena, ~200 px) — não no
// "cofre" do script, que tem limite de espaço.
function fotosSheet() {
  return getOrCreateSheet(getOrCreateSpreadsheet(), "Fotos", ["Login", "Foto (data URL)", "Atualizada em"]);
}

function getFotosMap() {
  var sh = fotosSheet();
  var last = sh.getLastRow();
  var map = {};
  if (last < 2) return map;
  var vals = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < vals.length; i++) if (vals[i][0]) map[String(vals[i][0])] = String(vals[i][1] || "");
  return map;
}

function getFoto(uname) {
  var sh = fotosSheet();
  var last = sh.getLastRow();
  if (last < 2) return "";
  var cell = sh.getRange(2, 1, last - 1, 1).createTextFinder(uname).matchEntireCell(true).findNext();
  return cell ? String(sh.getRange(cell.getRow(), 2).getValue() || "") : "";
}

function setFoto(uname, dataUrl) {
  var sh = fotosSheet();
  var last = sh.getLastRow();
  var cell = last < 2 ? null : sh.getRange(2, 1, last - 1, 1).createTextFinder(uname).matchEntireCell(true).findNext();
  if (!dataUrl) { if (cell) sh.deleteRow(cell.getRow()); return; }
  if (cell) {
    sh.getRange(cell.getRow(), 2, 1, 2).setValues([[dataUrl, new Date()]]);
  } else {
    sh.appendRow([uname, dataUrl, new Date()]);
  }
}

// Quem pode trocar: o próprio usuário, o admin e o gestor (técnicos da empresa dele).
function saveFoto(data) {
  var caller = authCaller(data);
  if (!caller) return jsonOut({ error: "não autorizado" });
  var target = normUser(data.target || caller.username);
  var foto = (data.foto || "").toString();
  if (foto && (foto.indexOf("data:image/") !== 0 || foto.length > 45000)) {
    return jsonOut({ error: "foto inválida ou muito grande" });
  }
  var allowed = target === caller.username || caller.role === "admin";
  if (!allowed && caller.role === "gestor") {
    var rec = getUserRec(target);
    allowed = !!rec && equipeDoGestor_(rec) && sameEmpresa(rec.empresa, caller.empresa);
  }
  if (!allowed) return jsonOut({ error: "não autorizado" });
  if (target !== "admin" && !getUserRec(target)) return jsonOut({ error: "usuário não encontrado" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    setFoto(target, foto);
    return jsonOut({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/* ============ EMPRESAS E PERGUNTAS PERSONALIZADAS (Administrador) ============ */
function slugEmpresa(nome) {
  return (nome || "").toString().trim().toLowerCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function saveEmpresa(data) {
  if (!requireAdmin(data)) return jsonOut({ error: "não autorizado" });
  var nome = (data.nome || "").toString().trim();
  if (!nome) return jsonOut({ error: "informe o nome da empresa" });
  var slug = slugEmpresa(nome);
  if (!slug) return jsonOut({ error: "nome de empresa inválido" });
  // O envio por e-mail é definido pelo gestor (saveEmailEmpresa); só muda aqui se vier no pedido.
  var emailDestino = data.emailDestino === undefined ? undefined : (data.emailDestino || "").toString().trim();
  var emailAtivo = data.emailAtivo === undefined ? undefined : !!data.emailAtivo;
  if (emailAtivo && !emailDestino) return jsonOut({ error: "informe o(s) e-mail(s) de destino ou desligue o envio por e-mail" });
  var driveEmails = parseEmails(data.driveEmails);
  var contratoInicio = (data.contratoInicio || "").toString().trim();
  var contratoFim = (data.contratoFim || "").toString().trim();
  if (contratoInicio && !dataValida(contratoInicio)) return jsonOut({ error: "data de início do contrato inválida" });
  if (contratoFim && !dataValida(contratoFim)) return jsonOut({ error: "data de validade do contrato inválida" });
  if (contratoInicio && contratoFim && contratoFim < contratoInicio) return jsonOut({ error: "a validade não pode ser antes do início do contrato" });
  var avisoDias = Math.min(90, Math.max(1, parseInt(data.avisoDias, 10) || 30));
  var carenciaDias = Math.min(30, Math.max(0, parseInt(data.carenciaDias, 10) || 0));
  var documentosNovos = null;
  if (Array.isArray(data.documentos)) {
    documentosNovos = DOC_TIPOS.filter(function (k) { return data.documentos.indexOf(k) !== -1; });
    if (!documentosNovos.length) return jsonOut({ error: "a empresa precisa ter pelo menos 1 documento contratado" });
  }
  var teste = (data.teste === undefined) ? undefined : !!data.teste;
  var alertaHoras = (data.alertaHoras === undefined) ? undefined : Math.min(Math.max(parseInt(data.alertaHoras, 10) || 24, 1), 720);
  var relatorioMensal = (data.relatorioMensal === undefined) ? undefined : !!data.relatorioMensal;
  var suspenso = !!data.suspenso;
  var suspensoMsg = (data.suspensoMsg || "").toString().trim();
  var liberacao = (data.liberacao && typeof data.liberacao === "object") ? normLiberacao_(data.liberacao) : undefined;
  if (liberacao && liberacao.ativo && !liberacao.docs.length) return jsonOut({ error: "escolha ao menos 1 documento que exige liberação, ou desligue a liberação" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var anterior = {};
    try { anterior = JSON.parse(props.getProperty("empresa_" + slug) || "{}"); } catch (e0) { anterior = {}; }
    var sync = { shared: [], falhas: [], url: "" };
    try {
      sync = sincronizarAcessoPasta(nome, driveEmails, anterior.driveShared || []);
    } catch (e1) {
      sync.falhas.push("não foi possível compartilhar a pasta: " + e1.message);
      sync.shared = anterior.driveShared || [];
    }
    var rec = {
      nome: nome,
      perguntas: (data.perguntas && typeof data.perguntas === "object") ? data.perguntas : (anterior.perguntas || {}),   // perguntas agora são editadas em "Questionários" (savePerguntas)
      documentos: documentosNovos || docsDaEmpresa(anterior),
      teste: (teste === undefined) ? !!anterior.teste : teste,
      relatorioMensal: (relatorioMensal === undefined) ? !!anterior.relatorioMensal : relatorioMensal,
      alertaHoras: (alertaHoras === undefined) ? (anterior.alertaHoras || 24) : alertaHoras,
      aviso: (anterior.aviso || "").toString(),            // aviso aos técnicos: só o GESTOR edita
      avisoEm: anterior.avisoEm || 0,
      avisoGestor: (data.avisoGestor || "").toString(),    // aviso ao gestor: só o ADMIN edita
      suspenso: suspenso,
      suspensoMsg: suspensoMsg,
      contratoInicio: contratoInicio,
      contratoFim: contratoFim,
      avisoDias: avisoDias,
      carenciaDias: carenciaDias,
      themeColor: (data.themeColor || "").toString(),
      themeColor2: (data.themeColor2 || "").toString(),
      ptPrefix: (data.ptPrefix || "").toString(),
      emailAtivo: emailAtivo === undefined ? !!anterior.emailAtivo : emailAtivo,
      emailDestino: emailDestino === undefined ? (anterior.emailDestino || "") : emailDestino,
      driveEmails: driveEmails.join(", "),
      driveShared: sync.shared,
      driveFolderUrl: sync.url || anterior.driveFolderUrl || "",
      liberacao: liberacao || normLiberacao_(anterior.liberacao),
      updatedAt: new Date().toISOString()
    };
    props.setProperty("empresa_" + slug, JSON.stringify(rec));
    if (typeof data.logoDataUrl === "string") {
      if (data.logoDataUrl) props.setProperty("empresa_" + slug + "_logo", data.logoDataUrl);
      else props.deleteProperty("empresa_" + slug + "_logo");
    }
    if (typeof data.logoDataUrl2 === "string") {
      if (data.logoDataUrl2) props.setProperty("empresa_" + slug + "_logo2", data.logoDataUrl2);
      else props.deleteProperty("empresa_" + slug + "_logo2");
    }
    return jsonOut({ ok: true, driveShared: sync.shared, driveFalhas: sync.falhas, driveFolderUrl: rec.driveFolderUrl });
  } catch (err) {
    return jsonOut({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// Administrador: edita só as perguntas dos documentos de UMA empresa já cadastrada.
function savePerguntas(data) {
  if (!requireAdmin(data)) return jsonOut({ error: "não autorizado" });
  var nome = (data.nome || "").toString().trim();
  var slug = slugEmpresa(nome);
  if (!slug) return jsonOut({ error: "empresa inválida" });
  if (!data.perguntas || typeof data.perguntas !== "object") return jsonOut({ error: "perguntas inválidas" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty("empresa_" + slug);
    if (!raw) return jsonOut({ error: "empresa não encontrada — cadastre-a em Empresas" });
    var rec = JSON.parse(raw);
    rec.perguntas = data.perguntas;
    rec.updatedAt = new Date().toISOString();
    props.setProperty("empresa_" + slug, JSON.stringify(rec));
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// O gestor edita só o aviso aos técnicos da PRÓPRIA empresa.
/* ============ MENSAGENS (chat do gestor com o colaborador) ============ */
var MENSAGENS_HEADERS = ["ID", "Data/Hora", "Empresa", "De (login)", "De (nome)", "Para (login)", "Texto", "Documento em aberto", "Lida"];
function mensagensSheet() { return getOrCreateSheet(getOrCreateSpreadsheet(), "Mensagens", MENSAGENS_HEADERS); }
function msgObj_(v) {
  return { id: String(v[0]), ts: Number(v[1]), empresa: String(v[2] || ""), de: normUser(v[3]), deNome: String(v[4] || ""), para: normUser(v[5]),
           texto: String(v[6] || ""), abertoId: String(v[7] || ""), lida: v[8] === true || String(v[8]) === "1" || String(v[8]).toLowerCase() === "true" };
}
function enviarMensagem_(caller, paraLogin, paraRec, texto, abertoId) {
  var id = "M" + Date.now().toString(36) + Math.floor(Math.random() * 1000);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    mensagensSheet().appendRow([id, Date.now(), caller.empresa || paraRec.empresa || "", caller.username, caller.nome || caller.username, paraLogin, texto, abertoId || "", "0"]);
  } finally { lock.releaseLock(); }
  var enviouPush = false;
  try {
    if (pushAtivo_() && paraRec.fcmToken) {
      sendFcmPush(paraRec.fcmToken, "Mensagem de " + (caller.nome || caller.username), texto.slice(0, 140));
      enviouPush = true;
    }
  } catch (e) { Logger.log("push da mensagem falhou: " + e.message); }
  return { id: id, push: enviouPush };
}
function enviarMensagem(data) {
  var caller = authCaller(data);
  if (!caller || caller.role === "admin") return jsonOut({ error: "não autorizado" });
  var texto = (data.texto || "").toString().trim().slice(0, 500);
  if (!texto) return jsonOut({ error: "escreva a mensagem" });
  var para = normUser(data.para);
  var raw = PropertiesService.getScriptProperties().getProperty("user_" + para);
  if (!raw) return jsonOut({ error: "destinatário não encontrado" });
  var rec = JSON.parse(raw);
  if (!sameEmpresa(rec.empresa, caller.empresa)) return jsonOut({ error: "não autorizado" });
  if (caller.role === "gestor" || caller.role === "tst") {
    if (rec.role !== "tecnico") return jsonOut({ error: "não autorizado" });
  } else { // colaborador só responde a um gestor ou técnico de segurança que já falou com ele
    if (rec.role !== "gestor" && rec.role !== "tst") return jsonOut({ error: "não autorizado" });
    var sh = mensagensSheet(), n = sh.getLastRow(), falou = false;
    if (n > 1) {
      var vals = sh.getRange(2, 1, n - 1, 6).getValues();
      for (var i = 0; i < vals.length; i++) if (normUser(vals[i][3]) === para && normUser(vals[i][5]) === caller.username) { falou = true; break; }
    }
    if (!falou) return jsonOut({ error: "aguarde o gestor iniciar a conversa" });
  }
  var r = enviarMensagem_(caller, para, rec, texto, (data.abertoId || "").toString().slice(0, 40));
  return jsonOut({ ok: true, id: r.id, push: r.push });
}
function listMensagens(data) {
  var caller = authCaller(data);
  if (!caller || caller.role === "admin") return jsonOut({ error: "não autorizado" });
  var sh = mensagensSheet(), n = sh.getLastRow(), out = [], nao = 0;
  if (n > 1) {
    var vals = sh.getRange(2, 1, n - 1, MENSAGENS_HEADERS.length).getValues();
    var nomes = {};
    for (var i = vals.length - 1; i >= 0 && out.length < 300; i--) {
      var m = msgObj_(vals[i]);
      if (m.de !== caller.username && m.para !== caller.username) continue;
      if (m.para === caller.username && !m.lida) nao++;
      out.push(m);
    }
    out.reverse();
    var props = PropertiesService.getScriptProperties();
    out.forEach(function (m) {
      var outro = m.de === caller.username ? m.para : m.de;
      if (!(outro in nomes)) { try { nomes[outro] = JSON.parse(props.getProperty("user_" + outro)).nome || outro; } catch (e) { nomes[outro] = outro; } }
      m.outro = outro; m.outroNome = nomes[outro];
    });
  }
  return jsonOut({ ok: true, mensagens: out, naoLidas: nao, eu: caller.username });
}
function marcarLidas(data) {
  var caller = authCaller(data);
  if (!caller || caller.role === "admin") return jsonOut({ error: "não autorizado" });
  var com = normUser(data.com);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = mensagensSheet(), n = sh.getLastRow();
    if (n > 1) {
      var vals = sh.getRange(2, 1, n - 1, MENSAGENS_HEADERS.length).getValues();
      for (var i = 0; i < vals.length; i++) {
        if (normUser(vals[i][3]) === com && normUser(vals[i][5]) === caller.username && String(vals[i][8]) !== "1") sh.getRange(i + 2, 9).setValue("1");
      }
    }
  } finally { lock.releaseLock(); }
  return jsonOut({ ok: true });
}

/* ============ DOCUMENTOS EM ABERTO (iniciados e ainda não enviados) ============
   O celular do colaborador avisa quando começa a preencher um documento
   (docAberto) e o servidor tira o aviso quando o documento chega (saveDocuments)
   ou quando o rascunho é excluído (docFechado). O gestor consulta com listAbertos. */
var ABERTOS_HEADERS = ["ID", "Login", "Nome", "Empresa", "Tipo", "Número", "Local", "Início", "Atualizado"];
var ABERTOS_VALIDADE_MS = 14 * 86400000;
function abertosSheet() { return getOrCreateSheet(getOrCreateSpreadsheet(), "Em aberto", ABERTOS_HEADERS); }
function acharAberto_(sh, id) {
  var n = sh.getLastRow();
  if (n < 2) return 0;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return 0;
}
function removerAberto_(id) {
  id = String(id || "");
  if (!id) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = abertosSheet();
    var r = acharAberto_(sh, id);
    if (r) sh.deleteRow(r);
  } finally { lock.releaseLock(); }
}
function docAberto(data) {
  var caller = authCaller(data);
  if (!caller || caller.role === "admin" || !caller.empresa) return jsonOut({ error: "não autorizado" });
  var id = (data.abertoId || "").toString().slice(0, 40);
  if (!id) return jsonOut({ error: "id ausente" });
  var inicio = parseInt(data.inicioTs, 10);
  if (!inicio || inicio > Date.now() + 60000) inicio = Date.now();
  var linha = [id, caller.username, caller.nome || caller.username, caller.empresa,
    (data.tipoDocumento || "").toString().slice(0, 60), (data.numero || "").toString().slice(0, 40),
    (data.area || "").toString().slice(0, 160), inicio, Date.now()];
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = abertosSheet();
    var r = acharAberto_(sh, id);
    if (r) {
      var antigo = sh.getRange(r, 1, 1, ABERTOS_HEADERS.length).getValues()[0];
      if (normUser(antigo[1]) !== caller.username) return jsonOut({ error: "não autorizado" });
      linha[7] = antigo[7] || inicio;
      sh.getRange(r, 1, 1, ABERTOS_HEADERS.length).setValues([linha]);
    } else {
      sh.appendRow(linha);
    }
  } catch (err) { return jsonOut({ error: err.message }); }
  finally { lock.releaseLock(); }
  return jsonOut({ ok: true });
}
function docFechado(data) {
  var caller = authCaller(data);
  if (!caller) return jsonOut({ error: "não autorizado" });
  var id = (data.abertoId || "").toString();
  var sh = abertosSheet();
  var r = acharAberto_(sh, id);
  if (r) {
    var linha = sh.getRange(r, 1, 1, ABERTOS_HEADERS.length).getValues()[0];
    var dono = normUser(linha[1]);
    // O gestor da empresa também pode tirar um aviso que ficou preso.
    var gestorDaEmpresa = caller.role === "gestor" && sameEmpresa(linha[3], caller.empresa);
    if (caller.role !== "admin" && dono !== caller.username && !gestorDaEmpresa) return jsonOut({ error: "não autorizado" });
    removerAberto_(id);
  }
  return jsonOut({ ok: true });
}
function listAbertos(data) {
  var caller = authCaller(data);
  if (!caller || (caller.role !== "admin" && caller.role !== "gestor")) return jsonOut({ error: "não autorizado" });
  var empresaNome = caller.role === "gestor" ? caller.empresa : (data.empresa || "").toString().trim();
  if (!empresaNome) return jsonOut({ error: "empresa não informada" });
  var slug = slugEmpresa(empresaNome);
  var sh = abertosSheet();
  var n = sh.getLastRow();
  var out = [], velhas = [];
  var agora = Date.now();
  if (n > 1) {
    var vals = sh.getRange(2, 1, n - 1, ABERTOS_HEADERS.length).getValues();
    var zap = {};
    var props = PropertiesService.getScriptProperties();
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i];
      if (!v[0]) continue;
      if (agora - Number(v[8] || v[7]) > ABERTOS_VALIDADE_MS) { velhas.push(String(v[0])); continue; }
      if (slugEmpresa(String(v[3] || "")) !== slug) continue;
      var login = normUser(v[1]);
      if (!(login in zap)) {
        var raw = props.getProperty("user_" + login), z = "", tem = false;
        try { var ur = JSON.parse(raw); z = String(ur.whatsapp || ""); tem = !!ur.fcmToken; } catch (e) {}
        zap[login] = { z: z, tem: tem };
      }
      out.push({ abertoId: String(v[0]), login: login, nome: String(v[2] || ""), tipoDocumento: String(v[4] || ""),
        numero: String(v[5] || ""), area: String(v[6] || ""), inicioTs: Number(v[7]), atualizadoTs: Number(v[8]),
        push: zap[login].tem && pushAtivo_() });
    }
  }
  velhas.forEach(function (id) { try { removerAberto_(id); } catch (e) {} });
  out.sort(function (a, b) { return a.inicioTs - b.inicioTs; });
  return jsonOut({ ok: true, abertos: out, serverNow: agora, alertaHoras: getEmpresaExtras(empresaNome).alertaHoras || 24 });
}
function lembrarColaborador(data) {
  var caller = authCaller(data);
  if (!caller || (caller.role !== "admin" && caller.role !== "gestor")) return jsonOut({ error: "não autorizado" });
  var login = normUser(data.login);
  var raw = PropertiesService.getScriptProperties().getProperty("user_" + login);
  if (!raw) return jsonOut({ error: "colaborador não encontrado" });
  var rec = JSON.parse(raw);
  if (caller.role === "gestor" && !(rec.role === "tecnico" && sameEmpresa(rec.empresa, caller.empresa))) return jsonOut({ error: "não autorizado" });
  var doc = (data.documento || "documento").toString().slice(0, 60);
  var msg = enviarMensagem_(caller, login, rec, "Lembrete: finalize e envie o documento em aberto (" + doc + ").", (data.abertoId || "").toString());
  return jsonOut({ ok: true, push: msg.push, aviso: msg.push ? "" : (!pushAtivo_() ? "Mensagem entregue dentro do app. As notificações do celular ainda não foram configuradas no servidor." : "Mensagem entregue dentro do app. Este colaborador ainda não ativou as notificações do celular.") });
}


/* ============ RELATÓRIO MENSAL POR E-MAIL (PDF) ============
   No dia 1 de cada mês, às 7h, cada empresa com "Relatório mensal" ligado
   recebe um PDF com o resumo do mês anterior. Vai para o e-mail de destino
   da empresa ou, se não houver, para o e-mail dos gestores dela. */
var MESES_PT = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
function esc_(t) { return String(t == null ? "" : t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
function destinosRelatorio(empresaNome) {
  var ext = getEmpresaExtras(empresaNome);
  var lista = parseEmails(ext.emailDestino || "");
  if (lista.length) return lista;
  var props = PropertiesService.getScriptProperties().getProperties();
  var slug = slugEmpresa(empresaNome);
  Object.keys(props).forEach(function (k) {
    if (k.indexOf("user_") !== 0) return;
    try {
      var rec = JSON.parse(props[k]);
      if (rec.role === "gestor" && slugEmpresa(rec.empresa || "") === slug && rec.email) lista = lista.concat(parseEmails(rec.email));
    } catch (e) {}
  });
  return lista;
}
function resumoDoMes(empresaNome, ano, mes) { // mes: 1-12
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName(slugSheetTabName(empresaNome));
  var out = { total: 0, porTipo: { pt: 0, veiculo: 0, escada: 0, outro: 0 }, porColab: {}, naoConf: {} };
  if (!sheet) return out;
  ensureAuditHeaders(sheet);
  var n = sheet.getLastRow();
  if (n < 2) return out;
  var vals = sheet.getRange(2, 1, n - 1, AUDIT_HEADERS.length).getValues();
  for (var i = 0; i < vals.length; i++) {
    var ts = vals[i][0] instanceof Date ? vals[i][0] : new Date(vals[i][0]);
    if (isNaN(ts.getTime())) continue;
    var a = parseInt(Utilities.formatDate(ts, TZ_CONTRATO, "yyyy"), 10), m = parseInt(Utilities.formatDate(ts, TZ_CONTRATO, "MM"), 10);
    if (a !== ano || m !== mes) continue;
    out.total++;
    out.porTipo[tipoDocKey(vals[i][1])]++;
    var quem = String(vals[i][6] || vals[i][11] || "—");
    out.porColab[quem] = (out.porColab[quem] || 0) + 1;
    if (vals[i][13]) String(vals[i][13]).split(" | ").forEach(function (t) { if (t) out.naoConf[t] = (out.naoConf[t] || 0) + 1; });
  }
  return out;
}
function htmlRelatorio(empresaNome, ano, mes, r) {
  function linhas(obj, max) {
    var ks = Object.keys(obj).sort(function (x, y) { return obj[y] - obj[x]; }).slice(0, max || 50);
    if (!ks.length) return '<tr><td colspan="2" style="color:#666">Nenhum registro.</td></tr>';
    return ks.map(function (k) { return '<tr><td>' + esc_(k) + '</td><td style="text-align:right;width:70px"><b>' + obj[k] + '</b></td></tr>'; }).join("");
  }
  var tipos = { "Permissão de Trabalho": r.porTipo.pt, "Checklist de Veículo": r.porTipo.veiculo, "Checklist de Escada": r.porTipo.escada };
  if (r.porTipo.outro) tipos["Outros"] = r.porTipo.outro;
  var th = 'style="text-align:left;background:#12254B;color:#fff;padding:6px 8px;font-size:12px"';
  return '<html><body style="font-family:Arial,sans-serif;color:#111;padding:24px">' +
    '<div style="background:#12254B;color:#fff;padding:14px 18px;border-radius:8px"><div style="font-size:12px;opacity:.8">DocSeg Digital · Relatório mensal</div><div style="font-size:22px;font-weight:bold;margin-top:2px">' + esc_(empresaNome) + '</div><div style="font-size:13px;margin-top:2px">' + MESES_PT[mes - 1] + ' de ' + ano + '</div></div>' +
    '<p style="font-size:34px;margin:18px 0 0"><b>' + r.total + '</b> <span style="font-size:14px;color:#555">documentos enviados no mês</span></p>' +
    '<h3 style="margin:18px 0 6px;font-size:14px">Por tipo de documento</h3><table style="width:100%;border-collapse:collapse" border="1" bordercolor="#ddd" cellpadding="6"><tr><th ' + th + '>Tipo</th><th ' + th + '>Qtd.</th></tr>' + linhas(tipos) + '</table>' +
    '<h3 style="margin:18px 0 6px;font-size:14px">Por colaborador</h3><table style="width:100%;border-collapse:collapse" border="1" bordercolor="#ddd" cellpadding="6"><tr><th ' + th + '>Colaborador</th><th ' + th + '>Qtd.</th></tr>' + linhas(r.porColab) + '</table>' +
    '<h3 style="margin:18px 0 6px;font-size:14px">Não conformidades mais comuns</h3><table style="width:100%;border-collapse:collapse" border="1" bordercolor="#ddd" cellpadding="6"><tr><th ' + th + '>Item</th><th ' + th + '>Vezes</th></tr>' + linhas(r.naoConf, 8) + '</table>' +
    '<p style="font-size:10px;color:#777;margin-top:22px">Gerado automaticamente pelo DocSeg Digital.</p></body></html>';
}
function enviarRelatorioEmpresa(empresaNome, ano, mes) {
  var dest = destinosRelatorio(empresaNome);
  if (!dest.length) return { enviado: false, motivo: "sem e-mail de destino (cadastre na empresa ou no gestor)" };
  var r = resumoDoMes(empresaNome, ano, mes);
  var nomePdf = "Relatorio_" + slugEmpresa(empresaNome) + "_" + ano + "-" + (mes < 10 ? "0" : "") + mes + ".pdf";
  var pdf = Utilities.newBlob(htmlRelatorio(empresaNome, ano, mes, r), "text/html", "relatorio.html").getAs("application/pdf").setName(nomePdf);
  GmailApp.sendEmail(dest.join(","), "[DocSeg Digital] Relatório de " + MESES_PT[mes - 1] + "/" + ano + " - " + empresaNome,
    "Segue o resumo de " + MESES_PT[mes - 1] + " de " + ano + ": " + r.total + " documento(s) enviado(s) pela equipe.\n\nO PDF está em anexo.",
    { attachments: [pdf], name: "DocSeg Digital" });
  return { enviado: true, total: r.total, para: dest.length };
}
function mesAnterior_() {
  var agora = new Date();
  var a = parseInt(Utilities.formatDate(agora, TZ_CONTRATO, "yyyy"), 10), m = parseInt(Utilities.formatDate(agora, TZ_CONTRATO, "MM"), 10);
  m--; if (m < 1) { m = 12; a--; }
  return { ano: a, mes: m };
}
function enviarRelatoriosMensais() {
  var ref = mesAnterior_();
  var props = PropertiesService.getScriptProperties().getProperties();
  Object.keys(props).forEach(function (k) {
    if (k.indexOf("empresa_") !== 0 || /_logo2?$/.test(k)) return;
    try {
      var rec = JSON.parse(props[k]);
      if (!rec.relatorioMensal || empresaSuspensa(rec.nome)) return;
      enviarRelatorioEmpresa(rec.nome, ref.ano, ref.mes);
    } catch (e) { Logger.log("relatório falhou (" + k + "): " + e.message); }
  });
}
function relatorioAgora(data) {
  if (!requireAdmin(data)) return jsonOut({ error: "não autorizado" });
  var nome = (data.nome || "").toString().trim();
  if (!nome) return jsonOut({ error: "informe a empresa" });
  var ref = mesAnterior_();
  var m = /^(\d{4})-(\d{2})$/.exec((data.mes || "").toString());
  if (m) { ref = { ano: parseInt(m[1], 10), mes: parseInt(m[2], 10) }; }
  try {
    var r = enviarRelatorioEmpresa(nome, ref.ano, ref.mes);
    if (!r.enviado) return jsonOut({ error: r.motivo });
    return jsonOut({ ok: true, total: r.total, para: r.para, mes: MESES_PT[ref.mes - 1] + "/" + ref.ano });
  } catch (e) { return jsonOut({ error: e.message }); }
}

/* ============ BACKUP (usuários, empresas, perguntas, avisos) ============
   Guarda um arquivo .json por dia na pasta "_Backups" do Drive e mantém os
   30 mais recentes. Não inclui senhas do administrador nem chaves do Firebase. */
var BACKUP_PREFIXOS = ["user_", "empresa_"];
var BACKUP_MANTER = 30;
function pastaBackups() { return subpasta(getRaiz(), "_Backups"); }
function montarBackup() {
  var all = PropertiesService.getScriptProperties().getProperties();
  var dados = {};
  Object.keys(all).forEach(function (k) {
    for (var i = 0; i < BACKUP_PREFIXOS.length; i++) {
      if (k.indexOf(BACKUP_PREFIXOS[i]) === 0) { dados[k] = all[k]; break; }
    }
  });
  return { sistema: "DocSeg Digital", versao: 1, quando: new Date().toISOString(), chaves: Object.keys(dados).length, dados: dados };
}
function fazerBackup() {
  var b = montarBackup();
  var nome = "backup_" + Utilities.formatDate(new Date(), TZ_CONTRATO, "yyyy-MM-dd_HHmm") + ".json";
  var f = pastaBackups().createFile(Utilities.newBlob(JSON.stringify(b), "application/json", nome));
  var lista = [], it = pastaBackups().getFiles();
  while (it.hasNext()) { var x = it.next(); lista.push({ f: x, t: x.getDateCreated().getTime() }); }
  lista.sort(function (a, c) { return c.t - a.t; });
  for (var i = BACKUP_MANTER; i < lista.length; i++) { try { lista[i].f.setTrashed(true); } catch (e) {} }
  PropertiesService.getScriptProperties().setProperty("backup_ultimo", JSON.stringify({ quando: b.quando, nome: nome, chaves: b.chaves }));
  return { id: f.getId(), nome: nome, chaves: b.chaves, quando: b.quando };
}
function fazerBackupDiario() { try { fazerBackup(); } catch (e) { Logger.log("backup falhou: " + e.message); } }
function backupNow(data) {
  if (!requireAdmin(data)) return jsonOut({ error: "não autorizado" });
  try { var r = fazerBackup(); return jsonOut({ ok: true, backup: r }); } catch (e) { return jsonOut({ error: e.message }); }
}
function listBackups(data) {
  if (!requireAdmin(data)) return jsonOut({ error: "não autorizado" });
  var out = [], it = pastaBackups().getFiles();
  while (it.hasNext()) {
    var f = it.next();
    out.push({ id: f.getId(), nome: f.getName(), tamanho: f.getSize(), quando: Utilities.formatDate(f.getDateCreated(), TZ_CONTRATO, "dd/MM/yyyy HH:mm"), t: f.getDateCreated().getTime() });
  }
  out.sort(function (a, b) { return b.t - a.t; });
  var ultimo = {};
  try { ultimo = JSON.parse(PropertiesService.getScriptProperties().getProperty("backup_ultimo") || "{}"); } catch (e) {}
  return jsonOut({ ok: true, backups: out.slice(0, BACKUP_MANTER), ultimo: ultimo });
}
// Restaurar exige a senha do admin de novo e antes guarda um backup do estado atual.
function restoreBackup(data) {
  if (!requireAdmin(data)) return jsonOut({ error: "não autorizado" });
  if ((data.confirmPasswordHash || "") !== getAdminPasswordHash()) return jsonOut({ error: "senha do administrador incorreta" });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var arq = DriveApp.getFileById((data.fileId || "").toString());
    var b = JSON.parse(arq.getBlob().getDataAsString());
    if (!b || b.sistema !== "DocSeg Digital" || !b.dados) return jsonOut({ error: "arquivo de backup inválido" });
    fazerBackup();
    var props = PropertiesService.getScriptProperties(), n = 0;
    Object.keys(b.dados).forEach(function (k) {
      for (var i = 0; i < BACKUP_PREFIXOS.length; i++) {
        if (k.indexOf(BACKUP_PREFIXOS[i]) === 0) { props.setProperty(k, b.dados[k]); n++; break; }
      }
    });
    return jsonOut({ ok: true, restauradas: n });
  } catch (e) {
    return jsonOut({ error: e.message });
  } finally {
    lock.releaseLock();
  }
}

/* ============ SUPORTE (chamados) ============
   Qualquer usuário ativo abre chamados. O admin vê todos e muda o status.
   Registro na aba "Chamados" da planilha; o print vai para a pasta
   "_Suporte" do Drive; cada chamado novo avisa o e-mail de suporte. */
var SUPORTE_EMAIL = "docsegdigital@gmail.com";
var CHAMADOS_HEADERS = ["ID", "Data/Hora", "Login", "Nome", "Perfil", "Empresa", "Tipo", "Descrição", "Versão do app", "Aparelho", "Último erro de envio", "WhatsApp", "Print (link)", "Status", "Conversa", "Apagado pelo usuário"];
var CH_OCULTO = 15; // coluna (base 0) "Apagado pelo usuário": some da lista dele, o admin continua vendo

function chamadosSheet() {
  var sh = getOrCreateSheet(getOrCreateSpreadsheet(), "Chamados", CHAMADOS_HEADERS);
  try {
    if (sh.getRange(1, 15).getValue() !== "Conversa") sh.getRange(1, 15).setValue("Conversa");
    if (sh.getRange(1, CHAMADOS_HEADERS.length).getValue() !== CHAMADOS_HEADERS[CH_OCULTO]) sh.getRange(1, CHAMADOS_HEADERS.length).setValue(CHAMADOS_HEADERS[CH_OCULTO]);
  } catch (e) {}
  return sh;
}
function conversaDe_(v) {
  try { var a = JSON.parse(String(v || "[]")); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}
function chamadoRow(v, tz) {
  var dt = v[1];
  var dataTxt = (dt instanceof Date) ? Utilities.formatDate(dt, tz, "dd/MM/yyyy HH:mm") : String(dt || "");
  return { id: String(v[0]), data: dataTxt, login: String(v[2] || ""), nome: String(v[3] || ""), perfil: v[4], empresa: String(v[5] || ""), tipo: String(v[6] || ""), descricao: String(v[7] || ""),
           versao: String(v[8] || ""), aparelho: String(v[9] || ""), erro: String(v[10] || ""), whatsapp: String(v[11] || ""), print: String(v[12] || ""), status: v[13] || "aberto", conversa: conversaDe_(v[14]) };
}

function saveChamado(data) {
  var caller = authCaller(data);
  if (!caller) return jsonOut({ error: "não autorizado" });
  var tipo = (data.tipo || "").toString().slice(0, 40);
  var descricao = (data.descricao || "").toString().trim().slice(0, 2000);
  if (!descricao) return jsonOut({ error: "descreva o que aconteceu" });
  var whats = (data.whatsapp || "").toString().replace(/\D/g, "").slice(0, 15);
  var versao = (data.versao || "").toString().slice(0, 20);
  var aparelho = (data.aparelho || "").toString().slice(0, 200);
  var erro = (data.ultimoErro || "").toString().slice(0, 300);
  var id = "C" + Utilities.formatDate(new Date(), TZ_CONTRATO, "yyMMddHHmmss") + Math.floor(Math.random() * 90 + 10);
  var printUrl = "";
  var m = /^data:(image\/(?:jpeg|png));base64,(.+)$/.exec((data.printDataUrl || "").toString());
  if (m && m[2].length < 3000000) {
    try {
      var pasta = subpasta(getRaiz(), "_Suporte");
      var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], id + (m[1] === "image/png" ? ".png" : ".jpg"));
      var f = pasta.createFile(blob);
      printUrl = f.getUrl();
    } catch (e) { Logger.log("print do chamado não salvo: " + e.message); }
  }
  var perfil = caller.role === "gestor" ? "Gestor" : (caller.role === "tst" ? "Técnico de Segurança" : (caller.role === "admin" ? "Administrador" : "Colaborador"));
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    chamadosSheet().appendRow([id, new Date(), caller.username, caller.nome, perfil, caller.empresa, tipo, descricao, versao, aparelho, erro, whats, printUrl, "aberto", "[]", ""]);
  } catch (err) {
    return jsonOut({ error: err.message });
  } finally {
    lock.releaseLock();
  }
  try {
    var corpo = "Novo chamado " + id + "\n\n" +
      "Quem: " + caller.nome + " (" + caller.username + ") - " + perfil + "\n" +
      "Empresa: " + (caller.empresa || "-") + "\n" +
      "Tipo: " + tipo + "\n" +
      "WhatsApp: " + (whats || "não informado") + (whats ? "  https://wa.me/55" + whats : "") + "\n\n" +
      descricao + "\n\n" +
      "App: " + versao + " | Aparelho: " + aparelho + "\n" +
      "Último erro de envio: " + (erro || "-") + "\n" +
      "Print: " + (printUrl || "sem print");
    GmailApp.sendEmail(SUPORTE_EMAIL, "[DocSeg Suporte] " + tipo + " - " + caller.nome + (caller.empresa ? " (" + caller.empresa + ")" : ""), corpo, { name: "DocSeg Digital" });
  } catch (mailErr) { Logger.log("e-mail do chamado não enviado: " + mailErr.message); }
  return jsonOut({ ok: true, id: id });
}

function listChamados(data) {
  var caller = authCaller(data);
  if (!caller) return jsonOut({ error: "não autorizado" });
  var sh = chamadosSheet();
  var tz = Session.getScriptTimeZone();
  var n = sh.getLastRow();
  var out = [];
  if (n > 1) {
    var vals = sh.getRange(2, 1, n - 1, CHAMADOS_HEADERS.length).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (!vals[i][0]) continue;
      if (caller.role !== "admin" && (normUser(vals[i][2]) !== caller.username || String(vals[i][CH_OCULTO]) === "1")) continue;
      var c = chamadoRow(vals[i], tz);
      if (caller.role !== "admin") { c.print = ""; c.whatsapp = ""; c.erro = ""; c.aparelho = ""; }
      out.push(c);
      if (out.length >= 200) break;
    }
  }
  return jsonOut({ ok: true, chamados: out });
}

function responderChamado(data) {
  var caller = authCaller(data);
  if (!caller) return jsonOut({ error: "não autorizado" });
  var id = (data.id || "").toString();
  var texto = (data.texto || "").toString().trim().slice(0, 1000);
  if (!texto) return jsonOut({ error: "escreva a mensagem" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var dono = "", assunto = "", deAdmin = caller.role === "admin";
  try {
    var sh = chamadosSheet();
    var n = sh.getLastRow();
    if (n < 2) return jsonOut({ error: "chamado não encontrado" });
    var vals = sh.getRange(2, 1, n - 1, CHAMADOS_HEADERS.length).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) !== id) continue;
      if (!deAdmin && normUser(vals[i][2]) !== caller.username) return jsonOut({ error: "não autorizado" });
      var conv = conversaDe_(vals[i][14]);
      if (conv.length >= 60) return jsonOut({ error: "limite de mensagens neste chamado; abra um novo" });
      conv.push({ de: deAdmin ? "suporte" : "usuario", nome: deAdmin ? "Suporte DocSeg" : (caller.nome || caller.username), ts: Date.now(), texto: texto });
      sh.getRange(i + 2, 15).setValue(JSON.stringify(conv)); // coluna "Conversa"
      if (!deAdmin) sh.getRange(i + 2, 14).setValue("aberto"); // resposta do usuário reabre
      dono = normUser(vals[i][2]); assunto = String(vals[i][6] || "chamado");
      var out = chamadoRow(sh.getRange(i + 2, 1, 1, CHAMADOS_HEADERS.length).getValues()[0], Session.getScriptTimeZone());
      lock.releaseLock(); lock = null;
      if (deAdmin) {
        pushParaUsuario_(dono, "Suporte DocSeg", "Respondeu seu chamado: " + texto.slice(0, 90));
      } else {
        try { GmailApp.sendEmail(SUPORTE_EMAIL, "[DocSeg Suporte] Nova resposta no chamado " + id, caller.nome + " respondeu:\n\n" + texto, { name: "DocSeg Digital" }); } catch (e) {}
      }
      if (!deAdmin) { out.print = ""; out.whatsapp = ""; out.erro = ""; out.aparelho = ""; }
      return jsonOut({ ok: true, chamado: out });
    }
    return jsonOut({ error: "chamado não encontrado" });
  } catch (err) {
    return jsonOut({ error: err.message });
  } finally {
    if (lock) lock.releaseLock();
  }
}

function updateChamado(data) {
  if (!requireAdmin(data)) return jsonOut({ error: "não autorizado" });
  var id = (data.id || "").toString();
  var status = data.status === "resolvido" ? "resolvido" : "aberto";
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = chamadosSheet();
    var n = sh.getLastRow();
    if (n < 2) return jsonOut({ error: "chamado não encontrado" });
    var ids = sh.getRange(2, 1, n - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) {
        sh.getRange(i + 2, 14).setValue(status);
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ error: "chamado não encontrado" });
  } catch (err) {
    return jsonOut({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function saveAviso(data) {
  var caller = authCaller(data);
  if (!caller || caller.role !== "gestor") return jsonOut({ error: "não autorizado" });
  var slug = slugEmpresa(caller.empresa);
  if (!slug) return jsonOut({ error: "empresa não encontrada" });
  var aviso = (data.aviso || "").toString().trim();
  if (aviso.length > 600) return jsonOut({ error: "aviso muito longo (máximo 600 caracteres)" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty("empresa_" + slug);
    var rec = raw ? JSON.parse(raw) : { nome: caller.empresa, perguntas: {} };
    rec.aviso = aviso;
    rec.avisoEm = aviso ? Date.now() : 0; // o aviso aparece para a equipe por 24 h a partir daqui
    rec.updatedAt = new Date().toISOString();
    props.setProperty("empresa_" + slug, JSON.stringify(rec));
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// Admin vê tudo; o gestor vê só a própria empresa (sem as perguntas).
function listEmpresas(params) {
  var caller = authCaller(params);
  if (!caller || (caller.role !== "admin" && caller.role !== "gestor")) return jsonOut({ error: "não autorizado" });
  var all = PropertiesService.getScriptProperties().getProperties();
  var empresas = [];
  for (var k in all) {
    if (k.indexOf("empresa_") === 0 && k.indexOf("_logo") === -1) {
      try {
        var rec = JSON.parse(all[k]);
        var slug = k.substring(8);
        if (caller.role === "gestor" && slug !== slugEmpresa(caller.empresa)) continue;
        empresas.push({
          slug: slug,
          nome: rec.nome,
          perguntas: rec.perguntas || {},
          documentos: docsDaEmpresa(rec),
          teste: !!rec.teste,
          relatorioMensal: !!rec.relatorioMensal,
          alertaHoras: rec.alertaHoras || 24,
          aviso: rec.aviso || "",
          avisoEm: rec.avisoEm || 0,
          avisoGestor: rec.avisoGestor || "",
          suspenso: !!rec.suspenso,
          suspensoMsg: rec.suspensoMsg || "",
          contratoInicio: rec.contratoInicio || "",
          contratoFim: rec.contratoFim || "",
          avisoDias: rec.avisoDias != null ? rec.avisoDias : 30,
          carenciaDias: rec.carenciaDias || 0,
          contratoVencido: contratoInfo(rec).vencido,
          diasRestantes: contratoInfo(rec).diasRestantes,
          themeColor: rec.themeColor || "",
          themeColor2: rec.themeColor2 || "",
          ptPrefix: rec.ptPrefix || "",
          emailAtivo: !!rec.emailAtivo,
          emailDestino: rec.emailDestino || "",
          driveEmails: rec.driveEmails || "",
          driveFolderUrl: rec.driveFolderUrl || "",
          liberacao: normLiberacao_(rec.liberacao),
          logoDataUrl: all["empresa_" + slug + "_logo"] || "",
          logoDataUrl2: all["empresa_" + slug + "_logo2"] || ""
        });
      } catch (e) { /* ignora registro corrompido */ }
    }
  }
  return jsonOut({ ok: true, empresas: empresas });
}

function deleteEmpresa(data) {
  if (!requireAdmin(data)) return jsonOut({ error: "não autorizado" });
  var slug = slugEmpresa(data.nome || "");
  if (!slug) return jsonOut({ error: "nome inválido" });
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty("empresa_" + slug);
  props.deleteProperty("empresa_" + slug + "_logo");
  props.deleteProperty("empresa_" + slug + "_logo2");
  return jsonOut({ ok: true });
}

function getEmpresaExtras(nome) {
  var out = { alertaHoras: 24, teste: false, documentos: DOC_TIPOS.slice(), perguntas: {}, aviso: "", avisoEm: 0, avisoGestor: "", suspenso: false, suspensoMsg: "", contratoInicio: "", contratoFim: "", avisoDias: 30, carenciaDias: 0, contratoVencido: false, diasRestantes: null, themeColor: "", themeColor2: "", ptPrefix: "", logoDataUrl: "", logoDataUrl2: "", emailAtivo: false, emailDestino: "", liberacao: normLiberacao_(null) };
  if (!nome) return out;
  var slug = slugEmpresa(nome);
  if (!slug) return out;
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty("empresa_" + slug);
  if (raw) {
    try {
      var rec = JSON.parse(raw);
      out.perguntas = rec.perguntas || {};
      out.documentos = docsDaEmpresa(rec);
      out.teste = !!rec.teste;
      out.alertaHoras = rec.alertaHoras || 24;
      out.aviso = rec.aviso || "";
      out.avisoEm = rec.avisoEm || 0;
      out.avisoGestor = rec.avisoGestor || "";
      out.suspenso = !!rec.suspenso;
      out.suspensoMsg = rec.suspensoMsg || "";
      out.contratoInicio = rec.contratoInicio || "";
      out.contratoFim = rec.contratoFim || "";
      out.avisoDias = rec.avisoDias != null ? rec.avisoDias : 30;
      out.carenciaDias = rec.carenciaDias || 0;
      var ci = contratoInfo(rec);
      out.contratoVencido = ci.vencido;
      out.diasRestantes = ci.diasRestantes;
      out.themeColor = rec.themeColor || "";
      out.themeColor2 = rec.themeColor2 || "";
      out.ptPrefix = rec.ptPrefix || "";
      out.emailAtivo = !!rec.emailAtivo;
      out.emailDestino = rec.emailDestino || "";
      out.liberacao = normLiberacao_(rec.liberacao);
    } catch (e) { /* ignora registro corrompido */ }
  }
  out.logoDataUrl = props.getProperty("empresa_" + slug + "_logo") || "";
  out.logoDataUrl2 = props.getProperty("empresa_" + slug + "_logo2") || "";
  return out;
}

/* ============ NOTIFICAÇÕES PUSH (lembrete diário da PT) ============ */
// Guarda o token deste aparelho junto ao registro do usuário (exige o mesmo
// login/senha já usados no app, para não aceitar token de quem quer).
function saveFcmToken(data) {
  var username = (data.username || "").toString().trim().toLowerCase();
  var passwordHash = data.passwordHash || "";
  if (!username || !passwordHash) return jsonOut({ error: "não autorizado" });
  var props = PropertiesService.getScriptProperties();
  var key = "user_" + username;
  var raw = props.getProperty(key);
  if (!raw) {
    return jsonOut({ error: "cadastre este usuário em Usuários antes de ativar notificações" });
  }
  var rec = JSON.parse(raw);
  if (rec.passwordHash !== passwordHash) return jsonOut({ error: "não autorizado" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    rec.fcmToken = data.fcmToken || "";
    rec.fcmTokenUpdatedAt = new Date().toISOString();
    props.setProperty(key, JSON.stringify(rec));
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// Gera um access token OAuth2 a partir da Conta de Serviço do Firebase
// (client_email + private_key guardados no cofre do script), para poder
// chamar a API do Firebase Cloud Messaging (FCM) sem precisar de biblioteca
// externa nenhuma.
function getFcmAccessToken() {
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty("fcm_client_email");
  var privateKey = props.getProperty("fcm_private_key");
  if (!email || !privateKey) {
    throw new Error("Configure fcm_client_email e fcm_private_key nas Propriedades do Script");
  }
  privateKey = privateKey.replace(/\\n/g, "\n");

  function base64url(strOrObj) {
    var json = typeof strOrObj === "string" ? strOrObj : JSON.stringify(strOrObj);
    return Utilities.base64EncodeWebSafe(json).replace(/=+$/, "");
  }

  var now = Math.floor(Date.now() / 1000);
  var header = { alg: "RS256", typ: "JWT" };
  var claimSet = {
    iss: email,
    scope: "https://www.googleapis.com/auth/firebase.messaging",
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600
  };
  var toSign = base64url(header) + "." + base64url(claimSet);
  var signatureBytes = Utilities.computeRsaSha256Signature(toSign, privateKey);
  var signature = Utilities.base64EncodeWebSafe(signatureBytes).replace(/=+$/, "");
  var jwt = toSign + "." + signature;

  var resp = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt
    },
    muteHttpExceptions: true
  });
  var json = JSON.parse(resp.getContentText());
  if (!json.access_token) throw new Error("Falha ao obter token do Firebase: " + resp.getContentText());
  return json.access_token;
}

// Envia uma notificação push para um token específico via FCM (API v1).
function sendFcmPush(token, title, body) {
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty("fcm_project_id") || "docseg-app";
  var accessToken = getFcmAccessToken();
  var message = {
    message: {
      token: token,
      notification: { title: title, body: body },
      data: { tipo: "docseg", titulo: String(title), corpo: String(body) },
      webpush: { headers: { Urgency: "high" }, notification: { title: title, body: body, tag: "docseg-aviso", renotify: true, icon: "/icon-192.png" }, fcm_options: { link: "/" } }
    }
  };
  var resp = UrlFetchApp.fetch(
    "https://fcm.googleapis.com/v1/projects/" + projectId + "/messages:send",
    {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + accessToken },
      payload: JSON.stringify(message),
      muteHttpExceptions: true
    }
  );
  if (resp.getResponseCode() >= 300) {
    var det = resp.getContentText();
    try { var j = JSON.parse(det); det = (j.error && (j.error.message || j.error.status)) || det; } catch (e) {}
    throw new Error("Firebase respondeu " + resp.getResponseCode() + ": " + String(det).slice(0, 200));
  }
  return resp.getContentText();
}

// Push só funciona depois que as chaves do Firebase forem cadastradas nas
// Propriedades do script. Sem elas, estas funções simplesmente não fazem nada.
function pushAtivo_() {
  var props = PropertiesService.getScriptProperties();
  return !!(props.getProperty("fcm_client_email") && props.getProperty("fcm_private_key"));
}
function pushParaUsuario_(login, titulo, corpo) {
  try {
    if (!pushAtivo_()) return;
    var raw = PropertiesService.getScriptProperties().getProperty("user_" + normUser(login));
    if (!raw) return;
    var rec = JSON.parse(raw);
    if (rec.fcmToken) sendFcmPush(rec.fcmToken, titulo, corpo);
  } catch (e) { Logger.log("push falhou: " + e.message); }
}
function pushGestores_(empresaNome, titulo, corpo) {
  try {
    if (!pushAtivo_()) return;
    var all = PropertiesService.getScriptProperties().getProperties();
    var slug = slugEmpresa(empresaNome);
    Object.keys(all).forEach(function (k) {
      if (k.indexOf("user_") !== 0) return;
      try {
        var rec = JSON.parse(all[k]);
        if (rec.role === "gestor" && rec.fcmToken && slugEmpresa(rec.empresa || "") === slug) sendFcmPush(rec.fcmToken, titulo, corpo);
      } catch (e) {}
    });
  } catch (e) { Logger.log("push gestores falhou: " + e.message); }
}

// Função chamada pelo gatilho diário (dias úteis, ~7h da manhã) — manda o
// lembrete da PT para todo usuário que já tenha um token salvo.
function enviarLembretesDiarios() {
  var diaSemana = new Date().getDay(); // 0=domingo ... 6=sábado
  if (diaSemana === 0 || diaSemana === 6) return; // não notifica fim de semana

  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  for (var k in all) {
    if (k.indexOf("user_") === 0) {
      try {
        var rec = JSON.parse(all[k]);
        if (rec.fcmToken) {
          sendFcmPush(
            rec.fcmToken,
            "DocSeg Digital",
            "Bom dia! Não esqueça de abrir a Permissão de Trabalho (PT) de hoje antes de iniciar as atividades."
          );
        }
      } catch (e) {
        Logger.log("Falha ao notificar " + k + ": " + e.message);
      }
    }
  }
}


/* ============ LIBERAÇÃO DE DOCUMENTOS (Apto / Não apto) ============
   Ligada por empresa (cadastro da empresa) e por tipo de documento. Quando
   ligada, o documento enviado fica "aguardando": o PDF espera na pasta
   "_Pendentes" (que o cliente não vê) e os gestores recebem um push. O
   gestor avalia e assina (na hora ou com o carimbo + PIN); o app dele monta
   o PDF final com o quadro "Liberação da atividade" e só então ele vai para
   a pasta da empresa — "Não Aptas"/"Não Aptos" quando recusado.
   Tudo fica na aba "Liberações" da planilha. */
var LIB_ABA = "Liberações";
var LIB_HEADERS = ["ID", "Enviado em", "Empresa", "Tipo de documento", "Número", "Revisão", "Login", "Técnico",
  "Área/Local", "Resumo", "Não conformidades", "Arquivo pendente", "Status", "Decidido por (login)", "Decidido por",
  "Decidido em", "Forma", "Motivo", "O que fazer", "Pode reenviar", "Arquivo final", "Lembrete enviado", "Dados"];
// Colunas (base 0) da aba "Liberações"
var LB = { id: 0, ts: 1, empresa: 2, tipo: 3, numero: 4, rev: 5, login: 6, nome: 7, area: 8, resumo: 9, nc: 10,
  arqPend: 11, status: 12, decLogin: 13, decNome: 14, decTs: 15, forma: 16, motivo: 17, fazer: 18, reenviar: 19,
  arqFinal: 20, lembrete: 21, dados: 22 };
var LIB_MODOS = ["escolha", "hora", "carimbo"];
var PIN_MAX_FALHAS = 5;
var PIN_BLOQUEIO_MS = 15 * 60000;

// Técnico de Segurança (perfil "tst"): quando a empresa tem ao menos um, só
// eles liberam e são avisados; o gestor passa a acompanhar. Sem nenhum, o
// gestor libera como antes.
function equipeDoGestor_(rec) { return rec.role === "tecnico" || rec.role === "tst"; }
function usuariosDaEmpresa_(empresaNome, role) {
  var all = PropertiesService.getScriptProperties().getProperties(), slug = slugEmpresa(empresaNome), out = [];
  Object.keys(all).forEach(function (k) {
    if (k.indexOf("user_") !== 0) return;
    try { var rec = JSON.parse(all[k]); if (rec.role === role && slugEmpresa(rec.empresa || "") === slug) out.push(rec); } catch (e) {}
  });
  return out;
}
function empresaTemTst_(empresaNome) { return !!empresaNome && usuariosDaEmpresa_(empresaNome, "tst").length > 0; }
function podeLiberar_(caller) {
  if (caller.role === "tst") return true;
  return caller.role === "gestor" && !empresaTemTst_(caller.empresa);
}
function pushLiberadores_(empresaNome, titulo, corpo) {
  try {
    if (!pushAtivo_()) return;
    var tsts = usuariosDaEmpresa_(empresaNome, "tst");
    (tsts.length ? tsts : usuariosDaEmpresa_(empresaNome, "gestor")).forEach(function (rec) {
      if (rec.fcmToken) { try { sendFcmPush(rec.fcmToken, titulo, corpo); } catch (e) { Logger.log("push liberador falhou: " + e.message); } }
    });
  } catch (e) { Logger.log("push liberadores falhou: " + e.message); }
}

function normLiberacao_(x) {
  x = x || {};
  var docs = Array.isArray(x.docs) ? DOC_TIPOS.filter(function (k) { return x.docs.indexOf(k) !== -1; }) : [];
  return {
    ativo: !!x.ativo, docs: docs,
    modo: LIB_MODOS.indexOf(x.modo) !== -1 ? x.modo : "escolha",
    lembreteMin: Math.min(240, Math.max(5, parseInt(x.lembreteMin, 10) || 15))
  };
}
function liberacaoExigida_(empresaNome, tipoKey) {
  var l = getEmpresaExtras(empresaNome).liberacao;
  return l.ativo && l.docs.indexOf(tipoKey) !== -1;
}
function pastaPendentes_() { return subpasta(getRaiz(), "_Pendentes"); }
function libSheet_() { return getOrCreateSheet(getOrCreateSpreadsheet(), LIB_ABA, LIB_HEADERS); }
function tituloLib_(tipoDocumento, numero) {
  return tipoDocKey(tipoDocumento) === "pt" ? "PT Nº " + (numero || "S/N") : String(tipoDocumento || "Documento");
}
function libDados_(v) { try { return JSON.parse(String(v[LB.dados] || "{}")); } catch (e) { return {}; } }
function libObj_(v) {
  var d = libDados_(v);
  return {
    id: String(v[LB.id]), ts: Number(v[LB.ts]), empresa: String(v[LB.empresa] || ""), tipoDocumento: String(v[LB.tipo] || ""),
    tipo: tipoDocKey(v[LB.tipo]), numero: String(v[LB.numero] || ""), revisao: Number(v[LB.rev]) || 0,
    login: normUser(v[LB.login]), nome: String(v[LB.nome] || ""), area: String(v[LB.area] || ""), resumo: String(v[LB.resumo] || ""),
    naoConformes: v[LB.nc] ? String(v[LB.nc]).split(" | ") : [], fileId: String(v[LB.arqFinal] || v[LB.arqPend] || ""),
    status: String(v[LB.status] || "aguardando"), decididoPor: String(v[LB.decNome] || ""), decididoTs: Number(v[LB.decTs]) || 0,
    forma: String(v[LB.forma] || ""), motivo: String(v[LB.motivo] || ""), oQueFazer: String(v[LB.fazer] || ""),
    podeReenviar: String(v[LB.reenviar]) === "1", dataDoc: d.ptData || "", horaIni: d.horaIni || "", horaFim: d.horaFim || ""
  };
}
// Linha (base 1) e valores de uma liberação pelo ID do documento, ou null.
function acharLiberacao_(id) {
  id = String(id || "");
  if (!id) return null;
  var sh = libSheet_(), n = sh.getLastRow();
  if (n < 2) return null;
  var cell = sh.getRange(2, 1, n - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  if (!cell) return null;
  return { row: cell.getRow(), v: sh.getRange(cell.getRow(), 1, 1, LIB_HEADERS.length).getValues()[0] };
}
function acharLiberacaoPorArquivo_(fileId) {
  var sh = libSheet_(), n = sh.getLastRow();
  if (n < 2) return null;
  var vals = sh.getRange(2, 1, n - 1, LIB_HEADERS.length).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][LB.arqPend]) === fileId || String(vals[i][LB.arqFinal]) === fileId) {
      return { empresa: String(vals[i][LB.empresa] || ""), login: normUser(vals[i][LB.login]) };
    }
  }
  return null;
}
// Chamada dentro do lock de saveDocuments.
function registrarPendente_(caller, empresaNome, it, file) {
  var sh = libSheet_();
  var nc = Array.isArray(it.naoConformes) ? it.naoConformes.slice(0, 40).map(function (x) { return String(x).slice(0, 160); }) : [];
  var dados = {
    ptData: it.ptData || "", horaIni: String(it.ptHoraIni || "").slice(0, 5), horaFim: String(it.ptHoraFim || "").slice(0, 5),
    motivoEnc: String(it.ptMotivo || "").slice(0, 60), pdfNome: file.getName(), revisaoDe: String(it.revisaoDe || "")
  };
  sh.appendRow([String(it.docId), Date.now(), empresaNome, it.tipoDocumento || "", it.ptNumero || "", Math.max(0, parseInt(it.revisao, 10) || 0),
    caller.username, caller.nome || caller.username, it.ptArea || "", String(it.resumo || "").slice(0, 300), nc.join(" | "),
    file.getId(), "aguardando", "", "", "", "", "", "", "", "", "", JSON.stringify(dados)]);
  // Era a correção de uma recusa: a recusa anterior sai da lista "corrigir".
  if (it.revisaoDe) {
    var ant = acharLiberacao_(it.revisaoDe);
    if (ant && normUser(ant.v[LB.login]) === caller.username && String(ant.v[LB.status]) === "nao_apto") {
      sh.getRange(ant.row, LB.status + 1).setValue("reenviada");
    }
  }
}

function listLiberacoes(data) {
  var caller = authCaller(data);
  if (!caller || (caller.role !== "admin" && caller.role !== "gestor" && caller.role !== "tst")) return jsonOut({ error: "não autorizado" });
  var empresaNome = caller.role === "admin" ? (data.empresa || "").toString().trim() : caller.empresa;
  if (!empresaNome) return jsonOut({ error: "empresa não informada" });
  var slug = slugEmpresa(empresaNome);
  var sh = libSheet_(), n = sh.getLastRow(), pendentes = [], recentes = [];
  var desde = Date.now() - 7 * 86400000;
  if (n > 1) {
    var vals = sh.getRange(2, 1, n - 1, LIB_HEADERS.length).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (!vals[i][LB.id] || slugEmpresa(String(vals[i][LB.empresa] || "")) !== slug) continue;
      var o = libObj_(vals[i]);
      if (o.status === "aguardando") pendentes.push(o);
      else if (o.ts >= desde && recentes.length < 50) recentes.push(o);
    }
  }
  pendentes.sort(function (a, b) { return a.ts - b.ts; });
  return jsonOut({ ok: true, pendentes: pendentes, recentes: recentes, config: getEmpresaExtras(empresaNome).liberacao,
    carimbo: caller.role !== "admin" ? carimboResumo_(caller.username, caller.rec) : null,
    podeLiberar: podeLiberar_(caller), temTst: empresaTemTst_(empresaNome), serverNow: Date.now() });
}

function minhasLiberacoes(data) {
  var caller = authCaller(data);
  if (!caller || caller.role === "admin") return jsonOut({ error: "não autorizado" });
  var sh = libSheet_(), n = sh.getLastRow(), out = [];
  var desde = Date.now() - 3 * 86400000;
  if (n > 1) {
    var vals = sh.getRange(2, 1, n - 1, LIB_HEADERS.length).getValues();
    for (var i = vals.length - 1; i >= 0 && out.length < 30; i--) {
      if (normUser(vals[i][LB.login]) !== caller.username) continue;
      var o = libObj_(vals[i]);
      if (o.status === "aguardando" || o.status === "nao_apto" || o.ts >= desde) out.push(o);
    }
  }
  return jsonOut({ ok: true, liberacoes: out, serverNow: Date.now() });
}

function decidirLiberacao(data) {
  var caller = authCaller(data);
  if (!caller || !podeLiberar_(caller)) return jsonOut({ error: caller && caller.role === "gestor" ? "nesta empresa quem libera é o Técnico de Segurança" : "só o Técnico de Segurança ou o gestor da empresa pode liberar" });
  var decisao = data.decisao === "apto" ? "apto" : (data.decisao === "nao_apto" ? "nao_apto" : "");
  if (!decisao) return jsonOut({ error: "decisão inválida" });
  var forma = data.forma === "carimbo" ? "carimbo" : "hora";
  var motivo = (data.motivo || "").toString().trim().slice(0, 80);
  var oQueFazer = (data.oQueFazer || "").toString().trim().slice(0, 500);
  if (decisao === "nao_apto" && !motivo) return jsonOut({ error: "informe o motivo" });
  var m = /^[A-Za-z0-9+\/=]+$/.exec((data.pdfFinalBase64 || "").toString());
  if (!m) return jsonOut({ error: "PDF final ausente" });

  var config = getEmpresaExtras(caller.empresa).liberacao;
  if (config.modo !== "escolha" && config.modo !== forma) {
    return jsonOut({ error: config.modo === "carimbo" ? "esta empresa exige o carimbo digital" : "esta empresa exige a assinatura na hora" });
  }
  if (forma === "hora" && String(data.assinaturaHora || "").indexOf("data:image/") !== 0) return jsonOut({ error: "assine antes de confirmar" });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var item, dados, finalFile;
  try {
    var achado = acharLiberacao_(data.id);
    if (!achado || !sameEmpresa(achado.v[LB.empresa], caller.empresa)) return jsonOut({ error: "documento não encontrado" });
    item = libObj_(achado.v);
    if (item.status !== "aguardando") {
      return jsonOut({ error: "este documento já foi avaliado" + (item.decididoPor ? " por " + item.decididoPor : ""), item: item });
    }
    var rec = getUserRec(caller.username);
    var carimbo = rec.carimbo || {};
    if (forma === "carimbo") {
      if (!carimbo.registro || !getAssinaturaCarimbo_(caller.username)) return jsonOut({ error: "cadastre seu carimbo no perfil antes de usar" });
      var pin = verificarPin_(caller.username, rec, (data.pinHash || "").toString());
      if (pin.error) return jsonOut({ error: pin.error, pinErro: true });
    }
    dados = libDados_(achado.v);
    var nomeArq = String(dados.pdfNome || (item.numero || "documento") + ".pdf");
    if (decisao === "nao_apto") nomeArq = nomeArq.replace(/\.pdf$/i, "") + "_NAO_APTO.pdf";
    var blob = Utilities.newBlob(Utilities.base64Decode(data.pdfFinalBase64), "application/pdf", nomeArq);
    var pasta = pastaDoMes(caller.empresa);
    if (decisao === "nao_apto") pasta = subpasta(pasta, item.tipo === "pt" ? "Não Aptas" : "Não Aptos");
    finalFile = pasta.createFile(blob);
    try { DriveApp.getFileById(String(achado.v[LB.arqPend])).setTrashed(true); } catch (e) { /* já não existe */ }

    var agora = Date.now();
    var quem = (rec.nome || caller.nome) + (carimbo.registro ? " (" + (carimbo.orgao ? carimbo.orgao + " " : "") + carimbo.registro + ")" : "");
    var podeReenviar = decisao === "nao_apto" && data.podeReenviar !== false;
    libSheet_().getRange(achado.row, LB.status + 1, 1, 9).setValues([[decisao, caller.username, quem, agora, forma,
      motivo, oQueFazer, podeReenviar ? "1" : "0", finalFile.getId()]]);
    item = libObj_(libSheet_().getRange(achado.row, 1, 1, LIB_HEADERS.length).getValues()[0]);

    registrarNaPlanilha({
      tipoDocumento: item.tipoDocumento, ptNumero: item.numero, ptData: dados.ptData, ptArea: item.area,
      empresa: caller.empresa, usuario: item.nome, login: item.login, resumo: item.resumo,
      pdfNome: nomeArq, fileId: finalFile.getId(), docId: item.id,
      horaIni: dados.horaIni, horaFim: dados.horaFim, motivo: dados.motivoEnc, naoConformes: item.naoConformes,
      liberacao: decisao === "apto" ? "Apto" : "Não apto", liberadoPor: quem,
      liberadoEm: Utilities.formatDate(new Date(agora), TZ_CONTRATO, "dd/MM/yyyy HH:mm"),
      motivoNaoApto: decisao === "nao_apto" ? motivo + (oQueFazer ? " — " + oQueFazer : "") : ""
    });
  } finally {
    lock.releaseLock();
  }

  var titulo = tituloLib_(item.tipoDocumento, item.numero), fem = item.tipo === "pt";
  try {
    var ext = getEmpresaExtras(caller.empresa);
    if (ext.emailAtivo && ext.emailDestino) {
      GmailApp.sendEmail(ext.emailDestino,
        "[DocSeg Digital] " + titulo + " — " + (decisao === "apto" ? "APTO" : "NÃO APTO") + " - " + caller.empresa,
        "Enviado por " + item.nome + " e avaliado por " + item.decididoPor + " pelo DocSeg Digital." +
        (decisao === "nao_apto" ? "\n\nMotivo: " + motivo + (oQueFazer ? "\nO que fazer: " + oQueFazer : "") : ""),
        { attachments: [finalFile.getBlob()], name: "DocSeg Digital" });
    }
  } catch (mailErr) { Logger.log("Falha ao enviar e-mail da liberação: " + mailErr.message); }
  if (decisao === "apto") {
    pushParaUsuario_(item.login, "✅ " + titulo + (fem ? " liberada" : " liberado"), "Liberado por " + item.decididoPor + ". Pode iniciar a atividade.");
  } else {
    pushParaUsuario_(item.login, "⛔ " + titulo + (fem ? " não liberada" : " não liberado"), "Não inicie. Motivo: " + motivo + (oQueFazer ? ". " + oQueFazer.slice(0, 90) : ""));
  }
  return jsonOut({ ok: true, item: item });
}

// Depois de um "Não apto": o técnico (ou o gestor) encerra o documento como
// "Não executado" quando não vai corrigir (ex.: a própria pessoa não está apta).
function encerrarLiberacao(data) {
  var caller = authCaller(data);
  if (!caller || caller.role === "admin") return jsonOut({ error: "não autorizado" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var achado = acharLiberacao_(data.id);
    if (!achado) return jsonOut({ error: "documento não encontrado" });
    var dono = normUser(achado.v[LB.login]) === caller.username;
    var gestor = (caller.role === "gestor" || caller.role === "tst") && sameEmpresa(achado.v[LB.empresa], caller.empresa);
    if (!dono && !gestor) return jsonOut({ error: "não autorizado" });
    if (String(achado.v[LB.status]) !== "nao_apto") return jsonOut({ error: "só um documento não apto pode ser encerrado" });
    libSheet_().getRange(achado.row, LB.status + 1).setValue("nao_executada");
    return jsonOut({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Gatilho a cada 5 min (instalarGatilhos): lembra os gestores, uma vez, dos
// documentos que esperam liberação há mais tempo que o definido na empresa.
function lembrarLiberacoesPendentes() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  var avisos = [];
  try {
    var sh = libSheet_(), n = sh.getLastRow();
    if (n < 2) return;
    var vals = sh.getRange(2, 1, n - 1, LIB_HEADERS.length).getValues(), agora = Date.now(), cfg = {};
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i];
      if (String(v[LB.status]) !== "aguardando" || String(v[LB.lembrete]) === "1") continue;
      var emp = String(v[LB.empresa] || "");
      if (!(emp in cfg)) cfg[emp] = getEmpresaExtras(emp).liberacao.lembreteMin;
      var min = Math.floor((agora - Number(v[LB.ts])) / 60000);
      if (min < cfg[emp]) continue;
      sh.getRange(i + 2, LB.lembrete + 1).setValue("1");
      avisos.push({ empresa: emp, titulo: tituloLib_(v[LB.tipo], v[LB.numero]), nome: String(v[LB.nome] || ""), min: min });
    }
  } finally {
    lock.releaseLock();
  }
  avisos.forEach(function (a) {
    pushLiberadores_(a.empresa, "⏰ " + a.titulo + " aguardando há " + a.min + " min", a.nome + " está esperando a liberação para começar.");
  });
}

/* ---- Carimbo digital do liberador (gestor) ----
   Nome, função, órgão e nº de registro ficam no cadastro do usuário; a
   imagem da assinatura fica na aba "Carimbos" (o cofre do script tem
   limite de tamanho). O PIN é criado pela própria pessoa e guardado com
   sal — nunca chega por SMS nem fica no aparelho. */
function carimbosSheet_() { return getOrCreateSheet(getOrCreateSpreadsheet(), "Carimbos", ["Login", "Assinatura (data URL)", "Atualizada em"]); }
function getAssinaturaCarimbo_(uname) {
  var sh = carimbosSheet_(), n = sh.getLastRow();
  if (n < 2) return "";
  var cell = sh.getRange(2, 1, n - 1, 1).createTextFinder(uname).matchEntireCell(true).findNext();
  return cell ? String(sh.getRange(cell.getRow(), 2).getValue() || "") : "";
}
function setAssinaturaCarimbo_(uname, dataUrl) {
  var sh = carimbosSheet_(), n = sh.getLastRow();
  var cell = n < 2 ? null : sh.getRange(2, 1, n - 1, 1).createTextFinder(uname).matchEntireCell(true).findNext();
  if (!dataUrl) { if (cell) sh.deleteRow(cell.getRow()); return; }
  if (cell) sh.getRange(cell.getRow(), 2, 1, 2).setValues([[dataUrl, new Date()]]);
  else sh.appendRow([uname, dataUrl, new Date()]);
}
function carimboResumo_(uname, rec) {
  var c = (rec && rec.carimbo) || {};
  // O nome do carimbo é sempre o nome completo do cadastro (o gestor corrige em Colaboradores).
  return { nome: (rec && rec.nome) || c.nome || "", funcao: c.funcao || "", orgao: c.orgao || "", registro: c.registro || "",
    temAssinatura: !!getAssinaturaCarimbo_(uname), temPin: !!(rec && rec.pinHash) };
}
function verificarPin_(uname, rec, pinHash) {
  if (!rec.pinHash) return { error: "crie o seu PIN em Perfil › Meu carimbo" };
  var agora = Date.now();
  if (rec.pinBloqueadoAte && rec.pinBloqueadoAte > agora) {
    return { error: "PIN bloqueado por excesso de tentativas. Tente em " + Math.ceil((rec.pinBloqueadoAte - agora) / 60000) + " min" };
  }
  var props = PropertiesService.getScriptProperties();
  if (sha256Hex(rec.pinSalt + ":" + pinHash) !== rec.pinHash) {
    rec.pinFalhas = (rec.pinFalhas || 0) + 1;
    var msg;
    if (rec.pinFalhas >= PIN_MAX_FALHAS) {
      rec.pinFalhas = 0; rec.pinBloqueadoAte = agora + PIN_BLOQUEIO_MS;
      msg = "PIN incorreto. Bloqueado por 15 min";
    } else {
      msg = "PIN incorreto (restam " + (PIN_MAX_FALHAS - rec.pinFalhas) + " tentativas)";
    }
    props.setProperty("user_" + uname, JSON.stringify(rec));
    return { error: msg };
  }
  if (rec.pinFalhas || rec.pinBloqueadoAte) {
    rec.pinFalhas = 0; rec.pinBloqueadoAte = 0;
    props.setProperty("user_" + uname, JSON.stringify(rec));
  }
  return { ok: true };
}
function getCarimbo(data) {
  var caller = authCaller(data);
  if (!caller || (caller.role !== "gestor" && caller.role !== "tst")) return jsonOut({ error: "não autorizado" });
  var out = carimboResumo_(caller.username, caller.rec);
  out.assinatura = getAssinaturaCarimbo_(caller.username);
  return jsonOut({ ok: true, carimbo: out });
}
function saveCarimbo(data) {
  var caller = authCaller(data);
  if (!caller || (caller.role !== "gestor" && caller.role !== "tst")) return jsonOut({ error: "não autorizado" });
  var assinatura = data.assinatura;
  if (typeof assinatura === "string" && assinatura && (assinatura.indexOf("data:image/") !== 0 || assinatura.length > 45000)) {
    return jsonOut({ error: "assinatura inválida ou muito grande" });
  }
  var novoPin = (data.novoPinHash || "").toString();
  if (novoPin && !/^[0-9a-f]{64}$/.test(novoPin)) return jsonOut({ error: "PIN inválido" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var rec = getUserRec(caller.username);
    if (novoPin) {
      // Criar ou trocar o PIN pede a senha de login de novo (serve também de "esqueci o PIN").
      if ((data.senhaHash || "") !== rec.passwordHash) return jsonOut({ error: "senha de login incorreta" });
      rec.pinSalt = Utilities.getUuid();
      rec.pinHash = sha256Hex(rec.pinSalt + ":" + novoPin);
      rec.pinFalhas = 0; rec.pinBloqueadoAte = 0;
    }
    rec.carimbo = {
      nome: rec.nome || caller.nome, funcao: (data.funcao || "").toString().trim().slice(0, 80),
      orgao: (data.orgao || "").toString().trim().slice(0, 20), registro: (data.registro || "").toString().trim().slice(0, 30),
      atualizadoEm: new Date().toISOString()
    };
    PropertiesService.getScriptProperties().setProperty("user_" + caller.username, JSON.stringify(rec));
    if (typeof assinatura === "string") setAssinaturaCarimbo_(caller.username, assinatura);
    return jsonOut({ ok: true, carimbo: carimboResumo_(caller.username, rec) });
  } finally {
    lock.releaseLock();
  }
}

/* ============ E-MAIL DOS PDFs (definido pelo gestor) ============ */
function saveEmailEmpresa(data) {
  var caller = authCaller(data);
  if (!caller || caller.role !== "gestor") return jsonOut({ error: "não autorizado" });
  var ativo = !!data.emailAtivo;
  var lista = parseEmails(data.emailDestino);
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var ruins = lista.filter(function (e) { return !re.test(e); });
  if (ruins.length) return jsonOut({ error: "e-mail inválido: " + ruins.join(", ") });
  if (ativo && !lista.length) return jsonOut({ error: "informe pelo menos um e-mail ou desligue o envio" });
  var slug = slugEmpresa(caller.empresa);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty("empresa_" + slug);
    if (!raw) return jsonOut({ error: "empresa não encontrada" });
    var rec = JSON.parse(raw);
    rec.emailAtivo = ativo;
    rec.emailDestino = lista.join(", ");
    rec.updatedAt = new Date().toISOString();
    props.setProperty("empresa_" + slug, JSON.stringify(rec));
    return jsonOut({ ok: true, emailAtivo: rec.emailAtivo, emailDestino: rec.emailDestino });
  } finally {
    lock.releaseLock();
  }
}

/* ============ SONS DO APP (definidos pelo administrador, valem para todos) ============ */
var SONS_OPCOES = ["sino", "plim", "duplo", "suave", "nenhum"];
var SONS_PADRAO = { enviar: "sino", liberar: "plim", salvar: "nenhum" };
function getSons_() {
  var s = {};
  try { s = JSON.parse(PropertiesService.getScriptProperties().getProperty("app_sons") || "{}"); } catch (e) {}
  var out = {};
  Object.keys(SONS_PADRAO).forEach(function (k) { out[k] = SONS_OPCOES.indexOf(s[k]) !== -1 ? s[k] : SONS_PADRAO[k]; });
  return out;
}
function saveSons(data) {
  if (!requireAdmin(data)) return jsonOut({ error: "não autorizado" });
  var s = {};
  Object.keys(SONS_PADRAO).forEach(function (k) { s[k] = SONS_OPCOES.indexOf(data[k]) !== -1 ? data[k] : SONS_PADRAO[k]; });
  PropertiesService.getScriptProperties().setProperty("app_sons", JSON.stringify(s));
  return jsonOut({ ok: true, sons: s });
}

/* ============ APAGAR CHAMADOS DE SUPORTE ============
   Usuário: o chamado some só da lista dele (o admin continua vendo).
   Admin: o chamado sai da lista e vai para a aba "Chamados excluídos". */
function chamadosExcluidosSheet_() {
  return getOrCreateSheet(getOrCreateSpreadsheet(), "Chamados excluídos", CHAMADOS_HEADERS.concat(["Apagado em"]));
}
function apagarChamadoLinha_(sh, row, caller) {
  if (caller.role === "admin") {
    var v = sh.getRange(row, 1, 1, CHAMADOS_HEADERS.length).getValues()[0];
    chamadosExcluidosSheet_().appendRow(v.concat([new Date()]));
    sh.deleteRow(row);
  } else {
    sh.getRange(row, CH_OCULTO + 1).setValue("1");
  }
}
function apagarChamado(data) {
  var caller = authCaller(data);
  if (!caller) return jsonOut({ error: "não autorizado" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = chamadosSheet(), n = sh.getLastRow();
    if (n < 2) return jsonOut({ error: "chamado não encontrado" });
    var vals = sh.getRange(2, 1, n - 1, CHAMADOS_HEADERS.length).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) !== String(data.id || "")) continue;
      if (caller.role !== "admin" && normUser(vals[i][2]) !== caller.username) return jsonOut({ error: "não autorizado" });
      apagarChamadoLinha_(sh, i + 2, caller);
      return jsonOut({ ok: true });
    }
    return jsonOut({ error: "chamado não encontrado" });
  } finally {
    lock.releaseLock();
  }
}
// "Limpar histórico": o usuário esconde todos os dele; o admin apaga os já resolvidos.
function limparChamados(data) {
  var caller = authCaller(data);
  if (!caller) return jsonOut({ error: "não autorizado" });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = chamadosSheet(), n = sh.getLastRow(), qtd = 0;
    if (n < 2) return jsonOut({ ok: true, apagados: 0 });
    var vals = sh.getRange(2, 1, n - 1, CHAMADOS_HEADERS.length).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var meu = normUser(vals[i][2]) === caller.username;
      var alvo = caller.role === "admin" ? String(vals[i][13]) === "resolvido" : (meu && String(vals[i][CH_OCULTO]) !== "1");
      if (!alvo) continue;
      apagarChamadoLinha_(sh, i + 2, caller);
      qtd++;
    }
    return jsonOut({ ok: true, apagados: qtd });
  } finally {
    lock.releaseLock();
  }
}

/* ============ NUMERAÇÃO CENTRALIZADA DA PT ============ */
// Guarda o contador no "cofre" do próprio script (PropertiesService),
// compartilhado entre todos os celulares que usam esta mesma implantação.
// A numeração é separada automaticamente por EMPRESA (cada empresa cadastrada
// tem sua própria sequência, começando do 0001) — isso é identificado pela
// empresa cadastrada no usuário que está logado, sem precisar configurar nada
// a mais. Se por algum motivo a chamada não informar a empresa (ex: usuário
// sem empresa definida), cai no comportamento antigo, baseado no prefixo
// digitado manualmente em Configurações — só pra não quebrar instalações
// antigas.
// Antes era por GET e sem login: qualquer um com o endereço do servidor podia
// pular ou zerar a numeração. Agora vem por POST (doPost) e exige login.
// O doGet só responde para avisar quem ainda está com o app antigo.
function doGet(e) {
  return jsonOut({ error: "atualize o app para a versão mais nova" });
}

function numeracao(data) {
  try {
    var caller = authCaller(data);
    if (!caller) return jsonOut({ error: "não autorizado" });
    var action = data.action;
    if (action !== "nextNumber" && caller.role !== "admin") return jsonOut({ error: "não autorizado" });
    var prefix = (data.prefix || "").toString();
    // Quem não é admin só mexe na numeração da própria empresa.
    var empresaParam = caller.role === "admin" ? (data.empresa || "").toString() : caller.empresa;
    var empresaSlug = empresaParam ? slugEmpresa(empresaParam) : "";
    var key = empresaSlug ? ("pt_counter_emp_" + empresaSlug) : ("pt_counter_" + prefix);
    var props = PropertiesService.getScriptProperties();

    if (action === "nextNumber") {
      // Reserva e devolve o próximo número, com trava para evitar
      // que dois celulares recebam o mesmo número ao mesmo tempo.
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var current = parseInt(props.getProperty(key) || "0", 10);
        var next = current + 1;
        props.setProperty(key, String(next));
        return jsonOut({ number: next });
      } finally {
        lock.releaseLock();
      }
    }

    if (action === "peekNumber") {
      // Só consulta qual seria o próximo número, sem reservar nada.
      var currentPeek = parseInt(props.getProperty(key) || "0", 10);
      return jsonOut({ number: currentPeek + 1 });
    }

    if (action === "setNumber") {
      // Ajuste manual feito pelo responsável, via app (Configurações).
      var value = parseInt(data.value, 10);
      if (isNaN(value) || value < 1) {
        return jsonOut({ error: "valor inválido" });
      }
      props.setProperty(key, String(value - 1));
      return jsonOut({ ok: true, nextWillBe: value });
    }

    if (action === "sheetUrl") {
      // Devolve o link da planilha central de registros, para o app
      // mostrar um atalho em Configurações. Cria a planilha se ainda não existir.
      var sheet = getOrCreateAuditSheet();
      return jsonOut({ url: sheet.getParent().getUrl() });
    }

    return jsonOut({ error: "ação inválida" });
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}
