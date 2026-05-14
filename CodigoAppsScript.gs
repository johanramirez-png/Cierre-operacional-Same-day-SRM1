var CONFIG = {
  srm1: {
    gridDocId        : "01KRGV19Z3RQE74SFNHJD6R76S",
    gridSkillVersion : "3.6.3",
    processes        : ["SRM13-SD","SRM15-SD","SRM17-SD","SRM19-SD"],
    pmMap            : {"SRM13-SD":"PM1","SRM15-SD":"PM2","SRM17-SD":"PM3","SRM19-SD":"PM4"},
    driveFileName    : "Dashboard_Historico_SameDay_MLC.html",
    ghFile           : "Dashboard_Historico_SameDay_MLC.html",
    fallbackLastDate : "2026-05-12"
  },
  srm2: {
    gridDocId        : "PENDIENTE",
    gridSkillVersion : "3.6.3",
    processes        : ["SRM2A-SD","SRM2B-SD","SRM2C-SD"],
    pmMap            : {"SRM2A-SD":"PM1","SRM2B-SD":"PM2","SRM2C-SD":"PM3"},
    driveFileName    : "Dashboard_Historico_SameDay_MLC_SRM2.html",
    ghFile           : "Dashboard_Historico_SameDay_MLC_SRM2.html",
    fallbackLastDate : "2026-05-13"
  },
  gridEndpoint    : "https://grid.melioffice.com/api/v1/engine/run/json",
  sheetName       : "Registro V2",
  debounceMinutos : 5
};

function getTokens() {
  var props = PropertiesService.getScriptProperties();
  return {
    srm1Token : props.getProperty("grid_token_srm1") || "",
    srm2Token : props.getProperty("grid_token_srm2") || "",
    srm2DocId : props.getProperty("grid_docid_srm2") || CONFIG.srm2.gridDocId
  };
}

function getGitHubConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    token : props.getProperty("gh_token") || "",
    owner : props.getProperty("gh_owner") || "",
    repo  : props.getProperty("gh_repo")  || "",
    branch: props.getProperty("gh_branch") || "main"
  };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Dashboards SD")
    .addItem("Actualizar SRM1 + SRM2",        "actualizarAmbos")
    .addSeparator()
    .addItem("Actualizar solo SRM1",           "actualizarSRM1")
    .addItem("Actualizar solo SRM2",           "actualizarSRM2")
    .addSeparator()
    .addItem("Configurar token Grid",          "configurarTokens")
    .addItem("Configurar GitHub",              "configurarGitHub")
    .addItem("Instalar trigger automatico",    "instalarTriggerAuto")
    .addItem("Desinstalar trigger automatico", "desinstalarTriggerAuto")
    .addSeparator()
    .addItem("Diagnostico del sistema",        "diagnosticar")
    .addItem("Ver log ultima ejecucion",       "verLog")
    .addToUi();
}

function actualizarAmbos() {
  var r1 = procesarServicio("srm1");
  var r2 = procesarServicio("srm2");
  mostrarResumen([r1, r2]);
}
function actualizarSRM1() { mostrarResumen([procesarServicio("srm1")]); }
function actualizarSRM2() { mostrarResumen([procesarServicio("srm2")]); }

function verLog() {
  var log = PropertiesService.getScriptProperties().getProperty("ultimo_log") || "Sin ejecuciones registradas.";
  SpreadsheetApp.getUi().alert("Log", log, SpreadsheetApp.getUi().ButtonSet.OK);
}

function configurarTokens() {
  var ui     = SpreadsheetApp.getUi();
  var props  = PropertiesService.getScriptProperties();
  var tokens = getTokens();

  var r1 = ui.prompt(
    "Paso 1 de 2 - Token Grid",
    "Pega tu token grid_sk_...\n(actual: " + (tokens.srm1Token ? tokens.srm1Token.slice(0,12) + "..." : "vacio") + ")",
    ui.ButtonSet.OK_CANCEL
  );
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var nuevoToken = r1.getResponseText().trim();
  if (nuevoToken) {
    props.setProperty("grid_token_srm1", nuevoToken);
    props.setProperty("grid_token_srm2", nuevoToken);
  }

  var r2 = ui.prompt(
    "Paso 2 de 2 - DocId SRM2 (opcional)",
    "Pega el doc_id del archivo SRM2 en Grid\n(actual: " + tokens.srm2DocId + ")\nDeja en blanco para mantener.",
    ui.ButtonSet.OK_CANCEL
  );
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var nuevoDocId = r2.getResponseText().trim();
  if (nuevoDocId && nuevoDocId !== "PENDIENTE") {
    props.setProperty("grid_docid_srm2", nuevoDocId);
  }

  ui.alert("Listo", "Tokens guardados.", ui.ButtonSet.OK);
}

function configurarGitHub() {
  var ui  = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var gh  = getGitHubConfig();

  var r1 = ui.prompt(
    "GitHub - Paso 1 de 3",
    "Nombre de usuario u organizacion GitHub\n(actual: " + (gh.owner || "vacio") + ")",
    ui.ButtonSet.OK_CANCEL
  );
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var owner = r1.getResponseText().trim();
  if (owner) props.setProperty("gh_owner", owner);

  var r2 = ui.prompt(
    "GitHub - Paso 2 de 3",
    "Nombre del repositorio\n(actual: " + (gh.repo || "vacio") + ")",
    ui.ButtonSet.OK_CANCEL
  );
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var repo = r2.getResponseText().trim();
  if (repo) props.setProperty("gh_repo", repo);

  var r3 = ui.prompt(
    "GitHub - Paso 3 de 3",
    "Personal Access Token (ghp_...)\nNecesita permisos: repo (contents:write)\n(actual: " + (gh.token ? gh.token.slice(0,12) + "..." : "vacio") + ")",
    ui.ButtonSet.OK_CANCEL
  );
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  var token = r3.getResponseText().trim();
  if (token) props.setProperty("gh_token", token);

  var ownerFinal = owner || gh.owner;
  var repoFinal  = repo  || gh.repo;
  ui.alert(
    "GitHub configurado",
    "URL del dashboard en linea:\nhttps://" + ownerFinal + ".github.io/" + repoFinal + "/\n\nActiva GitHub Pages en:\nhttps://github.com/" + ownerFinal + "/" + repoFinal + "/settings/pages",
    ui.ButtonSet.OK
  );
}

function onCambioHoja(e) {
  if (e && e.changeType && e.changeType !== "INSERT_ROW" && e.changeType !== "EDIT") {
    return;
  }
  var props    = PropertiesService.getScriptProperties();
  var ahora    = Date.now();
  var ultimaEj = parseInt(props.getProperty("ultima_ejecucion_auto") || "0");
  var minDelta = CONFIG.debounceMinutos * 60 * 1000;

  if ((ahora - ultimaEj) < minDelta) {
    if (!props.getProperty("trigger_diferido_activo")) {
      props.setProperty("trigger_diferido_activo", "1");
      ScriptApp.newTrigger("ejecutarActualizacionDiferida")
               .timeBased()
               .after(CONFIG.debounceMinutos * 60 * 1000)
               .create();
    }
    return;
  }

  props.setProperty("ultima_ejecucion_auto", String(ahora));
  props.deleteProperty("trigger_diferido_activo");
  procesarServicio("srm1");
  procesarServicio("srm2");
}

function ejecutarActualizacionDiferida() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty("trigger_diferido_activo");
  props.setProperty("ultima_ejecucion_auto", String(Date.now()));
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "ejecutarActualizacionDiferida") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  procesarServicio("srm1");
  procesarServicio("srm2");
}

function instalarTriggerAuto() {
  var ui = SpreadsheetApp.getUi();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onCambioHoja") {
      ui.alert("Ya estaba activo", "El trigger automatico ya estaba instalado.", ui.ButtonSet.OK);
      return;
    }
  }
  ScriptApp.newTrigger("onCambioHoja")
           .forSpreadsheet(SpreadsheetApp.getActive())
           .onChange()
           .create();
  ui.alert("Trigger instalado", "Los dashboards se actualizaran automaticamente al cargar datos.", ui.ButtonSet.OK);
}

function desinstalarTriggerAuto() {
  var ui       = SpreadsheetApp.getUi();
  var triggers = ScriptApp.getProjectTriggers();
  var borrados = 0;
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === "onCambioHoja" || fn === "ejecutarActualizacionDiferida") {
      ScriptApp.deleteTrigger(triggers[i]);
      borrados++;
    }
  }
  ui.alert("Listo", borrados > 0 ? "Triggers eliminados: " + borrados : "No habia triggers activos.", ui.ButtonSet.OK);
}

function procesarServicio(serviceKey) {
  var cfg    = CONFIG[serviceKey];
  var tokens = getTokens();
  var log    = [];
  var gridToken = (serviceKey === "srm1") ? tokens.srm1Token : tokens.srm2Token;
  var gridDocId = (serviceKey === "srm2") ? tokens.srm2DocId : cfg.gridDocId;

  log.push("[" + new Date().toLocaleString("es-CL") + "] Procesando " + serviceKey.toUpperCase());

  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.sheetName);
    if (!sheet) throw new Error("Hoja " + CONFIG.sheetName + " no encontrada");

    var data = sheet.getDataRange().getValues();

    var htmlContent   = null;
    var dashboardFile = null;
    var files = DriveApp.getFilesByName(cfg.driveFileName);
    if (files.hasNext()) {
      dashboardFile = files.next();
      htmlContent   = dashboardFile.getBlob().getDataAsString();
      log.push("  Dashboard encontrado: " + cfg.driveFileName);
    } else {
      log.push("  ERROR: Dashboard no encontrado en Drive: " + cfg.driveFileName);
      guardarLog(log.join("\n"));
      return { service: serviceKey.toUpperCase(), nuevos: 0, ok: false,
               msg: "Dashboard no encontrado en Drive: " + cfg.driveFileName };
    }

    var lastDate = extraerUltimaFecha(htmlContent, cfg.fallbackLastDate);
    log.push("  Ultima fecha: " + lastDate);

    var newRecords = extraerRegistros(data, cfg.processes, cfg.pmMap, lastDate);
    log.push("  Registros nuevos: " + newRecords.length);

    if (newRecords.length === 0) {
      log.push("  Sin datos nuevos desde " + lastDate);
      guardarLog(log.join("\n"));
      return { service: serviceKey.toUpperCase(), nuevos: 0, ok: true,
               msg: "Sin datos nuevos desde " + lastDate };
    }

    for (var i = 0; i < newRecords.length; i++) {
      var r = newRecords[i];
      log.push("    " + r.f + " | " + r.pm + " | " + r.p + " | SHP=" + r.shp);
    }

    htmlContent = actualizarHTML(htmlContent, newRecords);
    dashboardFile.setContent(htmlContent);
    log.push("  Drive actualizado OK");

    var gridMsg = subirAGrid(htmlContent, gridDocId, gridToken, cfg);
    log.push("  Grid: " + gridMsg);

    var ghMsg = subirAGitHub(htmlContent, cfg.ghFile, newRecords[0].f);
    log.push("  GitHub: " + ghMsg);

    guardarLog(log.join("\n"));
    return { service: serviceKey.toUpperCase(), nuevos: newRecords.length,
             ok: true, msg: "Drive OK. Grid: " + gridMsg + ". GitHub: " + ghMsg };

  } catch (e) {
    log.push("  ERROR: " + e.message);
    guardarLog(log.join("\n"));
    return { service: serviceKey.toUpperCase(), nuevos: 0, ok: false, msg: e.message };
  }
}

function subirAGitHub(htmlContent, filename, fechaActualizada) {
  var gh = getGitHubConfig();
  if (!gh.token)  return "token GitHub no configurado - usa menu Configurar GitHub";
  if (!gh.owner)  return "owner GitHub no configurado";
  if (!gh.repo)   return "repo GitHub no configurado";

  var baseUrl = "https://api.github.com/repos/" + gh.owner + "/" + gh.repo + "/contents/" + filename;
  var headers = {
    "Authorization"        : "token " + gh.token,
    "Accept"               : "application/vnd.github.v3+json",
    "Content-Type"         : "application/json",
    "X-GitHub-Api-Version" : "2022-11-28"
  };

  try {
    var sha = null;
    var getResp = UrlFetchApp.fetch(baseUrl + "?ref=" + gh.branch, {
      method: "get", headers: headers, muteHttpExceptions: true
    });
    if (getResp.getResponseCode() === 200) {
      sha = JSON.parse(getResp.getContentText()).sha;
    }

    // FIX: usar base64Encode directo para evitar truncamiento en archivos grandes (>100KB)
    var contentB64 = Utilities.base64Encode(htmlContent);
    var commitMsg  = "Update " + filename + " - cierre " + fechaActualizada;
    var body = { message: commitMsg, content: contentB64, branch: gh.branch };
    if (sha) body.sha = sha;

    var putResp = UrlFetchApp.fetch(baseUrl, {
      method: "put", headers: headers, payload: JSON.stringify(body), muteHttpExceptions: true
    });
    var code = putResp.getResponseCode();
    if (code === 200 || code === 201) {
      var ghUrl = "https://" + gh.owner + ".github.io/" + gh.repo + "/" + filename;
      return "OK - " + ghUrl;
    }
    return "Error HTTP " + code + ": " + putResp.getContentText().slice(0, 120);

  } catch (e) {
    return "Error: " + e.message;
  }
}

function extraerRegistros(data, processes, pmMap, lastDateStr) {
  var targetSet = {};
  for (var i = 0; i < processes.length; i++) targetSet[processes[i]] = true;

  var lastDate = new Date(lastDateStr + "T00:00:00");
  var records  = [];

  for (var row = 2; row < data.length; row++) {
    var fecha   = data[row][1];
    var proceso = String(data[row][2] || "").trim();
    if (!fecha || !proceso || !targetSet[proceso]) continue;
    var fechaDate = (fecha instanceof Date) ? fecha : new Date(fecha);
    if (isNaN(fechaDate.getTime()) || fechaDate <= lastDate) continue;
    records.push(buildRecord(data[row], proceso, pmMap, fechaDate));
  }

  records.sort(function(a, b) {
    if (b.f !== a.f) return b.f > a.f ? 1 : -1;
    return a.pm > b.pm ? 1 : -1;
  });
  return records;
}

function buildRecord(row, proceso, pmMap, fechaDate) {
  function fmtT(v) {
    if (!v) return null;
    if (v instanceof Date) return Utilities.formatDate(v, "America/Santiago", "HH:mm");
    var s = String(v);
    if (s.match(/^\d{2}:\d{2}/)) return s.slice(0, 5);
    return null;
  }
  function fmtI(v) { return (v !== null && v !== "") ? parseInt(v) : null; }
  function fmtF(v) { return (v !== null && v !== "") ? Math.round(parseFloat(v) * 10000) / 10000 : null; }

  return {
    f  : Utilities.formatDate(fechaDate, "America/Santiago", "yyyy-MM-dd"),
    p  : proceso,
    pm : pmMap[proceso],
    lv : String(row[0] || ""),
    shp: fmtI(row[3]),
    rut: fmtI(row[4]),
    spr: fmtF(row[5]),
    shi: fmtT(row[22]),
    rte: fmtT(row[23]),
    il : fmtT(row[24]),
    fl : fmtT(row[25]),
    fs : fmtT(row[26]),
    id : fmtT(row[30]),
    fd : fmtT(row[31]),
    dd : fmtT(row[32]),
    flt: fmtI(row[28]),
    aus: fmtF(row[20])
  };
}

function extraerUltimaFecha(html, fallback) {
  var m1 = html.match(/const CIERRE_RAW_DATA = \[{"f":"([^"]+)"/);
  if (m1) return m1[1];
  var m2 = html.match(/const CIERRE_RAW = \[{"f":"([^"]+)"/);
  if (m2) return m2[1];
  return fallback;
}

function actualizarHTML(html, newRecords) {
  var parts = [];
  for (var i = 0; i < newRecords.length; i++) {
    parts.push(JSON.stringify(newRecords[i]));
  }
  var newJson = parts.join(",");

  var keys = ["const CIERRE_RAW_DATA = [", "const CIERRE_RAW = ["];
  for (var k = 0; k < keys.length; k++) {
    if (html.indexOf(keys[k]) !== -1) {
      html = html.replace(keys[k], keys[k] + newJson + ",");
      break;
    }
  }

  var tableParts = [];
  for (var j = 0; j < newRecords.length; j++) {
    if (newRecords[j].shp && newRecords[j].shp > 0) {
      tableParts.push(JSON.stringify(newRecords[j]));
    }
  }
  if (tableParts.length > 0 && html.indexOf('"table":[') !== -1) {
    html = html.replace('"table":[', '"table":[' + tableParts.join(",") + ",");
  }

  var nowStr = Utilities.formatDate(new Date(), "America/Santiago", "yyyy-MM-dd HH:mm");
  html = html.replace(/Actualizado: \d{4}-\d{2}-\d{2} \d{2}:\d{2}/g, "Actualizado: " + nowStr);

  var mostRecent = newRecords[0].f;
  html = html.replace(
    /<option value="\d{4}-\d{2}-\d{2}" selected>/,
    '<option value="' + mostRecent + '" selected>'
  );
  return html;
}

function subirAGrid(htmlContent, gridDocId, gridToken, cfg) {
  if (!gridDocId || gridDocId === "PENDIENTE") return "doc_id no configurado - omitido";
  if (!gridToken) return "token no configurado - usa menu Configurar token Grid";

  var authHdr   = {"Authorization": "Bearer " + gridToken, "Content-Type": "application/json"};
  var fileBytes = Utilities.newBlob(htmlContent, "text/html").getBytes();
  var fileSize  = fileBytes.length;
  var ep        = "https://grid.melioffice.com/api/v1/engine/run/json";

  try {
    var body1 = JSON.stringify({
      skill_version    : cfg.gridSkillVersion,
      doc_id           : gridDocId,
      file_new_version : true,
      presigned_upload : { filename: cfg.driveFileName, content_type: "text/html", file_size: fileSize }
    });
    var r1 = UrlFetchApp.fetch(ep, {method:"post", headers:authHdr, payload:body1, muteHttpExceptions:true});
    var c1 = r1.getResponseCode();
    if (c1 >= 400) return "paso1 HTTP " + c1 + ": " + r1.getContentText().slice(0,120);
    var d1 = JSON.parse(r1.getContentText());
    if (!d1.data || !d1.data.upload_url) return "paso1 sin upload_url: " + r1.getContentText().slice(0,120);

    var r2 = UrlFetchApp.fetch(d1.data.upload_url, {
      method: "put", contentType: "text/html", payload: fileBytes, muteHttpExceptions: true
    });
    var c2 = r2.getResponseCode();
    if (c2 !== 200 && c2 !== 204) return "paso2 PUT HTTP " + c2;

    var body3 = JSON.stringify({
      skill_version: cfg.gridSkillVersion, doc_id: d1.doc_id,
      confirm_presigned: true, file_size: fileSize
    });
    var r3 = UrlFetchApp.fetch(ep, {method:"post", headers:authHdr, payload:body3, muteHttpExceptions:true});
    var d3 = JSON.parse(r3.getContentText());
    if (d3.ok) return "Subido a Grid version " + (d3.version || "OK");
    return "paso3: " + r3.getContentText().slice(0, 120);

  } catch (e) {
    return "Error: " + e.message;
  }
}

function diagnosticar() {
  var ui     = SpreadsheetApp.getUi();
  var msgs   = [];
  var tokens = getTokens();
  var gh     = getGitHubConfig();

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetName);
    msgs.push(sheet
      ? "OK: Hoja " + CONFIG.sheetName + " (" + sheet.getLastRow() + " filas)"
      : "ERROR: Hoja " + CONFIG.sheetName + " no encontrada");
  } catch(e) { msgs.push("ERROR hoja: " + e.message); }

  try {
    msgs.push(DriveApp.getFilesByName(CONFIG.srm1.driveFileName).hasNext()
      ? "OK: Drive SRM1 encontrado"
      : "ERROR: Drive SRM1 no encontrado");
  } catch(e) { msgs.push("ERROR Drive SRM1: " + e.message); }

  try {
    msgs.push(DriveApp.getFilesByName(CONFIG.srm2.driveFileName).hasNext()
      ? "OK: Drive SRM2 encontrado"
      : "ERROR: Drive SRM2 no encontrado");
  } catch(e) { msgs.push("ERROR Drive SRM2: " + e.message); }

  var hayTrigger = false;
  var triggers   = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === "onCambioHoja") { hayTrigger = true; break; }
  }
  msgs.push(hayTrigger ? "OK: Trigger automatico ACTIVO" : "AVISO: Trigger NO instalado");
  msgs.push(tokens.srm1Token ? "OK: Token Grid configurado" : "AVISO: Token Grid no configurado");
  msgs.push(tokens.srm2DocId !== "PENDIENTE" ? "OK: DocId SRM2: " + tokens.srm2DocId : "AVISO: DocId SRM2 pendiente");

  if (gh.owner && gh.repo) {
    msgs.push("OK: GitHub -> " + gh.owner + "/" + gh.repo);
    msgs.push(gh.token ? "OK: Token GitHub configurado" : "AVISO: Token GitHub no configurado");
    msgs.push("URL publica: https://" + gh.owner + ".github.io/" + gh.repo + "/");
  } else {
    msgs.push("AVISO: GitHub no configurado - usa menu Configurar GitHub");
  }

  ui.alert("Diagnostico Same Day MLC", msgs.join("\n"), ui.ButtonSet.OK);
}

function mostrarResumen(results) {
  var lines = ["Actualizacion finalizada\n"];
  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    lines.push((r.ok ? "OK" : "AVISO") + " - " + r.service + ": " + r.nuevos + " registro(s)");
    lines.push("   -> " + r.msg);
  }
  SpreadsheetApp.getUi().alert("Dashboards Same Day MLC", lines.join("\n"), SpreadsheetApp.getUi().ButtonSet.OK);
}

function guardarLog(text) {
  PropertiesService.getScriptProperties().setProperty("ultimo_log", text);
}
