// ================================================================
// CRM Alunos VR — Google Apps Script Backend v6
// Fix: normalizar datetime de células Date do Sheets → pt-BR dd/MM/yyyy HH:mm:ss
// Sincroniza: history, contacts, followups, base de alunos (centralizada), usuários
// ================================================================

var SHEET_HISTORY  = 'historico';
var SHEET_CONTACTS = 'cadastros';
var SHEET_META     = 'meta';

// ── SERVIDOR HTML — serve o app diretamente via GAS ─────────────
// Acesse: URL_DO_GAS?action=app
// Isso permite usar o GAS como host do HTML sem depender do GitHub
function serveApp() {
  try {
    // O HTML está embedado como propriedade ou retorna redirect
    var html = HtmlService.createHtmlOutput(getAppHtml())
      .setTitle('EngajaVR — Estácio Volta Redonda')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    return html;
  } catch(e) {
    return HtmlService.createHtmlOutput('<p>Erro ao carregar: ' + e + '</p>');
  }
}

function getAppHtml() {
  // Retorna o HTML armazenado nas PropertiesService ou DriveApp
  var raw = storeGet('AppHTML');
  if (raw) return raw;
  return '<html><body><h2>EngajaVR</h2><p>HTML não carregado. Contate o administrador.</p></body></html>';
}

// ── ROTEADOR ────────────────────────────────────────────────────
function doGet(e) {
  // Se acessado sem parâmetros ou com action=app, servir o HTML
  if (!e || !e.parameter || !e.parameter.action || e.parameter.action === 'app') {
    return serveApp();
  }
  return handleRequest(e);
}
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  var action = 'ping';
  var body   = {};
  try {
    if (e && e.parameter && e.parameter.action) action = e.parameter.action;
    if (e && e.postData && e.postData.contents) {
      var raw = e.postData.contents;
      body = JSON.parse(raw);
      if (body.action) action = body.action;
    } else if (e && e.parameter) {
      // GET com parâmetros
      body = e.parameter;
    }
  } catch(err) {
    return respond({ ok: false, error: 'Erro ao parsear JSON: ' + err.message });
  }
  try {
    var result;
    switch(action) {
      case 'ping':            result = ping();                  break;
      case 'getHistory':      result = getHistory(body);        break;
      case 'getContacts':     result = getContacts();           break;
      case 'saveAction':      result = saveAction(body);        break;
      case 'saveContact':     result = saveContact(body);       break;
      case 'clearAll':        result = clearAll();              break;
      case 'getFollowups':    result = getFollowups(body);      break;
      case 'saveFollowup':    result = saveFollowup(body);      break;
      case 'saveEventos':     result = saveEventos(body);       break;
      case 'getEventos':      result = getEventos();            break;
      case 'saveModelos':     result = saveModelos(body);       break;
      case 'getModelos':      result = getModelos();            break;
      case 'saveAtendentes':  result = saveAtendentes(body);   break;
      case 'getAtendentes':   result = getAtendentes();        break;
      case 'renameAtendente': result = renameAtendente(body);  break;
      case 'saveHistoryFull': result = saveHistoryFull(body);   break;
      case 'saveRelatEmail':  result = saveRelatEmail(body);    break;
      case 'uploadAppHTML':   result = uploadAppHTML(body);    break;
      case 'publicarBase':    result = publicarBase(body);     break;
      case 'getBasePublicada': result = getBasePublicada(body); break;
      case 'importarCSVSharePoint': result = importarCSVSharePoint(body); break;
      case 'getSharePointCSV':      result = getSharePointCSV(body);      break;
      case 'getTarefas':      result = getTarefas();            break;
      case 'saveTarefas':     result = saveTarefas(body);       break;
      case 'getTutoriais':    result = getTutoriais();          break;
      case 'saveTutoriais':   result = saveTutoriais(body);     break;
      case 'reconstruirHistory': result = reconstruirHistory(body); break;
      case 'diagnostico':       result = diagnostico();          break;
      case 'restaurarDeAbas':   result = restaurarDeAbas(body);  break;
      case 'saveActionLote':  result = saveActionLote(body);    break;
      case 'sendEmail':       result = sendEmailAction(body);   break;
      case 'sendEmailLote':   result = sendEmailLote(body);     break;
      case 'salvarBase':   result = salvarBase(body.base, body.dados, body.parte || 0, body.total || 1); break;
      case 'carregarBase': result = carregarBase(body.base); break;
      case 'getBaseTipo':  result = getBaseTipo(body.base);  break;
      case 'getUsuarios':  result = getUsuarios(); break;
      case 'saveUsuarios': result = saveUsuarios(body); break;
      case 'getTCEData':      result = getTCEData(body.base); break;
      case 'saveTCEData':     result = saveTCEData(body.base, body.data); break;
      case 'getEvasaoData':    result = getEvasaoData(body.base); break;
      case 'saveEvasaoData':   result = saveEvasaoData(body.base, body.data); break;
      case 'mergeEvasaoData':  result = mergeEvasaoData(body.base, body.data); break;
      case 'saveSnaps':       result = saveSnaps(body.base, body.snaps);      break;
      case 'getSnaps':        result = getSnaps(body.base);                   break;
            default: result = { ok: false, error: 'Ação desconhecida: ' + action };
    }
    return respond(result);
  } catch(err) {
    return respond({ ok: false, error: err.message });
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── UTILITÁRIOS ──────────────────────────────────────────────────

// Normaliza qualquer valor de data para string pt-BR "dd/MM/yyyy HH:mm:ss"
// Aceita: Date object do Sheets, string ISO, string pt-BR, string JS Date.toString()
function normalizarDatetime(val) {
  var tz = Session.getScriptTimeZone();
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, tz, 'dd/MM/yyyy HH:mm:ss');
  }
  var s = String(val);
  // Já está no formato pt-BR dd/MM/yyyy
  if (s.match(/^\d{2}\/\d{2}\/\d{4}/)) return s;
  // ISO: yyyy-mm-dd
  var mISO = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (mISO) return mISO[3] + '/' + mISO[2] + '/' + mISO[1];
  // JS Date.toString(): "Wed Sep 16 2026 18:35:01 GMT..."
  var ts = Date.parse(s);
  if (!isNaN(ts)) {
    return Utilities.formatDate(new Date(ts), tz, 'dd/MM/yyyy HH:mm:ss');
  }
  return s;
}

// Extrai apenas a parte de data (dd/MM/yyyy) de um valor normalizado
function normalizarDate(val) {
  var dt = normalizarDatetime(val);
  if (!dt) return '';
  return dt.split(' ')[0] || dt.split(',')[0] || dt;
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Planilha não encontrada.');
  return ss;
}

function getOrCreate(name, headers) {
  var ss = getSheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.appendRow(headers);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold');
    }
  }
  return sh;
}

// ── STORAGE ──────────────────────────────────────────────────────
// Hierarquia de armazenamento:
//   1. PropertiesService  → dados pequenos (<9KB): snapshots, timestamps, configs
//   2. Google Drive (JSON) → dados grandes (base de alunos, histórico)
//      Arquivo nomeado "engajavr_<key>.json" na pasta raiz do Drive do script
//   3. Planilha em chunks → fallback legado (mantido para leitura retrocompatível)
//
// Esta abordagem elimina o erro "20.000.000 de células" do Sheets.

// Retorna (ou cria) a pasta "EngajaVR_Storage" no Drive do usuário proprietário do script
function getDriveFolder() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('_drive_folder_id');
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch(e) {}
  }
  // Cria a pasta uma única vez e grava o ID
  var folder = DriveApp.createFolder('EngajaVR_Storage');
  props.setProperty('_drive_folder_id', folder.getId());
  return folder;
}

// Lê o arquivo "engajavr_<key>.json" do Drive
function driveGet(key) {
  try {
    var folder = getDriveFolder();
    var files = folder.getFilesByName('engajavr_' + key + '.json');
    if (files.hasNext()) {
      var f = files.next();
      return f.getBlob().getDataAsString('UTF-8');
    }
  } catch(e) { Logger.log('driveGet erro: ' + e.message); }
  return null;
}

// Grava (ou substitui) "engajavr_<key>.json" no Drive
function driveSet(key, value) {
  try {
    var folder = getDriveFolder();
    var fileName = 'engajavr_' + key + '.json';
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch(e) {}
    try {
      var files = folder.getFilesByName(fileName);
      while (files.hasNext()) { files.next().setTrashed(true); } // remove versão antiga
      folder.createFile(fileName, value, 'application/json');
    } finally {
      try { lock.releaseLock(); } catch(e) {}
    }
    return true;
  } catch(e) {
    Logger.log('driveSet erro: ' + e.message);
    return false;
  }
}

function storeGet(key) {
  // 1. PropertiesService (rápido, para valores pequenos)
  try {
    var v = PropertiesService.getScriptProperties().getProperty(key);
    if (v) return v;
  } catch(e) {}

  // 2. Drive (para valores grandes — base de alunos, histórico)
  var dv = driveGet(key);
  if (dv) return dv;

  // 3. Planilha em chunks (retrocompatibilidade com dados gravados pelo código antigo)
  try {
    var ss = getSheet();
    var shName = '_kv_' + key;
    var sh = ss.getSheetByName(shName);
    if (sh) {
      var result = '';
      var lastRow = sh.getLastRow();
      for (var i = 1; i <= lastRow; i++) result += sh.getRange(i, 1).getValue();
      if (result) {
        // Migra para Drive e remove a aba da planilha para liberar células
        driveSet(key, result);
        try { ss.deleteSheet(sh); } catch(e) {}
        return result;
      }
    }
  } catch(e) {}

  return null;
}

function storeSet(key, value) {
  var val = String(value || '');

  // 1. PropertiesService para valores pequenos (<9KB)
  if (val.length < 9000) {
    try {
      PropertiesService.getScriptProperties().setProperty(key, val);
      return;
    } catch(e) {}
  }

  // 2. Drive para valores grandes (sem limite de células do Sheets)
  driveSet(key, val);

  // Remove chave legada do PropertiesService se existir (evita conflito de leitura)
  try { PropertiesService.getScriptProperties().deleteProperty(key); } catch(e) {}
}

// ── PING ────────────────────────────────────────────────────────
function ping() {
  return { ok: true, message: 'Conectado com sucesso!', planilha: getSheet().getName(), ts: Date.now() };
}

// ── HISTÓRICO ───────────────────────────────────────────────────
function getHistory(data) {
  var base = (data && data.base) || 'impar';
  var hist = {};

  // FONTE 0: _kv_HistoryJSON_{base} (chunks — fonte primária da base solicitada)
  var rawPrimario = storeGet('HistoryJSON_' + base);
  if (rawPrimario) {
    try {
      var primario = JSON.parse(rawPrimario);
      Object.keys(primario).forEach(function(mat) {
        if (!hist[mat]) hist[mat] = [];
        var existIds = {};
        hist[mat].forEach(function(a){ existIds[a.id||(a.datetime+a.atend)] = true; });
        primario[mat].forEach(function(a) {
          var k = a.id||(a.datetime+a.atend);
          if (!existIds[k]) hist[mat].push(a);
        });
      });
    } catch(e) {}
  }

  // FONTE 1: Aba "historico" legada (acionamentos anteriores ao GAS v4)
  var shH = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('historico');
  if (shH && shH.getLastRow() > 1) {
    var rows = shH.getDataRange().getValues();
    var hdrs = rows[0].map(String);
    for (var i = 1; i < rows.length; i++) {
      var obj = {};
      hdrs.forEach(function(h,j){ obj[h] = rows[i][j]; });
      var mat = String(obj['matricula'] || obj['COD_MATRICULA'] || '');
      if (!mat) continue;
      if (!hist[mat]) hist[mat] = [];
      hist[mat].push({
        id:       obj['id'] || String(obj['datetime']||'') + '_' + mat,
        datetime: String(obj['datetime'] || ''),
        date:     String(obj['date'] || (String(obj['datetime']||'')).split(' ')[0] || ''),
        atend:    String(obj['atend'] || obj['atendente'] || ''),
        canal:    String(obj['canal'] || obj['tipo'] || ''),
        tipo:     String(obj['canal'] || obj['tipo'] || ''),
        motivo:   String(obj['motivo'] || ''),
        obs:      String(obj['obs'] || ''),
        nps:      String(obj['nps'] || ''),
        nps_tipo: '',
        matricula: mat,
        base:     String(obj['base'] || 'impar')
      });
    }
  }

  // FONTE 2: Aba "Acionamentos" (nova — filtrar pela base solicitada)
  var shA = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Acionamentos');
  if (shA && shA.getLastRow() > 1) {
    var rowsA = shA.getDataRange().getValues();
    for (var i = 1; i < rowsA.length; i++) {
      var r = rowsA[i];
      var mat = String(r[1] || '');
      var rowBase = String(r[6] || 'impar');
      if (!mat || rowBase !== base) continue;
      if (!hist[mat]) hist[mat] = [];
      var dtNorm = normalizarDatetime(r[0]);
      var entry = {
        id:       String(r[8] || dtNorm + '_' + mat),
        datetime: dtNorm,
        date:     normalizarDate(r[0]),
        atend:    String(r[2] || ''),
        canal:    String(r[3] || ''),
        tipo:     String(r[3] || ''),
        motivo:   String(r[4] || ''),
        obs:      String(r[5] || ''),
        nps:      String(r[7] || ''),
        nps_tipo: '',
        matricula: mat,
        base:     rowBase
      };
      var exists = hist[mat].some(function(a){ return a.id === entry.id; });
      if (!exists) hist[mat].push(entry);
    }
  }

  return { ok: true, data: hist, ts: Date.now() };
}

function saveAction(body) {
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch(e) { return { ok: false, error: 'JSON inválido' }; } }
  if (!body) return { ok: false, error: 'body ausente' };
  var base = body.base || 'impar';
  var key  = 'HistoryJSON_' + base;

  // Salvar na planilha Acionamentos
  var sh = getOrCreate('Acionamentos',
    ['datetime','matricula','atendente','canal','motivo','obs','base','nps','id']);
  var _tz = Session.getScriptTimeZone();
  var _now = new Date();
  var entry = {
    id:       body.id        || (Date.now() + '_' + Math.random().toString(36).substr(2,5)),
    datetime: body.datetime  || Utilities.formatDate(_now, _tz, 'dd/MM/yyyy HH:mm:ss'),
    date:     body.date      || Utilities.formatDate(_now, _tz, 'dd/MM/yyyy'),
    atend:    body.atend     || '',
    canal:    body.canal     || body.tipo || '',
    tipo:     body.canal     || body.tipo || '',
    motivo:   body.motivo    || '',
    obs:      body.obs       || '',
    nps:      body.nps       !== undefined ? body.nps : '',
    nps_tipo: body.nps_tipo  || '',
    matricula: body.matricula || '',
    base:     base
  };
  sh.appendRow([entry.datetime, entry.matricula, entry.atend, entry.canal,
                entry.motivo, entry.obs, entry.base, entry.nps, entry.id]);

  // Atualizar HistoryJSON
  var raw  = storeGet(key);
  var hist = raw ? JSON.parse(raw) : {};
  var mat  = entry.matricula;
  if (mat) {
    if (!hist[mat]) hist[mat] = [];
    hist[mat].push(entry);
    storeSet(key, JSON.stringify(hist));
  }
  return { ok: true, id: entry.id };
}

function saveActionLote(data) {
  var entries = data.entries;
  if (!entries || !entries.length) return { ok: false, error: 'Nenhum entry' };
  var base = data.base || 'impar';
  var key  = 'HistoryJSON_' + base;
  var raw  = storeGet(key);
  var hist = raw ? JSON.parse(raw) : {};
  var sh   = getOrCreate('Acionamentos',
    ['datetime','matricula','atendente','canal','motivo','obs','base','nps','id']);

  entries.forEach(function(e) {
    sh.appendRow([e.datetime||'', e.matricula||'', e.atend||'',
                  e.canal||'', e.motivo||'', e.obs||'', e.base||base, e.nps||'', e.id||'']);
    if (e.matricula) {
      if (!hist[e.matricula]) hist[e.matricula] = [];
      hist[e.matricula].push(e);
    }
  });
  storeSet(key, JSON.stringify(hist));
  return { ok: true, saved: entries.length };
}

function saveHistoryFull(data) {
  if (!data.history) return { ok: false, error: 'history obrigatório' };
  var base = data.base || 'impar';
  storeSet('HistoryJSON_' + base, JSON.stringify(data.history));
  // Reconstruir planilha
  var sh = getOrCreate('Acionamentos',
    ['datetime','matricula','atendente','canal','motivo','obs','base','nps','id']);
  var lastRow = sh.getLastRow();
  if (lastRow > 1) sh.deleteRows(2, lastRow - 1);
  Object.entries(data.history).forEach(function(entry) {
    var mat = entry[0], arr = entry[1];
    arr.forEach(function(a) {
      sh.appendRow([a.datetime||'', mat, a.atend||'', a.canal||'',
                    a.motivo||'', a.obs||'', a.base||base, a.nps||'', a.id||'']);
    });
  });
  return { ok: true };
}

// ── DIAGNÓSTICO — listar tudo que existe ───────────────────────
function diagnostico() {
  var ss = getSheet();
  var sheets = ss.getSheets().map(function(s){ return s.getName(); });

  // Verificar PropertiesService
  var props = PropertiesService.getScriptProperties().getProperties();
  var propKeys = Object.keys(props);

  // Contar acionamentos em cada fonte
  var resultado = {
    abas: sheets,
    propKeys: propKeys,
    fontes: {}
  };

  // Aba Acionamentos
  var shAc = ss.getSheetByName('Acionamentos');
  if (shAc) resultado.fontes['aba_Acionamentos'] = shAc.getLastRow() - 1 + ' linhas';

  // Aba historico (legada)
  var shH = ss.getSheetByName('historico');
  if (shH) resultado.fontes['aba_historico'] = shH.getLastRow() - 1 + ' linhas';

  // PropertiesService keys com 'History'
  propKeys.forEach(function(k) {
    if (k.indexOf('History') >= 0 || k.indexOf('history') >= 0) {
      try {
        var v = JSON.parse(props[k]);
        var total = Object.values(v).reduce(function(s,arr){ return s+(arr.length||0); }, 0);
        resultado.fontes['prop_' + k] = total + ' acionamentos';
      } catch(e) {
        resultado.fontes['prop_' + k] = 'erro ao parsear';
      }
    }
  });

  // Abas _kv_ (chunks grandes)
  sheets.forEach(function(name) {
    if (name.indexOf('_kv_') === 0) {
      var sh = ss.getSheetByName(name);
      resultado.fontes['chunk_' + name] = sh.getLastRow() + ' chunks';
    }
  });

  return { ok: true, resultado: resultado };
}

// ── RESTAURAR DE QUALQUER ABA/FONTE ─────────────────────────────
function restaurarDeAbas(data) {
  var ss = getSheet();
  var hist = {};
  var total = 0;

  // FONTE 1: _kv_HistoryJSON_impar (chunks — onde os dados realmente estão)
  var rawKv = storeGet('HistoryJSON_impar');
  if (rawKv) {
    try {
      var kvData = JSON.parse(rawKv);
      Object.keys(kvData).forEach(function(mat) {
        if (!hist[mat]) hist[mat] = [];
        kvData[mat].forEach(function(a) { hist[mat].push(a); total++; });
      });
    } catch(e) { Logger.log('Erro HistoryJSON_impar: ' + e); }
  }

  // FONTE 2: HistoryJSON legado (sem sufixo)
  var rawLeg = storeGet('HistoryJSON');
  if (rawLeg) {
    try {
      var leg = JSON.parse(rawLeg);
      Object.keys(leg).forEach(function(mat) {
        if (!hist[mat]) hist[mat] = [];
        var existIds = new Set(hist[mat].map(function(a){ return a.id||(a.datetime+a.atend); }));
        leg[mat].forEach(function(a) {
          var k = a.id||(a.datetime+a.atend);
          if (!existIds.has(k)) { hist[mat].push(a); total++; }
        });
      });
    } catch(e) {}
  }

  // FONTE 3: Aba historico (GAS v2)
  var shH = ss.getSheetByName('historico');
  if (shH && shH.getLastRow() > 1) {
    var rows = shH.getDataRange().getValues();
    var hdrs = rows[0].map(String);
    for (var i = 1; i < rows.length; i++) {
      var obj = {}; hdrs.forEach(function(h,j){ obj[h]=rows[i][j]; });
      var mat = String(obj['matricula']||obj['COD_MATRICULA']||'');
      if (!mat) continue;
      if (!hist[mat]) hist[mat] = [];
      var e1 = { id: obj['id']||String(obj['datetime']||'')+'_'+mat,
        datetime:String(obj['datetime']||''), date:String(obj['date']||(String(obj['datetime']||'')).split(' ')[0]||''),
        atend:String(obj['atend']||obj['atendente']||''), canal:String(obj['canal']||obj['tipo']||''),
        tipo:String(obj['canal']||obj['tipo']||''), motivo:String(obj['motivo']||''),
        obs:String(obj['obs']||''), nps:String(obj['nps']||''), nps_tipo:'',
        matricula:mat, base:String(obj['base']||'impar') };
      if (!hist[mat].some(function(a){ return a.id===e1.id; })) { hist[mat].push(e1); total++; }
    }
  }

  // FONTE 4: Aba Acionamentos (nova)
  var shA = ss.getSheetByName('Acionamentos');
  if (shA && shA.getLastRow() > 1) {
    var rowsA = shA.getDataRange().getValues();
    for (var i = 1; i < rowsA.length; i++) {
      var r = rowsA[i];
      var mat = String(r[1]||'');
      if (!mat) continue;
      if (!hist[mat]) hist[mat] = [];
      var dt2 = normalizarDatetime(r[0]);
      var e2 = { id:String(r[8]||dt2+'_'+mat),
        datetime:dt2, date:normalizarDate(r[0]),
        atend:String(r[2]||''), canal:String(r[3]||''), tipo:String(r[3]||''),
        motivo:String(r[4]||''), obs:String(r[5]||''), nps:String(r[7]||''),
        nps_tipo:'', matricula:mat, base:String(r[6]||'impar') };
      if (!hist[mat].some(function(a){ return a.id===e2.id; })) { hist[mat].push(e2); total++; }
    }
  }

  if (total === 0) return { ok: false, error: 'Nenhum dado encontrado. Fontes verificadas: _kv_HistoryJSON_impar, HistoryJSON, historico, Acionamentos' };
  var totalFinal = Object.values(hist).reduce(function(s,arr){ return s+arr.length; },0);
  return { ok: true, recuperados: total, totalFinal: totalFinal, data: hist };
}


// ── RECONSTRUIR HISTORY DA PLANILHA ────────────────────────────
// Lê a aba Acionamentos e reconstrói o HistoryJSON por base
function reconstruirHistory(data) {
  // Ler o HistoryJSON legado (salvo pelo GAS anterior sem sufixo de base)
  var rawLegado = storeGet('HistoryJSON');
  var total = 0;
  var resultado = {};

  if (!rawLegado) {
    // Tentar também via células da planilha legada
    return { ok: true, total: 0, msg: 'Nenhum histórico legado encontrado em HistoryJSON' };
  }

  var legado = {};
  try { legado = JSON.parse(rawLegado); } catch(e) {
    return { ok: false, error: 'Erro ao parsear HistoryJSON: ' + e.message };
  }

  // Mesclar legado com HistoryJSON_impar (assume base impar para legado)
  var keyImpar = 'HistoryJSON_impar';
  var rawAtual = storeGet(keyImpar);
  var atual = rawAtual ? JSON.parse(rawAtual) : {};

  Object.keys(legado).forEach(function(mat) {
    if (!atual[mat]) atual[mat] = [];
    var existIds = new Set(atual[mat].map(function(a){ return a.id || (a.datetime+a.atend); }));
    legado[mat].forEach(function(a) {
      var k = a.id || (a.datetime+a.atend);
      if (!existIds.has(k)) {
        // Garantir campo base
        if (!a.base) a.base = 'impar';
        atual[mat].push(a);
        total++;
      }
    });
  });

  storeSet(keyImpar, JSON.stringify(atual));
  resultado['impar'] = Object.keys(atual).length + ' alunos, ' + total + ' novos acionamentos mesclados';

  return { ok: true, total: total, bases: resultado };
}

// ── CONTATOS ────────────────────────────────────────────────────
function getContacts() {
  var sh = getOrCreate('cadastros',
    ['matricula','cel','fone','email','atualizado_em']);
  var rows = sh.getDataRange().getValues();
  if (rows.length < 2) return { ok: true, data: {}, ts: Date.now() };
  var hdrs = rows[0].map(String);
  var result = {};
  for (var i = 1; i < rows.length; i++) {
    var mat = String(rows[i][0] || '');
    if (!mat) continue;
    var obj = {};
    for (var j = 0; j < hdrs.length; j++) obj[hdrs[j]] = rows[i][j];
    result[mat] = obj;
  }
  return { ok: true, data: result, ts: Date.now() };
}

function saveContact(body) {
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch(e) { return { ok: false, error: 'JSON inválido' }; } }
  if (!body) return { ok: false, error: 'body ausente' };
  var sh  = getOrCreate('cadastros',
    ['matricula','cel','fone','email','atualizado_em']);
  var rows = sh.getDataRange().getValues();
  var mat  = String(body.matricula || '');
  if (!mat) return { ok: false, error: 'matricula obrigatória' };
  var now  = new Date().toISOString();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === mat) {
      sh.getRange(i+1,2,1,4).setValues([[body.cel||'',body.fone||'',body.email||'',now]]);
      return { ok: true, updated: true };
    }
  }
  sh.appendRow([mat, body.cel||'', body.fone||'', body.email||'', now]);
  return { ok: true, created: true };
}

// ── FOLLOW-UPS ──────────────────────────────────────────────────
function getFollowups(data) {
  var base = (data && data.base) || 'impar';
  var raw  = storeGet('FollowupsJSON_' + base);
  return { ok: true, data: raw ? JSON.parse(raw) : {} };
}

function saveFollowup(data) {
  // Defensive: data pode vir como string ou objeto
  if (typeof data === 'string') {
    try { data = JSON.parse(data); } catch(e) { return { ok: false, error: 'JSON inválido: ' + e.message }; }
  }
  if (!data || !data.followups) return { ok: false, error: 'followups ausente', received: JSON.stringify(data) };
  var base = data.base || 'impar';
  storeSet('FollowupsJSON_' + base, JSON.stringify(data.followups));
  return { ok: true, base: base };
}

// ── CALENDÁRIO DE EVENTOS ───────────────────────────────────────
function saveEventos(data) {
  if (!data || !data.eventos) return { ok: false, error: 'eventos ausente' };
  storeSet('EventosJSON', JSON.stringify(data.eventos));
  return { ok: true, saved: data.eventos.length };
}

function getEventos() {
  var raw = storeGet('EventosJSON');
  return { ok: true, data: raw ? JSON.parse(raw) : [] };
}

// ── MODELOS DE MENSAGEM ─────────────────────────────────────────
function saveModelos(data) {
  if (!data || !data.modelos) return { ok: false, error: 'modelos ausente' };
  storeSet('ModelosJSON', JSON.stringify(data.modelos));
  return { ok: true, saved: data.modelos.length };
}

function getModelos() {
  var raw = storeGet('ModelosJSON');
  return { ok: true, data: raw ? JSON.parse(raw) : [] };
}

// ── ATENDENTES ──────────────────────────────────────────────────
function saveAtendentes(data) {
  if (!data || !data.atendentes) return { ok: false, error: 'atendentes ausente' };
  storeSet('AtendentesJSON', JSON.stringify(data.atendentes));
  return { ok: true, saved: data.atendentes.length };
}

function getAtendentes() {
  var raw = storeGet('AtendentesJSON');
  return { ok: true, data: raw ? JSON.parse(raw) : [] };
}

// Renomear atendente em TODOS os registros do GAS
function renameAtendente(data) {
  if (!data || !data.antigo || !data.novo) return { ok: false, error: 'antigo e novo são obrigatórios' };
  var antigo = data.antigo;
  var novo   = data.novo;
  var count  = 0;

  // 1. Renomear na lista de atendentes
  var rawAtends = storeGet('AtendentesJSON');
  if (rawAtends) {
    try {
      var atends = JSON.parse(rawAtends);
      var idx = atends.indexOf(antigo);
      if (idx >= 0) { atends[idx] = novo; atends.sort(); storeSet('AtendentesJSON', JSON.stringify(atends)); }
    } catch(e) {}
  }

  // 2. Renomear em todos os HistoryJSON (impar, par, ct)
  ['impar','par','ct'].forEach(function(base) {
    var raw = storeGet('HistoryJSON_' + base);
    if (!raw) return;
    try {
      var hist = JSON.parse(raw);
      var changed = false;
      Object.keys(hist).forEach(function(mat) {
        hist[mat].forEach(function(a) {
          if (a.atend === antigo) { a.atend = novo; changed = true; count++; }
        });
      });
      if (changed) storeSet('HistoryJSON_' + base, JSON.stringify(hist));
    } catch(e) {}
  });

  // 3. Renomear na aba Acionamentos
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Acionamentos');
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][2]) === antigo) {
          sh.getRange(i+1, 3).setValue(novo);
          count++;
        }
      }
    }
  } catch(e) {}

  // 4. Renomear em FollowupsJSON
  ['impar','par','ct'].forEach(function(base) {
    var raw = storeGet('FollowupsJSON_' + base);
    if (!raw) return;
    try {
      var fus = JSON.parse(raw);
      var changed = false;
      Object.keys(fus).forEach(function(mat) {
        fus[mat].forEach(function(fu) {
          if (fu.atend === antigo) { fu.atend = novo; changed = true; count++; }
        });
      });
      if (changed) storeSet('FollowupsJSON_' + base, JSON.stringify(fus));
    } catch(e) {}
  });

  // 5. Renomear em Tarefas
  var rawTf = storeGet('Tarefas');
  if (rawTf) {
    try {
      var tarefas = JSON.parse(rawTf);
      var changed = false;
      tarefas.forEach(function(t) {
        if (t.atend === antigo) { t.atend = novo; changed = true; count++; }
        if (t.criadoPor === antigo) { t.criadoPor = novo; changed = true; }
      });
      if (changed) storeSet('Tarefas', JSON.stringify(tarefas));
    } catch(e) {}
  }

  Logger.log('✅ Renomeado "' + antigo + '" → "' + novo + '" em ' + count + ' registros no GAS');
  return { ok: true, renomeados: count };
}

// ── LIMPAR TUDO ──────────────────────────────────────────────────
function clearAll() {
  var ss = getSheet();
  ['historico','cadastros','meta','Acionamentos','Emails'].forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });
  // Limpar PropertiesService
  PropertiesService.getScriptProperties().deleteAllProperties();
  return { ok: true };
}

// ── ENVIO DE E-MAIL ──────────────────────────────────────────────
function sendEmailAction(data) {
  if (typeof data === 'string') { try { data = JSON.parse(data); } catch(e) { return { ok: false, error: 'JSON inválido' }; } }
  if (!data || !data.to) return { ok: false, error: 'Destinatário ausente', received: JSON.stringify(data) };
  var to  = data.to;
  var sub = data.subject || 'CRM Alunos VR';
  var txt = data.body    || '';
  GmailApp.sendEmail(to, sub, txt, { name: 'CRM Alunos — Estácio VR' });
  return { ok: true };
}

function sendEmailLote(data) {
  if (typeof data === 'string') { try { data = JSON.parse(data); } catch(e) { return { ok: false, error: 'JSON inválido' }; } }
  if (!data) return { ok: false, error: 'data ausente' };
  var emails = data.emails;
  if (!emails || !emails.length) return { ok: false, error: 'Nenhum e-mail', received: JSON.stringify(data) };
  var enviados = 0, erros = [];
  emails.forEach(function(item) {
    if (!item.to) return;
    try {
      GmailApp.sendEmail(item.to, item.subject||'CRM Alunos VR', item.body||'',
        { name: 'CRM Alunos — Estácio VR' });
      enviados++;
    } catch(e) { erros.push(item.to + ': ' + e.message); }
  });
  return { ok: true, enviados: enviados, erros: erros.length ? erros : null };
}

// ── RELATÓRIO DIÁRIO AUTOMÁTICO ──────────────────────────────────
// Trigger: enviarRelatorioDiario | Diário | 18h–19h
function enviarRelatorioDiario() {
  var tz = Session.getScriptTimeZone();
  var agora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  var email = PropertiesService.getScriptProperties().getProperty('RELAT_EMAIL')
              || 'marques.marcela@estacio.br';

  Logger.log('[' + agora + '] Iniciando envio do relatório para: ' + email);
  logRelatório('INÍCIO', agora, email, '—');

  var relat;
  try {
    relat = buildRelatorioDiario();
  } catch(e) {
    Logger.log('Erro ao gerar relatório: ' + e);
    logRelatório('ERRO buildRelatorioDiario', agora, email, e.toString());
    return;
  }

  if (!relat) {
    Logger.log('Sem acionamentos hoje — relatório não enviado.');
    logRelatório('SEM ACIONAMENTOS', agora, email, 'Nenhum acionamento registrado hoje');
    return;
  }

  try {
    // MailApp é mais confiável em triggers que GmailApp
    MailApp.sendEmail({
      to: email,
      subject: relat.subject,
      body: relat.body,
      htmlBody: relat.html,
      name: 'EngajaVR — Estácio Volta Redonda'
    });
    Logger.log('✅ Relatório enviado com sucesso para: ' + email);
    logRelatório('ENVIADO', agora, email, relat.subject);
  } catch(e) {
    Logger.log('❌ Erro ao enviar: ' + e);
    logRelatório('ERRO ENVIO', agora, email, e.toString());
  }
}

// Registra log de envio na planilha para diagnóstico
function logRelatório(status, horario, email, obs) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('log_relatorio');
    if (!sh) {
      sh = ss.insertSheet('log_relatorio');
      sh.appendRow(['Data/Hora', 'Status', 'Destinatário', 'Observação']);
    }
    sh.appendRow([horario, status, email, obs]);
  } catch(e) {}
}

// Criar trigger diário automaticamente — rodar UMA vez após implantar o GAS
function criarTriggerDiario() {
  // Remover triggers antigos da mesma função para não duplicar
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarRelatorioDiario') {
      ScriptApp.deleteTrigger(t);
      Logger.log('Trigger antigo removido.');
    }
  });
  // Criar novo trigger às 18h
  ScriptApp.newTrigger('enviarRelatorioDiario')
    .timeBased()
    .everyDays(1)
    .atHour(20)
    .inTimezone(Session.getScriptTimeZone())
    .create();
  Logger.log('✅ Trigger criado: enviarRelatorioDiario todo dia às 20h');
}

// Testar envio imediato (para verificar se está funcionando)
function testarRelatorio() {
  var email = PropertiesService.getScriptProperties().getProperty('RELAT_EMAIL')
              || 'marques.marcela@estacio.br';
  Logger.log('Testando envio para: ' + email);
  enviarRelatorioDiario();
  Logger.log('Teste concluído — verifique a caixa de entrada e a aba log_relatorio na planilha');
}

function buildRelatorioDiario() {
  var tz        = Session.getScriptTimeZone();
  var hoje      = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  var hojeLabel = Utilities.formatDate(new Date(), tz, "EEEE, dd 'de' MMMM 'de' yyyy");
  var baseMap   = { impar:'Semestre Ímpar', par:'Semestre Par', ct:'Curso Técnico' };
  var atends = {}, total = 0;
  var seenIds = {};

  // FONTE PRINCIPAL: aba Acionamentos (mais confiável, não tem problema de chunks)
  try {
    var shA = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Acionamentos');
    if (shA && shA.getLastRow() > 1) {
      var rowsA = shA.getDataRange().getValues();
      // headers: datetime, matricula, atendente, canal, motivo, obs, base, nps, id
      for (var i = 1; i < rowsA.length; i++) {
        var r = rowsA[i];
        var cellVal = r[0];
        var dataAc = '';
        // r[0] pode ser Date object ou string
        if (cellVal instanceof Date) {
          dataAc = Utilities.formatDate(cellVal, tz, 'dd/MM/yyyy');
        } else {
          var d = String(cellVal || '');
          var matchBR = d.match(/(\d{2})\/(\d{2})\/(\d{4})/);
          dataAc = matchBR ? (matchBR[1] + '/' + matchBR[2] + '/' + matchBR[3]) : '';
        }
        if (!dataAc || dataAc !== hoje) continue;

        var rowId = String(r[8] || '') || (d + String(r[1]) + String(r[2]));
        if (seenIds[rowId]) continue;
        seenIds[rowId] = true;

        var nome  = String(r[2] || '(sem nome)');
        var canal = String(r[3] || 'Outro');
        var base  = String(r[6] || 'impar');
        if (!atends[nome]) atends[nome] = { total:0, bases:{}, canais:{} };
        atends[nome].total++;
        atends[nome].bases[base]   = (atends[nome].bases[base]  ||0)+1;
        atends[nome].canais[canal] = (atends[nome].canais[canal]||0)+1;
        total++;
      }
    }
  } catch(e) { Logger.log('Erro ao ler aba Acionamentos: ' + e); }

  // FONTE SECUNDÁRIA: _kv_HistoryJSON (para acionamentos que não foram para a aba)
  ['impar','par','ct'].forEach(function(base) {
    var raw = storeGet('HistoryJSON_' + base);
    if (!raw) return;
    var hist;
    try { hist = JSON.parse(raw); } catch(e) { return; }
    Object.keys(hist).forEach(function(mat) {
      (hist[mat] || []).forEach(function(a) {
        var uid = a.id || ((a.datetime||'') + mat + (a.atend||''));
        if (seenIds[uid]) return; // já contado via aba
        var d = String(a.datetime || a.date || '');
        var matchBR2 = d.match(/(\d{2})\/(\d{2})\/(\d{4})/);
        var dataAc2 = matchBR2 ? (matchBR2[1]+'/'+matchBR2[2]+'/'+matchBR2[3]) : '';
        if (!dataAc2 || dataAc2 !== hoje) return;
        seenIds[uid] = true;
        var aBase = a.base || base;
        var nome2  = a.atend || '(sem nome)';
        var canal2 = a.canal || a.tipo || 'Outro';
        if (!atends[nome2]) atends[nome2] = { total:0, bases:{}, canais:{} };
        atends[nome2].total++;
        atends[nome2].bases[aBase]   = (atends[nome2].bases[aBase]  ||0)+1;
        atends[nome2].canais[canal2] = (atends[nome2].canais[canal2]||0)+1;
        total++;
      });
    });
  });

  if (total === 0) return null;
  var sorted = Object.keys(atends).sort(function(a,b){ return atends[b].total-atends[a].total; });

  var txt = 'RELATÓRIO — EngajaVR\n' + '='.repeat(40) + '\n';
  txt += 'Data: ' + hojeLabel + '\nTotal: ' + total + ' acionamentos\n\n';
  sorted.forEach(function(nome,i) {
    var d = atends[nome];
    txt += (i===0?'🥇':i===1?'🥈':i===2?'🥉':'#'+(i+1)+'.') + ' ' + nome + ' — ' + d.total + '\n';
    Object.keys(d.bases).forEach(function(b){ txt += '  ['+(baseMap[b]||b)+'] '+d.bases[b]+'\n'; });
    txt += '\n';
  });

  return {
    subject: '📊 Produtividade ' + hoje + ' — ' + total + ' acionamentos',
    body: txt,
    html: '<html><body style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">' +
      '<h2 style="color:#3b5bdb">📊 ' + hojeLabel + '</h2>' +
      '<p><strong>' + total + ' acionamentos</strong></p>' +
      sorted.map(function(nome,i){
        var d=atends[nome];
        return '<div style="border:1px solid #e9ecef;border-radius:8px;padding:12px;margin-bottom:8px">' +
          '<strong>'+(i===0?'🥇 ':i===1?'🥈 ':i===2?'🥉 ':'#'+(i+1)+'. ')+nome+'</strong>' +
          ' <span style="color:#3b5bdb;font-size:20px;font-weight:700">'+d.total+'</span>' +
          '<div style="font-size:12px;color:#666;margin-top:4px">' +
          Object.keys(d.bases).map(function(b){ return (baseMap[b]||b)+': '+d.bases[b]; }).join(' · ') + ' | ' +
          Object.keys(d.canais).map(function(c){ return c+': '+d.canais[c]; }).join(', ') +
          '</div></div>';
      }).join('') +
      '</body></html>'
  };
}

// Admin publica a base de alunos (CSV processado) para todos os atendentes baixarem
function publicarBase(data) {
  if (!data || !data.db || !data.base) return { ok: false, error: 'db e base são obrigatórios' };
  var base = data.base; // 'impar', 'par' ou 'ct'
  storeSet('BasePublicada_' + base, JSON.stringify({
    db: data.db,
    identificacao: data.identificacao || '',
    totalAlunos: Array.isArray(data.db) ? data.db.length : 0,
    publicadoEm: new Date().toISOString(),
    publicadoPor: data.publicadoPor || ''
  }));
  Logger.log('✅ Base publicada: ' + base + ' — ' + (Array.isArray(data.db) ? data.db.length : 0) + ' alunos');
  return { ok: true, totalAlunos: Array.isArray(data.db) ? data.db.length : 0 };
}

// Atendentes buscam a base publicada mais recente
function getBasePublicada(data) {
  var base = (data && data.base) || 'impar';
  var raw = storeGet('BasePublicada_' + base);
  if (!raw) return { ok: false, error: 'Nenhuma base publicada para: ' + base };
  try {
    var obj = JSON.parse(raw);
    return { ok: true, data: obj };
  } catch(e) {
    return { ok: false, error: 'Erro ao ler base: ' + e };
  }
}

// Upload do HTML do app para o GAS (chamado pelo admin após atualização)
function uploadAppHTML(data) {
  if (!data || !data.html) return { ok: false, error: 'html ausente' };
  // Salvar em chunks no Drive (HTML pode ter 400KB+)
  storeSet('AppHTML', data.html);
  Logger.log('✅ AppHTML atualizado: ' + data.html.length + ' chars');
  return { ok: true, chars: data.html.length };
}

// ── IMPORTAÇÃO VIA SHAREPOINT (Power Automate) ───────────────────
// O Power Automate chama esta função com o conteúdo do CSV em base64
function importarCSVSharePoint(data) {
  if (!data || !data.csvBase64) return { ok: false, error: 'csvBase64 ausente' };
  var base = data.base || 'impar'; // 'impar', 'par' ou 'ct'
  var identificacao = data.identificacao || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  try {
    // Decodificar base64 → texto CSV
    var csvBytes = Utilities.base64Decode(data.csvBase64);
    var csvText = Utilities.newBlob(csvBytes).getDataAsString('UTF-8');

    // Salvar CSV raw na planilha para diagnóstico
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSh = ss.getSheetByName('log_sharepoint') || ss.insertSheet('log_sharepoint');
    logSh.appendRow([
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      base,
      identificacao,
      'CSV recebido: ' + csvText.length + ' chars'
    ]);

    // Salvar CSV na chave de SharePoint para que os clientes possam buscar
    storeSet('SharePointCSV_' + base, JSON.stringify({
      csv: csvText,
      identificacao: identificacao,
      updatedAt: new Date().toISOString(),
      base: base
    }));

    Logger.log('✅ CSV importado via SharePoint — base: ' + base + ', ' + csvText.length + ' chars');
    return { ok: true, base: base, chars: csvText.length, identificacao: identificacao };

  } catch(e) {
    Logger.log('❌ Erro importarCSVSharePoint: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// Retorna o CSV mais recente do SharePoint para um atendente importar
function getSharePointCSV(data) {
  var base = (data && data.base) || 'impar';
  var raw = storeGet('SharePointCSV_' + base);
  if (!raw) return { ok: false, error: 'Nenhum CSV disponível para base: ' + base };
  try {
    var obj = JSON.parse(raw);
    return { ok: true, data: obj };
  } catch(e) {
    return { ok: false, error: 'Erro ao ler CSV: ' + e };
  }
}


function getTarefas() {
  var raw = storeGet('Tarefas');
  if (!raw) return { ok: true, data: [] };
  try { return { ok: true, data: JSON.parse(raw) }; } catch(e) { return { ok: true, data: [] }; }
}

function saveTarefas(data) {
  if (!data || !data.tarefas) return { ok: false, error: 'sem dados' };
  storeSet('Tarefas', JSON.stringify(data.tarefas));
  return { ok: true };
}

// ── TUTORIAIS ─────────────────────────────────────────────────────
function getTutoriais() {
  var raw = storeGet('Tutoriais');
  if (!raw) return { ok: true, data: [] };
  try { return { ok: true, data: JSON.parse(raw) }; } catch(e) { return { ok: true, data: [] }; }
}

function saveTutoriais(data) {
  if (!data || !data.tutoriais) return { ok: false, error: 'sem dados' };
  storeSet('Tutoriais', JSON.stringify(data.tutoriais));
  return { ok: true };
}

// ── CONFIGURAÇÃO DE E-MAIL DO RELATÓRIO ─────────────────────────
function saveRelatEmail(data) {
  if (!data || !data.email) return { ok: false, error: 'e-mail ausente' };
  PropertiesService.getScriptProperties().setProperty('RELAT_EMAIL', data.email);
  Logger.log('E-mail do relatório configurado: ' + data.email);
  return { ok: true, email: data.email };
}

// ── CONFIGURAÇÃO ─────────────────────────────────────────────────
function configurarEmailRelatorio() {
  var email = 'marques.marcela@estacio.br';
  PropertiesService.getScriptProperties().setProperty('RELAT_EMAIL', email);
  Logger.log('✅ E-mail configurado: ' + email);
}


// ================================================================
// BASE CENTRALIZADA — v5
// Permite que coordenador importe uma vez e atendentes baixem
// automaticamente ao fazer login.
// ================================================================

var BASE_KEYS = {
  impar: 'crm_base_impar',
  par:   'crm_base_par',
  ct:    'crm_base_ct'
};

// ── SALVAR BASE (chunked) ────────────────────────────────────────
// Recebe a base em partes de 500 registros.
// parte=0 começa uma nova base; partes seguintes acumulam.
// salvarBase: acumula partes em _temp na planilha e na última parte consolida via storeSet
// Não usa DriveApp — tudo vai para abas _kv_* da planilha (sem limite prático de tamanho)
function salvarBase(base, dados, parte, total) {
  try {
    if (!base || !dados || !Array.isArray(dados)) return { ok: false, msg: 'Dados inválidos' };
    var chave = BASE_KEYS[base];
    if (!chave) return { ok: false, msg: 'Base desconhecida: ' + base };

    var acumulado = [];
    if (parte > 0) {
      try {
        var rawTemp = storeGet(chave + '_temp');
        if (rawTemp) acumulado = JSON.parse(rawTemp);
      } catch(e) { acumulado = []; }
    }

    acumulado = acumulado.concat(dados);

    if (parte < total - 1) {
      // Partes intermediárias: acumula em _temp (usando planilha se necessário)
      storeSet(chave + '_temp', JSON.stringify(acumulado));
      return { ok: true, msg: 'Parte ' + (parte + 1) + '/' + total + ' recebida.' };
    }

    // Última parte: salva definitivamente e limpa temp
    storeSet(chave, JSON.stringify(acumulado));
    // Salva timestamp de publicação (usado pelos atendentes para detectar base nova)
    var ts = new Date().getTime();
    storeSet(chave + '_ts', String(ts));
    // Limpa aba temporária se existir
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var tempSh = ss.getSheetByName('_kv_' + chave + '_temp');
      if (tempSh) ss.deleteSheet(tempSh);
    } catch(e) {}

    Logger.log('✅ Base "' + base + '" salva: ' + acumulado.length + ' registros (ts=' + ts + ')');
    return { ok: true, msg: 'Base "' + base + '" salva com ' + acumulado.length + ' registros.', ts: ts };
  } catch(e) {
    return { ok: false, msg: 'Erro ao salvar base: ' + e.message };
  }
}

// getBaseTipo: retorna apenas o timestamp da última publicação (leve, sem baixar a base inteira)
function getBaseTipo(base) {
  var chave = BASE_KEYS[base];
  if (!chave) return { ok: false, msg: 'Base desconhecida: ' + base };
  var tsRaw = storeGet(chave + '_ts');
  var ts = tsRaw ? parseInt(tsRaw) : 0;
  return { ok: true, base: base, ts: ts };
}

// carregarBase: lê via storeGet (planilha ou PropertiesService)
function carregarBase(base) {
  try {
    var chave = BASE_KEYS[base];
    if (!chave) return { ok: false, msg: 'Base desconhecida: ' + base };

    var raw = storeGet(chave);
    if (raw) {
      try { return { ok: true, dados: JSON.parse(raw) }; }
      catch(e) { return { ok: false, msg: 'Erro ao parsear base: ' + e.message }; }
    }

    return { ok: true, dados: [] };
  } catch(e) {
    return { ok: false, msg: 'Erro ao carregar base: ' + e.message };
  }
}

// ================================================================
// USUÁRIOS — perfis de acesso (admin/coordenador/apoio/atendente)
// ================================================================

// ================================================================
// TCE DATA — dados de estágio por matrícula (compartilhado entre atendentes)
// ================================================================
function getTCEData(base) {
  var key = 'TCEData_' + (base || 'impar');
  var raw = storeGet(key);
  if (!raw) return { ok: true, data: {} };
  try { return { ok: true, data: JSON.parse(raw) }; }
  catch(e) { return { ok: true, data: {} }; }
}

function saveTCEData(base, data) {
  if (!data) return { ok: false, error: 'data ausente' };
  var key = 'TCEData_' + (base || 'impar');
  storeSet(key, JSON.stringify(data));
  return { ok: true };
}

function getEvasaoData(base) {
  var key = 'EvasaoData_' + (base || 'impar');
  var raw = storeGet(key);
  if (!raw) return { ok: true, data: [] };
  try { return { ok: true, data: JSON.parse(raw) }; }
  catch(e) { return { ok: true, data: [] }; }
}

function saveEvasaoData(base, data) {
  if (!data) return { ok: false, error: 'data ausente' };
  var key = 'EvasaoData_' + (base || 'impar');
  storeSet(key, JSON.stringify(data));
  return { ok: true };
}

// Merge server-side: recebe registros do cliente, mescla com o que já está salvo
// Para cada id, vence quem tem _ts maior — evita que um atendente sobrescreva o outro
function mergeEvasaoData(base, incoming) {
  if (!incoming || !Array.isArray(incoming)) return { ok: false, error: 'data ausente' };
  var key = 'EvasaoData_' + (base || 'impar');
  var raw = storeGet(key);
  var atual = [];
  try { if (raw) atual = JSON.parse(raw); } catch(e) { atual = []; }

  // Monta mapa pelo id
  var mapa = {};
  atual.forEach(function(e){ mapa[e.id] = e; });
  incoming.forEach(function(e) {
    if (!e.id) return;
    var existente = mapa[e.id];
    if (!existente || (e._ts || 0) > (existente._ts || 0)) {
      mapa[e.id] = e;
    }
  });

  var merged = Object.keys(mapa).map(function(k){ return mapa[k]; });
  storeSet(key, JSON.stringify(merged));
  return { ok: true, total: merged.length };
}

function getUsuarios() {
  var raw = storeGet('UsuariosJSON');
  if (!raw) return { ok: true, usuarios: [] };
  try { return { ok: true, usuarios: JSON.parse(raw) }; }
  catch(e) { return { ok: true, usuarios: [] }; }
}

function saveUsuarios(data) {
  if (!data || !data.usuarios) return { ok: false, error: 'usuarios ausente' };
  storeSet('UsuariosJSON', JSON.stringify(data.usuarios));
  return { ok: true, saved: data.usuarios.length };
}

// ── SNAPSHOTS DE EVOLUÇÃO DA BASE ────────────────────────────────
// Armazena o histórico de evolução no PropertiesService do GAS
// (até 9KB por propriedade; comprimimos mantendo só os últimos 20 snaps)

function saveSnaps(base, snaps) {
  if (!base || !snaps) return { ok: false, error: 'base ou snaps ausente' };
  var key = 'snaps_' + base;
  try {
    var json = JSON.stringify(snaps);
    storeSet(key, json);
    return { ok: true, saved: snaps.length };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function getSnaps(base) {
  if (!base) return { ok: false, error: 'base ausente' };
  var key = 'snaps_' + base;
  try {
    var raw = storeGet(key);
    if (!raw) return { ok: true, snaps: [] };
    var snaps = JSON.parse(raw);
    return { ok: true, snaps: Array.isArray(snaps) ? snaps : [] };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}
