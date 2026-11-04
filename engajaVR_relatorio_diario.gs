// ================================================================
// EngajaVR — Relatório Diário Automático
// Cole este código no mesmo projeto GAS do backend.
// Configure um trigger de tempo: enviarRelatorioDiario() → todo dia às 17h30
// ================================================================

var EMAIL_DESTINO = 'marques.marcela@estacio.br';
var BASES         = ['impar', 'par', 'ct'];

// ── TRIGGER MANUAL ───────────────────────────────────────────────
// Chame uma vez para criar o trigger automático de 17h30:
//   configurarTriggerRelatorio()
function configurarTriggerRelatorio() {
  // Remove triggers antigos da mesma função para não duplicar
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarRelatorioDiario') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('enviarRelatorioDiario')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .nearMinute(30)
    .inTimezone('America/Sao_Paulo')
    .create();
  Logger.log('Trigger criado: enviarRelatorioDiario às 17h30 (Brasília)');
}

// ── FUNÇÃO PRINCIPAL ─────────────────────────────────────────────
function enviarRelatorioDiario() {
  var tz     = Session.getScriptTimeZone();
  var hoje   = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  var hojeISO= Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var dadosBases   = coletarDadosBases(hojeISO);
  var dadosHistory = coletarDadosHistory(hojeISO, tz);
  var htmlBody     = gerarHTMLRelatorio(hoje, dadosBases, dadosHistory);
  var textBody     = gerarTextoRelatorio(hoje, dadosBases, dadosHistory);

  MailApp.sendEmail({
    to:      EMAIL_DESTINO,
    subject: '📊 EngajaVR — Relatório do dia ' + hoje,
    body:    textBody,
    htmlBody: htmlBody
  });

  Logger.log('Relatório enviado para ' + EMAIL_DESTINO + ' em ' + hoje);
}

// ── COLETA DE DADOS DAS BASES ────────────────────────────────────
function coletarDadosBases(hojeISO) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var res = {};

  BASES.forEach(function(base) {
    var info = { total: 0, engajados: 0, inadimplentes: 0, evadidos: 0, renovados: 0 };
    try {
      // Tenta ler aba da base (nomes possíveis: impar, par, ct)
      var sh = ss.getSheetByName(base) || ss.getSheetByName('base_' + base);
      if (!sh) return;
      var data = sh.getDataRange().getValues();
      if (data.length < 2) return;
      var hdr = data[0].map(function(h){ return String(h).trim().toUpperCase(); });

      var iStatus    = hdr.indexOf('STATUS');
      var iSit       = hdr.indexOf('SITUACAO');
      var iInad      = hdr.indexOf('INADIMPLENTE');
      var iEvadido   = hdr.indexOf('EVADIDO');
      var iRenovado  = hdr.indexOf('RENOVADO');

      info.total = data.length - 1;
      for (var i = 1; i < data.length; i++) {
        var row  = data[i];
        var sit  = iSit >= 0  ? String(row[iSit]||'').toLowerCase()  : '';
        var st   = iStatus >= 0 ? String(row[iStatus]||'').toLowerCase() : '';
        if (iEvadido >= 0  && (row[iEvadido]===1||String(row[iEvadido]).toLowerCase()==='sim')) info.evadidos++;
        if (iInad >= 0     && (row[iInad]===1||String(row[iInad]).toLowerCase()==='sim'))      info.inadimplentes++;
        if (iRenovado >= 0 && (row[iRenovado]===1||String(row[iRenovado]).toLowerCase()==='sim')) info.renovados++;
        if (sit.includes('engaj') || st.includes('engaj')) info.engajados++;
      }
    } catch(e) {
      info.erro = e.message;
    }
    res[base] = info;
  });

  return res;
}

// ── COLETA DE DADOS DO HISTÓRICO ─────────────────────────────────
function coletarDadosHistory(hojeISO, tz) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var res    = { totalHoje: 0, porAtendente: {}, porCanal: {}, porBase: {}, ultimos: [] };

  // Aba "historico" — colunas esperadas:
  // 0:datetime 1:atendente 2:matricula 3:nome 4:canal 5:mensagem 6:base 7:tipo 8:id
  var sh = ss.getSheetByName('historico') || ss.getSheetByName('Acionamentos') || ss.getSheetByName('acionamentos');
  if (!sh) return res;

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return res;

  var hdr     = data[0].map(function(h){ return String(h).trim().toUpperCase(); });
  var iDt     = Math.max(hdr.indexOf('DATETIME'), hdr.indexOf('DATA'), 0);
  var iAtend  = hdr.indexOf('ATENDENTE');
  var iCanal  = hdr.indexOf('CANAL');
  var iBase   = hdr.indexOf('BASE');
  var iNome   = hdr.indexOf('NOME');
  var iMat    = hdr.indexOf('MATRICULA');

  var hoje    = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dtRaw = row[iDt];
    var dtStr = '';
    if (dtRaw instanceof Date) {
      dtStr = Utilities.formatDate(dtRaw, tz, 'yyyy-MM-dd');
    } else {
      var s = String(dtRaw||'');
      // dd/MM/yyyy → yyyy-MM-dd
      var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      if (m) dtStr = m[3]+'-'+m[2]+'-'+m[1];
      else   dtStr = s.split('T')[0].split(' ')[0];
    }
    if (dtStr !== hojeISO) continue;

    res.totalHoje++;
    var atend = iAtend >= 0 ? String(row[iAtend]||'Sem atendente') : 'Sem atendente';
    var canal = iCanal >= 0 ? String(row[iCanal]||'Outros')        : 'Outros';
    var base  = iBase  >= 0 ? String(row[iBase] ||'—')             : '—';
    var nome  = iNome  >= 0 ? String(row[iNome] ||'')              : '';
    var mat   = iMat   >= 0 ? String(row[iMat]  ||'')              : '';

    res.porAtendente[atend] = (res.porAtendente[atend]||0) + 1;
    res.porCanal[canal]     = (res.porCanal[canal]||0)     + 1;
    res.porBase[base]       = (res.porBase[base]||0)       + 1;
    hoje.push({ atend: atend, canal: canal, base: base, nome: nome, mat: mat });
  }

  // Últimos 10 acionamentos do dia
  res.ultimos = hoje.slice(-10).reverse();
  return res;
}

// ── GERADOR HTML ─────────────────────────────────────────────────
function gerarHTMLRelatorio(hoje, bases, hist) {
  var estiloTabela = 'width:100%;border-collapse:collapse;margin-bottom:20px;font-size:13px';
  var estiloTh     = 'background:#1c3b6e;color:#fff;padding:8px 12px;text-align:left';
  var estiloTd     = 'padding:8px 12px;border-bottom:1px solid #e9ecef';
  var estiloTdNum  = 'padding:8px 12px;border-bottom:1px solid #e9ecef;text-align:right;font-weight:700';

  // Seção de bases
  var rowsBases = '';
  BASES.forEach(function(b) {
    var d = bases[b] || {};
    rowsBases +=
      '<tr>' +
      '<td style="'+estiloTd+';font-weight:600">'+b.toUpperCase()+'</td>' +
      '<td style="'+estiloTdNum+'">'+( d.total||0 )+'</td>' +
      '<td style="'+estiloTdNum+'">'+( d.renovados||0 )+'</td>' +
      '<td style="'+estiloTdNum+'">'+( d.inadimplentes||0 )+'</td>' +
      '<td style="'+estiloTdNum+'">'+( d.evadidos||0 )+'</td>' +
      '</tr>';
  });

  // Seção de acionamentos por atendente
  var rowsAtend = '';
  Object.keys(hist.porAtendente).sort().forEach(function(a){
    rowsAtend +=
      '<tr>' +
      '<td style="'+estiloTd+'">'+a+'</td>' +
      '<td style="'+estiloTdNum+'">'+hist.porAtendente[a]+'</td>' +
      '</tr>';
  });
  if (!rowsAtend) rowsAtend = '<tr><td colspan="2" style="'+estiloTd+';color:#999">Nenhum acionamento registrado hoje.</td></tr>';

  // Seção por canal
  var rowsCanal = '';
  Object.keys(hist.porCanal).sort().forEach(function(c){
    rowsCanal += '<tr><td style="'+estiloTd+'">'+c+'</td><td style="'+estiloTdNum+'">'+hist.porCanal[c]+'</td></tr>';
  });

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>body{font-family:Arial,sans-serif;background:#f4f6f9;margin:0;padding:20px}' +
    '.card{background:#fff;border-radius:10px;padding:20px;margin-bottom:20px;box-shadow:0 2px 8px rgba(0,0,0,.08)}' +
    'h1{color:#1c3b6e;font-size:20px;margin:0 0 4px} .sub{color:#868e96;font-size:13px;margin:0 0 20px}' +
    '.badge{display:inline-block;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:700}' +
    '.badge-blue{background:#e7f5ff;color:#1864ab} .badge-green{background:#ebfbee;color:#2b8a3e}' +
    'h2{color:#1c3b6e;font-size:15px;margin:0 0 12px}' +
    '</style></head><body>' +
    '<div class="card"><h1>📊 EngajaVR — Relatório Diário</h1>' +
    '<p class="sub">Data: <strong>'+hoje+'</strong> &nbsp;|&nbsp; Total de acionamentos hoje: ' +
    '<span class="badge badge-blue">'+hist.totalHoje+'</span></p></div>' +

    '<div class="card"><h2>📋 Situação das Bases</h2>' +
    '<table style="'+estiloTabela+'">' +
    '<thead><tr>' +
    '<th style="'+estiloTh+'">Base</th>' +
    '<th style="'+estiloTh+'">Total</th>' +
    '<th style="'+estiloTh+'">Renovados</th>' +
    '<th style="'+estiloTh+'">Inadimplentes</th>' +
    '<th style="'+estiloTh+'">Evadidos</th>' +
    '</tr></thead><tbody>'+rowsBases+'</tbody></table></div>' +

    '<div class="card"><h2>👥 Acionamentos por Atendente</h2>' +
    '<table style="'+estiloTabela+'">' +
    '<thead><tr><th style="'+estiloTh+'">Atendente</th><th style="'+estiloTh+'">Acionamentos</th></tr></thead>' +
    '<tbody>'+rowsAtend+'</tbody></table></div>' +

    '<div class="card"><h2>📡 Acionamentos por Canal</h2>' +
    '<table style="'+estiloTabela+'">' +
    '<thead><tr><th style="'+estiloTh+'">Canal</th><th style="'+estiloTh+'">Qtd</th></tr></thead>' +
    '<tbody>'+rowsCanal+'</tbody></table></div>' +

    '<div style="color:#868e96;font-size:11px;text-align:center;margin-top:10px">' +
    'Gerado automaticamente pelo EngajaVR · Estácio Volta Redonda</div>' +
    '</body></html>';
}

// ── GERADOR TEXTO SIMPLES ────────────────────────────────────────
function gerarTextoRelatorio(hoje, bases, hist) {
  var linhas = [
    '📊 EngajaVR — Relatório Diário: ' + hoje,
    'Total de acionamentos hoje: ' + hist.totalHoje,
    '',
    '=== BASES ===',
  ];
  BASES.forEach(function(b){
    var d = bases[b]||{};
    linhas.push(b.toUpperCase() + ': ' + (d.total||0) + ' alunos | Renovados: ' + (d.renovados||0) +
      ' | Inadimplentes: ' + (d.inadimplentes||0) + ' | Evadidos: ' + (d.evadidos||0));
  });
  linhas.push('', '=== ACIONAMENTOS POR ATENDENTE ===');
  Object.keys(hist.porAtendente).sort().forEach(function(a){
    linhas.push(a + ': ' + hist.porAtendente[a] + ' acionamentos');
  });
  linhas.push('', '=== CANAIS ===');
  Object.keys(hist.porCanal).sort().forEach(function(c){
    linhas.push(c + ': ' + hist.porCanal[c]);
  });
  linhas.push('', '---', 'Gerado automaticamente pelo EngajaVR');
  return linhas.join('\n');
}
