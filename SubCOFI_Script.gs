/**
 * SubCOFI – CRESS-SP Seccional São José do Rio Preto
 * Menu de botões da planilha de gestão de demandas.
 * Cole este código em Extensões › Apps Script (veja o tutorial).
 * Versão 3.1: sem colunas automáticas repetidas; vários nomes em Quem fez/Quem encaminha; localidade 'Municípios em geral'; página pública via doPost com código de acesso (Listas G2); mapas se atualizam ao abrir a planilha (dentro das células); pauta abre numa janela; formulário web de relatório de atividade com documento .docx (v4.1); PDFs ficam na pasta da planilha se ela já for 'SubCOFI – Relatórios'.
 * Versão 3: planilha simplificada (só Demandas e Ações e eventos), mapas pintados por município e visitas.
 */

var ABA = {
  dem: 'Demandas', acoes: 'Ações e eventos',
  mun: 'Municípios', painel: 'Painel', pauta: 'Pauta'
};
var PASTA_RELATORIOS = 'SubCOFI – Relatórios';
var COR = '#7A1F2B';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('SubCOFI')
    .addItem('Gerar relatório em PDF…', 'abrirJanelaRelatorio')
    .addItem('Gerar pauta da próxima reunião', 'gerarPauta')
    .addItem('Atualizar painel e mapa', 'atualizarPainel')
    .addSeparator()
    .addItem('Link do formulário de relatório', 'abrirLinkFormulario')
    .addToUi();
}

/* ---------------- leitura das abas ---------------- */

function linhas_(nome) {
  var sh = planilha_().getSheetByName(nome);
  var v = sh.getDataRange().getValues();
  var cab = v.shift();
  return v.filter(function (l) { return l[0] !== '' || l[1] !== '' || l[2] !== ''; })
    .map(function (l) { var o = {}; cab.forEach(function (c, i) { o[c] = l[i]; }); return o; });
}

function ehData_(x) { return Object.prototype.toString.call(x) === '[object Date]' && !isNaN(x); }
function noPeriodo_(d, ini, fim) { return ehData_(d) && d >= ini && d <= fim; }
function fmt_(d) { return ehData_(d) ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy') : ''; }
function hoje_() { var h = new Date(); h.setHours(0, 0, 0, 0); return h; }

function contar_(lista, campo) {
  var m = {};
  lista.forEach(function (o) { var k = o[campo] || '(não informado)'; m[k] = (m[k] || 0) + 1; });
  return Object.keys(m).map(function (k) { return [k, m[k]]; }).sort(function (a, b) { return b[1] - a[1]; });
}

function municipios_() {
  return linhas_(ABA.mun).map(function (o) {
    return { loc: o['Localidade'], sede: o['Município-sede'], lat: Number(o['Latitude']), lon: Number(o['Longitude']) };
  });
}

/* ---------------- cálculo do período ---------------- */

function dadosPeriodo_(ini, fim) {
  var dem = linhas_(ABA.dem), ac = linhas_(ABA.acoes);
  var h = hoje_();
  var d = dem.filter(function (o) { return noPeriodo_(o['Data de entrada'], ini, fim); });
  var delib = dem.filter(function (o) { return noPeriodo_(o['Deliberada em'], ini, fim); });
  var r = delib;
  var a = ac.filter(function (o) { return noPeriodo_(o['Data'], ini, fim); });
  var vencidas = dem.filter(function (o) {
    return ehData_(o['Prazo']) && o['Prazo'] < h &&
      ['Concluída', 'Arquivada'].indexOf(o['Situação']) < 0;
  });
  var dias = delib.filter(function (o) { return ehData_(o['Data de entrada']); })
    .map(function (o) { return Math.round((o['Deliberada em'] - o['Data de entrada']) / 864e5); }).filter(function (x) { return x >= 0; });
  var datasReu = {};
  r.forEach(function (o) { datasReu[fmt_(o['Deliberada em'])] = 1; });

  var porLoc = {}, acLoc = {};
  d.forEach(function (o) { porLoc[o['Localidade']] = (porLoc[o['Localidade']] || 0) + 1; });
  a.forEach(function (o) { acLoc[o['Localidade']] = (acLoc[o['Localidade']] || 0) + 1; });
  var mun = municipios_().filter(function (m) { return m.lat; }); // sem 'Municípios em geral'
  var gerais = d.concat(a).filter(function (o) { return o['Localidade'] === 'Municípios em geral'; }).length;
  var semRegistro = mun.filter(function (m) { return !porLoc[m.loc] && !acLoc[m.loc]; }).map(function (m) { return m.loc; });

  return {
    ini: ini, fim: fim, dem: d, delib: delib, reu: r, acoes: a, vencidas: vencidas,
    mediaDias: dias.length ? Math.round(dias.reduce(function (s, x) { return s + x; }, 0) / dias.length) : null,
    nReunioes: Object.keys(datasReu).length,
    porLoc: porLoc, acLoc: acLoc, mun: mun, semRegistro: semRegistro, gerais: gerais
  };
}

/* ---------------- mapa (contornos dos municípios, IBGE) ---------------- */

// Tipos de ação que contam como visita (mapa verde)
function ehVisita_(tipo) { return /^visita/i.test(String(tipo || '')); }

// Soma por município-sede (distritos contam na sede)
function porSede_(lista, campoLoc) {
  var sede = {};
  municipios_().forEach(function (m) { sede[m.loc] = m.sede; });
  var out = {};
  lista.forEach(function (o) {
    var s = sede[o[campoLoc]] || o[campoLoc];
    if (s) out[s] = (out[s] || 0) + 1;
  });
  return out;
}

function rotuloPeriodo_(ini, fim) { return fmt_(ini) + ' a ' + fmt_(fim); }

function malhaDaSeccional_() {
  var sedes = {};
  municipios_().forEach(function (m) { sedes[m.sede] = 1; });
  return { W: MALHA.W, H: MALHA.H, rows: MALHA.rows.filter(function (r) { return sedes[r[0]]; }) };
}

// Dados que a janela usa para desenhar os 3 mapas
function dadosMapa(iniIso, fimIso) {
  var ini = ehData_(iniIso) ? iniIso : paraData_(iniIso), fim = ehData_(fimIso) ? fimIso : paraData_(fimIso, true);
  var D = dadosPeriodo_(ini, fim);
  var acoes = D.acoes.filter(function (o) { return !ehVisita_(o['Tipo de ação']); });
  var visitas = D.acoes.filter(function (o) { return ehVisita_(o['Tipo de ação']); });
  var p = rotuloPeriodo_(ini, fim);
  return {
    M: malhaDaSeccional_(), dem: porSede_(D.dem, 'Localidade'), ac: porSede_(acoes, 'Localidade'), vis: porSede_(visitas, 'Localidade'),
    t1: 'Demandas por município – ' + p, t2: 'Ações e eventos por município – ' + p,
    t3: 'Visitas (membros da SubCOFI e fiscal) – ' + p
  };
}

// Código que roda no navegador (desenha o mapa e devolve PNG em base64)
var CLIENTE_MAPA_ = '<script>' + desenharMapa.toString() +
  'function tresMapas(D){return [desenharMapa(D.M,D.dem,D.t1,"vinho"),desenharMapa(D.M,D.ac,D.t2,"azul"),desenharMapa(D.M,D.vis,D.t3,"verde")]}' +
  '</script>';

function desenharMapa(M, valores, titulo, paleta) {
  var ESC = { vinho: ['#F3D9DD', '#E0A3AD', '#C46272', '#9A3043', '#651326'],
              azul: ['#D7E7F6', '#9DC4E7', '#5C9BD6', '#2E6DA5', '#16497A'],
              verde: ['#DAF0DA', '#A7D9A7', '#63B363', '#2F8A3C', '#155C22'] }[paleta];
  var faixas = [[1, 1, '1'], [2, 2, '2'], [3, 4, '3 a 4'], [5, 9, '5 a 9'], [10, 1e9, '10 ou mais']];
  function cor(v) { for (var i = 0; i < faixas.length; i++) if (v >= faixas[i][0] && v <= faixas[i][1]) return ESC[i]; return '#E9E9E9'; }
  var mg = 20, topo = 64, rod = 74;
  var c = document.createElement('canvas');
  c.width = M.W + 2 * mg; c.height = M.H + topo + rod;
  var g = c.getContext('2d');
  g.fillStyle = '#FFFFFF'; g.fillRect(0, 0, c.width, c.height);
  g.fillStyle = '#222'; g.font = 'bold 28px Arial'; g.fillText(titulo, mg, 40);
  g.save(); g.translate(mg, topo);
  M.rows.forEach(function (r) {
    var v = valores[r[0]] || 0, p = new Path2D(r[4]);
    g.fillStyle = v ? cor(v) : '#E9E9E9'; g.fill(p);
    g.strokeStyle = '#FFFFFF'; g.lineWidth = 1.3; g.stroke(p);
  });
  g.font = 'bold 13px Arial'; g.textAlign = 'center'; g.lineJoin = 'round';
  M.rows.forEach(function (r) {
    var v = valores[r[0]] || 0; if (!v) return;
    var t = r[0] + ' (' + v + ')';
    g.strokeStyle = 'rgba(255,255,255,0.9)'; g.lineWidth = 4; g.strokeText(t, r[2], r[3] + 4);
    g.fillStyle = '#111'; g.fillText(t, r[2], r[3] + 4);
  });
  g.restore();
  var y = topo + M.H + 30, x = mg;
  g.font = '16px Arial'; g.textAlign = 'left';
  [['#E9E9E9', 'Sem registro']].concat(faixas.map(function (f, i) { return [ESC[i], f[2]]; })).forEach(function (it) {
    g.fillStyle = it[0]; g.fillRect(x, y - 15, 22, 18); g.strokeStyle = '#999'; g.lineWidth = 1; g.strokeRect(x, y - 15, 22, 18);
    g.fillStyle = '#333'; g.fillText(it[1], x + 30, y); x += 46 + g.measureText(it[1]).width;
  });
  g.fillStyle = '#666'; g.font = '13px Arial';
  g.fillText('Distritos contam no município-sede. Mostra o que foi registrado, não todas as necessidades do território.', mg, y + 30);
  return c.toDataURL('image/png').split(',')[1];
}

/* ---------------- relatório em PDF ---------------- */

function abrirJanelaRelatorio() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:14px">' +
    '<p>Escolha o período do relatório:</p>' +
    '<p>De <input type="date" id="i"> até <input type="date" id="f"></p>' +
    '<p style="font-size:12px">Atalhos: ' +
    '<a href="#" onclick="atalho(\'mes\')">mês atual</a> • ' +
    '<a href="#" onclick="atalho(\'tri\')">trimestre atual</a> • ' +
    '<a href="#" onclick="atalho(\'ano\')">ano atual</a></p>' +
    '<button id="b" onclick="gerar()" style="background:' + COR + ';color:#fff;border:0;padding:8px 14px;border-radius:4px">Gerar PDF</button>' +
    '<p id="msg"></p></div>' + CLIENTE_MAPA_ +
    '<script>' +
    'function iso(d){return d.toISOString().slice(0,10)}' +
    'function atalho(t){var h=new Date(),y=h.getFullYear(),m=h.getMonth(),a,b;' +
    'if(t=="mes"){a=new Date(y,m,1);b=new Date(y,m+1,0)}' +
    'else if(t=="tri"){var q=Math.floor(m/3)*3;a=new Date(y,q,1);b=new Date(y,q+3,0)}' +
    'else{a=new Date(y,0,1);b=new Date(y,11,31)}' +
    'document.getElementById("i").value=iso(new Date(a-a.getTimezoneOffset()*6e4));' +
    'document.getElementById("f").value=iso(new Date(b-b.getTimezoneOffset()*6e4));return false}' +
    'function msg(t){document.getElementById("msg").innerHTML=t}' +
    'function erro(e){msg("Erro: "+e.message);document.getElementById("b").disabled=false}' +
    'function gerar(){var i=document.getElementById("i").value,f=document.getElementById("f").value;' +
    'if(!i||!f){alert("Preencha as duas datas.");return}' +
    'document.getElementById("b").disabled=true;msg("Desenhando os mapas…");' +
    'google.script.run.withFailureHandler(erro).withSuccessHandler(function(D){' +
    'var im=tresMapas(D);msg("Montando o PDF… pode levar alguns segundos.");' +
    'google.script.run.withFailureHandler(erro).withSuccessHandler(function(u){' +
    'msg("Pronto! <a target=_blank href=\'"+u+"\'>Abrir o PDF</a> (salvo na pasta ' + PASTA_RELATORIOS + ')")})' +
    '.gerarRelatorio(i,f,im)}).dadosMapa(i,f)}' +
    '</script>').setWidth(440).setHeight(240);
  SpreadsheetApp.getUi().showModalDialog(html, 'Relatório SubCOFI');
}

function paraData_(iso, fimDoDia) {
  var p = iso.split('-');
  return fimDoDia ? new Date(p[0], p[1] - 1, p[2], 23, 59, 59) : new Date(p[0], p[1] - 1, p[2]);
}

function tabela_(titulo, pares, total) {
  if (!pares.length) return '<h3>' + titulo + '</h3><p class="v">Nenhum registro no período.</p>';
  var max = Math.max.apply(null, pares.map(function (p) { return p[1]; }));
  var h = '<h3>' + titulo + '</h3><table class="t">';
  pares.forEach(function (p) {
    var w = Math.max(4, Math.round(260 * p[1] / max));
    h += '<tr><td class="n">' + p[0] + '</td><td class="q">' + p[1] + '</td>' +
      '<td><table cellspacing="0" cellpadding="0"><tr><td style="background:' + COR + ';width:' + w + 'px;height:12px"></td></tr></table></td></tr>';
  });
  return h + '</table>';
}

function porPeriodo_(lista, campoData) {
  var m = {};
  lista.forEach(function (o) {
    var d = o[campoData]; if (!ehData_(d)) return;
    var k = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
    m[k] = (m[k] || 0) + 1;
  });
  return Object.keys(m).sort().map(function (k) { return [k, m[k]]; });
}

function porTrimestre_(lista, campoData) {
  var m = {};
  lista.forEach(function (o) {
    var d = o[campoData]; if (!ehData_(d)) return;
    var k = d.getFullYear() + '-T' + (Math.floor(d.getMonth() / 3) + 1);
    m[k] = (m[k] || 0) + 1;
  });
  return Object.keys(m).sort().map(function (k) { return [k, m[k]]; });
}

function gerarRelatorio(iniIso, fimIso, imgs) {
  var ini = paraData_(iniIso), fim = paraData_(fimIso, true);
  if (fim < ini) throw new Error('A data final é anterior à inicial.');
  var D = dadosPeriodo_(ini, fim);
  var sit = contar_(D.dem, 'Situação');
  var apoio = D.dem.filter(function (o) { return o['Situação'] === 'Precisa de apoio'; }).length;

  var css = '<style>body{font-family:Arial;font-size:10.5pt;color:#222}h1{color:' + COR + ';font-size:17pt;margin:0}' +
    'h2{color:' + COR + ';font-size:13pt;border-bottom:2px solid ' + COR + ';margin-top:18px}h3{font-size:11pt;margin:10px 0 4px}' +
    '.t td{padding:2px 6px;font-size:10pt}.n{width:260px}.q{text-align:right;font-weight:bold;width:30px}' +
    '.k td{border:1px solid #ccc;padding:6px 10px;text-align:center}.k b{font-size:15pt;color:' + COR + '}' +
    '.v{color:#777;font-style:italic}.p{font-size:8.5pt;color:#555}</style>';

  var kp = [['Demandas recebidas', D.dem.length], ['Deliberadas', D.delib.length], ['Reuniões com deliberação', D.nReunioes],
    ['Ações e eventos', D.acoes.filter(function (o) { return !ehVisita_(o['Tipo de ação']); }).length],
    ['Visitas', D.acoes.filter(function (o) { return ehVisita_(o['Tipo de ação']); }).length], ['Com pedido de apoio', apoio], ['Para todos os municípios', D.gerais], ['Encaminhamentos vencidos (hoje)', D.vencidas.length]];
  var k = '<table class="k" cellspacing="0"><tr>' + kp.map(function (x) { return '<td><b>' + x[1] + '</b><br>' + x[0] + '</td>'; }).join('') + '</tr></table>';

  var html = css +
    '<h1>Relatório de atividades – SubCOFI</h1>' +
    '<p>CRESS-SP • Seccional São José do Rio Preto<br><b>Período: ' + fmt_(ini) + ' a ' + fmt_(fim) + '</b> • Gerado em ' + fmt_(new Date()) + '</p>' +
    '<h2>Resumo</h2>' + k +
    '<p>Tempo médio entre a entrada e a deliberação (demandas deliberadas no período): <b>' +
    (D.mediaDias === null ? '–' : D.mediaDias + ' dias') + '</b></p>' +
    '<h2>Demandas</h2>' +
    tabela_('Por mês de entrada', porPeriodo_(D.dem, 'Data de entrada')) +
    tabela_('Por trimestre de entrada', porTrimestre_(D.dem, 'Data de entrada')) +
    tabela_('Por tipo', contar_(D.dem, 'Tipo de demanda')) +
    tabela_('Situação atual (das recebidas no período)', sit) +
    tabela_('Análises relatadas, por integrante', contar_(D.dem, 'Relator/a')) +
    tabela_('Por localidade', contar_(D.dem, 'Localidade')) +
    '<h2>Ações, eventos e visitas</h2>' +
    tabela_('Por tipo', contar_(D.acoes, 'Tipo de ação')) +
    tabela_('Por mês', porPeriodo_(D.acoes, 'Data')) +
    tabela_('Por localidade', contar_(D.acoes, 'Localidade')) +
    '<h2>Território</h2>' +
    (imgs || []).map(function (b64) { return '<p><img src="data:image/png;base64,' + b64 + '" width="600"></p>'; }).join('') +
    '<h3>Localidades sem nenhum registro no período (' + D.semRegistro.length + ' de ' + D.mun.length + ')</h3>' +
    '<p style="font-size:9pt">' + (D.semRegistro.join(' • ') || 'Nenhuma.') + '</p>' +
    '<p class="p">Relatório de uso interno da SubCOFI, gerado a partir da planilha de gestão. Não substitui o registro oficial da fiscalização.</p>';

  var nome = 'Relatório SubCOFI ' + Utilities.formatDate(ini, Session.getScriptTimeZone(), 'yyyy-MM-dd') +
    ' a ' + Utilities.formatDate(fim, Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.pdf';
  var pdf = Utilities.newBlob(html, 'text/html', 'r.html').getAs('application/pdf').setName(nome);
  return pasta_().createFile(pdf).getUrl();
}

function pasta_() {
  var arq = DriveApp.getFileById(planilha_().getId());
  var pais = arq.getParents();
  var pai = pais.hasNext() ? pais.next() : DriveApp.getRootFolder();
  if (pai.getName() === PASTA_RELATORIOS) return pai; // planilha já está dentro da pasta de relatórios
  var it = pai.getFoldersByName(PASTA_RELATORIOS);
  return it.hasNext() ? it.next() : pai.createFolder(PASTA_RELATORIOS);
}

/* ---------------- pauta ---------------- */

function gerarPauta() {
  var ss = SpreadsheetApp.getActive();
  var velha = ss.getSheetByName('Pauta');
  if (velha) ss.deleteSheet(velha); // versões antigas criavam uma aba; agora a pauta abre numa janela
  var dem = linhas_(ABA.dem), h = hoje_();
  var abertas = ['Recebida', 'Em análise', 'Para a pauta', 'Precisa de apoio', 'Deliberada', 'Em encaminhamento'];
  var dias = function (d) { return ehData_(d) ? Math.round((h - d) / 864e5) : ''; };

  var b1 = dem.filter(function (o) { return o['Situação'] === 'Para a pauta'; });
  var b2 = dem.filter(function (o) {
    return ['Recebida', 'Em análise'].indexOf(o['Situação']) >= 0 && dias(o['Data de entrada']) >= 30;
  }).sort(function (a, b) { return a['Data de entrada'] - b['Data de entrada']; });
  var b3 = dem.filter(function (o) {
    var p = o['Prazo'];
    return ehData_(p) && abertas.indexOf(o['Situação']) >= 0 && (p - h) / 864e5 <= 15;
  }).sort(function (a, b) { return a['Prazo'] - b['Prazo']; });
  var b4 = dem.filter(function (o) { return o['Situação'] === 'Precisa de apoio'; });

  function esc(x) { return String(x === undefined || x === null ? '' : x).replace(/&/g, '&amp;').replace(/</g, '&lt;'); }
  function bloco(t, lista, extra, val) {
    var h2 = '<h3>' + t + ' (' + lista.length + ')</h3>';
    if (!lista.length) return h2 + '<p class="v">Nada neste item.</p>';
    var tb = '<table><tr><th>Código</th><th>Localidade</th><th>Tipo</th><th>Relator/a</th><th>Situação</th><th>' + extra + '</th></tr>';
    lista.forEach(function (o) {
      tb += '<tr><td>' + esc(o['Código']) + '</td><td>' + esc(o['Localidade']) + '</td><td>' + esc(o['Tipo de demanda']) +
        '</td><td>' + esc(o['Relator/a']) + '</td><td>' + esc(o['Situação']) + '</td><td>' + esc(val(o)) + '</td></tr>';
    });
    return h2 + tb + '</table>';
  }
  var corpo =
    bloco('1. Análises prontas para deliberar', b1, 'Entrada', function (o) { return fmt_(o['Data de entrada']); }) +
    bloco('2. Demandas paradas há 30 dias ou mais', b2, 'Dias parada', function (o) { return dias(o['Data de entrada']); }) +
    bloco('3. Prazos vencidos ou vencendo em 15 dias', b3, 'Prazo', function (o) { return fmt_(o['Prazo']); }) +
    bloco('4. Pedidos de apoio', b4, 'Entrada', function (o) { return fmt_(o['Data de entrada']); });
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial;font-size:12px;color:#222}h2{color:' + COR + ';margin:0 0 4px}' +
    'h3{color:' + COR + ';font-size:13px;margin:14px 0 4px}table{border-collapse:collapse;width:100%}' +
    'th,td{border:1px solid #ddd;padding:3px 5px;text-align:left}th{background:#F6EDEF}.v{color:#777;font-style:italic;margin:0}' +
    'button{background:' + COR + ';color:#fff;border:0;padding:7px 12px;border-radius:4px;cursor:pointer}' +
    '@media print{button{display:none}}</style>' +
    '<h2>Pauta da próxima reunião – SubCOFI</h2><div>Gerada em ' + fmt_(new Date()) + '</div>' + corpo +
    '<p style="margin-top:14px"><button onclick="window.print()">Imprimir / salvar em PDF</button></p>')
    .setWidth(820).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Pauta SubCOFI');
}

/* ---------------- painel ---------------- */

// Chamado automaticamente toda vez que a planilha é aberta (gatilho instalável).
function atualizarAoAbrir() { atualizarPainel(); }

// Cria o gatilho de abertura uma única vez (também funciona em cópias da planilha).
function garantirGatilho_() {
  var ss = SpreadsheetApp.getActive();
  var existe = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'atualizarAoAbrir'; });
  if (!existe) ScriptApp.newTrigger('atualizarAoAbrir').forSpreadsheet(ss).onOpen().create();
}

function atualizarPainel() {
  guardarIdPlanilha_();
  try { garantirGatilho_(); } catch (e) { /* sem permissão para criar gatilho: segue só a atualização */ }
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA.painel);
  var ini = sh.getRange('C3').getValue(), fim = sh.getRange('C4').getValue();
  if (!ehData_(ini) || !ehData_(fim)) { SpreadsheetApp.getUi().alert('Preencha as datas em C3 e C4 do Painel.'); return; }
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;font-size:14px" id="m">Desenhando os mapas…</div>' + CLIENTE_MAPA_ +
    '<script>' +
    'function erro(e){document.getElementById("m").innerText="Erro: "+e.message}' +
    'google.script.run.withFailureHandler(erro).withSuccessHandler(function(D){' +
    'google.script.run.withFailureHandler(erro).withSuccessHandler(function(){google.script.host.close()})' +
    '.salvarMapasPainel(tresMapas(D))}).dadosMapaPainel();' +
    '</script>').setWidth(260).setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(html, 'Atualizando o Painel');
}

function dadosMapaPainel() {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA.painel);
  var ini = sh.getRange('C3').getValue(), fim = sh.getRange('C4').getValue();
  fim = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate(), 23, 59, 59);
  return dadosMapa(ini, fim);
}

// Os mapas ficam DENTRO de células mescladas (E7:L26, E28:L47, E49:L68): assim a troca aparece na hora.
function salvarMapasPainel(imgs) {
  var sh = SpreadsheetApp.getActive().getSheetByName(ABA.painel);
  sh.getImages().forEach(function (im) { im.remove(); }); // remove mapas antigos "soltos"
  imgs.forEach(function (b64, i) {
    var topo = 7 + i * 21, faixa = sh.getRange(topo, 5, 20, 8);
    if (!faixa.isPartOfMerge()) { faixa.merge(); }
    faixa.setHorizontalAlignment('center').setVerticalAlignment('middle');
    var img = SpreadsheetApp.newCellImage().setSourceUrl('data:image/png;base64,' + b64)
      .setAltTextTitle('Mapa ' + (i + 1)).build();
    sh.getRange(topo, 5).setValue(img);
  });
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast('Mapas do Painel atualizados.', 'SubCOFI');
}


/* ================= FORMULÁRIO DE RELATÓRIO DE ATIVIDADE (página web) ================= */
// A página é publicada em Implantar › Nova implantação › App da Web.
// Ela registra a atividade na aba "Ações e eventos", salva as fotos, gera o relatório editável (Google Docs/.docx)
// no modelo "Relatório de participação em evento" e entrega o e-mail pronto para copiar.

var PASTA_FOTOS = 'SubCOFI – Fotos de ações';

function planilha_() {
  var ss = SpreadsheetApp.getActive();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('ID_PLANILHA');
  if (!id) throw new Error('Abra a planilha e use uma vez o menu SubCOFI antes de usar o formulário.');
  return SpreadsheetApp.openById(id);
}

function guardarIdPlanilha_() {
  try { PropertiesService.getScriptProperties().setProperty('ID_PLANILHA', SpreadsheetApp.getActive().getId()); } catch (e) {}
}

function abrirLinkFormulario() {
  guardarIdPlanilha_();
  var url = '';
  try { url = String(planilha_().getSheetByName('Listas').getRange('F2').getValue()).trim(); } catch (e) {}
  if (!url) { try { url = ScriptApp.getService().getUrl(); } catch (e) {} }
  var corpo = url
    ? '<p>Link do formulário (envie à equipe):</p><p><input style="width:100%" value="' + url + '" onclick="this.select()" readonly></p>' +
      '<p><a href="' + url + '" target="_blank">Abrir o formulário</a></p>'
    : '<p>O formulário ainda não foi publicado. Veja o guia: Extensões › Apps Script › Implantar › Nova implantação › App da Web.</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<div style="font-family:Arial;font-size:14px">' + corpo + '</div>')
    .setWidth(480).setHeight(180), 'Formulário de relatório');
}

function doGet() {
  return HtmlService.createHtmlOutput(paginaFormulario_())
    .setTitle('Relatório de atividade – SubCOFI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function valoresColuna_(aba, col) {
  var sh = planilha_().getSheetByName(aba);
  return sh.getRange(2, col, Math.max(sh.getLastRow() - 1, 1), 1).getValues()
    .map(function (l) { return String(l[0]).trim(); }).filter(String);
}

function emailDestino_() {
  try { return String(planilha_().getSheetByName('Listas').getRange('E2').getValue()).trim(); } catch (e) { return ''; }
}

// Código de acesso da Seccional (aba Listas, célula G2). Fica só na planilha, nunca na página pública.
function codigoOk_(c) {
  var cod = String(planilha_().getSheetByName('Listas').getRange('G2').getValue()).trim();
  return cod && String(c || '').trim().toUpperCase() === cod.toUpperCase();
}

// Recebe os pedidos da página pública (GitHub Pages). Sem o código certo, não devolve nem grava nada.
function doPost(e) {
  var r;
  try {
    var q = JSON.parse(e.postData.contents);
    if (!codigoOk_(q.codigo)) throw new Error('Código de acesso inválido.');
    if (q.acao === 'dados') r = { ok: true, dados: dadosFormulario() };
    else if (q.acao === 'registrar') r = { ok: true, resultado: registrarRelatorio(q.dados) };
    else throw new Error('Ação desconhecida.');
  } catch (err) { r = { ok: false, erro: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function dadosFormulario() {
  return {
    municipios: valoresColuna_(ABA.mun, 1),
    tipos: valoresColuna_('Listas', 3),
    equipe: valoresColuna_('Equipe', 1),
    destino: emailDestino_()
  };
}

function pastaFilha_(pai, nome) {
  var it = pai.getFoldersByName(nome);
  return it.hasNext() ? it.next() : pai.createFolder(nome);
}

function esc_(x) { return String(x === undefined || x === null ? '' : x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br>'); }

function dataExtenso_(d) {
  var m = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
  return d.getDate() + ' de ' + m[d.getMonth()] + ' de ' + d.getFullYear();
}

function registrarRelatorio(d) {
  ['atividade', 'data', 'municipio', 'tipo', 'relato', 'nome', 'emailPessoa'].forEach(function (k) {
    if (!d[k] || !String(d[k]).trim()) throw new Error('Preencha o campo obrigatório: ' + k);
  });
  var p = d.data.split('-'); var data = new Date(p[0], p[1] - 1, p[2], 12); // meio-dia evita troca de dia pelo fuso
  var dataTxt = fmt_(data) + (d.horario ? ' – ' + d.horario : '');
  var quem = (d.quem || []).join(', ');

  // 1. Fotos no Drive
  var pastaFotos = pastaFilha_(pastaFilha_(pasta_(), PASTA_FOTOS),
    Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd') + ' – ' + d.atividade.slice(0, 60));
  var blobs = (d.fotos || []).map(function (f, i) {
    var blob = Utilities.newBlob(Utilities.base64Decode(f.b64), 'image/jpeg', 'foto_' + (i + 1) + '.jpg');
    pastaFotos.createFile(blob);
    return blob;
  });

  // 2. Documento editável (Google Docs, baixável em .docx) no modelo "Relatório de participação em evento"
  var nomeDoc = 'Relatório – ' + Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd') + ' – ' + d.atividade.slice(0, 60);
  var doc = DocumentApp.create(nomeDoc), body = doc.getBody();
  body.setFontFamily('Arial').setFontSize(11);
  var titulo = body.getParagraphs()[0];
  titulo.setText('RELATÓRIO DE PARTICIPAÇÃO EM EVENTO');
  titulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titulo.setSpacingAfter(12);
  titulo.editAsText().setBold(true).setFontSize(14);
  function linha(rot, val) {
    if (!val) return;
    var pl = body.appendParagraph('');
    pl.appendText(rot + ': ').setBold(true);
    pl.appendText(String(val)).setBold(false);
  }
  linha('Atividade', d.atividade); linha('Data/horário', dataTxt);
  linha('Local', [d.formato, d.local].filter(String).join(' – ')); linha('Endereço', d.endereco);
  linha('Município', d.municipio); linha('Tipo de ação', d.tipo); linha('Temática', d.tematica);
  linha('Público', d.publico); linha('Convidados integrantes da mesa', d.mesa);
  linha('Total de participantes', d.total); linha('Participação da SubCOFI', quem);
  function par(texto, negrito, alinh, antes) {
    var p = body.appendParagraph(texto || '');
    if (texto) p.editAsText().setBold(!!negrito);
    if (alinh) p.setAlignment(alinh);
    if (antes) p.setSpacingBefore(antes);
    return p;
  }
  var A = DocumentApp.HorizontalAlignment;
  par(d.relato, false, A.JUSTIFY, 12);
  if (blobs.length) {
    par('Registro fotográfico', true, null, 12);
    blobs.forEach(function (b) {
      var img = par('', false).appendInlineImage(b);
      var w = 300; img.setHeight(Math.round(img.getHeight() * w / img.getWidth())).setWidth(w);
    });
  }
  par((d.cidade || 'São José do Rio Preto') + ', ' + dataExtenso_(new Date()) + '.', false, A.CENTER, 30);
  par(d.nome, false, A.CENTER, 18);
  if (d.cargo) par(d.cargo, false, A.CENTER);
  doc.saveAndClose();
  var arq = DriveApp.getFileById(doc.getId());
  arq.moveTo(pasta_());
  arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // para o CRESS abrir pelo link
  var docxUrl = 'https://docs.google.com/document/d/' + doc.getId() + '/export?format=docx';

  // 3. Linha na aba Ações e eventos (com links do relatório e das fotos)
  var sh = planilha_().getSheetByName(ABA.acoes);
  if (!sh.getRange('F1').getValue()) sh.getRange('F1:G1').setValues([['Relatório', 'Fotos']]);
  var colA = sh.getRange(2, 1, sh.getMaxRows() - 1, 1).getValues();
  var r = 2; while (r - 2 < colA.length && colA[r - 2][0] !== '') r++;
  sh.getRange(r, 1, 1, 5).setValues([[data, d.municipio, d.tipo, quem, d.atividade]]);
  sh.getRange(r, 6, 1, 2).setFormulas([['=HYPERLINK("' + arq.getUrl() + '";"Abrir relatório")', '=HYPERLINK("' + pastaFotos.getUrl() + '";"Abrir fotos")']]);

  // 4. E-mail pronto para copiar
  var assunto = 'Relatório de participação em evento – ' + d.atividade + ' – ' + d.municipio + ' – ' + fmt_(data);
  var resumo = [['Data/horário', dataTxt], ['Local', [d.formato, d.local].filter(String).join(' – ')], ['Temática', d.tematica],
    ['Público', d.publico], ['Total de participantes', d.total], ['Participação da SubCOFI', quem]].filter(function (x) { return x[1]; });
  var texto = 'Prezadas e prezados,\n\nEncaminho o relatório de participação em evento referente à atividade "' + d.atividade +
    '", realizada em ' + fmt_(data) + ' (' + d.municipio + ').\n\n' +
    resumo.map(function (x) { return x[0] + ': ' + x[1]; }).join('\n') +
    '\n\nRelatório completo: ' + arq.getUrl() + '\n\nAtenciosamente,\n' + d.nome + (d.cargo ? '\n' + d.cargo : '') + '\n' + d.emailPessoa;
  var htmlEmail = '<p>Prezadas e prezados,</p><p>Encaminho o relatório de participação em evento referente à atividade <b>' + esc_(d.atividade) +
    '</b>, realizada em ' + fmt_(data) + ' (' + esc_(d.municipio) + ').</p><p>' +
    resumo.map(function (x) { return '<b>' + x[0] + ':</b> ' + esc_(x[1]); }).join('<br>') +
    '</p><p><a href="' + arq.getUrl() + '">Relatório completo</a></p><p>Atenciosamente,<br>' + esc_(d.nome) +
    (d.cargo ? '<br>' + esc_(d.cargo) : '') + '<br>' + esc_(d.emailPessoa) + '</p>';

  return { para: d.destino || '', assunto: assunto, texto: texto, html: htmlEmail,
    docUrl: arq.getUrl(), docxUrl: docxUrl };
}

function paginaFormulario_() {
  return HTML_FORMULARIO_;
}

var HTML_FORMULARIO_ = "<!DOCTYPE html>\n<html><head><meta charset=\"utf-8\"><base target=\"_top\">\n<style>\n:root{--v:#7A1F2B;--c:#F6EDEF;--t:#222;--m:#666}\n*{box-sizing:border-box}\nbody{font-family:Arial,Helvetica,sans-serif;color:var(--t);background:#fafafa;margin:0;padding:16px}\n.w{max-width:720px;margin:0 auto;background:#fff;border:1px solid #e5e5e5;border-radius:10px;padding:20px}\nh1{color:var(--v);font-size:21px;margin:0 0 4px}\n.s{color:var(--m);font-size:13px;margin:0 0 16px}\nlabel{display:block;font-weight:bold;font-size:14px;margin:14px 0 4px}\nlabel small{font-weight:normal;color:var(--m)}\ninput,select,textarea{width:100%;padding:9px;border:1px solid #ccc;border-radius:6px;font:inherit;font-size:15px}\ntextarea{min-height:120px}\n.g{display:grid;grid-template-columns:1fr 1fr;gap:0 12px}\n@media(max-width:560px){.g{grid-template-columns:1fr}}\n.chk label{display:inline-flex;align-items:center;gap:6px;font-weight:normal;margin:4px 14px 4px 0}\n.chk input{width:auto}\nbutton{background:var(--v);color:#fff;border:0;border-radius:6px;padding:11px 16px;font-size:15px;cursor:pointer}\nbutton.sec{background:#fff;color:var(--v);border:1px solid var(--v)}\n.av{background:var(--c);border-radius:6px;padding:10px;font-size:13px;margin-top:12px}\n#fotosPrev img{width:90px;height:90px;object-fit:cover;border-radius:6px;margin:4px}\n.caixa{border:1px solid #ddd;border-radius:6px;padding:10px;margin:6px 0;white-space:pre-wrap;font-size:14px;background:#fcfcfc}\n.lin{display:flex;justify-content:space-between;align-items:center;margin-top:12px}\n#msg{margin-top:12px;font-weight:bold}\n.oculto{display:none}\n</style></head><body><div class=\"w\">\n<div id=\"form\">\n<h1>Relatório de atividade – SubCOFI</h1>\n<p class=\"s\">CRESS-SP • Seccional São José do Rio Preto. Preencha, registre e copie o e-mail pronto para enviar ao CRESS-SP.</p>\n\n<label>Atividade *</label><input id=\"atividade\" placeholder=\"Ex.: 10ª Reunião Ordinária do CMAS\">\n<div class=\"g\">\n<div><label>Data *</label><input id=\"data\" type=\"date\"></div>\n<div><label>Horário</label><input id=\"horario\" placeholder=\"Ex.: 14h às 17h\"></div>\n</div>\n<div class=\"g\">\n<div><label>Formato</label><select id=\"formato\"><option>Presencial</option><option>Virtual</option><option>Híbrido</option></select></div>\n<div><label>Local</label><input id=\"local\" placeholder=\"Ex.: Câmara Municipal / Google Meet\"></div>\n</div>\n<label>Endereço</label><input id=\"endereco\">\n<div class=\"g\">\n<div><label>Município *</label><select id=\"municipio\"></select></div>\n<div><label>Tipo de ação *</label><select id=\"tipo\"></select></div>\n</div>\n<label>Temática</label><input id=\"tematica\">\n<label>Público</label><input id=\"publico\" placeholder=\"Ex.: Representantes do CRESS, CRP e CMAS\">\n<label>Convidados integrantes da mesa</label><textarea id=\"mesa\" style=\"min-height:60px\"></textarea>\n<label>Total de participantes</label><input id=\"total\" type=\"number\" min=\"0\">\n<label>Participação da SubCOFI <small>(marque todos)</small></label><div class=\"chk\" id=\"quem\"></div>\n<label>Relato da atividade * <small>(pautas, discussões, encaminhamentos)</small></label><textarea id=\"relato\"></textarea>\n<label>Fotos <small>(até 10; são reduzidas antes de enviar)</small></label><input id=\"fotos\" type=\"file\" accept=\"image/*\" multiple>\n<div id=\"fotosPrev\"></div>\n<div class=\"av\"><label style=\"margin:0;font-weight:normal\"><input type=\"checkbox\" id=\"consent\" style=\"width:auto\"> As pessoas que aparecem nas fotos autorizaram o uso das imagens. Não incluí dados sigilosos no relato.</label></div>\n\n<div class=\"g\">\n<div><label>Seu nome *</label><input id=\"nome\"></div>\n<div><label>Cargo/função</label><input id=\"cargo\" placeholder=\"Ex.: Coordenadora da Seccional\"></div>\n</div>\n<div class=\"g\">\n<div><label>Seu e-mail *</label><input id=\"emailPessoa\" type=\"email\"></div>\n<div><label>Cidade da assinatura</label><input id=\"cidade\" value=\"São José do Rio Preto\"></div>\n</div>\n<label>E-mail do CRESS-SP (destino)</label><input id=\"destino\" type=\"email\">\n\n<div class=\"lin\"><span class=\"s\" style=\"margin:0\">* obrigatório</span><button id=\"enviar\" onclick=\"registrar()\">Registrar e gerar relatório</button></div>\n<div id=\"msg\"></div>\n</div>\n\n<div id=\"resultado\" class=\"oculto\">\n<h1>Pronto! Atividade registrada.</h1>\n<p class=\"s\">Abra o seu e-mail, clique em \"Novo e-mail\" e cole cada parte. Se precisar anexar, use \"Baixar Word (.docx)\".</p>\n<label>Para</label><div class=\"caixa\" id=\"rPara\"></div><button class=\"sec\" onclick=\"copiarTexto('rPara')\">Copiar destinatário</button>\n<label>Assunto</label><div class=\"caixa\" id=\"rAssunto\"></div><button class=\"sec\" onclick=\"copiarTexto('rAssunto')\">Copiar assunto</button>\n<label>Texto</label><div class=\"caixa\" id=\"rTexto\"></div><button class=\"sec\" onclick=\"copiarHtml()\">Copiar texto</button>\n<div class=\"lin\"><button onclick=\"baixarDocx()\">Baixar Word (.docx)</button><button class=\"sec\" onclick=\"location.reload()\">Registrar outra atividade</button></div>\n<div id=\"msg2\" class=\"s\" style=\"margin-top:10px\"></div>\n</div>\n</div>\n<script>\nvar R = null, FOTOS = [];\nfunction el(i){return document.getElementById(i)}\nfunction opcoes(sel, lista){ el(sel).innerHTML = '<option value=\"\"></option>' + lista.map(function(x){return '<option>'+x.replace(/</g,'&lt;')+'</option>'}).join(''); }\nfunction carregar(D){\n  opcoes('municipio', D.municipios); opcoes('tipo', D.tipos);\n  el('quem').innerHTML = D.equipe.map(function(n,i){return '<label><input type=\"checkbox\" value=\"'+n.replace(/\"/g,'&quot;')+'\"> '+n+'</label>'}).join('');\n  el('destino').value = D.destino || '';\n}\nfunction reduzir(arq){\n  return new Promise(function(ok, falha){\n    var fr = new FileReader();\n    fr.onload = function(){ var img = new Image(); img.onload = function(){\n      var m = 1400, e = Math.min(1, m / Math.max(img.width, img.height));\n      var c = document.createElement('canvas'); c.width = Math.round(img.width*e); c.height = Math.round(img.height*e);\n      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);\n      ok(c.toDataURL('image/jpeg', 0.75).split(',')[1]);\n    }; img.onerror = falha; img.src = fr.result; };\n    fr.onerror = falha; fr.readAsDataURL(arq);\n  });\n}\nfunction mostrarFotos(){\n  var arqs = Array.prototype.slice.call(el('fotos').files, 0, 10);\n  el('fotosPrev').innerHTML = 'Preparando fotos…';\n  Promise.all(arqs.map(reduzir)).then(function(b){\n    FOTOS = b; el('fotosPrev').innerHTML = b.map(function(x){return '<img src=\"data:image/jpeg;base64,'+x+'\">'}).join('');\n  }).catch(function(){ el('fotosPrev').innerHTML = 'Não foi possível ler uma das fotos.'; });\n}\nfunction registrar(){\n  var ids = ['atividade','data','horario','formato','local','endereco','municipio','tipo','tematica','publico','mesa','total','relato','nome','cargo','emailPessoa','cidade','destino'];\n  var d = {}; ids.forEach(function(i){ d[i] = el(i).value.trim(); });\n  d.quem = Array.prototype.slice.call(el('quem').querySelectorAll('input:checked')).map(function(x){return x.value});\n  var falta = [['atividade','Atividade'],['data','Data'],['municipio','Município'],['tipo','Tipo de ação'],['relato','Relato'],['nome','Seu nome'],['emailPessoa','Seu e-mail']]\n    .filter(function(x){return !d[x[0]]}).map(function(x){return x[1]});\n  if (falta.length){ el('msg').textContent = 'Preencha: ' + falta.join(', ') + '.'; return; }\n  if (FOTOS.length && !el('consent').checked){ el('msg').textContent = 'Confirme a autorização de uso das imagens.'; return; }\n  d.fotos = FOTOS.map(function(b){return {b64:b}});\n  el('enviar').disabled = true; el('msg').textContent = 'Registrando e gerando o relatório… pode levar até 1 minuto.';\n  google.script.run.withSuccessHandler(function(r){\n    R = r; el('form').className = 'oculto'; el('resultado').className = '';\n    el('rPara').textContent = r.para || '(digite o e-mail do CRESS-SP)';\n    el('rAssunto').textContent = r.assunto; el('rTexto').innerHTML = r.html; window.scrollTo(0,0);\n  }).withFailureHandler(function(e){ el('enviar').disabled = false; el('msg').textContent = 'Erro: ' + e.message; }).registrarRelatorio(d);\n}\nfunction copiarNo(no){\n  var s = window.getSelection(), r = document.createRange(); r.selectNodeContents(no); s.removeAllRanges(); s.addRange(r);\n  var ok = false; try { ok = document.execCommand('copy'); } catch(e){}\n  s.removeAllRanges(); el('msg2').textContent = ok ? 'Copiado! Agora cole no seu e-mail (Ctrl+V).' : 'Selecione o texto e use Ctrl+C.';\n}\nfunction copiarTexto(id){ copiarNo(el(id)); }\nfunction copiarHtml(){ copiarNo(el('rTexto')); }\nfunction baixarDocx(){ window.open(R.docxUrl, '_blank'); }\nel('fotos').addEventListener('change', mostrarFotos);\nif (window.google && google.script) google.script.run.withSuccessHandler(carregar).dadosFormulario();\n</script></body></html>\n";

/* Contornos simplificados dos municípios-sede da Seccional (IBGE, malha municipal).
   Formato: [município-sede, código IBGE, x do rótulo, y do rótulo, contorno SVG]. Não editar à mão. */
var MALHA = {"W":1000.0,"H":723.6,"rows":[["Adolfo","3500204",525.1,574.9,"M549.4,569.9L546,566.7L543.9,571.6L540.6,572L535.1,566.2L538.9,551.3L536.1,545.9L530.2,543.8L526.6,547.1L514.5,541.6L509.5,544.3L510.6,551.5L505.7,556.5L506.5,559.4L494.5,564L499.8,576.4L495.2,578.9L495,581.6L487.3,579.7L487.7,587.4L493.9,598.1L501,594.1L504.9,596.3L504.9,604.9L508,605.6L515.4,599.9L544,604.7L543,597L546.8,590.5L545.3,582.6L548.4,582L551,576.9L549.4,569.9Z"],["Altair","3500907",717.1,287.8,"M757.5,306.4L752.2,305.4L748.7,301.9L746.6,294.2L750,293.6L752.1,288.6L750.1,279.8L746.8,279L740.1,270.8L741.2,268.6L734.8,258.7L727.3,256.7L723.8,258.6L723.1,251.9L725.3,248.3L722.2,244.3L718.4,243.1L713.4,248.1L709.8,259.1L706.5,261.5L702.4,262.4L694.8,259.8L691,265.4L679.6,262.4L675.5,264.5L678.1,266.8L678.5,276.2L686.1,280.3L681.5,287L682.5,288.9L691.2,292.8L697.4,303.5L705.4,306.9L711.5,315.1L722.4,314.1L724.9,322.2L732.8,319.1L736.1,324.3L748.7,332.5L751.6,327.5L754.4,326.9L757.5,306.4Z"],["Alvares Florence","3501202",420.3,195.0,"M393.9,210.9L396.7,211.1L401.9,217L414.1,219.9L418.8,233.7L435,235.4L436.2,240.4L452.6,242.3L466.9,231.1L468.2,227.3L456.6,221L449.3,219L446.3,220.6L446,213L449,206.7L446.3,205L444.6,192.4L451.4,178.1L452.6,165.1L463.6,158.5L458.5,155.7L450.9,159.4L449,157.5L439.2,159.8L432.3,167L427.7,167.5L420.8,160.4L420,155.7L417.1,154.9L417.3,151L409.2,150L399.4,155.2L386.7,153.6L387.7,159.1L391.8,162.3L390.9,175.5L392.7,185.9L395.9,190.7L396.7,200.7L392.6,207.9L393.9,210.9Z"],["Américo de Campos","3501806",477.0,195.3,"M508.3,161.1L498.6,161.1L494.3,163.8L494.7,171.5L483.7,180.8L484.6,185.7L468.4,190.6L467,193L460.7,192.4L459,185.7L468.1,174.5L466.1,161.9L463.6,158.5L452.6,165.1L451.4,178.1L444.6,192.4L446.3,205L449,206.7L446,213L446.3,220.6L449.3,219L464.3,224.1L468.2,227.3L466.9,231.1L469.7,228.8L479.1,228.6L485.1,220.1L491.9,217.9L492.3,214.8L506,205.2L511.5,191.9L518.4,188.9L522.9,179.6L510.5,173.9L508.5,169.5L514.5,167.4L508.3,161.1Z"],["Aparecida D'Oeste","3502606",67.5,267.3,"M42.4,249.1L46,259L45.6,274L50.1,279.3L51.2,294L53.3,297L68.7,294.5L70.2,290.8L74,290.1L77.2,292.1L76.1,297.2L80.5,296.8L81.3,291.6L84.5,291.4L88.1,285.1L89.9,251.7L96.8,241.6L81.7,245.9L76.9,252.5L71.3,252.1L68.2,249L69.6,243.8L52.8,236.3L42.4,249.1Z"],["Ariranha","3503703",828.4,532.0,"M809.4,549.9L816.4,551.2L820.1,548.3L830.8,552L840,550L839.5,558.3L846.4,560.2L852.9,558.7L862.8,550.7L852.1,537.7L848.2,523.1L844.5,522.6L840.4,516.5L840.5,512.3L834.6,508L832.5,515L825.9,523.8L819.5,526.1L813.9,525.2L813.4,515.8L808.8,518.5L804.5,527.8L806.7,534.9L810.1,536.3L806.3,545.6L809.4,549.9Z"],["Aspásia","3503950",137.7,156.1,"M144.9,133.8L142.1,137.9L130.5,135.7L128.9,144L125.1,147.3L122.4,166.9L117,173.5L120.3,175.3L123.2,173.1L131.3,173.8L136.2,171.5L136.5,166.4L141.9,160.4L147.1,161.1L152.3,158L152.9,144.2L144.9,133.8Z"],["Bady Bassitt","3504602",601.9,440.4,"M573,445.3L576.2,444.3L581.6,447L594.6,460.6L620,451.1L625.1,443.1L630.3,440.2L617.5,432.4L609.7,430.8L605.1,435L604.9,428.3L597.5,422.5L597.9,420.2L583.6,423.4L571.9,429.1L573,445.3Z"],["Barretos","3505500",882.8,282.1,"M929.1,198.3L916.2,199L912.5,196.4L911,203.8L901.5,209.4L891.3,208.7L889.1,206.3L876.8,212.2L867.7,222.4L859.8,225.2L858.6,221.7L856.1,221.1L849.8,224L848.4,228.3L837.3,231L830.7,229.1L823.2,230.2L819.6,233.1L812.5,230.6L801.5,236.6L799.9,243.5L794.6,251.1L795.1,256.8L792.5,257.9L792.8,266L794.3,271.5L801.3,279.2L801.4,287.1L804.3,287.9L814.4,302.3L809.2,323L810.9,333.3L815.2,335.9L823.7,333.5L828,336.9L830.3,347.6L826.2,357.6L834.8,366.8L842,368L847.4,362.8L853.5,361.2L874.7,347.4L885.8,346.6L887.9,351.8L898.7,350.6L908.1,343.2L909,333.9L919.3,330.7L929.2,317.3L935,316.4L940.6,307L956.9,298.6L964.9,304.1L968.9,300L975.7,305.7L988.3,304L985.8,298.5L974.7,292.4L972.6,286.4L966.9,284.8L951.8,259.8L940.6,247.7L937.7,231.8L939.2,218.4L932.5,210.6L929.1,198.3Z"],["Bálsamo","3504800",559.2,353.4,"M559.4,320.1L554.3,323.8L554.1,328L546.5,337L546.6,340.8L541,346.5L541.7,352.1L539.1,352.3L533.6,360.9L533.4,370.4L530.3,373.7L541.7,377.6L548,385.7L555.3,382.3L558.5,376L577.2,366.7L580,359L581,349.2L576.3,345.8L578.9,338.8L577.4,336.4L569.2,336.4L564.3,333.8L563.2,323.7L559.4,320.1Z"],["Cajobi","3509304",807.9,418.0,"M787.2,425.1L795.7,432.9L810.5,433.1L823.5,441.7L830.6,441.1L845.5,425.8L841.8,419.8L845.5,415.7L843.7,412.1L845.2,405.1L826.6,403.5L810.3,406.5L808.3,402.5L800.9,401.1L798.9,398.5L792.7,399L782.6,394.9L772.2,416.4L772.6,419.7L775.1,422L786.4,421.1L787.2,425.1Z"],["Cardoso","3510708",420.5,109.0,"M463.6,158.5L467.2,143.8L461.8,126L465.8,112.3L461.1,110.8L458.2,105.4L459.6,102.6L453.1,99.2L453.6,95.5L445.2,94.5L441.5,100L435.2,98.3L433.8,92L436.4,82.6L434.6,80L437.5,76.4L431.8,68.9L430.1,60.1L407.3,53.5L395.4,56L380.8,52.2L376.6,66.3L376.6,77L370.2,80.5L369.6,86.9L356.9,91.8L356.8,98.6L364.2,100.5L364.7,107.4L370.1,111.5L380.3,107.1L384.4,119L382.1,124.1L386.5,127.8L388.4,136.1L385.4,143.9L386.7,153.6L399.4,155.2L409.2,150L417.3,151L417.1,154.9L420,155.7L420.8,160.4L427.7,167.5L432.3,167L439.2,159.8L449,157.5L450.9,159.4L458.5,155.7L463.6,158.5Z"],["Catanduva","3511102",760.7,516.6,"M750.2,563.1L765.6,559.9L767.1,555.6L764,543.5L780,529.7L783.5,530.4L781.7,524.4L783.7,522.2L782.6,510.6L784.9,507.7L791.1,510L795.4,508.5L803.9,493.5L818.6,482.2L819.5,479.6L813,472.9L804.2,468.6L793.8,479.8L783.2,486.7L776.2,487.9L774.2,491.3L767,486.7L762.6,494.1L757.1,497.7L747.1,498.8L744.7,504.9L738,511.8L737.7,518.1L733.2,522L733.5,524.9L739.8,525.9L746.4,535L739.7,540.6L732.8,539.5L721.4,550.8L732.9,558.9L744.4,560L750.2,563.1Z"],["Catiguá","3511201",735.0,491.9,"M767,486.7L760.8,475.7L751.3,475.2L750.7,470.9L741,471L737.7,467.2L718.9,477.3L714.3,476.7L711.9,481.1L707,483.4L706.3,488.9L711.1,498.9L720.7,508.8L723.8,508.8L727,516.2L737.7,518.1L738,511.8L744.7,504.9L747.1,498.8L757.1,497.7L762.6,494.1L767,486.7Z"],["Cedral","3511300",651.6,435.3,"M662.1,409.4L662.6,413.1L656.2,420.9L648.1,419.7L641.6,426.5L640.5,432.3L634.7,437.5L634.4,440.6L625.1,443.1L620,451.1L625.4,459.9L633.3,459.4L638.8,467.3L648.5,470.4L651.3,473.6L655.6,469.1L656.5,454.5L666.6,448.7L663.5,445.2L665.3,434.7L683.8,430.6L689,425L688.2,418.7L700.7,408L703.3,399.3L689.5,396.1L684,399.3L680.7,404.8L674.2,404.8L662.1,409.4Z"],["Colômbia","3512100",862.3,186.9,"M929.1,198.3L930.3,189.1L925.6,178.1L910.4,168L899.3,151L893.2,148.1L884.5,150.8L880.2,147.9L865.4,146.7L847.4,141.5L832.2,146.8L815,146.5L805.7,153L806.4,166.3L804.3,172.9L799.7,176.6L795.7,184.7L793.9,197.7L798.5,205.8L797.5,220.2L801.5,236.6L812.5,230.6L819.6,233.1L823.2,230.2L830.7,229.1L837.3,231L848.4,228.3L849.8,224L856.1,221.1L858.6,221.7L859.8,225.2L867.7,222.4L876.8,212.2L889.1,206.3L891.3,208.7L901.5,209.4L911,203.8L912.5,196.4L916.2,199L929.1,198.3Z"],["Cosmorama","3512902",479.0,245.7,"M522.3,183.5L518.4,188.9L511.5,191.9L506,205.2L492.3,214.8L491.9,217.9L485.1,220.1L479.1,228.6L469.7,228.8L449.3,243.4L446.3,249L448.2,261.5L434,267L425.2,278.5L423.6,289.3L425.4,293.2L423.5,297.3L426.9,302L430.8,300.8L439.5,304.7L442.8,309.9L449,304.4L454.1,303.5L459.7,294.4L467.4,287.5L481.3,285.5L488.8,288.6L490.6,277.8L496.1,276.3L501.9,270.2L504.1,260.3L510.5,248L509,242.5L510.7,234.5L517,222.6L533.4,207.4L531.5,203.9L533.6,201.3L531.2,199.2L529.4,201.6L527.5,199.8L522.3,183.5Z"],["Dirce Reis","3513850",172.1,256.0,"M156.4,251.4L157.6,270.1L162.2,281.7L164.8,278.2L170.7,280.9L176,270.7L183.6,264.2L187.4,254.2L186.1,245.4L190.8,236.1L177.3,234.9L177.4,227.7L170.6,238.4L159.5,244.6L156.4,251.4Z"],["Dolcinópolis","3514205",208.3,125.5,"M187.5,137.9L191.8,141.7L202.1,143.2L210.5,151.4L223.9,144.6L223.9,138.9L211.5,114.2L213.4,105.1L211,97.8L199,111.5L196.5,119.2L199.6,127.2L187.5,137.9Z"],["Elisiário","3514924",719.5,523.5,"M721.4,550.8L732.8,539.5L739.7,540.6L746.4,535L739.8,525.9L733.5,524.9L733.2,522L737.7,518.1L727,516.2L723.8,508.8L720.7,508.8L711.1,498.9L711.8,505.2L708.8,512.1L710,518.4L699.9,532.3L699.2,536.9L714.9,541.7L721.4,550.8Z"],["Embáuba","3514957",811.6,448.7,"M828.6,443.6L815.6,437.9L810.5,433.1L795.7,432.9L787.2,425.1L786.6,433.9L792.7,446.4L795.9,446.1L794.9,450.9L805,463.2L804.2,468.6L809.7,472.2L823.8,461.4L827.5,450.9L828.6,443.6Z"],["Estrela D'Oeste","3515202",249.6,182.6,"M275.2,213.5L278.3,199.5L267.3,183.3L272,170.5L268.6,167.4L266.7,155.1L263.3,149.4L262.9,138.6L255.7,123.7L248.5,139.3L235.5,144.8L230.3,155.1L229.8,158.9L235.3,166.1L230.1,174.7L232.7,188.8L226.2,194L224.7,198.8L229.3,204.6L230.3,211.6L228.1,215.3L229,233.4L242.8,241.7L246.3,229.8L256.3,219.8L257.2,215.5L266.3,215.1L270.1,211.5L275.2,213.5Z"],["Fernandópolis","3515509",308.2,193.2,"M344.7,204L340.8,190.3L334.2,186.5L337.7,182.1L337.9,176.8L314.2,167.9L318,161.4L316.2,151.6L301.3,139.7L300,135.8L304.3,125.1L288.8,121.4L286.2,118L279.3,124.4L264.2,122.7L257.5,124.6L262.9,138.6L263.3,149.4L266.7,155.1L268.6,167.4L272,170.5L267.3,183.3L278.3,199.5L275.2,213.5L281.2,226.7L279,230.5L281.3,239.7L276.7,247.3L274.5,248.8L270.8,246.7L266,250.9L262.2,257L263.2,264.8L273.4,265.9L278.5,263.5L279.8,266.3L283,265.5L287.1,258.1L294.8,255.5L303.4,241.8L310.7,240L310.4,236.6L318,229.9L314.2,222.6L322.6,212.3L321.9,210L327.5,205.8L334.3,208.4L344.7,204Z"],["Guapiaçu","3517505",683.0,359.1,"M719.8,385.2L716.4,382.1L719.9,377.3L719.9,373.2L714,365.9L711.7,366.9L708.6,355.8L710.7,345.5L707.8,342.4L710.2,338.1L707.7,335.6L715.6,332.4L709.4,327.8L712.5,322.6L709.4,319L711.5,315.1L710.4,312.1L707.3,311.4L696.2,328.7L691,331.6L690.1,337.1L680,338.5L674.9,343.7L658.7,344L656.4,347.1L657.1,357.3L653.2,378.9L659.3,397.4L656.3,401.7L662.1,409.4L674.2,404.8L680.7,404.8L684,399.3L689.5,396.1L703.3,399.3L715.9,391.1L719.8,385.2Z"],["Guaraci","3517901",745.8,227.0,"M806.8,292.3L801.4,287.1L801.3,279.2L794.3,271.5L792.8,266L791.8,260.7L796.1,252.7L788.2,251.8L767.2,238.2L764.6,233.1L767.7,184.7L764,163.5L759,149.4L748.1,142.8L730.4,142.8L710.9,187.5L720.5,194.8L721.7,213.4L727,220.2L723.5,236.7L724.9,240.2L722.2,244.3L725.3,248.3L723.1,251.9L723.8,258.6L727.3,256.7L734.8,258.7L741.2,268.6L740.1,270.8L746.8,279L750.1,279.8L752.1,288.6L750,293.6L746.6,294.2L748.7,301.9L752.2,305.4L767.4,310L768.9,304.1L772.2,302.4L793.3,295.9L796,292.9L806.8,292.3Z"],["Guarani D'Oeste","3518008",270.6,107.5,"M290.9,90.8L271.8,98.9L266.7,96.7L261.7,105.6L245.4,104.9L255.7,123.7L279.3,124.4L293.3,111.6L295,93.7L290.9,90.8Z"],["Ibira","3519402",674.1,493.9,"M640.9,514.9L662.9,517.9L672.9,513.9L684.1,520.5L686.2,519.1L688,521.9L687.7,530.1L692.8,535.6L699.2,536.9L699.9,532.3L710,518.4L708.8,512.1L711.8,505.2L711.1,498.9L707.1,493.2L702.2,496.2L699.8,494.8L692.1,486.1L688.8,476.2L679.4,467.5L677.8,460.7L674,459.6L666.6,451.3L662,450.4L656.5,454.5L655.6,469.1L652.9,473.1L642.7,469.7L641.2,476L645.2,481.2L650.2,495.6L649.5,502L642.2,504.7L640.9,514.9Z"],["Icém","3519808",689.0,223.4,"M643.8,203.2L645.3,209.5L650.6,213.6L656.4,214.4L656.2,219.5L649.2,225.6L648.1,229.8L656.1,228.4L659.3,234.7L658.6,245L664.6,247.2L674.5,242.3L673.8,245.4L669.4,247.6L672.8,250.5L667.9,255.3L673.9,258.4L675.5,264.5L679.6,262.4L691,265.4L694.8,259.8L702.4,262.4L706.5,261.5L709.8,259.1L713.4,248.1L718.4,243.1L722.2,244.3L724.9,240.2L723.5,236.7L727,220.2L721.7,213.4L720.5,194.8L710.9,187.5L692.8,203.2L688.4,204.4L681.4,200.6L673.5,200.8L661.2,183L654.8,186.2L647,201.4L643.8,203.2Z"],["Indiaporã","3520707",301.5,63.5,"M278,97.8L295.8,86.8L300.5,90.2L302.1,96.1L310.8,101.4L317.6,101.3L322.9,96L329.4,77.7L327.4,71.2L329.9,69L328,55.2L330.3,54.2L330.8,43.9L335.1,36.5L326.5,33.4L319.5,33.1L308.9,37L296.6,33.9L287,36L276,34.2L275.4,45.8L270.6,52.4L273.2,57.3L278,97.8Z"],["Ipiguá","3521150",611.1,327.9,"M587.7,344.7L600.3,351.8L602.2,363.3L605.5,362.5L606.7,358L613,353.1L622.5,360.4L629.9,343.3L624.6,340.7L628,326.9L623.7,315.6L606.4,301.6L601.8,294.7L594.6,291.7L595.5,302.9L600.9,308.8L602.8,316.6L593.7,328.9L590.4,341.7L587.7,344.7Z"],["Irapuã","3521507",617.3,570.7,"M569,553.5L599.1,585.1L604.6,587.6L612.3,608.5L623,599.1L626.4,602.1L636.4,600.2L638.9,597.1L639.2,586.5L652,579.3L649.9,568.8L642.7,564.5L641.1,561L637.2,560.4L634.6,556.5L624.2,558.1L622.3,550L615.3,547.1L612.1,535.3L595.9,538.5L589.5,534.1L591.4,537.4L589,541.3L581.4,545.8L574.4,545L569,553.5Z"],["Itajobi","3521903",738.3,606.7,"M750.2,563.1L744.4,560L732.9,558.9L729.1,565.7L721.2,570.4L720.9,573.5L709.5,580.5L696.9,575.8L678.3,585.4L676.4,589L670.3,590L675.5,602.7L687.4,605.5L692.1,616.1L689.2,620.2L692.5,623.8L725.1,622.1L729.7,632.8L729.5,638L724.6,643.8L734.4,653L739.3,654.2L739.5,652.1L747.1,650.6L757.1,640.7L759.8,640.6L758.8,633.8L761.4,625.3L764.5,628.7L771.9,628.9L772.5,618.9L789.7,607.9L781.4,598.2L782.9,589.7L781.1,578.3L768.5,572.9L765.9,568.5L756.3,575.6L750.2,568.9L750.2,563.1Z"],["Jaci","3524501",543.6,444.6,"M527,455.5L529.3,460.5L534.9,463.4L535.3,467.3L538.1,467.2L538,474.6L539.7,475.5L545.3,475.7L550.3,472.8L554.8,475.7L561.8,471L567.2,456.8L551.8,437L562.3,422.8L561.6,415.7L557.2,415.7L552.8,412.4L530,420.4L533.4,425.4L530.5,428.3L527,455.5Z"],["Jales","3524808",208.7,185.0,"M210.5,151.4L202.1,143.2L191.8,141.7L178.1,134L180.9,146L177.7,161.3L185,184.4L183.3,189.4L175.3,189.3L175.2,192.8L170.2,196.2L168.5,203.8L164.9,202.7L161.8,209L163.5,220.8L177.4,227.7L177.3,234.9L190.8,236.1L202.9,228.2L214.4,234.9L225.2,230.3L229,233.4L228.1,215.3L230.3,211.6L229.3,204.6L224.7,198.8L226.2,194L232.7,188.8L232.8,185.5L230.1,174.7L222.6,177.9L212.2,166.5L210.5,151.4Z"],["José Bonifácio","3525706",484.5,500.5,"M523.2,526.1L535.6,512.2L548.1,510.2L550.6,497.6L549,495L553.2,489.8L551.2,485.1L554.8,475.7L550.3,472.8L545.3,475.7L538,474.6L538.1,467.2L535.3,467.3L534.9,463.4L529.3,460.5L527,455.5L520.2,464L520.1,459.2L515.8,453.2L505.2,453.2L496.4,457.1L497.4,450.8L501.8,447.4L494.4,440.4L480.3,443.8L465.9,441.8L458.4,444.2L458,450.7L446.8,468.4L448.4,473.3L432.3,495L428.2,492.8L423.4,495.2L416.1,504.1L415.5,513.2L410.8,517.2L408,530.5L404.4,534.7L403,551.1L399.6,552.9L419.2,555.2L424.8,552.6L427.1,557.3L438.6,561.7L442.8,565.6L447.8,565.8L457.9,558.7L464,543.7L472,546.2L480.9,525L485,524.6L490.9,517.7L512.6,517.9L514,516L518.1,517.8L517.8,527.9L523.2,526.1Z"],["Macaubal","3528106",407.7,410.1,"M429.8,412.6L427.5,395.1L431.4,388.6L430.4,374.7L424.5,367.2L423.3,360.6L419.9,361L414.6,369.4L405.6,370.3L399.6,366.9L389.9,374.7L386.2,387.5L388,404.5L383.9,413.5L385.4,424.4L381.6,432.8L382.8,435.3L378.4,440.1L380.1,443.7L383.7,444.6L385,440.8L388.5,444.7L384.8,458L390.7,456.9L398.8,447.1L405.3,432.2L415.5,428.2L429.8,412.6Z"],["Macedônia","3528205",334.3,119.7,"M317.2,162.6L324.2,156.6L333.6,160.5L340,157.2L343.8,146L353.3,143.6L358.2,139.8L358,133.6L362,129.3L378.6,122.5L383.9,116.4L381.1,113.4L380.3,107.1L370.1,111.5L364.7,107.4L364.2,100.5L361.7,98.9L352.6,99.4L349.8,102.7L340.4,98.3L330.2,85.6L328.6,78.9L324.7,92.8L317.6,101.3L310.8,101.4L302.1,96.1L300.5,90.2L295.8,86.8L290.9,90.8L295,93.7L294.7,108.6L286.2,118L288.8,121.4L304.3,125.1L300,135.8L301.3,139.7L316.2,151.6L317.2,162.6Z"],["Marapoama","3528858",710.2,563.2,"M702.2,537.9L702,548.8L695.3,551.2L692.3,560.7L674.8,578.3L670.3,590L676.4,589L678.3,585.4L696.9,575.8L709.5,580.5L729.1,565.7L732.9,558.9L730.6,555.5L721.4,550.8L714.9,541.7L702.2,537.9Z"],["Marinópolis","3529104",99.1,269.7,"M96.8,241.6L89.9,251.7L88.1,285.1L84.5,291.4L91,291.6L93.1,293.7L94.7,288.9L109,285L108.2,271.9L110.9,267.6L110.6,244.5L96.8,241.6Z"],["Mendonça","3529500",552.4,542.8,"M569,553.5L574.4,545L585.4,544.2L591.4,537.4L587.5,527.2L576.8,519.1L574.1,522.3L573.9,527.9L557.4,527L550.9,524.3L550.8,521.6L546.5,518.7L541.2,520.5L532.9,515.1L529.5,520.7L508.2,536.6L509.5,544.3L516.8,541.7L526.6,547.1L530.2,543.8L536.1,545.9L538.9,551.3L535.1,566.2L540.6,572L543.9,571.6L546,566.7L549.4,569.9L554.5,569.3L553.9,566.2L557.7,566.9L565,554.4L569,553.5Z"],["Meridiano","3529609",328.8,238.3,"M344.7,204L334.3,208.4L327.5,205.8L321.9,210L322.6,212.3L314.2,222.6L318,229.9L310.4,236.6L310.7,240L303.4,241.8L294.8,255.5L287.1,258.1L285.3,263.5L296.1,262.5L304.3,267.5L306.8,265.8L308.6,271.7L315.8,273.3L320.5,266.6L327.3,266.8L328.2,263.6L337.2,257L341.8,257.3L348,246L346.8,232L363.6,227.3L367.9,205.4L367.5,202.7L355.9,206.9L344.7,204Z"],["Mesópolis","3529658",177.0,67.4,"M164.4,85.6L173,89L175.4,96.3L186.4,102.4L191.2,91.5L199.3,81L185,47.4L178.2,38.4L177,31L170.7,34.8L161.3,48.6L151.2,52.1L159.8,68.2L158.2,70.2L159,78.9L164.4,85.6Z"],["Mira Estrela","3530003",353.5,69.7,"M327.4,71.2L330.2,85.6L336.8,94.7L349.8,102.7L356.8,98.6L356.9,91.8L369.6,86.9L370.2,80.5L376.6,77L376.6,66.3L380.8,52.2L361.8,36.5L335.1,36.5L330.8,43.9L330.3,54.2L328,55.2L329.9,69L327.4,71.2Z"],["Mirassol","3530300",573.3,401.4,"M581,349.2L577.2,366.7L558.5,376L555.3,382.3L548,385.7L546.9,393.4L552.7,398.3L549,414.3L552.8,412.4L563.1,417.7L562.3,422.8L551.8,437L567.2,456.8L573.9,440.5L571.9,429.1L583.6,423.4L599.5,419.7L599.2,410.2L589.9,384.8L594.9,375.3L595.5,368.8L590,363.7L589.1,345.9L587.7,344.7L581,349.2Z"],["Mirassolândia","3530409",581.9,313.4,"M585.5,280.5L576.4,274.2L564.1,288.3L561.8,319L559.4,320.1L563.2,323.7L564.3,333.8L569.2,336.4L577.4,336.4L578.9,338.8L576.3,345.8L581,349.2L590.4,341.7L593.7,328.9L602.8,316.6L600.9,308.8L595.5,302.9L594.6,291.7L585.4,285.5L585.5,280.5Z"],["Monte Aprazível","3531407",506.7,362.5,"M455.9,408.7L464.1,402.8L467.8,403.1L476.1,412.3L478.8,407.9L485.1,412L487.7,408.2L494.2,407.4L498.5,408.3L503.3,413.1L502.9,400.8L509.5,402.5L524.3,395.1L528,390.1L519.6,379L533.4,370.4L532.2,364.8L537.7,354.5L530.6,357.3L524.2,355.1L522,347.9L500.6,327.5L491.5,327.5L487.3,322.8L482,322.8L476.9,319.6L471.1,310.6L459.7,314.1L454.2,318.9L457.9,326L440.8,339.7L438.6,346.6L431,355.5L426.4,351.2L415.1,351.4L422.4,357.2L424.5,367.2L430.4,374.7L430.7,369.7L439.3,361.1L449.4,364.7L449.8,359.1L459.1,350L460.9,341.7L467.9,343.5L481.6,358.6L479.6,364.7L480.7,374.9L477.2,384.3L478.7,388.5L476.2,394.4L468.1,396.6L447.4,390.8L442.3,375.7L435.6,373.9L430.4,375.2L431.4,388.6L427.5,395.1L429.8,412.6L438.7,407L448.1,409.7L450.2,406.2L455.9,408.7Z"],["Monções","3531001",365.4,418.9,"M338,418.5L343.4,418L349.1,421.3L355.6,419.3L358.9,422.8L363.5,423.2L373.4,436.3L374.4,443.1L382.8,435.3L381.6,432.8L383.7,431.2L385.9,418.2L383.9,413.5L386,412.3L387.9,399L385.9,396.8L373.6,397.2L372,399.3L368.6,395.2L363.4,394.1L342.7,409.8L338,418.5Z"],["Neves Paulista","3532504",519.7,418.2,"M530.3,373.7L519.6,379L528,390.1L524.3,395.1L509.5,402.5L502.9,400.8L503.4,420.1L500,425.4L500,431.5L494.6,437L494.4,440.4L501.8,447.4L497.4,450.8L496.4,457.1L505.2,453.2L515.8,453.2L520.1,459.2L520.2,464L528.7,451.8L530.5,428.3L533.4,425.4L530,420.4L549,414.3L552.7,398.3L546.9,393.4L549,388.9L546.7,382.1L537.9,375.5L530.3,373.7Z"],["Nhandeara","3532603",384.4,341.8,"M415.1,351.4L412.1,337L406.5,335.7L401.7,316.3L403.9,309.4L402.7,302.8L397.9,295.5L390.4,290.1L378.3,288.6L371.2,291.1L368.5,287.2L357.9,281.2L353.6,283.5L350.6,297.1L351.8,301.8L355.8,313.5L363.7,319.9L364.9,325.2L360.2,341L354.4,342.6L351.6,351.2L354.3,360.1L351.9,369.7L357.2,381.7L351.3,392.7L353.2,401.8L363.4,394.1L368.6,395.2L372,399.3L373.6,397.2L385.9,396.8L387.9,399L386.2,387.5L389.9,374.7L399.6,366.9L405.6,370.3L414.6,369.4L419.9,361L423.3,360.6L422.4,357.2L415.1,351.4Z"],["Nipoã","3532702",477.2,426.6,"M503.3,413.1L498.5,408.3L494.2,407.4L487.7,408.2L485.1,412L478.8,407.9L476.1,412.3L467.8,403.1L461.1,404.2L455.9,408.7L453.4,415.5L454.6,427.9L449.9,440.3L451.5,445.1L458.3,448.6L458.4,444.2L465.9,441.8L480.3,443.8L493.8,442L496.4,433.7L500,431.5L503.3,413.1Z"],["Nova Aliança","3532801",570.4,486.2,"M594.6,460.6L581.6,447L576.2,444.3L571.6,446L564.6,465.9L561.8,471L554.8,475.7L551.2,485.1L553.2,489.8L549,495L550.6,497.6L548.1,510.2L535.6,512.2L532.9,515.1L541.2,520.5L546.5,518.7L550.8,521.6L550.9,524.3L557.4,527L573.9,527.9L574.1,522.3L581.3,516.6L585.3,508.3L584.8,503.9L590.1,499.4L588.4,481.3L590.9,466.7L594.6,460.6Z"],["Nova Canaã Paulista","3532843",67.7,225.4,"M48.9,235.5L69.6,243.8L68.2,249L71.3,252.1L76.9,252.5L81.7,245.9L95,242.6L82.1,205.6L65.9,197.8L58.1,215.2L53,216.2L49.2,220.5L47.5,225.1L48.9,235.5Z"],["Nova Granada","3533007",638.9,256.8,"M678.1,266.8L673.9,258.4L667.9,255.3L672.8,250.5L669.4,247.6L673.8,245.4L674.5,242.3L664.6,247.2L660.6,246.7L657.6,239.2L659.3,234.7L656.1,228.4L648.1,229.8L649.2,225.6L656.2,219.5L656.4,214.4L650.6,213.6L645.3,209.5L644.2,199.4L633.9,206L611.9,228.6L607.8,237.1L611.6,250.9L608.6,251.9L606.9,258.3L596,264.7L588.6,275.6L583.1,277.2L585.5,280.5L585.4,285.5L590.2,290.2L601.8,294.7L606.4,301.6L623.7,315.6L632.5,316.3L647.3,306.5L660.7,303.1L674.9,291.8L682.5,288.9L681.5,287L686.1,280.3L678.5,276.2L678.1,266.8Z"],["Nova Luzitânia","3533304",310.8,416.6,"M338,418.5L325.9,413.1L320.3,415.1L310.9,404.2L295.6,403L280.9,409.1L276.3,407.5L281,417.7L298,417.7L297.6,421.8L302.6,430.3L315.6,423.7L317.6,427.2L331.9,428.3L338.3,423.3L338,418.5Z"],["Novais","3533254",786.8,461.3,"M786.6,433.9L781.6,433.5L771.1,439.9L768.3,444.9L766.9,453.6L770.1,459.3L767.8,478L771.7,487.4L775.5,489.1L793.8,479.8L805.5,465.6L800.1,455.6L794.9,450.9L795.9,446.1L792.7,446.4L786.6,433.9Z"],["Novo Horizonte","3533502",646.4,647.3,"M713.1,657.2L729.5,638L729.7,632.8L725.1,622.1L692.5,623.8L689.2,620.2L692.1,616.1L687.4,605.5L675.5,602.7L670.3,590L672.1,586.6L657.3,575.5L652.2,576.8L652,579.3L645.2,584.7L639.2,586.5L638.9,597.1L636.4,600.2L626.4,602.1L623,599.1L607.9,612L606.6,610.6L598.9,625.4L592.7,631.1L587.7,631.4L586.9,634.5L577.6,633L567.5,637.3L558.4,627.9L576.9,655.2L601.4,672.3L606.8,674.4L618.5,670.7L631.9,674.8L636.6,679.9L642.7,702.2L650.3,708.8L652.8,714.5L659.4,720.3L676.4,723.6L677.8,722.8L672.7,718.4L667.8,704.9L671.4,703.6L675.5,697.6L687,694L690.4,690.8L691.5,686.3L698.3,678.8L700.3,668.9L706.4,663.7L712.2,662.4L713.1,657.2Z"],["Olímpia","3533908",769.2,353.7,"M806.8,292.3L796,292.9L793.3,295.9L772.2,302.4L768.9,304.1L767.4,310L757.5,306.4L754.4,326.9L751.6,327.5L748.7,332.5L736.1,324.3L732.8,319.1L724.9,322.2L722.4,314.1L711.5,315.1L709.4,319L712.5,322.6L709.4,327.8L715.6,332.4L707.7,335.6L710.2,338.1L707.8,342.4L710.7,345.5L708.6,355.8L711.7,366.9L714,365.9L719.9,373.2L719.9,377.3L716.4,382.1L721.8,387.6L720.7,394.2L730.6,394.4L746,409.4L750,411.1L753.9,408.3L759.7,408.4L772.2,416.4L782.6,394.9L792.7,399L797.5,391.1L810.4,379.2L818.1,379.1L824,361.7L828.2,361.3L826.2,357.6L830.3,347.6L828,336.9L823.7,333.5L815.2,335.9L810.9,333.3L809.2,323L814.4,302.3L806.8,292.3Z"],["Onda Verde","3534005",663.9,318.0,"M623.7,315.6L630.9,329.9L630.8,337.3L644.7,345L674.9,343.7L680,338.5L690.1,337.1L691,331.6L696.2,328.7L707.3,311.4L705.4,306.9L697.4,303.5L691.2,292.8L682.5,288.9L674.9,291.8L660.7,303.1L647.3,306.5L632.5,316.3L623.7,315.6Z"],["Orindiúva","3534203",622.6,163.4,"M589.7,155.4L590.5,171.1L595.8,176L602,167.2L607.4,169.8L610.5,175.7L618.9,179.2L629.2,179.9L628.5,187.8L635.6,189.3L639.4,193.1L639.1,195.8L644.2,199.4L643.8,203.2L647,201.4L654.8,186.2L661.2,183L659.2,171.9L647.1,148L644.7,128.8L634.8,130L629.7,136.2L624.6,135.2L619.4,139.6L593.2,133.4L591.8,135L595.7,144.8L589.7,155.4Z"],["Ouroeste","3534757",253.9,55.3,"M278,97.8L273.2,57.3L270.6,52.4L275.4,45.8L276,34.2L270.4,32.4L247.5,8.1L234.2,2.4L225,17.9L224.8,25.2L227.3,34L231.8,38.3L234.8,51.9L233.3,53.3L238.9,58.2L236.9,66.7L241.6,70.8L243.9,103L255.6,106.4L261.7,105.6L266.7,96.7L271.8,98.9L278,97.8Z"],["Palestina","3535002",583.6,214.2,"M644.2,199.4L641.3,198.7L635.6,189.3L628.5,187.8L629.2,179.9L618.9,179.2L610.5,175.7L607.4,169.8L602,167.2L595.8,176L590.5,171.1L590.3,167.5L583.4,168.8L579,175.3L576.1,174.7L576.7,171.9L569.6,165.7L563.2,169.4L557.9,164.6L555.1,166.1L536.2,155.2L532.3,155.8L530.8,151.6L528.5,153.5L524.5,150.7L521.7,152.9L520.1,149.3L518.7,152.1L507.3,155.4L506.7,159L514.5,167.4L508.5,169.5L510.5,173.9L522.9,179.6L522.3,183.5L527.5,199.8L529.4,201.6L531.2,199.2L533.6,201.3L531.5,203.9L536.1,211.8L548.3,218.3L552.6,232.1L555.9,232.5L555.5,236.2L558.8,238.3L556.7,243.4L557.5,249.3L561.1,257.9L565.9,260.9L567.8,271.2L573.4,269.2L576.4,274.2L583.1,277.2L588.6,275.6L596,264.7L606.9,258.3L608.6,251.9L611.6,250.9L607.8,237.1L611.9,228.6L633.9,206L644.2,199.4Z"],["Palmares Paulista","3535101",815.1,502.8,"M808.8,518.5L813.4,515.8L813.9,525.2L819.5,526.1L825.9,523.8L832.5,515L834.6,508L833.8,504.2L831.2,504.6L829,497.2L822.2,493.2L823.9,490.2L819.1,486.1L818.6,482.2L803.9,493.5L800.9,501.4L795.9,506.6L799.9,510.7L799.5,517.6L808.8,518.5Z"],["Palmeira D'Oeste","3535200",133.4,252.0,"M106.3,287.1L108.1,293.1L110.9,291.1L121.8,292.3L122.4,295.5L126.4,293.7L126.4,291.1L132.4,292.4L137.4,285.5L144.6,288.3L149.9,284.6L158.4,287.5L162.2,281.7L157.6,270.1L157.8,257.9L153.3,245.9L153.6,235.2L143.5,228.2L140.9,214.1L134.9,208.9L106.6,213.7L101.1,210.6L105.3,219.2L102,224.3L102.5,227.2L98.3,230.6L91.4,231.8L95,242.6L110.6,244.5L110.9,267.6L108.2,271.9L109,285L106.3,287.1Z"],["Paranapuã","3535903",180.6,109.5,"M201.1,79.4L191.2,91.5L186.4,102.4L175.4,96.3L173,89L164.4,85.6L159.1,104.3L161.4,111.3L160.4,117.5L174.9,127L178.1,134L187.5,137.9L199.6,127.2L196.5,119.2L199,111.5L211,97.8L212,91.6L207.8,85.7L205,85.2L204.3,81.7L201.1,79.4Z"],["Paraíso","3535705",837.4,472.6,"M819.5,479.6L819.1,486.1L823.9,490.2L822.2,493.2L829,497.2L831.2,504.6L833.8,504.2L836,501.1L852.3,494.4L855.6,484.4L864.1,471L860.8,467.6L860.7,463.4L857.5,461.6L858.2,459L851.4,458.9L845.8,450.3L838.6,445.9L837.7,447.3L830.6,441.1L827.1,441.7L828.6,443.6L827.5,450.9L823.8,461.4L809.7,472.2L813,472.9L819.5,479.6Z"],["Parisi","3536257",381.8,182.0,"M393.9,210.9L392.6,207.9L396.7,200.7L395.9,190.7L392.7,185.9L390.9,175.5L391.8,162.3L387.7,159.1L386.7,153.6L385,161.6L380.4,162.7L371,178.1L372.2,187.5L367.6,196.8L377.5,212.6L383.2,205.2L390.9,212.4L393.9,210.9Z"],["Paulo de Faria","3536604",599.8,111.7,"M664.6,72.4L659.2,69.6L647.7,69L626.6,79L598.7,77.4L591.8,75.7L577.9,65.2L575.3,55.8L570.3,51.4L556.9,48.1L534.8,53.5L541.7,64.4L542.3,70.1L548.5,75.6L557.3,92.9L557.1,110L551.2,117L548.6,115.8L546.8,117.7L542.1,114.2L539.9,128.6L533.4,140.8L530.8,151.6L532.3,155.8L536.2,155.2L555.1,166.1L557.9,164.6L563.2,169.4L569.6,165.7L576.7,171.9L576.1,174.7L579,175.3L583.4,168.8L590.3,167.5L589.7,155.4L595.7,144.8L591.8,135L593.2,133.4L619.4,139.6L624.6,135.2L629.7,136.2L634.8,130L644.7,128.8L644.2,107.2L647.7,90.8L652.5,86L660.2,86.8L665.2,84L664.6,72.4Z"],["Pedranópolis","3536901",350.5,162.1,"M383.9,116.4L378.6,122.5L362,129.3L358,133.6L358.2,139.8L353.3,143.6L343.8,146L340,157.2L333.6,160.5L324.2,156.6L314.1,166.2L325.2,173.1L337.9,176.8L337.7,182.1L334.2,186.5L340.8,190.3L344.7,204L355.9,206.9L368.7,200.5L367.6,196.8L372.2,187.5L371,178.1L380.4,162.7L385,161.6L386.7,153.6L385.4,143.9L388.4,136.1L386.5,127.8L382.1,124.1L383.9,116.4Z"],["Pindorama","3538105",785.6,544.5,"M795.9,506.6L791.1,510L784.9,507.7L782.6,510.6L783.7,522.2L781.7,524.4L783.5,530.4L780,529.7L764,543.5L767.1,555.6L765.6,559.9L750.2,563.1L750.2,568.9L756.3,575.6L765.9,568.5L768.5,572.9L776.4,575.6L782.6,580.8L790.1,566.7L796.6,568L798.8,566.4L800.8,568.4L802.7,562.4L808.4,556.7L804.2,553.7L809.4,549.9L806.3,545.6L810.1,536.3L806.7,534.9L804.5,527.8L808.8,518.5L799.5,517.6L799.9,510.7L795.9,506.6Z"],["Planalto","3539608",418.5,474.4,"M376.9,465.3L398,478.8L402,483.8L403.3,500L397.7,504.1L410.8,517.2L415.5,513.2L416.1,504.1L420.5,497.9L428.2,492.8L432.3,495L445.7,477.7L448.4,473.3L446.8,468.4L458.3,448.6L451.5,445.1L449.9,440.3L438.6,434.6L424,435.9L421.2,440.7L416.8,441.6L411.9,430.7L405.3,432.2L398.8,447.1L390.7,456.9L382.5,459.1L376.9,465.3Z"],["Poloni","3539905",456.5,367.2,"M430.4,375.2L435.6,373.9L442.3,375.7L447.4,390.8L468.1,396.6L476.2,394.4L478.7,388.5L477.2,384.3L480.7,374.9L479.6,364.7L481.6,358.6L467.9,343.5L460.9,341.7L459.1,350L449.8,359.1L449.4,364.7L439.3,361.1L430.7,369.7L430.4,375.2Z"],["Pontalinda","3540259",212.8,256.5,"M242.8,241.7L225.2,230.3L214.4,234.9L202.9,228.2L198.1,229.9L187.3,241.8L187.4,254.2L183.6,264.2L176,270.7L173.2,278.3L177.3,282.6L184.8,284L191.6,282L198.5,286L201.8,276.6L207.8,278.6L236.2,269.9L240.2,245.7L242.8,241.7Z"],["Pontes Gestal","3540309",489.4,152.4,"M463.6,158.5L466.1,161.9L468.1,174.5L459,185.7L460.7,192.4L467,193L468.4,190.6L484.6,185.7L483.7,180.8L494.7,171.5L494.3,163.8L498.6,161.1L508.3,161.1L507.3,155.4L518.7,152.1L520.8,145.6L518.4,143L510,143.7L508.7,142.3L510.9,136.6L503.6,133.1L504.1,128.7L500.1,127.1L496,130.7L490,127.6L486.4,118.4L477.2,113.7L472.5,117.2L467.2,111.3L462.4,120L461.8,126L467.2,143.8L463.6,158.5Z"],["Populina","3540408",209.8,49.6,"M177,31L178.2,38.4L185,47.4L199.3,81L201.1,79.4L205,85.2L207.8,85.7L212,91.6L211.4,95.3L219.3,98.3L230,92.6L241.7,77.9L241.6,70.8L236.9,66.7L238.9,58.2L233.3,53.3L234.8,51.9L231.8,38.3L227.3,34L224.8,25.2L225,17.9L234.2,2.4L228.8,0L218.2,6.1L211.7,6.3L201.8,11.5L190,13.9L177,31Z"],["Potirendaba","3540804",619.2,495.1,"M642.7,469.7L633.3,459.4L625.4,459.9L620,451.1L596.2,458.8L590.9,466.7L588.4,481.3L590.1,499.4L584.8,503.9L585.3,508.3L581.3,516.6L576.8,519.1L587.5,527.2L589.5,534.1L595.9,538.5L612.1,535.3L617.2,524.9L625.5,528.4L631.8,525.2L637.2,516.6L640.9,514.9L642.2,504.7L649.5,502L650.2,495.6L645.2,481.2L641.2,476L642.7,469.7Z"],["Riolândia","3544202",507.7,104.0,"M461.1,110.8L467.2,111.3L472.5,117.2L477.2,113.7L486.4,118.4L490,127.6L496,130.7L500.1,127.1L504.1,128.7L503.6,133.1L510.9,136.6L508.7,142.3L510,143.7L518.4,143L521.7,152.9L524.5,150.7L528.5,153.5L539.9,128.6L542.1,114.2L546.8,117.7L548.6,115.8L551.2,117L557.1,110L557.3,92.9L548.5,75.6L542.3,70.1L541.7,64.4L534.8,53.5L527.2,58.1L474.3,55.3L455.4,58.6L448.8,63.7L436.5,63.2L430.1,60.1L431.8,68.9L437.5,76.4L434.6,80L436.4,82.6L433.8,92L435.2,98.3L441.5,100L445.2,94.5L453.6,95.5L453.1,99.2L459.6,102.6L458.2,105.4L461.1,110.8Z"],["Rubinéia","3544509",22.0,183.2,"M49.7,238.3L35.2,207.1L35.6,198.8L30,190.4L34.7,179.8L39,178.8L47.1,171.1L46.1,159.3L48.5,145.3L56.4,144.7L68.3,134L45.9,124.5L44.5,121.5L38.7,120.9L28.6,141.8L18.1,174L14.4,179.7L0,191.2L7.5,196.4L10.9,209.6L21.4,223L31.6,231.5L33.4,238.4L42.4,249.1L49.7,238.3Z"],["Sales","3544806",576.6,595.7,"M612.3,608.5L604.6,587.6L599.1,585.1L569,553.5L565,554.4L557.7,566.9L553.9,566.2L554.5,569.3L549.4,569.9L551,576.9L548.4,582L545.3,582.6L546.8,590.5L543,597L544,604.7L529.4,601.7L538,608.9L545.5,611.3L554,624.9L557.9,626.4L567.5,637.3L570.2,634.5L586.9,634.5L587.7,631.4L592.7,631.1L598.9,625.4L603.1,614.3L612.3,608.5Z"],["Santa Adélia","3545605",816.0,588.4,"M781.4,598.2L792.6,613.1L799.4,617L814.2,617.5L816.7,622L823.8,624.2L827,622.8L835,629.7L841.8,627.1L845.1,616.3L844.7,602.3L850.2,587.1L849,566.1L845.9,563.4L846.4,560.2L839.5,558.3L840,550L830.8,552L820.1,548.3L816.4,551.2L809.4,549.9L804.2,553.7L808.4,556.7L802.7,562.4L800.8,568.4L798.8,566.4L788.1,568.6L781.2,582.2L782.9,589.7L781.4,598.2Z"],["Santa Albertina","3545704",138.1,95.7,"M94.8,71.7L106.7,78.6L109.9,83L111.5,91.6L122.3,103.3L124.2,115L132,127.3L130.5,135.7L134.4,137.7L143.8,137.2L145.9,129.6L155.9,117.6L155.5,111.5L159.4,107.6L164.4,85.6L159,78.9L158.2,70.2L159.8,68.2L153.8,54.2L151.2,52.1L116.5,60.1L94.8,71.7Z"],["Santa Clara D'Oeste","3546108",70.0,110.9,"M68.3,134L71.6,138.2L77.9,137.1L80.7,134L85.1,134.7L88.1,121.2L91,122.7L101.6,104.5L96.8,94.8L79.9,80.4L72,81.9L58.6,95.5L50.2,98L38.9,116.8L38.7,120.9L44.5,121.5L45.9,124.5L68.3,134Z"],["Santa Fé do Sul","3546603",48.8,185.5,"M68.3,134L56.4,144.7L48.5,145.3L46.1,159.3L47.1,171.1L39,178.8L34.7,179.8L30,190.4L35.6,198.8L35.2,207.1L49.7,238.3L47.5,225.1L53,216.2L58.1,215.2L63.3,205.8L61.6,192L68.8,179.6L69.3,172.1L83.7,155.7L87.4,147.5L85.3,143.9L85.1,134.7L80.7,134L77.9,137.1L71.6,138.2L68.3,134Z"],["Santa Rita D'Oeste","3547403",109.9,119.4,"M79.9,80.4L80.3,82.7L91.5,89.4L100.6,100.7L101.6,104.5L97.8,108.2L97.2,113.6L91,122.7L88.1,121.2L84.4,138.4L88.6,151.2L102.4,166.9L106.1,166.4L108.2,164.4L106.4,154.5L110.7,148.2L113.5,146L125.1,147.3L128.9,144L132,127.3L124.2,115L122.3,103.3L111.5,91.6L109.9,83L106.7,78.6L94.8,71.7L86.3,79L79.9,80.4Z"],["Santa Rita d'Oeste","3547403",109.9,119.4,"M79.9,80.4L80.3,82.7L91.5,89.4L100.6,100.7L101.6,104.5L97.8,108.2L97.2,113.6L91,122.7L88.1,121.2L84.4,138.4L88.6,151.2L102.4,166.9L106.1,166.4L108.2,164.4L106.4,154.5L110.7,148.2L113.5,146L125.1,147.3L128.9,144L132,127.3L124.2,115L122.3,103.3L111.5,91.6L109.9,83L106.7,78.6L94.8,71.7L86.3,79L79.9,80.4Z"],["Santa Salete","3547650",139.0,184.3,"M129,173.7L125.3,183.3L127,196.7L115.9,212.5L123.1,212.2L131.2,208.1L146,188.2L150.7,187L156.6,181.1L157.4,168.7L153.5,165.4L152.3,158L147.1,161.1L141.9,160.4L136.5,166.4L136.2,171.5L129,173.7Z"],["Santana da Ponte Pensa","3547205",113.4,180.5,"M102.4,166.9L102.6,174.9L96.2,186.6L97.2,192.2L89.5,199.2L89.8,203.9L106.6,213.7L115.9,212.5L127,196.7L125.3,183.3L129,173.7L123.2,173.1L120.3,175.3L117,173.5L122.4,166.9L125.1,147.3L113.5,146L110.7,148.2L106.4,154.5L108.2,164.4L102.4,166.9Z"],["Sebastianópolis do Sul","3551306",430.3,328.6,"M403.8,300.9L403.9,309.4L401.7,316.3L406.5,335.7L412.1,337L415.1,351.4L426.4,351.2L431,355.5L438.6,346.6L440.8,339.7L457.9,326L454.2,318.9L445.7,314.4L442.6,307.2L430.8,300.8L426.9,302L424,298.9L417.8,298.5L412.4,304.1L403.8,300.9Z"],["Severínia","3551900",829.4,383.9,"M828.2,361.3L824,361.7L818.1,379.1L810.4,379.2L797.5,391.1L792.7,399L798.9,398.5L800.9,401.1L808.3,402.5L810.3,406.5L826.6,403.5L845.2,405.1L844.5,401.3L854.3,394.8L854.7,386.1L851.8,376.7L847.8,371.4L844,371.8L842,368L834.8,366.8L828.2,361.3Z"],["São Francisco","3549003",159.7,226.3,"M131.2,208.1L140.9,214.1L143.5,228.2L152.2,232.6L153.3,245.9L156.4,251.4L159.5,244.6L170.6,238.4L177.4,227.7L174.6,225L165.1,223.2L161.8,209L156.9,211.5L152,210.3L149.1,204.5L139.3,202L134.2,203.6L131.2,208.1Z"],["São José do Rio Preto","3549805",622.6,384.0,"M599.5,419.7L597.5,422.5L604.9,428.3L605.1,435L609.7,430.8L634.4,440.6L634.7,437.5L640.5,432.3L641.6,426.5L648.1,419.7L656.2,420.9L662.6,413.1L656.3,401.7L659.3,397.4L653.2,378.9L657.1,357.3L656.4,347.1L658.7,344L649.3,345.8L636.7,342.1L630.8,337.3L630.9,329.9L628,326.9L624.6,340.7L629.9,343.3L622.5,360.4L613,353.1L606.7,358L605.5,362.5L602.2,363.3L600.3,351.8L589.1,345.9L590,363.7L595.5,368.8L594.9,375.3L589.9,384.8L599.2,410.2L599.5,419.7Z"],["São João das Duas Pontes","3549201",261.8,240.7,"M275.2,213.5L270.1,211.5L266.3,215.1L257.2,215.5L256.3,219.8L246.3,229.8L240.2,245.7L237.2,268.7L242.7,266.1L250.1,267.8L263.2,264.8L262.2,257L266,250.9L270.8,246.7L274.5,248.8L276.7,247.3L281.3,239.7L279,230.5L281.2,226.7L275.2,213.5Z"],["São João de Iracema","3549250",266.6,284.4,"M236.2,269.9L233.8,286.1L237.3,289.8L242.1,288.9L248,294.5L248.5,298.5L255,301.3L274,302.1L276.8,300.4L294.4,308.1L298.2,304L295.1,295.8L295.9,288.8L299.5,283.8L294.6,263.7L289.1,262L279.8,266.3L278.5,263.5L273.4,265.9L263.2,264.8L250.1,267.8L242.7,266.1L236.2,269.9Z"],["Tabapuã","3552601",742.8,441.6,"M714.3,476.7L718.9,477.3L737.7,467.2L741,471L750.7,470.9L751.3,475.2L760.8,475.7L764.6,484.7L774.2,491.3L775.5,489.1L771.7,487.4L767.8,478L770.1,459.3L766.9,453.6L768.3,444.9L771.1,439.9L781.6,433.5L786.6,433.9L786.4,421.1L775.1,422L772.2,416.4L759.7,408.4L753.9,408.3L750,411.1L746,409.4L730.6,394.4L725.8,394.4L721.2,410L717.9,412.3L721.1,420.1L716.2,424.4L713.6,434.7L716.2,439.1L715,443.2L706.6,454.5L708.4,461.3L711.3,463.7L707.9,465.9L714.3,476.7Z"],["Tanabi","3553401",529.8,281.2,"M454.2,318.9L463,312.4L471.1,310.6L476.9,319.6L482,322.8L487.3,322.8L491.5,327.5L500.6,327.5L522,347.9L524.2,355.1L530.6,357.3L541.7,352.1L541,346.5L546.6,340.8L546.5,337L554.1,328L554.3,323.8L561.8,319L564.1,288.3L576.4,274.2L573.4,269.2L567.8,271.2L565.9,260.9L561.1,257.9L556.7,243.4L558.8,238.3L555.5,236.2L555.9,232.5L552.6,232.1L548.3,218.3L540.8,215.2L533.4,207.4L517,222.6L510.7,234.5L509,242.5L510.5,248L504.1,260.3L501.9,270.2L496.1,276.3L490.6,277.8L488.8,288.6L481.3,285.5L467.4,287.5L459.7,294.4L454.1,303.5L449,304.4L442.8,309.9L445.7,314.4L454.2,318.9Z"],["Taquaritinga","3553708",906.6,634.9,"M941,663.9L946.2,648.5L954,642.4L958.5,632.7L967.8,625.4L968.5,618.6L978.7,619.8L991.4,613.5L1000,604.2L992,599.3L991,593L985.7,589.1L975.8,590.4L972.7,587.6L958.9,588.1L954.9,597.3L950.2,598.8L940.4,593.1L935.6,598.6L928.2,597.9L918.5,601.5L913.7,595.9L905.1,596.1L896.2,606.6L893.3,617L884.1,614.7L878.1,605.8L870.1,613L866.7,612.7L864.5,617L856.3,618.2L852.1,626.8L855.4,631.9L855.4,644.7L858.5,647.2L859.2,653.7L866.7,652L874.6,654L875.1,668.4L873.2,674.5L877.4,674.2L884.6,682.5L889.7,671.6L894.1,668L894,664.4L905.7,663.7L910.6,669.6L916.3,669.2L922.4,665.1L941,663.9Z"],["Três Fronteiras","3554904",79.7,189.3,"M62.9,202.9L65.9,197.8L82.1,205.6L89,229.6L91.4,231.8L98.3,230.6L102.5,227.2L102,224.3L105.3,219.2L101.1,210.6L89.8,203.9L89.5,199.2L97.2,192.2L96.2,186.6L102.6,174.9L102.4,166.9L87.4,147.5L83.7,155.7L69.3,172.1L68.8,179.6L61.6,192L62.9,202.9Z"],["Turmalina","3555307",229.5,110.6,"M211.4,95.3L213.4,105.1L211.5,114.2L224.6,142.8L235.5,144.8L248.5,139.3L255.7,123.7L243.9,103L241.9,93.2L244.1,88L241.7,77.9L225.3,95.7L219.3,98.3L211.4,95.3Z"],["Uchôa","3555604",690.5,440.8,"M662,450.4L666.6,451.3L674,459.6L677.8,460.7L679.4,467.5L688.8,476.2L692.1,486.1L699.8,494.8L702.2,496.2L707.1,493.2L707,483.4L711.9,481.1L714.3,476.7L707.9,465.9L711.3,463.7L708.4,461.3L706.6,454.5L715,443.2L716.2,439.1L713.6,434.7L716.2,424.4L721.1,420.1L717.9,412.3L721.2,410L721.5,404.2L725.5,399.2L725.8,394.4L720.7,394.2L721.8,387.6L718.7,383.4L719.8,385.2L715.9,391.1L703.3,399.3L700.7,408L688.2,418.7L689,425L683.8,430.6L665.3,434.7L663.5,445.2L666.6,448.7L662,450.4Z"],["União Paulista","3555703",437.4,423.0,"M455.9,408.7L450.2,406.2L448.1,409.7L438.7,407L429.8,412.6L411.9,430.7L416.8,441.6L421.2,440.7L424,435.9L438.6,434.6L449.9,440.3L454.6,427.9L453.4,415.5L455.9,408.7Z"],["Urupês","3556008",659.8,550.6,"M672.1,586.6L674.8,578.3L692.3,560.7L695.3,551.2L702,548.8L702.2,537.9L692.8,535.6L687.7,530.1L688,521.9L686.2,519.1L684.1,520.5L672.9,513.9L662.9,517.9L646.4,514.3L637.2,516.6L631.8,525.2L625.5,528.4L617.2,524.9L615.8,531L611.7,533.2L612.5,538.6L615.3,547.1L622.3,550L624.2,558.1L634.6,556.5L649.9,568.8L652.2,576.8L657.3,575.5L672.1,586.6Z"],["Urânia","3555802",165.1,159.7,"M159.4,107.6L155.5,111.5L155.9,117.6L145.9,129.6L144.9,133.8L152.9,144.2L153.5,165.4L157.4,168.7L156.6,181.1L150.7,187L146,188.2L134.2,203.6L146.1,203.1L149.1,204.5L152,210.3L156.9,211.5L164.5,206L164.9,202.7L168.5,203.8L170.2,196.2L175.2,192.8L175.3,189.3L183.3,189.4L185,184.4L177.7,161.3L180.9,146L178.1,134L174.9,127L160.4,117.5L161.4,111.3L159.4,107.6Z"],["Valentim Gentil","3556107",364.7,241.3,"M336.8,280.3L349.6,267.3L362.9,264L365.3,260.3L377.1,256.3L385,246.7L382.7,244.4L381.7,226.4L376.7,219.4L377.5,212.6L368.7,200.5L363.6,227.3L346.8,232L348,246L341.8,257.3L337.2,257L324,268.1L333.4,272.3L337.1,276.5L336.8,280.3Z"],["Vitória Brasil","3556958",221.5,160.1,"M235.5,144.8L224.6,142.8L210.5,151.4L212.2,166.5L222.6,177.9L230.1,174.7L235.3,166.1L229.8,158.9L232.2,148.6L235.5,144.8Z"],["Votuporanga","3557105",413.0,254.9,"M353.6,283.5L360.6,281.9L371.2,291.1L378.3,288.6L390.4,290.1L397.9,295.5L402.7,302.8L403.8,300.9L412.4,304.1L417.8,298.5L424,298.9L425.2,278.5L427.7,273.9L431.9,271.7L434,267L448.2,261.5L446.3,249L449.3,243.4L436.2,240.4L435,235.4L418.8,233.7L414.1,219.9L401.9,217L396.7,211.1L390.9,212.4L383.2,205.2L377.5,212.6L376.7,219.4L381.7,226.4L382.7,244.4L385,246.7L377.1,256.3L365.3,260.3L362.9,264L349.6,267.3L343.2,272.8L341.4,277.1L343,275.6L346.3,277.6L347.1,283.3L352.2,280.5L353.6,283.5Z"],["Zacarias","3557154",378.8,507.7,"M410.8,517.2L397.7,504.1L403.3,500L402,483.8L398,478.8L376.9,465.3L369.9,473L367.5,476.7L368.5,478.9L364.4,481.5L357.4,494.2L357.3,505.8L355.3,509.5L342.5,519.8L348.7,539.5L356.7,546.7L364,550.5L386.9,546.9L399.6,552.9L403,551.1L404.4,534.7L408,530.5L410.8,517.2Z"],["Álvares Florence","3501202",420.3,195.0,"M393.9,210.9L396.7,211.1L401.9,217L414.1,219.9L418.8,233.7L435,235.4L436.2,240.4L452.6,242.3L466.9,231.1L468.2,227.3L456.6,221L449.3,219L446.3,220.6L446,213L449,206.7L446.3,205L444.6,192.4L451.4,178.1L452.6,165.1L463.6,158.5L458.5,155.7L450.9,159.4L449,157.5L439.2,159.8L432.3,167L427.7,167.5L420.8,160.4L420,155.7L417.1,154.9L417.3,151L409.2,150L399.4,155.2L386.7,153.6L387.7,159.1L391.8,162.3L390.9,175.5L392.7,185.9L395.9,190.7L396.7,200.7L392.6,207.9L393.9,210.9Z"]]};
