// ============================================================
// [UL] Dashboard_Taxonomias — 01_Config.gs
// Menu, credenciais, gatilhos e popup de progresso
// ============================================================

const JIRA_BASE = 'https://storm-x.atlassian.net';
const PROJETO   = 'UL';
const CAMPOS    = 'summary,parent,status,issuetype,created,updated,assignee,reporter,duedate,subtasks,' +
                  'customfield_10184,customfield_10185,customfield_10186,customfield_10444,' +
                  'customfield_11545,customfield_10188,customfield_10807';

// ── HTML do popup de progresso ────────────────────────────────
const POPUP_HTML = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
html,body{
  width:640px;height:520px;overflow:hidden;
  background:#fff;color:#111827;
  font-family:'Inter','Segoe UI',system-ui,sans-serif;
  font-size:14px;line-height:1.5;
}
body{padding:28px 32px 24px;display:flex;flex-direction:column;gap:0;}
.logo{font-size:11px;font-weight:600;color:#6b7280;letter-spacing:.3px;margin-bottom:16px;}
.logo strong{color:#374151;}
h1{font-size:19px;font-weight:700;color:#111827;margin-bottom:4px;letter-spacing:-.3px;}
.sub{font-size:12px;color:#6b7280;margin-bottom:20px;line-height:1.5;}
.steps-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px 14px;margin-bottom:14px;}
.step{display:flex;align-items:flex-start;gap:10px;padding:12px 14px;border-radius:10px;background:#f9fafb;border:1.5px solid #e5e7eb;transition:border-color .25s,background .25s;}
.done.step{background:#f0fdf4;border-color:#86efac;}
.active.step{background:#eff6ff;border-color:#93c5fd;}
.si{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;flex-shrink:0;border:2px solid #d1d5db;background:#f9fafb;color:#9ca3af;transition:all .25s;}
.done .si{background:#dcfce7;border-color:#16a34a;color:#16a34a;font-size:15px;}
.active .si{background:#eff6ff;border-color:#3b82f6;animation:pulse 1.4s infinite;}
.active .si .spin{animation:spin .7s linear infinite;}
.sb{flex:1;min-width:0;}
.st{font-size:13px;font-weight:700;color:#9ca3af;transition:color .25s;}
.done .st{color:#16a34a;}.active .st{color:#111827;}
.sd{font-size:12px;color:#9ca3af;margin-top:3px;line-height:1.4;overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;}
.active .sd{color:#4b5563;}
.pw{background:#f3f4f6;border-radius:99px;height:6px;overflow:hidden;margin-bottom:4px;}
.pb{height:100%;background:linear-gradient(90deg,#3b82f6,#8b5cf6);border-radius:99px;transition:width .6s ease;width:0%;}
.pl{font-size:11px;color:#9ca3af;text-align:right;margin-bottom:14px;font-weight:500;}
.result{display:none;border-radius:12px;padding:16px 20px;text-align:center;}
.result.ok{display:flex;flex-direction:column;align-items:center;background:#f0fdf4;border:1.5px solid #86efac;}
.result.erro{display:flex;flex-direction:column;align-items:center;background:#fef2f2;border:1.5px solid #fca5a5;}
.result-title{font-size:15px;font-weight:700;margin-bottom:6px;}
.ok .result-title{color:#15803d;}.erro .result-title{color:#991b1b;}
.result-detail{font-size:12px;color:#4b5563;margin-bottom:14px;line-height:1.6;max-width:100%;word-break:break-word;}
.erro .result-detail{color:#7f1d1d;}
.btn-ok{background:#3b82f6;color:#fff;border:none;border-radius:8px;padding:9px 32px;font-size:13px;font-weight:700;cursor:pointer;transition:background .2s;}
.btn-ok:hover{background:#2563eb;}
.btn-ok.erro{background:#ef4444;}
.btn-ok.erro:hover{background:#dc2626;}
.spin{display:inline-block;width:13px;height:13px;border:2px solid #3b82f6;border-top-color:transparent;border-radius:50%;}
@keyframes spin{to{transform:rotate(360deg)}}
@keyframes pulse{0%,100%{box-shadow:0 0 0 0 rgba(59,130,246,.25)}50%{box-shadow:0 0 0 6px rgba(59,130,246,0)}}
</style></head><body>
<div class="logo">📋 <strong>[UL] Dashboard_Taxonomias</strong> &nbsp;·&nbsp; Unilever BR · Grasp x StormX</div>
<h1>Sincronizando dados</h1>
<p class="sub">Execuções encadeadas — cada etapa roda no próprio contexto, sem risco de timeout.</p>
<div class="steps-grid">
  <div class="step active" id="s1">
    <div class="si"><div class="spin"></div></div>
    <div class="sb">
      <div class="st">1 · Buscando no Jira</div>
      <div class="sd" id="d1">Conectando...</div>
    </div>
  </div>
  <div class="step" id="s2">
    <div class="si">2</div>
    <div class="sb">
      <div class="st">2 · Gravando RAW</div>
      <div class="sd" id="d2">Aguardando etapa 1...</div>
    </div>
  </div>
  <div class="step" id="s3">
    <div class="si">3</div>
    <div class="sb">
      <div class="st">3 · Calculando PAINEL</div>
      <div class="sd" id="d3">Aguardando etapa 2...</div>
    </div>
  </div>
  <div class="step" id="s4">
    <div class="si">4</div>
    <div class="sb">
      <div class="st">4 · Finalizando</div>
      <div class="sd" id="d4">Aguardando etapa 3...</div>
    </div>
  </div>
</div>
<div class="pw"><div class="pb" id="pb"></div></div>
<div class="pl" id="pl">Iniciando...</div>
<div class="result" id="result">
  <div class="result-title" id="rTitulo">—</div>
  <div class="result-detail" id="rDetalhe"></div>
  <button class="btn-ok" id="btnOk" onclick="google.script.host.close()">OK — Fechar</button>
</div>
<script>
var MAP = {
  buscando:   {idx:0, p:12},
  gravando:   {idx:1, p:42},
  calculando: {idx:2, p:72},
  formatando: {idx:3, p:90},
  concluido:  {idx:4, p:100}
};
var ultima = '', iv, erros = 0;

function setBar(p, l) {
  document.getElementById('pb').style.width = p + '%';
  document.getElementById('pl').textContent = l;
}
function markDone(n) {
  var e = document.getElementById('s' + n);
  e.className = 'step done';
  e.querySelector('.si').textContent = '✓';
}
function markActive(n, det) {
  var e = document.getElementById('s' + n);
  e.className = 'step active';
  e.querySelector('.si').innerHTML = '<div class="spin"></div>';
  var d = document.getElementById('d' + n);
  if (d && det) d.textContent = det;
}
function showResult(ok, titulo, detalhe) {
  clearInterval(iv);
  var r = document.getElementById('result');
  r.className = 'result ' + (ok ? 'ok' : 'erro');
  document.getElementById('rTitulo').textContent = titulo;
  document.getElementById('rDetalhe').textContent = detalhe || '';
  var btn = document.getElementById('btnOk');
  if (!ok) btn.className = 'btn-ok erro';
  btn.focus();
}

function tick() {
  google.script.run
    .withSuccessHandler(function(prog) {
      erros = 0;
      var et  = prog.etapa   || '';
      var det = prog.detalhe || '';
      if (et === ultima && et !== 'buscando' && et !== 'gravando' && et !== 'calculando' && et !== 'formatando') return;
      ultima = et;
      var c = MAP[et];
      if (!c) return;
      if (et === 'concluido') {
        for (var i = 1; i <= 4; i++) markDone(i);
        setBar(100, '100% — concluído');
        showResult(true, '✅ Sincronização concluída!', det);
        return;
      }
      if (et === 'erro') {
        setBar(100, '');
        showResult(false, '❌ Erro na sincronização', det || 'Verifique o log no Apps Script (Execuções) para mais detalhes.');
        return;
      }
      for (var j = 1; j <= c.idx; j++) markDone(j);
      markActive(c.idx + 1, det);
      setBar(c.p, c.p + '% concluído');
    })
    .withFailureHandler(function() {
      erros++;
      if (erros >= 5) {
        showResult(false, '❌ Sem resposta do servidor',
          'O Apps Script não está respondendo. Verifique as Execuções no editor.');
      }
    })
    .getProgresso();
}

// Evita pré-seleção do X do modal
document.addEventListener('DOMContentLoaded', function() { document.body.focus(); });

tick();
iv = setInterval(tick, 2000);
</script></body></html>`;

// ── Menu ──────────────────────────────────────────────────────
function onOpen() {
  try { _limparAbasExtrasInterno(); } catch(e) {}
  _verificarTokenUrgente();
  SpreadsheetApp.getUi()
    .createMenu('📋 Dashboard UL')
    .addItem('1. Configurar credenciais', 'configurarCredenciais')
    .addItem('2. Testar conexão',         'testarConexaoMenu')
    .addSeparator()
    .addItem('3. Sincronizar',            'abrirInstalador')
    .addItem('4. Atualizar TABELA',       'atualizarAbasMenu')
    .addSeparator()
    .addItem('Abrir painel',              'abrirPainel')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Avançado')
      .addItem('Sync incremental',        'sincronizarJira')
      .addItem('Recriar DE_PARA',         'inicializarDeParaMenu')
      .addItem('Reconfigurar acionadores','configurarAcionadores')
      .addSeparator()
      .addItem('Limpar abas extras',      'limparAbasExtras'))
    .addToUi();
}

// ── Acionadores ───────────────────────────────────────────────
function configurarAcionadores() {
  PropertiesService.getScriptProperties().setProperty('ACAO_STATUS', 'aguardando');
  PropertiesService.getScriptProperties().setProperty('ACAO_MSG', '');
  ScriptApp.newTrigger('_executarConfigurarAcionadores').timeBased().after(1000).create();
  const html = HtmlService.createHtmlOutput(_acaoPopupHtml('Reconfigurando acionadores...', 'acionadores'))
    .setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Reconfigurar acionadores');
}
function _executarConfigurarAcionadores() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '_executarConfigurarAcionadores') ScriptApp.deleteTrigger(t);
  });
  try {
    _configurarAcionadoresInterno();
    PropertiesService.getScriptProperties().setProperties({
      ACAO_STATUS: 'ok',
      ACAO_MSG: 'Acionadores reconfigurados. Apenas sincronizarJira (a cada 1h) está ativo.'
    });
  } catch(e) {
    PropertiesService.getScriptProperties().setProperties({ ACAO_STATUS: 'erro', ACAO_MSG: e.message });
  }
}

function _agendarGatilho() {
  _configurarAcionadoresInterno();
}

function _configurarAcionadoresInterno() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (['sincronizarJira'].includes(fn) || fn.startsWith('_executar') || fn === '_sincronizarEtapa2' || fn === '_etapa2GravarPainel' || fn === '_etapa2GravarTabela' || fn === '_etapa2Finalizar' || fn === '_gravarRAWDoBusfer' || fn === '_buscarProximaPagina') return;
    ScriptApp.deleteTrigger(t);
  });
  const jaExiste = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'sincronizarJira' && t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK);
  if (!jaExiste) ScriptApp.newTrigger('sincronizarJira').timeBased().everyHours(1).create();
  Logger.log('Acionadores reconfigurados: sincronizarJira a cada 1h.');
}

// ── Popup genérico para ações rápidas ────────────────────────
function getStatusAcao() {
  const props = PropertiesService.getScriptProperties();
  return {
    status: props.getProperty('ACAO_STATUS') || 'aguardando',
    msg:    props.getProperty('ACAO_MSG')    || ''
  };
}

function _acaoPopupHtml(msg, tipo) {
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">' +
  '<style>' +
  '*{box-sizing:border-box;margin:0;padding:0;}' +
  'body{background:#fff;color:#111827;font-family:Inter,"Segoe UI",system-ui,sans-serif;font-size:13px;padding:24px;overflow:hidden;}' +
  '.logo{font-size:10px;font-weight:600;color:#6b7280;margin-bottom:16px;}' +
  '.logo strong{color:#374151;}' +
  '.loading{text-align:center;padding:16px 0;}' +
  '.spin{display:inline-block;width:28px;height:28px;border:3px solid #3b82f6;border-top-color:transparent;border-radius:50%;animation:spin .8s linear infinite;margin-bottom:12px;}' +
  '@keyframes spin{to{transform:rotate(360deg)}}' +
  '.loading p{font-size:13px;color:#6b7280;}' +
  '.result{display:none;border-radius:10px;padding:14px 16px;text-align:center;}' +
  '.result.ok{background:#f0fdf4;border:1.5px solid #86efac;}' +
  '.result.erro{background:#fef2f2;border:1.5px solid #fca5a5;}' +
  '.result h2{font-size:14px;font-weight:700;margin-bottom:6px;}' +
  '.result.ok h2{color:#15803d;}.result.erro h2{color:#991b1b;}' +
  '.result p{font-size:12px;color:#4b5563;line-height:1.5;margin-bottom:12px;}' +
  '.btn{background:#3b82f6;color:#fff;border:none;border-radius:6px;padding:7px 20px;font-size:12px;font-weight:700;cursor:pointer;}' +
  '.btn.erro{background:#ef4444;}' +
  '</style></head><body>' +
  '<div class="logo">📋 <strong>[UL] Dashboard_Taxonomias</strong> · Unilever BR · Grasp x StormX</div>' +
  '<div class="loading" id="loading"><div class="spin"></div><p>' + msg + '</p></div>' +
  '<div class="result" id="result"><h2 id="rTitulo">—</h2><p id="rMsg">—</p>' +
  '<button class="btn" id="btnF" onclick="google.script.host.close()">Fechar</button></div>' +
  '<script>' +
  'var iv=setInterval(function(){' +
  'google.script.run.withSuccessHandler(function(s){' +
  'if(s.status==="ok"||s.status==="erro"){' +
  'clearInterval(iv);' +
  'document.getElementById("loading").style.display="none";' +
  'var r=document.getElementById("result");r.style.display="block";' +
  'r.className="result "+(s.status==="ok"?"ok":"erro");' +
  'document.getElementById("rTitulo").textContent=s.status==="ok"?"✅ Concluído":"❌ Erro";' +
  'document.getElementById("rMsg").textContent=s.msg;' +
  'if(s.status==="erro")document.getElementById("btnF").className="btn erro";' +
  '}}).withFailureHandler(function(){}).getStatusAcao();' +
  '},1500);' +
  '<\/script></body></html>';
}

// atualizarAbas com popup
function atualizarAbasMenu() {
  PropertiesService.getScriptProperties().setProperty('ACAO_STATUS', 'aguardando');
  PropertiesService.getScriptProperties().setProperty('ACAO_MSG', '');
  ScriptApp.newTrigger('_executarAtualizarAbas').timeBased().after(1000).create();
  const html = HtmlService.createHtmlOutput(_acaoPopupHtml('Atualizando TABELA...', 'abas'))
    .setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Atualizar TABELA');
}
function _executarAtualizarAbas() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '_executarAtualizarAbas') ScriptApp.deleteTrigger(t);
  });
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    _gravarTabela(ss);
    PropertiesService.getScriptProperties().setProperties({
      ACAO_STATUS: 'ok',
      ACAO_MSG: 'Aba TABELA atualizada com sucesso.'
    });
  } catch(e) {
    PropertiesService.getScriptProperties().setProperties({ ACAO_STATUS: 'erro', ACAO_MSG: String(e) });
  }
}

// inicializarDeParaForcar com popup
function inicializarDeParaMenu() {
  PropertiesService.getScriptProperties().setProperty('ACAO_STATUS', 'aguardando');
  PropertiesService.getScriptProperties().setProperty('ACAO_MSG', '');
  ScriptApp.newTrigger('_executarRecriarDePara').timeBased().after(1000).create();
  const html = HtmlService.createHtmlOutput(_acaoPopupHtml('Recriando aba DE_PARA...', 'depara'))
    .setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Recriar DE_PARA');
}
function _executarRecriarDePara() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '_executarRecriarDePara') ScriptApp.deleteTrigger(t);
  });
  try {
    inicializarDeParaForcar();
    PropertiesService.getScriptProperties().setProperties({
      ACAO_STATUS: 'ok',
      ACAO_MSG: 'Aba DE_PARA recriada com todos os mapeamentos padrão.'
    });
  } catch(e) {
    PropertiesService.getScriptProperties().setProperties({ ACAO_STATUS: 'erro', ACAO_MSG: String(e) });
  }
}

// ── Popup de sincronização ────────────────────────────────────
function abrirInstalador() {
  _setProgresso('aguardando', 'Iniciando em instantes...');
  ScriptApp.newTrigger('sincronizarCompleto').timeBased().after(2000).create();
  const html = HtmlService.createHtmlOutput(POPUP_HTML).setWidth(640).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Sincronizando Dashboard UL');
}

function abrirPainel() {
  const url = 'https://falssp.github.io/ul-dashboard-taxonomias/';
  const html = HtmlService.createHtmlOutput(
    '<style>*{box-sizing:border-box;margin:0;padding:0;font-family:Inter,sans-serif;}' +
    'body{background:#0f1117;display:flex;align-items:center;justify-content:center;min-height:100vh;}' +
    '.card{background:#181c27;border:1px solid #2a3050;border-radius:12px;padding:28px 32px;text-align:center;width:360px;}' +
    'p{color:#a0aabf;font-size:13px;margin-bottom:18px;line-height:1.5;}' +
    'a{display:inline-block;background:#4f7ef8;color:#fff;text-decoration:none;border-radius:7px;' +
    'padding:10px 28px;font-size:13px;font-weight:700;}' +
    'a:hover{opacity:.85;}</style>' +
    '<div class="card">' +
    '<p>Clique para abrir o painel em uma nova aba:</p>' +
    '<a href="' + url + '" target="_blank" onclick="setTimeout(function(){google.script.host.close();},300)">Abrir painel ↗</a>' +
    '</div>'
  ).setWidth(400).setHeight(160);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Dashboard UL');
}

// ── Credenciais ───────────────────────────────────────────────
const SETUP_HTML = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{background:#fff;color:#111827;font-family:"Inter","Segoe UI",system-ui,sans-serif;font-size:13px;padding:22px 26px;overflow:hidden;}
.logo{font-size:10px;font-weight:600;color:#6b7280;margin-bottom:12px;}
.logo strong{color:#374151;}
h1{font-size:17px;font-weight:700;color:#111827;margin-bottom:4px;}
.sub{font-size:12px;color:#6b7280;margin-bottom:16px;line-height:1.4;}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:0 16px;}
.field{margin-bottom:10px;}
.field.full{grid-column:1/-1;}
label{display:block;font-size:10px;font-weight:600;color:#374151;text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px;}
input{width:100%;background:#f9fafb;border:1.5px solid #d1d5db;color:#111827;border-radius:7px;padding:8px 10px;font-size:13px;outline:none;transition:border-color .2s;font-family:inherit;}
input:focus{border-color:#3b82f6;background:#fff;}
.hint{font-size:10px;color:#9ca3af;margin-top:3px;}
.btn{width:100%;background:#3b82f6;color:#fff;border:none;border-radius:7px;padding:10px;font-size:13px;font-weight:700;cursor:pointer;margin-top:6px;transition:background .2s;}
.btn:hover{background:#2563eb;}
.btn:disabled{background:#93c5fd;cursor:default;}
.loading{display:none;margin-top:16px;text-align:center;color:#6b7280;font-size:13px;}
.spin{display:inline-block;width:16px;height:16px;border:2px solid #3b82f6;border-top-color:transparent;border-radius:50%;animation:spin .7s linear infinite;margin-right:8px;vertical-align:middle;}
@keyframes spin{to{transform:rotate(360deg)}}
.result{display:none;margin-top:14px;border-radius:8px;padding:14px 16px;text-align:center;}
.result.ok{background:#f0fdf4;border:1.5px solid #86efac;}
.result.erro{background:#fef2f2;border:1.5px solid #fca5a5;}
.result h2{font-size:14px;font-weight:700;margin-bottom:5px;}
.result.ok h2{color:#15803d;}.result.erro h2{color:#991b1b;}
.result p{font-size:12px;color:#4b5563;line-height:1.5;margin-bottom:10px;}
.btn-fechar{background:#3b82f6;color:#fff;border:none;border-radius:6px;padding:7px 22px;font-size:12px;font-weight:700;cursor:pointer;}
.btn-fechar.erro{background:#ef4444;}
</style></head><body>
<div class="logo">📋 <strong>[UL] Dashboard_Taxonomias</strong> · Unilever BR · Grasp x StormX</div>
<h1>Configurar credenciais</h1>
<p class="sub">E-mail Atlassian + API Token do Jira. Salvo de forma segura no Apps Script.</p>
<div id="form">
  <div class="grid">
    <div class="field full">
      <label>E-mail Atlassian</label>
      <input type="email" id="email" placeholder="seu@email.com" autocomplete="off">
    </div>
    <div class="field full">
      <label>API Token</label>
      <input type="password" id="token" placeholder="Cole o token aqui..." autocomplete="off">
      <div class="hint">id.atlassian.com → Segurança → Tokens de API</div>
    </div>
    <div class="field full">
      <label>Validade do token</label>
      <input type="date" id="expira" autocomplete="off">
      <div class="hint">Confirme a data de expiração — normalmente hoje + 1 ano</div>
    </div>
  </div>
  <button class="btn" id="btnSalvar" onclick="salvar()">Salvar e testar conexão</button>
</div>
<div class="loading" id="loading">
  <div class="spin"></div>Testando conexão e salvando...
</div>
<div class="result" id="result">
  <h2 id="rTitulo">—</h2>
  <p id="rMsg">—</p>
  <button class="btn-fechar" id="btnFechar" onclick="fechar()">Fechar</button>
</div>
<script>
document.addEventListener('DOMContentLoaded', function() {
  var d = new Date(); d.setDate(d.getDate() + 365);
  document.getElementById('expira').value = d.toISOString().substring(0,10);
  document.getElementById('email').addEventListener('keydown', function(e){ if(e.key==='Enter') document.getElementById('token').focus(); });
  document.getElementById('token').addEventListener('keydown', function(e){ if(e.key==='Enter') document.getElementById('expira').focus(); });
  document.getElementById('expira').addEventListener('keydown', function(e){ if(e.key==='Enter') salvar(); });
});
var iv;
function salvar() {
  var email  = document.getElementById('email').value.trim();
  var token  = document.getElementById('token').value.trim();
  var expira = document.getElementById('expira').value.trim();
  if (!email)  { alert('Informe o e-mail.'); return; }
  if (!token)  { alert('Informe o token.'); return; }
  if (!expira) { alert('Informe a validade do token.'); return; }
  document.getElementById('form').style.display = 'none';
  document.getElementById('loading').style.display = 'block';
  google.script.run
    .withSuccessHandler(function() { iv = setInterval(verificar, 1500); })
    .withFailureHandler(function(e) { mostrarErro('Erro ao iniciar: ' + e.message); })
    .agendarSalvarCredenciais(email, token, expira);
}
function verificar() {
  google.script.run
    .withSuccessHandler(function(s) {
      if (s.status === 'ok')        { clearInterval(iv); mostrarOk(s.msg); }
      else if (s.status === 'erro') { clearInterval(iv); mostrarErro(s.msg); }
    })
    .withFailureHandler(function(){})
    .getStatusCredenciais();
}
function mostrarOk(msg) {
  document.getElementById('loading').style.display = 'none';
  var res = document.getElementById('result');
  res.className = 'result ok'; res.style.display = 'block';
  document.getElementById('rTitulo').textContent = '✅ Credenciais salvas!';
  document.getElementById('rMsg').textContent = msg;
}
function mostrarErro(msg) {
  document.getElementById('loading').style.display = 'none';
  document.getElementById('form').style.display = 'block';
  var res = document.getElementById('result');
  res.className = 'result erro'; res.style.display = 'block';
  document.getElementById('rTitulo').textContent = '❌ Erro na conexão';
  document.getElementById('rMsg').textContent = msg;
  document.getElementById('btnFechar').className = 'btn-fechar erro';
}
function fechar() { google.script.host.close(); }
</script></body></html>`;

function configurarCredenciais() {
  PropertiesService.getScriptProperties().deleteProperty('CRED_STATUS');
  PropertiesService.getScriptProperties().deleteProperty('CRED_MSG');
  const html = HtmlService.createHtmlOutput(SETUP_HTML).setWidth(480).setHeight(470);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Configurar Dashboard UL');
}

function agendarSalvarCredenciais(email, token, expira) {
  PropertiesService.getScriptProperties().setProperties({
    CRED_EMAIL_TEMP:  email,
    CRED_TOKEN_TEMP:  token,
    CRED_EXPIRA_TEMP: expira || '',
    CRED_STATUS: 'aguardando',
    CRED_MSG: ''
  });
  ScriptApp.newTrigger('_executarSalvarCredenciais').timeBased().after(1000).create();
}

function _executarSalvarCredenciais() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '_executarSalvarCredenciais') ScriptApp.deleteTrigger(t);
  });
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('CRED_EMAIL_TEMP') || '';
  const token = props.getProperty('CRED_TOKEN_TEMP') || '';
  props.deleteProperty('CRED_EMAIL_TEMP');
  props.deleteProperty('CRED_TOKEN_TEMP');

  if (!email || !token) {
    props.setProperties({ CRED_STATUS: 'erro', CRED_MSG: 'Credenciais não encontradas.' });
    return;
  }
  const teste = _testarConexao(email, token);
  if (!teste.ok) {
    props.setProperties({ CRED_STATUS: 'erro', CRED_MSG: 'Conexão falhou: ' + teste.msg });
    return;
  }
  PropertiesService.getUserProperties().setProperties({ JIRA_EMAIL: email, JIRA_API_TOKEN: token });
  inicializarDePara();
  _agendarGatilho();
  try { props.setProperty('WEBAPP_URL', ScriptApp.getService().getUrl()); } catch(e) {}
  const expiraStr = props.getProperty('CRED_EXPIRA_TEMP') || '';
  props.deleteProperty('CRED_EXPIRA_TEMP');
  const expiraFinal = expiraStr || (() => {
    const d = new Date(); d.setDate(d.getDate() + 365);
    return d.toISOString().substring(0, 10);
  })();
  props.setProperty('TOKEN_EXPIRA', expiraFinal);
  props.deleteProperty('TOKEN_AVISO_ENVIADO');
  const expiraFmt = expiraFinal.split('-').reverse().join('/');
  props.setProperties({
    CRED_STATUS: 'ok',
    CRED_MSG: teste.total + ' issues no Jira. Token válido até ' + expiraFmt + '. Sincronização iniciada — abra o painel em ~10 minutos.'
  });
  ScriptApp.newTrigger('sincronizarCompleto').timeBased().after(2000).create();
}

function getStatusCredenciais() {
  const props = PropertiesService.getScriptProperties();
  return {
    status: props.getProperty('CRED_STATUS') || 'aguardando',
    msg:    props.getProperty('CRED_MSG')    || ''
  };
}

// ── Aviso de expiração do token ──────────────────────────────
function _diasParaExpirar() {
  const exp = PropertiesService.getScriptProperties().getProperty('TOKEN_EXPIRA');
  if (!exp) return null;
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  const vence = new Date(exp); vence.setHours(0,0,0,0);
  return Math.round((vence - hoje) / 86400000);
}

function _verificarTokenUrgente() {
  try {
    const dias = _diasParaExpirar();
    if (dias === null) return;
    if (dias <= 0) {
      SpreadsheetApp.getUi().alert(
        '🔴 Token do Jira EXPIRADO!\n\n' +
        'O sync automático está falhando.\n' +
        'Vá em Menu → 1. Configurar credenciais e gere um novo token em id.atlassian.com'
      );
    } else if (dias <= 3) {
      SpreadsheetApp.getUi().alert(
        '⚠️ Token do Jira vence em ' + dias + ' dia' + (dias === 1 ? '' : 's') + '!\n\n' +
        'Renove em: id.atlassian.com → Segurança → Tokens de API\n' +
        'Depois: Menu → 1. Configurar credenciais'
      );
    }
  } catch(e) {}
}

function verificarAvisoToken() {
  try {
    const dias = _diasParaExpirar();
    if (dias === null || dias > 10) return;

    const props = PropertiesService.getScriptProperties();
    const hoje  = new Date().toISOString().substring(0, 10);

    props.setProperty('TOKEN_AVISO_DIAS', String(dias));

    const ultimoEnvio = props.getProperty('TOKEN_AVISO_ENVIADO') || '';
    if (ultimoEnvio === hoje) return;

    const { email } = _creds();
    const assunto = dias <= 0
      ? '[URGENTE] Token do Jira EXPIRADO — Dashboard UL parou de sincronizar'
      : '[Aviso] Token do Jira vence em ' + dias + ' dia' + (dias === 1 ? '' : 's') + ' — Dashboard UL';

    const corpo = dias <= 0
      ? 'O token de API do Jira usado pelo Dashboard_Taxonomias UL <b>expirou</b>.<br><br>' +
        'O sync automático está falhando desde a expiração.<br><br>' +
        '<b>Para corrigir:</b><br>' +
        '1. Acesse <a href="https://id.atlassian.com">id.atlassian.com</a> → Segurança → Tokens de API<br>' +
        '2. Crie um novo token (validade máxima: 1 ano)<br>' +
        '3. Na planilha: Menu 📋 Dashboard UL → 1. Configurar credenciais<br>' +
        '4. Insira o e-mail e o novo token'
      : 'O token de API do Jira usado pelo Dashboard_Taxonomias UL vence em <b>' + dias + ' dia' + (dias === 1 ? '' : 's') + '</b>.<br><br>' +
        'Renove antes que o sync automático pare de funcionar.<br><br>' +
        '<b>Como renovar:</b><br>' +
        '1. Acesse <a href="https://id.atlassian.com">id.atlassian.com</a> → Segurança → Tokens de API<br>' +
        '2. Crie um novo token (validade máxima: 1 ano)<br>' +
        '3. Na planilha: Menu 📋 Dashboard UL → 1. Configurar credenciais<br>' +
        '4. Insira o e-mail e o novo token';

    GmailApp.sendEmail(email, assunto, '', { htmlBody: corpo, name: '[UL] Dashboard_Taxonomias' });
    props.setProperty('TOKEN_AVISO_ENVIADO', hoje);
    Logger.log('Aviso de token enviado para ' + email + ' — ' + dias + ' dias restantes.');
  } catch(e) {
    Logger.log('Erro ao verificar aviso de token: ' + e);
  }
}

function getAvisoToken() {
  const props = PropertiesService.getScriptProperties();
  const dias = props.getProperty('TOKEN_AVISO_DIAS');
  if (!dias) return null;
  const d = parseInt(dias, 10);
  if (d > 15) return null;
  return { dias: d, expira: props.getProperty('TOKEN_EXPIRA') || '' };
}

// ── Limpeza de abas desnecessárias ───────────────────────────
const ABAS_PERMITIDAS = [
  '🏠 INICIO', 'PAINEL', '📋 TABELA',
  'RAW_PAI', 'RAW_FILHO', '⚠️ INCORRETOS', 'DE_PARA', '_SYNC_BUFFER'
];

function _limparAbasExtrasInterno() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const permitidas = new Set([
    '🏠 INICIO','PAINEL','📋 TABELA',
    'RAW_PAI','RAW_FILHO','⚠️ INCORRETOS','DE_PARA','_SYNC_BUFFER'
  ]);
  ss.getSheets().forEach(function(sh) {
    if (!permitidas.has(sh.getName()) && ss.getSheets().length > 1) {
      try { ss.deleteSheet(sh); Logger.log('Aba extra removida: ' + sh.getName()); }
      catch(e) {}
    }
  });
}

function limparAbasExtras() {
  PropertiesService.getScriptProperties().setProperty('ACAO_STATUS', 'aguardando');
  PropertiesService.getScriptProperties().setProperty('ACAO_MSG', '');
  ScriptApp.newTrigger('_executarLimparAbas').timeBased().after(1000).create();
  const html = HtmlService.createHtmlOutput(_acaoPopupHtml('Verificando e removendo abas desnecessárias...', 'limpeza'))
    .setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Limpar abas');
}

function _executarLimparAbas() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '_executarLimparAbas') ScriptApp.deleteTrigger(t);
  });
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const removidas = [];
    ss.getSheets().forEach(sh => {
      const nome = sh.getName();
      if (!ABAS_PERMITIDAS.includes(nome)) {
        if (ss.getSheets().length > 1) {
          ss.deleteSheet(sh);
          removidas.push(nome);
        }
      }
    });
    const msg = removidas.length > 0
      ? 'Removidas: ' + removidas.join(', ')
      : 'Nenhuma aba extra encontrada — planilha já está limpa.';
    PropertiesService.getScriptProperties().setProperties({ ACAO_STATUS: 'ok', ACAO_MSG: msg });
  } catch(e) {
    PropertiesService.getScriptProperties().setProperties({ ACAO_STATUS: 'erro', ACAO_MSG: String(e) });
  }
}

function testarConexaoMenu() {
  PropertiesService.getScriptProperties().setProperty('ACAO_STATUS', 'aguardando');
  PropertiesService.getScriptProperties().setProperty('ACAO_MSG', '');
  ScriptApp.newTrigger('_executarTestarConexao').timeBased().after(1000).create();
  const html = HtmlService.createHtmlOutput(_acaoPopupHtml('Testando conexão com o Jira...', 'testar'))
    .setWidth(400).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Testar conexão');
}
function _executarTestarConexao() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '_executarTestarConexao') ScriptApp.deleteTrigger(t);
  });
  try {
    const { email, token } = _creds();
    const r = _testarConexao(email, token);
    const props = PropertiesService.getScriptProperties();
    if (r.ok) {
      props.setProperties({ ACAO_STATUS: 'ok', ACAO_MSG: r.total + ' issues encontradas no projeto UL. Conexão funcionando normalmente.' });
    } else {
      props.setProperties({ ACAO_STATUS: 'erro', ACAO_MSG: r.msg });
    }
  } catch(e) {
    PropertiesService.getScriptProperties().setProperties({ ACAO_STATUS: 'erro', ACAO_MSG: e.message });
  }
}

function _creds() {
  const p = PropertiesService.getUserProperties();
  const email = p.getProperty('JIRA_EMAIL'), token = p.getProperty('JIRA_API_TOKEN');
  if (!email || !token) throw new Error('Configure as credenciais: Menu > 1. Configurar credenciais');
  return { email, token };
}

function _testarConexao(email, token) {
  const auth = Utilities.base64Encode(email + ':' + token);
  const respMe = UrlFetchApp.fetch(
    JIRA_BASE + '/rest/api/3/myself',
    { method: 'GET', headers: { Authorization: 'Basic ' + auth, Accept: 'application/json' }, muteHttpExceptions: true }
  );
  if (respMe.getResponseCode() !== 200) {
    const code = respMe.getResponseCode();
    const msg = code === 401 ? 'Token inválido ou expirado (401)' :
                code === 403 ? 'Sem permissão de acesso (403)' :
                'Erro HTTP ' + code;
    return { ok: false, msg: msg };
  }
  const syncTotal = PropertiesService.getScriptProperties().getProperty('SYNC_TOTAL');
  if (syncTotal) return { ok: true, total: syncTotal };
  const respSearch = UrlFetchApp.fetch(
    JIRA_BASE + '/rest/api/3/search/jql?jql=' + encodeURIComponent('project=' + PROJETO) + '&maxResults=1&fields=summary',
    { method: 'GET', headers: { Authorization: 'Basic ' + auth, Accept: 'application/json' }, muteHttpExceptions: true }
  );
  if (respSearch.getResponseCode() !== 200) return { ok: false, msg: 'Erro ao acessar projeto: HTTP ' + respSearch.getResponseCode() };
  return { ok: true, total: 'acesso confirmado' };
}

function salvarCredenciais(email, token) {
  PropertiesService.getUserProperties().setProperties({ JIRA_EMAIL: email, JIRA_API_TOKEN: token });
  const r = _testarConexao(email, token);
  if (!r.ok) return { ok: false, msg: r.msg };
  inicializarDePara();
  _agendarGatilho();
  try { PropertiesService.getScriptProperties().setProperty('WEBAPP_URL', ScriptApp.getService().getUrl()); } catch(e) {}
  ScriptApp.newTrigger('sincronizarCompleto').timeBased().after(5 * 1000).create();
  return { ok: true, msg: r.total + ' issues no Jira. Sincronização iniciada automaticamente — aguarde ~10 minutos e abra o painel.' };
}

function getCredentials() {
  const u = PropertiesService.getUserProperties();
  return { email: u.getProperty('JIRA_EMAIL') || '', token: u.getProperty('JIRA_API_TOKEN') || '' };
}
