// ============================================================
// [UL] Dashboard_Taxonomias — 02_Sync.gs
// Deploy automático via clasp
// Busca Jira paginada em execuções encadeadas (escala ilimitada)
// ============================================================

const TIPOS_VALIDOS = ['Tarefa', 'Subtarefa'];

const BUFFER_ABA = '_SYNC_BUFFER';

const PROP_NEXT_TOKEN  = 'SYNC_NEXT_TOKEN';
const PROP_PAGINA      = 'SYNC_PAGINA';
const PROP_ACUMULADO   = 'SYNC_ACUMULADO';
const PROP_JQL_EXTRA   = 'SYNC_JQL_EXTRA';
const PROP_SYNC_TOTAL  = 'SYNC_TOTAL';
const PROP_ULTIMA_SYNC = 'ULTIMA_SYNC';

// ── Progresso (lido pelo popup a cada 2s) ────────────────────
function _setProgresso(etapa, detalhe) {
  PropertiesService.getScriptProperties().setProperties({
    PROG_ETAPA:   etapa,
    PROG_DETALHE: detalhe || '',
    PROG_TS:      new Date().toISOString()
  });
}

function getProgresso() {
  const p = PropertiesService.getScriptProperties();
  return {
    etapa:   p.getProperty('PROG_ETAPA')   || '',
    detalhe: p.getProperty('PROG_DETALHE') || '',
    ts:      p.getProperty('PROG_TS')      || ''
  };
}

// ── Buffer ────────────────────────────────────────────────────
function _bufferCriar(ss) {
  const old = ss.getSheetByName(BUFFER_ABA);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(BUFFER_ABA);
  sh.hideSheet();
  sh.getRange(1, 1).setValue('JSON_ISSUE');
  return sh;
}

function _bufferAppend(ss, issues) {
  if (!issues.length) return;
  let sh = ss.getSheetByName(BUFFER_ABA);
  if (!sh) sh = _bufferCriar(ss);
  const rows = issues.map(i => {
    const f = i.fields || {};
    const slim = {
      key: i.key,
      fields: {
        summary:           f.summary           || '',
        issuetype:         f.issuetype         ? {name: f.issuetype.name || ''}  : null,
        status:            f.status            ? {name: f.status.name    || ''}  : null,
        assignee:          f.assignee          ? {displayName: f.assignee.displayName || ''} : null,
        reporter:          f.reporter          ? {displayName: f.reporter.displayName || ''} : null,
        parent:            f.parent            ? {key: f.parent.key || '',
                             fields: {summary: (f.parent.fields || {}).summary || ''}} : null,
        created:           f.created           || '',
        updated:           f.updated           || '',
        duedate:           f.duedate           || '',
        customfield_10184: f.customfield_10184 != null ? f.customfield_10184 : '',
        customfield_10185: f.customfield_10185 ? {value: f.customfield_10185.value || ''} : null,
        customfield_10186: f.customfield_10186 ? {value: f.customfield_10186.value || ''} : null,
        customfield_10444: f.customfield_10444 ? {value: f.customfield_10444.value || ''} : null,
        customfield_10188: f.customfield_10188 ? {value: f.customfield_10188.value || ''} : null,
        customfield_10807: f.customfield_10807 ? {value: f.customfield_10807.value || ''} : null,
        customfield_11545: f.customfield_11545 || null
      }
    };
    return [JSON.stringify(slim)];
  });
  const last = sh.getLastRow();
  sh.getRange(last + 1, 1, rows.length, 1).setValues(rows);
}

function _bufferLer(ss) {
  const sh = ss.getSheetByName(BUFFER_ABA);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(r => { try { return JSON.parse(r[0]); } catch(e) { return null; } })
    .filter(Boolean);
}

function _bufferDeletar(ss) {
  const sh = ss.getSheetByName(BUFFER_ABA);
  if (sh) try { ss.deleteSheet(sh); } catch(e) {}
}

function _cacheSet(key, value) {
  try { CacheService.getScriptCache().put(key, value, 21600); } catch(e) {}
}
function _cacheGet(key) {
  try { return CacheService.getScriptCache().get(key); } catch(e) { return null; }
}

// ── Limpeza de triggers ───────────────────────────────────────
function _limparTriggersBusca() {
  const LIMPAR = [
    '_buscarProximaPagina','_gravarRAWDoBusfer',
    '_sincronizarEtapa2','_etapa2GravarPainel','_etapa2GravarTabela','_etapa2Finalizar',
    'sincronizarCompleto'
  ];
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (!LIMPAR.includes(fn)) return;
    ScriptApp.deleteTrigger(t);
  });
}

function _limparEstado(ss) {
  _limparTriggersBusca();
  _bufferDeletar(ss);
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_NEXT_TOKEN);
  props.deleteProperty(PROP_PAGINA);
  props.deleteProperty(PROP_ACUMULADO);
  props.deleteProperty(PROP_JQL_EXTRA);
  _cacheSet('SYNC_NEXT_TOKEN', '');
}

// ── SYNC COMPLETO ─────────────────────────────────────────────
function sincronizarCompleto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _limparEstado(ss);
  // Apaga ULTIMA_SYNC para que o incremental de 1h não interfira durante o completo
  PropertiesService.getScriptProperties().deleteProperty(PROP_ULTIMA_SYNC);
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_JQL_EXTRA, '');
  props.setProperty(PROP_PAGINA, '0');
  props.setProperty(PROP_ACUMULADO, '0');
  _bufferCriar(ss);
  _setProgresso('buscando', 'Conectando ao Jira...');
  Logger.log('Sync completo — iniciando busca paginada encadeada...');
  _executarPaginaBusca(ss, '', null, 0, 0);
}

// ── SYNC INCREMENTAL (1h) ─────────────────────────────────────
function sincronizarJira() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _limparEstado(ss);

  try { verificarAvisoToken(); } catch(e) { Logger.log('Aviso token: ' + e); }

  if (!PropertiesService.getScriptProperties().getProperty(PROP_ULTIMA_SYNC)) {
    sincronizarCompleto(); return;
  }
  const raw = ss.getSheetByName('RAW_FILHO');
  if (!raw || raw.getLastRow() < 2) { sincronizarCompleto(); return; }

  const jqlExtra = 'updated >= "-1h"';
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_JQL_EXTRA, jqlExtra);
  props.setProperty(PROP_PAGINA, '0');
  props.setProperty(PROP_ACUMULADO, '0');
  _bufferCriar(ss);
  _setProgresso('buscando', 'Sync incremental — conectando...');
  Logger.log('Sync incremental (1h) — iniciando busca...');
  _executarPaginaBusca(ss, jqlExtra, null, 0, 0);
}

// ── NÚCLEO DA BUSCA ───────────────────────────────────────────
function _executarPaginaBusca(ss, jqlExtra, nextToken, paginaBase, acumuladoBase) {
  const { email, token } = _creds();
  const auth  = Utilities.base64Encode(email + ':' + token);
  const jql   = 'project=' + PROJETO + (jqlExtra ? ' AND ' + jqlExtra : '') + ' ORDER BY key ASC';
  const LIMITE_MS = 4.5 * 60 * 1000;
  const t0 = Date.now();

  let pagina     = paginaBase;
  let acumulado  = acumuladoBase;
  let token_next = nextToken;
  let isLast     = false;
  const lote     = [];

  while (true) {
    if (Date.now() - t0 > LIMITE_MS) {
      Logger.log('Limite de tempo atingido — acumulado=' + acumulado + ', agendando continuação...');
      break;
    }
    let url = JIRA_BASE + '/rest/api/3/search/jql?jql=' + encodeURIComponent(jql) +
              '&maxResults=100&fields=' + CAMPOS;
    if (token_next) url += '&nextPageToken=' + encodeURIComponent(token_next);

    const resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { Authorization: 'Basic ' + auth, Accept: 'application/json' },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      Logger.log('Erro HTTP ' + resp.getResponseCode());
      _setProgresso('erro', 'Erro HTTP ' + resp.getResponseCode() + ' ao buscar no Jira.');
      _limparEstado(ss);
      return;
    }

    const data   = JSON.parse(resp.getContentText());
    const issues = data.issues || [];
    if (!issues.length) { isLast = true; break; }

    lote.push(...issues);
    pagina++;
    acumulado += issues.length;

    Logger.log('Página ' + pagina + ' | issues=' + issues.length + ' | acumulado=' + acumulado);
    _setProgresso('buscando', 'Página ' + pagina + ' — ' + acumulado + ' issues carregadas...');

    if (data.isLast === true) { isLast = true; break; }
    token_next = data.nextPageToken;
    if (!token_next) { isLast = true; break; }
    Utilities.sleep(150);
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    [PROP_NEXT_TOKEN]: token_next || '',
    [PROP_PAGINA]:     String(pagina),
    [PROP_ACUMULADO]:  String(acumulado)
  });
  _cacheSet('SYNC_NEXT_TOKEN', token_next || '');

  if (lote.length) _bufferAppend(ss, lote);

  if (isLast) {
    if (acumulado === 0) {
      Logger.log('Sync incremental: 0 issues novas — dados já atualizados.');
      _setProgresso('concluido', '0 issues novas — dados já atualizados.');
      _bufferDeletar(ss);
      return;
    }
    Logger.log('Busca completa: ' + acumulado + ' issues — iniciando gravação...');
    _setProgresso('gravando', 'Busca concluída (' + acumulado + ' issues) — gravando RAW...');
    props.setProperty(PROP_SYNC_TOTAL, String(acumulado));
    ScriptApp.newTrigger('_gravarRAWDoBusfer').timeBased().after(2000).create();
  } else {
    Logger.log('Agendando próxima página a partir da ' + (pagina + 1) + '...');
    ScriptApp.newTrigger('_buscarProximaPagina').timeBased().after(2000).create();
  }
}

// ── CONTINUAÇÃO DA BUSCA ──────────────────────────────────────
function _buscarProximaPagina() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '_buscarProximaPagina') ScriptApp.deleteTrigger(t);
  });
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const jqlExtra   = props.getProperty(PROP_JQL_EXTRA)   || '';
  const paginaBase = parseInt(props.getProperty(PROP_PAGINA)    || '0', 10);
  const acumBase   = parseInt(props.getProperty(PROP_ACUMULADO) || '0', 10);
  const nextToken  = _cacheGet('SYNC_NEXT_TOKEN') || props.getProperty(PROP_NEXT_TOKEN) || null;
  if (!nextToken) {
    Logger.log('nextPageToken perdido — reiniciando sync completo.');
    _setProgresso('buscando', 'Token perdido — reiniciando busca...');
    sincronizarCompleto();
    return;
  }
  Logger.log('Continuando busca a partir da página ' + (paginaBase + 1) + ' | acumulado=' + acumBase);
  _executarPaginaBusca(ss, jqlExtra, nextToken, paginaBase, acumBase);
}

// ── GRAVAÇÃO DO RAW ───────────────────────────────────────────
function _gravarRAWDoBusfer() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '_gravarRAWDoBusfer') ScriptApp.deleteTrigger(t);
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Lendo buffer...');
  const issues = _bufferLer(ss);
  if (!issues.length) {
    _setProgresso('erro', 'Buffer vazio — execute a sincronização novamente.');
    return;
  }
  _setProgresso('gravando', 'Separando ' + issues.length + ' issues em pais, filhos e incorretos...');
  Logger.log('Buffer: ' + issues.length + ' issues — separando e gravando RAW...');
  const jqlExtra    = PropertiesService.getScriptProperties().getProperty(PROP_JQL_EXTRA) || '';
  const incremental = jqlExtra !== '';
  const { pais, filhos, incorretos } = _separar(issues, incremental);

  const nPais       = Object.keys(pais).length;
  const nFilhos     = filhos.length;
  const nIncorretos = incorretos.length;
  const nTotal      = nPais + nFilhos + nIncorretos;
  Logger.log('══ AUDITORIA ══ Buffer=' + issues.length + ' | Pais=' + nPais + ' | Filhos=' + nFilhos + ' | Incorretos=' + nIncorretos + ' | Total=' + nTotal + (nTotal===issues.length?' ✅':' ⚠️ DIVERGÊNCIA'));

  filhos.forEach(f => {
    const pk = (f.fields.parent || {}).key;
    if (pk && !pais[pk]) pais[pk] = _stubPai(pk, (f.fields.parent.fields || {}).summary || pk);
  });

  _setProgresso('gravando', 'Gravando ' + nPais + ' tarefas pai (RAW_PAI)...');
  _gravarRAWPai(ss, Object.values(pais));

  _setProgresso('gravando', 'Gravando ' + nFilhos + ' subtarefas (RAW_FILHO)...');
  _gravarRAWFilho(ss, filhos);

  _setProgresso('gravando', nIncorretos > 0 ? 'Gravando ' + nIncorretos + ' issues com tipo incorreto...' : 'Sem issues com tipo incorreto.');
  _gravarIncorretos(ss, incorretos);

  PropertiesService.getScriptProperties().setProperty(PROP_ULTIMA_SYNC, new Date().toISOString());
  _bufferDeletar(ss);
  Logger.log('RAW gravado. Agendando cálculo do PAINEL...');
  _setProgresso('calculando', 'Calculando PAINEL...');
  ScriptApp.newTrigger('_etapa2GravarPainel').timeBased().after(2000).create();
}

// ── ETAPA 2a: PAINEL ─────────────────────────────────────────
function _etapa2GravarPainel() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '_etapa2GravarPainel') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Etapa 2a: gravando PAINEL...');
  _setProgresso('calculando', 'Calculando PAINEL...');

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const rawP = ss.getSheetByName('RAW_PAI');
  const rawF = ss.getSheetByName('RAW_FILHO');

  if (!rawP || !rawF) {
    _setProgresso('erro', 'RAW não encontrado — execute a sincronização novamente.');
    return;
  }

  const pais = {};
  rawP.getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]) pais[r[0]] = _rowParaIssuePai(r);
  });

  const hdrF  = rawF.getRange(1, 1, 1, rawF.getLastColumn()).getValues()[0];
  const filhos = [];
  rawF.getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]) filhos.push(_rowParaIssueFilho(r, hdrF));
  });

  const nPaisE2 = Object.keys(pais).length;
  const nFilhosE2 = filhos.length;
  _setProgresso('calculando', 'Calculando PAINEL com ' + nPaisE2 + ' campanhas e ' + nFilhosE2 + ' subtarefas...');
  _gravarPainel(ss, Object.values(pais), filhos);
  Logger.log('PAINEL gravado. Agendando TABELA...');
  _setProgresso('calculando', 'Gerando TABELA...');
  ScriptApp.newTrigger('_etapa2GravarTabela').timeBased().after(2000).create();
}

// ── ETAPA 2b: TABELA ─────────────────────────────────────────
function _etapa2GravarTabela() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '_etapa2GravarTabela') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Etapa 2b: gravando TABELA...');
  _setProgresso('calculando', 'Gerando TABELA...');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const painel = ss.getSheetByName('PAINEL');
  const nLinhas = painel ? Math.max(0, painel.getLastRow() - 1) : 0;
  _setProgresso('calculando', 'Gerando TABELA com ' + nLinhas + ' linhas...');
  try { _gravarTabela(ss); } catch(e) { Logger.log('TABELA: ' + e); }

  Logger.log('TABELA gravada. Agendando finalização...');
  _setProgresso('formatando', 'Finalizando...');
  ScriptApp.newTrigger('_etapa2Finalizar').timeBased().after(2000).create();
}

// ── ETAPA 2c: INICIO + FORMATAÇÃO + CONCLUÍDO ────────────────
function _etapa2Finalizar() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '_etapa2Finalizar') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Etapa 2c: INICIO + reordenação...');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  // INICIO — só no sync completo (JQL_EXTRA vazio = completo)
  const jqlExtra = props.getProperty(PROP_JQL_EXTRA) || '';
  const isCompleto = jqlExtra === '';

  if (isCompleto) {
    _setProgresso('formatando', 'Atualizando manual (INICIO)...');
    try { _criarAbaInicio(ss); } catch(e) { Logger.log('INICIO: ' + e); }
  }

  // Reordena abas
  try {
    const ORDEM = ['🏠 INICIO', 'PAINEL', '📋 TABELA', 'RAW_PAI', 'RAW_FILHO', 'DE_PARA', '⚠️ INCORRETOS'];
    ORDEM.slice().reverse().forEach(nome => {
      const sh = ss.getSheetByName(nome);
      if (!sh) return;
      sh.activate();
      ss.moveActiveSheet(1);
    });
  } catch(e) { Logger.log('Reordenação abas: ' + e); }

  // SYNC_TOTAL = RAW_PAI + RAW_FILHO + INCORRETOS (número real gravado)
  try {
    const rawP = ss.getSheetByName('RAW_PAI');
    const rawF = ss.getSheetByName('RAW_FILHO');
    const rawI = ss.getSheetByName('⚠️ INCORRETOS');
    const nP = rawP && rawP.getLastRow() > 1 ? rawP.getLastRow() - 1 : 0;
    const nF = rawF && rawF.getLastRow() > 1 ? rawF.getLastRow() - 1 : 0;
    const nI = rawI && rawI.getLastRow() > 1 ? rawI.getLastRow() - 1 : 0;
    const totalReal = nP + nF + nI;
    if (totalReal > 0) {
      props.setProperty(PROP_SYNC_TOTAL, String(totalReal));
      Logger.log('SYNC_TOTAL recalculado: ' + nP + ' pais + ' + nF + ' filhos + ' + nI + ' incorretos = ' + totalReal);
    }
  } catch(e) { Logger.log('Erro ao recalcular SYNC_TOTAL: ' + e); }

  // Formatação leve — só limpeza de abas extras (sem setRowHeight em massa)
  _setProgresso('formatando', 'Limpando abas extras...');
  try {
    const PERMITIDAS = new Set(['🏠 INICIO','PAINEL','📋 TABELA','RAW_PAI','RAW_FILHO','DE_PARA','⚠️ INCORRETOS','_SYNC_BUFFER']);
    ss.getSheets().forEach(sh => {
      if (!PERMITIDAS.has(sh.getName()) && ss.getSheets().length > 1) {
        try { ss.deleteSheet(sh); } catch(e) {}
      }
    });
  } catch(e) { Logger.log('Limpeza: ' + e); }

  const total = props.getProperty(PROP_SYNC_TOTAL) || '?';
  Logger.log('Sync concluído: ' + total + ' issues.');
  _setProgresso('concluido', total + ' issues sincronizadas · PAINEL e TABELA atualizados.');
}

// ── ETAPA 2 legada — mantida para compatibilidade ─────────────
// Redireciona para o novo fluxo encadeado
function _sincronizarEtapa2() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === '_sincronizarEtapa2') ScriptApp.deleteTrigger(t);
  });
  Logger.log('_sincronizarEtapa2 legada — redirecionando para _etapa2GravarPainel...');
  ScriptApp.newTrigger('_etapa2GravarPainel').timeBased().after(1000).create();
}

// ── Separação de issues por tipo ──────────────────────────────
function _separar(issues, incremental) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pais = {}, filhos = [], incorretos = [];

  issues.forEach(i => {
    const tipo = (i.fields.issuetype || {}).name || '';
    if (!TIPOS_VALIDOS.includes(tipo)) { incorretos.push(i); return; }
    if (tipo === 'Tarefa' || !i.fields.parent) pais[i.key] = i;
    else filhos.push(i);
  });

  if (incremental) {
    const rawP = ss.getSheetByName('RAW_PAI'), rawF = ss.getSheetByName('RAW_FILHO');
    const paisKeys       = new Set(Object.keys(pais));
    const filhosKeys     = new Set(filhos.map(f => f.key));
    const incorretosKeys = new Set(incorretos.map(i => i.key));

    if (rawP && rawP.getLastRow() > 1) {
      rawP.getDataRange().getValues().slice(1).forEach(r => {
        if (r[0] && !paisKeys.has(r[0])) {
          pais[r[0]] = _rowParaIssuePai(r);
          paisKeys.add(r[0]);
        }
      });
    }
    if (rawF && rawF.getLastRow() > 1) {
      const hdr = rawF.getRange(1, 1, 1, rawF.getLastColumn()).getValues()[0];
      const ki  = hdr.indexOf('KEY');
      rawF.getDataRange().getValues().slice(1).forEach(r => {
        if (r[ki] && !filhosKeys.has(r[ki])) {
          filhos.push(_rowParaIssueFilho(r, hdr));
          filhosKeys.add(r[ki]);
        }
      });
    }
    const rawI = ss.getSheetByName('⚠️ INCORRETOS');
    if (rawI && rawI.getLastRow() > 1) {
      const hdrI = rawI.getRange(1, 1, 1, rawI.getLastColumn()).getValues()[0];
      const kiI  = hdrI.indexOf('KEY');
      rawI.getDataRange().getValues().slice(1).forEach(r => {
        if (r[kiI] && !incorretosKeys.has(r[kiI])) {
          incorretos.push({ key: r[kiI], fields: {
            issuetype: { name: r[hdrI.indexOf('TIPO')] || '' },
            summary:   r[hdrI.indexOf('SUMMARY')] || '',
            status:    { name: r[hdrI.indexOf('STATUS')] || '' },
            assignee:  { displayName: r[hdrI.indexOf('RESPONSÁVEL')] || '' },
            updated:   r[hdrI.indexOf('ATUALIZADO')] || ''
          }});
          incorretosKeys.add(r[kiI]);
        }
      });
    }
    Logger.log('Merge incremental: ' + Object.keys(pais).length + ' pais, ' + filhos.length + ' filhos, ' + incorretos.length + ' incorretos.');
  }
  return { pais, filhos, incorretos };
}

// ── Extração de campos custom ─────────────────────────────────
function _cx(f) {
  let plataforma = '';
  if (Array.isArray(f.customfield_11545)) {
    plataforma = f.customfield_11545.map(v => v.value || v).filter(Boolean).join(' | ');
  } else if (f.customfield_11545 && f.customfield_11545.value) {
    plataforma = f.customfield_11545.value;
  }
  return {
    relator:    (f.reporter  || {}).displayName || '',
    duedate:    f.duedate || '',
    qtdTax:     f.customfield_10184 != null ? f.customfield_10184 : '',
    statusAdp:  (f.customfield_10185 || {}).value || '',
    bu:         (f.customfield_10186 || {}).value || '',
    marcaJira:  (f.customfield_10444 || {}).value || '',
    plataforma: plataforma,
    agencia:    (f.customfield_10188 || {}).value || '',
    campMidia:  (f.customfield_10807 || {}).value || ''
  };
}

// ── Reconstrução de issues a partir das abas ──────────────────
function _rowParaIssuePai(r) {
  return { key: r[0], fields: {
    issuetype: { name: 'Tarefa' }, summary: r[2] || '', parent: null,
    status: { name: r[3] || '' }, assignee: r[4] ? { displayName: r[4] } : null,
    created: r[5], updated: r[6], reporter: { displayName: r[7] || '' },
    duedate: r[8] || '', customfield_10184: r[9] || '',
    customfield_10185: r[10] ? { value: r[10] } : null,
    customfield_10186: r[11] ? { value: r[11] } : null,
    customfield_10444: r[12] ? { value: r[12] } : null,
    customfield_11545: r[13] ? { value: r[13] } : null,
    customfield_10188: r[14] ? { value: r[14] } : null,
    customfield_10807: r[15] ? { value: r[15] } : null
  }};
}

function _rowParaIssueFilho(r, hdr) {
  const g = k => { const i = hdr.indexOf(k); return i >= 0 ? r[i] : ''; };
  return { key: g('KEY'), fields: {
    issuetype: { name: 'Subtarefa' }, summary: g('SUMMARY') || '',
    parent: { key: g('PARENT_KEY'), fields: { summary: g('PARENT_SUMMARY') || '' } },
    status: { name: g('STATUS') || '' },
    assignee: g('RESPONSÁVEL') ? { displayName: g('RESPONSÁVEL') } : null,
    created: g('CRIADO'), updated: g('ATUALIZADO'),
    reporter: { displayName: g('RELATOR') || '' },
    duedate: g('DATA LIMITE') || '', customfield_10184: g('QTD TAXONOMIA') || '',
    customfield_10185: g('STATUS ADP')    ? { value: g('STATUS ADP') }    : null,
    customfield_10186: g('BU')            ? { value: g('BU') }            : null,
    customfield_10444: g('MARCA (JIRA)')  ? { value: g('MARCA (JIRA)') }  : null,
    customfield_11545: g('PLATAFORMA')    ? { value: g('PLATAFORMA') }    : null,
    customfield_10188: g('AGÊNCIA_UL')    ? { value: g('AGÊNCIA_UL') }    : null,
    customfield_10807: g('CAMP. MÍDIA')   ? { value: g('CAMP. MÍDIA') }   : null
  }};
}

function _stubPai(key, summary) {
  return { key, fields: {
    issuetype: { name: 'Tarefa' }, summary, parent: null,
    status: { name: '' }, assignee: null, created: '', updated: '',
    reporter: null, duedate: '', customfield_10184: '',
    customfield_10185: null, customfield_10186: null, customfield_10444: null,
    customfield_11545: null, customfield_10188: null, customfield_10807: null
  }};
}

function _fmt(d) {
  if (!d) return '';
  try { return Utilities.formatDate(new Date(d), 'America/Sao_Paulo', 'yyyy-MM-dd'); }
  catch(e) { return String(d).substring(0, 10); }
}

// ── Exportar dados para planilha externa ──────────────────────
function exportarDados(params) {
  try {
    const { url, nomeAba, colunas, linhas } = params;
    let ss;
    if (url && url.trim()) {
      try {
        const id = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (!id) throw new Error('URL inválida');
        ss = SpreadsheetApp.openById(id[1]);
      } catch(e) {
        Logger.log('Erro ao abrir planilha do usuário: ' + e + ' — criando nova.');
        ss = null;
      }
    }
    const nomeFinal = nomeAba || ('Export UL ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'));
    if (!ss) ss = SpreadsheetApp.create(nomeFinal);
    const existing = ss.getSheetByName(nomeFinal);
    if (existing) ss.deleteSheet(existing);
    const sh = ss.insertSheet(nomeFinal);
    const todasCols = ['PAI_KEY','PAI_SUMMARY','BU','AGÊNCIA','MARCA (JIRA)','CAMP. MÍDIA',
                       'STATUS ADP (PAI)','FILHO_KEY','SUMMARY (FILHO)','VEÍCULO','STATUS',
                       'RESPONSÁVEL','RELATOR','QTD TAX.','PLATAFORMA','DATA LIMITE','ÚLT. ATU.'];
    const colsSel = (colunas && colunas.length) ? colunas : todasCols;
    sh.getRange(1, 1, 1, colsSel.length).setValues([colsSel])
      .setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold').setFontSize(10);
    sh.setFrozenRows(1);
    if (linhas && linhas.length) {
      const rows = linhas.map(r => colsSel.map(c => r[c] !== undefined ? r[c] : ''));
      sh.getRange(2, 1, rows.length, colsSel.length).setValues(rows).setFontSize(10);
      for (let i = 0; i < rows.length; i++) {
        sh.getRange(i+2, 1, 1, colsSel.length).setBackground(i%2===0 ? '#eef2ff' : '#fff');
      }
    }
    try { sh.autoResizeColumns(1, colsSel.length); } catch(e) {}
    const ssUrl = ss.getUrl();
    Logger.log('Exportado para: ' + ssUrl);
    return { ok: true, url: ssUrl, nome: nomeFinal, total: linhas ? linhas.length : 0 };
  } catch(e) {
    Logger.log('Erro exportarDados: ' + e);
    return { ok: false, erro: String(e) };
  }
}

// ── Busca simples (usado por getDadosPeriodo) ─────────────────
function _buscar(jqlExtra) {
  const { email, token } = _creds();
  const auth = Utilities.base64Encode(email + ':' + token);
  const jql  = 'project=' + PROJETO + (jqlExtra ? ' AND ' + jqlExtra : '') + ' ORDER BY key ASC';
  const issues = [];
  let nextToken = null;
  let pagina = 0;
  while (true) {
    let url = JIRA_BASE + '/rest/api/3/search/jql?jql=' + encodeURIComponent(jql) +
              '&maxResults=100&fields=' + CAMPOS;
    if (nextToken) url += '&nextPageToken=' + encodeURIComponent(nextToken);
    const resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { Authorization: 'Basic ' + auth, Accept: 'application/json' },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) break;
    const data = JSON.parse(resp.getContentText());
    issues.push(...(data.issues || []));
    pagina++;
    if (data.isLast || !data.nextPageToken) break;
    nextToken = data.nextPageToken;
    Utilities.sleep(150);
    if (pagina > 50) break; // safeguard
  }
  return issues;
}
