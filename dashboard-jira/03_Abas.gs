// ============================================================
// [UL] Dashboard_Taxonomias — 03_Abas.gs
// Gravação, formatação e proteção de todas as abas
// ============================================================

// ── Colunas que devem ter alinhamento CENTRAL (horizontal) ───
const COLS_CENTRO = ['KEY','TIPO','STATUS','CRIADO','ATUALIZADO','DATA LIMITE',
  'QTD TAXONOMIA','QTD TAX.','STATUS ADP','STATUS ADP (PAI)','BU','MARCA (JIRA)',
  'PLATAFORMA','ÚLT. ATU.','PARENT_KEY','FILHO_KEY','PAI_KEY','VEÍCULO','CAMP. MÍDIA'];

function _alinharColunas(sh, hdr, nRows) {
  if (nRows < 2) return;
  hdr.forEach((h, i) => {
    const col = i + 1;
    const alinhamento = COLS_CENTRO.includes(h) ? 'center' : 'left';
    sh.getRange(2, col, nRows - 1, 1).setHorizontalAlignment(alinhamento);
  });
}

function _wrapColunas(sh, hdr, nRows) {
  const WRAP_COLS = ['SUMMARY','PAI_SUMMARY','PARENT_SUMMARY','SUMMARY (FILHO)','ASSIGNEE','RELATOR','RESPONSÁVEL','AGÊNCIA_UL','AGÊNCIA'];
  hdr.forEach((h, i) => {
    const col = i + 1;
    const wrap = WRAP_COLS.includes(h) ? SpreadsheetApp.WrapStrategy.WRAP : SpreadsheetApp.WrapStrategy.CLIP;
    sh.getRange(1, col, Math.max(nRows, 1), 1).setWrapStrategy(wrap);
  });
}

function _protegerCabecalho(sh) {
  try {
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
    const prot = sh.getRange(1, 1, 1, sh.getLastColumn()).protect();
    prot.setDescription('Cabeçalho — não editar');
    prot.setWarningOnly(true);
  } catch(e) {}
}

// ── RAW PAI ───────────────────────────────────────────────────
function _gravarRAWPai(_ss, pais) {
  _ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = _ss.getSheetByName('RAW_PAI'); if (old) _ss.deleteSheet(old);
  const sh = _ss.insertSheet('RAW_PAI'); sh.setTabColor('#2E5BCD');
  const H = ['KEY','TIPO','SUMMARY','STATUS','ASSIGNEE','CRIADO','ATUALIZADO',
             'RELATOR','DATA LIMITE','QTD TAXONOMIA','STATUS ADP','BU','MARCA (JIRA)',
             'PLATAFORMA','AGÊNCIA_UL','CAMP. MÍDIA'];
  const rows = [H];
  pais.forEach(i => {
    const f = i.fields, c = _cx(f);
    rows.push([i.key,(f.issuetype||{}).name||'',f.summary||'',(f.status||{}).name||'',
      (f.assignee||{}).displayName||'',_fmt(f.created),_fmt(f.updated),
      c.relator,c.duedate,c.qtdTax,c.statusAdp,c.bu,c.marcaJira,c.plataforma,c.agencia,c.campMidia]);
  });
  sh.getRange(1,1,rows.length,H.length).setValues(rows);
  _estilizarSheet(sh, rows.length, H.length, true);
  [80,75,360,110,160,90,90,180,90,60,110,120,110,100,110,120].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  _alinharColunas(sh, H, rows.length);
  _wrapColunas(sh, H, rows.length);
  try { const f=sh.getFilter(); if(f) f.remove(); } catch(e) {}
  try { sh.getRange(1,1,rows.length,H.length).createFilter(); } catch(e) {}
  _sinalizar(sh, rows.length, H, ['STATUS ADP','BU','MARCA (JIRA)','AGÊNCIA_UL']);
  _protegerCabecalho(sh);
  _trim(sh, rows.length, H.length);
}

// ── RAW FILHO ─────────────────────────────────────────────────
function _gravarRAWFilho(ss, filhos) {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName('RAW_FILHO'); if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet('RAW_FILHO'); sh.setTabColor('#7C3AED');
  const H = ['KEY','TIPO','SUMMARY','PARENT_KEY','PARENT_SUMMARY','STATUS','ASSIGNEE',
             'CRIADO','ATUALIZADO','RELATOR','DATA LIMITE','QTD TAXONOMIA','STATUS ADP',
             'BU','MARCA (JIRA)','PLATAFORMA','AGÊNCIA_UL','CAMP. MÍDIA'];
  const rows = [H];
  filhos.forEach(i => {
    const f = i.fields, pf = (f.parent||{}).fields||{}, c = _cx(f);
    rows.push([i.key,(f.issuetype||{}).name||'',f.summary||'',(f.parent||{}).key||'',
      pf.summary||'',(f.status||{}).name||'',(f.assignee||{}).displayName||'',
      _fmt(f.created),_fmt(f.updated),c.relator,c.duedate,c.qtdTax,
      c.statusAdp,c.bu,c.marcaJira,c.plataforma,c.agencia,c.campMidia]);
  });
  ss.getRange ? null : (ss = SpreadsheetApp.getActiveSpreadsheet());
  sh.getRange(1,1,rows.length,H.length).setValues(rows);
  _estilizarSheet(sh, rows.length, H.length, true);
  [80,75,340,80,300,110,160,90,90,180,90,60,110,120,110,100,110,120].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  _alinharColunas(sh, H, rows.length);
  _wrapColunas(sh, H, rows.length);
  try { const f=sh.getFilter(); if(f) f.remove(); } catch(e) {}
  try { sh.getRange(1,1,rows.length,H.length).createFilter(); } catch(e) {}
  _sinalizar(sh, rows.length, H, ['STATUS ADP','BU','MARCA (JIRA)','AGÊNCIA_UL']);
  _protegerCabecalho(sh);
  _trim(sh, rows.length, H.length);
}

// ── DE_PARA ───────────────────────────────────────────────────
function inicializarDePara() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('DE_PARA')) return;
  _gravarDePara(ss.insertSheet('DE_PARA'));
}
function inicializarDeParaForcar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName('DE_PARA');
  if (s) ss.deleteSheet(s);
  _gravarDePara(ss.insertSheet('DE_PARA'));
  SpreadsheetApp.getUi().alert('DE_PARA recriada.');
}
function _gravarDePara(sh) {
  const d = [
    ['TIPO','TOKEN_JIRA','NOME_OFICIAL'],
    ['MARCA','Dove','Dove'],['MARCA','Dove Skincare','Dove'],['MARCA','Dove Diva','Dove'],
    ['MARCA','Dove MSTB','Dove'],['MARCA','Dove Deos','Dove'],['MARCA','Dove Men Care','Dove'],
    ['MARCA','Debeers','Dove'],['MARCA','Dove Debeers','Dove'],['MARCA','Dune','Dove'],['MARCA','DMC','Dove'],
    ['MARCA','Seda','Seda'],['MARCA','Seda Core','Seda'],['MARCA','Seda Steel','Seda'],['MARCA','RES','Seda'],
    ['MARCA','Closeup','Closeup'],['MARCA','Close Up','Closeup'],['MARCA','Close-Up','Closeup'],
    ['MARCA','Hellmanns',"Hellmann's"],['MARCA',"Hellmann's","Hellmann's"],['MARCA','HELLMANNS',"Hellmann's"],
    ['MARCA','Copa Supreme',"Hellmann's"],['MARCA','OMO','OMO'],['MARCA','Alchemist','OMO'],
    ['MARCA','CIF','CIF'],['MARCA','Rexona','Rexona'],['MARCA','Lux','Lux'],
    ['MARCA','Clear','Clear'],['MARCA','Crown','Clear'],['MARCA','Crown World Cup','Clear'],
    ['MARCA','Knorr','Knorr'],['MARCA','KNORR','Knorr'],
    ['VEICULO','YT','YouTube'],['VEICULO','Youtube','YouTube'],['VEICULO','YouTube','YouTube'],
    ['VEICULO','Meta','Meta'],['VEICULO','Tiktok','TikTok'],['VEICULO','TikTok','TikTok'],
    ['VEICULO','DV360','DV360'],['VEICULO','Pinterest','Pinterest'],['VEICULO','Social','Social'],
    ['VEICULO','Flashtalking','Flashtalking'],['VEICULO','Flashtalkng','Flashtalking'],
    ['VEICULO','MP','Mídia Programática'],['VEICULO','Programmatic','Mídia Programática'],
    ['VEICULO','Midia Programatica','Mídia Programática'],
    ['VEICULO','Amazon DSP','Amazon DSP'],['VEICULO','Amz Ads','Amazon DSP'],
    ['VEICULO','Google Search','Google Search'],['VEICULO','Netflix','Netflix'],
    ['VEICULO','Spotify','Spotify'],['VEICULO','Meliads','Meliads'],['VEICULO','UOL','UOL'],
    ['VEICULO','Futbol Sites','Futbol Sites'],['VEICULO','365 Scores','365 Scores'],
    ['VEICULO','GETV','GETV'],['VEICULO','Globoplay','Globoplay'],
    ['VEICULO','Tag Kantar','Tag Kantar'],['VEICULO','B&W','B&W'],
  ];
  sh.getRange(1,1,d.length,3).setValues(d);
  _estilizarSheet(sh, d.length, 3, true);
  [90,220,200].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.getRange(1,1,d.length,3).setVerticalAlignment('middle');
  sh.getRange(2,1,d.length-1,1).setHorizontalAlignment('center');
  sh.getRange(2,2,d.length-1,2).setHorizontalAlignment('left');
  sh.setFrozenRows(1);
  _protegerCabecalho(sh);
  _trim(sh, d.length, 3);
}

function _extrairTokens(texto, mapa) {
  const found = new Set();
  for (const token of Object.keys(mapa)) {
    const re = new RegExp('(?:^|[\\s|\\-/,])' + token.replace(/[.*+?^${}()|[\]\\]/g,'\\$&') + '(?:$|[\\s|\\-/,])', 'i');
    if (re.test(' ' + texto + ' ')) found.add(mapa[token]);
  }
  return [...found];
}

function _carregarDePara(tipo) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DE_PARA');
  if (!sh) return {};
  const map = {};
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]===tipo && r[1] && r[2]) map[r[1].toString().toLowerCase().trim()] = r[2].toString().trim();
  });
  return map;
}

// ── PAINEL ────────────────────────────────────────────────────
function _gravarPainel(ss, pais, filhos) {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName('PAINEL'); if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet('PAINEL');
  const dv = _carregarDePara('VEICULO');

  // Sem filtro de data — PAINEL contém histórico completo (igual ao Jira)

  const H = ['PAI_KEY','PAI_SUMMARY','BU','AGÊNCIA','MARCA (JIRA)','CAMP. MÍDIA','STATUS ADP (PAI)',
             'FILHO_KEY','SUMMARY (FILHO)','VEÍCULO','STATUS','RESPONSÁVEL','RELATOR',
             'QTD TAX.','PLATAFORMA','DATA LIMITE','ÚLT. ATU.'];
  const mapaPai = {};
  pais.forEach(p => {
    const c = _cx(p.fields);
    mapaPai[p.key] = { key:p.key, summary:p.fields.summary||'',
      bu:c.bu, agencia:c.agencia, marcaJira:c.marcaJira, campMidia:c.campMidia,
      statusAdp:c.statusAdp, assignee:(p.fields.assignee||{}).displayName||'',
      relator:c.relator, duedate:c.duedate, plataforma:c.plataforma, qtdTax:c.qtdTax,
      status:(p.fields.status||{}).name||'', updated:_fmt(p.fields.updated) };
  });
  const porPai = {};
  filhos.forEach(f => {
    const pk = (f.fields.parent||{}).key||'SEM_PAI';
    if (!porPai[pk]) porPai[pk] = [];
    porPai[pk].push(f);
  });
  const rows = [H];
  Object.keys(mapaPai).sort().forEach(pk => {
    const pai = mapaPai[pk], fd = porPai[pk]||[];
    if (!fd.length) {
      rows.push([pai.key,pai.summary,pai.bu||'⚠️ vazio',pai.agencia||'⚠️ vazio',
        pai.marcaJira||'⚠️ vazio',pai.campMidia||'',pai.statusAdp||'⚠️ vazio',
        '','(sem subtarefas)','',pai.status,pai.assignee,pai.relator,
        pai.qtdTax||'',pai.plataforma||'',pai.duedate||'',pai.updated||'']);
    } else {
      fd.forEach(f => {
        const ff = f.fields, cf = _cx(ff);
        const bu=cf.bu||pai.bu||'⚠️ vazio', ag=cf.agencia||pai.agencia||'⚠️ vazio';
        const mk=cf.marcaJira||pai.marcaJira||'⚠️ vazio', cm=cf.campMidia||pai.campMidia||'';
        const sa=cf.statusAdp||pai.statusAdp||'⚠️ vazio', pl=cf.plataforma||pai.plataforma||'';
        const qt=cf.qtdTax!==''?cf.qtdTax:pai.qtdTax||'', dd=cf.duedate||pai.duedate||'';
        const rl=cf.relator||pai.relator||'', upd=_fmt(ff.updated);
        const sum=ff.summary||'';
        const vv=_extrairTokens(sum,dv); if(!vv.length) vv.push('⚠️ Não mapeado');
        vv.forEach(vei => rows.push([pai.key,pai.summary,bu,ag,mk,cm,sa,
          f.key,sum,vei,(ff.status||{}).name||'',(ff.assignee||{}).displayName||'',rl,qt,pl,dd,upd]));
      });
    }
  });
  // Órfãos
  (porPai['SEM_PAI']||[]).forEach(f => {
    const ff=f.fields, cf=_cx(ff), sum=ff.summary||'';
    const vv=_extrairTokens(sum,dv); if(!vv.length) vv.push('⚠️ Não mapeado');
    vv.forEach(vei => rows.push(['','(sem pai)',cf.bu||'⚠️ vazio',cf.agencia||'⚠️ vazio',
      cf.marcaJira||'⚠️ vazio',cf.campMidia||'',cf.statusAdp||'⚠️ vazio',
      f.key,sum,vei,(ff.status||{}).name||'',(ff.assignee||{}).displayName||'',cf.relator||'',
      cf.qtdTax||'',cf.plataforma||'',cf.duedate||'',_fmt(ff.updated)]));
  });
  sh.getRange(1,1,rows.length,H.length).setValues(rows);
  _estilizarSheet(sh, rows.length, H.length, true);
  [80,340,115,115,120,115,115,80,320,115,115,155,155,55,100,90,90].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(1); sh.setFrozenColumns(2);
  _alinharColunas(sh, H, rows.length);
  _wrapColunas(sh, H, rows.length);
  try { const f=sh.getFilter(); if(f) f.remove(); } catch(e) {}
  try { sh.getRange(1,1,rows.length,H.length).createFilter(); } catch(e) {}
  sh.getRange(rows.length+2,1).setValue('⏱ Sync: '+new Date().toLocaleString('pt-BR')).setFontColor('#888').setFontStyle('italic');
  _protegerCabecalho(sh);
  _trim(sh, rows.length+2, H.length);
}

// ── TABELA — atualização manual ───────────────────────────────
function atualizarAbas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try { _gravarTabela(ss); } catch(e) { Logger.log('TABELA: '+e); }
  SpreadsheetApp.getUi().alert('Aba TABELA atualizada.');
}

function _gravarVisual(ss) {
  // VISUAL eliminado — visualização feita pelo painel HTML
  // Remove aba legada se existir
  const old = ss.getSheetByName('📊 VISUAL'); if (old) ss.deleteSheet(old);
  Logger.log('VISUAL eliminado — usando painel HTML.');
}


// ── 🏠 INICIO — manual gerado automaticamente ────────────────
function _criarAbaInicio(ss) {
  const old = ss.getSheetByName('🏠 INICIO'); if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet('🏠 INICIO');
  sh.setTabColor('#2E5BCD');

  const now = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  const syncTotal   = PropertiesService.getScriptProperties().getProperty('SYNC_TOTAL')   || '?';
  const tokenExpira = PropertiesService.getScriptProperties().getProperty('TOKEN_EXPIRA') || '?';

  const linhas = [
    ['📋 [UL] Dashboard_Taxonomias'],
    ['Unilever BR · Grasp x StormX'],
    [''],
    ['⏱ Última atualização: ' + now + ' · ' + syncTotal + ' issues sincronizadas'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['📌 O QUE É ESTE ARQUIVO'],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['Este arquivo sincroniza automaticamente os tickets do Jira (projeto UL) e organiza os dados em abas.'],
    ['O painel visual fica em: https://falssp.github.io/ul-dashboard-taxonomias/'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['🗂 ABAS'],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['🏠 INICIO      →  Este manual'],
    ['PAINEL        →  Fonte de dados do painel HTML'],
    ['📋 TABELA     →  Cópia formatada do PAINEL para consulta direta'],
    ['RAW_PAI       →  Todas as Tarefas do Jira (campanhas)'],
    ['RAW_FILHO     →  Todas as Subtarefas do Jira (tickets de trabalho)'],
    ['DE_PARA       →  Dicionário de normalização — edite para mapear veículos e marcas'],
    ['⚠️ INCORRETOS →  Issues com tipo inválido — corrija no Jira'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['⚙️ MENU — 📋 Dashboard UL'],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['1. Configurar credenciais  →  E-mail Atlassian + API Token do Jira'],
    ['2. Testar conexão          →  Verifica se o token ainda funciona'],
    ['3. Sincronizar             →  Força sync completo imediato'],
    ['4. Atualizar TABELA        →  Regenera TABELA sem rebuscar o Jira'],
    ['Abrir painel               →  Abre https://falssp.github.io/ul-dashboard-taxonomias/'],
    ['Avançado › Sync incremental         →  Força sync da última 1h'],
    ['Avançado › Recriar DE_PARA          →  Restaura mapeamentos padrão'],
    ['Avançado › Reconfigurar acionadores →  Recria o trigger de 1h'],
    ['Avançado › Limpar abas extras       →  Remove abas fora do padrão'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['🔄 SYNC — roda automaticamente a cada 1h'],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['Incremental: busca só issues atualizadas na última hora e faz merge.'],
    ['Completo: Menu → 3. Sincronizar (todas as issues, ~10 minutos).'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['🗺 DE_PARA — como mapear'],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['Aba DE_PARA → nova linha:   TIPO | TOKEN_JIRA | NOME_OFICIAL'],
    ['Exemplo:   VEICULO | YT | YouTube   ou   MARCA | Dove Deos | Dove'],
    ['Se aparecer "⚠️ Não mapeado" no painel, adicione o token aqui.'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['🔑 TOKEN DO JIRA — válido até: ' + tokenExpira],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['E-mail automático quando faltam 10 dias para vencer.'],
    ['Renovar: id.atlassian.com → Segurança → Tokens de API → Menu → 1. Configurar credenciais'],
    [''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['📞 SUPORTE'],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    [''],
    ['Repositório: https://github.com/falssp/ul-dashboard-taxonomias'],
    ['Responsável: StormX Data & Tech · Projeto: Unilever BR · Grasp x StormX'],
    [''],
  ];

  const nRows = linhas.length;
  const range = sh.getRange(1, 1, nRows, 1);
  range.setValues(linhas);
  range
    .setFontFamily('Inter, Arial, sans-serif')
    .setFontSize(11)
    .setFontColor('#e8ecf4')
    .setBackground('#0f1117')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  sh.getRange(1, 1).setFontSize(18).setFontWeight('bold').setFontColor('#ffffff');
  sh.getRange(2, 1).setFontSize(12).setFontColor('#6b7a9e');
  sh.getRange(4, 1).setFontSize(10).setFontColor('#4f7ef8').setFontStyle('italic');

  const idxSep = [], idxTit = [];
  linhas.forEach((l, i) => {
    const v = String(l[0]);
    if (v.startsWith('━')) idxSep.push(i + 1);
    else if (v.match(/^[📌🗂⚙️🔄🗺🔑📞]/u)) idxTit.push(i + 1);
  });
  idxSep.forEach(ln => sh.getRange(ln, 1).setFontColor('#2a3050').setFontSize(9));
  idxTit.forEach(ln => sh.getRange(ln, 1).setFontSize(12).setFontWeight('bold').setFontColor('#4f7ef8'));

  sh.setColumnWidth(1, 720);
  _trim(sh, nRows, 1);
}
// ── TABELA ────────────────────────────────────────────────────
function _gravarTabela(ss) {
  const painel=ss.getSheetByName('PAINEL'); if(!painel) return;
  const old=ss.getSheetByName('📋 TABELA'); if(old) ss.deleteSheet(old);
  const sh=ss.insertSheet('📋 TABELA');
  const data=painel.getDataRange().getValues();
  if(data.length<2) return;
  const hdr=data[0].map(String);
  sh.getRange(1,1,1,hdr.length).setValues([hdr])
    .setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1,30); sh.setFrozenRows(1); sh.setFrozenColumns(2);
  const rows=data.slice(1).filter(r=>r[0]!==undefined&&r[0]!==''&&!String(r[0]).startsWith('⏱'));
  if(!rows.length) return;
  sh.getRange(2,1,rows.length,hdr.length).setValues(rows).setFontSize(10);
  for(let i=0;i<rows.length;i++){
    sh.getRange(i+2,1,1,hdr.length)
      .setBackground(i%2===0?'#f0f4ff':'#fff')
      .setVerticalAlignment('middle');
  }
  _alinharColunas(sh, hdr, rows.length+1);
  _wrapColunas(sh, hdr, rows.length+1);
  [80,340,115,115,120,115,115,80,320,115,115,155,155,55,100,90,90].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  try { const f=sh.getFilter(); if(f) f.remove(); } catch(e) {}
  try { sh.getRange(1,1,rows.length+1,hdr.length).createFilter(); } catch(e) {}
  _protegerCabecalho(sh);
  _trim(sh,rows.length+1,hdr.length);
}

// ── INCORRETOS ────────────────────────────────────────────────
function _gravarIncorretos(ss, incorretos) {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  const nome = '⚠️ INCORRETOS';
  const old = ss.getSheetByName(nome); if (old) ss.deleteSheet(old);
  if (!incorretos.length) return;
  const sh = ss.insertSheet(nome); sh.setTabColor('#DC2626');
  const H = ['KEY','TIPO','SUMMARY','STATUS','RESPONSÁVEL','ATUALIZADO','LINK JIRA'];
  const rows = [H];
  incorretos.forEach(i => {
    const f = i.fields;
    rows.push([i.key,(f.issuetype||{}).name||'',f.summary||'',(f.status||{}).name||'',
      (f.assignee||{}).displayName||'',_fmt(f.updated),JIRA_BASE + '/browse/' + i.key]);
  });
  sh.getRange(1,1,rows.length,H.length).setValues(rows);
  _estilizarSheet(sh, rows.length, H.length, true);
  [80,100,340,120,155,90,300].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(1);
  _alinharColunas(sh, H, rows.length);
  _wrapColunas(sh, H, rows.length);
  try { sh.getRange(1,1,rows.length,H.length).createFilter(); } catch(e) {}
  _protegerCabecalho(sh);
  _trim(sh, rows.length, H.length);
  Logger.log('INCORRETOS: ' + incorretos.length + ' issues com tipo inválido.');
}

// ── Helpers de estilo ─────────────────────────────────────────
function _estilizarSheet(sh, nRows, nCols, isNova) {
  sh.getRange(1,1,1,nCols)
    .setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(10);
  sh.setRowHeight(1,30);
  for(let i=2;i<=nRows;i++){
    sh.getRange(i,1,1,nCols)
      .setVerticalAlignment('middle').setFontSize(10)
      .setBackground(i%2===0?'#f0f4ff':'#fff');
    if(isNova) sh.setRowHeight(i,24);
  }
}

function _sinalizar(sh, nRows, hdr, colNames) {
  if (nRows < 2) return;
  colNames.forEach(nome => {
    const col = hdr.indexOf(nome);
    if (col < 0) return;
    const vals = sh.getRange(2, col+1, nRows-1, 1).getValues();
    vals.forEach((row, i) => {
      if (!row[0] || String(row[0]).trim() === '') {
        sh.getRange(i+2, col+1)
          .setValue('⚠️ vazio')
          .setBackground('#fff3cd')
          .setFontColor('#856404')
          .setHorizontalAlignment('center');
      }
    });
  });
}

function _trim(sh, usedRows, usedCols) {
  try {
    const mr=sh.getMaxRows(), mc=sh.getMaxColumns();
    if(mr>usedRows+1) sh.deleteRows(usedRows+2, mr-usedRows-1);
    if(mc>usedCols)   sh.deleteColumns(usedCols+1, mc-usedCols);
  } catch(e) {}
}
