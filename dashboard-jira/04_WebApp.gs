// ============================================================
// [UL] Dashboard_Taxonomias — 04_WebApp.gs
// Dados para o painel HTML e rotas do Web App
// ============================================================

function getDados() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PAINEL');
  if (!sh) return { erro: 'Aba PAINEL não encontrada. Execute a sincronização primeiro.' };
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return { erro: 'PAINEL vazio.' };
  const vals = sh.getRange(1,1,lastRow,lastCol).getValues();
  const hdr = vals[0].map(String);
  const rows = [];
  const iPaiKey = hdr.indexOf('PAI_KEY');
  const iFilhoKey = hdr.indexOf('FILHO_KEY');
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    const paiKey   = iPaiKey   >= 0 ? String(r[iPaiKey]   || '') : '';
    const filhoKey = iFilhoKey >= 0 ? String(r[iFilhoKey] || '') : '';
    if (!paiKey && !filhoKey) continue;
    if (paiKey.startsWith('⏱')) continue;
    const o = {};
    hdr.forEach((h,j) => { o[h] = r[j] !== null && r[j] !== undefined ? r[j] : ''; });
    rows.push(o);
  }
  const shInc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⚠️ INCORRETOS');
  const nIncorretos = shInc && shInc.getLastRow() > 1 ? shInc.getLastRow() - 1 : 0;
  // Calcula total Jira direto das abas RAW (mais confiável que SYNC_TOTAL)
  const ss2 = SpreadsheetApp.getActiveSpreadsheet();
  const rawP2 = ss2.getSheetByName('RAW_PAI');
  const rawF2 = ss2.getSheetByName('RAW_FILHO');
  const rawI2 = ss2.getSheetByName('⚠️ INCORRETOS');
  const nP2 = rawP2 && rawP2.getLastRow() > 1 ? rawP2.getLastRow() - 1 : 0;
  const nF2 = rawF2 && rawF2.getLastRow() > 1 ? rawF2.getLastRow() - 1 : 0;
  const nI2 = rawI2 && rawI2.getLastRow() > 1 ? rawI2.getLastRow() - 1 : 0;
  const totalJira = nP2 + nF2 + nI2 || null;

  return {
    sync: new Date().toISOString(),
    total: rows.length,
    totalJira: totalJira,
    totalIncorretos: nIncorretos,
    avisoToken: getAvisoToken(),
    rows
  };
}

function getIncorretos() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⚠️ INCORRETOS');
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(String);
  return vals.slice(1).map(r => {
    const o = {};
    hdr.forEach((h,j) => { o[h] = r[j] !== null && r[j] !== undefined ? r[j] : ''; });
    return o;
  });
}

function getDadosPeriodo(de, ate, campo) {
  const jqlParts = [];
  if (de)  jqlParts.push(campo + ' >= "' + de + '"');
  if (ate) jqlParts.push(campo + ' <= "' + ate + '"');
  const issues = _buscar(jqlParts.join(' AND '));
  if (!issues.length) return { sync: new Date().toISOString(), total: 0, rows: [] };
  const pais = {}, filhos = [];
  issues.forEach(i => {
    if ((i.fields.issuetype||{}).name==='Tarefa'||!i.fields.parent) pais[i.key]=i;
    else filhos.push(i);
  });
  filhos.forEach(f => {
    const pk=(f.fields.parent||{}).key;
    if(pk&&!pais[pk]) pais[pk]=_stubPai(pk,(f.fields.parent.fields||{}).summary||pk);
  });
  const dv=_carregarDePara('VEICULO');
  const H=['PAI_KEY','PAI_SUMMARY','BU','AGÊNCIA','MARCA (JIRA)','CAMP. MÍDIA','STATUS ADP (PAI)',
           'FILHO_KEY','SUMMARY (FILHO)','VEÍCULO','STATUS','RESPONSÁVEL','RELATOR',
           'QTD TAX.','PLATAFORMA','DATA LIMITE','ÚLT. ATU.'];
  const rows=[];
  const mapaPai={};
  Object.values(pais).forEach(p=>{
    const c=_cx(p.fields);
    mapaPai[p.key]={key:p.key,summary:p.fields.summary||'',bu:c.bu,agencia:c.agencia,
      marcaJira:c.marcaJira,campMidia:c.campMidia,statusAdp:c.statusAdp,
      assignee:(p.fields.assignee||{}).displayName||'',relator:c.relator,
      duedate:c.duedate,plataforma:c.plataforma,qtdTax:c.qtdTax,
      status:(p.fields.status||{}).name||'',updated:_fmt(p.fields.updated)};
  });
  const porPai={};
  filhos.forEach(f=>{const pk=(f.fields.parent||{}).key||'SEM_PAI';if(!porPai[pk])porPai[pk]=[];porPai[pk].push(f);});
  Object.keys(mapaPai).sort().forEach(pk=>{
    const pai=mapaPai[pk],fd=porPai[pk]||[];
    if(!fd.length){
      const o={};H.forEach((h,i)=>o[h]=[pai.key,pai.summary,pai.bu||'⚠️ vazio',pai.agencia||'⚠️ vazio',
        pai.marcaJira||'⚠️ vazio',pai.campMidia||'',pai.statusAdp||'⚠️ vazio',
        '','(sem subtarefas)','',pai.status,pai.assignee,pai.relator,
        pai.qtdTax||'',pai.plataforma||'',pai.duedate||'',pai.updated||''][i]);
      rows.push(o);
    } else {
      fd.forEach(f=>{
        const ff=f.fields,cf=_cx(ff),sum=ff.summary||'';
        const vv=_extrairTokens(sum,dv);if(!vv.length)vv.push('⚠️ Não mapeado');
        vv.forEach(vei=>{
          const bu=cf.bu||pai.bu||'⚠️ vazio',ag=cf.agencia||pai.agencia||'⚠️ vazio';
          const mk=cf.marcaJira||pai.marcaJira||'⚠️ vazio',cm=cf.campMidia||pai.campMidia||'';
          const sa=cf.statusAdp||pai.statusAdp||'⚠️ vazio',pl=cf.plataforma||pai.plataforma||'';
          const qt=cf.qtdTax!==''?cf.qtdTax:pai.qtdTax||'',dd=cf.duedate||pai.duedate||'';
          const o={};
          H.forEach((h,i)=>o[h]=[pai.key,pai.summary,bu,ag,mk,cm,sa,f.key,sum,vei,
            (ff.status||{}).name||'',(ff.assignee||{}).displayName||'',cf.relator||'',
            qt,pl,dd,_fmt(ff.updated)][i]);
          rows.push(o);
        });
      });
    }
  });
  return { sync: new Date().toISOString(), total: rows.length, rows };
}

// ── Formatação da planilha ────────────────────────────────────
const ABAS_PERMITIDAS_SET = new Set([
  '🏠 INICIO','PAINEL','📋 TABELA',
  'RAW_PAI','RAW_FILHO','⚠️ INCORRETOS','DE_PARA','_SYNC_BUFFER'
]);

function formatarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove abas não permitidas
  ss.getSheets().forEach(function(sh) {
    const nome = sh.getName();
    if (!ABAS_PERMITIDAS_SET.has(nome) && ss.getSheets().length > 1) {
      try { ss.deleteSheet(sh); Logger.log('Aba extra removida: ' + nome); }
      catch(e) {}
    }
  });

  // Remove aba RAW legada
  try { const r = ss.getSheetByName('RAW'); if (r) ss.deleteSheet(r); } catch(e) {}

  ss.getSheets().forEach(function(sh) {
    const nome = sh.getName();
    if (nome === '🏠 INICIO') return;
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return;
    try {
      const maxR = sh.getMaxRows(), maxC = sh.getMaxColumns();
      if (maxR > lastRow + 1) sh.deleteRows(lastRow + 2, maxR - lastRow - 1);
      if (maxC > lastCol)     sh.deleteColumns(lastCol + 1, maxC - lastCol);
    } catch(e) {}
    // Formatação em batch — sem loop linha a linha
    const range = sh.getRange(1, 1, lastRow, lastCol);
    range.setVerticalAlignment('middle');
    if (lastRow >= 1) {
      sh.getRange(1, 1, 1, lastCol).setHorizontalAlignment('center').setFontWeight('bold');
      sh.setRowHeight(1, 28);
    }
    // Altura das linhas de dados em uma única chamada batch
    if (lastRow > 1) {
      try { sh.setRowHeightsForced(2, lastRow - 1, 22); } catch(e) {
        // fallback se setRowHeightsForced não disponível
        sh.setRowHeight(2, 22);
      }
    }
    Logger.log('Formatando: ' + nome + ' (' + lastRow + ' linhas)');
  });
}

// ── Web App ───────────────────────────────────────────────────
function doGet(e) {
  const p = (e && e.parameter) || {};

  // Status do progresso
  if (p.action === 'status') {
    return ContentService.createTextOutput(JSON.stringify(getProgresso())).setMimeType(ContentService.MimeType.JSON);
  }

  // API de dados — JSONP (GitHub Pages) ou JSON (fetch direto)
  let payload;
  if (p.acao === 'periodo') {
    const dp = getDadosPeriodo(p.de||'', p.ate||'', p.campo||'updated');
    payload  = { dados: dp, periodo: { de:p.de||'', ate:p.ate||'', campo:p.campo||'updated' }, incorretos: getIncorretos() };
  } else {
    payload  = { dados: getDados(), periodo: null, incorretos: getIncorretos() };
  }
  const json = JSON.stringify(payload);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ── POST — exportar via fetch ────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const result = exportarDados(body);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,erro:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}
