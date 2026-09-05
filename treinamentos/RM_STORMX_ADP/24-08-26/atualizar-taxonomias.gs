/**
 * DICIONÁRIO DE TAXONOMIAS UL — Corp (StormX)
 * - Mantém IMPORTRANGE na aba automaticamente (recria se deletado)
 * - Lê dados e commita no GitHub
 * - Planilha fica com dados visíveis para uso humano
 */

const GITHUB_TOKEN = 'COLE_SEU_TOKEN_AQUI';
const SOURCE_ID    = '1qIJIAz8UnYxHsRk1I5eRl1S9oPbewJhi7l7PjDnvgs0';
const ABA          = 'Galielo e Freetext';

const TARGETS = [
  { owner: 'falssp', repo: 'nc-tool',    path: 'treinamentos/RM_STORMX_ADP/24-08-26/taxonomias-ul.html' },
  { owner: 'falssp', repo: 'Projeto-NC', path: 'Treinamentos/RM_STORMX_ADP/24-08-26/taxonomias-ul.html' },
];

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NC Tool')
    .addItem('▶ Atualizar dicionário de taxonomias', 'atualizarTaxonomias')
    .addToUi();
  // Garante que a aba existe e está protegida ao abrir
  try { garantirImportRange(); } catch(e) { Logger.log('onOpen: ' + e.message); }
}

// ── GARANTIR IMPORTRANGE ──────────────────────────────────────
// Cria ou recria a aba com IMPORTRANGE se estiver ausente ou vazia.
// Chamada automaticamente antes de cada atualização.
function garantirImportRange() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let aba   = ss.getSheetByName(ABA);

  // Cria a aba se não existir
  if(!aba) {
    aba = ss.insertSheet(ABA);
    Logger.log('Aba "' + ABA + '" criada.');
  }

  // Verifica se A1 tem o IMPORTRANGE correto
  const a1    = aba.getRange('A1');
  const atual = a1.getFormula();
  const esperada = '=IMPORTRANGE("' + SOURCE_ID + '";"' + ABA + '!A:G")';

  if(atual.replace(/\s/g,'').toLowerCase() !== esperada.replace(/\s/g,'').toLowerCase()) {
    a1.setFormula(esperada);
    SpreadsheetApp.flush();
    Logger.log('IMPORTRANGE configurado.');

    // Aguarda carregar (pode precisar de autorização manual na primeira vez)
    Utilities.sleep(3000);

    // Tenta autorizar programaticamente
    try {
      const token = ScriptApp.getOAuthToken();
      UrlFetchApp.fetch(
        'https://docs.google.com/spreadsheets/d/' + ss.getId() +
        '/externaldata/addimportrangepermissions?key=' + SOURCE_ID,
        { method: 'POST', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
      );
    } catch(e) { Logger.log('Autorização manual pode ser necessária: ' + e.message); }

    SpreadsheetApp.flush();
    Utilities.sleep(2000);
  }

  // Verifica se dados carregaram
  const val = aba.getRange('A2').getDisplayValue();
  if(!val || val === '#ERROR!' || val === 'Carregando…') {
    throw new Error(
      'IMPORTRANGE ainda não autorizado.\n\n' +
      '1. Clique na célula A1 da aba "' + ABA + '"\n' +
      '2. Clique em "Permitir acesso"\n' +
      '3. Aguarde carregar e tente novamente.'
    );
  }

  // Oculta a aba (invisível para usuários comuns)
  aba.hideSheet();

  // Protege a aba: bloqueia edição e deleção para todos exceto o dono do script
  const protecoes = aba.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if(protecoes.length === 0) {
    const prot = aba.protect();
    prot.setDescription('Dados de taxonomia — gerenciado pelo NC Tool');
    // Remove todos os editores exceto o dono
    const me = Session.getEffectiveUser();
    prot.addEditor(me);
    prot.removeEditors(prot.getEditors().filter(e => e.getEmail() !== me.getEmail()));
    // Se for planilha de domínio, bloqueia edição para o domínio também
    if(prot.canDomainEdit()) prot.setDomainEdit(false);
    Logger.log('Aba protegida e oculta.');
  }

  return aba;
}

// ── WEBAPP ───────────────────────────────────────────────────
function doGet() {
  try {
    const dados     = lerAba(garantirImportRange());
    const htmlBase  = buscarHTMLBase();
    const htmlFinal = injetarDATA(htmlBase, JSON.stringify(dados, null, 2));
    return HtmlService.createHtmlOutput(htmlFinal)
      .setTitle('Dicionário de Taxonomias UL')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch(e) {
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:2rem;color:#c00">Erro: ' + e.message + '</p>'
    );
  }
}

// ── ATUALIZAR ─────────────────────────────────────────────────
function atualizarTaxonomias() {
  const ui = SpreadsheetApp.getUi();
  try {
    const aba       = garantirImportRange();
    const dados     = lerAba(aba);
    const htmlBase  = buscarHTMLBase();
    const htmlFinal = injetarDATA(htmlBase, JSON.stringify(dados, null, 2));

    const erros = [];
    TARGETS.forEach(t => {
      try { commitGitHub(t.owner, t.repo, t.path, htmlFinal); }
      catch(e) { erros.push(t.repo + ': ' + e.message); }
    });

    const msg = erros.length
      ? '⚠ Erros:\n' + erros.join('\n')
      : '✓ Dicionário atualizado!\n\nhttps://falssp.github.io/nc-tool/treinamentos/RM_STORMX_ADP/24-08-26/taxonomias-ul.html';
    ui.alert('NC Tool', msg, ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
    Logger.log('ERRO: ' + e.message + '\n' + e.stack);
  }
}

// ── LER ABA ──────────────────────────────────────────────────
function lerAba(aba) {
  const rows = aba.getDataRange().getValues();
  if(rows.length < 2) throw new Error('Aba vazia. Verifique o IMPORTRANGE.');

  const header = rows[0].map(h => h.toString().toLowerCase().trim());
  function col(kw) { return header.findIndex(h => h.includes(kw)); }

  const iCampo = col('campo'), iDesc = col('descri'), iOpt = col('op');
  const iAbr = col('abrev'), iRel = col('rela'), iSig = col('signif'), iPlat = col('plataforma');

  if(iCampo < 0 || iOpt < 0 || iSig < 0)
    throw new Error('Colunas obrigatórias não encontradas.\nCabeçalho: ' + header.join(' | '));

  const blocos = [];
  let campoAtual = null;

  for(let i = 1; i < rows.length; i++) {
    const r     = rows[i];
    const campo = String(r[iCampo] || '').trim();
    const opt   = String(r[iOpt]   || '').trim();
    const sig   = String(r[iSig]   || '').trim();
    const plat  = iPlat >= 0 ? String(r[iPlat] || '').trim() : '';
    const abr   = iAbr  >= 0 ? String(r[iAbr]  || '').trim() : '';
    const rel   = iRel  >= 0 ? String(r[iRel]  || '').trim() : '';
    const desc  = iDesc >= 0 ? String(r[iDesc]  || '').trim() : '';

    if(!opt && !sig) continue;

    if(campo && campo !== campoAtual) {
      campoAtual = campo;
      blocos.push({ id: slugify(campo), field: campo, desc: desc, items: [] });
    }

    if(blocos.length && opt) {
      const item = { opt, sig, plat };
      if(abr) item.abr = abr;
      if(rel) item.rel = rel;
      blocos[blocos.length - 1].items.push(item);
    }
  }

  if(!blocos.length) throw new Error('Nenhum dado processado. Verifique o IMPORTRANGE.');
  Logger.log(blocos.length + ' campos, ' + blocos.reduce((a,b) => a + b.items.length, 0) + ' opções.');
  return blocos;
}

function slugify(s) {
  return s.toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[^a-z0-9\s-]/g,'').trim().replace(/\s+/g,'-');
}

// ── HTML ─────────────────────────────────────────────────────
function buscarHTMLBase() {
  const t   = TARGETS[0];
  const url = 'https://api.github.com/repos/' + t.owner + '/' + t.repo + '/contents/' + t.path;
  const r   = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'token ' + GITHUB_TOKEN, Accept: 'application/vnd.github.v3+json' },
    muteHttpExceptions: true
  });
  if(r.getResponseCode() !== 200)
    throw new Error('HTML base não encontrado no GitHub (' + r.getResponseCode() + ').');
  const json = JSON.parse(r.getContentText());
  return Utilities.newBlob(Utilities.base64Decode(json.content.replace(/\n/g,''))).getDataAsString();
}

function injetarDATA(html, dataJS) {
  const re = /const DATA\s*=\s*\[[\s\S]*?\];\s*\n/;
  if(!re.test(html)) throw new Error('"const DATA = [...]" não encontrado no HTML base.');
  return html.replace(re, 'const DATA = ' + dataJS + ';\n');
}

// ── GITHUB ───────────────────────────────────────────────────
function commitGitHub(owner, repo, path, content) {
  const base    = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + path;
  const headers = {
    Authorization:  'token ' + GITHUB_TOKEN,
    Accept:         'application/vnd.github.v3+json',
    'Content-Type': 'application/json'
  };

  const get = UrlFetchApp.fetch(base, { headers, muteHttpExceptions: true });
  const sha = get.getResponseCode() === 200 ? JSON.parse(get.getContentText()).sha : null;

  const body = {
    message: 'atualiza dicionário de taxonomias UL [GAS ' + new Date().toISOString().slice(0,10) + ']',
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch:  'main'
  };
  if(sha) body.sha = sha;

  const put  = UrlFetchApp.fetch(base, {
    method: 'PUT', headers, payload: JSON.stringify(body), muteHttpExceptions: true
  });
  const code = put.getResponseCode();
  if(code !== 200 && code !== 201)
    throw new Error('GitHub ' + code + ': ' + put.getContentText().substring(0,200));
}
