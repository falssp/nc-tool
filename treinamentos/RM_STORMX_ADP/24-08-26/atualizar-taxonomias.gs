/**
 * DICIONÁRIO DE TAXONOMIAS UL — Corp (StormX)
 * - Sem menu visível — tudo via trigger mensal
 * - Planilha protegida contra edição
 * - Lê dados via openById e commita no GitHub
 *
 * SETUP (uma única vez no editor do Apps Script):
 *   1. Cole este código
 *   2. Execute: setup()
 *   3. Pronto — trigger mensal e proteção configurados automaticamente
 */

const GITHUB_TOKEN = 'COLE_SEU_TOKEN_AQUI';
const SOURCE_ID    = '1qIJIAz8UnYxHsRk1I5eRl1S9oPbewJhi7l7PjDnvgs0';
const ABA          = 'Galielo e Freetext';

const TARGETS = [
  { owner: 'falssp', repo: 'nc-tool',    path: 'treinamentos/RM_STORMX_ADP/24-08-26/taxonomias-ul.html' },
  { owner: 'falssp', repo: 'Projeto-NC', path: 'Treinamentos/RM_STORMX_ADP/24-08-26/taxonomias-ul.html' },
];

// ── SETUP (rodar uma única vez) ───────────────────────────────
function setup() {
  configurarTrigger();
  protegerPlanilha();
  Logger.log('✓ Setup concluído — trigger mensal e proteção ativos.');
}

// ── TRIGGER MENSAL ────────────────────────────────────────────
function configurarTrigger() {
  // Remove triggers existentes para evitar duplicatas
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'atualizarTaxonomias')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Cria trigger: todo dia 1 do mês entre 08:00 e 09:00
  ScriptApp.newTrigger('atualizarTaxonomias')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  Logger.log('✓ Trigger mensal configurado — todo dia 1 às 08h.');
}

// ── PROTEGER PLANILHA ─────────────────────────────────────────
function protegerPlanilha() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const me  = Session.getEffectiveUser();

  ss.getSheets().forEach(sheet => {
    // Remove proteções existentes
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .forEach(p => p.remove());

    // Cria nova proteção
    const prot = sheet.protect();
    prot.setDescription('Protegido pelo NC Tool — não editar manualmente');
    prot.removeEditors(prot.getEditors());
    prot.addEditor(me);
    if(prot.canDomainEdit()) prot.setDomainEdit(false);
  });

  Logger.log('✓ Todas as abas protegidas contra edição.');
}

// ── WEBAPP ───────────────────────────────────────────────────
function doGet() {
  try {
    const dados     = lerAba();
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

// ── ATUALIZAR (chamada pelo trigger ou manualmente) ───────────
function atualizarTaxonomias() {
  try {
    const dados     = lerAba();
    const htmlBase  = buscarHTMLBase();
    const htmlFinal = injetarDATA(htmlBase, JSON.stringify(dados, null, 2));

    TARGETS.forEach(t => commitGitHub(t.owner, t.repo, t.path, htmlFinal));
    Logger.log('✓ Dicionário atualizado em ' + new Date().toISOString());
  } catch(e) {
    Logger.log('ERRO: ' + e.message + '\n' + e.stack);
    // Envia email de erro para o dono do script
    MailApp.sendEmail(
      Session.getEffectiveUser().getEmail(),
      '[NC Tool] Erro na atualização do dicionário',
      'Erro em ' + new Date().toLocaleString('pt-BR') + ':\n\n' + e.message
    );
  }
}

// ── LER ABA DA PLANILHA ORIGINAL ─────────────────────────────
function lerAba() {
  let ss;
  try {
    ss = SpreadsheetApp.openById(SOURCE_ID);
  } catch(e) {
    throw new Error('Sem acesso à planilha original: ' + e.message);
  }

  const aba = ss.getSheetByName(ABA);
  if(!aba) throw new Error('Aba "' + ABA + '" não encontrada.');

  const rows = aba.getDataRange().getValues();
  if(rows.length < 2) throw new Error('Aba vazia.');

  const header = rows[0].map(h => h.toString().toLowerCase().trim());
  function col(kw) { return header.findIndex(h => h.includes(kw)); }

  const iCampo = col('campo'), iDesc = col('descri'), iOpt = col('op');
  const iAbr = col('abrev'), iRel = col('rela'), iSig = col('signif'), iPlat = col('plataforma');

  if(iCampo < 0 || iOpt < 0 || iSig < 0)
    throw new Error('Colunas obrigatórias não encontradas. Cabeçalho: ' + header.join(' | '));

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

  if(!blocos.length) throw new Error('Nenhum dado processado.');
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
  // Procura o início do bloco DATA
  const startMarker = 'const DATA = ';
  const startIdx = html.indexOf(startMarker);
  if(startIdx === -1) throw new Error('"const DATA = [...]" não encontrado no HTML base. Tamanho: ' + html.length);

  // Procura o fechamento do array
  let depth = 0, endIdx = -1;
  for(let i = startIdx + startMarker.length; i < html.length; i++) {
    if(html[i] === '[') depth++;
    else if(html[i] === ']') {
      depth--;
      if(depth === 0) { endIdx = i + 1; break; }
    }
  }
  if(endIdx === -1) throw new Error('Fim do bloco DATA não encontrado.');

  // Pula o ; e possível \n
  while(endIdx < html.length && (html[endIdx] === ';' || html[endIdx] === '\n' || html[endIdx] === '\r')) endIdx++;

  return html.substring(0, startIdx) + 'const DATA = ' + dataJS + ';\n' + html.substring(endIdx);
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

  const put = UrlFetchApp.fetch(base, {
    method: 'PUT', headers, payload: JSON.stringify(body), muteHttpExceptions: true
  });
  const code = put.getResponseCode();
  if(code !== 200 && code !== 201)
    throw new Error('GitHub ' + code + ': ' + put.getContentText().substring(0,200));
}
