/**
 * DICIONÁRIO DE TAXONOMIAS UL — Atualizador automático
 * Google Apps Script
 *
 * Fluxo:
 *   1. Lê a planilha de taxonomias (Sheets da UL)
 *   2. Gera o bloco DATA[] atualizado
 *   3. Busca o SHA atual do arquivo no GitHub
 *   4. Faz commit do HTML atualizado nos dois repositórios
 *
 * CONFIGURAÇÃO:
 *   - Substitua GITHUB_TOKEN pelo seu Personal Access Token (scope: repo)
 *   - Verifique SHEET_TAB se a aba mudar de nome
 */

const GITHUB_TOKEN = 'SEU_TOKEN_AQUI';  // GitHub PAT (repo scope)
const SHEET_ID     = '1qIJIAz8UnYxHsRk1I5eRl1S9oPbewJhi7l7PjDnvgs0';
const SHEET_TAB    = 'Plano';           // ajuste se a aba tiver outro nome

const TARGETS = [
  { owner: 'falssp', repo: 'nc-tool',    path: 'treinamentos/RM_STORMX_ADP/24-08-26/taxonomias-ul.html' },
  { owner: 'falssp', repo: 'Projeto-NC', path: 'Treinamentos/RM_STORMX_ADP/24-08-26/taxonomias-ul.html' },
];

// ── MENU NA PLANILHA ──────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NC Tool')
    .addItem('Atualizar dicionário de taxonomias', 'atualizarTaxonomias')
    .addToUi();
}

// ── PONTO DE ENTRADA ──────────────────────────────────────────
function atualizarTaxonomias() {
  try {
    Logger.log('▶ Lendo planilha…');
    const dataJS = lerPlanilha();

    Logger.log('▶ Buscando HTML base no GitHub…');
    const htmlBase = buscarHTMLBase();

    Logger.log('▶ Injetando dados atualizados…');
    const htmlAtualizado = injetarDATA(htmlBase, dataJS);

    Logger.log('▶ Fazendo commit nos repositórios…');
    const erros = [];
    TARGETS.forEach(t => {
      try {
        commitGitHub(t.owner, t.repo, t.path, htmlAtualizado);
        Logger.log('  ✓ ' + t.owner + '/' + t.repo);
      } catch(e) {
        Logger.log('  ✗ ' + t.repo + ': ' + e.message);
        erros.push(t.repo + ': ' + e.message);
      }
    });

    const msg = erros.length
      ? 'Commit com erros:\n' + erros.join('\n')
      : '✓ Dicionário atualizado nos dois repositórios!';
    SpreadsheetApp.getUi().alert(msg);

  } catch(e) {
    SpreadsheetApp.getUi().alert('Erro: ' + e.message);
    Logger.log('ERRO: ' + e.message + '\n' + e.stack);
  }
}

// ── LER PLANILHA ──────────────────────────────────────────────
function lerPlanilha() {
  const ss  = SpreadsheetApp.openById(SHEET_ID);
  const aba = ss.getSheetByName(SHEET_TAB);
  if(!aba) throw new Error('Aba "' + SHEET_TAB + '" não encontrada.');

  const rows   = aba.getDataRange().getValues();
  const header = rows[0].map(h => h.toString().toLowerCase());

  function col(kw) {
    const i = header.findIndex(h => h.includes(kw.toLowerCase()));
    return i >= 0 ? i : -1;
  }

  const iCampo = col('campo');
  const iDesc  = col('descri');
  const iOpt   = col('op');
  const iAbr   = col('abrev');
  const iRel   = col('rela');
  const iSig   = col('signif');
  const iPlat  = col('plataforma');

  if([iCampo, iOpt, iSig].some(x => x < 0))
    throw new Error('Colunas obrigatórias não encontradas. Verifique o cabeçalho da aba.');

  const blocos = [];
  let campoAtual = null;

  for(let i = 1; i < rows.length; i++) {
    const r     = rows[i];
    const campo = (r[iCampo] || '').toString().trim();
    const opt   = (r[iOpt]   || '').toString().trim();
    const sig   = iSig  >= 0 ? (r[iSig]  || '').toString().trim() : '';
    const plat  = iPlat >= 0 ? (r[iPlat] || '').toString().trim() : '';
    const abr   = iAbr  >= 0 ? (r[iAbr]  || '').toString().trim() : '';
    const rel   = iRel  >= 0 ? (r[iRel]  || '').toString().trim() : '';

    if(!opt && !sig) continue;

    if(campo && campo !== campoAtual) {
      campoAtual = campo;
      blocos.push({
        id:    slugify(campo),
        field: campo,
        desc:  iDesc >= 0 ? (r[iDesc] || '').toString().trim() : '',
        items: []
      });
    }

    if(blocos.length && opt) {
      const item = { opt, sig, plat };
      if(abr) item.abr = abr;
      if(rel) item.rel = rel;
      blocos[blocos.length - 1].items.push(item);
    }
  }

  if(!blocos.length) throw new Error('Nenhum dado lido. Verifique a aba e as colunas.');
  return JSON.stringify(blocos, null, 2);
}

function slugify(s) {
  return s.toLowerCase().replace(/[^a-z0-9\s-]/g,'').trim().replace(/\s+/g,'-');
}

// ── BUSCAR HTML BASE ──────────────────────────────────────────
function buscarHTMLBase() {
  const t   = TARGETS[0];
  const url = 'https://api.github.com/repos/' + t.owner + '/' + t.repo + '/contents/' + t.path;
  const r   = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'token ' + GITHUB_TOKEN, Accept: 'application/vnd.github.v3+json' },
    muteHttpExceptions: true
  });
  if(r.getResponseCode() !== 200)
    throw new Error('Arquivo HTML não encontrado no GitHub. Faça o primeiro deploy manualmente.');
  const json = JSON.parse(r.getContentText());
  return Utilities.newBlob(Utilities.base64Decode(json.content.replace(/\n/g,''))).getDataAsString();
}

// ── INJETAR DATA ──────────────────────────────────────────────
function injetarDATA(html, dataJS) {
  const re = /const DATA\s*=\s*\[[\s\S]*?\];\s*
/;
  if(!re.test(html)) throw new Error('Marcador "const DATA = [...]" não encontrado no HTML.');
  return html.replace(re, 'const DATA = ' + dataJS + ';
');
}

// ── COMMIT GITHUB ─────────────────────────────────────────────
function commitGitHub(owner, repo, path, content) {
  const base    = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + path;
  const headers = {
    Authorization:  'token ' + GITHUB_TOKEN,
    Accept:         'application/vnd.github.v3+json',
    'Content-Type': 'application/json'
  };

  // Busca SHA atual
  const get = UrlFetchApp.fetch(base, { headers, muteHttpExceptions: true });
  const sha = get.getResponseCode() === 200 ? JSON.parse(get.getContentText()).sha : null;

  const body = {
    message: 'atualiza dicionário de taxonomias [GAS ' + new Date().toISOString() + ']',
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch:  'main'
  };
  if(sha) body.sha = sha;

  const put  = UrlFetchApp.fetch(base, { method: 'PUT', headers, payload: JSON.stringify(body), muteHttpExceptions: true });
  const code = put.getResponseCode();
  if(code !== 200 && code !== 201)
    throw new Error('GitHub API ' + code + ': ' + put.getContentText().substring(0, 150));
}
