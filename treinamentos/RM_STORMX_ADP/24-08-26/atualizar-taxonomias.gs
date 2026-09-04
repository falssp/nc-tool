/**
 * DICIONÁRIO DE TAXONOMIAS UL — Atualizador v2
 * Cole este script em uma planilha PRIVADA @stormx.com.br
 *
 * SETUP (uma única vez):
 *   1. Crie uma planilha privada em sua conta @stormx.com.br
 *   2. Crie uma aba chamada "Galileo e Freetext"
 *   3. Na célula A1 cole:
 *      =IMPORTRANGE("1qIJIAz8UnYxHsRk1I5eRl1S9oPbewJhi7l7PjDnvgs0","Galileo e Freetext!A:G")
 *   4. Autorize o IMPORTRANGE quando solicitado
 *   5. Cole este script em Extensões → Apps Script
 *   6. Preencha GITHUB_TOKEN abaixo
 *   7. Execute atualizarTaxonomias() para autorizar e testar
 */

const GITHUB_TOKEN = 'COLE_SEU_TOKEN_AQUI';  // Gere em: github.com/settings/tokens (scope: repo)

const ABA = 'Galileo e Freetext';  // nome da aba na plan privada

const TARGETS = [
  { owner: 'falssp', repo: 'nc-tool',    path: 'treinamentos/RM_STORMX_ADP/24-08-26/taxonomias-ul.html' },
  { owner: 'falssp', repo: 'Projeto-NC', path: 'Treinamentos/RM_STORMX_ADP/24-08-26/taxonomias-ul.html' },
];

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NC Tool')
    .addItem('▶ Atualizar dicionário de taxonomias', 'atualizarTaxonomias')
    .addItem('⚙ Verificar conexão com GitHub', 'verificarGitHub')
    .addToUi();
}

// ── PONTO DE ENTRADA ─────────────────────────────────────────
function atualizarTaxonomias() {
  const ui = SpreadsheetApp.getUi();
  try {
    Logger.log('▶ Lendo aba "' + ABA + '"…');
    const dados = lerAba(ABA);

    Logger.log('▶ ' + dados.length + ' campos. Buscando HTML base…');
    const htmlBase = buscarHTMLBase();

    Logger.log('▶ Injetando dados…');
    const htmlFinal = injetarDATA(htmlBase, JSON.stringify(dados, null, 2));

    Logger.log('▶ Fazendo commit…');
    const erros = [];
    TARGETS.forEach(t => {
      try {
        commitGitHub(t.owner, t.repo, t.path, htmlFinal);
        Logger.log('  ✓ ' + t.owner + '/' + t.repo);
      } catch(e) {
        Logger.log('  ✗ ' + t.repo + ': ' + e.message);
        erros.push(t.repo + ': ' + e.message);
      }
    });

    const msg = erros.length
      ? '⚠ Commit com erros:\n' + erros.join('\n')
      : '✓ Dicionário atualizado!\n\nhttps://falssp.github.io/nc-tool/treinamentos/RM_STORMX_ADP/24-08-26/taxonomias-ul.html';
    ui.alert('NC Tool', msg, ui.ButtonSet.OK);

  } catch(e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
    Logger.log('ERRO: ' + e.message + '\n' + e.stack);
  }
}

function verificarGitHub() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = UrlFetchApp.fetch('https://api.github.com/user', {
      headers: { Authorization: 'token ' + GITHUB_TOKEN, Accept: 'application/vnd.github.v3+json' },
      muteHttpExceptions: true
    });
    const d = JSON.parse(r.getContentText());
    if(r.getResponseCode() === 200) {
      ui.alert('GitHub OK', 'Conectado como: ' + d.login, ui.ButtonSet.OK);
    } else {
      ui.alert('GitHub ERRO', 'HTTP ' + r.getResponseCode() + ': ' + d.message, ui.ButtonSet.OK);
    }
  } catch(e) {
    ui.alert('Erro', e.message, ui.ButtonSet.OK);
  }
}

// ── LER ABA ──────────────────────────────────────────────────
function lerAba(nomeAba) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nomeAba);
  if(!aba) throw new Error(
    'Aba "' + nomeAba + '" não encontrada.\n\n' +
    'Crie a aba com esse nome exato e cole na A1:\n' +
    '=IMPORTRANGE("1qIJIAz8UnYxHsRk1I5eRl1S9oPbewJhi7l7PjDnvgs0","Galileo e Freetext!A:G")'
  );

  const rows = aba.getDataRange().getValues();
  if(rows.length < 2) throw new Error('Aba "' + nomeAba + '" vazia. Verifique o IMPORTRANGE.');

  const header = rows[0].map(h => h.toString().toLowerCase().trim());

  function col(kw) {
    const i = header.findIndex(h => h.includes(kw));
    return i >= 0 ? i : -1;
  }

  const iCampo = col('campo');
  const iDesc  = col('descri');
  const iOpt   = col('op');
  const iAbr   = col('abrev');
  const iRel   = col('rela');
  const iSig   = col('signif');
  const iPlat  = col('plataforma');

  if(iCampo < 0 || iOpt < 0 || iSig < 0)
    throw new Error(
      'Colunas obrigatórias não encontradas.\nCabeçalho detectado: ' + header.join(' | ') +
      '\n\nEsperado: Campo | Descrição | Opções | Abreviações | Relação | Significado | Plataforma'
    );

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

  if(!blocos.length) throw new Error('Nenhum dado processado. Verifique se o IMPORTRANGE está carregado.');
  Logger.log('  ' + blocos.length + ' campos lidos, ' + blocos.reduce((a,b) => a + b.items.length, 0) + ' opções no total.');
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
