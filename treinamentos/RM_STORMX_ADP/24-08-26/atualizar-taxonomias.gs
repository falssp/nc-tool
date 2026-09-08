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

  // Aplica siglas do ADP onde estão vazias
  const SIGLAS = {
  'Awareness': 'awa',
  'Awareness/Consideration': 'awacons',
  'Consideration': 'cons',
  'Conversion': 'Conversion',
  'Convert': 'con',
  'Inform': 'Inf',
  'Inspire': 'Ins',
  'Advise': 'adv',
  '1PD-DID': '1pd-did',
  '1PD-DAD': '1pd-dad',
  '1PD-DID+DAD': '1pd-diddad',
  'No-1PD': 'No-1PD',
  'CRM': 'CRM',
  'Behavioural': 'Behav',
  'Lookalike': 'LLike',
  'CombinedAudiences': 'CombAud',
  'DMP': 'DMP',
  'Collected in-platform': 'CollPform',
  'Custom': 'Cust',
  'Demographic': 'Demog',
  'In-Market': 'InMark',
  'Interest/Affinity': 'IntAff',
  'In-Market & Lifestyle': 'InmLst',
  'Lifestyle/Life Events': 'Lstyle',
  'OMP': 'OMP',
  'OMPPMP': 'OMPPMP',
  'PMP': 'PMP',
  'Prefer': 'Prefer',
  'ProgGuar': 'ProgGuar',
  'GooglePref': 'GooglePref',
  'TrVAction': 'TrVAction',
  'TrVAuc': 'TrVAuc',
  'TrVReach': 'TrVReach',
  'TrVInsRes': 'TrVInsRes',
  'OpAuc': 'OpAuc',
  'ReaFreq': 'ReaFreq',
  'Reserv': 'Reserv',
  'ResBuy': 'ResBuy',
  'Auc': 'Auc',
  'Spn': 'Spn',
  'First-Impression': 'First-Impression',
  'CPA': 'CPA',
  'CPI': 'CPI',
  'CPL': 'CPL',
  'CPM': 'CPM',
  'CPD': 'CPD',
  'CPC': 'CPC',
  'CPV': 'CPV',
  'Flat-Fee/Time-Based': 'Flat-Fee/Time-Based',
  'vCPM': 'vCPM',
  'Zero-cost': 'Zero-cost',
  'CPH': 'CPH',
  'Brand': 'Brand',
  'Retailer(via-MikMak)': 'RetMM',
  'D2C': 'D2C',
  'External': 'Ext',
  'None': 'None',
  'Retail': 'restaurant-week',
  'Shoppable-Tool-Provider': 'STP',
  'Social': 'teads',
  'Category': 'Cat',
  'TikTok-Shop': 'TikTok-Shop',
  'Non-shoppable': 'Non-shoppable',
  'na': 'na',
  'Shoppable': 'Shoppable',
  'Adimo': 'Adimo',
  'Cartwire': 'Cartwire',
  'Dotter': 'Dotter',
  'MikMak-–-Shoppable-Recipe': 'MMRec',
  'Chicory-–-Shoppable-Recipe': 'ChicRec',
  'Northfork-–-Shoppable-Recipe': 'NfrkRec',
  'FIU-–-Shoppable-Recipe': 'FIURec',
  'MikMak': 'MikMak',
  'PriceSpider': 'PriceSpider',
  'Shopalyst': 'shopalyst',
  'Whisk': 'Whisk',
  'Retailers cadastrados': 'Retail',
  'All': 'All',
  'Branded Effect': 'BranEffect',
  'Bumper': 'Bump',
  'Browse': 'Browse',
  'Companion Banner': 'CompBan',
  'Content-Suggestion': 'Content-Suggestion',
  'DCO - Display': 'DCODisp',
  'DCO-Video': 'DCO-Video',
  'Discovery': 'Disc',
  'DOOH-Display': 'DOOH-Display',
  'DOOH-Video': 'DOOH-Video',
  'eCommerce-Responsive': 'eCommerce-Responsive',
  'Efficient Reach': 'EfReach',
  'First View': '1stView',
  'Hashtag Challenge': 'TagChal',
  'In-Banner Video': 'InBanVid',
  'In-Feed': 'InFeed',
  'In-Feed Video': 'InFeedVid',
  'In-Read': 'InRead',
  'Instream': 'InStr',
  'Interstitial': 'Interstit',
  'Masthead': 'Mhead',
  'Mixed Outstream': 'MixOut',
  'Multiple': 'Multi',
  'News Feed App Series Placement': 'NewsFeed',
  'Non-Skippable In-Stream': 'NonSkipInStr',
  'NonSkip': 'NSkip',
  'Outstream': 'OutSt',
  'Pangle Placement': 'Pangle',
  'Promoted Spotlight': 'PromSpot',
  'Promoted Trend': 'PromTrend',
  'Pulse': 'Pulse',
  'Rewarded-Display': 'Rewarded-Display',
  'Rewarded-Video': 'Rewarded-Video',
  'Rich-Media-Display': 'Rich-Media-Display',
  'Roadblock/Takeover': 'RbkTkover',
  'Search': 'revista-hoteis',
  'Shorts Video Ads': 'ShortsVid',
  'Skippable In-Stream': 'SkipInStr',
  'Sponsorship': 'Spons',
  'Standard-Display': 'Standard-Display',
  'Static Banner': 'SttBan',
  'TikTok Placement': 'TikTok',
  'TopFeed': 'TopFeed',
  'TopView': 'TopView',
  'Tracking-Placements': 'Tracking-Placements',
  'App-Campaign': 'appcamp',
  'Audio-NonSkippable': 'audnoskp',
  'Audio-Skippable': 'audskp',
  'AudioEverywhere-NonSkippable': 'audevrynoskp',
  'AudioEverywhere-Skippable': 'audevryskp',
  'Binge-Ads-Display': 'bngead',
  'Branded-Buzz': 'bbuzz',
  'Branded-Effect': 'beffect',
  'Branded-Mission': 'bmission',
  'Bumper-NonSkippable': 'bump',
  'Bumper-TargetFrequency-NonSkippable': 'bumptfnonskip',
  'Bumper-TargetFrequency-Skippable': 'bumptfskip',
  'Carousel': 'carousel',
  'Collection': 'collect',
  'Conversational Card': 'concard',
  'CuttingEdge-Display': 'cuttedge',
  'Demand Gen': 'dgen',
  'Destaque-Premium-Globoplay-Display': 'dstqprgplay',
  'Diaria-Display': 'diariadsply',
  'Display-Card': 'dispcard',
  'Dynamic Product Ads': 'dyprod',
  'EyeMax-Skippable': 'eymx',
  'Feed': 'feed',
  'Feed-Display': 'fed-dply',
  'Feed-NonSkippable': 'Feed-NonSkippable',
  'Feed-Stories': 'fed-stor',
  'Feed-Stories-Reels': 'fed-stor-reel',
  'Feed-Video-NonSkippable': 'feednskp',
  'Feed-Video-Skippable': 'feedskp',
  'FirstScreenAdTile-Dynamic': 'fstscradtl',
  'FirstScreenImmersivePlus-Dynamic': 'fstscrimmplus',
  'FirstScreenImmesiveMasthead-Dynamic': 'fstscrimmmast',
  'FrameAds-Globoplay-Display': 'framegplay',
  'Idea Ads': 'idea',
  'Instant Experience': 'instexp',
  'InStream-NonSkippable': 'instm-nskipyt',
  'InStream-OutStream-NonSkippable': 'inoutnoskp',
  'InStream-OutStream-Skippable': 'inoutskp',
  'InStream-Skippable': 'instm-skipyt',
  'Lead Ads': 'lead-ads',
  'Max Width Video': 'maxvid',
  'Messenger Ads': 'msg-ads',
  'Mosaico-Globoplay-Display': 'mscogplay',
  'Native-Display': 'ntvdsply',
  'Native-Video-NonSkippable': 'ntvnoskp',
  'Native-Video-Skippable': 'ntvskp',
  'PauseAds-Globoplay-Display': 'pseadsgplay',
  'PauseAds-Skippable': 'psadskp',
  'Pmax': 'pmax',
  'Premiere Spotlight': 'premsptght',
  'Primeirissima-Globoplay-Skippable': 'primgplay',
  'Product-Spotlight': 'prodspot',
  'Promoted Poll': 'prompoll',
  'Promoted-Video-NonSkippable': 'promvidnonskip',
  'Promoted-Video-Skippable': 'promvidskip',
  'Push-Notification': 'pushntf',
  'Quiz Ads': 'quis-ads',
  'Reels': 'reel',
  'Reels-Stories': 'reels-stor',
  'Rewarded-Video-NonSkippable': 'rwdnoskip',
  'Rewarded-Video-Skippable': 'rwrdskp',
  'RichMedia-Display': 'richdsply',
  'Segundissima-Globoplay-Skippable': 'segungplay',
  'Shopping': 'shop',
  'Sidebar-Contextual': 'sbarcont',
  'Sponsored-Answer-Card': 'spscard',
  'Standard Image': 'standimg',
  'Standard Width Video': 'standvid',
  'Standard-Display-IAB': 'standdisply',
  'Stories': 'stor',
  'Takeover-NonSkippable': 'tkovernonskpi',
  'TikTok-Pulse-Skippable': 'tkpulsesk',
  'TikTok-Shop-Skippable': 'tkshopsk',
  'Top-Feed': 'tfeed',
  'Top-Feed-Skippable': 'tfeed',
  'Top-View-Skippable': 'tview',
  'Trend-Takeover': 'trendtkover',
  'Video-Instream-NonSkippable': 'vid-instm-nskipyt',
  'Video-Instream-Skippable': 'vid-instm-skipyt',
  'Video-NonSkippable': 'vidnonskip',
  'Video-Outstream-NonSkippable': 'outsnsk',
  'Video-Outstream-Skippable': 'outssk',
  'Video-PartnershipAds-NonSkippable': 'prtnradsnskip',
  'Video-PartnershipAds-Skippable': 'prtnradsskip',
  'Video-Skippable': 'vidskip',
  'Video-SponsorshipAds-NonSkippable': 'spnsradsnskip',
  'Video-SponsorshipAds-Skippable': 'spnsradsskip',
  'VRC-NonSkippable': 'vrcnskip',
  'VRC-Shorts-NonSkippable': 'vrcshtnskip',
  'VRC-Shorts-Skippable': 'vrcshtskip',
  'VRC-Skippable': 'vrcskip',
  'VVC-NonSkippable': 'vvcnskip',
  'VVC-Shorts-NonSkippable': 'vvcshtnskip',
  'VVC-Shorts-Skippable': 'vvcshtskip',
  'VVC-Skippable': 'vvcskip',
  'YouTube-Masthead-CPH': 'mstdyt-cph',
  'YouTube-Masthead-Skippable': 'mstdyt',
  'YT-Shorts-NonSkippable': 'ytshortsnskip',
  'YT-Shorts-Skippable': 'ytshorts',
  'Texto': 'Texto',
  'Video': 'vinijr',
  'Email': 'Email',
  'Image': 'Image',
  'Rich Mix': 'Rmix',
  'Shopping Campaigns with Partners': 'SCwP',
  'Shopping PLA': 'ShopPLA',
  'Shopping Showcase Ads': 'ShopShow',
  'Audio': '365scores',
  'Display': 'Disp',
  'NativeDisplay': 'NatDisp',
  'NativeVideo': 'NatVid',
  'carousel': 'carousel',
  'collection': 'collect',
  'dynamic-product-ad-{NA-Only}': 'DynProdNA',
  'event-response': 'EventResp',
  'image': 'Image',
  'instant-experiences': 'instant-experiences',
  'lead-generation': 'lead-generation',
  'offer-ads': 'offer-ads',
  'slideshow': 'Sshow',
  'stories': 'stor',
  'video': 'vinijr',
  'Collection Ad': 'Collection',
  'Display Cards': 'DispCard',
  'Dynamic Showcase Ad': 'DSA',
  'In-Feed Ads': 'InFeed',
  'Instant Page': 'InstPage',
  'Interactive Cards': 'InteraCard',
  'Live Shopping Ads': 'LSA',
  'Product Shopping Ads': 'PSA',
  'Spark Ads': 'SparkAds',
  'SuperLike': 'SuperLike',
  'Video Shopping Ads': 'VSA',
  'Voting Cards': 'VotCard',
  'Top-feed Ads': 'TopFeed',
  'Idea': 'Idea',
  'Max. Video': 'MaxVid',
  'Shopping Ads': 'Shop',
  'Standard Video': 'StandVid',
  'App Card': 'AppCard',
  'Direct Message Card': 'DMCard',
  'Promoted Account': 'PromAc',
  'Promoted Moments': 'ProMom',
  'Website Card': 'WebCard',
  'GIF': 'GIF',
  'Cpas-Rmkt': 'Cpas-Rmkt',
  'E-mail-Marketing': 'E-mail-Marketing',
  'InApp': 'InApp',
  'Produto-Patrocinado': 'Produto-Patrocinado',
  'Rec': 'Rec',
  'Rec-1P': 'Rec-1P',
  'Sponsored-Brand-Product-Collection': 'SponBPC',
  'Sponsored-Brand-Store-Spotlight': 'SponBSS',
  'Sponsored-Brand-Video': 'SponBV',
  'Sponsored-Brands': 'Sponsored-Brands',
  'Sponsored-Display': 'Sponsored-Display',
  'Sponsored-Products': 'Sponsored-Products',
  'TP-Checkout': 'TP-Checkout',
  'TP-Search': 'TP-Search',
  'N-A': 'N-A',
  'Playable': 'Playable',
  'InstantExperience': 'InstantExperience',
  'Carrossel': 'Carrossel',
  'html': 'html',
  'ResponsiveEcom': 'ResponsiveEcom',
  'Widget-Count': 'Widget-Count',
  'Widget-Countdown': 'Widget-Countdown',
  'Insterscroller': 'Insterscroller',
  'Immersive-Banner': 'Immersive-Banner',
  'Brand Generic': 'BrndGen',
  'Brand Product': 'BrndProd',
  'Competitor': 'Comp',
  'Dynamic Search Ads': 'DSA',
  'Generic': 'Gen',
  'Generic Product': 'GenProd',
  'Broad Match': 'BM',
  'Dynamic Search Ads + Exact': 'DSAE',
  'Dynamic Search Ads + Broad': 'DSAB',
  'Dynamic Search Ads + Broad + Exact': 'DSABE',
  'Dynamic Search Ads + Broad + Exact + Phrase': 'DSABPE',
  'Dynamic Search Ads + Broad + Phrase': 'DSABP',
  'Dynamic Search Ads + Phrase': 'DSAP',
  'Dynamic Search Ads + Phrase + Exact': 'DSAPE',
  'Exact Match': 'EM',
  'Exact Match + Broad': 'EMB',
  'Exact Match + Broad + Phrase': 'EMBP',
  'Exact Match + Phrase': 'EMP',
  'Phrase Match': 'PM',
  'Phrase Match + Broad': 'PMB',
  'vpaid': 'vpaid',
  'vast': 'vast',
  'pixel': 'pixel',
  'native': 'native',
  'audio': '365scores',
  'multi-tracker': 'multi-tracker',
  'video-pixel': 'video-pixel',
  'audio-pixel': 'audio-pixel',
  'Audience': 'Aud',
  'Contextual': 'CNT',
  'Domains-List': 'Domains-List',
  'Run-of-Network': 'Run-of-Network',
  'Ad Recall': 'AdRcl',
  'App Event': 'AppEve',
  'Brand Consideration': 'BrndCons',
  'Call': 'Call',
  'Catalog Sales': 'CatSal',
  'Clicks': 'Clicks',
  'Conversations': 'Conversations',
  'Conversion Manual': 'ConvMnl',
  'Conversion Manual Revenue': 'ConvMnlRev',
  'Conversion Search': 'ConvSrch',
  'Conversion Smart': 'ConvSmt',
  'Conversion Tik Tok Shop Clicks': 'ConvTTSClk',
  'Conversion Tik Tok Shop Chekouts': 'ConvTTSChk',
  'Conversion Tik Tok Shop Purchase': 'ConvTTSPur',
  'Conversion Tik Tok Shop Revenue': 'ConvTTSRev',
  'Conversions': 'Conversions',
  'Daily Reach': 'DlyRch',
  'Event': 'Event',
  'Impressions': 'Impressions',
  'Install': 'Install',
  'Interaction': 'Interaction',
  'Interaction with the community Follow': 'IntrFollow',
  'Interaction with the community Page Visit': 'IntrPgVst',
  'Leads': 'Leads',
  'Link Clicks': 'LnkClk',
  'Manual Clicks': 'MnlClk',
  'Manual Leads': 'MnlLead',
  'Manual Trafic': 'MnlTrf',
  'Messages': 'Message',
  'Page View': 'PgView',
  'Profile View': 'ProfView',
  'Reach': 'Reach',
  'Reach & Frequence': 'R&F',
  'Reach & Frequence Non Skip': 'R&FNSkp',
  'Reach & Frequence Skip': 'R&FSkp',
  'Reach Non Skip': 'RNSkp',
  'Reach Skip': 'RSkp',
  'Reminder': 'Reminder',
  'Revenue': 'Revenue',
  'Search Clicks': 'SrchClk',
  'Search Trafic': 'SrchTrf',
  'Smart Clicks': 'SmtClk',
  'Smart Leads': 'SmtLead',
  'Smart Trafic': 'SmtTrf',
  'Views 6 Seg': 'V6s',
  'Views 15 Seg': 'V15s',
  'Views 10 Seg': 'V10s',
  'Views 2 Seg': 'V2s',
  'Views Conclusion': 'VConc',
  'Views Non Skip': 'VNSkp',
  'Views Skip': 'VSkp',
  'Brand Say': 'BSAY',
  'Influencer Handle (Influencer) / Other say': 'InfHndOSay',
  'Brand Handle (Influencer) / Brand say': 'BrdHndBSay',
  'dark': 'dark',
  'boosted': 'BSTD',
  'brd': 'brd',
  'inf': 'inf',
  'ugc': 'ugc',
  'regiao': 'regiao',
  'estado': 'estado',
  'cidades': 'cidades',
  'nacional': 'nacional',
  'cid-aju': 'cid-aju',
  'cid-apg': 'cid-apg',
  'cid-aps': 'cid-aps',
  'cid-bel': 'cid-bel',
  'cid-bel-mao': 'cid-bel-mao',
  'cid-bel-slz': 'cid-bel-slz',
  'cid-bhz': 'cid-bhz',
  'cid-bhz-jdf': 'cid-bhz-jdf',
  'cid-bhz-udi': 'cid-bhz-udi',
  'cid-bnu': 'cid-bnu',
  'cid-bnu-fln': 'cid-bnu-fln',
  'cid-brb': 'cid-brb',
  'cid-brb-cgb': 'cid-brb-cgb',
  'cid-brb-gyn': 'cid-brb-gyn',
  'cid-cau': 'cid-cau',
  'cid-cax': 'cid-cax',
  'cid-cgb': 'cid-cgb',
  'cid-cgr': 'cid-cgr',
  'cid-con': 'cid-con',
  'cid-cpq': 'cid-cpq',
  'cid-cpq-rao': 'cid-cpq-rao',
  'cid-cpq-sod': 'cid-cpq-sod',
  'cid-cwb': 'cid-cwb',
  'cid-cwb-fln': 'cid-cwb-fln',
  'cid-cwb-poa': 'cid-cwb-poa',
  'cid-cxj': 'cid-cxj',
  'cid-fec': 'cid-fec',
  'cid-fln': 'cid-fln',
  'cid-for': 'cid-for',
  'cid-gru': 'cid-gru',
  'cid-gyn': 'cid-gyn',
  'cid-imp': 'cid-imp',
  'cid-jdf': 'cid-jdf',
  'cid-jdo': 'cid-jdo',
  'cid-joi': 'cid-joi',
  'cid-joi-bnu': 'cid-joi-bnu',
  'cid-joi-fln': 'cid-joi-fln',
  'cid-jpa': 'cid-jpa',
  'cid-jpa-nat': 'cid-jpa-nat',
  'cid-jpa-rec': 'cid-jpa-rec',
  'cid-ldb': 'cid-ldb',
  'cid-mao': 'cid-mao',
  'cid-mao-cgb': 'cid-mao-cgb',
  'cid-mcz': 'cid-mcz',
  'cid-mgf': 'cid-mgf',
  'cid-nat': 'cid-nat',
  'cid-nat-for': 'cid-nat-for',
  'cid-nig': 'cid-nig',
  'cid-nit': 'cid-nit',
  'cid-osa': 'cid-osa',
  'cid-pet': 'cid-pet',
  'cid-pnz': 'cid-pnz',
  'cid-poa': 'cid-poa',
  'cid-poa-fln': 'cid-poa-fln',
  'cid-rao': 'cid-rao',
  'cid-rao-sod': 'cid-rao-sod',
  'cid-rec': 'cid-rec',
  'cid-rec-for': 'cid-rec-for',
  'cid-rio': 'cid-rio',
  'cid-rio-bhz': 'cid-rio-bhz',
  'cid-rio-brb': 'cid-rio-brb',
  'cid-rio-rec': 'cid-rio-rec',
  'cid-rio-ssa': 'cid-rio-ssa',
  'cid-sao': 'cid-sao',
  'cid-sao-bhz': 'cid-sao-bhz',
  'cid-sao-brb': 'cid-sao-brb',
  'cid-sao-cpq': 'cid-sao-cpq',
  'cid-sao-cwb': 'cid-sao-cwb',
  'cid-sao-for': 'cid-sao-for',
  'cid-sao-gru': 'cid-sao-gru',
  'cid-sao-osa': 'cid-sao-osa',
  'cid-sao-poa': 'cid-sao-poa',
  'cid-sao-rec': 'cid-sao-rec',
  'cid-sao-rio': 'cid-sao-rio',
  'cid-sao-sbc': 'cid-sao-sbc',
  'cid-sao-ssa': 'cid-sao-ssa',
  'cid-sbc': 'cid-sbc',
  'cid-slz': 'cid-slz',
  'cid-sod': 'cid-sod',
  'cid-ssa': 'cid-ssa',
  'cid-ssa-for': 'cid-ssa-for',
  'cid-ssa-rec': 'cid-ssa-rec',
  'cid-ssz': 'cid-ssz',
  'cid-stm': 'cid-stm',
  'cid-the': 'cid-the',
  'cid-udi': 'cid-udi',
  'cid-vix': 'cid-vix',
  'cid-vix-bhz': 'cid-vix-bhz',
  'cid-vix-rio': 'cid-vix-rio',
  'est-ac': 'est-ac',
  'est-al': 'est-al',
  'est-am': 'est-am',
  'est-ap': 'est-ap',
  'est-ba': 'est-ba',
  'est-ce': 'est-ce',
  'est-df': 'est-df',
  'est-es': 'est-es',
  'est-go': 'est-go',
  'est-ma': 'est-ma',
  'est-mg': 'est-mg',
  'est-ms': 'est-ms',
  'est-mt': 'est-mt',
  'est-pa': 'est-pa',
  'est-pb': 'est-pb',
  'est-pe': 'est-pe',
  'est-pi': 'est-pi',
  'est-pr': 'est-pr',
  'est-rj': 'est-rj',
  'est-rn': 'est-rn',
  'est-ro': 'est-ro',
  'est-rr': 'est-rr',
  'est-rs': 'est-rs',
  'est-sc': 'est-sc',
  'est-se': 'est-se',
  'est-sp': 'est-sp',
  'est-to': 'est-to',
  'nac': 'nac',
  'reg-co': 'reg-co',
  'reg-co-se': 'reg-co-se',
  'reg-ne': 'reg-ne',
  'reg-ne-se': 'reg-ne-se',
  'reg-ne-se-su': 'reg-ne-se-su',
  'reg-no': 'reg-no',
  'reg-no-ne': 'reg-no-ne',
  'reg-no-se': 'reg-no-se',
  'reg-no-se-su': 'reg-no-se-su',
  'reg-se': 'reg-se',
  'reg-su': 'reg-su',
  'reg-su-se': 'reg-su-se'
};
  blocos.forEach(b => b.items.forEach(it => {
    if(!it.abr) {
      if(SIGLAS[it.opt]) it.abr = SIGLAS[it.opt];
    }
  }));

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
