/* =====================================================================
   Gerador de Documentos — Infobarra
   - Contrato de Comodato e Recibo de Equipamento
   - Importa base dos .mdb (JSON), permite adicionar/editar (localStorage)
   - Gera DOCX (substituição de tokens no template) e PDF (réplica HTML)
   ===================================================================== */
'use strict';

/* ----------------------------- Utilidades ----------------------------- */
const $  = (sel, ctx = document) => ctx.querySelector(sel);
const $$ = (sel, ctx = document) => Array.from(ctx.querySelectorAll(sel));

const MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho',
               'Agosto','Setembro','Outubro','Novembro','Dezembro'];

function hojeLongoPt(){
  const d = new Date();
  return `${d.getDate()} de ${MESES[d.getMonth()]} de ${d.getFullYear()}`;
}

/** Escapa texto para inserção segura em XML (DOCX) */
function escXml(s){
  return String(s ?? '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&apos;')
    .replace(/\r\n|\r|\n/g,' ');
}
/** Escapa texto para HTML (PDF) */
function escHtml(s){
  return String(s ?? '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
/** Chave normalizada de coluna (remove espaços, maiúsculas) p/ acesso robusto */
function normKey(s){ return String(s).toUpperCase().replace(/\s+/g,'').trim(); }

let toastTimer = null;
function toast(msg, isErr = false){
  const t = $('#toast');
  t.textContent = msg;
  t.classList.toggle('err', !!isErr);
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 2600);
}

/* ----------------------- Definição dos documentos ---------------------- */
/* Cada campo: { token (no template), col (chave no JSON), label, type } */
const DOCS = {
  comodato: {
    title: 'Contrato de Comodato',
    sub: 'Comodato de equipamentos — pessoa física ou jurídica.',
    templateUrl: 'templates/template_contrato.docx',
    dataUrl: 'data/clientes_comodato.json',
    storeKey: 'infobarra_comodato_v1',
    fileBase: 'Contrato_Comodato',
    sections: [
      { title: 'Identificação do Comodatário', grid: 'g3', fields: [
        { token:'CONTRATO',      col:'CONTRATO',      label:'Nº do Contrato' },
        { token:'CLIENTE_',      col:'CLIENTE ',      label:'Cliente (Pessoa Física)', span:2 },
        { token:'RAZÃO_SOCIAL',  col:'RAZÃO SOCIAL',  label:'Razão Social (Pessoa Jurídica)', span:2 },
        { token:'CPFCNPJ',       col:'CPF/CNPJ',      label:'CPF / CNPJ' },
        { token:'RG_IE',         col:'RG /IE',        label:'RG / I.E.' },
        { token:'ORG_EXPEDIDOR', col:'ORG EXPEDIDOR', label:'Órgão Expedidor' },
        { token:'ESTADO_CIVIL',  col:'ESTADO CIVIL',  label:'Estado Civil' },
        { token:'PROFISSÃO',     col:'PROFISSÃO',     label:'Profissão' },
      ]},
      { title: 'Endereço', grid: 'g3', fields: [
        { token:'RUA',    col:'RUA',    label:'Rua / Logradouro', span:2 },
        { token:'NUMERO', col:'NUMERO', label:'Número' },
        { token:'BAIRRO', col:'BAIRRO', label:'Bairro' },
        { token:'CIDADE', col:'CIDADE', label:'Cidade / UF', span:2 },
      ]},
      { title: 'Contato', grid: 'g2', fields: [
        { token:'TELEFONE', col:'TELEFONE', label:'Telefone' },
        { token:'EMAIL',    col:'EMAIL',    label:'E-mail' },
      ]},
    ],
    equip: [
      { chk:{token:'ONU_E_FONTE', col:'ONU E FONTE'}, label:'ONU com Fonte',            q:{token:'QUANT_01',col:'QUANT 01'}, m:{token:'MARCA_01',col:'MARCA 01'}, mo:{token:'MODELO_01',col:'MODELO 01'} },
      { chk:{token:'ONU_WIFI',    col:'ONU WIFI'},    label:'ONU Wireless com Fonte',   q:{token:'QUANT_02',col:'QUANT 02'}, m:{token:'MARCA_02',col:'MARCA 02'}, mo:{token:'MODELO_02',col:'MODELO 02'} },
      { chk:{token:'ROTEADOR',    col:'ROTEADOR'},    label:'Roteador com Fonte',       q:{token:'QUANT_03',col:'QUANT 03'}, m:{token:'MARCA_03',col:'MARCA 03'}, mo:{token:'MODELO_03',col:'MODELO 03'} },
      { chk:{token:'AIRGRID',     col:'AIRGRID'},     label:'Airgrid completa c/ fonte',q:{token:'QUANT_04',col:'QUANT 04'}, m:{token:'MARCA_04',col:'MARCA 04'}, mo:{token:'MODELO_04',col:'MODELO 04'} },
      { chk:{token:'OUTROS',      col:'OUTROS'},      label:'Outros', q:{token:'QUANT_05',col:'QUANT 05'}, m:{token:'MARCA_05',col:'MARCA 05'}, mo:{token:'MODELO_05',col:'MODELO 05'} },
    ],
    date: { token:'DATA', col:'DATA' },
  },

  recibo: {
    title: 'Recibo de Equipamento',
    sub: 'Recibo de entrega/devolução de equipamento.',
    templateUrl: 'templates/template_recibo.docx',
    dataUrl: 'data/clientes_recibo.json',
    storeKey: 'infobarra_recibo_v1',
    fileBase: 'Recibo_Equipamento',
    sections: [
      { title: 'Identificação do Comodatário', grid: 'g3', fields: [
        { token:'CONTRATO', col:'CONTRATO', label:'Nº do Contrato' },
        { token:'CLIENTE_', col:'CLIENTE ', label:'Nome / Razão Social', span:2 },
        { token:'CPF',      col:'CPF',      label:'CPF / CNPJ' },
        { token:'CONTATO',  col:'CONTATO',  label:'Contato', span:2 },
      ]},
      { title: 'Endereço', grid: 'g3', fields: [
        { token:'RUA',    col:'RUA',    label:'Rua / Logradouro', span:2 },
        { token:'NUMERO', col:'NUMERO', label:'Número' },
        { token:'BAIRRO', col:'BAIRRO', label:'Bairro' },
        { token:'CIDADE', col:'CIDADE', label:'Cidade / UF', span:2 },
      ]},
    ],
    equip: [
      { chk:{token:'ONU',      col:'ONU'},      label:'ONU',           q:{token:'QT_01',col:'QT 01'}, m:{token:'MARCA_01',col:'MARCA 01'}, mo:{token:'MODELO_01',col:'MODELO 01'} },
      { chk:{token:'ONU_WIFI', col:'ONU WIFI'}, label:'ONU Wireless',  q:{token:'QT_02',col:'QT 02'}, m:{token:'MARCA_02',col:'MARCA 02'}, mo:{token:'MODELO_02',col:'MODELO 02'} },
      { chk:{token:'ROTEADOR', col:'ROTEADOR'}, label:'Roteador',      q:{token:'QT_03',col:'QT 03'}, m:{token:'MARCA_03',col:'MARCA 03'}, mo:{token:'MODELO_03',col:'MODELO 03'} },
      { chk:{token:'FONTE_',   col:'FONTE '},   label:'Fonte',         q:{token:'QT_04',col:'QT 04'}, m:{token:'MARCA_04',col:'MARCA 04'}, mo:{token:'MODELO_04',col:'MODELO 04'} },
      { chk:{token:'OUTROS',   col:'OUTROS'},   label:'Outros', q:{token:'QT_05',col:'QT 05'}, m:{token:'MARCA_05',col:'MARCA 05'}, mo:{token:'MODELO_05',col:'MODELO 05'} },
    ],
    date: { token:'DATA', col:'DATA' },
  },
};

/* ----------------------------- Estado ----------------------------- */
const state = {
  doc: 'comodato',     // documento ativo
  clients: { comodato: [], recibo: [] },   // registros (seed + locais)
  activeIdx: -1,       // índice do cliente selecionado na lista filtrada
  filtered: [],        // lista filtrada atual
  baseMarcas: {},      // catálogo base (catalogo.json): { marcaDisplay: [modelos] }
  catalogByKey: {},    // { normMarca: [modelos] }  (base + aprendidos)
  allMarcas: [],       // marcas (forma de exibição) ordenadas
  allModelos: [],      // união de todos os modelos
  view: 'doc',         // 'doc' | 'estoque'
  lastDeductSig: '',   // evita baixa dupla de estoque (DOCX + PDF do mesmo doc)
  operador: null,      // operador logado { id, nome, role }
  estoque: [],         // cache do estoque (memória); persiste local OU Supabase
  catalogUser: {},     // catálogo aprendido (memória); persiste local OU Supabase
};

/* Normaliza marca para casar variações (TP-LINK / TP LINK / TPLINK) */
function normMarca(s){ return String(s||'').toUpperCase().replace(/[^A-Z0-9]/g,''); }

/* Carrega o catálogo Marca->Modelos (data/catalogo.json) */
const USER_CAT_KEY = 'infobarra_catalogo_user_v1';
function loadUserCatalogLocal(){ try { return JSON.parse(localStorage.getItem(USER_CAT_KEY) || '{}'); } catch(e){ return {}; } }
function loadUserCatalog(){ return state.catalogUser || {}; }   // cache em memória (base p/ rebuild)
function saveUserCatalog(o){
  state.catalogUser = o || {};
  if (_supa()) syncCatalogRemote(state.catalogUser);
  else localStorage.setItem(USER_CAT_KEY, JSON.stringify(state.catalogUser));
}

/* Catálogo aprendido na nuvem (tabela catalogo, PK = marca) */
async function loadUserCatalogRemote(){
  try {
    const { data, error } = await sb.from('catalogo').select('marca,modelos');
    if (error) throw error;
    const o = {};
    for (const r of (data || [])){ o[normMarca(r.marca)] = { display: r.marca, modelos: r.modelos || [] }; }
    return o;
  } catch(e){ console.warn('loadUserCatalogRemote:', e.message); return {}; }
}
async function syncCatalogRemote(o){
  try {
    const rows = Object.values(o || {})
      .map(u => ({ marca: (u.display || '').trim(), modelos: u.modelos || [] }))
      .filter(r => r.marca);
    if (rows.length){ const { error } = await sb.from('catalogo').upsert(rows, { onConflict: 'marca' }); if (error) throw error; }
  } catch(e){ console.warn('syncCatalogRemote:', e.message); }
}

async function loadCatalog(){
  try {
    const r = await fetch('data/catalogo.json', { cache:'no-store' });
    const j = await r.json();
    state.baseMarcas = j.marcas || {};
  } catch(e){
    console.warn('Catálogo base não carregado:', e);
    state.baseMarcas = {};
  }
  state.catalogUser = _supa() ? await loadUserCatalogRemote() : loadUserCatalogLocal();
  rebuildCatalog();
}

/* Recompõe o catálogo em memória: base (catalogo.json) + aprendidos (localStorage) */
function rebuildCatalog(){
  const byKey = {}, display = {}, allMod = new Set();
  const addMarca = d => { const k = normMarca(d); if (!byKey[k]){ byKey[k] = []; display[k] = d; } return k; };
  const addModel = (k, m) => { m = (m||'').trim(); if (m && !byKey[k].some(x => x.toUpperCase() === m.toUpperCase())){ byKey[k].push(m); allMod.add(m); } };

  for (const d in state.baseMarcas){ const k = addMarca(d); (state.baseMarcas[d]||[]).forEach(m => addModel(k, m)); }
  const user = loadUserCatalog();
  for (const uk in user){ const u = user[uk]; const k = addMarca(u.display || uk); (u.modelos||[]).forEach(m => addModel(k, m)); }

  for (const k in byKey) byKey[k].sort((a,b) => a.localeCompare(b,'pt-BR'));
  state.catalogByKey = byKey;
  state.allMarcas = Object.keys(byKey).map(k => display[k]).sort((a,b) => a.localeCompare(b,'pt-BR'));
  state.allModelos = Array.from(allMod).sort((a,b) => a.localeCompare(b,'pt-BR'));
}

/* Aprende uma marca/modelo nova (persistido). Retorna true se algo mudou. */
function learnCatalog(marca, modelo){
  marca = (marca||'').trim(); modelo = (modelo||'').trim();
  if (!marca) return false;
  const key = normMarca(marca);
  const user = loadUserCatalog();
  let changed = false;
  if (!user[key]){ user[key] = { display: marca, modelos: [] }; changed = true; }
  if (modelo){
    const inBase = (state.catalogByKey[key] || []).some(m => m.toUpperCase() === modelo.toUpperCase());
    const inUser = (user[key].modelos || []).some(m => m.toUpperCase() === modelo.toUpperCase());
    if (!inBase && !inUser){ user[key].modelos.push(modelo); changed = true; }
  }
  if (changed){ saveUserCatalog(user); rebuildCatalog(); }
  return changed;
}

/* Aprende todas as marcas/modelos preenchidas num documento */
function learnFromValues(vals, def){
  let changed = false;
  for (const row of def.equip){ if (learnCatalog(vals[row.m.token], vals[row.mo.token])) changed = true; }
  return changed;
}

/* Lê valor de um registro pela coluna, com fallback por chave normalizada */
function recVal(rec, col){
  if (rec == null) return '';
  if (col in rec) return rec[col] ?? '';
  const nk = normKey(col);
  for (const k in rec){ if (normKey(k) === nk) return rec[k] ?? ''; }
  return '';
}

/* Nome de exibição de um registro */
function displayName(rec){
  return (recVal(rec,'CLIENTE ') || recVal(rec,'RAZÃO SOCIAL') || '(sem nome)').trim();
}

/* Registro tem conteúdo real? (descarta linhas em branco vindas do Access) */
function clientHasContent(r){
  const nome = (recVal(r,'CLIENTE ') || recVal(r,'RAZÃO SOCIAL')).trim();
  return !!(nome || recVal(r,'CONTRATO').trim() || recVal(r,'CPF/CNPJ').trim() || recVal(r,'CPF').trim());
}

/* --------------------------- Carga de dados --------------------------- */
async function loadData(docKey){
  const def = DOCS[docKey];

  // Modo nuvem: lê os clientes da tabela `clientes` (RLS exige login → [] sem sessão)
  if (_supa()){
    try {
      const { data, error } = await sb.from('clientes').select('dados').eq('tipo', docKey);
      if (error) throw error;
      state.clients[docKey] = (data || [])
        .map(row => row.dados || {})
        .filter(clientHasContent)
        .sort((a,b) => displayName(a).localeCompare(displayName(b),'pt-BR'));
    } catch(e){
      console.warn('loadData (supabase):', e.message);
      state.clients[docKey] = state.clients[docKey] || [];
    }
    return;
  }

  // Modo local: seed (JSON) + edições do localStorage, dedup por CONTRATO
  let seed = [];
  try {
    const r = await fetch(def.dataUrl, { cache:'no-store' });
    if (r.ok) seed = await r.json();
  } catch(e){ console.warn('Falha ao carregar seed', def.dataUrl, e); }

  let local = [];
  try { local = JSON.parse(localStorage.getItem(def.storeKey) || '[]'); } catch(e){}

  const byContrato = new Map();
  for (const r of seed)  byContrato.set(normKey(recVal(r,'CONTRATO')) || Symbol(), r);
  for (const r of local) byContrato.set(normKey(recVal(r,'CONTRATO')) || Symbol(), r);

  state.clients[docKey] = Array.from(byContrato.values())
    .filter(clientHasContent)
    .sort((a,b) => displayName(a).localeCompare(displayName(b),'pt-BR'));
}

/* Persiste registros locais (modo local). Na nuvem, persistência é no Supabase. */
function saveLocal(docKey){
  if (_supa()) return;
  const def = DOCS[docKey];
  const local = state.clients[docKey].filter(r => r._local);
  localStorage.setItem(def.storeKey, JSON.stringify(local));
}

/* --------------------------- Render: lista --------------------------- */
function renderList(){
  const def = DOCS[state.doc];
  const q = normKey($('#search').value);
  const all = state.clients[state.doc];

  state.filtered = !q ? all : all.filter(r => {
    const hay = normKey([
      displayName(r), recVal(r,'CPF/CNPJ'), recVal(r,'CPF'), recVal(r,'CONTRATO')
    ].join(' '));
    return hay.includes(q);
  });

  $('#list-count').textContent = state.filtered.length;
  const box = $('#client-list');

  if (!state.filtered.length){
    box.innerHTML = `<div class="list-empty">Nenhum cliente encontrado.</div>`;
    return;
  }

  box.innerHTML = state.filtered.map((r, i) => {
    const doc = recVal(r,'CPF/CNPJ') || recVal(r,'CPF') || '—';
    const contrato = recVal(r,'CONTRATO') || '—';
    const active = i === state.activeIdx ? ' active' : '';
    return `<div class="client-item${active}" data-i="${i}">
      <div class="ci-name">${escHtml(displayName(r))}</div>
      <div class="ci-meta"><b>${escHtml(contrato)}</b><span>${escHtml(doc)}</span></div>
    </div>`;
  }).join('');

  $$('.client-item', box).forEach(el => {
    el.addEventListener('click', () => selectClient(parseInt(el.dataset.i, 10)));
  });
}

/* --------------------------- Render: formulário --------------------------- */
function renderForm(){
  const def = DOCS[state.doc];
  const form = $('#doc-form');
  $('#form-title').textContent = def.title;
  $('#form-sub').textContent = def.sub;

  let html = '';

  // Seções de campos
  for (const sec of def.sections){
    html += `<div class="card"><div class="card-head"><span class="dot"></span><h2>${escHtml(sec.title)}</h2></div>
      <div class="grid ${sec.grid}">`;
    for (const f of sec.fields){
      const span = f.span ? ` style="grid-column:span ${f.span}"` : '';
      html += `<div class="field"${span}>
        <label>${escHtml(f.label)}</label>
        <input type="text" data-token="${f.token}" />
      </div>`;
    }
    html += `</div></div>`;
  }

  // Equipamentos
  const marcaOpts = state.allMarcas.map(m => `<option value="${escHtml(m)}"></option>`).join('');
  html += `<div class="card"><div class="card-head"><span class="dot"></span><h2>Objeto do Comodato — Equipamentos</h2></div>
    <div class="equip-head"><span></span><span>Equipamento</span><span>Qtd</span><span>Marca</span><span>Modelo</span></div>
    <datalist id="dl-marcas">${marcaOpts}</datalist>`;
  def.equip.forEach((row, i) => {
    const labelCell = row.desc
      ? `<div class="erow-label">${escHtml(row.label)}:
           <input type="text" data-token="${row.desc.token}" class="erow-desc" placeholder="descrição do equipamento" /></div>`
      : `<div class="erow-label">${escHtml(row.label)}</div>`;
    html += `<div class="equip-row">
      <div class="chk" data-chk="${row.chk.token}" role="checkbox" tabindex="0" aria-label="${escHtml(row.label)}"></div>
      ${labelCell}
      <input type="text" data-token="${row.q.token}" placeholder="0" inputmode="numeric" />
      <input type="text" data-token="${row.m.token}" list="dl-marcas" data-marca-row="${i}" placeholder="Marca ▾" autocomplete="off" />
      <input type="text" data-token="${row.mo.token}" list="dl-modelos-${i}" placeholder="Modelo ▾" autocomplete="off" />
      <datalist id="dl-modelos-${i}"></datalist>
    </div>`;
  });
  html += `</div>`;

  // Data
  html += `<div class="card"><div class="card-head"><span class="dot"></span><h2>Data do Documento</h2></div>
    <div class="grid g2">
      <div class="field"><label>Data por extenso (sai no documento)</label>
        <input type="text" data-token="${def.date.token}" id="data-extenso" placeholder="ex.: 5 de Junho de 2026" /></div>
      <div class="field"><label>&nbsp;</label>
        <button type="button" class="btn btn-soft" id="btn-hoje">📅 Usar data de hoje</button></div>
    </div></div>`;

  form.innerHTML = html;

  // Checkboxes interativos
  $$('.chk', form).forEach(c => {
    const toggle = () => c.classList.toggle('on');
    c.addEventListener('click', toggle);
    c.addEventListener('keydown', e => { if (e.key === ' ' || e.key === 'Enter'){ e.preventDefault(); toggle(); }});
  });
  $('#btn-hoje').addEventListener('click', () => { $('#data-extenso').value = hojeLongoPt(); });

  // Comboboxes: ao mudar a Marca, filtra os Modelos daquela marca
  $$('input[data-marca-row]', form).forEach(inp => {
    const i = parseInt(inp.dataset.marcaRow, 10);
    inp.addEventListener('input', () => refreshModeloList(i));
    refreshModeloList(i);
  });
}

/* Repopula o datalist de Modelos da linha i conforme a Marca escolhida */
function refreshModeloList(i){
  const form = $('#doc-form');
  const marcaInp = form.querySelector(`[data-marca-row="${i}"]`);
  const dl = form.querySelector(`#dl-modelos-${i}`);
  if (!marcaInp || !dl) return;
  const key = normMarca(marcaInp.value);
  const models = (key && state.catalogByKey[key]) ? state.catalogByKey[key] : [];   // só expande com a marca escolhida
  dl.innerHTML = (models || []).map(x => `<option value="${escHtml(x)}"></option>`).join('');
}

/* Preenche o formulário com um registro (ou limpa) */
function fillForm(rec){
  const form = $('#doc-form');
  const def = DOCS[state.doc];

  $$('input[data-token]', form).forEach(inp => { inp.value = ''; });
  $$('.chk', form).forEach(c => c.classList.remove('on'));

  const refreshAll = () => DOCS[state.doc].equip.forEach((_, i) => refreshModeloList(i));

  if (!rec){
    $('#data-extenso').value = hojeLongoPt();
    refreshAll();
    return;
  }

  // Campos de texto: mapeia token -> coluna
  const tokenToCol = buildTokenColMap(def);
  $$('input[data-token]', form).forEach(inp => {
    const tk = inp.dataset.token;
    const col = tokenToCol[tk];
    inp.value = col ? recVal(rec, col) : '';
  });

  // Checkboxes
  $$('.chk', form).forEach(c => {
    const tk = c.dataset.chk;
    const col = tokenToCol[tk];
    const v = normKey(recVal(rec, col));
    if (v === 'X') c.classList.add('on');
  });

  // Data
  const dv = recVal(rec, def.date.col);
  $('#data-extenso').value = dv || hojeLongoPt();

  refreshAll();   // sincroniza os modelos com as marcas carregadas
}

/* Mapa token -> coluna (a partir da definição do doc) */
function buildTokenColMap(def){
  const map = {};
  for (const sec of def.sections) for (const f of sec.fields) map[f.token] = f.col;
  for (const row of def.equip){ map[row.chk.token] = row.chk.col; map[row.q.token] = row.q.col; map[row.m.token] = row.m.col; map[row.mo.token] = row.mo.col; if (row.desc) map[row.desc.token] = row.desc.col; }
  map[def.date.token] = def.date.col;
  return map;
}

/* Coleta valores do formulário -> { TOKEN: valor } */
function collectValues(){
  const form = $('#doc-form');
  const vals = {};
  $$('input[data-token]', form).forEach(inp => { vals[inp.dataset.token] = inp.value.trim(); });
  $$('.chk', form).forEach(c => { vals[c.dataset.chk] = c.classList.contains('on') ? 'X' : ''; });
  return vals;
}

/* Coleta valores -> registro (chaves = colunas do JSON) p/ salvar */
function collectRecord(){
  const def = DOCS[state.doc];
  const vals = collectValues();
  const tokenToCol = buildTokenColMap(def);
  const rec = { _local: true };
  for (const tk in vals){ rec[tokenToCol[tk] || tk] = vals[tk]; }
  return rec;
}

/* --------------------------- Seleção --------------------------- */
function selectClient(i){
  state.activeIdx = i;
  fillForm(state.filtered[i]);
  renderList();
  $('.form-scroll').scrollTop = 0;
}

function newClient(){
  state.activeIdx = -1;
  fillForm(null);
  renderList();
  toast('Novo cliente — preencha e clique em Salvar ou gere o documento.');
}

async function saveClient(){
  const def = DOCS[state.doc];
  const rec = collectRecord();
  const contrato = normKey(recVal(rec,'CONTRATO'));
  if (!displayName(rec) || displayName(rec) === '(sem nome)'){
    toast('Informe ao menos o nome/razão social.', true); return;
  }

  // Substitui registro existente com mesmo contrato, senão adiciona
  const arr = state.clients[def === DOCS.comodato ? 'comodato' : 'recibo'] || state.clients[state.doc];
  let idx = -1;
  if (contrato) idx = arr.findIndex(r => normKey(recVal(r,'CONTRATO')) === contrato);
  if (idx >= 0){ rec._local = true; arr[idx] = rec; }
  else arr.push(rec);

  arr.sort((a,b) => displayName(a).localeCompare(displayName(b),'pt-BR'));
  saveLocal(state.doc);
  if (_supa()) await upsertClienteRemote(state.doc, rec);
  learnFromValues(collectValues(), def);   // aprende marcas/modelos digitados
  renderList();
  toast(_supa() ? 'Cliente salvo na nuvem. ✓' : 'Cliente salvo na base local. ✓');
}

/* Upsert de 1 cliente na nuvem (casa por tipo+contrato; senão insere) */
async function upsertClienteRemote(docKey, rec){
  try {
    const contrato = (recVal(rec,'CONTRATO') || '').trim();
    const dados = Object.assign({}, rec); delete dados._local;
    const row = {
      tipo: docKey,
      contrato: contrato || null,
      nome: displayName(rec),
      cpf: (recVal(rec,'CPF/CNPJ') || recVal(rec,'CPF') || '').trim(),
      dados,
      updated_at: new Date().toISOString(),
    };
    if (contrato){
      const { data: ex } = await sb.from('clientes').select('id').eq('tipo', docKey).eq('contrato', contrato).limit(1).maybeSingle();
      if (ex){ const { error } = await sb.from('clientes').update(row).eq('id', ex.id); if (error) throw error; return; }
    }
    const { error } = await sb.from('clientes').insert(row);
    if (error) throw error;
  } catch(e){ toast('Falha ao salvar na nuvem: ' + e.message, true); }
}

/* Importa a base local (JSON + edições) para a nuvem — admin, roda 1 vez */
async function importLocalClientsToCloud(){
  if (!_supa()){ toast('Importação só no modo nuvem.', true); return; }
  if (!confirm('Importar a base local (comodato + recibo) para a nuvem?\nPode rodar 1 vez; contratos já existentes são atualizados.')) return;
  toast('Importando base local para a nuvem…');
  let total = 0;
  try {
    for (const docKey of ['comodato','recibo']){
      const def = DOCS[docKey];
      let seed = [];
      try { const r = await fetch(def.dataUrl, { cache:'no-store' }); if (r.ok) seed = await r.json(); } catch(_){}
      let local = []; try { local = JSON.parse(localStorage.getItem(def.storeKey) || '[]'); } catch(_){}
      const byC = new Map();
      for (const rec of [...seed, ...local]){ if (clientHasContent(rec)) byC.set(normKey(recVal(rec,'CONTRATO')) || Symbol(), rec); }
      const recs = Array.from(byC.values());

      const { data: existing } = await sb.from('clientes').select('id,contrato').eq('tipo', docKey);
      const exMap = new Map((existing || []).map(e => [normKey(e.contrato || ''), e.id]));

      const ins = [], upd = [];
      for (const rec of recs){
        const c = normKey(recVal(rec,'CONTRATO'));
        const dados = Object.assign({}, rec); delete dados._local;
        const row = { tipo: docKey, contrato: (recVal(rec,'CONTRATO') || '').trim() || null, nome: displayName(rec), cpf: (recVal(rec,'CPF/CNPJ') || recVal(rec,'CPF') || '').trim(), dados };
        if (c && exMap.has(c)) upd.push({ id: exMap.get(c), row }); else ins.push(row);
      }
      for (let i = 0; i < ins.length; i += 200){
        const slice = ins.slice(i, i + 200);
        const { error } = await sb.from('clientes').insert(slice);
        if (error) throw new Error(docKey + ': ' + error.message);
        total += slice.length;
      }
      for (const u of upd){
        const { error } = await sb.from('clientes').update(u.row).eq('id', u.id);
        if (error) throw new Error(docKey + ': ' + error.message);
        total++;
      }
    }
    toast('Importados/atualizados ' + total + ' clientes na nuvem. ✓');
    await loadData('comodato'); await loadData('recibo'); renderList();
  } catch(e){ toast('Erro na importação: ' + e.message, true); }
}

/* --------------------------- Geração DOCX --------------------------- */
/* Testemunhas do contrato — vêm do config local (config.local.js), nunca do repo */
function witnessTokens(){
  const w = (window.APP_CONFIG && window.APP_CONFIG.witnesses) || [];
  return {
    TEST1_NOME: (w[0] && w[0].name) || '', TEST1_CPF: (w[0] && w[0].cpf) || '',
    TEST2_NOME: (w[1] && w[1].name) || '', TEST2_CPF: (w[1] && w[1].cpf) || '',
  };
}

/* Preenche o template DOCX com os valores e devolve um Blob (.docx) */
async function buildDocxBlob(vals, def){
  const resp = await fetch(def.templateUrl, { cache:'no-store' });
  if (!resp.ok) throw new Error('Template não encontrado: ' + def.templateUrl);
  const zip = await JSZip.loadAsync(await resp.arrayBuffer());
  let xml = await zip.file('word/document.xml').async('string');
  const all = Object.assign(witnessTokens(), vals);   // valores do formulário têm prioridade
  Object.keys(all).forEach(tk => { xml = xml.split('{{' + tk + '}}').join(escXml(all[tk])); });
  xml = xml.replace(/\{\{[^}]+\}\}/g, '');   // limpa tokens não preenchidos
  zip.file('word/document.xml', xml);
  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  });
}

async function generateDocx(){
  const def = DOCS[state.doc];
  const vals = collectValues();
  learnFromValues(vals, def);
  const btn = $('#btn-docx');
  btn.disabled = true; const old = btn.textContent; btn.textContent = 'Gerando…';
  try {
    const blob = await buildDocxBlob(vals, def);
    downloadBlob(blob, fileName(vals, def, 'docx'));
    const baixa = deductStock(vals, def);
    logOperacao(vals, def);
    toast('DOCX gerado. ✓' + baixaTxt(baixa));
  } catch(e){
    console.error(e); toast('Erro ao gerar DOCX: ' + e.message, true);
  } finally {
    btn.disabled = false; btn.textContent = old;
  }
}

/* --------------------------- Geração PDF --------------------------- */
/* O PDF é gerado pelo Word (servidor /__pdf), vetorial e fiel ao oficial. */

async function generatePdf(){
  const def = DOCS[state.doc];
  const vals = collectValues();
  learnFromValues(vals, def);
  const btn = $('#btn-pdf'); btn.disabled = true; const old = btn.textContent; btn.textContent = 'Gerando…';
  try {
    const blob = await buildDocxBlob(vals, def);
    let pdfBlob = null;
    // Caminho principal: servidor converte o DOCX em PDF via Word (vetorial, igual ao oficial)
    try {
      const r = await fetch('/__pdf', { method:'POST', headers:{ 'Content-Type':'application/octet-stream' }, body: blob });
      if (!r.ok) throw new Error((await r.text()).slice(0, 200));
      const b = await r.blob();
      if (b.size > 1500) pdfBlob = b; else throw new Error('PDF vazio');
    } catch(srvErr){
      console.warn('PDF via Word/servidor indisponível, usando modo alternativo:', srvErr);
    }
    if (pdfBlob){
      downloadBlob(pdfBlob, fileName(vals, def, 'pdf'));
      const baixa = deductStock(vals, def);
      logOperacao(vals, def);
      toast('PDF gerado. ✓' + baixaTxt(baixa));
    } else {
      toast('PDF precisa do Word + servidor (abra pelo INICIAR.bat). Use “Baixar DOCX” como alternativa.', true);
    }
  } catch(e){
    console.error(e); toast('Erro ao gerar PDF: ' + e.message, true);
  } finally {
    btn.disabled = false; btn.textContent = old;
  }
}

/* Imprime direto: gera o PDF (Word) e manda para a impressora, sem baixar */
async function printDoc(){
  const def = DOCS[state.doc];
  const vals = collectValues();
  learnFromValues(vals, def);
  const btn = $('#btn-print'); btn.disabled = true; const old = btn.textContent; btn.textContent = 'Preparando…';
  try {
    const blob = await buildDocxBlob(vals, def);
    let pdfBlob = null;
    try {
      const r = await fetch('/__pdf', { method:'POST', headers:{ 'Content-Type':'application/octet-stream' }, body: blob });
      if (r.ok){ const b = await r.blob(); if (b.size > 1500) pdfBlob = b; }
    } catch(e){ /* servidor/Word indisponível */ }
    if (!pdfBlob){ toast('Impressão direta precisa do Word/servidor (INICIAR.bat). Use Baixar PDF.', true); return; }
    printBlob(pdfBlob);
    const baixa = deductStock(vals, def);
    logOperacao(vals, def);
    toast('Enviado para impressão. 🖨️' + baixaTxt(baixa));
  } catch(e){
    console.error(e); toast('Erro ao imprimir: ' + e.message, true);
  } finally {
    btn.disabled = false; btn.textContent = old;
  }
}
function printBlob(blob){
  const url = URL.createObjectURL(blob);
  let ifr = document.getElementById('print-frame');
  if (!ifr){
    ifr = document.createElement('iframe'); ifr.id = 'print-frame';
    ifr.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden;';
    document.body.appendChild(ifr);
  }
  ifr.onload = () => {
    try { ifr.contentWindow.focus(); ifr.contentWindow.print(); }
    catch(e){ window.open(url, '_blank'); }   // fallback: abre em nova aba p/ imprimir
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  };
  ifr.src = url;
}

function fileName(vals, def, ext){
  const nome = (vals.CLIENTE_ || vals['RAZÃO_SOCIAL'] || 'cliente').replace(/[\\/:*?"<>|]+/g,' ').trim().slice(0,40);
  const contrato = (vals.CONTRATO || '').replace(/[\\/:*?"<>|]+/g,'').trim();
  return `${def.fileBase}_${contrato ? contrato + '_' : ''}${nome}.${ext}`.replace(/\s+/g,'_');
}

function downloadBlob(blob, name){
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

/* --------------------------- Backup / Restore --------------------------- */
/* Menu do Backup: escolher JSON ou .mdb (Access) */
function showBackupMenu(){
  document.querySelector('#backup-menu')?.remove();
  const btn = $('#btn-export-base');
  const r = btn.getBoundingClientRect();
  const menu = document.createElement('div');
  menu.id = 'backup-menu'; menu.className = 'popmenu';
  menu.innerHTML = `
    <div class="popmenu-title">Exportar backup</div>
    <button data-fmt="json">📄 Backup local (.json)</button>
    <button data-fmt="mdb-comodato">🗄️ Base Comodato (.mdb)</button>
    <button data-fmt="mdb-recibo">🗄️ Base Recibo (.mdb)</button>`;
  document.body.appendChild(menu);
  menu.style.left = Math.round(r.left) + 'px';
  menu.style.bottom = Math.round(window.innerHeight - r.top + 6) + 'px';
  const close = () => { menu.remove(); document.removeEventListener('mousedown', onDoc, true); };
  const onDoc = e => { if (!menu.contains(e.target) && e.target !== btn) close(); };
  setTimeout(() => document.addEventListener('mousedown', onDoc, true), 0);
  menu.querySelectorAll('button').forEach(b => b.addEventListener('click', () => {
    const f = b.dataset.fmt; close();
    if (f === 'json') exportBaseJson();
    else if (f === 'mdb-comodato') exportBaseMdb('comodato');
    else if (f === 'mdb-recibo') exportBaseMdb('recibo');
  }));
}

function exportBaseJson(){
  const data = { comodato: state.clients.comodato.filter(r=>r._local),
                 recibo:   state.clients.recibo.filter(r=>r._local) };
  downloadBlob(new Blob([JSON.stringify(data,null,2)],{type:'application/json'}),
               'infobarra_base_local.json');
  toast('Backup local (.json) exportado. ✓');
}

async function exportBaseMdb(type){
  const records = state.clients[type].map(r => {
    const o = {}; for (const k in r){ if (!k.startsWith('_')) o[k] = r[k]; } return o;
  });
  if (!records.length){ toast('Sem registros de ' + type + ' para exportar.', true); return; }
  toast('Gerando .mdb de ' + (type === 'comodato' ? 'Comodato' : 'Recibo') + '…');
  try {
    const r = await fetch('/__mkmdb', { method:'POST', headers:{ 'Content-Type':'application/json' },
                                        body: JSON.stringify({ type, records }) });
    if (!r.ok) throw new Error((await r.text()).slice(0, 200));
    const blob = await r.blob();
    if (blob.size < 1000) throw new Error('arquivo vazio');
    downloadBlob(blob, `base_${type}.mdb`);
    toast('.mdb de ' + (type === 'comodato' ? 'Comodato' : 'Recibo') + ' exportado. ✓');
  } catch(e){
    console.error(e);
    toast('Erro ao gerar .mdb: ' + e.message + (location.protocol === 'file:' ? ' (rode pelo INICIAR.bat)' : ''), true);
  }
}
function importBase(file){
  const name = (file.name || '').toLowerCase();
  if (name.endsWith('.mdb')) { importMdb(file); return; }   // banco Access
  const fr = new FileReader();
  fr.onload = () => {
    try {
      const data = JSON.parse(fr.result);
      ['comodato','recibo'].forEach(k => {
        if (Array.isArray(data[k])){
          const def = DOCS[k];
          localStorage.setItem(def.storeKey, JSON.stringify(data[k].map(r=>({...r,_local:true}))));
        }
      });
      toast('Base restaurada. Recarregando…');
      setTimeout(()=>location.reload(), 800);
    } catch(e){ toast('Arquivo inválido: ' + e.message, true); }
  };
  fr.readAsText(file);
}

/* Detecta se os registros de um .mdb são de Comodato ou Recibo pelas colunas */
function detectMdbType(records){
  const keys = new Set();
  records.slice(0, 8).forEach(r => Object.keys(r).forEach(k => keys.add(normKey(k))));
  const has = arr => arr.some(k => keys.has(k));
  if (has(['CONTATO','QT01','QT02','FONTE'])) return 'recibo';
  if (has(['ESTADOCIVIL','QUANT01','RAZÃOSOCIAL','ORGEXPEDIDOR','PROFISSÃO'])) return 'comodato';
  return null;
}

/* Importa um banco Access (.mdb) via servidor (Access COM) → base local */
async function importMdb(file){
  toast('Lendo banco Access (.mdb)…');
  try {
    const r = await fetch('/__mdb', { method:'POST', headers:{ 'Content-Type':'application/octet-stream' }, body: file });
    if (!r.ok) throw new Error((await r.text()).slice(0, 200));
    let records = await r.json();
    if (records && !Array.isArray(records)) records = [records];   // 1 registro vem como objeto
    if (!Array.isArray(records) || !records.length) throw new Error('o arquivo não contém registros.');
    const type = detectMdbType(records);
    if (!type){ toast('Não reconheci se o .mdb é de Comodato ou Recibo.', true); return; }
    const def = DOCS[type];
    localStorage.setItem(def.storeKey, JSON.stringify(records.map(rec => ({ ...rec, _local:true }))));
    const label = type === 'comodato' ? 'Comodato' : 'Recibo';
    toast(`${records.length} registros de ${label} importados. Recarregando…`);
    setTimeout(() => location.reload(), 1000);
  } catch(e){
    console.error(e);
    toast('Erro ao importar .mdb: ' + e.message + (location.protocol==='file:' ? ' (rode pelo INICIAR.bat)' : ''), true);
  }
}

/* --------------------------- Troca de documento --------------------------- */
async function switchDoc(docKey){
  state.doc = docKey;
  state.activeIdx = -1;
  $$('.doc-tab').forEach(b => b.classList.toggle('active', b.dataset.doc === docKey));
  if (!state.clients[docKey].length) await loadData(docKey);
  renderForm();
  fillForm(null);
  $('#search').value = '';
  renderList();
}

/* --------------------------- View (Documento / Estoque / Relatórios) --------------------------- */
function setView(v){
  state.view = v;
  const isDoc = v === 'doc';
  $('#doc-form').style.display         = isDoc ? '' : 'none';
  $('#estoque-view').style.display     = v === 'estoque' ? 'flex' : 'none';
  $('#relatorios-view').style.display  = v === 'relatorios' ? 'flex' : 'none';
  $('#action-bar').style.display       = isDoc ? '' : 'none';
  $('#clientes-section').style.display = isDoc ? '' : 'none';
  const foot = document.querySelector('.side-foot'); if (foot) foot.style.display = isDoc ? '' : 'none';
  if (!isDoc){
    $$('.doc-tab').forEach(b => b.classList.toggle('active', b.dataset.view === v));
  }
  if (v === 'estoque'){
    $('#form-title').textContent = 'Estoque de Equipamentos';
    $('#form-sub').textContent = 'Controle de quantidade por marca e modelo.';
    renderEstoque();
  } else if (v === 'relatorios'){
    $('#form-title').textContent = 'Relatórios';
    $('#form-sub').textContent = 'Termos por operador e status de assinatura.';
    renderRelatorios();
  }
}

/* --------------------------- Estoque (local OU Supabase) --------------------------- */
const ESTOQUE_KEY = 'infobarra_estoque_v1';
function loadEstoque(){ return state.estoque || []; }   // cache em memória
function loadEstoqueLocal(){ try { return JSON.parse(localStorage.getItem(ESTOQUE_KEY) || '[]'); } catch(e){ return []; } }
let _estoqueSyncChain = Promise.resolve();
function saveEstoque(arr){
  state.estoque = arr || [];
  if (_supa()){
    const snap = state.estoque.slice();
    _estoqueSyncChain = _estoqueSyncChain.then(() => syncEstoqueRemote(snap)).catch(() => {});
  } else {
    localStorage.setItem(ESTOQUE_KEY, JSON.stringify(state.estoque));
  }
}

/* Lê o estoque da nuvem (tabela estoque) */
async function loadEstoqueRemote(){
  try {
    const { data, error } = await sb.from('estoque').select('marca,modelo,qtd').order('marca');
    if (error) throw error;
    return (data || []).map(r => ({ marca: r.marca, modelo: r.modelo, qtd: parseInt(r.qtd,10) || 0 }));
  } catch(e){ console.warn('loadEstoqueRemote:', e.message); return []; }
}
/* Sincroniza o array do estoque com a tabela (update por id, insere novos, apaga removidos) */
async function syncEstoqueRemote(arr){
  try {
    const { data: rows, error } = await sb.from('estoque').select('id,marca,modelo');
    if (error) throw error;
    const keyOf = (m, mo) => normMarca(m) + '|' + String(mo).toUpperCase().trim();
    const dbByKey = new Map((rows || []).map(r => [keyOf(r.marca, r.modelo), r.id]));
    const seen = new Set();
    for (const it of arr){
      const k = keyOf(it.marca, it.modelo);
      const qtd = parseInt(it.qtd, 10) || 0;
      if (dbByKey.has(k)){
        const id = dbByKey.get(k); seen.add(id);
        const { error: e2 } = await sb.from('estoque').update({ qtd, updated_at: new Date().toISOString() }).eq('id', id);
        if (e2) throw e2;
      } else {
        const { data: ins, error: e3 } = await sb.from('estoque').insert({ marca: it.marca, modelo: it.modelo, qtd }).select('id').maybeSingle();
        if (e3) throw e3;
        if (ins) seen.add(ins.id);
      }
    }
    const toDel = (rows || []).filter(r => !seen.has(r.id)).map(r => r.id);
    if (toDel.length){ const { error: e4 } = await sb.from('estoque').delete().in('id', toDel); if (e4) throw e4; }
  } catch(e){ console.warn('syncEstoqueRemote:', e.message); toast('Falha ao salvar estoque na nuvem.', true); }
}

function renderEstoque(){
  const items = loadEstoque();
  const marcaOpts = state.allMarcas.map(m => `<option value="${escHtml(m)}"></option>`).join('');
  const total = items.reduce((s,i) => s + (parseInt(i.qtd,10)||0), 0);
  const rows = items.length ? items.map((it,idx) => `<tr>
      <td>${escHtml(it.marca)}</td>
      <td>${escHtml(it.modelo)}</td>
      <td class="qtd"><input class="qtd-input" type="number" min="0" value="${escHtml(String(it.qtd))}" data-i="${idx}" title="Clique para editar a quantidade"></td>
      <td class="acts">
        <button class="stock-mini" data-act="dec" data-i="${idx}" title="Diminuir">−</button>
        <button class="stock-mini" data-act="inc" data-i="${idx}" title="Aumentar">+</button>
        <button class="stock-mini del" data-act="del" data-i="${idx}" title="Remover">🗑</button>
      </td></tr>`).join('')
    : `<tr><td colspan="4" class="stock-empty">Nenhum item em estoque. Adicione acima.</td></tr>`;

  $('#estoque-view').innerHTML = `
    <div class="card">
      <div class="card-head"><span class="dot"></span><h2>Adicionar ao estoque</h2></div>
      <div class="stock-add">
        <div class="field"><label>Marca</label>
          <input id="st-marca" list="st-dl-marcas" placeholder="Marca ▾" autocomplete="off"></div>
        <div class="field"><label>Modelo</label>
          <input id="st-modelo" list="st-dl-modelos" placeholder="Modelo ▾" autocomplete="off"></div>
        <div class="field"><label>Quantidade</label>
          <input id="st-qtd" type="number" min="1" value="1"></div>
        <button class="btn-add" id="st-add">+ Adicionar</button>
      </div>
      <datalist id="st-dl-marcas">${marcaOpts}</datalist>
      <datalist id="st-dl-modelos"></datalist>
    </div>
    <div class="card">
      <div class="card-head"><span class="dot"></span><h2>Estoque atual</h2>
        <span class="action-spacer"></span>
        <span class="stock-total">Total: <b>${total}</b> equipamentos</span></div>
      <table class="stock">
        <thead><tr><th>Marca</th><th>Modelo</th><th style="text-align:center">Qtd</th><th></th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;

  const stMarca = $('#st-marca');
  const refreshStModelos = () => {
    const k = normMarca(stMarca.value);
    const ms = (k && state.catalogByKey[k]) ? state.catalogByKey[k] : [];   // vazio sem marca
    $('#st-dl-modelos').innerHTML = ms.map(x => `<option value="${escHtml(x)}"></option>`).join('');
  };
  stMarca.addEventListener('input', refreshStModelos); refreshStModelos();
  $('#st-add').addEventListener('click', addStock);
  $('#st-qtd').addEventListener('keydown', e => { if (e.key === 'Enter') addStock(); });
  $$('#estoque-view .stock-mini').forEach(b =>
    b.addEventListener('click', () => stockAct(b.dataset.act, parseInt(b.dataset.i,10))));
  $$('#estoque-view .qtd-input').forEach(inp =>
    inp.addEventListener('change', () => setStockQty(parseInt(inp.dataset.i,10), inp.value)));
}

/* Dá baixa no estoque dos equipamentos marcados ao gerar um documento.
   Retorna um resumo (string) do que foi descontado, ou ''. */
function deductStock(vals, def){
  const list = [];
  for (const row of def.equip){
    const on = String(vals[row.chk.token]||'').trim().toUpperCase() === 'X';
    const marca = (vals[row.m.token]||'').trim();
    const modelo = (vals[row.mo.token]||'').trim();
    const qty = parseInt(vals[row.q.token], 10) || 0;
    if (on && marca && modelo && qty > 0) list.push({ marca, modelo, qty });
  }
  if (!list.length) return '';
  const sig = (vals.CONTRATO||'') + '|' + state.doc + '|' + JSON.stringify(list);
  if (sig === state.lastDeductSig) return '';   // mesmo documento gerado de novo → não desconta 2×
  const items = loadEstoque();
  const done = [];
  for (const it of list){
    const idx = items.findIndex(s => normMarca(s.marca) === normMarca(it.marca)
                                   && String(s.modelo).toUpperCase().trim() === it.modelo.toUpperCase());
    if (idx >= 0){
      items[idx].qtd = Math.max(0, (parseInt(items[idx].qtd,10)||0) - it.qty);
      done.push(`${it.marca} ${it.modelo} −${it.qty}`);
    }
  }
  if (done.length){ saveEstoque(items); state.lastDeductSig = sig; if (state.view === 'estoque') renderEstoque(); }
  return done.join(', ');
}
function baixaTxt(b){ return b ? ' · baixa estoque: ' + b : ''; }

/* Edita a quantidade direto na tabela (digitar + clicar fora) */
function setStockQty(i, val){
  const items = loadEstoque();
  if (!items[i]) return;
  items[i].qtd = Math.max(0, parseInt(val, 10) || 0);
  saveEstoque(items);
  renderEstoque();
}

function addStock(){
  const marca = $('#st-marca').value.trim();
  const modelo = $('#st-modelo').value.trim();
  const qtd = parseInt($('#st-qtd').value, 10) || 0;
  if (!marca || !modelo){ toast('Informe marca e modelo.', true); return; }
  if (qtd <= 0){ toast('Quantidade inválida.', true); return; }
  const items = loadEstoque();
  const i = items.findIndex(it => normMarca(it.marca) === normMarca(marca)
                                 && String(it.modelo).toUpperCase().trim() === modelo.toUpperCase());
  if (i >= 0) items[i].qtd = (parseInt(items[i].qtd,10)||0) + qtd;
  else items.push({ marca, modelo, qtd });
  saveEstoque(items);
  const aprendeu = learnCatalog(marca, modelo);   // memoriza marca/modelo novo
  renderEstoque();
  toast(aprendeu ? 'Estoque atualizado + catálogo aprendido. ✓' : 'Estoque atualizado. ✓');
}

function stockAct(act, i){
  const items = loadEstoque();
  if (!items[i]) return;
  if (act === 'del'){
    if (!confirm(`Remover "${items[i].marca} ${items[i].modelo}" do estoque?`)) return;
    items.splice(i, 1);
  }
  else if (act === 'inc') items[i].qtd = (parseInt(items[i].qtd,10)||0) + 1;
  else if (act === 'dec') items[i].qtd = Math.max(0, (parseInt(items[i].qtd,10)||0) - 1);
  saveEstoque(items);
  renderEstoque();
}

/* =====================================================================
   RELATÓRIOS — login por operador + termos (assinado / quem criou)
   Dados compartilhados no servidor local (data/store/*.json via /__store);
   pronto para migrar para nuvem trocando storeGet/storeSet.
   ===================================================================== */
const SESSAO_KEY = 'infobarra_sessao';
let _rel = { ops: [], operadores: [], filtros: { q:'', op:'', tipo:'', status:'' } };

async function storeGet(key){
  try { const r = await fetch('/__store?key=' + key, { cache:'no-store' }); if (!r.ok) return null; return await r.json(); }
  catch(e){ return null; }
}
async function storeSet(key, obj){
  try { const r = await fetch('/__store?key=' + key, { method:'POST', headers:{ 'Content-Type':'application/json' }, body: JSON.stringify(obj) }); return r.ok; }
  catch(e){ return false; }
}
async function sha256(s){
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(String(s)));
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2,'0')).join('');
}
function loadSessao(){ try { return JSON.parse(localStorage.getItem(SESSAO_KEY) || 'null'); } catch(e){ return null; } }
function saveSessao(o){ if (o) localStorage.setItem(SESSAO_KEY, JSON.stringify(o)); else localStorage.removeItem(SESSAO_KEY); }
function fmtData(iso){
  if (!iso) return '—';
  const d = new Date(iso); if (isNaN(d)) return '—';
  const p = n => String(n).padStart(2,'0');
  return `${p(d.getDate())}/${p(d.getMonth()+1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

/* Registra/atualiza o termo gerado (1 linha por tipo+contrato) */
async function logOperacao(vals, def){
  const tipo = state.doc;
  const contrato = (vals.CONTRATO || '').trim();
  const cliente = (vals.CLIENTE_ || vals['RAZÃO_SOCIAL'] || '').trim();
  const operador = state.operador ? state.operador.nome : '—';

  if (_supa()){
    try {
      const now = new Date().toISOString();
      const operador_id = state.operador ? state.operador.id : null;
      if (contrato){
        const { data: ex } = await sb.from('operacoes').select('id').eq('tipo', tipo).eq('contrato', contrato).limit(1).maybeSingle();
        if (ex){ await sb.from('operacoes').update({ cliente, operador, operador_id, data: now }).eq('id', ex.id); return; }
      }
      await sb.from('operacoes').insert({ tipo, contrato: contrato || null, cliente, operador, operador_id, data: now });
    } catch(e){ console.warn('logOperacao (supabase):', e.message); }
    return;
  }

  try {
    const ops = (await storeGet('operacoes')) || [];
    const k = tipo + '|' + contrato.toUpperCase();
    const op = contrato ? ops.find(o => (o.tipo + '|' + String(o.contrato||'').toUpperCase()) === k) : null;
    if (op){ op.cliente = cliente || op.cliente; op.operador = operador; op.data = new Date().toISOString(); }
    else ops.push({ id: uid(), tipo, contrato, cliente, operador, data: new Date().toISOString(), assinado:false, assinadoEm:null, anexo:null });
    await storeSet('operacoes', ops);
  } catch(e){ console.warn('logOperacao falhou:', e); }
}

/* Lê as operações (relatórios) da nuvem ou do store local */
async function loadOperacoes(){
  if (_supa()){
    try {
      const { data, error } = await sb.from('operacoes').select('*').order('data', { ascending: false });
      if (error) throw error;
      return (data || []).map(o => ({
        id: o.id, tipo: o.tipo, contrato: o.contrato, cliente: o.cliente,
        operador: o.operador, data: o.data, assinado: !!o.assinado,
        assinadoEm: o.assinado_em, anexo: o.anexo_url,
      }));
    } catch(e){ console.warn('loadOperacoes (supabase):', e.message); return []; }
  }
  return (await storeGet('operacoes')) || [];
}
function uid(){ return (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2)); }

/* --- Operadores / auth ---
   Com Supabase configurado: login = Supabase Auth (e-mail + senha); perfis na
   tabela `perfis`. Sem Supabase: gate local (hash SHA-256 no /__store). */
function _supa(){ return !!(window.hasSupabase && window.hasSupabase()); }

async function getOperadores(){
  if (_supa()){
    try {
      const { data, error } = await sb.from('perfis').select('id,nome,role').order('nome');
      if (error) throw error;
      return (data || []).map(p => ({ id:p.id, nome:p.nome, role:p.role }));
    } catch(e){ console.warn('getOperadores (supabase):', e.message); return []; }
  }
  return (await storeGet('operadores')) || [];
}

/* Carrega nome/role do usuário logado (tabela perfis; fallback p/ metadata/e-mail) */
async function loadPerfilFromSupabase(user){
  let nome = (user.user_metadata && user.user_metadata.nome) || (user.email || '').split('@')[0];
  let role = 'operador';
  try {
    const { data } = await sb.from('perfis').select('nome,role').eq('id', user.id).maybeSingle();
    if (data){ nome = data.nome || nome; role = data.role || role; }
  } catch(e){ /* perfil pode não existir ainda */ }
  return { id:user.id, nome, role, email:user.email };
}

async function criarOperador(nome, senha, role, email){
  if (_supa()){
    email = (email||'').trim();
    nome  = (nome||'').trim();
    if (!email || !senha){ toast('Informe e-mail e senha.', true); return false; }
    try {
      const { data, error } = await sb.functions.invoke('admin-create-user', {
        body: { email, password: senha, nome, role: role || 'operador' },
      });
      let errMsg = null;
      if (error){
        errMsg = error.message;
        try { const j = await error.context.json(); if (j && j.error) errMsg = j.error; } catch(_){}
      } else if (data && data.error){ errMsg = data.error; }
      if (errMsg){ toast('Não criou: ' + errMsg, true); return false; }
      return true;
    } catch(e){ toast('Falha ao criar usuário: ' + e.message, true); return false; }
  }
  nome = (nome||'').trim();
  if (!nome || !senha){ toast('Informe nome e senha.', true); return false; }
  const ops = await getOperadores();
  if (ops.some(x => x.nome.toLowerCase() === nome.toLowerCase())){ toast('Já existe operador com esse nome.', true); return false; }
  ops.push({ id: uid(), nome, hash: await sha256(senha), role: role || 'operador' });
  if (!(await storeSet('operadores', ops))){ toast('Falha ao salvar (rode pelo INICIAR.bat).', true); return false; }
  return true;
}

/* Login. No Supabase: id = e-mail; valida o reCAPTCHA (se configurado) antes de entrar. */
async function loginOperador(id, senha, captchaToken){
  if (_supa()){
    const email = (id||'').trim();
    if (!email || !senha){ toast('Informe e-mail e senha.', true); return false; }
    if (recaptchaEnabled()){
      if (!captchaToken){ toast('Confirme o “Não sou robô”.', true); return false; }
      if (!(await verifyCaptcha(captchaToken))){ toast('Falha no reCAPTCHA — tente de novo.', true); resetCaptcha(); return false; }
    }
    const { data, error } = await sb.auth.signInWithPassword({ email, password: senha });
    if (error){ toast('Login falhou: ' + (error.message || 'verifique e-mail/senha'), true); resetCaptcha(); return false; }
    state.operador = await loadPerfilFromSupabase(data.user);
    toast('Bem-vindo, ' + state.operador.nome + '.');
    onLoggedIn();
    return true;
  }
  const ops = await getOperadores();
  const o = ops.find(x => x.nome.toLowerCase() === (id||'').trim().toLowerCase());
  if (!o){ toast('Operador não encontrado.', true); return false; }
  if (await sha256(senha) !== o.hash){ toast('Senha incorreta.', true); return false; }
  state.operador = { id:o.id, nome:o.nome, role:o.role };
  saveSessao(state.operador);
  toast('Bem-vindo, ' + o.nome + '.');
  onLoggedIn();
  return true;
}

async function logoutOperador(){
  if (_supa()){ try { await sb.auth.signOut(); } catch(e){} }
  state.operador = null; saveSessao(null); toast('Sessão encerrada.'); updateOpStatus(); gateLogin();
}

async function onLoggedIn(){
  $('#login-overlay').classList.remove('show');
  updateOpStatus();
  if (_supa()){
    await loadData('comodato'); await loadData('recibo');
    state.estoque = await loadEstoqueRemote();
    state.catalogUser = await loadUserCatalogRemote(); rebuildCatalog();
    renderList();
    if (state.view === 'estoque') renderEstoque();
  }
  if (state.view === 'relatorios') renderRelatorios();
}

/* --- reCAPTCHA v2 (opcional; só ativa se config.js tiver a site key) --- */
function recaptchaEnabled(){
  // desligado por config (enabled:false) OU sem site key → sem captcha
  if (!(window.RECAPTCHA && window.RECAPTCHA.enabled && window.RECAPTCHA.siteKey)) return false;
  // em localhost (dev) também fica OFF — evita erro de "domínio inválido"
  const h = location.hostname;
  if (h === 'localhost' || h === '127.0.0.1' || h === '') return false;
  return true;
}
let _captchaWidgetId = null;
function renderCaptchaInto(elId){
  if (!recaptchaEnabled()) return;
  const tryRender = () => {
    if (!(window.grecaptcha && grecaptcha.render)) return setTimeout(tryRender, 200);
    const el = document.getElementById(elId);
    if (!el || el.dataset.rendered) return;
    el.dataset.rendered = '1';
    _captchaWidgetId = grecaptcha.render(el, { sitekey: window.RECAPTCHA.siteKey });
  };
  tryRender();
}
function getCaptchaToken(){
  try { return (window.grecaptcha && _captchaWidgetId !== null) ? grecaptcha.getResponse(_captchaWidgetId) : ''; }
  catch(e){ return ''; }
}
function resetCaptcha(){ try { if (window.grecaptcha && _captchaWidgetId !== null) grecaptcha.reset(_captchaWidgetId); } catch(e){} }
async function verifyCaptcha(token){
  try {
    const { data, error } = await sb.functions.invoke('verify-recaptcha', { body:{ token } });
    // function indisponível (ainda não deployada / fora do ar) → não tranca o login
    if (error){ console.warn('verify-recaptcha indisponível, liberando login:', error.message); return true; }
    return !!(data && data.success);   // deployada: vale o que o Google disser
  } catch(e){ console.warn('verifyCaptcha exceção, liberando login:', e.message); return true; }
}

/* Portão de login: mostra overlay até o operador entrar (necessário para criar/editar) */
async function gateLogin(){
  const ov = $('#login-overlay');
  if (state.operador){ ov.classList.remove('show'); updateOpStatus(); return; }

  if (_supa()){
    ov.innerHTML = `<div class="lo-card">
      <img src="assets/infobarra.png" class="lo-logo" alt="Infobarra">
      <h2>Entrar</h2>
      <p class="muted">Acesse com seu e-mail e senha para criar e editar documentos.</p>
      <input id="lg-email" type="email" placeholder="E-mail" autocomplete="username">
      <input id="lg-senha" type="password" placeholder="Senha" autocomplete="current-password">
      <label class="lo-remember"><input type="checkbox" id="lg-remember" ${window.sbRemember() ? 'checked' : ''}> Lembrar login neste dispositivo</label>
      ${recaptchaEnabled() ? '<div id="lg-captcha" class="lo-captcha"></div>' : ''}
      <button class="btn btn-primary" id="lg-go">Entrar</button></div>`;
    ov.classList.add('show');
    const submit = async () => {
      window.sbRemember($('#lg-remember').checked);
      await loginOperador($('#lg-email').value, $('#lg-senha').value, getCaptchaToken());
    };
    $('#lg-go').onclick = submit;
    ['lg-email','lg-senha'].forEach(id => $('#'+id).addEventListener('keydown', e => { if (e.key === 'Enter') submit(); }));
    renderCaptchaInto('lg-captcha');
    setTimeout(() => { const n = $('#lg-email'); if (n) n.focus(); }, 60);
    return;
  }

  // ---- modo local (sem Supabase) ----
  ov.innerHTML = `<div class="lo-card"><p class="muted">Carregando…</p></div>`;
  ov.classList.add('show');
  let ops = null, ok = true;
  try { const r = await fetch('/__store?key=operadores', { cache:'no-store' }); ops = r.ok ? await r.json() : []; }
  catch(e){ ok = false; }
  if (!ok){
    ov.innerHTML = `<div class="lo-card">
      <img src="assets/infobarra.png" class="lo-logo" alt="Infobarra">
      <h2>Servidor indisponível</h2>
      <p class="muted">Para entrar e criar documentos, abra o app pelo <b>INICIAR.bat</b>.</p>
      <button class="btn btn-soft" id="lo-retry">Tentar de novo</button></div>`;
    $('#lo-retry').onclick = gateLogin;
    return;
  }
  _rel.operadores = ops || [];
  const first = _rel.operadores.length === 0;
  ov.innerHTML = `<div class="lo-card">
    <img src="assets/infobarra.png" class="lo-logo" alt="Infobarra">
    <h2>${first ? 'Primeiro acesso' : 'Entrar'}</h2>
    <p class="muted">${first ? 'Crie o operador administrador para começar.' : 'Entre com sua conta para criar e editar documentos.'}</p>
    <input id="lg-nome" placeholder="${first ? 'Seu nome' : 'Operador'}" autocomplete="username">
    <input id="lg-senha" type="password" placeholder="Senha" autocomplete="current-password">
    <button class="btn btn-primary" id="lg-go">${first ? 'Criar administrador' : 'Entrar'}</button></div>`;
  const submit = async () => {
    const nome = $('#lg-nome').value, senha = $('#lg-senha').value;
    if (first){ if (await criarOperador(nome, senha, 'admin')) await loginOperador(nome, senha); }
    else await loginOperador(nome, senha);
  };
  $('#lg-go').onclick = submit;
  ['lg-nome','lg-senha'].forEach(id => $('#'+id).addEventListener('keydown', e => { if (e.key === 'Enter') submit(); }));
  setTimeout(() => { const n = $('#lg-nome'); if (n) n.focus(); }, 60);
}
function updateOpStatus(){
  const el = $('#op-status'); if (!el) return;
  if (state.operador){
    el.innerHTML = `<span class="op-name" title="${escHtml(state.operador.role==='admin'?'Administrador':'Operador')}">👤 ${escHtml(state.operador.nome)}</span><button class="op-logout">Sair</button>`;
    el.querySelector('.op-logout').onclick = logoutOperador;
  } else {
    el.innerHTML = `<button class="op-login">🔒 Entrar para usar</button>`;
    el.querySelector('.op-login').onclick = gateLogin;
  }
}
async function removerOperador(id){
  if (_supa()){
    toast('No modo nuvem, remova operadores em Authentication → Users no painel do Supabase.', true);
    return;
  }
  if (!confirm('Remover este operador?')) return;
  let ops = await getOperadores();
  ops = ops.filter(o => o.id !== id);
  await storeSet('operadores', ops);
  _rel.operadores = ops;
  renderRelatorios();
}

/* --- Render --- */
async function renderRelatorios(){
  const box = $('#relatorios-view');
  if (location.protocol === 'file:'){
    box.innerHTML = `<div class="card"><p class="muted">Os relatórios exigem rodar pelo <b>INICIAR.bat</b> (servidor local).</p></div>`;
    return;
  }
  box.innerHTML = `<div class="card"><p class="muted">Carregando…</p></div>`;
  _rel.operadores = await getOperadores();
  if (!state.operador){
    const first = _rel.operadores.length === 0;
    box.innerHTML = relLoginHtml(first);
    wireLogin(first);
    return;
  }
  _rel.ops = await loadOperacoes();
  box.innerHTML = relReportHtml();
  wireReport();
}

function relLoginHtml(isFirst){
  if (isFirst){
    return `<div class="card rel-login">
      <div class="card-head"><span class="dot"></span><h2>Primeiro acesso</h2></div>
      <p class="muted">Crie o operador <b>administrador</b> para liberar os relatórios.</p>
      <input id="lg-nome" placeholder="Seu nome" autocomplete="off">
      <input id="lg-senha" type="password" placeholder="Senha" autocomplete="new-password">
      <button class="btn btn-primary" data-act="criaradmin">Criar administrador</button>
    </div>`;
  }
  return `<div class="card rel-login">
    <div class="card-head"><span class="dot"></span><h2>Acesso aos Relatórios</h2></div>
    <p class="muted">Entre com seu usuário de operador.</p>
    <input id="lg-nome" placeholder="Operador" autocomplete="username">
    <input id="lg-senha" type="password" placeholder="Senha" autocomplete="current-password">
    <button class="btn btn-primary" data-act="entrar">Entrar</button>
  </div>`;
}
function wireLogin(isFirst){
  const box = $('#relatorios-view');
  const submit = async () => {
    const nome = $('#lg-nome').value, senha = $('#lg-senha').value;
    if (isFirst){ if (await criarOperador(nome, senha, 'admin')) await loginOperador(nome, senha); }
    else await loginOperador(nome, senha);
  };
  box.onclick = e => { if (e.target.closest('[data-act="entrar"],[data-act="criaradmin"]')) submit(); };
  ['lg-nome','lg-senha'].forEach(id => $('#'+id) && $('#'+id).addEventListener('keydown', e => { if (e.key === 'Enter') submit(); }));
}

function relReportHtml(){
  const me = state.operador, isAdmin = me.role === 'admin';
  const nomes = [...new Set(_rel.ops.map(o => o.operador).filter(Boolean))].sort();
  const opOptions = ['<option value="">Todos os operadores</option>'].concat(nomes.map(n => `<option>${escHtml(n)}</option>`)).join('');
  return `
  <div class="card rel-head">
    <span class="rel-badge">${escHtml(me.nome)}</span>
    <span class="rel-role">${isAdmin ? 'Administrador' : 'Operador'}</span>
    <span class="action-spacer"></span>
    ${isAdmin ? '<button class="btn btn-soft" data-act="gerops">👤 Operadores</button>' : ''}
    <button class="btn btn-soft" data-act="logout">Sair</button>
  </div>
  <div id="rel-admin"></div>
  <div class="card">
    <div class="rel-filters">
      <input id="rel-q" placeholder="Buscar contrato ou cliente…" autocomplete="off">
      <select id="rel-op">${opOptions}</select>
      <select id="rel-tipo"><option value="">Todos os tipos</option><option value="comodato">Comodato</option><option value="recibo">Recibo</option></select>
      <select id="rel-status"><option value="">Todos</option><option value="1">Assinados</option><option value="0">Não assinados</option></select>
    </div>
    <table class="rel-table">
      <thead><tr><th>Data</th><th>Operador</th><th>Tipo</th><th>Contrato</th><th>Cliente</th><th>Status</th><th></th></tr></thead>
      <tbody id="rel-tbody"></tbody>
    </table>
    <div class="rel-total" id="rel-total"></div>
  </div>`;
}
function fillRelTable(){
  const f = _rel.filtros;
  let rows = _rel.ops.slice().sort((a,b) => String(b.data||'').localeCompare(String(a.data||'')));
  if (f.op) rows = rows.filter(o => o.operador === f.op);
  if (f.tipo) rows = rows.filter(o => o.tipo === f.tipo);
  if (f.status !== '') rows = rows.filter(o => (o.assinado ? '1' : '0') === f.status);
  if (f.q){ const q = f.q.toLowerCase(); rows = rows.filter(o => (String(o.contrato||'') + ' ' + String(o.cliente||'')).toLowerCase().includes(q)); }
  $('#rel-tbody').innerHTML = rows.length ? rows.map(o => {
    const st = o.assinado ? `<span class="rel-st on">✔ Assinado</span>` : `<span class="rel-st off">○ Não assinado</span>`;
    const link = o.anexo ? ` <button class="rel-anexo" data-act="ver-anexo" data-id="${o.id}" title="Ver PDF assinado">📎</button>` : '';
    return `<tr>
      <td>${fmtData(o.data)}</td>
      <td>${escHtml(o.operador||'—')}</td>
      <td>${o.tipo === 'comodato' ? 'Comodato' : 'Recibo'}</td>
      <td class="mono">${escHtml(o.contrato||'—')}</td>
      <td>${escHtml(o.cliente||'—')}</td>
      <td>${st}${link}</td>
      <td class="rel-acts">
        <button class="stock-mini" data-act="toggle" data-id="${o.id}" title="${o.assinado?'Marcar como NÃO assinado':'Marcar como assinado'}">${o.assinado?'↺':'✔'}</button>
        <button class="stock-mini" data-act="anexar" data-id="${o.id}" title="Anexar PDF assinado">📎</button>
      </td></tr>`;
  }).join('') : `<tr><td colspan="7" class="stock-empty">Nenhum termo registrado ainda. Gere um documento para aparecer aqui.</td></tr>`;
  const tot = rows.length, ass = rows.filter(o => o.assinado).length;
  $('#rel-total').innerHTML = `<b>${tot}</b> termo(s) · <b>${ass}</b> assinado(s) · <b>${tot-ass}</b> pendente(s)`;
}
function wireReport(){
  fillRelTable();
  const f = _rel.filtros;
  $('#rel-q').value = f.q; $('#rel-op').value = f.op; $('#rel-tipo').value = f.tipo; $('#rel-status').value = f.status;
  $('#rel-q').addEventListener('input', e => { f.q = e.target.value; fillRelTable(); });
  $('#rel-op').addEventListener('change', e => { f.op = e.target.value; fillRelTable(); });
  $('#rel-tipo').addEventListener('change', e => { f.tipo = e.target.value; fillRelTable(); });
  $('#rel-status').addEventListener('change', e => { f.status = e.target.value; fillRelTable(); });
  $('#relatorios-view').onclick = async e => {
    const b = e.target.closest('[data-act]'); if (!b) return;
    const act = b.dataset.act, id = b.dataset.id;
    if (act === 'logout') logoutOperador();
    else if (act === 'gerops') toggleAdminPanel();
    else if (act === 'toggle') toggleAssinado(id);
    else if (act === 'anexar') anexarPdf(id);
    else if (act === 'addop') addOperadorFromForm();
    else if (act === 'delop') removerOperador(id);
    else if (act === 'import-clients') importLocalClientsToCloud();
    else if (act === 'ver-anexo') openAnexo(id);
  };
}
async function toggleAssinado(id){
  const o = _rel.ops.find(x => x.id === id); if (!o) return;
  o.assinado = !o.assinado; o.assinadoEm = o.assinado ? new Date().toISOString() : null;
  if (_supa()){
    try { await sb.from('operacoes').update({ assinado: o.assinado, assinado_em: o.assinadoEm }).eq('id', id); }
    catch(e){ toast('Falha ao atualizar na nuvem: ' + e.message, true); }
  } else {
    await storeSet('operacoes', _rel.ops);
  }
  fillRelTable();
}
async function openAnexo(id){
  const o = _rel.ops.find(x => x.id === id); if (!o || !o.anexo) return;
  if (_supa()){
    try {
      const { data, error } = await sb.storage.from('assinados').createSignedUrl(o.anexo, 60 * 60);
      if (error) throw error;
      window.open(data.signedUrl, '_blank');
    } catch(e){ toast('Erro ao abrir anexo: ' + e.message, true); }
  } else {
    window.open(o.anexo, '_blank');
  }
}
function anexarPdf(id){
  const inp = document.createElement('input'); inp.type = 'file'; inp.accept = 'application/pdf,.pdf';
  inp.onchange = async () => {
    const file = inp.files[0]; if (!file) return;
    const o = _rel.ops.find(x => x.id === id); if (!o) return;
    const name = (o.tipo + '_' + (o.contrato || o.id) + '_assinado.pdf').replace(/[^A-Za-z0-9._-]/g, '_');
    toast('Enviando PDF assinado…');
    if (_supa()){
      try {
        const path = o.tipo + '/' + name;
        const { error: upErr } = await sb.storage.from('assinados').upload(path, file, { upsert: true, contentType: 'application/pdf' });
        if (upErr) throw upErr;
        o.anexo = path; o.assinado = true; o.assinadoEm = new Date().toISOString();
        const { error } = await sb.from('operacoes').update({ assinado: true, assinado_em: o.assinadoEm, anexo_url: path }).eq('id', id);
        if (error) throw error;
        fillRelTable(); toast('PDF assinado anexado. ✓');
      } catch(e){ toast('Erro ao anexar: ' + e.message, true); }
      return;
    }
    try {
      const r = await fetch('/__upload?name=' + encodeURIComponent(name), { method:'POST', headers:{ 'Content-Type':'application/pdf' }, body: file });
      if (!r.ok) throw new Error((await r.text()).slice(0,120));
      const j = await r.json();
      o.anexo = j.path; o.assinado = true; o.assinadoEm = new Date().toISOString();
      await storeSet('operacoes', _rel.ops);
      fillRelTable(); toast('PDF assinado anexado. ✓');
    } catch(e){ toast('Erro ao anexar: ' + e.message, true); }
  };
  inp.click();
}
function toggleAdminPanel(){
  const el = $('#rel-admin');
  if (el.innerHTML){ el.innerHTML = ''; return; }
  const list = _rel.operadores.map(o => `<tr>
      <td>${escHtml(o.nome)}</td><td>${o.role === 'admin' ? 'Administrador' : 'Operador'}</td>
      <td class="rel-acts">${o.id !== state.operador.id ? `<button class="stock-mini del" data-act="delop" data-id="${o.id}" title="Remover">🗑</button>` : '<span class="muted">você</span>'}</td>
    </tr>`).join('');
  const cloud = _supa();
  el.innerHTML = `<div class="card">
    <div class="card-head"><span class="dot"></span><h2>Operadores</h2></div>
    <table class="rel-table"><thead><tr><th>Nome</th><th>Perfil</th><th></th></tr></thead><tbody>${list}</tbody></table>
    <p class="muted" style="margin:4px 0 0">Só administradores criam usuários. Marque <b>Administrador</b> para dar acesso total.</p>
    <div class="rel-addop">
      <input id="op-nome" placeholder="Nome" autocomplete="off">
      ${cloud ? '<input id="op-email" type="email" placeholder="E-mail" autocomplete="off">' : ''}
      <input id="op-senha" type="password" placeholder="Senha" autocomplete="new-password">
      <select id="op-role"><option value="operador">Operador</option><option value="admin">Administrador</option></select>
      <button class="btn-add" data-act="addop">+ Adicionar</button>
    </div>
    ${cloud ? '<div style="margin-top:12px;border-top:1px solid var(--line);padding-top:10px;display:flex;gap:10px;align-items:center;flex-wrap:wrap"><button class="btn btn-soft" data-act="import-clients">⬆ Importar clientes locais → nuvem</button><span class="muted" style="font-size:12px">Migração: rode 1 vez.</span></div>' : ''}
  </div>`;
}
async function addOperadorFromForm(){
  const nome = $('#op-nome').value;
  const email = $('#op-email') ? $('#op-email').value : '';
  const senha = $('#op-senha').value, role = $('#op-role').value;
  if (await criarOperador(nome, senha, role, email)){ _rel.operadores = await getOperadores(); toast('Usuário criado. ✓'); renderRelatorios(); }
}

/* --------------------------- Inicialização --------------------------- */
async function restoreSession(){
  if (_supa()){
    try {
      const { data } = await sb.auth.getSession();
      const user = data && data.session && data.session.user;
      state.operador = user ? await loadPerfilFromSupabase(user) : null;
    } catch(e){ state.operador = null; }
  } else {
    state.operador = loadSessao();
  }
}
async function init(){
  await restoreSession();
  // Eventos UI
  $('#doc-switch').addEventListener('click', e => {
    const b = e.target.closest('.doc-tab'); if (!b) return;
    if (b.dataset.view) setView(b.dataset.view);
    else if (b.dataset.doc){ setView('doc'); switchDoc(b.dataset.doc); }
  });
  $('#search').addEventListener('input', () => { state.activeIdx = -1; renderList(); });
  $('#btn-new').addEventListener('click', newClient);
  $('#btn-clear').addEventListener('click', () => { state.activeIdx=-1; fillForm(null); renderList(); });
  $('#btn-save').addEventListener('click', saveClient);
  $('#btn-print').addEventListener('click', printDoc);
  $('#btn-docx').addEventListener('click', generateDocx);
  $('#btn-pdf').addEventListener('click', generatePdf);
  $('#btn-export-base').addEventListener('click', showBackupMenu);
  $('#btn-import-base').addEventListener('click', () => $('#file-import').click());
  $('#file-import').addEventListener('change', e => { if (e.target.files[0]) importBase(e.target.files[0]); });

  await loadData('comodato');
  await loadData('recibo');
  await loadCatalog();
  state.estoque = _supa() ? await loadEstoqueRemote() : loadEstoqueLocal();
  renderForm();
  fillForm(null);
  renderList();
  updateOpStatus();
  await gateLogin();   // exige login para criar/editar (portão do app)
}

document.addEventListener('DOMContentLoaded', init);
