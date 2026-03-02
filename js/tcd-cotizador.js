// ══════════════════════════════════════════════════════════
// TCD COTIZADOR — JS
// ══════════════════════════════════════════════════════════

const CLIENTES_API_URL = 'https://script.google.com/macros/s/AKfycbzRlpra4Oi0fb1Gnx3zXBr3S2MD-VMH4DiYQqr7G3sQI88eJEogYPTVodw5-pIj-G_J/exec';

// ── Estado ──────────────────────────────────────────────
const state = {
  order: [],
  currentProduct: null,
  scannerActive: false,
  html5QrCode: null,
  productos: [],
  clientes: [],
  currentCliente: null,
  // MULTI-FILTRO: objeto { CAMPO: 'valor', ... }
  activeFilters: {},
  searchTimer: null,
  inputMode: 'cam',
  extDebounce: null,
  searchCart: new Map(),
  clienteTimer: null,
};

// ── DOM ──────────────────────────────────────────────────
const $ = (id) => document.getElementById(id);

const els = {
  clienteLoading:   $('clienteLoading'),
  clienteSearchWrap:$('clienteSearchWrap'),
  clienteInput:     $('clienteInput'),
  btnClearCliente:  $('btnClearCliente'),
  clienteResults:   $('clienteResults'),
  clienteCard:      $('clienteCard'),
  ccFantasia:       $('ccFantasia'),
  ccRazons:         $('ccRazons'),
  ccCodigo:         $('ccCodigo'),
  ccZona:           $('ccZona'),
  ccIva:            $('ccIva'),
  ccLocalidad:      $('ccLocalidad'),
  btnChangeCliente: $('btnChangeCliente'),
  eCliente:         $('eCliente'),
  productCard:      $('productCard'),
  pcStateLabel:     $('pcStateLabel'),
  pcDescripcion:    $('pcDescripcion'),
  pcProveedor:      $('pcProveedor'),
  pcGramaje:        $('pcGramaje'),
  pcUxb:            $('pcUxb'),
  btnClearProduct:  $('btnClearProduct'),
  qtyRow:           $('qtyRow'),
  fQty:             $('fQty'),
  eQty:             $('eQty'),
  btnAdd:           $('btnAdd'),
  orderTable:       $('orderTable'),
  orderBody:        $('orderBody'),
  tableEmpty:       $('tableEmpty'),
  totalesPanel:     $('totalesPanel'),
  totalSuper:       $('totalSuper'),
  totalMay:         $('totalMay'),
  totalesAhorro:    $('totalesAhorro'),
  ahorroMonto:      $('ahorroMonto'),
  ahorroPct:        $('ahorroPct'),
  btnExport:        $('btnExport'),
  btnClearOrder:    $('btnClearOrder'),
  badgeCount:       $('badgeCount'),
  toast:            $('toast'),
  toastMsg:         $('toastMsg'),
  toastDot:         $('toastDot'),
  scanOverlay:      $('scanOverlay'),
  btnOpenScanner:   $('btnOpenScanner'),
  btnCloseScanner:  $('btnCloseScanner'),
  btnManualEntry:   $('btnManualEntry'),
  camErr:           $('camErr'),
  sHint:            $('sHint'),
  searchOverlay:    $('searchOverlay'),
  btnOpenSearch:    $('btnOpenSearch'),
  btnCloseSearch:   $('btnCloseSearch'),
  searchInput:      $('searchInput'),
  btnClearSearch:   $('btnClearSearch'),
  searchResults:    $('searchResults'),
  searchStateEmpty: $('searchStateEmpty'),
  searchStateNone:  $('searchStateNone'),
  searchTermDisplay:$('searchTermDisplay'),
  searchCount:      $('searchCount'),
  searchCommitBar:  $('searchCommitBar'),
  scbLabel:         $('scbLabel'),
  btnCommitSearch:  $('btnCommitSearch'),
  modeTabCam:       $('modeTabCam'),
  modeTabExt:       $('modeTabExt'),
  panelCam:         $('panelCam'),
  panelExt:         $('panelExt'),
  extInput:         $('extInput'),
  btnExtClear:      $('btnExtClear'),
  extStatus:        $('extStatus'),
  extStatusText:    $('extStatusText'),
};

// ══════════════════════════════════════════════
// ── Init
// ══════════════════════════════════════════════
async function init() {
  await Promise.all([cargarProductos(), cargarClientes()]);

  const btnTheme = $('btnThemeToggle');
  const applyTheme = (light) => {
    document.documentElement.classList.toggle('light', light);
    btnTheme.textContent = light ? '🌙' : '☀️';
    btnTheme.title = light ? 'Cambiar a modo oscuro' : 'Cambiar a modo claro';
    localStorage.setItem('theme', light ? 'light' : 'dark');
  };
  applyTheme(localStorage.getItem('theme') === 'light');
  btnTheme.addEventListener('click', () => {
    applyTheme(!document.documentElement.classList.contains('light'));
  });

  els.btnExportPDF = $('btnExportPDF');
  els.btnExportPDF.addEventListener('click', exportPDF);
  els.btnExport.addEventListener('click', exportExcel);
  els.btnClearOrder.addEventListener('click', clearOrder);

  els.btnClearProduct.addEventListener('click', resetCurrentProduct);
  els.fQty.addEventListener('input', () => els.eQty.classList.remove('show'));
  els.fQty.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); addToOrder(); } });
  els.btnAdd.addEventListener('click', addToOrder);

  els.btnOpenScanner.addEventListener('click', openScanner);
  els.btnCloseScanner.addEventListener('click', closeScanner);
  els.btnManualEntry.addEventListener('click', closeScanner);

  els.btnOpenSearch.addEventListener('click', openSearchModal);
  els.btnCloseSearch.addEventListener('click', closeSearchModal);
  els.searchOverlay.addEventListener('click', (e) => {
    if (e.target === els.searchOverlay) closeSearchModal();
  });

  els.searchInput.addEventListener('input', onSearchInput);
  els.btnClearSearch.addEventListener('click', () => {
    els.searchInput.value = '';
    els.btnClearSearch.style.display = 'none';
    if (!Object.keys(state.activeFilters).length) {
      document.querySelector('.sf-chip[data-field="all"]').classList.add('sf-chip--active');
    }
    renderSearchResults();
    els.searchInput.focus();
  });

  // ══════════════════════════════════════════════
  // CHIPS MULTI-FILTRO
  // Flujo:
  //  1. Escribís un valor en el input
  //  2. Hacés clic en un chip (Sector, Proveedor, etc.)
  //  3. Ese valor queda "fijado" como filtro para ese campo
  //  4. Podés escribir otro valor y fijar otro campo → filtros AND
  //  5. Clic en un chip activo → lo elimina
  // ══════════════════════════════════════════════
  document.querySelectorAll('.sf-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      const field = chip.dataset.field;

      if (field === 'all') {
        state.activeFilters = {};
        els.searchInput.value = '';
        els.btnClearSearch.style.display = 'none';
        document.querySelectorAll('.sf-chip').forEach(c => {
          c.classList.remove('sf-chip--active');
          c.textContent = c.dataset.label;
        });
        chip.classList.add('sf-chip--active');
        renderSearchResults();
        return;
      }

      // Si el chip ya tiene filtro activo → quitarlo
      if (state.activeFilters[field] !== undefined) {
        delete state.activeFilters[field];
        chip.classList.remove('sf-chip--active');
        chip.textContent = chip.dataset.label;
        if (!Object.keys(state.activeFilters).length && !els.searchInput.value.trim()) {
          document.querySelector('.sf-chip[data-field="all"]').classList.add('sf-chip--active');
        }
        renderSearchResults();
        return;
      }

      // Fijar el texto del input para este campo
      const val = els.searchInput.value.trim();
      if (!val) {
        showToast('wrn', `Escribí un valor y luego presioná "${chip.dataset.label}" para filtrarlo`);
        return;
      }
      state.activeFilters[field] = val;
      chip.classList.add('sf-chip--active');
      chip.textContent = `${chip.dataset.label}: ${val}`;
      document.querySelector('.sf-chip[data-field="all"]').classList.remove('sf-chip--active');

      els.searchInput.value = '';
      els.btnClearSearch.style.display = 'none';
      els.searchInput.focus();
      renderSearchResults();
    });
  });

  els.btnCommitSearch.addEventListener('click', commitSearchCart);

  function setInputMode(mode) {
    state.inputMode = mode;
    localStorage.setItem('inputMode', mode);
    const isCam = mode === 'cam';
    els.modeTabCam.classList.toggle('mode-tab--active', isCam);
    els.modeTabExt.classList.toggle('mode-tab--active', !isCam);
    els.panelCam.style.display = isCam ? '' : 'none';
    els.panelExt.style.display = isCam ? 'none' : '';
    if (!isCam) {
      resetCurrentProduct();
      setTimeout(() => els.extInput.focus(), 80);
      setExtStatus('ready', 'Listo — esperando lectura');
    }
  }
  setInputMode(localStorage.getItem('inputMode') || 'cam');
  els.modeTabCam.addEventListener('click', () => setInputMode('cam'));
  els.modeTabExt.addEventListener('click', () => setInputMode('ext'));

  els.extInput.addEventListener('input', () => {
    const val = els.extInput.value.trim();
    els.btnExtClear.style.display = val ? 'flex' : 'none';
    if (!val) { setExtStatus('ready', 'Listo — esperando lectura'); return; }
    setExtStatus('reading', 'Leyendo código...');
    els.extInput.classList.add('ext-input--reading');
    els.extInput.classList.remove('ext-input--found');
    clearTimeout(state.extDebounce);
    state.extDebounce = setTimeout(() => triggerExtLookup(val), 120);
  });
  els.extInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      clearTimeout(state.extDebounce);
      const val = els.extInput.value.trim();
      if (val) triggerExtLookup(val);
    }
  });
  els.btnExtClear.addEventListener('click', () => {
    els.extInput.value = '';
    els.btnExtClear.style.display = 'none';
    els.extInput.classList.remove('ext-input--reading', 'ext-input--found');
    resetCurrentProduct();
    setExtStatus('ready', 'Listo — esperando lectura');
    els.extInput.focus();
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if (els.searchOverlay.classList.contains('open')) closeSearchModal();
      if (els.scanOverlay.classList.contains('open')) closeScanner();
    }
    if (state.inputMode === 'ext'
      && !els.searchOverlay.classList.contains('open')
      && !els.scanOverlay.classList.contains('open')
      && document.activeElement !== els.extInput
      && document.activeElement !== els.fQty
      && e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
      els.extInput.focus();
    }
  });
}

document.addEventListener('DOMContentLoaded', init);

// ══════════════════════════════════════════════
// ── Clientes
// ══════════════════════════════════════════════
const CLIENTES_CACHE_KEY = 'tcd_clientes_v1';
const CLIENTES_CACHE_HRS = 1;

async function cargarClientes() {
  try {
    const raw = localStorage.getItem(CLIENTES_CACHE_KEY);
    if (raw) {
      const { data, ts } = JSON.parse(raw);
      const ageHrs = (Date.now() - ts) / 3_600_000;
      if (ageHrs < CLIENTES_CACHE_HRS && data?.length > 0) {
        state.clientes = data.filter(c => !c.inactivo);
        initClienteSelector(`${state.clientes.length} clientes`);
        return;
      }
    }
  } catch(e) {}

  try {
    const res  = await fetch(CLIENTES_API_URL + '?action=getClientes');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || 'Error en API');
    state.clientes = (json.data || []).filter(c => !c.inactivo);
    localStorage.setItem(CLIENTES_CACHE_KEY, JSON.stringify({ data: json.data || [], ts: Date.now() }));
    initClienteSelector(`${state.clientes.length} clientes`);
  } catch (err) {
    console.error('Error al cargar clientes:', err);
    state.clientes = [];
    initClienteSelector(null, err.message);
  }
}

function selectCliente(cliente) {
  state.currentCliente = cliente;
  els.clienteCard.style.display       = 'block';
  els.clienteSearchWrap.style.display = 'none';
  els.clienteResults.style.display    = 'none';
  els.ccFantasia.textContent  = cliente.fantasia  || cliente.razons || '—';
  els.ccRazons.textContent    = cliente.razons    || '—';
  els.ccCodigo.textContent    = cliente.codigo    || '—';
  els.ccZona.textContent      = cliente.zona      || '—';
  els.ccIva.textContent       = cliente.iva       || '—';
  els.ccLocalidad.textContent = cliente.localidad || '—';
  els.eCliente.classList.remove('show');
}

async function refrescarClientes() {
  localStorage.removeItem(CLIENTES_CACHE_KEY);
  await cargarClientes();
}

function initClienteSelector(successMsg, errorMsg) {
  els.clienteLoading.style.display    = 'none';
  els.clienteSearchWrap.style.display = 'block';

  const statusEl = document.getElementById('clienteStatus');
  if (statusEl) {
    if (errorMsg) {
      statusEl.textContent   = `⚠ No se pudieron cargar los clientes: ${errorMsg}`;
      statusEl.className     = 'cliente-status cliente-status--err';
      statusEl.style.display = 'block';
    } else if (successMsg) {
      statusEl.textContent   = `✓ ${successMsg}`;
      statusEl.className     = 'cliente-status cliente-status--ok';
      statusEl.style.display = 'block';
      setTimeout(() => { statusEl.style.display = 'none'; }, 3000);
    }
  }

  els.clienteInput.addEventListener('input', () => {
    const val = els.clienteInput.value;
    els.btnClearCliente.style.display = val ? 'flex' : 'none';
    clearTimeout(state.clienteTimer);
    state.clienteTimer = setTimeout(() => renderClienteResults(val), 100);
  });
  els.clienteInput.addEventListener('focus', () => {
    if (els.clienteInput.value.length >= 1) renderClienteResults(els.clienteInput.value);
  });
  els.btnClearCliente.addEventListener('click', () => {
    els.clienteInput.value = '';
    els.btnClearCliente.style.display = 'none';
    els.clienteResults.style.display  = 'none';
    els.clienteInput.focus();
  });
  els.btnChangeCliente.addEventListener('click', () => {
    state.currentCliente = null;
    els.clienteCard.style.display       = 'none';
    els.clienteSearchWrap.style.display = 'block';
    els.clienteInput.value              = '';
    els.btnClearCliente.style.display   = 'none';
    els.clienteResults.style.display    = 'none';
    setTimeout(() => els.clienteInput.focus(), 80);
  });
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.cliente-search-wrap')) {
      els.clienteResults.style.display = 'none';
    }
  });
}

function renderClienteResults(query) {
  const q = normalize(query);
  if (q.length === 0) { els.clienteResults.style.display = 'none'; return; }
  if (!state.clientes.length) {
    els.clienteResults.innerHTML = `<li class="cr-empty">⚠ No hay clientes cargados.</li>`;
    els.clienteResults.style.display = 'block';
    return;
  }
  const matches = state.clientes.filter(c =>
    normalize(c.fantasia).includes(q) || normalize(c.razons).includes(q)
  ).slice(0, 50);

  if (!matches.length) {
    els.clienteResults.innerHTML = `<li class="cr-empty">Sin resultados para "<strong>${esc(query)}</strong>"</li>`;
    els.clienteResults.style.display = 'block';
    return;
  }
  els.clienteResults.innerHTML = matches.map(c => {
    const nombre = c.fantasia || c.razons || c.codigo;
    const razons = (c.fantasia && c.razons && c.fantasia !== c.razons) ? c.razons : '';
    return `
      <li class="cr-item" tabindex="0" data-codigo="${esc(c.codigo)}">
        <div class="cr-main">
          <div class="cr-fantasia">${highlight(nombre, q)}</div>
          ${razons ? `<div class="cr-razons">${highlight(razons, q)}</div>` : ''}
          <div class="cr-meta">
            ${c.codigo    ? `<span class="cr-chip">${esc(c.codigo)}</span>` : ''}
            ${c.localidad ? `<span class="cr-chip cr-chip--loc">${esc(c.localidad)}</span>` : ''}
            ${c.zona      ? `<span class="cr-chip cr-chip--zona">${esc(c.zona)}</span>` : ''}
            ${c.iva       ? `<span class="cr-chip cr-chip--iva">${esc(c.iva)}</span>` : ''}
          </div>
        </div>
      </li>`;
  }).join('');
  els.clienteResults.style.display = 'block';
  els.clienteResults.querySelectorAll('.cr-item').forEach(li => {
    const codigo = li.dataset.codigo;
    const seleccionar = () => {
      const cliente = state.clientes.find(c => c.codigo === codigo);
      if (cliente) selectCliente(cliente);
    };
    li.addEventListener('click', seleccionar);
    li.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); seleccionar(); }
      if (e.key === 'ArrowDown') { e.preventDefault(); (li.nextElementSibling || li).focus(); }
      if (e.key === 'ArrowUp')   { e.preventDefault(); (li.previousElementSibling || li).focus(); }
    });
  });
}

// ══════════════════════════════════════════════
// ── Productos
// ══════════════════════════════════════════════
async function cargarProductos() {
  try {
    const res  = await fetch('../data/productos.json');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    state.productos = json.data || [];
  } catch (err) {
    console.error('Error al cargar productos.json:', err);
    showToast('err', 'No se pudo cargar la base de productos.');
  }
}

function normalizeCode(val) {
  if (val == null || val === '') return '';
  return String(val).trim().replace(/^0+(?=\d)/, '');
}

async function fetchProducto(code) {
  const normCode = normalizeCode(code);
  return state.productos.find(p => {
    const normEAN = normalizeCode(p.EAN);
    const normINT = normalizeCode(p.INTERNO);
    return normEAN === normCode || normINT === normCode;
  }) || null;
}

function cartKey(p) {
  return `${p.EAN ?? ''}_${p.INTERNO ?? ''}_${normalize(p.DESCRIPCION ?? '')}`;
}

async function triggerExtLookup(code) {
  els.extInput.classList.remove('ext-input--reading', 'ext-input--found');
  setExtStatus('reading', `Buscando: ${code}`);
  showCardLoading();
  const data = await fetchProducto(code);
  if (data) {
    setCurrentProduct(data);
    els.extInput.classList.add('ext-input--found');
    setExtStatus('found', `✓ ${data.DESCRIPCION ? data.DESCRIPCION.slice(0, 40) : 'Producto encontrado'}`);
    showToast('ok', 'Producto encontrado');
    setTimeout(() => els.fQty.focus(), 150);
  } else {
    state.currentProduct = null;
    showCardNotFound();
    hideQtyRow();
    setExtStatus('notfound', '⚠ Código no encontrado — intentá la búsqueda manual');
    showToast('wrn', 'Código no encontrado — intentá buscarlo manualmente.');
  }
}

function setExtStatus(type, text) {
  els.extStatus.className = 'ext-status';
  if (type) els.extStatus.classList.add(`ext-status--${type}`);
  els.extStatusText.textContent = text;
}

async function triggerLookup(code) {
  showCardLoading();
  const data = await fetchProducto(code);
  if (data) {
    setCurrentProduct(data);
    showToast('ok', 'Producto encontrado');
    setTimeout(() => els.fQty.focus(), 150);
  } else {
    state.currentProduct = null;
    showCardNotFound();
    hideQtyRow();
    showToast('wrn', 'Código no encontrado — intentá buscarlo manualmente.');
  }
}

function setCurrentProduct(data) {
  state.currentProduct = data;
  fillProductCard(data);
  showCardFound();
  showQtyRow();
}

function fillProductCard(data) {
  els.pcDescripcion.textContent = data.DESCRIPCION || '—';
  els.pcProveedor.textContent   = data.PROVEEDOR   || '—';
  els.pcGramaje.textContent     = data.GRAMAJE      || '—';
  els.pcUxb.textContent         = data.UXB != null  ? `${data.UXB} u.` : '—';
  const pvp    = data['PVP SUPER'];
  const pvpMay = data['PVP MAYORISTA'];
  const pvpEl    = $('pcPvp');
  const pvpMayEl = $('pcPvpMay');
  if (pvpEl)    pvpEl.textContent    = pvp    != null ? `$${fmt(pvp)}`    : '—';
  if (pvpMayEl) pvpMayEl.textContent = pvpMay != null ? `$${fmt(pvpMay)}` : '—';
}

function resetCurrentProduct() {
  state.currentProduct = null;
  hideProductCard();
  hideQtyRow();
  els.fQty.value = '';
}

function showCardLoading() {
  els.productCard.classList.add('visible');
  els.productCard.dataset.state = 'loading';
  els.pcStateLabel.textContent   = 'Buscando...';
  els.btnClearProduct.style.display = 'none';
}
function showCardFound() {
  els.productCard.classList.add('visible');
  els.productCard.dataset.state = 'found';
  els.pcStateLabel.textContent   = 'Producto encontrado';
  els.btnClearProduct.style.display = 'flex';
}
function showCardNotFound() {
  els.productCard.classList.add('visible');
  els.productCard.dataset.state = 'notfound';
  els.pcStateLabel.textContent   = 'No encontrado';
  els.btnClearProduct.style.display = 'flex';
}
function hideProductCard() {
  els.productCard.classList.remove('visible');
  els.productCard.dataset.state = '';
  els.btnClearProduct.style.display = 'none';
}
function showQtyRow() { els.qtyRow.style.display = 'flex'; }
function hideQtyRow() { els.qtyRow.style.display = 'none'; }

// ══════════════════════════════════════════════
// ── Search Modal — MULTI-FILTRO (AND)
// ══════════════════════════════════════════════
function openSearchModal() {
  state.searchCart.clear();
  state.activeFilters = {};
  els.searchOverlay.classList.add('open');
  document.body.style.overflow = 'hidden';
  els.searchInput.value = '';
  els.btnClearSearch.style.display = 'none';
  els.searchCount.textContent = '';
  document.querySelectorAll('.sf-chip').forEach(c => {
    c.classList.remove('sf-chip--active');
    c.textContent = c.dataset.label;
  });
  document.querySelector('.sf-chip[data-field="all"]').classList.add('sf-chip--active');
  renderSearchResults();
  updateCommitBar();
  setTimeout(() => els.searchInput.focus(), 120);
}

function closeSearchModal() {
  els.searchOverlay.classList.remove('open');
  document.body.style.overflow = '';
}

function onSearchInput() {
  const query = els.searchInput.value;
  els.btnClearSearch.style.display = query.length ? 'flex' : 'none';
  if (query.trim()) {
    document.querySelector('.sf-chip[data-field="all"]').classList.remove('sf-chip--active');
  } else if (!Object.keys(state.activeFilters).length) {
    document.querySelector('.sf-chip[data-field="all"]').classList.add('sf-chip--active');
  }
  clearTimeout(state.searchTimer);
  state.searchTimer = setTimeout(() => renderSearchResults(), 150);
}

function renderSearchResults() {
  const inputQuery   = normalize(els.searchInput.value.trim());
  const activeFields = Object.entries(state.activeFilters);
  const hasAny       = activeFields.length > 0 || inputQuery.length >= 2;

  renderFilterTags();

  if (!hasAny) {
    els.searchStateEmpty.style.display = 'flex';
    els.searchStateNone.style.display  = 'none';
    els.searchResults.style.display    = 'none';
    els.searchCount.textContent        = '';
    return;
  }

  // Filtros fijados (AND) + texto libre adicional
  const matches = state.productos.filter(p => {
    for (const [field, val] of activeFields) {
      const q = normalize(val);
      const pass = field === 'EAN'
        ? normalize(String(p.EAN ?? '')).includes(q) || normalize(String(p.INTERNO ?? '')).includes(q)
        : normalize(String(p[field] ?? '')).includes(q);
      if (!pass) return false;
    }
    if (inputQuery.length >= 2) {
      return normalize(p.DESCRIPCION).includes(inputQuery) ||
             normalize(p.PROVEEDOR  ).includes(inputQuery) ||
             normalize(p.SECTOR  ?? '').includes(inputQuery) ||
             normalize(p.SECCION ?? '').includes(inputQuery) ||
             normalize(String(p.EAN     ?? '')).includes(inputQuery) ||
             normalize(String(p.INTERNO ?? '')).includes(inputQuery);
    }
    return true;
  });

  els.searchStateEmpty.style.display = 'none';

  if (!matches.length) {
    els.searchStateNone.style.display  = 'flex';
    els.searchTermDisplay.textContent  = buildFilterDescription();
    els.searchResults.style.display    = 'none';
    els.searchCount.textContent        = 'Sin resultados';
    return;
  }

  els.searchStateNone.style.display = 'none';
  els.searchResults.style.display   = 'block';

  const shown = matches.slice(0, 80);
  const total = matches.length;
  els.searchCount.textContent = total > 80
    ? `${total} resultados — mostrando los primeros 80`
    : `${total} resultado${total !== 1 ? 's' : ''}`;

  const q = inputQuery;

  els.searchResults.innerHTML = shown.map((p, i) => {
    const key    = cartKey(p);
    const inCart = state.searchCart.has(key);
    const qtyVal = inCart ? (state.searchCart.get(key).qty || '') : '';

    const desc    = highlight(p.DESCRIPCION || '—', q);
    const prov    = highlight(p.PROVEEDOR   || '—', q);
    const eanRaw  = String(p.EAN     ?? '');
    const intRaw  = String(p.INTERNO ?? '');
    const eanHL   = highlight(eanRaw, q);
    const intHL   = highlight(intRaw, q);

    const eanBadge = eanRaw ? `<span class="sr-ean">EAN ${eanHL}</span>` : '';
    const intBadge = intRaw && intRaw !== eanRaw ? `<span class="sr-ean sr-ean--int">INT ${intHL}</span>` : '';
    const sect     = p.SECTOR  ? `<span class="sr-tag">${esc(p.SECTOR)}</span>`  : '';
    const secc     = p.SECCION ? `<span class="sr-tag">${esc(p.SECCION)}</span>` : '';
    const gramaje  = p.GRAMAJE ? `<span class="sr-tag sr-tag--gramaje">${esc(p.GRAMAJE)}</span>` : '';
    const pvp      = p['PVP SUPER']     != null ? `<span class="sr-tag sr-tag--pvp">Super $${fmt(p['PVP SUPER'])}</span>`        : '';
    const pvpMay   = p['PVP MAYORISTA'] != null ? `<span class="sr-tag sr-tag--pvp-may">May. $${fmt(p['PVP MAYORISTA'])}</span>` : '';

    return `
      <li class="sr-item${inCart ? ' sr-item--selected' : ''}" role="option" tabindex="0" data-idx="${i}">
        <div class="sr-check">
          <svg class="sr-check-icon" viewBox="0 0 18 18" fill="none">
            <rect x="1" y="1" width="16" height="16" rx="4" stroke="currentColor" stroke-width="1.8"/>
            <path class="sr-check-tick" d="M4.5 9l3 3 6-6" stroke="currentColor" stroke-width="2"
                  stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        </div>
        <div class="sr-main">
          <div class="sr-desc">${desc}</div>
          <div class="sr-meta">
            <span class="sr-prov">${prov}</span>
            ${eanBadge}${intBadge}
          </div>
        </div>
        <div class="sr-right">
          <div class="sr-tags">${gramaje}${sect}${secc}${pvp}${pvpMay}</div>
          <div class="sr-qty-wrap${inCart ? ' sr-qty-wrap--visible' : ''}">
            <input type="number" class="sr-qty-input" data-key="${esc(key)}"
              value="${qtyVal}" placeholder="Cant." min="1" step="1"
              inputmode="numeric" aria-label="Cantidad"/>
          </div>
        </div>
      </li>`;
  }).join('');

  els.searchResults.querySelectorAll('.sr-item').forEach((li, i) => {
    const p   = shown[i];
    const key = cartKey(p);
    li.addEventListener('click', (e) => {
      if (e.target.closest('.sr-qty-input')) return;
      toggleCartItem(p, key, li);
    });
    li.addEventListener('keydown', (e) => {
      if ((e.key === 'Enter' || e.key === ' ') && !e.target.closest('.sr-qty-input')) {
        e.preventDefault(); toggleCartItem(p, key, li);
      }
      if (e.key === 'ArrowDown') { e.preventDefault(); (li.nextElementSibling || li).focus(); }
      if (e.key === 'ArrowUp')   { e.preventDefault(); (li.previousElementSibling || li).focus(); }
    });
    const qtyInput = li.querySelector('.sr-qty-input');
    if (qtyInput) {
      qtyInput.addEventListener('click',   (e) => e.stopPropagation());
      qtyInput.addEventListener('keydown', (e) => e.stopPropagation());
      qtyInput.addEventListener('input', () => {
        const val = parseFloat(qtyInput.value);
        if (state.searchCart.has(key)) {
          state.searchCart.get(key).qty = (!isNaN(val) && val > 0) ? val : 0;
          updateCommitBar();
        }
      });
    }
  });
}

// ── Tags removibles de filtros activos ──
function renderFilterTags() {
  let tagsEl = document.getElementById('activeFilterTags');
  if (!tagsEl) {
    tagsEl = document.createElement('div');
    tagsEl.id = 'activeFilterTags';
    tagsEl.className = 'active-filter-tags';
    const wrap = document.querySelector('.search-results-wrap');
    if (wrap) wrap.parentElement.insertBefore(tagsEl, wrap);
  }
  const entries = Object.entries(state.activeFilters);
  if (!entries.length) { tagsEl.innerHTML = ''; return; }

  tagsEl.innerHTML = `
    <div class="aft-row">
      <span class="aft-hint">Filtros activos:</span>
      ${entries.map(([field, val]) => `
        <button class="aft-tag" data-field="${esc(field)}">
          <span class="aft-label">${esc(field)}:</span>
          <strong>${esc(val)}</strong>
          <span class="aft-x">✕</span>
        </button>`).join('')}
      <button class="aft-clear-all">Limpiar todo</button>
    </div>`;

  tagsEl.querySelectorAll('.aft-tag').forEach(btn => {
    btn.addEventListener('click', () => {
      const f = btn.dataset.field;
      delete state.activeFilters[f];
      const chip = document.querySelector(`.sf-chip[data-field="${f}"]`);
      if (chip) { chip.classList.remove('sf-chip--active'); chip.textContent = chip.dataset.label; }
      if (!Object.keys(state.activeFilters).length && !els.searchInput.value.trim()) {
        document.querySelector('.sf-chip[data-field="all"]').classList.add('sf-chip--active');
      }
      renderSearchResults();
    });
  });

  tagsEl.querySelector('.aft-clear-all')?.addEventListener('click', () => {
    state.activeFilters = {};
    els.searchInput.value = '';
    els.btnClearSearch.style.display = 'none';
    document.querySelectorAll('.sf-chip').forEach(c => {
      c.classList.remove('sf-chip--active');
      c.textContent = c.dataset.label;
    });
    document.querySelector('.sf-chip[data-field="all"]').classList.add('sf-chip--active');
    renderSearchResults();
  });
}

function buildFilterDescription() {
  const parts = Object.entries(state.activeFilters).map(([f, v]) => `${f}="${v}"`);
  if (els.searchInput.value.trim()) parts.push(`"${els.searchInput.value.trim()}"`);
  return parts.join(' + ');
}

function toggleCartItem(producto, key, liEl) {
  if (state.searchCart.has(key)) {
    state.searchCart.delete(key);
    liEl.classList.remove('sr-item--selected');
    liEl.querySelector('.sr-qty-wrap')?.classList.remove('sr-qty-wrap--visible');
    const inp = liEl.querySelector('.sr-qty-input');
    if (inp) inp.value = '';
  } else {
    state.searchCart.set(key, { producto, qty: 0 });
    liEl.classList.add('sr-item--selected');
    liEl.querySelector('.sr-qty-wrap')?.classList.add('sr-qty-wrap--visible');
    const inp = liEl.querySelector('.sr-qty-input');
    if (inp) setTimeout(() => { inp.focus(); inp.select(); }, 50);
  }
  if (navigator.vibrate) navigator.vibrate(25);
  updateCommitBar();
}

function updateCommitBar() {
  const total  = state.searchCart.size;
  const conQty = [...state.searchCart.values()].filter(e => e.qty > 0).length;
  const sinQty = total - conQty;
  els.searchCommitBar.classList.toggle('has-items', total > 0);
  if (!total) {
    els.scbLabel.textContent     = '0 productos seleccionados';
    els.btnCommitSearch.disabled = true;
    els.btnCommitSearch.innerHTML = btnCommitHTML('Agregar a la lista');
    return;
  }
  els.scbLabel.textContent = `${total} seleccionado${total !== 1 ? 's' : ''}` +
    (sinQty > 0 ? ` · ${sinQty} sin cantidad` : '');
  els.btnCommitSearch.disabled  = conQty === 0;
  els.btnCommitSearch.innerHTML = btnCommitHTML(
    conQty > 0 ? `Agregar ${conQty} a la lista` : 'Agregá las cantidades'
  );
}

function btnCommitHTML(label) {
  return `${label}
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"
         stroke-linecap="round" stroke-linejoin="round" width="15" height="15">
      <line x1="12" y1="5" x2="12" y2="19"/>
      <line x1="5" y1="12" x2="19" y2="12"/>
    </svg>`;
}

function commitSearchCart() {
  const toAdd = [...state.searchCart.values()].filter(e => e.qty > 0);
  if (!toAdd.length) return;
  let added = 0, updated = 0;
  toAdd.forEach(({ producto: p, qty }) => {
    const existing = findOrderItem(p);
    if (existing) { existing.cantidad += qty; updated++; }
    else { state.order.push(buildOrderItem(p, qty)); added++; }
  });
  renderTable(); updateBadge();
  const parts = [];
  if (added)   parts.push(`${added} nuevo${added !== 1 ? 's' : ''}`);
  if (updated) parts.push(`${updated} actualizado${updated !== 1 ? 's' : ''}`);
  showToast('ok', parts.join(' · ') + ' en la lista');
  if (navigator.vibrate) navigator.vibrate([60, 30, 60]);
  state.searchCart.clear();
  closeSearchModal();
}

function buildOrderItem(p, cantidad) {
  return {
    id:          Date.now() + Math.random(),
    ean:         p.EAN         || p.INTERNO || '',
    interno:     p.INTERNO     || '',
    descripcion: p.DESCRIPCION || '',
    proveedor:   p.PROVEEDOR   || '',
    gramaje:     p.GRAMAJE     || '',
    uxb:         p.UXB != null ? p.UXB : '',
    pvpSuper:    p['PVP SUPER']      != null ? Number(p['PVP SUPER'])      : null,
    pvpMay:      p['PVP MAYORISTA']  != null ? Number(p['PVP MAYORISTA'])  : null,
    cantidad,
  };
}

// ══════════════════════════════════════════════
// ── Pedido
// ══════════════════════════════════════════════
function addToOrder() {
  const p   = state.currentProduct;
  const qty = parseFloat(els.fQty.value.replace(',', '.'));
  if (isNaN(qty) || qty <= 0) { els.eQty.classList.add('show'); els.fQty.focus(); return; }
  const existing = findOrderItem(p);
  if (existing) {
    existing.cantidad += qty;
    renderTable(); updateBadge();
    showToast('ok', `"${p?.DESCRIPCION || ''}" — cantidad actualizada (total: ×${existing.cantidad})`);
  } else {
    state.order.push(buildOrderItem(p, qty));
    renderTable(); updateBadge();
    showToast('ok', `"${p?.DESCRIPCION || ''}" agregado (×${qty})`);
  }
  if (navigator.vibrate) navigator.vibrate(60);
  resetCurrentProduct();
  els.fQty.value = '';
  if (state.inputMode === 'ext') {
    els.extInput.value = '';
    els.btnExtClear.style.display = 'none';
    els.extInput.classList.remove('ext-input--reading', 'ext-input--found');
    setExtStatus('ready', 'Listo — esperando lectura');
    setTimeout(() => els.extInput.focus(), 80);
  }
}

function removeItem(id) {
  state.order = state.order.filter(i => i.id !== id);
  renderTable(); updateBadge();
}

function clearOrder() {
  if (!state.order.length) return;
  if (!confirm('¿Limpiar toda la lista del pedido?')) return;
  state.order = [];
  renderTable(); updateBadge();
  showToast('wrn', 'Lista vaciada');
}

function findOrderItem(p) {
  const ean     = String(p?.EAN    || p?.INTERNO || '').trim();
  const interno = String(p?.INTERNO || '').trim();
  return state.order.find(item => {
    if (ean     && String(item.ean).trim()     === ean)     return true;
    if (interno && String(item.interno).trim() === interno) return true;
    return false;
  }) || null;
}

function renderTable() {
  const has = state.order.length > 0;
  els.tableEmpty.style.display   = has ? 'none'  : 'block';
  els.orderTable.style.display   = has ? 'table' : 'none';
  els.totalesPanel.style.display = has ? 'block' : 'none';
  els.btnExport.disabled         = !has;
  els.btnExportPDF.disabled      = !has;
  els.btnClearOrder.disabled     = !has;

  let sumSuper = 0, sumMay = 0;
  els.orderBody.innerHTML = state.order.map(item => {
    const subSuper = item.pvpSuper != null ? item.pvpSuper * item.cantidad : null;
    const subMay   = item.pvpMay   != null ? item.pvpMay   * item.cantidad : null;
    if (subSuper != null) sumSuper += subSuper;
    if (subMay   != null) sumMay   += subMay;
    const fmtSuper    = item.pvpSuper != null ? `$${fmt(item.pvpSuper)}` : '—';
    const fmtMay      = item.pvpMay   != null ? `$${fmt(item.pvpMay)}`   : '—';
    const fmtSubSuper = subSuper      != null ? `$${fmt(subSuper)}`       : '—';
    const fmtSubMay   = subMay        != null ? `$${fmt(subMay)}`         : '—';
    return `
      <tr>
        <td class="td-code">${esc(String(item.ean || '—'))}</td>
        <td class="td-desc">${esc(item.descripcion)}</td>
        <td class="td-prov">${esc(String(item.gramaje || '—'))}</td>
        <td class="td-prov">${item.uxb !== '' ? item.uxb + ' u.' : '—'}</td>
        <td class="td-qty">${item.cantidad}</td>
        <td class="td-price td-super">${fmtSuper}</td>
        <td class="td-price td-sub td-super">${fmtSubSuper}</td>
        <td class="td-price td-may">${fmtMay}</td>
        <td class="td-price td-sub td-may">${fmtSubMay}</td>
        <td class="td-actions"><button class="btn-remove" onclick="removeItem(${item.id})">✕</button></td>
      </tr>`;
  }).join('');
  els.totalSuper.textContent = `$${fmt(sumSuper)}`;
  els.totalMay.textContent   = `$${fmt(sumMay)}`;
  if (sumSuper > 0 && sumMay > 0 && sumSuper > sumMay) {
    const ahorro = sumSuper - sumMay;
    const pct    = ((ahorro / sumSuper) * 100).toFixed(1);
    els.ahorroMonto.textContent     = `$${fmt(ahorro)}`;
    els.ahorroPct.textContent       = `${pct}%`;
    els.totalesAhorro.style.display = 'flex';
  } else {
    els.totalesAhorro.style.display = 'none';
  }
}

function updateBadge() {
  const n = state.order.length;
  els.badgeCount.textContent = `${n} ${n === 1 ? 'ítem' : 'ítems'}`;
  els.badgeCount.classList.toggle('show', n > 0);
}

// ══════════════════════════════════════════════
// ── Scanner
// ══════════════════════════════════════════════
function openScanner() {
  if (state.scannerActive) return;
  els.scanOverlay.classList.add('open');
  els.camErr.classList.remove('show');
  document.body.style.overflow = 'hidden';
  setTimeout(startScanner, 250);
}

async function startScanner() {
  try {
    state.html5QrCode = new Html5Qrcode('reader');
    await state.html5QrCode.start(
      { facingMode: 'environment' },
      {
        fps: 10, qrbox: { width: 260, height: 120 },
        formatsToSupport: [
          Html5QrcodeSupportedFormats.EAN_13, Html5QrcodeSupportedFormats.EAN_8,
          Html5QrcodeSupportedFormats.UPC_A,  Html5QrcodeSupportedFormats.UPC_E,
          Html5QrcodeSupportedFormats.CODE_128, Html5QrcodeSupportedFormats.CODE_39,
          Html5QrcodeSupportedFormats.ITF,
        ],
        aspectRatio: 4 / 3,
      },
      onBarcodeScanned, null
    );
    state.scannerActive = true;
  } catch (err) {
    console.warn('Scanner error:', err);
    if (els.sHint) els.sHint.style.display = 'none';
    els.camErr.classList.add('show');
  }
}

function onBarcodeScanned(decodedText) {
  if (navigator.vibrate) navigator.vibrate([80, 30, 80]);
  closeScanner();
  showToast('ok', '¡Código escaneado!');
  triggerLookup(decodedText);
}

async function closeScanner() {
  if (state.html5QrCode) {
    try { if (state.scannerActive) await state.html5QrCode.stop(); state.html5QrCode.clear(); } catch (_) {}
    state.html5QrCode = null;
  }
  state.scannerActive = false;
  els.scanOverlay.classList.remove('open');
  document.body.style.overflow = '';
}

// ══════════════════════════════════════════════
// ── Helpers
// ══════════════════════════════════════════════
function normalize(str) {
  if (!str) return '';
  return String(str).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function fmt(num) {
  return Number(num).toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function highlight(text, query) {
  if (!text || !query) return esc(text);
  const normText  = normalize(text);
  const normQuery = normalize(query);
  const result = []; let i = 0;
  while (i < text.length) {
    const idx = normText.indexOf(normQuery, i);
    if (idx === -1) { result.push(esc(text.slice(i))); break; }
    result.push(esc(text.slice(i, idx)));
    result.push(`<mark>${esc(text.slice(idx, idx + query.length))}</mark>`);
    i = idx + query.length;
  }
  return result.join('');
}

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

let toastTimer = null;
function showToast(type, msg, duration = 3000) {
  els.toastMsg.textContent = msg;
  els.toast.className      = `toast ${type}`;
  void els.toast.offsetWidth;
  els.toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => els.toast.classList.remove('show'), duration);
}

// ══════════════════════════════════════════════
// ── Export Excel — SIN PRECIOS NI TOTALES DE PRECIO
// ══════════════════════════════════════════════
async function exportExcel() {
  if (!state.order.length) return;
  if (!state.currentCliente) {
    els.eCliente.classList.add('show');
    showToast('err', 'Seleccioná un cliente antes de exportar.');
    return;
  }

  const now       = new Date();
  const fecha     = now.toLocaleDateString('es-AR');
  const fechaFile = now.toISOString().slice(0, 10);
  const remitoId  = now.getTime();
  const c         = state.currentCliente;
  const clienteNombre = c.fantasia || c.razons || c.codigo;
  const filename  = `PED_${clienteNombre.replace(/\s+/g,'_')}_${fechaFile}_${remitoId}.xlsx`;

  const workbook = new ExcelJS.Workbook();
  const ws       = workbook.addWorksheet('Pedido');

  ws.pageSetup = {
    paperSize: 9, orientation: 'portrait', fitToPage: true,
    fitToWidth: 1, fitToHeight: 0,
    margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.3, footer: 0.3 },
  };

  ws.columns = [
    { key: 'barcode', width: 20   },
    { key: 'sp1',     width: 0.01 },
    { key: 'ean',     width: 16   },
    { key: 'desc',    width: 48   },
    { key: 'sp2',     width: 0.01 },
    { key: 'gram',    width: 12   },
    { key: 'uxb',     width: 7    },
    { key: 'cant',    width: 9    },
  ];

  const thinBorder   = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
  const centerMiddle = { horizontal:'center', vertical:'middle', wrapText:true };

  function styleCell(cell, opts = {}) {
    cell.alignment = centerMiddle;
    if (!opts.noborder) cell.border = thinBorder;
    if (opts.bold)      cell.font   = { ...(cell.font||{}), bold: true };
  }

  function barcodeDataUrl(code) {
    const canvas = document.createElement('canvas');
    const f = /^\d{13}$/.test(code) ? 'ean13' : 'code128';
    try { JsBarcode(canvas, String(code), { format: f,         displayValue: false, margin: 4, width: 2, height: 56 }); }
    catch { JsBarcode(canvas, String(code), { format: 'code128', displayValue: false, margin: 4, width: 2, height: 56 }); }
    return canvas.toDataURL('image/png');
  }

  const lastCol = 'H';

  [
    [1, `PEDIDO — ${clienteNombre.toUpperCase()} — ${fecha}`, true,  18],
    [2, `Cód.: ${c.codigo}  |  CUIT: ${c.cuit||'—'}  |  IVA: ${c.iva||'—'}`, false, 11],
    [3, `Dirección: ${c.direccion||'—'}  |  Tel.: ${c.telefono||'—'}`, false, 11],
    [4, `ID: ${remitoId}`, false, 10],
  ].forEach(([row, val, bold, size]) => {
    ws.mergeCells(`A${row}:${lastCol}${row}`);
    const cell = ws.getCell(`A${row}`);
    cell.value = val; cell.font = { bold, size };
    styleCell(cell);
  });

  const hdrs = { A:'COD-BAR', B:'', C:'EAN', D:'DESCRIPCION', E:'', F:'GRAMAJE', G:'UxB', H:'CANT' };
  for (const [col, val] of Object.entries(hdrs)) {
    const cell = ws.getCell(`${col}5`);
    cell.value = val;
    cell.font  = { bold: true, size: 9 };
    styleCell(cell, { noborder: col === 'B' || col === 'E' });
    if (col !== 'B' && col !== 'E') cell.border = thinBorder;
  }
  ws.getRow(5).height = 18;

  for (let i = 0; i < state.order.length; i++) {
    const item     = state.order[i];
    const rowIndex = 6 + i;
    const eanCode  = String(item.ean || '');
    const ROW_H    = 52;
    ws.getRow(rowIndex).height = ROW_H;

    ws.getCell(`A${rowIndex}`).value  = '';
    ws.getCell(`A${rowIndex}`).border = thinBorder;
    if (eanCode) {
      const b64 = barcodeDataUrl(eanCode).split(',')[1];
      const img = workbook.addImage({ base64: b64, extension: 'png' });
      ws.addImage(img, { tl: { col: 0.1, row: rowIndex - 1 + 0.08 }, ext: { width: 130, height: ROW_H * 0.82 }, editAs: 'oneCell' });
    }
    ws.getCell(`B${rowIndex}`).value = '';

    const setVal = (col, val, opts = {}) => {
      const cell = ws.getCell(`${col}${rowIndex}`);
      cell.value = val; styleCell(cell, opts);
      if (opts.mono) cell.font = { ...(cell.font||{}), name: 'Courier New', size: 10 };
    };

    setVal('C', eanCode,           { mono: true });
    setVal('D', item.descripcion || '');
    ws.getCell(`E${rowIndex}`).value = '';
    setVal('F', item.gramaje || '');
    setVal('G', item.uxb !== '' ? item.uxb : '');
    setVal('H', item.cantidad, { bold: true });
  }

  // Total — solo cantidad, sin precios
  const totRow = 6 + state.order.length;
  ws.mergeCells(`A${totRow}:G${totRow}`);
  const tc = ws.getCell(`A${totRow}`);
  tc.value = 'TOTAL ÍTEMS'; tc.font = { bold: true, size: 11 }; styleCell(tc);
  const tcQ = ws.getCell(`H${totRow}`);
  tcQ.value = state.order.reduce((s, i) => s + i.cantidad, 0);
  tcQ.font  = { bold: true, size: 11 }; styleCell(tcQ);
  ws.getRow(totRow).height = 22;

  try {
    const buf  = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/octet-stream' });
    saveAs(blob, filename);
    showToast('ok', 'Excel descargado.');
  } catch (err) {
    console.error(err); showToast('err', 'No se pudo generar el Excel.');
  }
}

// ══════════════════════════════════════════════
// ── Export PDF — SIN PRECIOS NI TOTALES DE PRECIO
// ══════════════════════════════════════════════
async function exportPDF() {
  if (!state.order.length) return;
  if (!state.currentCliente) {
    els.eCliente.classList.add('show');
    showToast('err', 'Seleccioná un cliente antes de exportar.');
    return;
  }

  const now       = new Date();
  const fecha     = now.toLocaleDateString('es-AR');
  const fechaFile = now.toISOString().slice(0, 10);
  const remitoId  = now.getTime();
  const c         = state.currentCliente;
  const clienteNombre = c.fantasia || c.razons || c.codigo;
  const filename  = `PED_${clienteNombre.replace(/\s+/g,'_')}_${fechaFile}_${remitoId}.pdf`;

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'pt', format: 'a4', orientation: 'portrait' });

  const PAGE_W  = doc.internal.pageSize.getWidth();
  const MARGIN  = 36;
  const TABLE_W = PAGE_W - MARGIN * 2;

  doc.setFont('helvetica', 'bold');   doc.setFontSize(14);
  doc.text(`PEDIDO — ${clienteNombre.toUpperCase()} — ${fecha}`, MARGIN, 34);
  doc.setFont('helvetica', 'normal'); doc.setFontSize(9);
  doc.text(`Cód: ${c.codigo}  |  CUIT: ${c.cuit||'—'}  |  IVA: ${c.iva||'—'}  |  ${c.localidad||''}`, MARGIN, 50);
  doc.text(`ID: ${remitoId}`, MARGIN, 63);

  const W = {
    barcode: TABLE_W * 0.17,
    ean:     TABLE_W * 0.14,
    desc:    TABLE_W * 0.42,
    gram:    TABLE_W * 0.12,
    uxb:     TABLE_W * 0.07,
    cant:    TABLE_W * 0.08,
  };
  const ROW_H = 34;
  const rows  = [];

  for (const item of state.order) {
    const eanCode = String(item.ean || '');
    let barcodeImg = null;
    if (eanCode) {
      try {
        const canvas = document.createElement('canvas');
        const f = /^\d{13}$/.test(eanCode) ? 'ean13' : 'code128';
        try { JsBarcode(canvas, eanCode, { format: f,         displayValue: false, height: 38, margin: 2, width: 1.4 }); }
        catch { JsBarcode(canvas, eanCode, { format: 'code128', displayValue: false, height: 38, margin: 2, width: 1.4 }); }
        barcodeImg = canvas.toDataURL('image/png');
      } catch(e) {}
    }
    rows.push([
      { content: '', barcode: barcodeImg },
      eanCode,
      item.descripcion || '—',
      item.gramaje     || '—',
      item.uxb !== ''  ? String(item.uxb) : '—',
      item.cantidad,
    ]);
  }

  const totalCant = state.order.reduce((s, i) => s + i.cantidad, 0);
  rows.push([
    { content: 'TOTAL ÍTEMS', colSpan: 5, styles: { fontStyle:'bold', halign:'right' } },
    { content: String(totalCant), styles: { fontStyle:'bold' } },
  ]);

  doc.autoTable({
    startY: 76,
    margin: { left: MARGIN, right: MARGIN },
    head: [['Cód. Barras', 'EAN', 'Descripción', 'Gramaje', 'UxB', 'Cant.']],
    body: rows,
    styles: { halign:'center', valign:'middle', fontSize: 8, cellPadding: 2.5, minCellHeight: ROW_H, overflow:'ellipsize' },
    headStyles: { fillColor:[30,30,50], textColor:240, fontStyle:'bold', fontSize:9 },
    alternateRowStyles: { fillColor:[245,245,250] },
    columnStyles: {
      0: { cellWidth: W.barcode },
      1: { cellWidth: W.ean    },
      2: { cellWidth: W.desc,  halign:'left' },
      3: { cellWidth: W.gram   },
      4: { cellWidth: W.uxb    },
      5: { cellWidth: W.cant,  fontStyle:'bold' },
    },
    didDrawCell: (data) => {
      if (data.section !== 'body' || data.column.index !== 0) return;
      const img = data.cell.raw?.barcode;
      if (!img) return;
      const pad = 3;
      doc.addImage(img, 'PNG', data.cell.x + pad, data.cell.y + pad,
        data.cell.width - pad*2, data.cell.height - pad*2);
    },
  });

  const pdfBlob = doc.output('blob');
  doc.save(filename);
  showToast('ok', 'PDF generado.');
}

// ── Drive upload ──────────────────────────────
function blobToBase64(blob) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onloadend = () => res(reader.result.split(',')[1]);
    reader.onerror = rej;
    reader.readAsDataURL(blob);
  });
}

