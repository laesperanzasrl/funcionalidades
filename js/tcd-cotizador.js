// ══════════════════════════════════════════════════════════
// TCD COTIZADOR — JS
// ══════════════════════════════════════════════════════════

// ⚠️ Reemplazá esta URL con la URL de tu Web App de Apps Script
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
  searchField: 'all',
  searchTimer: null,
  inputMode: 'cam',
  extDebounce: null,
  searchCart: new Map(),
  clienteTimer: null,
};

// ── DOM ──────────────────────────────────────────────────
const $ = (id) => document.getElementById(id);

const els = {
  // cliente
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
  // producto
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
  // tabla
  orderTable:       $('orderTable'),
  orderBody:        $('orderBody'),
  tableEmpty:       $('tableEmpty'),
  // totales
  totalesPanel:     $('totalesPanel'),
  totalSuper:       $('totalSuper'),
  totalMay:         $('totalMay'),
  totalesAhorro:    $('totalesAhorro'),
  ahorroMonto:      $('ahorroMonto'),
  ahorroPct:        $('ahorroPct'),
  // footer
  btnExport:        $('btnExport'),
  btnClearOrder:    $('btnClearOrder'),
  badgeCount:       $('badgeCount'),
  toast:            $('toast'),
  toastMsg:         $('toastMsg'),
  toastDot:         $('toastDot'),
  // scanner
  scanOverlay:      $('scanOverlay'),
  btnOpenScanner:   $('btnOpenScanner'),
  btnCloseScanner:  $('btnCloseScanner'),
  btnManualEntry:   $('btnManualEntry'),
  camErr:           $('camErr'),
  sHint:            $('sHint'),
  // search modal
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
  // mode
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
  // Cargar en paralelo
  await Promise.all([cargarProductos(), cargarClientes()]);

  // Theme
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

  // PDF / Export
  els.btnExportPDF = $('btnExportPDF');
  els.btnExportPDF.addEventListener('click', exportPDF);
  els.btnExport.addEventListener('click', exportExcel);
  els.btnClearOrder.addEventListener('click', clearOrder);

  // Scanner / Search
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
    renderSearchResults('');
    els.searchInput.focus();
  });

  document.querySelectorAll('.sf-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      document.querySelectorAll('.sf-chip').forEach(c => c.classList.remove('sf-chip--active'));
      chip.classList.add('sf-chip--active');
      state.searchField = chip.dataset.field;
      onSearchInput();
    });
  });

  els.btnCommitSearch.addEventListener('click', commitSearchCart);

  // Mode switcher
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

  // External scanner
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

// ⚠️ Reemplazá con el file_id que te logueó generarYGuardar
const CLIENTES_JSON_ID  = '1u5LD7FEaACm6fa6lV1ERfukwFSOBcpu6';
const CLIENTES_JSON_URL = `https://drive.google.com/uc?export=download&id=${CLIENTES_JSON_ID}`;
const CLIENTES_CACHE_KEY = 'tcd_clientes_v1';
const CLIENTES_CACHE_HRS = 6;

async function cargarClientes() {
  // 1. localStorage — instantáneo si el caché es fresco
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

  // 2. Fetch directo al JSON en Drive
  try {
    const res  = await fetch(CLIENTES_JSON_URL);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (!json.ok) throw new Error('JSON inválido');

    state.clientes = (json.data || []).filter(c => !c.inactivo);

    localStorage.setItem(CLIENTES_CACHE_KEY, JSON.stringify({
      data: json.data || [],
      ts:   Date.now()
    }));

    initClienteSelector(`${state.clientes.length} clientes`);

  } catch (err) {
    console.error('Error al cargar clientes:', err);
    state.clientes = [];
    initClienteSelector(null, err.message);
  }
}

// Forzar actualización manual (si querés un botón de "Actualizar")
async function refrescarClientes() {
  localStorage.removeItem(CLIENTES_CACHE_KEY);
  await cargarClientes();
}

function initClienteSelector(successMsg, errorMsg) {
  els.clienteLoading.style.display  = 'none';
  els.clienteSearchWrap.style.display = 'block';

  // Mostrar estado de carga
  const statusEl = document.getElementById('clienteStatus');
  if (statusEl) {
    if (errorMsg) {
      statusEl.textContent  = `⚠ No se pudieron cargar los clientes: ${errorMsg}`;
      statusEl.className    = 'cliente-status cliente-status--err';
      statusEl.style.display = 'block';
    } else if (successMsg) {
      statusEl.textContent  = `✓ ${successMsg}`;
      statusEl.className    = 'cliente-status cliente-status--ok';
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
    if (els.clienteInput.value.length >= 1) {
      renderClienteResults(els.clienteInput.value);
    }
  });

  els.btnClearCliente.addEventListener('click', () => {
    els.clienteInput.value = '';
    els.btnClearCliente.style.display = 'none';
    els.clienteResults.style.display  = 'none';
    els.clienteInput.focus();
  });

  els.btnChangeCliente.addEventListener('click', () => {
    state.currentCliente = null;
    els.clienteCard.style.display        = 'none';
    els.clienteSearchWrap.style.display  = 'block';
    els.clienteInput.value               = '';
    els.btnClearCliente.style.display    = 'none';
    els.clienteResults.style.display     = 'none';
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

  // Con 0 caracteres, ocultar
  if (q.length === 0) {
    els.clienteResults.style.display = 'none';
    return;
  }

  // Sin clientes cargados
  if (!state.clientes.length) {
    els.clienteResults.innerHTML = `<li class="cr-empty">⚠ No hay clientes cargados. Verificá la URL del script.</li>`;
    els.clienteResults.style.display = 'block';
    return;
  }

  // Búsqueda SOLO por fantasia y razons (como pediste)
  const matches = state.clientes.filter(c =>
    normalize(c.fantasia).includes(q) ||
    normalize(c.razons).includes(q)
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
            ${c.codigo   ? `<span class="cr-chip">${esc(c.codigo)}</span>`           : ''}
            ${c.localidad? `<span class="cr-chip cr-chip--loc">${esc(c.localidad)}</span>` : ''}
            ${c.zona     ? `<span class="cr-chip cr-chip--zona">${esc(c.zona)}</span>`    : ''}
            ${c.iva      ? `<span class="cr-chip cr-chip--iva">${esc(c.iva)}</span>`      : ''}
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
    const normEAN  = normalizeCode(p.EAN);
    const normINT  = normalizeCode(p.INTERNO);
    return normEAN === normCode || normINT === normCode;
  }) || null;
}

function cartKey(p) {
  return `${p.EAN ?? ''}_${p.INTERNO ?? ''}_${normalize(p.DESCRIPCION ?? '')}`;
}

// ── External scanner ─────────────────────────
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
// ── Search Modal
// ══════════════════════════════════════════════
function openSearchModal() {
  state.searchCart.clear();
  els.searchOverlay.classList.add('open');
  document.body.style.overflow = 'hidden';
  els.searchInput.value = '';
  els.btnClearSearch.style.display = 'none';
  els.searchCount.textContent = '';
  renderSearchResults('');
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
  clearTimeout(state.searchTimer);
  state.searchTimer = setTimeout(() => renderSearchResults(query), 150);
}

function renderSearchResults(query) {
  const q     = normalize(query);
  const field = state.searchField;

  if (q.length < 2) {
    els.searchStateEmpty.style.display = 'flex';
    els.searchStateNone.style.display  = 'none';
    els.searchResults.style.display    = 'none';
    els.searchCount.textContent        = '';
    return;
  }

  const matches = state.productos.filter(p => {
    if (field === 'all') {
      return normalize(p.DESCRIPCION).includes(q) ||
             normalize(p.PROVEEDOR).includes(q)   ||
             normalize(p.SECTOR).includes(q)       ||
             normalize(p.SECCION).includes(q)      ||
             normalize(String(p.EAN ?? '')).includes(q) ||
             normalize(String(p.INTERNO ?? '')).includes(q);
    }
    if (field === 'EAN') {
      return normalize(String(p.EAN ?? '')).includes(q) ||
             normalize(String(p.INTERNO ?? '')).includes(q);
    }
    return normalize(String(p[field] ?? '')).includes(q);
  });

  els.searchStateEmpty.style.display = 'none';

  if (!matches.length) {
    els.searchStateNone.style.display  = 'flex';
    els.searchTermDisplay.textContent  = `"${query}"`;
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

  els.searchResults.innerHTML = shown.map((p, i) => {
    const key     = cartKey(p);
    const inCart  = state.searchCart.has(key);
    const qtyVal  = inCart ? (state.searchCart.get(key).qty || '') : '';

    const desc    = highlight(p.DESCRIPCION || '—', q);
    const prov    = highlight(p.PROVEEDOR   || '—', q);
    const eanRaw  = String(p.EAN    ?? '');
    const intRaw  = String(p.INTERNO ?? '');
    const eanHL   = highlight(eanRaw, q);
    const intHL   = highlight(intRaw, q);

    const eanBadge  = eanRaw ? `<span class="sr-ean">EAN ${eanHL}</span>` : '';
    const intBadge  = intRaw && intRaw !== eanRaw
      ? `<span class="sr-ean sr-ean--int">INT ${intHL}</span>` : '';
    const sect    = p.SECTOR  ? `<span class="sr-tag">${esc(p.SECTOR)}</span>`  : '';
    const secc    = p.SECCION ? `<span class="sr-tag">${esc(p.SECCION)}</span>` : '';
    const gramaje = p.GRAMAJE ? `<span class="sr-tag sr-tag--gramaje">${esc(p.GRAMAJE)}</span>` : '';
    const pvp     = p['PVP SUPER']      != null
      ? `<span class="sr-tag sr-tag--pvp" title="Precio normal">Super $${fmt(p['PVP SUPER'])}</span>` : '';
    const pvpMay  = p['PVP MAYORISTA']  != null
      ? `<span class="sr-tag sr-tag--pvp-may" title="Precio mayorista">May. $${fmt(p['PVP MAYORISTA'])}</span>` : '';

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
    els.scbLabel.textContent = '0 productos seleccionados';
    els.btnCommitSearch.disabled   = true;
    els.btnCommitSearch.innerHTML  = btnCommitHTML('Agregar a la lista');
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
    if (existing) {
      existing.cantidad += qty;
      updated++;
    } else {
      state.order.push(buildOrderItem(p, qty));
      added++;
    }
  });

  renderTable();
  updateBadge();
  const parts = [];
  if (added)   parts.push(`${added} nuevo${added !== 1 ? 's' : ''}`);
  if (updated) parts.push(`${updated} actualizado${updated !== 1 ? 's' : ''}`);
  showToast('ok', parts.join(' · ') + ' en la lista');
  if (navigator.vibrate) navigator.vibrate([60, 30, 60]);

  state.searchCart.clear();
  closeSearchModal();
}

// ── Construye un item del pedido con precios ──────────────
function buildOrderItem(p, cantidad) {
  return {
    id:          Date.now() + Math.random(),
    ean:         p.EAN       || p.INTERNO || '',
    interno:     p.INTERNO   || '',
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
  if (isNaN(qty) || qty <= 0) {
    els.eQty.classList.add('show');
    els.fQty.focus();
    return;
  }

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
  const ean    = String(p?.EAN    || p?.INTERNO || '').trim();
  const interno = String(p?.INTERNO || '').trim();
  return state.order.find(item => {
    if (ean    && String(item.ean).trim()    === ean)    return true;
    if (interno && String(item.interno).trim() === interno) return true;
    return false;
  }) || null;
}

// ── Render tabla + totales ────────────────────
function renderTable() {
  const has = state.order.length > 0;
  els.tableEmpty.style.display  = has ? 'none' : 'block';
  els.orderTable.style.display  = has ? 'table' : 'none';
  els.totalesPanel.style.display = has ? 'block' : 'none';
  els.btnExport.disabled        = !has;
  els.btnExportPDF.disabled     = !has;
  els.btnClearOrder.disabled    = !has;

  let sumSuper = 0;
  let sumMay   = 0;

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

  // Totales
  els.totalSuper.textContent = `$${fmt(sumSuper)}`;
  els.totalMay.textContent   = `$${fmt(sumMay)}`;

  // Ahorro
  if (sumSuper > 0 && sumMay > 0 && sumSuper > sumMay) {
    const ahorro = sumSuper - sumMay;
    const pct    = ((ahorro / sumSuper) * 100).toFixed(1);
    els.ahorroMonto.textContent = `$${fmt(ahorro)}`;
    els.ahorroPct.textContent   = `${pct}%`;
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
// ── Export Excel
// ══════════════════════════════════════════════
async function exportExcel() {
  if (!state.order.length) return;

  const now        = new Date();
  const fecha      = now.toLocaleDateString('es-AR');
  const fechaFile  = now.toISOString().slice(0, 10);
  const remitoId   = now.getTime();

  if (!state.currentCliente) {
    els.eCliente.classList.add('show');
    showToast('err', 'Seleccioná un cliente antes de exportar.');
    return;
  }

  const c         = state.currentCliente;
  const clienteNombre = c.fantasia || c.razons || c.codigo;
  const filename  = `COT_${clienteNombre.replace(/\s+/g,'_')}_${fechaFile}_${remitoId}.xlsx`;

  const workbook  = new ExcelJS.Workbook();
  const ws        = workbook.addWorksheet('Cotización');

  ws.pageSetup = {
    paperSize: 9, orientation: 'portrait', fitToPage: true,
    fitToWidth: 1, fitToHeight: 0,
    margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.3, footer: 0.3 },
  };

  ws.columns = [
    { key: 'barcode', width: 18   },
    { key: 'sp1',     width: 0.01 },
    { key: 'ean',     width: 16   },
    { key: 'desc',    width: 42   },
    { key: 'sp2',     width: 0.01 },
    { key: 'gram',    width: 12   },
    { key: 'uxb',     width: 6    },
    { key: 'cant',    width: 7    },
    { key: 'pvpS',    width: 12   },
    { key: 'subS',    width: 14   },
    { key: 'pvpM',    width: 12   },
    { key: 'subM',    width: 14   },
  ];

  const thinBorder    = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
  const centerMiddle  = { horizontal:'center', vertical:'middle', wrapText:true };

  function styleCell(cell, opts = {}) {
    cell.alignment = centerMiddle;
    if (!opts.noborder) cell.border = thinBorder;
    if (opts.bold)      cell.font   = { ...(cell.font||{}), bold: true };
  }

  function barcodeDataUrl(code) {
    const canvas = document.createElement('canvas');
    const fmt    = /^\d{13}$/.test(code) ? 'ean13' : 'code128';
    try { JsBarcode(canvas, String(code), { format: fmt,       displayValue: false, margin: 4, width: 2, height: 56 }); }
    catch { JsBarcode(canvas, String(code), { format: 'code128', displayValue: false, margin: 4, width: 2, height: 56 }); }
    return canvas.toDataURL('image/png');
  }

  const lastCol = 'L';

  // Cabeceras
  [
    [1, `COTIZACIÓN — ${clienteNombre.toUpperCase()} — ${fecha}`, true,  18],
    [2, `Cód.: ${c.codigo}  |  CUIT: ${c.cuit||'—'}  |  IVA: ${c.iva||'—'}`, false, 11],
    [3, `Dirección: ${c.direccion||'—'}  |  Tel.: ${c.telefono||'—'}`, false, 11],
    [4, `ID: ${remitoId}`, false, 10],
  ].forEach(([row, val, bold, size]) => {
    ws.mergeCells(`A${row}:${lastCol}${row}`);
    const cell = ws.getCell(`A${row}`);
    cell.value = val; cell.font = { bold, size };
    styleCell(cell);
  });

  // Fila títulos
  const hdrs = { A:'COD-BAR', B:'', C:'EAN', D:'DESCRIPCION', E:'', F:'GRAMAJE',
                 G:'UxB', H:'CANT', I:'P.SUPER', J:'SUB.SUPER', K:'P.MAY', L:'SUB.MAY' };
  for (const [col, val] of Object.entries(hdrs)) {
    const cell = ws.getCell(`${col}5`);
    cell.value = val;
    cell.font  = { bold: true, size: 9 };
    styleCell(cell, { noborder: col === 'B' || col === 'E' });
    if (col !== 'B' && col !== 'E') cell.border = thinBorder;
  }
  ws.getRow(5).height = 18;

  // Datos
  let sumSuper = 0, sumMay = 0;
  for (let i = 0; i < state.order.length; i++) {
    const item     = state.order[i];
    const rowIndex = 6 + i;
    const eanCode  = String(item.ean || '');
    const ROW_H    = 52;
    ws.getRow(rowIndex).height = ROW_H;

    const subS = item.pvpSuper != null ? item.pvpSuper * item.cantidad : null;
    const subM = item.pvpMay   != null ? item.pvpMay   * item.cantidad : null;
    if (subS != null) sumSuper += subS;
    if (subM != null) sumMay   += subM;

    // A — barcode
    ws.getCell(`A${rowIndex}`).value = ''; ws.getCell(`A${rowIndex}`).border = thinBorder;
    if (eanCode) {
      const b64 = barcodeDataUrl(eanCode).split(',')[1];
      const img = workbook.addImage({ base64: b64, extension: 'png' });
      ws.addImage(img, { tl: { col: 0.1, row: rowIndex - 1 + 0.08 }, ext: { width: 118, height: ROW_H * 0.82 }, editAs: 'oneCell' });
    }
    ws.getCell(`B${rowIndex}`).value = '';

    const setVal = (col, val, opts = {}) => {
      const cell = ws.getCell(`${col}${rowIndex}`);
      cell.value = val; styleCell(cell, opts);
      if (opts.mono) cell.font = { ...(cell.font||{}), name: 'Courier New', size: 10 };
    };

    setVal('C', eanCode,       { mono: true });
    setVal('D', item.descripcion || '');
    ws.getCell(`E${rowIndex}`).value = '';
    setVal('F', item.gramaje   || '');
    setVal('G', item.uxb !== '' ? item.uxb : '');
    setVal('H', item.cantidad,   { bold: true });
    setVal('I', item.pvpSuper != null ? item.pvpSuper : '');
    setVal('J', subS          != null ? subS           : '', { bold: true });
    setVal('K', item.pvpMay   != null ? item.pvpMay    : '');
    setVal('L', subM          != null ? subM            : '', { bold: true });

    // Formato moneda
    ['I','J','K','L'].forEach(col => {
      ws.getCell(`${col}${rowIndex}`).numFmt = '"$"#,##0.00';
    });
  }

  // Fila totales
  const totRow = 6 + state.order.length;
  ws.mergeCells(`A${totRow}:H${totRow}`);
  const tc = ws.getCell(`A${totRow}`);
  tc.value = 'TOTALES'; tc.font = { bold: true, size: 11 }; styleCell(tc);

  const setTot = (col, val) => {
    const cell = ws.getCell(`${col}${totRow}`);
    cell.value = val; cell.numFmt = '"$"#,##0.00';
    cell.font  = { bold: true, size: 11 }; styleCell(cell);
  };
  setTot('J', sumSuper); setTot('L', sumMay);
  ws.getRow(totRow).height = 22;

  try {
    const buf  = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/octet-stream' });
    saveAs(blob, filename);
    showToast('ok', 'Excel descargado.');
    await uploadToDrive(blob, filename);
  } catch (err) {
    console.error(err); showToast('err', 'No se pudo generar el Excel.');
  }
}

// ── Export PDF ────────────────────────────────
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
  const filename  = `COT_${clienteNombre.replace(/\s+/g,'_')}_${fechaFile}_${remitoId}.pdf`;

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'pt', format: 'a4', orientation: 'landscape' });

  const PAGE_W = doc.internal.pageSize.getWidth();
  const MARGIN = 36;
  const TABLE_W = PAGE_W - MARGIN * 2;

  // Header
  doc.setFont('helvetica', 'bold');   doc.setFontSize(14);
  doc.text(`COTIZACIÓN — ${clienteNombre.toUpperCase()} — ${fecha}`, MARGIN, 34);
  doc.setFont('helvetica', 'normal'); doc.setFontSize(9);
  doc.text(`Cód: ${c.codigo}  |  CUIT: ${c.cuit||'—'}  |  IVA: ${c.iva||'—'}  |  ${c.localidad||''}`, MARGIN, 50);
  doc.text(`ID: ${remitoId}`, MARGIN, 63);

  const W = {
    barcode: TABLE_W * 0.12, ean:  TABLE_W * 0.09,
    desc:    TABLE_W * 0.24, prov: TABLE_W * 0.12,
    interno: TABLE_W * 0.07, gram: TABLE_W * 0.07,
    uxb:     TABLE_W * 0.06, cant: TABLE_W * 0.06,
    pvpS:    TABLE_W * 0.07, subS: TABLE_W * 0.08,
    pvpM:    TABLE_W * 0.07, subM: TABLE_W * 0.09,
  };
  const ROW_H = 34;
  let sumSuper = 0, sumMay = 0;
  const rows = [];

  for (const item of state.order) {
    const eanCode  = String(item.ean || '');
    const subS     = item.pvpSuper != null ? item.pvpSuper * item.cantidad : null;
    const subM     = item.pvpMay   != null ? item.pvpMay   * item.cantidad : null;
    if (subS != null) sumSuper += subS;
    if (subM != null) sumMay   += subM;

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
      item.descripcion  || '—',
      item.proveedor    || '—',
      item.interno      || '—',
      item.gramaje      || '—',
      item.uxb !== ''   ? String(item.uxb) : '—',
      item.cantidad,
      item.pvpSuper != null ? `$${fmt(item.pvpSuper)}` : '—',
      subS != null          ? `$${fmt(subS)}`          : '—',
      item.pvpMay != null   ? `$${fmt(item.pvpMay)}`   : '—',
      subM != null          ? `$${fmt(subM)}`          : '—',
    ]);
  }

  // Fila totales
  rows.push([
    { content: 'TOTALES', colSpan: 9, styles: { fontStyle:'bold', halign:'right' } },
    { content: `$${fmt(sumSuper)}`, styles: { fontStyle:'bold', textColor:[34,197,94] } },
    { content: '' },
    { content: `$${fmt(sumMay)}`, styles: { fontStyle:'bold', textColor:[96,165,250] } },
  ]);

  doc.autoTable({
    startY: 76,
    margin: { left: MARGIN, right: MARGIN },
    head: [['Cód. Barras','EAN','Descripción','Proveedor','Interno','Gramaje','UxB','Cant.','P.Super','Sub.Super','P.May.','Sub.May.']],
    body: rows,
    styles: { halign:'center', valign:'middle', fontSize:7, cellPadding:2.5, minCellHeight: ROW_H, overflow:'ellipsize' },
    headStyles: { fillColor:[30,30,50], textColor:240, fontStyle:'bold', fontSize:8 },
    alternateRowStyles: { fillColor:[245,245,250] },
    columnStyles: {
      0: { cellWidth: W.barcode }, 1: { cellWidth: W.ean     },
      2: { cellWidth: W.desc, halign:'left' }, 3: { cellWidth: W.prov, halign:'left' },
      4: { cellWidth: W.interno }, 5: { cellWidth: W.gram    },
      6: { cellWidth: W.uxb    }, 7: { cellWidth: W.cant     },
      8: { cellWidth: W.pvpS   }, 9: { cellWidth: W.subS, fontStyle:'bold' },
      10:{ cellWidth: W.pvpM   }, 11:{ cellWidth: W.subM, fontStyle:'bold' },
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
  await uploadToDrive(pdfBlob, filename);
  showToast('ok', 'PDF generado y subido a Drive.');
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

async function uploadToDrive(blob, filename) {
  const url = "https://script.google.com/macros/s/AKfycbzIRnuzUWTjG38fxttQbkvJ7Br_wYSYs5UeaJO9EHnncy7jr8vQivcfQLot0xqSzwsq/exec";
  try {
    const base64Data = await blobToBase64(blob);
    const res  = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ filename, mimeType: blob.type, file: base64Data }),
    });
    const json = await res.json();
    if (json.ok) showToast('ok', `Archivo subido a Drive: ${filename}`);
    else showToast('err', 'Error al subir a Drive: ' + (json.error || 'Desconocido'));
  } catch (err) {
    console.error(err); showToast('err', 'Error al conectar con Google Drive.');
  }
}