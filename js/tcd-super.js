// ── Estado ──
const state = {
  order: [],
  currentProduct: null,
  lookupTimer: null,
  scannerActive: false,
  html5QrCode: null,
  productos: [],
  searchField: 'all',
  searchTimer: null,
  inputMode: 'cam',    // 'cam' | 'ext'
  extDebounce: null,     // timer for external scanner debounce
  searchCart: new Map(),
};

// ── DOM ──
const $ = (id) => document.getElementById(id);

const els = {
  fConcepto: $('fConcepto'),
  eConcepto: $('eConcepto'),
  fSucursal: $('fSucursal'),
  eSucursal: $('eSucursal'),
  productCard: $('productCard'),
  pcStateLabel: $('pcStateLabel'),
  pcDescripcion: $('pcDescripcion'),
  pcProveedor: $('pcProveedor'),
  pcGramaje: $('pcGramaje'),
  pcUxb: $('pcUxb'),
  btnClearProduct: $('btnClearProduct'),
  qtyRow: $('qtyRow'),
  fQty: $('fQty'),
  eQty: $('eQty'),
  btnAdd: $('btnAdd'),
  orderTable: $('orderTable'),
  orderBody: $('orderBody'),
  tableEmpty: $('tableEmpty'),
  btnExport: $('btnExport'),
  btnClearOrder: $('btnClearOrder'),
  badgeCount: $('badgeCount'),
  toast: $('toast'),
  toastMsg: $('toastMsg'),
  toastDot: $('toastDot'),
  // scanner
  scanOverlay: $('scanOverlay'),
  btnOpenScanner: $('btnOpenScanner'),
  btnCloseScanner: $('btnCloseScanner'),
  btnManualEntry: $('btnManualEntry'),
  camErr: $('camErr'),
  sHint: $('sHint'),
  // search modal
  searchOverlay: $('searchOverlay'),
  btnOpenSearch: $('btnOpenSearch'),
  btnCloseSearch: $('btnCloseSearch'),
  searchInput: $('searchInput'),
  btnClearSearch: $('btnClearSearch'),
  searchResults: $('searchResults'),
  searchStateEmpty: $('searchStateEmpty'),
  searchStateNone: $('searchStateNone'),
  searchTermDisplay: $('searchTermDisplay'),
  searchCount: $('searchCount'),
  // carrito del modal
  searchCommitBar: $('searchCommitBar'),
  scbLabel: $('scbLabel'),
  btnCommitSearch: $('btnCommitSearch'),
  // mode switcher
  modeTabCam: $('modeTabCam'),
  modeTabExt: $('modeTabExt'),
  panelCam: $('panelCam'),
  panelExt: $('panelExt'),
  extInput: $('extInput'),
  btnExtClear: $('btnExtClear'),
  extStatus: $('extStatus'),
  extStatusDot: $('extStatusDot'),
  extStatusText: $('extStatusText'),
};

// ── Init ──
async function init() {
  await cargarProductos();

  // ── Theme toggle ──
  const btnTheme = $('btnThemeToggle');
  const applyTheme = (light) => {
    document.documentElement.classList.toggle('light', light);
    btnTheme.textContent = light ? '🌙' : '☀️';
    btnTheme.title = light ? 'Cambiar a modo oscuro' : 'Cambiar a modo claro';
    localStorage.setItem('theme', light ? 'light' : 'dark');
  };
  const savedTheme = localStorage.getItem('theme');
  applyTheme(savedTheme === 'light');
  btnTheme.addEventListener('click', () => {
    applyTheme(!document.documentElement.classList.contains('light'));
  });

  els.btnExportPDF = $('btnExportPDF');
  els.btnExportPDF.addEventListener('click', exportPDF);

  els.fConcepto.addEventListener('input', () => els.eConcepto.classList.remove('show'));
  els.fSucursal.addEventListener('change', () => {
    els.eSucursal.classList.remove('show');
    els.fSucursal.classList.remove('err');
  });

  // Scanner: tarjeta + cantidad
  els.btnClearProduct.addEventListener('click', resetCurrentProduct);
  els.fQty.addEventListener('input', () => els.eQty.classList.remove('show'));
  els.fQty.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); addToOrder(); }
  });
  els.btnAdd.addEventListener('click', addToOrder);

  // Modales
  els.btnOpenScanner.addEventListener('click', openScanner);
  els.btnCloseScanner.addEventListener('click', closeScanner);
  els.btnManualEntry.addEventListener('click', closeScanner);

  els.btnOpenSearch.addEventListener('click', openSearchModal);
  els.btnCloseSearch.addEventListener('click', closeSearchModal);
  els.searchOverlay.addEventListener('click', (e) => {
    if (e.target === els.searchOverlay) closeSearchModal();
  });

  // Búsqueda
  els.searchInput.addEventListener('input', onSearchInput);
  els.btnClearSearch.addEventListener('click', () => {
    els.searchInput.value = '';
    els.btnClearSearch.style.display = 'none';
    renderSearchResults('');
    els.searchInput.focus();
  });

  // Chips de filtro
  document.querySelectorAll('.sf-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      document.querySelectorAll('.sf-chip').forEach(c => c.classList.remove('sf-chip--active'));
      chip.classList.add('sf-chip--active');
      state.searchField = chip.dataset.field;
      onSearchInput();
    });
  });

  // Botón confirmar carrito
  els.btnCommitSearch.addEventListener('click', commitSearchCart);

  // ── Mode switcher (Cámara / Lector Externo) ──
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

  const savedMode = localStorage.getItem('inputMode') || 'cam';
  setInputMode(savedMode);
  els.modeTabCam.addEventListener('click', () => setInputMode('cam'));
  els.modeTabExt.addEventListener('click', () => setInputMode('ext'));

  // ── External scanner input logic ──
  els.extInput.addEventListener('input', () => {
    const val = els.extInput.value.trim();
    els.btnExtClear.style.display = val ? 'flex' : 'none';

    if (!val) { setExtStatus('ready', 'Listo — esperando lectura'); return; }

    // Indicate "reading" while chars come in
    setExtStatus('reading', 'Leyendo código...');
    els.extInput.classList.add('ext-input--reading');
    els.extInput.classList.remove('ext-input--found');

    // Debounce: scanners dump chars very fast then stop;
    // wait 120ms of silence before triggering lookup
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

  // Escape + auto-focus ext input when a scanner key arrives
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if (els.searchOverlay.classList.contains('open')) closeSearchModal();
      if (els.scanOverlay.classList.contains('open')) closeScanner();
    }
    // If ext mode is active and user isn't focused on qty or ext input,
    // redirect any printable key to the ext input automatically
    if (state.inputMode === 'ext'
      && !els.searchOverlay.classList.contains('open')
      && !els.scanOverlay.classList.contains('open')
      && document.activeElement !== els.extInput
      && document.activeElement !== els.fQty
      && e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
      els.extInput.focus();
    }
  });

  els.btnExport.addEventListener('click', exportExcel);
  els.btnClearOrder.addEventListener('click', clearOrder);
}

document.addEventListener('DOMContentLoaded', init);

// ══════════════════════════════════════════════
// ── Productos
// ══════════════════════════════════════════════

async function cargarProductos() {
  try {
    const res = await fetch('../data/productos.json');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    state.productos = json.data || [];
  } catch (err) {
    console.error('Error al cargar productos.json:', err);
    showToast('err', 'No se pudo cargar la base de productos.');
  }
}

// Strips leading zeros and returns a trimmed string for safe comparison
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

// Clave única por producto para el carrito temporal
function cartKey(p) {
  return `${p.EAN ?? ''}_${p.INTERNO ?? ''}_${normalize(p.DESCRIPCION ?? '')}`;
}

// ── External scanner: lookup ──
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
  const el = els.extStatus;
  el.className = 'ext-status';
  if (type) el.classList.add(`ext-status--${type}`);
  els.extStatusText.textContent = text;
}

// ── Scanner: lookup ──
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
  els.pcProveedor.textContent = data.PROVEEDOR || '—';
  els.pcGramaje.textContent = data.GRAMAJE || '—';
  els.pcUxb.textContent = data.UXB != null ? `${data.UXB} u.` : '—';
  const pvp = data['PVP SUPER'];
  const pvpEl = $('pcPvp');
  if (pvpEl) pvpEl.textContent = pvp != null
    ? `$${Number(pvp).toLocaleString('es-AR', { minimumFractionDigits: 2 })}`
    : '—';

  const pvpMay = data['PVP MAYORISTA'];
  const pvpMayEl = $('pcPvpMay');
  if (pvpMayEl) pvpMayEl.textContent = pvpMay != null
    ? `$${Number(pvpMay).toLocaleString('es-AR', { minimumFractionDigits: 2 })}`
    : '—';
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
  els.pcStateLabel.textContent = 'Buscando...';
  els.btnClearProduct.style.display = 'none';
}
function showCardFound() {
  els.productCard.classList.add('visible');
  els.productCard.dataset.state = 'found';
  els.pcStateLabel.textContent = 'Producto encontrado';
  els.btnClearProduct.style.display = 'flex';
}
function showCardNotFound() {
  els.productCard.classList.add('visible');
  els.productCard.dataset.state = 'notfound';
  els.pcStateLabel.textContent = 'No encontrado';
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
  const q = normalize(query);
  const field = state.searchField;

  if (q.length < 2) {
    els.searchStateEmpty.style.display = 'flex';
    els.searchStateNone.style.display = 'none';
    els.searchResults.style.display = 'none';
    els.searchCount.textContent = '';
    return;
  }

  const matches = state.productos.filter(p => {
    if (field === 'all') {
      return (
        normalize(p.DESCRIPCION).includes(q) ||
        normalize(p.PROVEEDOR).includes(q) ||
        normalize(p.SECTOR).includes(q) ||
        normalize(p.SECCION).includes(q) ||
        normalize(String(p.EAN ?? '')).includes(q) ||
        normalize(String(p.INTERNO ?? '')).includes(q)
      );
    }
    if (field === 'EAN') {
      return (
        normalize(String(p.EAN ?? '')).includes(q) ||
        normalize(String(p.INTERNO ?? '')).includes(q)
      );
    }
    return normalize(String(p[field] ?? '')).includes(q);
  });

  els.searchStateEmpty.style.display = 'none';

  if (matches.length === 0) {
    els.searchStateNone.style.display = 'flex';
    els.searchTermDisplay.textContent = `"${query}"`;
    els.searchResults.style.display = 'none';
    els.searchCount.textContent = 'Sin resultados';
    return;
  }

  els.searchStateNone.style.display = 'none';
  els.searchResults.style.display = 'block';

  const shown = matches.slice(0, 80);
  const total = matches.length;
  els.searchCount.textContent = total > 80
    ? `${total} resultados — mostrando los primeros 80`
    : `${total} resultado${total !== 1 ? 's' : ''}`;

  els.searchResults.innerHTML = shown.map((p, i) => {
    const key = cartKey(p);
    const inCart = state.searchCart.has(key);
    const qtyVal = inCart ? (state.searchCart.get(key).qty || '') : '';

    const desc = highlight(p.DESCRIPCION || '—', q);
    const prov = highlight(p.PROVEEDOR || '—', q);
    const eanRaw = String(p.EAN ?? '');
    const intRaw = String(p.INTERNO ?? '');
    const eanHL = highlight(eanRaw, q);
    const intHL = highlight(intRaw, q);

    const eanBadge = eanRaw ? `<span class="sr-ean">EAN ${eanHL}</span>` : '';
    const intBadge = intRaw && intRaw !== eanRaw ? `<span class="sr-ean sr-ean--int">INT ${intHL}</span>` : '';
    const sect = p.SECTOR ? `<span class="sr-tag">${esc(p.SECTOR)}</span>` : '';
    const secc = p.SECCION ? `<span class="sr-tag">${esc(p.SECCION)}</span>` : '';
    const gramaje = p.GRAMAJE ? `<span class="sr-tag sr-tag--gramaje">${esc(p.GRAMAJE)}</span>` : '';
    const pvp = p['PVP SUPER'] != null
      ? `<span class="sr-tag sr-tag--pvp" title="Precio normal">Normal $${Number(p['PVP SUPER']).toLocaleString('es-AR', { minimumFractionDigits: 2 })}</span>`
      : '';
    const pvpMay = p['PVP MAYORISTA'] != null
      ? `<span class="sr-tag sr-tag--pvp-may" title="Precio mayorista">Mayorista $${Number(p['PVP MAYORISTA']).toLocaleString('es-AR', { minimumFractionDigits: 2 })}</span>`
      : '';

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
            <input
              type="number"
              class="sr-qty-input"
              data-key="${esc(key)}"
              value="${qtyVal}"
              placeholder="Cant."
              min="1" step="1"
              inputmode="numeric"
              aria-label="Cantidad"
            />
          </div>
        </div>

      </li>`;
  }).join('');

  // Eventos
  els.searchResults.querySelectorAll('.sr-item').forEach((li, i) => {
    const p = shown[i];
    const key = cartKey(p);

    // Click en la fila → toggle (excepto sobre el input)
    li.addEventListener('click', (e) => {
      if (e.target.closest('.sr-qty-input')) return;
      toggleCartItem(p, key, li);
    });

    // Teclado
    li.addEventListener('keydown', (e) => {
      if ((e.key === 'Enter' || e.key === ' ') && !e.target.closest('.sr-qty-input')) {
        e.preventDefault();
        toggleCartItem(p, key, li);
      }
      if (e.key === 'ArrowDown') { e.preventDefault(); (li.nextElementSibling || li).focus(); }
      if (e.key === 'ArrowUp') { e.preventDefault(); (li.previousElementSibling || li).focus(); }
    });

    // Input de cantidad
    const qtyInput = li.querySelector('.sr-qty-input');
    if (qtyInput) {
      qtyInput.addEventListener('click', (e) => e.stopPropagation());
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

// Agrega o quita un producto del carrito temporal
function toggleCartItem(producto, key, liEl) {
  if (state.searchCart.has(key)) {
    state.searchCart.delete(key);
    liEl.classList.remove('sr-item--selected');
    const wrap = liEl.querySelector('.sr-qty-wrap');
    if (wrap) wrap.classList.remove('sr-qty-wrap--visible');
    const inp = liEl.querySelector('.sr-qty-input');
    if (inp) inp.value = '';
  } else {
    state.searchCart.set(key, { producto, qty: 0 });
    liEl.classList.add('sr-item--selected');
    const wrap = liEl.querySelector('.sr-qty-wrap');
    if (wrap) wrap.classList.add('sr-qty-wrap--visible');
    const inp = liEl.querySelector('.sr-qty-input');
    if (inp) setTimeout(() => { inp.focus(); inp.select(); }, 50);
  }
  if (navigator.vibrate) navigator.vibrate(25);
  updateCommitBar();
}

// Actualiza el texto y estado de la barra de confirmación
function updateCommitBar() {
  const total = state.searchCart.size;
  const conQty = [...state.searchCart.values()].filter(e => e.qty > 0).length;
  const sinQty = total - conQty;

  // Clase para iluminar label
  els.searchCommitBar.classList.toggle('has-items', total > 0);

  if (total === 0) {
    els.scbLabel.textContent = '0 productos seleccionados';
    els.btnCommitSearch.disabled = true;
    els.btnCommitSearch.innerHTML = btnCommitHTML('Agregar a la lista');
    return;
  }

  els.scbLabel.textContent = `${total} seleccionado${total !== 1 ? 's' : ''}` +
    (sinQty > 0 ? ` · ${sinQty} sin cantidad` : '');

  els.btnCommitSearch.disabled = conQty === 0;
  const label = conQty > 0 ? `Agregar ${conQty} a la lista` : 'Agregá las cantidades';
  els.btnCommitSearch.innerHTML = btnCommitHTML(label);
}

function btnCommitHTML(label) {
  return `${label}
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"
         stroke-linecap="round" stroke-linejoin="round" width="15" height="15">
      <line x1="12" y1="5" x2="12" y2="19"/>
      <line x1="5" y1="12" x2="19" y2="12"/>
    </svg>`;
}

// Confirma el carrito: envía todos los items con qty > 0 al pedido principal
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
      state.order.push({
        id: Date.now() + Math.random(),
        ean: p.EAN || p.INTERNO || '',
        interno: p.INTERNO || '',
        descripcion: p.DESCRIPCION || '',
        proveedor: p.PROVEEDOR || '',
        gramaje: p.GRAMAJE || '',
        uxb: p.UXB != null ? p.UXB : '',
        cantidad: qty,
      });
      added++;
    }
  });

  renderTable();
  updateBadge();

  const parts = [];
  if (added) parts.push(`${added} nuevo${added !== 1 ? 's' : ''}`);
  if (updated) parts.push(`${updated} actualizado${updated !== 1 ? 's' : ''}`);
  showToast('ok', parts.join(' · ') + ` en la lista`);
  if (navigator.vibrate) navigator.vibrate([60, 30, 60]);

  state.searchCart.clear();
  closeSearchModal();
}

// ── Helpers ──
function normalize(str) {
  if (!str) return '';
  return String(str).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function highlight(text, query) {
  if (!text || !query) return esc(text);
  const normText = normalize(text);
  const normQuery = normalize(query);
  const result = [];
  let i = 0;
  while (i < text.length) {
    const idx = normText.indexOf(normQuery, i);
    if (idx === -1) { result.push(esc(text.slice(i))); break; }
    result.push(esc(text.slice(i, idx)));
    result.push(`<mark>${esc(text.slice(idx, idx + query.length))}</mark>`);
    i = idx + query.length;
  }
  return result.join('');
}

// ── Busca un ítem ya existente en el pedido por EAN o INTERNO ──
function findOrderItem(p) {
  const ean = String(p?.EAN || p?.INTERNO || '').trim();
  const interno = String(p?.INTERNO || '').trim();
  return state.order.find(item => {
    if (ean && String(item.ean).trim() === ean) return true;
    if (interno && String(item.interno).trim() === interno) return true;
    return false;
  }) || null;
}

// ── Pedido (scanner) ──
function addToOrder() {
  const p = state.currentProduct;
  const qty = parseFloat(els.fQty.value.replace(',', '.'));
  if (isNaN(qty) || qty <= 0) { els.eQty.classList.add('show'); els.fQty.focus(); return; }

  const existing = findOrderItem(p);
  if (existing) {
    existing.cantidad += qty;
    renderTable();
    updateBadge();
    showToast('ok', `"${p?.DESCRIPCION || ''}" — cantidad actualizada (total: ×${existing.cantidad})`);
  } else {
    state.order.push({
      id: Date.now(),
      ean: p?.EAN || p?.INTERNO || '',
      interno: p?.INTERNO || '',
      descripcion: p?.DESCRIPCION || '',
      proveedor: p?.PROVEEDOR || '',
      gramaje: p?.GRAMAJE || '',
      uxb: p?.UXB != null ? p.UXB : '',
      cantidad: qty,
    });
    renderTable();
    updateBadge();
    showToast('ok', `"${p?.DESCRIPCION || ''}" agregado (×${qty})`);
  }
  if (navigator.vibrate) navigator.vibrate(60);
  resetCurrentProduct();
  els.fQty.value = '';

  // In ext mode: reset the input and refocus it for the next scan
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
  renderTable();
  updateBadge();
}

function clearOrder() {
  if (!state.order.length) return;
  if (!confirm('¿Limpiar toda la lista del pedido?')) return;
  state.order = [];
  renderTable();
  updateBadge();
  showToast('wrn', 'Lista vaciada');
}

function renderTable() {
  const has = state.order.length > 0;
  els.tableEmpty.style.display = has ? 'none' : 'block';
  els.orderTable.style.display = has ? 'table' : 'none';
  els.btnExport.disabled = !has;
  els.btnExportPDF.disabled = !has;
  els.btnClearOrder.disabled = !has;

  els.orderBody.innerHTML = state.order.map(item => `
    <tr>
      <td class="td-code">${esc(String(item.ean || '—'))}</td>
      <td class="td-desc">${esc(item.descripcion)}</td>
      <td class="td-prov">${esc(String(item.gramaje || '—'))}</td>
      <td class="td-prov">${item.uxb !== '' ? item.uxb + ' u.' : '—'}</td>
      <td class="td-qty">${item.cantidad}</td>
      <td class="td-actions"><button class="btn-remove" onclick="removeItem(${item.id})">✕</button></td>
    </tr>`).join('');
}

function updateBadge() {
  const n = state.order.length;
  els.badgeCount.textContent = `${n} ${n === 1 ? 'ítem' : 'ítems'}`;
  els.badgeCount.classList.toggle('show', n > 0);
}

// ── Scanner ──
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
        fps: 10,
        qrbox: { width: 260, height: 120 },
        formatsToSupport: [
          Html5QrcodeSupportedFormats.EAN_13,
          Html5QrcodeSupportedFormats.EAN_8,
          Html5QrcodeSupportedFormats.UPC_A,
          Html5QrcodeSupportedFormats.UPC_E,
          Html5QrcodeSupportedFormats.CODE_128,
          Html5QrcodeSupportedFormats.CODE_39,
          Html5QrcodeSupportedFormats.ITF,
        ],
        aspectRatio: 4 / 3,
      },
      onBarcodeScanned,
      null
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
    try { if (state.scannerActive) await state.html5QrCode.stop(); state.html5QrCode.clear(); } catch (_) { }
    state.html5QrCode = null;
  }
  state.scannerActive = false;
  els.scanOverlay.classList.remove('open');
  document.body.style.overflow = '';
}

// ── Utils ──
let toastTimer = null;
function showToast(type, msg, duration = 3000) {
  els.toastMsg.textContent = msg;
  els.toast.className = `toast ${type}`;
  void els.toast.offsetWidth;
  els.toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => els.toast.classList.remove('show'), duration);
}

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── Export Excel ──
// ── Export Excel ──
// Los códigos de barras se renderizan con la fuente "Libre Barcode 128".
// Si la fuente no está instalada en la PC, la celda muestra el EAN en texto plano (igualmente útil).
// Descarga gratuita: https://fonts.google.com/specimen/Libre+Barcode+128
//
// VENTAJAS sobre el método anterior (imágenes flotantes):
//   ✅ La macro de consolidado copia correctamente las filas (End(xlUp) detecta el valor)
//   ✅ PasteSpecial copia el contenido sin perder datos
//   ✅ Archivos más livianos (sin imágenes embebidas)
//   ✅ Fallback legible si la fuente no está instalada

async function exportExcel() {
  if (!state.order.length) return;

  // ── Validaciones ──────────────────────────────────────────
  const sucursal = els.fSucursal.value;
  if (!sucursal) {
    els.eSucursal.classList.add('show');
    els.fSucursal.classList.add('err');
    els.fSucursal.focus();
    showToast('err', 'Seleccioná una sucursal antes de exportar.');
    return;
  }

  const concepto = els.fConcepto.value.trim();
  if (!concepto) {
    els.eConcepto.classList.add('show');
    els.fConcepto.classList.add('err');
    els.fConcepto.focus();
    showToast('err', 'Ingresá un concepto general.');
    return;
  }

  // ── Metadatos ─────────────────────────────────────────────
  const now          = new Date();
  const fecha        = now.toLocaleDateString('es-AR');
  const fechaFile    = now.toISOString().slice(0, 10);
  const remitoId     = now.getTime();
  const conceptoFile = concepto.replace(/\s+/g, '_').toLowerCase();
  const filename     = `${sucursal}_${fechaFile}_${conceptoFile}_${remitoId}.xlsx`;

  // ── Workbook ──────────────────────────────────────────────
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Pedido', { views: [{ state: 'normal' }] });

  ws.pageSetup = {
    paperSize: 9,
    orientation: 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.3, footer: 0.3 },
  };

  // ── Layout de columnas ────────────────────────────────────
  // A : Código de barras  → texto con fuente "Libre Barcode 128"
  // B : Espaciador angosto
  // C : EAN (texto plano, fuente normal)
  // D : Descripción
  // E : Espaciador angosto
  // F : Gramaje
  // G : UxB
  // H : Cantidad
  ws.columns = [
    { key: 'barcode', width: 30   },  // A – ancho generoso para el barcode
    { key: 'sp1',     width: 0.05  },  // B – espaciador visual
    { key: 'ean',     width: 16   },  // C
    { key: 'desc',    width: 48   },  // D
    { key: 'sp2',     width: 0.05  },  // E – espaciador visual
    { key: 'gram',    width: 14   },  // F
    { key: 'uxb',     width: 6    },  // G
    { key: 'cant',    width: 8    },  // H
  ];

  // ── Helpers ───────────────────────────────────────────────
  const thinBorder = {
    top:    { style: 'thin' },
    left:   { style: 'thin' },
    bottom: { style: 'thin' },
    right:  { style: 'thin' },
  };

  const centerMiddle = { horizontal: 'center', vertical: 'middle', wrapText: false };

  /**
   * Aplica alineación centrada y borde fino a una celda.
   * @param {ExcelJS.Cell} cell
   * @param {{ noborder?: boolean, bold?: boolean }} opts
   */
  function styleCell(cell, opts = {}) {
    cell.alignment = centerMiddle;
    if (!opts.noborder) cell.border = thinBorder;
    if (opts.bold)      cell.font   = { ...(cell.font || {}), bold: true };
  }

  // ── Filas de cabecera (1–3) ───────────────────────────────
  const lastCol = 'H';

  [
    [1, `REMITO/PEDIDO — ${sucursal.toUpperCase()} — ${fecha}`, true,  12],
    [2, `Concepto: ${concepto}`,                                false, 10],
    [3, `ID: ${remitoId}`,                                      false,  9],
  ].forEach(([row, val, bold, size]) => {
    ws.mergeCells(`A${row}:${lastCol}${row}`);
    const cell  = ws.getCell(`A${row}`);
    cell.value  = val;
    cell.font   = { bold, size };
    styleCell(cell);
  });

  // ── Fila de títulos (fila 4) ──────────────────────────────
  const headerMap = {
    A: 'COD-BAR',
    B: '',
    C: 'EAN',
    D: 'DESCRIPCION',
    E: '',
    F: 'GRAMAJE',
    G: 'UxB',
    H: 'CANTIDAD',
  };

  for (const [col, val] of Object.entries(headerMap)) {
    const cell  = ws.getCell(`${col}4`);
    cell.value  = val;
    cell.font   = { bold: true, size: 10 };
    cell.alignment = centerMiddle;
    // Espaciadores sin borde
    if (col !== 'B' && col !== 'E') cell.border = thinBorder;
  }
  ws.getRow(4).height = 18;

  // ── Filas de datos (fila 5 en adelante) ───────────────────
  for (let i = 0; i < state.order.length; i++) {
    const item     = state.order[i];
    const rowIndex = 5 + i;
    const eanCode  = String(item.ean || '');

    // Alto máximo 20 pt — el barcode y el resto del contenido se ajustan a ese límite
    ws.getRow(rowIndex).height = 30;

    // ── A — Código de barras como texto con fuente Libre Barcode 128 ──
    // Con height 20 el size 14 genera barras que entran justas y son escaneables.
    // Si la fuente no está instalada, se ve el EAN en texto plano.
    const cellA     = ws.getCell(`A${rowIndex}`);
    cellA.value     = eanCode;
    cellA.border    = thinBorder;
    cellA.alignment = { horizontal: 'center', vertical: 'middle' };
    cellA.font      = {
      name: 'Libre Barcode 128',
      size: 14,                   // ajustado para caber en height 20
      color: { argb: 'FF000000' },
    };

    // ── B — Espaciador sin borde ──
    ws.getCell(`B${rowIndex}`).value = '';

    // ── C — EAN en texto plano, fuente monoespaciada ──
    const cellC     = ws.getCell(`C${rowIndex}`);
    cellC.value     = eanCode;
    styleCell(cellC);
    cellC.font      = { name: 'Courier New', size: 10 };

    // ── D — Descripción ──
    const cellD     = ws.getCell(`D${rowIndex}`);
    cellD.value     = item.descripcion || '';
    styleCell(cellD);

    // ── E — Espaciador sin borde ──
    ws.getCell(`E${rowIndex}`).value = '';

    // ── F — Gramaje ──
    const cellF     = ws.getCell(`F${rowIndex}`);
    cellF.value     = item.gramaje || '';
    styleCell(cellF);

    // ── G — UxB ──
    const cellG     = ws.getCell(`G${rowIndex}`);
    cellG.value     = item.uxb !== '' ? item.uxb : '';
    styleCell(cellG);

    // ── H — Cantidad ──
    const cellH     = ws.getCell(`H${rowIndex}`);
    cellH.value     = item.cantidad;
    styleCell(cellH);
    cellH.font      = { bold: true, size: 10 };
  }

  // Autofit columna D (descripción)
  autoFitColumns(ws, [4]);

  // ── Guardar y subir ────────────────────────────────────────
  try {
    const buf  = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/octet-stream' });
    saveAs(blob, filename);
    showToast('ok', 'Excel descargado.');
    await uploadToDrive(blob, filename);
  } catch (err) {
    console.error(err);
    showToast('err', 'No se pudo generar el Excel.');
  }
}

function autoFitColumns(worksheet, cols = []) {
  cols.forEach(n => { const col = worksheet.getColumn(n); let max = 10; col.eachCell({ includeEmpty: true }, c => { const v = c.value ? c.value.toString() : ''; max = Math.max(max, v.length + 2); }); col.width = max; });
}

// ── Export PDF ──
async function exportPDF() {
  if (!state.order.length) return;

  const now          = new Date();
  const fecha        = now.toLocaleDateString('es-AR');
  const fechaFile    = now.toISOString().slice(0, 10);
  const remitoId     = now.getTime();

  const sucursal = els.fSucursal.value;
  if (!sucursal) {
    els.eSucursal.classList.add('show'); els.fSucursal.classList.add('err');
    els.fSucursal.focus();
    showToast('err', 'Seleccioná una sucursal antes de exportar.');
    return;
  }
  const concepto = els.fConcepto.value.trim();
  if (!concepto) {
    els.eConcepto.classList.add('show'); els.fConcepto.classList.add('err');
    els.fConcepto.focus();
    showToast('err', 'Ingresá un concepto general.');
    return;
  }

  const conceptoFile = concepto.replace(/\s+/g, '_').toLowerCase();
  const filename     = `${sucursal}_${fechaFile}_${conceptoFile}_${remitoId}.pdf`;

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'pt', format: 'a4', orientation: 'landscape' });

  // ── Geometría ─────────────────────────────────────────────
  // A4 landscape = 841.89 × 595.28 pt
  const PAGE_W  = doc.internal.pageSize.getWidth();   // 841.89
  const MARGIN  = 36;
  const TABLE_W = PAGE_W - MARGIN * 2;                // ~769.89 pt

  // ── Encabezado ────────────────────────────────────────────
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(14);
  doc.text(`REMITO/PEDIDO — ${sucursal.toUpperCase()} — ${fecha}`, MARGIN, 34);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.text(`Concepto: ${concepto}`, MARGIN, 50);
  doc.text(`ID: ${remitoId}`,       MARGIN, 63);

  // ── Widths proporcionales — deben sumar 1.0 ───────────────
  //  Col 0 Barcode     15 %
  //  Col 1 EAN          9 %
  //  Col 2 Descripción 27 %
  //  Col 3 Proveedor   15 %
  //  Col 4 Interno      8 %
  //  Col 5 Gramaje      9 %
  //  Col 6 UxB          8.5 %
  //  Col 7 Cant.        8.5 %
  //  ─────────────────────────
  //  Total            100 %
  const W = {
    barcode:  TABLE_W * 0.150,
    ean:      TABLE_W * 0.090,
    desc:     TABLE_W * 0.270,
    prov:     TABLE_W * 0.150,
    interno:  TABLE_W * 0.080,
    gramaje:  TABLE_W * 0.090,
    uxb:      TABLE_W * 0.085,
    cant:     TABLE_W * 0.085,
  };

  // ── Filas ─────────────────────────────────────────────────
  const ROW_H = 34;
  const rows  = [];

  for (const item of state.order) {
    const eanCode = String(item.ean || '');
    let barcodeImg = null;

    if (eanCode) {
      try {
        const canvas = document.createElement('canvas');
        const fmt    = /^\d{13}$/.test(eanCode) ? 'ean13' : 'code128';
        try {
          JsBarcode(canvas, eanCode, { format: fmt,       displayValue: false, height: 38, margin: 2, width: 1.4 });
        } catch {
          JsBarcode(canvas, eanCode, { format: 'code128', displayValue: false, height: 38, margin: 2, width: 1.4 });
        }
        barcodeImg = canvas.toDataURL('image/png');
      } catch (e) { /* si falla, celda queda vacía */ }
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
    ]);
  }

  // ── Tabla ─────────────────────────────────────────────────
  doc.autoTable({
    startY: 76,
    margin: { left: MARGIN, right: MARGIN },

    head: [[
      'Cód. Barras', 'EAN', 'Descripción', 'Proveedor',
      'Interno', 'Gramaje', 'UxB', 'Cant.',
    ]],
    body: rows,

    styles: {
      halign:        'center',
      valign:        'middle',
      fontSize:       7.5,
      cellPadding:    3,
      minCellHeight:  ROW_H,
      overflow:      'ellipsize',
    },
    headStyles: {
      fillColor:  [30, 30, 50],
      textColor:  240,
      fontStyle: 'bold',
      fontSize:   8,
    },
    alternateRowStyles: {
      fillColor: [245, 245, 250],
    },

    columnStyles: {
      0: { cellWidth: W.barcode  },
      1: { cellWidth: W.ean      },
      2: { cellWidth: W.desc,    halign: 'left' },
      3: { cellWidth: W.prov,    halign: 'left' },
      4: { cellWidth: W.interno  },
      5: { cellWidth: W.gramaje  },
      6: { cellWidth: W.uxb      },
      7: { cellWidth: W.cant,    fontStyle: 'bold' },
    },

    didDrawCell: (data) => {
      if (data.section !== 'body') return;
      if (data.column.index !== 0) return;
      const img = data.cell.raw?.barcode;
      if (!img) return;

      const pad = 3;
      const iw  = data.cell.width  - pad * 2;
      const ih  = data.cell.height - pad * 2;
      doc.addImage(img, 'PNG', data.cell.x + pad, data.cell.y + pad, iw, ih);
    },
  });

  const pdfBlob = doc.output('blob');
  doc.save(filename);
  await uploadToDrive(pdfBlob, filename);
  showToast('ok', 'PDF generado y subido a Drive.');
}

// ── Drive upload ──
function blobToBase64(blob) {
  return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onloadend = () => resolve(reader.result.split(',')[1]); reader.onerror = reject; reader.readAsDataURL(blob); });
}
async function uploadToDrive(blob, filename) {
  const url = "https://script.google.com/macros/s/AKfycbzIRnuzUWTjG38fxttQbkvJ7Br_wYSYs5UeaJO9EHnncy7jr8vQivcfQLot0xqSzwsq/exec";
  try {
    const base64Data = await blobToBase64(blob);
    const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ filename, mimeType: blob.type, file: base64Data }) });
    const json = await res.json();
    if (json.ok) showToast('ok', `Archivo subido a Drive: ${filename}`);
    else showToast('err', 'Error al subir a Drive: ' + (json.error || 'Desconocido'));
  } catch (err) { console.error(err); showToast('err', 'Error al conectar con Google Drive.'); }
}