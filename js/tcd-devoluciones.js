const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwj0qEOm9THbYxw0TYek2Oot3dlL1wn7YmPLtYknFzrGBQJXFnd-kh7yxXtFgYFyC-B/exec';

// ─────────────────────────────────────────────
// Estado global
// ─────────────────────────────────────────────
const state = {
    photoBase64: null,
    photoMime: null,
    photoName: null,
    scannerActive: false,
    html5QrCode: null,
    submitting: false,
    sentCount: parseInt(localStorage.getItem('devCount') || '0'),
    productData: null,
    productos: [],
    inputMode: 'cam',
    extDebounce: null,
    searchTimer: null,
    searchField: 'all',
    isPesable: false,
    pesoKg: null,
};

// ─────────────────────────────────────────────
// Elementos del DOM
// ─────────────────────────────────────────────
const $ = (id) => document.getElementById(id);

const els = {
    scanOverlay: $('scanOverlay'),
    btnOpenScanner: $('btnOpenScanner'),
    btnCloseScanner: $('btnCloseScanner'),
    btnManualEntry: $('btnManualEntry'),
    camErr: $('camErr'),
    sHint: $('sHint'),
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
    fBarcode: $('fBarcode'),
    scannedOk: $('scannedOk'),
    modeTabCam: $('modeTabCam'),
    modeTabExt: $('modeTabExt'),
    panelCam: $('panelCam'),
    panelExt: $('panelExt'),
    extInput: $('extInput'),
    btnExtClear: $('btnExtClear'),
    extStatus: $('extStatus'),
    extStatusText: $('extStatusText'),
    productCard: $('productCard'),
    pcState: $('pcState'),
    pcDescripcion: $('pcDescripcion'),
    pcProveedor: $('pcProveedor'),
    pcGramaje: $('pcGramaje'),
    pcUxb: $('pcUxb'),
    pcSector: $('pcSector'),
    pcEan: $('pcEan'),
    fPhoto: $('fPhoto'),
    photoZone: $('photoZone'),
    photoPlaceholder: $('photoPlaceholder'),
    photoPreview: $('photoPreview'),
    photoImg: $('photoImg'),
    photoName: $('photoName'),
    btnCamera: $('btnCamera'),
    btnGallery: $('btnGallery'),
    btnRemovePhoto: $('btnRemovePhoto'),
    fEmail: $('fEmail'),
    fDesc: $('fDesc'),
    fQty: $('fQty'),
    qtyPesableHint: $('qtyPesableHint'),
    fExp: $('fExp'),
    fBranch: $('fBranch'),
    fEvent: $('fEvent'),
    fDiscount: $('fDiscount'),
    discountWrap: $('discountWrap'),
    fNote: $('fNote'),
    fLot: $('fLot'),
    fComment: $('fComment'),
    btnSubmit: $('btnSubmit'),
    btnSubmitText: $('btnSubmitText'),
    btnClear: $('btnClear'),
    toast: $('toast'),
    toastMsg: $('toastMsg'),
    toastDot: $('toastDot'),
    counterBadge: $('counterBadge'),
    cfgBanner: $('cfgBanner'),
};

const MOTIVOS_VENCIMIENTO = ['CONTROL DE VENCIMIENTO'];
const SECTORES_PESABLES_KEYS = ['FIAMBRE', 'CARNICER', 'VERDULER', 'FRUTA'];

// ─────────────────────────────────────────────
// INIT
// ─────────────────────────────────────────────
async function init() {
    await cargarProductos();

    if (APPS_SCRIPT_URL !== 'PEGA_AQUI_TU_URL_DE_APPS_SCRIPT') {
        if (els.cfgBanner) els.cfgBanner.style.display = 'none';
    }

    updateCounter();

    if (!dentroDelHorario()) {
        showToast('wrn', mensajeHorario(), 9000);
    }

    // ── Mode switcher ──
    function setInputMode(mode) {
        state.inputMode = mode;
        localStorage.setItem('devInputMode', mode);
        const isCam = mode === 'cam';
        els.modeTabCam.classList.toggle('mode-tab--active', isCam);
        els.modeTabExt.classList.toggle('mode-tab--active', !isCam);
        els.panelCam.style.display = isCam ? '' : 'none';
        els.panelExt.style.display = isCam ? 'none' : '';
        if (!isCam) {
            hideProductCard();
            els.fBarcode.value = '';
            els.scannedOk.classList.remove('show');
            setTimeout(() => els.extInput.focus(), 80);
            setExtStatus('ready', 'Listo — esperando lectura');
        }
    }
    const savedMode = localStorage.getItem('devInputMode') || 'cam';
    setInputMode(savedMode);
    els.modeTabCam.addEventListener('click', () => setInputMode('cam'));
    els.modeTabExt.addEventListener('click', () => setInputMode('ext'));

    // ── Scanner events ──
    els.btnOpenScanner.addEventListener('click', openScanner);
    els.btnCloseScanner.addEventListener('click', closeScanner);
    els.btnManualEntry.addEventListener('click', closeScanner);

    // ── Search modal events ──
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

    // ── External scanner input ──
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
        els.fBarcode.value = '';
        els.scannedOk.classList.remove('show');
        hideProductCard();
        setExtStatus('ready', 'Listo — esperando lectura');
        els.extInput.focus();
    });

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            if (els.searchOverlay.classList.contains('open')) closeSearchModal();
            if (els.scanOverlay.classList.contains('open')) closeScanner();
        }
        const active = document.activeElement;
        const isTypingElsewhere = active && (
            active.tagName === 'INPUT' ||
            active.tagName === 'SELECT' ||
            active.tagName === 'TEXTAREA'
        );
        if (state.inputMode === 'ext'
            && !els.searchOverlay.classList.contains('open')
            && !els.scanOverlay.classList.contains('open')
            && !isTypingElsewhere
            && e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
            els.extInput.focus();
        }
    });

    // ── Photo events ──
    els.btnCamera.addEventListener('click', () => triggerFileInput(true));
    els.btnGallery.addEventListener('click', () => triggerFileInput(false));
    els.btnRemovePhoto.addEventListener('click', removePhoto);
    els.photoZone.addEventListener('click', onPhotoZoneClick);
    els.fPhoto.addEventListener('change', onPhotoSelected);

    els.photoZone.addEventListener('dragover', (e) => { e.preventDefault(); els.photoZone.classList.add('drag-over'); });
    els.photoZone.addEventListener('dragleave', () => els.photoZone.classList.remove('drag-over'));
    els.photoZone.addEventListener('drop', (e) => {
        e.preventDefault();
        els.photoZone.classList.remove('drag-over');
        const file = e.dataTransfer.files[0];
        if (file) processPhoto(file);
    });

    // ── Tipo de registro ──
    els.fEvent.addEventListener('change', () => {
        const val = els.fEvent.value;
        applyDateRestriction(val);
        toggleDiscountInput(val);
        toggleBdBadge(val);
        updateSubmitButtonLabel(val);
        clearFieldError('fEvent');
        if (els.fExp.value) validateField('fExp');
    });

    // ── Botón principal: enviar ──
    els.btnSubmit.addEventListener('click', submitSingle);
    els.btnClear.addEventListener('click', clearForm);

    // Live validation
    const requiredFields = ['fEmail', 'fQty', 'fExp', 'fBranch', 'fEvent'];
    requiredFields.forEach((id) => {
        const el = $(id);
        if (!el) return;
        el.addEventListener('blur', () => validateField(id));
        el.addEventListener('input', () => {
            if (id === 'fQty' && el.value.includes(',')) {
                const pos = el.selectionStart;
                el.value = el.value.replace(',', '.');
                try { el.setSelectionRange(pos, pos); } catch (_) { }
            }
            clearFieldError(id);
        });
    });
    ['fBranch', 'fEvent'].forEach((id) => {
        const el = $(id);
        if (!el) return;
        el.addEventListener('change', () => clearFieldError(id));
    });
    els.fExp.addEventListener('change', () => validateField('fExp'));
}

document.addEventListener('DOMContentLoaded', init);

// ─────────────────────────────────────────────
// UI HELPERS — BADGE Y LABEL
// ─────────────────────────────────────────────

function toggleBdBadge(motivo) {
    const badge = $('bdBadge');
    if (!badge || !motivo) { badge && (badge.style.display = 'none'); return; }
    const esVen = MOTIVOS_VENCIMIENTO.includes(motivo);
    badge.className = 'bd-badge ' + (esVen ? 'bd-badge--ven' : 'bd-badge--dev');
    badge.textContent = esVen
        ? '📋 Se guardará en: Control de Vencimientos'
        : '📦 Se guardará en: Devoluciones / Acciones';
    badge.style.display = 'flex';
}

function updateSubmitButtonLabel(motivo) {
    if (!els.btnSubmitText) return;
    const esVen = MOTIVOS_VENCIMIENTO.includes(motivo);
    els.btnSubmitText.textContent = esVen
        ? 'Enviar control de vencimiento'
        : 'Enviar registro';
}

// ─────────────────────────────────────────────
// PRODUCTOS
// ─────────────────────────────────────────────
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

function normalizeCode(val) {
    if (val == null || val === '') return '';
    return String(val).trim().replace(/^0+(?=\d)/, '');
}

async function fetchProducto(code) {
    const normCode = normalizeCode(code);
    return state.productos.find(p =>
        normalizeCode(p.EAN) === normCode || normalizeCode(p.INTERNO) === normCode
    ) || null;
}

// ─────────────────────────────────────────────
// DETECCIÓN DE PRODUCTOS PESABLES
// ─────────────────────────────────────────────
function parsePesableEAN(ean) {
    const s = String(ean).trim().replace(/\s/g, '');
    if (s.length !== 13 || !s.startsWith('2')) return null;
    const internoRaw = s.slice(2, 7);
    const pesoRaw = s.slice(7, 12);
    const interno = String(parseInt(internoRaw, 10));
    const pesoGr = parseInt(pesoRaw, 10);
    if (isNaN(pesoGr) || pesoGr <= 0) return null;
    return { interno, pesoGr, pesoKg: pesoGr / 1000 };
}

function esSectorPesable(data) {
    const sectorUp = (data.SECTOR || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const seccionUp = (data.SECCION || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const eanStr = String(data.EAN || data.INTERNO || '');
    return (
        SECTORES_PESABLES_KEYS.some(k => sectorUp.includes(k) || seccionUp.includes(k)) ||
        eanStr.startsWith('23')
    );
}

async function triggerLookup(code) {
    showProductCardLoading();
    const pesable = parsePesableEAN(code);
    const lookupCode = pesable ? pesable.interno : code;
    const data = await fetchProducto(lookupCode);
    if (data) fillProductData(data, pesable);
    else showProductCardNotFound();
}

async function triggerExtLookup(code) {
    els.extInput.classList.remove('ext-input--reading', 'ext-input--found');
    setExtStatus('reading', `Buscando: ${code}`);
    showProductCardLoading();
    const pesable = parsePesableEAN(code);
    const lookupCode = pesable ? pesable.interno : code;
    const data = await fetchProducto(lookupCode);
    if (data) {
        els.fBarcode.value = code;
        clearFieldError('fBarcode');
        els.scannedOk.classList.add('show');
        els.extInput.classList.add('ext-input--found');
        setExtStatus('found', `✓ ${data.DESCRIPCION ? data.DESCRIPCION.slice(0, 40) : 'Producto encontrado'}`);
        fillProductData(data, pesable);
        if (navigator.vibrate) navigator.vibrate([60, 20, 60]);
        setTimeout(() => pesable ? els.fExp.focus() : els.fQty.focus(), 150);
    } else {
        els.fBarcode.value = '';
        state.productData = null;
        showProductCardNotFound();
        setExtStatus('notfound', '⚠ Código no encontrado — intentá la búsqueda manual');
        showToast('wrn', 'Código no encontrado — intentá buscarlo manualmente.');
    }
}

function fillProductData(data, pesable = null) {
    state.productData = data;
    state.isPesable = !!pesable;
    state.pesoKg = pesable ? pesable.pesoKg : null;

    els.fBarcode.value = String(data.EAN || data.INTERNO || '');
    els.fDesc.value = data.DESCRIPCION || '';
    clearFieldError('fBarcode');

    if (pesable) {
        els.fQty.value = pesable.pesoKg.toFixed(3);
        els.fQty.readOnly = true;
        els.fQty.classList.add('qty--pesable');
        els.fQty.setAttribute('step', '0.001');
        els.fQty.setAttribute('min', '0.001');
        clearFieldError('fQty');
        if (els.qtyPesableHint) {
            els.qtyPesableHint.textContent = `⚖ Pesable · ${pesable.pesoGr.toLocaleString('es-AR')} g = ${pesable.pesoKg.toFixed(3)} kg`;
            els.qtyPesableHint.style.display = 'flex';
        }
        const hintLabel = $('fQtyHint');
        if (hintLabel) hintLabel.textContent = '(kg — automático)';
        const eQtyEl = $('eQty');
        if (eQtyEl) eQtyEl.textContent = 'Peso inválido.';
    } else {
        els.fQty.readOnly = false;
        els.fQty.classList.remove('qty--pesable');

        const sectorPesable = esSectorPesable(data);
        state.isPesable = sectorPesable;

        if (sectorPesable) {
            els.fQty.setAttribute('step', '0.001');
            els.fQty.setAttribute('min', '0.001');
            if (els.qtyPesableHint) {
                const sectName = (data.SECTOR || data.SECCION || 'Pesable');
                els.qtyPesableHint.textContent = `⚖ ${sectName} — ingresá el peso en kg (ej: 1.250)`;
                els.qtyPesableHint.style.display = 'flex';
            }
            const hintLabel = $('fQtyHint');
            if (hintLabel) hintLabel.textContent = '(kg)';
            const eQtyEl = $('eQty');
            if (eQtyEl) eQtyEl.textContent = 'Ingresá el peso en kg (mayor a 0).';
        } else {
            els.fQty.setAttribute('step', '1');
            els.fQty.setAttribute('min', '1');
            if (els.qtyPesableHint) els.qtyPesableHint.style.display = 'none';
            const hintLabel = $('fQtyHint');
            if (hintLabel) hintLabel.textContent = '(enteros)';
            const eQtyEl = $('eQty');
            if (eQtyEl) eQtyEl.textContent = 'Número ≥ 1';
        }
    }

    els.pcDescripcion.textContent = data.DESCRIPCION || '-';
    els.pcProveedor.textContent = data.PROVEEDOR || '-';
    els.pcGramaje.textContent = data.GRAMAJE || '-';
    els.pcUxb.textContent = data.UXB != null ? `${data.UXB} u.` : '-';

    const sector = data.SECTOR || '';
    const seccion = data.SECCION || '';
    els.pcSector.textContent = sector && seccion ? `${sector} › ${seccion}` : (sector || seccion || '-');

    const eanParts = [];
    if (data.EAN) eanParts.push(`EAN ${data.EAN}`);
    if (data.INTERNO && String(data.INTERNO) !== String(data.EAN)) eanParts.push(`INT ${data.INTERNO}`);
    els.pcEan.textContent = eanParts.length ? eanParts.join('  ·  ') : '-';

    const pvpEl = $('pcPvp');
    if (pvpEl) {
        const pvp = data['PVP SUPER'];
        pvpEl.textContent = pvp != null
            ? `$${Number(pvp).toLocaleString('es-AR', { minimumFractionDigits: 2 })}` : '-';
    }

    if (pesable) {
        showProductCardFound();
        els.pcState.textContent = `Pesable · ${pesable.pesoKg.toFixed(3)} kg`;
        showToast('ok', `Pesable encontrado · ${pesable.pesoKg.toFixed(3)} kg`);
    } else {
        showProductCardFound();
        if (state.isPesable) {
            els.pcState.textContent = `Producto pesable — ${sector || seccion}`;
            showToast('ok', 'Producto pesable — ingresá el peso en kg');
        } else {
            showToast('ok', 'Producto encontrado');
        }
    }
}

function showProductCardLoading() {
    els.productCard.classList.add('visible');
    els.productCard.dataset.state = 'loading';
    els.pcState.textContent = 'Buscando producto...';
}
function showProductCardFound() {
    els.productCard.classList.add('visible');
    els.productCard.dataset.state = 'found';
    els.pcState.textContent = 'Producto encontrado';
}
function showProductCardNotFound() {
    els.productCard.classList.add('visible');
    els.productCard.dataset.state = 'notfound';
    els.pcState.textContent = 'No encontrado';
    ['pcDescripcion', 'pcProveedor', 'pcGramaje', 'pcUxb', 'pcEan'].forEach(k => { els[k].textContent = '-'; });
    showToast('wrn', 'Producto no encontrado en la base de datos.');
}
function hideProductCard() {
    els.productCard.classList.remove('visible');
    els.productCard.dataset.state = '';
    state.productData = null;
}

// ─────────────────────────────────────────────
// EXT STATUS
// ─────────────────────────────────────────────
function setExtStatus(type, text) {
    els.extStatus.className = 'ext-status';
    if (type) els.extStatus.classList.add(`ext-status--${type}`);
    els.extStatusText.textContent = text;
}

// ─────────────────────────────────────────────
// SEARCH MODAL
// ─────────────────────────────────────────────
function openSearchModal() {
    els.searchOverlay.classList.add('open');
    document.body.style.overflow = 'hidden';
    els.searchInput.value = '';
    els.btnClearSearch.style.display = 'none';
    els.searchCount.textContent = '';
    renderSearchResults('');
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
                normalize(p.DESCRIPCION).includes(q) || normalize(p.PROVEEDOR).includes(q) ||
                normalize(p.SECTOR).includes(q) || normalize(p.SECCION).includes(q) ||
                normalize(String(p.EAN ?? '')).includes(q) || normalize(String(p.INTERNO ?? '')).includes(q)
            );
        }
        if (field === 'EAN') return normalize(String(p.EAN ?? '')).includes(q) || normalize(String(p.INTERNO ?? '')).includes(q);
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
        const desc = highlight(p.DESCRIPCION || '—', q);
        const prov = highlight(p.PROVEEDOR || '—', q);
        const eanRaw = String(p.EAN ?? '');
        const intRaw = String(p.INTERNO ?? '');
        const eanBadge = eanRaw ? `<span class="sr-ean">EAN ${highlight(eanRaw, q)}</span>` : '';
        const intBadge = intRaw && intRaw !== eanRaw ? `<span class="sr-ean sr-ean--int">INT ${highlight(intRaw, q)}</span>` : '';
        const sect = p.SECTOR ? `<span class="sr-tag">${esc(p.SECTOR)}</span>` : '';
        const secc = p.SECCION ? `<span class="sr-tag">${esc(p.SECCION)}</span>` : '';
        const gramaje = p.GRAMAJE ? `<span class="sr-tag sr-tag--gramaje">${esc(p.GRAMAJE)}</span>` : '';
        const pvp = p['PVP SUPER'] != null
            ? `<span class="sr-tag sr-tag--pvp">$${Number(p['PVP SUPER']).toLocaleString('es-AR', { minimumFractionDigits: 2 })}</span>` : '';
        const isPes = esSectorPesable(p);
        const pesBadge = isPes ? `<span class="sr-tag sr-tag--pesable">⚖ kg</span>` : '';
        return `
        <li class="sr-item" role="option" tabindex="0" data-idx="${i}">
            <div class="sr-main">
                <div class="sr-desc">${desc}</div>
                <div class="sr-meta"><span class="sr-prov">${prov}</span>${eanBadge}${intBadge}</div>
            </div>
            <div class="sr-right">
                <div class="sr-tags">${gramaje}${sect}${secc}${pesBadge}${pvp}</div>
                <svg class="sr-arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"
                     stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"/></svg>
            </div>
        </li>`;
    }).join('');

    els.searchResults.querySelectorAll('.sr-item').forEach((li, i) => {
        const p = shown[i];
        li.addEventListener('click', () => selectSearchProduct(p));
        li.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); selectSearchProduct(p); }
            if (e.key === 'ArrowDown') { e.preventDefault(); (li.nextElementSibling || li).focus(); }
            if (e.key === 'ArrowUp') { e.preventDefault(); (li.previousElementSibling || li).focus(); }
        });
    });
}

function selectSearchProduct(p) {
    els.scannedOk.classList.add('show');
    if (navigator.vibrate) navigator.vibrate([60, 20, 60]);
    fillProductData(p);
    closeSearchModal();
    setTimeout(() => {
        if (!els.fQty.value) els.fQty.focus();
        else if (!els.fExp.value) els.fExp.focus();
    }, 200);
}

// ─────────────────────────────────────────────
// BARCODE SCANNER
// ─────────────────────────────────────────────
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
        const config = {
            fps: 10,
            qrbox: { width: 260, height: 120 },
            formatsToSupport: [
                Html5QrcodeSupportedFormats.EAN_13, Html5QrcodeSupportedFormats.EAN_8,
                Html5QrcodeSupportedFormats.UPC_A, Html5QrcodeSupportedFormats.UPC_E,
                Html5QrcodeSupportedFormats.CODE_128, Html5QrcodeSupportedFormats.CODE_39,
                Html5QrcodeSupportedFormats.QR_CODE, Html5QrcodeSupportedFormats.ITF,
            ],
            aspectRatio: 4 / 3,
            disableFlip: false,
        };
        await state.html5QrCode.start({ facingMode: 'environment' }, config, onBarcodeScanned, null);
        state.scannerActive = true;
    } catch (err) {
        console.warn('Scanner error:', err);
        if (els.sHint) els.sHint.style.display = 'none';
        els.camErr.classList.add('show');
    }
}

function onBarcodeScanned(decodedText) {
    clearFieldError('fBarcode');
    els.scannedOk.classList.add('show');
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

// ─────────────────────────────────────────────
// PHOTO HANDLING
// ─────────────────────────────────────────────
function triggerFileInput(camera) {
    if (camera) els.fPhoto.setAttribute('capture', 'environment');
    else els.fPhoto.removeAttribute('capture');
    els.fPhoto.value = '';
    els.fPhoto.click();
}
function onPhotoZoneClick(e) {
    if (e.target.closest('.photo-preview')) return;
    if (e.target.closest('.photo-btns')) return;
    if (!state.photoBase64) triggerFileInput(false);
}
function onPhotoSelected(e) {
    const file = e.target.files[0];
    if (file) processPhoto(file);
}
function processPhoto(file) {
    if (!file.type.startsWith('image/')) { showToast('err', 'El archivo debe ser una imagen.'); return; }
    const MAX_MB = 5;
    if (file.size > MAX_MB * 1024 * 1024) { $('ePhoto').classList.add('show'); showToast('err', 'La imagen supera los 5 MB.'); return; }
    $('ePhoto').classList.remove('show');
    const reader = new FileReader();
    reader.onload = (ev) => {
        compressImage(ev.target.result, file.type, (compressedBase64, mime) => {
            state.photoBase64 = compressedBase64;
            state.photoMime = mime;
            state.photoName = file.name;
            els.photoImg.src = `data:${mime};base64,${compressedBase64}`;
            els.photoName.textContent = file.name;
            els.photoPreview.classList.add('show');
            els.photoPlaceholder.classList.add('hide');
        });
    };
    reader.readAsDataURL(file);
}
function compressImage(dataUrl, originalMime, callback) {
    const img = new Image();
    img.onload = () => {
        const MAX_DIM = 1200;
        let { width, height } = img;
        if (width > MAX_DIM || height > MAX_DIM) {
            if (width > height) { height = Math.round(height * MAX_DIM / width); width = MAX_DIM; }
            else { width = Math.round(width * MAX_DIM / height); height = MAX_DIM; }
        }
        const canvas = document.createElement('canvas');
        canvas.width = width; canvas.height = height;
        canvas.getContext('2d').drawImage(img, 0, 0, width, height);
        callback(canvas.toDataURL('image/jpeg', 0.75).split(',')[1], 'image/jpeg');
    };
    img.onerror = () => callback(dataUrl.split(',')[1], originalMime);
    img.src = dataUrl;
}
function removePhoto() {
    state.photoBase64 = null; state.photoMime = null; state.photoName = null;
    els.fPhoto.value = '';
    els.photoPreview.classList.remove('show');
    els.photoPlaceholder.classList.remove('hide');
    els.photoImg.src = '';
}

// ─────────────────────────────────────────────
// VALIDACIÓN
// ─────────────────────────────────────────────
const validationRules = {
    fEmail: { required: true, type: 'emailOrName' },
    fBarcode: { required: true },
    fQty: { required: true, min: 1, integer: true },
    fExp: { required: true },
    fBranch: { required: true },
    fEvent: { required: true },
    fDiscount: { required: false, min: 1, max: 99, integer: true },
};

const errorMap = {
    fEmail: 'eEmail', fBarcode: 'eBarcode', fQty: 'eQty',
    fExp: 'eExp', fBranch: 'eBranch', fEvent: 'eEvent', fDiscount: 'eDiscount',
};

function getTomorrowStr() {
    const d = new Date();
    d.setDate(d.getDate() + 1);
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

function validateField(fieldId) {
    if (fieldId === 'fQty') {
        validationRules.fQty = state.isPesable
            ? { required: true, min: 0.001 }
            : { required: true, min: 1, integer: true };
    }
    const rule = validationRules[fieldId];
    if (!rule) return true;
    const el = $(fieldId);
    if (!el) return true;
    const val = el.value.trim();
    let ok = true;

    if (rule.required && !val) ok = false;

    if (ok && rule.type === 'emailOrName' && val) {
        const isEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(val);
        const isName = val.trim().length >= 3;
        if (!isEmail && !isName) ok = false;
    }
    const numVal = Number(val.replace(',', '.'));
    if (ok && rule.min !== undefined && (isNaN(numVal) || numVal < rule.min)) ok = false;
    if (ok && rule.max !== undefined && numVal > rule.max) ok = false;
    if (ok && rule.integer && val && !Number.isInteger(numVal)) ok = false;

    if (fieldId === 'fExp' && ok && val && MOTIVOS_VENCIMIENTO.includes(els.fEvent.value)) {
        const minStr = getTomorrowStr();
        if (val < minStr) {
            ok = false;
            const errEl = $('eExp');
            if (errEl) errEl.textContent = 'Para Control de Vencimiento la fecha debe ser a partir de mañana.';
        }
    } else {
        const errEl = $('eExp');
        if (errEl && errEl.textContent.includes('mañana')) errEl.textContent = 'Seleccioná la fecha.';
    }

    if (fieldId === 'fBarcode') {
        const errEl = $('eBarcode');
        if (errEl) errEl.classList.toggle('show', !ok);
        return ok;
    }

    el.classList.toggle('err', !ok);
    const errEl = $(errorMap[fieldId]);
    if (errEl) errEl.classList.toggle('show', !ok);
    return ok;
}

function clearFieldError(fieldId) {
    const el = $(fieldId);
    if (!el) return;
    el.classList.remove('err');
    const errEl = $(errorMap[fieldId]);
    if (errEl) errEl.classList.remove('show');
}

function validateAll() {
    let allOk = true;

    validationRules.fQty = state.isPesable
        ? { required: true, min: 0.001 }
        : { required: true, min: 1, integer: true };

    for (const fieldId of Object.keys(validationRules)) {
        if (fieldId === 'fDiscount') {
            if (els.discountWrap.style.display !== 'none') {
                const val = els.fDiscount.value.trim();
                const num = Number(val);
                const ok = val !== '' && !isNaN(num) && Number.isInteger(num) && num >= 1 && num <= 99;
                els.fDiscount.classList.toggle('err', !ok);
                const errEl = $('eDiscount');
                if (errEl) errEl.classList.toggle('show', !ok);
                if (!ok) allOk = false;
            }
            continue;
        }
        if (!validateField(fieldId)) allOk = false;
    }
    return allOk;
}

// ─────────────────────────────────────────────
// CONSTRUIR PAYLOAD
// ─────────────────────────────────────────────
function buildPayload() {
    return {
        action: 'submitForm',
        usuario: els.fEmail.value.trim(),
        sucursal: els.fBranch.value,
        event: (() => {
            const base = els.fEvent.value;
            if (base === 'OTRO DESCUENTO' && els.fDiscount.value.trim()) {
                return `OTRO DESCUENTO ${els.fDiscount.value.trim()}% OFF`;
            }
            return base;
        })(),
        cantidad: els.fQty.value.trim().replace(',', '.'),
        esPesable: state.isPesable,
        unidadCantidad: state.isPesable ? 'kg' : 'unidades',
        fechaVenc: els.fExp.value,
        aclaracion: els.fNote.value.trim(),
        lote: els.fLot.value.trim(),
        comentarios: els.fComment.value.trim(),
        ean: els.fBarcode.value.trim(),
        descripcion: els.fDesc.value.trim(),
        codInterno: state.productData?.INTERNO || null,
        proveedor: state.productData?.PROVEEDOR || null,
        codProv: state.productData?.['COD.PROVEEDOR'] || null,
        gramaje: state.productData?.GRAMAJE || null,
        uxb: state.productData?.UXB || null,
        sector: state.productData?.SECTOR || null,
        seccion: state.productData?.SECCION || null,
        photoBase64: state.photoBase64 || null,
        photoMime: state.photoMime || null,
        photoName: state.photoName || null,
    };
}

// ─────────────────────────────────────────────
// ENVÍO DIRECTO — un producto por vez
// ─────────────────────────────────────────────
async function submitSingle() {
    if (state.submitting) return;

    if (!dentroDelHorario()) {
        showToast('err', mensajeHorario(), 7000);
        return;
    }

    if (!validateAll()) {
        showToast('err', 'Corregí los campos marcados antes de enviar.');
        const firstErr = document.querySelector('.ferr.show');
        if (firstErr) firstErr.scrollIntoView({ behavior: 'smooth', block: 'center' });
        return;
    }

    const payload = buildPayload();
    setLoading(true);

    try {
        if (APPS_SCRIPT_URL === 'PEGA_AQUI_TU_URL_DE_APPS_SCRIPT') {
            await delay(800);
            onSubmitSuccess();
        } else {
            const res = await fetch(APPS_SCRIPT_URL, {
                method: 'POST',
                body: JSON.stringify(payload),
            });
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const data = await res.json();
            if (data.success) onSubmitSuccess();
            else throw new Error(data.message || 'Error desconocido');
        }
    } catch (err) {
        console.error('submitSingle error:', err);
        showToast('err', 'Error al enviar. Revisá la conexión e intentá de nuevo.', 6000);
    } finally {
        setLoading(false);
    }
}

function onSubmitSuccess() {
    state.sentCount++;
    localStorage.setItem('devCount', state.sentCount);
    updateCounter();
    showToast('ok', '✓ Registro enviado correctamente', 5000);

    // Conservar email, sucursal y tipo de registro para el próximo envío
    const savedEmail = els.fEmail.value;
    const savedBranch = els.fBranch.value;
    const savedEvent = els.fEvent.value;
    const savedDiscount = els.fDiscount.value;

    clearForm();

    els.fEmail.value = savedEmail;
    els.fBranch.value = savedBranch;
    els.fEvent.value = savedEvent;
    els.fDiscount.value = savedDiscount;
    toggleBdBadge(savedEvent);
    toggleDiscountInput(savedEvent);
    updateSubmitButtonLabel(savedEvent);
    applyDateRestriction(savedEvent);

    // Scroll al tope del formulario
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ─────────────────────────────────────────────
// LIMPIAR FORMULARIO
// ─────────────────────────────────────────────
function clearForm() {
    document.querySelectorAll('input:not(#fPhoto):not(#extInput), select, textarea').forEach((el) => {
        if (el.type !== 'hidden') el.value = '';
        el.classList.remove('err');
    });
    els.fBarcode.value = '';
    els.fDesc.value = '';
    els.discountWrap.style.display = 'none';
    els.fDiscount.value = '';
    els.fExp.removeAttribute('min');
    els.fExp.removeAttribute('max');

    state.isPesable = false;
    state.pesoKg = null;
    els.fQty.readOnly = false;
    els.fQty.classList.remove('qty--pesable');
    els.fQty.setAttribute('step', '1');
    els.fQty.setAttribute('min', '1');
    if (els.qtyPesableHint) els.qtyPesableHint.style.display = 'none';
    const hintLabel = $('fQtyHint');
    if (hintLabel) hintLabel.textContent = '(enteros)';
    const eQtyEl = $('eQty');
    if (eQtyEl) eQtyEl.textContent = 'Número ≥ 1';

    document.querySelectorAll('.ferr').forEach((el) => el.classList.remove('show'));
    els.scannedOk.classList.remove('show');

    const badge = $('bdBadge');
    if (badge) badge.style.display = 'none';

    if (els.btnSubmitText) els.btnSubmitText.textContent = 'Enviar registro';

    if (els.extInput) {
        els.extInput.value = '';
        els.extInput.classList.remove('ext-input--reading', 'ext-input--found');
        els.btnExtClear.style.display = 'none';
        setExtStatus('ready', 'Listo — esperando lectura');
    }
    removePhoto();
    hideProductCard();
}

// ─────────────────────────────────────────────
// DISCOUNT INPUT TOGGLE
// ─────────────────────────────────────────────
function toggleDiscountInput(motivo) {
    const isDiscount = motivo === 'OTRO DESCUENTO';
    els.discountWrap.style.display = isDiscount ? 'block' : 'none';
    if (!isDiscount) { els.fDiscount.value = ''; clearFieldError('fDiscount'); }
    else setTimeout(() => els.fDiscount.focus(), 80);
}

// ─────────────────────────────────────────────
// RESTRICCIÓN DE FECHA
// ─────────────────────────────────────────────
function applyDateRestriction(motivo) {
    if (MOTIVOS_VENCIMIENTO.includes(motivo)) {
        const tomorrow = getTomorrowStr();
        els.fExp.setAttribute('min', tomorrow);
        els.fExp.removeAttribute('max');
        if (els.fExp.value && els.fExp.value < tomorrow) {
            els.fExp.value = '';
            showToast('wrn', 'La fecha de vencimiento debe ser a partir de mañana.');
        }
    } else {
        els.fExp.removeAttribute('min');
        els.fExp.removeAttribute('max');
    }
}

// ─────────────────────────────────────────────
// UI HELPERS
// ─────────────────────────────────────────────
function setLoading(on) {
    state.submitting = on;
    if (els.btnSubmit) els.btnSubmit.disabled = on;
    if (on) {
        if (typeof window.showSendingOverlay === 'function') window.showSendingOverlay();
    } else {
        if (typeof window.hideSendingOverlay === 'function') window.hideSendingOverlay();
    }
}

let toastTimer = null;
function showToast(type, msg, duration = 3500) {
    const { toast, toastMsg } = els;
    toastMsg.textContent = msg;
    toast.className = `toast ${type}`;
    void toast.offsetWidth;
    toast.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => toast.classList.remove('show'), duration);
}

function updateCounter() {
    if (state.sentCount > 0) {
        const label = state.sentCount === 1 ? 'registro enviado' : 'registros enviados';
        els.counterBadge.textContent = `${state.sentCount} ${label}`;
        els.counterBadge.classList.add('show');
    }
}

function delay(ms) { return new Promise((resolve) => setTimeout(resolve, ms)); }

// ─────────────────────────────────────────────
// SEARCH UTILS
// ─────────────────────────────────────────────
function normalize(str) {
    if (!str) return '';
    return String(str).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function highlight(text, query) {
    if (!text || !query) return esc(text || '');
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

function esc(str) {
    return String(str)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;')
        .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ─────────────────────────────────────────────
// RESTRICCIÓN HORARIA
// ─────────────────────────────────────────────
const HORARIO = {
    diasHabiles: [1, 2, 3, 4, 5], // 0=Dom, 1=Lun, 2=Mar, 3=Mié, 4=Jue, 5=Vie, 6=Sáb
    inicio: { h: 9, m: 0 },
    fin: { h: 15, m: 0 },
};

function dentroDelHorario() {
    const now = new Date();
    const dia = now.getDay();
    const hoy = now.getHours() * 60 + now.getMinutes();
    const ini = HORARIO.inicio.h * 60 + HORARIO.inicio.m;
    const fin = HORARIO.fin.h * 60 + HORARIO.fin.m;
    return HORARIO.diasHabiles.includes(dia) && hoy >= ini && hoy < fin;
}

function mensajeHorario() {
    const pad = n => String(n).padStart(2, '0');
    return `El formulario está disponible solo de lunes a viernes, `
        + `de ${pad(HORARIO.inicio.h)}:${pad(HORARIO.inicio.m)} `
        + `a ${pad(HORARIO.fin.h)}:${pad(HORARIO.fin.m)} hs.`;
}