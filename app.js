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
};

// ── DOM ──
const $ = (id) => document.getElementById(id);

const els = {
  fConcepto:      $('fConcepto'),
  eConcepto:      $('eConcepto'),
  fSucursal:      $('fSucursal'),
  eSucursal:      $('eSucursal'),
  productCard:    $('productCard'),
  pcStateLabel:   $('pcStateLabel'),
  pcDescripcion:  $('pcDescripcion'),
  pcProveedor:    $('pcProveedor'),
  pcGramaje:      $('pcGramaje'),
  pcUxb:          $('pcUxb'),
  btnClearProduct:$('btnClearProduct'),
  qtyRow:         $('qtyRow'),
  fQty:           $('fQty'),
  eQty:           $('eQty'),
  btnAdd:         $('btnAdd'),
  orderTable:     $('orderTable'),
  orderBody:      $('orderBody'),
  tableEmpty:     $('tableEmpty'),
  btnExport:      $('btnExport'),
  btnClearOrder:  $('btnClearOrder'),
  badgeCount:     $('badgeCount'),
  toast:          $('toast'),
  toastMsg:       $('toastMsg'),
  toastDot:       $('toastDot'),
  // scanner
  scanOverlay:    $('scanOverlay'),
  btnOpenScanner: $('btnOpenScanner'),
  btnCloseScanner:$('btnCloseScanner'),
  btnManualEntry: $('btnManualEntry'),
  camErr:         $('camErr'),
  sHint:          $('sHint'),
  // search modal
  searchOverlay:       $('searchOverlay'),
  btnOpenSearch:       $('btnOpenSearch'),
  btnCloseSearch:      $('btnCloseSearch'),
  searchInput:         $('searchInput'),
  btnClearSearch:      $('btnClearSearch'),
  searchResults:       $('searchResults'),
  searchStateEmpty:    $('searchStateEmpty'),
  searchStateNone:     $('searchStateNone'),
  searchTermDisplay:   $('searchTermDisplay'),
  searchCount:         $('searchCount'),
};

// ── Init ──
async function init() {
  await cargarProductos();

  els.btnExportPDF = $('btnExportPDF');
  els.btnExportPDF.addEventListener('click', exportPDF);

  els.fConcepto.addEventListener('input', () => els.eConcepto.classList.remove('show'));
  els.fSucursal.addEventListener('change', () => {
    els.eSucursal.classList.remove('show');
    els.fSucursal.classList.remove('err');
  });

  // Producto
  els.btnClearProduct.addEventListener('click', resetCurrentProduct);

  // Cantidad
  els.fQty.addEventListener('input', () => els.eQty.classList.remove('show'));
  els.fQty.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); addToOrder(); }
  });
  els.btnAdd.addEventListener('click', addToOrder);

  // Scanner
  els.btnOpenScanner.addEventListener('click', openScanner);
  els.btnCloseScanner.addEventListener('click', closeScanner);
  els.btnManualEntry.addEventListener('click', closeScanner);

  // Search modal
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

  // Filter chips
  document.querySelectorAll('.sf-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      document.querySelectorAll('.sf-chip').forEach(c => c.classList.remove('sf-chip--active'));
      chip.classList.add('sf-chip--active');
      state.searchField = chip.dataset.field;
      onSearchInput();
    });
  });

  // Cerrar modales con Escape
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if (els.searchOverlay.classList.contains('open')) closeSearchModal();
      if (els.scanOverlay.classList.contains('open')) closeScanner();
    }
  });

  els.btnExport.addEventListener('click', exportExcel);
  els.btnClearOrder.addEventListener('click', clearOrder);
}

document.addEventListener('DOMContentLoaded', init);

// ── Carga de productos ──
async function cargarProductos() {
  try {
    const res = await fetch('bd/productos.json');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    state.productos = json.data || [];
  } catch (err) {
    console.error('Error al cargar productos.json:', err);
    showToast('err', 'No se pudo cargar la base de productos.');
  }
}

// ── Búsqueda por código (EAN / INTERNO) ──
async function fetchProducto(code) {
  const codigo = parseInt(code);
  return state.productos.find(p =>
    p.EAN === codigo ||
    p.INTERNO === codigo ||
    String(p.EAN) === code
  ) || null;
}

// ── Lookup desde scanner ──
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

// ── Producto actual ──
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
}

function resetCurrentProduct() {
  state.currentProduct = null;
  hideProductCard();
  hideQtyRow();
  els.fQty.value = '';
}

function showCardLoading()  {
  els.productCard.classList.add('visible');
  els.productCard.dataset.state = 'loading';
  els.pcStateLabel.textContent = 'Buscando...';
  els.btnClearProduct.style.display = 'none';
}
function showCardFound()    {
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
function hideProductCard()  {
  els.productCard.classList.remove('visible');
  els.productCard.dataset.state = '';
  els.btnClearProduct.style.display = 'none';
}

function showQtyRow() { els.qtyRow.style.display = 'flex'; }
function hideQtyRow() { els.qtyRow.style.display = 'none'; }

// ══════════════════════════════════════════════
// ── Search Modal ──
// ══════════════════════════════════════════════

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
      return (
        normalize(p.DESCRIPCION).includes(q) ||
        normalize(p.PROVEEDOR).includes(q)   ||
        normalize(p.SECTOR).includes(q)      ||
        normalize(p.SECCION).includes(q)     ||
        normalize(String(p.EAN    ?? '')).includes(q) ||
        normalize(String(p.INTERNO ?? '')).includes(q)
      );
    }
    // Campo EAN: busca tanto EAN como INTERNO
    if (field === 'EAN') {
      return (
        normalize(String(p.EAN     ?? '')).includes(q) ||
        normalize(String(p.INTERNO ?? '')).includes(q)
      );
    }
    return normalize(String(p[field] ?? '')).includes(q);
  });

  els.searchStateEmpty.style.display = 'none';

  if (matches.length === 0) {
    els.searchStateNone.style.display = 'flex';
    els.searchTermDisplay.textContent = `"${query}"`;
    els.searchResults.style.display   = 'none';
    els.searchCount.textContent       = 'Sin resultados';
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
    const desc    = highlight(p.DESCRIPCION || '—', q);
    const prov    = highlight(p.PROVEEDOR   || '—', q);
    const eanRaw  = String(p.EAN ?? '');
    const eanHL   = highlight(eanRaw, q);
    const intRaw  = String(p.INTERNO ?? '');
    const intHL   = highlight(intRaw, q);
    const sect    = p.SECTOR  ? `<span class="sr-tag">${esc(p.SECTOR)}</span>`  : '';
    const secc    = p.SECCION ? `<span class="sr-tag">${esc(p.SECCION)}</span>` : '';
    const gramaje = p.GRAMAJE ? `<span class="sr-tag sr-tag--gramaje">${esc(p.GRAMAJE)}</span>` : '';

    const eanBadge = eanRaw
      ? `<span class="sr-ean">EAN ${eanHL}</span>`
      : '';
    const intBadge = intRaw && intRaw !== eanRaw
      ? `<span class="sr-ean sr-ean--int">INT ${intHL}</span>`
      : '';

    return `
      <li class="sr-item" role="option" tabindex="0" data-idx="${i}">
        <div class="sr-main">
          <div class="sr-desc">${desc}</div>
          <div class="sr-meta">
            <span class="sr-prov">${prov}</span>
            ${eanBadge}${intBadge}
          </div>
        </div>
        <div class="sr-tags">${gramaje}${sect}${secc}</div>
        <div class="sr-arrow">›</div>
      </li>`;
  }).join('');

  els.searchResults.querySelectorAll('.sr-item').forEach((li, i) => {
    const selectItem = () => selectSearchResult(shown[i]);
    li.addEventListener('click', selectItem);
    li.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); selectItem(); }
      if (e.key === 'ArrowDown') { e.preventDefault(); (li.nextElementSibling || li).focus(); }
      if (e.key === 'ArrowUp')   { e.preventDefault(); (li.previousElementSibling || li).focus(); }
    });
  });
}

function selectSearchResult(producto) {
  closeSearchModal();
  setCurrentProduct(producto);
  showToast('ok', 'Producto seleccionado');
  if (navigator.vibrate) navigator.vibrate([60, 20, 60]);
  setTimeout(() => els.fQty.focus(), 150);
}

// ── Helpers ──
function normalize(str) {
  if (!str) return '';
  return String(str).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function highlight(text, query) {
  if (!text || !query) return esc(text);
  const normText  = normalize(text);
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

// ── Gestión del pedido ──
function addToOrder() {
  const p   = state.currentProduct;
  const qty = parseFloat(els.fQty.value.replace(',', '.'));

  if (isNaN(qty) || qty <= 0) {
    els.eQty.classList.add('show');
    els.fQty.focus();
    return;
  }

  const code = p?.EAN || p?.INTERNO || '';
  state.order.push({
    id:          Date.now(),
    ean:         p?.EAN        || code,
    interno:     p?.INTERNO    || '',
    descripcion: p?.DESCRIPCION || '',
    proveedor:   p?.PROVEEDOR   || '',
    gramaje:     p?.GRAMAJE     || '',
    uxb:         p?.UXB != null ? p.UXB : '',
    cantidad:    qty,
  });

  renderTable();
  updateBadge();
  showToast('ok', `"${p?.DESCRIPCION || code}" agregado (×${qty})`);
  resetCurrentProduct();
  els.fQty.value = '';
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

// ── Render tabla ──
function renderTable() {
  const has = state.order.length > 0;
  els.tableEmpty.style.display  = has ? 'none'  : 'block';
  els.orderTable.style.display  = has ? 'table' : 'none';
  els.btnExport.disabled        = !has;
  els.btnExportPDF.disabled     = !has;
  els.btnClearOrder.disabled    = !has;

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
    try { if (state.scannerActive) await state.html5QrCode.stop(); state.html5QrCode.clear(); } catch (_) {}
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
async function exportExcel() {
  if (!state.order.length) return;

  const now       = new Date();
  const fecha     = now.toLocaleDateString('es-AR');
  const fechaFile = now.toISOString().slice(0, 10);
  const remitoId  = now.getTime();

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

  const conceptoFile = concepto.replace(/\s+/g, '_').toLowerCase();
  const filename = `${sucursal}_${fechaFile}_${conceptoFile}_${remitoId}.xlsx`;

  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Pedido', { views: [{ state: 'normal' }] });

  ws.pageSetup = {
    paperSize: 9, orientation: 'portrait',
    fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    margins: { left:0.3, right:0.3, top:0.4, bottom:0.4, header:0.3, footer:0.3 }
  };
  ws.pageSetup.printTitlesRow = '1:3';

  ws.mergeCells('A1:F1');
  ws.getCell('A1').value = `REMITO/PEDIDO — ${sucursal.toUpperCase()} — ${fecha}`;
  ws.getCell('A1').alignment = { horizontal:'center', vertical:'middle' };
  ws.getCell('A1').font = { bold:true, size:14 };

  ws.mergeCells('A2:F2');
  ws.getCell('A2').value = `Concepto: ${concepto}`;
  ws.getCell('A2').alignment = { horizontal:'center' };

  ws.mergeCells('A3:F3');
  ws.getCell('A3').value = `ID: ${remitoId}`;
  ws.getCell('A3').alignment = { horizontal:'center' };

  ws.getRow(4).values = ['COD-BAR','EAN','DESCRIPCION','GRAMAJE','UxB','CANTIDAD'];
  ws.getRow(4).font = { bold:true };
  ws.columns = [
    { key:'img',  width:18 }, { key:'ean',  width:15 },
    { key:'desc', width:50 }, { key:'gram', width:15 },
    { key:'uxb',  width:5  }, { key:'cant', width:8  }
  ];

  function barcodeDataUrl(code, opts = {}) {
    const canvas = document.createElement('canvas');
    const format = /^\d{13}$/.test(code) ? 'ean13' : 'code128';
    try {
      JsBarcode(canvas, String(code), { format, displayValue:false, margin:0, width:opts.width||2, height:opts.height||60, fontSize:12 });
    } catch {
      JsBarcode(canvas, String(code), { format:'code128', displayValue:false, margin:0, width:opts.width||2, height:opts.height||60 });
    }
    return canvas.toDataURL('image/png');
  }

  for (let i = 0; i < state.order.length; i++) {
    const item     = state.order[i];
    const rowIndex = 5 + i;
    const eanCode  = item.ean || '';
    const base64   = barcodeDataUrl(eanCode, { width:2, height:60 }).split(',')[1];
    const imageId  = workbook.addImage({ base64, extension:'png' });

    ws.getCell(`A${rowIndex}`).value = '';
    ws.addImage(imageId, { tl:{ col:0.15, row:rowIndex-1+0.1 }, ext:{ width:100, height:45 }, editAs:'oneCell' });
    ws.getCell(`B${rowIndex}`).value = String(eanCode);
    ws.getCell(`C${rowIndex}`).value = item.descripcion || '';
    ws.getCell(`D${rowIndex}`).value = item.gramaje || '';
    ws.getCell(`E${rowIndex}`).value = item.uxb !== '' ? item.uxb : '';
    ws.getCell(`F${rowIndex}`).value = item.cantidad;
    ws.getRow(rowIndex).height = 55;
  }

  ws.eachRow({ includeEmpty:true }, row => {
    row.eachCell({ includeEmpty:true }, cell => {
      cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
      cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    });
  });

  autoFitColumns(ws, [3]);

  try {
    const buf  = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buf], { type:'application/octet-stream' });
    saveAs(blob, filename);
    showToast('ok', 'Excel descargado con códigos de barra.');
    await uploadToDrive(blob, filename);
  } catch (err) {
    console.error('Error Excel:', err);
    showToast('err', 'No se pudo generar el Excel.');
  }
}

function autoFitColumns(worksheet, cols = []) {
  cols.forEach(n => {
    const col = worksheet.getColumn(n);
    let max = 10;
    col.eachCell({ includeEmpty:true }, c => {
      const v = c.value ? c.value.toString() : '';
      max = Math.max(max, v.length + 2);
    });
    col.width = max;
  });
}

// ── Export PDF ──
async function exportPDF() {
  if (!state.order.length) return;

  const now       = new Date();
  const fecha     = now.toLocaleDateString('es-AR');
  const fechaFile = now.toISOString().slice(0, 10);
  const remitoId  = now.getTime();

  const sucursal = els.fSucursal.value;
  if (!sucursal) {
    els.eSucursal.classList.add('show'); els.fSucursal.classList.add('err'); els.fSucursal.focus();
    showToast('err', 'Seleccioná una sucursal antes de exportar.'); return;
  }
  const concepto = els.fConcepto.value.trim();
  if (!concepto) {
    els.eConcepto.classList.add('show'); els.fConcepto.classList.add('err'); els.fConcepto.focus();
    showToast('err', 'Ingresá un concepto general.'); return;
  }

  const conceptoFile = concepto.replace(/\s+/g, '_').toLowerCase();
  const filename = `${sucursal}_${fechaFile}_${conceptoFile}_${remitoId}.pdf`;

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit:'pt', format:'a4', orientation:'landscape' });

  doc.setFontSize(16); doc.text(`REMITO/PEDIDO — ${sucursal.toUpperCase()} — ${fecha}`, 40, 40);
  doc.setFontSize(12); doc.text(`Concepto: ${concepto}`, 40, 60); doc.text(`ID: ${remitoId}`, 40, 80);

  const rows = [];
  for (const item of state.order) {
    const canvas = document.createElement('canvas');
    const format = /^\d{13}$/.test(item.ean) ? 'ean13' : 'code128';
    JsBarcode(canvas, String(item.ean), { format, displayValue:false, height:40 });
    rows.push([
      { content:'', barcode: canvas.toDataURL('image/png') },
      item.ean, item.descripcion, item.proveedor, item.interno, item.gramaje, item.uxb, item.cantidad
    ]);
  }

  doc.autoTable({
    startY: 110,
    head: [['Código','EAN','Descripción','Proveedor','Interno','Gramaje','UxB','Cant.']],
    body: rows, rowHeight: 30,
    styles: { halign:'center', valign:'middle', cellPadding:{ top:6, bottom:6 } },
    didDrawCell: (data) => {
      if (data.column.index === 0 && data.cell.raw?.barcode) {
        const iw = 110, ih = 30;
        doc.addImage(data.cell.raw.barcode, 'PNG',
          data.cell.x + (data.cell.width - iw) / 2,
          data.cell.y + (data.cell.height - ih) / 2,
          iw, ih);
      }
    },
    columnStyles: { 0:{cellWidth:120}, 1:{cellWidth:70}, 2:{cellWidth:220}, 3:{cellWidth:120}, 4:{cellWidth:80}, 5:{cellWidth:70}, 6:{cellWidth:70}, 7:{cellWidth:60} }
  });

  const pdfBlob = doc.output('blob');
  doc.save(filename);
  await uploadToDrive(pdfBlob, filename);
  showToast('ok', 'PDF generado y subido a Drive.');
}

// ── Drive upload ──
function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result.split(',')[1]);
    reader.onerror  = reject;
    reader.readAsDataURL(blob);
  });
}

async function uploadToDrive(blob, filename) {
  const url = "https://script.google.com/macros/s/AKfycbzIRnuzUWTjG38fxttQbkvJ7Br_wYSYs5UeaJO9EHnncy7jr8vQivcfQLot0xqSzwsq/exec";
  try {
    const base64Data = await blobToBase64(blob);
    const res  = await fetch(url, {
      method:'POST',
      headers:{ 'Content-Type':'text/plain;charset=utf-8' },
      body: JSON.stringify({ filename, mimeType: blob.type, file: base64Data })
    });
    const json = await res.json();
    if (json.ok) showToast('ok', `Archivo subido a Drive: ${filename}`);
    else showToast('err', 'Error al subir a Drive: ' + (json.error || 'Desconocido'));
  } catch (err) {
    console.error('Error Drive:', err);
    showToast('err', 'Error al conectar con Google Drive.');
  }
}