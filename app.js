// ── Estado ──
// ── Estado ──
const state = {
  order: [],
  currentProduct: null,
  lookupTimer: null,
  scannerActive: false,
  html5QrCode: null,
  productos: [],      // 👈 nuevo
};

// ── DOM ──
const $ = (id) => document.getElementById(id);

const els = {
  fConcepto: $('fConcepto'),
  eConcepto: $('eConcepto'),
  fSucursal: $('fSucursal'),
  eSucursal: $('eSucursal'),
  fBarcode: $('fBarcode'),
  scannedOk: $('scannedOk'),
  productCard: $('productCard'),
  pcStateLabel: $('pcStateLabel'),
  pcDescripcion: $('pcDescripcion'),
  pcProveedor: $('pcProveedor'),
  pcGramaje: $('pcGramaje'),
  pcUxb: $('pcUxb'),
  fDesc: $('fDesc'),
  eDesc: $('eDesc'),
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
  scanOverlay: $('scanOverlay'),
  btnOpenScanner: $('btnOpenScanner'),
  btnCloseScanner: $('btnCloseScanner'),
  btnManualEntry: $('btnManualEntry'),
  camErr: $('camErr'),
  sHint: $('sHint'),
};

// ── Init ──
async function init() {
  await cargarProductos();  
  els.fConcepto.addEventListener('input', () => {
    els.eConcepto.classList.remove('show');
  });
  els.btnExportPDF = $('btnExportPDF');
  els.btnExportPDF.addEventListener('click', exportPDF);
  els.btnOpenScanner.addEventListener('click', openScanner);
  els.btnCloseScanner.addEventListener('click', closeScanner);
  els.btnManualEntry.addEventListener('click', closeScanner);

  els.fSucursal.addEventListener('change', () => {
    els.eSucursal.classList.remove('show');
    els.fSucursal.classList.remove('err');
  });

  els.fBarcode.addEventListener('input', () => {
    els.scannedOk.classList.remove('show');
    hideProductCard();
    resetCurrentProduct();
    clearTimeout(state.lookupTimer);
    const code = els.fBarcode.value.trim();
    if (!code) { updateAddBtn(); return; }
    state.lookupTimer = setTimeout(() => triggerLookup(code), 600);
  });

  els.fBarcode.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      clearTimeout(state.lookupTimer);
      const code = els.fBarcode.value.trim();
      if (code) triggerLookup(code);
    }
  });

  els.fDesc.addEventListener('input', () => {
    els.eDesc.classList.remove('show');
    updateAddBtn();
  });

  els.fQty.addEventListener('input', () => els.eQty.classList.remove('show'));

  els.fQty.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); addToOrder(); }
  });

  els.btnAdd.addEventListener('click', addToOrder);
  els.btnExport.addEventListener('click', exportExcel);
  els.btnClearOrder.addEventListener('click', clearOrder);
}

document.addEventListener('DOMContentLoaded', init);

// ── Product lookup ──
async function triggerLookup(code) {
  showCardLoading();
  const data = await fetchProducto(code);
  if (data) {
    state.currentProduct = data;
    fillProductCard(data);
    showCardFound();
    showToast('ok', 'Producto encontrado');
    els.fQty.focus();
  } else {
    state.currentProduct = null;
    showCardNotFound();
    els.fDesc.value = '';
    els.fDesc.focus();
    showToast('wrn', 'Producto no encontrado — ingresá la descripción.');
  }
  updateAddBtn();
}

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

async function fetchProducto(code) {
  const codigo = parseInt(code);
  const producto = state.productos.find(p =>
    p.EAN === codigo ||
    p.INTERNO === codigo ||
    String(p.EAN) === code
  );
  return producto || null;
}

function fillProductCard(data) {
  els.pcDescripcion.textContent = data.DESCRIPCION || '-';
  els.pcProveedor.textContent = data.PROVEEDOR || '-';
  els.pcGramaje.textContent = data.GRAMAJE || '-';
  els.pcUxb.textContent = data.UXB != null ? `${data.UXB} u.` : '-';
  els.fDesc.value = data.DESCRIPCION || '';
  els.eDesc.classList.remove('show');
}

function resetCurrentProduct() {
  state.currentProduct = null;
  els.fDesc.value = '';
  ['pcDescripcion', 'pcProveedor', 'pcGramaje', 'pcUxb'].forEach(k => els[k].textContent = '—');
}

function showCardLoading() { els.productCard.classList.add('visible'); els.productCard.dataset.state = 'loading'; els.pcStateLabel.textContent = 'Buscando...'; }
function showCardFound() { els.productCard.classList.add('visible'); els.productCard.dataset.state = 'found'; els.pcStateLabel.textContent = 'Producto encontrado'; }
function showCardNotFound() { els.productCard.classList.add('visible'); els.productCard.dataset.state = 'notfound'; els.pcStateLabel.textContent = 'No encontrado'; }
function hideProductCard() { els.productCard.classList.remove('visible'); els.productCard.dataset.state = ''; }

function updateAddBtn() {
  els.btnAdd.disabled = !(els.fDesc.value.trim() && els.fBarcode.value.trim());
}

// ── Order management ──
function addToOrder() {
  const desc = els.fDesc.value.trim();
  const code = els.fBarcode.value.trim();
  const qty = parseFloat(els.fQty.value.replace(',', '.'));
  let valid = true;

  if (!desc) { els.eDesc.classList.add('show'); valid = false; }
  if (isNaN(qty) || qty <= 0) {
    els.eQty.classList.add('show');
    if (valid) els.fQty.focus();
    valid = false;
  }
  if (!valid) return;

  const p = state.currentProduct;
  state.order.push({
    id: Date.now(),
    ean: p?.EAN || code,
    interno: p?.INTERNO || '',
    descripcion: desc,
    proveedor: p?.PROVEEDOR || '',
    gramaje: p?.GRAMAJE || '',
    uxb: p?.UXB != null ? p.UXB : '',
    cantidad: qty,
  });

  renderTable();
  updateBadge();
  showToast('ok', `"${desc}" agregado (×${qty})`);
  resetForm();
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

function resetForm() {
  els.fBarcode.value = '';
  els.fDesc.value = '';
  els.fQty.value = '';
  els.scannedOk.classList.remove('show');
  hideProductCard();
  resetCurrentProduct();
  updateAddBtn();
  els.fBarcode.focus();
}

// ── Render table ──
function renderTable() {
  const has = state.order.length > 0;
  els.tableEmpty.style.display = has ? 'none' : 'block';
  els.orderTable.style.display = has ? 'table' : 'none';
  els.btnExport.disabled = !has;
  els.btnExportPDF.disabled = !has;
  els.btnClearOrder.disabled = !has;

  els.orderBody.innerHTML = state.order.map(item => `
    <tr>
      <td class="td-code">${esc(item.ean || '—')}</td>
      <td class="td-desc">${esc(item.descripcion)}</td>
      <td class="td-prov">${esc(String(item.gramaje)) || '—'}</td>
      <td class="td-prov">${item.uxb !== '' ? item.uxb + ' u.' : '—'}</td>
      <td class="td-qty">${item.cantidad}</td>
      <td class="td-actions"><button class="btn-remove" onclick="removeItem(${item.id})">✕</button></td>
    </tr>
  `).join('');
}

function updateBadge() {
  const n = state.order.length;
  els.badgeCount.textContent = `${n} ${n === 1 ? 'ítem' : 'ítems'}`;
  els.badgeCount.classList.toggle('show', n > 0);
}

// ── Export Excel ──
async function exportExcel() {

  const now = new Date();
  const fecha = now.toLocaleDateString('es-AR');
  const fechaFile = now.toISOString().slice(0, 10);
  const remitoId = now.getTime();
  if (!state.order.length) return;

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
    paperSize: 9,
    orientation: 'portrait',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: {
      left: 0.3, right: 0.3,
      top: 0.4, bottom: 0.4,
      header: 0.3, footer: 0.3,
    }
  };

  ws.pageSetup.printTitlesRow = '1:3';

  ws.mergeCells('A1:F1');
  ws.getCell('A1').value = `REMITO/PEDIDO — ${sucursal.toUpperCase()} — ${fecha}`;
  ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell('A1').font = { bold: true, size: 14 };

  ws.mergeCells('A2:F2');
  ws.getCell('A2').value = `Concepto: ${concepto}`;
  ws.getCell('A2').alignment = { horizontal: 'center' };

  ws.mergeCells('A3:F3');
  ws.getCell('A3').value = `ID: ${remitoId}`;
  ws.getCell('A3').alignment = { horizontal: 'center' };

  const headerRowIndex = 4;
  ws.getRow(headerRowIndex).values = [
    'COD-BAR',
    'EAN',
    'DESCRIPCION',
    'GRAMAJE',
    'UxB',
    'CANTIDAD'
  ];
  ws.getRow(headerRowIndex).font = { bold: true };

  ws.columns = [
    { key: 'img', width: 18 },
    { key: 'ean', width: 15 },
    { key: 'desc', width: 50 },
    { key: 'gram', width: 15 },
    { key: 'uxb', width: 5 },
    { key: 'cant', width: 8 }
  ];

  function barcodeDataUrl(code, opts = {}) {
    const canvas = document.createElement('canvas');
    const format = (/^\d{13}$/.test(code) ? 'ean13' : 'code128');
    try {
      JsBarcode(canvas, String(code), {
        format,
        displayValue: false,
        margin: 0,
        width: opts.width || 2,
        height: opts.height || 60,
        fontSize: 12
      });
    } catch (err) {
      JsBarcode(canvas, String(code), {
        format: 'code128',
        displayValue: false,
        margin: 0,
        width: opts.width || 2,
        height: opts.height || 60
      });
    }
    return canvas.toDataURL('image/png');
  }

  const startRow = headerRowIndex + 1;

  for (let i = 0; i < state.order.length; i++) {
    const item = state.order[i];
    const rowIndex = startRow + i;
    const eanCode = item.ean || '';

    const dataUrl = barcodeDataUrl(eanCode, { width: 2, height: 60 });
    const base64 = dataUrl.split(',')[1];

    const imageId = workbook.addImage({
      base64: base64,
      extension: 'png'
    });

    ws.getCell(`A${rowIndex}`).value = '';

    ws.addImage(imageId, {
      tl: { col: 0.15, row: rowIndex - 1 + 0.1 },
      ext: { width: 100, height: 45 },
      editAs: 'oneCell'
    });

    ws.getCell(`B${rowIndex}`).value = String(eanCode);
    ws.getCell(`C${rowIndex}`).value = item.descripcion || '';
    ws.getCell(`D${rowIndex}`).value = item.gramaje || '';
    ws.getCell(`E${rowIndex}`).value = item.uxb !== '' ? item.uxb : '';
    ws.getCell(`F${rowIndex}`).value = item.cantidad;

    ws.getRow(rowIndex).height = 55;
  }

  ws.eachRow({ includeEmpty: true }, row => {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true
      };
    });
  });

  ws.eachRow({ includeEmpty: true }, row => {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
  });

  autoFitColumns(ws, [3]);

  try {
    const buf = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/octet-stream' });
    saveAs(blob, filename);
    showToast('ok', 'Excel descargado con códigos de barra.');
    await uploadToDrive(blob, filename);
  } catch (err) {
    console.error('Error al generar Excel con imágenes:', err);
    showToast('err', 'No se pudo generar el Excel con imágenes.');
  }
}

function autoFitColumns(worksheet, cols = []) {
  cols.forEach(colNumber => {
    const column = worksheet.getColumn(colNumber);
    let maxLength = 10;

    column.eachCell({ includeEmpty: true }, (cell) => {
      const val = cell.value ? cell.value.toString() : "";
      maxLength = Math.max(maxLength, val.length + 2);
    });

    column.width = maxLength;
  });
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
  els.fBarcode.value = decodedText;
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

async function exportPDF() {

  if (!state.order.length) return;

  const now = new Date();
  const fecha = now.toLocaleDateString('es-AR');
  const fechaFile = now.toISOString().slice(0, 10);
  const remitoId = now.getTime();

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
  const filename = `${sucursal}_${fechaFile}_${conceptoFile}_${remitoId}.pdf`;

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    unit: 'pt',
    format: 'a4',
    orientation: 'landscape'
  });

  doc.setFontSize(16);
  doc.text(`REMITO/PEDIDO — ${sucursal.toUpperCase()} — ${fecha}`, 40, 40);

  doc.setFontSize(12);
  doc.text(`Concepto: ${concepto}`, 40, 60);
  doc.text(`ID: ${remitoId}`, 40, 80);

  const rows = [];

  for (const item of state.order) {
    const canvas = document.createElement('canvas');
    const format = /^\d{13}$/.test(item.ean) ? 'ean13' : 'code128';

    JsBarcode(canvas, String(item.ean), {
      format,
      displayValue: false,
      height: 40
    });

    rows.push([
      { content: '', barcode: canvas.toDataURL('image/png') },
      item.ean,
      item.descripcion,
      item.proveedor,
      item.interno,
      item.gramaje,
      item.uxb,
      item.cantidad
    ]);
  }

  doc.autoTable({
    startY: 110,
    head: [['Código', 'EAN', 'Descripción', 'Proveedor', 'Interno', 'Gramaje', 'UxB', 'Cant.']],
    body: rows,
    rowHeight: 30,
    styles: {
      halign: 'center',
      valign: 'middle',
      cellPadding: { top: 6, bottom: 6 }
    },

    didDrawCell: (data) => {
      if (data.column.index === 0 && data.cell.raw?.barcode) {
        const imgWidth = 110;
        const imgHeight = 30;

        const x = data.cell.x + (data.cell.width - imgWidth) / 2;
        const y = data.cell.y + (data.cell.height - imgHeight) / 2;

        doc.addImage(
          data.cell.raw.barcode,
          'PNG',
          x,
          y,
          imgWidth,
          imgHeight
        );
      }
    },

    columnStyles: {
      0: { cellWidth: 120 },
      1: { cellWidth: 70 },
      2: { cellWidth: 220 },
      3: { cellWidth: 120 },
      4: { cellWidth: 80 },
      5: { cellWidth: 70 },
      6: { cellWidth: 70 },
      7: { cellWidth: 60 }
    }
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
    reader.onloadend = () => {
      const base64data = reader.result.split(',')[1];
      resolve(base64data);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

async function uploadToDrive(blob, filename) {
  const url = "https://script.google.com/macros/s/AKfycbzIRnuzUWTjG38fxttQbkvJ7Br_wYSYs5UeaJO9EHnncy7jr8vQivcfQLot0xqSzwsq/exec";

  try {
    const base64Data = await blobToBase64(blob);

    const payload = {
      filename: filename,
      mimeType: blob.type,
      file: base64Data
    };

    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "text/plain;charset=utf-8",
      },
      body: JSON.stringify(payload)
    });

    const json = await res.json();

    if (json.ok) {
      showToast('ok', `Archivo subido a Drive: ${filename}`);
    } else {
      showToast('err', "Error al subir a Drive: " + (json.error || "Desconocido"));
    }
  } catch (err) {
    console.error("Error subida Drive:", err);
    showToast('err', "Error al conectar con Google Drive. Revisá la consola.");
  }
}