// ══════════════════════════════════════════════════════════
//  STATE
// ══════════════════════════════════════════════════════════
let allData = [];
let filteredOp = [];
let filteredNc = [];
let currentPreset = 'today';
let dateFrom = null, dateTo = null;
let opPage = 1, ncPage = 1;
const PAGE_SIZE = 50;
let selectedIds = new Set();
let modalCallback = null;
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwj0qEOm9THbYxw0TYek2Oot3dlL1wn7YmPLtYknFzrGBQJXFnd-kh7yxXtFgYFyC-B/exec';

// ── Auto-refresh ──
const AUTO_REFRESH_MS = 30 * 60 * 1000; // 30 minutos
let _autoRefreshTimer = null;
let _countdownTimer  = null;
let _nextRefreshAt   = null;

// ══════════════════════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════════════════════
window.onload = () => {
  showPanel('operaciones');
  loadData();
};

// ══════════════════════════════════════════════════════════
//  NAV
// ══════════════════════════════════════════════════════════
const PANEL_TITLES = {
  config: '⚙️ Configuración',
  operaciones: '📋 Panel General',
  compras: '📊 Métricas',
  nc: '🧾 Gestión N/C',
};

function showPanel(id) {
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById('panel-' + id).classList.add('active');
  const btn = Array.from(document.querySelectorAll('.nav-item')).find(b => b.textContent.trim().includes(PANEL_TITLES[id]?.replace(/^[^\s]+\s/, '') || ''));
  if (btn) btn.classList.add('active');
  document.getElementById('topbarTitle').textContent = PANEL_TITLES[id] || id;
  closeSidebar();
}

function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('backdrop').classList.toggle('open');
}
function closeSidebar() {
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('backdrop').classList.remove('open');
}

// ══════════════════════════════════════════════════════════
//  AUTO-REFRESH
// ══════════════════════════════════════════════════════════
function startAutoRefresh() {
  stopAutoRefresh(); // limpia cualquier timer previo
  _nextRefreshAt = Date.now() + AUTO_REFRESH_MS;
  _autoRefreshTimer = setTimeout(function() {
    loadData();
  }, AUTO_REFRESH_MS);
  startCountdown();
}

function stopAutoRefresh() {
  if (_autoRefreshTimer)  { clearTimeout(_autoRefreshTimer);  _autoRefreshTimer = null; }
  if (_countdownTimer)    { clearInterval(_countdownTimer);   _countdownTimer   = null; }
}

function startCountdown() {
  if (_countdownTimer) clearInterval(_countdownTimer);
  _countdownTimer = setInterval(function() {
    var remaining = _nextRefreshAt - Date.now();
    if (remaining <= 0) { clearInterval(_countdownTimer); return; }
    var m = Math.floor(remaining / 60000);
    var s = Math.floor((remaining % 60000) / 1000);
    var syncEl = document.getElementById('syncTime');
    if (syncEl) {
      var base = syncEl.dataset.base || '';
      syncEl.textContent = base + '  ·  🔄 ' + m + ':' + String(s).padStart(2,'0');
    }
  }, 1000);
}


function showAlert(containerId, type, msg) {
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = `<div class="alert alert-${type}">${msg}</div>`;
  setTimeout(() => { if (el) el.innerHTML = ''; }, 5000);
}

// ══════════════════════════════════════════════════════════
//  DATA LOADING
// ══════════════════════════════════════════════════════════
async function loadData() {
  const btn = document.getElementById('syncBtn');
  btn.classList.add('loading');
  document.getElementById('syncIcon').textContent = '⏳';
  stopAutoRefresh();

  showSpinner('Cargando datos...');
  try {
    const url = `${SCRIPT_URL}?action=getHistorial&bd=devoluciones&_t=${Date.now()}`;
    const res = await fetch(url);
    const json = await res.json();
    if (!json.success) throw new Error(json.message || 'Error del servidor');
    processData(json.data || []);
    const timeStr = 'Actualizado: ' + new Date().toLocaleTimeString('es-AR');
    const syncEl = document.getElementById('syncTime');
    syncEl.dataset.base = timeStr;
    syncEl.textContent = timeStr;
    document.getElementById('statusDot').style.cssText = 'background:var(--green);box-shadow:0 0 6px var(--green);animation:pulse 2s infinite';
    startAutoRefresh();
  } catch (err) {
    const syncEl = document.getElementById('syncTime');
    syncEl.dataset.base = '⚠️ Error de conexión';
    syncEl.textContent  = '⚠️ Error de conexión';
    document.getElementById('statusDot').style.cssText = 'background:var(--red);box-shadow:none;';
    console.error('loadData error:', err.message);
    startAutoRefresh();
  } finally {
    hideSpinner();
    btn.classList.remove('loading');
    document.getElementById('syncIcon').textContent = '🔄';
  }
}

function processData(raw) {
  allData = raw.map(r => ({
    id: r['ID'] || '',
    fecha: parseDate(r['FECHA REGISTRO'] || r['FECHA'] || ''),
    sucursal: r['SUCURSAL'] || '',
    usuario: r['USUARIO'] || '',
    ean: String(r['EAN'] || ''),
    codInterno: r['COD. INTERNO'] || '',
    descripcion: r['DESCRIPCION'] || '',
    gramaje: r['GRAMAJE'] || '',
    cantidad: r['CANTIDAD'] || '',
    fechaVenc: r['FECHA VENC.'] || '',
    sector: r['SECTOR'] || '',
    seccion: r['SECCION'] || '',
    proveedor: r['PROVEEDOR'] || '',
    codProv: r['COD. PROVEEDOR'] || '',
    motivo: r['MOTIVO'] || '',
    lote: r['LOTE'] || '',
    aclaracion: r['ACLARACION'] || '',
    comentarios: r['COMENTARIOS'] || '',
    estado: r['ESTADO'] || 'PENDIENTE',
    obsNC: r['OBSERVACION N/C'] || '',
  })).filter(r => r.id);

  document.getElementById('recordCount').textContent = allData.length + ' registros';
  populateFilterSelects();
  setPreset(currentPreset, document.querySelector('.date-preset.active'));
  renderComprasPanel();
  renderProvTable();
  applyNcFilters();
}

function parseDate(str) {
  if (!str) return null;
  const s = String(str).trim();
  let m;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(parseInt(m[1]), parseInt(m[2])-1, parseInt(m[3]));
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
  return null;
}

function populateFilterSelects() {
  const sucursales = [...new Set(allData.map(r => r.sucursal).filter(Boolean))].sort();
  const proveedores = [...new Set(allData.map(r => r.proveedor).filter(Boolean))].sort();
  const motivos = [...new Set(allData.map(r => r.motivo).filter(Boolean))].sort();
  fillSelect('fSucursal', sucursales);
  fillSelect('fProveedor', proveedores);
  fillSelect('fMotivo', motivos);
  fillSelect('ncProveedor', proveedores);
  fillSelect('ncSucursal', sucursales);
  fillSelect('ncMotivo', motivos);
}

function fillSelect(id, values) {
  const sel = document.getElementById(id);
  if (!sel) return;
  const cur = sel.value;
  while (sel.options.length > 1) sel.remove(1);
  values.forEach(v => {
    const o = new Option(v, v);
    sel.add(o);
  });
  sel.value = cur;
}

// ══════════════════════════════════════════════════════════
//  DATE PRESETS
// ══════════════════════════════════════════════════════════
function setPreset(preset, btn) {
  currentPreset = preset;
  document.querySelectorAll('.date-preset').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');

  const now = new Date(); now.setHours(0,0,0,0);
  document.getElementById('customDateGroup').style.display = 'none';
  document.getElementById('customDateGroup2').style.display = 'none';

  if (preset === 'custom') {
    dateFrom = null; dateTo = null;
    document.getElementById('customDateGroup').style.display = '';
    document.getElementById('customDateGroup2').style.display = '';
  } else if (preset === 'all') {
    dateFrom = null; dateTo = null;
  } else {
    const days = {today:0, yesterday:1, '3d':3, '7d':7, '14d':14, '30d':30}[preset] ?? 0;
    dateTo = new Date(); dateTo.setHours(23,59,59,999);
    if (preset === 'yesterday') {
      const y = new Date(now); y.setDate(y.getDate()-1);
      dateFrom = y;
      dateTo = new Date(y); dateTo.setHours(23,59,59,999);
    } else {
      dateFrom = new Date(now); dateFrom.setDate(dateFrom.getDate()-days);
    }
  }
  applyFilters();
}

document.addEventListener('DOMContentLoaded', () => {
  const dpg = document.querySelector('.date-preset-group');
  if (dpg) {
    const customBtn = document.createElement('button');
    customBtn.className = 'date-preset';
    customBtn.textContent = 'Personalizado';
    customBtn.onclick = () => setPreset('custom', customBtn);
    dpg.appendChild(customBtn);
  }
});

// ══════════════════════════════════════════════════════════
//  FILTER & RENDER — OPERACIONES
// ══════════════════════════════════════════════════════════
function applyFilters() {
  const suc = document.getElementById('fSucursal').value;
  const prov = document.getElementById('fProveedor').value;
  const mot = document.getElementById('fMotivo').value;
  const est = document.getElementById('fEstado').value;
  const search = (document.getElementById('opSearch').value || '').toLowerCase();
  let from = dateFrom, to = dateTo;
  if (currentPreset === 'custom') {
    const f = document.getElementById('fDateFrom').value;
    const t = document.getElementById('fDateTo').value;
    from = f ? new Date(f) : null;
    to = t ? new Date(t + 'T23:59:59') : null;
  }
  filteredOp = allData.filter(r => {
    if (suc && r.sucursal !== suc) return false;
    if (prov && r.proveedor !== prov) return false;
    if (mot && r.motivo !== mot) return false;
    if (est && r.estado !== est) return false;
    if (from && r.fecha && r.fecha < from) return false;
    if (to && r.fecha && r.fecha > to) return false;
    if (search) {
      const hay = [r.descripcion, r.ean, r.lote, r.proveedor, r.sucursal, r.motivo].join(' ').toLowerCase();
      if (!hay.includes(search)) return false;
    }
    return true;
  });
  opPage = 1;
  renderOpTable();
}

function clearFilters() {
  ['fSucursal','fProveedor','fMotivo','fEstado'].forEach(id => { document.getElementById(id).value = ''; });
  document.getElementById('opSearch').value = '';
  setPreset('all', document.querySelectorAll('.date-preset')[6]);
}

function renderOpTable() {
  const start = (opPage-1)*PAGE_SIZE;
  const page = filteredOp.slice(start, start+PAGE_SIZE);
  const tbody = document.getElementById('opTableBody');

  if (filteredOp.length === 0) {
    tbody.innerHTML = '<div class="empty-state"><div class="icon">🔍</div><p>No hay registros con los filtros aplicados.</p></div>';
    document.getElementById('opPagination').style.display = 'none';
    return;
  }

  const html = `<div class="table-scroll"><table>
    <thead><tr>
      <th>ID</th><th>Fecha</th><th>Sucursal</th>
      <th>Descripción</th><th>EAN</th>
      <th>Cantidad</th><th>Proveedor</th>
      <th class="th-venc">⏰ Vencimiento</th><th>Motivo</th><th>Lote</th><th>Estado</th>
    </tr></thead>
    <tbody>${page.map(r => `
      <tr>
        <td class="cell-id">${r.id}</td>
        <td class="text-mono" style="font-size:11px">${formatDateDisplay(r.fecha)}</td>
        <td><span class="tag">${r.sucursal || '—'}</span></td>
        <td class="cell-main">${r.descripcion || '—'}</td>
        <td class="text-mono" style="font-size:11px">${r.ean || '—'}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace">${r.cantidad || '—'}</td>
        <td>${r.proveedor || '—'}</td>
        <td class="td-venc">${vencBadge(r.fechaVenc)}</td>
        <td>${motivoBadge(r.motivo)}</td>
        <td class="text-mono" style="font-size:11px">${r.lote || '—'}</td>
        <td>${estadoBadge(r.estado)}</td>
      </tr>`).join('')}
    </tbody>
  </table></div>`;
  tbody.innerHTML = html;
  renderPagination('opPagination', filteredOp.length, opPage, p => { opPage=p; renderOpTable(); });
}

// ══════════════════════════════════════════════════════════
//  COMPRAS PANEL
// ══════════════════════════════════════════════════════════
function renderComprasPanel() {
  const d = allData;
  document.getElementById('k-total').textContent = d.length;
  document.getElementById('k-nc').textContent = d.filter(r => r.estado==='N/C RECIBIDA').length;
  document.getElementById('k-pend').textContent = d.filter(r => r.estado==='PENDIENTE' || r.estado==='EN GESTION').length;
  document.getElementById('k-rech').textContent = d.filter(r => r.estado==='RECHAZADA').length;
  document.getElementById('k-prov').textContent = new Set(d.map(r=>r.proveedor).filter(Boolean)).size;
  document.getElementById('k-suc').textContent = new Set(d.map(r=>r.sucursal).filter(Boolean)).size;

  renderBarChart('chartMotivo', countBy(d, 'motivo'), '#4B7BFF');
  renderBarChart('chartSucursal', countBy(d, 'sucursal'), '#22C55E');
  renderBarChart('chartSector', countBy(d, 'sector'), '#A855F7');

  const ncData = d.filter(r => r.estado==='N/C RECIBIDA');
  renderBarChart('chartProvNC', countBy(ncData, 'proveedor'), '#F59E0B', 10);
}

function countBy(arr, key, limit) {
  const map = {};
  arr.forEach(r => { const v = r[key] || '(sin datos)'; map[v] = (map[v]||0)+1; });
  let entries = Object.entries(map).sort((a,b)=>b[1]-a[1]);
  if (limit) entries = entries.slice(0,limit);
  return entries;
}

function renderBarChart(containerId, entries, color, maxItems=15) {
  const el = document.getElementById(containerId);
  if (!entries.length) { el.innerHTML='<div class="empty-state" style="padding:20px"><p>Sin datos</p></div>'; return; }
  const max = entries[0][1];
  const items = entries.slice(0, maxItems);
  el.innerHTML = items.map(([label, val]) => `
    <div class="bar-row">
      <div class="bar-label tooltip-wrap" title="${label}">${label}<span class="tooltip-text">${label}</span></div>
      <div class="bar-track">
        <div class="bar-fill" style="width:${Math.max(8,Math.round(val/max*100))}%;background:${color}">
          <span class="bar-num">${val}</span>
        </div>
      </div>
    </div>`).join('');
}

function renderProvTable() {
  const search = (document.getElementById('provSearch')?.value || '').toLowerCase();
  const provs = [...new Set(allData.map(r=>r.proveedor).filter(Boolean))].sort();
  const rows = provs.filter(p => p.toLowerCase().includes(search)).map(prov => {
    const d = allData.filter(r => r.proveedor===prov);
    const pend = d.filter(r=>r.estado==='PENDIENTE').length;
    const gest = d.filter(r=>r.estado==='EN GESTION').length;
    const nc = d.filter(r=>r.estado==='N/C RECIBIDA').length;
    const rech = d.filter(r=>r.estado==='RECHAZADA').length;
    const cob = d.length > 0 ? Math.round((nc/d.length)*100) : 0;
    return { prov, total:d.length, pend, gest, nc, rech, cob };
  });
  if (!rows.length) {
    document.getElementById('provTableBody').innerHTML = '<tr><td colspan="7"><div class="empty-state"><p>Sin datos</p></div></td></tr>';
    return;
  }
  document.getElementById('provTableBody').innerHTML = rows.map(r => `
    <tr class="prov-table-row">
      <td>${r.prov}</td>
      <td class="text-right text-mono">${r.total}</td>
      <td class="text-right">${r.pend > 0 ? `<span class="badge badge-pendiente">${r.pend}</span>` : '—'}</td>
      <td class="text-right">${r.gest > 0 ? `<span class="badge badge-en-gestion">${r.gest}</span>` : '—'}</td>
      <td class="text-right">${r.nc > 0 ? `<span class="badge badge-nc-recibida">${r.nc}</span>` : '—'}</td>
      <td class="text-right">${r.rech > 0 ? `<span class="badge badge-rechazada">${r.rech}</span>` : '—'}</td>
      <td style="min-width:120px">
        <div class="nc-progress">
          <div class="progress-track"><div class="progress-fill" style="width:${r.cob}%"></div></div>
          <span class="text-mono" style="font-size:11px;min-width:34px">${r.cob}%</span>
        </div>
      </td>
    </tr>`).join('');
}

// ══════════════════════════════════════════════════════════
//  N/C PANEL
// ══════════════════════════════════════════════════════════
function applyNcFilters() {
  const prov = document.getElementById('ncProveedor')?.value || '';
  const est = document.getElementById('ncEstado')?.value || '';
  const suc = document.getElementById('ncSucursal')?.value || '';
  const mot = document.getElementById('ncMotivo')?.value || '';
  const search = (document.getElementById('ncSearch')?.value || '').toLowerCase();

  filteredNc = allData.filter(r => {
    if (prov && r.proveedor !== prov) return false;
    if (est && r.estado !== est) return false;
    if (suc && r.sucursal !== suc) return false;
    if (mot && r.motivo !== mot) return false;
    if (search) {
      const hay = [r.descripcion, r.ean, r.lote, r.proveedor, r.sucursal, r.id].join(' ').toLowerCase();
      if (!hay.includes(search)) return false;
    }
    return true;
  });
  ncPage = 1;
  renderNcTable();
}

function renderNcTable() {
  const start = (ncPage-1)*PAGE_SIZE;
  const page = filteredNc.slice(start, start+PAGE_SIZE);
  const tbody = document.getElementById('ncTableBody');

  if (filteredNc.length === 0) {
    tbody.innerHTML = '<div class="empty-state"><div class="icon">🔍</div><p>No hay registros con los filtros aplicados.</p></div>';
    document.getElementById('ncPagination').style.display = 'none';
    return;
  }

  const html = `<div class="table-scroll"><table>
    <thead><tr>
      <th class="checkbox-col"></th>
      <th>ID</th><th>Fecha</th><th>Sucursal</th>
      <th>Descripción</th><th>EAN</th>
      <th>Cant.</th><th>Venc.</th><th>Proveedor</th>
      <th>Vencimiento</th><th>Motivo</th><th>Lote</th><th>Estado</th><th>Obs. N/C</th>
    </tr></thead>
    <tbody>${page.map(r => `
      <tr>
        <td class="checkbox-col"><input type="checkbox" ${selectedIds.has(r.id)?'checked':''} onchange="toggleRow('${r.id}',this)"></td>
        <td class="cell-id">${r.id}</td>
        <td class="text-mono" style="font-size:11px">${formatDateDisplay(r.fecha)}</td>
        <td><span class="tag">${r.sucursal || '—'}</span></td>
        <td class="cell-main">${r.descripcion || '—'}</td>
        <td class="text-mono" style="font-size:11px">${r.ean || '—'}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace">${r.cantidad || '—'}</td>
        <td class="text-mono" style="font-size:11px">${formatVenc(r.fechaVenc)}</td>
        <td>${r.proveedor || '—'}</td>
        <td><span class="tag" style="font-size:10px">${r.motivo || '—'}</span></td>
        <td class="text-mono" style="font-size:11px">${r.lote || '—'}</td>
        <td>${estadoBadge(r.estado)}</td>
        <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;color:var(--text2);">${r.obsNC || '—'}</td>
      </tr>`).join('')}
    </tbody>
  </table></div>`;
  tbody.innerHTML = html;
  renderPagination('ncPagination', filteredNc.length, ncPage, p => { ncPage=p; renderNcTable(); });
}

function toggleRow(id, checkbox) {
  if (checkbox.checked) selectedIds.add(id);
  else selectedIds.delete(id);
  updateBulkBar();
}

function toggleSelectAll() {
  const checked = document.getElementById('selectAll').checked;
  filteredNc.forEach(r => { if (checked) selectedIds.add(r.id); else selectedIds.delete(r.id); });
  renderNcTable();
  updateBulkBar();
}

function clearSelection() {
  selectedIds.clear();
  document.getElementById('selectAll').checked = false;
  renderNcTable();
  updateBulkBar();
}

function updateBulkBar() {
  const bar = document.getElementById('bulkBar');
  const count = selectedIds.size;
  document.getElementById('bulkCount').textContent = count + ' registro' + (count===1?'':'s') + ' seleccionado' + (count===1?'':'s');
  bar.classList.toggle('visible', count > 0);
}

// ══════════════════════════════════════════════════════════
//  BULK ACTION → POST a Apps Script
// ══════════════════════════════════════════════════════════
function applyBulkAction() {
  const estado = document.getElementById('bulkEstado').value;
  const obs = document.getElementById('bulkObs').value.trim();
  if (!estado && !obs) { alert('Seleccioná un estado o escribí una observación.'); return; }
  if (selectedIds.size === 0) { alert('No hay registros seleccionados.'); return; }

  const ids = [...selectedIds];
  document.getElementById('modalTitle').textContent = '✏️ Confirmar actualización masiva';
  document.getElementById('modalText').textContent = `Se actualizarán ${ids.length} registro${ids.length===1?'':'s'}:`;
  document.getElementById('modalDetails').innerHTML = `
    ${estado ? `<div class="alert alert-info" style="margin-bottom:8px">Estado → <strong>${estado}</strong></div>` : ''}
    ${obs ? `<div class="alert alert-info">Obs. N/C → <em>${obs}</em></div>` : ''}
    <p style="font-size:11px;color:var(--text3);margin-top:8px">IDs: ${ids.slice(0,8).join(', ')}${ids.length>8?'…':''}</p>
  `;
  modalCallback = () => executeBulkUpdate(ids, estado, obs);
  const confirmBtn = document.getElementById('modalConfirmBtn');
  confirmBtn.onclick = function() {
    if (modalCallback) { var cb = modalCallback; modalCallback = null; cb(); }
  };
  document.getElementById('confirmModal').classList.add('open');
}

async function executeBulkUpdate(ids, estado, obs) {
  document.getElementById('confirmModal').classList.remove('open');
  modalCallback = null;

  showSpinner('Procesando 0 / ' + ids.length + '...');

  let success = 0, errors = 0;
  for (var i = 0; i < ids.length; i++) {
    var id = ids[i];
    showSpinner('Procesando ' + (i+1) + ' / ' + ids.length + '...');
    try {
      var payload = { action: 'updateRecord', id: id };
      if (estado) payload.estado = estado;
      if (obs) payload.observacionNC = obs;

      var res = await fetch(SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify(payload),
        headers: { 'Content-Type': 'text/plain;charset=utf-8' }
      });
      var json = await res.json();
      if (json.success) {
        success++;
        var r = allData.find(function(x){ return x.id === id; });
        if (r) {
          if (estado) r.estado = estado;
          if (obs) r.obsNC = obs;
        }
      } else {
        errors++;
        console.warn('Error en ID ' + id + ':', json.message);
      }
    } catch(e) {
      errors++;
      console.error('Fetch error ID ' + id + ':', e);
    }
  }

  hideSpinner();
  clearSelection();
  applyNcFilters();
  renderComprasPanel();
  renderProvTable();
  applyFilters();

  var ok = errors === 0;
  var msg = (ok ? '✅ ' : '⚠️ ') +
    success + ' actualizado' + (success === 1 ? '' : 's') +
    (errors ? ' · ' + errors + ' error' + (errors === 1 ? '' : 'es') : '');
  var alertEl = document.createElement('div');
  alertEl.className = 'alert alert-' + (ok ? 'success' : 'error');
  alertEl.style.cssText = 'position:fixed;bottom:20px;right:20px;z-index:300;min-width:280px;box-shadow:0 4px 20px rgba(0,0,0,.5);';
  alertEl.textContent = msg;
  document.body.appendChild(alertEl);
  setTimeout(function(){ if (alertEl.parentNode) alertEl.remove(); }, 5000);
}

function closeModal() {
  document.getElementById('confirmModal').classList.remove('open');
  modalCallback = null;
}

document.getElementById('confirmModal').addEventListener('click', function(e) {
  if (e.target === this) closeModal();
});

// ══════════════════════════════════════════════════════════
//  SPINNER
// ══════════════════════════════════════════════════════════
function showSpinner(msg) {
  document.getElementById('spinnerMsg').textContent = msg || 'Procesando...';
  document.getElementById('spinnerOverlay').classList.add('active');
}
function hideSpinner() {
  document.getElementById('spinnerOverlay').classList.remove('active');
}

// ══════════════════════════════════════════════════════════
//  EAN-13 SVG BARCODE GENERATOR (pure JS, no deps)
// ══════════════════════════════════════════════════════════
var _EAN13 = (function() {
  var L = ['0001101','0011001','0010011','0111101','0100011','0110001','0101111','0111011','0110111','0001011'];
  var G = ['0100111','0110011','0011011','0100001','0011101','0111001','0000101','0010001','0001001','0010111'];
  var R = ['1110010','1100110','1101100','1000010','1011100','1001110','1010000','1000100','1001000','1110100'];
  var PARITY = ['LLLLLL','LLGLGG','LLGGLG','LLGGGL','LGLLGG','LGGLLG','LGGGLL','LGLGLG','LGLGGL','LGGLGL'];

  function checkDigit(s) {
    var d = s.substring(0,12).split('').map(Number);
    var t = d.reduce(function(acc,v,i){ return acc + v * (i%2===0?1:3); }, 0);
    return ((10 - (t%10)) % 10).toString();
  }

  function encode(ean) {
    var s = String(ean).replace(/\D/g,'');
    if (s.length === 12) s = s + checkDigit(s);
    if (s.length !== 13) return null;
    var parity = PARITY[parseInt(s[0])];
    var bits = '101';
    for (var i=1; i<=6; i++) {
      var d = parseInt(s[i]);
      bits += parity[i-1]==='L' ? L[d] : G[d];
    }
    bits += '01010';
    for (var i=7; i<=12; i++) bits += R[parseInt(s[i])];
    bits += '101';
    return { bits: bits, digits: s };
  }

  function toSVG(ean, opts) {
    opts = opts || {};
    var barW = opts.barW || 2;
    var h = opts.h || 70;
    var textH = 12;
    var totalH = h + textH + 4;

    var enc = encode(ean);
    if (!enc) return '<svg width="20" height="' + totalH + '"><text y="12" font-size="9" fill="red">EAN inv.</text></svg>';

    var totalW = enc.bits.length * barW + 20;
    var xOff = 10;
    var bars = '';
    for (var i=0; i<enc.bits.length; i++) {
      if (enc.bits[i] === '1') {
        bars += '<rect x="' + (xOff + i*barW) + '" y="0" width="' + barW + '" height="' + h + '" fill="#000"/>';
      }
    }
    var d = enc.digits;
    var labelY = h + textH;
    var leftNum = '<text x="' + (xOff - 4) + '" y="' + labelY + '" font-size="9" font-family="monospace" text-anchor="end">' + d[0] + '</text>';
    var leftCenter = xOff + (3 + 3*7) * barW;
    var leftGroup = '<text x="' + leftCenter + '" y="' + labelY + '" font-size="9" font-family="monospace" text-anchor="middle">' + d.substring(1,7) + '</text>';
    var rightCenter = xOff + (3+42+5 + 3*7) * barW;
    var rightGroup = '<text x="' + rightCenter + '" y="' + labelY + '" font-size="9" font-family="monospace" text-anchor="middle">' + d.substring(7,13) + '</text>';

    return '<svg xmlns="http://www.w3.org/2000/svg" width="' + totalW + '" height="' + totalH + '" viewBox="0 0 ' + totalW + ' ' + totalH + '">' +
      bars + leftNum + leftGroup + rightGroup + '</svg>';
  }

  return { toSVG: toSVG };
})();

// ══════════════════════════════════════════════════════════
//  EXPORT
// ══════════════════════════════════════════════════════════
function exportToExcel(mode) {
  if (mode === 'op') {
    exportOpXlsx(filteredOp);
  } else {
    exportNcCsv();
  }
}

// ── Helper: genera PNG de código de barras usando JsBarcode + canvas ──
function barcodeDataUrl(ean) {
  var canvas = document.createElement('canvas');
  var code = String(ean || '').replace(/\D/g, '');
  if (!code) return null;
  var format = /^\d{13}$/.test(code) ? 'ean13' : (/^\d{8}$/.test(code) ? 'ean8' : 'code128');
  try {
    JsBarcode(canvas, code, {
      format: format,
      displayValue: false,
      margin: 6,
      width: 3,
      height: 90
    });
  } catch(e) {
    try {
      JsBarcode(canvas, code, { format: 'code128', displayValue: false, margin: 6, width: 3, height: 90 });
    } catch(e2) { return null; }
  }
  return canvas.toDataURL('image/png').split(',')[1]; // base64 sin prefijo
}

// ── EXPORT PANEL GENERAL → .xlsx con imagen de barcode por fila ──
async function exportOpXlsx(rows) {
  if (!rows.length) { alert('No hay registros para exportar.'); return; }

  var workbook = new ExcelJS.Workbook();
  var ws = workbook.addWorksheet('Panel General');

  // Configuración de página A4 horizontal, listo para imprimir
  ws.pageSetup = {
    paperSize: 9,           // A4
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 }
  };
  ws.pageSetup.printTitlesRow = '1:1';

  // Anchos de columnas: COD-BAR | Fecha | Sucursal | Descripción | Gramaje | Cantidad | Fecha Venc. | Proveedor | Motivo | Lote | Estado
  ws.columns = [
    { key: 'bc',    width: 26 },
    { key: 'fecha', width: 12 },
    { key: 'suc',   width: 18 },
    { key: 'desc',  width: 36 },
    { key: 'gram',  width: 10 },
    { key: 'cant',  width: 10 },
    { key: 'venc',  width: 13 },
    { key: 'prov',  width: 24 },
    { key: 'mot',   width: 22 },
    { key: 'lote',  width: 14 },
    { key: 'est',   width: 16 }
  ];

  // ── Fila de encabezado ──
  var headerRow = ws.addRow(['COD-BAR', 'Fecha', 'Sucursal', 'Descripción', 'Gramaje', 'Cantidad', 'Fecha Venc.', 'Proveedor', 'Motivo', 'Lote', 'Estado']);
  headerRow.height = 22;
  headerRow.eachCell(function(cell) {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E2D5A' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = {
      bottom: { style: 'medium', color: { argb: 'FF4B7BFF' } }
    };
  });

  // ── Filas de datos ──
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var rowIndex = i + 2; // fila 1 = header
    var isEven = i % 2 === 1;
    var fillColor = isEven ? 'FFF0F4FF' : 'FFFFFFFF';

    // Celda vacía para la imagen (columna A)
    var dataRow = ws.addRow([
      '',  // COD-BAR → imagen insertada aparte
      r.fecha ? r.fecha.toLocaleDateString('es-AR') : '',
      r.sucursal    || '',
      r.descripcion || '',
      r.gramaje     || '',
      r.cantidad    || '',
      formatVenc(r.fechaVenc),
      r.proveedor   || '',
      r.motivo      || '',
      r.lote        || '',
      r.estado      || ''
    ]);
    dataRow.height = 70;

    dataRow.eachCell({ includeEmpty: true }, function(cell, colNumber) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
      cell.alignment = { vertical: 'middle', horizontal: colNumber === 1 ? 'center' : 'left', wrapText: colNumber === 4 };
      cell.font = { size: 10 };
      cell.border = {
        bottom: { style: 'thin', color: { argb: 'FFDDDDDD' } },
        right:  { style: 'thin', color: { argb: 'FFDDDDDD' } }
      };
    });

    // Estado con color de fondo
    var estCell = dataRow.getCell(11);
    var estColors = {
      'PENDIENTE':   { bg: 'FFFFF3CD', fg: 'FF856404' },
      'EN GESTION':  { bg: 'FFD1ECF1', fg: 'FF0C5460' },
      'N/C RECIBIDA':{ bg: 'FFD4EDDA', fg: 'FF155724' },
      'RECHAZADA':   { bg: 'FFF8D7DA', fg: 'FF721C24' }
    };
    var ec = estColors[r.estado];
    if (ec) {
      estCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ec.bg } };
      estCell.font = { size: 10, bold: true, color: { argb: ec.fg } };
    }

    // ── Insertar imagen de código de barras ──
    var b64 = barcodeDataUrl(r.ean);
    if (b64) {
      var imageId = workbook.addImage({ base64: b64, extension: 'png' });
      ws.addImage(imageId, {
        tl: { col: 0.08, row: rowIndex - 1 + 0.06 },
        ext: { width: 145, height: 60 },
        editAs: 'oneCell'
      });
    }
  }

  // ── Freeze primera fila ──
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];

  // ── Generar y descargar ──
  try {
    var buf = await workbook.xlsx.writeBuffer();
    var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'registros_panel_' + formatIso(new Date()) + '.xlsx';
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch(err) {
    console.error(err);
    alert('Error al generar el Excel: ' + err.message);
  }
}

function exportNcCsv() {
  var rows = selectedIds.size > 0
    ? filteredNc.filter(function(r){ return selectedIds.has(r.id); })
    : filteredNc;
  if (!rows.length) { alert('No hay registros para exportar.'); return; }

  var headers = ['ID','Fecha','Sucursal','EAN','Descripción','Gramaje','Cantidad',
    'Fecha Venc.','Proveedor','Motivo','Lote','Estado','Obs. N/C'];
  var wsData = [headers];
  rows.forEach(function(r) {
    wsData.push([
      r.id,
      r.fecha ? r.fecha.toLocaleDateString('es-AR') : '',
      r.sucursal, r.ean, r.descripcion, r.gramaje, r.cantidad,
      formatVenc(r.fechaVenc), r.proveedor, r.motivo, r.lote, r.estado, r.obsNC
    ]);
  });
  var wb = XLSX.utils.book_new();
  var ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [
    {wch:22},{wch:12},{wch:18},{wch:16},{wch:32},{wch:10},{wch:10},
    {wch:14},{wch:22},{wch:22},{wch:14},{wch:16},{wch:28}
  ];
  XLSX.utils.book_append_sheet(wb, ws, 'Gestión NC');
  XLSX.writeFile(wb, 'gestion_nc_' + formatIso(new Date()) + '.xlsx');
}

// ══════════════════════════════════════════════════════════
//  HELPERS UI
// ══════════════════════════════════════════════════════════
function estadoBadge(estado) {
  const map = {
    'PENDIENTE': 'pendiente',
    'EN GESTION': 'en-gestion',
    'N/C RECIBIDA': 'nc-recibida',
    'RECHAZADA': 'rechazada',
  };
  const cls = map[estado] || 'pendiente';
  return `<span class="badge badge-${cls}">${estado || 'PENDIENTE'}</span>`;
}

// ── Badge de motivo con color por grupo ──
function motivoBadge(motivo) {
  if (!motivo) return '<span class="motivo-badge motivo-otro">—</span>';
  const m = motivo.toUpperCase();
  let cls = 'motivo-otro';
  if (m === 'ACCION 2X1' || m === 'VENCIDO ACCION 2X1')           cls = 'motivo-2x1';
  else if (m === 'ACCION 50% OFF' || m === 'VENCIDO ACCION 50% OFF') cls = 'motivo-50off';
  else if (m === 'OTRO DESCUENTO')                                   cls = 'motivo-descuento';
  else if (m === 'VENCIDO')                                          cls = 'motivo-vencido';
  else if (m === 'ROTO/DAÑADO' || m === 'MAL ESTADO')               cls = 'motivo-danado';
  else if (m.startsWith('DECOMISO'))                                 cls = 'motivo-decomiso';
  return `<span class="motivo-badge ${cls}">${motivo}</span>`;
}

// ── Celda de vencimiento con semáforo de urgencia ──
function vencBadge(str) {
  const txt = formatVenc(str);
  if (txt === '—') return `<span class="venc-badge venc-nd">—</span>`;

  // Parsear la fecha formateada DD-MM-YYYY
  const parts = txt.split('-');
  if (parts.length !== 3) return `<span class="venc-badge venc-nd">${txt}</span>`;
  const d = new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0]));
  const diff = Math.ceil((d - new Date()) / 86400000); // días restantes

  let cls = 'venc-ok';
  if (diff < 0)        cls = 'venc-expired';
  else if (diff <= 7)  cls = 'venc-critical';
  else if (diff <= 30) cls = 'venc-warn';

  const icon = diff < 0 ? '💀' : diff <= 7 ? '🔴' : diff <= 30 ? '🟡' : '🟢';
  return `<span class="venc-badge ${cls}">${icon} ${txt}</span>`;
}

function formatVenc(str) {
  if (!str) return '—';
  const s = String(str).trim();
  var m;
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[3].padStart(2,'0') + '-' + m[2].padStart(2,'0') + '-' + m[1];
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
  if (m) return m[1].padStart(2,'0') + '-' + m[2].padStart(2,'0') + '-' + m[3];
  return s;
}

function formatDateDisplay(d) {
  if (!d) return '—';
  return d.toLocaleDateString('es-AR', { day:'2-digit', month:'2-digit', year:'2-digit' });
}

const _pageCbs = {};
function renderPagination(containerId, total, current, callback) {
  const el = document.getElementById(containerId);
  const totalPages = Math.ceil(total/PAGE_SIZE);
  if (totalPages <= 1) { el.style.display='none'; return; }
  _pageCbs[containerId] = callback;
  el.style.display = 'flex';
  const shown = Math.min(current*PAGE_SIZE, total);
  let html = '<span>Mostrando ' + ((current-1)*PAGE_SIZE+1) + '\u2013' + shown + ' de ' + total + '</span><div class="page-btns">';
  html += '<button class="page-btn" data-pg="' + (current-1) + '" data-cb="' + containerId + '" ' + (current===1?'disabled':'') + '>←</button>';
  for (let i=Math.max(1,current-2); i<=Math.min(totalPages,current+2); i++) {
    html += '<button class="page-btn ' + (i===current?'active':'') + '" data-pg="' + i + '" data-cb="' + containerId + '">' + i + '</button>';
  }
  html += '<button class="page-btn" data-pg="' + (current+1) + '" data-cb="' + containerId + '" ' + (current===totalPages?'disabled':'') + '>→</button>';
  html += '</div>';
  el.innerHTML = html;
  el.querySelectorAll('button[data-pg]').forEach(function(btn) {
    btn.addEventListener('click', function() {
      if (btn.disabled) return;
      var pg = parseInt(btn.getAttribute('data-pg'));
      var cb = _pageCbs[btn.getAttribute('data-cb')];
      if (cb) cb(pg);
    });
  });
}

// ══════════════════════════════════════════════════════════
//  DEMO DATA
// ══════════════════════════════════════════════════════════
function loadDemoData() {
  const sucursales = ['SM PALERMO', 'SM BELGRANO', 'SM CABALLITO', 'SM ALMAGRO'];
  const provs = ['ARCOR SA', 'MOLINOS RIO DE LA PLATA', 'UNILEVER ARG', 'MASTELLONE', 'DANONE ARG', 'LA SERENISIMA'];
  const motivos = ['PRODUCTO VENCIDO','RETIRO PROVEEDOR','2X1 DEVOLUCION','PRODUCTO DAÑADO','PROMO LIQUIDACION','ERROR DE PEDIDO'];
  const sectores = ['LACTEOS','BEBIDAS','FIAMBRES','PANADERIA','ROTISERIA','CONGELADOS'];
  const estados = ['PENDIENTE','EN GESTION','N/C RECIBIDA','RECHAZADA'];
  const descrips = ['Leche Entera 1L','Yogur Natural 400g','Manteca 200g','Queso Cremoso 400g','Dulce de Leche 400g','Galletitas Oreo 118g','Aceite Girasol 900ml','Jabón Líquido 750ml','Agua Mineral 1.5L','Jugo Tang Naranja'];

  const data = [];
  const now = new Date();
  for (let i=0; i<120; i++) {
    const d = new Date(now);
    d.setDate(d.getDate() - Math.floor(Math.random()*45));
    const id = `DEV-${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}-${String(i+1).padStart(5,'0')}`;
    const est = estados[Math.floor(Math.random()*4)];
    data.push({
      id,
      fecha: d,
      sucursal: sucursales[i%sucursales.length],
      usuario: `user${i%5+1}@empresa.com`,
      ean: String(7790000000000 + i*13),
      codInterno: 'CI-' + String(1000+i),
      descripcion: descrips[i%descrips.length],
      gramaje: ['1L','400g','200g','118g','900ml'][i%5],
      cantidad: String(Math.floor(Math.random()*20)+1),
      fechaVenc: formatIso(new Date(now.getTime() + (Math.random()-0.3)*30*86400000)),
      sector: sectores[i%sectores.length],
      seccion: 'SEC-' + String(i%3+1),
      proveedor: provs[i%provs.length],
      codProv: 'P' + String(100+i%provs.length),
      motivo: motivos[i%motivos.length],
      lote: 'L' + String(2024000+i),
      aclaracion: '',
      comentarios: '',
      estado: est,
      obsNC: est==='N/C RECIBIDA' ? 'NC recibida OK' : '',
    });
  }
  allData = data;
  document.getElementById('recordCount').textContent = allData.length + ' registros (DEMO)';
  document.getElementById('syncTime').textContent = '🎭 Modo demo — sin auto-actualización';
  document.getElementById('syncTime').dataset.base = '🎭 Modo demo — sin auto-actualización';
  document.getElementById('statusDot').style.cssText = 'background:#A855F7;box-shadow:0 0 6px #A855F7;';
  stopAutoRefresh(); // en modo demo no tiene sentido el auto-refresh
  populateFilterSelects();
  setPreset('all', document.querySelectorAll('.date-preset')[6]);
  renderComprasPanel();
  renderProvTable();
  applyNcFilters();
  showAlert('configAlert', 'info', '🎭 Datos de demostración cargados. 120 registros ficticios.');
  showPanel('operaciones');
}

function formatIso(d) {
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}