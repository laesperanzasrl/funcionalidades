// ══════════════════════════════════════════════════════
//  STATE
// ══════════════════════════════════════════════════════

let accDateMode = 'registro';
let loteMode = false;
let loteSeleccionados = new Map();
let loteQueue = [];
let SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwj0qEOm9THbYxw0TYek2Oot3dlL1wn7YmPLtYknFzrGBQJXFnd-kh7yxXtFgYFyC-B/exec";
let devData = [];
let venData = [];
let devMap = new Map();
let venGrupos = [];
let openGrupos = new Set();
let openSucs = new Set();

let venProdGrupos = [];
let filteredVenProd = [];
let openProds = new Set();

let filteredVen = [];
let filteredAcc = [];
let filteredNc = [];
let selectedIds = new Set();
let modalCb = null;

let accPreset = 'today';
let accDateFrom = null, accDateTo = null;
let accPage = 1, ncPage = 1;
const PAGE_SIZE = 50;

// ── ESTADO FILTROS VENCIMIENTOS ──────────────────────
let vencPreset = 'all';
let vencDateMode = 'registro';
let vencDateFrom = null;
let vencDateTo = null;
// ─────────────────────────────────────────────────────

let _autoTimer = null, _cdTimer = null, _nextRefresh = null;
const AUTO_MS = 30 * 60 * 1000;

const URG_ORDER = { VENCIDO: 0, CRITICO: 1, URGENTE: 2, PROXIMO: 3, ATENCION: 4, NORMAL: 5 };
const URG_LABELS = { VENCIDO: 'VENCIDO', CRITICO: 'CRÍTICO', URGENTE: 'URGENTE', PROXIMO: 'PRÓXIMO', ATENCION: 'ATENCIÓN', NORMAL: 'NORMAL' };
const BAR_CLRS = { VENCIDO: 'var(--venc-v-fg)', CRITICO: 'var(--venc-c-fg)', URGENTE: 'var(--venc-u-fg)', PROXIMO: 'var(--venc-p-fg)', ATENCION: 'var(--venc-a-fg)', NORMAL: 'var(--venc-n-fg)' };

const SUCURSALES_LIST = ['HIPER', 'CENTRO', 'RIBERA', 'MAYORISTA'];
let currentTransferData = null;
let currentAdjustData = null;

let currentAccTransferData = null;

let _currentReplicarData = null;

// ── LOG DATA (operaciones del día) ───────────────────
let venLogs = [];
let _histVenActiveTab = 'transfers';
let _histAccActiveTab = 'transfers';

// ══════════════════════════════════════════════════════
//  HELPER: ¿Es producto pesable?
// ══════════════════════════════════════════════════════
const SECTORES_PESABLES = ['FIAMBRE', 'CARNICER', 'VERDULER', 'FRUTA', 'PESCAD', 'ROTISERI'];

// Devuelve true solo cuando un DEV es hijo de una TRANSFERENCIA entre sucursales.
// Los DEV con ID_ORIGEN = 'VEN-...' son acciones originales creadas desde el panel de
// Vencimientos: NO son transferencias, deben mostrarse normalmente y aparecer en N/C.
function esTransferHijo(r) {
  return !!(r.idOrigen && String(r.idOrigen).startsWith('DEV-'));
}

function esPesable(item) {
  if (!item) return false;
  const ean = String(item.ean || item.EAN || '').trim();
  if (ean.startsWith('23')) return true;
  const sector = (item.sector || item.SECTOR || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const seccion = (item.seccion || item.SECCION || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return SECTORES_PESABLES.some(k => sector.includes(k) || seccion.includes(k));
}

// ══════════════════════════════════════════════════════
//  LOG OPERACIONES — Historial del día
// ══════════════════════════════════════════════════════

function _isToday(ts) {
  if (!ts) return false;
  const d = new Date(ts);
  if (isNaN(d)) return false;
  const t = new Date();
  return d.getFullYear() === t.getFullYear() && d.getMonth() === t.getMonth() && d.getDate() === t.getDate();
}

function _extractFechaVenc(idVen) {
  if (!idVen) return null;
  const m = String(idVen).match(/VEN-\d+-(\d{8})-/);
  if (!m) return null;
  const s = m[1];
  return `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}`;
}

function _fmtTime(ts) {
  if (!ts) return '—';
  const d = new Date(ts);
  return isNaN(d) ? '—' : d.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' });
}

function processVenLogs(raw) {
  venLogs = (raw || []).map(r => ({
    timestamp: r['TIMESTAMP'] || r['timestamp'] || '',
    tipo: String(r['TIPO'] || r['tipo'] || ''),
    idVen: r['ID_VEN'] || r['id_ven'] || '',
    idDev: r['ID_DEV'] || r['id_dev'] || '',
    descripcion: r['DESCRIPCION'] || r['descripcion'] || '',
    ean: String(r['EAN'] || r['ean'] || ''),
    sucursal: r['SUCURSAL'] || r['sucursal'] || '',
    cantAnterior: parseFloat(String(r['CANT_ANTERIOR'] ?? r['cant_anterior'] ?? '')),
    cantNueva: parseFloat(String(r['CANT_NUEVA'] ?? r['cant_nueva'] ?? '')),
    delta: parseFloat(String(r['DELTA'] ?? r['delta'] ?? '')),
    destino: r['DESTINO'] || r['destino'] || '',
    usuario: r['USUARIO'] || r['usuario'] || '',
    notas: r['NOTAS'] || r['notas'] || '',
  }));
  _updateHistBadges();
}

function _updateHistBadges() {
  const today = venLogs.filter(r => _isToday(r.timestamp));

  const VEN_TYPES = [
    'TRANSFERENCIA_SALIDA', 'TRANSFERENCIA_ENTRADA', 'TRANSFERENCIA_NUEVA_SUCURSAL',
    'AJUSTE_POSITIVO', 'AJUSTE_NEGATIVO', 'AJUSTE_SIN_CAMBIO', 'AJUSTE_STOCK_A_CERO',
    'CONTROL_VEN_INSERT', 'CONTROL_VEN_UPSERT', 'ACCION_DESDE_VEN'
  ];
  const ACC_TYPES = [
    'ACC_TRANSFER_SALIDA', 'ACC_TRANSFERENCIA_SALIDA',
    'ACC_TRANSFER_NUEVO_REGISTRO', 'ACC_TRANSFER_ENTRADA_EXISTENTE',
    'ACC_MARCADO_VENDIDO',
    'ACC_AJUSTE_POSITIVO', 'ACC_AJUSTE_NEGATIVO', 'ACC_AJUSTE_A_CERO', 'ACC_AJUSTE_SIN_CAMBIO',
    'DEVOLUCION_CON_DESCUENTO_VEN', 'DEVOLUCION_VINCULACION_MANUAL',
    'VYC_INSERT', 'VYC_CON_DESCUENTO_VEN',
    'ACCION_DESDE_VEN', 'REGISTRO_VENCIMIENTO_MANUAL', 'AUTO_VENCIMIENTO_DIARIO'
  ];
  const venCnt = today.filter(r => VEN_TYPES.includes(r.tipo)).length;
  const accCnt = today.filter(r => ACC_TYPES.includes(r.tipo)).length;

  const venBadge = document.getElementById('histVenBadge');
  if (venBadge) { venBadge.textContent = venCnt; venBadge.style.display = venCnt ? '' : 'none'; }
  const accBadge = document.getElementById('histAccBadge');
  if (accBadge) { accBadge.textContent = accCnt; accBadge.style.display = accCnt ? '' : 'none'; }
}

// ── Abrir / cerrar modales ───────────────────────────
function openHistVen() {
  _histVenActiveTab = 'transfers';
  const today = new Date().toLocaleDateString('es-AR', { weekday: 'long', day: 'numeric', month: 'long' });
  const el = document.getElementById('histVenDate');
  if (el) el.textContent = today.toUpperCase();

  // Forzar pestaña activa visualmente
  document.querySelectorAll('#histVenModal .hist-tab').forEach(t => {
    t.classList.toggle('active', t.dataset.tab === 'transfers');
  });

  document.getElementById('histVenModal').classList.add('open');
  _renderHistVen();
}
function closeHistVenModal(e) {
  if (!e || e.target === document.getElementById('histVenModal'))
    document.getElementById('histVenModal').classList.remove('open');
}
function openHistAcc() {
  _histAccActiveTab = 'transfers';
  const today = new Date().toLocaleDateString('es-AR', { weekday: 'long', day: 'numeric', month: 'long' });
  const el = document.getElementById('histAccDate');
  if (el) el.textContent = today.toUpperCase();

  // Forzar pestaña activa visualmente
  document.querySelectorAll('#histAccModal .hist-tab').forEach(t => {
    t.classList.toggle('active', t.dataset.tab === 'transfers');
  });

  document.getElementById('histAccModal').classList.add('open');
  _renderHistAcc();
}
function closeHistAccModal(e) {
  if (!e || e.target === document.getElementById('histAccModal'))
    document.getElementById('histAccModal').classList.remove('open');
}

function switchHistVenTab(tab) {
  _histVenActiveTab = tab;
  document.querySelectorAll('#histVenModal .hist-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === tab));
  _renderHistVen();
}
function switchHistAccTab(tab) {
  _histAccActiveTab = tab;
  document.querySelectorAll('#histAccModal .hist-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === tab));
  _renderHistAcc();
}

// ── Render VEN ───────────────────────────────────────
function _renderHistVen() {
  const today = venLogs.filter(r => _isToday(r.timestamp));
  const VEN_TRANSFER_TYPES = ['TRANSFERENCIA_SALIDA', 'TRANSFERENCIA_ENTRADA', 'TRANSFERENCIA_NUEVA_SUCURSAL'];
  const VEN_ADJ_TYPES = ['AJUSTE_POSITIVO', 'AJUSTE_NEGATIVO', 'AJUSTE_SIN_CAMBIO', 'AJUSTE_STOCK_A_CERO'];
  const transfers = today.filter(r => VEN_TRANSFER_TYPES.includes(r.tipo));
  const ajustes = today.filter(r => VEN_ADJ_TYPES.includes(r.tipo));

  // Update tab badges
  document.getElementById('histVenTabBadgeT').textContent = transfers.filter(r => r.tipo === 'TRANSFERENCIA_SALIDA').length;
  document.getElementById('histVenTabBadgeA').textContent = ajustes.length;

  const el = document.getElementById('histVenContent');
  // ── Actualizar badge de registros ──────────────────────────
  const VEN_REG_TYPES = ['CONTROL_VEN_INSERT', 'CONTROL_VEN_UPSERT'];
  const registros = today.filter(r => VEN_REG_TYPES.includes(r.tipo));
  document.getElementById('histVenTabBadgeR').textContent = registros.length;

  if (_histVenActiveTab === 'registros') {
    if (!registros.length) {
      el.innerHTML = `<div class="hist-empty"><div class="hist-empty-icon">📋</div><div>Sin registros de control de vencimiento hoy</div></div>`;
      return;
    }
    el.innerHTML = registros.map(r => {
      const _prod = venData.find(v => String(v.ean || '').trim() === r.ean)
        || devData.find(d => String(d.ean || '').trim() === r.ean) || {};
      const gramaje = _prod.gramaje || '';
      const proveedor = _prod.proveedor || '';
      const fechaVenc = _extractFechaVenc(r.idVen);
      const isUpsert = r.tipo === 'CONTROL_VEN_UPSERT';
      const deltaStr = !isNaN(r.delta)
        ? (r.delta > 0 ? '+' + r.delta : String(r.delta)) + ' u.'
        : null;

      return `<div class="hist-op-card registro-ven">
        <div class="hist-op-row">
          <span class="hist-op-time">${_fmtTime(r.timestamp)}</span>
          <div class="hist-op-desc">${esc(r.descripcion)}</div>
          ${isUpsert
          ? `<span class="hist-chip" style="color:var(--amber);border-color:rgba(245,166,35,.25);background:rgba(245,166,35,.07)">↺ SUMA</span>`
          : `<span class="hist-chip pos">✚ NUEVO</span>`}
        </div>
        <div class="hist-op-meta">
          <span class="hist-chip ean">🔖 ${esc(r.ean)}</span>
          <span class="hist-chip suc-orig">🏪 ${esc(r.sucursal)}</span>
          ${gramaje ? `<span class="hist-chip">⚖️ ${esc(gramaje)}</span>` : ''}
          ${proveedor ? `<span class="hist-chip">🏢 ${esc(proveedor)}</span>` : ''}
          ${fechaVenc ? `<span class="hist-chip venc">📅 ${fechaVenc}</span>` : ''}
          ${!isNaN(r.cantAnterior) ? `<span class="hist-chip">Antes: ${r.cantAnterior}</span>` : ''}
          ${!isNaN(r.cantNueva) ? `<span class="hist-chip qty">Ahora: ${r.cantNueva}</span>` : ''}
          ${deltaStr ? `<span class="hist-chip ${r.delta > 0 ? 'pos' : 'neg'}" style="font-weight:800">Δ ${deltaStr}</span>` : ''}
          ${r.usuario ? `<span class="hist-chip" style="color:var(--text3)">👤 ${esc(r.usuario)}</span>` : ''}
          <button class="btn-replicar" onclick="abrirModalReplicar('${esc(r.idVen)}','${esc(r.descripcion)}','${esc(r.ean)}','${esc(r.sucursal)}','${fechaVenc || ''}','${esc(gramaje)}')">
            📋 Replicar en mi sucursal
          </button>
        </div>
      </div>`;
    }).join('');
    return;
  }

  if (_histVenActiveTab === 'transfers') {
    const salidas = today.filter(r => r.tipo === 'TRANSFERENCIA_SALIDA');
    if (!salidas.length) {
      el.innerHTML = `<div class="hist-empty"><div class="hist-empty-icon">⇄</div><div>Sin transferencias registradas hoy</div></div>`;
      return;
    }
    el.innerHTML = salidas.map(r => {

      const _prod = venData.find(v => String(v.ean || '').trim() === r.ean)
        || devData.find(d => String(d.ean || '').trim() === r.ean) || {};
      const gramaje = _prod.gramaje || '';
      const proveedor = _prod.proveedor || '';
      const fechaVenc = _extractFechaVenc(r.idVen);
      const qty = Math.abs(isNaN(r.delta) ? (r.cantAnterior - r.cantNueva) : r.delta);
      return `<div class="hist-op-card transfer">
        <div class="hist-op-row">
                  <span class="hist-op-time">${_fmtTime(r.timestamp)}</span>
          <div class="hist-op-desc">${esc(r.descripcion)}</div>

          <span class="hist-chip ean">🔖 ${esc(r.ean)}</span>
          ${gramaje ? `<span class="hist-chip">⚖️ ${esc(gramaje)}</span>` : ''}
          ${proveedor ? `<span class="hist-chip">🏢 ${esc(proveedor)}</span>` : ''}
        </div>
        <div class="hist-op-meta">
          <span class="hist-chip ean">🔖 ${esc(r.ean)}</span>
          <span class="hist-chip suc-orig">📤 ${esc(r.sucursal)}</span>
          <span class="hist-arrow">→</span>
          <span class="hist-chip suc-dest">📥 ${esc(r.destino || '—')}</span>
          ${fechaVenc ? `<span class="hist-chip venc">📅 ${fechaVenc}</span>` : ''}
          <span class="hist-chip qty">Cant: ${qty} u.</span>
        </div>
      </div>`;
    }).join('');
  } else {
    if (!ajustes.length) {
      el.innerHTML = `<div class="hist-empty"><div class="hist-empty-icon">⚖</div><div>Sin ajustes registrados hoy</div></div>`;
      return;
    }
    el.innerHTML = ajustes.map(r => {

      const _prod = venData.find(v => String(v.ean || '').trim() === r.ean)
        || devData.find(d => String(d.ean || '').trim() === r.ean) || {};
      const gramaje = _prod.gramaje || '';
      const proveedor = _prod.proveedor || '';
      const isPos = r.tipo === 'AJUSTE_POSITIVO';
      const isNeg = r.tipo === 'AJUSTE_NEGATIVO' || r.tipo === 'AJUSTE_STOCK_A_CERO';

      const tipoLabel = {
        AJUSTE_POSITIVO: '＋ Positivo',
        AJUSTE_NEGATIVO: '－ Negativo',
        AJUSTE_STOCK_A_CERO: '⬛ A cero',
        AJUSTE_SIN_CAMBIO: '＝ Sin cambio',
        ACCION_DESDE_VEN: '📋 Acción registrada',
        REGISTRO_VENCIMIENTO_MANUAL: '🚨 Vencimiento manual'
      }[r.tipo] || r.tipo;

      const cardCls = isPos ? 'ajuste-pos' : isNeg ? 'ajuste-neg' : '';
      const chipCls = isPos ? 'pos' : isNeg ? 'neg' : '';
      const fechaVenc = _extractFechaVenc(r.idVen);
      // Delta: preferimos el campo delta, sino lo calculamos
      const deltaRaw = !isNaN(r.delta) ? r.delta : (!isNaN(r.cantNueva) && !isNaN(r.cantAnterior) ? r.cantNueva - r.cantAnterior : null);
      const deltaStr = deltaRaw !== null
        ? (deltaRaw > 0 ? `+${deltaRaw}` : `${deltaRaw}`) + ' u.'
        : null;
      const deltaChipCls = deltaRaw !== null ? (deltaRaw > 0 ? 'pos' : deltaRaw < 0 ? 'neg' : '') : '';
      return `<div class="hist-op-card ${cardCls}">
        <div class="hist-op-row">
    
          <div class="hist-op-desc">${esc(r.descripcion)}</div>
                    <span class="hist-chip ean">🔖 ${esc(r.ean)}</span>
          ${gramaje ? `<span class="hist-chip">⚖️ ${esc(gramaje)}</span>` : ''}
          ${proveedor ? `<span class="hist-chip">🏢 ${esc(proveedor)}</span>` : ''}
        </div>
        <div class="hist-op-meta">          <span class="hist-chip ean">🔖 ${esc(r.ean)}</span>
          <span class="hist-chip suc-orig">🏪 ${esc(r.sucursal)}</span>
          <span class="hist-chip ${chipCls}">${tipoLabel}</span>
          ${fechaVenc ? `<span class="hist-chip venc">📅 ${fechaVenc}</span>` : ''}
          <span class="hist-chip">Antes: ${isNaN(r.cantAnterior) ? '—' : r.cantAnterior}</span>
          <span class="hist-chip qty">Ahora: ${isNaN(r.cantNueva) ? '—' : r.cantNueva}</span>
          ${deltaStr ? `<span class="hist-chip ${deltaChipCls}" style="font-weight:800">Δ ${deltaStr}</span>` : ''}
             <span class="hist-op-time">${_fmtTime(r.timestamp)}</span>
        </div>
      </div>`;
    }).join('');
  }
}

// ── Render ACC ───────────────────────────────────────
function _renderHistAcc() {
  const today = venLogs.filter(r => _isToday(r.timestamp));

  const ACC_TRANSFER_TYPES = ['ACC_TRANSFER_SALIDA', 'ACC_TRANSFERENCIA_SALIDA'];
  const ACC_AJUSTE_TYPES = ['ACC_AJUSTE_POSITIVO', 'ACC_AJUSTE_NEGATIVO', 'ACC_AJUSTE_A_CERO', 'ACC_AJUSTE_SIN_CAMBIO'];

  // ── Registros del día ─────────────────────────────────────
  const ACC_REG_TYPES = [
    'DEVOLUCION_CON_DESCUENTO_VEN', 'DEVOLUCION_VINCULACION_MANUAL',
    'VYC_INSERT', 'VYC_CON_DESCUENTO_VEN',
    'ACCION_DESDE_VEN', 'REGISTRO_VENCIMIENTO_MANUAL', 'AUTO_VENCIMIENTO_DIARIO'
  ];
  const accRegistros = today.filter(r => ACC_REG_TYPES.includes(r.tipo))
    .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

  document.getElementById('histAccTabBadgeReg').textContent = accRegistros.length;
  document.getElementById('histAccTabBadgeT').textContent = accTransfers.length + accAjustes.length;
  document.getElementById('histAccTabBadgeV').textContent = ventasDedup.length;

  if (_histAccActiveTab === 'registros') {
    if (!accRegistros.length) {
      el.innerHTML = `<div class="hist-empty"><div class="hist-empty-icon">📋</div><div>Sin registros de acciones hoy</div></div>`;
      return;
    }
    el.innerHTML = accRegistros.map(r => {
      const _prod = devData.find(d => String(d.ean || '').trim() === r.ean)
        || venData.find(v => String(v.ean || '').trim() === r.ean) || {};
      const gramaje = _prod.gramaje || '';
      const proveedor = _prod.proveedor || '';

      const tipoLabel = {
        'DEVOLUCION_CON_DESCUENTO_VEN': '📦 Devolución',
        'DEVOLUCION_VINCULACION_MANUAL': '📦 Devolución manual',
        'VYC_INSERT': '💰 Venta / Consumo',
        'VYC_CON_DESCUENTO_VEN': '💰 Venta / Consumo',
        'ACCION_DESDE_VEN': '⚡ Acción desde vencimientos',
        'REGISTRO_VENCIMIENTO_MANUAL': '🚨 Vencimiento manual',
        'AUTO_VENCIMIENTO_DIARIO': '🤖 Auto-vencimiento',
      }[r.tipo] || r.tipo;

      const motivoMatch = r.notas ? r.notas.match(/Motivo[^:]*:\s*([^·\|]+)/) : null;
      const motivo = motivoMatch ? motivoMatch[1].trim() : '';
      const isAuto = r.tipo === 'AUTO_VENCIMIENTO_DIARIO';

      return `<div class="hist-op-card registro-acc">
        <div class="hist-op-row">
          <span class="hist-op-time">${_fmtTime(r.timestamp)}</span>
          <div class="hist-op-desc">${esc(r.descripcion)}</div>
          <span class="hist-chip ${isAuto ? '' : 'pos'}" style="${isAuto ? 'color:var(--text3)' : ''}">${tipoLabel}</span>
        </div>
        <div class="hist-op-meta">
          <span class="hist-chip ean">🔖 ${esc(r.ean)}</span>
          <span class="hist-chip suc-orig">🏪 ${esc(r.sucursal)}</span>
          ${gramaje ? `<span class="hist-chip">⚖️ ${esc(gramaje)}</span>` : ''}
          ${proveedor ? `<span class="hist-chip">🏢 ${esc(proveedor)}</span>` : ''}
          ${motivo ? `<span class="hist-chip motivo">${esc(motivo)}</span>` : ''}
          ${!isNaN(r.cantAnterior) ? `<span class="hist-chip">Antes: ${r.cantAnterior}</span>` : ''}
          ${!isNaN(r.cantNueva) ? `<span class="hist-chip qty">Ahora: ${r.cantNueva}</span>` : ''}
          ${r.idDev ? `<span class="hist-chip" style="color:var(--purple);border-color:rgba(166,124,255,.25);background:rgba(166,124,255,.07)">📋 ${esc(r.idDev)}</span>` : ''}
          ${r.idVen ? `<span class="hist-chip" style="color:var(--cyan);border-color:rgba(34,212,232,.2);background:rgba(34,212,232,.06)">📦 ${esc(r.idVen)}</span>` : ''}
          ${r.usuario ? `<span class="hist-chip" style="color:var(--text3)">👤 ${esc(r.usuario)}</span>` : ''}
        </div>
      </div>`;
    }).join('');
    return;
  }

  const accTransfers = today.filter(r => ACC_TRANSFER_TYPES.includes(r.tipo));
  const accVentas = today.filter(r => r.tipo === 'ACC_MARCADO_VENDIDO');
  const accAjustes = today.filter(r => ACC_AJUSTE_TYPES.includes(r.tipo));

  // Dedup ventas
  const ventaMap = new Map();
  accVentas.forEach(r => ventaMap.set((r.idDev || '') + '|' + r.sucursal, r));
  const ventasDedup = [...ventaMap.values()].sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

  const totalTransfers = accTransfers.length + accAjustes.length;
  document.getElementById('histAccTabBadgeT').textContent = totalTransfers;
  document.getElementById('histAccTabBadgeV').textContent = ventasDedup.length;

  const el = document.getElementById('histAccContent');

  if (_histAccActiveTab === 'transfers') {
    const todos = [...accTransfers, ...accAjustes].sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    if (!todos.length) {
      el.innerHTML = `<div class="hist-empty"><div class="hist-empty-icon">⇄</div><div>Sin transferencias ni ajustes de acciones hoy</div></div>`;
      return;
    }

    el.innerHTML = todos.map(r => {
      const _prod = devData.find(d => String(d.ean || '').trim() === r.ean)
        || venData.find(v => String(v.ean || '').trim() === r.ean) || {};
      const gramaje = _prod.gramaje || '';
      const proveedor = _prod.proveedor || '';

      const esAjuste = ACC_AJUSTE_TYPES.includes(r.tipo);
      const qty = Math.abs(isNaN(r.delta) ? (r.cantAnterior - r.cantNueva) : r.delta);
      const deltaRaw = !isNaN(r.delta) ? r.delta
        : (!isNaN(r.cantNueva) && !isNaN(r.cantAnterior) ? r.cantNueva - r.cantAnterior : null);
      const deltaStr = deltaRaw !== null
        ? (deltaRaw > 0 ? '+' + deltaRaw : String(deltaRaw)) + ' u.' : null;
      const deltaChipCls = deltaRaw !== null ? (deltaRaw > 0 ? 'pos' : deltaRaw < 0 ? 'neg' : '') : '';

      const motivoMatch = r.notas ? r.notas.match(/Motivo[^:]*:\s*([^·\|]+)/) : null;
      const motivo = motivoMatch ? motivoMatch[1].trim() : '';

      if (esAjuste) {
        const tipoLabel = {
          ACC_AJUSTE_POSITIVO: '＋ Ajuste positivo',
          ACC_AJUSTE_NEGATIVO: '－ Ajuste negativo',
          ACC_AJUSTE_A_CERO: '⬛ A cero',
          ACC_AJUSTE_SIN_CAMBIO: '＝ Sin cambio',
        }[r.tipo] || r.tipo;
        const isPos = r.tipo === 'ACC_AJUSTE_POSITIVO';
        const isNeg = r.tipo === 'ACC_AJUSTE_NEGATIVO' || r.tipo === 'ACC_AJUSTE_A_CERO';
        const chipCls = isPos ? 'pos' : isNeg ? 'neg' : '';

        return `<div class="hist-op-card ${isPos ? 'ajuste-pos' : isNeg ? 'ajuste-neg' : ''}">
          <div class="hist-op-row">
            <span class="hist-op-time">${_fmtTime(r.timestamp)}</span>
            <div class="hist-op-desc">${esc(r.descripcion)}</div>
          </div>
          <div class="hist-op-meta">
            <span class="hist-chip ean">🔖 ${esc(r.ean)}</span>
            <span class="hist-chip suc-orig">🏪 ${esc(r.sucursal)}</span>
            <span class="hist-chip ${chipCls}">${tipoLabel}</span>
            ${gramaje ? `<span class="hist-chip">⚖️ ${esc(gramaje)}</span>` : ''}
            ${proveedor ? `<span class="hist-chip">🏢 ${esc(proveedor)}</span>` : ''}
            <span class="hist-chip">Antes: ${isNaN(r.cantAnterior) ? '—' : r.cantAnterior}</span>
            <span class="hist-chip qty">Ahora: ${isNaN(r.cantNueva) ? '—' : r.cantNueva}</span>
            ${deltaStr ? `<span class="hist-chip ${deltaChipCls}" style="font-weight:800">Δ ${deltaStr}</span>` : ''}
            ${r.idDev ? `<span class="hist-chip" style="color:var(--purple);border-color:rgba(166,124,255,.25);background:rgba(166,124,255,.07)">📋 ${esc(r.idDev)}</span>` : ''}
          </div>
        </div>`;
      }

      // Transferencia
      return `<div class="hist-op-card transfer">
        <div class="hist-op-row">
          <span class="hist-op-time">${_fmtTime(r.timestamp)}</span>
          <div class="hist-op-desc">${esc(r.descripcion)}</div>
        </div>
        <div class="hist-op-meta">
          <span class="hist-chip ean">🔖 ${esc(r.ean)}</span>
          <span class="hist-chip suc-orig">📤 ${esc(r.sucursal)}</span>
          <span class="hist-arrow">→</span>
          <span class="hist-chip suc-dest">📥 ${esc(r.destino || '—')}</span>
          <span class="hist-chip qty">Cant: ${qty} u.</span>
          ${gramaje ? `<span class="hist-chip">⚖️ ${esc(gramaje)}</span>` : ''}
          ${proveedor ? `<span class="hist-chip">🏢 ${esc(proveedor)}</span>` : ''}
          ${motivo ? `<span class="hist-chip">${esc(motivo)}</span>` : ''}
          ${r.idDev ? `<span class="hist-chip" style="color:var(--purple);border-color:rgba(166,124,255,.25);background:rgba(166,124,255,.07)">📋 ${esc(r.idDev)}</span>` : ''}
        </div>
      </div>`;
    }).join('');

  } else {
    // Ventas / ajustes finales
    if (!ventasDedup.length) {
      el.innerHTML = `<div class="hist-empty"><div class="hist-empty-icon">🏷</div><div>Sin ventas ni ajustes hoy</div></div>`;
      return;
    }

    el.innerHTML = ventasDedup.map(r => {
      const _prod = devData.find(d => String(d.ean || '').trim() === r.ean)
        || venData.find(v => String(v.ean || '').trim() === r.ean) || {};
      const gramaje = _prod.gramaje || '';
      const proveedor = _prod.proveedor || '';

      const vendMatch = r.notas ? r.notas.match(/Vendidas:\s*(\d+)/) : null;
      const vencMatch = r.notas ? r.notas.match(/Vencidas en gondola:\s*(\d+)/) : null;
      const motivoMatch = r.notas ? r.notas.match(/Motivo[^:]*:\s*([^|]+)/) : null;
      const vendidas = vendMatch ? vendMatch[1] : '—';
      const vencidas = vencMatch ? vencMatch[1] : '—';
      const motivo = motivoMatch ? motivoMatch[1].trim() : '';
      const base = isNaN(r.cantAnterior) ? '—' : r.cantAnterior;

      return `<div class="hist-op-card venta">
        <div class="hist-op-row">
          <span class="hist-op-time">${_fmtTime(r.timestamp)}</span>
          <div class="hist-op-desc">${esc(r.descripcion)}</div>
        </div>
        <div class="hist-op-meta">
          <span class="hist-chip ean">🔖 ${esc(r.ean)}</span>
          <span class="hist-chip suc-orig">🏪 ${esc(r.sucursal)}</span>
          ${gramaje ? `<span class="hist-chip">⚖️ ${esc(gramaje)}</span>` : ''}
          ${proveedor ? `<span class="hist-chip">🏢 ${esc(proveedor)}</span>` : ''}
          ${motivo ? `<span class="hist-chip">${esc(motivo)}</span>` : ''}
          <span class="hist-chip">Base: ${base}</span>
          <span class="hist-chip pos">✓ Vendidas: ${vendidas}</span>
          <span class="hist-chip neg">⛔ Góndola: ${vencidas}</span>
          ${r.idDev ? `<span class="hist-chip" style="color:var(--purple);border-color:rgba(166,124,255,.25);background:rgba(166,124,255,.07)">📋 ${esc(r.idDev)}</span>` : ''}
        </div>
      </div>`;
    }).join('');
  }
}

window.onload = () => {
  showPanel('resumen');
  if (SCRIPT_URL) {
    document.getElementById('configModal').classList.remove('open');
    loadData();
  } else {
    document.getElementById('configModal').classList.add('open');
  }
  _updateThemeBtn(localStorage.getItem('nexus_theme') || 'dark');
};

// ══════════════════════════════════════════════════════
//  MODAL REPLICAR REGISTRO
// ══════════════════════════════════════════════════════
function abrirModalReplicar(idVen, desc, ean, sucOrigen, fechaVenc, gramaje) {
  _currentReplicarData = { idVen, desc, ean, sucOrigen, fechaVenc, gramaje };

  const descEl = document.getElementById('repRegDesc');
  if (descEl) descEl.textContent = desc || '—';

  const infoEl = document.getElementById('repRegInfo');
  if (infoEl) {
    infoEl.innerHTML = [
      ean ? `<span>EAN: <strong>${esc(ean)}</strong></span>` : '',
      gramaje ? `<span>⚖️ ${esc(gramaje)}</span>` : '',
      fechaVenc ? `<span>📅 Vence: <strong>${fechaVenc}</strong></span>` : '',
      sucOrigen ? `<span class="hist-chip origen-suc">Origen: ${esc(sucOrigen)}</span>` : '',
    ].filter(Boolean).join('<span style="color:var(--text3)"> · </span>');
  }

  const dateEl = document.getElementById('repRegDate');
  if (dateEl) dateEl.textContent = new Date().toLocaleDateString('es-AR', { weekday: 'long', day: 'numeric', month: 'long' }).toUpperCase();

  // Reset campos
  const sucEl = document.getElementById('repRegSuc');
  const cantEl = document.getElementById('repRegCant');
  const nombEl = document.getElementById('repRegNombre');
  const loteEl = document.getElementById('repRegLote');
  const unitEl = document.getElementById('repRegCantUnit');

  if (sucEl) sucEl.value = '';
  if (cantEl) cantEl.value = '';
  if (nombEl) nombEl.value = '';
  if (loteEl) loteEl.value = '';
  if (unitEl) unitEl.textContent = gramaje ? `(kg)` : `(unidades)`;

  // Ocultar errores
  ['repRegSucErr', 'repRegCantErr', 'repRegNombreErr'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });

  document.getElementById('modalReplicarReg').classList.add('open');
  setTimeout(() => { if (sucEl) sucEl.focus(); }, 150);
}

function cerrarModalReplicar(evt) {
  if (evt && evt.target !== document.getElementById('modalReplicarReg')) return;
  document.getElementById('modalReplicarReg').classList.remove('open');
  _currentReplicarData = null;
}

async function ejecutarReplicar() {
  if (!_currentReplicarData) return;
  const { idVen, desc, ean, fechaVenc, gramaje } = _currentReplicarData;

  const sucEl = document.getElementById('repRegSuc');
  const cantEl = document.getElementById('repRegCant');
  const nombEl = document.getElementById('repRegNombre');
  const loteEl = document.getElementById('repRegLote');

  const suc = sucEl?.value || '';
  const cant = parseFloat((cantEl?.value || '').replace(',', '.'));
  const nombre = (nombEl?.value || '').trim().toUpperCase();

  // Validaciones
  let ok = true;
  if (!suc) {
    document.getElementById('repRegSucErr').style.display = 'block';
    ok = false;
  }
  if (!cant || cant <= 0 || isNaN(cant)) {
    document.getElementById('repRegCantErr').style.display = 'block';
    ok = false;
  }
  if (!nombre || nombre.length < 2) {
    document.getElementById('repRegNombreErr').style.display = 'block';
    ok = false;
  }
  if (!ok) return;

  // Cerrar todos los modales abiertos antes de mostrar el loader
  [
    'modalReplicarReg',
    'histVenModal',
    'histAccModal',
    'devDetailModal',
    'configModal',
    'confirmModal'
  ].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove('open');
    if (el && el.style.display) el.style.display = '';
  });
  document.body.style.overflow = '';
  showSpinner('Replicando registro…');

  try {
    // Buscar el VEN original para obtener todos los datos del producto
    const venOrig = venData.find(v => v.id === idVen) || {};

    const body = {
      action: 'submitForm',
      tipoRegistro: 'CONTROL DE VENCIMIENTO',
      sucursal: suc,
      usuario: nombre,
      ean: ean,
      fechaVenc: fechaVenc,
      descripcion: desc,
      gramaje: gramaje || venOrig.gramaje || '',
      codInterno: venOrig.codInterno || '',
      sector: venOrig.sector || '',
      seccion: venOrig.seccion || '',
      proveedor: venOrig.proveedor || '',
      cantidad: cant,
      lote: loteEl?.value?.trim() || venOrig.lote || '',
      esReplica: true,
    };

    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body),
    });
    const json = await res.json();

    if (json.success) {
      // Cerrar ambos modales
      document.getElementById('modalReplicarReg').classList.remove('open');
      document.getElementById('histVenModal').classList.remove('open');
      document.body.style.overflow = '';
      _currentReplicarData = null;

      showToast(true, `✓ Replicado en ${suc} · ${cant} u. · ${json.action === 'updated' ? 'Sumado a stock existente' : 'Nuevo registro creado'}`);
      await loadData();

    } else if (json.limitReached) {
      showToast(false, json.message || 'Límite diario alcanzado para esta sucursal.');
    } else {
      showToast(false, json.message || 'Error al replicar');
    }
  } catch (e) {
    showToast(false, 'Error de red al replicar');
    console.error(e);
  }
  hideSpinner();
}

// ══════════════════════════════════════════════════════
//  NAV
// ══════════════════════════════════════════════════════
const PANEL_TITLES = {
  resumen: 'Resumen General', vencimientos: 'Control de Vencimientos',
  acciones: 'Registros de Acciones', metricas: 'Métricas Compras', nc: 'Gestión N/C'
};
function showPanel(id) {
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.sb-item').forEach(b => b.classList.remove('active'));
  document.getElementById('panel-' + id).classList.add('active');
  const nav = document.getElementById('nav-' + id);
  if (nav) nav.classList.add('active');
  document.getElementById('topbarTitle').textContent = PANEL_TITLES[id] || id;
  closeSidebar();
  // Toggle sticky history buttons
  const btnVen = document.getElementById('histVenBtn');
  const btnAcc = document.getElementById('histAccBtn');
  if (btnVen) btnVen.classList.toggle('visible', id === 'vencimientos');
  if (btnAcc) btnAcc.classList.toggle('visible', id === 'acciones');
}
function toggleSidebar() { document.getElementById('sidebar').classList.toggle('open'); document.getElementById('backdrop').classList.toggle('open'); }
function closeSidebar() { document.getElementById('sidebar').classList.remove('open'); document.getElementById('backdrop').classList.remove('open'); }

function goVencFilter(urg) { document.getElementById('vf-urg').value = urg || ''; applyVencFilters(); showPanel('vencimientos'); }
function filterAutoVencidos() {
  document.getElementById('af-motivo').value = 'VENCIDO';
  setAccPreset('all', document.querySelectorAll('#acc-presets .dp')[6]);
  showPanel('acciones');
}
function filterAccEstado(est) { document.getElementById('af-estado').value = est || ''; setAccPreset('all', document.querySelectorAll('#acc-presets .dp')[6]); showPanel('acciones'); }

// ══════════════════════════════════════════════════════
//  AUTO-REFRESH
// ══════════════════════════════════════════════════════
function startAutoRefresh() {
  if (_autoTimer) clearTimeout(_autoTimer);
  if (_cdTimer) clearInterval(_cdTimer);
  _nextRefresh = Date.now() + AUTO_MS;
  _autoTimer = setTimeout(loadData, AUTO_MS);
  _cdTimer = setInterval(() => {
    const rem = _nextRefresh - Date.now();
    if (rem <= 0) { clearInterval(_cdTimer); return; }
    const m = Math.floor(rem / 60000), s = Math.floor((rem % 60000) / 1000);
    const el = document.getElementById('syncTime');
    if (el) el.textContent = (el.dataset.base || '') + ' · 🔄 ' + m + ':' + String(s).padStart(2, '0');
  }, 1000);
}

// ══════════════════════════════════════════════════════
//  DATA LOADING
// ══════════════════════════════════════════════════════
async function loadData() {
  if (!SCRIPT_URL) { openConfig(); return; }
  const btn = document.getElementById('syncBtn');
  btn.classList.add('loading');
  document.getElementById('syncIcon').textContent = '⏳';
  showSpinner('Cargando datos…');
  if (_autoTimer) clearTimeout(_autoTimer);
  if (_cdTimer) clearInterval(_cdTimer);

  try {
    const [devRes, venRes, logRes] = await Promise.all([
      fetch(`${SCRIPT_URL}?action=getHistorial&bd=devoluciones&_t=${Date.now()}`),
      fetch(`${SCRIPT_URL}?action=getHistorial&bd=vencimientos&_t=${Date.now()}`),
      fetch(`${SCRIPT_URL}?action=getLogs&bd=vencimientos&_t=${Date.now()}`).catch(() => null)
    ]);
    const [devJson, venJson] = await Promise.all([devRes.json(), venRes.json()]);
    if (!devJson.success) throw new Error(devJson.message || 'Error BD Devoluciones');
    if (!venJson.success) throw new Error(venJson.message || 'Error BD Vencimientos');
    processDevData(devJson.data || []);
    processVenData(venJson.data || []);
    // Logs (graceful: si el endpoint no existe, no falla todo)
    if (logRes) {
      try {
        const logJson = await logRes.json();
        if (logJson?.success) processVenLogs(logJson.data || []);
      } catch (_) { }
    }
    updateAll();
    const timeStr = 'Actualizado: ' + new Date().toLocaleTimeString('es-AR');
    const el = document.getElementById('syncTime');
    el.dataset.base = timeStr;
    el.textContent = timeStr;
    document.getElementById('statusDot').className = 'status-dot online';
    startAutoRefresh();
  } catch (err) {
    const el = document.getElementById('syncTime');
    el.dataset.base = '⚠️ Error';
    el.textContent = '⚠️ Error de conexión';
    document.getElementById('statusDot').className = 'status-dot';
    console.error(err);
    startAutoRefresh();
  } finally {
    hideSpinner();
    btn.classList.remove('loading');
    document.getElementById('syncIcon').textContent = '🔄';
  }
}

function processDevData(raw) {
  devData = raw.map(r => {
    // CANT_DISPONIBLE: usa el campo nuevo si existe, sino cae a CANTIDAD
    const cdRaw = r['CANT_DISPONIBLE'];
    const cantDisponible = (cdRaw !== undefined && cdRaw !== '' && cdRaw !== null)
      ? (parseFloat(String(cdRaw).replace(',', '.')) || 0)
      : (parseFloat(String(r['CANTIDAD'] || '0').replace(',', '.')) || 0);

    return {
      id: r['ID'] || '',
      fecha: parseDate(r['FECHA REGISTRO'] || r['FECHA'] || ''),
      sucursal: r['SUCURSAL'] || '',
      usuario: r['USUARIO'] || '',
      ean: String(r['EAN'] || ''),
      codInterno: r['COD. INTERNO'] || '',
      descripcion: r['DESCRIPCION'] || '',
      gramaje: r['GRAMAJE'] || '',
      cantidad: r['CANTIDAD'] || '',
      cantDisponible,
      fechaVenc: r['FECHA VENC.'] || '',
      sector: r['SECTOR'] || '',
      seccion: r['SECCION'] || '',
      proveedor: r['PROVEEDOR'] || '',
      codProv: r['COD. PROVEEDOR'] || '',
      motivo: r['MOTIVO'] || '',
      lote: r['LOTE'] || '',
      aclaracion: r['ACLARACION'] || '',
      comentarios: r['COMENTARIOS'] || '',
      fotoRaw: r['ARCHIVO ADJUNTO'] || '',
      estado: r['ESTADO'] || 'PENDIENTE',
      obsNC: r['OBSERVACION N/C'] || '',
      idOrigen: r['ID_ORIGEN'] || '',
      sucOrigenTransfer: r['SUC_ORIGEN'] || '',
      cantVendida: parseFloat(String(r['CANT_VENDIDA'] || '0').replace(',', '.')) || 0,
      cantVencidaGondola: parseFloat(String(r['CANT_VENCIDA_GONDOLA'] || '0').replace(',', '.')) || 0,
    };
  }).filter(r => r.id);
  devMap.clear();
  devData.forEach(r => devMap.set(r.id, r));
  document.getElementById('recCount').textContent = devData.length + ' acciones · ' + venData.length + ' vencimientos';
}

function processVenData(raw) {
  venData = raw.map(r => {
    // Soporta tanto headers con espacios ("FECHA VENCIMIENTO") como con guión bajo ("FECHA_VENC")
    // para ser compatible con distintas configuraciones del sheet VEN.
    const fv = r['FECHA VENCIMIENTO'] || r['FECHA VENC.'] || r['FECHA_VENC'] || '';
    const dias = calcDias(fv);
    return {
      id: r['ID'] || '',
      fechaReg: parseRegDate(r['FECHA REGISTRO'] || r['FECHA_REG'] || ''),
      sucursal: String(r['SUCURSAL'] || ''),
      usuario: r['USUARIO'] || '',
      ean: String(r['EAN'] || ''),
      codInterno: r['COD. INTERNO'] || r['COD_INTERNO'] || '',
      descripcion: r['DESCRIPCION'] || '',
      gramaje: r['GRAMAJE'] || '',
      cantidad: r['CANTIDAD'] || '',
      fechaVenc: fv,
      sector: r['SECTOR'] || '',
      seccion: r['SECCION'] || '',
      proveedor: r['PROVEEDOR'] || '',
      lote: r['LOTE'] || '',
      aclaracion: r['ACLARACION'] || '',
      estadoGest: String(r['ESTADO GESTION'] || r['ESTADO_GEST'] || 'ACTIVO').toUpperCase(),
      idAccion: r['ID ACCION'] || r['ID_ACCION'] || '',
      _dias: dias,
      _urg: getUrg(dias),
    };
  });
  buildVenGrupos();
  document.getElementById('recCount').textContent = devData.length + ' acciones · ' + venData.length + ' vencimientos';
}

function updateAll() {
  populateFilterSelects();
  updateResumenKPIs();
  renderAlertFeed();
  applyVencFilters();
  setAccPreset(accPreset, document.querySelector('#acc-presets .dp.active'));
  renderMetricasPanel();
  renderProvTable();
  applyNcFilters();
  updateNavBadges();
}

// ══════════════════════════════════════════════════════
//  URGENCIA
// ══════════════════════════════════════════════════════
function calcDias(str) {
  if (!str) return null;
  const s = String(str).trim();
  let d;
  if (/^\d{2}[-\/]\d{2}[-\/]\d{4}/.test(s)) {
    const [dd, mm, yyyy] = s.split(/[-\/]/);
    d = new Date(+yyyy, +mm - 1, +dd);
  } else if (/^\d{4}[-\/]\d{2}[-\/]\d{2}/.test(s)) {
    d = new Date(s.slice(0, 10));
  } else {
    d = new Date(s);
  }
  if (isNaN(d)) return null;
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0); d.setHours(0, 0, 0, 0);
  // +1: se cuenta el día de vencimiento como plazo válido
  return Math.round((d - hoy) / 86400000) + 1;
}

function getUrg(dias) {
  if (dias === null) return 'NORMAL';
  if (dias <= 0) return 'VENCIDO';   // venció ayer o antes
  if (dias <= 7) return 'CRITICO';   // vence hoy…7 días  (antes: ≤6)
  if (dias <= 15) return 'URGENTE';   // 8–15 días          (antes: ≤14)
  if (dias <= 22) return 'PROXIMO';   // 16–22 días         (antes: ≤21)
  if (dias <= 46) return 'ATENCION';  // 23–46 días         (antes: ≤45)
  return 'NORMAL';
}

// ══════════════════════════════════════════════════════
//  BUILD VEN GRUPOS
// ══════════════════════════════════════════════════════
function buildVenGrupos() {
  const sucMap = new Map();
  venData.forEach(r => {
    const k = r.ean + '||' + r.fechaVenc + '||' + r.sucursal;
    if (!sucMap.has(k)) sucMap.set(k, { ean: r.ean, fechaVenc: r.fechaVenc, suc: r.sucursal, registros: [] });
    sucMap.get(k).registros.push(r);
  });

  const sucEntries = [];
  sucMap.forEach(e => {
    e.registros.sort((a, b) => {
      if (!a.fechaReg && !b.fechaReg) return 0;
      if (!a.fechaReg) return 1;
      if (!b.fechaReg) return -1;
      return b.fechaReg - a.fechaReg;
    });
    e.latest = e.registros[0];
    const cantActual = parseFloat(String(e.latest.cantidad || 0).replace(',', '.')) || 0;
    if (cantActual <= 0) return;
    e.controles = e.registros;
    e.diasDesde = diasDesde(e.latest.fechaReg);
    sucEntries.push(e);
  });

  const map = new Map();
  sucEntries.forEach(se => {
    const k = se.ean + '||' + se.fechaVenc;
    if (!map.has(k)) map.set(k, { key: k, ean: se.ean, fechaVenc: se.fechaVenc, sucursales: [] });
    map.get(k).sucursales.push(se);
  });

  venGrupos = [...map.values()]
    .filter(g => g.sucursales.length > 0)
    .map(g => {
      g.dias = g.sucursales[0].latest._dias;
      g.urg = getUrg(g.dias);
      g.desc = g.sucursales.reduce((b, se) => se.latest.descripcion || b, '—');
      g.prov = g.sucursales.reduce((b, se) => se.latest.proveedor || b, '');
      g.gramaje = g.sucursales.reduce((b, se) => se.latest.gramaje || b, '');
      g.totalU = g.sucursales.reduce((s, se) => s + (parseFloat(String(se.latest.cantidad || 0).replace(',', '.')) || 0), 0);
      const ests = g.sucursales.map(se => se.latest.estadoGest.replace(/\s+/g, '_'));
      if (ests.some(e => e === 'ACTIVO')) g.grupoEstado = 'ACTIVO';
      else if (ests.every(e => e === 'RETIRADO')) g.grupoEstado = 'RETIRADO';
      else g.grupoEstado = 'ACCION_TOMADA';
      g.maxDiasDesde = Math.max(...g.sucursales.map(se => se.diasDesde ?? 0));
      g.hasAccion = g.sucursales.some(se => {
        if (se.latest.idAccion) return true;
        const eanG = String(se.latest.ean || '').trim();
        const fvG = String(se.latest.fechaVenc || '').trim().slice(0, 10);
        const sucG = (se.suc || '').toUpperCase().trim();
        return devData.some(d =>
          String(d.ean || '').trim() === eanG &&
          String(d.fechaVenc || '').trim().slice(0, 10) === fvG &&
          (d.sucursal || '').toUpperCase().trim() === sucG
        );
      });
      return g;
    });
}

function diasDesde(date) {
  if (!date) return null;
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const d = new Date(date); d.setHours(0, 0, 0, 0);
  return Math.max(0, Math.round((hoy - d) / 86400000));
}

// ══════════════════════════════════════════════════════
//  RESUMEN PANEL
// ══════════════════════════════════════════════════════
function updateResumenKPIs() {
  const activos = venGrupos.filter(g => g.grupoEstado === 'ACTIVO');
  const vc = { VENCIDO: 0, CRITICO: 0, URGENTE: 0, PROXIMO: 0, ATENCION: 0, NORMAL: 0 };
  activos.forEach(g => vc[g.urg] = (vc[g.urg] || 0) + 1);
  document.getElementById('rv-total').textContent = activos.length;
  document.getElementById('rv-vencido').textContent = vc.VENCIDO;
  document.getElementById('rv-critico').textContent = vc.CRITICO;
  document.getElementById('rv-urgente').textContent = vc.URGENTE;
  document.getElementById('rv-proximo').textContent = vc.PROXIMO;
  document.getElementById('rv-atencion').textContent = vc.ATENCION;
  document.getElementById('rv-normal').textContent = vc.NORMAL;
  document.getElementById('ra-total').textContent = devData.length;
  document.getElementById('ra-nc').textContent = devData.filter(r => r.estado === 'N/C RECIBIDA').length;
  document.getElementById('ra-pend').textContent = devData.filter(r => r.estado === 'PENDIENTE').length;
  document.getElementById('ra-gest').textContent = devData.filter(r => r.estado === 'EN GESTION').length;
  document.getElementById('ra-rech').textContent = devData.filter(r => r.estado === 'RECHAZADA').length;
  const vinc = venData.filter(r => r.idAccion && devMap.has(r.idAccion)).length;
  document.getElementById('ra-vinc').textContent = vinc;
  const autoV = devData.filter(r => r.usuario === 'SISTEMA AUTO').length;
  const elAutoV = document.getElementById('ra-autov'); if (elAutoV) elAutoV.textContent = autoV;
}

function renderAlertFeed() {
  const critical = venGrupos.filter(g => g.grupoEstado === 'ACTIVO' && (g.urg === 'VENCIDO' || g.urg === 'CRITICO' || g.urg === 'URGENTE')).sort((a, b) => URG_ORDER[a.urg] - URG_ORDER[b.urg] || (a.dias ?? 99) - (b.dias ?? 99)).slice(0, 12);
  const el = document.getElementById('alertFeed');
  if (!critical.length) { el.innerHTML = '<div class="empty"><div class="ei">✅</div><p>Sin alertas críticas activas</p></div>'; return; }
  el.innerHTML = '<div class="alert-feed">' + critical.map(g => `
    <div class="alert-item ${g.urg}" onclick="goVencFilter('${g.urg}')">
      <div class="ai-top">
        <span class="ai-desc" title="${esc(g.desc)}">${esc(g.desc)}</span>
        <span class="urg ${g.urg}">${g.dias !== null ? (g.dias <= 0 ? 'VENCIDO' : g.dias + 'd') : ''} ${URG_LABELS[g.urg]}</span>
      </div>
      <div class="ai-meta">${esc(g.ean)} · ${esc(g.prov || '—')} · ${g.sucursales.length} suc.</div>
      <div class="ai-badges">
        ${g.sucursales.map(se => `<span class="suc-b ${(se.latest.estadoGest || 'ACTIVO').replace(/\s+/g, '-')}">${esc(se.suc)}</span>`).join('')}
        ${g.hasAccion ? `<span class="linked-indicator">🔗 acción vinculada</span>` : ''}
      </div>
    </div>`).join('') + '</div>';
}

function updateNavBadges() {
  const critCount = venGrupos.filter(g => g.grupoEstado === 'ACTIVO' && (g.urg === 'VENCIDO' || g.urg === 'CRITICO')).length;
  const pendCount = devData.filter(r => r.estado === 'PENDIENTE').length;
  const critBadge = document.getElementById('nav-badge-crit');
  const pendBadge = document.getElementById('nav-badge-pend');
  if (critCount > 0) { critBadge.textContent = critCount; critBadge.style.display = ''; } else critBadge.style.display = 'none';
  if (pendCount > 0) { pendBadge.textContent = pendCount; pendBadge.style.display = ''; } else pendBadge.style.display = 'none';
}

// ══════════════════════════════════════════════════════
//  VENCIMIENTOS — FILTROS DE FECHA
// ══════════════════════════════════════════════════════
function setVencDateMode(mode) {
  vencDateMode = mode;
  const btnReg = document.getElementById('vf-mode-reg');
  const btnVenc = document.getElementById('vf-mode-venc');
  const label = document.getElementById('vf-period-label');
  if (mode === 'registro') {
    btnReg.style.background = 'var(--blue)';
    btnReg.style.color = '#fff';
    btnVenc.style.background = 'var(--surface)';
    btnVenc.style.color = 'var(--text3)';
    if (label) label.textContent = 'Período — fecha de registro';
  } else {
    btnVenc.style.background = 'var(--blue)';
    btnVenc.style.color = '#fff';
    btnReg.style.background = 'var(--surface)';
    btnReg.style.color = 'var(--text3)';
    if (label) label.textContent = 'Período — fecha de vencimiento';
  }
  applyVencFilters();
}

function setVencPreset(preset, btn) {
  vencPreset = preset;
  document.querySelectorAll('#venc-presets .dp').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');
  const fromWrap = document.getElementById('vf-date-from-wrap');
  const toWrap = document.getElementById('vf-date-to-wrap');
  fromWrap.style.display = preset === 'custom' ? '' : 'none';
  toWrap.style.display = preset === 'custom' ? '' : 'none';
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  if (preset === 'all') {
    vencDateFrom = null; vencDateTo = null;
  } else if (preset === 'today') {
    vencDateFrom = new Date(hoy);
    vencDateTo = new Date(hoy); vencDateTo.setHours(23, 59, 59, 999);
  } else if (preset === 'yesterday') {
    const y = new Date(hoy); y.setDate(y.getDate() - 1);
    vencDateFrom = y;
    vencDateTo = new Date(y); vencDateTo.setHours(23, 59, 59, 999);
  } else if (preset === 'custom') {
    vencDateFrom = null; vencDateTo = null;
  } else {
    const days = { '3d': 3, '7d': 7, '14d': 14, '30d': 30 }[preset] ?? 0;
    vencDateFrom = new Date(hoy); vencDateFrom.setDate(vencDateFrom.getDate() - days);
    vencDateTo = new Date(); vencDateTo.setHours(23, 59, 59, 999);
  }
  applyVencFilters();
}

// ══════════════════════════════════════════════════════
//  VENCIMIENTOS — FILTRADO PRINCIPAL
// ══════════════════════════════════════════════════════
function filterByUrg(urg) {
  document.getElementById('vf-urg').value = urg || '';
  applyVencFilters();
}

function applyVencFilters() {
  const buscar = (document.getElementById('vf-buscar').value || '').toLowerCase().trim();
  const prov = document.getElementById('vf-prov').value;
  const suc = document.getElementById('vf-suc').value;
  const urg = document.getElementById('vf-urg').value;
  const estado = document.getElementById('vf-estado').value;
  let from = vencDateFrom, to = vencDateTo;
  if (vencPreset === 'custom') {
    const f = document.getElementById('vf-date-from').value;
    const t = document.getElementById('vf-date-to').value;
    from = f ? new Date(f) : null;
    to = t ? new Date(t + 'T23:59:59') : null;
  }
  filteredVen = venGrupos.filter(g => {
    if (buscar) {
      const hay = [g.ean, g.desc, g.prov,
      ...g.sucursales.flatMap(se => se.controles.map(r => r.lote || ''))
      ].join(' ').toLowerCase();
      if (!hay.includes(buscar)) return false;
    }
    if (prov && !g.sucursales.some(se => se.latest.proveedor === prov)) return false;
    if (suc && !g.sucursales.some(se => se.suc === suc)) return false;
    if (urg && g.urg !== urg) return false;
    if (estado) {
      if (estado === 'ACTIVO' && g.grupoEstado !== 'ACTIVO') return false;
      if (estado === 'ACCION TOMADA' && g.grupoEstado !== 'ACCION_TOMADA') return false;
      if (estado === 'RETIRADO' && g.grupoEstado !== 'RETIRADO') return false;
    }
    if (from || to) {
      let pasa = false;
      if (vencDateMode === 'registro') {
        pasa = g.sucursales.some(se =>
          se.controles.some(r => {
            if (!r.fechaReg) return false;
            const d = r.fechaReg;
            if (from && d < from) return false;
            if (to && d > to) return false;
            return true;
          })
        );
      } else {
        const fv = parseFechaVenc(g.fechaVenc);
        if (!fv) return false;
        if (from && fv < from) return false;
        if (to && fv > to) return false;
        pasa = true;
      }
      if (!pasa) return false;
    }
    return true;
  });
  filteredVen.sort((a, b) => {
    const d = URG_ORDER[a.urg] - URG_ORDER[b.urg];
    return d !== 0 ? d : (a.dias ?? 9999) - (b.dias ?? 9999);
  });
  filteredVenProd = regroupByEan(filteredVen);
  renderVencTable();
}

function parseFechaVenc(str) {
  if (!str) return null;
  const s = String(str).trim();
  let m;
  m = s.match(/^(\d{2})[-\/](\d{2})[-\/](\d{4})/);
  if (m) { const d = new Date(+m[3], +m[2] - 1, +m[1]); d.setHours(0, 0, 0, 0); return d; }
  m = s.match(/^(\d{4})[-\/](\d{2})[-\/](\d{2})/);
  if (m) { const d = new Date(+m[1], +m[2] - 1, +m[3]); d.setHours(0, 0, 0, 0); return d; }
  return null;
}

function clearVencFilters() {
  ['vf-buscar', 'vf-prov', 'vf-suc', 'vf-urg', 'vf-estado'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  const allBtn = document.querySelector('#venc-presets .dp');
  setVencPreset('all', allBtn);
  setVencDateMode('registro');
}

function regroupByEan(groups) {
  const map = new Map();
  groups.forEach(g => {
    if (!map.has(g.ean)) map.set(g.ean, { ean: g.ean, desc: g.desc, prov: g.prov, gramaje: g.gramaje || '', vencGrupos: [] });
    map.get(g.ean).vencGrupos.push(g);
  });
  return [...map.values()].map(p => {
    p.vencGrupos.sort((a, b) => URG_ORDER[a.urg] - URG_ORDER[b.urg] || (a.dias ?? 9999) - (b.dias ?? 9999));
    p.worstUrg = p.vencGrupos[0]?.urg ?? 'NORMAL';
    p.worstDias = p.vencGrupos[0]?.dias ?? null;
    p.allSucs = [...new Set(p.vencGrupos.flatMap(g => g.sucursales.map(s => s.suc)))];
    p.totalU = p.vencGrupos.reduce((s, g) => s + (g.totalU || 0), 0);
    return p;
  }).sort((a, b) => URG_ORDER[a.worstUrg] - URG_ORDER[b.worstUrg] || (a.worstDias ?? 9999) - (b.worstDias ?? 9999));
}

function getSucClass(suc) {
  const s = (suc || '').toUpperCase();
  if (s.includes('HIPER')) return 'suc-hiper';
  if (s.includes('CENTRO')) return 'suc-centro';
  if (s.includes('RIBERA')) return 'suc-ribera';
  if (s.includes('MAYORISTA')) return 'suc-mayorista';
  return 'suc-default';
}
function getSucColorVar(suc) {
  const s = (suc || '').toUpperCase();
  if (s.includes('HIPER')) return '#ffd166';
  if (s.includes('CENTRO')) return '#60a5fa';
  if (s.includes('RIBERA')) return '#f87171';
  if (s.includes('MAYORISTA')) return '#c084fc';
  return '#475569';
}

function toggleProd(row, ean) {
  const opening = !openProds.has(ean);
  if (opening) openProds.add(ean); else openProds.delete(ean);
  row.classList.toggle('open', opening);
  document.querySelectorAll('tr.venc-grupo-row, tr.suc-row, tr.ctrl-row').forEach(r => {
    if (r.dataset.prodEan !== ean) return;
    if (!opening) {
      r.classList.remove('visible', 'open', 'suc-open');
      if (r.dataset.grupoKey) openGrupos.delete(r.dataset.grupoKey);
      if (r.dataset.sucKey) openSucs.delete(r.dataset.sucKey);
    } else {
      if (r.classList.contains('venc-grupo-row')) r.classList.add('visible');
    }
  });
}

// ══════════════════════════════════════════════════════
//  VENCIMIENTOS — RENDER TABLE
// ══════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════
//  VENCIMIENTOS — RENDER TABLE
// ══════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════
//  VENCIMIENTOS — RENDER TABLE  (reemplaza la función completa)
// ══════════════════════════════════════════════════════
function renderVencTable() {
  const wrap = document.getElementById('vencTableWrap');
  const sinAccion = filteredVen.filter(g => g.grupoEstado === 'ACTIVO').length;
  const conAccion = filteredVen.filter(g => g.grupoEstado !== 'ACTIVO').length;
  document.getElementById('venc-count').innerHTML =
    `<strong>${filteredVenProd.length}</strong> producto${filteredVenProd.length !== 1 ? 's' : ''}`;
  document.getElementById('venc-detail').innerHTML =
    `⚠ <strong>${sinAccion}</strong> grupos sin acción · ✓ <strong>${conAccion}</strong> gestionados`;

  if (!filteredVenProd.length) {
    wrap.innerHTML = `
      <div class="table-head-bar">
        <h3>⏰ Vencimientos</h3>
        <button class="btn btn-secondary btn-sm" onclick="exportVencXlsx()">⬇ Exportar Excel</button>
      </div>
      <div class="empty"><div class="ei">🔍</div><p>Sin resultados con los filtros actuales</p></div>`;
    return;
  }

  // 8 columnas: +1 "Acciones vinculadas" al final
  const NCOLS = 8;

  const headBar = `
    <div class="table-head-bar">
      <h3>⏰ Vencimientos <span style="font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--text3);font-weight:400;margin-left:8px">${filteredVenProd.length} productos · ${filteredVen.length} grupos</span></h3>
      <button class="btn btn-secondary btn-sm" onclick="exportVencXlsx()" title="Exportar tabla filtrada a Excel">⬇ Exportar Excel</button>
    </div>`;

  // ─── DESKTOP TABLE ───────────────────────────────────
  let tableHtml = `
    <div class="venc-desktop">
    <div class="table-scroll"><table class="venc-main-table"><thead><tr>
      <th style="width:180px">
        <div>⏰ Urgencia</div>
        <div style="font-size:8px;font-weight:400;color:var(--text3);margin-top:2px;letter-spacing:.04em;text-transform:none">días al vencimiento más crítico</div>
      </th>
      <th>
        <div>EAN</div>
        <div style="font-size:8px;font-weight:400;color:var(--text3);margin-top:2px;letter-spacing:.04em;text-transform:none">código de barras</div>
      </th>
      <th>
        <div>📦 Descripción</div>
        <div style="font-size:8px;font-weight:400;color:var(--text3);margin-top:2px;letter-spacing:.04em;text-transform:none">nombre del producto</div>
      </th>
      <th class="c-right">
        <div>📊 Total u.</div>
        <div style="font-size:8px;font-weight:400;color:var(--text3);margin-top:2px;letter-spacing:.04em;text-transform:none">stock en todas las sucursales</div>
      </th>
      <th>
        <div>🏭 Proveedor</div>
        <div style="font-size:8px;font-weight:400;color:var(--text3);margin-top:2px;letter-spacing:.04em;text-transform:none">empresa / marca</div>
      </th>
      <th>
        <div>🏪 Sucursales</div>
        <div style="font-size:8px;font-weight:400;color:var(--text3);margin-top:2px;letter-spacing:.04em;text-transform:none">donde está el stock</div>
      </th>
      <th>
        <div>📅 Fechas de vencimiento</div>
        <div style="font-size:8px;font-weight:400;color:var(--text3);margin-top:2px;letter-spacing:.04em;text-transform:none">todos los lotes detectados</div>
      </th>
      <th>
        <div>🔗 Acciones</div>
        <div style="font-size:8px;font-weight:400;color:var(--text3);margin-top:2px;letter-spacing:.04em;text-transform:none">vinculadas al producto</div>
      </th>
    </tr></thead><tbody>`;

  filteredVenProd.forEach(prod => {
    const isProdOpen = openProds.has(prod.ean);
    const barPct = prod.worstDias === null ? 0
      : Math.max(0, Math.min(100, Math.round(prod.worstDias / 90 * 100)));
    const sucBadges = prod.allSucs
      .map(s => `<span class="suc-b ${getSucClass(s)}">${esc(s)}</span>`).join('');
    const vencMini = prod.vencGrupos.map(g =>
      `<span class="urg ${g.urg}" style="font-size:9px;padding:2px 5px">${fmtDateOnly(g.fechaVenc)} ${g.dias !== null ? (g.dias <= 0 ? '⚠' : g.dias + 'd') : ''}</span>`
    ).join(' ');
    const prodTotal = prod.vencGrupos.reduce((sum, g) =>
      sum + g.sucursales.reduce((s2, se) =>
        s2 + (parseFloat(String(se.latest.cantidad || 0).replace(',', '.')) || 0), 0), 0);

    // ── Acumular TODAS las acciones vinculadas al producto (todas las sucursales, todos los grupos) ──
    const seenDevIds = new Set();
    const allLinkedDevs = [];
    prod.vencGrupos.forEach(g => {
      g.sucursales.forEach(se => {
        const _eanF = String(se.latest.ean || '').trim();
        const _fvF = fmtDateOnly(se.latest.fechaVenc);
        const _sucF = (se.latest.sucursal || '').toUpperCase().trim();
        devData
          .filter(d =>
            String(d.ean || '').trim() === _eanF &&
            fmtDateOnly(d.fechaVenc) === _fvF &&
            (d.sucursal || '').toUpperCase().trim() === _sucF
          )
          .sort((a, b) => {
            if (!a.fecha && !b.fecha) return 0;
            if (!a.fecha) return 1;
            if (!b.fecha) return -1;
            return b.fecha - a.fecha;
          })
          .forEach(d => {
            if (!seenDevIds.has(d.id)) {
              seenDevIds.add(d.id);
              allLinkedDevs.push(d);
            }
          });
      });
    });

    const allLinkedDevsHtml = allLinkedDevs.length > 0
      ? `<div style="display:flex;flex-wrap:wrap;gap:4px;align-items:center">
          ${allLinkedDevs.map(d => {
        const autoLabel = d.usuario === 'SISTEMA AUTO' ? ' ⚡' : '';
        const sucLabel = d.sucursal ? `<span style="font-size:8px;opacity:.7;margin-left:2px">[${esc(d.sucursal)}]</span>` : '';
        return `<span class="dev-chip" onclick="openDevModal(event,'${esc(d.id)}')" title="Acción ${esc(d.id)} · ${esc(d.sucursal || '')} · ${esc(d.motivo || '')}">${esc(d.id)}${autoLabel} ${sucLabel}↗</span>`;
      }).join('')}
         </div>`
      : `<span style="color:var(--text3);font-family:'IBM Plex Mono',monospace;font-size:9px">—</span>`;

    tableHtml += `
    <tr class="prod-row${isProdOpen ? ' open' : ''}" onclick="toggleProd(this,'${esc(prod.ean)}')">
      <td>
        <div class="dias-cell">
          <span class="dias-num ${prod.worstUrg}">${prod.worstDias !== null ? prod.worstDias : '—'}</span>
          <div class="dias-right">
            <span class="urg ${prod.worstUrg}">${URG_LABELS[prod.worstUrg]}</span>
            <div class="dias-bar"><div class="dias-bar-fill bar-colors-${prod.worstUrg}" style="width:${barPct}%"></div></div>
          </div>
        </div>
      </td>
      <td class="c-mono" style="font-size:11px;color:var(--text2)">${esc(prod.ean)}</td>
      <td class="c-main">
        <span class="expand-icon">▶</span>${esc(prod.desc)}
        ${prod.gramaje ? `<div style="font-size:9px;color:var(--text3);font-family:'IBM Plex Mono',monospace;margin-top:2px;letter-spacing:.04em">${esc(prod.gramaje)}</div>` : ''}
      </td>
      <td class="c-right" style="vertical-align:middle">${qtyRiskHtml(prodTotal, prod.worstDias)}</td>
      <td style="max-width:130px;overflow:hidden;text-overflow:ellipsis;color:var(--text2)" title="${esc(prod.prov)}">${esc(prod.prov || '—')}</td>
      <td><div class="suc-badges">${sucBadges}</div></td>
      <td>${vencMini}</td>
      <td style="vertical-align:middle;min-width:120px">${allLinkedDevsHtml}</td>
    </tr>`;

    prod.vencGrupos.forEach(g => {
      const isGrupoOpen = openGrupos.has(g.key);
      const sucFilter = document.getElementById('vf-suc').value;
      const sucVis = sucFilter ? g.sucursales.filter(se => se.suc === sucFilter) : g.sucursales;
      const gBarPct = g.dias === null ? 0
        : Math.max(0, Math.min(100, Math.round(g.dias / 90 * 100)));
      const gSucBadges = sucVis
        .map(se => `<span class="suc-b ${getSucClass(se.suc)}">${esc(se.suc || '—')}</span>`).join('');
      const grupoTotal = sucVis.reduce((s, se) =>
        s + (parseFloat(String(se.latest.cantidad || 0).replace(',', '.')) || 0), 0);

      tableHtml += `
      <tr class="venc-grupo-row${isProdOpen ? ' visible' : ''}${isGrupoOpen ? ' open' : ''}"
          data-prod-ean="${esc(prod.ean)}" data-grupo-key="${esc(g.key)}"
          onclick="toggleGrupo(this,'${esc(g.key)}')">
        <td>
          <div class="dias-cell">
            <span class="dias-num ${g.urg}" style="font-size:14px">${g.dias !== null ? g.dias : '—'}</span>
            <div class="dias-right">
              <span class="urg ${g.urg}" style="font-size:9px">${URG_LABELS[g.urg]}</span>
              <div class="dias-bar"><div class="dias-bar-fill bar-colors-${g.urg}" style="width:${gBarPct}%"></div></div>
            </div>
          </div>
        </td>
        <td colspan="2" style="color:var(--text2)">
          <span class="expand-icon">▶</span>
          <span style="font-family:'IBM Plex Mono',monospace;font-size:11px">Vence: <strong>${fmtDateOnly(g.fechaVenc)}</strong></span>
          <span style="font-size:10px;color:var(--text3);margin-left:10px">${g.sucursales.length} suc · ${grupoTotal} u.</span>
        </td>
        <td class="c-right c-mono">${grupoTotal}</td>
        <td></td>
        <td><div class="suc-badges">${gSucBadges}</div></td>
        <td></td>
        <td></td>
      </tr>`;

      sucVis.forEach(se => {
        const f = se.latest;
        const est = (f.estadoGest || 'ACTIVO').replace(/\s+/g, '-');
        const nControles = se.controles.length;
        const hasHistory = nControles > 1;
        const sucKey = g.key + '||' + se.suc;
        const isSucOpen = openSucs.has(sucKey);
        const showRow = isProdOpen && isGrupoOpen;
        const sucColor = getSucColorVar(se.suc);
        const cantActualSuc = parseFloat(String(f.cantidad || 0).replace(',', '.')) || 0;
        const ctrlBadge = `<span class="ctrl-badge ${hasHistory ? 'multi' : 'single'}">${nControles} control${nControles !== 1 ? 'es' : ''}</span>`;

        const _eanF = String(f.ean || '').trim();
        const _fvF = fmtDateOnly(f.fechaVenc);
        const _sucF = (f.sucursal || '').toUpperCase().trim();
        const linkedDevs = devData.filter(d =>
          String(d.ean || '').trim() === _eanF &&
          fmtDateOnly(d.fechaVenc) === _fvF &&
          (d.sucursal || '').toUpperCase().trim() === _sucF
        ).sort((a, b) => {
          if (!a.fecha && !b.fecha) return 0;
          if (!a.fecha) return 1;
          if (!b.fecha) return -1;
          return b.fecha - a.fecha;
        });

        const devChip = linkedDevs.length > 0
          ? `<div style="display:flex;flex-wrap:wrap;gap:4px;margin-top:2px">${linkedDevs.map(d => {
            const autoLabel = d.usuario === 'SISTEMA AUTO' ? ' ⚡' : '';
            return `<span class="dev-chip" onclick="openDevModal(event,'${esc(d.id)}')" title="${esc(d.id)}">${esc(d.id)}${autoLabel} ↗</span>`;
          }).join('')}</div>`
          : `<span style="color:var(--text3);font-family:'IBM Plex Mono',monospace;font-size:10px">Sin acción</span>`;

        const loteCheckbox = loteMode ? `
          <input type="checkbox" data-suc-key="${esc(sucKey)}"
            ${loteSeleccionados.has(sucKey) ? 'checked' : ''}
            onclick="event.stopPropagation();toggleLoteRow('${esc(sucKey)}','${esc(f.id)}','${esc(f.descripcion || '')}','${esc(f.ean || '')}','${esc(f.fechaVenc || '')}','${esc(f.sucursal || '')}',${cantActualSuc},this)"
            style="cursor:pointer;accent-color:#60a5fa;width:14px;height:14px;flex-shrink:0">
        ` : '';

        const btnRegistrarVencimiento = (!loteMode && g.urg === 'VENCIDO' && cantActualSuc > 0)
          ? `<button onclick="registrarVencimientoManual(event,'${esc(f.id)}','${esc(f.sucursal)}',${cantActualSuc},'${esc(f.descripcion || '')}','${esc(f.ean || '')}','${esc(f.fechaVenc || '')}')"
               class="btn-action" style="background:#3d0000;color:#ff4444;border:1px solid #9b1c1c;font-weight:800;padding:6px 12px;font-size:11px">
               🚨 Registrar Vencimiento
             </button>`
          : '';

        tableHtml += `
        <tr class="suc-row${showRow ? ' visible' : ''}${isSucOpen ? ' suc-open' : ''}"
            data-suc-key="${esc(sucKey)}" data-prod-ean="${esc(prod.ean)}" data-grupo-key="${esc(g.key)}">
          <td colspan="${NCOLS}" style="padding:0">
            <div class="suc-card${hasHistory ? ' has-history' : ''}${isSucOpen ? ' suc-open' : ''}" style="border-left-color:${sucColor}"
                 ${hasHistory ? `onclick="toggleSuc(event,'${esc(sucKey)}')"` : ''}>
              <div class="suc-card-hdr" style="border-left:3px solid ${sucColor}20">
                <div class="suc-card-hdr-l">
                  ${loteCheckbox}
                  ${hasHistory ? `<span class="suc-exp-icon">▶</span>` : `<span style="color:var(--text3);font-size:10px">↳</span>`}
                  <span class="suc-name" style="color:${sucColor}">${esc(f.sucursal || '—')}</span>
                  ${ctrlBadge}
                  ${staleBadgeHtml(se.diasDesde, true)}
                </div>
                <div class="suc-card-hdr-r">
                  <span class="eg ${est}">${est.replace(/-/g, ' ')}</span>
                </div>
              </div>
              <div class="suc-card-body">
                <div class="sf"><div class="sf-lbl">Controlado por</div><div class="sf-val${!f.usuario ? ' empty' : ''}">${f.usuario || '—'}</div></div>
                <div class="sf"><div class="sf-lbl">Fecha control</div><div class="sf-val mono">${fmtDate(f.fechaReg || '')}</div></div>
                <div class="sf"><div class="sf-lbl">Unidades</div><div class="sf-val mono" id="qty-${f.id}">${cantActualSuc}</div></div>
                <div class="sf"><div class="sf-lbl">Lote</div><div class="sf-val mono${!f.lote ? ' empty' : ''}">${f.lote || '—'}</div></div>
                ${!loteMode ? `
<div class="sf medium">
  <div class="sf-lbl">Ajuste / Movimientos</div>
  <div class="sf-val" style="display:flex;gap:5px;flex-wrap:wrap;margin-top:3px">
    <button onclick="abrirModalTransferencia(event,'${f.id}','${esc(f.sucursal)}',${cantActualSuc})" class="btn-action btn-transfer-sm">⇄ Transferir</button>
    <button onclick="ajustarStock(event,'${f.id}',1)"  class="btn-action btn-adj-pos">＋ Ajuste</button>
    <button onclick="ajustarStock(event,'${f.id}',-1)" class="btn-action btn-adj-neg">－ Ajuste</button>
    ${cantActualSuc > 0 && est !== 'ACCIONADO-TOTAL' ? `<button onclick="abrirModalAccionVen(event,'${f.id}','${esc(f.sucursal)}',${cantActualSuc},'${esc(f.descripcion || '')}','${esc(f.ean || '')}','${esc(f.fechaVenc || '')}')" class="btn-action" style="background:rgba(204, 116, 255, 0.46);color:white;border:1px solid rgb(110, 0, 173)">📦 Acción</button>` : ''}
    ${btnRegistrarVencimiento}
  </div>
</div>` : `<div class="sf"><div class="sf-lbl">Modo Lote</div><div class="sf-val" style="color:var(--text3);font-family:'IBM Plex Mono',monospace;font-size:10px;margin-top:3px">☑ Seleccionado</div></div>`}
                <div class="sf wide"><div class="sf-lbl">Acción${linkedDevs.length > 1 ? 'es' : ''} vinculada${linkedDevs.length > 1 ? 's' : ''} (${linkedDevs.length || 'sin acción'})</div><div class="sf-val" style="margin-top:2px">${devChip}</div></div>
                ${f.aclaracion ? `<div class="sf wide"><div class="sf-lbl">Aclaración</div><div class="sf-val">${esc(f.aclaracion)}</div></div>` : ''}
              </div>
            </div>
          </td>
        </tr>`;

        se.controles.forEach((ctrl, idx) => {
          const isLatest = idx === 0;
          const ctrlEst = (ctrl.estadoGest || 'ACTIVO').replace(/\s+/g, '-');
          const ctrlAcc = ctrl.idAccion || '';
          const showCtrl = showRow && isSucOpen;
          let deltaHtml = '';
          if (idx < se.controles.length - 1) {
            const older = se.controles[idx + 1];
            if (ctrl.fechaReg && older.fechaReg) {
              const diff = Math.round((ctrl.fechaReg - older.fechaReg) / 86400000);
              const cls = diff > 0 ? 'pos' : diff < 0 ? 'neg' : 'neu';
              deltaHtml = `<span class="delta ${cls}">${diff > 0 ? '+' : ''}${diff}d vs anterior</span>`;
            }
          }
          const ctrlChip = ctrlAcc
            ? `<span class="dev-chip" onclick="openDevModal(event,'${esc(ctrlAcc)}')">${esc(ctrlAcc)} ↗</span>`
            : '';
          tableHtml += `
          <tr class="ctrl-row${showCtrl ? ' visible' : ''}"
              data-suc-key="${esc(sucKey)}" data-prod-ean="${esc(prod.ean)}" data-grupo-key="${esc(g.key)}">
            <td colspan="${NCOLS}" style="padding:0">
              <div class="ctrl-card${isLatest ? ' is-latest' : ''}">
                <div class="ctrl-num-wrap">
                  <div style="display:flex;flex-direction:column;gap:3px">
                    <span style="font-family:'IBM Plex Mono',monospace;font-size:8px;letter-spacing:1px;text-transform:uppercase;color:var(--text3)">N° Ctrl</span>
                    <div style="display:flex;align-items:center;gap:5px">
                      <span class="ctrl-n${isLatest ? ' latest' : ''}">C${nControles - idx}</span>
                      ${isLatest ? `<span class="ult-badge">✓ ÚLTIMO</span>` : ''}
                    </div>
                  </div>
                </div>
                <div class="ctrl-fields">
                  <div class="ctrl-field"><div class="cf-lbl">Fecha</div><div class="cf-val mono${isLatest ? ' latest' : ''}">${fmtDate(ctrl.fechaReg || '')}</div></div>
                  <div class="ctrl-field"><div class="cf-lbl">Usuario</div><div class="cf-val${isLatest ? ' latest' : ''}">${esc(ctrl.usuario || '—')}</div></div>
                  <div class="ctrl-field"><div class="cf-lbl">Unidades</div><div class="cf-val${isLatest ? ' latest' : ''}">${ctrl.cantidad ? ctrl.cantidad + ' u' : '—'}</div></div>
                  <div class="ctrl-field"><div class="cf-lbl">Lote</div><div class="cf-val mono${isLatest ? ' latest' : ''}">${esc(ctrl.lote || '—')}</div></div>
                  ${isLatest
              ? `<div class="ctrl-field"><div class="cf-lbl">Días desde</div><div style="margin-top:2px">${staleBadgeHtml(se.diasDesde, true)}</div></div>`
              : (deltaHtml ? `<div class="ctrl-field"><div class="cf-lbl">Intervalo</div><div style="margin-top:2px">${deltaHtml}</div></div>` : '')
            }
                  <div class="ctrl-field"><div class="cf-lbl">Estado</div><div class="cf-val${isLatest ? ' latest' : ''}"><span class="eg ${ctrlEst}" style="font-size:8px;padding:2px 5px">${ctrlEst.replace(/-/g, ' ')}</span></div></div>
                  ${ctrlAcc ? `<div class="ctrl-field"><div class="cf-lbl">Acción</div><div style="margin-top:2px">${ctrlChip}</div></div>` : ''}
                </div>
              </div>
            </td>
          </tr>`;
        });
      });
    });
  });

  tableHtml += '</tbody></table></div></div>'; // close .venc-desktop

  // ─── MOBILE CARDS ────────────────────────────────────
  let mobileHtml = '<div class="venc-mobile-cards">';

  filteredVenProd.forEach(prod => {
    const prodId = 'vmp-' + prod.ean.replace(/\D/g, '');
    const isOpen = openProds.has(prod.ean);
    const prodTotal = prod.vencGrupos.reduce((sum, g) =>
      sum + g.sucursales.reduce((s2, se) =>
        s2 + (parseFloat(String(se.latest.cantidad || 0).replace(',', '.')) || 0), 0), 0);

    // Acciones vinculadas al producto (mobile)
    const mSeenIds = new Set();
    const mAllLinkedDevs = [];
    prod.vencGrupos.forEach(g => {
      g.sucursales.forEach(se => {
        const _eanF = String(se.latest.ean || '').trim();
        const _fvF = fmtDateOnly(se.latest.fechaVenc);
        const _sucF = (se.latest.sucursal || '').toUpperCase().trim();
        devData
          .filter(d =>
            String(d.ean || '').trim() === _eanF &&
            fmtDateOnly(d.fechaVenc) === _fvF &&
            (d.sucursal || '').toUpperCase().trim() === _sucF
          )
          .forEach(d => {
            if (!mSeenIds.has(d.id)) { mSeenIds.add(d.id); mAllLinkedDevs.push(d); }
          });
      });
    });

    const mLinkedHtml = mAllLinkedDevs.length > 0
      ? `<div style="display:flex;flex-wrap:wrap;gap:4px;margin-top:6px;padding-top:6px;border-top:1px solid var(--border)">
          <span style="font-family:'IBM Plex Mono',monospace;font-size:8px;color:var(--text3);letter-spacing:.08em;text-transform:uppercase;width:100%;margin-bottom:2px">🔗 Acciones vinculadas</span>
          ${mAllLinkedDevs.map(d => {
        const autoLabel = d.usuario === 'SISTEMA AUTO' ? ' ⚡' : '';
        return `<span class="dev-chip" onclick="openDevModal(event,'${esc(d.id)}')">${esc(d.id)}${autoLabel} <span style="font-size:8px;opacity:.7">[${esc(d.sucursal || '')}]</span> ↗</span>`;
      }).join('')}
         </div>`
      : '';

    mobileHtml += `
    <div class="vmc-prod vmc-urg-${prod.worstUrg}" id="${prodId}">

      <div class="vmc-prod-hdr" onclick="toggleVMProd('${esc(prod.ean)}','${prodId}')">
        <div class="vmc-dias-block">
          <span class="vmc-dias-num ${prod.worstUrg}">${prod.worstDias !== null ? prod.worstDias : '?'}</span>
          <span class="vmc-dias-label">días</span>
        </div>
        <div class="vmc-prod-info">
          <div class="vmc-prod-desc">${esc(prod.desc)}</div>
          ${prod.gramaje ? `<div style="font-size:10px;color:var(--text3);font-family:'IBM Plex Mono',monospace;margin-top:1px">${esc(prod.gramaje)}</div>` : ''}
          <div class="vmc-prod-meta">
            <span class="urg ${prod.worstUrg}" style="font-size:9px">${URG_LABELS[prod.worstUrg]}</span>
            ${prod.prov ? `<span class="vmc-prov">${esc(prod.prov)}</span>` : ''}
          </div>
          <div class="vmc-prod-sucs">
            ${prod.allSucs.map(s => `<span class="suc-b ${getSucClass(s)}">${esc(s)}</span>`).join('')}
            <span class="vmc-stock-chip">${prodTotal} u. total</span>
          </div>
          <div class="vmc-fechas-mini">
            ${prod.vencGrupos.map(g => `
              <span class="vmc-fecha-chip ${g.urg}">
                ${fmtDateOnly(g.fechaVenc)}
                <span class="vmc-fecha-dias">${g.dias !== null ? (g.dias <= 0 ? 'VENC' : g.dias + 'd') : '?'}</span>
              </span>`).join('')}
          </div>
          ${mLinkedHtml}
        </div>
        <div class="vmc-expand-btn${isOpen ? ' open' : ''}">
          <span class="vmc-expand-icon">›</span>
        </div>
      </div>

      <div class="vmc-grupos${isOpen ? ' open' : ''}">
        ${prod.vencGrupos.map((g, gi) => {
      const grupoId = prodId + '-g' + gi;
      const isGrupoOpen = openGrupos.has(g.key);
      const sucFilter = document.getElementById('vf-suc') ? document.getElementById('vf-suc').value : '';
      const sucVis = sucFilter ? g.sucursales.filter(se => se.suc === sucFilter) : g.sucursales;
      const grupoTotal = sucVis.reduce((s, se) => s + (parseFloat(String(se.latest.cantidad || 0).replace(',', '.')) || 0), 0);

      return `
          <div class="vmc-grupo" id="${grupoId}">
            <div class="vmc-grupo-hdr" onclick="toggleVMGrupo('${esc(g.key)}','${grupoId}')">
              <div class="vmc-grupo-urg-bar vmc-urg-bar-${g.urg}"></div>
              <div class="vmc-grupo-info">
                <div class="vmc-grupo-fecha">
                  <span style="font-size:9px;color:var(--text3);font-family:'IBM Plex Mono',monospace;letter-spacing:.04em">VENCE</span>
                  <span style="font-family:'IBM Plex Mono',monospace;font-size:14px;font-weight:700;color:var(--text)">${fmtDateOnly(g.fechaVenc)}</span>
                </div>
                <div class="vmc-grupo-badges">
                  <span class="urg ${g.urg}" style="font-size:9px">${g.dias !== null ? (g.dias <= 0 ? 'VENCIDO' : g.dias + ' días') : '?'} · ${URG_LABELS[g.urg]}</span>
                  <span class="vmc-stock-chip">${grupoTotal} u.</span>
                  <span style="font-family:'IBM Plex Mono',monospace;font-size:9px;color:var(--text3)">${sucVis.length} suc.</span>
                  ${g.hasAccion ? `<span class="linked-indicator">🔗 acción</span>` : ''}
                </div>
                <div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:4px">
                  ${sucVis.map(se => `<span class="suc-b ${getSucClass(se.suc)}">${esc(se.suc)}</span>`).join('')}
                </div>
              </div>
              <span class="vmc-group-arrow${isGrupoOpen ? ' open' : ''}">›</span>
            </div>

            <div class="vmc-sucs${isGrupoOpen ? ' open' : ''}">
              ${sucVis.map((se, si) => {
        const f = se.latest;
        const est = (f.estadoGest || 'ACTIVO').replace(/\s+/g, '-');
        const cantActualSuc = parseFloat(String(f.cantidad || 0).replace(',', '.')) || 0;
        const sucKey = g.key + '||' + se.suc;
        const sucColor = getSucColorVar(se.suc);
        const nControles = se.controles.length;
        const hasHistory = nControles > 1;
        const isSucOpen = openSucs.has(sucKey);
        const sucId = grupoId + '-s' + si;

        const _eanF = String(f.ean || '').trim();
        const _fvF = fmtDateOnly(f.fechaVenc);
        const _sucF = (f.sucursal || '').toUpperCase().trim();
        const linkedDevs = devData.filter(d =>
          String(d.ean || '').trim() === _eanF &&
          fmtDateOnly(d.fechaVenc) === _fvF &&
          (d.sucursal || '').toUpperCase().trim() === _sucF
        );

        const btnRegistrar = (!loteMode && g.urg === 'VENCIDO' && cantActualSuc > 0)
          ? `<button onclick="registrarVencimientoManual(event,'${esc(f.id)}','${esc(f.sucursal)}',${cantActualSuc},'${esc(f.descripcion || '')}','${esc(f.ean || '')}','${esc(f.fechaVenc || '')}')"
                       class="vmc-btn vmc-btn-venc">🚨 Registrar Vencido</button>` : '';

        const loteCheck = loteMode ? `
                  <input type="checkbox" ${loteSeleccionados.has(sucKey) ? 'checked' : ''}
                    onclick="event.stopPropagation();toggleLoteRow('${esc(sucKey)}','${esc(f.id)}','${esc(f.descripcion || '')}','${esc(f.ean || '')}','${esc(f.fechaVenc || '')}','${esc(f.sucursal || '')}',${cantActualSuc},this)"
                    style="width:18px;height:18px;cursor:pointer;accent-color:#60a5fa;flex-shrink:0">` : '';

        return `
                <div class="vmc-suc-card" id="${sucId}" style="border-left-color:${sucColor}">
                  <div class="vmc-suc-hdr" onclick="toggleVMSuc('${esc(sucKey)}','${sucId}',event)">
                    ${loteCheck}
                    <span class="vmc-suc-name" style="color:${sucColor}">${esc(f.sucursal || '—')}</span>
                    <span class="vmc-suc-qty-badge ${est === 'ACTIVO' ? 'qty-active' : 'qty-done'}">${cantActualSuc} u.</span>
                    <span class="eg ${est}" style="font-size:9px;margin-left:auto">${est.replace(/-/g, ' ')}</span>
                    ${staleBadgeHtml(se.diasDesde, true)}
                    ${hasHistory ? `<span class="vmc-hist-btn${isSucOpen ? ' open' : ''}">📋 ${nControles}</span>` : ''}
                  </div>
                  <div class="vmc-suc-detail">
                    <div class="vmc-detail-grid">
                      <div class="vmc-dfield"><div class="vmc-dlbl">Usuario</div><div class="vmc-dval">${f.usuario || '—'}</div></div>
                      <div class="vmc-dfield"><div class="vmc-dlbl">Fecha control</div><div class="vmc-dval mono">${fmtDate(f.fechaReg || '')}</div></div>
                      <div class="vmc-dfield"><div class="vmc-dlbl">EAN</div><div class="vmc-dval mono">${esc(f.ean || '—')}</div></div>
                      <div class="vmc-dfield"><div class="vmc-dlbl">Lote</div><div class="vmc-dval mono">${f.lote || '—'}</div></div>
                      ${f.aclaracion ? `<div class="vmc-dfield vmc-dfield-wide"><div class="vmc-dlbl">Aclaración</div><div class="vmc-dval">${esc(f.aclaracion)}</div></div>` : ''}
                    </div>
                    ${!loteMode ? `
                    <div class="vmc-action-btns">
                      <button onclick="abrirModalTransferencia(event,'${f.id}','${esc(f.sucursal)}',${cantActualSuc})" class="vmc-btn vmc-btn-transfer">⇄ Transferir</button>
                      <button onclick="ajustarStock(event,'${f.id}',1)"  class="vmc-btn vmc-btn-pos">＋ Ajuste</button>
                      <button onclick="ajustarStock(event,'${f.id}',-1)" class="vmc-btn vmc-btn-neg">－ Ajuste</button>
                      ${cantActualSuc > 0 && est !== 'ACCIONADO-TOTAL' ? `<button onclick="abrirModalAccionVen(event,'${f.id}','${esc(f.sucursal)}',${cantActualSuc},'${esc(f.descripcion || '')}','${esc(f.ean || '')}','${esc(f.fechaVenc || '')}')" class="vmc-btn" style="background:rgba(204, 116, 255, 0.46);color:white;border:1px solid rgb(110, 0, 173)">📦 Acción</button>` : ''}
                      ${btnRegistrar}
                    </div>` : ''}

                    ${linkedDevs.length > 0 ? `
                    <div class="vmc-linked-devs">
                      <div class="vmc-dlbl" style="margin-bottom:4px">Acciones vinculadas</div>
                      <div style="display:flex;flex-wrap:wrap;gap:5px">
                        ${linkedDevs.map(d => {
          const autoLabel = d.usuario === 'SISTEMA AUTO' ? ' ⚡' : '';
          return `<span class="dev-chip" onclick="openDevModal(event,'${esc(d.id)}')">${esc(d.id)}${autoLabel} ↗</span>`;
        }).join('')}
                      </div>
                    </div>` : ''}
                    ${hasHistory ? `
                    <div class="vmc-history${isSucOpen ? ' open' : ''}">
                      <div class="vmc-history-title">📋 Historial de controles</div>
                      ${se.controles.map((ctrl, idx) => {
          const isLatest = idx === 0;
          return `
                        <div class="vmc-ctrl-row${isLatest ? ' latest' : ''}">
                          <div class="vmc-ctrl-num">C${nControles - idx}${isLatest ? ' ✓' : ''}</div>
                          <div class="vmc-ctrl-fields">
                            <span class="vmc-ctrl-val mono">${fmtDate(ctrl.fechaReg || '')}</span>
                            <span class="vmc-ctrl-val">${ctrl.usuario || '—'}</span>
                            <span class="vmc-ctrl-val mono">${ctrl.cantidad ? ctrl.cantidad + ' u.' : '—'}</span>
                            ${ctrl.lote ? `<span class="vmc-ctrl-val mono">${esc(ctrl.lote)}</span>` : ''}
                          </div>
                        </div>`;
        }).join('')}
                    </div>` : ''}
                  </div>
                </div>`;
      }).join('')}
            </div>
          </div>`;
    }).join('')}
      </div>

    </div>`; // end vmc-prod
  });

  mobileHtml += '</div>'; // end .venc-mobile-cards

  wrap.innerHTML = headBar + tableHtml + mobileHtml;
}


// ══════════════════════════════════════════════════════
//  TOGGLE HELPERS — MOBILE VENCIMIENTOS
//  Agregar al final de dashboard.js
// ══════════════════════════════════════════════════════

function toggleVMProd(ean, prodId) {
  const opening = !openProds.has(ean);
  if (opening) openProds.add(ean); else openProds.delete(ean);

  const card = document.getElementById(prodId);
  if (!card) return;
  const grupos = card.querySelector('.vmc-grupos');
  const btn = card.querySelector('.vmc-expand-btn');
  if (grupos) grupos.classList.toggle('open', opening);
  if (btn) btn.classList.toggle('open', opening);

  // Colapsar sub-niveles al cerrar
  if (!opening) {
    card.querySelectorAll('.vmc-grupo').forEach(g => {
      g.querySelector('.vmc-sucs')?.classList.remove('open');
      g.querySelector('.vmc-group-arrow')?.classList.remove('open');
    });
    card.querySelectorAll('.vmc-suc-card').forEach(s => {
      s.querySelector('.vmc-suc-detail')?.classList.remove('open');
      s.querySelector('.vmc-history')?.classList.remove('open');
      s.querySelector('.vmc-hist-btn')?.classList.remove('open');
    });
    // Limpiar Sets de estado
    openGrupos.forEach(k => { if (filteredVenProd.find(p => p.ean === ean)?.vencGrupos.some(g => g.key === k)) openGrupos.delete(k); });
  }
}

function toggleVMGrupo(key, grupoId) {
  const opening = !openGrupos.has(key);
  if (opening) openGrupos.add(key); else openGrupos.delete(key);

  const g = document.getElementById(grupoId);
  if (!g) return;
  g.querySelector('.vmc-sucs')?.classList.toggle('open', opening);
  g.querySelector('.vmc-group-arrow')?.classList.toggle('open', opening);
}

function toggleVMSuc(sucKey, sucId, evt) {
  // Evitar que chips/botones propagen
  if (evt.target.classList.contains('dev-chip') ||
    evt.target.classList.contains('vmc-btn') ||
    evt.target.tagName === 'BUTTON' ||
    evt.target.tagName === 'INPUT') return;

  const opening = !openSucs.has(sucKey);
  if (opening) openSucs.add(sucKey); else openSucs.delete(sucKey);

  const card = document.getElementById(sucId);
  if (!card) return;
  card.querySelector('.vmc-suc-detail')?.classList.toggle('open', opening);
  card.querySelector('.vmc-hist-btn')?.classList.toggle('open', opening);

  // El historial se muestra dentro del detalle cuando está abierto
  if (!opening) {
    card.querySelector('.vmc-history')?.classList.remove('open');
  } else {
    // Abrir historial automáticamente si hay historial
    card.querySelector('.vmc-history')?.classList.add('open');
  }
}

function toggleGrupo(row, key) {
  const opening = !openGrupos.has(key);
  if (opening) openGrupos.add(key); else openGrupos.delete(key);
  row.classList.toggle('open', opening);
  document.querySelectorAll('tr.suc-row, tr.ctrl-row').forEach(r => {
    if (r.dataset.grupoKey !== key) return;
    if (!opening) { r.classList.remove('visible', 'suc-open'); if (r.dataset.sucKey) openSucs.delete(r.dataset.sucKey); }
    else { if (r.classList.contains('suc-row')) r.classList.add('visible'); }
  });
}

function toggleSuc(evt, sucKey) {
  if (evt.target.classList.contains('dev-chip')) return;
  const opening = !openSucs.has(sucKey);
  if (opening) openSucs.add(sucKey); else openSucs.delete(sucKey);
  document.querySelectorAll(`tr.suc-row[data-suc-key="${sucKey}"]`).forEach(r => r.classList.toggle('suc-open', opening));
  document.querySelectorAll(`tr.ctrl-row[data-suc-key="${sucKey}"]`).forEach(r => r.classList.toggle('visible', opening));
}

// ══════════════════════════════════════════════════════
//  EXPORT VENCIMIENTOS
// ══════════════════════════════════════════════════════
async function exportVencXlsx() {
  if (!filteredVenProd.length) { alert('No hay registros para exportar con los filtros actuales.'); return; }
  const rows = [];
  filteredVenProd.forEach(prod => {
    prod.vencGrupos.forEach(g => {
      g.sucursales.forEach(se => {
        const f = se.latest;
        rows.push({
          ean: prod.ean, descripcion: prod.desc, proveedor: prod.prov || '—',
          urgencia: g.urg, dias: g.dias !== null ? g.dias : '—',
          fechaVenc: fmtDateOnly(g.fechaVenc), sucursal: se.suc,
          cantidad: parseFloat(String(f.cantidad || 0).replace(',', '.')) || 0,
          estadoGest: f.estadoGest || 'ACTIVO', lote: f.lote || '—',
          usuario: f.usuario || '—', fechaReg: fmtDate(f.fechaReg || ''),
          aclaracion: f.aclaracion || '—', nControles: se.controles.length,
        });
      });
    });
  });
  if (!rows.length) { alert('Sin datos para exportar.'); return; }
  const wb = new ExcelJS.Workbook();
  wb.creator = 'NEXUS v3.0'; wb.created = new Date();
  const ws = wb.addWorksheet('Vencimientos');
  ws.pageSetup = {
    paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    margins: { left: .3, right: .3, top: .4, bottom: .4, header: .2, footer: .2 }
  };
  ws.columns = [
    { key: 'ean', width: 16 }, { key: 'descripcion', width: 36 }, { key: 'proveedor', width: 22 },
    { key: 'urgencia', width: 12 }, { key: 'dias', width: 8 }, { key: 'fechaVenc', width: 14 },
    { key: 'sucursal', width: 14 }, { key: 'cantidad', width: 10 }, { key: 'estadoGest', width: 16 },
    { key: 'lote', width: 14 }, { key: 'usuario', width: 18 }, { key: 'fechaReg', width: 18 },
    { key: 'nControles', width: 10 }, { key: 'aclaracion', width: 26 },
  ];
  const hRow = ws.addRow(['EAN', 'Descripción', 'Proveedor', 'Urgencia', 'Días', 'F. Vencimiento', 'Sucursal', 'Unidades', 'Estado', 'Lote', 'Usuario', 'F. Registro', 'N° Ctrl', 'Aclaración']);
  hRow.height = 22;
  hRow.eachCell(c => {
    c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D111F' } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = { bottom: { style: 'medium', color: { argb: 'FF4f8eff' } } };
  });
  const urgColors = {
    VENCIDO: { bg: 'FFFF4444' }, CRITICO: { bg: 'FFFF7C2A' }, URGENTE: { bg: 'FFFBBF24' },
    PROXIMO: { bg: 'FFA3E635' }, ATENCION: { bg: 'FF38BDF8' }, NORMAL: { bg: 'FF4ADE80' },
  };
  const estadoColors = {
    'ACTIVO': { bg: 'FFD1ECF1', fg: 'FF0C5460' },
    'ACCION_TOMADA': { bg: 'FFD4EDDA', fg: 'FF155724' },
    'RETIRADO': { bg: 'FFE2E3E5', fg: 'FF383D41' },
  };
  rows.forEach((r, i) => {
    const isEven = i % 2 === 1;
    const dRow = ws.addRow([r.ean, r.descripcion, r.proveedor, r.urgencia, r.dias, r.fechaVenc,
    r.sucursal, r.cantidad, r.estadoGest, r.lote, r.usuario, r.fechaReg, r.nControles, r.aclaracion]);
    dRow.height = 18;
    dRow.eachCell({ includeEmpty: true }, (c, col) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEven ? 'FFF5F7FF' : 'FFFFFFFF' } };
      c.alignment = { vertical: 'middle', horizontal: col === 8 || col === 5 || col === 13 ? 'center' : 'left', wrapText: col === 2 };
      c.font = { size: 10 };
      c.border = { bottom: { style: 'thin', color: { argb: 'FFDDDDDD' } }, right: { style: 'thin', color: { argb: 'FFDDDDDD' } } };
    });
    const uc = urgColors[r.urgencia];
    if (uc) { const cell = dRow.getCell(4); cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + uc.bg.slice(2) } }; cell.font = { size: 10, bold: true, color: { argb: 'FFFFFFFF' } }; }
    const normEst = (r.estadoGest || '').replace(/\s+/g, '_');
    const ec = estadoColors[normEst];
    if (ec) { const cell = dRow.getCell(9); cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ec.bg } }; cell.font = { size: 10, bold: true, color: { argb: ec.fg } }; }
  });
  const totRow = ws.addRow(['', `TOTAL: ${rows.length} registros`, '', '', '',
    '', '', rows.reduce((s, r) => s + (r.cantidad || 0), 0), '', '', '', '', '', '']);
  totRow.height = 20;
  totRow.eachCell({ includeEmpty: true }, c => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D111F' } };
    c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
    c.alignment = { vertical: 'middle' };
  });
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];
  const wsMeta = wb.addWorksheet('Info exportación');
  wsMeta.addRow(['Exportado desde', 'NEXUS v3.0 — Control de Vencimientos']);
  wsMeta.addRow(['Fecha exportación', new Date().toLocaleString('es-AR')]);
  wsMeta.addRow(['Filtro aplicado — fecha por', vencDateMode === 'registro' ? 'Fecha de Registro' : 'Fecha de Vencimiento']);
  wsMeta.addRow(['Preset período', vencPreset]);
  wsMeta.addRow(['Modo fecha desde', vencDateFrom ? vencDateFrom.toLocaleDateString('es-AR') : 'Sin límite']);
  wsMeta.addRow(['Modo fecha hasta', vencDateTo ? vencDateTo.toLocaleDateString('es-AR') : 'Sin límite']);
  wsMeta.addRow(['Filtro urgencia', document.getElementById('vf-urg').value || 'Todas']);
  wsMeta.addRow(['Filtro proveedor', document.getElementById('vf-prov').value || 'Todos']);
  wsMeta.addRow(['Filtro sucursal', document.getElementById('vf-suc').value || 'Todas']);
  wsMeta.addRow(['Filtro estado', document.getElementById('vf-estado').value || 'Todos']);
  wsMeta.addRow(['Total productos', filteredVenProd.length]);
  wsMeta.addRow(['Total grupos/fechas', filteredVen.length]);
  wsMeta.addRow(['Total registros exportados', rows.length]);
  wsMeta.columns = [{ width: 34 }, { width: 36 }];
  try {
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'vencimientos_' + fmtIso(new Date()) + '.xlsx';
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  } catch (err) { alert('Error al generar Excel: ' + err.message); console.error(err); }
}

// ══════════════════════════════════════════════════════
//  REGISTRAR VENCIMIENTO MANUAL
// ══════════════════════════════════════════════════════
async function registrarVencimientoManual(evt, venId, suc, cant, desc, ean, fechaVenc) {
  evt.stopPropagation();
  document.getElementById('cm-title').textContent = '🚨 Registrar producto como Vencido';
  document.getElementById('cm-text').textContent = `¿Registrar ${cant} u. de "${desc}" en ${suc} como VENCIDO?`;
  document.getElementById('cm-details').innerHTML = `
    <div style="background:#3d0000;border:1px solid #9b1c1c;padding:10px 14px;border-radius:7px;font-size:12px;color:#ff8080;margin-bottom:8px">
      <strong>Esta acción va a:</strong>
      <ul style="margin-top:6px;padding-left:16px;line-height:1.8">
        <li>Crear un registro de devolución por <strong>VENCIDO</strong> con ${cant} u.</li>
        <li>Poner el stock de esta sucursal en <strong>0</strong></li>
        <li>El producto desaparecerá del dashboard (stock = 0)</li>
      </ul>
    </div>
    <div style="font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--text3)">
      EAN: ${esc(ean)} · SUC: ${esc(suc)} · Vence: ${fmtDateOnly(fechaVenc)}
    </div>`;
  modalCb = async () => {
    document.getElementById('confirmModal').classList.remove('open');
    showSpinner('Registrando vencimiento…');
    try {
      const res = await fetch(SCRIPT_URL, {
        method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'registrarVencimientoManual', idVen: venId })
      });
      const json = await res.json();
      if (json.success) { showToast(true, `Vencimiento registrado: ${cant} u. → ${json.devId}`); await loadData(); }
      else showToast(false, json.message || 'Error al registrar vencimiento');
    } catch (e) { showToast(false, 'Error de red al registrar vencimiento'); console.error(e); }
    hideSpinner();
  };
  document.getElementById('cm-confirm').onclick = () => { if (modalCb) { const cb = modalCb; modalCb = null; cb(); } };
  document.getElementById('confirmModal').classList.add('open');
}

// ══════════════════════════════════════════════════════
//  ACCIONES PANEL
// ══════════════════════════════════════════════════════
function setAccPreset(preset, btn) {
  accPreset = preset;
  document.querySelectorAll('#acc-presets .dp').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');

  const fromWrap = document.getElementById('acc-date-from');
  const toWrap = document.getElementById('acc-date-to');
  fromWrap.style.display = preset === 'custom' ? '' : 'none';
  toWrap.style.display = preset === 'custom' ? '' : 'none';

  const now = new Date(); now.setHours(0, 0, 0, 0);

  if (preset === 'custom' || preset === 'all') {
    accDateFrom = null; accDateTo = null;
  } else if (preset === 'yesterday') {
    const y = new Date(now); y.setDate(y.getDate() - 1);
    accDateFrom = y;
    accDateTo = new Date(y); accDateTo.setHours(23, 59, 59, 999);
  } else if (preset === 'today') {
    accDateFrom = new Date(now);
    accDateTo = new Date(); accDateTo.setHours(23, 59, 59, 999);
  } else {
    const days = { '3d': 3, '7d': 7, '14d': 14, '30d': 30 }[preset] ?? 0;
    accDateFrom = new Date(now); accDateFrom.setDate(accDateFrom.getDate() - days);
    accDateTo = new Date(); accDateTo.setHours(23, 59, 59, 999);
  }
  applyAccFilters();
}

function applyAccFilters() {
  const suc = document.getElementById('af-suc')?.value || '';
  const prov = document.getElementById('af-prov')?.value || '';
  const mot = document.getElementById('af-motivo')?.value || '';
  const est = document.getElementById('af-estado')?.value || '';
  const urg = document.getElementById('af-urg')?.value || '';

  // Búsqueda unificada: input superior (nuevo) + input del head-bar (existente)
  const searchTop = (document.getElementById('acc-search-top')?.value || '').toLowerCase().trim();
  const searchOld = (document.getElementById('acc-search')?.value || '').toLowerCase().trim();
  const search = searchTop || searchOld;

  // Fechas
  let from = accDateFrom, to = accDateTo;
  if (accPreset === 'custom') {
    const f = document.getElementById('af-date-from')?.value;
    const t = document.getElementById('af-date-to')?.value;
    from = f ? new Date(f) : null;
    to = t ? new Date(t + 'T23:59:59') : null;
  }

  filteredAcc = devData.filter(r => {

    // Ocultar registros transferidos por completo (cantDisponible=0),
    // EXCEPTO los VENDIDO que siempre deben aparecer (en verde)
    if (r.cantDisponible <= 0 && r.estado !== 'VENDIDO') return false;

    // Filtros de select
    if (suc && r.sucursal !== suc) return false;
    if (prov && r.proveedor !== prov) return false;
    if (mot && r.motivo !== mot) return false;
    if (est && r.estado !== est) return false;

    // Urgencia: calculada desde fechaVenc del registro de acción
    if (urg) {
      const dias = calcDias(r.fechaVenc);
      const urgRec = getUrg(dias);
      if (urgRec !== urg) return false;
    }

    // Filtro de fecha con modo (registro vs vencimiento)
    if (from || to) {
      if (accDateMode === 'registro') {
        if (from && r.fecha && r.fecha < from) return false;
        if (to && r.fecha && r.fecha > to) return false;
      } else {
        // modo vencimiento: parsear fechaVenc del registro
        const fv = parseFechaVenc(r.fechaVenc);
        if (!fv) return false;
        if (from && fv < from) return false;
        if (to && fv > to) return false;
      }
    }

    // Búsqueda de texto
    if (search) {
      const hay = [r.descripcion, r.ean, r.lote, r.proveedor, r.sucursal, r.motivo, r.id]
        .join(' ').toLowerCase();
      if (!hay.includes(search)) return false;
    }

    return true;
  });

  accPage = 1;
  renderAccTable();
}

function clearAccFilters() {
  ['af-suc', 'af-prov', 'af-motivo', 'af-estado', 'af-urg'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });

  // Limpiar ambos inputs de búsqueda
  const st = document.getElementById('acc-search-top');
  const so = document.getElementById('acc-search');
  if (st) st.value = '';
  if (so) so.value = '';

  // Resetear modo fecha a "registro"
  setAccDateMode('registro');

  // Resetear preset a "Todos"
  const allBtn = document.querySelector('#acc-presets .dp:nth-child(7)'); // botón "Todos"
  setAccPreset('all', allBtn);
}

function clearNcFilters() {
  ['ncf-prov', 'ncf-estado', 'ncf-suc', 'ncf-motivo'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  const searchEl = document.getElementById('ncf-search');
  if (searchEl) searchEl.value = '';

  applyNcFilters();
}

function setAccDateMode(mode) {
  accDateMode = mode;

  const btnReg = document.getElementById('af-mode-reg');
  const btnVenc = document.getElementById('af-mode-venc');
  const label = document.getElementById('af-period-label');

  if (mode === 'registro') {
    if (btnReg) { btnReg.style.background = 'var(--blue)'; btnReg.style.color = '#fff'; }
    if (btnVenc) { btnVenc.style.background = 'var(--surface)'; btnVenc.style.color = 'var(--text3)'; }
    if (label) label.textContent = 'Período — fecha de registro';
  } else {
    if (btnVenc) { btnVenc.style.background = 'var(--blue)'; btnVenc.style.color = '#fff'; }
    if (btnReg) { btnReg.style.background = 'var(--surface)'; btnReg.style.color = 'var(--text3)'; }
    if (label) label.textContent = 'Período — fecha de vencimiento';
  }

  applyAccFilters();
}


function renderAccTable() {
  const start = (accPage - 1) * PAGE_SIZE;
  const page = filteredAcc.slice(start, start + PAGE_SIZE);
  const tbody = document.getElementById('accTableBody');

  if (!filteredAcc.length) {
    tbody.innerHTML = '<div class="empty"><div class="ei">🔍</div><p>No hay registros con los filtros aplicados</p></div>';
    document.getElementById('accPagination').style.display = 'none';
    return;
  }

  const venByDevId = new Map();
  venData.forEach(v => { if (v.idAccion) venByDevId.set(v.idAccion, v); });

  // ── Desktop ──────────────────────────────────────────────────
  const tableHtml = `
    <div class="table-scroll acc-desktop">
      <table>
        <thead><tr>
          <th>Fecha</th>
          <th>Sucursal</th>
          <th>Descripción</th>
          <th>EAN</th>
          <th class="c-right">Disp.</th>
          <th>Vencimiento</th>
          <th>Proveedor</th>
          <th>Motivo</th>
          <th>Lote</th>
          <th>Estado</th>
        <th>Venc. vinculado</th>
<th>Acciones</th>  
        </tr></thead>
        <tbody>
          ${page.map(r => {
    const hasVen = venByDevId.has(r.id);
    const cantDisp = r.cantDisponible;
    const esHijo = esTransferHijo(r);
    const diasDt = calcDias(r.fechaVenc);
    const esVencido = diasDt !== null && diasDt <= 0;
    const puedeTransf = cantDisp > 0 && r.estado !== 'VENDIDO' && (diasDt === null || diasDt >= 0);
    const puedeVentaAjuste = (diasDt !== null && diasDt < 0) && r.estado !== 'VENDIDO' && cantDisp > 0;
    const yaRegistrado = r.estado === 'VENDIDO' && (r.cantVendida > 0 || r.cantVencidaGondola > 0);
    const rowStyle = r.estado === 'VENDIDO'
      ? 'background:rgba(34,197,94,.06);border-left:2px solid rgba(34,197,94,.3);'
      : esVencido ? 'background:rgba(248,113,113,.04);border-left:2px solid rgba(248,113,113,.25);'
        : '';

    return `<tr style="${rowStyle}">
              <td class="c-mono" style="font-size:11px;white-space:nowrap">${fmtDateDisp(r.fecha)}</td>
              <td>
                <span class="suc-b ${getSucClass(r.sucursal)}">${esc(r.sucursal || '—')}</span>
                ${esHijo ? `<br><span class="transfer-origin-badge" style="font-size:9px">⇄ de ${esc(r.sucOrigenTransfer || '?')}</span>` : ''}
              </td>
              <td class="c-main">${esc(r.descripcion || '—')}</td>
              <td class="c-mono" style="font-size:11px">${r.ean || '—'}</td>
              <td class="c-right c-mono">${cantDisp}</td>
              <td>${vencBadge(r.fechaVenc)}</td>
              <td style="max-width:130px;overflow:hidden;text-overflow:ellipsis">${esc(r.proveedor || '—')}</td>
              <td>${motivoBadge(r.motivo)}</td>
              <td class="c-mono" style="font-size:11px">${r.lote || '—'}</td>
              <td>${estadoBadge(r.estado)}</td>
              <td>${hasVen
        ? `<span class="ven-chip" onclick="openVenDetailFromAcc('${esc(r.id)}')">📦 Ver venc. ↗</span>`
        : '<span style="color:var(--text3);font-size:10px;font-family:\'IBM Plex Mono\',monospace">—</span>'
      }</td>
              <td>
                <div style="display:flex;gap:4px;align-items:center;flex-wrap:wrap">
                  ${puedeTransf
        ? `<button onclick="abrirModalTransferenciaAcc(event,'${esc(r.id)}','${esc(r.sucursal)}',${cantDisp},'${esc(r.descripcion || '')}')"
                         class="btn-action btn-transfer-sm">⇄ Transferir</button>`
        : ''}
                  ${puedeVentaAjuste
        ? `<button onclick="abrirModalVentaAjuste(event,'${esc(r.id)}','${esc(r.descripcion || '')}',${cantDisp})"
                         class="btn-action btn-vender">📊 Venta/Ajuste</button>`
        : ''}
                  ${yaRegistrado
        ? `<span style="font-size:10px;color:#94a3b8;font-family:'IBM Plex Mono',monospace;line-height:1.5">
             ✓ ${r.cantVendida}${r.cantVendida === 1 ? ' u.' : ' u.'} vendida${r.cantVendida !== 1 ? 's' : ''}<br>
             <span style="color:#f87171">⛔ ${r.cantVencidaGondola} venc. góndola</span>
           </span>`
        : r.estado === 'VENDIDO'
          ? `<span style="font-size:10px;color:#94a3b8;font-family:'IBM Plex Mono',monospace">✓ vendido</span>`
          : ''}
                  ${esVencido && r.estado !== 'VENDIDO'
        ? `<span style="font-size:10px;color:#f87171;font-family:'IBM Plex Mono',monospace">⛔ vencido</span>`
        : ''}
                </div>
              </td>
            </tr>`;
  }).join('')}
        </tbody>
      </table>
    </div>`;

  // ── Mobile: cards ────────────────────────────────────────────
  const cardsHtml = `
    <div class="acc-mobile-cards">
      ${page.map(r => {
    const hasVen = venByDevId.has(r.id);
    const cantDisp = r.cantDisponible;
    const esHijo = esTransferHijo(r);
    const dias = calcDias(r.fechaVenc);
    const urg = getUrg(dias);
    const esVencido = dias !== null && dias <= 0;
    const puedeTransf = cantDisp > 0 && r.estado !== 'VENDIDO' && (dias === null || dias >= 0);
    const puedeVentaAjuste = (dias !== null && dias < 0) && r.estado !== 'VENDIDO' && cantDisp > 0;
    const yaRegistrado = r.estado === 'VENDIDO' && (r.cantVendida > 0 || r.cantVencidaGondola > 0);
    const cardVendidoStyle = r.estado === 'VENDIDO'
      ? 'border-color:rgba(34,197,94,.3);background:rgba(34,197,94,.04);'
      : esVencido ? 'border-color:rgba(248,113,113,.3);background:rgba(248,113,113,.04);'
        : '';
    return `
        <div class="acc-card acc-card-urg-${urg}" style="${cardVendidoStyle}">
          <div class="acc-card-top">
            <div class="acc-card-desc">${esc(r.descripcion || '—')}</div>
            ${estadoBadge(r.estado)}
          </div>
          <div class="acc-card-meta">
            <span class="suc-b ${getSucClass(r.sucursal)}">${esc(r.sucursal || '—')}</span>
            ${esHijo ? `<span class="transfer-origin-badge">⇄ de ${esc(r.sucOrigenTransfer || '?')}</span>` : ''}
            <span class="acc-card-fecha">${fmtDateDisp(r.fecha)}</span>
            ${motivoBadge(r.motivo)}
          </div>
          <div class="acc-card-grid">
            <div class="acc-card-field">
              <div class="acc-card-lbl">EAN</div>
              <div class="acc-card-val mono">${r.ean || '—'}</div>
            </div>
            <div class="acc-card-field">
              <div class="acc-card-lbl">Disponible</div>
              <div class="acc-card-val mono">${cantDisp}</div>
            </div>
            <div class="acc-card-field acc-card-field-wide">
              <div class="acc-card-lbl">Fecha vencimiento</div>
              <div class="acc-card-val">${vencBadge(r.fechaVenc)}</div>
            </div>
            <div class="acc-card-field">
              <div class="acc-card-lbl">Proveedor</div>
              <div class="acc-card-val">${esc(r.proveedor || '—')}</div>
            </div>
            ${r.lote ? `<div class="acc-card-field">
              <div class="acc-card-lbl">Lote</div>
              <div class="acc-card-val mono">${esc(r.lote)}</div>
            </div>` : ''}
          </div>
          ${(hasVen || puedeTransf || puedeVentaAjuste || esVencido || yaRegistrado) ? `
          <div class="acc-card-footer" style="display:flex;gap:6px;flex-wrap:wrap">
  ${hasVen ? `<span class="ven-chip ven-chip-full" onclick="openVenDetailFromAcc('${esc(r.id)}')">📦 Ver venc. ↗</span>` : ''}
  ${puedeTransf ? `<button onclick="abrirModalTransferenciaAcc(event,'${esc(r.id)}','${esc(r.sucursal)}',${cantDisp},'${esc(r.descripcion || '')}')" class="btn-action btn-transfer-sm">⇄ Transferir</button>` : ''}
  ${puedeVentaAjuste ? `<button onclick="abrirModalVentaAjuste(event,'${esc(r.id)}','${esc(r.descripcion || '')}',${cantDisp})" class="btn-action btn-vender">📊 Venta/Ajuste</button>` : ''}
  ${yaRegistrado ? `<span style="font-size:10px;color:#94a3b8;font-family:'IBM Plex Mono',monospace;align-self:center">✓ ${r.cantVendida}u. vend. · <span style="color:#f87171">⛔ ${r.cantVencidaGondola}u. góndola</span></span>` : r.estado === 'VENDIDO' ? `<span style="font-size:10px;color:#94a3b8;font-family:'IBM Plex Mono',monospace;align-self:center">✓ vendido</span>` : ''}
</div>` : ''}
        </div>`;
  }).join('')}
    </div>`;

  tbody.innerHTML = tableHtml + cardsHtml;
  renderPagination('accPagination', filteredAcc.length, accPage, p => { accPage = p; renderAccTable(); });
}


// ══════════════════════════════════════════════════════
//  TRANSFERENCIA DE ACCIONES ENTRE SUCURSALES
// ══════════════════════════════════════════════════════
function abrirModalTransferenciaAcc(evt, id, sucOrigen, cantDisponible, desc) {
  if (evt) evt.stopPropagation();
  const item = devData.find(d => d.id === id);
  const pesable = esPesable(item);
  currentAccTransferData = { id, sucOrigen, cantDisponible, desc, pesable };

  const select = document.getElementById('destSucursalAcc');
  const inputQty = document.getElementById('transferQtyAcc');
  const labelMax = document.getElementById('transferMaxLabelAcc');
  const infoEl = document.getElementById('transferAccInfo');

  select.innerHTML = '';
  SUCURSALES_LIST.filter(s => s !== sucOrigen).forEach(s => {
    const opt = document.createElement('option');
    opt.value = s; opt.textContent = s;
    select.appendChild(opt);
  });

  if (infoEl) infoEl.textContent = desc + ' · ' + sucOrigen;
  labelMax.textContent = 'Máximo disponible: ' + cantDisponible + (pesable ? ' kg' : ' u.');
  inputQty.max = cantDisponible;
  inputQty.value = cantDisponible;
  inputQty.step = pesable ? '0.001' : '1';
  inputQty.min = pesable ? '0.001' : '1';

  document.getElementById('modalTransferAcc').style.display = 'flex';
  setTimeout(() => inputQty.focus(), 100);
}

function cerrarModalTransferAcc() {
  document.getElementById('modalTransferAcc').style.display = 'none';
  currentAccTransferData = null;
}

async function ejecutarTransferenciaAcc() {
  const dest = document.getElementById('destSucursalAcc').value;
  const raw = document.getElementById('transferQtyAcc').value.replace(',', '.');
  const cant = parseFloat(raw);
  const { id, cantDisponible, pesable, sucOrigen } = currentAccTransferData;

  if (!cant || cant <= 0 || isNaN(cant)) {
    showToast(false, pesable ? 'Ingresá un peso válido en kg.' : 'Ingresá una cantidad válida.');
    return;
  }
  if (!pesable && !Number.isInteger(cant)) {
    showToast(false, 'Solo se aceptan cantidades enteras.');
    return;
  }
  if (cant > cantDisponible) {
    showToast(false, `Máximo disponible: ${cantDisponible}${pesable ? ' kg' : ' u.'}`);
    document.getElementById('transferQtyAcc').value = cantDisponible;
    return;
  }

  cerrarModalTransferAcc();
  showSpinner('Procesando transferencia...');
  try {
    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'transferirAccion', idOrigen: id, sucursalDestino: dest, cantidad: cant })
    });
    const json = await res.json();
    if (json.success) {
      showToast(true, `${cant}${pesable ? ' kg' : ' u.'} transferido/s a ${dest}`);
      await loadData();
    } else {
      showToast(false, json.message || 'Error en la transferencia');
    }
  } catch (e) { showToast(false, 'Error de red'); console.error(e); }
  hideSpinner();
}

// ══════════════════════════════════════════════════════
//  MODAL VENTA / AJUSTE (solo para productos VENCIDOS)
// ══════════════════════════════════════════════════════
let currentVentaAjusteData = null;

function abrirModalVentaAjuste(evt, id, desc, cantDisponible) {
  if (evt) evt.stopPropagation();
  const item = devData.find(d => d.id === id);
  const pesable = esPesable(item);
  const cantBase = parseFloat(String(cantDisponible).replace(',', '.')) || 0;
  currentVentaAjusteData = { id, desc, cantTotal: cantBase, pesable };
  document.getElementById('vaModalTitle').textContent = 'Registrar Venta / Ajuste';
  document.getElementById('vaModalDesc').textContent = desc;
  document.getElementById('vaModalCantTotal').textContent = cantBase + (pesable ? ' kg' : ' u.');
  const inp = document.getElementById('vaQty');
  inp.value = '0';
  inp.max = cantBase;
  inp.step = pesable ? '0.001' : '1';
  inp.min = '0';
  calcVentaAjuste();
  document.getElementById('modalVentaAjuste').style.display = 'flex';
  setTimeout(() => inp.focus(), 100);
}

function calcVentaAjuste() {
  if (!currentVentaAjusteData) return;
  const { cantTotal, pesable } = currentVentaAjusteData;
  const raw = (document.getElementById('vaQty').value || '0').replace(',', '.');
  const vendida = Math.max(0, Math.min(parseFloat(raw) || 0, cantTotal));
  const vencida = Math.max(0, cantTotal - vendida);
  const dec = pesable ? 3 : 0;
  document.getElementById('vaResultVendida').textContent = vendida.toFixed(dec) + (pesable ? ' kg' : ' u.');
  document.getElementById('vaResultVencida').textContent = vencida.toFixed(dec) + (pesable ? ' kg' : ' u.');
  document.getElementById('vaResultVendida').style.color = vendida > 0 ? '#4ade80' : 'var(--text3)';
  document.getElementById('vaResultVencida').style.color = vencida > 0 ? '#f87171' : 'var(--text3)';
}

function cerrarModalVentaAjuste() {
  document.getElementById('modalVentaAjuste').style.display = 'none';
  currentVentaAjusteData = null;
}

async function ejecutarVentaAjuste() {
  if (!currentVentaAjusteData) return;
  const { id, cantTotal, pesable } = currentVentaAjusteData;
  const raw = (document.getElementById('vaQty').value || '0').replace(',', '.');
  const cantVendida = parseFloat(raw) || 0;
  if (cantVendida < 0 || cantVendida > cantTotal) {
    showToast(false, 'La cantidad vendida debe estar entre 0 y ' + cantTotal);
    return;
  }
  if (!pesable && !Number.isInteger(cantVendida)) {
    showToast(false, 'Solo se aceptan cantidades enteras para productos no pesables.');
    return;
  }
  cerrarModalVentaAjuste();
  showSpinner('Registrando venta/ajuste...');
  try {
    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'marcarVendidoAccion', id, cantidadVendida: cantVendida })
    });
    const json = await res.json();
    if (json.success) {
      const vencida = json.cantVencidaGondola !== undefined ? json.cantVencidaGondola : (cantTotal - cantVendida);
      showToast(true, 'Vendidas: ' + cantVendida + ' | Venc. gondola: ' + vencida);
      await loadData();
    } else {
      showToast(false, json.message || 'Error al registrar venta/ajuste');
    }
  } catch (e) { showToast(false, 'Error de red'); console.error(e); }
  hideSpinner();
}

// ══════════════════════════════════════════════════════
//  VEN DETAIL MODAL (desde panel acciones)
//  ── Agregar estas funciones al final de dashboard.js ──
// ══════════════════════════════════════════════════════

/** Abre el modal de detalle de vencimiento vinculado a una devolucion */
function openVenDetailFromAcc(devId) {
  // 1) Buscar por idAccion exacto
  let ven = venData.find(v => v.idAccion === devId);

  // 2) Fallback: buscar por EAN + fechaVenc + sucursal del devRecord
  if (!ven) {
    const devRec = devData.find(d => d.id === devId);
    if (devRec) {
      const eanD = String(devRec.ean || '').trim();
      const fvD = fmtDateOnly(devRec.fechaVenc);
      const sucD = (devRec.sucursal || '').toUpperCase().trim();
      ven = venData.find(v =>
        String(v.ean || '').trim() === eanD &&
        fmtDateOnly(v.fechaVenc) === fvD &&
        (v.sucursal || '').toUpperCase().trim() === sucD
      );
    }
  }

  if (!ven) { showToast(false, 'No se encontró el vencimiento vinculado'); return; }
  openVenModal(ven);
}

/** Renderiza el modal con los datos de un registro de vencimiento */
function openVenModal(ven) {
  document.getElementById('vm-id').textContent = 'VENCIMIENTO · ' + (ven.id || '—');
  document.getElementById('vm-desc').textContent = ven.descripcion || '—';

  const dias = ven._dias;
  const urg = ven._urg || 'NORMAL';

  // Helper de campo
  const vff = (lbl, val, mono = false) => {
    const empty = !val || String(val).trim() === '';
    return `<div class="dev-field">
      <div class="dev-fl">${lbl}</div>
      <div class="dev-fv${mono ? ' mono' : ''}${empty ? ' empty' : ''}">${empty ? '—' : esc(String(val))}</div>
    </div>`;
  };

  const diasTxt = dias === null ? '' :
    dias > 0 ? `${dias} día${dias !== 1 ? 's' : ''} restante${dias !== 1 ? 's' : ''}` :
      dias === 0 ? 'Vence HOY' :
        `Vencido hace ${Math.abs(dias)} día${Math.abs(dias) !== 1 ? 's' : ''}`;

  document.getElementById('vm-body').innerHTML = `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;flex-wrap:wrap">
      <span class="urg ${urg}">${URG_LABELS[urg]}</span>
      ${diasTxt ? `<span style="font-family:'IBM Plex Mono',monospace;font-size:12px;color:var(--text2)">${diasTxt}</span>` : ''}
      <span class="eg ${(ven.estadoGest || 'ACTIVO').replace(/\s+/g, '-')}" style="margin-left:auto">${ven.estadoGest || 'ACTIVO'}</span>
    </div>
    <div class="dev-grid">
      ${vff('Sucursal', ven.sucursal)}
      ${vff('Usuario', ven.usuario)}
      ${vff('Fecha registro', fmtDate(ven.fechaReg || ''))}
      ${vff('Fecha vencimiento', fmtDateOnly(ven.fechaVenc), true)}
      ${vff('EAN', ven.ean, true)}
      ${vff('Cód. interno', ven.codInterno, true)}
      ${vff('Gramaje', ven.gramaje)}
      ${vff('Cantidad', ven.cantidad)}
      ${vff('Sector', ven.sector)}
      ${vff('Sección', ven.seccion)}
      ${vff('Proveedor', ven.proveedor)}
      ${vff('Lote', ven.lote, true)}
      ${ven.aclaracion ? vff('Aclaración', ven.aclaracion) : ''}
    </div>
    ${ven.idAccion ? `
    <div style="margin-top:12px">
      <div style="font-family:'IBM Plex Mono',monospace;font-size:9px;color:var(--text3);letter-spacing:1px;text-transform:uppercase;margin-bottom:5px">Acción vinculada</div>
      <span class="dev-chip" onclick="switchToDevModal('${esc(ven.idAccion)}')">${esc(ven.idAccion)} ↗</span>
    </div>` : ''}
  `;

  document.getElementById('venDetailModal').classList.add('open');
}

function closeVenModal(evt) {
  if (evt && evt.target !== document.getElementById('venDetailModal')) return;
  document.getElementById('venDetailModal').classList.remove('open');
}

/** Cierra el modal de vencimiento y abre el de devolución correspondiente */
function switchToDevModal(devId) {
  document.getElementById('venDetailModal').classList.remove('open');
  // Pequeño delay para que el cierre sea visible antes de abrir el otro modal
  setTimeout(() => {
    const fakeEvt = { stopPropagation: () => { } };
    openDevModal(fakeEvt, devId);
  }, 150);
}


// ══════════════════════════════════════════════════════
//  MÉTRICAS PANEL
// ══════════════════════════════════════════════════════
function renderMetricasPanel() {
  const d = devData;
  document.getElementById('mk-total').textContent = d.length;
  document.getElementById('mk-nc').textContent = d.filter(r => r.estado === 'N/C RECIBIDA').length;
  document.getElementById('mk-pend').textContent = d.filter(r => r.estado === 'PENDIENTE' || r.estado === 'EN GESTION').length;
  document.getElementById('mk-rech').textContent = d.filter(r => r.estado === 'RECHAZADA').length;
  document.getElementById('mk-prov').textContent = new Set(d.map(r => r.proveedor).filter(Boolean)).size;
  document.getElementById('mk-suc').textContent = new Set(d.map(r => r.sucursal).filter(Boolean)).size;
  renderBarChart('ch-motivo', countBy(d, 'motivo'), '#4f8eff');
  renderBarChart('ch-sucursal', countBy(d, 'sucursal'), '#22d87a');
  renderBarChart('ch-sector', countBy(d, 'sector'), '#a67cff');
  renderBarChart('ch-prov-nc', countBy(d.filter(r => r.estado === 'N/C RECIBIDA'), 'proveedor'), '#f5a623', 10);
}
function countBy(arr, key, limit) {
  const map = {};
  arr.forEach(r => { const v = r[key] || '(sin datos)'; map[v] = (map[v] || 0) + 1; });
  let ent = Object.entries(map).sort((a, b) => b[1] - a[1]);
  if (limit) ent = ent.slice(0, limit);
  return ent;
}
function renderBarChart(id, entries, color, maxItems = 15) {
  const el = document.getElementById(id);
  if (!entries.length) { el.innerHTML = '<div class="empty" style="padding:14px"><p>Sin datos</p></div>'; return; }
  const max = entries[0][1];
  el.innerHTML = entries.slice(0, maxItems).map(([label, val]) => `
    <div class="bar-row">
      <div class="bar-lbl" title="${esc(label)}">${esc(label)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(6, Math.round(val / max * 100))}%;background:${color}"><span class="bar-num">${val}</span></div></div>
    </div>`).join('');
}
function renderProvTable() {
  const search = (document.getElementById('prov-search')?.value || '').toLowerCase();
  const provs = [...new Set(devData.map(r => r.proveedor).filter(Boolean))].sort();
  const rows = provs.filter(p => p.toLowerCase().includes(search)).map(prov => {
    const d = devData.filter(r => r.proveedor === prov);
    const pend = d.filter(r => r.estado === 'PENDIENTE').length;
    const gest = d.filter(r => r.estado === 'EN GESTION').length;
    const nc = d.filter(r => r.estado === 'N/C RECIBIDA').length;
    const rech = d.filter(r => r.estado === 'RECHAZADA').length;
    const cob = d.length ? Math.round(nc / d.length * 100) : 0;
    return { prov, total: d.length, pend, gest, nc, rech, cob };
  });
  if (!rows.length) { document.getElementById('provTableBody').innerHTML = '<tr><td colspan="7"><div class="empty"><p>Sin datos</p></div></td></tr>'; return; }
  document.getElementById('provTableBody').innerHTML = rows.map(r => `
    <tr>
      <td style="font-weight:600;color:var(--text)">${esc(r.prov)}</td>
      <td class="c-right c-mono">${r.total}</td>
      <td class="c-right">${r.pend ? `<span class="badge b-pend">${r.pend}</span>` : '—'}</td>
      <td class="c-right">${r.gest ? `<span class="badge b-gest">${r.gest}</span>` : '—'}</td>
      <td class="c-right">${r.nc ? `<span class="badge b-nc">${r.nc}</span>` : '—'}</td>
      <td class="c-right">${r.rech ? `<span class="badge b-rech">${r.rech}</span>` : '—'}</td>
      <td style="min-width:120px"><div class="nc-progress"><div class="prog-track"><div class="prog-fill" style="width:${r.cob}%"></div></div><span class="c-mono" style="font-size:10px;min-width:32px">${r.cob}%</span></div></td>
    </tr>`).join('');
}

// ══════════════════════════════════════════════════════
//  N/C PANEL
// ══════════════════════════════════════════════════════
function applyNcFilters() {
  const prov = document.getElementById('ncf-prov')?.value || '';
  const est = document.getElementById('ncf-estado')?.value || '';
  const suc = document.getElementById('ncf-suc')?.value || '';
  const mot = document.getElementById('ncf-motivo')?.value || '';
  const search = (document.getElementById('ncf-search')?.value || '').toLowerCase();
  filteredNc = devData.filter(r => {
    // Solo registros originales — excluir hijos de transferencia DEV→DEV.
    // Los DEV creados desde el panel de Vencimientos tienen ID_ORIGEN = '' (igual que tcd-devoluciones)
    // y SÍ deben aparecer aquí para que compras/proveedores puedan gestionar el reclamo.
    if (esTransferHijo(r)) return false;
    if (prov && r.proveedor !== prov) return false;
    if (est && r.estado !== est) return false;
    if (suc && r.sucursal !== suc) return false;
    if (mot && r.motivo !== mot) return false;
    if (search) { const hay = [r.descripcion, r.ean, r.lote, r.proveedor, r.sucursal, r.id].join(' ').toLowerCase(); if (!hay.includes(search)) return false; }
    return true;
  });
  ncPage = 1;
  renderNcTable();
}
function renderNcTable() {
  const start = (ncPage - 1) * PAGE_SIZE;
  const page = filteredNc.slice(start, start + PAGE_SIZE);
  const tbody = document.getElementById('ncTableBody');
  if (!filteredNc.length) { tbody.innerHTML = '<div class="empty"><div class="ei">🔍</div><p>Sin resultados</p></div>'; document.getElementById('ncPagination').style.display = 'none'; return; }

  // Distribución actual de cantDisponible por sucursal para un registro padre
  function getDistrib(r) {
    const hijos = devData.filter(h => h.idOrigen === r.id && h.cantDisponible > 0);
    const distrib = [];
    if (r.cantDisponible > 0) distrib.push({ suc: r.sucursal, cant: r.cantDisponible });
    hijos.forEach(h => distrib.push({ suc: h.sucursal, cant: h.cantDisponible }));
    return distrib;
  }

  tbody.innerHTML = `<div class="table-scroll"><table>
    <thead><tr><th style="width:36px"></th><th>ID</th><th>Fecha</th><th>Sucursal orig.</th>
    <th>Descripción</th><th>EAN</th><th class="c-right">Cant. total</th><th>Distribución actual</th><th>Vend./Venc. góndola</th><th>Venc.</th>
    <th>Proveedor</th><th>Motivo</th><th>Lote</th><th>Estado</th><th>Obs. N/C</th></tr></thead>
    <tbody>${page.map(r => {
    const distrib = getDistrib(r);
    const tieneHijos = devData.some(h => h.idOrigen === r.id);
    const distribHtml = tieneHijos
      ? distrib.map(s =>
        `<span style="display:inline-flex;align-items:center;gap:3px;font-size:9px;
           background:rgba(79,142,255,.1);border:1px solid rgba(79,142,255,.2);
           border-radius:4px;padding:2px 6px;color:var(--text2);
           font-family:'IBM Plex Mono',monospace;white-space:nowrap;margin:1px">
           <span style="font-weight:700;color:var(--text)">${s.cant}</span>&nbsp;${esc(s.suc)}
           </span>`).join('')
      : `<span style="color:var(--text3);font-size:11px">—</span>`;
    // Desglose vendido vs vencido en góndola para NC
    const cantVend = r.cantVendida || 0;
    const cantVencG = r.cantVencidaGondola || 0;
    const tieneDesglose = r.estado === 'VENDIDO' && (cantVend > 0 || cantVencG > 0);
    const desgloseHtml = tieneDesglose
      ? `<div style="display:flex;flex-direction:column;gap:3px;min-width:130px">
           <span style="display:inline-flex;align-items:center;gap:4px;font-size:9px;
             background:rgba(74,222,128,.08);border:1px solid rgba(74,222,128,.2);
             border-radius:4px;padding:2px 7px;color:#4ade80;
             font-family:'IBM Plex Mono',monospace;white-space:nowrap">
             ✓ <strong>${cantVend}</strong>&nbsp;vendidas (${esc(r.motivo || '—')})
           </span>
           ${cantVencG > 0 ? `<span style="display:inline-flex;align-items:center;gap:4px;font-size:9px;
             background:rgba(248,113,113,.08);border:1px solid rgba(248,113,113,.2);
             border-radius:4px;padding:2px 7px;color:#f87171;
             font-family:'IBM Plex Mono',monospace;white-space:nowrap">
             ⛔ <strong>${cantVencG}</strong>&nbsp;venc. góndola
           </span>` : ''}
         </div>`
      : r.estado === 'VENDIDO'
        ? `<span style="color:#94a3b8;font-size:11px;font-family:'IBM Plex Mono',monospace">✓ vendido</span>`
        : `<span style="color:var(--text3);font-size:11px">—</span>`;
    const rowStyle = r.estado === 'VENDIDO'
      ? 'background:rgba(34,197,94,.06);border-left:3px solid rgba(34,197,94,.35);'
      : '';
    return `
      <tr style="${rowStyle}">
        <td><input type="checkbox" ${selectedIds.has(r.id) ? 'checked' : ''} onchange="toggleRow('${r.id}',this)"></td>
        <td class="c-id">${r.id}</td>
        <td class="c-mono" style="font-size:11px">${fmtDateDisp(r.fecha)}</td>
        <td><span class="b-tag">${r.sucursal || '—'}</span></td>
        <td class="c-main">${esc(r.descripcion || '—')}</td>
        <td class="c-mono" style="font-size:11px">${r.ean || '—'}</td>
        <td class="c-right c-mono" style="font-weight:700">${r.cantidad || '—'}</td>
        <td style="min-width:130px">${distribHtml}</td>
        <td style="min-width:140px">${desgloseHtml}</td>
        <td>${vencBadge(r.fechaVenc)}</td>
        <td>${esc(r.proveedor || '—')}</td>
        <td>${motivoBadge(r.motivo)}</td>
        <td class="c-mono" style="font-size:11px">${r.lote || '—'}</td>
        <td>${estadoBadge(r.estado)}</td>
        <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;color:var(--text2);font-size:11px">${r.obsNC || '—'}</td>
      </tr>`;
  }).join('')}
    </tbody></table></div>`;
  renderPagination('ncPagination', filteredNc.length, ncPage, p => { ncPage = p; renderNcTable(); });
}
function toggleRow(id, cb) { if (cb.checked) selectedIds.add(id); else selectedIds.delete(id); updateBulkBar(); }
function toggleSelectAll() { const c = document.getElementById('selectAll').checked; filteredNc.forEach(r => { if (c) selectedIds.add(r.id); else selectedIds.delete(r.id); }); renderNcTable(); updateBulkBar(); }
function clearSelection() { selectedIds.clear(); document.getElementById('selectAll').checked = false; renderNcTable(); updateBulkBar(); }
function updateBulkBar() { const n = selectedIds.size; document.getElementById('bulkCount').textContent = n + ' registro' + (n === 1 ? '' : 's') + ' seleccionado' + (n === 1 ? '' : 's'); document.getElementById('bulkBar').classList.toggle('visible', n > 0); }

function applyBulkAction() {
  const estado = document.getElementById('bulkEstado').value;
  const obs = document.getElementById('bulkObs').value.trim();
  if (!estado && !obs) { alert('Seleccioná un estado o escribí una observación.'); return; }
  if (!selectedIds.size) { alert('No hay registros seleccionados.'); return; }
  const ids = [...selectedIds];
  document.getElementById('cm-title').textContent = '✏️ Confirmar actualización masiva';
  document.getElementById('cm-text').textContent = `Se actualizarán ${ids.length} registro${ids.length === 1 ? '' : 's'}:`;
  document.getElementById('cm-details').innerHTML = `
    ${estado ? `<div style="background:var(--blue-d);border:1px solid var(--blue);padding:8px 12px;border-radius:7px;font-size:12px;margin-bottom:6px">Estado → <strong>${esc(estado)}</strong></div>` : ''}
    ${obs ? `<div style="background:var(--blue-d);border:1px solid var(--blue);padding:8px 12px;border-radius:7px;font-size:12px">Obs N/C → <em>${esc(obs)}</em></div>` : ''}
    <p style="font-size:10px;color:var(--text3);margin-top:8px;font-family:'IBM Plex Mono',monospace">IDs: ${ids.slice(0, 8).join(', ')}${ids.length > 8 ? '…' : ''}</p>`;
  modalCb = () => executeBulkUpdate(ids, estado, obs);
  document.getElementById('cm-confirm').onclick = () => { if (modalCb) { const cb = modalCb; modalCb = null; cb(); } };
  document.getElementById('confirmModal').classList.add('open');
}

async function executeBulkUpdate(ids, estado, obs) {
  document.getElementById('confirmModal').classList.remove('open');
  modalCb = null;
  showSpinner('Procesando 0 / ' + ids.length + '…');
  let ok = 0, err = 0;
  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    showSpinner('Procesando ' + (i + 1) + ' / ' + ids.length + '…');
    try {
      const payload = { action: 'updateRecord', id };
      if (estado) payload.estado = estado;
      if (obs) payload.observacionNC = obs;
      const res = await fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload), headers: { 'Content-Type': 'text/plain;charset=utf-8' } });
      const json = await res.json();
      if (json.success) { ok++; const r = devData.find(x => x.id === id); if (r) { if (estado) r.estado = estado; if (obs) r.obsNC = obs; } }
      else err++;
    } catch (e) { err++; }
  }
  hideSpinner();
  clearSelection();
  applyNcFilters(); renderMetricasPanel(); renderProvTable(); applyAccFilters(); updateResumenKPIs(); updateNavBadges();
  showToast(err === 0, ok + ' actualizado' + (ok === 1 ? '' : 's') + (err ? ' · ' + err + ' error' + (err === 1 ? '' : 's') : ''));
}

function closeModal() { document.getElementById('confirmModal').classList.remove('open'); modalCb = null; }
document.getElementById('confirmModal').addEventListener('click', e => { if (e.target === document.getElementById('confirmModal')) closeModal(); });

// ══════════════════════════════════════════════════════
//  DEV DETAIL MODAL
// ══════════════════════════════════════════════════════
function openDevModal(evt, id) {
  evt.stopPropagation();
  const rec = devMap.get(id);
  document.getElementById('dd-id').textContent = id;
  document.getElementById('dd-desc').textContent = rec ? (rec.descripcion || '—') : '—';
  const body = document.getElementById('dd-body');
  if (!rec) { body.innerHTML = `<div style="color:var(--text3);font-size:12px;text-align:center;padding:20px">No se encontró <strong>${esc(id)}</strong> en BD Devoluciones.</div>`; document.getElementById('devDetailModal').classList.add('open'); return; }
  const estadoCls = (rec.estado || 'PENDIENTE').replace(/[\s\/]/g, '-');
  const fotoMatch = String(rec.fotoRaw || '').match(/https?:\/\/[^"')]+/);
  const nc = rec.obsNC || '';
  const f = (lbl, val, opts = {}) => {
    const isEmpty = !val || val === '';
    return `<div class="dev-field${opts.full ? ' full' : ''}"><div class="dev-fl">${lbl}</div><div class="dev-fv${opts.mono ? ' mono' : ''}${isEmpty ? ' empty' : ''}">${isEmpty ? '—' : esc(String(val))}</div></div>`;
  };
  body.innerHTML = `
    <div class="dev-estado-badge ${estadoCls}">${rec.estado || 'PENDIENTE'}</div>
    <div class="dev-grid">
      ${f('Sucursal', rec.sucursal)}${f('Usuario', rec.usuario)}
      ${f('Fecha registro', fmtDateDisp(rec.fecha))}${f('Motivo', rec.motivo)}
      ${f('EAN', rec.ean, { mono: true })}${f('Cód. interno', rec.codInterno, { mono: true })}
      ${f('Gramaje', rec.gramaje)}${f('Cantidad', rec.cantidad)}
      ${f('Fecha vencimiento', fmtVenc(rec.fechaVenc), { mono: true })}${f('Sector/Sección', [rec.sector, rec.seccion].filter(Boolean).join(' / '))}
      ${f('Proveedor', rec.proveedor)}${f('Cód. proveedor', rec.codProv, { mono: true })}
      ${f('Lote', rec.lote, { mono: true })}${f('Aclaración', rec.aclaracion)}
      ${rec.comentarios ? f('Comentarios', rec.comentarios, { full: true }) : ''}
    </div>
    ${fotoMatch ? `<a class="photo-link" href="${fotoMatch[0]}" target="_blank">📎 Ver foto adjunta</a>` : ''}
    ${nc ? `<div class="nc-obs-box"><div class="nc-obs-lbl">Observación N/C</div><div style="font-size:12px;color:var(--text)">${esc(nc)}</div></div>` : ''}
  `;
  document.getElementById('devDetailModal').classList.add('open');
}
function closeDevModal(evt) { if (evt && evt.target !== document.getElementById('devDetailModal')) return; document.getElementById('devDetailModal').classList.remove('open'); }

// ══════════════════════════════════════════════════════
//  POPULATE SELECTS
// ══════════════════════════════════════════════════════
function populateFilterSelects() {
  const provs = [...new Set([...devData.map(r => r.proveedor), ...venData.map(r => r.proveedor)].filter(Boolean))].sort();
  const sucs = [...new Set([...devData.map(r => r.sucursal), ...venData.map(r => r.sucursal)].filter(Boolean))].sort();
  const motivos = [...new Set(devData.map(r => r.motivo).filter(Boolean))].sort();
  const venSucs = [...new Set(venData.map(r => r.sucursal).filter(Boolean))].sort();
  const venProvs = [...new Set(venData.map(r => r.proveedor).filter(Boolean))].sort();
  fillSel('af-suc', sucs); fillSel('af-prov', provs); fillSel('af-motivo', motivos);
  fillSel('vf-prov', venProvs); fillSel('vf-suc', venSucs);
  fillSel('ncf-prov', provs); fillSel('ncf-suc', sucs); fillSel('ncf-motivo', motivos);
}
function fillSel(id, vals) { const s = document.getElementById(id); if (!s) return; const cur = s.value; s.innerHTML = '<option value="">Todos/Todas</option>'; vals.forEach(v => s.add(new Option(v, v))); s.value = cur; }

// ══════════════════════════════════════════════════════
//  STALE BADGE
// ══════════════════════════════════════════════════════
function staleStyle(dias) {
  if (dias === null) return { fg: '#525e72', bg: '#1c2030', br: '#252a38', icon: '?' };
  if (dias <= 3) return { fg: '#4ade80', bg: '#0a1f0a', br: '#14532d', icon: '✓' };
  if (dias <= 7) return { fg: '#a3e635', bg: '#172100', br: '#2d4a00', icon: '●' };
  if (dias <= 14) return { fg: '#d4e635', bg: '#1e2200', br: '#3d4700', icon: '●' };
  if (dias <= 21) return { fg: '#fbbf24', bg: '#2d2200', br: '#78350f', icon: '●' };
  if (dias <= 30) return { fg: '#fb923c', bg: '#2d1500', br: '#7c2d12', icon: '!' };
  if (dias <= 45) return { fg: '#f87171', bg: '#2d0c0c', br: '#7f1d1d', icon: '!!' };
  return { fg: '#ff3333', bg: '#3d0000', br: '#9b1c1c', icon: '⚠' };
}
function staleBadgeHtml(dias, mini = false) {
  const s = staleStyle(dias);
  const txt = s.icon + ' ' + (dias === null ? 'sin fecha' : `hace ${dias}d`);
  return `<span class="stale" style="color:${s.fg};background:${s.bg};border-color:${s.br};font-size:${mini ? '9px' : '10px'}">${txt}</span>`;
}

// ══════════════════════════════════════════════════════
//  COMBINED RISK (días × cantidad)
// ══════════════════════════════════════════════════════
/**
 * Calcula el riesgo combinado considerando tanto los días al vencimiento
 * como la cantidad de unidades. Un producto con muchas unidades y
 * bastantes días puede ser igual de riesgoso que pocos días con pocas unidades.
 *
 * Factor multiplicador según días restantes:
 *   vencido   → ×10   (cualquier cantidad es crítica)
 *   1–6 días  → ×4    (crítico)
 *   7–14 días → ×2    (urgente)
 *   15–21 días→ ×1    (normal-urgente)
 *   22–45 días→ ×0.4  (necesita bastante cantidad para alertar)
 *   46–90 días→ ×0.15 (necesita cantidad muy alta para alertar)
 *   > 90 días → ×0.03
 */
function getCombinedRisk(dias, qty) {
  if (!qty || qty <= 0) return 'NORMAL';
  if (dias === null) return 'NORMAL';

  let factor;
  if (dias <= 0) factor = 10;
  else if (dias <= 6) factor = 4;
  else if (dias <= 14) factor = 2;
  else if (dias <= 21) factor = 1;
  else if (dias <= 45) factor = 0.4;
  else if (dias <= 90) factor = 0.15;
  else factor = 0.03;

  const score = qty * factor;

  if (score >= 1500) return 'CRITICO';
  if (score >= 400) return 'URGENTE';
  if (score >= 80) return 'PROXIMO';
  if (score >= 20) return 'ATENCION';
  return 'NORMAL';
}

/**
 * Renderiza la celda de cantidad con color + badge de riesgo combinado.
 * Solo muestra badge cuando el riesgo combinado es >= PROXIMO.
 * @param {number} qty       - cantidad de unidades
 * @param {number|null} dias - días al vencimiento
 * @param {string} size      - 'normal' | 'small' para distintos contextos
 */
function qtyRiskHtml(qty, dias, size = 'normal') {
  const risk = getCombinedRisk(dias, qty);
  const showBadge = risk !== 'NORMAL';

  const RISK_LABELS = {
    CRITICO: '🔴 RIESGO ALTO',
    URGENTE: '🟠 RIESGO',
    PROXIMO: '🟡 ATENDER',
    ATENCION: '🔵 MONITOREAR',
  };

  if (size === 'small') {
    // Para suc-card
    return `<div class="sf-qty-wrap">
      <span class="sf-qty-num qr-${risk}" id="${arguments[3] ? 'qty-' + arguments[3] : ''}">${qty}</span>
      ${showBadge ? `<span class="qty-risk-badge qr-${risk}">${RISK_LABELS[risk]}</span>` : ''}
    </div>`;
  }

  // Para tabla principal
  return `<div class="qty-cell">
    <span class="qty-num qr-${risk}">${qty}</span>
    ${showBadge ? `<span class="qty-risk-badge qr-${risk}">${RISK_LABELS[risk]}</span>` : `<span style="font-family:'IBM Plex Mono',monospace;font-size:9px;color:var(--text3)">u.</span>`}
  </div>`;
}

// ══════════════════════════════════════════════════════
//  BADGE HELPERS
// ══════════════════════════════════════════════════════
function estadoBadge(est) {
  const map = {
    'PENDIENTE': 'b-pend',
    'EN GESTION': 'b-gest',
    'N/C RECIBIDA': 'b-nc',
    'RECHAZADA': 'b-rech',
    'VENDIDO': 'b-vendido'
  };
  return `<span class="badge ${map[est] || 'b-pend'}">${esc(est || 'PENDIENTE')}</span>`;
}

function motivoBadge(m) {
  if (!m) return `<span class="mb mb-otro">—</span>`;
  const u = m.toUpperCase();
  let cls = 'mb-otro';
  if (u === 'ACCION 2X1' || u === 'VENCIDO ACCION 2X1') cls = 'mb-2x1';
  else if (u === 'ACCION 50% OFF' || u === 'VENCIDO ACCION 50% OFF') cls = 'mb-50off';
  else if (u === 'OTRO DESCUENTO') cls = 'mb-desc';
  else if (u === 'VENCIDO') cls = 'mb-venc';
  else if (u === 'ROTO/DAÑADO' || u === 'MAL ESTADO') cls = 'mb-dano';
  return `<span class="mb ${cls}">${esc(m)}</span>`;
}
function vencBadge(str) {
  const txt = fmtVenc(str);
  if (txt === '—') return `<span class="vb vb-nd">—</span>`;
  const parts = txt.split('-');
  if (parts.length !== 3) return `<span class="vb-nd">${txt}</span>`;
  const d = new Date(+parts[2], +parts[1] - 1, +parts[0]);
  const diff = Math.ceil((d - new Date()) / 86400000);
  let cls = 'vb-ok';
  if (diff < 0) cls = 'vb-exp'; else if (diff <= 7) cls = 'vb-crit'; else if (diff <= 30) cls = 'vb-warn';
  const icon = diff < 0 ? '💀' : diff <= 7 ? '🔴' : diff <= 30 ? '🟡' : '🟢';
  return `<span class="vb ${cls}">${icon} ${txt}</span>`;
}

// ══════════════════════════════════════════════════════
//  DATE / FORMAT HELPERS
// ══════════════════════════════════════════════════════
function parseDate(str) {
  if (!str) return null;
  const s = String(str).trim(); let m;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/); if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}
function parseRegDate(str) {
  if (!str) return null;
  const m = str.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], +m[6]);
  const m2 = str.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m2) return new Date(+m2[3], +m2[2] - 1, +m2[1]);
  const d = new Date(str); return isNaN(d) ? null : d;
}
function fmtDate(v) {
  if (!v) return '—';
  if (v instanceof Date) return v.toLocaleString('es-AR', { day: '2-digit', month: '2-digit', year: '2-digit', hour: '2-digit', minute: '2-digit' });
  const s = String(v).trim();
  const m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}:\d{2}))?/);
  if (m1) return m1[4] ? `${m1[1]}-${m1[2]}-${m1[3]} ${m1[4]}` : `${m1[1]}-${m1[2]}-${m1[3]}`;
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[\sT]+(\d{2}:\d{2}))?/);
  if (m2) return m2[4] ? `${m2[3]}-${m2[2]}-${m2[1]} ${m2[4]}` : `${m2[3]}-${m2[2]}-${m2[1]}`;
  return s.slice(0, 16);
}
function fmtDateOnly(str) { if (!str) return '—'; return fmtDate(str).split(' ')[0]; }
function fmtDateDisp(d) { if (!d) return '—'; if (d instanceof Date) return d.toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: '2-digit' }); return fmtDate(d); }
function fmtVenc(str) {
  if (!str) return '—';
  const s = String(str).trim(); let m;
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return m[3].padStart(2, '0') + '-' + m[2].padStart(2, '0') + '-' + m[1];
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/); if (m) return m[1].padStart(2, '0') + '-' + m[2].padStart(2, '0') + '-' + m[3];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return m[1].padStart(2, '0') + '-' + m[2].padStart(2, '0') + '-' + m[3];
  return s;
}
function esc(s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

// ══════════════════════════════════════════════════════
//  PAGINATION
// ══════════════════════════════════════════════════════
const _pageCbs = {};
function renderPagination(containerId, total, current, cb) {
  const el = document.getElementById(containerId);
  const tp = Math.ceil(total / PAGE_SIZE);
  if (tp <= 1) { el.style.display = 'none'; return; }
  _pageCbs[containerId] = cb;
  el.style.display = 'flex';
  const shown = Math.min(current * PAGE_SIZE, total);
  let html = `<span>Mostrando ${(current - 1) * PAGE_SIZE + 1}–${shown} de ${total}</span><div class="page-btns">`;
  html += `<button class="page-btn" data-pg="${current - 1}" data-cb="${containerId}" ${current === 1 ? 'disabled' : ''}>←</button>`;
  for (let i = Math.max(1, current - 2); i <= Math.min(tp, current + 2); i++) html += `<button class="page-btn${i === current ? ' active' : ''}" data-pg="${i}" data-cb="${containerId}">${i}</button>`;
  html += `<button class="page-btn" data-pg="${current + 1}" data-cb="${containerId}" ${current === tp ? 'disabled' : ''}>→</button></div>`;
  el.innerHTML = html;
  el.querySelectorAll('button[data-pg]').forEach(btn => btn.addEventListener('click', () => { if (!btn.disabled) { const c = _pageCbs[btn.dataset.cb]; if (c) c(+btn.dataset.pg); } }));
}

// ══════════════════════════════════════════════════════
//  EXPORT ACCIONES
// ══════════════════════════════════════════════════════
function barcodeDataUrl(ean) {
  const canvas = document.createElement('canvas');
  const code = String(ean || '').replace(/\D/g, '');
  if (!code) return null;
  const fmt = /^\d{13}$/.test(code) ? 'ean13' : (/^\d{8}$/.test(code) ? 'ean8' : 'code128');
  try { JsBarcode(canvas, code, { format: fmt, displayValue: false, margin: 6, width: 3, height: 90 }); }
  catch (e) { try { JsBarcode(canvas, code, { format: 'code128', displayValue: false, margin: 6, width: 3, height: 90 }); } catch (e2) { return null; } }
  return canvas.toDataURL('image/png').split(',')[1];
}
async function exportAccXlsx() {
  const rows = filteredAcc;
  if (!rows.length) { alert('No hay registros para exportar.'); return; }
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Acciones');
  ws.pageSetup = { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: .3, right: .3, top: .4, bottom: .4, header: .2, footer: .2 } };
  ws.columns = [{ key: 'bc', width: 26 }, { key: 'fecha', width: 12 }, { key: 'suc', width: 16 }, { key: 'desc', width: 34 }, { key: 'gram', width: 10 }, { key: 'cant', width: 10 }, { key: 'venc', width: 13 }, { key: 'prov', width: 24 }, { key: 'mot', width: 22 }, { key: 'lote', width: 14 }, { key: 'est', width: 16 }];
  const hRow = ws.addRow(['COD-BAR', 'Fecha', 'Sucursal', 'Descripción', 'Gramaje', 'Cantidad', 'Fecha Venc.', 'Proveedor', 'Motivo', 'Lote', 'Estado']);
  hRow.height = 22;
  hRow.eachCell(c => { c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D111F' } }; c.alignment = { horizontal: 'center', vertical: 'middle' }; c.border = { bottom: { style: 'medium', color: { argb: 'FF4f8eff' } } }; });
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i]; const isEven = i % 2 === 1;
    const dRow = ws.addRow(['', r.fecha ? r.fecha.toLocaleDateString('es-AR') : '', r.sucursal || '', r.descripcion || '', r.gramaje || '', r.cantidad || '', fmtVenc(r.fechaVenc), r.proveedor || '', r.motivo || '', r.lote || '', r.estado || '']);
    dRow.height = 70;
    dRow.eachCell({ includeEmpty: true }, (c, col) => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEven ? 'FFF5F7FF' : 'FFFFFFFF' } }; c.alignment = { vertical: 'middle', horizontal: col === 1 ? 'center' : 'left', wrapText: col === 4 }; c.font = { size: 10 }; c.border = { bottom: { style: 'thin', color: { argb: 'FFDDDDDD' } }, right: { style: 'thin', color: { argb: 'FFDDDDDD' } } }; });
    const ec = dRow.getCell(11); const ecMap = { 'PENDIENTE': { bg: 'FFFFF3CD', fg: 'FF856404' }, 'EN GESTION': { bg: 'FFD1ECF1', fg: 'FF0C5460' }, 'N/C RECIBIDA': { bg: 'FFD4EDDA', fg: 'FF155724' }, 'RECHAZADA': { bg: 'FFF8D7DA', fg: 'FF721C24' } };
    const ecv = ecMap[r.estado]; if (ecv) { ec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ecv.bg } }; ec.font = { size: 10, bold: true, color: { argb: ecv.fg } }; }
    const b64 = barcodeDataUrl(r.ean); if (b64) { const imgId = wb.addImage({ base64: b64, extension: 'png' }); ws.addImage(imgId, { tl: { col: .08, row: i + 1 + .06 }, ext: { width: 145, height: 60 }, editAs: 'oneCell' }); }
  }
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];
  try {
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'acciones_' + fmtIso(new Date()) + '.xlsx'; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  } catch (err) { alert('Error al generar Excel: ' + err.message); }
}
async function exportNcXlsx() {
  const rows = selectedIds.size > 0 ? filteredNc.filter(r => selectedIds.has(r.id)) : filteredNc;
  if (!rows.length) { alert('No hay registros para exportar.'); return; }

  // ── Para cada fila calcular totales consolidados (padre + hijos transferidos) ──
  const rowsConsolidados = rows.map(r => {
    const hijosAll = devData.filter(h => h.idOrigen === r.id);
    const totalVend = (r.cantVendida || 0) + hijosAll.reduce((s, h) => s + (h.cantVendida || 0), 0);
    const totalVencG = (r.cantVencidaGondola || 0) + hijosAll.reduce((s, h) => s + (h.cantVencidaGondola || 0), 0);
    return { r, totalVend, totalVencG };
  });

  const wb = new ExcelJS.Workbook();
  wb.creator = 'NEXUS'; wb.created = new Date();
  const ws = wb.addWorksheet('Gestión NC');

  ws.pageSetup = {
    paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    margins: { left: .3, right: .3, top: .4, bottom: .4, header: .2, footer: .2 }
  };

  ws.columns = [
    { key: 'sucursal', width: 16 },
    { key: 'descripcion', width: 36 },
    { key: 'ean', width: 16 },
    { key: 'cantTotal', width: 12 },
    { key: 'vendidas', width: 12 },
    { key: 'vencGondola', width: 14 },
    { key: 'venc', width: 14 },
    { key: 'proveedor', width: 26 },
    { key: 'motivo', width: 22 },
    { key: 'lote', width: 14 },
  ];

  // ── Encabezado ──────────────────────────────────────────
  const hRow = ws.addRow([
    'Sucursal orig.', 'Descripción', 'EAN',
    'Cant. total', 'Vendidas', 'Venc. góndola',
    'Vencimiento', 'Proveedor', 'Motivo', 'Lote'
  ]);
  hRow.height = 24;
  hRow.eachCell(c => {
    c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D111F' } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = { bottom: { style: 'medium', color: { argb: 'FF4f8eff' } } };
  });

  // ── Datos ────────────────────────────────────────────────
  rowsConsolidados.forEach(({ r, totalVend, totalVencG }, i) => {
    const isEven = i % 2 === 1;
    const dRow = ws.addRow([
      r.sucursal || '—',
      r.descripcion || '—',
      r.ean || '—',
      r.cantidad || 0,
      totalVend,
      totalVencG,
      fmtVenc(r.fechaVenc),
      r.proveedor || '—',
      r.motivo || '—',
      r.lote || '—',
    ]);
    dRow.height = 20;

    dRow.eachCell({ includeEmpty: true }, (c, col) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEven ? 'FFF5F7FF' : 'FFFFFFFF' } };
      c.alignment = { vertical: 'middle', horizontal: [4, 5, 6].includes(col) ? 'center' : 'left', wrapText: col === 2 };
      c.font = { size: 10 };
      c.border = { bottom: { style: 'thin', color: { argb: 'FFDDDDDD' } }, right: { style: 'thin', color: { argb: 'FFDDDDDD' } } };
    });

    // Columna Vendidas → verde
    const cellVend = dRow.getCell(5);
    if (totalVend > 0) {
      cellVend.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD4EDDA' } };
      cellVend.font = { size: 10, bold: true, color: { argb: 'FF155724' } };
    }

    // Columna Venc. góndola → rojo si > 0
    const cellVencG = dRow.getCell(6);
    if (totalVencG > 0) {
      cellVencG.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8D7DA' } };
      cellVencG.font = { size: 10, bold: true, color: { argb: 'FF721C24' } };
    }

    // Columna Vencimiento: tachado si ya venció
    const vencCell = dRow.getCell(7);
    const diasV = calcDias(r.fechaVenc);
    if (diasV !== null && diasV <= 0) {
      vencCell.font = { size: 10, strike: true, color: { argb: 'FFFF4444' } };
      vencCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3D0000' } };
    }

    // Columna Motivo → color badge
    const motivoCell = dRow.getCell(9);
    const mU = (r.motivo || '').toUpperCase();
    if (mU.includes('2X1')) { motivoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } }; motivoCell.font = { size: 10, bold: true, color: { argb: 'FF93C5FD' } }; }
    else if (mU.includes('50%')) { motivoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B1F6A' } }; motivoCell.font = { size: 10, bold: true, color: { argb: 'FFCA9FFC' } }; }
    else if (mU.includes('VENC')) { motivoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3D0000' } }; motivoCell.font = { size: 10, bold: true, color: { argb: 'FFFF4444' } }; }
    else if (mU.includes('ROT') || mU.includes('DA') || mU.includes('ESTADO')) { motivoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3D2200' } }; motivoCell.font = { size: 10, bold: true, color: { argb: 'FFFB923C' } }; }
  });

  // ── Fila de totales ──────────────────────────────────────
  const sumCant = rowsConsolidados.reduce((s, { r }) => s + (parseFloat(String(r.cantidad || 0).replace(',', '.')) || 0), 0);
  const sumVend = rowsConsolidados.reduce((s, { totalVend }) => s + totalVend, 0);
  const sumVencG = rowsConsolidados.reduce((s, { totalVencG }) => s + totalVencG, 0);

  const totRow = ws.addRow(['', `TOTAL: ${rows.length} registros`, '', sumCant, sumVend, sumVencG, '', '', '', '']);
  totRow.height = 22;
  totRow.eachCell({ includeEmpty: true }, (c, col) => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D111F' } };
    c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
    c.alignment = { vertical: 'middle', horizontal: [4, 5, 6].includes(col) ? 'center' : 'left' };
  });
  const tVend = totRow.getCell(5);
  tVend.font = { bold: true, color: { argb: 'FF4ADE80' }, size: 10 };
  const tVencG = totRow.getCell(6);
  tVencG.font = { bold: true, color: { argb: 'FFF87171' }, size: 10 };

  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];

  try {
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'gestion_nc_' + fmtIso(new Date()) + '.xlsx';
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  } catch (err) { alert('Error al generar Excel: ' + err.message); console.error(err); }
}
function fmtIso(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }

// ══════════════════════════════════════════════════════
//  SPINNER / TOAST
// ══════════════════════════════════════════════════════
function showSpinner(msg) { document.getElementById('spinnerMsg').textContent = msg || 'Procesando…'; document.getElementById('spinnerOverlay').classList.add('active'); }
function hideSpinner() { document.getElementById('spinnerOverlay').classList.remove('active'); }
function showToast(isOk, msg) {
  const t = document.createElement('div'); t.className = 'toast ' + (isOk ? 'success' : 'error'); t.textContent = (isOk ? '✅ ' : '⚠️ ') + msg;
  document.body.appendChild(t); setTimeout(() => { if (t.parentNode) t.remove(); }, 5000);
}

// ══════════════════════════════════════════════════════
//  CONFIG
// ══════════════════════════════════════════════════════
function openConfig() { document.getElementById('urlInput').value = SCRIPT_URL; document.getElementById('configModal').classList.add('open'); }
function saveConfig() {
  const url = document.getElementById('urlInput').value.trim();
  if (!url.startsWith('http')) { document.getElementById('urlInput').style.borderColor = 'var(--red)'; return; }
  SCRIPT_URL = url; localStorage.setItem('nexus_script_url', url);
  document.getElementById('configModal').classList.remove('open');
  loadData();
}

// ══════════════════════════════════════════════════════
//  AJUSTE DE STOCK — con soporte pesable
// ══════════════════════════════════════════════════════
function ajustarStock(evt, id, delta) {
  if (evt) evt.stopPropagation();
  const elQty = document.getElementById(`qty-${id}`);
  const cantActual = elQty ? (parseFloat(elQty.textContent) || 0) : 0;
  const item = venData.find(v => v.id === id);
  const suc = item ? item.sucursal : '';
  const pesable = esPesable(item);

  currentAdjustData = { id, delta, cantActual, suc, pesable };

  const isPos = delta > 0;
  document.getElementById('adjModalTitle').textContent = isPos ? '＋ Ajuste positivo' : '－ Ajuste negativo';
  document.getElementById('adjModalTitle').style.color = isPos ? 'var(--green)' : 'var(--red)';
  document.getElementById('adjModalSubtitle').textContent =
    `${esc(suc)}${pesable ? ' · ⚖ Pesable — ingresá en kg' : ' · Sin restricciones de cantidad'}`;
  document.getElementById('adjCurrentStock').textContent = cantActual;

  const input = document.getElementById('adjustQty');
  input.value = pesable ? '0.001' : 1;
  input.step = pesable ? '0.001' : '1';
  input.min = pesable ? '0.001' : '1';

  const btn = document.getElementById('adjConfirmBtn');
  btn.style.background = isPos ? 'var(--green)' : 'var(--red)';
  btn.style.color = isPos ? '#000' : '#fff';

  document.getElementById('modalAdjust').style.display = 'flex';
  setTimeout(() => {
    input.focus();
    const ley = document.getElementById('adjustLeyenda');
    if (ley) ley.value = '';
  }, 100);
}

// ══════════════════════════════════════════════════════
//  MODAL TRANSFERENCIA — con soporte pesable
// ══════════════════════════════════════════════════════
function abrirModalTransferencia(evt, id, origen, max) {
  if (evt) evt.stopPropagation();
  const item = venData.find(v => v.id === id);
  const pesable = esPesable(item);
  currentTransferData = { id, origen, max, pesable };

  const modal = document.getElementById('modalTransfer');
  const select = document.getElementById('destSucursal');
  const inputQty = document.getElementById('transferQty');
  const labelMax = document.getElementById('transferMaxLabel');

  if (!modal || !select) { console.error('No se encontró el modal de transferencia'); return; }

  select.innerHTML = '';
  ['HIPER', 'CENTRO', 'RIBERA', 'MAYORISTA', 'PROVEEDOR', 'OTRO']
    .filter(s => s !== origen)
    .forEach(s => {
      const opt = document.createElement('option');
      opt.value = s; opt.textContent = s;
      select.appendChild(opt);
    });

  labelMax.textContent = `Disponible: ${max}${pesable ? ' kg' : ' u.'}`;
  inputQty.max = max;
  inputQty.value = max;
  inputQty.step = pesable ? '0.001' : '1';
  inputQty.min = pesable ? '0.001' : '1';

  const title = modal.querySelector('h3');
  if (title) title.textContent = `⇄ Transferir Stock${pesable ? ' ⚖ (kg)' : ''}`;

  modal.style.display = 'flex';
}

function cerrarModalTransfer() { const modal = document.getElementById('modalTransfer'); if (modal) modal.style.display = 'none'; }

// ══════════════════════════════════════════════════════
//  EJECUTAR TRANSFERENCIA — con soporte pesable
// ══════════════════════════════════════════════════════
async function ejecutarTransferencia() {
  const dest = document.getElementById('destSucursal').value;
  const raw = document.getElementById('transferQty').value.replace(',', '.');
  const cant = parseFloat(raw);
  const { max, pesable } = currentTransferData;

  if (!cant || cant <= 0 || isNaN(cant)) {
    showToast(false, pesable ? 'Ingresá un peso válido en kg.' : 'Ingresá una cantidad válida.');
    return;
  }
  if (!pesable && !Number.isInteger(cant)) {
    showToast(false, 'Este producto solo acepta cantidades enteras.');
    return;
  }
  if (cant > max) {
    showToast(false, `No podés transferir más de ${max}${pesable ? ' kg' : ' u.'}`);
    document.getElementById('transferQty').value = max;
    document.getElementById('transferQty').focus();
    return;
  }

  showSpinner('Procesando transferencia...'); cerrarModalTransfer();
  try {
    const res = await fetch(SCRIPT_URL, {
      method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'transferStock', idOrigen: currentTransferData.id, sucursalDestino: dest, cantidad: cant })
    });
    const result = await res.json();
    if (result.success) {
      showToast(true, `Transferencia: ${cant}${pesable ? ' kg' : ' u.'} → ${dest}`);
      await loadData();
    } else {
      showToast(false, result.message || 'Error en la operación');
    }
  } catch (e) { showToast(false, 'Error de red'); console.error(e); }
  hideSpinner();
}

function cerrarModalAdjust() { document.getElementById('modalAdjust').style.display = 'none'; currentAdjustData = null; }

// ══════════════════════════════════════════════════════
//  EJECUTAR AJUSTE — con soporte pesable
// ══════════════════════════════════════════════════════
async function ejecutarAjuste() {
  const { id, delta, cantActual, pesable } = currentAdjustData;
  const raw = document.getElementById('adjustQty').value.replace(',', '.');
  const qty = parseFloat(raw);

  if (!qty || qty <= 0 || isNaN(qty)) {
    showToast(false, pesable ? 'Ingresá un peso válido en kg (ej: 1.250)' : 'Ingresá una cantidad válida (≥ 1)');
    return;
  }
  if (!pesable && !Number.isInteger(qty)) {
    showToast(false, 'Este producto solo acepta cantidades enteras.');
    return;
  }

  const nuevaCant = parseFloat(parseFloat((cantActual + delta * qty).toFixed(3)).toFixed(3));
  const leyenda = (document.getElementById('adjustLeyenda')?.value || '').trim();
  cerrarModalAdjust();
  showSpinner('Actualizando stock...');
  try {
    const response = await fetch(SCRIPT_URL, {
      method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'updateStock', id, nuevaCant, leyenda: leyenda || undefined })
    });
    const json = await response.json();
    if (json.success) {
      showToast(true, `Stock: ${cantActual} → ${nuevaCant}${pesable ? ' kg' : ' u.'}${leyenda ? ' · ' + leyenda : ''}`);
      await loadData();
    } else {
      showToast(false, json.message || 'Error al actualizar');
    }
  } catch (e) { showToast(false, 'Error de red al actualizar stock'); console.error(e); }
  hideSpinner();
}

// ══════════════════════════════════════════════════════
//  MODAL ACCIÓN DESDE CONTROL VENCIMIENTOS
// ══════════════════════════════════════════════════════
let _currentAccionVenData = null;

function abrirModalAccionVen(evt, venId, sucursal, cantActual, desc, ean, fechaVenc) {
  if (evt) evt.stopPropagation();

  // ── Restricción horaria (NO aplica para MAYORISTA) ──────────
  if (sucursal !== 'MAYORISTA') {
    const _ahora = new Date();
    const _dia = _ahora.getDay();
    const _mins = _ahora.getHours() * 60 + _ahora.getMinutes();
    const _diasValidos = [1, 2, 3, 4, 5, 6]; // Lun-Sáb
    const _ini = 9 * 60; // 09:00
    const _fin = 16 * 60; // 12:00
    if (!_diasValidos.includes(_dia) || _mins < _ini || _mins >= _fin) {
      showToast(false, 'El registro de acciones está disponible lunes a sábado, de 09:00 a 16:00 hs.', 6000);
      return;
    }
  }
  const venRec = venData.find(v => v.id === venId) || {};
  const pesable = esPesable(venRec);

  _currentAccionVenData = {
    venId, sucursal, cantActual, desc, ean, fechaVenc, pesable,
    gramaje: venRec.gramaje || '',
    codInterno: venRec.codInterno || '',
    sector: venRec.sector || '',
    seccion: venRec.seccion || '',
    proveedor: venRec.proveedor || '',
    lote: venRec.lote || '',
    usuario: venRec.usuario || '',
  };

  // Header
  const descEl = document.getElementById('avDesc');
  if (descEl) descEl.textContent = desc || '—';

  // Pills de info
  const infoEl = document.getElementById('avInfo');
  if (infoEl) {
    const p = (lbl, val, c) => val
      ? `<span><span style="color:var(--text3)">${lbl}:</span> <strong style="color:${c || 'var(--text)'}">${esc(String(val))}</strong></span>`
      : '';
    infoEl.innerHTML = [
      p('Suc', sucursal, 'var(--cyan)'),
      p('Stock', cantActual + (pesable ? ' kg' : ' u.'), '#60a5fa'),
      p('Vence', fmtDateOnly(fechaVenc), 'var(--text2)'),
      p('EAN', ean, 'var(--text3)'),
      venRec.lote ? p('Lote', venRec.lote, 'var(--text3)') : '',
    ].filter(Boolean).join(' &nbsp;·&nbsp; ');
  }

  // Reset campos
  const sel = document.getElementById('avMotivo');
  if (sel) sel.value = '';
  const cantEl = document.getElementById('avCantidad');
  if (cantEl) {
    cantEl.value = cantActual;
    cantEl.max = cantActual;
    cantEl.step = pesable ? '0.001' : '1';
    cantEl.min = pesable ? '0.001' : '1';
  }
  const maxEl = document.getElementById('avCantMax');
  if (maxEl) maxEl.textContent = `Máximo disponible: ${cantActual}${pesable ? ' kg' : ' u.'}`;
  const leyEl = document.getElementById('avLeyenda');
  if (leyEl) leyEl.value = '';
  const nombEl = document.getElementById('avNombre');
  if (nombEl) nombEl.value = '';
  const nombErrEl = document.getElementById('avNombreErr');
  if (nombErrEl) nombErrEl.style.display = 'none';

  document.getElementById('modalAccionVen').style.display = 'flex';
  setTimeout(() => { if (sel) sel.focus(); }, 120);
}

function cerrarModalAccionVen() {
  document.getElementById('modalAccionVen').style.display = 'none';
  _currentAccionVenData = null;
}

async function ejecutarAccionVen() {
  if (!_currentAccionVenData) return;

  const motivo = (document.getElementById('avMotivo')?.value || '').trim();
  if (!motivo) { showToast(false, 'Seleccioná un motivo antes de confirmar.'); return; }

  const { venId, sucursal, cantActual, desc, ean, fechaVenc,
    gramaje, codInterno, sector, seccion, proveedor, lote, pesable } = _currentAccionVenData;

  const raw = (document.getElementById('avCantidad')?.value || '0').replace(',', '.');
  const cantidad = parseFloat(raw);
  if (!cantidad || cantidad <= 0 || isNaN(cantidad)) {
    showToast(false, 'Ingresá una cantidad válida mayor a 0.'); return;
  }
  if (!pesable && !Number.isInteger(cantidad)) {
    showToast(false, 'Solo se aceptan cantidades enteras para este producto.'); return;
  }
  if (cantidad > cantActual) {
    showToast(false, `No podés registrar más de ${cantActual}${pesable ? ' kg' : ' u.'} disponibles.`); return;
  }

  const leyenda = (document.getElementById('avLeyenda')?.value || '').trim();

  const nombre = (document.getElementById('avNombre')?.value || '').trim();
  if (!nombre) {
    const errEl = document.getElementById('avNombreErr');
    if (errEl) errEl.style.display = 'block';
    document.getElementById('avNombre')?.focus();
    return;
  }
  const errElNomb = document.getElementById('avNombreErr');
  if (errElNomb) errElNomb.style.display = 'none';

  // Formatear fechaVenc como YYYY-MM-DD
  const fvParsed = parseFechaVenc(fechaVenc);
  const fechaVencFmt = fvParsed
    ? `${fvParsed.getFullYear()}-${String(fvParsed.getMonth() + 1).padStart(2, '0')}-${String(fvParsed.getDate()).padStart(2, '0')}`
    : (fechaVenc || '');

  cerrarModalAccionVen();
  showSpinner('Registrando acción…');

  try {
    const body = {
      action: 'registrarAccionDesdeVen',
      idVen: venId,
      motivo,
      cantidad,
      leyenda: leyenda || '',
      usuario: nombre.toUpperCase(),
      // Datos del producto (el backend los usa si no los lee de la fila VEN)
      ean,
      fechaVenc: fechaVencFmt,
      sucursal,
      descripcion: desc,
      gramaje,
      codInterno,
      sector,
      seccion,
      proveedor,
      lote,
    };

    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    });
    const json = await res.json();

    if (json.success) {
      const esVentaConsumo = motivo.toUpperCase().includes('VENTA') || motivo.toUpperCase().includes('CONSUMO');
      const destinoLabel = esVentaConsumo ? ' → 📊 Ventas/Consumos' : '';
      showToast(true, `Acción registrada: ${cantidad}${pesable ? ' kg' : ' u.'} · ${motivo}${leyenda ? ' · ' + leyenda : ''}${destinoLabel}`);
      await loadData();
      // Volver al panel de vencimientos para ver el resultado
      showPanel('vencimientos');
    } else {
      showToast(false, json.message || 'Error al registrar acción');
    }
  } catch (e) {
    showToast(false, 'Error de red al registrar acción');
    console.error(e);
  }
  hideSpinner();
}

// ══════════════════════════════════════════════════════
//  MODO LOTE
// ══════════════════════════════════════════════════════
function toggleLoteMode() {
  loteMode = !loteMode;
  const btn = document.getElementById('btnLoteMode');
  btn.classList.toggle('active', loteMode);
  btn.textContent = loteMode ? '✕ Salir de Lote' : '🗂 Modo Lote';
  if (!loteMode) { loteSeleccionados.clear(); loteQueue = []; document.getElementById('loteModeBar').classList.remove('visible'); document.getElementById('loteQueueWrap').classList.remove('visible'); }
  renderVencTable();
}

function toggleLoteRow(sucKey, id, desc, ean, fechaVenc, suc, cantActual, checkbox) {
  if (checkbox.checked) loteSeleccionados.set(sucKey, { id, desc, ean, fechaVenc, suc, cantActual });
  else { loteSeleccionados.delete(sucKey); loteQueue = loteQueue.filter(op => op.sucKey !== sucKey); renderLoteQueue(); }
  updateLoteModeBar();
}

function updateLoteModeBar() {
  const n = loteSeleccionados.size;
  document.getElementById('loteCount').textContent = `${n} sucursal${n !== 1 ? 'es' : ''} seleccionada${n !== 1 ? 's' : ''}`;
  document.getElementById('loteModeBar').classList.toggle('visible', n > 0);
  loteSeleccionados.forEach((item, sucKey) => {
    const yaEsta = loteQueue.some(op => op.sucKey === sucKey);
    if (!yaEsta) loteQueue.push({ sucKey, id: item.id, desc: item.desc, ean: item.ean, fechaVenc: item.fechaVenc, suc: item.suc, cantActual: item.cantActual, op: 'adj_pos', cant: 1, dest: '' });
  });
  renderLoteQueue();
}

function limpiarSeleccionLote() { loteSeleccionados.clear(); loteQueue = []; renderLoteQueue(); document.getElementById('loteModeBar').classList.remove('visible'); renderVencTable(); }
function vaciarColaLote() { loteQueue = []; loteSeleccionados.clear(); renderLoteQueue(); document.getElementById('loteModeBar').classList.remove('visible'); renderVencTable(); }

// ══════════════════════════════════════════════════════
//  RENDER LOTE QUEUE — con soporte pesable
// ══════════════════════════════════════════════════════
function renderLoteQueue() {
  const tbody = document.getElementById('loteQueueBody');
  const wrap = document.getElementById('loteQueueWrap');
  document.getElementById('loteQueueCount').textContent = `${loteQueue.length} operación${loteQueue.length !== 1 ? 'es' : ''}`;
  document.getElementById('btnEjecutarLote').disabled = loteQueue.length === 0;

  if (!loteQueue.length) {
    wrap.classList.remove('visible');
    tbody.innerHTML = '<tr><td colspan="9"><div class="empty" style="padding:20px"><p>Seleccioná filas y configurá operaciones</p></div></td></tr>';
    return;
  }
  wrap.classList.add('visible');

  tbody.innerHTML = loteQueue.map((op, idx) => {
    // ── Detectar pesable para este ítem ──
    const itemLote = venData.find(v => v.id === op.id);
    const opPesable = esPesable(itemLote);
    const minCant = op.op === 'adj_directo' ? 0 : (opPesable ? 0.001 : 1);
    const stepCant = opPesable ? '0.001' : '1';

    const sucursales = ['HIPER', 'CENTRO', 'RIBERA', 'MAYORISTA', 'PROVEEDOR', 'OTRO'].filter(s => s !== op.suc);
    const destOpts = sucursales.map(s => `<option value="${s}" ${op.dest === s ? 'selected' : ''}>${s}</option>`).join('');
    const destCell = op.op === 'transfer'
      ? `<select class="lote-dest-select" onchange="loteQueue[${idx}].dest=this.value"><option value="">— Elegir —</option>${destOpts}</select>`
      : op.op === 'adj_directo'
        ? `<span style="font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--cyan)">→ reemplaza stock</span>`
        : `<span style="color:var(--text3);font-size:10px">—</span>`;

    const qtyVal = op.op === 'adj_directo'
      ? (op.cantDirecta ?? op.cantActual)
      : op.cant;

    // ── onchange adaptado: parseFloat para pesables, parseInt para el resto ──
    const qtyOnChange = op.op === 'adj_directo'
      ? `loteQueue[${idx}].cantDirecta=parseFloat(this.value.replace(',','.'))||0`
      : opPesable
        ? `loteQueue[${idx}].cant=parseFloat(this.value.replace(',','.'))||0.001`
        : `loteQueue[${idx}].cant=parseInt(this.value)||1`;

    const qtyTitle = op.op === 'adj_directo'
      ? `<div style="font-size:9px;color:var(--cyan);font-family:'IBM Plex Mono',monospace;margin-top:2px;">stock final${opPesable ? ' (kg)' : ''}</div>`
      : `<div style="font-size:9px;color:var(--text3);font-family:'IBM Plex Mono',monospace;margin-top:2px;">${opPesable ? 'delta kg' : 'delta'}</div>`;

    const pesBadge = opPesable
      ? `<span style="font-size:9px;color:var(--cyan);font-family:'IBM Plex Mono',monospace;margin-left:4px">⚖</span>`
      : '';

    return `<tr>
      <td class="c-main" style="max-width:180px">${esc(op.desc)}${pesBadge}</td>
      <td class="c-mono" style="font-size:10px">${esc(op.ean)}</td>
      <td class="c-mono" style="font-size:10px">${fmtDateOnly(op.fechaVenc)}</td>
      <td><span class="suc-b ${getSucClass(op.suc)}">${esc(op.suc)}</span></td>
      <td class="c-right c-mono">${op.cantActual}</td>
      <td>
        <select class="lote-op-select" onchange="loteQueue[${idx}].op=this.value;renderLoteQueue()">
          <option value="adj_pos"     ${op.op === 'adj_pos' ? 'selected' : ''}>＋ Ajuste positivo</option>
          <option value="adj_neg"     ${op.op === 'adj_neg' ? 'selected' : ''}>－ Ajuste negativo</option>
          <option value="adj_directo" ${op.op === 'adj_directo' ? 'selected' : ''}>✎ Ajuste directo</option>
          <option value="transfer"    ${op.op === 'transfer' ? 'selected' : ''}>⇄ Transferencia</option>
        </select>
      </td>
      <td><div>
        <input type="number" class="lote-qty-input"
          min="${minCant}" step="${stepCant}" value="${qtyVal}"
          onchange="${qtyOnChange}" oninput="${qtyOnChange}">
        ${qtyTitle}
      </div></td>
      <td>${destCell}</td>
      <td><button class="lote-remove" onclick="quitarDeLote(${idx})" title="Quitar">✕</button></td>
    </tr>`;
  }).join('');
}

function quitarDeLote(idx) {
  const op = loteQueue[idx];
  loteSeleccionados.delete(op.sucKey);
  loteQueue.splice(idx, 1);
  renderLoteQueue(); updateLoteModeBar();
  const cb = document.querySelector(`input[data-suc-key="${op.sucKey}"]`);
  if (cb) cb.checked = false;
}

// ══════════════════════════════════════════════════════
//  EJECUTAR LOTE — con soporte pesable
// ══════════════════════════════════════════════════════
async function ejecutarLote() {
  if (!loteQueue.length) return;

  // ── Validaciones ──
  for (const op of loteQueue) {
    const itemVal = venData.find(v => v.id === op.id);
    const opPes = esPesable(itemVal);
    const minCant = opPes ? 0.001 : 1;

    if (op.op === 'adj_directo') {
      const val = parseFloat(String(op.cantDirecta ?? '').replace(',', '.'));
      if (isNaN(val) || val < 0) {
        showToast(false, `Stock final inválido en: ${op.desc} (${op.suc})`); return;
      }
    } else {
      const val = parseFloat(String(op.cant ?? '').replace(',', '.'));
      if (!val || val < minCant || isNaN(val)) {
        showToast(false, `Cantidad inválida en: ${op.desc} (${op.suc})${opPes ? ' — mínimo 0.001 kg' : ''}`); return;
      }
      if (!opPes && !Number.isInteger(val)) {
        showToast(false, `${op.desc} (${op.suc}) solo acepta cantidades enteras.`); return;
      }
    }

    if (op.op === 'transfer') {
      if (!op.dest) { showToast(false, `Falta destino en: ${op.desc} (${op.suc})`); return; }
      if (op.cantActual <= 0) { showToast(false, `Sin stock en: ${op.desc} (${op.suc})`); return; }
      const cantNum = parseFloat(String(op.cant).replace(',', '.'));
      if (cantNum > op.cantActual) { showToast(false, `Stock insuficiente en: ${op.desc} (${op.suc})`); return; }
    }
  }

  const total = loteQueue.length;
  showSpinner(`Procesando 0 / ${total}…`);
  const resultados = []; let ok = 0, err = 0;

  for (let i = 0; i < loteQueue.length; i++) {
    const op = loteQueue[i];
    showSpinner(`Procesando ${i + 1} / ${total}…`);

    // Parsear cantidades con soporte de coma
    const opCant = parseFloat(String(op.cant ?? 0).replace(',', '.'));
    const opDirecta = parseFloat(String(op.cantDirecta ?? op.cantActual).replace(',', '.'));

    try {
      let payload, nuevaCant;
      const redondear = n => parseFloat(parseFloat(n).toFixed(3));
      if (op.op === 'adj_pos') {
        nuevaCant = redondear(op.cantActual + opCant);
        payload = { action: 'updateStock', id: op.id, nuevaCant };
      } else if (op.op === 'adj_neg') {
        nuevaCant = redondear(op.cantActual - opCant);
        payload = { action: 'updateStock', id: op.id, nuevaCant };
      } else if (op.op === 'adj_directo') {
        nuevaCant = redondear(opDirecta);
        payload = { action: 'updateStock', id: op.id, nuevaCant };
      } else {
        nuevaCant = redondear(op.cantActual - opCant);
        payload = { action: 'transferStock', idOrigen: op.id, sucursalDestino: op.dest, cantidad: redondear(opCant) };
      }

      const res = await fetch(SCRIPT_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(payload) });
      const json = await res.json();
      if (json.success) {
        ok++;
        resultados.push({ ...op, resultado: 'OK', nuevaCant: nuevaCant ?? '—' });
        const item = venData.find(v => v.id === op.id);
        if (item && op.op !== 'transfer') item.cantidad = nuevaCant;
      } else {
        err++;
        resultados.push({ ...op, resultado: 'ERROR: ' + (json.message || '?'), nuevaCant: '—' });
      }
    } catch (e) {
      err++;
      resultados.push({ ...op, resultado: 'ERROR: red', nuevaCant: '—' });
    }
  }

  hideSpinner();
  exportLoteExcel(resultados);
  showToast(err === 0, `${ok} op${ok !== 1 ? 's' : ''} ejecutada${ok !== 1 ? 's' : ''}${err ? ` · ${err} error${err !== 1 ? 'es' : ''}` : ''}`);
  // Limpiar y salir del modo lote completamente
  loteQueue = []; loteSeleccionados.clear(); loteMode = false;
  const btnLote = document.getElementById('btnLoteMode');
  if (btnLote) { btnLote.classList.remove('active'); btnLote.textContent = '🗂 Modo Lote'; }
  document.getElementById('loteModeBar').classList.remove('visible');
  document.getElementById('loteQueueWrap').classList.remove('visible');
  await loadData();
}

function exportLoteExcel(resultados) {
  if (!resultados.length) return;
  const headers = ['Producto', 'EAN', 'Vencimiento', 'Sucursal', 'Stock Anterior', 'Operación', 'Cantidad', 'Destino', 'Nuevo Stock', 'Resultado', 'Fecha'];
  const now = new Date().toLocaleString('es-AR');
  const wsData = [headers, ...resultados.map(op => [
    op.desc, op.ean, fmtDateOnly(op.fechaVenc), op.suc, op.cantActual,
    op.op === 'adj_pos' ? 'Ajuste +' : op.op === 'adj_neg' ? 'Ajuste -' : op.op === 'adj_directo' ? 'Ajuste directo' : 'Transferencia',
    op.cant, op.dest || '—', op.nuevaCant, op.resultado, now
  ])];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{ wch: 32 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 24 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Lote');
  XLSX.writeFile(wb, 'operaciones_lote_' + fmtIso(new Date()) + '.xlsx');
}

function abrirModalLote() {
  if (!loteQueue.length) return;
  document.getElementById('loteQueueWrap').classList.add('visible');
  const modal = document.getElementById('modalConfirmLote');
  const n = loteQueue.length;
  document.getElementById('confirmLoteResumen').textContent = `Estás por ejecutar ${n} operación${n !== 1 ? 'es' : ''} en lote. ¿Confirmar?`;
  modal.style.display = 'flex';
}

function cerrarModalConfirmLote() { document.getElementById('modalConfirmLote').style.display = 'none'; }
function confirmarEjecutarLote() { cerrarModalConfirmLote(); ejecutarLote(); }


// ══════════════════════════════════════════════════════
//  THEME TOGGLE
// ══════════════════════════════════════════════════════
function toggleTheme() {
  const html = document.documentElement;
  const isLight = html.getAttribute('data-theme') === 'light';
  const newTheme = isLight ? 'dark' : 'light';
  html.setAttribute('data-theme', newTheme);
  localStorage.setItem('nexus_theme', newTheme);
  _updateThemeBtn(newTheme);
}

function _updateThemeBtn(theme) {
  const icon = document.getElementById('themeIcon');
  const label = document.getElementById('themeLabel');
  const hint = document.getElementById('themeHint');
  if (!icon) return;
  if (theme === 'light') {
    icon.textContent = '🌙';
    label.textContent = 'Modo Oscuro';
    hint.textContent = 'LIGHT';
  } else {
    icon.textContent = '☀️';
    label.textContent = 'Modo Claro';
    hint.textContent = 'DARK';
  }
}

// Aplicar tema guardado al cargar la página
(function () {
  const saved = localStorage.getItem('nexus_theme') || 'dark';
  document.documentElement.setAttribute('data-theme', saved);
  // El botón se actualiza en window.onload
  window.addEventListener('DOMContentLoaded', () => _updateThemeBtn(saved));
})();