// ══════════════════════════════════════════════════════════════
//  tcd-gestion.js  —  Tablero de Gestión de Proveedores
// ══════════════════════════════════════════════════════════════

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwj0qEOm9THbYxw0TYek2Oot3dlL1wn7YmPLtYknFzrGBQJXFnd-kh7yxXtFgYFyC-B/exec';
const AUTO_REFRESH_MS = 5 * 60 * 1000;

const $ = id => document.getElementById(id);

// ── Estado ──────────────────────────────────────────────────
const state = {
    data: [],             // todos los registros
    proveedores: {},      // agrupados por proveedor
    activeProv: null,     // proveedor seleccionado
    vinculados: {},       // mapa EAN → [registros] para vínculos
    modalRecord: null,    // registro abierto en modal
    modalEstado: null,    // estado seleccionado en modal
    modalCalcNC: 0,       // NC calculada
    saving: false,
    loading: false,
    sidebarFilter: 'todos',
    lastUpdate: null,
};

// ── INIT ────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    // Refresh
    $('btnRefresh').addEventListener('click', () => loadData(true));
    $('btnRetry').addEventListener('click', () => loadData(true));

    // Sidebar toggle
    $('btnToggleSidebar').addEventListener('click', toggleSidebar);
    $('btnBack').addEventListener('click', () => {
        // En mobile, volver al sidebar
        document.getElementById('sidebar').classList.remove('collapsed');
        showKanban(false);
    });

    // Búsqueda proveedor
    $('provSearch').addEventListener('input', renderProvList);

    // Filtros sidebar
    document.querySelectorAll('.sf-pill').forEach(pill => {
        pill.addEventListener('click', () => {
            document.querySelectorAll('.sf-pill').forEach(p => p.classList.remove('sf-pill--active'));
            pill.classList.add('sf-pill--active');
            state.sidebarFilter = pill.dataset.filter;
            renderProvList();
        });
    });

    // Modal
    $('modalClose').addEventListener('click', closeModal);
    $('modalCancel').addEventListener('click', closeModal);
    $('modalOverlay').addEventListener('click', e => { if (e.target === $('modalOverlay')) closeModal(); });
    $('modalSave').addEventListener('click', saveRecord);
    $('btnUseCalc').addEventListener('click', () => {
        $('modalMonto').value = state.modalCalcNC.toFixed(2);
    });

    // Estado selector en modal
    document.querySelectorAll('.est-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.est-btn').forEach(b => b.classList.remove('selected'));
            btn.classList.add('selected');
            state.modalEstado = btn.dataset.estado;
        });
    });

    // ESC cierra modal
    document.addEventListener('keydown', e => {
        if (e.key === 'Escape') closeModal();
    });

    loadData();
    if (AUTO_REFRESH_MS > 0) setInterval(() => loadData(false), AUTO_REFRESH_MS);
});

// ══════════════════════════════════════════════════════════════
//  CARGA DE DATOS
// ══════════════════════════════════════════════════════════════
async function loadData(showSpinner = true) {
    if (state.loading) return;
    state.loading = true;

    if (showSpinner) {
        $('boardLoading').style.display = 'flex';
        $('boardError').style.display = 'none';
        $('boardPlaceholder').style.display = 'none';
        showKanban(false);
    }
    setSyncStatus('loading');
    $('btnRefresh').classList.add('spinning');

    try {
        const url = `${APPS_SCRIPT_URL}?action=getHistorial&t=${Date.now()}`;
        const res = await fetch(url);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const json = await res.json();
        if (!json.success) throw new Error(json.message || 'Error del servidor');

        state.data = json.data || [];
        state.lastUpdate = new Date();

        processData();
        renderProvList();

        $('boardLoading').style.display = 'none';
        $('boardError').style.display = 'none';

        // Restaurar proveedor activo si existía
        if (state.activeProv && state.proveedores[state.activeProv]) {
            renderKanban(state.activeProv);
        } else {
            $('boardPlaceholder').style.display = 'flex';
            showKanban(false);
        }

        setSyncStatus('ok');
        showToast('ok', `${state.data.length} registros cargados`);

    } catch (err) {
        console.error('[Gestión]', err);
        $('boardLoading').style.display = 'none';
        $('boardError').style.display = 'flex';
        $('errorMsg').textContent = err.message;
        setSyncStatus('err');
        showToast('err', 'Error al cargar datos');
    } finally {
        state.loading = false;
        $('btnRefresh').classList.remove('spinning');
    }
}

// ══════════════════════════════════════════════════════════════
//  PROCESAMIENTO
// ══════════════════════════════════════════════════════════════
function processData() {
    // Agrupar por proveedor
    state.proveedores = {};
    state.data.forEach(r => {
        const prov = (r.PROVEEDOR || 'Sin proveedor').trim();
        if (!state.proveedores[prov]) {
            state.proveedores[prov] = {
                nombre: prov,
                cod: r['COD. PROVEEDOR'] || '',
                registros: [],
            };
        }
        state.proveedores[prov].registros.push(r);
    });

    // Detectar vínculos: mismo EAN + mismo proveedor → tiene acción Y vencimiento/decomiso
    // Un "vínculo" es cuando el mismo producto (EAN) de un proveedor tiene
    // un registro de ACCION y también uno de VENCIDO — puede implicar reclamo doble si no se maneja bien
    state.vinculados = {};
    Object.values(state.proveedores).forEach(prov => {
        const byEAN = {};
        prov.registros.forEach(r => {
            const ean = String(r.EAN || r['COD. INTERNO'] || '').trim();
            if (!ean) return;
            const key = `${prov.nombre}||${ean}`;
            if (!byEAN[key]) byEAN[key] = [];
            byEAN[key].push(r);
        });
        // Guardar solo los grupos que tienen más de 1 tipo de motivo (acción + vencido, etc.)
        Object.entries(byEAN).forEach(([key, registros]) => {
            if (registros.length < 2) return;
            const motivos = new Set(registros.map(r => motivoGrupo(r.MOTIVO)));
            // Vínculo relevante: tiene acción Y (vencido O decomiso)
            if (motivos.has('Acción') && (motivos.has('Vencido') || motivos.has('Decomiso'))) {
                registros.forEach(r => {
                    const id = r.ID;
                    if (!state.vinculados[id]) state.vinculados[id] = [];
                    // Agregar los otros registros del mismo grupo como vínculos
                    registros.filter(rr => rr.ID !== id).forEach(rr => {
                        if (!state.vinculados[id].find(v => v.ID === rr.ID)) {
                            state.vinculados[id].push(rr);
                        }
                    });
                });
            }
        });
    });
}

// ══════════════════════════════════════════════════════════════
//  SIDEBAR — LISTA DE PROVEEDORES
// ══════════════════════════════════════════════════════════════
function renderProvList() {
    const query  = $('provSearch').value.toLowerCase().trim();
    const filter = state.sidebarFilter;

    let provs = Object.values(state.proveedores);

    // Filtro de texto
    if (query) provs = provs.filter(p => p.nombre.toLowerCase().includes(query));

    // Filtro por tipo
    if (filter === 'pendientes') {
        provs = provs.filter(p => p.registros.some(r => normEstado(r.ESTADO) === 'PENDIENTE'));
    }
    if (filter === 'vinculados') {
        provs = provs.filter(p => p.registros.some(r => state.vinculados[r.ID]?.length > 0));
    }

    // Ordenar: pendientes primero, luego alfabético
    provs.sort((a, b) => {
        const pa = a.registros.filter(r => normEstado(r.ESTADO) === 'PENDIENTE').length;
        const pb = b.registros.filter(r => normEstado(r.ESTADO) === 'PENDIENTE').length;
        return pb - pa || a.nombre.localeCompare(b.nombre);
    });

    const list = $('provList');
    list.innerHTML = '';

    if (provs.length === 0) {
        list.innerHTML = `<li style="padding:20px 12px;font-size:12px;color:var(--text3);text-align:center">Sin resultados</li>`;
        return;
    }

    provs.forEach((p, i) => {
        const pendientes = p.registros.filter(r => normEstado(r.ESTADO) === 'PENDIENTE').length;
        const enGestion  = p.registros.filter(r => normEstado(r.ESTADO) === 'EN GESTION').length;
        const nc         = p.registros.filter(r => normEstado(r.ESTADO) === 'N/C RECIBIDA').length;
        const tieneVink  = p.registros.some(r => state.vinculados[r.ID]?.length > 0);

        const li = document.createElement('li');
        li.className = 'prov-item' + (state.activeProv === p.nombre ? ' active' : '');
        li.style.animationDelay = `${i * 0.03}s`;
        li.innerHTML = `
            <span class="prov-item-name">${esc(p.nombre)}</span>
            <div class="prov-item-meta">
                ${p.cod ? `<span class="prov-item-cod">COD ${p.cod}</span>` : ''}
                <div class="prov-item-counts">
                    ${pendientes ? `<span class="pic-badge pic-badge--p">${pendientes}P</span>` : ''}
                    ${enGestion  ? `<span class="pic-badge pic-badge--g">${enGestion}G</span>` : ''}
                    ${nc         ? `<span class="pic-badge pic-badge--n">${nc}N/C</span>` : ''}
                    ${tieneVink  ? `<span class="pic-badge pic-badge--vink">⚠ VIN</span>` : ''}
                </div>
            </div>
        `;
        li.addEventListener('click', () => selectProveedor(p.nombre));
        list.appendChild(li);
    });
}

function selectProveedor(nombre) {
    state.activeProv = nombre;
    // Actualizar selección visual
    document.querySelectorAll('.prov-item').forEach(li => li.classList.remove('active'));
    document.querySelectorAll('.prov-item').forEach(li => {
        if (li.querySelector('.prov-item-name')?.textContent.trim() === nombre) {
            li.classList.add('active');
        }
    });
    renderKanban(nombre);
    $('boardPlaceholder').style.display = 'none';
    // En mobile cerrar sidebar
    if (window.innerWidth <= 700) {
        document.getElementById('sidebar').classList.add('collapsed');
    }
}

function toggleSidebar() {
    document.getElementById('sidebar').classList.toggle('collapsed');
}

// ══════════════════════════════════════════════════════════════
//  KANBAN
// ══════════════════════════════════════════════════════════════
function renderKanban(provNombre) {
    showKanban(true);

    const prov = state.proveedores[provNombre];
    if (!prov) return;

    const registros = prov.registros;

    // Header
    $('provHeaderName').textContent = prov.nombre;
    $('provHeaderMeta').textContent = `${prov.cod ? `COD ${prov.cod} · ` : ''}${registros.length} registros · ${new Set(registros.map(r => r.SUCURSAL)).size} sucursales`;

    // KPIs del header
    const montoNC     = registros.reduce((s, r) => s + (parseFloat(r['MONTO N/C']) || 0), 0);
    const costoTotal  = registros.reduce((s, r) => s + calcCostoTotal(r), 0);
    const pendientesN = registros.filter(r => normEstado(r.ESTADO) === 'PENDIENTE').length;

    $('provHeaderKpis').innerHTML = `
        <div class="phkpi">
            <span class="phkpi-val" style="color:var(--yellow)">${pendientesN}</span>
            <span class="phkpi-lbl">Pendientes</span>
        </div>
        <div class="phkpi">
            <span class="phkpi-val" style="color:var(--accent)">${fmtPeso(costoTotal)}</span>
            <span class="phkpi-lbl">Costo estimado</span>
        </div>
        <div class="phkpi">
            <span class="phkpi-val" style="color:var(--green)">${fmtPeso(montoNC)}</span>
            <span class="phkpi-lbl">N/C recibida</span>
        </div>
    `;

    // Alerta de vínculos
    const vinksCount = registros.filter(r => state.vinculados[r.ID]?.length > 0).length;
    if (vinksCount > 0) {
        $('vinkAlert').style.display = 'flex';
        $('vinkAlertSub').textContent = `${vinksCount} registros tienen vínculos acción+vencimiento. Revisalos antes de reclamar para no duplicar el monto de la N/C.`;
    } else {
        $('vinkAlert').style.display = 'none';
    }

    // Columnas
    const cols = ['PENDIENTE', 'EN GESTION', 'N/C RECIBIDA', 'RECHAZADA'];
    const colIds = { 'PENDIENTE': 'PENDIENTE', 'EN GESTION': 'EN_GESTION', 'N/C RECIBIDA': 'NC_RECIBIDA', 'RECHAZADA': 'RECHAZADA' };

    cols.forEach(colEstado => {
        const colId  = colIds[colEstado];
        const colEl  = $(`col-${colId}`);
        const cntEl  = $(`cnt-${colId}`);
        const amtEl  = $(`amt-${colId}`);
        colEl.innerHTML = '';

        const colRegs = registros.filter(r => normEstado(r.ESTADO) === colEstado);
        cntEl.textContent = colRegs.length;

        // Monto de la columna
        const colMonto = colRegs.reduce((s, r) => {
            if (colEstado === 'N/C RECIBIDA') return s + (parseFloat(r['MONTO N/C']) || 0);
            return s + calcNCsugerida(r).total;
        }, 0);
        amtEl.textContent = colMonto > 0 ? fmtPeso(colMonto) : '';

        colRegs.forEach((r, i) => {
            const card = buildCard(r, i);
            colEl.appendChild(card);
        });

        // Columna vacía
        if (colRegs.length === 0) {
            const empty = document.createElement('div');
            empty.style.cssText = 'padding:16px 8px;text-align:center;font-size:11px;color:var(--text3)';
            empty.textContent = 'Sin registros';
            colEl.appendChild(empty);
        }
    });
}

function buildCard(r, idx) {
    const card = document.createElement('div');
    card.className = 'kcard';
    card.style.animationDelay = `${idx * 0.05}s`;
    card.dataset.id = r.ID;

    const vinculados = state.vinculados[r.ID] || [];
    const isLinked   = vinculados.length > 0;
    if (isLinked) card.classList.add('is-linked');

    const nc    = calcNCsugerida(r);
    const ncStr = r['MONTO N/C'] ? fmtPeso(parseFloat(r['MONTO N/C'])) : '';
    const cantidad = fmtCantidad(r);

    card.innerHTML = `
        ${isLinked ? `<div class="kcard-vink-tag">🔗 VINCULADO</div>` : ''}
        <div class="kcard-top" style="${isLinked ? 'margin-top:10px' : ''}">
            <div class="kcard-desc" title="${esc(r.DESCRIPCION || '')}">${esc(r.DESCRIPCION || '—')}</div>
            <span class="kcard-motivo ${motivoClass(r.MOTIVO)}">${esc(motivoCorto(r.MOTIVO))}</span>
        </div>
        <div class="kcard-grid">
            <div>
                <div class="kcard-field-lbl">Cantidad</div>
                <div class="kcard-field-val accent">${cantidad}</div>
            </div>
            <div>
                <div class="kcard-field-lbl">Vencimiento</div>
                <div class="kcard-field-val">${esc(r['FECHA VENC.'] || r['FECHA VENCIMIENTO'] || '—')}</div>
            </div>
            <div>
                <div class="kcard-field-lbl">Costo neto</div>
                <div class="kcard-field-val">${r['COSTO NETO'] ? fmtPeso(parseFloat(r['COSTO NETO'])) : '—'}</div>
            </div>
            <div>
                <div class="kcard-field-lbl">N/C estimada</div>
                <div class="kcard-field-val accent">${fmtPeso(nc.total)}</div>
            </div>
        </div>
        ${isLinked ? `
        <div class="kcard-linked-records">
            <div class="klr-title">🔗 Registros vinculados</div>
            ${vinculados.map(v => `
                <div class="klr-item">${esc(v.ID)} · ${motivoCorto(v.MOTIVO)} · ${fmtCantidad(v)}</div>
            `).join('')}
        </div>` : ''}
        <div class="kcard-footer">
            <span class="kcard-id">${esc(r.ID || '')}</span>
            ${ncStr ? `<span class="kcard-nc">✓ ${ncStr}</span>` : ''}
            <span class="kcard-sucursal">${esc(r.SUCURSAL || '')}</span>
        </div>
    `;

    card.addEventListener('click', () => openModal(r));
    return card;
}

function showKanban(show) {
    $('kanbanWrapper').style.display = show ? 'flex' : 'none';
}

// ══════════════════════════════════════════════════════════════
//  MODAL DE EDICIÓN
// ══════════════════════════════════════════════════════════════
function openModal(r) {
    state.modalRecord = r;
    state.modalEstado = normEstado(r.ESTADO);

    $('modalTitle').textContent = r.DESCRIPCION || 'Editar registro';

    // Info del producto
    const costoNeto = parseFloat(r['COSTO NETO']) || 0;
    const iva       = parseFloat(r['IVA %']) || 0;
    const costoFinal = costoNeto * (1 + iva / 100);

    $('modalProductInfo').innerHTML = `
        <div class="mpi-full">
            <div class="mpi-field-lbl">Descripción</div>
            <div class="mpi-field-val">${esc(r.DESCRIPCION || '—')}</div>
        </div>
        <div>
            <div class="mpi-field-lbl">Proveedor</div>
            <div class="mpi-field-val">${esc(r.PROVEEDOR || '—')}</div>
        </div>
        <div>
            <div class="mpi-field-lbl">Motivo</div>
            <div class="mpi-field-val">${esc(r.MOTIVO || '—')}</div>
        </div>
        <div>
            <div class="mpi-field-lbl">Cantidad</div>
            <div class="mpi-field-val mono">${fmtCantidad(r)}</div>
        </div>
        <div>
            <div class="mpi-field-lbl">Vencimiento</div>
            <div class="mpi-field-val mono">${esc(r['FECHA VENC.'] || r['FECHA VENCIMIENTO'] || '—')}</div>
        </div>
        <div>
            <div class="mpi-field-lbl">Costo neto</div>
            <div class="mpi-field-val mono">${fmtPeso(costoNeto)}</div>
        </div>
        <div>
            <div class="mpi-field-lbl">Costo c/IVA ${iva}%</div>
            <div class="mpi-field-val mono">${fmtPeso(costoFinal)}</div>
        </div>
        <div>
            <div class="mpi-field-lbl">Sucursal · ID</div>
            <div class="mpi-field-val mono">${esc(r.SUCURSAL || '')} · ${esc(r.ID || '')}</div>
        </div>
    `;

    // Calculadora NC
    buildNCCalc(r);

    // Vínculos
    const vincs = state.vinculados[r.ID] || [];
    if (vincs.length > 0) {
        $('modalVinkWarn').style.display = 'flex';
        $('modalVinkList').innerHTML = vincs.map(v => {
            const g = motivoGrupo(v.MOTIVO);
            const bgMap = { 'Acción': '#a855f7', 'Vencido': '#f97316', 'Decomiso': '#ef4444', 'Otro': '#818cf8' };
            const color = bgMap[g] || '#818cf8';
            return `
            <div class="mvw-linked-item">
                <span class="mvw-li-badge" style="background:${color}22;color:${color}">${esc(motivoCorto(v.MOTIVO))}</span>
                <div class="mvw-li-info">
                    <div class="mvw-li-desc">${esc(v.DESCRIPCION || '—')}</div>
                    <div class="mvw-li-id">${esc(v.ID)} · ${fmtCantidad(v)} · Venc: ${esc(v['FECHA VENC.'] || v['FECHA VENCIMIENTO'] || '')}</div>
                </div>
                <span style="font-family:var(--mono);font-size:10.5px;color:var(--accent)">${fmtPeso(calcNCsugerida(v).total)}</span>
            </div>`;
        }).join('');
    } else {
        $('modalVinkWarn').style.display = 'none';
    }

    // Estado
    document.querySelectorAll('.est-btn').forEach(b => {
        b.classList.toggle('selected', b.dataset.estado === state.modalEstado);
    });

    // Monto actual
    $('modalMonto').value = r['MONTO N/C'] || '';

    // Observación actual
    $('modalObservacion').value = r['OBSERVACION N/C'] || '';

    // Abrir
    $('modalOverlay').classList.add('open');
    document.body.style.overflow = 'hidden';
}

function buildNCCalc(r) {
    const nc    = calcNCsugerida(r);
    const rows  = nc.detalles;
    state.modalCalcNC = nc.total;

    $('ncCalc').innerHTML = rows.map(row => `
        <div class="nc-calc-row">
            <span class="ncr-label">${esc(row.label)}</span>
            <span class="ncr-formula">${esc(row.formula)}</span>
            <span class="ncr-value ${row.isTotal ? 'total' : row.isInfo ? 'info' : ''}">${esc(row.value)}</span>
        </div>
    `).join('') + (nc.total > 0 ? `
        <div class="nc-calc-row" style="background:rgba(34,197,94,0.05)">
            <span class="ncr-label" style="font-weight:700;color:var(--text)">Total N/C sugerida</span>
            <span class="ncr-formula"></span>
            <span class="ncr-value total">${fmtPeso(nc.total)}</span>
        </div>
    ` : '');
}

function closeModal() {
    $('modalOverlay').classList.remove('open');
    document.body.style.overflow = '';
    state.modalRecord = null;
}

// ══════════════════════════════════════════════════════════════
//  GUARDAR CAMBIOS → Apps Script
// ══════════════════════════════════════════════════════════════
async function saveRecord() {
    if (state.saving || !state.modalRecord) return;
    const r       = state.modalRecord;
    const nuevoEstado  = state.modalEstado;
    const montoNC      = parseFloat($('modalMonto').value) || 0;
    const observacion  = $('modalObservacion').value.trim();

    state.saving = true;
    const btn = $('modalSave');
    btn.classList.add('saving');
    btn.disabled = true;

    try {
        const payload = {
            action: 'updateRecord',
            id: r.ID,
            estado: nuevoEstado,
            montoNC: montoNC || '',
            observacionNC: observacion,
        };

        const res = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify(payload),
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        if (!data.success) throw new Error(data.message || 'Error del servidor');

        // Actualizar en local (sin recargar todo)
        const idx = state.data.findIndex(d => d.ID === r.ID);
        if (idx !== -1) {
            state.data[idx].ESTADO           = nuevoEstado;
            state.data[idx]['MONTO N/C']     = montoNC || '';
            state.data[idx]['OBSERVACION N/C'] = observacion;
        }

        processData();
        renderProvList();
        renderKanban(state.activeProv);

        closeModal();
        showToast('ok', 'Registro actualizado correctamente');

    } catch (err) {
        console.error('[Guardar]', err);
        showToast('err', `Error al guardar: ${err.message}`);
    } finally {
        state.saving = false;
        btn.classList.remove('saving');
        btn.disabled = false;
    }
}

// ══════════════════════════════════════════════════════════════
//  CALCULADORA DE N/C
// ══════════════════════════════════════════════════════════════
function calcNCsugerida(r) {
    const costoNeto = parseFloat(r['COSTO NETO']) || 0;
    const iva       = parseFloat(r['IVA %']) || 0;
    const costoFinal = costoNeto * (1 + iva / 100);
    const cant      = parseFloat(r.CANTIDAD) || 0;
    const motivo    = (r.MOTIVO || '').toUpperCase();
    const grupo     = motivoGrupo(r.MOTIVO);
    const detalles  = [];
    let total = 0;

    detalles.push({
        label: 'Costo neto unitario',
        formula: `${fmtPeso(costoNeto)} × (1 + ${iva}% IVA)`,
        value: fmtPeso(costoFinal),
    });
    detalles.push({
        label: `Cantidad${r['UNIDAD CANTIDAD'] === 'kg' ? ' (kg)' : ''}`,
        formula: '',
        value: fmtCantidad(r),
    });

    if (grupo === 'Acción') {
        // Acción 2x1: empresa reconoce 50% por cada unidad accionada
        // Acción 50% OFF: empresa reconoce 50%
        // Otro descuento: extraer % del string del motivo
        let pct = 50;
        const matchPct = motivo.match(/(\d+)%/);
        if (matchPct) pct = parseFloat(matchPct[1]);
        if (motivo.includes('2X1') || motivo.includes('2 X 1')) pct = 50;

        const nc = costoFinal * cant * (pct / 100);
        total = nc;
        detalles.push({
            label: `Descuento ${pct}% por acción`,
            formula: `${fmtPeso(costoFinal)} × ${cant} × ${pct}%`,
            value: fmtPeso(nc),
        });
        detalles.push({
            label: '⚠ Verificar vencimiento de la acción',
            formula: '',
            value: 'Ver vínculos',
            isInfo: true,
        });

    } else if (grupo === 'Vencido') {
        // Vencimiento sin acción: 100% del costo final
        const nc = costoFinal * cant;
        total = nc;
        detalles.push({
            label: 'Vencimiento (100% costo c/IVA)',
            formula: `${fmtPeso(costoFinal)} × ${cant}`,
            value: fmtPeso(nc),
        });

    } else if (grupo === 'Decomiso') {
        // Decomiso: 100% del costo final (consultar acuerdo con proveedor)
        const nc = costoFinal * cant;
        total = nc;
        detalles.push({
            label: 'Decomiso (100% costo c/IVA)',
            formula: `${fmtPeso(costoFinal)} × ${cant}`,
            value: fmtPeso(nc),
        });
        detalles.push({
            label: 'Verificar acuerdo con proveedor',
            formula: '',
            value: 'Puede variar',
            isInfo: true,
        });

    } else {
        detalles.push({
            label: 'Motivo no estándar — calcular manualmente',
            formula: '',
            value: '—',
            isInfo: true,
        });
    }

    return { total, detalles };
}

function calcCostoTotal(r) {
    const costoNeto  = parseFloat(r['COSTO NETO']) || 0;
    const iva        = parseFloat(r['IVA %']) || 0;
    const costoFinal = costoNeto * (1 + iva / 100);
    const cant       = parseFloat(r.CANTIDAD) || 0;
    return costoFinal * cant;
}

// ══════════════════════════════════════════════════════════════
//  HELPERS UI
// ══════════════════════════════════════════════════════════════
function setSyncStatus(type) {
    const badge = $('syncBadge');
    const text  = $('syncText');
    badge.className = 'sync-badge';
    if (type === 'ok') {
        badge.classList.add('ok');
        const t = state.lastUpdate;
        text.textContent = t ? `Actualizado ${t.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' })}` : 'Actualizado';
    } else if (type === 'err') {
        badge.classList.add('err');
        text.textContent = 'Error de conexión';
    } else {
        text.textContent = 'Cargando...';
    }
}

let toastTimer;
function showToast(type, msg, ms = 3500) {
    const t = $('toast');
    $('toastMsg').textContent = msg;
    t.className = `toast ${type} show`;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), ms);
}

// ══════════════════════════════════════════════════════════════
//  HELPERS LÓGICA
// ══════════════════════════════════════════════════════════════
function motivoGrupo(motivo) {
    if (!motivo) return 'Otro';
    const m = motivo.toUpperCase();
    if (m.includes('ACCION') || m.includes('ACCIÓN') || m.includes('OFF') || m.includes('2X1')) return 'Acción';
    if (m.includes('VENCIDO') || m.includes('VENCIMIENTO')) return 'Vencido';
    if (m.includes('DECOMISO') || m.includes('ROTO') || m.includes('MAL ESTADO')) return 'Decomiso';
    return 'Otro';
}

function motivoCorto(motivo) {
    if (!motivo) return '—';
    const map = {
        'ACCION 2X1': '2×1', 'VENCIDO ACCION 2X1': 'Vto. 2×1',
        'ACCION 50% OFF': '50% OFF', 'VENCIDO ACCION 50% OFF': 'Vto. 50%',
        'VENCIDO': 'Vencido', 'ROTO/DAÑADO': 'Roto',
        'MAL ESTADO': 'Mal estado',
        'DECOMISO FRUTA Y VERDURA': 'Dec. F&V',
        'DECOMISO CARNICERIA': 'Dec. Carn.',
        'DECOMISO FIAMBRERIA': 'Dec. Fiam.',
        'DECOMISO': 'Decomiso',
        'OTRO': 'Otro',
    };
    return map[motivo.toUpperCase()] || motivo.slice(0, 14);
}

function motivoClass(motivo) {
    const g = motivoGrupo(motivo);
    return { 'Acción': 'mot-accion', 'Vencido': 'mot-vencido', 'Decomiso': 'mot-decomiso', 'Otro': 'mot-otro' }[g];
}

function normEstado(estado) {
    if (!estado || estado.trim() === '') return 'PENDIENTE';
    return estado.trim().toUpperCase();
}

function fmtPeso(n) {
    if (!n || isNaN(n)) return '$0';
    return '$' + n.toLocaleString('es-AR', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function fmtCantidad(r) {
    const cant = r.CANTIDAD;
    if (cant === '' || cant == null) return '—';
    const n = parseFloat(cant);
    if (isNaN(n)) return String(cant);
    const esKg = String(r['UNIDAD CANTIDAD'] || '').toLowerCase() === 'kg'
              || (r.GRAMAJE && String(r.GRAMAJE).includes('gramos') && n < 20 && !Number.isInteger(n));
    return esKg ? `${n.toFixed(3)} kg` : `${n} u.`;
}

function esc(str) {
    return String(str)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;')
        .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}