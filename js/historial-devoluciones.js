// ══════════════════════════════════════════════════════════════
//  tcd-gestion.js v2 — Tablero de Gestión de Proveedores
//  - Fix vinculación: VENCIDO se chequea ANTES que ACCION
//  - Drag & Drop entre columnas (HTML5 nativo)
//  - Cards compactas tipo lista
//  - Dashboard de resumen
// ══════════════════════════════════════════════════════════════

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwj0qEOm9THbYxw0TYek2Oot3dlL1wn7YmPLtYknFzrGBQJXFnd-kh7yxXtFgYFyC-B/exec';
const AUTO_REFRESH_MS = 5 * 60 * 1000;

const $ = id => document.getElementById(id);

const state = {
    data: [],
    proveedores: {},
    activeProv: null,
    vinculados: {},
    modalRecord: null,
    modalEstado: null,
    modalCalcNC: 0,
    saving: false,
    loading: false,
    sidebarFilter: 'todos',
    lastUpdate: null,
    activeView: 'kanban',
    drag: { id: null, srcEstado: null, record: null },
};

// ══════════════════════════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
    $('btnRefresh').addEventListener('click', () => loadData(true));
    $('btnRetry').addEventListener('click',  () => loadData(true));
    $('btnToggleSidebar').addEventListener('click', toggleSidebar);
    $('btnBack').addEventListener('click', () => {
        document.getElementById('sidebar').classList.remove('collapsed');
        showKanban(false);
    });
    $('provSearch').addEventListener('input', renderProvList);

    document.querySelectorAll('.sf-pill').forEach(pill => {
        pill.addEventListener('click', () => {
            document.querySelectorAll('.sf-pill').forEach(p => p.classList.remove('sf-pill--active'));
            pill.classList.add('sf-pill--active');
            state.sidebarFilter = pill.dataset.filter;
            renderProvList();
        });
    });

    document.querySelectorAll('.view-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            document.querySelectorAll('.view-tab').forEach(t => t.classList.remove('view-tab--active'));
            tab.classList.add('view-tab--active');
            setView(tab.dataset.view);
        });
    });

    $('modalClose').addEventListener('click',  closeModal);
    $('modalCancel').addEventListener('click', closeModal);
    $('modalOverlay').addEventListener('click', e => { if (e.target === $('modalOverlay')) closeModal(); });
    $('modalSave').addEventListener('click',   saveRecord);
    $('btnUseCalc').addEventListener('click', () => { $('modalMonto').value = state.modalCalcNC.toFixed(2); });

    document.querySelectorAll('.est-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.est-btn').forEach(b => b.classList.remove('selected'));
            btn.classList.add('selected');
            state.modalEstado = btn.dataset.estado;
        });
    });

    document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

    loadData();
    if (AUTO_REFRESH_MS > 0) setInterval(() => loadData(false), AUTO_REFRESH_MS);
});

function setView(v) {
    state.activeView = v;
    $('viewDashboard').style.display = v === 'dashboard' ? 'flex' : 'none';
    $('viewKanban').style.display    = v === 'kanban'    ? 'flex' : 'none';
    if (v === 'dashboard' && state.data.length) renderDashboard();
}

// ══════════════════════════════════════════════════════════════
//  CARGA DE DATOS
// ══════════════════════════════════════════════════════════════
async function loadData(showSpinner = true) {
    if (state.loading) return;
    state.loading = true;

    if (showSpinner) {
        $('boardLoading').style.display = 'flex';
        $('boardError').style.display   = 'none';
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
        $('boardError').style.display   = 'none';

        if (state.activeProv && state.proveedores[state.activeProv]) {
            renderKanban(state.activeProv);
        } else {
            $('boardPlaceholder').style.display = 'flex';
            showKanban(false);
        }

        if (state.activeView === 'dashboard') renderDashboard();
        setSyncStatus('ok');
        showToast('ok', `${state.data.length} registros cargados`);

    } catch (err) {
        console.error('[Gestión]', err);
        $('boardLoading').style.display = 'none';
        $('boardError').style.display   = 'flex';
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
    state.proveedores = {};
    state.data.forEach(r => {
        const prov = (r.PROVEEDOR || 'Sin proveedor').trim();
        if (!state.proveedores[prov]) {
            state.proveedores[prov] = { nombre: prov, cod: r['COD. PROVEEDOR'] || '', registros: [] };
        }
        state.proveedores[prov].registros.push(r);
    });

    // ── Detección de vínculos (FIX: normalizar código, ignorar ceros) ──
    state.vinculados = {};
    Object.values(state.proveedores).forEach(prov => {
        const byCode = {};
        prov.registros.forEach(r => {
            const rawInt = String(r['COD. INTERNO'] || '').trim().replace(/\.0+$/, '');
            const rawEAN = String(r.EAN || '').trim().replace(/\.0+$/, '');
            const codigo = (rawInt && rawInt !== '0') ? rawInt
                         : (rawEAN && rawEAN !== '0') ? rawEAN
                         : null;
            if (!codigo) return;
            const key = `${prov.nombre}||${codigo}`;
            if (!byCode[key]) byCode[key] = [];
            byCode[key].push(r);
        });

        Object.values(byCode).forEach(regs => {
            if (regs.length < 2) return;
            const motivos = new Set(regs.map(r => motivoGrupo(r.MOTIVO)));
            // Necesita tener al menos un 'Acción' Y un 'Vencido' para ser vínculo relevante
            if (motivos.has('Acción') && motivos.has('Vencido')) {
                regs.forEach(r => {
                    if (!state.vinculados[r.ID]) state.vinculados[r.ID] = [];
                    regs.filter(rr => rr.ID !== r.ID).forEach(rr => {
                        if (!state.vinculados[r.ID].find(v => v.ID === rr.ID)) {
                            state.vinculados[r.ID].push(rr);
                        }
                    });
                });
            }
        });
    });
}

// ══════════════════════════════════════════════════════════════
//  DASHBOARD
// ══════════════════════════════════════════════════════════════
function renderDashboard() {
    const all = state.data;
    if (!all.length) return;

    const pendientes = all.filter(r => normEstado(r.ESTADO) === 'PENDIENTE').length;
    const enGestion  = all.filter(r => normEstado(r.ESTADO) === 'EN GESTION').length;
    const ncRec      = all.filter(r => normEstado(r.ESTADO) === 'N/C RECIBIDA').length;
    const rechazada  = all.filter(r => normEstado(r.ESTADO) === 'RECHAZADA').length;
    const montoNC    = all.reduce((s, r) => s + (parseFloat(r['MONTO N/C']) || 0), 0);
    const costoEst   = all.reduce((s, r) => s + calcNCsugerida(r).total, 0);
    const vincsN     = Object.keys(state.vinculados).length;

    $('dashKpis').innerHTML = [
        { v: all.length, l: 'Total registros',    c: 'var(--text)' },
        { v: pendientes, l: 'Pendientes',          c: 'var(--yellow)' },
        { v: enGestion,  l: 'En gestión',          c: 'var(--blue)' },
        { v: ncRec,      l: 'N/C recibidas',       c: 'var(--green)' },
        { v: fmtPeso(costoEst), l: 'Costo estimado',      c: 'var(--accent)' },
        { v: fmtPeso(montoNC),  l: 'N/C total recibida',  c: 'var(--green)' },
        { v: vincsN,            l: '⚠ Vínculos detectados', c: 'var(--purple)' },
    ].map((k, i) => `
        <div class="dash-kpi" style="animation-delay:${i * .05}s">
            <div class="dash-kpi-val" style="color:${k.c}">${k.v}</div>
            <div class="dash-kpi-lbl">${k.l}</div>
        </div>`).join('');

    // Barras estado
    const maxE = Math.max(pendientes, enGestion, ncRec, rechazada, 1);
    $('dashEstados').innerHTML = [
        { l: 'Pendiente',    n: pendientes, c: '#eab308' },
        { l: 'En gestión',   n: enGestion,  c: '#3b82f6' },
        { l: 'N/C Recibida', n: ncRec,      c: '#22c55e' },
        { l: 'Rechazada',    n: rechazada,  c: '#ef4444' },
    ].map(e => `
        <div class="dash-bar-row">
            <span class="dash-bar-lbl">${e.l}</span>
            <div class="dash-bar-track"><div class="dash-bar-fill" style="width:${(e.n/maxE*100).toFixed(1)}%;background:${e.c}"></div></div>
            <span class="dash-bar-val">${e.n}</span>
        </div>`).join('');

    // Barras motivo
    const mc = {};
    all.forEach(r => { const g = motivoGrupo(r.MOTIVO); mc[g] = (mc[g] || 0) + 1; });
    const mcColors = { 'Acción':'#a855f7','Vencido':'#f97316','Decomiso':'#ef4444','Otro':'#818cf8' };
    const maxM = Math.max(...Object.values(mc), 1);
    $('dashMotivos').innerHTML = Object.entries(mc).sort(([,a],[,b]) => b - a).map(([l, n]) => `
        <div class="dash-bar-row">
            <span class="dash-bar-lbl">${l}</span>
            <div class="dash-bar-track"><div class="dash-bar-fill" style="width:${(n/maxM*100).toFixed(1)}%;background:${mcColors[l]||'#818cf8'}"></div></div>
            <span class="dash-bar-val">${n}</span>
        </div>`).join('');

    // Top proveedores
    const provData = Object.values(state.proveedores)
        .map(p => ({
            nombre: p.nombre,
            n: p.registros.length,
            est: p.registros.reduce((s,r) => s + calcNCsugerida(r).total, 0),
            pend: p.registros.filter(r => normEstado(r.ESTADO) === 'PENDIENTE').length,
        }))
        .sort((a,b) => b.est - a.est).slice(0, 10);

    $('dashProvList').innerHTML = provData.map((p, i) => `
        <div class="dash-prov-row">
            <span class="dash-prov-rank">#${i+1}</span>
            <span class="dash-prov-name">${esc(p.nombre)}</span>
            <div class="dash-prov-badges">
                ${p.pend ? `<span class="pic-badge pic-badge--p">${p.pend}P</span>` : ''}
                <span class="pic-badge" style="background:var(--surface3);color:var(--text3)">${p.n} reg</span>
            </div>
            <span class="dash-prov-amount">${fmtPeso(p.est)}</span>
        </div>`).join('') || '<p style="font-size:12px;color:var(--text3)">Sin datos</p>';

    // Próximos a vencer (pendientes)
    const hoy = new Date(); hoy.setHours(0,0,0,0);
    const upcoming = all
        .filter(r => normEstado(r.ESTADO) === 'PENDIENTE' && (r['FECHA VENC.'] || r['FECHA VENCIMIENTO']))
        .map(r => {
            const raw = String(r['FECHA VENC.'] || r['FECHA VENCIMIENTO'] || '');
            let d = null;
            if (raw.includes('-')) d = new Date(raw.split(' ')[0]);
            else if (raw.includes('/')) { const p = raw.split('/'); d = new Date(`${p[2]}-${p[1]}-${p[0]}`); }
            return { r, d };
        })
        .filter(x => x.d && !isNaN(x.d))
        .sort((a,b) => a.d - b.d)
        .slice(0, 12);

    $('dashUpcoming').innerHTML = upcoming.length ? upcoming.map(({ r, d }) => {
        const diff = Math.round((d - hoy) / 86400000);
        const cls  = diff < 0 ? 'expired' : diff <= 3 ? 'urgent' : diff <= 10 ? 'soon' : 'ok';
        const lbl  = diff < 0 ? `Venc. ${Math.abs(diff)}d` : diff === 0 ? 'HOY' : `${diff}d`;
        return `
        <div class="dash-upcoming-row">
            <span class="dash-up-days ${cls}">${lbl}</span>
            <div class="dash-up-info">
                <div class="dash-up-desc">${esc(r.DESCRIPCION || '—')}</div>
                <div class="dash-up-meta">${esc(r.PROVEEDOR||'')} · ${esc(r.SUCURSAL||'')} · ${fmtCantidad(r)}</div>
            </div>
            <span class="dash-up-nc">${fmtPeso(calcNCsugerida(r).total)}</span>
        </div>`;
    }).join('') : '<p style="font-size:12px;color:var(--text3);padding:8px 0">Sin registros pendientes con fecha</p>';
}

// ══════════════════════════════════════════════════════════════
//  SIDEBAR
// ══════════════════════════════════════════════════════════════
function renderProvList() {
    const q = ($('provSearch').value || '').toLowerCase().trim();
    const f = state.sidebarFilter;
    let provs = Object.values(state.proveedores);
    if (q) provs = provs.filter(p => p.nombre.toLowerCase().includes(q));
    if (f === 'pendientes') provs = provs.filter(p => p.registros.some(r => normEstado(r.ESTADO) === 'PENDIENTE'));
    if (f === 'vinculados') provs = provs.filter(p => p.registros.some(r => state.vinculados[r.ID]?.length > 0));
    provs.sort((a,b) => {
        const pa = a.registros.filter(r => normEstado(r.ESTADO) === 'PENDIENTE').length;
        const pb = b.registros.filter(r => normEstado(r.ESTADO) === 'PENDIENTE').length;
        return pb - pa || a.nombre.localeCompare(b.nombre);
    });

    const list = $('provList');
    list.innerHTML = '';
    if (!provs.length) {
        list.innerHTML = `<li style="padding:18px 10px;font-size:11px;color:var(--text3);text-align:center">Sin resultados</li>`;
        return;
    }
    provs.forEach((p, i) => {
        const pend = p.registros.filter(r => normEstado(r.ESTADO) === 'PENDIENTE').length;
        const gest = p.registros.filter(r => normEstado(r.ESTADO) === 'EN GESTION').length;
        const nc   = p.registros.filter(r => normEstado(r.ESTADO) === 'N/C RECIBIDA').length;
        const vink = p.registros.some(r => state.vinculados[r.ID]?.length > 0);
        const li = document.createElement('li');
        li.className = 'prov-item' + (state.activeProv === p.nombre ? ' active' : '');
        li.style.animationDelay = `${i * .025}s`;
        li.innerHTML = `
            <span class="prov-item-name">${esc(p.nombre)}</span>
            <div class="prov-item-meta">
                ${p.cod ? `<span class="prov-item-cod">COD ${p.cod}</span>` : ''}
                <div class="prov-item-counts">
                    ${pend ? `<span class="pic-badge pic-badge--p">${pend}P</span>` : ''}
                    ${gest ? `<span class="pic-badge pic-badge--g">${gest}G</span>` : ''}
                    ${nc   ? `<span class="pic-badge pic-badge--n">${nc}N/C</span>` : ''}
                    ${vink ? `<span class="pic-badge pic-badge--vink">⚠</span>` : ''}
                </div>
            </div>`;
        li.addEventListener('click', () => selectProveedor(p.nombre));
        list.appendChild(li);
    });
}

function selectProveedor(nombre) {
    state.activeProv = nombre;
    document.querySelectorAll('.prov-item').forEach(li => {
        li.classList.toggle('active', li.querySelector('.prov-item-name')?.textContent.trim() === nombre);
    });
    renderKanban(nombre);
    $('boardPlaceholder').style.display = 'none';
    if (window.innerWidth <= 700) document.getElementById('sidebar').classList.add('collapsed');
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
    const regs = prov.registros;

    $('provHeaderName').textContent = prov.nombre;
    $('provHeaderMeta').textContent = `${prov.cod ? `COD ${prov.cod} · ` : ''}${regs.length} registros · ${new Set(regs.map(r => r.SUCURSAL)).size} sucursal(es)`;

    const montoNC  = regs.reduce((s,r) => s + (parseFloat(r['MONTO N/C']) || 0), 0);
    const costoEst = regs.reduce((s,r) => s + calcNCsugerida(r).total, 0);
    const pends    = regs.filter(r => normEstado(r.ESTADO) === 'PENDIENTE').length;

    $('provHeaderKpis').innerHTML = `
        <div class="phkpi"><span class="phkpi-val" style="color:var(--yellow)">${pends}</span><span class="phkpi-lbl">Pendientes</span></div>
        <div class="phkpi"><span class="phkpi-val" style="color:var(--accent)">${fmtPeso(costoEst)}</span><span class="phkpi-lbl">Costo est.</span></div>
        <div class="phkpi"><span class="phkpi-val" style="color:var(--green)">${fmtPeso(montoNC)}</span><span class="phkpi-lbl">N/C recibida</span></div>`;

    const vinksN = regs.filter(r => state.vinculados[r.ID]?.length > 0).length;
    $('vinkAlert').style.display = vinksN ? 'flex' : 'none';
    if (vinksN) $('vinkAlertSub').textContent = `${vinksN} registros tienen vínculos acción+vencimiento. Revisalos antes de reclamar para no duplicar el monto.`;

    const COLS = {
        'PENDIENTE':    { id:'PENDIENTE',   amtFn: r => calcNCsugerida(r).total },
        'EN GESTION':   { id:'EN_GESTION',  amtFn: r => calcNCsugerida(r).total },
        'N/C RECIBIDA': { id:'NC_RECIBIDA', amtFn: r => parseFloat(r['MONTO N/C']) || 0 },
        'RECHAZADA':    { id:'RECHAZADA',   amtFn: () => 0 },
    };

    Object.entries(COLS).forEach(([estado, cfg]) => {
        const col     = $(`col-${cfg.id}`);
        const cnt     = $(`cnt-${cfg.id}`);
        const amt     = $(`amt-${cfg.id}`);
        const colRegs = regs.filter(r => normEstado(r.ESTADO) === estado);

        cnt.textContent = colRegs.length;
        const colMonto = colRegs.reduce((s,r) => s + cfg.amtFn(r), 0);
        amt.textContent = colMonto > 0 ? fmtPeso(colMonto) : '';

        col.innerHTML = '';
        if (!colRegs.length) {
            col.innerHTML = `<div class="kcol-empty">Vacío — arrastrá acá</div>`;
        } else {
            colRegs.forEach((r, i) => col.appendChild(buildCard(r, i)));
        }
        setupDropZone(col, estado);
    });
}

function buildCard(r, idx) {
    const card = document.createElement('div');
    card.className = 'kcard';
    card.style.animationDelay = `${idx * .035}s`;
    card.dataset.id     = r.ID;
    card.dataset.estado = normEstado(r.ESTADO);
    card.draggable = true;

    const linked = (state.vinculados[r.ID] || []).length > 0;
    if (linked) card.classList.add('is-linked');

    const nc    = calcNCsugerida(r);
    const ncStr = r['MONTO N/C'] ? fmtPeso(parseFloat(r['MONTO N/C'])) : '';

    card.innerHTML = `
        ${linked ? '<div class="kcard-vink-dot" title="Vinculado — clic para ver"></div>' : ''}
        <div class="kcard-drag" title="Arrastrar para cambiar estado">
            <svg viewBox="0 0 10 18" fill="currentColor" width="10" height="18">
                <circle cx="3" cy="3" r="1.2"/><circle cx="3" cy="9" r="1.2"/><circle cx="3" cy="15" r="1.2"/>
                <circle cx="7" cy="3" r="1.2"/><circle cx="7" cy="9" r="1.2"/><circle cx="7" cy="15" r="1.2"/>
            </svg>
        </div>
        <div class="kcard-motivo-dot" style="background:${motivoColor(r.MOTIVO)}"></div>
        <div class="kcard-main">
            <div class="kcard-desc" title="${esc(r.DESCRIPCION||'')}">${esc(r.DESCRIPCION||'—')}</div>
            <div class="kcard-meta">
                <span class="kcard-meta-item">${fmtCantidad(r)}</span>
                <span class="kcard-meta-sep">·</span>
                <span class="kcard-meta-item">${esc(motivoCorto(r.MOTIVO))}</span>
                ${r['FECHA VENC.']||r['FECHA VENCIMIENTO'] ? `<span class="kcard-meta-sep">·</span><span class="kcard-meta-item">${esc(r['FECHA VENC.']||r['FECHA VENCIMIENTO']||'')}</span>` : ''}
            </div>
        </div>
        <div class="kcard-right">
            ${ncStr ? `<span class="kcard-nc-val">✓${ncStr}</span>` : `<span class="kcard-nc-est">${fmtPeso(nc.total)}</span>`}
            <span class="kcard-suc">${esc(r.SUCURSAL||'')}</span>
        </div>`;

    card.addEventListener('click', e => { if (!e.target.closest('.kcard-drag')) openModal(r); });

    card.addEventListener('dragstart', e => {
        state.drag = { id: r.ID, srcEstado: normEstado(r.ESTADO), record: r };
        setTimeout(() => card.classList.add('dragging'), 0);
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', r.ID);
    });
    card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        document.querySelectorAll('.drop-placeholder').forEach(p => p.remove());
        document.querySelectorAll('.kanban-col').forEach(c => c.classList.remove('drag-over'));
    });

    return card;
}

function setupDropZone(colEl, estado) {
    colEl.addEventListener('dragover', e => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        const col = colEl.closest('.kanban-col');
        document.querySelectorAll('.kanban-col').forEach(c => c.classList.remove('drag-over'));
        col.classList.add('drag-over');
        document.querySelectorAll('.drop-placeholder').forEach(p => p.remove());
        const ph = document.createElement('div');
        ph.className = 'drop-placeholder';
        colEl.appendChild(ph);
    });

    colEl.addEventListener('dragleave', e => {
        const col = colEl.closest('.kanban-col');
        if (!col.contains(e.relatedTarget)) {
            col.classList.remove('drag-over');
            document.querySelectorAll('.drop-placeholder').forEach(p => p.remove());
        }
    });

    colEl.addEventListener('drop', async e => {
        e.preventDefault();
        document.querySelectorAll('.drop-placeholder').forEach(p => p.remove());
        colEl.closest('.kanban-col').classList.remove('drag-over');

        const { id, srcEstado, record } = state.drag;
        if (!id || srcEstado === estado) return;

        const idx = state.data.findIndex(d => d.ID === id);
        if (idx !== -1) state.data[idx].ESTADO = estado;
        processData(); renderProvList(); renderKanban(state.activeProv);
        showToast('wrn', `Movido a "${estado}" — guardando...`);

        try {
            const res = await fetch(APPS_SCRIPT_URL, {
                method: 'POST',
                body: JSON.stringify({
                    action: 'updateRecord', id, estado,
                    montoNC: record?.['MONTO N/C'] || '',
                    observacionNC: record?.['OBSERVACION N/C'] || '',
                }),
            });
            const data = await res.json();
            if (!data.success) throw new Error(data.message);
            showToast('ok', `Guardado en "${estado}"`);
        } catch (err) {
            if (idx !== -1) state.data[idx].ESTADO = srcEstado;
            processData(); renderProvList(); renderKanban(state.activeProv);
            showToast('err', `Error al guardar: ${err.message}`);
        }
    });
}

function showKanban(show) {
    $('kanbanWrapper').style.display = show ? 'flex' : 'none';
}

// ══════════════════════════════════════════════════════════════
//  MODAL
// ══════════════════════════════════════════════════════════════
function openModal(r) {
    state.modalRecord = r;
    state.modalEstado = normEstado(r.ESTADO);
    $('modalTitle').textContent = r.DESCRIPCION || 'Registro';
    $('modalDot').style.background = motivoColor(r.MOTIVO);
    $('modalDot').style.boxShadow  = `0 0 7px ${motivoColor(r.MOTIVO)}`;

    const cn = parseFloat(r['COSTO NETO']) || 0;
    const iv = parseFloat(r['IVA %']) || 0;
    const cf = cn * (1 + iv / 100);

    $('modalProductInfo').innerHTML = `
        <div class="mpi-full"><div class="mpi-field-lbl">Descripción</div><div class="mpi-field-val">${esc(r.DESCRIPCION||'—')}</div></div>
        <div><div class="mpi-field-lbl">Proveedor</div><div class="mpi-field-val">${esc(r.PROVEEDOR||'—')}</div></div>
        <div><div class="mpi-field-lbl">Motivo</div><div class="mpi-field-val">${esc(r.MOTIVO||'—')}</div></div>
        <div><div class="mpi-field-lbl">Cantidad</div><div class="mpi-field-val mono">${fmtCantidad(r)}</div></div>
        <div><div class="mpi-field-lbl">Vencimiento</div><div class="mpi-field-val mono">${esc(r['FECHA VENC.']||r['FECHA VENCIMIENTO']||'—')}</div></div>
        <div><div class="mpi-field-lbl">Costo neto</div><div class="mpi-field-val mono">${fmtPeso(cn)}</div></div>
        <div><div class="mpi-field-lbl">Costo c/IVA ${iv}%</div><div class="mpi-field-val mono">${fmtPeso(cf)}</div></div>
        <div><div class="mpi-field-lbl">Sucursal · ID</div><div class="mpi-field-val mono">${esc(r.SUCURSAL||'')} · ${esc(r.ID||'')}</div></div>`;

    buildNCCalc(r);

    const vincs = state.vinculados[r.ID] || [];
    $('modalVinkWarn').style.display = vincs.length ? 'flex' : 'none';
    if (vincs.length) {
        const cm = { 'Acción':'#a855f7','Vencido':'#f97316','Decomiso':'#ef4444','Otro':'#818cf8' };
        $('modalVinkList').innerHTML = vincs.map(v => {
            const col = cm[motivoGrupo(v.MOTIVO)] || '#818cf8';
            return `<div class="mvw-linked-item">
                <span class="mvw-li-badge" style="background:${col}22;color:${col}">${esc(motivoCorto(v.MOTIVO))}</span>
                <div class="mvw-li-info">
                    <div class="mvw-li-desc">${esc(v.DESCRIPCION||'—')}</div>
                    <div class="mvw-li-id">${esc(v.ID)} · ${fmtCantidad(v)} · Suc: ${esc(v.SUCURSAL||'')} · Est: ${fmtPeso(calcNCsugerida(v).total)}</div>
                </div>
                <span style="font-family:var(--mono);font-size:10px;color:var(--accent)">${fmtPeso(calcNCsugerida(v).total)}</span>
            </div>`;
        }).join('');
    }

    document.querySelectorAll('.est-btn').forEach(b => b.classList.toggle('selected', b.dataset.estado === state.modalEstado));
    $('modalMonto').value      = r['MONTO N/C'] || '';
    $('modalObservacion').value = r['OBSERVACION N/C'] || '';
    $('modalOverlay').classList.add('open');
    document.body.style.overflow = 'hidden';
}

function buildNCCalc(r) {
    const nc = calcNCsugerida(r);
    state.modalCalcNC = nc.total;
    $('ncCalc').innerHTML = nc.detalles.map(row => `
        <div class="nc-calc-row">
            <span class="ncr-label">${esc(row.label)}</span>
            <span class="ncr-formula">${esc(row.formula||'')}</span>
            <span class="ncr-value ${row.isInfo?'info':''}">${esc(row.value)}</span>
        </div>`).join('') +
        (nc.total > 0 ? `
        <div class="nc-calc-row" style="background:rgba(34,197,94,.04)">
            <span class="ncr-label" style="font-weight:700;color:var(--text)">Total N/C sugerida</span>
            <span class="ncr-formula"></span>
            <span class="ncr-value total">${fmtPeso(nc.total)}</span>
        </div>` : '');
}

function closeModal() {
    $('modalOverlay').classList.remove('open');
    document.body.style.overflow = '';
    state.modalRecord = null;
}

// ══════════════════════════════════════════════════════════════
//  GUARDAR
// ══════════════════════════════════════════════════════════════
async function saveRecord() {
    if (state.saving || !state.modalRecord) return;
    const r       = state.modalRecord;
    const nuevoE  = state.modalEstado;
    const montoNC = parseFloat($('modalMonto').value) || 0;
    const observ  = $('modalObservacion').value.trim();

    state.saving = true;
    const btn = $('modalSave');
    btn.classList.add('saving');
    btn.disabled = true;

    try {
        const res = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action:'updateRecord', id:r.ID, estado:nuevoE, montoNC:montoNC||'', observacionNC:observ }),
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        if (!data.success) throw new Error(data.message || 'Error del servidor');

        const idx = state.data.findIndex(d => d.ID === r.ID);
        if (idx !== -1) {
            state.data[idx].ESTADO             = nuevoE;
            state.data[idx]['MONTO N/C']       = montoNC || '';
            state.data[idx]['OBSERVACION N/C'] = observ;
        }
        processData(); renderProvList(); renderKanban(state.activeProv);
        closeModal();
        showToast('ok', 'Registro actualizado correctamente');
    } catch (err) {
        showToast('err', `Error: ${err.message}`);
    } finally {
        state.saving = false;
        btn.classList.remove('saving');
        btn.disabled = false;
    }
}

// ══════════════════════════════════════════════════════════════
//  CALCULADORA N/C
// ══════════════════════════════════════════════════════════════
function calcNCsugerida(r) {
    const cn = parseFloat(r['COSTO NETO']) || 0;
    const iv = parseFloat(r['IVA %']) || 0;
    const cf = cn * (1 + iv / 100);
    const ca = parseFloat(r.CANTIDAD) || 0;
    const g  = motivoGrupo(r.MOTIVO);
    const mu = (r.MOTIVO || '').toUpperCase();
    const det = [];
    let total = 0;

    det.push({ label:'Costo unitario c/IVA', formula:`${fmtPeso(cn)} + ${iv}% IVA`, value:fmtPeso(cf) });
    det.push({ label:'Cantidad', formula:'', value:fmtCantidad(r) });

    if (g === 'Acción') {
        let pct = 50;
        const m = mu.match(/(\d+)%/); if (m) pct = parseFloat(m[1]);
        if (mu.includes('2X1') || mu.includes('2 X 1')) pct = 50;
        total = cf * ca * (pct / 100);
        det.push({ label:`Descuento acción (${pct}%)`, formula:`${fmtPeso(cf)} × ${ca} × ${pct}%`, value:fmtPeso(total) });
        det.push({ label:'⚠ Verificar registros vinculados', formula:'', value:'Ver vínculos', isInfo:true });
    } else if (g === 'Vencido') {
        total = cf * ca;
        det.push({ label:'Vencimiento (100% costo c/IVA)', formula:`${fmtPeso(cf)} × ${ca}`, value:fmtPeso(total) });
    } else if (g === 'Decomiso') {
        total = cf * ca;
        det.push({ label:'Decomiso (100% costo c/IVA)', formula:`${fmtPeso(cf)} × ${ca}`, value:fmtPeso(total) });
        det.push({ label:'Verificar acuerdo con proveedor', formula:'', value:'Puede variar', isInfo:true });
    } else {
        det.push({ label:'Calcular manualmente', formula:'', value:'—', isInfo:true });
    }
    return { total, detalles: det };
}

// ══════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════

// ─── CRÍTICO: VENCIDO debe chequearse ANTES que ACCION
// porque "VENCIDO ACCION 50% OFF" contiene "ACCION" y "OFF"
// pero su grupo correcto es 'Vencido' ───────────────────────
function motivoGrupo(motivo) {
    if (!motivo) return 'Otro';
    const m = motivo.toUpperCase();
    if (m.includes('VENCIDO') || m.includes('VENCIMIENTO')) return 'Vencido';  // ← primero
    if (m.includes('ACCION') || m.includes('ACCIÓN') || m.includes('OFF') || m.includes('2X1')) return 'Acción';
    if (m.includes('DECOMISO') || m.includes('ROTO') || m.includes('MAL ESTADO')) return 'Decomiso';
    return 'Otro';
}

function motivoColor(motivo) {
    return { 'Acción':'#a855f7','Vencido':'#f97316','Decomiso':'#ef4444','Otro':'#818cf8' }[motivoGrupo(motivo)] || '#818cf8';
}

function motivoCorto(motivo) {
    if (!motivo) return '—';
    const map = {
        'ACCION 2X1':'2×1','VENCIDO ACCION 2X1':'Vto.2×1',
        'ACCION 50% OFF':'50%OFF','VENCIDO ACCION 50% OFF':'Vto.50%',
        'VENCIDO':'Vencido','ROTO/DAÑADO':'Roto','MAL ESTADO':'Mal estado',
        'DECOMISO FRUTA Y VERDURA':'Dec.F&V','DECOMISO CARNICERIA':'Dec.Carn.',
        'DECOMISO FIAMBRERIA':'Dec.Fiam.','DECOMISO':'Decomiso','OTRO':'Otro',
    };
    return map[motivo.toUpperCase()] || motivo.slice(0, 12);
}

function normEstado(estado) {
    if (!estado || !String(estado).trim()) return 'PENDIENTE';
    return String(estado).trim().toUpperCase();
}

function fmtPeso(n) {
    if (!n || isNaN(Number(n)) || Number(n) === 0) return '$0';
    return '$' + Number(n).toLocaleString('es-AR', { minimumFractionDigits:0, maximumFractionDigits:0 });
}

function fmtCantidad(r) {
    const c = r.CANTIDAD;
    if (c === '' || c == null) return '—';
    const n = parseFloat(c); if (isNaN(n)) return String(c);
    const esKg = String(r['UNIDAD CANTIDAD']||'').toLowerCase() === 'kg'
              || String(r['UNID. CANTIDAD']||'').toLowerCase() === 'kg';
    return esKg ? `${n.toFixed(3)} kg` : `${n} u.`;
}

function setSyncStatus(type) {
    const badge = $('syncBadge'), text = $('syncText');
    badge.className = 'sync-badge';
    if (type === 'ok') {
        badge.classList.add('ok');
        const t = state.lastUpdate;
        text.textContent = t ? `Act. ${t.toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit'})}` : 'Actualizado';
    } else if (type === 'err') { badge.classList.add('err'); text.textContent = 'Sin conexión'; }
    else { text.textContent = 'Cargando...'; }
}

let toastTimer;
function showToast(type, msg, ms = 3500) {
    const t = $('toast');
    $('toastMsg').textContent = msg;
    t.className = `toast ${type} show`;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), ms);
}

function esc(s) {
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}