// ── Reloj en tiempo real ──
function updateClock() {
  const now = new Date();
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  const ss = String(now.getSeconds()).padStart(2, '0');
  document.getElementById('clock').textContent = `${hh}:${mm}:${ss}`;
}

updateClock();
setInterval(updateClock, 1000);

// ── Fecha en footer ──
const opts = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
document.getElementById('footerDate').textContent =
  new Date().toLocaleDateString('es-AR', opts).toUpperCase();