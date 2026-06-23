// ============================================================
//  FRAGMENTO JS PARA INTEGRACIÓN CON GOOGLE SHEETS
//  Añadir a cada módulo HTML, dentro del <script> existente
// ============================================================
//
//  PASOS:
//  1. Sustituye SHEETS_URL por la URL real de tu Apps Script
//  2. Ajusta MODULE_ID y BLOQUE según el módulo
//  3. Este código ya llama a enviarASheets() dentro de marcarCompletado()
//     — solo tienes que añadir la función enviarASheets y modificar
//     marcarCompletado() como se muestra abajo
// ============================================================

// ── Constantes (ajustar en cada módulo) ─────────────────────

const SHEETS_URL  = 'PEGA_AQUÍ_LA_URL_DE_TU_APPS_SCRIPT';
const MODULE_ID   = 'b0-nat-op';   // ← cambia según el módulo
const BLOQUE      = 'b0';          // ← 'b0', 'b1' o 'b2'
const SESSION_KEY = 'modulos_3eso_b0'; // ← 'modulos_3eso_b0/b1/b2'

// ── Enviar al Sheets ─────────────────────────────────────────

function enviarASheets(moduloId, bloque) {
  const nombre    = sessionStorage.getItem('alumno_nombre')    || '';
  const apellidos = sessionStorage.getItem('alumno_apellidos') || '';
  const grupo     = sessionStorage.getItem('alumno_grupo')     || '';

  if (!nombre || !apellidos) return; // sin datos de alumno, no enviar

  const url = SHEETS_URL
    + '?accion=completar'
    + '&nombre='    + encodeURIComponent(nombre)
    + '&apellidos=' + encodeURIComponent(apellidos)
    + '&grupo='     + encodeURIComponent(grupo)
    + '&modulo='    + encodeURIComponent(moduloId)
    + '&bloque='    + encodeURIComponent(bloque);

  fetch(url, { method: 'GET', mode: 'no-cors' })
    .catch(() => {}); // silencioso — no interrumpir al alumno si falla
}

// ── Versión de marcarCompletado() con envío al Sheets ────────
//    (sustituye la función que ya tienes en el módulo)

function marcarCompletado() {
  const completados = JSON.parse(sessionStorage.getItem(SESSION_KEY) || '[]');
  if (!completados.includes(MODULE_ID)) completados.push(MODULE_ID);
  sessionStorage.setItem(SESSION_KEY, JSON.stringify(completados));

  // Enviar al Google Sheets
  enviarASheets(MODULE_ID, BLOQUE);

  // Enviar resumen global si todos los módulos del bloque están completos
  enviarResumenSiProcede();

  document.getElementById('completar-wrap').style.display = 'none';
  document.getElementById('completado-msg').style.display = 'block';
}

// ── Enviar resumen cuando un bloque se completa ──────────────

function enviarResumenSiProcede() {
  const TOTALES = { b0: 12, b1: 3, b2: 5 };
  const KEYS    = {
    b0: 'modulos_3eso_b0',
    b1: 'modulos_3eso_b1',
    b2: 'modulos_3eso_b2'
  };

  const b0 = JSON.parse(sessionStorage.getItem(KEYS.b0) || '[]');
  const b1 = JSON.parse(sessionStorage.getItem(KEYS.b1) || '[]');
  const b2 = JSON.parse(sessionStorage.getItem(KEYS.b2) || '[]');

  const b0pct = Math.round(b0.length / TOTALES.b0 * 100);
  const b1pct = Math.round(b1.length / TOTALES.b1 * 100);
  const b2pct = Math.round(b2.length / TOTALES.b2 * 100);

  // Solo enviar resumen si al menos un bloque está al 100%
  if (b0pct < 100 && b1pct < 100 && b2pct < 100) return;

  const nombre    = sessionStorage.getItem('alumno_nombre')    || '';
  const apellidos = sessionStorage.getItem('alumno_apellidos') || '';
  const grupo     = sessionStorage.getItem('alumno_grupo')     || '';

  const url = SHEETS_URL
    + '?accion=resumen'
    + '&nombre='    + encodeURIComponent(nombre)
    + '&apellidos=' + encodeURIComponent(apellidos)
    + '&grupo='     + encodeURIComponent(grupo)
    + '&b0pct='     + b0pct
    + '&b1pct='     + b1pct
    + '&b2pct='     + b2pct;

  fetch(url, { method: 'GET', mode: 'no-cors' })
    .catch(() => {});
}
