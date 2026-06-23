// ============================================================
//  APPS SCRIPT · SEGUIMIENTO 3º ESO · Las Chapas ECOS
//  Repaso de Verano 2025
// ============================================================
//
//  INSTRUCCIONES DE INSTALACIÓN:
//  1. Abre Google Sheets y crea un nuevo archivo llamado
//     "Seguimiento 3º ESO - Repaso Verano 2025"
//  2. Ve a Extensiones → Apps Script
//  3. Borra todo el contenido y pega este código completo
//  4. Guarda (Ctrl+S) con nombre "Seguimiento 3ESO"
//  5. Ejecuta la función "inicializarHojas" una sola vez:
//       - Haz clic en el menú desplegable junto al botón ▶
//       - Selecciona "inicializarHojas" y pulsa ▶
//       - Acepta los permisos que pida Google
//  6. Despliega como aplicación web:
//       - Implementar → Nueva implementación
//       - Tipo: Aplicación web
//       - Ejecutar como: Yo (tu cuenta)
//       - Quién tiene acceso: Cualquier usuario
//       - Pulsa "Implementar" y copia la URL que aparece
//  7. Pega esa URL en la constante SHEETS_URL de los módulos HTML
//
// ============================================================

// ── Cabeceras por bloque ─────────────────────────────────────

const CABECERAS = {

  'Bloque 0 - Números': [
    'Marca de tiempo', 'Nombre', 'Apellidos', 'Grupo',
    // Naturales
    'NAT - Operaciones',
    'NAT - Jerarquía',
    'NAT - Problemas',
    // Enteros
    'ENT - Operaciones',
    'ENT - Jerarquía',
    'ENT - Problemas',
    // Racionales
    'RAC - Operaciones',
    'RAC - Jerarquía',
    'RAC - Problemas',
    // Radicales
    'RAD - Notación potencias',
    'RAD - Extraer/introducir factores',
    'RAD - Operaciones',
    // Resumen
    'Módulos completados (de 12)',
    '% completado'
  ],

  'Bloque 1 - Álgebra': [
    'Marca de tiempo', 'Nombre', 'Apellidos', 'Grupo',
    'Polinomios y operaciones',
    'Productos notables',
    'Factorización de polinomios',
    'Módulos completados (de 3)',
    '% completado'
  ],

  'Bloque 2 - Ecuaciones': [
    'Marca de tiempo', 'Nombre', 'Apellidos', 'Grupo',
    'Ecuaciones 1er grado',
    'Ecuaciones 2º grado',
    'Ecuaciones polinómicas',
    'Ecuaciones racionales',
    'Ecuaciones irracionales',
    'Módulos completados (de 5)',
    '% completado'
  ],

  'Resumen general': [
    'Marca de tiempo', 'Nombre', 'Apellidos', 'Grupo',
    'B0 completado (%)',
    'B1 completado (%)',
    'B2 completado (%)',
    'Total módulos completados (de 20)',
    '% global'
  ]
};

// ── Mapeo: id de módulo → columna en su hoja ─────────────────
// (columna 1 = primera columna de datos, después de tiempo/nombre/apellidos/grupo)

const MODULO_COLUMNA = {
  // Bloque 0
  'b0-nat-op':   5,
  'b0-nat-jer':  6,
  'b0-nat-prob': 7,
  'b0-ent-op':   8,
  'b0-ent-jer':  9,
  'b0-ent-prob': 10,
  'b0-rac-op':   11,
  'b0-rac-jer':  12,
  'b0-rac-prob': 13,
  'b0-rad-pot':  14,
  'b0-rad-fac':  15,
  'b0-rad-op':   16,
  // Bloque 1
  'b1-poli':     5,
  'b1-pn':       6,
  'b1-fact':     7,
  // Bloque 2
  'b2-ec1':      5,
  'b2-ec2':      6,
  'b2-ecpoli':   7,
  'b2-ecrac':    8,
  'b2-ecirr':    9
};

const HOJA_POR_BLOQUE = {
  'b0': 'Bloque 0 - Números',
  'b1': 'Bloque 1 - Álgebra',
  'b2': 'Bloque 2 - Ecuaciones'
};

const TOTAL_POR_BLOQUE = { 'b0': 12, 'b1': 3, 'b2': 5 };

// ── Inicializar hojas ────────────────────────────────────────

function inicializarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.entries(CABECERAS).forEach(([nombreHoja, cabeceras]) => {
    let hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) {
      hoja = ss.insertSheet(nombreHoja);
    }
    // Solo escribe cabeceras si la hoja está vacía
    if (hoja.getLastRow() === 0) {
      hoja.appendRow(cabeceras);
      hoja.getRange(1, 1, 1, cabeceras.length)
          .setFontWeight('bold')
          .setBackground('#34A853')
          .setFontColor('white');
      hoja.setFrozenRows(1);
    }
  });

  // Eliminar la hoja "Hoja 1" por defecto si existe
  const hojaDefault = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hojaDefault) ss.deleteSheet(hojaDefault);

  SpreadsheetApp.getUi().alert('✅ Hojas creadas correctamente. Ya puedes desplegar como aplicación web.');
}

// ── Punto de entrada HTTP (GET) ──────────────────────────────

function doGet(e) {
  return manejarPeticion(e);
}

function doPost(e) {
  return manejarPeticion(e);
}

function manejarPeticion(e) {
  const params = e.parameter || {};

  try {
    const nombre    = params.nombre    || '';
    const apellidos = params.apellidos || '';
    const grupo     = params.grupo     || '';
    const moduloId  = params.modulo    || '';
    const bloque    = params.bloque    || '';   // 'b0', 'b1' o 'b2'
    const accion    = params.accion    || 'completar';

    if (accion === 'completar') {
      registrarModulo(nombre, apellidos, grupo, bloque, moduloId);
    } else if (accion === 'resumen') {
      const b0pct = parseFloat(params.b0pct || '0');
      const b1pct = parseFloat(params.b1pct || '0');
      const b2pct = parseFloat(params.b2pct || '0');
      registrarResumen(nombre, apellidos, grupo, b0pct, b1pct, b2pct);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Registrar módulo completado ──────────────────────────────

function registrarModulo(nombre, apellidos, grupo, bloque, moduloId) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHoja = HOJA_POR_BLOQUE[bloque];
  if (!nombreHoja) return;

  const hoja    = ss.getSheetByName(nombreHoja);
  if (!hoja) return;

  const col     = MODULO_COLUMNA[moduloId];
  const total   = TOTAL_POR_BLOQUE[bloque];
  const clave   = (nombre + '|' + apellidos + '|' + grupo).toLowerCase();

  // Buscar fila existente del alumno
  const datos   = hoja.getDataRange().getValues();
  let fila      = -1;
  for (let i = 1; i < datos.length; i++) {
    const filaKey = (datos[i][1] + '|' + datos[i][2] + '|' + datos[i][3]).toLowerCase();
    if (filaKey === clave) { fila = i + 1; break; }
  }

  if (fila === -1) {
    // Nueva fila
    const nuevaFila = new Array(CABECERAS[nombreHoja].length).fill('');
    nuevaFila[0] = new Date();
    nuevaFila[1] = nombre;
    nuevaFila[2] = apellidos;
    nuevaFila[3] = grupo;
    if (col) nuevaFila[col - 1] = '✓';
    hoja.appendRow(nuevaFila);
    fila = hoja.getLastRow();
  } else {
    // Actualizar fila existente
    if (col) hoja.getRange(fila, col).setValue('✓');
    hoja.getRange(fila, 1).setValue(new Date());
  }

  // Recalcular resumen de módulos completados
  const filaData = hoja.getRange(fila, 1, 1, CABECERAS[nombreHoja].length).getValues()[0];
  let completados = 0;
  // Las columnas de módulos van desde la 5 hasta 5+total-1
  for (let c = 4; c < 4 + total; c++) {
    if (filaData[c] === '✓') completados++;
  }
  const pct = Math.round(completados / total * 100);

  // Columnas de resumen: penúltima y última
  const colCompletados = CABECERAS[nombreHoja].length - 1;
  const colPct         = CABECERAS[nombreHoja].length;
  hoja.getRange(fila, colCompletados).setValue(completados);
  hoja.getRange(fila, colPct).setValue(pct + '%');
}

// ── Registrar resumen global ─────────────────────────────────

function registrarResumen(nombre, apellidos, grupo, b0pct, b1pct, b2pct) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Resumen general');
  if (!hoja) return;

  const clave = (nombre + '|' + apellidos + '|' + grupo).toLowerCase();
  const datos = hoja.getDataRange().getValues();
  let fila    = -1;
  for (let i = 1; i < datos.length; i++) {
    const filaKey = (datos[i][1] + '|' + datos[i][2] + '|' + datos[i][3]).toLowerCase();
    if (filaKey === clave) { fila = i + 1; break; }
  }

  // Módulos totales completados aproximados
  const b0comp = Math.round(b0pct / 100 * 12);
  const b1comp = Math.round(b1pct / 100 * 3);
  const b2comp = Math.round(b2pct / 100 * 5);
  const totalComp = b0comp + b1comp + b2comp;
  const pctGlobal = Math.round(totalComp / 20 * 100);

  const fila_vals = [
    new Date(), nombre, apellidos, grupo,
    b0pct + '%', b1pct + '%', b2pct + '%',
    totalComp, pctGlobal + '%'
  ];

  if (fila === -1) {
    hoja.appendRow(fila_vals);
  } else {
    hoja.getRange(fila, 1, 1, fila_vals.length).setValues([fila_vals]);
  }
}
