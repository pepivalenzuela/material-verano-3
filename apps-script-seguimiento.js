// ============================================================
//  APPS SCRIPT · SEGUIMIENTO 3º ESO · Las Chapas ECOS
//  Repaso de Verano 2025 — con Geometría
// ============================================================
//  INSTRUCCIONES:
//  1. Sheets → Extensiones → Apps Script
//  2. Pega este código completo y guarda
//  3. Ejecuta "inicializarHojas" una vez
//  4. Implementar → Nueva implementación → Aplicación web
//     Ejecutar como: Yo · Acceso: Cualquier usuario
//  5. Copia la URL y pégala en SHEETS_URL de los módulos HTML
// ============================================================

const CABECERAS = {
  'Bloque 0 - Números': [
    'Marca de tiempo','Nombre','Apellidos','Grupo',
    'NAT-Operaciones','NAT-Jerarquía','NAT-Problemas',
    'ENT-Operaciones','ENT-Jerarquía','ENT-Problemas',
    'RAC-Operaciones','RAC-Jerarquía','RAC-Problemas',
    'RAD-Potencias','RAD-Factores','RAD-Operaciones',
    'Completados (de 12)','% completado'
  ],
  'Bloque 1 - Álgebra': [
    'Marca de tiempo','Nombre','Apellidos','Grupo',
    'Polinomios','Productos notables','Factorización',
    'Completados (de 3)','% completado'
  ],
  'Bloque 2 - Ecuaciones': [
    'Marca de tiempo','Nombre','Apellidos','Grupo',
    'Ec. 1er grado','Ec. 2º grado','Ec. polinómicas','Ec. racionales','Ec. irracionales',
    'Completados (de 5)','% completado'
  ],
  'Bloque 3 - Geometría': [
    'Marca de tiempo','Nombre','Apellidos','Grupo',
    'Formas geométricas','Ángulos y rectas','Semejanza y Tales','Pitágoras','Trigonometría',
    'Completados (de 5)','% completado'
  ],
  'Resumen general': [
    'Marca de tiempo','Nombre','Apellidos','Grupo',
    'B0 %','B1 %','B2 %','B3 %',
    'Total completados (de 25)','% global'
  ]
};

const MODULO_COLUMNA = {
  'b0-nat-op':5,'b0-nat-jer':6,'b0-nat-prob':7,
  'b0-ent-op':8,'b0-ent-jer':9,'b0-ent-prob':10,
  'b0-rac-op':11,'b0-rac-jer':12,'b0-rac-prob':13,
  'b0-rad-pot':14,'b0-rad-fac':15,'b0-rad-op':16,
  'b1-poli':5,'b1-pn':6,'b1-fact':7,
  'b2-ec1':5,'b2-ec2':6,'b2-ecpoli':7,'b2-ecrac':8,'b2-ecirr':9,
  'geo1-formas':5,'geo2-angulos':6,'geo3-semejanza':7,'geo4-pitagoras':8,'geo5-trig':9
};

const HOJA_POR_BLOQUE = {
  'b0':'Bloque 0 - Números',
  'b1':'Bloque 1 - Álgebra',
  'b2':'Bloque 2 - Ecuaciones',
  'geo':'Bloque 3 - Geometría'
};

const TOTAL_POR_BLOQUE = {'b0':12,'b1':3,'b2':5,'geo':5};

function inicializarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.entries(CABECERAS).forEach(([nombre, cabs]) => {
    let h = ss.getSheetByName(nombre);
    if (!h) h = ss.insertSheet(nombre);
    if (h.getLastRow() === 0) {
      h.appendRow(cabs);
      h.getRange(1,1,1,cabs.length).setFontWeight('bold').setBackground('#6a1b9a').setFontColor('white');
      h.setFrozenRows(1);
    }
  });
  const def = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (def) ss.deleteSheet(def);
  SpreadsheetApp.getUi().alert('✅ Hojas creadas correctamente.');
}

function doGet(e) { return manejarPeticion(e); }
function doPost(e) { return manejarPeticion(e); }

function manejarPeticion(e) {
  const p = e.parameter || {};
  try {
    if ((p.accion||'completar') === 'completar') {
      registrarModulo(p.nombre||'',p.apellidos||'',p.grupo||'',p.bloque||'',p.modulo||'');
    } else if (p.accion === 'resumen') {
      registrarResumen(p.nombre||'',p.apellidos||'',p.grupo||'',
        parseFloat(p.b0pct||0),parseFloat(p.b1pct||0),parseFloat(p.b2pct||0),parseFloat(p.b3pct||0));
    }
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function registrarModulo(nombre, apellidos, grupo, bloque, moduloId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHoja = HOJA_POR_BLOQUE[bloque];
  if (!nombreHoja) return;
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return;
  const col = MODULO_COLUMNA[moduloId];
  const total = TOTAL_POR_BLOQUE[bloque];
  const clave = (nombre+'|'+apellidos+'|'+grupo).toLowerCase();
  const datos = hoja.getDataRange().getValues();
  let fila = -1;
  for (let i=1;i<datos.length;i++) {
    if ((datos[i][1]+'|'+datos[i][2]+'|'+datos[i][3]).toLowerCase()===clave){fila=i+1;break;}
  }
  if (fila===-1) {
    const nf = new Array(CABECERAS[nombreHoja].length).fill('');
    nf[0]=new Date();nf[1]=nombre;nf[2]=apellidos;nf[3]=grupo;
    if(col)nf[col-1]='✓';
    hoja.appendRow(nf);
    fila=hoja.getLastRow();
  } else {
    if(col)hoja.getRange(fila,col).setValue('✓');
    hoja.getRange(fila,1).setValue(new Date());
  }
  const fd=hoja.getRange(fila,1,1,CABECERAS[nombreHoja].length).getValues()[0];
  let comp=0;
  for(let c=4;c<4+total;c++){if(fd[c]==='✓')comp++;}
  const pct=Math.round(comp/total*100);
  hoja.getRange(fila,CABECERAS[nombreHoja].length-1).setValue(comp);
  hoja.getRange(fila,CABECERAS[nombreHoja].length).setValue(pct+'%');
}

function registrarResumen(nombre,apellidos,grupo,b0pct,b1pct,b2pct,b3pct) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const hoja=ss.getSheetByName('Resumen general');
  if(!hoja)return;
  const clave=(nombre+'|'+apellidos+'|'+grupo).toLowerCase();
  const datos=hoja.getDataRange().getValues();
  let fila=-1;
  for(let i=1;i<datos.length;i++){
    if((datos[i][1]+'|'+datos[i][2]+'|'+datos[i][3]).toLowerCase()===clave){fila=i+1;break;}
  }
  const b0c=Math.round(b0pct/100*12),b1c=Math.round(b1pct/100*3),b2c=Math.round(b2pct/100*5),b3c=Math.round(b3pct/100*5);
  const total=b0c+b1c+b2c+b3c;
  const vals=[new Date(),nombre,apellidos,grupo,b0pct+'%',b1pct+'%',b2pct+'%',b3pct+'%',total,Math.round(total/25*100)+'%'];
  if(fila===-1){hoja.appendRow(vals);}
  else{hoja.getRange(fila,1,1,vals.length).setValues([vals]);}
}
