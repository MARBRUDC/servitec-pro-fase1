/**
 * SERVITEC PRO - LIMPIEZA QA
 * Pegar este archivo como archivo adicional en Apps Script SOLO si deseas limpiar duplicados.
 * No reemplaza tu Code.gs principal.
 */
const SERVITEC_QA_SPREADSHEET_ID = '1affuNropDa7C-r2aUvoWMk0_y58YKggWhkjnwfR5U_w';

function servitecCrearBackup() {
  const file = DriveApp.getFileById(SERVITEC_QA_SPREADSHEET_ID);
  const copy = file.makeCopy('BACKUP_' + file.getName() + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss'));
  return copy.getUrl();
}

function servitecDeduplicarBase() {
  const ss = SpreadsheetApp.openById(SERVITEC_QA_SPREADSHEET_ID);
  const reglas = {
    EMPRESAS: function(o){ return o.RUC || o.ID_EMPRESA || o.id; },
    CLIENTES: function(o){ return o.ID_CLIENTE || o.RUC_DNI || o.id; },
    COTIZACIONES: function(o){ return o.NUMERO || o.ID_COTIZACION || o.id; },
    COTIZACION_PARTIDAS: function(o){ return o.ID_PARTIDA || ((o.ID_COTIZACION || '') + '-' + (o.ITEM || '')); },
    ORDENES: function(o){ return o.NUMERO || o.ID_ORDEN || o.id; },
    ORDEN_PARTIDAS: function(o){ return o.ID_PARTIDA || ((o.ID_ORDEN || '') + '-' + (o.ITEM || '')); }
  };

  Object.keys(reglas).forEach(function(nombre){
    const sh = ss.getSheetByName(nombre);
    if (!sh || sh.getLastRow() < 2) return;

    const values = sh.getRange(1,1,sh.getLastRow(),Math.max(3, sh.getLastColumn())).getValues();
    const headers = values[0].map(String);
    const jsonIndex = headers.indexOf('json');
    if (jsonIndex < 0) return;

    const seen = {};
    const rowsToDelete = [];
    for (let r = values.length - 1; r >= 1; r--) {
      const txt = values[r][jsonIndex];
      if (!txt) continue;
      let obj;
      try { obj = JSON.parse(String(txt)); } catch(e) { continue; }
      const key = String(reglas[nombre](obj) || '').trim();
      if (!key) continue;
      if (seen[key]) rowsToDelete.push(r + 1);
      else seen[key] = true;
    }
    rowsToDelete.forEach(function(row){ sh.deleteRow(row); });
  });

  return 'Deduplicación completada';
}

function servitecLimpiarOperacionPruebas() {
  const ss = SpreadsheetApp.openById(SERVITEC_QA_SPREADSHEET_ID);
  const limpiar = [
    'COTIZACIONES',
    'COTIZACION_PARTIDAS',
    'ORDENES',
    'ORDEN_PARTIDAS',
    'EJECUCION_TECNICA',
    'EVIDENCIAS',
    'ACTAS_CONFORMIDAD',
    'INFORMES_TECNICOS',
    'FACTURAS',
    'AUDITORIA',
    'SYNC_LOG'
  ];
  limpiar.forEach(function(nombre){
    const sh = ss.getSheetByName(nombre);
    if (!sh) return;
    if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  });
  return 'Hojas operativas limpiadas. EMPRESAS, CLIENTES, SEDES, USUARIOS, CATALOGOS y CORRELATIVOS se conservan.';
}

function servitecResetCorrelativoCotizacion207() {
  const ss = SpreadsheetApp.openById(SERVITEC_QA_SPREADSHEET_ID);
  const sh = ss.getSheetByName('CORRELATIVOS');
  if (!sh || sh.getLastRow() < 2) return 'No existe CORRELATIVOS';
  const vals = sh.getRange(2,1,sh.getLastRow()-1,3).getValues();
  for (let i=0;i<vals.length;i++) {
    let obj;
    try { obj = JSON.parse(String(vals[i][2] || '{}')); } catch(e) { continue; }
    if ((obj.tipo === 'COT' || obj.TIPO === 'COT') && (obj.sede === '001' || obj.SEDE === '001')) {
      obj.last = 207;
      obj.digits = 5;
      obj.updatedAt = new Date().toISOString();
      sh.getRange(i+2,3).setValue(JSON.stringify(obj));
      return 'Correlativo COT-001 reiniciado a último usado 207. Siguiente: COT-001-00208';
    }
  }
  return 'No se encontró correlativo COT-001';
}
