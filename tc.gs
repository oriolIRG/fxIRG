function forzarRecalculoTasas() {
  // GOOGLEFINANCE no se puede forzar directamente, pero
  // reescribir el valor de una celda adyacente obliga a Sheets
  // a recalcular todas las fórmulas dependientes de ella.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SHEETS.TASAS);
  const ts = new Date();
  s.getRange('I1').setValue('Recalculado: ' + Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
  SpreadsheetApp.flush(); // fuerza escritura inmediata
}

function crearTriggerTasas() {
  // Eliminar triggers previos del mismo tipo
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'forzarRecalculoTasas')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Trigger diario a las 8:00
  ScriptApp.newTrigger('forzarRecalculoTasas')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
    
  SpreadsheetApp.getUi().alert('✅ Trigger creado. Las tasas se refrescarán cada día a las 8:00.');
}