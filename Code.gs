// ============================================================
// FX TOOL HOTEL — Google Apps Script
// Solo COMPRA de divisa | Tipo BCE (GOOGLEFINANCE) + Markup %
// Cumplimiento Ley 10/2010 / RD 304/2014
// ============================================================

// ─────────────────────────────────────────────────────────────
// CONSTANTES
// ─────────────────────────────────────────────────────────────

const SHEETS = {
  CONFIG:    'Configuracion',
  TASAS:     'Tasas',
  OPS:       'Operaciones',
  CLIENTES:  'Clientes',
  AML:       'Registro_AML',
  DASHBOARD: 'Dashboard'
};

const AML = {
  // Umbrales si el hotel pierde la exención y pasa a ser sujeto obligado
  UMBRAL_ID:          1000,   // Art. 4 RD 304/2014 — identificación
  UMBRAL_DD_REF:      6000,   // Art. 19.2c RD 304/2014 — DD reforzada (cambio de divisa)
  UMBRAL_DECLARACION: 10000,  // Modelo S-1 Banco de España
  ANOS_CONSERVACION:  10,     // Art. 25 Ley 10/2010
  // Art. 3 RD 304/2014 — límites de la exención del hotel
  EXENCION_CLIENTE_TRIM: 1000,   // máximo por cliente por trimestre natural
  EXENCION_ANUAL:       100000,  // máximo volumen anual total
  AVISO_ANUAL:           80000,  // aviso preventivo (80% del límite)
  ALERTA_ANUAL:          95000   // alerta crítica (95% del límite)
};

const DIVISAS_INFO = {
  'GBP-ENG': { nombre: 'Libra Inglesa (Bank of England)',   iso: 'GBP', simbolo: '£',  gfTicker: 'CURRENCY:GBPEUR' },
  'GBP-SCO': { nombre: 'Libra Escocesa (RBS/BoS/Clyde)',    iso: 'GBP', simbolo: '£',  gfTicker: null },  // referencia GBP-ENG
  'GBP-NIR': { nombre: 'Libra Norirlandesa (BoI/Ulster)',   iso: 'GBP', simbolo: '£',  gfTicker: null },  // referencia GBP-ENG
  'USD':     { nombre: 'Dólar Americano',                   iso: 'USD', simbolo: '$',  gfTicker: 'CURRENCY:USDEUR' },
  'CAD':     { nombre: 'Dólar Canadiense',                  iso: 'CAD', simbolo: 'C$', gfTicker: 'CURRENCY:CADEUR' },
  'AUD':     { nombre: 'Dólar Australiano',                 iso: 'AUD', simbolo: 'A$', gfTicker: 'CURRENCY:AUDEUR' }
};

const ORDEN_DIVISAS = ['GBP-ENG', 'GBP-SCO', 'GBP-NIR', 'USD', 'CAD', 'AUD'];
const TIPOS_DOC = ['DNI', 'Pasaporte', 'NIE', 'Permiso Residencia', 'Carnet Conducir UE', 'Otro'];

// Columnas hoja Tasas (base-1)
const COL = {
  CODIGO:      1,  // A
  NOMBRE:      2,  // B
  ACTIVA:      3,  // C
  TIPO_BCE:    4,  // D — fórmula GOOGLEFINANCE
  MARKUP_PCT:  5,  // E — editable
  TASA_COMPRA: 6,  // F — fórmula =D*(1-E/100) — no editar
  ACT_FECHA:   7,  // G
  ACT_POR:     8   // H
};
const FILA_TASAS = 5; // primera fila de datos (cabecera en fila 4)

// ─────────────────────────────────────────────────────────────
// MENÚ
// ─────────────────────────────────────────────────────────────

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('FX Tool — Recepción')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💱 FX Tool')
    .addItem('📋 Nueva compra de divisa',       'showOperacionDialog')
    .addItem('📐 Ajustar markup de divisas',    'showTasasDialog')
    .addSeparator()
    .addItem('🔒 Ver alertas AML pendientes',   'irARegistroAML')
    .addItem('👤 Ver registro de clientes',     'irAClientes')
    .addItem('📊 Refrescar Dashboard',          'updateDashboard')
    .addSeparator()
    .addItem('⚙️  Gestionar divisas activas',   'irAConfiguracion')
    .addSeparator()
    .addItem('🧾 Reimprimir recibo anterior',    'reimprimir')
    .addSeparator()
    .addItem('🚀 Inicializar / Reparar Tool',   'setupSpreadsheet')
    .addToUi();
}

// ─────────────────────────────────────────────────────────────
// INICIALIZACIÓN
// ─────────────────────────────────────────────────────────────

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(SHEETS).forEach(n => { if (!ss.getSheetByName(n)) ss.insertSheet(n); });

  _setupConfig(ss);
  _setupTasas(ss);
  _setupOperaciones(ss);
  _setupClientes(ss);
  _setupAML(ss);
  _setupDashboard(ss);
  _setupHotelConfig(ss);
  _crearTriggerDiario();

  ss.setActiveSheet(ss.getSheetByName(SHEETS.DASHBOARD));
  SpreadsheetApp.getUi().alert(
    '✅ FX Tool iniciada.\n\n' +
    '• Los tipos BCE se actualizan solos vía GOOGLEFINANCE.\n' +
    '• Ajusta el markup (%) desde el menú o en la pestaña "Tasas".\n' +
    '• Las divisas activas se controlan en "Configuracion".'
  );
}

// ── Configuracion ─────────────────────────────────────────────
function _setupConfig(ss) {
  const s = ss.getSheetByName(SHEETS.CONFIG);
  s.clearContents(); s.clearFormats();

  s.getRange('A1').setValue('⚙️  CONFIGURACIÓN — DIVISAS Y TEMPORADAS').setFontSize(13).setFontWeight('bold');

  s.getRange('A3').setValue('▸ ESTADO DE DIVISAS').setFontWeight('bold').setFontSize(11);
  s.getRange('A4:D4').setValues([['Código','Descripción','¿Activa ahora?','Motivo bloqueo (opcional)']])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');

  const filasDiv = ORDEN_DIVISAS.map(cod => [
    cod, DIVISAS_INFO[cod].nombre,
    cod !== 'GBP-NIR',
    cod === 'GBP-NIR' ? 'Banco no acepta en temporada baja' : ''
  ]);
  s.getRange(5, 1, filasDiv.length, 4).setValues(filasDiv);
  s.getRange(5, 3, filasDiv.length, 1).insertCheckboxes();
  s.getRange(4, 1, filasDiv.length + 1, 4).setBorder(true,true,true,true,true,true);

  const filaT = 5 + filasDiv.length + 2;
  s.getRange(filaT,1).setValue('▸ RESTRICCIONES POR TEMPORADA').setFontWeight('bold').setFontSize(11);
  s.getRange(filaT+1,1).setValue('Las reglas de temporada tienen PRIORIDAD sobre el estado general de arriba.')
    .setFontColor('#d32f2f').setFontStyle('italic');
  s.getRange(filaT+2,1,1,5).setValues([['Código divisa','Nombre del periodo','Fecha inicio','Fecha fin','¿Activa en ese periodo?']])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');

  const ejemplos = [
    ['GBP-SCO','Temporada baja','2025-10-01','2026-03-31',false],
    ['GBP-SCO','Temporada alta','2026-04-01','2026-09-30',true ],
    ['GBP-NIR','Año completo',  '2025-01-01','2026-12-31',false],
    ['','','','',true], ['','','','',true]
  ];
  s.getRange(filaT+3,1,ejemplos.length,5).setValues(ejemplos);
  s.getRange(filaT+3,5,ejemplos.length,1).insertCheckboxes();
  s.getRange(filaT+3,3,ejemplos.length,2).setNumberFormat('yyyy-mm-dd');
  s.getRange(filaT+2,1,ejemplos.length+1,5).setBorder(true,true,true,true,true,true);
  s.autoResizeColumns(1,5);

  // ── Bloque trabajadores ─────────────────────────────────────
  const filaTrab = 5 + filasDiv.length + 2 + ejemplos.length + 6; // approximate — use lastRow
  const filaLastUsed = s.getLastRow() + 2;

  s.getRange(filaLastUsed, 1).setValue('▸ TRABAJADORES AUTORIZADOS')
    .setFontWeight('bold').setFontSize(11);
  s.getRange(filaLastUsed + 1, 1, 1, 2)
    .setValues([['Nombre completo', 'Cargo']])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');

  const trabEjemplos = [
    ['María García López', 'Recepcionista'],
    ['Carlos Martínez', 'Recepcionista'],
    ['Ana Sánchez', 'Jefe de Recepción'],
    ['', ''], ['', ''], ['', ''], ['', ''], ['', '']
  ];
  s.getRange(filaLastUsed + 2, 1, trabEjemplos.length, 2).setValues(trabEjemplos);
  s.getRange(filaLastUsed + 1, 1, trabEjemplos.length + 1, 2)
    .setBorder(true,true,true,true,true,true);

  // ── Bloque nacionalidades ────────────────────────────────────
  const filaNac = filaLastUsed + 2 + trabEjemplos.length + 2;
  s.getRange(filaNac, 1).setValue('▸ NACIONALIDADES (para dropdown)')
    .setFontWeight('bold').setFontSize(11);
  s.getRange(filaNac + 1, 1, 1, 1)
    .setValues([['Nacionalidad']])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');

  const nacs = [
    ['Británica'],['Americana'],['Australiana'],['Canadiense'],
    ['Española'],['Alemana'],['Francesa'],['Italiana'],['Holandesa'],
    ['Belga'],['Sueca'],['Noruega'],['Danesa'],['Irlandesa'],
    ['Portuguesa'],['Suiza'],['Austriaca'],['Polaca'],
    ['Rusa'],['China'],['Japonesa'],['Brasileña'],
    ['Argentina'],['Mexicana'],['Sudafricana'],[''],[''],[''],[''],['']
  ];
  s.getRange(filaNac + 2, 1, nacs.length, 1).setValues(nacs);
  s.getRange(filaNac + 1, 1, nacs.length + 1, 1)
    .setBorder(true,true,true,true,true,true);
  s.getRange(filaNac, 1).setNote('Añade o elimina filas según necesites. La lista se usa en el dropdown de la web app.');
}

// ── Tasas ─────────────────────────────────────────────────────
function _setupTasas(ss) {
  const s = ss.getSheetByName(SHEETS.TASAS);
  s.clearContents(); s.clearFormats();

  s.getRange('A1').setValue('💱  TASAS DE COMPRA — Tipo BCE (GOOGLEFINANCE) + Markup %').setFontSize(13).setFontWeight('bold');
  s.getRange('A2').setValue(
    'El tipo oficial (col. D) se actualiza solo. Edita únicamente el Markup % (col. E). ' +
    'La tasa aplicada (col. F, en azul) se calcula automáticamente: Tipo BCE × (1 − Markup/100).'
  ).setWrap(true).setFontColor('#5f6368').setFontStyle('italic');
  s.setRowHeight(2, 36);

  const headers = ['Código','Descripción','Activa','Tipo oficial BCE (€)','Markup (%)','Tasa compra aplicada (€)','Markup actualizado','Actualizado por'];
  s.getRange(4,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  s.setFrozenRows(4);

  // Markups por defecto — GBP-SCO y GBP-NIR algo más alto por menor liquidez bancaria
  const markupDefecto = { 'GBP-ENG':15,'GBP-SCO':18,'GBP-NIR':22,'USD':15,'CAD':15,'AUD':15 };

  ORDEN_DIVISAS.forEach((cod, i) => {
    const fila = FILA_TASAS + i;
    const info = DIVISAS_INFO[cod];

    s.getRange(fila, COL.CODIGO).setValue(cod);
    s.getRange(fila, COL.NOMBRE).setValue(info.nombre);
    s.getRange(fila, COL.ACTIVA).insertCheckboxes().setValue(cod !== 'GBP-NIR');

    // Col D: GOOGLEFINANCE para divisas con ticker propio;
    //        GBP-SCO y GBP-NIR referencian la celda D de GBP-ENG (misma divisa subyacente)
    if (info.gfTicker) {
      s.getRange(fila, COL.TIPO_BCE).setFormula(`=GOOGLEFINANCE("${info.gfTicker}")`);
    } else {
      s.getRange(fila, COL.TIPO_BCE).setFormula(`=D${FILA_TASAS}`); // referencia a GBP-ENG
    }

    // Col E: markup % — único valor editable por el usuario
    s.getRange(fila, COL.MARKUP_PCT).setValue(markupDefecto[cod]);

    // Col F: tasa aplicada = tipo_BCE * (1 - markup/100) — fórmula protegida
    s.getRange(fila, COL.TASA_COMPRA).setFormula(`=D${fila}*(1-E${fila}/100)`).setBackground('#e8f0fe');

    s.getRange(fila, COL.ACT_FECHA).setValue(new Date());
    s.getRange(fila, COL.ACT_POR).setValue('Setup inicial');
  });

  s.getRange(FILA_TASAS, COL.TIPO_BCE,    ORDEN_DIVISAS.length, 1).setNumberFormat('0.0000');
  s.getRange(FILA_TASAS, COL.MARKUP_PCT,  ORDEN_DIVISAS.length, 1).setNumberFormat('0.00"%"');
  s.getRange(FILA_TASAS, COL.TASA_COMPRA, ORDEN_DIVISAS.length, 1).setNumberFormat('0.0000');
  s.getRange(FILA_TASAS, COL.ACT_FECHA,   ORDEN_DIVISAS.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  s.getRange(4,1,ORDEN_DIVISAS.length+1,headers.length).setBorder(true,true,true,true,true,true);

  const notaFila = FILA_TASAS + ORDEN_DIVISAS.length + 1;
  s.getRange(notaFila,1,1,6).merge()
    .setValue('⚠️  La columna F (azul) contiene fórmulas — no la edites manualmente. Cambia solo el Markup % (col. E). Para bloquear divisas usa la pestaña "Configuracion".')
    .setFontColor('#d32f2f').setFontStyle('italic').setWrap(true);

  [90,280,70,150,85,190,160,130].forEach((w,i) => s.setColumnWidth(i+1,w));
}

// ── Operaciones ───────────────────────────────────────────────
function _setupOperaciones(ss) {
  const s = ss.getSheetByName(SHEETS.OPS);
  s.clearContents(); s.clearFormats();
  const headers = [
    'N° Op','Fecha','Hora','Empleado',
    'Nombre cliente','Tipo doc.','N° documento','Nacionalidad',
    'Habitación/Reserva',
    'Divisa','Importe divisa','Tipo BCE (€)','Markup (%)','Tasa aplicada (€)','€ pagados al cliente',
    'Nivel AML','ID verificado','DD Reforzada','Decl. S-1','Observaciones'
  ];
  s.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  s.setFrozenRows(1);
  [55,90,70,110,180,90,130,120,100,90,110,110,80,120,130,110,90,90,80,200].forEach((w,i)=>s.setColumnWidth(i+1,w));
}

// ── Clientes ──────────────────────────────────────────────────
function _setupClientes(ss) {
  const s = ss.getSheetByName(SHEETS.CLIENTES);
  s.clearContents(); s.clearFormats();
  const h = ['N° documento','Tipo doc.','Nombre completo','Nacionalidad',
    'Total acumulado (€)','N° operaciones','Primera operación','Última operación','Nivel riesgo','Marcado revisión'];
  s.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  s.setFrozenRows(1);
  s.getRange('A2').setValue(`⚠️  DATOS PERSONALES. Conservar ${AML.ANOS_CONSERVACION} años (Art. 25 Ley 10/2010).`)
    .setFontColor('#d32f2f').setFontStyle('italic');
  s.autoResizeColumns(1,h.length);
}

// ── Registro AML ──────────────────────────────────────────────
function _setupAML(ss) {
  const s = ss.getSheetByName(SHEETS.AML);
  s.clearContents(); s.clearFormats();
  s.getRange('A1').setValue('🔒  REGISTRO AML — Ley 10/2010 / RD 304/2014').setFontSize(13).setFontWeight('bold');
  s.getRange('A2').setValue('Revisar y gestionar TODAS las alertas antes del cierre del día.').setFontColor('#d32f2f').setFontStyle('italic');
  const h = ['Fecha/Hora','N° Op','Cliente','Doc. identidad','Importe op. (€)',
    'Tipo alerta','Descripción','Acción requerida','¿Gestionado?','Notas gestión','Gestor'];
  s.getRange(3,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#b71c1c').setFontColor('#fff');
  s.setFrozenRows(3);
  s.setColumnWidth(7,300); s.setColumnWidth(8,300); s.autoResizeColumns(1,6);
}

// ── Dashboard ─────────────────────────────────────────────────
function _setupDashboard(ss) {
  const s = ss.getSheetByName(SHEETS.DASHBOARD);
  s.clearContents(); s.clearFormats();
  s.getRange('A1').setValue('🏨  FX TOOL — DASHBOARD').setFontSize(16).setFontWeight('bold');
  s.getRange('A2').setValue('Última actualización:');
  s.getRange('B2').setFormula('=NOW()').setNumberFormat('dd/mm/yyyy hh:mm');
  s.getRange('A4').setValue('ESTADO DIVISAS Y TASAS HOY').setFontWeight('bold').setFontSize(12);
  s.getRange('A5:E5').setValues([['Divisa','Nombre','Estado','Tipo BCE','Tasa compra (con markup)']])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  s.getRange('G4').setValue('RESUMEN DIARIO').setFontWeight('bold').setFontSize(12);
  s.getRange('G5:H5').setValues([['Métrica','Valor']]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  updateDashboard();
}

// ─────────────────────────────────────────────────────────────
// DIVISAS ACTIVAS
// ─────────────────────────────────────────────────────────────

function isCurrencyActive(codigo) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  const hoy = new Date();
  const data = cfg.getRange(1,1,80,5).getValues();

  // 1. Buscar regla de temporada vigente
  let filaT = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0]==='Código divisa' && data[i][1]==='Nombre del periodo') { filaT=i+1; break; }
  }
  if (filaT >= 0) {
    for (let i = filaT; i < data.length; i++) {
      if (data[i][0]!==codigo || !data[i][2] || !data[i][3]) continue;
      const ini=new Date(data[i][2]), fin=new Date(data[i][3]); fin.setHours(23,59,59);
      if (hoy>=ini && hoy<=fin) return data[i][4]===true;
    }
  }
  // 2. Toggle general
  const divs = cfg.getRange('A5:C10').getValues();
  for (const r of divs) { if (r[0]===codigo) return r[2]===true; }
  return false;
}

function getDivisasEstado() {
  return ORDEN_DIVISAS.map(cod => ({
    codigo: cod, nombre: DIVISAS_INFO[cod].nombre,
    activa: isCurrencyActive(cod), simbolo: DIVISAS_INFO[cod].simbolo
  }));
}

// ─────────────────────────────────────────────────────────────
// TASAS
// ─────────────────────────────────────────────────────────────

function getTasas() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const s    = ss.getSheetByName(SHEETS.TASAS);
  const data = s.getRange(FILA_TASAS, 1, ORDEN_DIVISAS.length, COL.ACT_POR).getValues();
  return data.map(row => ({
    codigo:     row[COL.CODIGO-1],
    nombre:     row[COL.NOMBRE-1],
    activa:     isCurrencyActive(row[COL.CODIGO-1]),
    tipoBCE:    parseFloat(row[COL.TIPO_BCE-1])    || 0,
    markupPct:  parseFloat(row[COL.MARKUP_PCT-1])  || 0,
    tasaCompra: parseFloat(row[COL.TASA_COMPRA-1]) || 0
  }));
}

function getTasaPorCodigo(codigo) {
  return getTasas().find(t => t.codigo === codigo) || null;
}

/** Actualiza solo el markup %. La fórmula de Sheets recalcula la tasa al instante. */
function actualizarMarkupDesdeDialog(nuevos, empleado) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const s    = ss.getSheetByName(SHEETS.TASAS);
  const data = s.getRange(FILA_TASAS, COL.CODIGO, ORDEN_DIVISAS.length, 1).getValues();
  nuevos.forEach(n => {
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === n.codigo) {
        const fila = FILA_TASAS + i;
        s.getRange(fila, COL.MARKUP_PCT).setValue(parseFloat(n.markupPct));
        s.getRange(fila, COL.ACT_FECHA).setValue(new Date());
        s.getRange(fila, COL.ACT_POR).setValue(empleado || '');
        break;
      }
    }
  });
  return { ok: true };
}

// ─────────────────────────────────────────────────────────────
// OPERACIONES
// ─────────────────────────────────────────────────────────────

function registrarOperacion(d) {
  if (!isCurrencyActive(d.divisa))
    return { error: `La divisa ${d.divisa} no está aceptada actualmente. Consulta "Configuracion".` };

  const tasaObj = getTasaPorCodigo(d.divisa);
  if (!tasaObj || tasaObj.tasaCompra <= 0)
    return { error: `Tasa de compra no disponible para ${d.divisa}. Espera a que GOOGLEFINANCE cargue el tipo o verifica la hoja "Tasas".` };

  const importeDivisa = parseFloat(d.importeDivisa);
  const tipoBCE       = tasaObj.tipoBCE;
  const markupPct     = tasaObj.markupPct;
  const tasaAplicada  = tasaObj.tasaCompra;            // = tipoBCE * (1 - markup/100)
  const importeEur    = _round2(importeDivisa * tasaAplicada); // € que el hotel paga al cliente

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const opsSheet = ss.getSheetByName(SHEETS.OPS);
  const lastRow  = opsSheet.getLastRow();
  const opNum    = lastRow <= 1 ? 1 : (parseInt(opsSheet.getRange(lastRow,1).getValue())||0)+1;

  // ── Límite por operación: Art. 3b RD 304/2014 ──────────────────
  // Bloqueo duro: una sola operación no puede pasar de 1.000€
  if (importeEur > AML.EXENCION_CLIENTE_TRIM) {
    return {
      error: `Operación bloqueada: el importe de ${_fmt(importeEur)}€ supera el límite de ${_fmt(AML.EXENCION_CLIENTE_TRIM)}€ por operación (Art. 3b RD 304/2014). No se puede realizar esta operación.`
    };
  }

  // ── Límite acumulado trimestral ──────────────────────────────
  const acumTrimActual = _getAcumuladoClienteTrimestre(d.numDoc);
  if (_round2(acumTrimActual + importeEur) > AML.EXENCION_CLIENTE_TRIM) {
    const disponible = _round2(AML.EXENCION_CLIENTE_TRIM - acumTrimActual);
    return {
      error: `Operación bloqueada: este cliente ya ha cambiado ${_fmt(acumTrimActual)}€ este trimestre. Solo quedan ${_fmt(disponible)}€ disponibles hasta el límite de ${_fmt(AML.EXENCION_CLIENTE_TRIM)}€/trimestre (Art. 3b RD 304/2014).`
    };
  }

  const aml = _checkAML(d, importeEur, opNum);

  const now  = new Date();
  const fila = [
    opNum,
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss'),
    d.empleado,
    d.nombreCliente, d.tipoDoc, d.numDoc, d.nacionalidad,
    d.habitacion || '',
    d.divisa, importeDivisa,
    _round4(tipoBCE), markupPct, _round4(tasaAplicada), importeEur,
    aml.nivel,
    aml.idVerificado  ? 'SÍ' : 'NO',
    aml.ddReforzada   ? 'SÍ' : 'NO',
    aml.declaracionS1 ? '⚠️ REQUERIDA' : 'No',
    d.observaciones || ''
  ];
  opsSheet.appendRow(fila);

  const colores = {NORMAL:null,ID_REQUERIDO:'#fff9c4',DD_REFORZADA:'#ffe0b2',DECLARACION_S1:'#ffcdd2',SOSPECHOSO:'#b71c1c'};
  if (colores[aml.nivel]) opsSheet.getRange(opsSheet.getLastRow(),1,1,fila.length).setBackground(colores[aml.nivel]);

  _actualizarCliente(d, importeEur, opNum, now);
  if (aml.alertas.length > 0) _logAlertas(aml.alertas, opNum, d.nombreCliente, d.numDoc, importeEur, now);

  return { ok:true, opNum, divisa:d.divisa, importeDivisa,
    tipoBCE:_round4(tipoBCE), markupPct:markupPct.toFixed(2),
    tasaAplicada:_round4(tasaAplicada), importeEur, aml };
}

// ─────────────────────────────────────────────────────────────
// AML
// ─────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────
// HELPERS DE FECHA — trimestre natural y año
// ─────────────────────────────────────────────────────────────

/** Devuelve {inicio, fin} del trimestre natural que contiene 'fecha' */
function _getTrimestre(fecha) {
  const mes = fecha.getMonth(); // 0-11
  const anio = fecha.getFullYear();
  const trimestre = Math.floor(mes / 3); // 0=Q1, 1=Q2, 2=Q3, 3=Q4
  const inicio = new Date(anio, trimestre * 3, 1, 0, 0, 0);
  const fin    = new Date(anio, trimestre * 3 + 3, 0, 23, 59, 59);
  return { inicio, fin, label: `Q${trimestre + 1} ${anio}` };
}

/** Parsea una fecha almacenada como 'dd/MM/yyyy' o como objeto Date */
function _parseFechaOp(val) {
  if (val instanceof Date) return val;
  const s = String(val).trim();
  if (!s || s === '') return null;
  // Formato dd/MM/yyyy
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  return new Date(s); // fallback
}

/**
 * Suma importeEur de un cliente en el trimestre natural actual.
 * Lee directamente de Operaciones para máxima precisión.
 */
function _getAcumuladoClienteTrimestre(numDoc) {
  const trim = _getTrimestre(new Date());
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.OPS).getDataRange().getValues();
  let total = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][6]).trim() !== String(numDoc).trim()) continue;
    const f = _parseFechaOp(data[i][1]);
    if (f && f >= trim.inicio && f <= trim.fin) total += parseFloat(data[i][14]) || 0;
  }
  return _round2(total);
}

/**
 * Suma todo el volumen EUR del año en curso.
 * Usado para vigilar el límite de 100.000€ anuales (Art. 3c RD 304/2014).
 */
function _getVolumenAnual(anio) {
  anio = anio || new Date().getFullYear();
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.OPS).getDataRange().getValues();
  let total = 0;
  for (let i = 1; i < data.length; i++) {
    const f = _parseFechaOp(data[i][1]);
    if (f && f.getFullYear() === anio) total += parseFloat(data[i][14]) || 0;
  }
  return _round2(total);
}

// ─────────────────────────────────────────────────────────────
// LÓGICA AML — adaptada a la exención Art. 3 RD 304/2014
// ─────────────────────────────────────────────────────────────

function _checkAML(d, importeEur, opNum) {
  const r = { nivel:'NORMAL', idVerificado:false, ddReforzada:false, declaracionS1:false, alertas:[] };

  // ── 1. Límite por cliente por trimestre natural (Art. 3b) ──
  // Si un cliente supera 1.000€/trimestre, el hotel pierde la exención
  // para ESE cliente y debe identificarle formalmente.
  const acumTrim    = _getAcumuladoClienteTrimestre(d.numDoc);
  const totalTrim   = _round2(acumTrim + importeEur);
  const trim        = _getTrimestre(new Date());

  if (totalTrim > AML.EXENCION_CLIENTE_TRIM) {
    r.idVerificado = true;
    _setNivel(r, 'ID_REQUERIDO');
    r.alertas.push({
      tipo:   'LIMITE_CLIENTE_TRIM',
      desc:   `Cliente supera 1.000€ en ${trim.label} (ant: ${_fmt(acumTrim)}€ + esta: ${_fmt(importeEur)}€ = ${_fmt(totalTrim)}€)`,
      accion: 'Identificación formal obligatoria: verificar y archivar copia del documento de identidad (Art. 3b RD 304/2014).'
    });
  }

  // ── 2. DD reforzada a partir de 6.000€ acumulados en el trimestre ──
  // Art. 19.2c RD 304/2014: umbral específico para cambio de moneda.
  if (totalTrim > AML.UMBRAL_DD_REF) {
    r.ddReforzada = true;
    _setNivel(r, 'DD_REFORZADA');
    r.alertas.push({
      tipo:   'DD_REFORZADA_TRIM',
      desc:   `Acumulado trimestral del cliente supera 6.000€ (${_fmt(totalTrim)}€ en ${trim.label})`,
      accion: 'Diligencia debida reforzada: declaración de origen de fondos. Comunicar al Responsable de Cumplimiento (Art. 19.2c RD 304/2014).'
    });
  }

  // ── 3. Declaración S-1 — operación singular ≥ 10.000€ ──
  if (importeEur >= AML.UMBRAL_DECLARACION) {
    r.declaracionS1 = true;
    _setNivel(r, 'DECLARACION_S1');
    r.alertas.push({
      tipo:   'DECLARACION_S1',
      desc:   `Operación singular de ${_fmt(importeEur)}€ ≥ 10.000€`,
      accion: '⚠️ URGENTE: Modelo S-1 al Banco de España. NO entregar fondos sin autorización del Responsable de Cumplimiento.'
    });
  }

  // ── 4. Vigilancia del límite anual de 100.000€ (Art. 3c) ──
  const volAnual         = _getVolumenAnual();
  const totalAnualConEsta = _round2(volAnual + importeEur);
  const anio             = new Date().getFullYear();

  if (totalAnualConEsta > AML.EXENCION_ANUAL) {
    _setNivel(r, 'DECLARACION_S1');
    r.alertas.push({
      tipo:   'LIMITE_ANUAL_SUPERADO',
      desc:   `🚨 Volumen anual ${anio} supera 100.000€ (${_fmt(totalAnualConEsta)}€). El hotel pierde la exención del Art. 3 RD 304/2014.`,
      accion: '⚠️ URGENTE: Paralizar operaciones de cambio y consultar asesor legal / SEPBLAC inmediatamente. El hotel puede haber pasado a ser sujeto obligado pleno.'
    });
  } else if (totalAnualConEsta > AML.ALERTA_ANUAL) {
    r.alertas.push({
      tipo:   'ALERTA_LIMITE_ANUAL',
      desc:   `Volumen anual ${anio} en ${_fmt(totalAnualConEsta)}€ — quedan ${_fmt(AML.EXENCION_ANUAL - totalAnualConEsta)}€ para el límite de exención`,
      accion: 'Revisar si el volumen restante del año permite seguir operando. Consultar con asesor si se acerca al límite.'
    });
  } else if (totalAnualConEsta > AML.AVISO_ANUAL) {
    r.alertas.push({
      tipo:   'AVISO_LIMITE_ANUAL',
      desc:   `Volumen anual ${anio} en ${_fmt(totalAnualConEsta)}€ — 80% del límite de exención alcanzado`,
      accion: 'Monitorizar el volumen restante del año con atención.'
    });
  }

  // ── 5. Detección de fraccionamiento en 24h ──
  const frac = _checkFraccionadas(d.numDoc, importeEur);
  if (frac.sospechoso) {
    _setNivel(r, 'SOSPECHOSO');
    r.alertas.push({
      tipo:   'OPERACIONES_FRACCIONADAS',
      desc:   `Posibles operaciones fraccionadas: ${frac.numOps} ops en 24h = ${_fmt(frac.total)}€`,
      accion: '🚨 STOP: No realizar sin autorización. Comunicar al Responsable de Cumplimiento. Posible reporte al SEPBLAC.'
    });
  }

  return r;
}

function _setNivel(r, n) {
  const p = ['NORMAL','ID_REQUERIDO','DD_REFORZADA','DECLARACION_S1','SOSPECHOSO'];
  if (p.indexOf(n) > p.indexOf(r.nivel)) r.nivel = n;
}

function _checkFraccionadas(numDoc, importeActual) {
  const data    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.OPS).getDataRange().getValues();
  const hace24h = new Date(Date.now() - 86400000);
  let ops = 0, total = importeActual;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][6]).trim() !== String(numDoc).trim()) continue;
    const f = _parseFechaOp(data[i][1]);
    if (f && f >= hace24h) { ops++; total += parseFloat(data[i][14]) || 0; }
  }
  return { sospechoso: ops >= 2 && total >= AML.EXENCION_CLIENTE_TRIM, numOps: ops + 1, total };
}

function _logAlertas(alertas, opNum, cliente, numDoc, importeEur, fecha) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AML);
  alertas.forEach(a => {
    const fila=[Utilities.formatDate(fecha,Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm:ss'),
      opNum,cliente,numDoc,_round2(importeEur),a.tipo,a.desc,a.accion,false,'',''];
    s.appendRow(fila);
    const lr=s.getLastRow();
    const bg=['DECLARACION_S1_OP','DECLARACION_S1_ACUM','OPERACIONES_FRACCIONADAS'].includes(a.tipo)
      ?'#ffcdd2':a.tipo.includes('DD')?'#ffe0b2':'#fff9c4';
    s.getRange(lr,1,1,fila.length).setBackground(bg);
    if(a.tipo.includes('DECLARACION')||a.tipo==='OPERACIONES_FRACCIONADAS')
      s.getRange(lr,1,1,fila.length).setFontWeight('bold');
  });
}

// ─────────────────────────────────────────────────────────────
// CLIENTES
// ─────────────────────────────────────────────────────────────

function _actualizarCliente(d, importeEur, opNum, fecha) {
  const s    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CLIENTES);
  const data = s.getDataRange().getValues();
  let filaC  = -1;
  for (let i=1;i<data.length;i++) { if(String(data[i][0]).trim()===String(d.numDoc).trim()){filaC=i+1;break;} }
  const fechaFmt = Utilities.formatDate(fecha,Session.getScriptTimeZone(),'dd/MM/yyyy');
  const ant      = filaC===-1?0:parseFloat(data[filaC-1][4])||0;
  const nuevo    = ant+importeEur;
  const nivel    = _nivelRiesgo(nuevo);
  if (filaC===-1) {
    s.appendRow([d.numDoc,d.tipoDoc,d.nombreCliente,d.nacionalidad,_round2(nuevo),1,fechaFmt,fechaFmt,nivel,false]);
    const nr=s.getLastRow(); if(nr>2) s.getRange(nr,10).insertCheckboxes();
  } else {
    const ops=parseInt(s.getRange(filaC,6).getValue())+1;
    s.getRange(filaC,5).setValue(_round2(nuevo));
    s.getRange(filaC,6).setValue(ops);
    s.getRange(filaC,8).setValue(fechaFmt);
    s.getRange(filaC,9).setValue(nivel);
  }
}

function _nivelRiesgo(t) {
  if (t>=AML.UMBRAL_DECLARACION) return '🔴 ALTO';
  if (t>=AML.UMBRAL_DD_REF)      return '🟠 MEDIO-ALTO';
  if (t>=AML.UMBRAL_ID)          return '🟡 MEDIO';
  return '🟢 BAJO';
}

// ─────────────────────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────────────────────

function updateDashboard() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const dash     = ss.getSheetByName(SHEETS.DASHBOARD);
  const opsSheet = ss.getSheetByName(SHEETS.OPS);
  const amlSheet = ss.getSheetByName(SHEETS.AML);
  const hoy      = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy');

  const opsData = opsSheet.getDataRange().getValues();
  let totalHoy=0, numOpsHoy=0;
  for (let i=1;i<opsData.length;i++) { if(opsData[i][1]===hoy){numOpsHoy++;totalHoy+=parseFloat(opsData[i][13])||0;} }

  const amlData = amlSheet.getDataRange().getValues();
  let alertasPend=0;
  for (let i=3;i<amlData.length;i++) { if(amlData[i][8]===false||amlData[i][8]==='') alertasPend++; }

  const tasas = getTasas();
  ORDEN_DIVISAS.forEach((cod,idx) => {
    const t=tasas.find(x=>x.codigo===cod), fila=6+idx;
    dash.getRange(fila,1).setValue(cod);
    dash.getRange(fila,2).setValue(DIVISAS_INFO[cod].nombre);
    dash.getRange(fila,3).setValue(t.activa?'✅  ACTIVA':'🚫  BLOQUEADA')
      .setFontColor(t.activa?'#2e7d32':'#c62828').setFontWeight('bold');
    dash.getRange(fila,4).setValue(t.tipoBCE>0?t.tipoBCE.toFixed(4)+' €':'—');
    dash.getRange(fila,5).setValue(t.tasaCompra>0?t.tasaCompra.toFixed(4)+` € (−${t.markupPct}%)`:'—');
  });

  const res=[['Operaciones hoy',numOpsHoy],['€ pagados hoy',`${_fmt(totalHoy)} €`],
    ['Alertas AML sin gestionar',alertasPend],['Fecha',hoy]];
  dash.getRange(6,7,res.length,2).setValues(res);
  dash.getRange(8,8).setBackground(alertasPend>0?'#ffcdd2':'#c8e6c9')
    .setFontColor(alertasPend>0?'#c62828':'#2e7d32').setFontWeight('bold');
  dash.getRange('B2').setValue(new Date()).setNumberFormat('dd/mm/yyyy hh:mm');
}

// ─────────────────────────────────────────────────────────────
// DIÁLOGOS
// ─────────────────────────────────────────────────────────────

function showOperacionDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('OperacionDialog').setWidth(640).setHeight(700),
    '📋 Nueva Compra de Divisa');
}
function showTasasDialog() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('TasasDialog').setWidth(620).setHeight(540),
    '📐 Ajustar Markup de Divisas');
}
function getDialogData() {
  return { divisas: getDivisasEstado(), tasas: getTasas(), tiposDoc: TIPOS_DOC };
}
function procesarOperacion(datos) {
  try { const r=registrarOperacion(datos); updateDashboard(); return r; }
  catch(e) { return {error:e.message}; }
}
function guardarMarkup(nuevos, empleado) {
  try { actualizarMarkupDesdeDialog(nuevos,empleado); updateDashboard(); return {ok:true}; }
  catch(e) { return {error:e.message}; }
}

// ─────────────────────────────────────────────────────────────
// NAVEGACIÓN
// ─────────────────────────────────────────────────────────────

function irARegistroAML()   { _goTo(SHEETS.AML); }
function irAClientes()      { _goTo(SHEETS.CLIENTES); }
function irAConfiguracion() { _goTo(SHEETS.CONFIG); }
function _goTo(n) {
  const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
  if(s) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(s);
}

// ─────────────────────────────────────────────────────────────
// TRIGGER
// ─────────────────────────────────────────────────────────────

function _crearTriggerDiario() {
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='updateDashboard').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('updateDashboard').timeBased().everyDays(1).atHour(8).create();
}

// ─────────────────────────────────────────────────────────────
// UTILIDADES
// ─────────────────────────────────────────────────────────────

function _round2(n) { return Math.round(parseFloat(n)*100)/100; }
function _round4(n) { return Math.round(parseFloat(n)*10000)/10000; }
function _fmt(n)    { return parseFloat(n).toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2}); }
// ============================================================
// AÑADIR AL FINAL DE Code.gs
// Funciones para el recibo de compra (impresión A5)
// ============================================================

// ─────────────────────────────────────────────────────────────
// CONFIGURACIÓN DEL HOTEL
// (añade estas filas a _setupConfig si quieres que salgan
//  automáticamente; o rellena manualmente en la hoja Config)
// ─────────────────────────────────────────────────────────────

function _setupHotelConfig(ss) {
  // Esta función es llamada por setupSpreadsheet.
  // Si ya ejecutaste el setup, puedes rellenar manualmente las
  // celdas O1:P5 en la hoja Configuracion, o volver a ejecutar
  // setupSpreadsheet (no borra datos existentes).
  const s = ss.getSheetByName(SHEETS.CONFIG);

  // Buscamos la última fila con contenido para no machacar nada
  const ultimaFila = s.getLastRow() + 2;

  s.getRange(ultimaFila, 1).setValue('▸ DATOS DEL HOTEL (para el recibo)')
    .setFontWeight('bold').setFontSize(11);
  s.getRange(ultimaFila + 1, 1, 1, 2).setValues([['Campo', 'Valor']])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');

  const campos = [
    ['Nombre del hotel',  'Hotel Ejemplo'],
    ['Dirección',         'Calle Mayor, 1 — 07001 Palma de Mallorca'],
    ['Teléfono',          '+34 971 000 000'],
    ['Email',             'recepcion@hotelejemplo.com'],
    ['CIF / NIF',         'A12345678'],
    ['Texto pie recibo',  'Gracias por su visita. Thank you for your stay.']
  ];
  s.getRange(ultimaFila + 2, 1, campos.length, 2).setValues(campos);
  s.getRange(ultimaFila + 1, 1, campos.length + 1, 2).setBorder(true,true,true,true,true,true);
}

function getHotelConfig() {
  const s    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
  const data = s.getDataRange().getValues();
  const cfg  = {};
  const campos = ['Nombre del hotel','Dirección','Teléfono','Email','CIF / NIF','Texto pie recibo'];
  for (const row of data) {
    if (campos.includes(row[0])) cfg[row[0]] = row[1];
  }
  return cfg;
}

// ─────────────────────────────────────────────────────────────
// RECIBO — abrir e inyectar datos
// ─────────────────────────────────────────────────────────────

/**
 * Abre el modal del recibo para una operación concreta.
 * Se llama desde OperacionDialog.html tras confirmar la op.
 * Guardamos el opNum en ScriptProperties para que
 * getDatosRecibo() lo recupere (dialogs no admiten parámetros directos).
 */
function prepararYAbrirRecibo(opNum) {
  PropertiesService.getScriptProperties().setProperty('RECIBO_OP_NUM', String(opNum));
  const html = HtmlService.createHtmlOutputFromFile('Recibo')
    .setWidth(460).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, `🧾 Recibo — Operación N° ${opNum}`);
}

/**
 * Llamado desde Recibo.html vía google.script.run.
 * Lee la operación guardada y devuelve todos los datos al HTML.
 */
function getDatosRecibo() {
  const opNumStr = PropertiesService.getScriptProperties().getProperty('RECIBO_OP_NUM');
  if (!opNumStr) return { error: 'No se encontró ninguna operación pendiente de imprimir.' };

  const opNum = parseInt(opNumStr);
  return _buildDatosRecibo(opNum);
}

/**
 * Versión directa: recibe el opNum como parámetro.
 * Llamada desde OperacionDialog.html después de confirmar.
 * Evita el problema de los modales anidados en GAS.
 */
function getDatosReciboDirecto(opNum) {
  return _buildDatosRecibo(parseInt(opNum));
}

/**
 * Desde el menú: pide el N° de operación y abre el recibo.
 * Útil para reimprimir recibos anteriores.
 */
function reimprimir() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt('🧾 Reimprimir recibo', 'Introduce el N° de operación:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const opNum = parseInt(res.getResponseText().trim());
  if (isNaN(opNum) || opNum <= 0) { ui.alert('❌ Número de operación no válido.'); return; }

  const datos = _buildDatosRecibo(opNum);
  if (datos.error) { ui.alert('❌ ' + datos.error); return; }

  prepararYAbrirRecibo(opNum);
}

/**
 * Construye el objeto de datos para el recibo a partir del N° de op.
 */
function _buildDatosRecibo(opNum) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const opsSheet = ss.getSheetByName(SHEETS.OPS);
  const data     = opsSheet.getDataRange().getValues();

  // Buscar la fila con ese N° de operación (col 1 = índice 0)
  let fila = null;
  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][0]) === opNum) { fila = data[i]; break; }
  }
  if (!fila) return { error: `Operación N° ${opNum} no encontrada en el registro.` };

  // Columnas de la hoja Operaciones (base-0):
  // 0=N°Op, 1=Fecha, 2=Hora, 3=Empleado, 4=NombreCliente, 5=TipoDoc,
  // 6=NumDoc, 7=Nacionalidad, 8=Divisa, 9=ImporteDivisa,
  // 10=TipoBCE, 11=MarkupPct, 12=TasaAplicada, 13=ImporteEur,
  // 14=NivelAML, 15=IDVerificado, 16=DDReforzada, 17=DeclS1, 18=Obs

  const codigoDivisa = fila[9];
  const infoDiv      = DIVISAS_INFO[codigoDivisa] || { nombre: codigoDivisa, simbolo: '' };

  const hotel = getHotelConfig();

  return {
    // Hotel
    hotelNombre:    hotel['Nombre del hotel'] || '',
    hotelDireccion: hotel['Dirección']        || '',
    hotelTel:       hotel['Teléfono']         || '',
    hotelEmail:     hotel['Email']            || '',
    hotelCIF:       hotel['CIF / NIF']        || '',
    hotelPie:       hotel['Texto pie recibo'] || '',

    // Operación — fila[1] y fila[2] pueden ser Date objects si Sheets los reinterpreta
    opNum:    fila[0],
    fecha:    fila[1] instanceof Date
                ? Utilities.formatDate(fila[1], Session.getScriptTimeZone(), 'dd/MM/yyyy')
                : String(fila[1]),
    hora:     fila[2] instanceof Date
                ? Utilities.formatDate(fila[2], Session.getScriptTimeZone(), 'HH:mm:ss')
                : String(fila[2]),
    empleado: fila[3],

    // Cliente
    nombreCliente:  fila[4],
    tipoDoc:        fila[5],
    numDoc:         fila[6],
    nacionalidad:   fila[7],

    // Cambio
    divisa:         codigoDivisa,
    divisaNombre:   infoDiv.nombre,
    simbolo:        infoDiv.simbolo,
    habitacion:     String(fila[8] || ''),
    importeDivisa:  parseFloat(fila[10]) || 0,
    tipoBCE:        parseFloat(fila[11]) || 0,
    markupPct:      parseFloat(fila[12]) || 0,
    tasaAplicada:   parseFloat(fila[13]) || 0,
    importeEur:     parseFloat(fila[14]) || 0,

    // AML
    amlNivel:       fila[15] || 'NORMAL'
  };
}

/**
 * Lee trabajadores y nacionalidades de la hoja Configuracion.
 * Trabajadores: col A = nombre, col B = cargo (ambas rellenas).
 * Nacionalidades: col A solo, después de la cabecera "Nacionalidad".
 */
function getConfigLists() {
  const s    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
  const data = s.getDataRange().getValues();

  const trabajadores   = [];
  const nacionalidades = [];

  let inTrab = false, inNac = false;

  for (let i = 0; i < data.length; i++) {
    const colA = String(data[i][0] || '').trim();
    const colB = String(data[i][1] || '').trim();

    // Detectar inicio de sección trabajadores
    if (colA === 'Nombre completo' && colB === 'Cargo') {
      inTrab = true; inNac = false; continue;
    }

    // Detectar inicio de sección nacionalidades
    if (colA === 'Nacionalidad') {
      inNac = true; inTrab = false; continue;
    }

    // Cualquier cabecera ▸ cierra la sección activa
    if (colA.startsWith('▸')) {
      inTrab = false; inNac = false; continue;
    }

    // Trabajadores: necesita nombre (colA) Y cargo (colB) — así no se cuela ninguna otra lista
    if (inTrab && colA !== '' && colB !== '') {
      trabajadores.push(colA);
    }

    // Nacionalidades: solo colA rellena
    if (inNac && colA !== '') {
      nacionalidades.push(colA);
    }
  }

  return { trabajadores, nacionalidades };
}

function getWebAppData() {
  const hotel   = getHotelConfig();
  const divisas = getDivisasEstado();
  const tasas   = getTasas();
  let listas = { trabajadores: [], nacionalidades: [] };
  try { listas = getConfigLists(); } catch(e) { /* listas vacías si Config no está preparada */ }

  return {
    hotelNombre:    hotel['Nombre del hotel'] || '',
    hotelDireccion: hotel['Dirección']        || '',
    hotelTel:       hotel['Teléfono']         || '',
    hotelEmail:     hotel['Email']            || '',
    divisas, tasas, tiposDoc: TIPOS_DOC,
    trabajadores:   listas.trabajadores,
    nacionalidades: listas.nacionalidades
  };
}

function getDashboardData() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const opsSheet = ss.getSheetByName(SHEETS.OPS);
  const amlSheet = ss.getSheetByName(SHEETS.AML);
  const ahora    = new Date();
  const hoy      = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const anio     = ahora.getFullYear();
  const trim     = _getTrimestre(ahora);

  const opsData = opsSheet.getDataRange().getValues();
  let totalHoy = 0, numOpsHoy = 0, totalTrim = 0, numOpsTrim = 0;
  for (let i = 1; i < opsData.length; i++) {
    const f   = _parseFechaOp(opsData[i][1]);
    const eur = parseFloat(opsData[i][14]) || 0;
    if (!f) continue;
    const fechaStr = Utilities.formatDate(f, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    if (fechaStr === hoy)                            { numOpsHoy++; totalHoy += eur; }
    if (f >= trim.inicio && f <= trim.fin)           { numOpsTrim++; totalTrim += eur; }
  }

  const amlData = amlSheet.getDataRange().getValues();
  let alertasPend = 0;
  for (let i = 3; i < amlData.length; i++) {
    if (amlData[i][8] === false || amlData[i][8] === '') alertasPend++;
  }

  const volAnual      = _getVolumenAnual(anio);
  const pctAnual      = Math.min(100, Math.round(volAnual / AML.EXENCION_ANUAL * 100));
  const restante      = _round2(AML.EXENCION_ANUAL - volAnual);
  const estadoAnual   = volAnual > AML.EXENCION_ANUAL ? 'SUPERADO'
                      : volAnual > AML.ALERTA_ANUAL   ? 'CRITICO'
                      : volAnual > AML.AVISO_ANUAL    ? 'AVISO'
                      : 'OK';

  return {
    fecha: hoy, numOpsHoy, totalHoy: _round2(totalHoy), alertasPend,
    trimLabel: trim.label, totalTrim: _round2(totalTrim), numOpsTrim,
    anio, volAnual, pctAnual, restante, estadoAnual,
    limiteAnual: AML.EXENCION_ANUAL,
    tasas: getTasas()
  };
}

function getOperaciones() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const s    = ss.getSheetByName(SHEETS.OPS);
  const data = s.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1)
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => ({
      opNum:         row[0],
      fecha:         row[1] instanceof Date
                       ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'dd/MM/yyyy')
                       : String(row[1] || ''),
      hora:          row[2] instanceof Date
                       ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'HH:mm:ss')
                       : String(row[2] || ''),
      empleado:      String(row[3]  || ''),
      cliente:       String(row[4]  || ''),
      tipoDoc:       String(row[5]  || ''),
      numDoc:        String(row[6]  || ''),
      nacionalidad:  String(row[7]  || ''),
      habitacion:    String(row[8]  || ''),
      divisa:        String(row[9]  || ''),
      importeDivisa: parseFloat(row[10]) || 0,
      tipoBCE:       parseFloat(row[11]) || 0,
      markupPct:     parseFloat(row[12]) || 0,
      tasaAplicada:  parseFloat(row[13]) || 0,
      importeEur:    parseFloat(row[14]) || 0,
      amlNivel:      String(row[15] || 'NORMAL'),
      idVerificado:  String(row[16] || ''),
      ddReforzada:   String(row[17] || ''),
      declS1:        String(row[18] || ''),
      obs:           String(row[19] || '')
    }));
}