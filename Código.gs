// ════════════════════════════════════════════════════════════════
//  Bookbuster WMS v2 — Code.gs
//  Backend Google Apps Script
// ═══════════════════════════════════════════════════════════════


const SS = SpreadsheetApp.getActiveSpreadsheet();

const GENEROS = [
  'Ficción Moderna',
  'Novela Policial',
  'Novela Romántica',
  'Fantasía juvenil',
  'Cuentos y Relatos',
  'Clásicos',
  'Ensayos e Historia',
  'Biografías y Crónicas',
  'Autoayuda',
  'Negocios'
];


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 Bookbuster')
    .addItem('📦 Nuevo remito de entrega', 'abrirRemitoEntrega')
    .addItem('🔙 Nuevo remito de retiro',  'abrirRemitoRetiro')
    .addItem('📩 Gestionar encargo',        'abrirGestionarEncargo')
    .addItem('📊 Generar Reporte Mensual',  'abrirReporteMensual')
    .addItem('💰 Deudas de café',           'abrirDeudasCafe')
    .addSeparator()
    .addItem('⚙️ Configurar Sheet', 'setupSheet')
    .addToUi();
}


function abrirDeudasCafe() {
  const html = HtmlService.createHtmlOutputFromFile('deudas-cafe')
    .setWidth(480).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, '💰 Deudas de café');
}


function abrirReporteMensual() {
  const html = HtmlService.createHtmlOutputFromFile('reporte-mensual')
    .setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Reporte Mensual WMS');
}


function abrirGestionarEncargo() {
  const html = HtmlService.createHtmlOutputFromFile('encargo-admin')
    .setWidth(520).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, '📩 Gestionar encargo');
}


function doGet(e) {
  const params   = (e && e.parameter) || {};
  const catParam = (params.catalogo || '').toLowerCase().trim();


  const esCatalogo = !!catParam;
  const slug = esCatalogo ? catParam : (params.cafe || '').toLowerCase().trim();
  const cafe = slug ? getCafeBySlug_(slug) : null;


  if (!cafe) {
    return HtmlService.createHtmlOutput(
      '<body style="font-family:sans-serif;text-align:center;padding:3rem;background:#FAF7F2">' +
      '<h2>⚠️ Café no encontrado</h2>' +
      '<p style="color:#666">Verificá el link con Bookbuster.</p></body>'
    );
  }


  const tmpl = HtmlService.createTemplateFromFile(esCatalogo ? 'catalogo' : 'index');
  tmpl.cafeName = cafe.nombre;
  tmpl.cafeId   = String(cafe.id);


  return tmpl.evaluate()
    .setTitle((esCatalogo ? 'Catálogo · ' : 'Bookbuster · ') + cafe.nombre)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}


function getStockCafe(cafeId) {
  return getRows_('Stock').filter(
    r => String(r.id_cafe) === String(cafeId) &&
         r.ubicacion === 'en_cafe' &&
         !r.id_encargo
  );
}


// Para el catálogo público: solo lo que un visitante debería ver.
// Sin modalidad, costo_firme ni ningún otro dato interno del negocio.
function getStockPublico(cafeId) {
  const generoPorIsbn  = getGeneroCatalogoPorIsbn_();
  const resumenPorIsbn = getResumenCatalogoPorIsbn_();
  return getRows_('Stock')
    .filter(r =>
      String(r.id_cafe) === String(cafeId) &&
      r.ubicacion === 'en_cafe' &&
      !r.id_encargo
    )
    .map(r => ({
      isbn:      String(r.isbn      || ''),
      titulo:    String(r.titulo    || ''),
      autor:     String(r.autor     || ''),
      editorial: String(r.editorial || ''),
      genero:    String(r.genero || generoPorIsbn[normalizarIsbn_(r.isbn)] || ''),
      resumen:   resumenPorIsbn[normalizarIsbn_(r.isbn)] || ''
    }));
}


// Los ejemplares en Stock cargados antes de que existiera la columna
// "genero" (migrarGenero) quedaron con ese campo vacío. Como fallback,
// buscamos el género del mismo libro en el Catálogo por ISBN.
function normalizarIsbn_(isbn) {
  return String(isbn || '').replace(/\D/g, '');
}

function getGeneroCatalogoPorIsbn_() {
  const map = {};
  getRows_('Catálogo').forEach(r => {
    const isbn = normalizarIsbn_(r.isbn);
    if (isbn && r.genero) map[isbn] = String(r.genero);
  });
  return map;
}

// El "resumen" (por qué leerlo) vive solo en Catálogo, no en Stock.
function getResumenCatalogoPorIsbn_() {
  const map = {};
  getRows_('Catálogo').forEach(r => {
    const isbn = normalizarIsbn_(r.isbn);
    if (isbn && r.resumen) map[isbn] = String(r.resumen);
  });
  return map;
}


function getVentasRecientes(cafeId) {
  return getRows_('Ventas')
    .filter(r => String(r.id_cafe) === String(cafeId))
    .sort((a, b) => new Date(b.fecha) - new Date(a.fecha))
    .slice(0, 30);
}


function registrarVenta(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Sistema ocupado, intentá de nuevo.' };
  try {
    const esDonacion    = payload.modalidad === 'donacion';
    const comisionPct   = payload.tipoVenta === 'personal' ? 0 : (payload.modalidad === 'firme' ? 30 : 20);
    const comisionMonto = Math.round(payload.precioFinal * comisionPct / 100);
    const id            = newId_('V');

    const deudaAplica = payload.tipoPago === 'cafe' && payload.modalidad !== 'firme' && payload.tipoVenta !== 'personal';
    const deudaMonto   = deudaAplica ? (payload.precioFinal - comisionMonto) : 0;

    const donante = esDonacion && payload.idEjemplar ? getDonanteFromStock_(payload.idEjemplar) : null;

    appendRow_('Ventas', {
      id,
      fecha:               new Date(),
      id_cafe:             payload.cafeId,
      cafe:                payload.cafeName,
      id_ejemplar:         payload.idEjemplar || '',
      titulo:              payload.titulo,
      autor:               payload.autor || '',
      pvp:                 payload.pvp,
      descuento_pct:       payload.descuentoPct,
      precio_final:        payload.precioFinal,
      tipo_pago:           payload.tipoPago,
      modalidad:           payload.modalidad,
      comision_pct:        comisionPct,
      comision_monto:      comisionMonto,
      tipo_venta:          payload.tipoVenta || 'local',
      estado_comision:     'pendiente',
      fecha_pago_comision: '',
      notas:               payload.notas || '',
      deuda_monto:         deudaMonto,
      deuda_estado:        deudaAplica ? 'pendiente' : '',
      donante_nombre:      donante ? donante.donante_nombre   : '',
      donante_telefono:    donante ? donante.donante_telefono : ''
    });


    if (payload.idEjemplar) {
      setStockUbicacion_(payload.idEjemplar, 'vendido');
    }


    const donanteMonto = esDonacion ? Math.round(payload.precioFinal * 0.6) : 0;

    notificarTelegram_(
      (esDonacion ? '💰 <b>Nueva venta (📦 donación)</b>\n' : '💰 <b>Nueva venta</b>\n') +
      '☕ Café: ' + payload.cafeName + '\n' +
      '📚 ' + payload.titulo + (payload.autor ? ' — ' + payload.autor : '') + '\n' +
      '💵 Precio final: $' + payload.precioFinal +
      (payload.descuentoPct > 0 ? ' (−' + payload.descuentoPct + '%)' : '') + '\n' +
      (payload.tipoVenta === 'personal'
        ? '🏷️ Precio empleados — sin comisión, transferido completo a Bookbuster'
        : '💳 Pago: ' + (payload.tipoPago === 'transfer_bookbuster' ? 'Transferencia a Bookbuster' : (payload.modalidad === 'firme' ? 'Cobró el café — libro firme (sin transferencia pendiente)' : 'Cobró el café — te transfiere el neto')) + '\n' +
          '📊 Comisión café: $' + comisionMonto + ' (' + comisionPct + '%)') +
      (esDonacion
        ? '\n🎁 Donante: ' + (donante && donante.donante_nombre ? donante.donante_nombre : '(sin datos)') +
          (donante && donante.donante_telefono ? ' · ' + donante.donante_telefono : '') +
          '\n🎁 Parte del donante: $' + donanteMonto
        : '')
    );


    return { ok: true, id, comisionMonto, comisionPct };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


function registrarEncargo(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Sistema ocupado, intentá de nuevo.' };
  try {
    const id    = newId_('E');
    const canal = payload.canal === 'web' ? 'web' : 'cafe';
    const tipo  = ['compra', 'venta', 'contacto'].indexOf(payload.tipo) >= 0 ? payload.tipo : 'compra';


    appendRow_('Encargos', {
      id,
      fecha:            new Date(),
      id_cafe:          payload.cafeId,
      cafe:             payload.cafeName,
      descripcion:      payload.descripcion,
      cliente_nombre:   payload.clienteNombre,
      cliente_whatsapp: payload.clienteWhatsapp || '',
      notas:            payload.notas || '',
      estado:           'pendiente',
      canal:            canal,
      tipo:             tipo
    });


    const titulos = {
      compra:   canal === 'web' ? '📱 <b>Encargo desde el catálogo web</b>\n' : '📋 <b>Nuevo encargo</b>\n',
      venta:    '💰 <b>Alguien quiere vender sus libros</b>' + (canal === 'web' ? ' (catálogo web)' : '') + '\n',
      contacto: '💬 <b>Mensaje desde el catálogo web</b>\n'
    };
    const etiquetaMensaje = tipo === 'contacto' ? '📝 Mensaje: ' : '📚 Libro(s): ';

    notificarTelegram_(
      titulos[tipo] +
      '☕ Café: ' + payload.cafeName + '\n' +
      etiquetaMensaje + payload.descripcion + '\n' +
      (payload.clienteNombre   ? '👤 Nombre: '   + payload.clienteNombre   + '\n' : '') +
      (payload.clienteWhatsapp ? '📱 WhatsApp: ' + payload.clienteWhatsapp + '\n' : '') +
      (payload.notas ? '📝 Notas: ' + payload.notas : '')
    );


    return { ok: true, id };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


function setupSheet() {
  const ui = SpreadsheetApp.getUi();


  const SHEETS = {
    'Cafés':    ['id', 'nombre', 'slug', 'activo'],
    'Catálogo': ['isbn', 'titulo', 'autor', 'editorial', 'genero', 'resumen'],
    'Stock': [
      'id_ejemplar', 'isbn', 'titulo', 'autor', 'editorial', 'genero',
      'id_cafe', 'modalidad', 'costo_firme', 'ubicacion',
      'fecha_ingreso', 'id_remito', 'id_encargo',
      'donante_nombre', 'donante_telefono'
    ],
    'Ventas': [
      'id', 'fecha', 'id_cafe', 'cafe', 'id_ejemplar', 'titulo', 'autor',
      'pvp', 'descuento_pct', 'precio_final', 'tipo_pago', 'modalidad',
      'comision_pct', 'comision_monto', 'tipo_venta', 'canal', 'estado_comision',
      'fecha_pago_comision', 'notas',
      'deuda_monto', 'deuda_estado', 'donante_nombre', 'donante_telefono'
    ],
    'Encargos': [
      'id', 'fecha', 'id_cafe', 'cafe', 'descripcion',
      'cliente_nombre', 'cliente_whatsapp', 'notas', 'estado', 'canal', 'tipo', 'prepago'
    ]
  };


  for (const [name, headers] of Object.entries(SHEETS)) {
    let sh = SS.getSheetByName(name);
    if (!sh) {
      sh = SS.insertSheet(name);
    } else {
      sh.clearContents();
      sh.clearFormats();
      sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).removeCheckboxes();
      try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); } catch(e) {}
    }
    const headerRange = sh.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#2C1810');
    headerRange.setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, headers.length, 140);
  }


  setDropdown_('Cafés',    'activo',          ['TRUE', 'FALSE']);
  setDropdown_('Catálogo', 'genero',           GENEROS);
  setDropdown_('Stock',    'modalidad',        ['consignacion', 'firme', 'donacion']);
  setDropdown_('Stock',    'ubicacion',        ['en_cafe', 'vendido', 'retirado']);
  setDropdown_('Stock',    'genero',           GENEROS);
  setDropdown_('Ventas',   'tipo_pago',        ['transfer_bookbuster', 'cafe']);
  setDropdown_('Ventas',   'modalidad',        ['consignacion', 'firme', 'donacion']);
  setDropdown_('Ventas',   'estado_comision',  ['pendiente', 'pagada']);
  setDropdown_('Ventas',   'deuda_estado',     ['pendiente', 'pagada']);
  setDropdown_('Ventas',   'canal',            ['cafe', 'web']);
  setDropdown_('Encargos', 'estado',           ['pendiente', 'gestionado', 'recibido', 'entregado']);
  setDropdown_('Encargos', 'canal',            ['cafe', 'web']);
  setDropdown_('Encargos', 'tipo',             ['compra', 'venta', 'contacto']);


  const cafesSheet = SS.getSheetByName('Cafés');
  cafesSheet.getRange(2, 1, 1, 4).setValues([[1, 'Nombre del café', 'slug-del-cafe', 'TRUE']]);
  cafesSheet.getRange(2, 1, 1, 4).setFontColor('#999999').setFontStyle('italic');


  const defaultSheet = SS.getSheetByName('Hoja 1') || SS.getSheetByName('Sheet1');
  if (defaultSheet && SS.getSheets().length > 1) SS.deleteSheet(defaultSheet);


  ui.alert('✅ Listo', 'El Sheet está configurado.\n\nPróximo paso: completá tus cafés en la pestaña "Cafés"\ny pegá el código de la web app en Apps Script.', ui.ButtonSet.OK);
}


function setDropdown_(sheetName, colName, options) {
  const sh      = SS.getSheetByName(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col     = headers.indexOf(colName);
  if (col < 0) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, col + 1, 999, 1).setDataValidation(rule);
}


function getCafeBySlug_(slug) {
  return getRows_('Cafés').find(
    r => String(r.slug).toLowerCase() === slug && String(r.activo).toLowerCase() === 'true'
  ) || null;
}


function getRows_(sheetName) {
  const sh = SS.getSheetByName(sheetName);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0].map(h => String(h).trim());
  return vals.slice(1)
    .filter(r => r[0] !== '' && r[0] !== null && r[0] !== undefined)
    .map(r => Object.fromEntries(headers.map((h, i) => {
      const v = r[i];
      return [h, v instanceof Date ? v.toISOString() : v];
    })));
}


function appendRow_(sheetName, obj) {
  const sh      = SS.getSheetByName(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  sh.appendRow(headers.map(h => (obj[h] !== undefined ? obj[h] : '')));
}


function setStockUbicacion_(idEjemplar, ubicacion) {
  const sh      = SS.getSheetByName('Stock');
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0].map(h => String(h).trim());
  const idCol   = headers.indexOf('id_ejemplar');
  const ubCol   = headers.indexOf('ubicacion');
  if (idCol < 0 || ubCol < 0) return;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][idCol]) === String(idEjemplar)) {
      sh.getRange(i + 1, ubCol + 1).setValue(ubicacion);
      return;
    }
  }
}


function newId_(prefix) {
  return prefix + '-' + Date.now().toString(36).toUpperCase();
}


function getDonanteFromStock_(idEjemplar) {
  const s = getRows_('Stock').find(r => String(r.id_ejemplar) === String(idEjemplar));
  if (!s) return null;
  return {
    donante_nombre:   s.donante_nombre   || '',
    donante_telefono: s.donante_telefono || ''
  };
}


// ── Deudas de café (café cobró y todavía no te transfirió) ────────

function getDeudasPendientes(cafeId) {
  return getRows_('Ventas')
    .filter(r => r.deuda_estado === 'pendiente' && (!cafeId || String(r.id_cafe) === String(cafeId)))
    .sort((a, b) => new Date(a.fecha) - new Date(b.fecha))
    .map(r => ({
      id:     r.id,
      cafe:   r.cafe,
      id_cafe: r.id_cafe,
      titulo: r.titulo,
      fecha:  r.fecha,
      monto:  Number(r.deuda_monto) || 0
    }));
}


function marcarDeudasPagadas(ids) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'Sistema ocupado, intentá de nuevo.' };
  try {
    const sh      = SS.getSheetByName('Ventas');
    const vals    = sh.getDataRange().getValues();
    const headers = vals[0].map(h => String(h).trim());
    const idCol      = headers.indexOf('id');
    const deudaCol   = headers.indexOf('deuda_estado');
    if (idCol < 0 || deudaCol < 0) return { ok: false, error: 'Faltan columnas de deuda. Corré migrarDonacionYDeuda() primero.' };
    const idSet = new Set((ids || []).map(String));
    let actualizadas = 0;
    for (let i = 1; i < vals.length; i++) {
      if (idSet.has(String(vals[i][idCol]))) {
        const cell = sh.getRange(i + 1, deudaCol + 1);
        cell.clearDataValidations();
        cell.setValue('pagada');
        actualizadas++;
      }
    }
    return { ok: true, actualizadas };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


function getDeudaPendienteCafe(cafeId) {
  return getRows_('Ventas')
    .filter(r => r.deuda_estado === 'pendiente' && String(r.id_cafe) === String(cafeId))
    .reduce((sum, r) => sum + (Number(r.deuda_monto) || 0), 0);
}


// ── Migración segura: agrega columnas/opciones nuevas SIN borrar datos.
// Correr UNA sola vez desde el editor de Apps Script (seleccionar esta
// función en el desplegable de arriba y tocar "Ejecutar").
function migrarDonacionYDeuda() {
  agregarColumnasFaltantes_('Stock',   ['donante_nombre', 'donante_telefono']);
  agregarColumnasFaltantes_('Ventas',  ['deuda_monto', 'deuda_estado', 'donante_nombre', 'donante_telefono']);
  agregarColumnasFaltantes_('Remitos', ['donante_nombre', 'donante_telefono']);

  setDropdown_('Stock',  'modalidad',   ['consignacion', 'firme', 'donacion']);
  setDropdown_('Ventas', 'modalidad',   ['consignacion', 'firme', 'donacion']);
  setDropdown_('Ventas', 'deuda_estado', ['pendiente', 'pagada']);

  SpreadsheetApp.getUi().alert('✅ Listo', 'Se agregaron las columnas y opciones nuevas. No se borró ningún dato existente.', SpreadsheetApp.getUi().ButtonSet.OK);
}


// ── Migración segura: agrega la columna "canal" (cafe / web) a Encargos
// y Ventas, para poder diferenciar a futuro comisiones por origen del
// pedido. Correr UNA sola vez desde el editor de Apps Script.
function migrarCanalEncargo() {
  agregarColumnasFaltantes_('Encargos', ['canal']);
  agregarColumnasFaltantes_('Ventas',   ['canal']);

  setDropdown_('Encargos', 'canal', ['cafe', 'web']);
  setDropdown_('Ventas',   'canal', ['cafe', 'web']);

  SpreadsheetApp.getUi().alert('✅ Listo', 'Se agregó la columna "canal" (cafe / web) a Encargos y Ventas. No se borró ningún dato existente.', SpreadsheetApp.getUi().ButtonSet.OK);
}


// ── Migración segura: agrega la columna "genero" a Catálogo y Stock,
// con la lista fija de géneros como validación. Los libros ya cargados
// quedan con el campo vacío hasta que se completen a mano; los que se
// carguen desde ahora por Remito van a pedirlo obligatoriamente.
// Correr UNA sola vez desde el editor de Apps Script.
function migrarGenero() {
  agregarColumnasFaltantes_('Catálogo', ['genero']);
  agregarColumnasFaltantes_('Stock',    ['genero']);

  setDropdown_('Catálogo', 'genero', GENEROS);
  setDropdown_('Stock',    'genero', GENEROS);

  SpreadsheetApp.getUi().alert('✅ Listo', 'Se agregó la columna "genero" a Catálogo y Stock, con la lista fija como validación. No se borró ningún dato existente.', SpreadsheetApp.getUi().ButtonSet.OK);
}


// ── Migración segura: agrega la columna "resumen" (por qué leerlo, 2-3
// oraciones) a Catálogo. Los libros ya cargados quedan con el campo
// vacío hasta que se completen a mano. Correr UNA sola vez desde el
// editor de Apps Script.
function migrarResumen() {
  agregarColumnasFaltantes_('Catálogo', ['resumen']);

  SpreadsheetApp.getUi().alert('✅ Listo', 'Se agregó la columna "resumen" a Catálogo. No se borró ningún dato existente.', SpreadsheetApp.getUi().ButtonSet.OK);
}


// ── Migración segura: agrega la columna "prepago" a Encargos, para los
// encargos directos (cargados por Bookbuster sin pasar por un café) que
// ya fueron pagados de antemano y no llevan comisión. Correr UNA sola
// vez desde el editor de Apps Script.
function migrarEncargoDirecto() {
  agregarColumnasFaltantes_('Encargos', ['prepago']);

  SpreadsheetApp.getUi().alert('✅ Listo', 'Se agregó la columna "prepago" a Encargos. No se borró ningún dato existente.', SpreadsheetApp.getUi().ButtonSet.OK);
}


// ── Migración segura: agrega la columna "tipo" (compra / venta / contacto)
// a Encargos, para distinguir pedidos de compra de ofertas de venta y
// mensajes de contacto que llegan desde el catálogo web. Los encargos
// existentes (todos "compra" hasta ahora) quedan sin tocar. Correr UNA
// sola vez desde el editor de Apps Script.
function migrarTipoEncargo() {
  agregarColumnasFaltantes_('Encargos', ['tipo']);
  setDropdown_('Encargos', 'tipo', ['compra', 'venta', 'contacto']);

  SpreadsheetApp.getUi().alert('✅ Listo', 'Se agregó la columna "tipo" a Encargos. No se borró ningún dato existente.', SpreadsheetApp.getUi().ButtonSet.OK);
}


function agregarColumnasFaltantes_(sheetName, columnas) {
  const sh = SS.getSheetByName(sheetName);
  if (!sh) return;
  const lastCol = sh.getLastColumn();
  const headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim()) : [];
  columnas.forEach(col => {
    if (headers.indexOf(col) < 0) {
      const nextCol = sh.getLastColumn() + 1;
      const cell = sh.getRange(1, nextCol);
      cell.setValue(col);
      cell.setFontWeight('bold');
      cell.setBackground('#2C1810');
      cell.setFontColor('#FFFFFF');
    }
  });
}


function getPendingEncargos() {
  return getRows_('Encargos').filter(
    r => r.estado === 'pendiente' && (r.tipo || 'compra') === 'compra'
  );
}


function gestionarEncargo(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Sistema ocupado.' };
  try {
    const idEjemplar = newId_('EJ');
    appendRow_('Stock', {
      id_ejemplar:   idEjemplar,
      isbn:          payload.isbn || '',
      titulo:        payload.titulo,
      autor:         payload.autor || '',
      editorial:     payload.editorial || '',
      genero:        payload.genero || '',
      id_cafe:       payload.cafeId,
      modalidad:     'firme',
      costo_firme:   '',
      ubicacion:     'en_cafe',
      fecha_ingreso: new Date(),
      id_remito:     '',
      id_encargo:    payload.encargoId
    });
    setEncargoEstado_(payload.encargoId, 'gestionado');
    return { ok: true };
  } catch(err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


// Encargo cargado directo por Bookbuster (sin pasar por el catálogo de un
// café), para pedidos que llegan a vos por afuera. A diferencia del flujo
// normal (registrarEncargo → gestionarEncargo), acá ya sabés qué libro es,
// así que se crea el Encargo y el ejemplar en Stock en un solo paso, con
// estado "gestionado" directamente.
function crearEncargoDirecto(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Sistema ocupado.' };
  try {
    const id         = newId_('E');
    const idEjemplar = newId_('EJ');
    const prepago    = !!payload.prepago;

    appendRow_('Encargos', {
      id,
      fecha:            new Date(),
      id_cafe:          payload.cafeId,
      cafe:             payload.cafeName,
      descripcion:      payload.titulo,
      cliente_nombre:   payload.clienteNombre,
      cliente_whatsapp: payload.clienteWhatsapp || '',
      notas:            payload.notas || '',
      estado:           'gestionado',
      canal:            'cafe',
      tipo:             'compra',
      prepago:          prepago ? 'TRUE' : 'FALSE'
    });

    appendRow_('Stock', {
      id_ejemplar:   idEjemplar,
      isbn:          payload.isbn      || '',
      titulo:        payload.titulo,
      autor:         payload.autor     || '',
      editorial:     payload.editorial || '',
      genero:        payload.genero    || '',
      id_cafe:       payload.cafeId,
      modalidad:     'firme',
      costo_firme:   '',
      ubicacion:     'en_cafe',
      fecha_ingreso: new Date(),
      id_remito:     '',
      id_encargo:    id
    });

    notificarTelegram_(
      '📋 <b>Nuevo encargo</b>\n' +
      '☕ Café: ' + payload.cafeName + '\n' +
      '📚 Libro(s): ' + payload.titulo + (payload.autor ? ' — ' + payload.autor : '') + '\n' +
      '👤 Nombre: ' + payload.clienteNombre + '\n' +
      (payload.clienteWhatsapp ? '📱 WhatsApp: ' + payload.clienteWhatsapp + '\n' : '') +
      (payload.notas ? '📝 Notas: ' + payload.notas + '\n' : '') +
      (prepago
        ? '💰 Ya está pago'
        : '💰 Se paga en el café al momento de retirar — sin comisión (encargo directo de Bookbuster)')
    );

    return { ok: true, id };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


function getEncargosActivos(cafeId) {
  const encargos = getRows_('Encargos').filter(
    r => String(r.id_cafe) === String(cafeId) &&
         (r.estado === 'gestionado' || r.estado === 'recibido')
  );
  if (!encargos.length) return [];
  const stock = getRows_('Stock');
  return encargos.map(e => {
    const s = stock.find(r => String(r.id_encargo) === String(e.id));
    return Object.assign({}, e, {
      id_ejemplar: s ? s.id_ejemplar : '',
      titulo:      s ? s.titulo      : e.descripcion,
      autor:       s ? s.autor       : ''
    });
  });
}


function confirmarRecepcion(encargoId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Sistema ocupado.' };
  try {
    const encargo = getRows_('Encargos').find(r => String(r.id) === String(encargoId));
    if (!encargo) return { ok: false, error: 'Encargo no encontrado.' };
    setEncargoEstado_(encargoId, 'recibido');
    notificarTelegram_(
      '📦 <b>Encargo recibido</b>\n' +
      '☕ Café: ' + encargo.cafe + '\n' +
      '📚 ' + encargo.descripcion + '\n' +
      '👤 Cliente: ' + encargo.cliente_nombre +
      (encargo.cliente_whatsapp ? '\n📱 WhatsApp: ' + encargo.cliente_whatsapp : '') + '\n' +
      '✅ Ya podés avisarle que puede pasar a retirarlo.'
    );
    return { ok: true };
  } catch(err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


function registrarVentaEncargo(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Sistema ocupado.' };
  try {
    const prepago = !!payload.prepago;
    const encargo = getRows_('Encargos').find(r => String(r.id) === String(payload.encargoId));
    const canal   = encargo && encargo.canal === 'web' ? 'web' : 'cafe';

    // Un encargo directo (creado por Bookbuster, no por el café) nunca
    // lleva comisión, esté pago o no: la columna "prepago" solo se llena
    // en ese flujo (crearEncargoDirecto), así que su sola presencia ya
    // identifica el origen — no hace falta calcular nada.
    const esDirecto = !!(encargo && encargo.prepago !== '' && encargo.prepago !== undefined);

    // Un encargo directo sin prepago no tiene comisión, así que no tiene
    // sentido que el café lo cobre y después tenga que transferir el
    // 100% de vuelta: el cliente paga directo a Bookbuster.
    if (esDirecto && !prepago) payload.tipoPago = 'transfer_bookbuster';

    // TODO: cuando se definan comisiones distintas por canal, ramificar
    // acá según `canal` (hoy 30% para cafe y web por igual).
    const comisionPct   = (prepago || esDirecto) ? 0 : 30;
    const comisionMonto = (prepago || esDirecto) ? 0 : Math.round(payload.precioFinal * comisionPct / 100);
    const id            = newId_('V');

    const deudaAplica = !prepago && payload.tipoPago === 'cafe';
    const deudaMonto   = deudaAplica ? (payload.precioFinal - comisionMonto) : 0;

    appendRow_('Ventas', {
      id,
      fecha:               new Date(),
      id_cafe:             payload.cafeId,
      cafe:                payload.cafeName,
      id_ejemplar:         payload.idEjemplar || '',
      titulo:              payload.titulo,
      autor:               payload.autor || '',
      pvp:                 prepago ? '' : payload.pvp,
      descuento_pct:       prepago ? '' : payload.descuentoPct,
      precio_final:        prepago ? '' : payload.precioFinal,
      tipo_pago:           prepago ? 'prepago' : payload.tipoPago,
      modalidad:           'firme',
      comision_pct:        comisionPct,
      comision_monto:      comisionMonto,
      tipo_venta:          'encargo',
      canal:               canal,
      estado_comision:     'pendiente',
      fecha_pago_comision: '',
      notas:               '',
      deuda_monto:         deudaMonto,
      deuda_estado:        deudaAplica ? 'pendiente' : '',
      donante_nombre:      '',
      donante_telefono:    ''
    });
    if (payload.idEjemplar) setStockUbicacion_(payload.idEjemplar, 'vendido');
    setEncargoEstado_(payload.encargoId, 'entregado');
    notificarTelegram_(
      prepago
        ? '🎉 <b>Encargo entregado</b>\n' +
          '☕ Café: ' + payload.cafeName + '\n' +
          '📚 ' + payload.titulo + (payload.autor ? ' — ' + payload.autor : '') + '\n' +
          '💰 Ya estaba pago — sin comisión para el café'
        : '🎉 <b>Encargo entregado</b>\n' +
          (canal === 'web' && !esDirecto ? '📱 Origen: Catálogo web\n' : '') +
          '☕ Café: ' + payload.cafeName + '\n' +
          '📚 ' + payload.titulo + (payload.autor ? ' — ' + payload.autor : '') + '\n' +
          '💵 Precio final: $' + payload.precioFinal +
          (payload.descuentoPct > 0 ? ' (−' + payload.descuentoPct + '%)' : '') + '\n' +
          '💳 Pago: ' + (payload.tipoPago === 'transfer_bookbuster' ? 'Transferencia a Bookbuster' : 'Cobró el café') + '\n' +
          (esDirecto
            ? '💰 Sin comisión — encargo directo de Bookbuster'
            : '📊 Comisión: $' + comisionMonto + ' (' + comisionPct + '%)')
    );
    return { ok: true, id, comisionMonto, comisionPct };
  } catch(err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


function setEncargoEstado_(encargoId, estado) {
  const sh      = SS.getSheetByName('Encargos');
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0].map(h => String(h).trim());
  const idCol   = headers.indexOf('id');
  const estCol  = headers.indexOf('estado');
  if (idCol < 0 || estCol < 0) return;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][idCol]) === String(encargoId)) {
      const cell = sh.getRange(i + 1, estCol + 1);
      cell.clearDataValidations();
      cell.setValue(estado);
      return;
    }
  }
}


function notificarTelegram_(mensaje) {
  try {
    const token  = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
    const chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
    UrlFetchApp.fetch(
      'https://api.telegram.org/bot' + token + '/sendMessage',
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ chat_id: chatId, text: mensaje, parse_mode: 'HTML' })
      }
    );
  } catch(e) {}
}


function generarPdfReporteMensual(mesStr) { // Formato esperado: "YYYY-MM"
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'Sistema ocupado. Intentá en un minuto.' };

  try {
    const ventas = getRows_('Ventas');
    const stock = getRows_('Stock');
    const cafes = getRows_('Cafés'); // Traemos la tabla de Cafés para cruzar los nombres

    // Función para formatear plata: sin decimales y con punto de miles (ej: $ 15.000)
    const fmt = (n) => {
      let numStr = Math.round(Number(n) || 0).toString();
      return "$ " + numStr.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    };


    // Función para buscar el nombre real del café usando su ID
    const getCafeName = (id) => {
      const c = cafes.find(r => String(r.id) === String(id));
      return c ? String(c.nombre) : String(id);
    };

    // Filtros
    const isTestCafe = (c) => String(c || '').toLowerCase().includes('test');
    const isLowercase = (t) => /[a-z]/.test(t) && !/[A-Z]/.test(t);

    // --- 1. Filtrar Ventas del Mes ---
    let v_pvpTotal = 0;
    let v_precioFinal = 0;
    let v_comisiones = 0;
    let ventasPorCafe = {};
    let librosVendidos = {};

    // Contadores para el TOTAL FINAL de la tabla de ventas
    let totalLocal = 0, totalEncargo = 0, totalLibros = 0;


    ventas.forEach(v => {
      const fecha = new Date(v.fecha);
      const mesVenta = fecha.getFullYear() + '-' + String(fecha.getMonth() + 1).padStart(2, '0');
      const titulo = String(v.titulo || '');

      if (mesVenta === mesStr && !isTestCafe(v.cafe) && !isLowercase(titulo)) {
        const pvp = Number(v.pvp) || 0;
        const pf = Number(v.precio_final) || 0;
        const com = Number(v.comision_monto) || 0;

        // "personal" o vacío cae dentro de "local". Solo "encargo" suma en encargo.
        const tipoVenta = String(v.tipo_venta || '').toLowerCase();

        // Sumas globales
        v_pvpTotal += pvp;
        v_precioFinal += pf;
        v_comisiones += com;

        // Agrupar por café
        if (!ventasPorCafe[v.cafe]) ventasPorCafe[v.cafe] = { local: 0, encargo: 0, libros: 0, cobrado: 0, comision: 0 };

        if (tipoVenta === 'encargo') {
          ventasPorCafe[v.cafe].encargo += 1;
          totalEncargo += 1;
        } else {
          ventasPorCafe[v.cafe].local += 1;
          totalLocal += 1;
        }

        ventasPorCafe[v.cafe].libros += 1;
        totalLibros += 1;
        ventasPorCafe[v.cafe].cobrado += pf;
        ventasPorCafe[v.cafe].comision += com;

        // Agrupar por título
        if (!librosVendidos[titulo]) librosVendidos[titulo] = { autor: String(v.autor || ''), cantidad: 0 };
        librosVendidos[titulo].cantidad += 1;
      }
    });

    const v_descuentos = v_pvpTotal - v_precioFinal;
    const v_netoBookbuster = v_precioFinal - v_comisiones;


    // --- 2. Filtrar y Agrupar Stock ---
    let stockPorCafe = {};
    let totalStock = 0;


    stock.forEach(s => {
      const titulo = String(s.titulo || '');
      const nombreCafe = getCafeName(s.id_cafe); // Cruzamos ID con Nombre

      if (s.ubicacion === 'en_cafe' && !isTestCafe(nombreCafe) && !isLowercase(titulo)) {
        if (!stockPorCafe[nombreCafe]) stockPorCafe[nombreCafe] = 0;
        stockPorCafe[nombreCafe] += 1;
        totalStock += 1;
      }
    });


    // --- 3. Crear el Documento PDF ---
    const doc = DocumentApp.create('Reporte_Bookbuster_' + mesStr);
    const body = doc.getBody();
    body.setMarginTop(30).setMarginBottom(30).setMarginLeft(40).setMarginRight(40);

    // Título
    body.appendParagraph('BOOKBUSTER WMS').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(16).setBold(true);
    body.appendParagraph('Reporte Mensual de Operaciones: ' + mesStr).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);
    body.appendParagraph('');

    // --- BLOQUE 1: CADENA FINANCIERA ---
    body.appendParagraph('1. RESULTADO ECONÓMICO').editAsText().setBold(true).setFontSize(12);
    const tblFinanzas = body.appendTable([
      ['PVP Total (Precio de Tapa)', fmt(v_pvpTotal)],
      ['(-) Descuentos Realizados', fmt(v_descuentos)],
      ['(=) Total Cobrado en Cafés', fmt(v_precioFinal)],
      ['(-) Comisiones Pagadas', fmt(v_comisiones)],
      ['(=) NETO BOOKBUSTER', fmt(v_netoBookbuster)]
    ]);
    tblFinanzas.getRow(4).getCell(0).editAsText().setBold(true);
    tblFinanzas.getRow(4).getCell(1).editAsText().setBold(true);
    body.appendParagraph('');


    // --- BLOQUE 2: RANKING DE CAFÉS ---
    body.appendParagraph('2. RENDIMIENTO POR CAFÉ').editAsText().setBold(true).setFontSize(12);
    const rankingCafes = Object.entries(ventasPorCafe).sort((a, b) => b[1].cobrado - a[1].cobrado);
    if (rankingCafes.length > 0) {
      const dataCafes = [['Café', 'Local', 'Encargo', 'Total', 'Cobrado', 'Comisión']];

      rankingCafes.forEach(c => dataCafes.push([
        c[0],
        String(c[1].local),
        String(c[1].encargo),
        String(c[1].libros),
        fmt(c[1].cobrado),
        fmt(c[1].comision)
      ]));

      // FILA DE TOTALES GENERALES
      dataCafes.push(['TOTALES', String(totalLocal), String(totalEncargo), String(totalLibros), fmt(v_precioFinal), fmt(v_comisiones)]);


      const t = body.appendTable(dataCafes);
      t.getRow(0).editAsText().setBold(true);
      // Poner negrita en la última fila (Totales)
      t.getRow(t.getNumRows() - 1).editAsText().setBold(true);
    } else {
      body.appendParagraph('Sin ventas registradas este mes.');
    }
    body.appendParagraph('');


    // --- BLOQUE 3: STOCK DISPONIBLE ---
    body.appendParagraph('3. STOCK DISPONIBLE POR CAFÉ').editAsText().setBold(true).setFontSize(12);
    if (Object.keys(stockPorCafe).length > 0) {
      const dataStock = [['Café', 'Libros Disponibles']];

      Object.entries(stockPorCafe).sort((a, b) => b[1] - a[1]).forEach(s => dataStock.push([s[0], String(s[1])]));

      // FILA TOTAL STOCK EN LA CALLE
      dataStock.push(['TOTAL EN LA CALLE', String(totalStock)]);


      const ts = body.appendTable(dataStock);
      ts.getRow(0).editAsText().setBold(true);
      // Poner negrita en la última fila
      ts.getRow(ts.getNumRows() - 1).editAsText().setBold(true);
    } else {
      body.appendParagraph('Sin stock registrado.');
    }
    body.appendParagraph('');


    // --- BLOQUE 4: LIBROS MÁS VENDIDOS ---
    body.appendParagraph('4. TÍTULOS VENDIDOS (Agrupados)').editAsText().setBold(true).setFontSize(12);
    const rankingLibros = Object.entries(librosVendidos).sort((a, b) => b[1].cantidad - a[1].cantidad);

    if (rankingLibros.length > 0) {
      const dataLibros = [['Título', 'Autor', 'Cant.']];
      let totalControl = 0; // Para sumar la última fila

      rankingLibros.forEach(l => {
        dataLibros.push([l[0], l[1].autor, String(l[1].cantidad)]);
        totalControl += l[1].cantidad;
      });

      // FILA DE TOTAL PARA CONTROL
      dataLibros.push(['TOTAL', '', String(totalControl)]);


      const tl = body.appendTable(dataLibros);
      tl.getRow(0).editAsText().setBold(true);
      // Poner negrita en la última fila
      tl.getRow(tl.getNumRows() - 1).editAsText().setBold(true);
    } else {
      body.appendParagraph('Sin datos para agrupar.');
    }


    doc.saveAndClose();


    // Guardar como PDF
    const pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
    pdfBlob.setName('Bookbuster_Reporte_' + mesStr + '.pdf');
    const folder = DriveApp.getRootFolder();
    const pdfFile = folder.createFile(pdfBlob);

    // Borrar el temporal de Word
    DriveApp.getFileById(doc.getId()).setTrashed(true);


    return { ok: true, url: pdfFile.getUrl() };

  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}
