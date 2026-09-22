// ═══════════════════════════════════════════════════════════════
//  Bookbuster WMS v2 — Remitos.gs
// ═══════════════════════════════════════════════════════════════


function abrirRemitoEntrega() {
  const tmpl = HtmlService.createTemplateFromFile('remito');
  tmpl.modo  = 'entrega';
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(700).setHeight(620),
    '📦 Nuevo remito de entrega'
  );
}


function abrirRemitoRetiro() {
  const tmpl = HtmlService.createTemplateFromFile('remito');
  tmpl.modo  = 'retiro';
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(700).setHeight(620),
    '🔙 Nuevo remito de retiro'
  );
}


function getCafesLista() {
  return getRows_('Cafés')
    .filter(r => String(r.activo).toLowerCase() === 'true')
    .map(r => ({ id: String(r.id), nombre: String(r.nombre) }));
}


function getCatalogo() {
  return getRows_('Catálogo').map(r => ({
    isbn:      String(r.isbn      || ''),
    titulo:    String(r.titulo    || ''),
    autor:     String(r.autor     || ''),
    editorial: String(r.editorial || ''),
    genero:    String(r.genero    || '')
  }));
}


function getStockParaRetiro(cafeId) {
  return getRows_('Stock')
    .filter(r => String(r.id_cafe) === String(cafeId) && String(r.ubicacion) === 'en_cafe')
    .map(r => ({
      id_ejemplar: String(r.id_ejemplar  || ''),
      isbn:        String(r.isbn         || ''),
      titulo:      String(r.titulo       || ''),
      autor:       String(r.autor        || ''),
      editorial:   String(r.editorial    || ''),
      modalidad:   String(r.modalidad    || ''),
      costo_firme: String(r.costo_firme  || ''),
      id_remito:   String(r.id_remito    || '')
    }));
}


function agregarLibroAlCatalogo(libro) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, error: 'Sistema ocupado.' };
  try {
    if (GENEROS.indexOf(libro.genero) < 0) {
      return { ok: false, error: 'Elegí un género de la lista.' };
    }
    if (libro.isbn) {
      const isbn = String(libro.isbn).replace(/\D/g, '');
      const existe = getRows_('Catálogo').find(
        r => String(r.isbn || '').replace(/\D/g, '') === isbn
      );
      if (existe) return { ok: false, error: 'Ese ISBN ya existe en el Catálogo.' };
    }
    appendRow_('Catálogo', {
      isbn:      String(libro.isbn      || ''),
      titulo:    String(libro.titulo    || ''),
      autor:     String(libro.autor     || ''),
      editorial: String(libro.editorial || ''),
      genero:    String(libro.genero    || ''),
      resumen:   String(libro.resumen   || '')
    });
    return { ok: true, libro: {
      isbn:      String(libro.isbn      || ''),
      titulo:    String(libro.titulo    || ''),
      autor:     String(libro.autor     || ''),
      editorial: String(libro.editorial || ''),
      genero:    String(libro.genero    || ''),
      resumen:   String(libro.resumen   || '')
    }};
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


function crearRemitoEntrega(data) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'Sistema ocupado.' };
  try {
    const id    = newId_('REM');
    const fecha = new Date();


    const donanteNombre   = String(data.donanteNombre   || '');
    const donanteTelefono = String(data.donanteTelefono || '');

    appendRow_('Remitos', {
      id, fecha,
      id_cafe: data.cafeId,
      cafe:    data.cafeName,
      tipo:    'entrega',
      estado:  'pendiente',
      pdf_url: '',
      notas:   data.notas || '',
      donante_nombre:   donanteNombre,
      donante_telefono: donanteTelefono
    });


    for (const item of data.items) {
      const esDonacion = String(item.modalidad || '') === 'donacion';
      appendRow_('Stock', {
        id_ejemplar:   newId_('EJ'),
        isbn:          String(item.isbn       || ''),
        titulo:        String(item.titulo     || ''),
        autor:         String(item.autor      || ''),
        editorial:     String(item.editorial  || ''),
        genero:        String(item.genero     || ''),
        id_cafe:       data.cafeId,
        modalidad:     String(item.modalidad  || 'consignacion'),
        costo_firme:   String(item.costo_firme|| ''),
        ubicacion:     'en_cafe',
        fecha_ingreso: fecha,
        id_remito:     id,
        donante_nombre:   esDonacion ? donanteNombre   : '',
        donante_telefono: esDonacion ? donanteTelefono : ''
      });
    }


    const pdfUrl = generarPdf_(id, fecha, data.cafeName, data.items, 'entrega', data.notas || '');
    actualizarPdfRemito_(id, pdfUrl);


    return { ok: true, id, pdfUrl };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


function crearRemitoRetiro(data) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'Sistema ocupado.' };
  try {
    const id    = newId_('RET');
    const fecha = new Date();


    appendRow_('Remitos', {
      id, fecha,
      id_cafe: data.cafeId,
      cafe:    data.cafeName,
      tipo:    'retiro',
      estado:  'confirmado',
      pdf_url: '',
      notas:   data.notas || ''
    });


    for (const item of data.items) {
      setStockUbicacion_(item.id_ejemplar, 'retirado');
    }


    const pdfUrl = generarPdf_(id, fecha, data.cafeName, data.items, 'retiro', data.notas || '');
    actualizarPdfRemito_(id, pdfUrl);


    return { ok: true, id, pdfUrl };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}


function generarPdf_(id, fecha, cafeName, items, tipo, notas) {
  const tipoLabel = tipo === 'entrega' ? 'REMITO DE ENTREGA' : 'REMITO DE RETIRO';
  const fechaStr  = Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), 'dd/MM/yyyy');


  const doc  = DocumentApp.create('_TEMP_' + id);
  const body = doc.getBody();
  body.setMarginTop(28).setMarginBottom(28).setMarginLeft(36).setMarginRight(36);


  const p1 = body.appendParagraph('BOOKBUSTER');
  p1.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .editAsText().setFontSize(16).setBold(true);


  const p2 = body.appendParagraph(tipoLabel + '   N° ' + id);
  p2.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .editAsText().setFontSize(10).setBold(false);


  const meta = body.appendParagraph('Café: ' + cafeName + '     Fecha: ' + fechaStr + '     Libros: ' + items.length);
  meta.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setFontSize(9);


  body.appendParagraph('');


  const colHeaders = ['#', 'Título', 'Autor', 'ISBN', 'Modalidad'];


  const tableData = [colHeaders];
  items.forEach((item, i) => {
    const mod = item.modalidad === 'firme' ? 'Firme' : (item.modalidad === 'donacion' ? 'Donación' : 'Consignación');
    tableData.push([String(i + 1), item.titulo || '', item.autor || '', item.isbn || '-', mod]);
  });


  const table = body.appendTable(tableData);
  table.setBorderColor('#CCCCCC');


  const hRow = table.getRow(0);
  for (let c = 0; c < colHeaders.length; c++) {
    const cell = hRow.getCell(c);
    cell.setBackgroundColor('#2C1810');
    cell.editAsText().setBold(true).setFontSize(8).setForegroundColor('#FFFFFF');
    cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(4).setPaddingRight(4);
  }


  for (let r = 1; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.editAsText().setFontSize(8);
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(4).setPaddingRight(4);
      if (r % 2 === 0) cell.setBackgroundColor('#FAF7F2');
    }
  }


  body.appendParagraph('');


  if (notas) {
    body.appendParagraph('Notas: ' + notas).editAsText().setFontSize(8).setItalic(true);
    body.appendParagraph('');
  }


  body.appendParagraph('');
  const label1 = tipo === 'entrega' ? 'Entregó (Bookbuster)' : 'Retiró (Bookbuster)';
  const label2 = tipo === 'entrega' ? 'Recibió (Café)'       : 'Entregó (Café)';
  const sigTable = body.appendTable([
    ['\n\n_________________________\n' + label1, '\n\n_________________________\n' + label2]
  ]);
  sigTable.setBorderColor('#FFFFFF');
  sigTable.getRow(0).getCell(0).editAsText().setFontSize(9);
  sigTable.getRow(0).getCell(1).editAsText().setFontSize(9);


  doc.saveAndClose();


  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  pdfBlob.setName(id + '.pdf');


  const folder = DriveApp.getFolderById('1RMgUU5zR0x4po4zmUoFnCUZI_iPJJUeU');
  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  DriveApp.getFileById(doc.getId()).setTrashed(true);


  return pdfFile.getUrl();
}


function actualizarPdfRemito_(id, pdfUrl) {
  const sh      = SS.getSheetByName('Remitos');
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  const pdfCol  = headers.indexOf('pdf_url');
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][idCol]) === String(id)) {
      sh.getRange(i + 1, pdfCol + 1).setValue(pdfUrl);
      return;
    }
  }
}
