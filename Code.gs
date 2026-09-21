/***************************************
 * BOUTIQUE DE LENTES YAEL
 * Núcleo Apps Script - v1
 * Base: estructura v3
 ***************************************/

const HOJAS = {
  CONFIG: 'CONFIGURACION',
  CLIENTES: 'CLIENTES',
  ARMAZONES: 'ARMAZONES',
  MICAS: 'MICAS',
  ACCESORIOS: 'ACCESORIOS',
  VENTAS: 'VENTAS',
  ABONOS: 'ABONOS',
  CATALOGOS: 'CATALOGOS',
  REGLAS: 'REGLAS'
};

const OPTICST_SPREADSHEET_ID = '1qVSL4r5m9Jc9Ny208dX78_HgFRhlccq87eRi8Ekg16vcDdhjyyo4cfWi';

function ss_() {
  if (!OPTICST_SPREADSHEET_ID || OPTICST_SPREADSHEET_ID === 'PENDIENTE_ID_DE_GOOGLE_SHEETS') {
    throw new Error('Falta configurar OPTICST_SPREADSHEET_ID con el ID del Google Sheets.');
  }
  return SpreadsheetApp.openById(OPTICST_SPREADSHEET_ID);
}

function hoja_(nombre) {
  const sh = ss_().getSheetByName(nombre);
  if (!sh) throw new Error(`No existe la hoja "${nombre}".`);
  return sh;
}

function encabezados_(nombreHoja) {
  const sh = hoja_(nombreHoja);
  const lastCol = sh.getLastColumn();
  if (!lastCol) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0];
}

function indiceCampos_(nombreHoja) {
  const headers = encabezados_(nombreHoja);
  const map = {};
  headers.forEach((h, i) => {
    if (h !== '') map[String(h).trim()] = i;
  });
  return map;
}

function filaComoObjeto_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    if (h !== '') obj[String(h).trim()] = row[i];
  });
  return obj;
}

function leerHoja_(nombreHoja) {
  const sh = hoja_(nombreHoja);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  return values
    .filter(row => row.some(v => v !== ''))
    .map(row => filaComoObjeto_(headers, row));
}

function siguienteConsecutivo_(campo) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sh = hoja_(HOJAS.CONFIG);
  const map = indiceCampos_(HOJAS.CONFIG);
  const col = map[campo];
  if (col === undefined) throw new Error(`No existe "${campo}" en CONFIGURACION.`);
  const row = 2;
  const actual = Number(sh.getRange(row, col + 1).getValue()) || 0;
    sh.getRange(row, col + 1).setValue(actual + 1);
    return actual;
  } finally {
    lock.releaseLock();
  }
}

function nuevoId_(tipo) {
  const tipos = {
    cliente: ['Consecutivo_Cliente', 'CLI-', 6],
    venta: ['Consecutivo_Venta', 'VTA-', 6],
    abono: ['Consecutivo_Abono', 'ABO-', 6]
  };
  if (!tipos[tipo]) throw new Error(`Tipo de ID no soportado: ${tipo}`);
  const [campo, prefijo, ancho] = tipos[tipo];
  const numero = siguienteConsecutivo_(campo);
  return prefijo + String(numero).padStart(ancho, '0');
}

function buscarClientes(termino) {
  termino = String(termino || '').trim().toLowerCase();
  if (!termino) return [];

  return leerHoja_(HOJAS.CLIENTES)
    .filter(c => {
      const nombre = String(c.Nombre || '').toLowerCase();
      const telefono = String(c.Telefono || '').toLowerCase();
      const id = String(c.ID_Cliente || '').toLowerCase();
      return nombre.includes(termino) ||
             telefono.includes(termino) ||
             id.includes(termino);
    })
    .slice(0, 30)
    .map(clienteParaApp_);
}

function obtenerCliente(idCliente) {
  const cliente = leerHoja_(HOJAS.CLIENTES)
    .find(c => String(c.ID_Cliente) === String(idCliente));
  return cliente || null;
}

/* =========================
   DATOS SEGUROS PARA LA APP
   ========================= */

function clienteParaApp_(cliente) {
  if (!cliente) return null;
  return {
    ID_Cliente: String(cliente.ID_Cliente || ''),
    Nombre: String(cliente.Nombre || ''),
    Telefono: String(cliente.Telefono || ''),
    Ubicacion: String(cliente.Ubicacion || ''),
    Estado: String(cliente.Estado || '')
  };
}

function ventaParaApp_(venta) {
  if (!venta) return null;

  const cliente = leerHoja_(HOJAS.CLIENTES)
    .find(c => String(c.ID_Cliente || '') === String(venta.ID_Cliente || ''));

  return {
    ID_Venta: String(venta.ID_Venta || ''),
    Fecha: venta.Fecha ? String(venta.Fecha) : '',
    Hora: venta.Hora ? String(venta.Hora) : '',
    ID_Cliente: String(venta.ID_Cliente || ''),
    Nombre_Cliente: cliente ? String(cliente.Nombre || '') : '',
    Telefono_Cliente: cliente ? String(cliente.Telefono || '') : '',
    Armazon_ID: String(venta.Armazon_ID || ''),
    Armazon_Descripcion: String(venta.Armazon_Descripcion || ''),
    Precio_Armazon: Number(venta.Precio_Armazon) || 0,
    Mica_ID: String(venta.Mica_ID || ''),
    Mica_Descripcion: String(venta.Mica_Descripcion || ''),
    Precio_Mica: Number(venta.Precio_Mica) || 0,
    Rebisel_Montaje: String(venta.Rebisel_Montaje || ''),
    Precio_Rebisel: Number(venta.Precio_Rebisel) || 0,
    Accesorio_Descripcion: String(venta.Accesorio_Descripcion || ''),
    Precio_Accesorio: Number(venta.Precio_Accesorio) || 0,
    Total: Number(venta.Total) || 0,
    Pagado: Number(venta.Pagado) || 0,
    Saldo: Number(venta.Saldo) || 0,
    Estado: String(venta.Estado || ''),
    Proxima_Visita: venta.Proxima_Visita ? String(venta.Proxima_Visita) : ''
  };
}

function crearCliente(datos) {
  validarTexto_(datos.nombre, 'Nombre');
  validarTexto_(datos.telefono, 'Teléfono');

  const id = nuevoId_('cliente');
  const sh = hoja_(HOJAS.CLIENTES);
  const map = indiceCampos_(HOJAS.CLIENTES);
  const row = Array(sh.getLastColumn()).fill('');

  row[map.ID_Cliente] = id;
  row[map.Fecha_Registro] = new Date();
  row[map.Nombre] = String(datos.nombre).trim();
  row[map.Telefono] = String(datos.telefono).trim();
  row[map.Ubicacion] = String(datos.ubicacion || '').trim();
  row[map.Proxima_Visita] = datos.proximaVisita ? new Date(datos.proximaVisita) : '';
  row[map.Ultima_Venta] = '';
  row[map.Estado] = 'Activo';

  sh.appendRow(row);

return {
  ID_Cliente: id,
  Nombre: String(datos.nombre).trim(),
  Telefono: String(datos.telefono).trim(),
  Ubicacion: String(datos.ubicacion || '').trim()
};
}

function obtenerArmazones() {
  return leerHoja_(HOJAS.ARMAZONES).filter(x => String(x.Activo).toLowerCase() === 'sí');
}

function obtenerMicas() {
  return leerHoja_(HOJAS.MICAS).filter(x => String(x.Activo).toLowerCase() === 'sí');
}

function obtenerAccesorios() {
  return leerHoja_(HOJAS.ACCESORIOS).filter(x => String(x.Activo).toLowerCase() === 'sí');
}

function obtenerCatalogos() {
  return leerHoja_(HOJAS.CATALOGOS).filter(x => String(x.Activo).toLowerCase() === 'sí');
}

/**
 * Crea una venta.
 * Todos los conceptos son independientes:
 * armazón, mica, rebisel/montaje y accesorio.
 * Los precios siempre son manuales.
 */
function crearVenta(datos) {
  validarTexto_(datos.idCliente, 'Cliente');

  const armazonPrecio = dinero_(datos.precioArmazon);
  const micaPrecio = dinero_(datos.precioMica);
  const rebiselPrecio = dinero_(datos.precioRebisel);
  const accesorioPrecio = dinero_(datos.precioAccesorio);

  const hayArmazon = Boolean(datos.armazonId || datos.armazonDescripcion);
  const hayMica = Boolean(datos.micaId || datos.micaDescripcion);
  const hayRebisel = datos.rebiselMontaje === true || String(datos.rebiselMontaje).toLowerCase() === 'sí';
  const hayAccesorio = Boolean(datos.accesorioDescripcion);

  if (hayArmazon && armazonPrecio === null) throw new Error('El precio del armazón es obligatorio.');
  if (hayMica && micaPrecio === null) throw new Error('El precio de la mica es obligatorio.');
  if (hayRebisel && rebiselPrecio === null) throw new Error('El precio de rebisel/montaje es obligatorio.');
  if (hayAccesorio && accesorioPrecio === null) throw new Error('El precio del accesorio es obligatorio.');

  if (!hayArmazon && !hayMica && !hayRebisel && !hayAccesorio) {
    throw new Error('La venta debe contener al menos un concepto.');
  }

  const total = (armazonPrecio || 0) +
                (micaPrecio || 0) +
                (rebiselPrecio || 0) +
                (accesorioPrecio || 0);

  const idVenta = nuevoId_('venta');
  const fecha = new Date();

  const sh = hoja_(HOJAS.VENTAS);
  const map = indiceCampos_(HOJAS.VENTAS);
  const row = Array(sh.getLastColumn()).fill('');

  row[map.ID_Venta] = idVenta;
  row[map.Fecha] = fecha;
  row[map.Hora] = fecha;
  row[map.ID_Cliente] = datos.idCliente;

  row[map.Armazon_ID] = datos.armazonId || '';
  row[map.Armazon_Descripcion] = datos.armazonDescripcion || '';
  row[map.Precio_Armazon] = armazonPrecio === null ? '' : armazonPrecio;

  row[map.Mica_ID] = datos.micaId || '';
  row[map.Mica_Descripcion] = datos.micaDescripcion || '';
  row[map.Precio_Mica] = micaPrecio === null ? '' : micaPrecio;

  row[map.Rebisel_Montaje] = hayRebisel ? 'Sí' : '';
  row[map.Precio_Rebisel] = rebiselPrecio === null ? '' : rebiselPrecio;

  row[map.Accesorio_Descripcion] = datos.accesorioDescripcion || '';
  row[map.Precio_Accesorio] = accesorioPrecio === null ? '' : accesorioPrecio;

  row[map.Total] = total;
  row[map.Pagado] = 0;
  row[map.Saldo] = total;
  row[map.Estado] = 'PENDIENTE';
  row[map.Proxima_Visita] = datos.proximaVisita ? new Date(datos.proximaVisita) : '';

  // Buscar la primera fila disponible para la venta
const ultimaFilaHoja = sh.getLastRow();
let siguienteFila = 2;

if (ultimaFilaHoja >= 2) {
  const ids = sh
    .getRange(2, map.ID_Venta + 1, ultimaFilaHoja - 1, 1)
    .getDisplayValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === '') {
      siguienteFila = i + 2;
      break;
    }
  }

  // Si no hay espacios vacíos, usar la siguiente fila
  if (siguienteFila === 2 && String(ids[0][0]).trim() !== '') {
    siguienteFila = ultimaFilaHoja + 1;
  }
}

// Guardar la venta
sh.getRange(siguienteFila, 1, 1, row.length).setValues([row]);

  // Actualizar expediente del cliente.
  actualizarClienteTrasVenta_(datos.idCliente, idVenta, datos.proximaVisita);

  return ventaParaApp_(obtenerVenta(idVenta));
}

function obtenerVenta(idVenta) {
  const venta = leerHoja_(HOJAS.VENTAS)
    .find(v => String(v.ID_Venta) === String(idVenta));
  return venta || null;
}

function obtenerVentasCliente(idCliente) {
  return leerHoja_(HOJAS.VENTAS)
    .filter(v => String(v.ID_Cliente) === String(idCliente))
    .map(ventaParaApp_);
}

function buscarVentas(termino) {
  termino = String(termino || '').trim().toLowerCase();
  if (!termino) return [];

  const clientes = leerHoja_(HOJAS.CLIENTES);
  const ventas = leerHoja_(HOJAS.VENTAS);

  return ventas
    .map(v => {
      const cliente = clientes.find(c =>
        String(c.ID_Cliente || '') === String(v.ID_Cliente || '')
      );
      return { venta: v, cliente: cliente || null };
    })
    .filter(item => {
      const v = item.venta;
      const c = item.cliente;
      const campos = [
        v.ID_Venta,
        v.ID_Cliente,
        v.Estado,
        c ? c.Nombre : '',
        c ? c.Telefono : ''
      ];
      return campos.some(valor =>
        String(valor || '').toLowerCase().includes(termino)
      );
    })
    .slice(0, 30)
    .map(item => ventaParaApp_(item.venta));
}

function buscarVenta(idVenta) {
  return ventaParaApp_(obtenerVenta(idVenta));
}

function buscarVentasParaAbono(termino) {
  return buscarVentas(termino);
}

function obtenerAbonosVenta(idVenta) {
  return leerHoja_(HOJAS.ABONOS)
    .filter(a => String(a.ID_Venta || '') === String(idVenta || ''))
    .sort((a, b) => new Date(a.Fecha) - new Date(b.Fecha))
    .map(a => ({
      ID_Abono: String(a.ID_Abono || ''),
      ID_Venta: String(a.ID_Venta || ''),
      Fecha: a.Fecha ? String(a.Fecha) : '',
      Hora: a.Hora ? String(a.Hora) : '',
      Importe: Number(a.Importe) || 0,
      Observaciones: String(a.Observaciones || '')
    }));
}

function registrarAbono(datos) {
  validarTexto_(datos.idVenta, 'Venta');

  const importe = dinero_(datos.importe);
  if (importe === null || importe <= 0) {
    throw new Error('El importe del abono debe ser mayor que cero.');
  }

  const venta = obtenerVenta(datos.idVenta);
  if (!venta) throw new Error('No se encontró la venta.');

  const saldoActual = Number(venta.Saldo) || 0;
  if (importe > saldoActual) {
    throw new Error('El abono no puede ser mayor que el saldo pendiente.');
  }

  const idAbono = nuevoId_('abono');
  const fecha = new Date();
  const sh = hoja_(HOJAS.ABONOS);
  const map = indiceCampos_(HOJAS.ABONOS);
  const row = Array(sh.getLastColumn()).fill('');

  row[map.ID_Abono] = idAbono;
  row[map.ID_Venta] = datos.idVenta;
  row[map.Fecha] = fecha;
  row[map.Hora] = fecha;
  row[map.Importe] = importe;
  row[map.Observaciones] = String(datos.observaciones || '').trim();

  // Buscar la primera fila disponible para el abono
const ultimaFilaHoja = sh.getLastRow();
let siguienteFila = 2;

if (ultimaFilaHoja >= 2) {
  const ids = sh
    .getRange(2, map.ID_Abono + 1, ultimaFilaHoja - 1, 1)
    .getDisplayValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === '') {
      siguienteFila = i + 2;
      break;
    }
  }

  // Si no hay espacios vacíos, usar la siguiente fila
  if (siguienteFila === 2 && String(ids[0][0]).trim() !== '') {
    siguienteFila = ultimaFilaHoja + 1;
  }
}

// Guardar el abono
sh.getRange(siguienteFila, 1, 1, row.length).setValues([row]);

  actualizarTotalesVenta_(datos.idVenta);

  return {
    abono: {
      ID_Abono: idAbono,
      ID_Venta: String(datos.idVenta),
      Fecha: String(fecha),
      Hora: String(fecha),
      Importe: importe,
      Observaciones: String(datos.observaciones || '').trim()
    },
    venta: ventaParaApp_(obtenerVenta(datos.idVenta))
  };
}

function actualizarTotalesVenta_(idVenta) {
  const venta = obtenerVenta(idVenta);
  if (!venta) throw new Error('No se encontró la venta.');

  const totalPagado = leerHoja_(HOJAS.ABONOS)
    .filter(a => String(a.ID_Venta) === String(idVenta))
    .reduce((sum, a) => sum + (Number(a.Importe) || 0), 0);

  const total = Number(venta.Total) || 0;
  const saldo = Math.max(0, total - totalPagado);
  const estado = saldo <= 0 ? 'PAGADA' : 'PENDIENTE';

  const sh = hoja_(HOJAS.VENTAS);
  const map = indiceCampos_(HOJAS.VENTAS);
  const rowNumber = buscarFilaPorValor_(sh, map.ID_Venta + 1, idVenta);

  if (!rowNumber) throw new Error('No se pudo localizar la fila de la venta.');

  sh.getRange(rowNumber, map.Pagado + 1).setValue(totalPagado);
  sh.getRange(rowNumber, map.Saldo + 1).setValue(saldo);
  sh.getRange(rowNumber, map.Estado + 1).setValue(estado);

  return ventaParaApp_(obtenerVenta(idVenta));
}

function buscarFilaPorValor_(sh, col, value) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const values = sh.getRange(2, col, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(value)) return i + 2;
  }
  return null;
}

function actualizarClienteTrasVenta_(idCliente, idVenta, proximaVisita) {
  const sh = hoja_(HOJAS.CLIENTES);
  const map = indiceCampos_(HOJAS.CLIENTES);
  const rowNumber = buscarFilaPorValor_(sh, map.ID_Cliente + 1, idCliente);
  if (!rowNumber) return;

  sh.getRange(rowNumber, map.Ultima_Venta + 1).setValue(idVenta);
  if (proximaVisita) {
    sh.getRange(rowNumber, map.Proxima_Visita + 1).setValue(new Date(proximaVisita));
  }
}

function obtenerResumen(periodo) {
  const ventas = leerHoja_(HOJAS.VENTAS);
  const ahora = new Date();

  const filtradas = ventas.filter(v => pertenecePeriodo_(v.Fecha, periodo, ahora));

  return {
    periodo: periodo || 'todos',
    ventas: filtradas.length,
    total: filtradas.reduce((s, v) => s + (Number(v.Total) || 0), 0),
    cobrado: filtradas.reduce((s, v) => s + (Number(v.Pagado) || 0), 0),
    pendiente: filtradas.reduce((s, v) => s + (Number(v.Saldo) || 0), 0)
  };
}

function pertenecePeriodo_(valorFecha, periodo, ahora) {
  if (!periodo || periodo === 'todos') return true;
  const fecha = new Date(valorFecha);
  if (isNaN(fecha.getTime())) return false;

  if (periodo === 'anio') return fecha.getFullYear() === ahora.getFullYear();
  if (periodo === 'mes') {
    return fecha.getFullYear() === ahora.getFullYear() &&
           fecha.getMonth() === ahora.getMonth();
  }
  if (periodo === 'semana') {
    const inicio = new Date(ahora);
    const dia = inicio.getDay(); // domingo = 0
    inicio.setHours(0,0,0,0);
    inicio.setDate(inicio.getDate() - dia);
    const fin = new Date(inicio);
    fin.setDate(fin.getDate() + 7);
    return fecha >= inicio && fecha < fin;
  }
  return true;
}

function obtenerProximasVisitas() {
  return leerHoja_(HOJAS.CLIENTES)
    .filter(c => c.Proxima_Visita)
    .sort((a,b) => new Date(a.Proxima_Visita) - new Date(b.Proxima_Visita));
}

function dinero_(valor) {
  if (valor === '' || valor === null || valor === undefined) return null;
  const n = Number(String(valor).replace(/[$,\s]/g, ''));
  if (!Number.isFinite(n) || n < 0) throw new Error('Precio o importe inválido.');
  return Math.round(n * 100) / 100;
}

function validarTexto_(valor, campo) {
  if (String(valor || '').trim() === '') {
    throw new Error(`${campo} es obligatorio.`);
  }
}

/**
 * Prueba básica de estructura.
 * Ejecutar una vez desde Apps Script para comprobar que existen todas las hojas.
 */

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Boutique de Lentes Yael');
}

/**
 * Puente para GitHub Pages / PWA -> Apps Script.
 * Mantiene la interfaz google.script.run del frontend V23.
 */
function doPost(e) {
  if (e && e.parameter && e.parameter.opticst_api === '1') {
    return responderApi_(e.parameter.rid, e.parameter.fn, e.parameter.args);
  }
  return HtmlService.createHtmlOutput('OPTICST');
}

function responderApi_(rid, fn, argsJson) {
  const funciones = {
    buscarClientes: buscarClientes,
    crearCliente: crearCliente,
    crearVenta: crearVenta,
    buscarVenta: buscarVenta,
    buscarVentas: buscarVentas,
    buscarVentasParaAbono: buscarVentasParaAbono,
    obtenerCliente: obtenerCliente,
    obtenerVentasCliente: obtenerVentasCliente,
    obtenerAbonosVenta: obtenerAbonosVenta,
    registrarAbono: registrarAbono,
    obtenerResumen: obtenerResumen,
    obtenerCatalogosApp: obtenerCatalogosApp
  };

  try {
    if (!rid) throw new Error('Falta rid.');
    if (!funciones[fn]) throw new Error('Función no permitida: ' + fn);

    let args = [];
    if (argsJson) {
      args = JSON.parse(argsJson);
      if (!Array.isArray(args)) throw new Error('Los argumentos deben ser un arreglo.');
    }

    const resultado = funciones[fn].apply(null, args);
    return responderApiHtml_(rid, true, resultado, '');
  } catch (err) {
    return responderApiHtml_(rid, false, null, String(err && err.message ? err.message : err));
  }
}

function responderApiHtml_(rid, ok, data, error) {
  const payload = JSON.stringify({
    type: 'OPTICST_API_RESPONSE',
    opticst_api: true,
    rid: String(rid || ''),
    ok: !!ok,
    data: data,
    error: error || ''
  });
  return HtmlService
    .createHtmlOutput('<!doctype html><html><body><script>window.top.postMessage(' + payload + ', "*");</script></body></html>')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerCatalogosApp() {
  return {
    armazones: obtenerArmazones(),
    micas: obtenerMicas()
  };
}
