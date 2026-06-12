// ============================================================
// CUERO VIEJO — Apps Script Backend
// ============================================================

var SS = SpreadsheetApp.getActiveSpreadsheet();

// ─── doGet ─────────────────────────────────────────────────
function doGet(e) {
  // Modo API: si viene ?action=... devuelve JSON con CORS
  if (e && e.parameter && e.parameter.action) {
    var result = callAction(e.parameter.action, e.parameter.payload || '{}');
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Modo HTML: sirve la app (compatibilidad con deploy viejo)
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Cuero Viejo')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── doPost ────────────────────────────────────────────────
// Recibe POST con form-encoded: action=xxx&payload={...}
// Usado por la versión PWA hosteada externamente (GitHub Pages)
function doPost(e) {
  var action  = e && e.parameter && e.parameter.action  ? e.parameter.action  : '';
  var payload = e && e.parameter && e.parameter.payload ? e.parameter.payload : '{}';
  if (!action) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'Falta el parámetro action' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var result = callAction(action, payload);
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── callAction: entry point desde google.script.run ───────
function callAction(action, payloadJson) {
  try {
    var payload = {};
    if (payloadJson) {
      try { payload = JSON.parse(payloadJson); } catch(e) { payload = {}; }
    }
    var result = runAction(action, payload);
    return { ok: true, data: result };
  } catch(e) {
    Logger.log('callAction error: ' + action + ' — ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ─── runAction: dispatcher ─────────────────────────────────
function runAction(action, payload) {
  switch(action) {
    case 'ping':            return { status: 'ok', time: new Date().toISOString() };
    case 'getConfig':       return getConfig();
    case 'getCapsulas':     return getCapsulas();
    case 'getCategorias':   return getCategorias();
    case 'getTiposProducto':return getTiposProducto();
    case 'getEstados':      return getEstados();
    case 'getUnidades':     return getUnidades();
    case 'getVehiculos':    return getVehiculos();
    case 'getSocios':       return getSocios();
    case 'getDashboard':    return getDashboard();
    case 'getProductos':    return getProductos();
    case 'crearProducto':   return crearProducto(payload);
    case 'getStock':        return getStock();
    case 'getInsumos':      return getInsumos();
    case 'crearInsumo':     return crearInsumo(payload);
    case 'getCompras':      return getCompras();
    case 'registrarCompra': return registrarCompra(payload);
    case 'getLotesCuero':   return getLotesCuero();
    case 'crearLoteCuero':  return crearLoteCuero(payload);
    case 'getVentas':       return getVentas();
    case 'registrarVenta':  return registrarVenta(payload);
    case 'eliminarVenta': return eliminarVenta(payload);
    case 'getEgresos':      return getEgresos();
    case 'registrarEgreso': return registrarEgreso(payload);
    case 'getViajes':       return getViajes();
    case 'registrarViaje':  return registrarViaje(payload);
    case 'getActivos':      return getActivos();
    case 'crearActivo':     return crearActivo(payload);
    case 'getCostosFijos':  return getCostosFijos();
    case 'registrarCostoFijo': return registrarCostoFijo(payload);
    case 'getProduccion':   return getProduccion();
    case 'registrarProduccion': return registrarProduccion(payload);
    case 'getProduccionAgrupada': return getProduccionAgrupada();
    case 'editarProduccion': return editarProduccion(payload);
    case 'getCotizaciones': return getCotizaciones();
    case 'calcularCosteo':  return calcularCosteo(payload);
    case 'guardarCotizacion': return guardarCotizacion(payload);
    case 'getTareas':       return getTareas();
    case 'crearTarea':      return crearTarea(payload);
    case 'actualizarTarea': return actualizarTarea(payload);
    case 'getSets':         return getSets();
    case 'crearSet':        return crearSet(payload);
    case 'getKPIs':         return getKPIs();
    case 'getMovimientos':  return getMovimientos();
    case 'getFinanzas':     return getFinanzas();
    case 'actualizarConfig': return actualizarConfig(payload);
    case 'crearCapsula': return crearCapsula(payload);
    case 'eliminarCapsula': return eliminarCapsula(payload);
    case 'previsualizarEliminacionCapsula': return previsualizarEliminacionCapsula(payload);
    case 'eliminarCapsulaCompleta': return eliminarCapsulaCompleta(payload);
    case 'editarProducto': return editarProducto(payload);
    case 'getShowroom':            return getShowroom();
    case 'registrarMovimientoShowroom': return registrarMovimientoShowroom(payload);
    case 'getShowroomKPIs':        return getShowroomKPIs(payload);
    case 'registrarProduccionConTalles': return registrarProduccionConTalles(payload);
    case 'registrarVentaConTalle':       return registrarVentaConTalle(payload);
    case 'getVariantes':                 return getVariantes();
    case 'getProductosConVariantes':     return getProductosConVariantes();
    case 'eliminarCostoFijo':            return eliminarCostoFijo(payload);
    case 'getShowroomHistorial':         return getShowroomHistorial();
    case 'getEventosCalendar':           return getEventosCalendar(payload);
    default: throw new Error('Accion no reconocida: ' + action);
  }
}

// ─── HELPERS ───────────────────────────────────────────────
function getSheet(name) {
  var s = SS.getSheetByName(name);
  if (!s) throw new Error('Hoja no encontrada: ' + name);
  return s;
}

function sanitize(v) {
  if (v === null || v === undefined || v === '') return null;
  // Si es string con formato fecha yyyy-mm-dd, devolverlo directo sin convertir
if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v.trim())) {
  return v.trim();
}
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v;
  var s = String(v).trim();
  if (s === '') return null;
  // Convertir numeros con coma decimal (formato argentino)
  if (/^-?\d{1,3}(,\d+)?$/.test(s)) {
    var n = parseFloat(s.replace(',', '.'));
    if (!isNaN(n)) return n;
  }
  return s;
}

function normKey(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
}

function val(row, names) {
  if (!row) return null;

  for (var i = 0; i < names.length; i++) {
    if (row[names[i]] !== undefined && row[names[i]] !== null) return row[names[i]];
  }

  var wanted = names.map(normKey);
  var keys = Object.keys(row);

  for (var k = 0; k < keys.length; k++) {
    var nk = normKey(keys[k]);
    if (wanted.indexOf(nk) >= 0) return row[keys[k]];
  }

  return null;
}

function money(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;

  var s = String(v)
    .replace(/\$/g, '')
    .replace(/\s/g, '')
    .trim();

  if (s.indexOf(',') >= 0) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else {
    s = s.replace(/\./g, '');
  }

  var n = Number(s);
  return isNaN(n) ? 0 : n;
}

function normCapsula(v) {
  var s = String(v || '').trim();
  if (!s) return '';

  s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();

  if (/^\d+$/.test(s)) return 'capsula ' + Number(s);
  if (/^c\s*\d+$/.test(s)) return 'capsula ' + Number(s.replace(/\D/g, ''));
  if (/^capsula\s*\d+$/.test(s)) return 'capsula ' + Number(s.replace(/\D/g, ''));

  return s.replace(/\s+/g, ' ');
}

function capsulaDisplay(valor, capsulas) {
  var n = normCapsula(valor);
  if (!n) return '';

  for (var i = 0; i < capsulas.length; i++) {
    if (normCapsula(capsulas[i]) === n) return capsulas[i];
  }

  return String(valor || '').trim();
}

function getHeaderInfo(sheetName) {
  var s = getSheet(sheetName);
  var data = s.getDataRange().getValues();

  var headerRow = -1;

  for (var r = 0; r < Math.min(data.length, 10); r++) {
    var joined = data[r].map(function(x) { return normKey(x); }).join('|');

    if (
      joined.indexOf('id_') >= 0 ||
      joined.indexOf('fecha') >= 0 ||
      joined.indexOf('monto') >= 0 ||
      joined.indexOf('stock_actual') >= 0 ||
      joined.indexOf('cantidad_vendida') >= 0 ||
      joined.indexOf('clave') >= 0
    ) {
      headerRow = r;
      break;
    }
  }

  if (headerRow < 0) throw new Error('No encontré encabezados en ' + sheetName);

  return {
    sheet: s,
    data: data,
    headerRow: headerRow,
    headers: data[headerRow],
  };
}

function colIndex(headers, aliases) {
  var wanted = aliases.map(normKey);

  for (var i = 0; i < headers.length; i++) {
    if (wanted.indexOf(normKey(headers[i])) >= 0) return i;
  }

  return -1;
}

function colIndexRequired(headers, aliases, sheetName) {
  var c = colIndex(headers, aliases);
  if (c < 0) throw new Error('No encontré la columna "' + aliases[0] + '" en ' + sheetName);
  return c;
}

function sheetToArray(sheetName) {
  var s = getSheet(sheetName);
  var data = s.getDataRange().getValues();
  if (data.length < 2) return [];

  var headerRowIndex = -1;

  // Detecta automáticamente la fila de encabezados.
  // Sirve para hojas con headers en fila 1 o fila 3.
  for (var r = 0; r < Math.min(data.length, 10); r++) {
    var row = data[r].map(function(h) {
      return String(h || '').trim().toLowerCase();
    });

    var joined = row.join('|');

    if (
      joined.indexOf('id_') >= 0 ||
      joined.indexOf('fecha') >= 0 ||
      joined.indexOf('nombre_producto') >= 0 ||
      joined.indexOf('stock_actual') >= 0 ||
      joined.indexOf('monto') >= 0 ||
      joined.indexOf('costo') >= 0 ||
      joined.indexOf('categoria') >= 0 ||
      joined.indexOf('clave') >= 0
    ) {
      headerRowIndex = r;
      break;
    }
  }

  if (headerRowIndex < 0) {
    throw new Error('No encontré encabezados válidos en la hoja: ' + sheetName);
  }

  var headers = data[headerRowIndex].map(function(h) {
    return String(h || '').trim();
  });

  var rows = [];

  for (var i = headerRowIndex + 1; i < data.length; i++) {
    var row = data[i];

    var hasContent = row.some(function(c) {
      return c !== '' && c !== null && c !== undefined;
    });

    if (!hasContent) continue;

    var first = String(row[0] || '').trim().toUpperCase();
    if (first.indexOf('TOTAL') === 0) continue;

    var obj = {};

    for (var k = 0; k < headers.length; k++) {
      if (headers[k]) obj[headers[k]] = sanitize(row[k]);
    }

    rows.push(obj);
  }

  return rows;
}

function getLastId(sheetName, prefix) {
  var s = SS.getSheetByName(sheetName);
  if (!s) return prefix + '001';

  var lastRow = s.getLastRow();
  if (lastRow <= 1) return prefix + '001';

  var data = s.getRange(1, 1, lastRow, 1).getValues();
  var max = 0;

  for (var i = 0; i < data.length; i++) {
    var id = String(data[i][0] || '').trim();
    if (id.indexOf(prefix) === 0) {
      var num = parseInt(id.replace(prefix, '')) || 0;
      if (num > max) max = num;
    }
  }

  // Agregar SpreadsheetApp.flush() para forzar escritura antes de leer
  SpreadsheetApp.flush();

  var next = max + 1;
  var str = String(next);
  while (str.length < 3) str = '0' + str;
  return prefix + str;
}

function appendTo(sheetName, row) {
  var s = SS.getSheetByName(sheetName);
  if (!s) {
    Logger.log('appendTo: hoja no encontrada: ' + sheetName);
    throw new Error('Hoja no encontrada: ' + sheetName);
  }

  var lastCol = s.getLastColumn();
  var lastRow = s.getLastRow();

  Logger.log('appendTo ' + sheetName + ': lastCol=' + lastCol + ' lastRow=' + lastRow + ' rowLen=' + row.length);

  // Si la hoja está vacía (0 columnas), el problema es que no tiene headers
  if (lastCol === 0) {
    throw new Error('La hoja ' + sheetName + ' está vacía, sin encabezados.');
  }

  // Si el array tiene más columnas que la hoja, truncar
  if (row.length > lastCol) {
    Logger.log('Truncando row de ' + row.length + ' a ' + lastCol + ' columnas');
    row = row.slice(0, lastCol);
  }

  s.appendRow(row);
}

function findRow(sheetName, colIndex, value) {
  var s = getSheet(sheetName);
  var data = s.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]).trim() === String(value).trim()) return i + 1;
  }
  return -1;
}

// ─── CONFIG ────────────────────────────────────────────────
function getConfig() {
  var rows = sheetToArray('CONFIG');
  var cfg = {};

  rows.forEach(function(r) {
    var clave = r.clave || r.Clave || r.CLAVE;
    var valor = r.valor || r.Valor || r.VALOR;
    var activo = r.activo;

    if (!clave) return;
    if (activo === false || String(activo).toUpperCase() === 'FALSE') return;

    cfg[String(clave).trim()] = valor;
  });

  return cfg;
}

function actualizarConfig(p) {
  var s = getSheet('CONFIG');
  var data = s.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var claveCol = headers.indexOf('clave');
  var valorCol = headers.indexOf('valor');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][claveCol]).trim() === p.clave) {
      s.getRange(i + 1, valorCol + 1).setValue(p.valor);
      return true;
    }
  }
  return false;
}

function getCapsulas() {
  var cfg = getConfig();
  var raw = String(cfg['capsulas_activas'] || '');
  return raw.split(';').map(function(c) { return c.trim(); }).filter(Boolean);
}

function getSocios() {
  var cfg = getConfig();
  var raw = String(cfg['socios'] || '');
  return raw.split(';').map(function(s) { return s.trim(); }).filter(Boolean);
}

function getVehiculos() {
  return sheetToArray('VEHICULOS');
}

function getCategorias() {
  try {
    var rows = sheetToArray('INSUMOS_MAESTRO');
    var seen = {}, cats = [];
    rows.forEach(function(r) {
      if (r.categoria && !seen[r.categoria]) { seen[r.categoria] = true; cats.push(r.categoria); }
    });
    return cats.length ? cats : ['Cuero','Avios / herrajes','Hilos / cordones','Packaging','Etiquetas','Consumibles de produccion','Otros'];
  } catch(e) {
    return ['Cuero','Avios / herrajes','Hilos / cordones','Packaging','Etiquetas','Consumibles de produccion','Otros'];
  }
}

function getTiposProducto() {
  try {
    var rows = sheetToArray('PRODUCTOS_MAESTRO');
    var seen = {}, tipos = [];
    rows.forEach(function(r) {
      if (r.categoria && !seen[r.categoria]) { seen[r.categoria] = true; tipos.push(r.categoria); }
    });
    return tipos.length ? tipos : ['Indumentaria','Accesorios','Conjunto','Set Personalizado'];
  } catch(e) {
    return ['Indumentaria','Accesorios','Conjunto','Set Personalizado'];
  }
}

function getEstados() {
  return {
    stock: ['Sin stock','Stock bajo','OK'],
    produccion: ['En proceso','Completada','Pausada'],
    tarea: ['No iniciada','En curso','Completada'],
    set: ['Consulta','En proceso','Entregado','Cerrado']
  };
}

function getUnidades() {
  return ['u','m','m2','cm2','cm','litro','kg','g','rollo','lote','par'];
}

// ─── DASHBOARD ─────────────────────────────────────────────
function getDashboard() {
  var ventas = getVentas();
  var egresos = sheetToArray('EGRESOS');
  var stock = sheetToArray('STOCK_PRODUCTOS');
  var activos = sheetToArray('ACTIVOS_AMORTIZABLES');

  var capsulas = getCapsulas();

  var totalVentas = 0;
  var totalEgresos = 0;
  var amortMensual = 0;

  var ventasPorCapsula = {};
  capsulas.forEach(function(c) { ventasPorCapsula[c] = 0; });

  ventas.forEach(function(v) {
    var monto = money(val(v, ['monto', 'Monto', 'importe', 'Importe', 'total', 'Total']));
    var capRaw = val(v, ['capsula', 'Cápsula', 'Capsula', 'capsula_asociada', 'Capsula asociada']);
    var cap = capsulaDisplay(capRaw, capsulas) || 'Sin cápsula';

    totalVentas += monto;

    if (ventasPorCapsula[cap] === undefined) ventasPorCapsula[cap] = 0;
    ventasPorCapsula[cap] += monto;
  });

  egresos.forEach(function(e) {
    totalEgresos += money(val(e, ['costo', 'Costo', 'importe', 'Importe', 'monto', 'Monto']));
  });

  activos.forEach(function(a) {
    amortMensual += money(val(a, ['amortizacion_mensual', 'Amortizacion mensual', 'Amort. mensual']));
  });

  var stockSinStock = 0;
  var stockBajo = 0;
  var ingresoPotencial = 0;

  stock.forEach(function(p) {
    var estado = String(val(p, ['estado_stock', 'Estado stock']) || '');
    var stockActual = money(val(p, ['stock_actual', 'Stock actual', 'stock']));
    var precio = money(val(p, ['precio_venta_actual', 'Precio venta actual', 'precio']));

    if (estado === 'Sin stock') stockSinStock++;
    if (estado === 'Stock bajo') stockBajo++;

    ingresoPotencial += stockActual * precio;
  });

  var pendientePago = 0;
  ventas.forEach(function(v) {
    var pagado = val(v, ['pagado', 'Pagado']);
    var monto = money(val(v, ['monto', 'Monto', 'importe', 'Importe', 'total', 'Total']));

    if (pagado === 'No' || pagado === false || String(pagado).toLowerCase() === 'pendiente') {
      pendientePago += monto;
    }
  });

  var vendidosPorProd = {};
  ventas.forEach(function(v) {
    var k = String(val(v, ['nombre_producto', 'Producto', 'producto', 'Compra']) || 'Sin producto');
    vendidosPorProd[k] = (vendidosPorProd[k] || 0) + 1;
  });

  var topProductos = Object.keys(vendidosPorProd).map(function(k) {
    return { nombre: k, cant: vendidosPorProd[k] };
  });

  topProductos.sort(function(a, b) { return b.cant - a.cant; });
  topProductos = topProductos.slice(0, 5);

  var ventasPorMes = {};
  ventas.forEach(function(v) {
    var fecha = val(v, ['fecha', 'Fecha']);
    if (!fecha) return;

    var d = new Date(fecha);
    if (isNaN(d)) return;

    var k = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
    ventasPorMes[k] = (ventasPorMes[k] || 0) + money(val(v, ['monto', 'Monto', 'importe', 'Importe', 'total', 'Total']));
  });

  var ventasMes = Object.keys(ventasPorMes).sort().map(function(k) {
    return { mes: k, total: ventasPorMes[k] };
  });

  return {
    totalVentas: totalVentas,
    totalEgresos: totalEgresos,
    beneficioNeto: totalVentas - totalEgresos,
    amortizacionMensual: amortMensual,
    stockSinStock: stockSinStock,
    stockBajo: stockBajo,
    ingresoPotencial: ingresoPotencial,
    pendientePago: pendientePago,
    ventasPorCapsula: ventasPorCapsula,
    topProductos: topProductos,
    ventasMes: ventasMes,
    totalProductos: stock.length,
    totalVentasCount: ventas.length
  };
}

// ─── PRODUCTOS ─────────────────────────────────────────────
function getProductos() { return sheetToArray('PRODUCTOS_MAESTRO'); }

function crearProducto(p) {
  var s = getSheet('PRODUCTOS_MAESTRO');
  var id = getLastId('PRODUCTOS_MAESTRO', 'PRD-');

  s.appendRow([
    id,
    p.id_producto_base || id,
    p.nombre_producto,
    p.categoria,
    p.linea,
    p.capsula,
    p.talle || 'Unico',
    p.color || '',
    p.estado || 'Activo',
    p.precio_venta_actual || 0,
    p.margen_objetivo || 1.5,
    true,
    p.observaciones || '',
    p.sistema_talles || 'unico',
    p.talles_custom || ''
  ]);

  getSheet('STOCK_PRODUCTOS').appendRow([
    id,
    p.nombre_producto,
    p.capsula,
    0, 0, 0,
    p.precio_venta_actual || 0,
    0,
    'Sin stock'
  ]);

  // Crear variantes en STOCK_VARIANTES si tiene talles
  var talles = getTallesPorSistema(p.sistema_talles || 'unico', p.talles_custom || '');
  var sVar = SS.getSheetByName('STOCK_VARIANTES');
  if (sVar && talles.length > 0 && talles[0] !== 'Único') {
    talles.forEach(function(talle) {
      var idVar = getLastId('STOCK_VARIANTES', 'VAR-');
      sVar.appendRow([idVar, id, p.nombre_producto, talle, 0]);
    });
  }

  return { id: id };
}

// ─── STOCK ─────────────────────────────────────────────────
function getStock() { return sheetToArray('STOCK_PRODUCTOS'); }

function actualizarStock(idProducto, delta, tipo) {
  var info = getHeaderInfo('STOCK_PRODUCTOS');
  var s = info.sheet;
  var data = info.data;
  var headers = info.headers;

  var colId = colIndexRequired(headers, ['id_producto', 'ID'], 'STOCK_PRODUCTOS');
  var colVend = colIndexRequired(headers, ['cantidad_vendida', 'Vendidos'], 'STOCK_PRODUCTOS');
  var colStock = colIndexRequired(headers, ['stock_actual', 'Stock'], 'STOCK_PRODUCTOS');
  var colProd = colIndexRequired(headers, ['cantidad_producida', 'Producidos'], 'STOCK_PRODUCTOS');
  var colEstado = colIndex(headers, ['estado_stock', 'Estado']);

  for (var i = info.headerRow + 1; i < data.length; i++) {
    if (String(data[i][colId]).trim() !== String(idProducto).trim()) continue;

    var producido = Number(data[i][colProd]) || 0;
    var vendido = Number(data[i][colVend]) || 0;

    if (tipo === 'venta') vendido += delta;
    if (tipo === 'produccion') producido += delta;

    var stockActual = producido - vendido;

    s.getRange(i + 1, colVend + 1).setValue(vendido);
    s.getRange(i + 1, colProd + 1).setValue(producido);
    s.getRange(i + 1, colStock + 1).setValue(stockActual);

    if (colEstado >= 0) {
      var minimo = 5;
      try { minimo = Number(getConfig()['stock_minimo_alerta']) || 5; } catch(e) {}

      var estado = stockActual <= 0 ? 'Sin stock' : stockActual <= minimo ? 'Stock bajo' : 'OK';
      s.getRange(i + 1, colEstado + 1).setValue(estado);
    }

    return true;
  }

  throw new Error('No encontré el producto en STOCK_PRODUCTOS: ' + idProducto);
}

// ─── INSUMOS ───────────────────────────────────────────────
function getInsumos() { return sheetToArray('INSUMOS_MAESTRO'); }

function crearInsumo(p) {
  var id = getLastId('INSUMOS_MAESTRO', 'INS-');
  appendTo('INSUMOS_MAESTRO', [id, p.nombre_insumo, p.categoria, p.subcategoria||'',
    p.unidad, p.metodo_costeo||'Directo', p.stock_actual||0,
    p.costo_unitario_actual||0, p.stock_minimo||5, p.proveedor_default||'', true, p.observaciones||'']);
  return { id: id };
}

function actualizarStockInsumo(idInsumo, delta, nuevoCosto) {
  var s = getSheet('INSUMOS_MAESTRO');
  var data = s.getDataRange().getValues();

  // Detectar fila de headers automáticamente
  var headerRowIndex = -1;
  for (var r = 0; r < Math.min(data.length, 10); r++) {
    var joined = data[r].map(function(h) {
      return String(h || '').toLowerCase().trim();
    }).join('|');
    if (joined.indexOf('stock_actual') >= 0 || joined.indexOf('id_insumo') >= 0) {
      headerRowIndex = r;
      break;
    }
  }

  if (headerRowIndex < 0) {
    Logger.log('actualizarStockInsumo: no encontré headers en INSUMOS_MAESTRO');
    return;
  }

  var headers = data[headerRowIndex].map(function(h) { return String(h).trim(); });
  var colStock = headers.indexOf('stock_actual');
  var colCosto = headers.indexOf('costo_unitario_actual');

  if (colStock < 0) {
    Logger.log('actualizarStockInsumo: columna stock_actual no encontrada. Headers: ' + headers.join('|'));
    return;
  }

  for (var i = headerRowIndex + 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== String(idInsumo).trim()) continue;

    var stockActual = Number(data[i][colStock]) || 0;
    var costoActual = colCosto >= 0 ? (Number(data[i][colCosto]) || 0) : 0;

    if (delta > 0 && nuevoCosto && colCosto >= 0) {
      var nuevoStock = stockActual + delta;
      var nuevoCostoProm = nuevoStock > 0
        ? ((stockActual * costoActual) + (delta * nuevoCosto)) / nuevoStock
        : nuevoCosto;
      s.getRange(i + 1, colStock + 1).setValue(nuevoStock);
      s.getRange(i + 1, colCosto + 1).setValue(nuevoCostoProm);
    } else {
      s.getRange(i + 1, colStock + 1).setValue(Math.max(0, stockActual + delta));
    }

    break;
  }
}

// ─── COMPRAS ───────────────────────────────────────────────
function getCompras() { return sheetToArray('COMPRAS'); }

function registrarCompra(p) {
  var id = getLastId('COMPRAS', 'CMP-');
  var cantidad = Number(p.cantidad) || 0;
  var costoTotal = Number(p.costo_total) || 0;
  var costoUnit = cantidad > 0 ? costoTotal / cantidad : 0;

  appendTo('COMPRAS', [
    id,
    p.fecha,
    p.id_insumo,
    p.nombre_insumo,
    cantidad,
    p.unidad || '',
    costoTotal,
    costoUnit,
    p.proveedor || '',
    p.metodo_pago || '',
    p.pagado_por || '',
    p.capsula_asociada || '',
    p.tipo_costo || 'Variable',
    true,
    p.observaciones || ''
  ]);

  actualizarStockInsumo(p.id_insumo, cantidad, costoUnit);

  // MOVIMIENTOS_STOCK
  try {
    var sMov = SS.getSheetByName('MOVIMIENTOS_STOCK');
    if (sMov) {
      var idMov = getLastId('MOVIMIENTOS_STOCK', 'MOV-');
      appendTo('MOVIMIENTOS_STOCK', [
        idMov, p.fecha, p.id_insumo, p.nombre_insumo,
        'Entrada', cantidad, costoUnit, costoTotal,
        'Compra ' + id, p.observaciones || ''
      ]);
    }
  } catch(e) {
    Logger.log('MOVIMIENTOS_STOCK error: ' + e.message);
  }

  // EGRESOS — con categoría real del insumo
  try {
    var idEgr = getLastId('EGRESOS', 'EGR-');
    var categoriaEgreso = p.categoria_insumo || 'Materia prima';
    var subcategoriaEgreso = p.subcategoria_insumo || '';
    appendTo('EGRESOS', [
      idEgr, p.fecha, categoriaEgreso, subcategoriaEgreso,
      p.nombre_insumo, costoTotal, p.proveedor || '', p.pagado_por || '',
      p.metodo_pago || '', p.capsula_asociada || '',
      'Variable', p.id_insumo, true, 'Desde compra ' + id
    ]);
  } catch(e) {
    Logger.log('EGRESOS error: ' + e.message);
  }

  return { id: id };
}

// ─── CUEROS ────────────────────────────────────────────────
function getLotesCuero() { return sheetToArray('CUEROS_LOTES'); }

function crearLoteCuero(p) {
  var id = getLastId('CUEROS_LOTES', 'LOT-');
  var supUtil = Number(p.superficie_total) * (1 - Number(p.porcentaje_descarte || 0.2));
  var costoPorUnit = Number(p.costo_total) / supUtil;

  appendTo('CUEROS_LOTES', [
    id, p.fecha_compra, p.proveedor, p.tipo_cuero, p.color,
    p.costo_total, p.superficie_total, p.unidad_superficie || 'cm2',
    p.porcentaje_descarte || 0.2, supUtil, costoPorUnit, supUtil,
    p.capsula_asociada || '', p.observaciones || ''
  ]);

  // Crear egreso automático
  try {
    var idEgr = getLastId('EGRESOS', 'EGR-');
    appendTo('EGRESOS', [
      idEgr, p.fecha_compra, 'Materia prima', 'Cuero',
      p.tipo_cuero + ' ' + p.color + ' — ' + p.proveedor,
      p.costo_total, p.proveedor || '', '', '',
      p.capsula_asociada || 'Abarca para todas',
      'Variable', id, true, 'Lote cuero ' + id
    ]);
  } catch(e) {
    Logger.log('Error egreso cuero: ' + e.message);
  }

  return { id: id, costoPorUnidad: costoPorUnit };
}

// ─── VENTAS ────────────────────────────────────────────────
function getVentas() {
  return sheetToArray('VENTAS').filter(function(v) {
    var obs = String(val(v, ['observaciones', 'Observaciones']) || '');
    var activo = val(v, ['activo', 'Activo']);

    if (obs.indexOf('[ELIMINADA]') === 0) return false;
    if (activo === false || String(activo).toUpperCase() === 'FALSE') return false;

    return true;
  });
}

function registrarVenta(p) {
  var id = getLastId('VENTAS', 'VTA-');
  appendTo('VENTAS', [id, p.fecha, p.comprador, p.contacto||'',
    p.id_producto, p.nombre_producto, p.lugar_venta||'', p.nombre_lugar||'',
    p.edad||'', p.anio_nacimiento||'', p.monto, p.pagado||'Si',
    p.metodo_pago||'', p.capsula, p.descuento_aplicado||0, p.observaciones||'']);
  actualizarStock(p.id_producto, 1, 'venta');

  // Si la venta es del showroom, registrar retiro automático
  if (String(p.lugar_venta || '').toLowerCase() === 'showroom') {
    try {
      var sShw = SS.getSheetByName('SHOWROOM_MOVIMIENTOS');
      if (sShw) {
        var idShw = getLastId('SHOWROOM_MOVIMIENTOS', 'SHW-');
        sShw.appendRow([
  idShw,
  p.fecha,
  'Venta',
  p.id_producto,
  p.nombre_producto,
  p.capsula || '',
  1,
  'Venta registrada',
  id,
  '',
  p.talle || ''
]);
      }
    } catch(e) {
      Logger.log('Error registrando retiro showroom: ' + e.message);
    }
  }

  return { id: id };
}

function eliminarVenta(p) {
  var idVenta = p.id_venta || p.id || '';
  if (!idVenta) throw new Error('Falta id_venta');

  var info = getHeaderInfo('VENTAS');
  var s = info.sheet;
  var data = info.data;
  var headers = info.headers;

  var colId = colIndexRequired(headers, ['id_venta', 'ID', 'id'], 'VENTAS');
  var colObs = colIndex(headers, ['observaciones', 'Observaciones']);
  var colActivo = colIndex(headers, ['activo', 'Activo']);
  var colProd = colIndex(headers, ['id_producto', 'ID Producto', 'producto_id']);
  var colNombre = colIndex(headers, ['nombre_producto', 'Producto', 'Compra']);
  var colFecha = colIndex(headers, ['fecha', 'Fecha']);
  var colMonto = colIndex(headers, ['monto', 'Monto', 'importe', 'total']);

  for (var i = info.headerRow + 1; i < data.length; i++) {
    if (String(data[i][colId]).trim() !== String(idVenta).trim()) continue;

    var obsActual = colObs >= 0 ? String(data[i][colObs] || '') : '';
    if (obsActual.indexOf('[ELIMINADA]') === 0) {
      throw new Error('La venta ya estaba eliminada');
    }

    var idProducto = colProd >= 0 ? data[i][colProd] : '';
    var nombreProducto = colNombre >= 0 ? data[i][colNombre] : '';
    var fecha = colFecha >= 0 ? data[i][colFecha] : new Date();
    var monto = colMonto >= 0 ? data[i][colMonto] : '';

    if (idProducto) {
      actualizarStock(idProducto, -1, 'venta');
    }

    var marca = '[ELIMINADA] ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') +
      ' — ' +
      (p.motivo || 'Eliminada desde app');

    if (colObs >= 0) {
      s.getRange(i + 1, colObs + 1).setValue(marca + (obsActual ? ' | Antes: ' + obsActual : ''));
    }

    if (colActivo >= 0) {
      s.getRange(i + 1, colActivo + 1).setValue(false);
    }

    try {
      var idMov = getLastId('MOVIMIENTOS_STOCK', 'MOV-');
      appendTo('MOVIMIENTOS_STOCK', [
        idMov,
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        idProducto,
        nombreProducto,
        'Anulación venta',
        1,
        '',
        monto,
        'Venta anulada ' + idVenta,
        p.motivo || ''
      ]);
    } catch(e) {}

    return {
      ok: true,
      id_venta: idVenta,
      id_producto: idProducto,
      mensaje: 'Venta eliminada y stock actualizado'
    };
  }

  throw new Error('No encontré la venta: ' + idVenta);
}

// ─── EGRESOS ───────────────────────────────────────────────
function getEgresos() { return sheetToArray('EGRESOS'); }

function registrarEgreso(p) {
  var id = getLastId('EGRESOS', 'EGR-');
  appendTo('EGRESOS', [id, p.fecha, p.categoria, p.subcategoria||'',
    p.detalle, p.costo, p.proveedor||'', p.pagado_por||'', p.metodo_pago||'',
    p.capsula||'Abarca para todas', p.tipo_costo||'Variable',
    p.id_insumo||'', false, p.observaciones||'']);
  return { id: id };
}

// ─── VIAJES ────────────────────────────────────────────────
function getViajes() { return sheetToArray('VIAJES_NAFTA'); }

function registrarViaje(p) {
  var id = getLastId('VIAJES_NAFTA', 'NAF-');
  var vehs = getVehiculos();
  var veh = null;
  for (var i = 0; i < vehs.length; i++) {
    if (vehs[i].id_vehiculo === p.id_vehiculo) { veh = vehs[i]; break; }
  }
  var litC = 0, litR = 0;
  if (veh) {
    litC = p.km_ciudad ? Number(p.km_ciudad) / Number(veh.consumo_ciudad_km_l) : 0;
    litR = p.km_ruta ? Number(p.km_ruta) / Number(veh.consumo_ruta_km_l) : 0;
  }
  var costo = (litC + litR) * Number(p.valor_litro || 0);
  appendTo('VIAJES_NAFTA', [id, p.fecha, p.motivo, p.id_vehiculo,
    p.km_ciudad||0, p.km_ruta||0, p.valor_litro||0, litC, litR, costo,
    p.capsula_asociada||'Abarca para todas', p.criterio_reparto||'Por unidades',
    p.responsable||'', p.observaciones||'']);
  var idEgr = getLastId('EGRESOS', 'EGR-');
  appendTo('EGRESOS', [idEgr, p.fecha, 'Envios y logistica', 'Nafta',
    p.motivo, costo, p.responsable||'Cuero Viejo', p.responsable||'', 'Efectivo',
    p.capsula_asociada||'Abarca para todas', 'Variable', '', false, 'Viaje ' + id]);
  return { id: id, costoNafta: costo };
}

// ─── ACTIVOS ───────────────────────────────────────────────
function getActivos() { return sheetToArray('ACTIVOS_AMORTIZABLES'); }

function crearActivo(p) {
  var id = getLastId('ACTIVOS_AMORTIZABLES', 'ACT-');
  var amort = (Number(p.costo_compra) - Number(p.valor_residual||0)) / Number(p.vida_util_meses);
  appendTo('ACTIVOS_AMORTIZABLES', [id, p.nombre_activo, p.categoria||'Herramienta',
    p.fecha_compra, p.costo_compra, p.vida_util_meses, p.valor_residual||0,
    amort, p.centro_costo||'General', true, p.observaciones||'']);
  return { id: id, amortizacion_mensual: amort };
}

// ─── COSTOS FIJOS ──────────────────────────────────────────
function getCostosFijos() {
  var s = SS.getSheetByName('COSTOS_FIJOS');
  if (!s) return [];

  var data = s.getDataRange().getValues();

  var headerRowIndex = -1;
  for (var r = 0; r < Math.min(data.length, 10); r++) {
    var joined = data[r].map(function(h) {
      return String(h || '').toLowerCase().trim();
    }).join('|');
    if (joined.indexOf('id_costo_fijo') >= 0) {
      headerRowIndex = r;
      break;
    }
  }

  if (headerRowIndex < 0) return [];

  var headers = data[headerRowIndex].map(function(h) {
    return String(h || '').trim();
  });

  var colActivo = headers.map(function(h) {
    return String(h).toLowerCase().trim();
  }).indexOf('activo');

  var rows = [];
  for (var i = headerRowIndex + 1; i < data.length; i++) {
    var row = data[i];

    var hasContent = row.some(function(c) {
      return c !== '' && c !== null && c !== undefined;
    });
    if (!hasContent) continue;

    // Filtrar inactivos
    if (colActivo >= 0) {
      var activo = row[colActivo];
      if (activo === false || String(activo).toUpperCase() === 'FALSE') continue;
    }

    var obj = {};
    for (var k = 0; k < headers.length; k++) {
      if (headers[k]) obj[headers[k]] = sanitize(row[k]);
    }
    rows.push(obj);
  }

  return rows;
}

function registrarCostoFijo(p) {
  var id = getLastId('COSTOS_FIJOS', 'CF-');

  var s = getSheet('COSTOS_FIJOS');
  var headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(function(h) {
    return String(h || '').trim().toLowerCase();
  });
  var tieneColTipo = headers.indexOf('tipo_asignacion') >= 0;

  var row = [
    id,
    p.fecha,
    p.mes || '',
    p.categoria,
    p.subcategoria || '',
    p.detalle,
    p.importe,
    p.frecuencia || 'Único',
    p.centro_costo || 'General',
    p.capsula_asociada || 'Abarca para todas',
    p.criterio_reparto || 'Por unidades',
    true,
    p.observaciones || ''
  ];

  if (tieneColTipo) {
    row.push(p.tipo_asignacion || 'Estructural');
  }

  appendTo('COSTOS_FIJOS', row);

  try {
    var idEgr = getLastId('EGRESOS', 'EGR-');
    appendTo('EGRESOS', [
      idEgr, p.fecha, p.categoria || 'Costo fijo', p.subcategoria || '',
      p.detalle, p.importe, '', '', '',
      p.capsula_asociada || 'Abarca para todas',
      'Fijo', '', true, 'Desde costo fijo ' + id
    ]);
  } catch(e) {
    Logger.log('Error egreso costo fijo: ' + e.message);
  }

  return { id: id };
}

// ─── PRODUCCION ────────────────────────────────────────────
function getProduccion() {
  var s = SS.getSheetByName('PRODUCCION');
  if (!s) return [];

  return sheetToArray('PRODUCCION').filter(function(p) {
    var obs = String(val(p, ['observaciones', 'Observaciones']) || '');
    var activo = val(p, ['activo', 'Activo']);

    if (obs.indexOf('[ELIMINADA]') === 0) return false;
    if (activo === false || String(activo).toUpperCase() === 'FALSE') return false;

    return true;
  });
}

function registrarProduccion(p) {
  var sP = SS.getSheetByName('PRODUCCION') || SS.insertSheet('PRODUCCION');
  if (sP.getLastRow() === 0) {
    sP.appendRow(['id_produccion','fecha','id_producto','nombre_producto',
      'cantidad_producida','lote','estado','id_cotizacion_ref',
      'costo_estimado_unitario','observaciones']);
  }
  var id = getLastId('PRODUCCION', 'PRD-P');
  sP.appendRow([id, p.fecha, p.id_producto, p.nombre_producto,
    p.cantidad_producida, p.lote||'', p.estado||'Completada',
    p.id_cotizacion_ref||'', p.costo_estimado_unitario||0, p.observaciones||'']);
  actualizarStock(p.id_producto, Number(p.cantidad_producida), 'produccion');
  // Descontar insumos
  var recetas = sheetToArray('RECETAS_PRODUCTOS');
  for (var i = 0; i < recetas.length; i++) {
    if (recetas[i].id_producto !== p.id_producto) continue;
    var cantUsada = Number(recetas[i].cantidad_requerida) * Number(p.cantidad_producida);
    actualizarStockInsumo(recetas[i].id_insumo, -cantUsada, null);
  }
  return { id: id };
}

function getProduccionAgrupada() {
  var produccion = getProduccion();
  var stock = sheetToArray('STOCK_PRODUCTOS');
  var productos = sheetToArray('PRODUCTOS_MAESTRO');
  var capsulas = getCapsulas();

  var out = {};

  capsulas.forEach(function(c) {
    out[c] = {
      capsula: c,
      unidades_producidas: 0,
      unidades_vendidas: 0,
      stock_actual: 0,
      costo_estimado_total: 0,
      productos: []
    };
  });

  stock.forEach(function(s) {
    var cap = capsulaDisplay(val(s, ['capsula', 'Cápsula', 'Capsula']), capsulas) || 'Sin cápsula';

    if (!out[cap]) {
      out[cap] = {
        capsula: cap,
        unidades_producidas: 0,
        unidades_vendidas: 0,
        stock_actual: 0,
        costo_estimado_total: 0,
        productos: []
      };
    }

    var idProducto = val(s, ['id_producto', 'ID', 'id']);
    var nombreProducto = val(s, ['nombre_producto', 'Producto', 'producto']);
    var producido = money(val(s, ['cantidad_producida', 'Producidos']));
    var vendido = money(val(s, ['cantidad_vendida', 'Vendidos']));
    var stockActual = money(val(s, ['stock_actual', 'Stock']));
    var costoRef = money(val(s, ['costo_unitario_ref', 'Costo ref', 'costo_unitario']));
    var precio = money(val(s, ['precio_venta_actual', 'Precio']));
    var estado = val(s, ['estado_stock', 'Estado']) || '';

    var registrosProd = produccion.filter(function(p) {
      return String(val(p, ['id_producto', 'ID Producto'])).trim() === String(idProducto).trim();
    });

    var ultimo = registrosProd.length ? registrosProd[registrosProd.length - 1] : null;

    out[cap].unidades_producidas += producido;
    out[cap].unidades_vendidas += vendido;
    out[cap].stock_actual += stockActual;
    out[cap].costo_estimado_total += producido * costoRef;

    out[cap].productos.push({
      id_producto: idProducto,
      nombre_producto: nombreProducto,
      capsula: cap,
      cantidad_producida: producido,
      cantidad_vendida: vendido,
      stock_actual: stockActual,
      costo_unitario_ref: costoRef,
      precio_venta_actual: precio,
      estado_stock: estado,
      ultimo_lote: ultimo ? val(ultimo, ['lote', 'Lote']) : '',
      ultima_fecha: ultimo ? val(ultimo, ['fecha', 'Fecha']) : '',
      ultimo_estado: ultimo ? val(ultimo, ['estado', 'Estado']) : ''
    });
  });

  return Object.keys(out).map(function(k) { return out[k]; });
}

function editarProduccion(p) {
  var idProducto = p.id_producto || '';
  if (!idProducto) throw new Error('Falta id_producto');

  var nuevaCantidad = Number(p.cantidad_producida);
  if (isNaN(nuevaCantidad)) throw new Error('Cantidad producida inválida');

  var info = getHeaderInfo('STOCK_PRODUCTOS');
  var s = info.sheet;
  var data = info.data;
  var headers = info.headers;

  var colId = colIndexRequired(headers, ['id_producto', 'ID'], 'STOCK_PRODUCTOS');
  var colProd = colIndexRequired(headers, ['cantidad_producida', 'Producidos'], 'STOCK_PRODUCTOS');
  var colVend = colIndexRequired(headers, ['cantidad_vendida', 'Vendidos'], 'STOCK_PRODUCTOS');
  var colStock = colIndexRequired(headers, ['stock_actual', 'Stock'], 'STOCK_PRODUCTOS');
  var colEstado = colIndex(headers, ['estado_stock', 'Estado']);

  for (var i = info.headerRow + 1; i < data.length; i++) {
    if (String(data[i][colId]).trim() !== String(idProducto).trim()) continue;

    var cantidadAnterior = Number(data[i][colProd]) || 0;
    var vendido = Number(data[i][colVend]) || 0;
    var diferencia = nuevaCantidad - cantidadAnterior;
    var nuevoStock = nuevaCantidad - vendido;

    s.getRange(i + 1, colProd + 1).setValue(nuevaCantidad);
    s.getRange(i + 1, colStock + 1).setValue(nuevoStock);

    if (colEstado >= 0) {
      var minimo = 5;
      try { minimo = Number(getConfig()['stock_minimo_alerta']) || 5; } catch(e) {}

      var estado = nuevoStock <= 0 ? 'Sin stock' : nuevoStock <= minimo ? 'Stock bajo' : 'OK';
      s.getRange(i + 1, colEstado + 1).setValue(estado);
    }

    try {
      var idMov = getLastId('MOVIMIENTOS_STOCK', 'MOV-');
      appendTo('MOVIMIENTOS_STOCK', [
        idMov,
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        idProducto,
        p.nombre_producto || '',
        'Ajuste producción',
        diferencia,
        '',
        '',
        'Edición manual de producción',
        p.motivo || ''
      ]);
    } catch(e) {}

    try {
      var sP = SS.getSheetByName('PRODUCCION') || SS.insertSheet('PRODUCCION');

      if (sP.getLastRow() === 0) {
        sP.appendRow([
          'id_produccion','fecha','id_producto','nombre_producto',
          'cantidad_producida','lote','estado','id_cotizacion_ref',
          'costo_estimado_unitario','observaciones','activo'
        ]);
      }

      var idProd = getLastId('PRODUCCION', 'PRD-P');
      sP.appendRow([
        idProd,
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        idProducto,
        p.nombre_producto || '',
        diferencia,
        p.lote || '',
        'Ajuste',
        '',
        p.costo_estimado_unitario || '',
        'Ajuste producción. Antes: ' + cantidadAnterior + ' | Ahora: ' + nuevaCantidad + ' | Motivo: ' + (p.motivo || ''),
        true
      ]);
    } catch(e) {}

    return {
      ok: true,
      id_producto: idProducto,
      cantidad_anterior: cantidadAnterior,
      cantidad_nueva: nuevaCantidad,
      diferencia: diferencia,
      stock_actual: nuevoStock
    };
  }

  throw new Error('No encontré el producto en STOCK_PRODUCTOS: ' + idProducto);
}

// ─── COSTEO ────────────────────────────────────────────────
function getCotizaciones() {
  var s = SS.getSheetByName('COTIZACIONES');
  if (!s) return [];

  return sheetToArray('COTIZACIONES').filter(function(c) {
    var id = val(c, ['id_cotizacion', 'ID', 'id']);
    var producto = val(c, ['nombre_producto', 'Producto', 'producto']);
    var precio = money(val(c, ['precio_venta_final', 'Precio final', 'precio_final']));

    return id || producto || precio > 0;
  });
}

function calcularCosteo(p) {
  var cfg = getConfig();

  // ─── Parámetros editables desde CONFIG ───────────────────
  var capacidadPracticaAnual = Number(cfg['capacidad_practica_anual'] || 1400);
  var unidadesMetaCapsula    = Number(p.unidades_meta_capsula || cfg['unidades_meta_capsula'] || 15);
  var valorHora              = Number(p.valor_hora || cfg['valor_hora_mo'] || 7000);
  var margenObjetivo         = Number(p.margen_objetivo || cfg['margen_objetivo_porcentaje'] || 35) / 100;
  var comision               = Number(p.comision || cfg['comision_porcentaje'] || 0) / 100;
  var descuentoEsperado      = Number(p.descuento_esperado || cfg['descuento_esperado_porcentaje'] || 10) / 100;
  var colchonRiesgo          = Number(p.colchon_riesgo || cfg['colchon_riesgo_porcentaje'] || 5) / 100;

  var alertas = [];

  if (capacidadPracticaAnual <= 0) {
    capacidadPracticaAnual = 1400;
    alertas.push('capacidad_practica_anual es 0 o negativa — se usó fallback 1400 hs');
  }
  if (unidadesMetaCapsula <= 0) {
    unidadesMetaCapsula = 15;
    alertas.push('unidades_meta_capsula es 0 o negativa — se usó fallback 15 u');
  }
  if (comision + margenObjetivo + descuentoEsperado + colchonRiesgo >= 0.95) {
    alertas.push('La suma de comisión + margen + descuento + colchón supera el 95% — el precio sugerido puede ser irreal');
  }

  // ─── Horas del producto ───────────────────────────────────
  var horasProducto = Number(p.horas_mo || 0);

  // ─── 1. Costo variable directo ────────────────────────────
  var costoCuero = Number(p.superficie_cuero || 0) * Number(p.costo_cm2_cuero || 0);

  var costoInsumos = 0;
  if (p.insumos && p.insumos.length) {
    p.insumos.forEach(function(ins) {
      costoInsumos += Number(ins.cantidad || 0) * Number(ins.costo_unitario || 0) * (1 + Number(ins.merma || 0));
    });
  }

  var costoMO = horasProducto * Number(p.personas_mo || 1) * valorHora;
  var costoVarDirecto = costoCuero + costoInsumos + costoMO;

  // ─── 2. Tasa estructural por hora ─────────────────────────
  var cf = [];
  try { cf = getCostosFijos(); } catch(e) {}

  var totalEstructuralAnual = 0;
  cf.forEach(function(c) {
    var tipo = String(c.tipo_asignacion || '').trim();
    if (tipo !== 'Estructural') return;

    var importe    = Number(c.importe || 0);
    var frecuencia = String(c.frecuencia || 'Único').toLowerCase();

    if (frecuencia === 'mensual') {
      totalEstructuralAnual += importe * 12;
    } else {
      // Anual o Único: se absorbe en el año en curso
      totalEstructuralAnual += importe;
    }
  });

  // Amortizaciones anuales
  var activos = [];
  try { activos = getActivos(); } catch(e) {}
  var amortizacionAnual = 0;
  activos.forEach(function(a) {
    var activo = a.activo;
    if (activo === false || String(activo).toUpperCase() === 'FALSE') return;
    amortizacionAnual += Number(a.amortizacion_mensual || 0) * 12;
  });

  var totalAnualConAmort     = totalEstructuralAnual + amortizacionAnual;
  var tasaEstructuralPorHora = capacidadPracticaAnual > 0 ? totalAnualConAmort / capacidadPracticaAnual : 0;
  var costoEstructuralUnitario = tasaEstructuralPorHora * horasProducto;
  var amortUnitario = capacidadPracticaAnual > 0 ? (amortizacionAnual / capacidadPracticaAnual) * horasProducto : 0;

  // ─── 3. Tasa de cápsula por unidad ───────────────────────
  // Driver: unidades meta (no horas) — fotos e influencer benefician a todos por igual
  var totalCostosCapsula = 0;
  cf.forEach(function(c) {
    var tipo = String(c.tipo_asignacion || '').trim();
    if (tipo !== 'Cápsula') return;

    var capAsoc   = String(c.capsula_asociada || '').trim();
    var capActual = String(p.capsula || '').trim();
    if (!capActual) return;
    if (normCapsula(capAsoc) !== normCapsula(capActual)) return;

    totalCostosCapsula += Number(c.importe || 0);
  });

  // Viajes asociados a la cápsula
  var viajes = [];
  try { viajes = getViajes(); } catch(e) {}
  var totalViajesCapsula = 0;
  viajes.forEach(function(v) {
    var capAsoc   = String(v.capsula_asociada || '').trim();
    var capActual = String(p.capsula || '').trim();
    if (!capActual) return;
    if (normCapsula(capAsoc) === normCapsula(capActual)) {
      totalViajesCapsula += Number(v.costo_total_nafta || 0);
    }
  });

  var totalLoteCapsula    = totalCostosCapsula + totalViajesCapsula;
  var costoCapsulaUnitario = unidadesMetaCapsula > 0 ? totalLoteCapsula / unidadesMetaCapsula : 0;

  // ─── 4. Costo total unitario ──────────────────────────────
  var costoTotal = costoVarDirecto + costoEstructuralUnitario + costoCapsulaUnitario;

  if (costoTotal <= 0) {
    alertas.push('El costo total es 0 o negativo — verificá los datos ingresados');
  }

  // ─── 5. Precio por margen real sobre precio ───────────────
  var denominador    = 1 - comision - margenObjetivo - descuentoEsperado - colchonRiesgo;
  var precioMinimo   = comision > 0 ? costoTotal / (1 - comision) : costoTotal;
  var precioSugerido = denominador > 0 ? costoTotal / denominador : 0;

  if (precioSugerido <= costoTotal && denominador > 0) {
    alertas.push('El precio sugerido no cubre el costo total — revisá los parámetros de margen');
  }

  var gananciaUnitaria      = precioSugerido - costoTotal;
  var margenRealSobrePrecio = precioSugerido > 0 ? ((precioSugerido - costoTotal) / precioSugerido) * 100 : 0;
  var precioPromo10         = precioSugerido * 0.9;
  var precioPromo20         = precioSugerido * 0.8;
  var margenPostDesc10      = precioPromo10 > 0 ? ((precioPromo10 - costoTotal) / precioPromo10) * 100 : 0;
  var margenPostDesc20      = precioPromo20 > 0 ? ((precioPromo20 - costoTotal) / precioPromo20) * 100 : 0;

  return {
    // Costos directos
    costoCuero:              costoCuero,
    costoInsumos:            costoInsumos,
    costoMO:                 costoMO,
    costoVarDirecto:         costoVarDirecto,
    // Tasa estructural
    totalEstructuralAnual:   totalEstructuralAnual,
    amortizacionAnual:       amortizacionAnual,
    totalAnualConAmort:      totalAnualConAmort,
    tasaEstructuralPorHora:  tasaEstructuralPorHora,
    costoEstructuralUnitario:costoEstructuralUnitario,
    amortUnitario:           amortUnitario,
    // Tasa cápsula
    totalCostosCapsula:      totalCostosCapsula,
    totalViajesCapsula:      totalViajesCapsula,
    totalLoteCapsula:        totalLoteCapsula,
    unidadesMetaCapsula:     unidadesMetaCapsula,
    costoCapsulaUnitario:    costoCapsulaUnitario,
    // Total
    costoTotal:              costoTotal,
    // Precio
    margenObjetivo:          margenObjetivo * 100,
    comision:                comision * 100,
    descuentoEsperado:       descuentoEsperado * 100,
    colchonRiesgo:           colchonRiesgo * 100,
    precioMinimo:            precioMinimo,
    precioSugerido:          precioSugerido,
    gananciaUnitaria:        gananciaUnitaria,
    margenRealSobrePrecio:   margenRealSobrePrecio,
    margenSobreprecio:       margenRealSobrePrecio,
    // Promos
    precioPromo10:           precioPromo10,
    precioPromo20:           precioPromo20,
    margenPostDesc10:        margenPostDesc10,
    margenPostDesc20:        margenPostDesc20,
    // Gerencial
    capacidadPracticaAnual:  capacidadPracticaAnual,
    // Alertas
    alertas:                 alertas
  };
}

function guardarCotizacion(p) {
  var s = SS.getSheetByName('COTIZACIONES');
  if (!s) return { error: 'Hoja COTIZACIONES no encontrada' };

  var id = getLastId('COTIZACIONES', 'COT-');

  s.appendRow([
    id, p.fecha, p.id_producto || '', p.nombre_producto || '',
    p.costo_cuero || 0, p.costo_insumos || 0, p.costo_mano_obra || 0,
    p.costo_fijo_asignado || 0, p.amortizacion_asignada || 0, p.nafta_asignada || 0,
    p.costo_total_unitario || 0, p.margen_aplicado || 0, p.comision_aplicada || 0,
    p.descuento_esperado || 0, p.colchon_riesgo || 0,
    p.precio_venta_final || 0, true, p.observaciones || ''
  ]);

  if (p.es_producto_nuevo) {
    // Crear producto nuevo en PRODUCTOS_MAESTRO
    var idNuevo = getLastId('PRODUCTOS_MAESTRO', 'PRD-');
    appendTo('PRODUCTOS_MAESTRO', [
      idNuevo,           // id_producto
      idNuevo,           // id_producto_base
      p.nombre_producto, // nombre_producto
      '',                // categoria — completar después
      '',                // linea
      p.capsula || '',   // capsula
      'Unico',           // talle
      '',                // color
      'Activo',          // estado
      p.precio_venta_final || 0,  // precio_venta_actual
      0,                 // margen_objetivo
      true,              // activo
      'Creado desde costeo ' + id  // observaciones
    ]);

    // Crear fila en STOCK_PRODUCTOS
    appendTo('STOCK_PRODUCTOS', [
      idNuevo,
      p.nombre_producto,
      p.capsula || '',
      0,  // cantidad_producida
      0,  // cantidad_vendida
      0,  // stock_actual
      p.precio_venta_final || 0,
      p.costo_total_unitario || 0,
      'Sin stock'
    ]);

    return { id: id, id_producto_creado: idNuevo };
  }

  // Producto existente — actualizar precio y costo sin tocar ventas
  if (p.id_producto && p.precio_venta_final) {
    var rowPM = findRow('PRODUCTOS_MAESTRO', 0, p.id_producto);
    if (rowPM > 0) {
      var pm  = getSheet('PRODUCTOS_MAESTRO');
      var hPM = pm.getRange(1, 1, 1, pm.getLastColumn()).getValues()[0]
                  .map(function(h) { return String(h).trim(); });
      var colPrecio = hPM.indexOf('precio_venta_actual');
      if (colPrecio >= 0) pm.getRange(rowPM, colPrecio + 1).setValue(p.precio_venta_final);
    }

    var rowSP = findRow('STOCK_PRODUCTOS', 0, p.id_producto);
    if (rowSP > 0) {
      var sp  = getSheet('STOCK_PRODUCTOS');
      var hSP = sp.getRange(1, 1, 1, sp.getLastColumn()).getValues()[0]
                  .map(function(h) { return String(h).trim(); });
      var colCosto    = hSP.indexOf('costo_unitario_ref');
      var colPrecioSP = hSP.indexOf('precio_venta_actual');
      if (colCosto >= 0)    sp.getRange(rowSP, colCosto + 1).setValue(p.costo_total_unitario);
      if (colPrecioSP >= 0) sp.getRange(rowSP, colPrecioSP + 1).setValue(p.precio_venta_final);
    }
  }

  return { id: id };
}

// ─── TAREAS ────────────────────────────────────────────────
function getTareas() { return sheetToArray('TAREAS'); }

function crearTarea(p) {
  var id = getLastId('TAREAS', 'TSK-');
  appendTo('TAREAS', [
    id,
    p.tarea,
    p.descripcion || '',
    p.prioridad || 'Media',
    p.estado || 'No iniciada',
    p.fecha_creacion || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    p.fecha_vencimiento || '',
    p.responsable || '',
    false,
    p.origen || 'Manual',
    p.referencia || '',
    '' // id_evento_calendar
  ]);

  // Crear evento en Google Calendar
  try {
    crearEventoTarea({
      id_tarea: id,
      tarea: p.tarea,
      descripcion: p.descripcion || '',
      prioridad: p.prioridad || 'Media',
      responsable: p.responsable || '',
      fecha_vencimiento: p.fecha_vencimiento || ''
    });
  } catch(e) {
    Logger.log('Error creando evento para tarea ' + id + ': ' + e.message);
  }

  return { id: id };
}

function actualizarTarea(p) {
  var s = getSheet('TAREAS');
  var data = s.getDataRange().getValues();

  // Detectar fila de headers
  var headerRowIndex = -1;
  for (var r = 0; r < Math.min(data.length, 10); r++) {
    var joined = data[r].map(function(h) {
      return String(h || '').toLowerCase().trim();
    }).join('|');
    if (joined.indexOf('id_tarea') >= 0) {
      headerRowIndex = r;
      break;
    }
  }

  if (headerRowIndex < 0) throw new Error('No encontré headers en TAREAS');

  var headers = data[headerRowIndex].map(function(h) {
    return String(h).trim();
  });

  for (var i = headerRowIndex + 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== String(p.id_tarea).trim()) continue;

    var colEstado = headers.indexOf('estado');
    if (colEstado >= 0 && p.estado !== undefined) {
      s.getRange(i + 1, colEstado + 1).setValue(p.estado);
    }

    var colCompletada = headers.indexOf('completada');
    if (colCompletada >= 0 && p.completada !== undefined) {
      s.getRange(i + 1, colCompletada + 1).setValue(p.completada);
    }

    // Sincronizar con Google Calendar
    if (p.estado !== undefined) {
      try {
        actualizarEventoTarea(p.id_tarea, p.estado);
      } catch(e) {
        Logger.log('Error actualizando evento Calendar: ' + e.message);
      }
    }

    return { ok: true, id_tarea: p.id_tarea, estado: p.estado };
  }

  throw new Error('Tarea no encontrada: ' + p.id_tarea);
}

// ─── SETS ──────────────────────────────────────────────────
function getSets() { return sheetToArray('SETS_PERSONALIZADOS'); }

function crearSet(p) {
  var id = getLastId('SETS_PERSONALIZADOS', 'SET-');
  appendTo('SETS_PERSONALIZADOS', [id, p.fecha_consulta, p.nombre_cliente, p.contacto||'',
    p.instagram||'', p.tipo_cuero_preferido||'', p.color_preferido||'',
    p.fecha_necesaria||'', p.referencia_diseno||'', p.estado||'Consulta', '', '', p.observaciones||'']);
  return { id: id };
}

// ─── KPIs ──────────────────────────────────────────────────
function getKPIs() {
  var d = getDashboard();
  var ventas = getVentas();
  var stock = sheetToArray('STOCK_PRODUCTOS');
  var activos = sheetToArray('ACTIVOS_AMORTIZABLES');
  var ticketProm = ventas.length ? d.totalVentas / ventas.length : 0;
  var margenBruto = d.totalVentas > 0 ? ((d.totalVentas - d.totalEgresos) / d.totalVentas) * 100 : 0;
  var vendidos = 0, producidos = 0;
  stock.forEach(function(p) {
    vendidos += Number(p.cantidad_vendida) || 0;
    producidos += Number(p.cantidad_producida) || 0;
  });
  var rotacion = producidos > 0 ? (vendidos / producidos) * 100 : 0;
  return {
    totalVentas: d.totalVentas, totalEgresos: d.totalEgresos,
    beneficioNeto: d.beneficioNeto, amortizacionMensual: d.amortizacionMensual,
    stockSinStock: d.stockSinStock, stockBajo: d.stockBajo,
    ingresoPotencial: d.ingresoPotencial, pendientePago: d.pendientePago,
    totalProductos: d.totalProductos, totalVentasCount: d.totalVentasCount,
    ticketPromedio: ticketProm, margenBruto: margenBruto,
    rotacionInventario: rotacion, totalActivos: activos.length,
    amortizacionAnual: d.amortizacionMensual * 12
  };
}

// ─── MOVIMIENTOS ───────────────────────────────────────────
function getMovimientos() { return sheetToArray('MOVIMIENTOS_STOCK'); }

// ─── FINANZAS ──────────────────────────────────────────────
function getFinanzas() {
  var ventas = getVentas();
  var egresos = sheetToArray('EGRESOS');
  var capsulas = getCapsulas();

  var meses = {};
  var totalV = 0;
  var totalE = 0;

  ventas.forEach(function(v) {
    var fecha = val(v, ['fecha', 'Fecha']);
    var monto = money(val(v, ['monto', 'Monto', 'importe', 'Importe', 'total', 'Total']));

    totalV += monto;

    if (!fecha) return;
    var d = new Date(fecha);
    if (isNaN(d)) return;

    var k = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
    if (!meses[k]) meses[k] = { ventas: 0, egresos: 0 };
    meses[k].ventas += monto;
  });

  egresos.forEach(function(e) {
    var fecha = val(e, ['fecha', 'Fecha']);
    var costo = money(val(e, ['costo', 'Costo', 'importe', 'Importe', 'monto', 'Monto']));

    totalE += costo;

    if (!fecha) return;
    var d = new Date(fecha);
    if (isNaN(d)) return;

    var k = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
    if (!meses[k]) meses[k] = { ventas: 0, egresos: 0 };
    meses[k].egresos += costo;
  });

  var timeline = Object.keys(meses).sort().map(function(k) {
    return {
      mes: k,
      ventas: meses[k].ventas,
      egresos: meses[k].egresos,
      resultado: meses[k].ventas - meses[k].egresos
    };
  });

  var porCapsula = {};
  capsulas.forEach(function(c) {
    porCapsula[c] = { ventas: 0, egresos: 0 };
  });

  ventas.forEach(function(v) {
    var cap = capsulaDisplay(val(v, ['capsula', 'Cápsula', 'Capsula', 'capsula_asociada']), capsulas) || 'Sin cápsula';
    var monto = money(val(v, ['monto', 'Monto', 'importe', 'Importe', 'total', 'Total']));

    if (!porCapsula[cap]) porCapsula[cap] = { ventas: 0, egresos: 0 };
    porCapsula[cap].ventas += monto;
  });

  egresos.forEach(function(e) {
    var cap = capsulaDisplay(val(e, ['capsula', 'Cápsula', 'Capsula', 'capsula_asociada']), capsulas) || 'Sin cápsula';
    var costo = money(val(e, ['costo', 'Costo', 'importe', 'Importe', 'monto', 'Monto']));

    if (!porCapsula[cap]) porCapsula[cap] = { ventas: 0, egresos: 0 };
    porCapsula[cap].egresos += costo;
  });

  return {
    totalVentas: totalV,
    totalEgresos: totalE,
    beneficio: totalV - totalE,
    timeline: timeline,
    porCapsula: porCapsula
  };
}
function crearCapsula(p) {
  var nombre = String(p.nombre_capsula || '').trim();
  if (!nombre) throw new Error('Falta el nombre de la cápsula');

  var capsulas = getCapsulas();

  var existe = capsulas.some(function(c) {
    return normCapsula(c) === normCapsula(nombre);
  });

  if (existe) {
    throw new Error('La cápsula ya existe: ' + nombre);
  }

  capsulas.push(nombre);

  actualizarConfigFlexible('capsulas_activas', capsulas.join(';'));

  return {
    ok: true,
    nombre_capsula: nombre,
    capsulas: capsulas
  };
}

function actualizarConfigFlexible(clave, valor) {
  var info = getHeaderInfo('CONFIG');
  var s = info.sheet;
  var data = info.data;
  var headers = info.headers;

  var colClave = colIndexRequired(headers, ['clave', 'Clave'], 'CONFIG');
  var colValor = colIndexRequired(headers, ['valor', 'Valor'], 'CONFIG');

  for (var i = info.headerRow + 1; i < data.length; i++) {
    if (String(data[i][colClave]).trim() === String(clave).trim()) {
      s.getRange(i + 1, colValor + 1).setValue(valor);
      return true;
    }
  }

  var row = new Array(headers.length).fill('');
  row[colClave] = clave;
  row[colValor] = valor;

  var colCategoria = colIndex(headers, ['categoria', 'Categoría']);
  var colDesc = colIndex(headers, ['descripcion', 'Descripción']);
  var colActivo = colIndex(headers, ['activo', 'Activo']);

  if (colCategoria >= 0) row[colCategoria] = 'sistema';
  if (colDesc >= 0) row[colDesc] = 'Cápsulas activas separadas por ;';
  if (colActivo >= 0) row[colActivo] = true;

  s.appendRow(row);
  return true;
}

function eliminarCapsula(p) {
  var nombre = String(p.nombre_capsula || '').trim();
  if (!nombre) throw new Error('Falta el nombre de la cápsula');

  var capsulas = getCapsulas();

  var nuevas = capsulas.filter(function(c) {
    return normCapsula(c) !== normCapsula(nombre);
  });

  if (nuevas.length === capsulas.length) {
    throw new Error('No encontré la cápsula: ' + nombre);
  }

  actualizarConfigFlexible('capsulas_activas', nuevas.join(';'));

  return {
    ok: true,
    capsula_eliminada: nombre,
    capsulas: nuevas
  };
}

function previsualizarEliminacionCapsula(p) {
  var nombre = String(p.nombre_capsula || '').trim();
  if (!nombre) throw new Error('Falta el nombre de la cápsula');

  var productos = sheetToArray('PRODUCTOS_MAESTRO').filter(function(r) {
    return sameCapsulaFlexible(val(r, ['capsula', 'Cápsula', 'Capsula']), nombre);
  });

  var idsProductos = productos.map(function(r) {
    return String(val(r, ['id_producto', 'ID', 'id']) || '').trim();
  }).filter(Boolean);

  var ventas = sheetToArray('VENTAS').filter(function(r) {
    return sameCapsulaFlexible(val(r, ['capsula', 'Cápsula', 'Capsula', 'capsula_asociada']), nombre);
  });

  var stock = sheetToArray('STOCK_PRODUCTOS').filter(function(r) {
    var capMatch = sameCapsulaFlexible(val(r, ['capsula', 'Cápsula', 'Capsula']), nombre);
    var id = String(val(r, ['id_producto', 'ID', 'id']) || '').trim();
    return capMatch || idsProductos.indexOf(id) >= 0;
  });

  var egresos = sheetToArray('EGRESOS').filter(function(r) {
    return sameCapsulaFlexible(val(r, ['capsula', 'Cápsula', 'Capsula', 'capsula_asociada']), nombre);
  });

  var compras = sheetToArray('COMPRAS').filter(function(r) {
    return sameCapsulaFlexible(val(r, ['capsula_asociada', 'capsula', 'Cápsula', 'Capsula']), nombre);
  });

  var viajes = sheetToArray('VIAJES_NAFTA').filter(function(r) {
    return sameCapsulaFlexible(val(r, ['capsula_asociada', 'capsula', 'Cápsula', 'Capsula']), nombre);
  });

  var cueros = sheetToArray('CUEROS_LOTES').filter(function(r) {
    return sameCapsulaFlexible(val(r, ['capsula_asociada', 'capsula', 'Cápsula', 'Capsula']), nombre);
  });

  var costosFijos = sheetToArray('COSTOS_FIJOS').filter(function(r) {
    return sameCapsulaFlexible(val(r, ['capsula_asociada', 'capsula', 'Cápsula', 'Capsula']), nombre);
  });

  var produccion = sheetToArray('PRODUCCION').filter(function(r) {
    var id = String(val(r, ['id_producto', 'ID Producto', 'producto_id']) || '').trim();
    return idsProductos.indexOf(id) >= 0;
  });

  var cotizaciones = sheetToArray('COTIZACIONES').filter(function(r) {
    return sameCapsulaFlexible(val(r, ['capsula', 'Cápsula', 'Capsula']), nombre);
  });

  return {
    capsula: nombre,
    productos: productos.length,
    stock: stock.length,
    ventas: ventas.length,
    egresos: egresos.length,
    compras: compras.length,
    viajes: viajes.length,
    cueros: cueros.length,
    costos_fijos: costosFijos.length,
    produccion: produccion.length,
    cotizaciones: cotizaciones.length,
    ids_productos: idsProductos
  };
}

function sameCapsulaFlexible(a, b) {
  return normCapsula(a) === normCapsula(b);
}

function marcarFilasPorCapsula(nombreHoja, nombreCapsula, posiblesColsCapsula, marca) {
  var s = SS.getSheetByName(nombreHoja);
  if (!s) return 0;

  var info = getHeaderInfo(nombreHoja);
  var data = info.data;
  var headers = info.headers;

  var colCapsula = colIndex(headers, posiblesColsCapsula);
  if (colCapsula < 0) return 0;

  var colObs = colIndex(headers, ['observaciones', 'Observaciones', 'obs']);
  var colActivo = colIndex(headers, ['activo', 'Activo']);
  var eliminadas = 0;

  for (var i = info.headerRow + 1; i < data.length; i++) {
    if (!sameCapsulaFlexible(data[i][colCapsula], nombreCapsula)) continue;

    if (colObs >= 0) {
      var obsActual = String(data[i][colObs] || '');
      s.getRange(i + 1, colObs + 1).setValue(marca + (obsActual ? ' | Antes: ' + obsActual : ''));
    }

    if (colActivo >= 0) {
      s.getRange(i + 1, colActivo + 1).setValue(false);
    }

    eliminadas++;
  }

  return eliminadas;
}

function marcarFilasPorProducto(nombreHoja, idsProductos, posiblesColsProducto, marca) {
  var s = SS.getSheetByName(nombreHoja);
  if (!s || !idsProductos || !idsProductos.length) return 0;

  var info = getHeaderInfo(nombreHoja);
  var data = info.data;
  var headers = info.headers;

  var colProducto = colIndex(headers, posiblesColsProducto);
  if (colProducto < 0) return 0;

  var colObs = colIndex(headers, ['observaciones', 'Observaciones', 'obs']);
  var colActivo = colIndex(headers, ['activo', 'Activo']);
  var eliminadas = 0;

  for (var i = info.headerRow + 1; i < data.length; i++) {
    var id = String(data[i][colProducto] || '').trim();
    if (idsProductos.indexOf(id) < 0) continue;

    if (colObs >= 0) {
      var obsActual = String(data[i][colObs] || '');
      s.getRange(i + 1, colObs + 1).setValue(marca + (obsActual ? ' | Antes: ' + obsActual : ''));
    }

    if (colActivo >= 0) {
      s.getRange(i + 1, colActivo + 1).setValue(false);
    }

    eliminadas++;
  }

  return eliminadas;
}


function editarProducto(p) {
  var id = String(p.id_producto || '').trim();
  if (!id) throw new Error('Falta id_producto');

  var info = getHeaderInfo('PRODUCTOS_MAESTRO');
  var s = info.sheet;
  var data = info.data;
  var headers = info.headers;

  var colId = colIndexRequired(headers, ['id_producto', 'ID', 'id'], 'PRODUCTOS_MAESTRO');

  var campos = {
    nombre_producto: ['nombre_producto', 'Producto', 'producto', 'nombre'],
    categoria: ['categoria', 'Categoría'],
    linea: ['linea', 'Línea'],
    capsula: ['capsula', 'Cápsula', 'Capsula'],
    color: ['color', 'Color'],
    precio_venta_actual: ['precio_venta_actual', 'Precio', 'precio'],
    estado: ['estado', 'Estado'],
    margen_objetivo: ['margen_objetivo', 'Margen objetivo'],
    observaciones: ['observaciones', 'Observaciones']
  };

  for (var i = info.headerRow + 1; i < data.length; i++) {
    if (String(data[i][colId]).trim() !== id) continue;

    Object.keys(campos).forEach(function(key) {
      if (p[key] === undefined) return;

      var col = colIndex(headers, campos[key]);
      if (col >= 0) {
        s.getRange(i + 1, col + 1).setValue(p[key]);
      }
    });

    actualizarProductoEnStock(p);

    return {
      ok: true,
      id_producto: id
    };
  }

  throw new Error('No encontré el producto: ' + id);
}

function actualizarProductoEnStock(p) {
  var id = String(p.id_producto || '').trim();
  if (!id) return;

  var sh = SS.getSheetByName('STOCK_PRODUCTOS');
  if (!sh) return;

  var info = getHeaderInfo('STOCK_PRODUCTOS');
  var s = info.sheet;
  var data = info.data;
  var headers = info.headers;

  var colId = colIndex(headers, ['id_producto', 'ID', 'id']);
  if (colId < 0) return;

  var map = {
    nombre_producto: ['nombre_producto', 'Producto', 'producto', 'nombre'],
    capsula: ['capsula', 'Cápsula', 'Capsula'],
    precio_venta_actual: ['precio_venta_actual', 'Precio', 'precio'],
    costo_unitario_ref: ['costo_unitario_ref', 'Costo ref', 'costo_unitario']
  };

  for (var i = info.headerRow + 1; i < data.length; i++) {
    if (String(data[i][colId]).trim() !== id) continue;

    Object.keys(map).forEach(function(key) {
      if (p[key] === undefined) return;

      var col = colIndex(headers, map[key]);
      if (col >= 0) {
        s.getRange(i + 1, col + 1).setValue(p[key]);
      }
    });

    return;
  }
}

function eliminarCapsulaCompleta(p) {
  var nombre = String(p.nombre_capsula || '').trim();
  if (!nombre) throw new Error('Falta el nombre de la cápsula');

  // 1. Marcar como eliminadas las filas asociadas
  marcarFilasPorCapsula('PRODUCTOS_MAESTRO', nombre, ['capsula', 'Cápsula', 'Capsula']);
  marcarFilasPorCapsula('STOCK_PRODUCTOS', nombre, ['capsula', 'Cápsula', 'Capsula']);
  marcarFilasPorCapsula('VENTAS', nombre, ['capsula', 'Cápsula', 'Capsula', 'capsula_asociada']);
  marcarFilasPorCapsula('EGRESOS', nombre, ['capsula', 'Cápsula', 'Capsula', 'capsula_asociada']);
  marcarFilasPorCapsula('COMPRAS', nombre, ['capsula_asociada', 'capsula', 'Cápsula', 'Capsula']);
  marcarFilasPorCapsula('VIAJES_NAFTA', nombre, ['capsula_asociada', 'capsula', 'Cápsula', 'Capsula']);
  marcarFilasPorCapsula('LOTES_CUERO', nombre, ['capsula_asociada', 'capsula', 'Cápsula', 'Capsula']);
  marcarFilasPorCapsula('COSTOS_FIJOS', nombre, ['capsula_asociada', 'capsula', 'Cápsula', 'Capsula']);
  marcarFilasPorCapsula('COTIZACIONES', nombre, ['capsula', 'Cápsula', 'Capsula']);

  // 2. Sacar de cápsulas activas
  eliminarCapsula({ nombre_capsula: nombre });

  return {
    ok: true,
    capsula_eliminada: nombre
  };
}

function marcarFilasPorCapsula(nombreHoja, nombreCapsula, posiblesColsCapsula) {
  var sh = SS.getSheetByName(nombreHoja);
  if (!sh) return 0;

  var info = getHeaderInfo(nombreHoja);
  var s = info.sheet;
  var data = info.data;
  var headers = info.headers;

  var colCapsula = colIndex(headers, posiblesColsCapsula);
  if (colCapsula < 0) return 0;

  var colActivo = colIndex(headers, ['activo', 'Activo']);
  var colObs = colIndex(headers, ['observaciones', 'Observaciones', 'obs']);

  var marca = '[ELIMINADO POR CÁPSULA: ' + nombreCapsula + '] ' + new Date();

  var count = 0;

  for (var i = info.headerRow + 1; i < data.length; i++) {
    if (!sameCapsulaFlexible(data[i][colCapsula], nombreCapsula)) continue;

    if (colActivo >= 0) {
      s.getRange(i + 1, colActivo + 1).setValue(false);
    }

    if (colObs >= 0) {
      var obs = String(data[i][colObs] || '');
      s.getRange(i + 1, colObs + 1).setValue(marca + (obs ? ' | ' + obs : ''));
    }

    count++;
  }

  return count;
}

function eliminarCostoFijo(p) {
  var id = String(p.id_costo_fijo || '').trim();
  if (!id) throw new Error('Falta id_costo_fijo');

  var info = getHeaderInfo('COSTOS_FIJOS');
  var s = info.sheet;
  var data = info.data;
  var headers = info.headers;

  var colId = colIndex(headers, ['id_costo_fijo', 'ID', 'id']);
  var colActivo = colIndex(headers, ['activo', 'Activo']);

  var encontrado = false;
  for (var i = info.headerRow + 1; i < data.length; i++) {
    if (String(data[i][colId]).trim() !== id) continue;
    if (colActivo >= 0) {
      s.getRange(i + 1, colActivo + 1).setValue(false);
    }
    encontrado = true;
    break;
  }

  if (!encontrado) throw new Error('No encontré el costo fijo: ' + id);

  // Buscar y desactivar egreso asociado si existe
  try {
    var infoEgr = getHeaderInfo('EGRESOS');
    var sEgr = infoEgr.sheet;
    var dataEgr = infoEgr.data;
    var headersEgr = infoEgr.headers;

    var colObs = colIndex(headersEgr, ['observaciones', 'Observaciones']);
    var colActivoEgr = colIndex(headersEgr, ['activo', 'Activo']);
    var refBuscada = 'Desde costo fijo ' + id;

    for (var j = infoEgr.headerRow + 1; j < dataEgr.length; j++) {
      if (colObs >= 0 && String(dataEgr[j][colObs]).indexOf(refBuscada) >= 0) {
        if (colActivoEgr >= 0) {
          sEgr.getRange(j + 1, colActivoEgr + 1).setValue(false);
        }
        break;
      }
    }
  } catch(e) {
    Logger.log('No se encontró egreso asociado para ' + id + ': ' + e.message);
  }

  return { ok: true, id_costo_fijo: id };
}

// ─── SHOWROOM ──────────────────────────────────────────────

function getShowroom() {
  var s = SS.getSheetByName('SHOWROOM_MOVIMIENTOS');
  if (!s) return { inventario: [], movimientos: [] };

  var movimientos = sheetToArray('SHOWROOM_MOVIMIENTOS');
  var productos = sheetToArray('PRODUCTOS_MAESTRO');

  // Calcular stock showroom por producto Y por talle
  var stockMap = {};

  movimientos.forEach(function(m) {
    var id = String(m.id_producto || '').trim();
    if (!id) return;

    if (!stockMap[id]) {
      stockMap[id] = {
        id_producto: id,
        nombre_producto: m.nombre_producto || '',
        capsula_origen: m.capsula_origen || '',
        stock_showroom: 0,
        talles: {}
      };
    }

    var cantidad = Number(m.cantidad || 0);
    var tipo = String(m.tipo || '').trim();
    var talle = String(m.talle || '').trim() || 'Único';

    if (tipo === 'Ingreso') {
      stockMap[id].stock_showroom += cantidad;
      stockMap[id].talles[talle] = (stockMap[id].talles[talle] || 0) + cantidad;
    } else if (tipo === 'Retiro' || tipo === 'Venta') {
      stockMap[id].stock_showroom -= cantidad;
      stockMap[id].talles[talle] = (stockMap[id].talles[talle] || 0) - cantidad;
    }
  });

  var inventario = Object.values(stockMap)
    .filter(function(item) { return item.stock_showroom > 0; })
    .map(function(item) {
      var prod = null;
      for (var i = 0; i < productos.length; i++) {
        if (String(productos[i].id_producto || '').trim() === item.id_producto) {
          prod = productos[i];
          break;
        }
      }

      // Filtrar talles con stock positivo
      var tallesArray = Object.keys(item.talles)
        .filter(function(t) { return item.talles[t] > 0; })
        .map(function(t) { return { talle: t, cantidad: item.talles[t] }; });

      return {
        id_producto: item.id_producto,
        nombre_producto: item.nombre_producto,
        capsula_origen: item.capsula_origen,
        stock_showroom: item.stock_showroom,
        precio_venta: prod ? Number(prod.precio_venta_actual || 0) : 0,
        valor_exhibido: item.stock_showroom * (prod ? Number(prod.precio_venta_actual || 0) : 0),
        talles: tallesArray
      };
    });

  return {
    inventario: inventario,
    movimientos: movimientos.slice(-50).reverse()
  };
}

function registrarMovimientoShowroom(p) {
  var id = getLastId('SHOWROOM_MOVIMIENTOS', 'SHW-');

  var s = SS.getSheetByName('SHOWROOM_MOVIMIENTOS');
  if (!s) throw new Error('Hoja SHOWROOM_MOVIMIENTOS no encontrada');

  s.appendRow([
    id,
    p.fecha,
    p.tipo,
    p.id_producto,
    p.nombre_producto,
    p.capsula_origen || '',
    Number(p.cantidad) || 1,
    p.motivo || '',
    p.id_venta_ref || '',
    p.fecha_ingreso_ref || '',
    p.talle || ''
  ]);

  return { id: id };
}

function getShowroomKPIs(p) {
  var movimientos = [];
  try { movimientos = sheetToArray('SHOWROOM_MOVIMIENTOS'); } catch(e) {}

  var ventas = getVentas();
  var ventasShowroom = ventas.filter(function(v) {
    return String(val(v, ['lugar_venta', 'Canal']) || '').toLowerCase() === 'showroom';
  });

  var costosFijos = [];
  try { costosFijos = getCostosFijos(); } catch(e) {}

  // Filtrar por período si viene
  var desde = p && p.desde ? new Date(p.desde + 'T00:00:00') : null;
  var hasta = p && p.hasta ? new Date(p.hasta + 'T23:59:59') : null;

  var ventasFiltradas = ventasShowroom.filter(function(v) {
    if (!desde && !hasta) return true;
    var fecha = new Date(val(v, ['fecha', 'Fecha']));
    if (isNaN(fecha)) return false;
    if (desde && fecha < desde) return false;
    if (hasta && fecha > hasta) return false;
    return true;
  });

  var totalVentas = ventasFiltradas.reduce(function(s, v) {
    return s + money(val(v, ['monto', 'Monto']));
  }, 0);

  var ticketPromedio = ventasFiltradas.length ? totalVentas / ventasFiltradas.length : 0;

  // Costo del showroom en el período
  var costoAlquiler = 0;
  costosFijos.forEach(function(cf) {
    var detalle = String(cf.detalle || '').toLowerCase();
    var capsula = String(cf.capsula_asociada || '').toLowerCase();
    if (capsula.indexOf('fragmento') >= 0 || detalle.indexOf('showroom') >= 0 || detalle.indexOf('alquiler') >= 0) {
      var importe = Number(cf.importe || 0);
      var frecuencia = String(cf.frecuencia || '').toLowerCase();
      if (frecuencia === 'mensual') costoAlquiler += importe;
      else costoAlquiler += importe;
    }
  });

  // Stock showroom actual
  var stockMap = {};
  movimientos.forEach(function(m) {
    var id = String(m.id_producto || '').trim();
    if (!id) return;
    if (!stockMap[id]) stockMap[id] = 0;
    var cantidad = Number(m.cantidad || 0);
    var tipo = String(m.tipo || '');
    if (tipo === 'Ingreso') stockMap[id] += cantidad;
    else if (tipo === 'Retiro' || tipo === 'Venta') stockMap[id] -= cantidad;
  });

  var unidadesEnShowroom = Object.values(stockMap).reduce(function(s, v) {
    return s + Math.max(0, v);
  }, 0);

  // Producto más vendido en showroom
  var ventasPorProd = {};
  ventasFiltradas.forEach(function(v) {
    var nombre = String(val(v, ['nombre_producto', 'Producto']) || '—');
    ventasPorProd[nombre] = (ventasPorProd[nombre] || 0) + 1;
  });

  var masVendido = Object.keys(ventasPorProd).sort(function(a, b) {
    return ventasPorProd[b] - ventasPorProd[a];
  })[0] || '—';

  return {
    totalVentas: totalVentas,
    cantidadVentas: ventasFiltradas.length,
    ticketPromedio: ticketPromedio,
    costoAlquiler: costoAlquiler,
    resultadoNeto: totalVentas - costoAlquiler,
    unidadesEnShowroom: unidadesEnShowroom,
    masVendido: masVendido,
    rentable: totalVentas > costoAlquiler
  };
}

// ─── VARIANTES / TALLES ────────────────────────────────────

var SISTEMAS_TALLES = {
  letras:    ['S', 'M', 'L', 'XL', 'XXL'],
  numericos: ['Talle 1', 'Talle 2', 'Talle 3'],
  cintos:    ['110cm', '115cm', '120cm'],
  unico:     ['Único']
};

function getTallesPorSistema(sistema, tallesCustom) {
  if (sistema === 'manual' && tallesCustom) {
    return tallesCustom.split(',').map(function(t) { return t.trim(); }).filter(Boolean);
  }
  return SISTEMAS_TALLES[sistema] || ['Único'];
}

function getVariantes() {
  var s = SS.getSheetByName('STOCK_VARIANTES');
  if (!s) return [];
  return sheetToArray('STOCK_VARIANTES');
}

function getVariantesPorProducto(idProducto) {
  return getVariantes().filter(function(v) {
    return String(v.id_producto || '').trim() === String(idProducto).trim();
  });
}

function actualizarStockVariante(idProducto, talle, delta) {
  var s = SS.getSheetByName('STOCK_VARIANTES');
  if (!s) return;

  var data = s.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });

  var colId     = headers.indexOf('id_producto');
  var colTalle  = headers.indexOf('talle');
  var colStock  = headers.indexOf('stock_actual');
  var colIdVar  = headers.indexOf('id_variante');
  var colNombre = headers.indexOf('nombre_producto');

  // Buscar fila existente
  for (var i = 1; i < data.length; i++) {
    var idProd   = String(data[i][colId]   || '').trim();
    var talleRow = String(data[i][colTalle] || '').trim();

    if (idProd !== String(idProducto).trim()) continue;
    if (talleRow !== String(talle).trim()) continue;

    var stockActual = Number(data[i][colStock]) || 0;
    var nuevoStock  = Math.max(0, stockActual + delta);
    s.getRange(i + 1, colStock + 1).setValue(nuevoStock);
    return;
  }

  // No existe → crear fila nueva
  if (delta <= 0) return;

  var idVar = getLastId('STOCK_VARIANTES', 'VAR-');

  // Obtener nombre del producto
  var prods = sheetToArray('PRODUCTOS_MAESTRO');
  var nombreProducto = '';
  for (var j = 0; j < prods.length; j++) {
    if (String(prods[j].id_producto || '').trim() === String(idProducto).trim()) {
      nombreProducto = prods[j].nombre_producto || '';
      break;
    }
  }

  s.appendRow([idVar, idProducto, nombreProducto, talle, delta]);
}

function registrarProduccionConTalles(p) {
  // p.talles = [{ talle: 'S', cantidad: 3 }, { talle: 'M', cantidad: 2 }]
  var totalProducido = 0;

  (p.talles || []).forEach(function(t) {
    var cantidad = Number(t.cantidad) || 0;
    if (cantidad <= 0) return;

    actualizarStockVariante(p.id_producto, t.talle, cantidad);
    totalProducido += cantidad;
  });

  if (totalProducido <= 0) throw new Error('Ingresá al menos una unidad');

  // Registrar en PRODUCCION
  var sP = SS.getSheetByName('PRODUCCION') || SS.insertSheet('PRODUCCION');
  if (sP.getLastRow() === 0) {
    sP.appendRow(['id_produccion','fecha','id_producto','nombre_producto',
      'cantidad_producida','lote','estado','id_cotizacion_ref',
      'costo_estimado_unitario','observaciones']);
  }

  var id = getLastId('PRODUCCION', 'PRD-P');
  var tallesResumen = (p.talles || [])
    .filter(function(t) { return Number(t.cantidad) > 0; })
    .map(function(t) { return t.talle + ':' + t.cantidad; })
    .join(' | ');

  sP.appendRow([id, p.fecha, p.id_producto, p.nombre_producto,
    totalProducido, p.lote || '', p.estado || 'Completada',
    p.id_cotizacion_ref || '', p.costo_estimado_unitario || 0,
    'Talles: ' + tallesResumen + (p.observaciones ? ' · ' + p.observaciones : '')]);

    // Actualizar STOCK_PRODUCTOS total
  Logger.log('Buscando en STOCK_PRODUCTOS: "' + p.id_producto + '"');
  var info2 = getHeaderInfo('STOCK_PRODUCTOS');
  var colId2 = colIndex(info2.headers, ['id_producto', 'ID']);
  info2.data.forEach(function(row, i) {
    Logger.log('Fila ' + i + ': "' + row[colId2] + '"');
  });
  actualizarStock(p.id_producto, totalProducido, 'produccion');

  // Actualizar STOCK_PRODUCTOS total

  return { id: id, totalProducido: totalProducido };
}

function registrarVentaConTalle(p) {
  // Descontar de variante
  if (p.talle && p.talle !== 'Único') {
    actualizarStockVariante(p.id_producto, p.talle, -1);
  }
  // El resto lo maneja registrarVenta normalmente
  return registrarVenta(p);
}

function getProductosConVariantes() {
  var productos = sheetToArray('PRODUCTOS_MAESTRO');
  var variantes = getVariantes();

  return productos.map(function(p) {
    var idProd = String(p.id_producto || '').trim();
    var vars = variantes.filter(function(v) {
      return String(v.id_producto || '').trim() === idProd;
    });
    p.variantes = vars;
    p.tiene_talles = vars.length > 0;
    return p;
  });
}

// ─── GOOGLE CALENDAR & RECORDATORIOS ──────────────────────

var CALENDAR_ID = 'cueroviejocba@gmail.com';
var EMAIL_RECORDATORIO = 'cueroviejocba@gmail.com';

function crearEventoTarea(tarea) {
  try {
    var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) calendar = CalendarApp.getDefaultCalendar();

    var fechaVenc = tarea.fecha_vencimiento
      ? new Date(tarea.fecha_vencimiento + 'T09:00:00')
      : new Date(new Date().getTime() + 7 * 24 * 60 * 60 * 1000);

    var titulo = '📌 ' + tarea.tarea;
    var descripcion = [
      'Tarea: ' + tarea.tarea,
      'Prioridad: ' + (tarea.prioridad || 'Media'),
      'Responsable: ' + (tarea.responsable || 'Ambos'),
      tarea.descripcion ? 'Descripción: ' + tarea.descripcion : '',
      '',
      'Sistema Cuero Viejo — ID: ' + (tarea.id_tarea || '')
    ].filter(Boolean).join('\n');

    var evento = calendar.createAllDayEvent(titulo, fechaVenc, { description: descripcion });

    // Recordatorios según prioridad
    var minutosAntes = tarea.prioridad === 'Alta' ? 60 : tarea.prioridad === 'Media' ? 120 : 180;
    evento.addEmailReminder(minutosAntes);
    evento.addPopupReminder(30);

    // Guardar ID del evento en la tarea
    var idEvento = evento.getId();
    guardarIdEventoEnTarea(tarea.id_tarea, idEvento);

    return idEvento;
  } catch(e) {
    Logger.log('Error creando evento Calendar: ' + e.message);
    return null;
  }
}

function guardarIdEventoEnTarea(idTarea, idEvento) {
  try {
    var info = getHeaderInfo('TAREAS');
    var s = info.sheet;
    var data = info.data;
    var headers = info.headers;

    var colId = colIndex(headers, ['id_tarea', 'ID', 'id']);
    var colEvento = colIndex(headers, ['id_evento_calendar', 'id_evento']);

    if (colEvento < 0) {
      // La columna no existe, agregarla
      var lastCol = s.getLastColumn();
      s.getRange(info.headerRow + 1, lastCol + 1).setValue('id_evento_calendar');
      colEvento = lastCol;
    }

    for (var i = info.headerRow + 1; i < data.length; i++) {
      if (String(data[i][colId] || '').trim() === String(idTarea).trim()) {
        s.getRange(i + 1, colEvento + 1).setValue(idEvento);
        return;
      }
    }
  } catch(e) {
    Logger.log('Error guardando ID evento: ' + e.message);
  }
}

function actualizarEventoTarea(idTarea, nuevoEstado) {
  try {
    var info = getHeaderInfo('TAREAS');
    var data = info.data;
    var headers = info.headers;

    var colId = colIndex(headers, ['id_tarea', 'ID', 'id']);
    var colEvento = colIndex(headers, ['id_evento_calendar', 'id_evento']);
    var colTarea = colIndex(headers, ['tarea', 'Tarea']);

    if (colEvento < 0) return;

    for (var i = info.headerRow + 1; i < data.length; i++) {
      if (String(data[i][colId] || '').trim() !== String(idTarea).trim()) continue;

      var idEvento = String(data[i][colEvento] || '').trim();
      if (!idEvento) return;

      var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
      if (!calendar) calendar = CalendarApp.getDefaultCalendar();

      var evento = calendar.getEventById(idEvento);
      if (!evento) return;

      if (nuevoEstado === 'Completada') {
        // Cambiar título para marcar como completada
        var tituloActual = evento.getTitle();
        evento.setTitle(tituloActual.replace('📌', '✅'));
        evento.setDescription(evento.getDescription() + '\n\n✅ Completada');
      } else if (nuevoEstado === 'En curso') {
        var tituloActual2 = evento.getTitle();
        evento.setTitle(tituloActual2.replace('📌', '▶').replace('✅', '▶'));
      }

      return;
    }
  } catch(e) {
    Logger.log('Error actualizando evento Calendar: ' + e.message);
  }
}

function enviarRecordatorioPendientes() {
  try {
    var tareas = sheetToArray('TAREAS').filter(function(t) {
      var completada = t.completada;
      var estado = String(t.estado || '');
      if (completada === true || String(completada).toUpperCase() === 'TRUE') return false;
      if (estado === 'Completada') return false;
      return true;
    });

    if (!tareas.length) return;

    var prioColor = { Alta: '#A04040', Media: '#C4882A', Baja: '#5C7A40' };

    var porEstado = { 'En curso': [], 'No iniciada': [] };
    tareas.forEach(function(t) {
      var estado = t.estado || 'No iniciada';
      if (!porEstado[estado]) porEstado[estado] = [];
      porEstado[estado].push(t);
    });

    var html = `
      <div style="font-family:'Georgia',serif;max-width:560px;margin:0 auto;background:#FAF6EF;padding:32px;border-radius:12px">
        <div style="text-align:center;margin-bottom:28px">
          <div style="font-family:'Georgia',serif;font-size:28px;font-weight:300;letter-spacing:4px;color:#4A3728;text-transform:uppercase">
            Cuero Viejo
          </div>
          <div style="font-size:11px;letter-spacing:2px;color:#C4A882;margin-top:4px;text-transform:uppercase">
            Recordatorio de tareas
          </div>
        </div>

        <div style="background:white;border-radius:10px;padding:24px;margin-bottom:16px">
          <div style="font-size:13px;color:#9A836E;margin-bottom:16px">
            Tenés <strong style="color:#4A3728">${tareas.length} tarea${tareas.length !== 1 ? 's' : ''} pendiente${tareas.length !== 1 ? 's' : ''}</strong>
          </div>

          ${porEstado['En curso'].length ? `
            <div style="margin-bottom:20px">
              <div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#C4A882;margin-bottom:10px">En curso</div>
              ${porEstado['En curso'].map(function(t) { return `
                <div style="padding:10px 0;border-bottom:1px solid #F5EDE0;display:flex;align-items:flex-start;gap:10px">
                  <div style="width:3px;min-width:3px;height:36px;background:${prioColor[t.prioridad] || '#C4A882'};border-radius:2px"></div>
                  <div>
                    <div style="font-size:14px;color:#2C1F14;font-weight:500">${t.tarea || ''}</div>
                    <div style="font-size:12px;color:#9A836E;margin-top:2px">
                      ${t.responsable || ''}${t.fecha_vencimiento ? ' · Vence: ' + t.fecha_vencimiento : ''}
                    </div>
                  </div>
                </div>
              `; }).join('')}
            </div>
          ` : ''}

          ${porEstado['No iniciada'].length ? `
            <div>
              <div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#C4A882;margin-bottom:10px">Pendientes</div>
              ${porEstado['No iniciada'].map(function(t) { return `
                <div style="padding:10px 0;border-bottom:1px solid #F5EDE0;display:flex;align-items:flex-start;gap:10px">
                  <div style="width:3px;min-width:3px;height:36px;background:${prioColor[t.prioridad] || '#C4A882'};border-radius:2px"></div>
                  <div>
                    <div style="font-size:14px;color:#2C1F14;font-weight:500">${t.tarea || ''}</div>
                    <div style="font-size:12px;color:#9A836E;margin-top:2px">
                      ${t.responsable || ''}${t.fecha_vencimiento ? ' · Vence: ' + t.fecha_vencimiento : ''}
                    </div>
                  </div>
                </div>
              `; }).join('')}
            </div>
          ` : ''}
        </div>

        <div style="text-align:center;font-size:11px;color:#9A836E;letter-spacing:1px">
          CUERO VIEJO · SISTEMA DE GESTIÓN
        </div>
      </div>
    `;

    GmailApp.sendEmail(
      EMAIL_RECORDATORIO,
      '📌 Cuero Viejo — ' + tareas.length + ' tarea' + (tareas.length !== 1 ? 's' : '') + ' pendiente' + (tareas.length !== 1 ? 's' : ''),
      'Tenés ' + tareas.length + ' tareas pendientes en el sistema.',
      { htmlBody: html }
    );

    Logger.log('Recordatorio enviado: ' + tareas.length + ' tareas');
  } catch(e) {
    Logger.log('Error enviando recordatorio: ' + e.message);
  }
}

function configurarTriggerRecordatorio() {
  // Eliminar triggers anteriores del mismo nombre
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'enviarRecordatorioPendientes') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Crear trigger cada 2 días a las 9am
  ScriptApp.newTrigger('enviarRecordatorioPendientes')
    .timeBased()
    .everyDays(2)
    .atHour(9)
    .create();

  Logger.log('Trigger configurado: recordatorio cada 2 días a las 9am');
}

function getShowroomHistorial() {
  var movimientos = [];
  try { movimientos = sheetToArray('SHOWROOM_MOVIMIENTOS'); } catch(e) {}

  var ventas = getVentas();
  var ventasShowroom = ventas.filter(function(v) {
    return String(val(v, ['lugar_venta', 'Canal']) || '').toLowerCase() === 'showroom';
  });

  var costosFijos = [];
  try { costosFijos = getCostosFijos(); } catch(e) {}

  // Calcular alquiler mensual
  var alquilerMensual = 0;
  costosFijos.forEach(function(cf) {
    var detalle = String(cf.detalle || '').toLowerCase();
    var capsula = String(cf.capsula_asociada || '').toLowerCase();
    if (capsula.indexOf('fragmento') >= 0 || detalle.indexOf('showroom') >= 0 || detalle.indexOf('alquiler') >= 0) {
      var frecuencia = String(cf.frecuencia || '').toLowerCase();
      if (frecuencia === 'mensual') alquilerMensual += Number(cf.importe || 0);
    }
  });

  // Agrupar ventas por mes
  var mesesMap = {};

  ventasShowroom.forEach(function(v) {
    var fecha = val(v, ['fecha', 'Fecha']);
    if (!fecha) return;
    var d = new Date(fecha);
    if (isNaN(d)) return;
    var k = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);

    if (!mesesMap[k]) {
      mesesMap[k] = {
        mes: k,
        ventas: 0,
        cantidad: 0,
        alquiler: alquilerMensual,
        resultado: 0,
        rentable: false,
        productos: {}
      };
    }

    var monto = money(val(v, ['monto', 'Monto']));
    var nombre = String(val(v, ['nombre_producto', 'Producto']) || '—');

    mesesMap[k].ventas += monto;
    mesesMap[k].cantidad += 1;
    mesesMap[k].productos[nombre] = (mesesMap[k].productos[nombre] || 0) + 1;
  });

  // Calcular resultado y top producto por mes
  var ahora = new Date();
  var mesActual = ahora.getFullYear() + '-' + ('0' + (ahora.getMonth() + 1)).slice(-2);

  var historial = Object.keys(mesesMap).sort().map(function(k) {
    var m = mesesMap[k];
    m.resultado = m.ventas - m.alquiler;
    m.rentable = m.resultado >= 0;
    m.cerrado = k < mesActual;

    // Top producto
    var topProd = Object.keys(m.productos).sort(function(a, b) {
      return m.productos[b] - m.productos[a];
    })[0] || '—';
    m.top_producto = topProd;
    m.top_producto_cant = m.productos[topProd] || 0;

    // Movimientos del mes
    var movsMes = movimientos.filter(function(mov) {
      var fecha = val(mov, ['fecha', 'Fecha']);
      if (!fecha) return false;
      var d = new Date(fecha);
      if (isNaN(d)) return false;
      var km = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
      return km === k;
    });

    m.movimientos = movsMes.length;
    return m;
  });

  // Si el mes actual no tiene ventas todavía, agregarlo igual
  if (!mesesMap[mesActual]) {
    historial.push({
      mes: mesActual,
      ventas: 0,
      cantidad: 0,
      alquiler: alquilerMensual,
      resultado: -alquilerMensual,
      rentable: false,
      cerrado: false,
      top_producto: '—',
      top_producto_cant: 0,
      movimientos: 0
    });
  }

  return {
    historial: historial.reverse(), // más reciente primero
    alquilerMensual: alquilerMensual,
    mesActual: mesActual
  };
}

function onFormSubmit(e) {
  var row = e.values;
  
  var nombre          = row[1] || '';
  var contacto        = row[2] || '';
  var instagram       = row[3] || '';
  var tipo_cuero      = row[4] || '';
  var color           = row[5] || '';
  var fecha_necesaria = row[6] || '';
  var referencia      = row[7] || '';

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Crear tarea en TAREAS
  var sheetTodo = ss.getSheetByName('TAREAS');
  if (sheetTodo) {
    var lastIdTarea = sheetTodo.getLastRow();
    var idTarea = 'TSK-' + String(lastIdTarea).padStart(3, '0');

    sheetTodo.appendRow([
      idTarea,
      'Hablar con ' + nombre,
      'WhatsApp: ' + contacto +
      (instagram ? ' | IG: ' + instagram : '') +
      (tipo_cuero ? ' | Cuero: ' + tipo_cuero : '') +
      (color ? ' | Color: ' + color : '') +
      (referencia ? ' | Ref: ' + referencia : ''),
      'Alta',
      'No iniciada',
      Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'),
      fecha_necesaria,
      'Ambos',
      false,
      'Set personalizado',
      ''
    ]);
    Logger.log('Tarea creada: ' + idTarea);
  } else {
    Logger.log('ERROR: No encontré la hoja TAREAS');
  }

  // 2. Registrar en SETS
  var sheetSets = ss.getSheetByName('SETS_PERSONALIZADOS');
  if (sheetSets) {
    var lastIdSet = sheetSets.getLastRow();
    var idSet = 'SET-' + String(lastIdSet).padStart(3, '0');

    sheetSets.appendRow([
      idSet,
      Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'),
      nombre,
      contacto,
      instagram,
      tipo_cuero,
      color,
      fecha_necesaria,
      'Consulta',
      referencia,
      ''
    ]);
    Logger.log('Set creado: ' + idSet);
  } else {
    Logger.log('ERROR: No encontré la hoja SETS_PERSONALIZADOS');
  }

}

function debugProducto(p) {
  var s = SS.getSheetByName('PRODUCTOS_MAESTRO');
  var data = s.getDataRange().getValues();
  for (var i = 0; i < Math.min(data.length, 6); i++) {
    Logger.log('Fila ' + (i+1) + ': ' + data[i].join(' | '));
  }
  return true;
}


// ─── EVENTOS DE GOOGLE CALENDAR ───────────────────────────
// Trae eventos del mes pedido (o del mes actual si no se especifica)
// payload: { year: 2025, month: 6 }  (month = 1-12)
function getEventosCalendar(p) {
  try {
    var now = new Date();
    var year  = parseInt(p.year  || now.getFullYear(), 10);
    var month = parseInt(p.month || (now.getMonth() + 1), 10); // 1-based

    var inicio = new Date(year, month - 1, 1, 0, 0, 0);
    var fin    = new Date(year, month,     0, 23, 59, 59); // último día del mes

    var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) calendar = CalendarApp.getDefaultCalendar();

    var eventos = calendar.getEvents(inicio, fin);

    var lista = eventos.map(function(ev) {
      var start = ev.getStartTime();
      var end   = ev.getEndTime();
      return {
        id:     ev.getId(),
        titulo: ev.getTitle(),
        fecha:  Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        hora:   ev.isAllDayEvent() ? null : Utilities.formatDate(start, Session.getScriptTimeZone(), 'HH:mm'),
        todo_el_dia: ev.isAllDayEvent(),
        descripcion: ev.getDescription() || '',
        color:  ev.getColor() || ''
      };
    });

    return { ok: true, data: lista };
  } catch(e) {
    Logger.log('Error getEventosCalendar: ' + e.message);
    return { ok: false, error: e.message, data: [] };
  }
}
