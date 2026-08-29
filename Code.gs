/**
 * MENU DIGITAL - Backend con Google Sheets
 * =========================================
 * Pegá TODO este codigo en: Google Sheets -> Extensiones -> Apps Script
 * Despues: Implementar -> Nueva implementacion -> App web
 *   - Ejecutar como: Yo (tu cuenta)
 *   - Quien tiene acceso: Cualquier usuario
 * Implementas y copias la URL que te da (terminada en /exec).
 * Esa URL es la API_URL que va en el frontend.
 */

var HOJA_MENU = 'Menu';
var HOJA_PEDIDOS = 'Pedidos';
var HOJA_CONFIG = 'Config';

function doGet(e) {
  var accion = e && e.parameter && e.parameter.accion || 'menu';
  var out = { ok: true };
  try {
    if (accion === 'menu') {
      out.data = leerMenu();
    } else if (accion === 'pedidos') {
      out.data = leerPedidos(e.parameter.estado || '');
    } else if (accion === 'config') {
      out.data = leerConfig();
    } else if (accion === 'resumen') {
      out.data = resumenVentas(e.parameter);
    } else {
      out.ok = false; out.error = 'accion desconocida';
    }
  } catch (err) {
    out.ok = false; out.error = String(err);
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var out = { ok: true };
  try {
    var body = JSON.parse(e.postData.contents);
    var accion = body.accion;
    if (accion === 'crear_pedido') {
      out.data = crearPedido(body);
    } else if (accion === 'estado_pedido') {
      out.data = cambiarEstadoPedido(body);
    } else if (accion === 'guardar_producto') {
      out.data = guardarProducto(body);
    } else if (accion === 'borrar_producto') {
      out.data = borrarProducto(body);
    } else if (accion === 'agregar_categoria') {
      out.data = agregarCategoria(body);
    } else if (accion === 'borrar_categoria') {
      out.data = borrarCategoria(body);
    } else {
      out.ok = false; out.error = 'accion desconocida';
    }
  } catch (err) {
    out.ok = false; out.error = String(err);
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- utilidades ----------
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function asegurarPlanilla() {
  var s = ss();
  [[HOJA_MENU,['id','categoria','nombre','descripcion','precio','imagen','disponible']],
   [HOJA_PEDIDOS,['id','numero','nombre','telefono','items','total','nota','pago_metodo','pago_captura','pago_confirmado','estado','creado_en']],
   [HOJA_CONFIG,['clave','valor']]
  ].forEach(function(def){
    var nom = def[0]; var cols = def[1];
    var sh = s.getSheetByName(nom);
    if (!sh) { sh = s.insertSheet(nom); }
    if (sh.getLastRow() === 0) {
      sh.appendRow(cols);
    } else {
      // asegurar encabezados
      cols.forEach(function(c,i){ if (!sh.getRange(1,i+1).getValue()) sh.getRange(1,i+1).setValue(c); });
    }
  });
}

// ---------- MENU ----------
function leerMenu() {
  asegurarPlanilla();
  var sh = ss().getSheetByName(HOJA_MENU);
  var data = sh.getDataRange().getValues();
  var cats = {}; var orden = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[1]) continue; // sin categoria
    var cat = String(r[1]);
    if (!cats[cat]) { cats[cat] = { nombre: cat, productos: [] }; orden.push(cat); }
    cats[cat].productos.push({
      id: r[0], categoria: cat, nombre: r[2], descripcion: r[3],
      precio: Number(r[4]), imagen: r[5], disponible: r[6] === 0 ? 0 : 1
    });
  }
  return orden.map(function(c){ return cats[c]; });
}

function proximoIdMenu() {
  var sh = ss().getSheetByName(HOJA_MENU);
  var data = sh.getDataRange().getValues();
  var max = 100;
  for (var i = 1; i < data.length; i++) if (Number(data[i][0]) > max) max = Number(data[i][0]);
  return max + 1;
}

function guardarProducto(b) {
  asegurarPlanilla();
  var sh = ss().getSheetByName(HOJA_MENU);
  var id = b.id;
  var fila = null;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (Number(data[i][0]) === Number(id)) fila = i + 1; }
  var disp = (b.disponible === undefined || b.disponible === true || b.disponible === 1) ? 1 : 0;
  var nuevo = [ id ? Number(id) : proximoIdMenu(), b.categoria, b.nombre, b.descripcion || '', Number(b.precio), b.imagen || '', disp ];
  if (fila) sh.getRange(fila,1,1,nuevo.length).setValues([nuevo]);
  else sh.appendRow(nuevo);
  return { id: nuevo[0] };
}

function borrarProducto(b) {
  asegurarPlanilla();
  var sh = ss().getSheetByName(HOJA_MENU);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(b.id)) sh.deleteRow(i + 1);
  }
  return { ok: true };
}

function agregarCategoria(b) { return { ok: true }; } // categorias = valor de la columna, no requiere tabla aparte
function borrarCategoria(b) { return { ok: true }; }

// ---------- PEDIDOS ----------
function leerPedidos(estado) {
  asegurarPlanilla();
  var sh = ss().getSheetByName(HOJA_PEDIDOS);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    var ped = { id: r[0], numero: r[1], nombre: r[2], telefono: r[3], items: r[4], total: r[5], nota: r[6], pago_metodo: r[7], pago_captura: r[8], pago_confirmado: r[9], estado: r[10], creado_en: r[11] };
    if (!estado || ped.estado === estado) out.push(ped);
  }
  out.sort(function(a,b){ return Number(b.id) - Number(a.id); });
  return out;
}

function proximoNumero() {
  var sh = ss().getSheetByName(HOJA_PEDIDOS);
  var data = sh.getDataRange().getValues();
  var hoy = fechaLocal();
  var max = 0;
  for (var i = 1; i < data.length; i++) {
    var creado = String(data[i][11]);
    if (creado && creado.indexOf(hoy) === 0 && Number(data[i][1]) > max) max = Number(data[i][1]);
  }
  return max + 1;
}

function fechaLocal() {
  return Utilities.formatDate(new Date(), ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function crearPedido(b) {
  asegurarPlanilla();
  var sh = ss().getSheetByName(HOJA_PEDIDOS);
  var numero = proximoNumero();
  var id = proximoIdPedidos();
  sh.appendRow([ id, numero, b.nombre || '', b.telefono || '', JSON.stringify(b.items || []), Number(b.total) || 0, b.nota || '', b.pago_metodo || '', b.pago_captura || '', b.pago_confirmado ? 1 : 0, 'pendiente', fechaLocal() ]);
  return { id: id, numero: numero };
}

function proximoIdPedidos() {
  var sh = ss().getSheetByName(HOJA_PEDIDOS);
  var data = sh.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < data.length; i++) if (Number(data[i][0]) > max) max = Number(data[i][0]);
  return max + 1;
}

function cambiarEstadoPedido(b) {
  asegurarPlanilla();
  var sh = ss().getSheetByName(HOJA_PEDIDOS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(b.id)) {
      if (b.estado) sh.getRange(i+1, 11).setValue(b.estado);
      if (b.pago_confirmado !== undefined) sh.getRange(i+1, 10).setValue(b.pago_confirmado ? 1 : 0);
      return { ok: true };
    }
  }
  return { ok: false, error: 'no existe' };
}

// ---------- CONFIG ----------
function leerConfig() {
  asegurarPlanilla();
  var sh = ss().getSheetByName(HOJA_CONFIG);
  var data = sh.getDataRange().getValues();
  var cfg = { negocio: 'Mi Negocio', wppAdmin: '' };
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'negocio') cfg.negocio = data[i][1];
    if (data[i][0] === 'wppAdmin') cfg.wppAdmin = data[i][1];
  }
  if (!cfg.wppAdmin) { sh.appendRow(['wppAdmin','']); }
  if (!cfg.negocio) { sh.appendRow(['negocio','Mi Negocio']); }
  return cfg;
}

// ---------- RESUMEN DE VENTAS ----------
function resumenVentas(e) {
  var sh = ss().getSheetByName(HOJA_PEDIDOS);
  var data = sh.getDataRange().getValues();
  var hoy = fechaLocal().slice(0,10);
  var desde, hasta;
  if (e.fecha) { desde = e.fecha; hasta = e.fecha; }
  else if (e.desde && e.hasta) { desde = e.desde; hasta = e.hasta; }
  else { desde = hoy; hasta = hoy; }

  var totalGeneral = 0, totalPagar = 0, totalSinPagar = 0, pedidosTotales = 0, itemsVendidos = 0;
  var porMetodo = {}, porProducto = {};

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (r[10] !== 'entregado') continue;
    var fecha = String(r[11]).slice(0,10);
    if (fecha < desde || fecha > hasta) continue;
    pedidosTotales++;
    totalGeneral += Number(r[5]) || 0;
    if (r[9]) totalPagar += Number(r[5]) || 0; else totalSinPagar += Number(r[5]) || 0;
    var met = String(r[7]) || 's/n';
    porMetodo[met] = (porMetodo[met] || 0) + (Number(r[5]) || 0);
    var items = [];
    try { items = JSON.parse(r[4] || '[]'); } catch(err) {}
    items.forEach(function(it){
      itemsVendidos += Number(it.cantidad) || 0;
      if (!porProducto[it.nombre]) porProducto[it.nombre] = { cantidad: 0, monto: 0 };
      porProducto[it.nombre].cantidad += Number(it.cantidad) || 0;
      porProducto[it.nombre].monto += Number(it.subtotal || (it.precio * it.cantidad) || 0);
    });
  }

  return {
    pedidosTotales: pedidosTotales, itemsVendidos: itemsVendidos,
    totalGeneral: totalGeneral, totalPagar: totalPagar, totalSinPagar: totalSinPagar,
    porMetodo: Object.keys(porMetodo).map(function(k){ return { metodo: k, monto: porMetodo[k] }; }),
    porProducto: Object.keys(porProducto).map(function(k){ return { nombre: k, cantidad: porProducto[k].cantidad, monto: porProducto[k].monto }; })
      .sort(function(a,b){ return b.cantidad - a.cantidad; })
  };
}
