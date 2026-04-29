// ============================================================
// CARRANZA BARRIENTOS, S.A.S. DE C.V.
// Sistema de Gestión — Google Apps Script Backend
// Versión 2.0 — 2026
// ============================================================

const SHEETS = {
  PRODUCTS:       'Productos',
  SALES:          'Ventas',
  CONFIG:         'Configuracion',
  CATEGORIES:     'Categorias',
  LOG:            'Log',
  CLIENTS:        'Clientes',
  CONTACTS:       'Contactos',
  QUOTES:         'Cotizaciones',
  QUOTE_ITEMS:    'CotizacionItems',
  PURCHASES:      'Compras',
  PURCHASE_ITEMS: 'CompraItems',
  CXC:            'CuentasCobrar'
};

// ============================================================
// ENRUTADOR PRINCIPAL
// ============================================================
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  var params   = e.parameter || {};
  var postData = e.postData ? JSON.parse(e.postData.contents || '{}') : {};
  var data     = Object.assign({}, params, postData);
  var action   = data.action;
  var callback = data.callback;
  var result;

  try {
    // --- PRODUCTOS ---
    if      (action === 'getAll')            result = getAllData();
    else if (action === 'getProducts')       result = getProducts();
    else if (action === 'saveProduct')       result = saveProduct(data);
    else if (action === 'deleteProduct')     result = deleteProduct(data.id);
    else if (action === 'updateStock')       result = updateStock(data);
    else if (action === 'reactivate')        result = reactivateProducts();

    // --- VENTAS ---
    else if (action === 'recordSale')        result = recordSale(data);
    else if (action === 'getSales')          result = getSales(data);
    else if (action === 'deleteSale')        result = deleteSale(data.id);

    // --- CATEGORÍAS ---
    else if (action === 'getCategories')     result = getCategories();
    else if (action === 'saveCategory')      result = saveCategory(data);
    else if (action === 'deleteCategory')    result = deleteCategory(data.name);

    // --- CONFIGURACIÓN Y USUARIOS ---
    else if (action === 'getConfig')         result = getConfig();
    else if (action === 'saveConfig')        result = saveConfig(data);
    else if (action === 'getUsers')          result = getUsers();
    else if (action === 'saveUsers')         result = saveUsers(data);

    // --- CLIENTES ---
    else if (action === 'getClients')        result = getClients();
    else if (action === 'saveClient')        result = saveClient(data);
    else if (action === 'deleteClient')      result = deleteClient(data.id);

    // --- CONTACTOS POR CLIENTE ---
    else if (action === 'getContacts')       result = getContacts(data.cliente_id);
    else if (action === 'saveContact')       result = saveContact(data);
    else if (action === 'deleteContact')     result = deleteContact(data.id);

    // --- COTIZACIONES ---
    else if (action === 'getQuotes')         result = getQuotes();
    else if (action === 'saveQuote')         result = saveQuote(data);
    else if (action === 'deleteQuote')       result = deleteQuote(data.id);
    else if (action === 'getQuoteItems')     result = getQuoteItems(data.cotizacion_id);
    else if (action === 'saveQuoteItems')    result = saveQuoteItems(data);
    else if (action === 'getNextQuoteNum')   result = getNextQuoteNum();

    // --- COMPRAS ---
    else if (action === 'getPurchases')      result = getPurchases();
    else if (action === 'savePurchase')      result = savePurchase(data);
    else if (action === 'deletePurchase')    result = deletePurchase(data.id);
    else if (action === 'getPurchaseItems')  result = getPurchaseItems(data.compra_id);

    // --- CUENTAS POR COBRAR ---
    else if (action === 'getCxC')            result = getCxC();
    else if (action === 'saveCxC')           result = saveCxC(data);
    else if (action === 'updateCxC')         result = updateCxC(data);
    else if (action === 'deleteCxC')         result = deleteCxC(data.id);

    // --- MANTENIMIENTO ---
    else if (action === 'setup')             result = setupSheets();
    else if (action === 'repair')            result = repairAllSheets();

    else result = { success: false, error: 'Accion no reconocida: ' + action };

  } catch (err) {
    result = { success: false, error: err.toString() };
    logError(action, err.toString());
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SETUP Y MANTENIMIENTO
// ============================================================
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = getSheetDefinitions();

  defs.forEach(function(def) {
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.appendRow(def.headers);
      sheet.getRange(1, 1, 1, def.headers.length)
           .setFontWeight('bold')
           .setBackground('#1a3a5c')
           .setFontColor('#ffffff');
      // Datos iniciales de configuración
      if (def.name === 'Configuracion') {
        sheet.appendRow(['meta_diaria','500']);
        sheet.appendRow(['meta_mensual','10000']);
        sheet.appendRow(['stock_minimo_global','3']);
        sheet.appendRow(['moneda','USD']);
        sheet.appendRow(['nombre_empresa','CARRANZA BARRIENTOS, S.A.S. DE C.V.']);
        sheet.appendRow(['direccion','Residencial Altos de San Ernesto #51, Final Calle Motocross, El Salvador']);
        sheet.appendRow(['telefono1','(503) 6198-7622']);
        sheet.appendRow(['telefono2','(503) 7499-8889']);
        sheet.appendRow(['email','carrba.sv@gmail.com']);
        sheet.appendRow(['nit','0623-010725-118-6']);
        sheet.appendRow(['nrc','367151-0']);
        sheet.appendRow(['banco','Banco Agrícola']);
        sheet.appendRow(['cuenta_bancaria','005490414780']);
        sheet.appendRow(['correlativo_cotizacion','250']);
        sheet.appendRow(['forma_pago_default','50% anticipo / 50% contra entrega']);
        sheet.appendRow(['tiempo_entrega_default','15 días hábiles']);
        sheet.appendRow(['validez_cotizacion','15 días calendario']);
        sheet.appendRow(['porcentaje_ganancia_default','40']);
        sheet.appendRow(['asesor_default','Jesús Cristóbal Carranza']);
      }
      if (def.name === 'Categorias') {
        sheet.appendRow(['Botas Nacionales','Botas de seguridad fabricación nacional','false',new Date().toISOString()]);
        sheet.appendRow(['Botas Duramax','Botas de seguridad Duramax importadas','false',new Date().toISOString()]);
        sheet.appendRow(['Camisas','Camisas manga larga y corta','true',new Date().toISOString()]);
        sheet.appendRow(['Camisetas','Camisetas de algodón','true',new Date().toISOString()]);
        sheet.appendRow(['Gabachas','Gabachas de trabajo','true',new Date().toISOString()]);
        sheet.appendRow(['Pantalones','Pantalones de trabajo','true',new Date().toISOString()]);
        sheet.appendRow(['Accesorios','Accesorios varios','false',new Date().toISOString()]);
        sheet.appendRow(['Otros','Otros productos','false',new Date().toISOString()]);
      }
    }
  });

  // Usuarios por defecto
  saveDefaultUsers();
  return { success: true, message: 'Sistema configurado correctamente' };
}

function saveDefaultUsers() {
  var sheet = getSheet('Configuracion');
  var rows = sheet.getDataRange().getValues();
  var hasUsers = rows.some(function(r){ return String(r[0]).startsWith('user_'); });
  if (!hasUsers) {
    var users = [
      { user:'vcarranza1', pass:'v1ct@r',  name:'Victor Carranza',   role:'Maestro' },
      { user:'acarranza',  pass:'4le3105', name:'A. Carranza',        role:'Operador' },
      { user:'mcarranza',  pass:'M4rt@',   name:'M. Carranza',        role:'Operador' },
      { user:'ccarranza',  pass:'Dog1t0',  name:'C. Carranza',        role:'Operador' }
    ];
    users.forEach(function(u, i) {
      sheet.appendRow(['user_' + i, JSON.stringify(u)]);
    });
  }
}

function repairAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = getSheetDefinitions();
  defs.forEach(function(def) {
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) return;
    var rows = sheet.getDataRange().getValues();
    if (rows.length === 0) { sheet.appendRow(def.headers); return; }
    if (String(rows[0][0]).toLowerCase().trim() !== def.headers[0].toLowerCase()) {
      sheet.insertRowBefore(1);
      sheet.getRange(1,1,1,def.headers.length).setValues([def.headers]);
      sheet.getRange(1,1,1,def.headers.length)
           .setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
    }
  });
  return { success: true, message: 'Hojas reparadas' };
}

function getSheetDefinitions() {
  return [
    { name:'Productos',       headers:['id','nombre','categoria','precio_venta','costo','stock','stock_minimo','tallas','imagen_url','notas','fecha_creacion','activo'] },
    { name:'Ventas',          headers:['id','fecha','producto_id','producto_nombre','categoria','cantidad','precio_unitario','costo_unitario','total','ganancia','metodo_pago','notas','registrado_por'] },
    { name:'Configuracion',   headers:['clave','valor'] },
    { name:'Categorias',      headers:['nombre','descripcion','tiene_tallas','fecha_creacion'] },
    { name:'Log',             headers:['fecha','accion','detalle','error'] },
    { name:'Clientes',        headers:['id','nombre_empresa','tipo_factura','tipo_factura_otro','nit','registro','direccion','telefono','email','notas','fecha_creacion','activo'] },
    { name:'Contactos',       headers:['id','cliente_id','nombre','telefono','email','cargo','notas','fecha_creacion'] },
    { name:'Cotizaciones',    headers:['id','numero','fecha','cliente_id','cliente_nombre','asesor','valido_hasta','subtotal','iva_monto','total','estado','metodo_pago','tiempo_entrega','notas','registrado_por','fecha_creacion'] },
    { name:'CotizacionItems', headers:['id','cotizacion_id','producto_id','descripcion','imagen_url','cantidad','precio_unit','incluye_iva','subtotal_item'] },
    { name:'Compras',         headers:['id','fecha','proveedor','factura_ref','subtotal','iva_monto','total','notas','registrado_por','fecha_creacion'] },
    { name:'CompraItems',     headers:['id','compra_id','producto_id','producto_nombre','cantidad','costo_unit','total_item'] },
    { name:'CuentasCobrar',   headers:['id','cotizacion_num','cliente_id','cliente_nombre','monto_total','anticipo_recibido','saldo_pendiente','fecha_vencimiento','estado','notas','fecha_creacion'] }
  ];
}

function getAllData() {
  ensureSheets();
  return {
    success:     true,
    products:    getProducts().data,
    sales:       getSales({}).data,
    config:      getConfig().data,
    categories:  getCategories().data,
    clients:     getClients().data,
    quotes:      getQuotes().data,
    purchases:   getPurchases().data,
    cxc:         getCxC().data
  };
}

// ============================================================
// PRODUCTOS
// ============================================================
function getProducts() {
  var sheet = getSheet('Productos');
  var rows  = sheet.getDataRange().getValues();
  if (rows.length === 0) return { success:true, data:[] };
  var EXPECTED = ['id','nombre','categoria','precio_venta','costo','stock','stock_minimo','tallas','imagen_url','notas','fecha_creacion','activo'];
  var hasH = String(rows[0][0]).toLowerCase().trim() === 'id';
  var headers = hasH ? rows[0] : EXPECTED;
  var dataRows = hasH ? rows.slice(1) : rows;
  if (!hasH) {
    sheet.insertRowBefore(1);
    sheet.getRange(1,1,1,EXPECTED.length).setValues([EXPECTED]);
    sheet.getRange(1,1,1,EXPECTED.length).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
  }
  var products = dataRows.map(function(row) {
    var obj = {}; headers.forEach(function(h,i){ obj[h]=row[i]; });
    if (obj.tallas && typeof obj.tallas==='string' && obj.tallas.trim().startsWith('{')) {
      try { obj.tallas=JSON.parse(obj.tallas); } catch(e) { obj.tallas={}; }
    } else { obj.tallas={}; }
    return obj;
  }).filter(function(p) {
    if (!p.id || String(p.id).trim()==='') return false;
    var a = String(p.activo).toUpperCase().trim();
    return a!=='FALSE' && a!=='0';
  });
  return { success:true, data:products };
}

function saveProduct(data) {
  var sheet = getSheet('Productos');
  var rows  = sheet.getDataRange().getValues();
  var costo  = parseFloat(data.costo)||0;
  var precio = parseFloat(data.precio_venta)||0;
  var tallas = data.tallas ? (typeof data.tallas==='string' ? data.tallas : JSON.stringify(data.tallas)) : '{}';
  if (data.id && String(data.id).trim()!=='') {
    for (var i=1; i<rows.length; i++) {
      if (String(rows[i][0]).trim()===String(data.id).trim()) {
        sheet.getRange(i+1,1,1,12).setValues([[data.id,data.nombre,data.categoria,precio,costo,parseInt(data.stock)||0,parseInt(data.stock_minimo)||3,tallas,data.imagen_url||'',data.notas||'',rows[i][10],true]]);
        return { success:true, message:'Actualizado', id:data.id };
      }
    }
  }
  var newId = 'P'+Date.now();
  sheet.appendRow([newId,data.nombre,data.categoria,precio,costo,parseInt(data.stock)||0,parseInt(data.stock_minimo)||3,tallas,data.imagen_url||'',data.notas||'',new Date().toISOString(),true]);
  return { success:true, message:'Creado', id:newId };
}

function deleteProduct(id) {
  var sheet = getSheet('Productos');
  var rows  = sheet.getDataRange().getValues();
  for (var i=rows.length-1; i>=1; i--) {
    if (String(rows[i][0]).trim()===String(id).trim()) {
      sheet.getRange(i+1,12).setValue('FALSE');
      return { success:true };
    }
  }
  return { success:false, error:'Producto no encontrado: '+id };
}

function reactivateProducts() {
  var sheet = getSheet('Productos');
  var rows  = sheet.getDataRange().getValues();
  var count = 0;
  for (var i=1; i<rows.length; i++) {
    var activo = String(rows[i][11]).toUpperCase().trim();
    if (activo==='FALSE' || activo==='0') {
      sheet.getRange(i+1,12).setValue(true);
      count++;
    }
  }
  return { success:true, message:'Reactivados: '+count };
}

function updateStock(data) {
  var sheet = getSheet('Productos');
  var rowIdx = findProductRowIndex(sheet, data.id);
  if (rowIdx < 1) return { success:false, error:'Producto no encontrado' };
  if (data.talla) {
    var tv = sheet.getRange(rowIdx,8).getValue();
    var tobj = {}; try { tobj=JSON.parse(String(tv)); } catch(e) {}
    tobj[data.talla] = parseInt(data.stock)||0;
    sheet.getRange(rowIdx,8).setValue(JSON.stringify(tobj));
    var ts=0; Object.keys(tobj).forEach(function(t){ts+=parseInt(tobj[t])||0;});
    sheet.getRange(rowIdx,6).setValue(ts);
  } else {
    sheet.getRange(rowIdx,6).setValue(parseInt(data.stock)||0);
  }
  return { success:true };
}

// ============================================================
// VENTAS
// ============================================================
function recordSale(data) {
  var ss = getSheet('Ventas');
  var ps = getSheet('Productos');
  var items = typeof data.items==='string' ? JSON.parse(data.items) : (data.items||[]);
  var fecha = new Date().toISOString();
  var saleId = 'V'+Date.now();
  items.forEach(function(item) {
    var qty  = parseInt(item.cantidad)||1;
    var prod = findProduct(ps, item.producto_id) || {};
    var pu   = parseFloat(item.precio_unitario)||parseFloat(prod.precio_venta)||0;
    var cu   = parseFloat(prod.costo)||0;
    ss.appendRow([saleId+'_'+item.producto_id, fecha, item.producto_id, prod.nombre||item.descripcion, prod.categoria||'', qty, pu, cu, pu*qty, (pu-cu)*qty, data.metodo_pago||'Efectivo', data.notas||'', data.registrado_por||'Sistema']);
    var rowIdx = findProductRowIndex(ps, item.producto_id);
    if (rowIdx > 0) {
      if (item.talla) {
        var tv = ps.getRange(rowIdx,8).getValue();
        var tobj = {}; try { tobj=JSON.parse(String(tv)); } catch(e) {}
        tobj[item.talla] = Math.max(0,(parseInt(tobj[item.talla])||0)-qty);
        ps.getRange(rowIdx,8).setValue(JSON.stringify(tobj));
        var ts=0; Object.keys(tobj).forEach(function(t){ts+=parseInt(tobj[t])||0;});
        ps.getRange(rowIdx,6).setValue(ts);
      } else {
        var cur = parseInt(ps.getRange(rowIdx,6).getValue())||0;
        ps.getRange(rowIdx,6).setValue(Math.max(0,cur-qty));
      }
    }
  });
  return { success:true, saleId:saleId };
}

function getSales(data) {
  var HEADERS = ['id','fecha','producto_id','producto_nombre','categoria','cantidad','precio_unitario','costo_unitario','total','ganancia','metodo_pago','notas','registrado_por'];
  var rows = readSheetRows('Ventas', HEADERS);
  return { success:true, data:rows };
}

function deleteSale(id) {
  var sheet = getSheet('Ventas');
  var rows  = sheet.getDataRange().getValues();
  var idStr = String(id).trim();
  for (var i=rows.length-1; i>=1; i--) {
    if (String(rows[i][0]).trim().indexOf(idStr)===0) sheet.deleteRow(i+1);
  }
  return { success:true };
}

// ============================================================
// CONFIGURACIÓN
// ============================================================
function getConfig() {
  var sheet = getSheet('Configuracion');
  var rows  = sheet.getDataRange().getValues();
  var config = {};
  rows.slice(1).forEach(function(row){ if(row[0] && !String(row[0]).startsWith('user_')) config[row[0]]=row[1]; });
  return { success:true, data:config };
}

function saveConfig(data) {
  var sheet = getSheet('Configuracion');
  var rows  = sheet.getDataRange().getValues();
  var keys  = Object.keys(data).filter(function(k){ return k!=='action'&&k!=='callback'; });
  keys.forEach(function(key) {
    var found=false;
    for (var i=1;i<rows.length;i++) { if(rows[i][0]===key){sheet.getRange(i+1,2).setValue(data[key]);found=true;break;} }
    if (!found) sheet.appendRow([key,data[key]]);
  });
  return { success:true };
}

// ============================================================
// USUARIOS
// ============================================================
function getUsers() {
  var sheet = getSheet('Configuracion');
  var rows  = sheet.getDataRange().getValues();
  var users = [];
  for (var i=1; i<rows.length; i++) {
    if (String(rows[i][0]).startsWith('user_')) {
      try { users.push(JSON.parse(String(rows[i][1]))); } catch(e) {}
    }
  }
  return { success:true, data:users };
}

function saveUsers(data) {
  var sheet = getSheet('Configuracion');
  var rows  = sheet.getDataRange().getValues();
  var users = [];
  if (data.users) {
    try { users = typeof data.users==='string' ? JSON.parse(data.users) : data.users; } catch(e) {}
  } else {
    var idx=0;
    while (data['u'+idx]) { try { users.push(JSON.parse(data['u'+idx])); } catch(e) {} idx++; }
  }
  if (users.length===0) return { success:false, error:'No se recibieron usuarios' };
  for (var i=rows.length-1; i>=1; i--) {
    if (String(rows[i][0]).startsWith('user_')) sheet.deleteRow(i+1);
  }
  users.forEach(function(u,i) {
    sheet.appendRow(['user_'+i, JSON.stringify({ user:u.user, pass:u.pass, name:u.name, role:u.role||'Operador' })]);
  });
  return { success:true, message:'Usuarios guardados: '+users.length };
}

// ============================================================
// CATEGORÍAS
// ============================================================
function getCategories() {
  var sheet = getSheet('Categorias');
  var rows  = sheet.getDataRange().getValues();
  if (rows.length===0) return { success:true, data:[] };
  var EXPECTED = ['nombre','descripcion','tiene_tallas','fecha_creacion'];
  var hasH = String(rows[0][0]).toLowerCase().trim()==='nombre';
  var headers = hasH ? rows[0] : EXPECTED;
  var dataRows = hasH ? rows.slice(1) : rows;
  var cats = dataRows.map(function(row) {
    var obj={}; headers.forEach(function(h,i){obj[h]=row[i];});
    obj.tiene_tallas=String(obj.tiene_tallas).toLowerCase().trim();
    return obj;
  }).filter(function(c){ return c.nombre&&String(c.nombre).trim()!==''; });
  return { success:true, data:cats };
}

function saveCategory(data) {
  var sheet = getSheet('Categorias');
  var rows  = sheet.getDataRange().getValues();
  for (var i=1;i<rows.length;i++) {
    if (String(rows[i][0]).trim()===String(data.nombre).trim()) {
      sheet.getRange(i+1,1,1,4).setValues([[data.nombre,data.descripcion||'',data.tiene_tallas||'false',rows[i][3]]]);
      return { success:true };
    }
  }
  sheet.appendRow([data.nombre,data.descripcion||'',data.tiene_tallas||'false',new Date().toISOString()]);
  return { success:true };
}

function deleteCategory(name) {
  var sheet = getSheet('Categorias');
  var rows  = sheet.getDataRange().getValues();
  for (var i=1;i<rows.length;i++) {
    if (String(rows[i][0]).trim()===String(name).trim()) { sheet.deleteRow(i+1); return { success:true }; }
  }
  return { success:false, error:'Categoría no encontrada' };
}

// ============================================================
// CLIENTES
// ============================================================
function getClients() {
  var HEADERS = ['id','nombre_empresa','tipo_factura','tipo_factura_otro','nit','registro','direccion','telefono','email','notas','fecha_creacion','activo'];
  var rows = readSheetRows('Clientes', HEADERS);
  var clients = rows.filter(function(c) {
    if (!c.id || String(c.id).trim() === '') return false;
    var a = String(c.activo).toUpperCase().trim();
    return a !== 'FALSE' && a !== '0' && a !== '';
  });
  return { success:true, data:clients };
}

function saveClient(data) {
  var sheet = getSheet('Clientes');
  var rows  = sheet.getDataRange().getValues();
  if (data.id && String(data.id).trim()!=='') {
    for (var i=1; i<rows.length; i++) {
      if (String(rows[i][0]).trim()===String(data.id).trim()) {
        sheet.getRange(i+1,1,1,12).setValues([[
          data.id, data.nombre_empresa||'', data.tipo_factura||'Consumidor Final',
          data.tipo_factura_otro||'', data.nit||'', data.registro||'',
          data.direccion||'', data.telefono||'', data.email||'',
          data.notas||'', rows[i][10], true
        ]]);
        return { success:true, message:'Cliente actualizado', id:data.id };
      }
    }
  }
  var newId = 'C'+Date.now();
  sheet.appendRow([
    newId, data.nombre_empresa||'', data.tipo_factura||'Consumidor Final',
    data.tipo_factura_otro||'', data.nit||'', data.registro||'',
    data.direccion||'', data.telefono||'', data.email||'',
    data.notas||'', new Date().toISOString(), true
  ]);
  return { success:true, message:'Cliente creado', id:newId };
}

function deleteClient(id) {
  var sheet = getSheet('Clientes');
  var rows  = sheet.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (String(rows[i][0]).trim()===String(id).trim()) {
      sheet.getRange(i+1,12).setValue('FALSE');
      return { success:true };
    }
  }
  return { success:false, error:'Cliente no encontrado' };
}

// ============================================================
// CONTACTOS POR CLIENTE
// ============================================================
function getContacts(clienteId) {
  var HEADERS = ['id','cliente_id','nombre','telefono','email','cargo','notas','fecha_creacion'];
  var rows = readSheetRows('Contactos', HEADERS);
  var contacts = rows.filter(function(c) {
    if (!c.id || String(c.id).trim() === '') return false;
    if (clienteId) return String(c.cliente_id).trim() === String(clienteId).trim();
    return true;
  });
  return { success:true, data:contacts };
}

function saveContact(data) {
  var sheet = getSheet('Contactos');
  var rows  = sheet.getDataRange().getValues();
  if (data.id && String(data.id).trim()!=='') {
    for (var i=1; i<rows.length; i++) {
      if (String(rows[i][0]).trim()===String(data.id).trim()) {
        sheet.getRange(i+1,1,1,8).setValues([[
          data.id, data.cliente_id, data.nombre||'', data.telefono||'',
          data.email||'', data.cargo||'', data.notas||'', rows[i][7]
        ]]);
        return { success:true, id:data.id };
      }
    }
  }
  var newId = 'CT'+Date.now();
  sheet.appendRow([
    newId, data.cliente_id, data.nombre||'', data.telefono||'',
    data.email||'', data.cargo||'', data.notas||'', new Date().toISOString()
  ]);
  return { success:true, id:newId };
}

function deleteContact(id) {
  var sheet = getSheet('Contactos');
  var rows  = sheet.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (String(rows[i][0]).trim()===String(id).trim()) { sheet.deleteRow(i+1); return { success:true }; }
  }
  return { success:false, error:'Contacto no encontrado' };
}

// ============================================================
// COTIZACIONES
// ============================================================
function getNextQuoteNum() {
  var sheet = getSheet('Configuracion');
  var rows  = sheet.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (rows[i][0]==='correlativo_cotizacion') {
      var current = parseInt(rows[i][1])||250;
      sheet.getRange(i+1,2).setValue(current+1);
      return { success:true, numero:current };
    }
  }
  sheet.appendRow(['correlativo_cotizacion','251']);
  return { success:true, numero:250 };
}

function getQuotes() {
  var HEADERS = ['id','numero','fecha','cliente_id','cliente_nombre','asesor','valido_hasta','subtotal','iva_monto','total','estado','metodo_pago','tiempo_entrega','notas','registrado_por','fecha_creacion'];
  var rows = readSheetRows('Cotizaciones', HEADERS);
  var quotes = rows.filter(function(q){ return q.id && String(q.id).trim() !== ''; });
  return { success:true, data:quotes };
}

function saveQuote(data) {
  var sheet = getSheet('Cotizaciones');
  var rows  = sheet.getDataRange().getValues();
  var qDate = data.fecha || new Date().toISOString();
  if (data.id && String(data.id).trim()!=='') {
    for (var i=1; i<rows.length; i++) {
      if (String(rows[i][0]).trim()===String(data.id).trim()) {
        sheet.getRange(i+1,1,1,16).setValues([[
          data.id, data.numero, qDate, data.cliente_id||'', data.cliente_nombre||'',
          data.asesor||'', data.valido_hasta||'', parseFloat(data.subtotal)||0,
          parseFloat(data.iva_monto)||0, parseFloat(data.total)||0,
          data.estado||'Pendiente', data.metodo_pago||'', data.tiempo_entrega||'',
          data.notas||'', data.registrado_por||'', rows[i][15]
        ]]);
        return { success:true, id:data.id };
      }
    }
  }
  var newId = 'Q'+Date.now();
  sheet.appendRow([
    newId, data.numero, qDate, data.cliente_id||'', data.cliente_nombre||'',
    data.asesor||'', data.valido_hasta||'', parseFloat(data.subtotal)||0,
    parseFloat(data.iva_monto)||0, parseFloat(data.total)||0,
    data.estado||'Pendiente', data.metodo_pago||'', data.tiempo_entrega||'',
    data.notas||'', data.registrado_por||'', new Date().toISOString()
  ]);
  return { success:true, id:newId };
}

function saveQuoteItems(data) {
  var sheet = getSheet('CotizacionItems');
  var rows  = sheet.getDataRange().getValues();
  var cotId = data.cotizacion_id;
  var items = typeof data.items==='string' ? JSON.parse(data.items) : (data.items||[]);
  // Eliminar items existentes de esta cotización
  for (var i=rows.length-1; i>=1; i--) {
    if (String(rows[i][1]).trim()===String(cotId).trim()) sheet.deleteRow(i+1);
  }
  // Insertar nuevos items
  items.forEach(function(item) {
    sheet.appendRow([
      'QI'+Date.now()+Math.random().toString(36).substr(2,4),
      cotId, item.producto_id||'', item.descripcion||'',
      item.imagen_url||'', parseInt(item.cantidad)||1,
      parseFloat(item.precio_unit)||0, item.incluye_iva==='true'||item.incluye_iva===true,
      parseFloat(item.subtotal_item)||0
    ]);
  });
  return { success:true };
}

function getQuoteItems(cotizacionId) {
  var HEADERS = ['id','cotizacion_id','producto_id','descripcion','imagen_url','cantidad','precio_unit','incluye_iva','subtotal_item'];
  var rows = readSheetRows('CotizacionItems', HEADERS);
  return { success:true, data:rows.filter(function(item){
    return String(item.cotizacion_id).trim()===String(cotizacionId).trim();
  })};
}

function deleteQuote(id) {
  // Eliminar items primero
  var itemSheet = getSheet('CotizacionItems');
  var itemRows = itemSheet.getDataRange().getValues();
  for (var i=itemRows.length-1; i>=1; i--) {
    if (String(itemRows[i][1]).trim()===String(id).trim()) itemSheet.deleteRow(i+1);
  }
  // Eliminar encabezado
  var sheet = getSheet('Cotizaciones');
  var rows  = sheet.getDataRange().getValues();
  for (var j=rows.length-1; j>=1; j--) {
    if (String(rows[j][0]).trim()===String(id).trim()) { sheet.deleteRow(j+1); return { success:true }; }
  }
  return { success:false, error:'Cotización no encontrada' };
}

// ============================================================
// COMPRAS
// ============================================================
function getPurchases() {
  var HEADERS = ['id','fecha','proveedor','factura_ref','subtotal','iva_monto','total','notas','registrado_por','fecha_creacion'];
  var rows = readSheetRows('Compras', HEADERS);
  return { success:true, data:rows.filter(function(p){ return p.id && String(p.id).trim()!==''; }) };
}

function savePurchase(data) {
  var sheet     = getSheet('Compras');
  var itemSheet = getSheet('CompraItems');
  var prodSheet = getSheet('Productos');
  var items     = typeof data.items==='string' ? JSON.parse(data.items) : (data.items||[]);
  var newId     = 'COM'+Date.now();
  var fecha     = data.fecha || new Date().toISOString();

  sheet.appendRow([
    newId, fecha, data.proveedor||'', data.factura_ref||'',
    parseFloat(data.subtotal)||0, parseFloat(data.iva_monto)||0,
    parseFloat(data.total)||0, data.notas||'',
    data.registrado_por||'', new Date().toISOString()
  ]);

  items.forEach(function(item) {
    var qty = parseInt(item.cantidad)||0;
    var prod = findProduct(prodSheet, item.producto_id);
    var nombre = prod ? prod.nombre : (item.producto_nombre||'');
    itemSheet.appendRow([
      'CI'+Date.now()+Math.random().toString(36).substr(2,4),
      newId, item.producto_id||'', nombre,
      qty, parseFloat(item.costo_unit)||0, qty*(parseFloat(item.costo_unit)||0)
    ]);
    // Actualizar stock (sumar)
    if (item.producto_id) {
      var rowIdx = findProductRowIndex(prodSheet, item.producto_id);
      if (rowIdx > 0) {
        var cur = parseInt(prodSheet.getRange(rowIdx,6).getValue())||0;
        prodSheet.getRange(rowIdx,6).setValue(cur+qty);
        // Actualizar costo si se indicó
        if (item.actualizar_costo==='true'||item.actualizar_costo===true) {
          prodSheet.getRange(rowIdx,5).setValue(parseFloat(item.costo_unit)||0);
        }
      }
    }
  });
  return { success:true, id:newId };
}

function getPurchaseItems(compraId) {
  var HEADERS = ['id','compra_id','producto_id','producto_nombre','cantidad','costo_unit','total_item'];
  var rows = readSheetRows('CompraItems', HEADERS);
  return { success:true, data:rows.filter(function(item){
    return String(item.compra_id).trim()===String(compraId).trim();
  })};
}

function deletePurchase(id) {
  var itemSheet = getSheet('CompraItems');
  var itemRows  = itemSheet.getDataRange().getValues();
  for (var i=itemRows.length-1; i>=1; i--) {
    if (String(itemRows[i][1]).trim()===String(id).trim()) itemSheet.deleteRow(i+1);
  }
  var sheet = getSheet('Compras');
  var rows  = sheet.getDataRange().getValues();
  for (var j=rows.length-1; j>=1; j--) {
    if (String(rows[j][0]).trim()===String(id).trim()) { sheet.deleteRow(j+1); return { success:true }; }
  }
  return { success:false, error:'Compra no encontrada' };
}

// ============================================================
// CUENTAS POR COBRAR
// ============================================================
function getCxC() {
  var HEADERS = ['id','cotizacion_num','cliente_id','cliente_nombre','monto_total','anticipo_recibido','saldo_pendiente','fecha_vencimiento','estado','notas','fecha_creacion'];
  var rows = readSheetRows('CuentasCobrar', HEADERS);
  return { success:true, data:rows.filter(function(c){ return c.id && String(c.id).trim()!==''; }) };
}

function saveCxC(data) {
  var sheet = getSheet('CuentasCobrar');
  var newId = 'CXC'+Date.now();
  var monto    = parseFloat(data.monto_total)||0;
  var anticipo = parseFloat(data.anticipo_recibido)||0;
  sheet.appendRow([
    newId, data.cotizacion_num||'', data.cliente_id||'', data.cliente_nombre||'',
    monto, anticipo, monto-anticipo, data.fecha_vencimiento||'',
    data.estado||'Pendiente', data.notas||'', new Date().toISOString()
  ]);
  return { success:true, id:newId };
}

function updateCxC(data) {
  var sheet = getSheet('CuentasCobrar');
  var rows  = sheet.getDataRange().getValues();
  var headers = rows[0];
  for (var i=1; i<rows.length; i++) {
    if (String(rows[i][0]).trim()===String(data.id).trim()) {
      var monto    = parseFloat(data.monto_total) || parseFloat(rows[i][4]) || 0;
      var anticipo = parseFloat(data.anticipo_recibido) || parseFloat(rows[i][5]) || 0;
      sheet.getRange(i+1,1,1,11).setValues([[
        data.id, data.cotizacion_num||rows[i][1], data.cliente_id||rows[i][2],
        data.cliente_nombre||rows[i][3], monto, anticipo, monto-anticipo,
        data.fecha_vencimiento||rows[i][7], data.estado||rows[i][8],
        data.notas||rows[i][9], rows[i][10]
      ]]);
      return { success:true };
    }
  }
  return { success:false, error:'Registro CxC no encontrado' };
}

function deleteCxC(id) {
  var sheet = getSheet('CuentasCobrar');
  var rows  = sheet.getDataRange().getValues();
  for (var i=1; i<rows.length; i++) {
    if (String(rows[i][0]).trim()===String(id).trim()) { sheet.deleteRow(i+1); return { success:true }; }
  }
  return { success:false, error:'Registro no encontrado' };
}

// ============================================================
// UTILIDADES INTERNAS
// ============================================================
function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName(name); }
  return sheet;
}

// Read sheet rows safely — always uses expected headers, never trusts row 1 to be a header
function readSheetRows(sheetName, expectedHeaders) {
  var sheet = getSheet(sheetName);
  var all = sheet.getDataRange().getValues();
  if (all.length === 0) return [];
  
  // Detect if first row is a header (first cell matches first expected header, case-insensitive)
  var firstCell = String(all[0][0]).toLowerCase().trim();
  var expectedFirst = expectedHeaders[0].toLowerCase();
  var hasHeader = (firstCell === expectedFirst);
  
  // If no header row found, insert one now and use expected headers
  if (!hasHeader) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    sheet.getRange(1, 1, 1, expectedHeaders.length)
         .setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
    // all already has the data rows (without header), so use expectedHeaders directly
    return all.map(function(row) {
      var obj = {};
      expectedHeaders.forEach(function(h, i) { obj[h] = row[i] !== undefined ? row[i] : ''; });
      return obj;
    });
  }
  
  // Normal case: first row is header
  var headers = all[0];
  return all.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i] !== undefined ? row[i] : ''; });
    return obj;
  });
}

function ensureSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var needed = Object.values(SHEETS);
  var existing = ss.getSheets().map(function(s){ return s.getName(); });
  if (!needed.every(function(n){ return existing.indexOf(n)>=0; })) setupSheets();
}

function findProduct(sheet, id) {
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0]; var idStr = String(id).trim();
  for (var i=1; i<rows.length; i++) {
    if (String(rows[i][0]).trim()===idStr) {
      var obj={}; headers.forEach(function(h,idx){ obj[h]=rows[i][idx]; });
      if (obj.tallas&&typeof obj.tallas==='string'&&obj.tallas.trim().startsWith('{')) {
        try { obj.tallas=JSON.parse(obj.tallas); } catch(e){ obj.tallas={}; }
      }
      return obj;
    }
  }
  return null;
}

function findProductRowIndex(sheet, id) {
  var rows = sheet.getDataRange().getValues(); var idStr = String(id).trim();
  for (var i=0; i<rows.length; i++) { if(String(rows[i][0]).trim()===idStr) return i+1; }
  return -1;
}

function logError(action, error) {
  try {
    var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
    if (log) log.appendRow([new Date().toISOString(), action||'', '', error||'']);
  } catch(e) {}
}
