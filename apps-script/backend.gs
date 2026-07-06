/**
 * Backend de Poolerie — Portal de Pedidos
 * ========================================
 * Este archivo reemplaza al Apps Script actual ("Proyecto poolerie Mails y mas").
 * Sigue mandando los emails de pedido exactamente igual que antes, y además
 * agrega el login/gestión de usuarios contra una pestaña "Usuarios" de esta
 * misma planilla, en vez de tener los datos de clientes adentro del HTML.
 *
 * CONFIGURACIÓN ANTES DE USAR (una sola vez):
 * 1. Cambiá PEPPER de abajo por cualquier texto largo al azar, y no lo
 *    vuelvas a tocar nunca más (si lo cambiás después de la migración,
 *    todas las contraseñas existentes dejan de funcionar).
 * 2. SPREADSHEET_ID ya está completado con el ID de la planilla real
 *    (este script es standalone, no está atado a ninguna Sheet, así que
 *    no puede usar "la planilla activa" — necesita el ID explícito).
 * 3. Corré la función migrarDatosIniciales() una vez (▶ elegí la función
 *    en el desplegable de arriba, "Ejecutar"). Va a crear la pestaña
 *    "Usuarios" en esa planilla y cargar los 336 usuarios que hoy están
 *    en el HTML.
 * 4. Volvé a desplegar como Web App (Implementar > Nueva implementación)
 *    y pasale la URL nueva a Claude.
 */

const PEPPER = 'CAMBIAR_ESTO_POR_UN_TEXTO_LARGO_AL_AZAR_UNA_SOLA_VEZ';
const SPREADSHEET_ID = '1HQZdQST90P23rZegQ0TRxi6etoQMOOch';
const USERS_SHEET_NAME = 'Usuarios';
const PRODUCTS_SHEET_NAME = 'Productos'; // catálogo "publicado", vive en esta misma planilla
const SESSION_SECONDS = 6 * 60 * 60; // 6 horas — límite técnico de CacheService

// Planilla de precios EXTERNA (la de comparación con Vulcano) — de solo
// lectura, nunca se escribe ahí. El admin la edita libremente cuando
// quiera; solo se toma una "foto" cuando se aprieta "Actualizar precios".
const PRICE_SPREADSHEET_ID = '18iKGWNCx4Zdqeo0OvkBAnQ7sIbBvKbaV';
const PRICE_SHEET_GID = 1046618359;

// ═══════════════════════════════════════════
//  ENTRADA HTTP
// ═══════════════════════════════════════════
function doPost(e) {
  // Apps Script a veces reutiliza el mismo proceso entre requests separados
  // (variables globales "sobreviven"); se limpia el cache acá para que
  // cada request lea la planilla fresca al menos una vez, y solo evite
  // relecturas redundantes DENTRO de este mismo request.
  _sheetCache = null;
  _usersCache = null;
  let result;
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'sendOrder';
    switch (action) {
      case 'login':          result = handleLogin(data); break;
      case 'changePassword': result = handleChangePassword(data); break;
      case 'listVendors':    result = handleListVendors(data); break;
      case 'listClients':    result = handleListClients(data); break;
      case 'listUsers':      result = handleListUsers(data); break;
      case 'saveUser':       result = handleSaveUser(data); break;
      case 'deleteUser':     result = handleDeleteUser(data); break;
      case 'resetPassword':  result = handleResetPassword(data); break;
      case 'importClients':  result = handleImportClients(data); break;
      case 'getProducts':    result = handleGetProducts(data); break;
      case 'refreshPrices':  result = handleRefreshPrices(data); break;
      case 'sendOrder':      result = handleSendOrder(data); break;
      default:                result = { ok: false, error: 'Acción desconocida: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput('Backend Poolerie OK');
}

// ═══════════════════════════════════════════
//  ENVÍO DE EMAIL (igual que el script original)
// ═══════════════════════════════════════════
function handleSendOrder(data) {
  const user = resolveSession(data.token);
  if (!user) return { ok: false, error: 'Tu sesión venció, volvé a loguearte.' };
  const blob = Utilities.newBlob(
    Utilities.base64Decode(data.attachment),
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    data.filename
  );
  GmailApp.sendEmail(data.to, data.subject, '', {
    htmlBody: data.html,
    attachments: [blob],
    replyTo: data.replyTo || ''
  });
  return { ok: true };
}

// ═══════════════════════════════════════════
//  LOGIN / SESIÓN
// ═══════════════════════════════════════════
function handleLogin(data) {
  const user = findUserByLoginOrName(data.usuario);
  if (!user) return { ok: false, error: 'Usuario o contraseña incorrectos.' };
  const hash = hashPassword(data.password || '', user.salt);
  if (hash !== user.passwordHash) return { ok: false, error: 'Usuario o contraseña incorrectos.' };
  const token = createSession(user.clave);
  return { ok: true, token: token, user: publicUser(user) };
}

function handleChangePassword(data) {
  const user = resolveSession(data.token);
  if (!user) return { ok: false, error: 'Tu sesión venció, volvé a loguearte.' };
  const hashActual = hashPassword(data.actual || '', user.salt);
  if (hashActual !== user.passwordHash) return { ok: false, error: 'La contraseña actual no es correcta.' };
  if (!data.nueva || data.nueva.length < 4) return { ok: false, error: 'La contraseña debe tener al menos 4 caracteres.' };
  const salt = makeSalt();
  const hash = hashPassword(data.nueva, salt);
  updateUserFields(user.clave, { passwordHash: hash, salt: salt, passwordChanged: true });
  return { ok: true };
}

// ═══════════════════════════════════════════
//  LISTADOS (sin contraseñas)
// ═══════════════════════════════════════════
function handleListVendors(data) {
  const user = resolveSession(data.token);
  if (!user) return { ok: false, error: 'Tu sesión venció, volvé a loguearte.' };
  const vendors = readAllUsers()
    .filter(function (u) { return u.esVendedor; })
    .map(function (u) { return { key: u.clave, razonSocial: u.razonSocial, email: u.email }; });
  return { ok: true, vendors: vendors };
}

function handleListClients(data) {
  const user = resolveSession(data.token);
  if (!user) return { ok: false, error: 'Tu sesión venció, volvé a loguearte.' };
  if (!user.esVendedor && !user.esAdmin) return { ok: false, error: 'No autorizado.' };
  const clients = readAllUsers()
    .filter(function (u) { return !u.esVendedor; })
    .map(function (u) {
      return {
        key: u.clave, razonSocial: u.razonSocial, cuit: u.cuit, codigo: u.codigo,
        descuento: u.descuento, passwordChanged: u.passwordChanged
      };
    });
  return { ok: true, clients: clients };
}

// Combina listClients + listVendors en un solo viaje de ida y vuelta —
// la usa el panel de Admin para no pedir la planilla dos veces seguidas.
function handleListUsers(data) {
  const check = requireAdmin(data);
  if (check.error) return check.error;
  const all = readAllUsers();
  const vendors = all
    .filter(function (u) { return u.esVendedor; })
    .map(function (u) { return { key: u.clave, razonSocial: u.razonSocial, email: u.email }; });
  const clients = all
    .filter(function (u) { return !u.esVendedor; })
    .map(function (u) {
      return {
        key: u.clave, razonSocial: u.razonSocial, cuit: u.cuit, codigo: u.codigo,
        descuento: u.descuento, passwordChanged: u.passwordChanged
      };
    });
  return { ok: true, clients: clients, vendors: vendors };
}

// ═══════════════════════════════════════════
//  ADMIN — alta / baja / edición / reseteo / import
// ═══════════════════════════════════════════
function handleSaveUser(data) {
  const check = requireAdmin(data);
  if (check.error) return check.error;
  const key = normalizeLogin(data.key);
  if (!key) return { ok: false, error: 'Falta la clave de login.' };
  const existing = findUserRow(key);
  let passwordHash = existing ? existing.passwordHash : '';
  let salt = existing ? existing.salt : makeSalt();
  let passwordChanged = existing ? existing.passwordChanged : false;
  if (data.password) {
    salt = makeSalt();
    passwordHash = hashPassword(data.password, salt);
    passwordChanged = false; // contraseña puesta a mano por el admin: hay que volver a pedirla
  } else if (!existing) {
    return { ok: false, error: 'Falta la contraseña inicial.' };
  }
  upsertUserRow({
    clave: key, razonSocial: data.razonSocial || '', cuit: data.cuit || '', codigo: data.codigo || '',
    passwordHash: passwordHash, salt: salt, descuento: data.descuento || 32.5,
    esVendedor: !!data.isVendor, esAdmin: false, email: data.email || '', passwordChanged: passwordChanged
  }, existing ? existing.rowNum : null);
  return { ok: true };
}

function handleDeleteUser(data) {
  const check = requireAdmin(data);
  if (check.error) return check.error;
  const key = normalizeLogin(data.key);
  if (key === 'ezequiel sanchez') return { ok: false, error: 'No se puede eliminar al admin principal.' };
  const existing = findUserRow(key);
  if (!existing) return { ok: false, error: 'Usuario no encontrado.' };
  getUsersSheet().deleteRow(existing.rowNum);
  _usersCache = null;
  return { ok: true };
}

function handleResetPassword(data) {
  const check = requireAdmin(data);
  if (check.error) return check.error;
  const existing = findUserRow(normalizeLogin(data.key));
  if (!existing) return { ok: false, error: 'Usuario no encontrado.' };
  const newPass = (existing.cuit || '').replace(/-/g, '') || existing.clave;
  const salt = makeSalt();
  const hash = hashPassword(newPass, salt);
  updateUserFields(existing.clave, { passwordHash: hash, salt: salt, passwordChanged: false });
  return { ok: true, newPassword: newPass };
}

function handleImportClients(data) {
  const check = requireAdmin(data);
  if (check.error) return check.error;
  const rows = data.rows || [];
  let created = 0, updated = 0;
  rows.forEach(function (r) {
    const key = normalizeLogin(r.key);
    if (!key) return;
    const existing = findUserRow(key);
    let passwordHash, salt, passwordChanged;
    if (existing) {
      passwordHash = existing.passwordHash;
      salt = existing.salt;
      passwordChanged = existing.passwordChanged;
      updated++;
    } else {
      salt = makeSalt();
      passwordHash = hashPassword(r.pass, salt);
      passwordChanged = false;
      created++;
    }
    upsertUserRow({
      clave: key, razonSocial: r.rs, cuit: r.cuit, codigo: r.cod,
      passwordHash: passwordHash, salt: salt, descuento: r.dto,
      esVendedor: false, esAdmin: false, email: '', passwordChanged: passwordChanged
    }, existing ? existing.rowNum : null);
  });
  return { ok: true, created: created, updated: updated };
}

// ═══════════════════════════════════════════
//  CATÁLOGO DE PRODUCTOS (precios)
// ═══════════════════════════════════════════
// getProducts: lo puede pedir cualquiera (cliente o vendedor, incluso sin
// token) — no es información sensible, ya era visible en el HTML antes
// de esta migración. Lee el catálogo "publicado" en NUESTRA planilla.
function handleGetProducts(data) {
  const products = readAllProducts();
  const updatedAt = PropertiesService.getScriptProperties().getProperty('PRODUCTS_UPDATED_AT') || null;
  return { ok: true, products: products, updatedAt: updatedAt };
}

// refreshPrices: solo admin. Lee la planilla externa de comparación de
// precios (de SOLO LECTURA, nunca se escribe ahí) y publica una foto
// actual del catálogo en la pestaña "Productos" de nuestra planilla.
// Detección de categoría: una fila con texto en Código pero SIN precio
// válido se toma como título de categoría (ej. "CASCOS - LAGUNE") y se
// aplica a los productos siguientes hasta el próximo título.
function handleRefreshPrices(data) {
  const check = requireAdmin(data);
  if (check.error) return check.error;
  let priceSheet;
  try {
    const priceSs = SpreadsheetApp.openById(PRICE_SPREADSHEET_ID);
    priceSheet = priceSs.getSheets().find(function (s) { return s.getSheetId() === PRICE_SHEET_GID; });
  } catch (e) {
    return { ok: false, error: 'No se pudo abrir la planilla de precios: ' + e.message };
  }
  if (!priceSheet) return { ok: false, error: 'No se encontró la hoja de precios (revisar PRICE_SHEET_GID).' };

  const values = priceSheet.getDataRange().getValues();
  const products = [];
  let categoria = '';
  let categoriasDetectadas = 0;
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const codigo = String(row[0] || '').trim();
    const descripcion = String(row[2] || '').trim();
    const moneda = String(row[3] || '').trim();
    const precio = parseFloat(row[4]);
    if (!codigo) continue; // fila vacía o banner sin texto en Código
    if (!isNaN(precio) && precio > 0) {
      products.push({ codigo: codigo, descripcion: descripcion, categoria: categoria, precio: precio, moneda: moneda || 'US$' });
    } else {
      categoria = codigo; // fila título de categoría (sin precio válido)
      categoriasDetectadas++;
    }
  }
  if (!products.length) {
    return { ok: false, error: 'No se encontró ningún producto con precio válido — revisar el formato de la planilla.' };
  }
  writeAllProducts(products);
  const now = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty('PRODUCTS_UPDATED_AT', now);
  return { ok: true, count: products.length, categorias: categoriasDetectadas, updatedAt: now };
}

function getProductsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(PRODUCTS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PRODUCTS_SHEET_NAME);
    sh.appendRow(['Codigo', 'Descripcion', 'Categoria', 'Precio', 'Moneda']);
  }
  return sh;
}

function readAllProducts() {
  const values = getProductsSheet().getDataRange().getValues();
  const products = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;
    products.push({
      codigo: String(r[0]),
      descripcion: String(r[1] || ''),
      categoria: String(r[2] || ''),
      precio: parseFloat(r[3]) || 0,
      moneda: String(r[4] || 'US$')
    });
  }
  return products;
}

function writeAllProducts(products) {
  const sh = getProductsSheet();
  sh.clearContents();
  const header = ['Codigo', 'Descripcion', 'Categoria', 'Precio', 'Moneda'];
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  const rows = products.map(function (p) { return [p.codigo, p.descripcion, p.categoria, p.precio, p.moneda]; });
  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
}

// ═══════════════════════════════════════════
//  HELPERS — permisos / sesión
// ═══════════════════════════════════════════
function requireAdmin(data) {
  const user = resolveSession(data.token);
  if (!user) return { error: { ok: false, error: 'Tu sesión venció, volvé a loguearte.' } };
  if (!user.esAdmin) return { error: { ok: false, error: 'No autorizado.' } };
  return { user: user };
}

function createSession(clave) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, clave, SESSION_SECONDS);
  return token;
}

function resolveSession(token) {
  if (!token) return null;
  const clave = CacheService.getScriptCache().get('session_' + token);
  if (!clave) return null;
  return findUserRow(clave);
}

function publicUser(u) {
  return {
    key: u.clave, razonSocial: u.razonSocial, cuit: u.cuit, codigo: u.codigo,
    descuento: u.descuento, isVendor: u.esVendedor, isAdmin: u.esAdmin,
    email: u.email, passwordChanged: u.passwordChanged
  };
}

// ═══════════════════════════════════════════
//  HELPERS — normalización (igual que el HTML)
// ═══════════════════════════════════════════
function normalize(s) {
  return String(s || '').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function normalizeLogin(s) {
  return normalize(s).replace(/[.\-_+/\\,;:()]/g, '').replace(/\s+/g, ' ').trim();
}

// ═══════════════════════════════════════════
//  HELPERS — contraseñas
// ═══════════════════════════════════════════
function makeSalt() {
  return Utilities.getUuid();
}

function hashPassword(password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password + salt + PEPPER,
    Utilities.Charset.UTF_8
  );
  return bytes.map(function (b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

// ═══════════════════════════════════════════
//  HELPERS — planilla "Usuarios"
// ═══════════════════════════════════════════
// Columnas: Clave | RazonSocial | CUIT | Codigo | PasswordHash | Salt |
//           Descuento | EsVendedor | EsAdmin | Email | PasswordChanged
//
// _sheetCache/_usersCache memoizan la hoja y las filas parseadas durante
// UNA sola ejecución de doPost (varios handlers llaman a readAllUsers()
// más de una vez por request — sin esto, cada llamada releía toda la
// planilla desde cero, y eso era gran parte de la lentitud del panel
// de Admin, que además de todo hace 1-2 requests en paralelo).
let _sheetCache = null;
let _usersCache = null;

function getUsersSheet() {
  if (_sheetCache) return _sheetCache;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET_NAME);
    sh.appendRow(['Clave', 'RazonSocial', 'CUIT', 'Codigo', 'PasswordHash', 'Salt',
      'Descuento', 'EsVendedor', 'EsAdmin', 'Email', 'PasswordChanged']);
  }
  _sheetCache = sh;
  return _sheetCache;
}

function readAllUsers() {
  if (_usersCache) return _usersCache;
  const values = getUsersSheet().getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[0]) continue;
    rows.push({
      rowNum: i + 1,
      clave: normalizeLogin(r[0]),
      razonSocial: r[1],
      cuit: r[2] ? String(r[2]) : '',
      codigo: r[3] || '',
      passwordHash: r[4] || '',
      salt: r[5] || '',
      descuento: parseFloat(r[6]) || 0,
      esVendedor: r[7] === true || r[7] === 'TRUE',
      esAdmin: r[8] === true || r[8] === 'TRUE',
      email: r[9] || '',
      passwordChanged: r[10] === true || r[10] === 'TRUE'
    });
  }
  _usersCache = rows;
  return rows;
}

function findUserRow(clave) {
  const norm = normalizeLogin(clave);
  return readAllUsers().find(function (u) { return u.clave === norm; }) || null;
}

function findUserByLoginOrName(usuario) {
  const norm = normalizeLogin(usuario);
  const all = readAllUsers();
  let u = all.find(function (x) { return x.clave === norm; });
  if (!u) u = all.find(function (x) { return normalizeLogin(x.razonSocial) === norm; });
  return u || null;
}

function upsertUserRow(row, rowNum) {
  const sh = getUsersSheet();
  const values = [row.clave, row.razonSocial, row.cuit, row.codigo, row.passwordHash, row.salt,
    row.descuento, row.esVendedor, row.esAdmin, row.email, row.passwordChanged];
  if (rowNum) {
    sh.getRange(rowNum, 1, 1, values.length).setValues([values]);
  } else {
    sh.appendRow(values);
  }
  _usersCache = null; // la próxima lectura en esta misma ejecución debe ver el cambio
}

function updateUserFields(clave, fields) {
  const existing = findUserRow(clave);
  if (!existing) return;
  const merged = Object.assign({}, existing, fields);
  upsertUserRow(merged, existing.rowNum);
}

// ═══════════════════════════════════════════
//  MIGRACIÓN ÚNICA — correr una sola vez a mano
// ═══════════════════════════════════════════
// Carga los usuarios que hoy están hardcodeados en poolerie_pedidos.html.
// Si un usuario ya existe en la pestaña "Usuarios" (misma Clave), lo salta
// — así se puede volver a correr sin duplicar filas por error.
function migrarDatosIniciales() {
  const sh = getUsersSheet();
  const yaExisten = {};
  readAllUsers().forEach(function (u) { yaExisten[u.clave] = true; });

  function agregar(clave, razonSocial, cuit, codigo, passwordInicial, descuento, esVendedor, esAdmin, email) {
    const key = normalizeLogin(clave);
    if (yaExisten[key]) return;
    const salt = makeSalt();
    upsertUserRow({
      clave: key, razonSocial: razonSocial, cuit: cuit || '', codigo: codigo || '',
      passwordHash: hashPassword(passwordInicial, salt), salt: salt,
      descuento: descuento, esVendedor: esVendedor, esAdmin: esAdmin,
      email: email || '', passwordChanged: false
    }, null);
    yaExisten[key] = true;
  }

  // Vendedores + admin
  agregar('ezequiel sanchez', 'Ezequiel Sánchez', '', '', 'pool2026', 32.5, true, true, 'ezequielsanchez@indusplast.com');
  agregar('matias franchi', 'Matias Franchi', '', '', 'mati2026', 32.5, true, false, 'matiasfranchi@indusplast.com');
  agregar('lucas ñanez', 'Lucas Ñañez', '', '', 'lucas2026', 32.5, true, false, 'lnanez@indusplast.com');
  agregar('nicolas lorenzato', 'Nicolas Lorenzato', '', '', 'nico2026', 32.5, true, false, 'nlorenzato@indusplast.com');
  agregar('augusto camilo', 'Augusto Camilo', '', '', 'augusto2026', 32.5, true, false, 'augustocamilo@indusplast.com');

  // Clientes de ejemplo/demo
  agregar('distribuidora norte s.a.', 'Distribuidora Norte S.A.', '30-11111111-1', '', 'poolerie1', 32.5, false, false, '');
  agregar('piscinas del sur srl', 'Piscinas del Sur SRL', '30-22222222-2', '', 'poolerie2', 30.0, false, false, '');
  agregar('acuatica rosario s.a.', 'Acuática Rosario S.A.', '30-33333333-3', '', 'poolerie3', 35.0, false, false, '');
  agregar('demo', 'Cliente Demo', '00-00000000-0', '', 'demo1234', 32.5, false, false, '');

  // Base de 327 clientes reales: [clave, razonSocial, cuit, codigo, passwordInicial]
  const CLIENTS_BASE = [
    ["+ piletas s.a.s.","+ PILETAS S.A.S.","33718887009","C000270","33718887009"],
    ["123alagua s.a.s.","123ALAGUA S.A.S.","30718968077","C000260","30718968077"],
    ["acua piscinas sas","ACUA PISCINAS SAS","30717646947","C000251","30717646947"],
    ["aguas del parana srl","AGUAS DEL PARANA SRL","30715011499","C000034","30715011499"],
    ["aguilanti juan daniel","AGUILANTI JUAN DANIEL","23166347149","C000185","23166347149"],
    ["almada, alejandro","ALMADA, ALEJANDRO","20270037519","C000005","20270037519"],
    ["amado nicolas","AMADO NICOLAS","20430898737","C000264","20430898737"],
    ["area comercial","AREA COMERCIAL","1111111111","C000218","1111111111"],
    ["armaat soluciones hidraulicas s.r.l.","ARMAAT SOLUCIONES HIDRAULICAS S.R.L.","30718561279","C000290","30718561279"],
    ["armoa cristian luis","ARMOA CRISTIAN LUIS","20339387037","C000220","20339387037"],
    ["arrieta evelin mariel","ARRIETA EVELIN MARIEL","27323172027","C000181","27323172027"],
    ["asg s.a.","ASG S.A.","30715096419","C000316","30715096419"],
    ["asian s.a.","ASIAN S.A.","30719103401","C000301","30719103401"],
    ["ataf s. a. s.","ATAF S. A. S.","30718212835","C000150","30718212835"],
    ["auben (en formacion) s. r. l.","AUBEN (EN FORMACION) S. R. L.","30717379329","C000103","30717379329"],
    ["austral florida properties llc","AUSTRAL FLORIDA PROPERTIES LLC","853814772","C000120","853814772"],
    ["ayala rodolfo","AYALA RODOLFO","20301551259","C000278","20301551259"],
    ["baigorri mercedes nicolasa","BAIGORRI MERCEDES NICOLASA","27270526883","C000155","27270526883"],
    ["barcos gustavo antonio","BARCOS GUSTAVO ANTONIO","20178849167","C000213","20178849167"],
    ["barrigon andres","BARRIGON ANDRES","20117163106","C000072","20117163106"],
    ["battellino mario marcos","BATTELLINO MARIO MARCOS","20277820855","C000075","20277820855"],
    ["bauducco, diego  oscar","BAUDUCCO, DIEGO  OSCAR","20253659026","C000006","20253659026"],
    ["bauso dario ruben","BAUSO DARIO RUBEN","20252576534","C000023","20252576534"],
    ["beguiristain alejandro ceferino","BEGUIRISTAIN ALEJANDRO CEFERINO","20230722472","C000134","20230722472"],
    ["bergese hnos. srl.","BERGESE HNOS. SRL.","30715360744","C000211","30715360744"],
    ["bertonchini diego fabian","BERTONCHINI DIEGO FABIAN","20244025472","C000111","20244025472"],
    ["bezzari mauro agustin","BEZZARI MAURO AGUSTIN","20382507895","C000202","20382507895"],
    ["biarlo ignacio adrian","BIARLO IGNACIO ADRIAN","20266430184","C000232","20266430184"],
    ["blanda pablo daniel","BLANDA PABLO DANIEL","20321368728","C000288","20321368728"],
    ["blu srl","BLU SRL","30717361225","C000114","30717361225"],
    ["bonesso gustavo ruben","BONESSO GUSTAVO RUBEN","20219471514","C000323","20219471514"],
    ["brown pablo javier","BROWN PABLO JAVIER","20265228535","C000281","20265228535"],
    ["bulonera camba s.r.l","BULONERA CAMBA S.R.L","30567765918","C000128","30567765918"],
    ["busso claudio marcel","BUSSO CLAUDIO MARCEL","20232063867","C000244","20232063867"],
    ["c & m srl","C & M SRL","30711887160","C000007","30711887160"],
    ["cabana suiza s.r.l","CABAÑA SUIZA S.R.L","30717188019","C000298","30717188019"],
    ["cafieri aquiles adolfo","CAFIERI AQUILES ADOLFO","20253944987","C000233","20253944987"],
    ["caja transitoria - cheques","CAJA TRANSITORIA - CHEQUES","99999999999","CJTRNST","99999999999"],
    ["campanello daniel angel","CAMPANELLO DANIEL ANGEL","20182693430","C000262","20182693430"],
    ["cano daniel antonio","CANO DANIEL ANTONIO","20297481062","C000215","20297481062"],
    ["caribbean srl","CARIBBEAN SRL","30715341642","C000201","30715341642"],
    ["carletti victor manuel","CARLETTI VICTOR MANUEL","20240514797","C000133","20240514797"],
    ["carlos jarchum","CARLOS JARCHUM","10171493","C000148","10171493"],
    ["carrizo zappile facundo oscar","CARRIZO ZAPPILE FACUNDO OSCAR","20310551229","C000169","20310551229"],
    ["casa vigo sociedad por acciones simplificada","CASA VIGO SOCIEDAD POR ACCIONES SIMPLIFICADA","30717119718","C000183","30717119718"],
    ["casal fabio alejandro","CASAL FABIO ALEJANDRO","20169601942","C000268","20169601942"],
    ["castellano julio rene","CASTELLANO JULIO RENE","20336459568","C000102","20336459568"],
    ["castro ariel maximiliano","CASTRO ARIEL MAXIMILIANO","20307931347","C000269","20307931347"],
    ["ce&ce s.a.s.","CE&CE S.A.S.","30717091589","C000198","30717091589"],
    ["celiz rodrigo carlos","CELIZ RODRIGO CARLOS","20236649688","C000282","20236649688"],
    ["cementos del valle s.r.l","CEMENTOS DEL VALLE S.R.L","30716593564","C000261","30716593564"],
    ["ceramicos margarita srl","CERAMICOS MARGARITA SRL","30712181032","C000091","30712181032"],
    ["cerroclor s.r.l.","CERROCLOR S.R.L.","33707823319","C000203","33707823319"],
    ["cerutti juan manuel","CERUTTI JUAN MANUEL","20284461356","C000063","20284461356"],
    ["chaparro andres","CHAPARRO ANDRES","20243765677","C000176","20243765677"],
    ["charata s.a.s.","CHARATA S.A.S.","30718107012","C000229","30718107012"],
    ["chasilva srl","CHASILVA SRL","30716584573","C000073","30716584573"],
    ["colace david victorino","COLACE DAVID VICTORINO","20233710807","C000142","20233710807"],
    ["colangelo sebastian dario","COLANGELO SEBASTIAN DARIO","20289876716","C000308","20289876716"],
    ["colegio de arquitectos de la pcia de cordoba","COLEGIO DE ARQUITECTOS DE LA PCIA DE CORDOBA","30621592722","C000174","30621592722"],
    ["coleoni fernando gabriel y coleoni omar esteban s de h","COLEONI FERNANDO GABRIEL Y COLEONI OMAR ESTEBAN S DE H","30709758558","C000089","30709758558"],
    ["collazzo maria carina","COLLAZZO MARIA CARINA","27261880046","C000254","27261880046"],
    ["colombil karen natacha","COLOMBIL KAREN NATACHA","27350585635","C000225","27350585635"],
    ["comes fernando enrique","COMES FERNANDO ENRIQUE","20137907187","C000196","20137907187"],
    ["comprandoengrupo.net s.a.","COMPRANDOENGRUPO.NET S.A.","30712116818","C000105","30712116818"],
    ["constructora fibras rengo ltda","CONSTRUCTORA FIBRAS RENGO LTDA","55000000034","C000095","55000000034"],
    ["consumidor final","CONSUMIDOR FINAL","20000001","C000001","20000001"],
    ["corsar s. r. l.","CORSAR S. R. L.","30718115244","C000320","30718115244"],
    ["cosi lorenzo fabian","COSI LORENZO FABIAN","20226608606","C000253","20226608606"],
    ["costamagna leopoldo","COSTAMAGNA LEOPOLDO","20259185247","C000312","20259185247"],
    ["crisger seguridad sa","CRISGER SEGURIDAD SA","30715064177","C000309","30715064177"],
    ["crivello, daniel gustavo","CRIVELLO, DANIEL GUSTAVO","20236619584","C000008","20236619584"],
    ["cuadrado hugo antonio","CUADRADO HUGO ANTONIO","20102499876","C000025","20102499876"],
    ["cuatro b sa","CUATRO B SA","30710591039","C000311","30710591039"],
    ["culzoni sa","CULZONI SA","30692346021","C000310","30692346021"],
    ["czech, gaston","CZECH, GASTON","20386000205","C000082","20386000205"],
    ["dao s.a.","DAO S.A.","33719103699","C000284","33719103699"],
    ["decaral s.r.l.","DECARAL S.R.L.","30710502893","C000240","30710502893"],
    ["deltell andrƒs enrique","DELTELL ANDRƒS ENRIQUE","20205272853","C000058","20205272853"],
    ["demaria mauricio daniel","DEMARIA MAURICIO DANIEL","20333236940","C000267","20333236940"],
    ["desarrollo mevak pilar s.r.l.","DESARROLLO MEVAK PILAR S.R.L.","30718201205","C000257","30718201205"],
    ["dinosaurio s.a.","DINOSAURIO S.A.","30698471472","C000222","30698471472"],
    ["distribuciones cordoba s.a.s.","DISTRIBUCIONES CORDOBA S.A.S.","30716983281","C000136","30716983281"],
    ["distribuidora de la vega s.h.","DISTRIBUIDORA DE LA VEGA S.H.","33710921909","C000217","33710921909"],
    ["distribuidora libertad hogar sa","DISTRIBUIDORA LIBERTAD HOGAR SA","30710859635","C000319","30710859635"],
    ["doltin s.a.s.","DOLTIN S.A.S.","33716761539","C000227","33716761539"],
    ["domene matias leandro","DOMENE MATIAS LEANDRO","23355457079","C000166","23355457079"],
    ["ecop spa","ECOP SPA","769058745","C000048","769058745"],
    ["elektrim sa","ELEKTRIM SA","33586913609","C000171","33586913609"],
    ["emilio perrin srl","EMILIO PERRIN SRL","30709801232","C000086","30709801232"],
    ["ep navi srl","EP NAVI SRL","30714873497","C000057","30714873497"],
    ["errea jorge ricardo","ERREA JORGE RICARDO","20126597631","C000010","20126597631"],
    ["espmod s.a.s.","ESPMOD S.A.S.","30718905423","C000292","30718905423"],
    ["f m e fabiani sociedad de responsabilidad limitada","F M E FABIANI SOCIEDAD DE RESPONSABILIDAD LIMITADA","30688063708","C000237","30688063708"],
    ["fa dop s.a.s.","FA DOP S.A.S.","30716935155","C000168","30716935155"],
    ["faleglass s. r. l.","FALEGLASS S. R. L.","30717770672","C000153","30717770672"],
    ["fernandez luciano","FERNANDEZ LUCIANO","20229441931","C000164","20229441931"],
    ["fernandez nicolcs luis","FERNANDEZ NICOLçS LUIS","20282008298","C000056","20282008298"],
    ["fernandez,  marcos valentin","FERNANDEZ,  MARCOS VALENTIN","20385705914","C000051","20385705914"],
    ["fernandez, maria luisa","FERNANDEZ, MARIA LUISA","27277057242","C000093","27277057242"],
    ["fernandez, victor daniel","FERNANDEZ, VICTOR DANIEL","20291776168","C000038","20291776168"],
    ["fernando lujan s.r.l.","FERNANDO LUJAN S.R.L.","30714758418","C000152","30714758418"],
    ["ferrer lucas gabriel","FERRER LUCAS GABRIEL","20265411100","C000146","20265411100"],
    ["ferrero maria cecilia","FERRERO MARIA CECILIA","27278965436","C000238","27278965436"],
    ["ferreteria loewen sociedad simple","FERRETERIA LOEWEN SOCIEDAD SIMPLE","30716095920","C000219","30716095920"],
    ["ferreyra hnos sociedad por acciones simplificada","FERREYRA HNOS SOCIEDAD POR ACCIONES SIMPLIFICADA","30717467503","C000266","30717467503"],
    ["fibrar s.r.l.","FIBRAR S.R.L.","30718050932","C000195","30718050932"],
    ["fontana pamela anabel","FONTANA PAMELA ANABEL","27319080452","C000263","27319080452"],
    ["fresco srl","FRESCO SRL","218527030010","C000248","218527030010"],
    ["gallardo alexis","GALLARDO ALEXIS","20286560068","C000090","20286560068"],
    ["gallego, mauricio santiago","GALLEGO, MAURICIO SANTIAGO","20234334485","C000059","20234334485"],
    ["galvan lucas gabriel","GALVAN LUCAS GABRIEL","20402384841","C000252","20402384841"],
    ["gamin silvana vanina","GAMIN SILVANA VANINA","23325686774","C000014","23325686774"],
    ["gamu srl","GAMU SRL","30712220577","C000041","30712220577"],
    ["garro dante ariel","GARRO DANTE ARIEL","20228435296","C000012","20228435296"],
    ["genta, juan pablo jesus","GENTA, JUAN PABLO JESUS","20297395174","C000027","20297395174"],
    ["gigena daniel enrique","GIGENA DANIEL ENRIQUE","20208733312","C000191","20208733312"],
    ["gl trade and service sas","GL TRADE AND SERVICE SAS","30716590905","C000050","30716590905"],
    ["gomez marcos fabricio","GOMEZ MARCOS FABRICIO","20254519058","C000013","20254519058"],
    ["gomez roco y cia s r l","GOMEZ ROCO Y CIA S R L","30536158339","C000255","30536158339"],
    ["gonzalez diego german","GONZALEZ DIEGO GERMAN","20292451866","C000293","20292451866"],
    ["gonzalez martinez francisco leonardo","GONZALEZ MARTINEZ FRANCISCO LEONARDO","23261111209","C000173","23261111209"],
    ["grepa s.r.l.","GREPA S.R.L.","30716763222","C000178","30716763222"],
    ["grillo guillermo","GRILLO GUILLERMO","23348403729","C000273","23348403729"],
    ["grupo deltas & agro s.a.s.","GRUPO DELTAS & AGRO S.A.S.","33717561959","C000009","33717561959"],
    ["grupo habitar s.a.s.u","GRUPO HABITAR S.A.S.U","30717550117","C000119","30717550117"],
    ["grupo libertad srl","GRUPO LIBERTAD SRL","30710102569","C000223","30710102569"],
    ["grupo pomiglio sociedad de responsabilidad limitada s. r. l.","GRUPO POMIGLIO SOCIEDAD DE RESPONSABILIDAD LIMITADA S. R. L.","30710668619","C000315","30710668619"],
    ["grupo rc s.a.s.","GRUPO RC S.A.S.","30717969010","C000189","30717969010"],
    ["guridi nicolas","GURIDI NICOLÁS","20362940886","C000096","20362940886"],
    ["gutierrez melisa maria","GUTIERREZ MELISA MARIA","27349087176","C000279","27349087176"],
    ["hause mobel sa","HAUSE MOBEL SA","30685409794","C000172","30685409794"],
    ["hidropool piscinas s.a.","HIDROPOOL PISCINAS S.A.","30712490086","C000092","30712490086"],
    ["hotel calfucura sa","HOTEL CALFUCURA SA","30543170115","C000321","30543170115"],
    ["hoyos hector ramiro","HOYOS HECTOR RAMIRO","20167716866","C000087","20167716866"],
    ["indus py eas","INDUS PY EAS","55000000026","C000180","55000000026"],
    ["indusplast piscinas s. a.","INDUSPLAST PISCINAS S. A.","30710361939","C000015","30710361939"],
    ["industria stark srl","INDUSTRIA STARK SRL","30712602100","C000209","30712602100"],
    ["industrias sur srl","INDUSTRIAS SUR SRL","30709957992","C000083","30709957992"],
    ["ingemont s.r.l.","INGEMONT S.R.L.","30714752053","C000149","30714752053"],
    ["iribarria german horacio","IRIBARRIA GERMAN HORACIO","20278258441","C000194","20278258441"],
    ["irixcor sa","IRIXCOR SA","30710367368","C000175","30710367368"],
    ["isola adrian nicolas","ISOLA ADRIAN NICOLAS","20374819144","C000170","20374819144"],
    ["isola martin","ISOLA MARTIN","20231751719","C000216","20231751719"],
    ["j l farina e hijos sociedad anonima","J L FARINA E HIJOS SOCIEDAD ANONIMA","30707883312","C000132","30707883312"],
    ["jauregui diego andres","JAUREGUI DIEGO ANDRES","20285626898","C000100","20285626898"],
    ["jomax s.a.s.","JOMAX S.A.S.","30716608219","C000138","30716608219"],
    ["juma s. a. s.","JUMA S. A. S.","30718377621","C000186","30718377621"],
    ["keiner, sonia analia","KEINER, SONIA ANALIA","27289610826","C000046","27289610826"],
    ["klor piletas sociedad simple","KLOR PILETAS SOCIEDAD SIMPLE","30716884062","C000193","30716884062"],
    ["kolbe gamin s. a. s.","KOLBE GAMIN S. A. S.","30718157346","C000159","30718157346"],
    ["kreizer leonardo ariel","KREIZER LEONARDO ARIEL","20232654350","C000214","20232654350"],
    ["la casa del instalador s. r. l.","LA CASA DEL INSTALADOR S. R. L.","30717424200","C000258","30717424200"],
    ["la rotonda sas","LA ROTONDA SAS","285755000","C000236","285755000"],
    ["laboratorio labza srl","LABORATORIO LABZA SRL","30711135789","C000108","30711135789"],
    ["lafuente adolfo marcelo","LAFUENTE ADOLFO MARCELO","20246000361","C000043","20246000361"],
    ["lan - per s.a.","LAN - PER S.A.","30709782033","C000313","30709782033"],
    ["lasso anibal ruben","LASSO ANIBAL RUBEN","20285163200","C000228","20285163200"],
    ["latorre jorge ricardo","LATORRE JORGE RICARDO","20242873786","C000026","20242873786"],
    ["leal veronica gabriela","LEAL VERONICA GABRIELA","27279539848","C000068","27279539848"],
    ["leisure technics limited","LEISURE TECHNICS LIMITED","130833345","C000239","130833345"],
    ["litoral piscinas s.a","LITORAL PISCINAS S.A","30714925039","C000062","30714925039"],
    ["liva carolina andrea","LIVA CAROLINA ANDREA","27361903973","C000016","27361903973"],
    ["loewen carlos alberto","LOEWEN CARLOS ALBERTO","20338294361","C000230","20338294361"],
    ["lombo rodrigo","LOMBO RODRIGO","20925463567","C000289","20925463567"],
    ["maipu automotores s a","MAIPU AUTOMOTORES S A","30590493429","C000190","30590493429"],
    ["mangnus trade logistic s.a.","MANGNUS TRADE LOGISTIC S.A.","30715303678","C000182","30715303678"],
    ["maradona ricardo walter","MARADONA RICARDO WALTER","20160697637","C000104","20160697637"],
    ["marfi srl","MARFI SRL","30708779896","C000099","30708779896"],
    ["martinez ariel alejandro","MARTINEZ ARIEL ALEJANDRO","20233710815","C000017","20233710815"],
    ["martinez claudia lorena","MARTINEZ CLAUDIA LORENA","23243030374","C000245","23243030374"],
    ["martinez lallana emanuel ezequiel","MARTINEZ LALLANA EMANUEL EZEQUIEL","20381817378","C000144","20381817378"],
    ["martinuzzi silvana lorena","MARTINUZZI SILVANA LORENA","27278910488","C000208","27278910488"],
    ["matavos, ricardo hugo","MATAVOS, RICARDO HUGO","20108023172","C000024","20108023172"],
    ["mayer cisterna franco matias","MAYER CISTERNA FRANCO MATIAS","20370052515","C000140","20370052515"],
    ["mediterraneo insumos plasticos srl","MEDITERRANEO INSUMOS PLASTICOS SRL","30711115508","C000079","30711115508"],
    ["mega-plastic s.a.","MEGA-PLASTIC S.A.","33709614849","C000124","33709614849"],
    ["menghi bruno","MENGHI BRUNO","20320296995","C000187","20320296995"],
    ["merlo lucas dante","MERLO LUCAS DANTE","20390576898","C000107","20390576898"],
    ["miguel calderon e hijos sociedad anonima","MIGUEL CALDERON E HIJOS SOCIEDAD ANONIMA","30569436547","C000224","30569436547"],
    ["millan alvaro","MILLAN ALVARO","23178580639","C000055","23178580639"],
    ["molina silvana natalia","MOLINA SILVANA NATALIA","27296073704","C000291","27296073704"],
    ["molina, carlos fernando","MOLINA, CARLOS FERNANDO","20213256174","C000047","20213256174"],
    ["montanez ramon horacio","MONTAÑEZ RAMON HORACIO","20240381134","C000165","20240381134"],
    ["mosca, marcela viviana","MOSCA, MARCELA VIVIANA","27175990254","C000018","27175990254"],
    ["mosto germcn","MOSTO GERMçN","20263914040","C000117","20263914040"],
    ["mpyj srl","MPYJ SRL","30716310015","C000054","30716310015"],
    ["mr pool s.a.s","MR POOL S.A.S","55000000018","C000101","55000000018"],
    ["multicraft ra s.a.s.","MULTICRAFT RA S.A.S.","30717130614","C000300","30717130614"],
    ["municipalidad de malagueno","MUNICIPALIDAD DE MALAGUEÑO","30637237159","C000067","30637237159"],
    ["mutio, fausto","MUTIO, FAUSTO","20301641991","C000019","20301641991"],
    ["negri fernando leopoldo","NEGRI FERNANDO LEOPOLDO","20164406165","C000003","20164406165"],
    ["neuro vital s. r. l.","NEURO VITAL S. R. L.","30717815854","C000131","30717815854"],
    ["nicoletti dante nazareno","NICOLETTI DANTE NAZARENO","20297014227","C000147","20297014227"],
    ["novara elian nahuel","NOVARA ELIAN NAHUEL","20318195375","C000002","20318195375"],
    ["nuevos emprendimientos s.r.l.","NUEVOS EMPRENDIMIENTOS S.R.L.","30712506721","C000126","30712506721"],
    ["oegg, marcelo enrique","OEGG, MARCELO ENRIQUE","20275241297","C000060","20275241297"],
    ["ojeda luis misael","OJEDA LUIS MISAEL","23351968389","C000052","23351968389"],
    ["olaiz mariano","OLAIZ MARIANO","20324291718","C000113","20324291718"],
    ["ollaquindia mauro ezequiel","OLLAQUINDIA MAURO EZEQUIEL","20423047993","C000317","20423047993"],
    ["ortiz roberto fabian","ORTIZ ROBERTO FABIAN","20224307412","C000314","20224307412"],
    ["oviedo silva sebastian andres","OVIEDO SILVA SEBASTIAN ANDRES","23247183949","C000137","23247183949"],
    ["pace kevin carlos marcelo","PACE KEVIN CARLOS MARCELO","20371995332","C000033","20371995332"],
    ["pacher pablo luis","PACHER PABLO LUIS","20252082418","C000206","20252082418"],
    ["padilla xiomara natalia soledad","PADILLA XIOMARA NATALIA SOLEDAD","27371105099","C000081","27371105099"],
    ["palavecino dario gabriel","PALAVECINO DARIO GABRIEL","20270791779","C000044","20270791779"],
    ["pambuecor srl","PAMBUECOR SRL","30714243248","C000080","30714243248"],
    ["panero gabriela laura","PANERO GABRIELA LAURA","27306585067","C000256","27306585067"],
    ["perez fernando alfredo","PEREZ FERNANDO ALFREDO","20289892606","C000283","20289892606"],
    ["perez marcela soledad","PEREZ MARCELA SOLEDAD","27286564165","C000249","27286564165"],
    ["perri maria eugenia","PERRI MARIA EUGENIA","27262894636","C000188","27262894636"],
    ["petoletti, damian omar","PETOLETTI, DAMIAN OMAR","20323694320","C000031","20323694320"],
    ["picazo, luciano","PICAZO, LUCIANO","20277440440","C000045","20277440440"],
    ["piccolo ismael pablo","PICCOLO ISMAEL PABLO","20249566048","C000177","20249566048"],
    ["pilecor piscinas s.a.s","PILECOR PISCINAS S.A.S","30717018369","C000074","30717018369"],
    ["piletas y construcciones srl","PILETAS Y CONSTRUCCIONES SRL","30717027775","C000097","30717027775"],
    ["piloni tasso facundo","PILONI TASSO FACUNDO","20402636379","C000324","20402636379"],
    ["pipa hermanos s.r.l.","PIPA HERMANOS S.R.L.","30711733406","C000157","30711733406"],
    ["pire s. a.","PIRE S. A.","30715447416","C000295","30715447416"],
    ["pires juan carlos","PIRES JUAN CARLOS","20260676572","C000199","20260676572"],
    ["piscinas bora bora s.a.","PISCINAS BORA BORA S.A.","30717168077","C000109","30717168077"],
    ["piscinas hernan plast","PISCINAS HERNAN PLAST","30716798832","C000098","30716798832"],
    ["piscinas patagonicas srl","PISCINAS PATAGONICAS SRL","30714306819","C000040","30714306819"],
    ["piscinas santa fe sas","PISCINAS SANTA FE SAS","30716256509","C000066","30716256509"],
    ["piscinas sudeste s.a.s.","PISCINAS SUDESTE S.A.S.","30718132211","C000135","30718132211"],
    ["piscinas oceano sas","PISCINAS OCÉANO SAS","110443470018","C000250","110443470018"],
    ["plaquimet sa","PLAQUIMET SA","30600921335","C000127","30600921335"],
    ["plast car s.r.l.","PLAST CAR S.R.L.","30717285413","C000151","30717285413"],
    ["plasticos del sur s.r.l.","PLASTICOS DEL SUR S.R.L.","30708109580","C000158","30708109580"],
    ["playa y sol s.a.s.","PLAYA Y SOL S.A.S.","30716975165","C000297","30716975165"],
    ["plaza mayor sa","PLAZA MAYOR SA","30683972165","C000234","30683972165"],
    ["poligono industrial malagueno sa","POLIGONO INDUSTRIAL MALAGUEÑO SA","30712526366","C000305","30712526366"],
    ["ponce gustavo federico","PONCE GUSTAVO FEDERICO","20273038842","C000274","20273038842"],
    ["pool equip srl","POOL EQUIP SRL","30710031475","C000042","30710031475"],
    ["pool piscinas confiables zona sur sas","POOL PISCINAS CONFIABLES ZONA SUR SAS","30716556537","C000049","30716556537"],
    ["pooltan s. r. l.","POOLTAN S. R. L.","30717637204","C000118","30717637204"],
    ["precons s.r.l.","PRECONS S.R.L.","30710583265","C000306","30710583265"],
    ["prisma group sas","PRISMA GROUP SAS","30717194426","C000011","30717194426"],
    ["profin del sur s.a.","PROFIN DEL SUR S.A.","30710749066","C000179","30710749066"],
    ["provenzano gustavo ariel","PROVENZANO GUSTAVO ARIEL","20247209930","C000243","20247209930"],
    ["proyectos del sur s.r.l.","PROYECTOS DEL SUR S.R.L.","30715262823","C000247","30715262823"],
    ["pugacz pedro renƒ","PUGACZ PEDRO RENƒ","20239495770","C000028","20239495770"],
    ["quattroccolo angelo","QUATTROCCOLO ANGELO","20355882234","C000272","20355882234"],
    ["quintana luciano ezequiel","QUINTANA LUCIANO EZEQUIEL","20299669573","C000029","20299669573"],
    ["quinteros liliana roxana","QUINTEROS LILIANA ROXANA","29527244","C000088","29527244"],
    ["r y o valle s a","R Y O VALLE S A","30668229073","C000318","30668229073"],
    ["r.a muller y a.t bravo sociedad ley 19550","R.A MULLER Y A.T BRAVO SOCIEDAD LEY 19550","30717148483","C000077","30717148483"],
    ["raffaele natalino di giannantonio","RAFFAELE NATALINO DI GIANNANTONIO","20144972474","C000235","20144972474"],
    ["ramerez pablo andrƒs","RAMêREZ PABLO ANDRƒS","20346510995","C000110","20346510995"],
    ["recanatti gurmando bruno pablo","RECANATTI GURMANDO BRUNO PABLO","20928680755","C000241","20928680755"],
    ["recanatti vallejos paul michel","RECANATTI VALLEJOS PAUL MICHEL","20397145892","C000084","20397145892"],
    ["recanatti vallejos, brian jonathan","RECANATTI VALLEJOS, BRIAN JONATHAN","20364025999","C000116","20364025999"],
    ["reciclados plestick s.a.s.","RECICLADOS PLESTICK S.A.S.","30717423182","C000167","30717423182"],
    ["retamar valeria albina","RETAMAR VALERIA ALBINA","27310802218","C000070","27310802218"],
    ["ribes barberis agustin julian","RIBES BARBERIS AGUSTIN JULIAN","20415933410","C000271","20415933410"],
    ["rimaulo pablo martin","RIMAULO PABLO MARTIN","20277706777","C000161","20277706777"],
    ["rioja plast s r l","RIOJA PLAST S R L","30619360563","C000139","30619360563"],
    ["rion griselda sara","RION GRISELDA SARA","27137185038","C000039","27137185038"],
    ["rioplast s.r.l.","RIOPLAST S.R.L.","30716039397","C000160","30716039397"],
    ["robles leandro y weht carina vanesa","ROBLES LEANDRO Y WEHT CARINA VANESA","30715736981","C000071","30715736981"],
    ["rodriguez rudellat alberto","RODRIGUEZ RUDELLAT ALBERTO","20334112188","C000277","20334112188"],
    ["rojas dario andres","ROJAS DARIO ANDRES","20245793686","C000210","20245793686"],
    ["rojas sergio eduardo","ROJAS SERGIO EDUARDO","20171127115","C000130","20171127115"],
    ["rossini julieta maria","ROSSINI JULIETA MARIA","27296056869","C000265","27296056869"],
    ["rubic sa","RUBIC SA","30715985930","C000242","30715985930"],
    ["s.u.k. s. r. l.","S.U.K. S. R. L.","30711066140","C000325","30711066140"],
    ["sabrina neyra","SABRINA NEYRA","18831080","C000275","18831080"],
    ["salcedo sergio rodrigo","SALCEDO SERGIO RODRIGO","23286544479","C000004","23286544479"],
    ["saldos iniciales","SALDOS INICIALES",".","S900000","."],
    ["salto vanina soledad","SALTO VANINA SOLEDAD","27281331960","C000226","27281331960"],
    ["salvatico graciela rosa","SALVATICO GRACIELA ROSA","23242336534","C000123","23242336534"],
    ["san agustin la rioja s. r. l.","SAN AGUSTIN LA RIOJA S. R. L.","30717515842","C000141","30717515842"],
    ["sanchez facundo","SANCHEZ FACUNDO","20307723418","C000205","20307723418"],
    ["sapunar alicia noemi","SAPUNAR ALICIA NOEMI","27202917068","C000094","27202917068"],
    ["sassaroli federico","SASSAROLI FEDERICO","20351645165","C000259","20351645165"],
    ["satic sa","SATIC SA","30709354422","C000020","30709354422"],
    ["savino javiera","SAVINO JAVIERA","27307544976","C000064","27307544976"],
    ["sayago paulo alejandro","SAYAGO PAULO ALEJANDRO","20297864905","C000207","20297864905"],
    ["sayago, pablo david","SAYAGO, PABLO DAVID","20211124122","C000069","20211124122"],
    ["sburlatti adrian gustavo","SBURLATTI ADRIAN GUSTAVO","20308835937","C000307","20308835937"],
    ["schmidt oscar federico","SCHMIDT OSCAR FEDERICO","20171216053","C000125","20171216053"],
    ["sconfienza agustin","SCONFIENZA AGUSTIN","20315879591","C000280","20315879591"],
    ["scornavacche tissera maria jose","SCORNAVACCHE TISSERA MARIA JOSE","23453854264","C000197","23453854264"],
    ["segovia sandra nancy","SEGOVIA SANDRA NANCY","27227595553","C000129","27227595553"],
    ["servicios sanitarios integrales s. r. l.","SERVICIOS SANITARIOS INTEGRALES S. R. L.","30716005468","C000204","30716005468"],
    ["servifsi s.a.s.","SERVIFSI S.A.S.","30718756738","C000294","30718756738"],
    ["silva guillermo antonio","SILVA GUILLERMO ANTONIO","20286750630","C000121","20286750630"],
    ["silvestrini santos omar","SILVESTRINI SANTOS OMAR","20242060289","C000287","20242060289"],
    ["sofma sas","SOFMA SAS","30716909243","C000154","30716909243"],
    ["solvyp del centro s.r.l.","SOLVYP DEL CENTRO S.R.L.","30718251814","C000184","30718251814"],
    ["stabile hernandez dario lujan","STABILE HERNANDEZ DARIO LUJAN","20929324014","C000037","20929324014"],
    ["starlimm s.a.s.","STARLIMM S.A.S.","30718888022","C000276","30718888022"],
    ["sucesion de piccini eduardo","SUCESION DE PICCINI EDUARDO","23109547379","C000032","23109547379"],
    ["szeve leonardo fabian","SZEVE LEONARDO FABIAN","20336536783","C000304","20336536783"],
    ["tambores san lorenzo srl","TAMBORES SAN LORENZO SRL","30619789020","C000061","30619789020"],
    ["tarca leonardo alejandro","TARCA LEONARDO ALEJANDRO","20295314363","C000112","20295314363"],
    ["techworld profesional srl","TECHWORLD PROFESIONAL SRL","30712277102","C000156","30712277102"],
    ["tibor barath","TIBOR BARATH","28867807","C000053","28867807"],
    ["tiseira johana natali","TISEIRA JOHANA NATALI","27373834012","C000296","27373834012"],
    ["tit can gross sociedad anonima","TIT CAN GROSS SOCIEDAD ANONIMA","30711164401","C000322","30711164401"],
    ["todo piedra s.a.s.","TODO PIEDRA S.A.S.","30716811537","C000163","30716811537"],
    ["todo piscinas srl","TODO PISCINAS SRL","30716534800","C000065","30716534800"],
    ["todone juan bautista","Todone Juan Bautista","35706123","C000143","35706123"],
    ["trabucco ferrari tadeo","TRABUCCO FERRARI TADEO","23436931239","C000285","23436931239"],
    ["tupiscina s. a. s.","TUPISCINA S. A. S.","30717599868","C000115","30717599868"],
    ["u&f smart pools s.r.l.","U&F SMART POOLS S.R.L.","30718609174","C000231","30718609174"],
    ["upg s.a.s.","UPG S.A.S.","30717759261","C000162","30717759261"],
    ["valentini, franco","VALENTINI, FRANCO","20339431583","C000036","20339431583"],
    ["vazquez isidoro","VAZQUEZ ISIDORO","23382796179","C000030","23382796179"],
    ["vecchi, marcos sebastian","VECCHI, MARCOS SEBASTIAN","20231255762","C000085","20231255762"],
    ["veglia florencia soledad","VEGLIA FLORENCIA SOLEDAD","27351646336","C000286","27351646336"],
    ["velez mauro tomas","VELEZ MAURO TOMAS","20320866090","C000078","20320866090"],
    ["vera day ileana noel","VERA DAY ILEANA NOEL","23925324264","C000192","23925324264"],
    ["verzotti carlos evelio","VERZOTTI CARLOS EVELIO","20225374156","C000200","20225374156"],
    ["victoria desarrollos y servicios srl","VICTORIA DESARROLLOS Y SERVICIOS SRL","30717286762","C000212","30717286762"],
    ["villa green s.a.s","VILLA GREEN S.A.S","30717198839","C000076","30717198839"],
    ["visconti srl","VISCONTI SRL","30712513949","C000035","30712513949"],
    ["vital servicios srl","VITAL SERVICIOS SRL","30710975546","C000021","30710975546"],
    ["vivas rene alberto","VIVAS RENE ALBERTO","20210097210","C000299","20210097210"],
    ["viviendas y piletas sol s.r.l.","VIVIENDAS Y PILETAS SOL S.R.L.","30718135954","C000145","30718135954"],
    ["water technics limited","WATER TECHNICS LIMITED","GB 275377763","C000246","GB 275377763"],
    ["waterpool's s.a.s.","WATERPOOL'S S.A.S.","30716508508","C000106","30716508508"],
    ["xenlai s.a.","XENLAI S.A.","30719057094","C000302","30719057094"],
    ["xinlu s.a.","XINLU S.A.","30719116260","C000303","30719116260"],
    ["zerbini irma nora","ZERBINI IRMA NORA","23110509324","C000221","23110509324"],
    ["zeyca s a","ZEYCA S A","30589125572","C000122","30589125572"],
    ["zmf piscinas srl","ZMF PISCINAS SRL","30714939846","C000022","30714939846"]
  ];

  CLIENTS_BASE.forEach(function (c) {
    agregar(c[0], c[1], c[2], c[3], c[4], 32.5, false, false, '');
  });

  Logger.log('Migración terminada. Usuarios en la pestaña: ' + readAllUsers().length);
}
