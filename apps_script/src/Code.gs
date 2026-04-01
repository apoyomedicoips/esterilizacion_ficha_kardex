/**
 * Kardex Seguro - Google Apps Script backend
 */

const CONFIG = {
  SPREADSHEET_ID: '161uPwejp2D8YVO7OVgbVyqeUiqzt0fYBzfZG5ERmLLg',
  SESSION_TTL_SECONDS: 60 * 60 * 8,
  SESSION_PREFIX: 'SESSION_',
  USERS_PROP_KEY: 'KARDEX_USERS_JSON',
  PASSWORD_PEPPER_KEY: 'KARDEX_PASSWORD_PEPPER'
};

const SHEETS = {
  ITEMS: {
    name: 'ITEMS',
    headers: ['item_code', 'item_name', 'initial_stock', 'active', 'created_at']
  },
  SERVICES: {
    name: 'SERVICES',
    headers: ['service_name', 'active', 'created_at']
  },
  MOVEMENTS: {
    name: 'MOVEMENTS',
    headers: [
      'movement_id',
      'movement_date',
      'item_code',
      'service_name',
      'move_type',
      'quantity',
      'ticket_no',
      'notes',
      'username',
      'created_at'
    ]
  },
  AUDIT: {
    name: 'AUDIT_LOG',
    headers: ['created_at', 'username', 'action', 'entity', 'entity_id', 'details']
  }
};

function doGet() {
  const t = HtmlService.createTemplateFromFile('index');
  t.appVersion = '1.0.0';
  return t.evaluate().setTitle('Kardex Seguro').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getBootstrap(token) {
  ensureSchema_();
  const session = requireSession_(token, ['admin', 'operator', 'viewer']);
  return {
    user: { username: session.username, role: session.role },
    services: listServices_().filter((s) => s.active),
    items: listItems_().filter((i) => i.active),
    now: new Date().toISOString()
  };
}

function login(payload) {
  ensureSchema_();
  const username = String((payload && payload.username) || '').trim().toLowerCase();
  const password = String((payload && payload.password) || '');
  if (!username || !password) {
    throw new Error('Usuario o contrasena invalida.');
  }

  const users = getUsers_();
  const user = users[username];
  if (!user || !user.active) {
    throw new Error('Usuario o contrasena invalida.');
  }

  if (!verifyPassword_(password, user.passwordHash, user.salt)) {
    throw new Error('Usuario o contrasena invalida.');
  }

  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const expireAt = Date.now() + CONFIG.SESSION_TTL_SECONDS * 1000;
  const session = {
    username,
    role: user.role,
    issuedAt: new Date().toISOString(),
    expireAt
  };

  CacheService.getScriptCache().put(CONFIG.SESSION_PREFIX + token, JSON.stringify(session), CONFIG.SESSION_TTL_SECONDS);
  audit_('LOGIN', 'User', username, 'Login ok', username);

  return {
    token,
    user: { username, role: user.role },
    expireAt
  };
}

function logout(token) {
  const session = getSession_(token);
  if (session) {
    audit_('LOGOUT', 'User', session.username, 'Logout', session.username);
  }
  CacheService.getScriptCache().remove(CONFIG.SESSION_PREFIX + token);
  return { ok: true };
}

function getDashboard(payload) {
  const token = payload && payload.token;
  requireSession_(token, ['admin', 'operator', 'viewer']);

  const month = String((payload && payload.month) || '').trim();
  const parsed = parseMonth_(month);
  const year = parsed.year;
  const mon = parsed.month;

  const firstDay = new Date(year, mon - 1, 1);
  const lastDay = new Date(year, mon, 0);
  const daysInMonth = lastDay.getDate();
  const days = [];
  for (let d = 1; d <= daysInMonth; d++) days.push(d);

  const items = listItems_().filter((i) => i.active);
  const movements = listMovements_();

  const netBefore = {};
  const delta = {};

  movements.forEach((m) => {
    const md = parseDateSafe_(m.movement_date);
    if (!md) return;

    const itemCode = m.item_code;
    const sign = m.move_type === 'IN' ? 1 : -1;
    const qty = Number(m.quantity || 0);
    if (!itemCode || !qty) return;

    if (md < firstDay) {
      netBefore[itemCode] = (netBefore[itemCode] || 0) + sign * qty;
      return;
    }

    if (md > lastDay) return;

    const day = md.getDate();
    const key = itemCode + '|' + day;
    delta[key] = (delta[key] || 0) + sign * qty;
  });

  const rows = items.map((item) => {
    const code = item.item_code;
    let running = Number(item.initial_stock || 0) + Number(netBefore[code] || 0);
    const dayValues = [];
    for (let d = 1; d <= daysInMonth; d++) {
      running += Number(delta[code + '|' + d] || 0);
      dayValues.push(running);
    }
    return {
      item_code: code,
      item_name: item.item_name,
      start_balance: Number(item.initial_stock || 0) + Number(netBefore[code] || 0),
      day_values: dayValues,
      end_balance: running
    };
  });

  return {
    month: Utilities.formatDate(firstDay, Session.getScriptTimeZone(), 'yyyy-MM'),
    days,
    rows,
    recent: movements.slice(0, 100)
  };
}

function createMovement(payload) {
  const session = requireSession_(payload && payload.token, ['admin', 'operator']);
  const movementDate = String((payload && payload.movement_date) || '').trim();
  const itemCode = String((payload && payload.item_code) || '').trim().toUpperCase();
  const serviceName = String((payload && payload.service_name) || '').trim().toUpperCase();
  const moveType = String((payload && payload.move_type) || '').trim().toUpperCase();
  const quantity = Number((payload && payload.quantity) || 0);
  const ticketNo = String((payload && payload.ticket_no) || '').trim().slice(0, 80);
  const notes = String((payload && payload.notes) || '').trim().slice(0, 1000);

  if (!/^\d{4}-\d{2}-\d{2}$/.test(movementDate)) throw new Error('Fecha invalida');
  if (!itemCode) throw new Error('Item requerido');
  if (!serviceName) throw new Error('Servicio requerido');
  if (!(moveType === 'IN' || moveType === 'OUT')) throw new Error('Tipo invalido');
  if (!Number.isFinite(quantity) || quantity <= 0) throw new Error('Cantidad invalida');

  const ws = getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);
  ws.appendRow([
    Utilities.getUuid(),
    movementDate,
    itemCode,
    serviceName,
    moveType,
    Math.trunc(quantity),
    ticketNo,
    notes,
    session.username,
    new Date().toISOString()
  ]);

  audit_('CREATE', 'Movement', itemCode, `${moveType} qty=${quantity} service=${serviceName}`, session.username);
  return { ok: true };
}

function createItem(payload) {
  const session = requireSession_(payload && payload.token, ['admin']);
  const code = String((payload && payload.code) || '').trim().toUpperCase().slice(0, 30);
  const name = String((payload && payload.name) || '').trim().slice(0, 255);
  const initialStock = Number((payload && payload.initial_stock) || 0);

  if (!code || !name) throw new Error('Codigo y nombre requeridos');
  const exists = listItems_().some((i) => i.item_code === code);
  if (exists) throw new Error('Item ya existe');

  const ws = getSheet_(SHEETS.ITEMS.name, SHEETS.ITEMS.headers);
  ws.appendRow([code, name, Math.trunc(initialStock), true, new Date().toISOString()]);
  audit_('CREATE', 'Item', code, name, session.username);
  return { ok: true };
}

function createService(payload) {
  const session = requireSession_(payload && payload.token, ['admin']);
  const serviceName = String((payload && payload.service_name) || '').trim().toUpperCase().slice(0, 120);
  if (!serviceName) throw new Error('Servicio requerido');

  const exists = listServices_().some((s) => s.service_name.toUpperCase() === serviceName);
  if (exists) throw new Error('Servicio ya existe');

  const ws = getSheet_(SHEETS.SERVICES.name, SHEETS.SERVICES.headers);
  ws.appendRow([serviceName, true, new Date().toISOString()]);
  audit_('CREATE', 'Service', serviceName, '', session.username);
  return { ok: true };
}

function getAudit(payload) {
  requireSession_(payload && payload.token, ['admin']);
  return listAudit_(500);
}

function initSchema(payload) {
  const session = requireSession_(payload && payload.token, ['admin']);
  ensureSchema_();
  audit_('INIT', 'Spreadsheet', CONFIG.SPREADSHEET_ID, 'Schema initialized', session.username);
  return { ok: true };
}

function setupDefaultAdmin() {
  const users = getUsers_();
  if (users.user) return 'user already exists';
  const salt = Utilities.getUuid().replace(/-/g, '');
  const hash = hashPassword_('123', salt);
  users.user = { passwordHash: hash, salt: salt, role: 'admin', active: true };
  saveUsers_(users);
  return 'Default user created: user / 123';
}

function upsertUser(username, password, role, active) {
  const uname = String(username || '').trim().toLowerCase();
  if (!uname) throw new Error('username required');
  if (!password || password.length < 8) throw new Error('password too short');
  const validRole = ['admin', 'operator', 'viewer'];
  if (validRole.indexOf(role) === -1) throw new Error('invalid role');

  const users = getUsers_();
  const salt = Utilities.getUuid().replace(/-/g, '');
  users[uname] = {
    passwordHash: hashPassword_(password, salt),
    salt: salt,
    role: role,
    active: active !== false
  };
  saveUsers_(users);
  return { ok: true, username: uname, role: role };
}

function ensureSchema_() {
  getSheet_(SHEETS.ITEMS.name, SHEETS.ITEMS.headers);
  getSheet_(SHEETS.SERVICES.name, SHEETS.SERVICES.headers);
  getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);
  getSheet_(SHEETS.AUDIT.name, SHEETS.AUDIT.headers);
}

function spreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getSheet_(name, headers) {
  const ss = spreadsheet_();
  let ws = ss.getSheetByName(name);
  if (!ws) {
    ws = ss.insertSheet(name);
  }

  const currentHeaders = ws.getRange(1, 1, 1, headers.length).getValues()[0];
  const isEmptyHeader = currentHeaders.every((h) => !String(h || '').trim());
  if (isEmptyHeader) {
    ws.getRange(1, 1, 1, headers.length).setValues([headers]);
    ws.setFrozenRows(1);
  }

  return ws;
}

function getRowsAsObjects_(sheetDef) {
  const ws = getSheet_(sheetDef.name, sheetDef.headers);
  const lastRow = ws.getLastRow();
  if (lastRow < 2) return [];

  const data = ws.getRange(2, 1, lastRow - 1, sheetDef.headers.length).getValues();
  return data
    .filter((r) => String(r[0] || '').trim() !== '')
    .map((r) => {
      const obj = {};
      sheetDef.headers.forEach((h, idx) => (obj[h] = r[idx]));
      return obj;
    });
}

function listItems_() {
  return getRowsAsObjects_(SHEETS.ITEMS).map((r) => ({
    item_code: String(r.item_code || '').trim(),
    item_name: String(r.item_name || '').trim(),
    initial_stock: Number(r.initial_stock || 0),
    active: toBool_(r.active),
    created_at: String(r.created_at || '').trim()
  })).sort((a, b) => a.item_code.localeCompare(b.item_code));
}

function listServices_() {
  return getRowsAsObjects_(SHEETS.SERVICES).map((r) => ({
    service_name: String(r.service_name || '').trim(),
    active: toBool_(r.active),
    created_at: String(r.created_at || '').trim()
  })).sort((a, b) => a.service_name.localeCompare(b.service_name));
}

function listMovements_() {
  return getRowsAsObjects_(SHEETS.MOVEMENTS).map((r) => ({
    movement_id: String(r.movement_id || ''),
    movement_date: String(r.movement_date || ''),
    item_code: String(r.item_code || ''),
    service_name: String(r.service_name || ''),
    move_type: String(r.move_type || 'OUT').toUpperCase(),
    quantity: Number(r.quantity || 0),
    ticket_no: String(r.ticket_no || ''),
    notes: String(r.notes || ''),
    username: String(r.username || ''),
    created_at: String(r.created_at || '')
  })).sort((a, b) => String(b.created_at).localeCompare(String(a.created_at)));
}

function listAudit_(limit) {
  const rows = getRowsAsObjects_(SHEETS.AUDIT).map((r) => ({
    created_at: String(r.created_at || ''),
    username: String(r.username || ''),
    action: String(r.action || ''),
    entity: String(r.entity || ''),
    entity_id: String(r.entity_id || ''),
    details: String(r.details || '')
  }));

  rows.sort((a, b) => String(b.created_at).localeCompare(String(a.created_at)));
  return rows.slice(0, limit || 500);
}

function audit_(action, entity, entityId, details, username) {
  const ws = getSheet_(SHEETS.AUDIT.name, SHEETS.AUDIT.headers);
  ws.appendRow([new Date().toISOString(), username || '', action || '', entity || '', entityId || '', details || '']);
}

function requireSession_(token, roles) {
  const session = getSession_(token);
  if (!session) {
    throw new Error('Sesion expirada. Inicia sesion nuevamente.');
  }

  if (roles && roles.length && roles.indexOf(session.role) === -1) {
    throw new Error('No autorizado.');
  }
  return session;
}

function getSession_(token) {
  const t = String(token || '').trim();
  if (!t) return null;

  const cached = CacheService.getScriptCache().get(CONFIG.SESSION_PREFIX + t);
  if (!cached) return null;

  const session = JSON.parse(cached);
  if (!session || Date.now() > Number(session.expireAt || 0)) {
    CacheService.getScriptCache().remove(CONFIG.SESSION_PREFIX + t);
    return null;
  }

  return session;
}

function getUsers_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CONFIG.USERS_PROP_KEY);
  if (!raw) {
    return {};
  }

  const parsed = JSON.parse(raw);
  Object.keys(parsed).forEach((k) => {
    parsed[k].active = parsed[k].active !== false;
  });
  return parsed;
}

function saveUsers_(users) {
  PropertiesService.getScriptProperties().setProperty(CONFIG.USERS_PROP_KEY, JSON.stringify(users));
}

function getPepper_() {
  const props = PropertiesService.getScriptProperties();
  let pepper = props.getProperty(CONFIG.PASSWORD_PEPPER_KEY);
  if (!pepper) {
    pepper = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty(CONFIG.PASSWORD_PEPPER_KEY, pepper);
  }
  return pepper;
}

function hashPassword_(password, salt) {
  const pepper = getPepper_();
  const raw = salt + '|' + password + '|' + pepper;
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytesToHex_(bytes);
}

function verifyPassword_(password, expectedHash, salt) {
  const computed = hashPassword_(password, salt);
  return safeEquals_(computed, String(expectedHash || ''));
}

function safeEquals_(a, b) {
  if (a.length !== b.length) return false;
  let out = 0;
  for (let i = 0; i < a.length; i++) {
    out |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return out === 0;
}

function bytesToHex_(bytes) {
  return bytes.map((b) => {
    const v = (b + 256) % 256;
    return (v < 16 ? '0' : '') + v.toString(16);
  }).join('');
}

function toBool_(v) {
  const s = String(v || '').toLowerCase().trim();
  return s === 'true' || s === '1' || s === 'si' || s === 'yes' || s === 'verdadero';
}

function parseDateSafe_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s || ''))) return null;
  const [y, m, d] = String(s).split('-').map((x) => Number(x));
  return new Date(y, m - 1, d);
}

function parseMonth_(s) {
  if (/^\d{4}-\d{2}$/.test(String(s || ''))) {
    const y = Number(s.slice(0, 4));
    const m = Number(s.slice(5, 7));
    if (y >= 2020 && y <= 2100 && m >= 1 && m <= 12) return { year: y, month: m };
  }

  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth() + 1 };
}
