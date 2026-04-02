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
  const monthMovements = [];
  const serviceAgg = {};

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

    const service = String(m.service_name || '').trim().toUpperCase();
    monthMovements.push({
      movement_date: m.movement_date,
      day: day,
      item_code: itemCode,
      service_name: service,
      move_type: m.move_type,
      quantity: qty
    });

    if (!serviceAgg[service]) {
      serviceAgg[service] = { in_qty: 0, out_qty: 0 };
    }
    if (m.move_type === 'IN') serviceAgg[service].in_qty += qty;
    if (m.move_type === 'OUT') serviceAgg[service].out_qty += qty;
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

  const serviceSummary = Object.keys(serviceAgg)
    .map((name) => ({
      service_name: name,
      in_qty: serviceAgg[name].in_qty,
      out_qty: serviceAgg[name].out_qty,
      net_qty: serviceAgg[name].in_qty - serviceAgg[name].out_qty
    }))
    .sort((a, b) => b.out_qty - a.out_qty);

  return {
    month: Utilities.formatDate(firstDay, Session.getScriptTimeZone(), 'yyyy-MM'),
    days,
    rows,
    recent: movements.slice(0, 200),
    month_movements: monthMovements,
    service_summary: serviceSummary
  };
}

function getMonthlyClosure(payload) {
  const token = payload && payload.token;
  requireSession_(token, ['admin', 'operator', 'viewer']);

  const monthInput = String((payload && payload.month) || '').trim();
  const parsedInput = parseMonth_(monthInput);
  const monthKeyInput = Utilities.formatDate(new Date(parsedInput.year, parsedInput.month - 1, 1), Session.getScriptTimeZone(), 'yyyy-MM');
  const fingerprint = getMovementsFingerprint_();
  const cacheKey = 'CLOSURE_V3_' + monthKeyInput + '_' + fingerprint;
  const cached = CacheService.getScriptCache().get(cacheKey);
  if (cached) return JSON.parse(cached);
  const persisted = getPersistedClosureCache_(cacheKey);
  if (persisted) {
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(persisted), 300);
    return persisted;
  }

  const items = listItems_().filter((i) => i.active);
  const rawMovements = getMovementsFastRows_();
  const parsedMovements = [];
  let latestMonthKey = '';

  rawMovements.forEach((m) => {
    const md = parseDateSafe_(m.movement_date);
    if (!md) return;
    const itemCode = String(m.item_code || '').trim();
    const qty = Number(m.quantity || 0);
    const moveType = String(m.move_type || '').toUpperCase();
    if (!itemCode || !qty || !(moveType === 'IN' || moveType === 'OUT')) return;

    const monthKey = Utilities.formatDate(md, Session.getScriptTimeZone(), 'yyyy-MM');
    if (monthKey > latestMonthKey) latestMonthKey = monthKey;
    parsedMovements.push({ md, monthKey, itemCode, moveType, qty });
  });

  let effectiveYear = parsedInput.year;
  let effectiveMonth = parsedInput.month;
  let effectiveMonthKey = Utilities.formatDate(new Date(effectiveYear, effectiveMonth - 1, 1), Session.getScriptTimeZone(), 'yyyy-MM');

  const hasInSelectedMonth = parsedMovements.some((m) => m.monthKey === effectiveMonthKey);
  if (!hasInSelectedMonth && latestMonthKey) {
    effectiveYear = Number(latestMonthKey.slice(0, 4));
    effectiveMonth = Number(latestMonthKey.slice(5, 7));
    effectiveMonthKey = latestMonthKey;
  }

  const firstDay = new Date(effectiveYear, effectiveMonth - 1, 1);
  const lastDay = new Date(effectiveYear, effectiveMonth, 0);
  const daysInMonth = lastDay.getDate();

  const netBefore = {};
  const monthDaily = {};
  let movementCount = 0;

  parsedMovements.forEach((m) => {
    const sign = m.moveType === 'IN' ? 1 : -1;
    if (m.md < firstDay) {
      netBefore[m.itemCode] = (netBefore[m.itemCode] || 0) + sign * m.qty;
      return;
    }

    if (m.md > lastDay) return;
    if (m.monthKey !== effectiveMonthKey) return;

    if (!monthDaily[m.itemCode]) {
      monthDaily[m.itemCode] = { in_by_day: [], out_by_day: [] };
      for (let i = 0; i < daysInMonth; i++) {
        monthDaily[m.itemCode].in_by_day[i] = 0;
        monthDaily[m.itemCode].out_by_day[i] = 0;
      }
    }
    const dayIdx = m.md.getDate() - 1;
    if (dayIdx < 0 || dayIdx >= daysInMonth) return;
    if (m.moveType === 'IN') monthDaily[m.itemCode].in_by_day[dayIdx] += m.qty;
    if (m.moveType === 'OUT') monthDaily[m.itemCode].out_by_day[dayIdx] += m.qty;
    movementCount += 1;
  });

  const rows = items.map((item) => {
    const code = item.item_code;
    const startBalance = Number(item.initial_stock || 0) + Number(netBefore[code] || 0);
    const daily = monthDaily[code] || { in_by_day: [], out_by_day: [] };
    const inByDay = daily.in_by_day.length ? daily.in_by_day : fillZeros_(daysInMonth);
    const outByDay = daily.out_by_day.length ? daily.out_by_day : fillZeros_(daysInMonth);
    const totalIn = inByDay.reduce((a, b) => a + Number(b || 0), 0);
    const totalOut = outByDay.reduce((a, b) => a + Number(b || 0), 0);
    const endBalanceMonth = startBalance + totalIn - totalOut;
    return {
      item_code: code,
      item_name: item.item_name,
      stock_inicial_mes: startBalance,
      in_by_day: inByDay,
      out_by_day: outByDay,
      saldo_fin_mes: endBalanceMonth
    };
  });

  rows.sort((a, b) => {
    const an = Number(a.item_code);
    const bn = Number(b.item_code);
    const aNum = Number.isFinite(an) && String(a.item_code).trim() !== '';
    const bNum = Number.isFinite(bn) && String(b.item_code).trim() !== '';
    if (aNum && bNum) return an - bn;
    if (aNum) return -1;
    if (bNum) return 1;
    return String(a.item_code).localeCompare(String(b.item_code));
  });

  const out = {
    month: effectiveMonthKey,
    days_in_month: daysInMonth,
    movement_count: movementCount,
    rows: rows,
    auto_switched_month: effectiveMonthKey !== monthInput && !!latestMonthKey
  };
  CacheService.getScriptCache().put(cacheKey, JSON.stringify(out), 300);
  setPersistedClosureCache_(cacheKey, out);
  return out;
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
    movement_date: normalizeDateValue_(r.movement_date),
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

function listMovementsRaw_() {
  return getRowsAsObjects_(SHEETS.MOVEMENTS).map((r) => ({
    movement_date: r.movement_date,
    item_code: String(r.item_code || ''),
    move_type: String(r.move_type || 'OUT').toUpperCase(),
    quantity: Number(r.quantity || 0)
  }));
}

function getMovementsFastRows_() {
  const ws = getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);
  const lastRow = ws.getLastRow();
  if (lastRow < 2) return [];

  // Columns: B movement_date, C item_code, E move_type, F quantity
  const values = ws.getRange(2, 2, lastRow - 1, 5).getValues();
  const out = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const movementDate = r[0];
    const itemCode = String(r[1] || '').trim();
    const moveType = String(r[3] || '').toUpperCase();
    const qty = Number(r[4] || 0);
    if (!itemCode || !qty) continue;
    if (!(moveType === 'IN' || moveType === 'OUT')) continue;
    out.push({
      movement_date: movementDate,
      item_code: itemCode,
      move_type: moveType,
      quantity: qty
    });
  }
  return out;
}

function getMovementsFingerprint_() {
  const ws = getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);
  const lastRow = ws.getLastRow();
  if (lastRow < 2) return '0';
  const row = ws.getRange(lastRow, 1, 1, SHEETS.MOVEMENTS.headers.length).getValues()[0];
  return String(lastRow) + '|' + String(row[0] || '') + '|' + String(row[1] || '') + '|' + String(row[9] || '');
}

function fillZeros_(n) {
  const out = [];
  for (let i = 0; i < n; i++) out[i] = 0;
  return out;
}

function getPersistedClosureCache_(key) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('PERSIST_' + key);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (_e) {
    return null;
  }
}

function setPersistedClosureCache_(key, value) {
  const props = PropertiesService.getScriptProperties();
  try {
    props.setProperty('PERSIST_' + key, JSON.stringify(value));
  } catch (_e) {
    // Ignore if size limit is reached.
  }
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
  if (s instanceof Date && !isNaN(s.getTime())) {
    return new Date(s.getFullYear(), s.getMonth(), s.getDate());
  }

  const raw = String(s || '').trim();
  if (!raw) return null;

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    const [y, m, d] = raw.split('-').map((x) => Number(x));
    return new Date(y, m - 1, d);
  }

  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(raw)) {
    const [d, m, y] = raw.split('/').map((x) => Number(x));
    return new Date(y, m - 1, d);
  }

  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  }
  return null;
}

function normalizeDateValue_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const d = parseDateSafe_(v);
  if (!d) return String(v || '').trim();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
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

function importHistoricoMarzo2026_(sourceSpreadsheetId) {
  ensureSchema_();
  const tz = Session.getScriptTimeZone();
  const source = SpreadsheetApp.openById(sourceSpreadsheetId);
  const stock = source.getSheetByName('STOCK_ACUMULADO_DIARIO');
  const resumen = source.getSheetByName('ResumenMENSUALcierrediario');
  if (!stock || !resumen) throw new Error('Source sheets not found');

  const itemRows = [];
  const serviceName = 'HISTORICO_RECONSTRUIDO';
  const note = `Importado desde ${sourceSpreadsheetId}`;
  const nowIso = new Date().toISOString();

  const codeRange = stock.getRange(3, 1, 40, 2).getValues();
  let itemIdx = 0;
  for (let i = 0; i < codeRange.length; i++) {
    const rawCode = codeRange[i][0];
    const rawName = codeRange[i][1];
    if (!rawCode) continue;

    let code = String(rawCode).trim();
    if (typeof rawCode === 'number') code = String(Math.trunc(rawCode));
    const name = String(rawName || '').trim();

    const resumenRow = 15 + itemIdx * 3;
    const balances = resumen.getRange(resumenRow, 7, 1, 31).getValues()[0].map((v) => Number(v || 0));
    if (!balances[0]) continue;
    itemIdx++;

    itemRows.push({
      code,
      name,
      initialStock: balances[0],
      balances
    });
  }

  const itemsWs = getSheet_(SHEETS.ITEMS.name, SHEETS.ITEMS.headers);
  const servicesWs = getSheet_(SHEETS.SERVICES.name, SHEETS.SERVICES.headers);
  const movementsWs = getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);

  if (itemsWs.getLastRow() > 1) itemsWs.getRange(2, 1, itemsWs.getLastRow() - 1, SHEETS.ITEMS.headers.length).clearContent();
  if (movementsWs.getLastRow() > 1) movementsWs.getRange(2, 1, movementsWs.getLastRow() - 1, SHEETS.MOVEMENTS.headers.length).clearContent();

  const services = listServices_();
  const hasService = services.some((s) => String(s.service_name || '').toUpperCase() === serviceName);
  if (!hasService) servicesWs.appendRow([serviceName, true, nowIso]);

  const itemOut = itemRows.map((it) => [it.code, it.name, it.initialStock, true, nowIso]);
  if (itemOut.length) {
    itemsWs.getRange(2, 1, itemOut.length, SHEETS.ITEMS.headers.length).setValues(itemOut);
  }

  const movements = [];
  itemRows.forEach((it) => {
    for (let d = 2; d <= 31; d++) {
      const delta = Number(it.balances[d - 1] || 0) - Number(it.balances[d - 2] || 0);
      if (!delta) continue;
      const moveType = delta > 0 ? 'IN' : 'OUT';
      const qty = Math.abs(Math.trunc(delta));
      if (!qty) continue;
      const moveDate = Utilities.formatDate(new Date(2026, 2, d), tz, 'yyyy-MM-dd');
      movements.push([
        Utilities.getUuid(),
        moveDate,
        it.code,
        serviceName,
        moveType,
        qty,
        '',
        note,
        'system_import',
        nowIso
      ]);
    }
  });

  if (movements.length) {
    movementsWs.getRange(2, 1, movements.length, SHEETS.MOVEMENTS.headers.length).setValues(movements);
  }

  audit_('IMPORT', 'Historical', '2026-03', `items=${itemRows.length}, movements=${movements.length}`, 'system_import');
  return { items: itemRows.length, movements: movements.length, sourceSpreadsheetId };
}

function importHistoricoDetalladoMarzo2026_(sourceSpreadsheetId) {
  ensureSchema_();
  const tz = Session.getScriptTimeZone();
  const source = SpreadsheetApp.openById(sourceSpreadsheetId);
  const stock = source.getSheetByName('STOCK_ACUMULADO_DIARIO');
  const resumen = source.getSheetByName('ResumenMENSUALcierrediario');
  if (!stock || !resumen) throw new Error('Source sheets not found');

  const codeRows = stock.getRange(3, 1, 40, 2).getValues();
  const items = [];
  let idx = 0;
  for (let i = 0; i < codeRows.length; i++) {
    const rawCode = codeRows[i][0];
    const rawName = codeRows[i][1];
    if (!rawCode) continue;

    let code = String(rawCode).trim();
    if (typeof rawCode === 'number') code = String(Math.trunc(rawCode));
    const name = String(rawName || '').trim();

    const resumenRow = 15 + idx * 3;
    const balances = resumen.getRange(resumenRow, 7, 1, 31).getValues()[0].map((v) => Number(v || 0));
    idx++;
    if (!balances[0]) continue;
    items.push({ code, name, balances, initialStock: balances[0] });
  }

  const servicesSet = {};
  const allMovements = [];
  let adjustmentsIn = 0;
  let adjustmentsOut = 0;

  items.forEach((item) => {
    const sheetName = item.code === 'GUANTES' ? 'Guantes' : 'ITEM' + item.code;
    const ws = source.getSheetByName(sheetName);
    const outByDay = {};

    if (ws) {
      const lastCol = ws.getLastColumn();
      const lastRow = Math.min(ws.getLastRow(), 1200);
      const headers = ws.getRange(8, 1, 1, lastCol).getValues()[0];
      const serviceCols = [];
      for (let c = 4; c <= lastCol; c++) {
        const h = normalizeServiceName_(headers[c - 1]);
        if (!h) continue;
        if (h === 'SERVICIOS' || h === 'TOTAL DIA' || h === 'TOTAL') continue;
        serviceCols.push({ col: c, name: h });
        servicesSet[h] = true;
      }

      if (serviceCols.length > 0 && lastRow >= 9) {
        const table = ws.getRange(9, 1, lastRow - 8, lastCol).getValues();
        for (let r = 0; r < table.length; r++) {
          const row = table[r];
          const dateCell = row[0];
          if (String(dateCell || '').toUpperCase().indexOf('TOTAL') >= 0) break;
          if (!(dateCell instanceof Date)) continue;
          if (dateCell.getFullYear() !== 2026 || dateCell.getMonth() !== 2) continue;

          const day = dateCell.getDate();
          const moveDate = Utilities.formatDate(new Date(2026, 2, day), tz, 'yyyy-MM-dd');
          const ticket = String(row[1] || '').trim().slice(0, 80);
          let outSum = 0;

          serviceCols.forEach((s) => {
            const rawQty = row[s.col - 1];
            const qty = Number(rawQty || 0);
            if (!Number.isFinite(qty) || qty <= 0) return;
            const q = Math.round(qty);
            if (!q) return;

            allMovements.push([
              Utilities.getUuid(),
              moveDate,
              item.code,
              s.name,
              'OUT',
              q,
              ticket,
              `Detalle servicio desde ${sheetName}`,
              'system_import',
              new Date().toISOString()
            ]);
            outSum += q;
          });

          outByDay[day] = (outByDay[day] || 0) + outSum;
        }
      }
    }

    for (let d = 2; d <= 31; d++) {
      const prev = Number(item.balances[d - 2] || 0);
      const curr = Number(item.balances[d - 1] || 0);
      const delta = curr - prev;
      const outMeasured = Number(outByDay[d] || 0);
      const inNeeded = delta + outMeasured;
      const moveDate = Utilities.formatDate(new Date(2026, 2, d), tz, 'yyyy-MM-dd');

      if (inNeeded > 0) {
        const q = Math.round(inNeeded);
        if (q > 0) {
          const svc = 'REPROCESO / INGRESO';
          servicesSet[svc] = true;
          allMovements.push([
            Utilities.getUuid(),
            moveDate,
            item.code,
            svc,
            'IN',
            q,
            '',
            'Ajuste para cuadrar saldo diario',
            'system_import',
            new Date().toISOString()
          ]);
          adjustmentsIn += q;
        }
      } else if (inNeeded < 0) {
        const q = Math.round(Math.abs(inNeeded));
        if (q > 0) {
          const svc = 'AJUSTE SISTEMA';
          servicesSet[svc] = true;
          allMovements.push([
            Utilities.getUuid(),
            moveDate,
            item.code,
            svc,
            'OUT',
            q,
            '',
            'Ajuste negativo para cuadrar saldo diario',
            'system_import',
            new Date().toISOString()
          ]);
          adjustmentsOut += q;
        }
      }
    }
  });

  const itemsWs = getSheet_(SHEETS.ITEMS.name, SHEETS.ITEMS.headers);
  const servicesWs = getSheet_(SHEETS.SERVICES.name, SHEETS.SERVICES.headers);
  const movementsWs = getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);

  if (itemsWs.getLastRow() > 1) itemsWs.getRange(2, 1, itemsWs.getLastRow() - 1, SHEETS.ITEMS.headers.length).clearContent();
  if (servicesWs.getLastRow() > 1) servicesWs.getRange(2, 1, servicesWs.getLastRow() - 1, SHEETS.SERVICES.headers.length).clearContent();
  if (movementsWs.getLastRow() > 1) movementsWs.getRange(2, 1, movementsWs.getLastRow() - 1, SHEETS.MOVEMENTS.headers.length).clearContent();

  const nowIso = new Date().toISOString();
  if (items.length) {
    const itemData = items.map((i) => [i.code, i.name, i.initialStock, true, nowIso]);
    itemsWs.getRange(2, 1, itemData.length, SHEETS.ITEMS.headers.length).setValues(itemData);
  }

  const serviceList = Object.keys(servicesSet).sort();
  if (serviceList.length) {
    const svcData = serviceList.map((s) => [s, true, nowIso]);
    servicesWs.getRange(2, 1, svcData.length, SHEETS.SERVICES.headers.length).setValues(svcData);
  }

  if (allMovements.length) {
    movementsWs.getRange(2, 1, allMovements.length, SHEETS.MOVEMENTS.headers.length).setValues(allMovements);
  }

  audit_(
    'IMPORT',
    'HistoricalDetailed',
    '2026-03',
    `items=${items.length}, services=${serviceList.length}, movements=${allMovements.length}, inAdj=${adjustmentsIn}, outAdj=${adjustmentsOut}`,
    'system_import'
  );

  return {
    items: items.length,
    services: serviceList.length,
    movements: allMovements.length,
    adjustmentsIn,
    adjustmentsOut,
    sourceSpreadsheetId
  };
}

function normalizeServiceName_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const cleaned = s
    .normalize('NFD')
    .replace(/[\\u0300-\\u036f]/g, '')
    .replace(/\\s+/g, ' ')
    .toUpperCase()
    .trim();
  return cleaned;
}

function replaceDataForImport_(payload) {
  ensureSchema_();
  const items = (payload && payload.items) || [];
  const services = (payload && payload.services) || [];
  const movements = (payload && payload.movements) || [];

  const itemsWs = getSheet_(SHEETS.ITEMS.name, SHEETS.ITEMS.headers);
  const servicesWs = getSheet_(SHEETS.SERVICES.name, SHEETS.SERVICES.headers);
  const movementsWs = getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);

  if (itemsWs.getLastRow() > 1) itemsWs.getRange(2, 1, itemsWs.getLastRow() - 1, SHEETS.ITEMS.headers.length).clearContent();
  if (servicesWs.getLastRow() > 1) servicesWs.getRange(2, 1, servicesWs.getLastRow() - 1, SHEETS.SERVICES.headers.length).clearContent();
  if (movementsWs.getLastRow() > 1) movementsWs.getRange(2, 1, movementsWs.getLastRow() - 1, SHEETS.MOVEMENTS.headers.length).clearContent();

  if (items.length) itemsWs.getRange(2, 1, items.length, SHEETS.ITEMS.headers.length).setValues(items);
  if (services.length) servicesWs.getRange(2, 1, services.length, SHEETS.SERVICES.headers.length).setValues(services);
  if (movements.length) movementsWs.getRange(2, 1, movements.length, SHEETS.MOVEMENTS.headers.length).setValues(movements);

  audit_('IMPORT', 'HistoricalDetailed', '2026-03', `replace_data items=${items.length}, services=${services.length}, movements=${movements.length}`, 'system_import');
  return { items: items.length, services: services.length, movements: movements.length };
}
