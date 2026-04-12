/**
 * Kardex Seguro - Google Apps Script backend
 */

const CONFIG = {
  SPREADSHEET_ID: '161uPwejp2D8YVO7OVgbVyqeUiqzt0fYBzfZG5ERmLLg',
  SESSION_TTL_SECONDS: 60 * 60 * 8,
  SESSION_PREFIX: 'SESSION_',
  USERS_PROP_KEY: 'KARDEX_USERS_JSON',
  PASSWORD_PEPPER_KEY: 'KARDEX_PASSWORD_PEPPER',
  DATA_VERSION_KEY: 'KARDEX_DATA_VERSION',
  BACKUP_CSV_PROP_KEY: 'KARDEX_MOVEMENTS_BACKUP_CSV_ID',
  BACKUP_CSV_NAME: 'kardex_movements_backup.csv',
  EVIDENCE_FOLDER_PROP_KEY: 'KARDEX_EVIDENCE_FOLDER_ID',
  EVIDENCE_FOLDER_NAME: 'kardex_evidencias',
  MAX_EVIDENCE_BYTES: 8 * 1024 * 1024,
  DEFAULT_PROVIDER: 'CONSORCIO CDB',
  WARMUP_TRIGGER_FN: 'warmupDashboardCachesJob',
  WARMUP_TRIGGER_HOUR: 4
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
  CONSUMERS: {
    name: 'CONSUMERS',
    headers: ['consumer_name', 'active', 'created_at']
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
      'evidence_name',
      'evidence_url',
      'evidence_file_id',
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
  if (session.role === 'admin') {
    try {
      ensureWarmupTrigger_();
    } catch (_e) {
      // Do not block bootstrap if trigger setup fails.
    }
  }
  return {
    user: { username: session.username, role: session.role },
    services: listProviders_().filter((s) => s.active), // backward compatibility
    providers: listProviders_().filter((s) => s.active),
    consumers: listConsumers_().filter((s) => s.active),
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
  return getDashboardInternal_(month);
}

function getDashboardInternal_(month) {
  const parsed = parseMonth_(month);
  const year = parsed.year;
  const mon = parsed.month;
  const fingerprints = getMovementsFingerprint_() + '_' + getItemsFingerprint_() + '_' + getDataVersion_();
  const monthKey = Utilities.formatDate(new Date(year, mon - 1, 1), Session.getScriptTimeZone(), 'yyyy-MM');
  const cacheKey = buildCacheKey_('DASH_V6', monthKey, fingerprints);
  const scriptCache = CacheService.getScriptCache();
  const cached = scriptCache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  const persisted = getPersistedCache_(cacheKey);
  if (persisted) {
    scriptCache.put(cacheKey, JSON.stringify(persisted), 300);
    return persisted;
  }

  const firstDay = new Date(year, mon - 1, 1);
  const lastDay = new Date(year, mon, 0);
  const daysInMonth = lastDay.getDate();
  const days = [];
  for (let d = 1; d <= daysInMonth; d++) days.push(d);

  const items = listItems_().filter((i) => i.active);
  const movements = getMovementsFastRowsWithService_();

  const netBefore = {};
  const delta = {};
  const monthItemDayAgg = {};
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

    const dayKey = itemCode + '|' + day;
    if (!monthItemDayAgg[dayKey]) {
      monthItemDayAgg[dayKey] = { item_code: itemCode, day: day, in_qty: 0, out_qty: 0 };
    }
    if (m.move_type === 'IN') monthItemDayAgg[dayKey].in_qty += qty;
    if (m.move_type === 'OUT') monthItemDayAgg[dayKey].out_qty += qty;

    const service = String(m.service_name || '').trim().toUpperCase();
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

  const monthMovements = [];
  Object.keys(monthItemDayAgg).forEach((k) => {
    const r = monthItemDayAgg[k];
    if (r.in_qty > 0) {
      monthMovements.push({ day: r.day, item_code: r.item_code, move_type: 'IN', quantity: r.in_qty });
    }
    if (r.out_qty > 0) {
      monthMovements.push({ day: r.day, item_code: r.item_code, move_type: 'OUT', quantity: r.out_qty });
    }
  });
  monthMovements.sort((a, b) => Number(a.day || 0) - Number(b.day || 0) || String(a.item_code).localeCompare(String(b.item_code)));

  const recentLite = getRecentMovementsLite_(300);

  const out = {
    month: Utilities.formatDate(firstDay, Session.getScriptTimeZone(), 'yyyy-MM'),
    days,
    rows,
    recent: recentLite,
    month_movements: monthMovements,
    service_summary: serviceSummary
  };
  try {
    scriptCache.put(cacheKey, JSON.stringify(out), 300);
  } catch (_e) {
    // ignore cache errors
  }
  setPersistedCache_(cacheKey, out);
  return out;
}

function getMonthlyClosure(payload) {
  const token = payload && payload.token;
  requireSession_(token, ['admin', 'operator', 'viewer']);

  const monthInput = String((payload && payload.month) || '').trim();
  const parsedInput = parseMonth_(monthInput);
  const monthKeyInput = Utilities.formatDate(new Date(parsedInput.year, parsedInput.month - 1, 1), Session.getScriptTimeZone(), 'yyyy-MM');
  const fingerprint = getMovementsFingerprint_() + '_' + getItemsFingerprint_() + '_' + getDataVersion_();
  const cacheKey = buildCacheKey_('CLOSURE_V4', monthKeyInput, fingerprint);
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
  let serviceName = String((payload && payload.service_name) || '').trim().toUpperCase();
  const moveType = String((payload && payload.move_type) || '').trim().toUpperCase();
  const quantity = Number((payload && payload.quantity) || 0);
  const ticketNo = String((payload && payload.ticket_no) || '').trim().slice(0, 80);
  const notes = String((payload && payload.notes) || '').trim().slice(0, 1000);
  const evidence = payload && payload.evidence ? payload.evidence : null;

  if (!/^\d{4}-\d{2}-\d{2}$/.test(movementDate)) throw new Error('Fecha invalida');
  if (!itemCode) throw new Error('Item requerido');
  if (moveType === 'IN' && !serviceName) {
    serviceName = CONFIG.DEFAULT_PROVIDER;
  }
  if (!serviceName) throw new Error('Proveedor/area requerido');
  if (!(moveType === 'IN' || moveType === 'OUT')) throw new Error('Tipo invalido');
  if (!Number.isFinite(quantity) || quantity <= 0) throw new Error('Cantidad invalida');

  if (moveType === 'IN') {
    const providers = listProviders_().map((s) => String(s.service_name || '').toUpperCase());
    if (providers.indexOf(serviceName) === -1) {
      throw new Error('Proveedor no registrado. Agreguelo en Catalogo de proveedores.');
    }
  } else {
    const consumers = listConsumers_().map((s) => String(s.consumer_name || '').toUpperCase());
    if (consumers.indexOf(serviceName) === -1) {
      throw new Error('Dependencia consumidora no registrada. Agreguela en Catalogo de dependencias.');
    }
  }

  let evidenceName = '';
  let evidenceUrl = '';
  let evidenceFileId = '';
  if (evidence && evidence.data_base64) {
    const out = createEvidenceFile_(evidence, session.username);
    evidenceName = out.name;
    evidenceUrl = out.url;
    evidenceFileId = out.fileId;
  }

  const ws = getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);
  const rowData = [
    Utilities.getUuid(),
    movementDate,
    itemCode,
    serviceName,
    moveType,
    Math.trunc(quantity),
    ticketNo,
    notes,
    evidenceName,
    evidenceUrl,
    evidenceFileId,
    session.username,
    new Date().toISOString()
  ];
  ws.appendRow(rowData);
  try {
    appendMovementBackupCsv_(rowData);
  } catch (e) {
    // Keep data integrity: if CSV backup fails, revert the spreadsheet append.
    try {
      const lastRow = ws.getLastRow();
      const lastId = String(ws.getRange(lastRow, 1, 1, 1).getValue() || '');
      if (lastId === String(rowData[0])) {
        ws.deleteRow(lastRow);
      }
    } catch (_revertErr) {
      // Ignore rollback failure.
    }
    audit_('ERROR', 'MovementBackupCSV', itemCode, String(e && e.message ? e.message : e), session.username);
    throw new Error('No se pudo guardar la copia de seguridad CSV. El movimiento fue revertido.');
  }
  bumpDataVersion_();

  audit_(
    'CREATE',
    'Movement',
    itemCode,
    `${moveType} qty=${quantity} provider_or_area=${serviceName}; evidence=${evidenceFileId ? 'SI' : 'NO'}`,
    session.username
  );
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
  bumpDataVersion_();
  audit_('CREATE', 'Item', code, name, session.username);
  return { ok: true };
}

function createService(payload) {
  const session = requireSession_(payload && payload.token, ['admin']);
  const serviceName = String((payload && payload.service_name) || '').trim().toUpperCase().slice(0, 120);
  if (!serviceName) throw new Error('Proveedor requerido');

  const exists = listProviders_().some((s) => s.service_name.toUpperCase() === serviceName);
  if (exists) throw new Error('Proveedor ya existe');

  const ws = getSheet_(SHEETS.SERVICES.name, SHEETS.SERVICES.headers);
  ws.appendRow([serviceName, true, new Date().toISOString()]);
  bumpDataVersion_();
  audit_('CREATE', 'Provider', serviceName, '', session.username);
  return { ok: true };
}

function createConsumer(payload) {
  const session = requireSession_(payload && payload.token, ['admin']);
  const name = String((payload && payload.consumer_name) || '').trim().toUpperCase().slice(0, 120);
  if (!name) throw new Error('Dependencia requerida');
  const exists = listConsumers_().some((c) => String(c.consumer_name || '').toUpperCase() === name);
  if (exists) throw new Error('Dependencia ya existe');
  const ws = getSheet_(SHEETS.CONSUMERS.name, SHEETS.CONSUMERS.headers);
  ws.appendRow([name, true, new Date().toISOString()]);
  bumpDataVersion_();
  audit_('CREATE', 'Consumer', name, '', session.username);
  return { ok: true };
}

function listMovementsByMonth(payload) {
  const token = payload && payload.token;
  requireSession_(token, ['admin', 'operator']);
  const month = String((payload && payload.month) || '').trim();
  const parsed = parseMonth_(month);
  const first = new Date(parsed.year, parsed.month - 1, 1);
  const last = new Date(parsed.year, parsed.month, 0);

  const rows = listMovements_().filter((m) => {
    const d = parseDateSafe_(m.movement_date);
    if (!d) return false;
    return d >= first && d <= last;
  });
  rows.sort((a, b) => String(b.created_at || '').localeCompare(String(a.created_at || '')));
  return {
    month: Utilities.formatDate(first, Session.getScriptTimeZone(), 'yyyy-MM'),
    rows: rows
  };
}

function updateMovement(payload) {
  const session = requireSession_(payload && payload.token, ['admin', 'operator']);
  const movementId = String((payload && payload.movement_id) || '').trim();
  const movementDate = String((payload && payload.movement_date) || '').trim();
  const itemCode = String((payload && payload.item_code) || '').trim().toUpperCase();
  let serviceName = String((payload && payload.service_name) || '').trim().toUpperCase();
  const moveType = String((payload && payload.move_type) || '').trim().toUpperCase();
  const quantity = Number((payload && payload.quantity) || 0);
  const ticketNo = String((payload && payload.ticket_no) || '').trim().slice(0, 80);
  const notes = String((payload && payload.notes) || '').trim().slice(0, 1000);

  if (!movementId) throw new Error('ID de movimiento requerido');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(movementDate)) throw new Error('Fecha invalida');
  if (!itemCode) throw new Error('Item requerido');
  if (moveType === 'IN' && !serviceName) serviceName = CONFIG.DEFAULT_PROVIDER;
  if (!serviceName) throw new Error('Proveedor/area requerido');
  if (!(moveType === 'IN' || moveType === 'OUT')) throw new Error('Tipo invalido');
  if (!Number.isFinite(quantity) || quantity <= 0) throw new Error('Cantidad invalida');

  if (moveType === 'IN') {
    const providers = listProviders_().map((s) => String(s.service_name || '').toUpperCase());
    if (providers.indexOf(serviceName) === -1) {
      throw new Error('Proveedor no registrado. Agreguelo en Catalogo de proveedores.');
    }
  } else {
    const consumers = listConsumers_().map((s) => String(s.consumer_name || '').toUpperCase());
    if (consumers.indexOf(serviceName) === -1) {
      throw new Error('Dependencia consumidora no registrada. Agreguela en Catalogo de dependencias.');
    }
  }

  const ws = getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);
  const lastRow = ws.getLastRow();
  if (lastRow < 2) throw new Error('No hay movimientos para editar');
  const ids = ws.getRange(2, 1, lastRow - 1, 1).getValues();
  let rowNum = -1;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '') === movementId) {
      rowNum = i + 2;
      break;
    }
  }
  if (rowNum < 2) throw new Error('Movimiento no encontrado');

  const rowOld = ws.getRange(rowNum, 1, 1, SHEETS.MOVEMENTS.headers.length).getValues()[0];
  const oldObj = {
    movement_id: String(rowOld[0] || ''),
    movement_date: normalizeDateValue_(rowOld[1]),
    item_code: String(rowOld[2] || ''),
    service_name: String(rowOld[3] || ''),
    move_type: String(rowOld[4] || ''),
    quantity: Number(rowOld[5] || 0),
    ticket_no: String(rowOld[6] || ''),
    notes: String(rowOld[7] || '')
  };

  ws.getRange(rowNum, 2, 1, 7).setValues([[
    movementDate,
    itemCode,
    serviceName,
    moveType,
    Math.trunc(quantity),
    ticketNo,
    notes
  ]]);
  try {
    rebuildMovementBackupCsvFromSheet_();
  } catch (e) {
    // Rollback sheet update if backup cannot be rebuilt.
    ws.getRange(rowNum, 2, 1, 7).setValues([[
      oldObj.movement_date,
      oldObj.item_code,
      oldObj.service_name,
      oldObj.move_type,
      Math.trunc(Number(oldObj.quantity || 0)),
      oldObj.ticket_no,
      oldObj.notes
    ]]);
    throw new Error('No se pudo actualizar la copia CSV de seguridad. La edicion fue revertida.');
  }

  bumpDataVersion_();

  const newObj = {
    movement_id: movementId,
    movement_date: movementDate,
    item_code: itemCode,
    service_name: serviceName,
    move_type: moveType,
    quantity: Math.trunc(quantity),
    ticket_no: ticketNo,
    notes: notes
  };
  const changes = diffMovementForAudit_(oldObj, newObj);
  audit_('UPDATE', 'Movement', movementId, `editor=${session.username}; changes=${changes}`, session.username);
  return { ok: true };
}

function listUsersConfig(payload) {
  const session = requireSession_(payload && payload.token, ['admin']);
  requireMasterAdmin_(session);
  const users = getUsers_();
  return Object.keys(users).sort().map((u) => ({
    username: u,
    role: String(users[u].role || 'operator'),
    active: users[u].active !== false
  }));
}

function createUserConfig(payload) {
  const session = requireSession_(payload && payload.token, ['admin']);
  requireMasterAdmin_(session);
  const username = String((payload && payload.username) || '').trim().toLowerCase();
  const password = String((payload && payload.password) || '');
  const role = String((payload && payload.role) || 'operator').trim().toLowerCase();
  if (!username) throw new Error('Usuario requerido');
  if (!password || password.length < 3) throw new Error('Contrasena minima: 3 caracteres');
  const validRole = ['admin', 'operator', 'viewer'];
  if (validRole.indexOf(role) === -1) throw new Error('Rol invalido');

  const users = getUsers_();
  if (users[username]) throw new Error('Usuario ya existe');
  const salt = Utilities.getUuid().replace(/-/g, '');
  users[username] = {
    passwordHash: hashPassword_(password, salt),
    salt: salt,
    role: role,
    active: true
  };
  saveUsers_(users);
  audit_('CREATE', 'User', username, `role=${role}`, session.username);
  return { ok: true, username: username };
}

function changeMyPassword(payload) {
  const session = requireSession_(payload && payload.token, ['admin', 'operator', 'viewer']);
  const currentPassword = String((payload && payload.current_password) || '');
  const newPassword = String((payload && payload.new_password) || '');
  if (!newPassword || newPassword.length < 3) throw new Error('La nueva contrasena debe tener al menos 3 caracteres');

  const users = getUsers_();
  const user = users[session.username];
  if (!user) throw new Error('Usuario no encontrado');
  if (!verifyPassword_(currentPassword, user.passwordHash, user.salt)) {
    throw new Error('Contrasena actual incorrecta');
  }

  const salt = Utilities.getUuid().replace(/-/g, '');
  user.salt = salt;
  user.passwordHash = hashPassword_(newPassword, salt);
  users[session.username] = user;
  saveUsers_(users);
  audit_('PASSWORD_CHANGE', 'User', session.username, 'password updated', session.username);
  return { ok: true };
}

function getAudit(payload) {
  requireSession_(payload && payload.token, ['admin']);
  return listAudit_(500);
}

function initSchema(payload) {
  const session = requireSession_(payload && payload.token, ['admin']);
  ensureSchema_();
  ensureWarmupTrigger_();
  audit_('INIT', 'Spreadsheet', CONFIG.SPREADSHEET_ID, 'Schema initialized', session.username);
  return { ok: true };
}

function installWarmupScheduler(payload) {
  const session = requireSession_(payload && payload.token, ['admin']);
  ensureSchema_();
  const triggerInfo = ensureWarmupTrigger_();
  const warmed = warmupDashboardCachesJob();
  audit_('CONFIG', 'WarmupTrigger', triggerInfo.triggerId || '', `installed by ${session.username}`, session.username);
  return {
    ok: true,
    trigger: triggerInfo,
    warmup: warmed
  };
}

function warmupDashboardCachesJob() {
  ensureSchema_();
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const curr = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), 1), tz, 'yyyy-MM');
  const prev = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() - 1, 1), tz, 'yyyy-MM');
  const next = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() + 1, 1), tz, 'yyyy-MM');

  const months = [prev, curr, next];
  const warmed = [];
  for (let i = 0; i < months.length; i++) {
    const mk = months[i];
    try {
      const out = getDashboardInternal_(mk);
      warmed.push({ month: mk, ok: true, rows: (out.rows || []).length, movements: (out.month_movements || []).length });
    } catch (e) {
      warmed.push({ month: mk, ok: false, error: String(e && e.message ? e.message : e) });
    }
  }
  return {
    ok: true,
    at: new Date().toISOString(),
    warmed: warmed
  };
}

function ensureWarmupTrigger_() {
  const fn = CONFIG.WARMUP_TRIGGER_FN;
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === fn) {
      return { exists: true, triggerId: triggers[i].getUniqueId ? triggers[i].getUniqueId() : '' };
    }
  }

  const created = ScriptApp.newTrigger(fn)
    .timeBased()
    .atHour(Number(CONFIG.WARMUP_TRIGGER_HOUR || 4))
    .everyDays(1)
    .create();

  return { exists: false, triggerId: created.getUniqueId ? created.getUniqueId() : '' };
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
  if (!password || password.length < 3) throw new Error('password too short');
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
  getSheet_(SHEETS.CONSUMERS.name, SHEETS.CONSUMERS.headers);
  getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);
  getSheet_(SHEETS.AUDIT.name, SHEETS.AUDIT.headers);
  ensureDefaultProviders_();
  ensureDefaultConsumers_();
  ensureDefaultUsers_();
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
  } else {
    const nextHeaders = currentHeaders.slice();
    let changed = false;
    for (let i = 0; i < headers.length; i++) {
      const expected = String(headers[i] || '').trim();
      const current = String(nextHeaders[i] || '').trim();
      if (!current) {
        nextHeaders[i] = expected;
        changed = true;
      }
    }
    if (changed) {
      ws.getRange(1, 1, 1, headers.length).setValues([nextHeaders]);
    }
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

function listProviders_() {
  return getRowsAsObjects_(SHEETS.SERVICES).map((r) => ({
    service_name: String(r.service_name || '').trim(),
    active: toBool_(r.active),
    created_at: String(r.created_at || '').trim()
  })).sort((a, b) => a.service_name.localeCompare(b.service_name));
}

function listConsumers_() {
  return getRowsAsObjects_(SHEETS.CONSUMERS).map((r) => ({
    consumer_name: String(r.consumer_name || '').trim(),
    active: toBool_(r.active),
    created_at: String(r.created_at || '').trim()
  })).sort((a, b) => a.consumer_name.localeCompare(b.consumer_name));
}

function listServices_() {
  return listProviders_();
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
    evidence_name: String(r.evidence_name || ''),
    evidence_url: String(r.evidence_url || ''),
    evidence_file_id: String(r.evidence_file_id || ''),
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

function getMovementsFastRowsWithService_() {
  const ws = getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);
  const lastRow = ws.getLastRow();
  if (lastRow < 2) return [];

  // Columns: B movement_date, C item_code, D service_name, E move_type, F quantity
  const values = ws.getRange(2, 2, lastRow - 1, 5).getValues();
  const out = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const movementDate = r[0];
    const itemCode = String(r[1] || '').trim();
    const serviceName = String(r[2] || '').trim();
    const moveType = String(r[3] || '').toUpperCase();
    const qty = Number(r[4] || 0);
    if (!itemCode || !qty) continue;
    if (!(moveType === 'IN' || moveType === 'OUT')) continue;
    out.push({
      movement_date: movementDate,
      item_code: itemCode,
      service_name: serviceName,
      move_type: moveType,
      quantity: qty
    });
  }
  return out;
}

function getRecentMovementsLite_(limit) {
  const ws = getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);
  const lastRow = ws.getLastRow();
  if (lastRow < 2) return [];
  const n = Math.max(1, Math.min(Number(limit || 300), lastRow - 1));
  const start = lastRow - n + 1;
  // Only movement_date (B)
  const values = ws.getRange(start, 2, n, 1).getValues();
  const out = [];
  for (let i = values.length - 1; i >= 0; i--) {
    out.push({ movement_date: normalizeDateValue_(values[i][0]) });
  }
  return out;
}

function getMovementsFingerprint_() {
  const ws = getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);
  const lastRow = ws.getLastRow();
  if (lastRow < 2) return '0';
  const row = ws.getRange(lastRow, 1, 1, SHEETS.MOVEMENTS.headers.length).getValues()[0];
  return String(lastRow) + '|' + String(row[0] || '') + '|' + String(row[1] || '') + '|' + String(row[12] || '');
}

function getItemsFingerprint_() {
  const ws = getSheet_(SHEETS.ITEMS.name, SHEETS.ITEMS.headers);
  const lastRow = ws.getLastRow();
  if (lastRow < 2) return '0';
  const row = ws.getRange(lastRow, 1, 1, SHEETS.ITEMS.headers.length).getValues()[0];
  return String(lastRow) + '|' + String(row[0] || '') + '|' + String(row[1] || '') + '|' + String(row[4] || '');
}

function fillZeros_(n) {
  const out = [];
  for (let i = 0; i < n; i++) out[i] = 0;
  return out;
}

function getPersistedCache_(key) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('PERSIST_' + key);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (_e) {
    return null;
  }
}

function getDataVersion_() {
  const props = PropertiesService.getScriptProperties();
  const raw = String(props.getProperty(CONFIG.DATA_VERSION_KEY) || '0').trim();
  const n = Number(raw);
  return Number.isFinite(n) && n >= 0 ? String(Math.trunc(n)) : '0';
}

function bumpDataVersion_() {
  const props = PropertiesService.getScriptProperties();
  const n = Number(getDataVersion_()) + 1;
  props.setProperty(CONFIG.DATA_VERSION_KEY, String(Math.trunc(n)));
}

function setPersistedCache_(key, value) {
  const props = PropertiesService.getScriptProperties();
  try {
    props.setProperty('PERSIST_' + key, JSON.stringify(value));
  } catch (_e) {
    // Ignore if size limit is reached.
  }
}

function buildCacheKey_(prefix, monthKey, fingerprint) {
  const raw = String(prefix || '') + '|' + String(monthKey || '') + '|' + String(fingerprint || '');
  return String(prefix || 'K') + '_' + String(monthKey || '') + '_' + hashText_(raw).slice(0, 40);
}

function hashText_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text || ''), Utilities.Charset.UTF_8);
  return bytesToHex_(bytes);
}

function getPersistedClosureCache_(key) {
  return getPersistedCache_(key);
}

function setPersistedClosureCache_(key, value) {
  setPersistedCache_(key, value);
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

function ensureDefaultUsers_() {
  const users = getUsers_();
  let changed = false;

  // Legacy username normalization requested by operations.
  if (users.olga161718 && users.olga161718.active !== false) {
    users.olga161718.active = false;
    changed = true;
  }

  const defaults = [
    { username: 'user', password: '123', role: 'admin' },
    { username: 'florencia', password: '123', role: 'operator' },
    { username: 'perla', password: '456', role: 'operator' },
    { username: 'estela', password: '789', role: 'operator' },
    { username: 'gregoria', password: '101112', role: 'operator' },
    { username: 'nohelia', password: '131415', role: 'operator' },
    { username: 'olga', password: '161718', role: 'operator' },
    { username: 'claudia', password: '192021', role: 'operator' }
  ];

  for (let i = 0; i < defaults.length; i++) {
    const d = defaults[i];
    const uname = String(d.username || '').trim().toLowerCase();
    if (!uname || users[uname]) continue;
    const salt = Utilities.getUuid().replace(/-/g, '');
    users[uname] = {
      passwordHash: hashPassword_(String(d.password || ''), salt),
      salt: salt,
      role: String(d.role || 'operator'),
      active: true
    };
    changed = true;
  }
  if (changed) saveUsers_(users);
}

function ensureBackupCsvFile_() {
  const props = PropertiesService.getScriptProperties();
  const key = CONFIG.BACKUP_CSV_PROP_KEY;
  const knownId = String(props.getProperty(key) || '').trim();
  if (knownId) {
    try {
      return DriveApp.getFileById(knownId);
    } catch (_e) {
      // Continue and recreate.
    }
  }

  const header = [
    'movement_id',
    'movement_date',
    'item_code',
    'service_name',
    'move_type',
    'quantity',
    'ticket_no',
    'notes',
    'evidence_name',
    'evidence_url',
    'evidence_file_id',
    'username',
    'created_at'
  ].join(',') + '\n';

  let file = null;
  const ssFile = DriveApp.getFileById(CONFIG.SPREADSHEET_ID);
  const parents = ssFile.getParents();
  if (parents.hasNext()) {
    const folder = parents.next();
    const files = folder.getFilesByName(CONFIG.BACKUP_CSV_NAME);
    file = files.hasNext() ? files.next() : folder.createFile(CONFIG.BACKUP_CSV_NAME, header, MimeType.CSV);
  } else {
    const files = DriveApp.getFilesByName(CONFIG.BACKUP_CSV_NAME);
    file = files.hasNext() ? files.next() : DriveApp.createFile(CONFIG.BACKUP_CSV_NAME, header, MimeType.CSV);
  }

  props.setProperty(key, file.getId());
  if (file.getSize() === 0) {
    file.setContent(header);
  }
  return file;
}

function appendMovementBackupCsv_(rowData) {
  const file = ensureBackupCsvFile_();
  const line = rowData.map((v) => escapeCsvCell_(v)).join(',') + '\n';
  const current = file.getBlob().getDataAsString('UTF-8');
  file.setContent(current + line);
}

function ensureEvidenceFolder_() {
  const props = PropertiesService.getScriptProperties();
  const key = CONFIG.EVIDENCE_FOLDER_PROP_KEY;
  const knownId = String(props.getProperty(key) || '').trim();
  if (knownId) {
    try {
      return DriveApp.getFolderById(knownId);
    } catch (_e) {
      // Continue and recreate.
    }
  }

  const ssFile = DriveApp.getFileById(CONFIG.SPREADSHEET_ID);
  const parents = ssFile.getParents();
  let folder = null;
  if (parents.hasNext()) {
    const root = parents.next();
    const found = root.getFoldersByName(CONFIG.EVIDENCE_FOLDER_NAME);
    folder = found.hasNext() ? found.next() : root.createFolder(CONFIG.EVIDENCE_FOLDER_NAME);
  } else {
    const found = DriveApp.getFoldersByName(CONFIG.EVIDENCE_FOLDER_NAME);
    folder = found.hasNext() ? found.next() : DriveApp.createFolder(CONFIG.EVIDENCE_FOLDER_NAME);
  }
  props.setProperty(key, folder.getId());
  return folder;
}

function createEvidenceFile_(evidence, username) {
  const nameRaw = String((evidence && evidence.filename) || 'evidencia').trim();
  const mimeType = String((evidence && evidence.mime_type) || 'application/octet-stream').trim();
  const b64 = String((evidence && evidence.data_base64) || '').trim();
  if (!b64) throw new Error('Archivo de evidencia invalido');
  const bytes = Utilities.base64Decode(b64);
  if (!bytes || !bytes.length) throw new Error('Archivo de evidencia vacio');
  if (bytes.length > Number(CONFIG.MAX_EVIDENCE_BYTES || 0)) {
    throw new Error('El archivo de evidencia supera el limite de 8 MB');
  }
  const safeName = sanitizeFileName_(nameRaw);
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const fullName = `${stamp}_${String(username || 'user')}_${safeName}`;
  const folder = ensureEvidenceFolder_();
  const blob = Utilities.newBlob(bytes, mimeType, fullName);
  const file = folder.createFile(blob);
  return {
    fileId: file.getId(),
    url: file.getUrl(),
    name: fullName
  };
}

function sanitizeFileName_(name) {
  const s = String(name || 'evidencia').trim() || 'evidencia';
  return s
    .replace(/[\\/:*?"<>|]+/g, '_')
    .replace(/\\s+/g, '_')
    .slice(0, 120);
}

function rebuildMovementBackupCsvFromSheet_() {
  const ws = getSheet_(SHEETS.MOVEMENTS.name, SHEETS.MOVEMENTS.headers);
  const lastRow = ws.getLastRow();
  const file = ensureBackupCsvFile_();
  const header = SHEETS.MOVEMENTS.headers.join(',') + '\n';
  if (lastRow < 2) {
    file.setContent(header);
    return;
  }
  const rows = ws.getRange(2, 1, lastRow - 1, SHEETS.MOVEMENTS.headers.length).getValues();
  const body = rows.map((row) => row.map((v) => escapeCsvCell_(v)).join(',')).join('\n');
  file.setContent(header + body + '\n');
}

function escapeCsvCell_(value) {
  const raw = String(value == null ? '' : value);
  const escaped = raw.replace(/"/g, '""');
  if (/[",\n\r]/.test(escaped)) return '"' + escaped + '"';
  return escaped;
}

function diffMovementForAudit_(oldObj, newObj) {
  const keys = ['movement_date', 'item_code', 'service_name', 'move_type', 'quantity', 'ticket_no', 'notes'];
  const out = [];
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    const a = String(oldObj[k] == null ? '' : oldObj[k]);
    const b = String(newObj[k] == null ? '' : newObj[k]);
    if (a !== b) out.push(k + ':[' + a + ']=>[' + b + ']');
  }
  return out.length ? out.join(' | ') : 'sin cambios';
}

function ensureDefaultProviders_() {
  const providers = listProviders_();
  const key = String(CONFIG.DEFAULT_PROVIDER || '').toUpperCase();
  if (!key) return;
  const exists = providers.some((s) => String(s.service_name || '').toUpperCase() === key);
  if (exists) return;
  const ws = getSheet_(SHEETS.SERVICES.name, SHEETS.SERVICES.headers);
  ws.appendRow([key, true, new Date().toISOString()]);
}

function ensureDefaultConsumers_() {
  const consumers = listConsumers_();
  if (consumers.length > 0) return;

  const providers = listProviders_();
  const ws = getSheet_(SHEETS.CONSUMERS.name, SHEETS.CONSUMERS.headers);
  const nowIso = new Date().toISOString();
  if (providers.length) {
    const values = providers.map((p) => [String(p.service_name || '').toUpperCase(), true, nowIso]);
    ws.getRange(2, 1, values.length, SHEETS.CONSUMERS.headers.length).setValues(values);
    return;
  }
  ws.appendRow(['DEPENDENCIA GENERAL', true, nowIso]);
}

function requireMasterAdmin_(session) {
  if (!session || String(session.username || '').toLowerCase() !== 'user') {
    throw new Error('Solo el administrador principal puede acceder a esta configuracion.');
  }
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
        '',
        '',
        '',
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
              '',
              '',
              '',
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
            '',
            '',
            '',
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
            '',
            '',
            '',
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
  if (movements.length) {
    const normalized = movements.map((r) => normalizeImportedMovementRow_(r));
    movementsWs.getRange(2, 1, normalized.length, SHEETS.MOVEMENTS.headers.length).setValues(normalized);
  }

  audit_('IMPORT', 'HistoricalDetailed', '2026-03', `replace_data items=${items.length}, services=${services.length}, movements=${movements.length}`, 'system_import');
  return { items: items.length, services: services.length, movements: movements.length };
}

function normalizeImportedMovementRow_(row) {
  const inRow = Array.isArray(row) ? row : [];
  const out = inRow.slice(0, SHEETS.MOVEMENTS.headers.length);
  while (out.length < SHEETS.MOVEMENTS.headers.length) out.push('');
  return out;
}
