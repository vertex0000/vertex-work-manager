/**
 * Vertex Work Manager: Google Apps Script backend
 * Sheet ke andar: Extensions > Apps Script > is file ko paste karo.
 * Project Settings > Script properties > PIN = apna PIN
 *
 * Rate kaise decide hota hai (har nayi entry pe):
 *   1. Client ka special price (Client Prices tab)  -> agar set hai
 *   2. Client ka Discount % (Clients tab)           -> standard price pe % kam
 *   3. Standard price (Price List tab)              -> baaki sab clients
 * Entry ka rate "Saved rate" column mein lock ho jata hai, taaki price
 * badalne se purana hisaab na badle.
 */

const MASTER = 'Master';
const PRICES = 'Price List';
const PLANS = 'Plans';
const CLIENTS = 'Clients';
const CLIENT_PRICES = 'Client Prices';
const SETTINGS = 'Settings';
const CELL_LIMIT = 49000; // Google Sheet cell mein max ~50,000 characters

// Master tab columns (1-based)
const C = {
  date: 1, client: 2, phone: 3, work: 4, qty: 5, changes: 6,
  rate: 7, amount: 8, status: 9, payment: 10, notes: 11,
  month: 12, id: 13, saved: 14
};

function doGet() {
  return json({ ok: true, msg: 'Vertex API is running' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: 'Bad request' });
  }

  const pin = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!pin) return json({ ok: false, error: 'PIN not set. Script properties mein PIN add karo.' });
  if (String(req.pin).trim() !== String(pin).trim()) return json({ ok: false, error: 'Wrong PIN' });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = req.data || {};
    switch (req.action) {
      case 'getAll': break;
      case 'addEntry': addEntry(d); break;
      case 'updateEntry': updateEntry(d); break;
      case 'deleteEntry': deleteEntry(d.id); break;
      case 'addService': addService(d); break;
      case 'updateService': updateService(d); break;
      case 'deleteService': deleteService(d.name); break;
      case 'saveClient': saveClient(d); break;
      case 'deleteClient': deleteClient(d.name); break;
      case 'saveSettings': saveSettings(d); break;
      case 'savePlans': savePlans(d.rows); break;
      default: return json({ ok: false, error: 'Unknown action' });
    }
    SpreadsheetApp.flush();
    return json({ ok: true, data: getAll() });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- helpers ---------- */

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function book() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tab(name) {
  const s = book().getSheetByName(name);
  if (!s) throw new Error('Tab not found: ' + name);
  return s;
}
function num(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function isNum(v) { return v !== '' && v !== null && v !== undefined && !isNaN(Number(v)); }
function blankOrNum(v) { return isNum(v) ? Number(v) : ''; }
function low(s) { return String(s || '').trim().toLowerCase(); }
function tz() { return Session.getScriptTimeZone(); }

function toDate(s) {
  if (!s) return '';
  const p = String(s).split('-');
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}
function fmtDate(v) {
  if (!v) return '';
  const d = v instanceof Date ? v : new Date(v);
  return isNaN(d) ? String(v) : Utilities.formatDate(d, tz(), 'yyyy-MM-dd');
}
function newId() {
  return 'E' + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase();
}

// Last row that has something in the given column (ignores empty formula output)
function lastRowIn(sheet, col) {
  const n = sheet.getMaxRows();
  if (n < 2) return 1;
  const vals = sheet.getRange(2, col, n - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '' && vals[i][0] !== null) return i + 2;
  }
  return 1;
}

function ensureTab(name, headers) {
  let s = book().getSheetByName(name);
  if (!s) {
    s = book().insertSheet(name);
    s.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#141414').setFontColor('#ffffff');
    s.setFrozenRows(1);
  }
  if (s.getMaxColumns() < headers.length) s.insertColumnsAfter(s.getMaxColumns(), headers.length - s.getMaxColumns());
  const head = s.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach((h, i) => { if (!head[i]) s.getRange(1, i + 1).setValue(h); });
  return s;
}
function clientsTab() { return ensureTab(CLIENTS, ['Client name', 'Phone', 'Photo (set from website)', 'Notes', 'Discount %']); }
function cpTab() { return ensureTab(CLIENT_PRICES, ['Client name', 'Service', 'Special price (₹)']); }
function settingsTab() { return ensureTab(SETTINGS, ['Setting', 'Value']); }

/* ---------- prices ---------- */

function priceMap() {
  const s = tab(PRICES);
  const last = lastRowIn(s, 3);
  const m = {};
  if (last < 2) return m;
  s.getRange(2, 3, last - 1, 2).getValues().forEach(r => {
    if (r[0] !== '') m[String(r[0]).trim()] = (r[1] === '' ? '' : num(r[1]));
  });
  return m;
}

function getClientPrices() {
  const s = cpTab();
  const last = lastRowIn(s, 1);
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 3).getValues()
    .filter(r => r[0] !== '' && r[1] !== '' && isNum(r[2]))
    .map(r => ({ client: String(r[0]).trim(), service: String(r[1]).trim(), price: num(r[2]) }));
}

function clientDiscount(name) {
  const row = findClientRow(name);
  if (row < 0) return 0;
  const v = clientsTab().getRange(row, 5).getValue();
  return Math.max(0, Math.min(100, num(v)));
}

// Is client ke liye is service ka rate
function effectiveRate(client, work) {
  const sp = getClientPrices().find(p => low(p.client) === low(client) && p.service === String(work).trim());
  if (sp) return sp.price;
  const std = priceMap()[String(work).trim()];
  if (std === undefined || std === '') return '';
  const disc = clientDiscount(client);
  return disc > 0 ? Math.round(std * (1 - disc / 100)) : std;
}

// Old entries keep their old price: save current rate on rows that don't have one yet
function lockRates(workName, rate) {
  const s = tab(MASTER);
  const last = lastRowIn(s, C.client);
  if (last < 2 || rate === '' || rate === undefined) return;
  const rng = s.getRange(2, C.work, last - 1, C.saved - C.work + 1);
  const vals = rng.getValues();
  const savedCol = C.saved - C.work;
  let changed = false;
  vals.forEach(r => {
    if (String(r[0]).trim() === workName && (r[savedCol] === '' || r[savedCol] === null)) {
      r[savedCol] = rate;
      changed = true;
    }
  });
  if (changed) s.getRange(2, C.saved, last - 1, 1).setValues(vals.map(r => [r[savedCol]]));
}

/* ---------- read ---------- */

function getAll() {
  const m = tab(MASTER);
  const last = lastRowIn(m, C.client);
  let entries = [];

  if (last >= 2) {
    const rows = m.getRange(2, 1, last - 1, C.saved).getValues();
    const ids = [];
    let idsChanged = false;
    rows.forEach(r => {
      let id = r[C.id - 1];
      if (r[C.client - 1] !== '' && !id) { id = newId(); idsChanged = true; }
      ids.push([id || '']);
      r[C.id - 1] = id;
    });
    if (idsChanged) m.getRange(2, C.id, ids.length, 1).setValues(ids);

    entries = rows
      .filter(r => r[C.client - 1] !== '')
      .map(r => ({
        id: String(r[C.id - 1]),
        date: fmtDate(r[C.date - 1]),
        client: String(r[C.client - 1]).trim(),
        phone: String(r[C.phone - 1] || ''),
        work: String(r[C.work - 1] || ''),
        qty: r[C.qty - 1] === '' ? '' : num(r[C.qty - 1]),
        changes: r[C.changes - 1] === '' ? '' : num(r[C.changes - 1]),
        rate: num(r[C.rate - 1]),
        amount: num(r[C.amount - 1]),
        status: String(r[C.status - 1] || 'Pending'),
        payment: String(r[C.payment - 1] || 'Unpaid'),
        notes: String(r[C.notes - 1] || '')
      }));
  }

  const p = tab(PRICES);
  const pLast = lastRowIn(p, 3);
  const services = pLast < 2 ? [] : p.getRange(2, 1, pLast - 1, 4).getValues()
    .filter(r => r[2] !== '')
    .map(r => ({ name: String(r[2]).trim(), details: String(r[1] || '').trim(), rate: r[3] === '' ? null : num(r[3]) }));

  let plans = [];
  try { plans = tab(PLANS).getRange('A3:D15').getDisplayValues(); } catch (e) { plans = []; }

  return {
    entries, services, plans,
    clients: getClients(),
    clientPrices: getClientPrices(),
    settings: getSettings(),
    sheetUrl: book().getUrl()
  };
}

/* ---------- entries ---------- */

function writeEntry(row, d, savedRate) {
  const s = tab(MASTER);
  s.getRange(row, C.date, 1, 6).setValues([[
    toDate(d.date),
    String(d.client || '').trim(),
    d.phone ? "'" + String(d.phone).trim() : '',
    d.work || '',
    blankOrNum(d.qty),
    blankOrNum(d.changes)
  ]]);
  s.getRange(row, C.status, 1, 3).setValues([[d.status || 'Pending', d.payment || 'Unpaid', d.notes || '']]);
  s.getRange(row, C.saved).setValue(savedRate === undefined || savedRate === null ? '' : savedRate);
}

function addEntry(d) {
  if (!d.client) throw new Error('Client name zaroori hai');
  if (!d.work) throw new Error('Work type chuno');
  const s = tab(MASTER);
  const row = lastRowIn(s, C.client) + 1;
  if (row > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 100);
  const rate = isNum(d.rate) ? Number(d.rate) : effectiveRate(d.client, d.work);
  writeEntry(row, d, rate);
  s.getRange(row, C.id).setValue(newId());
}

function findEntryRow(id) {
  const s = tab(MASTER);
  const last = lastRowIn(s, C.client);
  if (last >= 2) {
    const ids = s.getRange(2, C.id, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  }
  throw new Error('Entry nahi mili. Refresh karke dobara try karo.');
}

function updateEntry(d) {
  const s = tab(MASTER);
  const row = findEntryRow(d.id);
  const oldWork = String(s.getRange(row, C.work).getValue()).trim();
  const oldClient = String(s.getRange(row, C.client).getValue()).trim();
  let saved = s.getRange(row, C.saved).getValue();
  if (isNum(d.rate)) saved = Number(d.rate);
  else if (oldWork !== d.work || low(oldClient) !== low(d.client)) saved = effectiveRate(d.client, d.work);
  writeEntry(row, d, saved);
}

function deleteEntry(id) {
  tab(MASTER).deleteRow(findEntryRow(id));
}

/* ---------- services (Price List) ---------- */

function findServiceRow(name) {
  const s = tab(PRICES);
  const last = lastRowIn(s, 3);
  if (last >= 2) {
    const names = s.getRange(2, 3, last - 1, 1).getValues();
    for (let i = 0; i < names.length; i++) if (low(names[i][0]) === low(name)) return i + 2;
  }
  return -1;
}

function rateFormula(row) {
  return '=IFERROR(VALUE(REGEXEXTRACT(TO_TEXT(B' + row + '),"\\d+")),"")';
}

function addService(d) {
  const name = String(d.name || '').trim();
  if (!name) throw new Error('Service name zaroori hai');
  if (findServiceRow(name) > 0) throw new Error('Ye service pehle se hai');
  const price = num(d.price);
  const s = tab(PRICES);
  const row = lastRowIn(s, 3) + 1;
  if (row > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 20);
  s.getRange(row, 1, 1, 3).setValues([[name, price + '/- ' + String(d.details || '').trim(), name]]);
  s.getRange(row, 4).setFormula(rateFormula(row));
  s.getRange(2, 1, 1, 4).copyTo(s.getRange(row, 1, 1, 4), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

function renameInColumn(sheet, col, oldName, newName) {
  const last = lastRowIn(sheet, col);
  if (last < 2) return;
  const rng = sheet.getRange(2, col, last - 1, 1);
  rng.setValues(rng.getValues().map(r => [String(r[0]).trim() === oldName ? newName : r[0]]));
}

function updateService(d) {
  const oldName = String(d.oldName || '').trim();
  const name = String(d.name || '').trim();
  if (!name) throw new Error('Service name zaroori hai');
  const s = tab(PRICES);
  const row = findServiceRow(oldName);
  if (row < 0) throw new Error('Service nahi mili. Refresh karo.');
  if (low(name) !== low(oldName) && findServiceRow(name) > 0) throw new Error('Is naam ki service pehle se hai');

  lockRates(oldName, priceMap()[oldName]);

  const oldA = String(s.getRange(row, 1).getValue()).trim();
  const oldB = String(s.getRange(row, 2).getValue());
  const price = num(d.price);
  const details = String(d.details === undefined ? oldB : d.details);
  let newB;
  if (details.trim() === oldB.trim() && /\d+/.test(oldB)) newB = oldB.replace(/\d+/, String(price));
  else newB = price + '/- ' + details.replace(/^\s*\d+\s*(\/-)?\s*/, '').trim();

  s.getRange(row, 2).setValue(newB);
  s.getRange(row, 3).setValue(name);
  if (oldA === oldName) s.getRange(row, 1).setValue(name);
  s.getRange(row, 4).setFormula(rateFormula(row));

  if (name !== oldName) {
    renameInColumn(tab(MASTER), C.work, oldName, name);
    renameInColumn(cpTab(), 2, oldName, name);
  }
}

function deleteRowsWhere(sheet, test) {
  const last = lastRowIn(sheet, 1);
  if (last < 2) return;
  const vals = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (test(vals[i])) sheet.deleteRow(i + 2);
}

function deleteService(name) {
  const row = findServiceRow(name);
  if (row < 0) throw new Error('Service nahi mili');
  const n = String(name).trim();
  lockRates(n, priceMap()[n]);
  tab(PRICES).deleteRow(row);
  deleteRowsWhere(cpTab(), r => String(r[1]).trim() === n);
}

/* ---------- clients (photo, phone, notes, discount, special prices) ---------- */

function getClients() {
  const s = clientsTab();
  const last = lastRowIn(s, 1);
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 5).getValues()
    .filter(r => r[0] !== '')
    .map(r => ({
      name: String(r[0]).trim(), phone: String(r[1] || ''), photo: String(r[2] || ''),
      notes: String(r[3] || ''), discount: num(r[4])
    }));
}

function findClientRow(name) {
  const s = clientsTab();
  const last = lastRowIn(s, 1);
  if (last >= 2) {
    const names = s.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < names.length; i++) if (low(names[i][0]) === low(name)) return i + 2;
  }
  return -1;
}

function saveClient(d) {
  const name = String(d.name || '').trim();
  if (!name) throw new Error('Client name zaroori hai');
  const oldName = String(d.oldName || name).trim();
  if (d.photo && String(d.photo).length > CELL_LIMIT) throw new Error('Photo bahut badi hai, choti photo try karo');
  const s = clientsTab();
  if (low(name) !== low(oldName) && findClientRow(name) > 0) throw new Error('Is naam ka client pehle se hai');

  let row = findClientRow(oldName);
  if (row < 0) {
    row = lastRowIn(s, 1) + 1;
    if (row > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 50);
  }
  const cur = s.getRange(row, 1, 1, 5).getValues()[0];
  s.getRange(row, 1, 1, 5).setValues([[
    name,
    d.phone !== undefined ? (d.phone ? "'" + String(d.phone).trim() : '') : cur[1],
    d.photo !== undefined && d.photo !== null ? d.photo : cur[2],
    d.notes !== undefined ? d.notes : cur[3],
    d.discount !== undefined ? (isNum(d.discount) && Number(d.discount) > 0 ? Math.min(100, Number(d.discount)) : '') : cur[4]
  ]]);

  const cp = cpTab();
  if (name !== oldName) {
    renameInColumn(tab(MASTER), C.client, oldName, name);
    renameInColumn(cp, 1, oldName, name);
  }

  // special prices: website se puri list aati hai, purani hata ke nayi likho
  if (Array.isArray(d.prices)) {
    deleteRowsWhere(cp, r => low(r[0]) === low(name));
    const rows = d.prices
      .filter(p => p && p.service && isNum(p.price))
      .map(p => [name, String(p.service).trim(), Number(p.price)]);
    if (rows.length) {
      const start = lastRowIn(cp, 1) + 1;
      if (start + rows.length > cp.getMaxRows()) cp.insertRowsAfter(cp.getMaxRows(), rows.length + 50);
      cp.getRange(start, 1, rows.length, 3).setValues(rows);
    }
  }
}

function deleteClient(name) {
  const row = findClientRow(name);
  if (row > 0) clientsTab().deleteRow(row);
  deleteRowsWhere(cpTab(), r => low(r[0]) === low(name));
}

/* ---------- settings (brand, logo, colour, target) ---------- */

function getSettings() {
  const s = settingsTab();
  const last = lastRowIn(s, 1);
  const o = {};
  if (last >= 2) s.getRange(2, 1, last - 1, 2).getValues().forEach(r => { if (r[0] !== '') o[String(r[0])] = r[1]; });
  return o;
}

function saveSettings(d) {
  const s = settingsTab();
  Object.keys(d).forEach(k => {
    let v = d[k];
    if (typeof v === 'object' && v !== null) v = JSON.stringify(v);
    if (String(v).length > CELL_LIMIT) throw new Error('Logo/image bahut badi hai, choti image try karo');
    const last = lastRowIn(s, 1);
    let row = -1;
    if (last >= 2) {
      const keys = s.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) if (String(keys[i][0]) === k) { row = i + 2; break; }
    }
    if (row < 0) row = last + 1;
    s.getRange(row, 1, 1, 2).setValues([[k, v]]);
  });
}

/* ---------- plans (text rows only, price rows are formulas) ---------- */

function savePlans(rows) {
  if (!rows || !rows.length) return;
  // Plans!A3:D13 = header + feature rows. Row 14-15 (price) formulas hain, unhe nahi chhedte.
  const data = rows.slice(0, 11).map(r => [0, 1, 2, 3].map(i => (r[i] === undefined ? '' : String(r[i]))));
  while (data.length < 11) data.push(['', '', '', '']);
  tab(PLANS).getRange(3, 1, 11, 4).setValues(data);
}
