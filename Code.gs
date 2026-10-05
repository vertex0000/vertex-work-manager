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
 *
 * Login (sirf PIN):
 *   Owner  -> Script properties > PIN
 *   Team   -> website ke Team page se banta hai. Team tab mein PIN hash ho ke
 *             save hota hai (asli PIN kahin nahi dikhta). Har member ka PIN alag.
 *
 * Files & links: Master (col P) aur Clients (col F) mein, ek line = "Label | https://..."
 *   "[client] " se shuru hone wali line client portal pe bhi dikhti hai.
 *
 * Client portal: har client ka alag link (Clients col G) + PIN (col H, hash).
 *   Client sirf apna kaam, client-visible files aur (on ho to) amount dekhta hai.
 *   Approve / Changes feedback Master col Q mein aata hai.
 */

const MASTER = 'Master';
const PRICES = 'Price List';
const PLANS = 'Plans';
const CLIENTS = 'Clients';
const CLIENT_PRICES = 'Client Prices';
const SETTINGS = 'Settings';
const TEAM = 'Team';
const MAX_FAILS = 15;      // itne galat PIN ke baad 10 minute lock
const CELL_LIMIT = 49000; // Google Sheet cell mein max ~50,000 characters

// Master tab columns (1-based)
const C = {
  date: 1, client: 2, phone: 3, work: 4, qty: 5, changes: 6,
  rate: 7, amount: 8, status: 9, payment: 10, notes: 11,
  month: 12, id: 13, saved: 14, by: 15, links: 16, feedback: 17
};

// Kaun kya kar sakta hai
const OWNER_ONLY = ['saveMember', 'deleteMember', 'saveSettings'];
const MEMBER_OK = ['getAll', 'addEntry', 'updateEntry', 'deleteEntry', 'saveClient'];
// Clients tab portal columns
const CP = { token: 7, pin: 8, on: 9, amounts: 10 };

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

  if (req.portal) return portalRequest(req);

  let me;
  try { me = auth(req.pin); } catch (err) { return json({ ok: false, error: String(err.message || err) }); }

  const action = String(req.action || '');
  if (me.role !== 'owner' && OWNER_ONLY.indexOf(action) >= 0) return json({ ok: false, error: 'Ye sirf owner kar sakta hai' });
  if (me.role === 'member' && MEMBER_OK.indexOf(action) < 0) return json({ ok: false, error: 'Is kaam ki permission nahi hai' });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = req.data || {};
    switch (action) {
      case 'getAll': break;
      case 'addEntry': addEntry(d, me); break;
      case 'updateEntry': updateEntry(d, me); break;
      case 'deleteEntry': deleteEntry(d.id, me); break;
      case 'addService': addService(d); break;
      case 'updateService': updateService(d); break;
      case 'deleteService': deleteService(d.name); break;
      case 'saveClient': saveClient(d, me); break;
      case 'saveMember': saveMember(d); break;
      case 'deleteMember': deleteMember(d.name); break;
      case 'deleteClient': deleteClient(d.name); break;
      case 'saveSettings': saveSettings(d); break;
      case 'savePlans': savePlans(d.rows); break;
      case 'setPortal': setPortal(d); break;
      default: return json({ ok: false, error: 'Unknown action' });
    }
    SpreadsheetApp.flush();
    return json({ ok: true, data: getAll(me) });
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
function clientsTab() { return ensureTab(CLIENTS, ['Client name', 'Phone', 'Photo (set from website)', 'Notes', 'Discount %', 'Files & links', 'Portal link code', 'Portal PIN (locked)', 'Portal on', 'Show amounts']); }
function cpTab() { return ensureTab(CLIENT_PRICES, ['Client name', 'Service', 'Special price (₹)']); }
function settingsTab() { return ensureTab(SETTINGS, ['Setting', 'Value']); }

/* ---------- login: owner PIN (Script properties) ya team member PIN (Team tab) ---------- */

function props() { return PropertiesService.getScriptProperties(); }
function salt() {
  let s = props().getProperty('SALT');
  if (!s) { s = Utilities.getUuid(); props().setProperty('SALT', s); }
  return s;
}
function hashPin(pin) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt() + '|' + String(pin).trim());
  return 'h:' + bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function ownerPin() {
  const p = props().getProperty('PIN');
  if (!p) throw new Error('PIN not set. Script properties mein PIN add karo.');
  return String(p).trim();
}
function ownerName() { const v = getSettings().ownerName; return v ? String(v).trim() : 'Owner'; }

function auth(pin) {
  pin = String(pin || '').trim();
  const owner = ownerPin();
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinFails') || 0);
  if (fails >= MAX_FAILS) throw new Error('Bahut baar galat PIN daala gaya. 10 minute baad try karo.');
  const ok = me => { if (fails) cache.put('pinFails', '0', 1); return me; };
  if (pin && pin === owner) return ok({ name: ownerName(), role: 'owner', view: 'all', money: true });
  if (pin) {
    const h = hashPin(pin);
    const m = readTeam().find(t => t.hash === h);
    if (m) {
      if (!m.active) throw new Error('Tumhara access band hai. Owner se baat karo.');
      return ok({ name: m.name, role: m.role, view: m.role === 'admin' ? 'all' : m.view, money: m.role === 'admin' ? true : m.money });
    }
  }
  cache.put('pinFails', String(fails + 1), 600);
  throw new Error('Wrong PIN');
}

/* ---------- team (Team tab) ---------- */

function teamTab() { return ensureTab(TEAM, ['Name', 'PIN (locked)', 'Role', 'Can see', 'Money', 'Active', 'Phone', 'Added on']); }
function yes(v) { return low(v) === 'yes' || v === true; }
function readTeam() {
  const s = teamTab();
  const last = lastRowIn(s, 1);
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 8).getValues().filter(r => r[0] !== '').map(r => ({
    name: String(r[0]).trim(), hash: String(r[1] || ''), role: low(r[2]) === 'admin' ? 'admin' : 'member',
    view: low(r[3]) === 'all' ? 'all' : 'own', money: yes(r[4]), active: r[5] === '' ? true : yes(r[5]),
    phone: String(r[6] || ''), added: fmtDate(r[7])
  }));
}
function findMemberRow(name) {
  const s = teamTab();
  const last = lastRowIn(s, 1);
  if (last >= 2) {
    const names = s.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < names.length; i++) if (low(names[i][0]) === low(name)) return i + 2;
  }
  return -1;
}

function saveMember(d) {
  const name = String(d.name || '').trim();
  const oldName = String(d.oldName || '').trim();
  if (!name) throw new Error('Member ka naam zaroori hai');
  if (low(name) === low(ownerName())) throw new Error('Ye naam owner ka hai, doosra naam rakho');
  const s = teamTab();
  let row = oldName ? findMemberRow(oldName) : -1;
  if (oldName && row < 0) throw new Error('Member nahi mila. Refresh karo.');
  const other = findMemberRow(name);
  if (other > 0 && other !== row) throw new Error('Is naam ka member pehle se hai');

  let hash = row > 0 ? String(s.getRange(row, 2).getValue()) : '';
  const pin = String(d.pin || '').trim();
  if (pin) {
    if (!/^\d{6,12}$/.test(pin)) throw new Error('PIN 6 se 12 number ka hona chahiye');
    if (pin === ownerPin()) throw new Error('Ye PIN owner ka hai, doosra PIN rakho');
    const h = hashPin(pin);
    if (readTeam().some(t => t.hash === h && low(t.name) !== low(oldName))) throw new Error('Ye PIN kisi aur member ka hai, doosra PIN rakho');
    hash = h;
  }
  if (!hash) throw new Error('Naye member ke liye PIN zaroori hai');

  if (row < 0) {
    row = lastRowIn(s, 1) + 1;
    if (row > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 20);
  }
  const added = row > 0 && s.getRange(row, 8).getValue() ? s.getRange(row, 8).getValue() : new Date();
  s.getRange(row, 1, 1, 8).setValues([[
    name, hash, d.role === 'admin' ? 'Admin' : 'Member', d.view === 'all' ? 'All' : 'Own',
    d.money ? 'Yes' : 'No', d.active === false ? 'No' : 'Yes', d.phone ? "'" + String(d.phone).trim() : '', added
  ]]);
  if (oldName && name !== oldName) renameInColumn(tab(MASTER), C.by, oldName, name);
}

function deleteMember(name) {
  const row = findMemberRow(name);
  if (row < 0) throw new Error('Member nahi mila');
  teamTab().deleteRow(row);
}

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

function masterTab() {
  const m = tab(MASTER);
  if (m.getMaxColumns() < C.feedback) m.insertColumnsAfter(m.getMaxColumns(), C.feedback - m.getMaxColumns());
  if (!m.getRange(1, C.by).getValue()) m.getRange(1, C.by).setValue('Added by');
  if (!m.getRange(1, C.links).getValue()) m.getRange(1, C.links).setValue('Files & links');
  if (!m.getRange(1, C.feedback).getValue()) m.getRange(1, C.feedback).setValue('Client feedback');
  return m;
}

/* ---------- files & links ---------- */

// Website se list aati hai [{label,url}], sheet mein "Label | https://..." har line pe
function parseLinkLines(v) {
  return String(v || '').split('\n').map(l => {
    let client = false;
    l = l.trim();
    if (/^\[client\]\s*/i.test(l)) { client = true; l = l.replace(/^\[client\]\s*/i, ''); }
    const i = l.lastIndexOf(' | ');
    return i >= 0 ? { label: l.slice(0, i), url: l.slice(i + 3), client } : { label: '', url: l, client };
  });
}
function cleanLinks(v) {
  const arr = Array.isArray(v) ? v : parseLinkLines(v);
  return arr
    .map(x => ({ label: String((x && x.label) || '').replace(/[|\r\n\[\]]+/g, ' ').trim().slice(0, 60), url: String((x && x.url) || '').trim(), client: !!(x && x.client) }))
    .filter(x => /^https?:\/\/[^\s]+$/i.test(x.url))
    .slice(0, 30)
    .map(x => (x.client ? '[client] ' : '') + (x.label ? x.label + ' | ' : '') + x.url)
    .join('\n');
}

function getAll(me) {
  me = me || { name: 'Owner', role: 'owner', view: 'all', money: true };
  const m = masterTab();
  const last = lastRowIn(m, C.client);
  let entries = [];

  if (last >= 2) {
    const rows = m.getRange(2, 1, last - 1, C.feedback).getValues();
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
        notes: String(r[C.notes - 1] || ''),
        by: String(r[C.by - 1] || '').trim(),
        links: String(r[C.links - 1] || ''),
        feedback: String(r[C.feedback - 1] || '')
      }));
  }

  const owner = ownerName();
  entries.forEach(e => { if (!e.by) e.by = owner; });
  if (me.role === 'member' && me.view !== 'all') entries = entries.filter(e => low(e.by) === low(me.name));
  if (!me.money) entries.forEach(e => { e.rate = null; e.amount = null; e.changes = ''; e.payment = ''; });

  const p = tab(PRICES);
  const pLast = lastRowIn(p, 3);
  const services = pLast < 2 ? [] : p.getRange(2, 1, pLast - 1, 4).getValues()
    .filter(r => r[2] !== '')
    .map(r => ({ name: String(r[2]).trim(), details: String(r[1] || '').trim(), rate: r[3] === '' ? null : num(r[3]) }));

  let plans = [];
  try { plans = tab(PLANS).getRange('A3:D15').getDisplayValues(); } catch (e) { plans = []; }
  if (!me.money) {                                  // paisa nahi dikhana to price bhi nahi
    services.forEach(sv => { sv.rate = null; sv.details = ''; });
    plans = [];
  }

  const team = readTeam();
  let clients = getClients();
  if (me.role === 'member') clients = clients.map(c => { const x = Object.assign({}, c); delete x.portal; return x; });
  if (!me.money) clients = clients.map(c => Object.assign({}, c, { discount: 0 }));
  return {
    entries, services, plans, clients,
    clientPrices: me.money ? getClientPrices() : [],
    settings: getSettings(),
    me: me,
    team: me.role === 'owner'
      ? team.map(t => ({ name: t.name, role: t.role, view: t.view, money: t.money, active: t.active, phone: t.phone, added: t.added }))
      : team.filter(t => t.active).map(t => ({ name: t.name, role: t.role })),
    sheetUrl: me.role === 'member' ? '' : book().getUrl()
  };
}

/* ---------- entries ---------- */

function writeEntry(row, d, savedRate) {
  const s = masterTab();
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
  if (d.links !== undefined) s.getRange(row, C.links).setValue(cleanLinks(d.links));
}

function addEntry(d, me) {
  me = me || { role: 'owner', money: true, name: '' };
  if (!d.client) throw new Error('Client name zaroori hai');
  if (!d.work) throw new Error('Work type chuno');
  const s = masterTab();
  const row = lastRowIn(s, C.client) + 1;
  if (row > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 100);
  if (me.role === 'member') {
    d.rate = '';                                   // member rate khud set nahi karta
    if (!me.money) { d.payment = 'Unpaid'; d.changes = ''; }
  }
  const rate = isNum(d.rate) ? Number(d.rate) : effectiveRate(d.client, d.work);
  writeEntry(row, d, rate);
  s.getRange(row, C.id).setValue(newId());
  s.getRange(row, C.by).setValue(me.name || ownerName());
}

function entryOwner(row) {
  const v = String(tab(MASTER).getRange(row, C.by).getValue() || '').trim();
  return v || ownerName();
}
function assertMine(row, me) {
  if (me && me.role === 'member' && low(entryOwner(row)) !== low(me.name)) throw new Error('Ye entry kisi aur ne add ki hai, tum ise badal nahi sakte');
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

function updateEntry(d, me) {
  const s = masterTab();
  const row = findEntryRow(d.id);
  assertMine(row, me);
  if (me && me.role === 'member') {
    d.rate = '';
    if (!me.money) {
      d.payment = s.getRange(row, C.payment).getValue() || 'Unpaid';
      const ch = s.getRange(row, C.changes).getValue();
      d.changes = ch === '' ? '' : ch;
    }
  }
  const oldWork = String(s.getRange(row, C.work).getValue()).trim();
  const oldClient = String(s.getRange(row, C.client).getValue()).trim();
  const oldStatus = String(s.getRange(row, C.status).getValue()).trim();
  const fb = String(s.getRange(row, C.feedback).getValue() || '');
  if (d.status === 'Delivered' && oldStatus !== 'Delivered' && /^changes/i.test(fb)) s.getRange(row, C.feedback).setValue(''); // dobara deliver -> client phir se review karega
  let saved = s.getRange(row, C.saved).getValue();
  if (isNum(d.rate)) saved = Number(d.rate);
  else if (oldWork !== d.work || low(oldClient) !== low(d.client)) saved = effectiveRate(d.client, d.work);
  writeEntry(row, d, saved);
}

function deleteEntry(id, me) {
  const row = findEntryRow(id);
  assertMine(row, me);
  tab(MASTER).deleteRow(row);
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
  return s.getRange(2, 1, last - 1, 10).getValues()
    .filter(r => r[0] !== '')
    .map(r => ({
      name: String(r[0]).trim(), phone: String(r[1] || ''), photo: String(r[2] || ''),
      notes: String(r[3] || ''), discount: num(r[4]), links: String(r[5] || ''),
      portal: { token: String(r[CP.token - 1] || ''), hasPin: !!r[CP.pin - 1], on: yes(r[CP.on - 1]), amounts: r[CP.amounts - 1] === '' ? true : yes(r[CP.amounts - 1]) }
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

function saveClient(d, me) {
  if (me && me.role === 'member') {
    delete d.discount; delete d.prices;            // pricing sirf owner/admin
    if (d.oldName && low(d.oldName) !== low(d.name)) throw new Error('Client ka naam sirf owner badal sakta hai');
  }
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
  const cur = s.getRange(row, 1, 1, 6).getValues()[0];
  s.getRange(row, 1, 1, 6).setValues([[
    name,
    d.phone !== undefined ? (d.phone ? "'" + String(d.phone).trim() : '') : cur[1],
    d.photo !== undefined && d.photo !== null ? d.photo : cur[2],
    d.notes !== undefined ? d.notes : cur[3],
    d.discount !== undefined ? (isNum(d.discount) && Number(d.discount) > 0 ? Math.min(100, Number(d.discount)) : '') : cur[4],
    d.links !== undefined ? cleanLinks(d.links) : cur[5]
  ]]);

  const cp = cpTab();
  if (name !== oldName) {
    renameInColumn(tab(MASTER), C.client, oldName, name);
    renameInColumn(cp, 1, oldName, name);
  }
  if (me && me.role === 'member') return;

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

/* ---------- client portal ---------- */

// Owner/admin: portal on/off, PIN, amount dikhana, naya link
function setPortal(d) {
  const name = String(d.client || '').trim();
  if (!name) throw new Error('Client chuno');
  let row = findClientRow(name);
  if (row < 0) { saveClient({ name: name }); row = findClientRow(name); }
  const s = clientsTab();
  const cur = s.getRange(row, 1, 1, 10).getValues()[0];
  let token = String(cur[CP.token - 1] || '');
  if (!token || d.newLink) token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  let pinHash = String(cur[CP.pin - 1] || '');
  const pin = String(d.pin || '').trim();
  if (pin) {
    if (!/^\d{4,8}$/.test(pin)) throw new Error('Client PIN 4 se 8 number ka ho');
    pinHash = hashPin('client|' + pin);
  }
  if (d.on && !pinHash) throw new Error('Portal on karne ke liye PIN set karo');
  s.getRange(row, CP.token, 1, 4).setValues([[
    token, pinHash,
    d.on ? 'Yes' : 'No',
    d.amounts === false ? 'No' : 'Yes'
  ]]);
}

function portalClient(token) {
  token = String(token || '').trim();
  if (!/^[a-f0-9]{32,48}$/i.test(token)) return null;
  const s = clientsTab();
  const last = lastRowIn(s, 1);
  if (last < 2) return null;
  const rows = s.getRange(2, 1, last - 1, 10).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][CP.token - 1]) === token) return { row: i + 2, r: rows[i] };
  }
  return null;
}

function portalRequest(req) {
  const cache = CacheService.getScriptCache();
  const key = 'pf_' + String(req.portal).slice(0, 48);
  const fails = Number(cache.get(key) || 0);
  if (fails >= 10) return json({ ok: false, error: 'Bahut baar galat PIN. 15 minute baad try karo.' });
  const found = portalClient(req.portal);
  const bad = msg => { cache.put(key, String(fails + 1), 900); return json({ ok: false, error: msg }); };
  if (!found) return bad('Link sahi nahi hai. Naya link maango.');
  const r = found.r;
  if (!yes(r[CP.on - 1])) return json({ ok: false, error: 'Portal abhi band hai. Agency se baat karo.' });
  if (!r[CP.pin - 1] || hashPin('client|' + String(req.pin || '').trim()) !== String(r[CP.pin - 1])) return bad('Wrong PIN');
  if (fails) cache.put(key, '0', 1);

  const client = String(r[0]).trim();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (req.action === 'portalFeedback') portalFeedback(client, req.data || {});
    else if (req.action !== 'portalGet') return json({ ok: false, error: 'Unknown action' });
    SpreadsheetApp.flush();
    return json({ ok: true, data: portalData(client, r) });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function portalData(client, r) {
  const amounts = r[CP.amounts - 1] === '' ? true : yes(r[CP.amounts - 1]);
  const st = getSettings();
  const all = getAll({ name: 'portal', role: 'owner', view: 'all', money: true }).entries;
  const entries = all.filter(e => low(e.client) === low(client)).map(e => ({
    id: e.id, date: e.date, work: e.work, qty: e.qty, status: e.status, feedback: e.feedback,
    amount: amounts ? e.amount : null, payment: amounts ? e.payment : null,
    files: parseLinkLines(e.links).filter(l => l.client && /^https?:\/\//i.test(l.url)).map(l => ({ label: l.label, url: l.url }))
  }));
  return {
    client: { name: client, photo: String(r[2] || '') },
    amounts: amounts,
    entries: entries,
    brand: { brandName: st.brandName || 'Vertex Media', logo: st.logo || '', accent: st.accent || '', ownerName: st.ownerName || '', phone: st.contactPhone || '', email: st.contactEmail || '', instagram: st.instagram || '' },
    upi: amounts && st.upiId ? { id: String(st.upiId), name: String(st.upiName || st.brandName || '') } : null
  };
}

function portalFeedback(client, d) {
  const s = masterTab();
  const row = findEntryRow(d.id);
  if (low(s.getRange(row, C.client).getValue()) !== low(client)) throw new Error('Ye kaam aapka nahi hai');
  if (String(s.getRange(row, C.status).getValue()) !== 'Delivered') throw new Error('Kaam deliver hone ke baad hi feedback de sakte ho');
  const note = String(d.comment || '').replace(/[\r\n|]+/g, ' ').trim().slice(0, 500);
  const date = Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd');
  if (d.decision === 'approve') {
    s.getRange(row, C.feedback).setValue('Approved | ' + date + (note ? ' | ' + note : ''));
  } else if (d.decision === 'changes') {
    if (!note) throw new Error('Kya change chahiye, likho');
    s.getRange(row, C.feedback).setValue('Changes | ' + date + ' | ' + note);
    s.getRange(row, C.status).setValue('In progress');
  } else throw new Error('Approve ya Changes chuno');
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
