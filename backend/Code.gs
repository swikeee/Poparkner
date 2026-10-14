/**
 * POPARKNER — Backend Google Apps Script
 * Agent Co · Podomoro Park Bandung
 *
 * Penyimpanan : Google Sheet (1 tab per data) + Google Drive (file dokumen MOU & KTP)
 * Keamanan    : login per user (password di-hash), token sesi 6 jam,
 *               dan SEMUA filter akses per role dilakukan di server.
 *
 * Cara pasang singkat:
 *  1. Buat Google Sheet baru → Extensions → Apps Script → tempel file ini → Save.
 *  2. Jalankan fungsi `setup` sekali (izinkan akses).
 *  3. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone → Deploy.
 *  4. Salin URL Web App → tempel di Poparkner (layar login → "Koneksi Google Sheet").
 */

const COLS = ['agencies', 'agents', 'visits', 'events', 'programs', 'bichecks', 'komisi'];
const SHEET = {
  agencies: 'Agencies', agents: 'Agents', visits: 'Visits', events: 'Events',
  programs: 'Programs', bichecks: 'BIChecking', komisi: 'Komisi'
};
const HEAD = {
  agencies: ['id', 'name', 'office', 'principal', 'phone', 'email', 'joined', 'docs'],
  agents:   ['id', 'name', 'agency', 'phone', 'email', 'joined', 'photo'],
  visits:   ['id', 'date', 'agent', 'consumer', 'phone', 'cluster', 'tipe', 'budget', 'dom', 'km', 'status', 'nextFU', 'notes'],
  events:   ['id', 'date', 'end', 'title', 'cat', 'loc', 'desc'],
  programs: ['id', 'title', 'tag', 'start', 'end', 'desc'],
  bichecks: ['id', 'date', 'agent', 'consumer', 'nik', 'bank', 'status', 'result', 'note', 'file'],
  komisi:   ['id', 'agent', 'consumer', 'cluster', 'tipe', 'unit', 'price', 'pct', 'nominal', 'close', 'stage', 'd1', 'd2', 'd3', 'cair', 'd4', 'note']
};
const JSON_F = { agencies: ['docs'], bichecks: ['file'] };
const NUM_F = { visits: ['km'], komisi: ['price', 'pct', 'nominal', 'stage'] };
const USER_HEAD = ['username', 'password_baru', 'role', 'nama', 'agency', 'agent', 'bank', 'aktif', 'password_hash', 'salt'];
const ROLES = ['agentco', 'principal', 'sales', 'bank'];
const BANK_NAMES = { BCA: 'BCA', BRI: 'BRI', MANDIRI: 'Mandiri', CIMB: 'CIMB Niaga', CIMBS: 'CIMB Niaga Syariah', BTN: 'BTN', MAYBANK: 'Maybank', PERMATA: 'Permata', DANAMON: 'Danamon', BSI: 'BSI' };
const TTL = 21600; // 6 jam
const PROPS = PropertiesService.getScriptProperties();

/* ===================== MENU & SETUP ===================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Poparkner')
    .addItem('1. Setup awal', 'setup')
    .addItem('Proses password user baru', 'hashPendingPasswords')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PROPS.setProperty('SS_ID', ss.getId());
  COLS.forEach(c => prepSheet_(ss, SHEET[c], HEAD[c]));
  const us = prepSheet_(ss, 'Users', USER_HEAD);
  us.getRange(1, 1, 1, USER_HEAD.length).setNote(null);
  us.getRange('B1').setNote('Isi password baru (teks biasa) di sini, lalu menu Poparkner → Proses password user baru. Kolom ini akan dikosongkan otomatis.');
  if (us.getLastRow() < 2) {
    const demo = [
      ['agentco', 'agentco123', 'agentco', 'Podomoro Park', '', '', '', 'TRUE'],
      ['principal.raywhite', 'principal123', 'principal', '', 'A1', '', '', 'TRUE'],
      ['sales.rina', 'sales123', 'sales', '', '', 'S1', '', 'TRUE'],
      ['bank.bca', 'bank123', 'bank', '', '', '', 'BCA', 'TRUE']
    ];
    us.getRange(2, 1, demo.length, 8).setValues(demo);
  }
  hashPendingPasswords();
  // Folder Drive
  if (!PROPS.getProperty('ROOT_FOLDER')) {
    const root = DriveApp.createFolder('Poparkner Files');
    PROPS.setProperty('ROOT_FOLDER', root.getId());
    PROPS.setProperty('MOU_FOLDER', root.createFolder('MOU Agency').getId());
    PROPS.setProperty('BI_FOLDER', root.createFolder('BI Checking').getId());
  }
  const s1 = ss.getSheetByName('Sheet1'); if (s1 && ss.getSheets().length > 1 && s1.getLastRow() === 0) ss.deleteSheet(s1);
  try { SpreadsheetApp.getUi().alert('Setup selesai ✅\n\nLanjut: Deploy → New deployment → Web app (Execute as: Me, Access: Anyone).'); } catch (e) {}
}

function prepSheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), Math.max(sh.getMaxColumns(), head.length)).setNumberFormat('@');
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, head.length).setValues([head]);
  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#002470').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

function hashPendingPasswords() {
  const sh = ss_().getSheetByName('Users'); if (!sh || sh.getLastRow() < 2) return;
  const rng = sh.getRange(2, 1, sh.getLastRow() - 1, USER_HEAD.length);
  const v = rng.getValues(); let n = 0;
  v.forEach(r => {
    if (String(r[1]).trim()) {
      const salt = Utilities.getUuid();
      r[9] = salt; r[8] = hash_(String(r[1]).trim(), salt); r[1] = ''; n++;
    }
  });
  rng.setValues(v);
  try { if (n) SpreadsheetApp.getActiveSpreadsheet().toast(n + ' password diproses'); } catch (e) {}
}

/* ===================== HTTP ===================== */
function doGet() { return out_({ ok: true, data: { app: 'Poparkner', status: 'online' } }); }

function doPost(e) {
  _all = null;
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return out_({ ok: false, error: 'Request tidak valid' }); }
  try {
    const a = body.action;
    if (a === 'ping') return out_({ ok: true, data: { app: 'Poparkner' } });
    if (a === 'login') return out_({ ok: true, data: login_(body) });
    const u = session_(body.token);
    if (!u) return out_({ ok: false, code: 'AUTH', error: 'Sesi berakhir, silakan login ulang' });
    let data;
    switch (a) {
      case 'logout':    CacheService.getScriptCache().remove(body.token); data = true; break;
      case 'bootstrap': data = { user: u, data: readFor_(u) }; break;
      case 'upsert':    data = withLock_(() => upsert_(u, body.col, body.row)); break;
      case 'remove':    data = withLock_(() => remove_(u, body.col, body.id)); break;
      case 'uploadDoc': data = withLock_(() => uploadDoc_(u, body)); break;
      case 'upload':    data = upload_(u, body); break;
      case 'file':      data = file_(u, body.id); break;
      case 'importAll': data = withLock_(() => importAll_(u, body.data)); break;
      default: throw new Error('Aksi tidak dikenal');
    }
    return out_({ ok: true, data: data });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }

/* ===================== AUTH ===================== */
function hash_(pw, salt) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8));
}

function login_(b) {
  const un = String(b.username || '').trim().toLowerCase(), pw = String(b.password || '');
  if (!un || !pw) throw new Error('Isi username dan password');
  const sh = ss_().getSheetByName('Users');
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, USER_HEAD.length).getValues() : [];
  const r = rows.find(x => String(x[0]).trim().toLowerCase() === un);
  if (!r || !r[8] || hash_(pw, r[9]) !== r[8]) throw new Error('Username atau password salah');
  if (String(r[7]).toUpperCase() === 'FALSE') throw new Error('Akun dinonaktifkan');
  const role = String(r[2]).trim();
  if (ROLES.indexOf(role) < 0) throw new Error('Role akun tidak valid');
  if (b.role && b.role !== role) throw new Error('Akun ini bukan untuk profil tersebut');
  const u = { username: un, role: role, agency: String(r[4] || ''), agent: String(r[5] || ''), bank: String(r[6] || ''), sub: String(r[3] || '') };
  if (!u.sub) {
    const A = readAll_();
    if (role === 'agentco') u.sub = 'Podomoro Park';
    if (role === 'principal') u.sub = (A.agencies.find(x => x.id === u.agency) || {}).name || u.agency;
    if (role === 'sales') u.sub = (A.agents.find(x => x.id === u.agent) || {}).name || u.agent;
    if (role === 'bank') u.sub = 'Bank ' + (BANK_NAMES[u.bank] || u.bank);
  }
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put(token, JSON.stringify(u), TTL);
  return { token: token, user: u };
}

function session_(t) {
  if (!t) return null;
  const c = CacheService.getScriptCache(), v = c.get(t);
  if (!v) return null;
  c.put(t, v, TTL); // perpanjang sesi
  return JSON.parse(v);
}

/* ===================== DATA LAYER ===================== */
let _ss = null, _all = null;
function ss_() { if (!_ss) { const id = PROPS.getProperty('SS_ID'); _ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet(); } return _ss; }

function readAll_() {
  if (_all) return _all;
  const tz = Session.getScriptTimeZone(), o = {};
  COLS.forEach(c => {
    const sh = ss_().getSheetByName(SHEET[c]);
    if (!sh || sh.getLastRow() < 2) { o[c] = []; return; }
    const v = sh.getDataRange().getValues(), h = v.shift();
    o[c] = v.filter(r => r[0] !== '').map(r => {
      const x = {};
      h.forEach((k, i) => {
        if (!k) return;
        let val = r[i];
        if (val instanceof Date) val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
        if ((JSON_F[c] || []).indexOf(k) >= 0) { try { val = val ? JSON.parse(val) : (k === 'docs' ? {} : null); } catch (e) { val = null; } }
        else if ((NUM_F[c] || []).indexOf(k) >= 0) val = val === '' ? 0 : Number(val);
        else val = val === null || val === undefined ? '' : String(val);
        x[k] = val;
      });
      return x;
    });
  });
  _all = o;
  return o;
}

function ser_(c, k, v) {
  if ((JSON_F[c] || []).indexOf(k) >= 0) return v ? JSON.stringify(v) : '';
  if (v === null || v === undefined) return '';
  return String(v);
}

function writeRow_(c, obj) {
  const sh = ss_().getSheetByName(SHEET[c]);
  let h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const extra = Object.keys(obj).filter(k => h.indexOf(k) < 0);
  if (extra.length) { sh.getRange(1, h.length + 1, 1, extra.length).setValues([extra]).setFontWeight('bold'); h = h.concat(extra); }
  const row = h.map(k => ser_(c, k, obj[k]));
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const i = ids.indexOf(String(obj.id));
  const r = i >= 0 ? i + 2 : sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, h.length).setNumberFormat('@').setValues([row]);
  if (_all) { const arr = _all[c], j = arr.findIndex(x => x.id === obj.id); if (j >= 0) arr[j] = obj; else arr.push(obj); }
}

function writeAll_(c, rows) {
  const sh = ss_().getSheetByName(SHEET[c]);
  const h = HEAD[c].slice();
  rows.forEach(o => Object.keys(o).forEach(k => { if (h.indexOf(k) < 0) h.push(k); }));
  sh.clearContents();
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#002470').setFontColor('#ffffff');
  if (rows.length) sh.getRange(2, 1, rows.length, h.length).setNumberFormat('@').setValues(rows.map(o => h.map(k => ser_(c, k, o[k]))));
}

/* ===================== FILTER BACA PER ROLE ===================== */
function stripAgency_(g) { return { id: g.id, name: g.name, office: g.office, principal: g.principal, phone: g.phone, email: g.email, joined: g.joined, docs: {} }; }

function readFor_(u) {
  const A = readAll_();
  if (u.role === 'agentco') return A;
  const o = { events: A.events, programs: A.programs };
  if (u.role === 'bank') {
    const bi = A.bichecks.filter(b => b.bank === u.bank);
    const cons = {}; bi.forEach(b => cons[b.consumer] = 1);
    o.bichecks = bi;
    o.visits = A.visits.filter(v => cons[v.consumer]);
    o.komisi = A.komisi.filter(k => cons[k.consumer]);
    const ids = {}; bi.concat(o.visits, o.komisi).forEach(x => ids[x.agent] = 1);
    o.agents = A.agents.filter(a => ids[a.id]).map(a => ({ id: a.id, name: a.name, agency: a.agency, photo: a.photo, phone: '', email: '', joined: '' }));
  } else {
    const ag = u.role === 'principal' ? A.agents.filter(a => a.agency === u.agency) : A.agents.filter(a => a.id === u.agent);
    const ids = {}; ag.forEach(a => ids[a.id] = 1);
    o.agents = ag;
    o.visits = A.visits.filter(v => ids[v.agent]);
    o.komisi = A.komisi.filter(k => ids[k.agent]);
    o.bichecks = A.bichecks.filter(b => ids[b.agent]);
  }
  const gid = {}; o.agents.forEach(a => gid[a.agency] = 1); if (u.agency) gid[u.agency] = 1;
  o.agencies = A.agencies.filter(g => gid[g.id]).map(g => (u.role === 'principal' && g.id === u.agency) ? g : stripAgency_(g));
  return o;
}

/* ===================== ATURAN TULIS PER ROLE ===================== */
function deny_(m) { throw new Error(m || 'Akses ditolak untuk role Anda'); }
function pick_(o, keys) { const x = {}; keys.forEach(k => { if (k in o) x[k] = o[k]; }); return x; }

function authorize_(u, c, row, old, A) {
  const r = u.role;
  if (r === 'agentco') return row;
  const mine = r === 'principal' ? A.agents.filter(a => a.agency === u.agency).map(a => a.id) : r === 'sales' ? [u.agent] : [];
  switch (c) {
    case 'agents':
      if (r === 'principal') { if (old && old.agency !== u.agency) deny_(); return Object.assign({}, old || {}, row, { agency: u.agency }); }
      if (r === 'sales') { if (!old || old.id !== u.agent) deny_(); return Object.assign({}, old, pick_(row, ['photo', 'phone', 'email'])); }
      break;
    case 'visits':
      if (r === 'principal' || r === 'sales') {
        if (old && mine.indexOf(old.agent) < 0) deny_();
        if (mine.indexOf(row.agent) < 0) deny_('Agent tidak sesuai');
        return row;
      }
      break;
    case 'bichecks':
      if (r === 'bank') { if (!old || old.bank !== u.bank) deny_(); return Object.assign({}, old, pick_(row, ['status', 'result', 'note'])); }
      if ((r === 'principal' || r === 'sales') && !old) {
        if (mine.indexOf(row.agent) < 0) deny_('Agent tidak sesuai');
        return Object.assign({}, row, { status: 'Diajukan', result: '', note: '' });
      }
      break;
    case 'komisi':
      if ((r === 'principal' || r === 'sales') && !old) {
        if (mine.indexOf(row.agent) < 0) deny_('Agent tidak sesuai');
        return Object.assign({}, row, { stage: 1, d2: '', d3: '', d4: '', cair: '', note: '' });
      }
      break;
  }
  deny_();
}

function upsert_(u, c, row) {
  if (COLS.indexOf(c) < 0 || !row || !row.id) throw new Error('Data tidak valid');
  const A = readAll_();
  const old = A[c].find(x => x.id === row.id) || null;
  const fin = authorize_(u, c, row, old, A);
  writeRow_(c, fin);
  return fin;
}

function remove_(u, c, id) {
  if (u.role !== 'agentco') deny_();
  if (COLS.indexOf(c) < 0) throw new Error('Data tidak valid');
  const sh = ss_().getSheetByName(SHEET[c]);
  if (sh.getLastRow() < 2) return true;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  if (i >= 0) sh.deleteRow(i + 2);
  return true;
}

/* ===================== FILE (Google Drive) ===================== */
function saveFile_(folderId, sub, b) {
  if (!b.data || String(b.data).indexOf('base64,') < 0) throw new Error('File tidak valid');
  const raw = String(b.data).split('base64,')[1];
  if (raw.length > 14 * 1024 * 1024) throw new Error('File terlalu besar (maks 10 MB)');
  let folder = DriveApp.getFolderById(folderId);
  if (sub) { const it = folder.getFoldersByName(sub); folder = it.hasNext() ? it.next() : folder.createFolder(sub); }
  const blob = Utilities.newBlob(Utilities.base64Decode(raw), b.type || 'application/octet-stream', b.name || 'file');
  const f = folder.createFile(blob);
  return { id: f.getId(), name: f.getName(), type: b.type || f.getMimeType(), at: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') };
}

function uploadDoc_(u, b) {
  const A = readAll_();
  const g = A.agencies.find(x => x.id === b.agencyId);
  if (!g) throw new Error('Agency tidak ditemukan');
  if (u.role !== 'agentco' && !(u.role === 'principal' && u.agency === g.id)) deny_();
  const docs = g.docs || {};
  if (docs[b.key] && docs[b.key].status === 'ok' && u.role !== 'agentco') deny_('Dokumen sudah disetujui');
  const meta = saveFile_(PROPS.getProperty('MOU_FOLDER'), g.name + ' (' + g.id + ')', b);
  docs[b.key] = { file: meta, status: 'pending', note: '' };
  g.docs = docs;
  writeRow_('agencies', g);
  return g;
}

function upload_(u, b) {
  if (b.kind !== 'bicheck') throw new Error('Jenis upload tidak dikenal');
  if (['agentco', 'principal', 'sales'].indexOf(u.role) < 0) deny_();
  return saveFile_(PROPS.getProperty('BI_FOLDER'), '', b);
}

function file_(u, id) {
  if (!id) throw new Error('File tidak valid');
  // Hanya boleh dibuka jika file direferensikan oleh data yang memang boleh dibaca user ini
  if (JSON.stringify(readFor_(u)).indexOf('"' + id + '"') < 0) deny_('Anda tidak punya akses ke file ini');
  const f = DriveApp.getFileById(id), blob = f.getBlob();
  return { data: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()), name: f.getName(), type: blob.getContentType() };
}

/* ===================== IMPORT DATA CONTOH ===================== */
function importAll_(u, data) {
  if (u.role !== 'agentco') deny_();
  if (!data) throw new Error('Data kosong');
  COLS.forEach(c => writeAll_(c, Array.isArray(data[c]) ? data[c] : []));
  _all = null;
  return true;
}
