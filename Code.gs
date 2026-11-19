/**
 * OHA ARTISAN BREAD — ระบบบริหารร้าน (Backend)
 * Google Apps Script + Google Sheets
 *
 * วิธีติดตั้ง (ทำครั้งเดียว):
 *  1) สร้าง Google Sheet ใหม่ → เมนู ส่วนขยาย (Extensions) → Apps Script
 *  2) ลบโค้ดเดิม วางไฟล์นี้ทั้งหมด → กดบันทึก
 *  3) เลือกฟังก์ชัน setup → กด Run → อนุญาตสิทธิ์ (Sheets + Drive)
 *  4) Deploy → New deployment → ประเภท Web app
 *       Execute as: Me   |   Who has access: Anyone
 *  5) คัดลอก Web app URL (ลงท้าย /exec) ไปใส่ใน index.html (ค่า API_URL)
 *
 * ข้อมูลแต่ละประเภทอยู่คนละชีต: 1 แถว = 1 รายการ (คอลัมน์ json)
 * อย่าแก้/ลบแถวในชีตด้วยมือขณะมีคนใช้งาน — ให้แก้ผ่านแอป
 */

const TZ = 'Asia/Bangkok';
const COLLS = ['users','items','moves','sales','attendance','ot','leaves','reports','goods','pmoves','wcust','worders','members','memberTx','expenses'];
const API_VERSION = 3;
const OWNER_WRITE = ['users','sales'];            // พนักงานแก้ไม่ได้
const STAFF_HIDDEN = ['sales'];                   // พนักงานมองไม่เห็น
const STAFF_DELETE = ['leaves','ot','pmoves','expenses'];    // พนักงานลบได้ (ยกเลิกคำขอ / ลบบันทึกของตัวเอง)
const OWNER_KV = ['settings'];
const SESSION_DAYS = 60;
const ADMIN_UID = 'u0';                          // อ๊อฟ — ดู/ตั้ง PIN ของทุกคนได้คนเดียว
const DEFAULT_USERS = [
  ['u4','โอ้','owner','เจ้าของร้าน','5678'],
  ['u5','เอ้','owner','เจ้าของร้าน','9012'],
  ['u6','พี่โม','owner','เจ้าของร้าน','2468'],
  ['u0','อ๊อฟ','owner','เจ้าของร้าน','1234'],
  ['u1','เปรี้ยว','staff','พนักงาน','1111'],
  ['u2','บีม','staff','พนักงาน','2222'],
  ['u3','จ๋า','staff','พนักงาน','3333']
];

/* ===================== ติดตั้ง ===================== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  COLLS.concat(['kv']).forEach(sh_);
  sysSheet_('pins', PIN_HEAD);
  sysSheet_('sessions', ['token', 'uid', 'expires']);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SALT')) props.setProperty('SALT', Utilities.getUuid());
  // เพิ่มผู้ใช้เริ่มต้นที่ยังไม่มี (รันซ้ำได้ ไม่ทับข้อมูลเดิม)
  const have = indexRows_('users');
  const missing = DEFAULT_USERS.filter(([id]) => !have[id]);
  if (missing.length) {
    const now = Date.now();
    const rows = missing.map(([id, name, role, pos, pin]) => {
      setPin_(id, pin);
      return [id, now, false, JSON.stringify({ id, name, role, pos })];
    });
    sh_('users').getRange(sh_('users').getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }
  backfillPins_();
  photoFolder_();
  ['Sheet1', 'ชีต1', 'แผ่นงาน1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1 && s.getLastRow() === 0) ss.deleteSheet(s);
  });
  Logger.log('ตั้งค่าเสร็จแล้ว — ต่อไป Deploy เป็น Web app');
}

/* ===================== Web app ===================== */
function doGet() { return out_({ ok: true, app: 'OHA ARTISAN BREAD', time: new Date().toISOString() }); }

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    const fn = ACTIONS[b.action];
    if (!fn) throw new Error('unknown action');
    return out_(Object.assign({ ok: true }, fn(b)));
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err), auth: !!(err && err.auth) });
  }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

const ACTIONS = {
  /* รายชื่อสำหรับหน้าเข้าสู่ระบบ (ไม่มี PIN) */
  users() {
    return { users: readRows_('users').filter(r => !r.del).map(r => { const u = JSON.parse(r.json); return { id: u.id, name: u.name, role: u.role, pos: u.pos }; }) };
  },

  login(b) {
    const uid = String(b.uid || ''), pin = String(b.pin || '');
    const cache = CacheService.getScriptCache(), fk = 'f_' + uid;
    const fails = Number(cache.get(fk) || 0);
    if (fails >= 6) throw new Error('ใส่ PIN ผิดหลายครั้ง กรุณารอ 10 นาที');
    const user = getUser_(uid);
    if (!user || getPinHash_(uid) !== hash_(uid, pin)) {
      cache.put(fk, String(fails + 1), 600);
      throw new Error('PIN ไม่ถูกต้อง');
    }
    cache.remove(fk);
    const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
    sysSheet_('sessions', ['token', 'uid', 'expires']).appendRow([token, uid, Date.now() + SESSION_DAYS * 864e5]);
    cache.put('t_' + token, uid, 21600);
    return { token, user: { id: user.id, name: user.name, role: user.role, pos: user.pos, admin: user.id === ADMIN_UID } };
  },

  logout(b) {
    if (!b.token) return {};
    CacheService.getScriptCache().remove('t_' + b.token);
    const sh = sysSheet_('sessions', ['token', 'uid', 'expires']), v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) if (v[i][0] === b.token) sh.deleteRow(i + 1);
    return {};
  },

  load(b) {
    const user = auth_(b.token), ts = Date.now() - 2000;
    const data = {};
    COLLS.forEach(c => {
      data[c] = (user.role !== 'owner' && STAFF_HIDDEN.indexOf(c) >= 0) ? [] :
        readRows_(c).filter(r => !r.del).map(r => JSON.parse(r.json)).filter(x => canSee_(c, x, user));
    });
    const kv = {};
    readRows_('kv').forEach(r => { if (!r.del) kv[r.id] = JSON.parse(r.json); });
    return { data, kv, ts, v: API_VERSION, me: Object.assign({}, user, { admin: user.id === ADMIN_UID }) };
  },

  changes(b) {
    const user = auth_(b.token), ts = Date.now() - 2000, since = Number(b.since || 0);
    const data = {};
    COLLS.forEach(c => {
      if (user.role !== 'owner' && STAFF_HIDDEN.indexOf(c) >= 0) return;
      const rs = readRows_(c).filter(r => r.ts > since);
      const out = rs.map(r => r.del ? { id: r.id, _deleted: true } : JSON.parse(r.json)).filter(x => x._deleted || canSee_(c, x, user));
      if (out.length) data[c] = out;
    });
    const kv = {};
    readRows_('kv').filter(r => r.ts > since).forEach(r => { kv[r.id] = r.del ? null : JSON.parse(r.json); });
    return { data, kv, ts };
  },

  sync(b) {
    const user = auth_(b.token), owner = user.role === 'owner';
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      const fixed = {}, ups = b.ups || {}, dels = b.dels || {};
      let skipped = 0;
      Object.keys(ups).forEach(c => {
        if (COLLS.indexOf(c) < 0) return;
        if (!owner && OWNER_WRITE.indexOf(c) >= 0) { skipped += ups[c].length; return; }
        const existing = indexRows_(c);
        const list = [];
        ups[c].forEach(rec => {
          if (!rec || !rec.id) return;
          rec.id = String(rec.id);
          const old = existing[rec.id] && !existing[rec.id].del ? JSON.parse(existing[rec.id].json) : null;
          if (c === 'expenses') { const r = fixExpense_(rec, old, user); if (!r) { skipped++; return; } (fixed[c] = fixed[c] || []).push(r); list.push(r); return; }
          if (c === 'attendance') { const r = fixAttendance_(rec, old, user, existing); if (!r) { skipped++; return; } (fixed[c] = fixed[c] || []).push(r); list.push(r); return; }
          if (!owner && (c === 'leaves' || c === 'ot') && rec.status !== 'pending' && (!old || old.status !== rec.status)) { skipped++; return; }
          list.push(rec);
        });
        writeRecords_(c, list, [], existing);
      });
      Object.keys(dels).forEach(c => {
        if (COLLS.indexOf(c) < 0) return;
        if (!owner && STAFF_DELETE.indexOf(c) < 0) { skipped += dels[c].length; return; }
        writeRecords_(c, [], dels[c].map(String));
      });
      const kvUps = [], kvDels = [];
      Object.keys(b.kv || {}).forEach(k => { if (!owner && OWNER_KV.indexOf(k) >= 0) { skipped++; return; } kvUps.push({ id: k, _v: b.kv[k] }); });
      (b.kvDel || []).forEach(k => { if (!owner && OWNER_KV.indexOf(k) >= 0) { skipped++; return; } kvDels.push(String(k)); });
      if (kvUps.length || kvDels.length) writeRecords_('kv', kvUps, kvDels, null, true);
      return { fixed, skipped, ts: Date.now() };
    } finally { lock.releaseLock(); }
  },

  setPin(b) {
    const user = auth_(b.token), uid = String(b.uid || ''), pin = String(b.pin || '');
    if (user.id !== ADMIN_UID) throw new Error('เฉพาะอ๊อฟเท่านั้นที่ตั้ง PIN ให้คนอื่นได้');
    if (!/^\d{4,6}$/.test(pin)) throw new Error('PIN ต้องเป็นตัวเลข 4–6 หลัก');
    setPin_(uid, pin);
    return {};
  },

  changePin(b) {
    const user = auth_(b.token), pin = String(b.pin || '');
    if (getPinHash_(user.id) !== hash_(user.id, String(b.old || ''))) throw new Error('PIN เดิมไม่ถูกต้อง');
    if (!/^\d{4,6}$/.test(pin)) throw new Error('PIN ต้องเป็นตัวเลข 4–6 หลัก');
    setPin_(user.id, pin);
    return {};
  },

  /* รายการ PIN ทุกคน — ดูได้เฉพาะ ADMIN */
  listPins(b) {
    const user = auth_(b.token);
    if (user.id !== ADMIN_UID) throw new Error('ไม่มีสิทธิ์ดู PIN');
    backfillPins_();
    const v = sysSheet_('pins', PIN_HEAD).getDataRange().getValues(), m = {};
    for (let i = 1; i < v.length; i++) m[String(v[i][0])] = { pin: String(v[i][2] || ''), at: Number(v[i][3]) || 0 };
    return { pins: m };
  },

  photo(b) {
    const user = auth_(b.token);
    const f = DriveApp.getFileById(String(b.id || ''));
    if (user.role !== 'owner' && f.getName().indexOf(user.id + '_') !== 0) throw new Error('ไม่มีสิทธิ์ดูรูปนี้');
    const blob = f.getBlob();
    return { dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
  }
};

/* ===================== รายจ่าย ===================== */
/** พนักงานเห็นเฉพาะรายจ่ายที่ตัวเองจ่าย/บันทึก */
function canSee_(c, x, user) {
  if (c !== 'expenses' || user.role === 'owner') return true;
  return x.payer === user.id || x.by === user.id;
}
function fixExpense_(rec, old, user) {
  if (user.role !== 'owner') {
    if (old && (old.by !== user.id || old.reimbursed)) return null;   // แก้ของคนอื่น/รายการที่เบิกแล้วไม่ได้
    rec.by = old ? old.by : user.id;
    rec.payer = user.id;
    rec.reimbursed = false; delete rec.reimbursedAt; delete rec.reimbursedBy;
  }
  if (typeof rec.receipt === 'string' && rec.receipt.indexOf('data:') === 0) {
    rec.receipt = 'drive:' + savePhoto_(rec.receipt, (rec.by || user.id) + '_exp_' + rec.date + '_' + rec.id);
  }
  return rec;
}

/* ===================== ลงเวลา: ใช้เวลาเซิร์ฟเวอร์ ===================== */
function fixAttendance_(rec, old, user, existing) {
  const nowIso = new Date().toISOString();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (user.role !== 'owner') {
    if (rec.uid !== user.id) return null;
    delete rec.editedBy;
    if (!old && Object.keys(existing).some(k => { const x = existing[k]; if (x.del) return false; const a = JSON.parse(x.json); return a.uid === user.id && a.date === today; })) return null; // ลงเวลาเข้าซ้ำในวันเดียวกันไม่ได้
    if (!old) { rec.in = nowIso; rec.date = today; if (rec.inLoc) rec.inLoc.at = nowIso; }
    else {
      rec.in = old.in; rec.date = old.date; rec.inLoc = old.inLoc;
      if (old.out) { rec.out = old.out; rec.outLoc = old.outLoc; }
      else if (rec.out) { rec.out = nowIso; if (rec.outLoc) rec.outLoc.at = nowIso; }
    }
    rec.server = true;
  }
  ['inLoc', 'outLoc'].forEach(k => {
    const l = rec[k];
    if (l && typeof l.photo === 'string' && l.photo.indexOf('data:') === 0) {
      l.photo = 'drive:' + savePhoto_(l.photo, rec.uid + '_' + rec.date + '_' + k);
    }
  });
  return rec;
}
function savePhoto_(dataUrl, name) {
  const m = dataUrl.match(/^data:(.+?);base64,(.*)$/);
  if (!m) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name + '.jpg');
  return photoFolder_().createFile(blob).getId();
}
function photoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('OHA ARTISAN BREAD — รูปเช็คอิน');
  props.setProperty('PHOTO_FOLDER', f.getId());
  return f;
}

/* ===================== ผู้ใช้ / PIN / session ===================== */
function authErr_(m) { const e = new Error(m || 'กรุณาเข้าสู่ระบบใหม่'); e.auth = true; return e; }
function auth_(token) {
  if (!token) throw authErr_();
  const cache = CacheService.getScriptCache();
  let uid = cache.get('t_' + token);
  if (!uid) {
    const v = sysSheet_('sessions', ['token', 'uid', 'expires']).getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (v[i][0] === token && Number(v[i][2]) > Date.now()) { uid = String(v[i][1]); break; }
    if (!uid) throw authErr_();
    cache.put('t_' + token, uid, 21600);
  }
  const u = getUser_(uid);
  if (!u) throw authErr_('ไม่พบผู้ใช้นี้แล้ว');
  return u;
}
function getUser_(uid) {
  const r = readRows_('users').find(x => x.id === uid && !x.del);
  return r ? JSON.parse(r.json) : null;
}
function hash_(uid, pin) {
  const salt = PropertiesService.getScriptProperties().getProperty('SALT') || '';
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + uid + '|' + pin, Utilities.Charset.UTF_8);
  return d.map(x => ('0' + ((x + 256) % 256).toString(16)).slice(-2)).join('');
}
const PIN_HEAD = ['uid', 'hash', 'pin', 'updatedAt'];
function getPinHash_(uid) {
  const v = sysSheet_('pins', PIN_HEAD).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === uid) return String(v[i][1]);
  return null;
}
/** เก็บ PIN ทั้งแบบเข้ารหัส (ใช้ตรวจตอนเข้าระบบ) และตัวจริง (ให้อ๊อฟเปิดดูเผื่อมีคนลืม) */
function setPin_(uid, pin) {
  const sh = sysSheet_('pins', PIN_HEAD), v = sh.getDataRange().getValues(), h = hash_(uid, pin), now = Date.now();
  sh.getRange('C:C').setNumberFormat('@');
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === uid) { sh.getRange(i + 1, 2, 1, 3).setValues([[h, String(pin), now]]); return; }
  sh.appendRow([uid, h, String(pin), now]);
}
/** แถวเก่าที่ยังไม่มี PIN ตัวจริง: ถ้ายังเป็น PIN เริ่มต้นอยู่ ให้เติมให้ */
function backfillPins_() {
  const sh = sysSheet_('pins', PIN_HEAD), v = sh.getDataRange().getValues();
  if (v[0].length < 4 || v[0][2] !== 'pin') sh.getRange(1, 1, 1, 4).setValues([PIN_HEAD]).setFontWeight('bold');
  sh.getRange('C:C').setNumberFormat('@');
  const def = {}; DEFAULT_USERS.forEach(([id, , , , pin]) => def[id] = pin);
  for (let i = 1; i < v.length; i++) {
    const uid = String(v[i][0]);
    if (!v[i][2] && def[uid] && String(v[i][1]) === hash_(uid, def[uid])) sh.getRange(i + 1, 3, 1, 2).setValues([[def[uid], Date.now()]]);
  }
}

/* ===================== Sheets storage ===================== */
function sh_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('D:D').setNumberFormat('@');
    sh.getRange(1, 1, 1, 4).setValues([['id', 'updatedAt', 'deleted', 'json']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function sysSheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function readRows_(name) {
  const sh = sh_(name), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 4).getValues().map((v, i) => ({
    id: String(v[0]), ts: Number(v[1]) || 0, del: v[2] === true || v[2] === 'TRUE', json: String(v[3] || '{}'), row: i + 2
  }));
}
function indexRows_(name) { const m = {}; readRows_(name).forEach(r => m[r.id] = r); return m; }

/** recs: รายการที่จะเขียน, delIds: id ที่จะลบ (เก็บเป็น tombstone) */
function writeRecords_(name, recs, delIds, existing, isKv) {
  if (!recs.length && !delIds.length) return;
  const sh = sh_(name), idx = existing || indexRows_(name), now = Date.now();
  const updates = [], appends = [];
  recs.forEach(rec => {
    const id = String(rec.id);
    const json = JSON.stringify(isKv ? rec._v : rec);
    if (json.length > 49000) throw new Error('ข้อมูลรายการ ' + name + '/' + id + ' ใหญ่เกินไป');
    const row = [id, now, false, json];
    if (idx[id]) updates.push([idx[id].row, row]); else { appends.push(row); idx[id] = { row: -1 }; }
  });
  delIds.forEach(id => { if (idx[id] && idx[id].row > 0) updates.push([idx[id].row, [id, now, true, '{}']]); });
  if (updates.length > 25) {
    const n = sh.getLastRow() - 1, all = sh.getRange(2, 1, n, 4).getValues();
    updates.forEach(([row, v]) => all[row - 2] = v);
    sh.getRange(2, 1, n, 4).setValues(all);
  } else updates.forEach(([row, v]) => sh.getRange(row, 1, 1, 4).setValues([v]));
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, 4).setValues(appends);
}
