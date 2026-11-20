/**
 * スラロームゲームの記録サーバー（Google Apps Script）
 * このスプレッドシートに、ユーザー・保存データ・タイムアタック記録を残します。
 *
 * 使い方：スプレッドシート →［拡張機能］→［Apps Script］にこのコードを貼り付け、
 *         ［デプロイ］→［新しいデプロイ］→ 種類「ウェブアプリ」
 *         実行するユーザー「自分」、アクセスできるユーザー「全員」でデプロイ。
 */
const SHEET_USERS = 'users';
const SHEET_DATA = 'data';
const SHEET_TA = 'ta';

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(header); s.setFrozenRows(1); }
  return s;
}
function users_() { return sheet_(SHEET_USERS, ['name', 'pinHash', 'salt', 'token', 'created', 'lastLogin', 'fails']); }
function data_() { return sheet_(SHEET_DATA, ['name', 'key', 'json', 'updated']); }
function ta_() { return sheet_(SHEET_TA, ['name', 'course', 't', 'r1', 'r2', 'at']); }

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hash_(pin, salt) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pin, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ((x + 256) % 256).toString(16).padStart(2, '0'); }).join('');
}
function cleanName_(n) { return String(n || '').replace(/[\u0000-\u001f<>"'&]/g, '').trim().slice(0, 14); }
function findUser_(name) {
  const s = users_(), v = s.getDataRange().getValues(), key = name.toLowerCase();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).toLowerCase() === key) return { row: i + 1, v: v[i] };
  return null;
}
function userByToken_(token) {
  if (!token) return null;
  const v = users_().getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][3] === token) return String(v[i][0]);
  return null;
}

function doGet(e) {
  const a = (e && e.parameter && e.parameter.action) || '';
  if (a === 'board') return out_(board_());
  if (a === 'players') return out_(players_());
  return out_({ ok: true, app: 'slalom' });
}

function doPost(e) {
  let q = {};
  try { q = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch (err) { return out_({ ok: false, error: 'json' }); }
  // live battle traffic is answered without the global lock so positions stay quick
  if (q.action === 'room_pos') return out_(room_pos_(q));
  if (q.action === 'room_state') return out_(room_state_(q));
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    switch (q.action) {
      case 'register': return out_(register_(q));
      case 'login': return out_(login_(q));
      case 'load': return out_(load_(q));
      case 'save': return out_(save_(q));
      case 'ta': return out_(ta_submit_(q));
      case 'board': return out_(board_());
      case 'room_create': return out_(room_create_(q));
      case 'room_join': return out_(room_join_(q));
      case 'room_start': return out_(room_start_(q));
      case 'room_leave': return out_(room_leave_(q));
    }
    return out_({ ok: false, error: 'unknown' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function register_(q) {
  const name = cleanName_(q.name), pin = String(q.pin || '');
  if (name.length < 2) return { ok: false, error: 'name' };
  if (!/^\d{4,8}$/.test(pin)) return { ok: false, error: 'pin' };
  if (findUser_(name)) return { ok: false, error: 'taken' };
  const salt = Utilities.getUuid(), token = Utilities.getUuid();
  users_().appendRow([name, hash_(pin, salt), salt, token, new Date(), new Date(), 0]);
  return { ok: true, name: name, token: token };
}

function login_(q) {
  const name = cleanName_(q.name), pin = String(q.pin || '');
  const u = findUser_(name);
  if (!u) { Utilities.sleep(600); return { ok: false, error: 'nouser' }; }
  const fails = Number(u.v[6] || 0);
  if (fails >= 10) return { ok: false, error: 'locked' };
  const s = users_();
  if (hash_(pin, u.v[2]) !== u.v[1]) {
    s.getRange(u.row, 7).setValue(fails + 1);
    Utilities.sleep(600);
    return { ok: false, error: 'pin' };
  }
  s.getRange(u.row, 6, 1, 2).setValues([[new Date(), 0]]);
  return { ok: true, name: String(u.v[0]), token: u.v[3] };
}

function load_(q) {
  const name = userByToken_(q.token);
  if (!name) return { ok: false, error: 'auth' };
  const v = data_().getDataRange().getValues(), out = {};
  for (let i = 1; i < v.length; i++) if (v[i][0] === name) out[v[i][1]] = v[i][2];
  return { ok: true, name: name, data: out };
}

function save_(q) {
  const name = userByToken_(q.token);
  if (!name) return { ok: false, error: 'auth' };
  const items = q.items || {};
  const s = data_(), v = s.getDataRange().getValues();
  Object.keys(items).slice(0, 20).forEach(function (key) {
    if (!/^slalomwc-[\w-]{1,60}$/.test(key)) return;
    const json = String(items[key] || '').slice(0, 45000);
    let row = -1;
    for (let i = 1; i < v.length; i++) if (v[i][0] === name && v[i][1] === key) { row = i + 1; break; }
    if (row > 0) s.getRange(row, 3, 1, 2).setValues([[json, new Date()]]);
    else { s.appendRow([name, key, json, new Date()]); v.push([name, key, json, new Date()]); }
  });
  return { ok: true };
}

function ta_submit_(q) {
  const name = userByToken_(q.token);
  if (!name) return { ok: false, error: 'auth' };
  const course = String(q.course || ''), t = Number(q.t), r1 = Number(q.r1), r2 = Number(q.r2);
  if (!/^(men|women)_\d{4}-\d{2}-\d{2}$/.test(course)) return { ok: false, error: 'course' };
  if (!(t > 15 && t < 400 && r1 > 5 && r2 > 5)) return { ok: false, error: 'time' };
  const s = ta_(), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === name && v[i][1] === course) {
      if (t < Number(v[i][2])) { s.getRange(i + 1, 3, 1, 4).setValues([[t, r1, r2, new Date()]]); return { ok: true, improved: true }; }
      return { ok: true, improved: false };
    }
  }
  s.appendRow([name, course, t, r1, r2, new Date()]);
  return { ok: true, improved: true };
}

function board_() {
  const v = ta_().getDataRange().getValues(), rows = [];
  for (let i = 1; i < v.length; i++) rows.push({ name: String(v[i][0]), course: v[i][1], t: Number(v[i][2]), r1: Number(v[i][3]), r2: Number(v[i][4]) });
  return { ok: true, rows: rows };
}

// Player ranking: each player's career summary for every season they have played (no personal data beyond the name)
function players_() {
  const v = data_().getDataRange().getValues(), rows = [], who = {};
  for (let i = 1; i < v.length; i++) {
    const key = String(v[i][1]);
    if (key === 'slalomwc-prefs') {
      try { const pf = JSON.parse(v[i][2] || '{}'); who[String(v[i][0])] = { nat: String(pf.nat || ''), gear: pf.gear || {} }; } catch (e) {}
      continue;
    }
    if (!/^slalomwc-2627-(men|women)-[a-z]+$/.test(key)) continue;
    try {
      const d = JSON.parse(v[i][2] || '{}');
      const done = d.done || {}, places = (d.places && d.places.you) || [];
      let races = 0, best = 0, vmax = 0, top10 = 0, rankSum = 0, ranked = 0;
      Object.keys(done).forEach(function (k) {
        const r = done[k]; if (!r) return;
        if (r.status !== 'dns') races++;
        if (r.rank){ if (!best || r.rank < best) best = r.rank; if (r.rank <= 10) top10++; rankSum += r.rank; ranked++; }
        if (r.v && r.v > vmax) vmax = r.v;
      });
      rows.push({ name: String(v[i][0]), key: key.replace('slalomwc-2627-', ''), pts: Number((d.pts && d.pts.you) || 0), races: races,
                  wins: Number(places[0] || 0), podiums: Number(places[0] || 0) + Number(places[1] || 0) + Number(places[2] || 0), best: best,
                  vmax: vmax, top10: top10, avg: ranked ? Math.round(rankSum / ranked * 10) / 10 : 0 });
    } catch (e) {}
  }
  return { ok: true, rows: rows, who: who };
}

// ---------------------------------------------------------------------
// Online battle rooms (kept in the script cache for up to 6 hours; nothing is written to the sheet)
// ---------------------------------------------------------------------
const RC_TTL = 21600;
function rc_() { return CacheService.getScriptCache(); }
function roomGet_(code) { const s = rc_().get('room_' + code); return s ? JSON.parse(s) : null; }
function roomPut_(r) { rc_().put('room_' + r.code, JSON.stringify(r), RC_TTL); }
function posKey_(code, name) { return 'pos_' + code + '_' + name; }
function roomView_(r) {
  const keys = r.members.map(function (m) { return posKey_(r.code, m); });
  const got = keys.length ? rc_().getAll(keys) : {}, pos = {};
  r.members.forEach(function (m) { const v = got[posKey_(r.code, m)]; if (v) pos[m] = JSON.parse(v); });
  return { ok: true, room: r, pos: pos, now: Date.now() };
}
function room_create_(q) {
  const name = userByToken_(q.token);
  if (!name) return { ok: false, error: 'auth' };
  const course = String(q.course || '');
  if (!/^(men|women)_\d{4}-\d{2}-\d{2}$/.test(course)) return { ok: false, error: 'course' };
  let code = '';
  for (let i = 0; i < 30; i++) { const c = String(1000 + Math.floor(Math.random() * 9000)); if (!roomGet_(c)) { code = c; break; } }
  if (!code) return { ok: false, error: 'busy' };
  const r = { code: code, host: name, course: course, members: [name], state: 'lobby', raceId: 0, startAt: 0, created: Date.now() };
  roomPut_(r);
  return roomView_(r);
}
function room_join_(q) {
  const name = userByToken_(q.token);
  if (!name) return { ok: false, error: 'auth' };
  const r = roomGet_(String(q.code || ''));
  if (!r) return { ok: false, error: 'noroom' };
  if (r.members.indexOf(name) < 0) {
    if (r.members.length >= 8) return { ok: false, error: 'full' };
    r.members.push(name); roomPut_(r);
  }
  return roomView_(r);
}
function room_state_(q) {
  const r = roomGet_(String(q.code || ''));
  if (!r) return { ok: false, error: 'noroom' };
  return roomView_(r);
}
function room_start_(q) {
  const name = userByToken_(q.token);
  const r = roomGet_(String(q.code || ''));
  if (!r) return { ok: false, error: 'noroom' };
  if (r.host !== name) return { ok: false, error: 'host' };
  if (q.course && /^(men|women)_\d{4}-\d{2}-\d{2}$/.test(String(q.course))) r.course = String(q.course);
  r.raceId += 1; r.state = 'race'; r.startAt = Date.now() + 9000;
  rc_().removeAll(r.members.map(function (m) { return posKey_(r.code, m); }));
  roomPut_(r);
  return roomView_(r);
}
function room_pos_(q) {
  const name = userByToken_(q.token);
  if (!name) return { ok: false, error: 'auth' };
  const r = roomGet_(String(q.code || ''));
  if (!r) return { ok: false, error: 'noroom' };
  if (r.members.indexOf(name) < 0) return { ok: false, error: 'member' };
  const num = function (x) { const n = Number(x); return isFinite(n) ? Math.round(n * 100) / 100 : 0; };
  const p = { raceId: Number(q.raceId) || 0, t: num(q.t), x: num(q.x), y: num(q.y), v: num(q.v), th: num(q.th), ng: Number(q.ng) || 0,
              fin: q.fin == null ? null : num(q.fin), out: q.out ? String(q.out).slice(0, 3) : '', ts: Date.now() };
  rc_().put(posKey_(r.code, name), JSON.stringify(p), RC_TTL);
  return roomView_(r);
}
function room_leave_(q) {
  const name = userByToken_(q.token);
  const r = roomGet_(String(q.code || ''));
  if (!r || !name) return { ok: true };
  r.members = r.members.filter(function (m) { return m !== name; });
  rc_().remove(posKey_(r.code, name));
  if (!r.members.length) { rc_().remove('room_' + r.code); return { ok: true }; }
  if (r.host === name) r.host = r.members[0];
  roomPut_(r);
  return { ok: true };
}
