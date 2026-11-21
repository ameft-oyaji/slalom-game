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
      case 'room_ev': return out_(room_ev_(q));
      case 'room_chat': return out_(room_chat_(q));
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
  for (let i = 1; i < v.length; i++) {
    const d = v[i][5] instanceof Date ? v[i][5] : (v[i][5] ? new Date(v[i][5]) : null);
    const at = d && !isNaN(d) ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') : '';   // the day this best was set
    rows.push({ name: String(v[i][0]), course: v[i][1], t: Number(v[i][2]), r1: Number(v[i][3]), r2: Number(v[i][4]), at: at });
  }
  return { ok: true, rows: rows };
}

// Player ranking: each player's career summary for every season they have played (no personal data beyond the name).
// Seasons a player reset are kept in "<season key>-arch" and still count toward the all-time figures
function seasonStat_(d) {
  const done = (d && d.done) || {}, places = (d && d.places && d.places.you) || [];
  const o = { races: 0, best: 0, vmax: 0, top10: 0, rankSum: 0, ranked: 0,
              wins: Number(places[0] || 0), podiums: Number(places[0] || 0) + Number(places[1] || 0) + Number(places[2] || 0) };
  Object.keys(done).forEach(function (k) {
    const r = done[k]; if (!r) return;
    if (r.status !== 'dns') o.races++;
    if (r.rank) { if (!o.best || r.rank < o.best) o.best = r.rank; if (r.rank <= 10) o.top10++; o.rankSum += r.rank; o.ranked++; }
    if (r.v && r.v > o.vmax) o.vmax = r.v;
  });
  return o;
}
function players_() {
  const v = data_().getDataRange().getValues(), who = {}, cur = {}, arch = {};
  for (let i = 1; i < v.length; i++) {
    const key = String(v[i][1]), name = String(v[i][0]);
    if (key === 'slalomwc-prefs') {
      try { const pf = JSON.parse(v[i][2] || '{}'); who[name] = { nat: String(pf.nat || ''), gear: pf.gear || {} }; } catch (e) {}
      continue;
    }
    const m = /^slalomwc-2627-((?:men|women)-[a-z]+(?:-gs)?)(-arch)?$/.exec(key);   // slalom, giant slalom (-gs), kept seasons (-arch)
    if (!m) continue;
    try {
      const d = JSON.parse(v[i][2] || (m[2] ? '[]' : '{}')), id = name + '\u0001' + m[1];
      if (m[2]) arch[id] = Array.isArray(d) ? d : []; else cur[id] = d;
    } catch (e) {}
  }
  const ids = {}; Object.keys(cur).forEach(function (k) { ids[k] = 1; }); Object.keys(arch).forEach(function (k) { ids[k] = 1; });
  const rows = Object.keys(ids).map(function (id) {
    const parts = id.split('\u0001'), d = cur[id] || {}, now = seasonStat_(d);
    const all = [now].concat((arch[id] || []).map(seasonStat_));
    const sum = function (f) { return all.reduce(function (a, o) { return a + o[f]; }, 0); };
    const bests = all.map(function (o) { return o.best; }).filter(function (b) { return b > 0; });
    const ranked = sum('ranked');
    return { name: parts[0], key: parts[1], pts: Number((d.pts && d.pts.you) || 0),
             races: sum('races'), wins: sum('wins'), podiums: sum('podiums'), top10: sum('top10'),
             best: bests.length ? Math.min.apply(null, bests) : 0,
             vmax: Math.max.apply(null, all.map(function (o) { return o.vmax; })),
             avg: ranked ? Math.round(sum('rankSum') / ranked * 10) / 10 : 0,
             seasons: all.length - (cur[id] ? 0 : 1),
             now: { races: now.races, wins: now.wins, podiums: now.podiums, best: now.best } };
  });
  return { ok: true, rows: rows, who: who };
}

// ---------------------------------------------------------------------
// Online battle rooms (kept in the script cache for up to 6 hours; nothing is written to the sheet)
// ---------------------------------------------------------------------
const RC_TTL = 21600;
// a World Cup course (gender_date), or a dual race on a rhythm set (gender_dual-SL|GS-seed)
const COURSE_RE = /^(men|women)_(\d{4}-\d{2}-\d{2}|dual-(SL|GS)-\d{1,6})$/;
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
  if (!COURSE_RE.test(course)) return { ok: false, error: 'course' };
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
  if (q.course && COURSE_RE.test(String(q.course))) r.course = String(q.course);
  r.run = Number(q.run) === 2 ? 2 : 1;                                 // dual race: which of the two runs
  if (q.ev != null) r.ev = evParse_(q.ev);                             // dual tournament / league: pairings and results so far
  r.raceId += 1; r.state = 'race'; r.startAt = Date.now() + 9000;
  rc_().removeAll(r.members.map(function (m) { return posKey_(r.code, m); }));
  roomPut_(r);
  return roomView_(r);
}
// a dual tournament or league among the room's members, worked out by the host's game and kept with the room
function evParse_(s) {
  s = s && typeof s === 'object' ? JSON.stringify(s) : String(s || '');
  if (!s || s === 'null' || s.length > 8000) return null;
  try { const e = JSON.parse(s); return e && (e.fmt === 'tour' || e.fmt === 'league') ? e : null; } catch (err) { return null; }
}
function room_ev_(q) {
  const name = userByToken_(q.token);
  const r = roomGet_(String(q.code || ''));
  if (!r) return { ok: false, error: 'noroom' };
  if (r.host !== name) return { ok: false, error: 'host' };
  r.ev = evParse_(q.ev);
  roomPut_(r);
  return roomView_(r);
}
// room chat: the last 30 lines, kept with the room
function room_chat_(q) {
  const name = userByToken_(q.token);
  if (!name) return { ok: false, error: 'auth' };
  const r = roomGet_(String(q.code || ''));
  if (!r) return { ok: false, error: 'noroom' };
  if (r.members.indexOf(name) < 0) return { ok: false, error: 'member' };
  const t = String(q.text || '').replace(/\s+/g, ' ').trim().slice(0, 40);
  if (!t) return { ok: false, error: 'empty' };
  r.chat = (r.chat || []).concat([{ n: name, t: t, ts: Date.now() }]).slice(-30);
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
              fin: q.fin == null ? null : num(q.fin), out: q.out ? String(q.out).slice(0, 3) : '', ts: Date.now(),
              t1: q.t1 == null ? null : num(q.t1), o1: q.o1 ? String(q.o1).slice(0, 3) : '',      // dual race run 2: the first run's result
              sp: q.sp ? 1 : 0, ch: q.ch ? String(q.ch).slice(0, 24) : '', cid: Number(q.cid) || 0 };  // a watcher's cheer (speech bubble)
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
