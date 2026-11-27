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
  if (a === 'fislist') return out_(fislist_());
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
      case 'room_ready': return out_(room_ready_(q));
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
  try { registrantList(); } catch (e) { /* the account is made whatever happens here */ }
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
  if (!/^(men|women)_\d{4}-\d{2}-\d{2}(_race)?$/.test(course)) return { ok: false, error: 'course' };   // _race: the best two-run total from the World Cup race itself
  if (!(t > 15 && t < 400 && r1 > 5 && r2 > 5)) return { ok: false, error: 'time' };
  const s = ta_(), v = s.getDataRange().getValues();
  // the course record before this one (time attack and race records together); a first-ever record is not news
  const base = course.replace(/_race$/, '');
  let prev = null;
  for (let i = 1; i < v.length; i++) {
    const c = String(v[i][1]).replace(/_race$/, ''), tt = Number(v[i][2]);
    if (c === base && tt > 0 && !/^zz-/.test(String(v[i][0])) && (!prev || tt < prev.t)) prev = { t: tt, name: String(v[i][0]) };   // test accounts do not count
  }
  const news = function () { if (prev && t < prev.t) recordNews_(name, course, t, prev, q.label); };
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === name && v[i][1] === course) {
      if (t < Number(v[i][2])) { s.getRange(i + 1, 3, 1, 4).setValues([[t, r1, r2, new Date()]]); news(); return { ok: true, improved: true }; }
      return { ok: true, improved: false };
    }
  }
  s.appendRow([name, course, t, r1, r2, new Date()]);
  news();
  return { ok: true, improved: true };
}

// ---- a new course record: tell the LINE Official Account's friends (broadcast).
// The channel access token is the script property LINE_TOKEN (Project Settings → Script properties); without it nothing is sent.
// Test accounts (zz-…) never send.
// Run this once from the editor (select "checkLine" and press 実行): it asks for permission to reach LINE
// and checks the token by reading the account's name. No message is sent
function checkLine() {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) { Logger.log('LINE_TOKEN がスクリプト プロパティにありません'); return; }
  const r = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  Logger.log(r.getResponseCode() === 200 ? 'OK：' + JSON.parse(r.getContentText()).displayName + ' に接続できました' : 'トークンが違います（' + r.getResponseCode() + '）');
}
function fmtT_(t) { const m = Math.floor(t / 60), s = (t - m * 60).toFixed(2); return m + ':' + (s < 10 ? '0' : '') + s; }
function recordNews_(name, course, t, prev, label) {
  try {
    if (/^zz-/.test(name)) return;
    const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
    if (!token) return;
    const where = String(label || '').replace(/[\r\n]/g, ' ').slice(0, 60) || course;
    const text = '🏁 コースレコード更新！\n' + where + (/_race$/.test(course) ? '（大会）' : '（タイムアタック）') + '\n\n' +
      fmtT_(t) + '　' + name + '\n' +
      '（これまで ' + fmtT_(prev.t) + '　' + prev.name + '、' + (prev.t - t).toFixed(2) + '秒更新）\n\n' +
      '▶ 挑戦する：https://ameft-oyaji.github.io/slalom-game/';
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ messages: [{ type: 'text', text: text }] })
    });
  } catch (e) { /* the record itself is saved whatever happens here */ }
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
const COURSE_RE = /^(men|women)_(\d{4}-\d{2}-\d{2}|dual-(SL|GS)-\d{1,6}|incol-(SL|GS|SG)-\d{1,6})$/;   // + the インカレ (up to 4 players)
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
    if (r.members.length >= (/_incol-/.test(r.course) ? 4 : 8)) return { ok: false, error: 'full' };
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
  // a new dual race (from the lobby, or "again" after one) waits until everyone else has pressed ready
  if (q.fresh && /_(dual|incol)-/.test(String(q.course || r.course))) {
    const rd = r.ready || {};
    if (r.members.some(function (m) { return m !== r.host && !rd[m]; })) return { ok: false, error: 'notready' };
  }
  if (q.course && COURSE_RE.test(String(q.course))) r.course = String(q.course);
  r.run = Number(q.run) === 2 ? 2 : 1;                                 // dual race: which of the two runs
  if (q.ev != null) r.ev = evParse_(q.ev);                             // dual tournament / league: pairings and results so far
  if (q.ik != null) r.ik = ikParse_(q.ik);                             // インカレ: whose turn it is and the players' results so far
  r.raceId += 1; r.state = 'race'; r.startAt = Date.now() + Math.min(30000, Math.max(9000, Number(q.lead) || 0));   // lead: a longer wait (the start list is shown first)
  r.ready = {};                                                        // everyone presses ready again before the next new race
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
function ikParse_(s) {
  s = String(s || '');
  if (!s || s === 'null' || s.length > 6000) return null;
  try { const o = JSON.parse(s); return o && typeof o === 'object' ? o : null; } catch (err) { return null; }
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
// a member says they are ready (or not) for the next dual race
function room_ready_(q) {
  const name = userByToken_(q.token);
  if (!name) return { ok: false, error: 'auth' };
  const r = roomGet_(String(q.code || ''));
  if (!r) return { ok: false, error: 'noroom' };
  if (r.members.indexOf(name) < 0) return { ok: false, error: 'member' };
  r.ready = r.ready || {};
  if (Number(q.on)) r.ready[name] = 1; else delete r.ready[name];
  if (q.uni && /^[A-Z]{3}$/.test(String(q.uni))) { r.uni = r.uni || {}; r.uni[name] = String(q.uni); }   // インカレ: the player's university
  if (q.fp != null) { r.fp = r.fp || {}; r.fp[name] = fpParse_(q.fp); }                  // the player's FIS points (start order)
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
  if (r.ready) delete r.ready[name];
  if (r.uni) delete r.uni[name];
  if (r.fp) delete r.fp[name];
  rc_().remove(posKey_(r.code, name));
  if (!r.members.length) { rc_().remove('room_' + r.code); return { ok: true }; }
  if (r.host === name) r.host = r.members[0];
  roomPut_(r);
  return { ok: true };
}

// ---------------------------------------------------------------------
// 登録者一覧 (for the owner only): a readable sheet of the registered players, rebuilt
// whenever someone registers, from the spreadsheet menu 「スラローム」→「登録者一覧を更新」,
// or by running registrantList from the editor. PINs, salts and tokens are never copied
const SHEET_LIST = '登録者一覧';
function onOpen() {
  SpreadsheetApp.getUi().createMenu('スラローム').addItem('登録者一覧を更新', 'registrantList').addToUi();
}
function registrantList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const u = users_().getDataRange().getValues(), d = data_().getDataRange().getValues(), t = ta_().getDataRange().getValues();
  const last = {}, nat = {}, recs = {};
  for (let i = 1; i < d.length; i++) {
    const n = String(d[i][0]), at = d[i][3] instanceof Date ? d[i][3] : null;
    if (at && (!last[n] || at > last[n])) last[n] = at;
    if (d[i][1] === 'slalomwc-prefs') { try { nat[n] = String(JSON.parse(d[i][2] || '{}').nat || ''); } catch (e) {} }
  }
  for (let i = 1; i < t.length; i++) { const n = String(t[i][0]); recs[n] = (recs[n] || 0) + 1; }
  const fmt = function (x) { return x instanceof Date ? Utilities.formatDate(x, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : ''; };
  const rows = [];
  for (let i = 1; i < u.length; i++) {
    const n = String(u[i][0]); if (!n) continue;
    const late = [u[i][5], last[n]].filter(function (x) { return x instanceof Date; }).sort(function (a, b) { return b - a; })[0];
    rows.push({ c: u[i][4] instanceof Date ? u[i][4] : new Date(0),
      v: [n, nat[n] || '', fmt(u[i][4]), fmt(late), recs[n] || 0,
          (/^zz-/.test(n) ? 'テスト用' : '') + (Number(u[i][6] || 0) >= 10 ? (/^zz-/.test(n) ? '・' : '') + 'PINロック中' : '')] });
  }
  rows.sort(function (a, b) { return b.c - a.c; });   // newest first
  let s = ss.getSheetByName(SHEET_LIST);
  if (!s) s = ss.insertSheet(SHEET_LIST);
  s.clear();
  const real = rows.filter(function (r) { return !/^zz-/.test(r.v[0]); }).length;
  s.getRange(1, 1).setValue('登録者 ' + real + '人（テスト用を除く）　更新：' + fmt(new Date()));
  const head = ['No.', '選手名', '国', '登録日', '最後に遊んだ日', 'コース記録の数', 'メモ'];
  s.getRange(2, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#0E1A2E').setFontColor('#ffffff');
  if (rows.length) s.getRange(3, 1, rows.length, head.length).setValues(rows.map(function (r, i) { return [rows.length - i].concat(r.v); }));
  s.setFrozenRows(2);
  s.autoResizeColumns(1, head.length);
  return real;
}

// ---------------------------------------------------------------------
// FIS points list: every player's race points (kept by the game in "slalomwc-fis") made into a list the FIS way:
// the average of the best two results of the last 8 weeks (one result: +20%), shown with the leader at 0.00.
// The official list comes out every Monday 0:00 (Japan time) and uses the results before then;
// "live" uses everything up to now. Test accounts (zz-…) are left out
// ---------------------------------------------------------------------
const FIS_MAX_ = { SL: 165, GS: 220, SG: 270 };
const FIS_KEEP_ = 56 * 864e5;
function fpParse_(s) {
  let o = {}; try { o = typeof s === 'object' ? s : JSON.parse(String(s)); } catch (e) { o = {}; }
  const out = {};
  ['SL', 'GS', 'SG'].forEach(function (d) { const v = Number(o && o[d]); if (o && o[d] != null && isFinite(v) && v >= 0 && v < 2000) out[d] = Math.round(v * 100) / 100; });
  return out;
}
function fisMonday_(now) {
  const j = new Date(now + 9 * 36e5), day = (j.getUTCDay() + 6) % 7;
  return Date.UTC(j.getUTCFullYear(), j.getUTCMonth(), j.getUTCDate()) - day * 864e5 - 9 * 36e5;
}
function fisResults_() {
  const v = data_().getDataRange().getValues(), res = {}, nat = {};
  for (let i = 1; i < v.length; i++) {
    const name = String(v[i][0]), key = String(v[i][1]);
    if (/^zz-/.test(name)) continue;
    if (key === 'slalomwc-prefs') { try { nat[name] = String(JSON.parse(v[i][2] || '{}').nat || ''); } catch (e) {} continue; }
    if (key !== 'slalomwc-fis') continue;
    try {
      const o = JSON.parse(v[i][2] || '{}');
      res[name] = (o.res || []).filter(function (x) { return x && /^(men|women)$/.test(x.g) && FIS_MAX_[x.d] && isFinite(+x.p) && +x.p >= 0 && +x.at > 0; })
        .map(function (x) { return { g: x.g, d: x.d, p: +x.p, at: +x.at }; });
    } catch (e) {}
  }
  return { res: res, nat: nat };
}
function fisVal_(ps) {
  if (!ps.length) return null;
  ps.sort(function (a, b) { return a - b; });
  return Math.round((ps.length >= 2 ? (ps[0] + ps[1]) / 2 : ps[0] * 1.2) * 100) / 100;
}
// the lists as they stood at time `at`: { 'men-SL': [{ name, nat, raw, p, rank, n }], ..., 'men-ALL': ... }
function fisListAt_(R, at) {
  const lists = {};
  Object.keys(R.res).forEach(function (name) {
    const mine = R.res[name].filter(function (x) { return x.at < at && x.at >= at - FIS_KEEP_; });
    ['men', 'women'].forEach(function (g) {
      const raw = {};
      ['SL', 'GS', 'SG'].forEach(function (d) {
        const ps = mine.filter(function (x) { return x.g === g && x.d === d; }).map(function (x) { return x.p; });
        const val = fisVal_(ps);
        if (val == null) return;
        raw[d] = val;
        (lists[g + '-' + d] = lists[g + '-' + d] || []).push({ name: name, nat: R.nat[name] || '', raw: val, n: ps.length });
      });
      if (raw.SL != null || raw.GS != null) {
        const all = Math.round(((raw.SL != null ? raw.SL : FIS_MAX_.SL) + (raw.GS != null ? raw.GS : FIS_MAX_.GS)) / 2 * 100) / 100;
        const n = mine.filter(function (x) { return x.g === g && (x.d === 'SL' || x.d === 'GS'); }).length;
        (lists[g + '-ALL'] = lists[g + '-ALL'] || []).push({ name: name, nat: R.nat[name] || '', raw: all, n: n });
      }
    });
  });
  Object.keys(lists).forEach(function (k) {
    const L = lists[k].sort(function (a, b) { return a.raw - b.raw || b.n - a.n; }), top = L[0].raw;
    L.forEach(function (r) { r.p = Math.round((r.raw - top) * 100) / 100; r.rank = 1 + L.filter(function (x) { return x.raw < r.raw; }).length; });
  });
  return lists;
}
function fisWithPrev_(cur, prev) {
  Object.keys(cur).forEach(function (k) {
    const pm = {}; (prev[k] || []).forEach(function (r) { pm[r.name] = r.rank; });
    cur[k].forEach(function (r) { r.pr = pm[r.name] != null ? pm[r.name] : null; });
  });
  return cur;
}
function fislist_() {
  const now = Date.now(), cut = fisMonday_(now), R = fisResults_();
  const list = fisWithPrev_(fisListAt_(R, cut), fisListAt_(R, cut - 7 * 864e5));
  const live = fisWithPrev_(fisListAt_(R, now + 1), list);
  fisWeekNews_(cut, list);
  return { ok: true, cut: cut, list: list, live: live };
}
// once a week, when the new official list is first asked for: its leaders go to the LINE Official Account's friends
function fisWeekNews_(cut, list) {
  try {
    const P = PropertiesService.getScriptProperties();
    if (P.getProperty('fisWeek') === String(cut)) return;
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(3000)) return;
    try {
      if (P.getProperty('fisWeek') === String(cut)) return;
      P.setProperty('fisWeek', String(cut));
      const names = { 'men-SL': '男子回転', 'men-GS': '男子大回転', 'men-SG': '男子スーパーG', 'men-ALL': '男子総合', 'women-SL': '女子回転', 'women-GS': '女子大回転', 'women-SG': '女子スーパーG', 'women-ALL': '女子総合' };
      const lines = Object.keys(names).filter(function (k) { return list[k] && list[k].length; }).map(function (k) {
        const tops = list[k].filter(function (r) { return r.rank === 1; }).map(function (r) { return r.name; }).join('・');
        return '👑 ' + names[k] + '：' + tops + '（' + list[k].length + '人中）';
      });
      if (!lines.length) return;
      const token = P.getProperty('LINE_TOKEN');
      if (!token) return;
      const d = new Date(cut + 9 * 36e5);
      const text = '📋 FISポイントリスト ' + (d.getUTCMonth() + 1) + '/' + d.getUTCDate() + ' 発表\n\n' + lines.join('\n') +
        '\n\nランキング →「FISポイント」で全員の順位が見られます\n▶ https://ameft-oyaji.github.io/slalom-game/';
      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ messages: [{ type: 'text', text: text }] })
      });
    } finally { lock.releaseLock(); }
  } catch (e) { /* the list itself is answered whatever happens here */ }
}
