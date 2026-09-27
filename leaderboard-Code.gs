/**
 * Pocket Plays – world leaderboard backend (Google Apps Script)
 *
 * 1. Create a Google Sheet → Extensions → Apps Script → paste this whole file → Save.
 * 2. Deploy → New deployment → Web app. Execute as: Me · Who has access: Anyone → Deploy → Authorize.
 * 2b. Before deploying, change OWNER_KEY below to your own password. You type it once in the cockpit.
 * 3. Copy the Web app URL (ends in /exec) and paste it into pocket.html (window.PP_LB_URL).
 * Owner dashboard: open leaderboard.html on your site and sign in with OWNER_KEY.
 * After changing this file: Deploy → Manage deployments → Edit → Version: New version → Deploy (URL stays the same).
 *
 * Stores only: a random device-player id, a player number (P1, P2…), total XP, weekly XP and games played.
 * No names, no personal details.
 */
const OWNER_KEY = 'xxx';   // needed to VIEW the leaderboard (cockpit only)
const SHEET = 'Players';
const HEAD = ['uid', 'player', 'xp', 'weekKey', 'weekStartXp', 'plays', 'updated', 'name', 'hexId'];
const TOP_N = 20;
const LOG = 'Log';   // one row per finished game: time, player, XP earned, total XP, game
const FEEDBACK = 'Feedback';   // thumbs up/down, reasons and comments from players

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  else if (!sh.getRange(1, 8).getValue()) sh.getRange(1, 8, 1, 2).setValues([['name', 'hexId']]);
  return sh;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function weekKey_() { return Utilities.formatDate(new Date(), 'Asia/Kolkata', "YYYY-'W'ww"); }
function int_(x, lo, hi) { x = Math.round(Number(x) || 0); return Math.max(lo, Math.min(hi, x)); }

function rows_() {
  const v = sheet_().getDataRange().getValues(); v.shift();
  return v.map((r, i) => ({ row: i + 2, uid: String(r[0]), player: Number(r[1]), xp: Number(r[2]) || 0, weekKey: String(r[3]), weekStartXp: Number(r[4]) || 0, plays: Number(r[5]) || 0, name: String(r[7] || ''), hid: String(r[8] || '') }));
}

function board_(uid) {
  const wk = weekKey_(), rows = rows_();
  const all = rows.filter(r => r.xp > 0).sort((a, b) => b.xp - a.xp || a.player - b.player);
  const week = rows.map(r => ({ r, w: r.weekKey === wk ? r.xp - r.weekStartXp : 0 })).filter(o => o.w > 0).sort((a, b) => b.w - a.w || a.r.player - b.r.player);
  const out = { ok: true, week: wk, players: all.length,
    top: all.slice(0, TOP_N).map((r, i) => ({ rank: i + 1, p: 'P' + r.player, name: r.name, xp: r.xp, plays: r.plays })),
    weekTop: week.slice(0, TOP_N).map((o, i) => ({ rank: i + 1, p: 'P' + o.r.player, name: o.r.name, xp: o.w })), weekPlayers: week.length, me: null };
  if (uid) {
    const i = all.findIndex(r => r.uid === uid), j = week.findIndex(o => o.r.uid === uid), r = rows.find(x => x.uid === uid);
    if (r) out.me = { p: 'P' + r.player, xp: r.xp, rank: i >= 0 ? i + 1 : null, weekXp: j >= 0 ? week[j].w : 0, weekRank: j >= 0 ? j + 1 : null };
  }
  return out;
}

function doPost(e) {
  let d; try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  if (d.type === 'fb') { feedback_(d); return json_({ ok: true }); }
  const uid = String(d.uid || '');
  if (!/^[a-z0-9]{12,32}$/.test(uid)) return json_({ ok: false, error: 'bad_id' });
  const name = /^[A-Za-z0-9]{2,16}$/.test(String(d.name || '')) ? String(d.name) : '', hid = /^[0-9a-f]{12}$/.test(String(d.hid || '')) ? String(d.hid) : '';
  const game = clean_(d.game, 40), sent = int_(d.xp, 0, 10000000), gain = int_(d.gain, 0, 100000), plays = int_(d.plays, 0, 10000000), wk = weekKey_();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_(), rows = rows_(), r = rows.find(x => x.uid === uid);
    if (!r) {
      const props = PropertiesService.getScriptProperties(), n = Number(props.getProperty('next') || 0) + 1;
      props.setProperty('next', String(n));
      const xp = sent;
      sh.appendRow([uid, n, xp, wk, Math.max(0, xp - gain), plays, new Date(), name, hid]);  // only this game's XP counts for the week
      log_('P' + n, gain, xp, game);
    } else {
      const xp = Math.max(r.xp, sent);
      const ws = r.weekKey === wk ? r.weekStartXp : r.xp;
      sh.getRange(r.row, 3, 1, 7).setValues([[xp, wk, ws, Math.max(r.plays, plays), new Date(), name || r.name, hid || r.hid]]);
      log_('P' + r.player, gain, xp, game);
    }
  } finally { lock.releaseLock(); }
  return json_({ ok: true });   // players' phones get no leaderboard data back
}

function clean_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').replace(/^[=+\-@\s]+/, '').trim().slice(0, max); }

function log_(player, gain, xp, game) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let lg = ss.getSheetByName(LOG);
  if (!lg) { lg = ss.insertSheet(LOG); lg.appendRow(['time', 'player', 'xpEarned', 'totalXp', 'game']); lg.setFrozenRows(1); }
  else if (!lg.getRange(1, 5).getValue()) lg.getRange(1, 5).setValue('game');
  lg.appendRow([new Date(), player, gain, xp, game || '']);
}

function feedback_(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let fb = ss.getSheetByName(FEEDBACK);
  if (!fb) { fb = ss.insertSheet(FEEDBACK); fb.appendRow(['time', 'player', 'game', 'rating', 'reasons', 'comment']); fb.setFrozenRows(1); }
  const uid = String(d.uid || ''), r = /^[a-z0-9]{12,32}$/.test(uid) ? rows_().find(x => x.uid === uid) : null;
  const nm = /^[A-Za-z0-9]{2,16}$/.test(String(d.name || '')) ? String(d.name) : '';
  fb.appendRow([new Date(), r ? (nm || r.name ? (nm || r.name) + ' (P' + r.player + ')' : 'P' + r.player) : nm, clean_(d.game, 60), clean_(d.rating, 20), clean_(d.reasons, 300), clean_(d.comment, 2000)]);
}

// Everything the owner dashboard (leaderboard.html) needs.
function stats_() {
  const wk = weekKey_(), tz = 'Asia/Kolkata';
  const players = sheet_().getDataRange().getValues().slice(1).map(r => ({
    p: 'P' + r[1], xp: Number(r[2]) || 0, weekXp: String(r[3]) === wk ? (Number(r[2]) || 0) - (Number(r[4]) || 0) : 0,
    plays: Number(r[5]) || 0, updated: r[6] instanceof Date ? r[6].toISOString() : String(r[6]), name: String(r[7] || ''), hid: String(r[8] || '') }));
  const lg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG);
  let log = [];
  if (lg && lg.getLastRow() > 1) {
    const n = Math.min(20000, lg.getLastRow() - 1);
    log = lg.getRange(lg.getLastRow() - n + 1, 1, n, 5).getValues();
  }
  const days = {};
  log.forEach(r => { if (!(r[0] instanceof Date)) return; const d = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd'); const o = days[d] || (days[d] = { games: 0, xp: 0, players: {} }); o.games++; o.xp += Number(r[2]) || 0; o.players[r[1]] = 1; });
  const daily = Object.keys(days).sort().slice(-30).map(d => ({ day: d, games: days[d].games, xp: days[d].xp, players: Object.keys(days[d].players).length }));
  const nameOf = {}; players.forEach(x => { nameOf[x.p] = x.name; });
  const recent = log.slice(-25).reverse().map(r => ({ time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), p: r[1], name: nameOf[r[1]] || '', gain: Number(r[2]) || 0, xp: Number(r[3]) || 0, game: String(r[4] || '') }));
  const games = {};
  const G = n => games[n] || (games[n] = { game: n, plays: 0, players: {}, xp: 0, up: 0, down: 0 });
  log.forEach(r => { if (!r[4]) return; const g = G(String(r[4])); g.plays++; g.players[r[1]] = 1; g.xp += Number(r[2]) || 0; });
  const fs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FEEDBACK);
  let fbRows = [];
  if (fs && fs.getLastRow() > 1) { const n = Math.min(5000, fs.getLastRow() - 1); fbRows = fs.getRange(fs.getLastRow() - n + 1, 1, n, 6).getValues(); }
  fbRows.forEach(r => { const rt = String(r[3]).toLowerCase(), name = String(r[2] || 'App'), k = rt === 'enjoying' || rt === 'up' ? 'up' : rt === 'not enjoying' || rt === 'down' ? 'down' : ''; if (k) G(name)[k]++; });
  const feedback = fbRows.slice(-50).reverse().map(r => ({ time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), p: String(r[1] || ''), game: String(r[2] || ''), rating: String(r[3] || ''), reasons: String(r[4] || ''), comment: String(r[5] || '') }));
  const gameList = Object.keys(games).map(k => { const g = games[k]; return { game: g.game, plays: g.plays, players: Object.keys(g.players).length, xp: g.xp, up: g.up, down: g.down }; });
  return { ok: true, week: wk, players: players, daily: daily, recent: recent, games: gameList, feedback: feedback, feedbackTotal: fbRows.length };
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.action === 'stats') {
    if (String(p.key || '') !== OWNER_KEY || OWNER_KEY === 'change-this-password') return json_({ ok: false, error: 'wrong_key' });
    return json_(stats_());
  }
  if (p.action === 'top') {
    if (String(p.key || '') !== OWNER_KEY || OWNER_KEY === 'change-this-password') return json_({ ok: false, error: 'wrong_key' });
    return json_(board_(String(p.uid || '')));
  }
  return json_({ ok: true, message: 'Pocket Plays leaderboard is running.' });
}
