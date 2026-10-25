/**
 * Greco Picks — Google Apps Script backend
 *
 * Paste into the Sheet's Extensions → Apps Script editor as Code.gs,
 * and add an HTML file named exactly "Index" with Index.html.
 *
 * Run from the editor, in this order:
 *   setup()          – builds the tabs (safe to re-run; moves existing players to the Players tab)
 *   dryRunPreview()  – pulls this week's games into a Preview tab ONLY (no other changes)
 *   refreshGames()   – loads this week's games for real and grades any finished picks
 *   installTriggers()– refreshes schedule/scores automatically every 30 minutes
 *   showLinks()      – after deploying, makes a personal link for every name on the Players tab
 *                      (add someone: type their name on the Players tab, run showLinks)
 *   rotateKeys()     – makes new links for EVERYONE; old links stop working
 *                      (one person only: clear their Key cell on Players, run showLinks)
 *   testFeeds()      – diagnostic: shows which data sources Google can reach
 *   testPush()       – sends a test notification to you (needs OneSignalKey on the Settings tab)
 *
 * Notifications (OneSignal): pick reminders 2 hours before each kickoff window (only to people
 * with unpicked games), "new week" when a new week's games load, and weekly results once a week
 * is final. They're checked every time refreshGames runs. Each person can turn types off in the app.
 *
 * The installable app (GitHub Pages) talks to this script through doGet(?api=...).
 * Put the app's address in AppURL on the Settings tab and showLinks will use it.
 */

const SCRIPT_VERSION = '2.9.9';   // shown in the app (⚙). Use the same number in the deployment description.
const TZ = 'America/New_York';
const TABS = { SETTINGS: 'Settings', PLAYERS: 'Players', GAMES: 'Games', PICKS: 'Picks', RECORDS: 'Records', PREVIEW: 'Preview',
  TIEBREAK: 'Tiebreakers', REACT: 'Reactions', GAMES_ARCH: 'Games Archive', PICKS_ARCH: 'Picks Archive' };
const REACTIONS = ['😂', '🔥', '💀', '🤡', '😤', '👀'];
const PLAYER_HEADERS = ['ID', 'Name', 'Key', 'Link', 'Active', 'Remind', 'NewWeek', 'Results', 'Evening', 'Groups', 'Hidden', 'Upsets'];
const EVENING_HOUR = 20;   // night-before reminder, 8 PM Eastern
const ONESIGNAL_APP_ID = 'e0fe6864-f2ca-4b41-8c87-00586a7634e4';
const REMIND_HOURS = 2;
const GAME_HEADERS = ['GameID', 'Season', 'SeasonType', 'Week', 'Kickoff', 'Away', 'AwayAbbr', 'AwayLogo', 'AwayScore',
  'Home', 'HomeAbbr', 'HomeLogo', 'HomeScore', 'Status', 'Detail', 'Winner', 'LastUpdated', 'AwayRecord', 'HomeRecord',
  'AwayRecordPre', 'HomeRecordPre'];
const PICK_HEADERS = ['GameID', 'Season', 'SeasonType', 'Week', 'Player', 'Pick', 'PickedAt', 'Override', 'Result'];
const G = {}; GAME_HEADERS.forEach((h, i) => { G[h] = i; });
const P = {}; PICK_HEADERS.forEach((h, i) => { P[h] = i; });

// ======================================================================
// Editor functions
// ======================================================================

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());

  const s = ensureSheet_(TABS.SETTINGS, ['Setting', 'Value']);
  if (s.getLastRow() < 2) {
    s.getRange('B:B').setNumberFormat('@');
    s.getRange(2, 1, 2, 2).setValues([['AdminPlayer', 'P1'], ['WebAppURL', '']]);
    s.setColumnWidth(1, 140);
    s.setColumnWidth(2, 520);
  }

  if (!('AppURL' in settings_())) setSetting_('AppURL', '');

  // Players tab (moves the original two players over, keeping their links working)
  const pl = ensureSheet_(TABS.PLAYERS, PLAYER_HEADERS);
  pl.getRange(1, 1, 1, PLAYER_HEADERS.length).setValues([PLAYER_HEADERS]).setFontWeight('bold');
  pl.getRange('C:D').setNumberFormat('@');
  if (!('OneSignalKey' in settings_())) setSetting_('OneSignalKey', '');
  if (pl.getLastRow() < 2) {
    const st = settings_();
    const rows = [];
    if (st.Player1Key) rows.push(['P1', st.Player1Name || 'Paul', st.Player1Key, st.Player1Link || '', true]);
    if (st.Player2Key) rows.push(['P2', st.Player2Name || 'Player 2', st.Player2Key, st.Player2Link || '', true]);
    if (!rows.length) rows.push(['P1', 'Paul', newKey_(), '', true]);
    pl.getRange(2, 1, rows.length, PLAYER_HEADERS.length).setValues(rows);
    pl.setColumnWidth(2, 140);
    pl.setColumnWidth(3, 280);
    pl.setColumnWidth(4, 520);
  }
  ensureSheet_(TABS.GAMES, GAME_HEADERS).getRange('A:A').setNumberFormat('@');
  ensureSheet_(TABS.PICKS, PICK_HEADERS).getRange('A:A').setNumberFormat('@');
  ensureSheet_(TABS.RECORDS, ['Season', 'Week']);

  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  SpreadsheetApp.flush();
  Logger.log('Setup done. Players live on the Players tab.');
}

/** Dry run: writes this week's games to the Preview tab and changes nothing else. */
function dryRunPreview() {
  const games = fetchScoreboard_();
  const sh = ensureSheet_(TABS.PREVIEW);
  sh.clear();
  const rows = [['Week', 'Kickoff (ET)', 'Away', 'Home', 'Status', 'Score (away-home)', 'Winner', 'Tie option']];
  games.forEach(g => rows.push([
    weekLabel_(g.seasonType, g.week) + ' ' + g.season,
    Utilities.formatDate(g.kickoff, TZ, 'EEE M/d h:mm a'),
    g.away + ' (' + g.awayAbbr + ')',
    g.home + ' (' + g.homeAbbr + ')',
    g.detail || g.status,
    g.status === 'pre' ? '' : g.awayScore + '-' + g.homeScore,
    g.winner,
    g.seasonType === 2 ? 'Yes' : 'No',
  ]));
  const rng = sh.getRange(1, 1, rows.length, rows[0].length);
  rng.setNumberFormat('@');
  rng.setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, rows[0].length);
  Logger.log('DRY RUN: ' + games.length + ' games written to the Preview tab. Games/Picks/Records were not touched.');
}

/** Pulls the current week (plus any earlier week with unfinished games), then grades picks. */
function refreshGames() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const current = fetchScoreboard_();
    if (current.length) {
      upsertGames_(current);
      PropertiesService.getScriptProperties().setProperty('CURRENT_WEEK', weekKey_(current[0]));
    }
    const fetched = new Set(current.map(weekKey_));
    staleWeeks_().filter(k => !fetched.has(k)).slice(0, 3).forEach(k => {
      const [season, type, week] = k.split('|');
      upsertGames_(fetchScoreboard_(season, type, week));
    });
    gradeAndRecord_();
    try { notifications_(); } catch (e) { Logger.log('Notifications skipped: ' + e.message); }
    try { archiveOldSeasons_(); } catch (e) { Logger.log('Archive skipped: ' + e.message); }
  } finally {
    lock.releaseLock();
  }
}

function installTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshGames')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshGames').timeBased().everyMinutes(30).create();
  Logger.log('Auto-refresh installed: every 30 minutes.');
}

function showLinks() {
  const s = settings_();
  const app = String(s.AppURL || '').trim().replace(/[?#].*$/, '');
  const web = String(s.WebAppURL || '').trim().replace(/\?.*$/, '');
  let url;
  if (app) {
    if (!/^https:\/\//.test(app)) throw new Error('AppURL on the Settings tab should start with https://');
    url = app.endsWith('/') ? app : app + '/';
  } else if (/^https:\/\/script\.google\.com\/.+\/exec$/.test(web)) {
    url = web;
  } else {
    throw new Error('Put the app address in AppURL (or the /exec URL in WebAppURL) on the Settings tab first.');
  }
  const sh = sheet_(TABS.PLAYERS);
  const vals = sh.getDataRange().getValues();
  let maxId = 0;
  vals.slice(1).forEach(r => { const m = /^P(\d+)$/i.exec(String(r[0]).trim()); if (m) maxId = Math.max(maxId, Number(m[1])); });
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (!String(r[1]).trim()) continue;                       // no name, skip row
    if (!String(r[0]).trim()) r[0] = 'P' + (++maxId);
    if (!/^[a-f0-9]{32}$/.test(String(r[2]).trim())) r[2] = newKey_();
    if (r[4] === '' || r[4] === null) r[4] = true;
    r[3] = url + '?k=' + String(r[2]).trim();
    Logger.log(r[1] + ': ' + r[3]);
  }
  sh.getRange('C:D').setNumberFormat('@');
  const width = PLAYER_HEADERS.length;
  const out = vals.map((r, i) => {
    const row = r.slice(0, width);
    while (row.length < width) row.push('');
    return i === 0 ? PLAYER_HEADERS.slice() : row;
  });
  sh.getRange(1, 1, out.length, width).setValues(out);
  PLAYERS_ = null;
}

/** New links for everyone. To reset just one person, clear their Key cell on Players and run showLinks. */
function rotateKeys() {
  const sh = sheet_(TABS.PLAYERS);
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 3, n, 1).setValues(Array.from({ length: n }, () => [newKey_()]));
  showLinks();
  Logger.log('New links made for everyone. The old links no longer work.');
}

// ======================================================================
// Web app
// ======================================================================

function doGet(e) {
  noteLiveVersion_();
  if (e && e.parameter && e.parameter.api) return api_(e.parameter);
  const key = (e && e.parameter && e.parameter.k) || '';
  if (!playerForKey_(key)) {
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:24px">This link isn\'t valid. Ask for a new one.</p>')
      .setTitle('Greco Picks')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const t = HtmlService.createTemplateFromFile('Index');
  t.key = key; // validated: 32 hex characters only
  t.state = JSON.stringify(getState(key, '')).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('Greco Picks')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/**
 * Only the deployed copy answers the app, so whatever version answers is the live one.
 * Writes it to LiveVersion on the Settings tab (only when it changes).
 */
function noteLiveVersion_() {
  try {
    if (settings_().LiveVersion !== SCRIPT_VERSION) {
      setSetting_('LiveVersion', SCRIPT_VERSION);
      setSetting_('LiveSince', Utilities.formatDate(new Date(), TZ, 'M/d/yyyy h:mm a'));
    }
  } catch (e) { /* never block the app over this */ }
}

/** JSON API used by the installable app: ?api=state|pick|admin&k=KEY&... */
function api_(p) {
  let out;
  try {
    let data;
    if (p.api === 'state') data = getState(p.k, p.wk || '');
    else if (p.api === 'pick') data = submitPick(p.k, p.g, p.p);
    else if (p.api === 'admin') data = adminSetPick(p.k, p.g, p.pl, p.p, p.wk || '');
    else if (p.api === 'history') data = history_(p.k, p.season || '');
    else if (p.api === 'prefs') data = setPref_(p.k, p.type, p.on);
    else if (p.api === 'testme') data = testMe_(p.k);
    else if (p.api === 'report') data = report_(p.k, p.text, p.info);
    else if (p.api === 'rename') data = rename_(p.k, p.name);
    else if (p.api === 'people') data = people_(p.k);
    else if (p.api === 'addperson') data = addPerson_(p.k, p.name);
    else if (p.api === 'newlink') data = newLink_(p.k, p.pl);
    else if (p.api === 'groups') data = setGroups_(p.k, p.pl, p.groups);
    else if (p.api === 'hide') data = setHidden_(p.k, p.pl, p.on);
    else if (p.api === 'active') data = setActive_(p.k, p.pl, p.on);
    else if (p.api === 'nudge') data = nudge_(p.k, p.pl);
    else if (p.api === 'tiebreak') data = setTiebreak_(p.k, p.wk, p.guess);
    else if (p.api === 'react') data = react_(p.k, p.g, p.e);
    else throw new Error('Unknown request.');
    out = { ok: true, data: data };
  } catch (err) {
    out = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function getState(key, wk) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  const roster = players_().filter(p => p.active);

  const games = readGames_();
  const byWeek = {};
  games.forEach(g => { byWeek[weekKey_(g)] = g; });
  const weeks = Object.keys(byWeek).map(k => {
    const g = byWeek[k];
    return { key: k, season: g.season, type: g.seasonType, week: g.week, label: weekLabel_(g.seasonType, g.week) + ' · ' + g.season };
  }).sort((a, b) => b.season - a.season || b.type - a.type || b.week - a.week);

  let sel = wk && byWeek[wk] ? wk : PropertiesService.getScriptProperties().getProperty('CURRENT_WEEK');
  if (!sel || !byWeek[sel]) sel = weeks.length ? weeks[0].key : '';

  const picks = readPicks_();
  const now = Date.now();
  const reactions = readReactions_();
  const weekGames = games.filter(g => weekKey_(g) === sel).sort((a, b) => a.kickoff - b.kickoff || a.id.localeCompare(b.id));
  const list = weekGames.map(g => gameView_(g, me, picks, now, roster, reactions));

  const parts = sel ? sel.split('|') : ['', '', ''];
  const board = leaderboard_(picks.vals, parts, roster);
  const mine = board.find(b => b.id === me.id) || { week: '0-0', season: '0-0' };

  const others = roster.filter(p => p.id !== me.id).sort((a, b) => a.name.localeCompare(b.name));
  const meP = players_().find(p => p.id === me.id) || { groups: [], hidden: [] };
  const groups = groupNames_();
  return {
    me: { id: me.id, name: me.name, isAdmin: me.isAdmin, pushId: 'gp_' + me.id, prefs: prefsFor_(me.id),
      groups: groupsOf_(meP), defaultGroup: me.isAdmin ? '' : (meP.groups[0] || ''), hidden: meP.hidden },
    groups: groups,
    adminId: adminId_(),
    my: { week: mine.week, season: mine.season },
    players: [meP].concat(others).map(p => ({ id: p.id, name: p.name, groups: groupsOf_(p) })),
    weeks: weeks,
    selected: sel,
    games: list,
    board: board,
    crown: crown_(games, picks.vals, roster),
    crowns: crowns_(games, picks.vals, roster),
    tiebreak: tiebreakView_(weekGames, sel, me, roster, now),
    last: lastWeekFor_(games, picks.vals, me, roster),
    version: SCRIPT_VERSION,
  };
}

/**
 * One game as seen by one player. Other players' picks are only sent once you've
 * picked that game (or it has kicked off). Picks can be changed until kickoff.
 */
function gameView_(g, me, picks, now, roster, reactions) {
  const mine = picks.map[g.id + '|' + me.id];
  const locked = now >= g.kickoff.getTime();
  const others = roster.filter(p => p.id !== me.id).map(p => ({ p: p, pk: picks.map[g.id + '|' + p.id] }));
  const othersPicked = others.filter(x => x.pk).length;
  const reveal = !!mine || locked;
  const out = {
    id: g.id,
    kickoff: g.kickoff.toISOString(),
    status: g.status,
    detail: g.detail,
    winner: g.winner,
    canTie: g.seasonType === 2,
    locked: locked,
    frozen: false,
    away: { name: g.away, abbr: g.awayAbbr, logo: g.awayLogo, score: g.awayScore, record: g.awayRecord },
    home: { name: g.home, abbr: g.homeAbbr, logo: g.homeLogo, score: g.homeScore, record: g.homeRecord },
    mine: mine ? { pick: mine.pick, result: mine.result, override: mine.override } : null,
    othersPicked: othersPicked,
    othersTotal: others.length,
    others: reveal ? others.filter(x => x.pk).map(x => ({ name: x.p.name, pick: x.pk.pick, result: x.pk.result, override: x.pk.override })) : [],
  };
  out.all = {};
  roster.forEach(p => {
    const pk = picks.map[g.id + '|' + p.id];
    if (!pk) return;
    out.all[p.id] = (reveal || p.id === me.id) ? { p: pk.pick, r: pk.result, o: pk.override } : { h: 1 };
  });
  if (locked) out.react = ((reactions || readReactions_())[g.id]) || {};
  if (me.isAdmin && locked) {
    out.admin = roster.map(p => ({ id: p.id, name: p.name, pick: (picks.map[g.id + '|' + p.id] || {}).pick || '' }));
  }
  return out;
}

/** Saves one pick and returns just that game's updated card (fast path for the phone). */
function submitPick(key, gameId, pick) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  pick = String(pick).toUpperCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const g = readGames_().find(x => x.id === String(gameId));
    if (!g) throw new Error('Game not found.');
    if (Date.now() >= g.kickoff.getTime()) throw new Error('This game has kicked off — picks are locked.');
    const allowed = [g.awayAbbr, g.homeAbbr].concat(g.seasonType === 2 ? ['TIE'] : []);
    if (allowed.indexOf(pick) < 0) throw new Error('That pick isn\'t allowed for this game.');
    const roster = players_().filter(p => p.active);
    const picks = readPicks_();
    writePick_(g, me.id, pick, false, picks);
    const fresh = readPicks_();
    try { allInCheck_(weekKey_(g), fresh); } catch (e) { Logger.log('All-in check: ' + e.message); }
    return { game: gameView_(g, me, fresh, Date.now(), roster) };
  } finally {
    lock.releaseLock();
  }
}

/** Admin only: set, change, mark MISSED, or CLEAR anyone's pick (any time). */
function adminSetPick(key, gameId, playerId, pick, wk) {
  const me = playerForKey_(key);
  if (!me || !me.isAdmin) throw new Error('Not allowed.');
  if (!players_().some(p => p.id === playerId)) throw new Error('Unknown player.');
  pick = String(pick).toUpperCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const g = readGames_().find(x => x.id === String(gameId));
    if (!g) throw new Error('Game not found.');
    const allowed = [g.awayAbbr, g.homeAbbr, 'MISSED', 'CLEAR'].concat(g.seasonType === 2 ? ['TIE'] : []);
    if (allowed.indexOf(pick) < 0) throw new Error('That pick isn\'t allowed for this game.');
    if (pick === 'CLEAR') {
      const picks = readPicks_();
      const hit = picks.map[g.id + '|' + playerId];
      if (hit) picks.sh.deleteRow(hit.row);
    } else {
      writePick_(g, playerId, pick, true);
      try { allInCheck_(weekKey_(g), readPicks_()); } catch (e) { Logger.log('All-in check: ' + e.message); }
    }
    gradeAndRecord_();
  } finally {
    lock.releaseLock();
  }
  return getState(key, wk);
}

/**
 * Everything the Stats tab needs for one season (or 'all'): every game that has kicked off
 * and every pick on those games. Games that haven't started are left out, so nothing hidden leaks.
 */
function history_(key, season) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  const games = readGames_().concat(readArchivedGames_());
  const seasons = Array.from(new Set(games.map(g => g.season))).sort((a, b) => b - a);
  const want = season === 'all' ? 'all' : (Number(season) || seasons[0] || 0);
  const now = Date.now();
  const gs = games.filter(g => (want === 'all' || g.season === want) && now >= g.kickoff.getTime());
  const ids = {};
  gs.forEach(g => { ids[g.id] = true; });
  const pickRows = readPicks_().vals.slice(1).concat(archivedPickRows_())
    .filter(r => ids[String(r[P.GameID])])
    .map(r => [String(r[P.GameID]), String(r[P.Player]).toUpperCase(), String(r[P.Pick]).toUpperCase(), String(r[P.Result]),
      r[P.PickedAt] instanceof Date ? r[P.PickedAt].toISOString() : '']);
  return {
    me: me.id,
    season: want,
    seasons: seasons,
    players: players_().map(p => ({ id: p.id, name: p.name, active: p.active, groups: groupsOf_(p) })),
    groups: groupNames_(),
    hidden: (players_().find(p => p.id === me.id) || { hidden: [] }).hidden,
    games: gs.map(g => ({
      id: g.id, season: g.season, type: g.seasonType, week: g.week, kickoff: g.kickoff.toISOString(),
      away: g.awayAbbr, home: g.homeAbbr, awayName: g.away, homeName: g.home,
      awayScore: g.awayScore, homeScore: g.homeScore, status: g.status, winner: g.winner,
      awayPre: g.awayRecordPre, homePre: g.homeRecordPre, awayLogo: g.awayLogo, homeLogo: g.homeLogo,
    })),
    picks: pickRows,
  };
}

/** Ranked by season win % (late joiners aren't penalized), then total wins. */
function leaderboard_(vals, parts, roster) {
  const t = {};
  roster.forEach(p => { t[p.id] = { id: p.id, name: p.name, wW: 0, wL: 0, sW: 0, sL: 0 }; });
  vals.slice(1).forEach(r => {
    const res = String(r[P.Result]);
    if (res !== 'W' && res !== 'L') return;
    const x = t[String(r[P.Player]).toUpperCase()];
    if (!x || String(r[P.Season]) !== parts[0]) return;
    x['s' + res]++;
    if (String(r[P.SeasonType]) === parts[1] && String(r[P.Week]) === parts[2]) x['w' + res]++;
  });
  const rows = Object.keys(t).map(k => {
    const x = t[k], n = x.sW + x.sL;
    return { id: x.id, name: x.name, week: x.wW + '-' + x.wL, season: x.sW + '-' + x.sL, wins: x.sW, pct: n ? x.sW / n : -1 };
  });
  rows.sort((a, b) => b.pct - a.pct || b.wins - a.wins || a.name.localeCompare(b.name));
  let rank = 0, prev = null;
  rows.forEach((r, i) => {
    const k = r.pct + '|' + r.wins;
    if (k !== prev) { rank = i + 1; prev = k; }
    r.rank = r.pct < 0 ? '' : rank;
    r.pctText = r.pct < 0 ? '—' : r.pct.toFixed(3).replace(/^0/, '');
  });
  return rows;
}

// ======================================================================
// ESPN feed
// ======================================================================

function feedUrls_(season, seasonType, week) {
  const q = season ? '?dates=' + season + '&seasontype=' + seasonType + '&week=' + week : '';
  const cdnPath = season ? '/_/week/' + week + '/year/' + season + '/seasontype/' + seasonType : '';
  return [
    { url: 'https://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard' + q },
    { url: 'https://site.web.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard' + q },
    { url: 'https://cdn.espn.com/core/nfl/scoreboard' + cdnPath + '?xhr=1', cdn: true },
  ];
}

const FETCH_OPTS = {
  muteHttpExceptions: true,
  followRedirects: true,
  headers: {
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'en-US,en;q=0.9',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36',
  },
};

/** Tries each ESPN address in turn and uses the first one that answers. */
function fetchScoreboard_(season, seasonType, week) {
  const errors = [];
  for (const f of feedUrls_(season, seasonType, week)) {
    try {
      const res = UrlFetchApp.fetch(f.url, FETCH_OPTS);
      if (res.getResponseCode() !== 200) { errors.push(res.getResponseCode() + ' ' + f.url); continue; }
      let data = JSON.parse(res.getContentText());
      if (f.cdn) data = (data.content && data.content.sbData) || {};
      if (!data.events) { errors.push('no games in reply ' + f.url); continue; }
      return parseGames_(data);
    } catch (e) {
      errors.push(e.message + ' ' + f.url);
    }
  }
  throw new Error('Could not reach any ESPN feed:\n' + errors.join('\n'));
}

/** Diagnostic: reports which data sources Google's servers can reach. Changes nothing. */
function testFeeds() {
  const urls = feedUrls_().map(f => f.url).concat([
    'https://www.thesportsdb.com/api/v1/json/3/eventsseason.php?id=4391&s=2026',
  ]);
  urls.forEach(u => {
    try {
      const r = UrlFetchApp.fetch(u, FETCH_OPTS);
      Logger.log(r.getResponseCode() + '  ' + u + '  (' + r.getContentText().length + ' chars)');
    } catch (e) {
      Logger.log('ERR  ' + u + '  ' + e.message);
    }
  });
}

function parseGames_(data) {
  const out = [];
  (data.events || []).forEach(ev => {
    const comp = (ev.competitions || [])[0];
    if (!comp || !comp.competitors) return;
    const home = comp.competitors.find(c => c.homeAway === 'home');
    const away = comp.competitors.find(c => c.homeAway === 'away');
    if (!home || !away || !home.team || !away.team) return;

    const seasonType = Number((ev.season && ev.season.type) || (data.season && data.season.type));
    if (seasonType !== 2 && seasonType !== 3) return;                         // skip preseason
    if (['AFC', 'NFC'].indexOf(home.team.abbreviation) >= 0) return;        // skip Pro Bowl

    const st = ((comp.status || ev.status || {}).type) || {};
    const completed = !!st.completed;
    const state = completed ? 'final' : (st.state || 'pre');
    let winner = '';
    if (completed) {
      if (home.winner) winner = home.team.abbreviation;
      else if (away.winner) winner = away.team.abbreviation;
      else if (String(home.score) === String(away.score)) winner = 'TIE';
    }
    const score = c => (state === 'pre' ? '' : Number(c.score || 0));
    const record = c => {
      const r = (c.records || []).find(x => x.type === 'total' || x.name === 'overall' || x.name === 'All Splits');
      return (r && r.summary) || '';
    };
    out.push({
      id: String(ev.id),
      season: Number((ev.season && ev.season.year) || (data.season && data.season.year)),
      seasonType: seasonType,
      week: Number((ev.week && ev.week.number) || (data.week && data.week.number)),
      kickoff: new Date(comp.date || ev.date),
      away: away.team.shortDisplayName || away.team.displayName,
      awayAbbr: away.team.abbreviation,
      awayLogo: away.team.logo || '',
      awayScore: score(away),
      home: home.team.shortDisplayName || home.team.displayName,
      homeAbbr: home.team.abbreviation,
      homeLogo: home.team.logo || '',
      homeScore: score(home),
      status: state,
      detail: st.shortDetail || st.detail || '',
      winner: winner,
      awayRecord: record(away),
      homeRecord: record(home),
    });
  });
  return out;
}

// ======================================================================
// Sheet helpers
// ======================================================================

let SS_ = null;
function ss_() {
  if (SS_) return SS_;
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  SS_ = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  return SS_;
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Missing tab "' + name + '". Run setup first.');
  return sh;
}

function ensureSheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers && sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

let SETTINGS_ = null;
function settings_() {
  if (SETTINGS_) return SETTINGS_;
  const vals = sheet_(TABS.SETTINGS).getDataRange().getValues();
  const o = {};
  vals.slice(1).forEach(r => { if (r[0] !== '') o[String(r[0]).trim()] = String(r[1]).trim(); });
  SETTINGS_ = o;
  return o;
}

function setSetting_(name, value) {
  SETTINGS_ = null;
  const sh = sheet_(TABS.SETTINGS);
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === name) { sh.getRange(i + 1, 2).setNumberFormat('@').setValue(value); return; }
  }
  sh.appendRow([name, value]);
}

function newKey_() { return Utilities.getUuid().replace(/-/g, '').toLowerCase(); }

let PLAYERS_ = null;
function players_() {
  if (PLAYERS_) return PLAYERS_;
  const vals = sheet_(TABS.PLAYERS).getDataRange().getValues();
  PLAYERS_ = vals.slice(1)
    .filter(r => String(r[0]).trim() && String(r[1]).trim())
    .map(r => ({
      id: String(r[0]).trim().toUpperCase(),
      name: String(r[1]).trim(),
      key: String(r[2]).trim(),
      active: isOn_(r[4]),
      remind: isOn_(r[5]),
      newweek: isOn_(r[6]),
      results: isOn_(r[7]),
      evening: isOn_(r[8]),
      groups: splitList_(r[9]),
      hidden: splitList_(r[10]).map(x => x.toUpperCase()),
      upsets: isOn_(r[11]),
    }));
  return PLAYERS_;
}

function playerForKey_(key) {
  if (!key || !/^[a-f0-9]{32}$/.test(key)) return null;
  const p = players_().find(x => x.key === key && x.active);
  if (!p) return null;
  return { id: p.id, name: p.name, isAdmin: (settings_().AdminPlayer || 'P1').toUpperCase() === p.id };
}

function weekKey_(g) { return [g.season, g.seasonType, g.week].join('|'); }

function weekLabel_(type, week) {
  if (Number(type) === 3) {
    return ({ 1: 'Wild Card', 2: 'Divisional', 3: 'Conference Championships', 5: 'Super Bowl' })[Number(week)] || ('Postseason Week ' + week);
  }
  return 'Week ' + week;
}

function rec_(t) { return t.W + '-' + t.L; }

function readGames_() {
  const vals = sheet_(TABS.GAMES).getDataRange().getValues();
  return vals.slice(1).filter(r => r[G.GameID] !== '').map(r => ({
    id: String(r[G.GameID]),
    season: Number(r[G.Season]),
    seasonType: Number(r[G.SeasonType]),
    week: Number(r[G.Week]),
    kickoff: r[G.Kickoff] instanceof Date ? r[G.Kickoff] : new Date(r[G.Kickoff]),
    away: String(r[G.Away]), awayAbbr: String(r[G.AwayAbbr]), awayLogo: String(r[G.AwayLogo]), awayScore: r[G.AwayScore],
    home: String(r[G.Home]), homeAbbr: String(r[G.HomeAbbr]), homeLogo: String(r[G.HomeLogo]), homeScore: r[G.HomeScore],
    status: String(r[G.Status]),
    detail: String(r[G.Detail]),
    winner: String(r[G.Winner]),
    awayRecord: rec2_(r[G.AwayRecord]),
    homeRecord: rec2_(r[G.HomeRecord]),
    awayRecordPre: rec2_(r[G.AwayRecordPre]),
    homeRecordPre: rec2_(r[G.HomeRecordPre]),
  }));
}

function upsertGames_(games) {
  if (!games.length) return;
  const sh = sheet_(TABS.GAMES);
  const width = GAME_HEADERS.length;
  const vals = sh.getDataRange().getValues().map(r => {
    const row = r.slice(0, width);
    while (row.length < width) row.push('');
    return row;
  });
  vals[0] = GAME_HEADERS.slice();
  const idx = {};
  for (let i = 1; i < vals.length; i++) idx[String(vals[i][G.GameID])] = i;
  const now = new Date();
  games.forEach(g => {
    // Team records going INTO the game: refreshed while the game hasn't started, then frozen.
    const old = idx[g.id] !== undefined ? vals[idx[g.id]] : null;
    const pre = g.status === 'pre';
    const awayPre = pre ? g.awayRecord : (old ? rec2_(old[G.AwayRecordPre]) : '');
    const homePre = pre ? g.homeRecord : (old ? rec2_(old[G.HomeRecordPre]) : '');
    const row = [g.id, g.season, g.seasonType, g.week, g.kickoff, g.away, g.awayAbbr, g.awayLogo, g.awayScore,
      g.home, g.homeAbbr, g.homeLogo, g.homeScore, g.status, g.detail, g.winner, now, g.awayRecord, g.homeRecord,
      awayPre, homePre];
    if (old) vals[idx[g.id]] = row;
    else { idx[g.id] = vals.length; vals.push(row); }
  });
  sh.getRange(1, G.AwayRecord + 1, vals.length, 4).setNumberFormat('@');
  sh.getRange(1, 1, vals.length, width).setValues(vals);
  sh.getRange(1, 1, 1, width).setFontWeight('bold');
  sh.getRange(2, G.Kickoff + 1, vals.length - 1, 1).setNumberFormat('ddd m/d h:mm AM/PM');
}

/** Weeks with games that kicked off 3h–21d ago but aren't final yet (e.g., Monday night after the feed rolls over). */
function staleWeeks_() {
  const now = Date.now();
  const keys = new Set();
  readGames_().forEach(g => {
    const t = g.kickoff.getTime();
    if (g.status !== 'final' && t < now - 3 * 3600e3 && t > now - 21 * 86400e3) keys.add(weekKey_(g));
  });
  return Array.from(keys);
}

function readPicks_() {
  const sh = sheet_(TABS.PICKS);
  const vals = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (r[P.GameID] === '') continue;
    map[String(r[P.GameID]) + '|' + String(r[P.Player]).toUpperCase()] = {
      row: i + 1,
      pick: String(r[P.Pick]).toUpperCase(),
      override: r[P.Override] === true || String(r[P.Override]).toUpperCase() === 'TRUE',
      result: String(r[P.Result]),
    };
  }
  return { sh: sh, vals: vals, map: map };
}

function writePick_(g, playerId, pick, override, picks) {
  picks = picks || readPicks_();
  const hit = picks.map[g.id + '|' + playerId];
  const row = [g.id, g.season, g.seasonType, g.week, playerId, pick, new Date(), override, ''];
  if (hit) picks.sh.getRange(hit.row, 1, 1, row.length).setValues([row]);
  else picks.sh.appendRow(row);
}

/** Grades every pick against final results, then rebuilds the Records tab. */
function gradeAndRecord_() {
  const games = {};
  readGames_().forEach(g => { games[g.id] = g; });
  const picks = readPicks_();
  const vals = picks.vals;
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (r[P.GameID] === '') continue;
    const g = games[String(r[P.GameID])];
    const pick = String(r[P.Pick]).toUpperCase().trim();
    let result = '';
    if (pick === 'MISSED') result = 'L';
    else if (pick && g && g.status === 'final' && g.winner) result = (pick === g.winner) ? 'W' : 'L';
    r[P.Result] = result;
  }
  if (vals.length > 1) {
    picks.sh.getRange(2, P.Result + 1, vals.length - 1, 1).setValues(vals.slice(1).map(r => [r[P.Result]]));
  }
  rebuildRecords_(vals);
}

function rebuildRecords_(pickVals) {
  pickVals = pickVals.concat(archivedPickRows_());   // archived seasons stay on the Records tab
  const players = players_();
  const ids = players.map(p => p.id);
  const weeks = {};
  pickVals.slice(1).forEach(r => {
    const res = String(r[P.Result]);
    const pl = String(r[P.Player]).toUpperCase();
    if ((res !== 'W' && res !== 'L') || ids.indexOf(pl) < 0) return;
    const k = [r[P.Season], r[P.SeasonType], r[P.Week]].join('|');
    const w = weeks[k] || (weeks[k] = { season: Number(r[P.Season]), type: Number(r[P.SeasonType]), week: Number(r[P.Week]), t: {} });
    const x = w.t[pl] || (w.t[pl] = { W: 0, L: 0 });
    x[res]++;
  });

  const sorted = Object.keys(weeks).map(k => weeks[k])
    .sort((a, b) => a.season - b.season || a.type - b.type || a.week - b.week);
  const rows = [['Season', 'Week'].concat(players.map(p => p.name))];
  const totalRows = [];
  let tot = null;
  const flush = () => {
    if (!tot) return;
    totalRows.push(rows.length + 1);
    rows.push([tot.season, 'SEASON TOTAL'].concat(ids.map(id => (tot.t[id] ? rec_(tot.t[id]) : ''))));
  };
  sorted.forEach(w => {
    if (tot && tot.season !== w.season) { flush(); tot = null; }
    if (!tot) tot = { season: w.season, t: {} };
    rows.push([w.season, weekLabel_(w.type, w.week)].concat(ids.map(id => (w.t[id] ? rec_(w.t[id]) : ''))));
    ids.forEach(id => {
      if (!w.t[id]) return;
      const x = tot.t[id] || (tot.t[id] = { W: 0, L: 0 });
      x.W += w.t[id].W; x.L += w.t[id].L;
    });
  });
  flush();

  const sh = ensureSheet_(TABS.RECORDS);
  sh.clear();
  const width = rows[0].length;
  if (width > 2) sh.getRange(1, 3, rows.length, width - 2).setNumberFormat('@');
  sh.getRange(1, 1, rows.length, width).setValues(rows);
  sh.getRange(1, 1, 1, width).setFontWeight('bold');
  totalRows.forEach(r => sh.getRange(r, 1, 1, width).setFontWeight('bold').setBackground('#eef2ff'));
  sh.setFrozenRows(1);
}

/** Team record as text; undoes Sheets turning "4-2" into a date. */
function rec2_(v) {
  if (v instanceof Date) return (v.getMonth() + 1) + '-' + v.getDate();
  return String(v || '');
}

function isOn_(v) { return !(v === false || String(v === undefined ? '' : v).trim().toUpperCase() === 'FALSE'); }

// ======================================================================
// Notifications (OneSignal)
// ======================================================================

const PREF_COLS = { remind: 6, newweek: 7, results: 8, evening: 9, upsets: 12 };   // column numbers on the Players tab

function prefsFor_(id) {
  const p = players_().find(x => x.id === id) || {};
  return { remind: p.remind !== false, newweek: p.newweek !== false, results: p.results !== false, evening: p.evening !== false, upsets: p.upsets !== false };
}

function setPref_(key, type, on) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  if (!PREF_COLS[type]) throw new Error('Unknown setting.');
  const sh = playersSheet_();
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() === me.id) {
      sh.getRange(i + 1, PREF_COLS[type]).setValue(String(on) === '1' || on === true);
      break;
    }
  }
  PLAYERS_ = null;
  return prefsFor_(me.id);
}

/** Sends one notification to the given players (by player ID). Returns OneSignal's reply text. */
function sendPush_(playerIds, title, body) {
  const key = String(settings_().OneSignalKey || '').trim();
  if (!key) throw new Error('Put your OneSignal REST API key in OneSignalKey on the Settings tab.');
  if (!playerIds.length) return 'nobody to send to';
  const app = String(settings_().AppURL || '').trim();
  const payload = {
    app_id: ONESIGNAL_APP_ID,
    target_channel: 'push',
    include_aliases: { external_id: playerIds.map(id => 'gp_' + id) },
    headings: { en: title },
    contents: { en: body },
    priority: 10,                    // deliver right away instead of when the phone next wakes
  };
  if (app) payload.url = app;
  const res = UrlFetchApp.fetch('https://api.onesignal.com/notifications?c=push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: (key.indexOf('os_v2_') === 0 ? 'Key ' : 'Basic ') + key },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const text = res.getResponseCode() + ' ' + res.getContentText().slice(0, 300);
  logNotice_(playerIds.join(','), title + ' — ' + body, text);
  return text;
}

function logNotice_(to, message, result) {
  const sh = ensureSheet_('Notifications', ['Sent', 'To', 'Message', 'OneSignal reply']);
  sh.appendRow([new Date(), to, message, result]);
  const extra = sh.getLastRow() - 501;                      // keep the log to ~500 rows
  if (extra > 0) sh.deleteRows(2, extra);
}

/** App: anyone can send a test notification to their own phone(s), at most once a minute. */
function testMe_(key) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  const cache = CacheService.getScriptCache();
  if (cache.get('testme_' + me.id)) throw new Error('You just sent one. Give it a minute.');
  cache.put('testme_' + me.id, '1', 60);
  const res = sendPush_([me.id], 'Greco Picks', 'Test notification: notifications are working on this phone! 🏈');
  if (res.indexOf('"id":""') >= 0) throw new Error('No phone is signed up for your notifications yet. Tap Turn on first.');
  return { sent: true };
}

/**
 * App: "Report a problem". Saves to the Issues tab, emails you (the script owner),
 * and sends you a notification. Max 5 reports per person per hour.
 */
function report_(key, text, info) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  text = String(text || '').trim().slice(0, 2000);
  if (!text) throw new Error('Type what happened first.');
  info = String(info || '').slice(0, 1500);
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('rep_' + me.id) || 0);
  if (n >= 5) throw new Error('Thanks, we already got your reports. Try again in a bit.');
  cache.put('rep_' + me.id, String(n + 1), 3600);
  const sh = ensureSheet_('Issues', ['Reported', 'From', 'What happened', 'Phone / app details', 'Status']);
  sh.appendRow([new Date(), me.name + ' (' + me.id + ')', text, info, 'New']);
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Greco Picks: problem reported by ' + me.name,
      me.name + ' wrote:\n\n' + text + '\n\n---\n' + info + '\n\n(Also saved on the Issues tab of the Sheet.)');
  } catch (e) { Logger.log('Report email failed: ' + e.message); }
  try {
    const admin = (settings_().AdminPlayer || 'P1').toUpperCase();
    if (admin !== me.id && String(settings_().OneSignalKey || '').trim()) sendPush_([admin], 'Problem reported', me.name + ': ' + text.slice(0, 120));
  } catch (e) { Logger.log('Report push failed: ' + e.message); }
  return { sent: true };
}

// ======================================================================
// Names and people (app)
// ======================================================================

/** Players tab, with the header row brought up to date (adds new columns like Evening). */
function playersSheet_() {
  const sh = sheet_(TABS.PLAYERS);
  const head = sh.getRange(1, 1, 1, PLAYER_HEADERS.length).getValues()[0];
  if (head.join('|') !== PLAYER_HEADERS.join('|')) sh.getRange(1, 1, 1, PLAYER_HEADERS.length).setValues([PLAYER_HEADERS]).setFontWeight('bold');
  return sh;
}

function cleanName_(name) {
  name = String(name || '').replace(/\s+/g, ' ').trim();
  if (!name) throw new Error('Type a name first.');
  if (name.length > 20) throw new Error('Keep names to 20 characters or less.');
  if (/[<>]/.test(name)) throw new Error('Names can\'t use < or >.');
  return name;
}

function nameTaken_(name, exceptId) {
  const n = name.toLowerCase();
  return players_().some(p => p.id !== exceptId && p.name.toLowerCase() === n);
}

function appBase_() {
  const s = settings_();
  const app = String(s.AppURL || '').trim().replace(/[?#].*$/, '');
  if (app) return app.endsWith('/') ? app : app + '/';
  const web = String(s.WebAppURL || '').trim().replace(/\?.*$/, '');
  return web;
}

/** Anyone: change your own name. History stays with you (everything is stored by player ID). */
function rename_(key, name) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  name = cleanName_(name);
  if (nameTaken_(name, me.id)) throw new Error('Someone already uses that name.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = playersSheet_();
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim().toUpperCase() === me.id) { sh.getRange(i + 1, 2).setValue(name); break; }
    }
    PLAYERS_ = null;
    rebuildRecords_(readPicks_().vals);   // Records tab picks up the new name right away
  } finally {
    lock.releaseLock();
  }
  return getState(key, '');
}

function requireAdmin_(key) {
  const me = playerForKey_(key);
  if (!me || !me.isAdmin) throw new Error('Not allowed.');
  return me;
}

/** Admin: everyone with their personal link. */
function people_(key) {
  requireAdmin_(key);
  const base = appBase_();
  return {
    base: base,
    groups: groupNames_(),
    adminId: adminId_(),
    people: players_().map(p => ({ id: p.id, name: p.name, active: p.active, groups: p.groups, link: base && p.key ? base + '?k=' + p.key : '' })),
  };
}

/** Admin: add a person and make their link. */
function addPerson_(key, name) {
  requireAdmin_(key);
  name = cleanName_(name);
  if (nameTaken_(name, '')) throw new Error('Someone already uses that name.');
  if (!appBase_()) throw new Error('Put the app address in AppURL on the Settings tab first.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = playersSheet_();
    let maxId = 0;
    players_().forEach(p => { const m = /^P(\d+)$/.exec(p.id); if (m) maxId = Math.max(maxId, Number(m[1])); });
    const id = 'P' + (maxId + 1), k = newKey_();
    sh.getRange('C:D').setNumberFormat('@');
    sh.appendRow([id, name, k, appBase_() + '?k=' + k, true]);
    PLAYERS_ = null;
  } finally {
    lock.releaseLock();
  }
  return people_(key);
}

/** Admin: give someone a new link. Their old link stops working; their picks and history stay. */
function newLink_(key, playerId) {
  const me = requireAdmin_(key);
  playerId = String(playerId || '').toUpperCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let k = '';
  try {
    const sh = playersSheet_();
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim().toUpperCase() === playerId) {
        k = newKey_();
        sh.getRange(i + 1, 3, 1, 2).setNumberFormat('@').setValues([[k, appBase_() + '?k=' + k]]);
        break;
      }
    }
    if (!k) throw new Error('Unknown player.');
    PLAYERS_ = null;
  } finally {
    lock.releaseLock();
  }
  const out = people_(playerId === me.id ? k : key);
  if (playerId === me.id) out.myNewKey = k;   // the app switches itself over to your new link
  return out;
}

// ======================================================================
// Groups and hiding
// ======================================================================

function splitList_(v) {
  return String(v === undefined || v === null ? '' : v).split(',').map(x => x.trim()).filter(Boolean);
}

/** "Greco" -> "Greco group" (left alone if the name already ends in "group"). */
function groupPhrase_(g) { return /group$/i.test(g) ? g : g + ' group'; }

function adminId_() { return (settings_().AdminPlayer || 'P1').toUpperCase(); }

/** Every group name in use, in the order they first appear on the Players tab. */
function groupNames_() {
  const out = [];
  players_().forEach(p => p.groups.forEach(g => { if (out.map(x => x.toLowerCase()).indexOf(g.toLowerCase()) < 0) out.push(g); }));
  return out;
}

/** Groups a player belongs to. The admin is in all of them. */
function groupsOf_(p) { return p.id === adminId_() ? groupNames_() : p.groups; }

function inGroup_(p, group) {
  if (!group) return true;
  return groupsOf_(p).some(g => g.toLowerCase() === group.toLowerCase());
}

/** Admin: set someone's groups (comma separated; blank = no group). */
function setGroups_(key, playerId, groups) {
  requireAdmin_(key);
  playerId = String(playerId || '').toUpperCase();
  const list = [];
  splitList_(groups).forEach(g => {
    g = g.replace(/\s+/g, ' ').slice(0, 20);
    if (/[<>]/.test(g)) throw new Error('Group names can\'t use < or >.');
    if (list.map(x => x.toLowerCase()).indexOf(g.toLowerCase()) < 0) list.push(g);
  });
  // Reuse the existing spelling of a group if it differs only by capitals.
  const known = groupNames_();
  const fixed = list.map(g => known.find(k => k.toLowerCase() === g.toLowerCase()) || g);
  const sh = playersSheet_();
  const vals = sh.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() === playerId) { sh.getRange(i + 1, 10).setValue(fixed.join(', ')); found = true; break; }
  }
  if (!found) throw new Error('Unknown player.');
  PLAYERS_ = null;
  return people_(key);
}

/**
 * Admin: retire someone who stopped playing (or bring them back). Retired people disappear from
 * everyone's Picks/Group/leaderboard and their link stops working; their past picks stay in Stats.
 */
function setActive_(key, playerId, on) {
  const me = requireAdmin_(key);
  playerId = String(playerId || '').toUpperCase();
  if (playerId === me.id) throw new Error('You can\'t retire yourself.');
  const sh = playersSheet_();
  const vals = sh.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() === playerId) { sh.getRange(i + 1, 5).setValue(String(on) === '1' || on === true); found = true; break; }
  }
  if (!found) throw new Error('Unknown player.');
  PLAYERS_ = null;
  return people_(key);
}

/** Anyone: hide or show someone on your own screens. */
function setHidden_(key, playerId, hide) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  playerId = String(playerId || '').toUpperCase();
  if (playerId === me.id) throw new Error('You can\'t hide yourself.');
  const sh = playersSheet_();
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() !== me.id) continue;
    const cur = splitList_(vals[i][10]).map(x => x.toUpperCase()).filter(x => x !== playerId);
    if (String(hide) === '1' || hide === true) cur.push(playerId);
    sh.getRange(i + 1, 11).setValue(cur.join(','));
    break;
  }
  PLAYERS_ = null;
  return { hidden: (players_().find(p => p.id === me.id) || { hidden: [] }).hidden };
}

/** 👑/💩 for everyone ('') and for each group. */
function crowns_(games, pickVals, roster) {
  const out = { '': crown_(games, pickVals, roster) };
  groupNames_().forEach(g => { out[g] = crown_(games, pickVals, roster.filter(p => inGroup_(p, g))); });
  return out;
}

/** Winner(s) 👑 and last place 💩 of the most recent week where every game is final. */
function crown_(games, pickVals, roster) {
  const byWeek = {};
  games.forEach(g => { (byWeek[weekKey_(g)] = byWeek[weekKey_(g)] || []).push(g); });
  const done = Object.keys(byWeek).filter(k => byWeek[k].every(g => g.status === 'final'))
    .sort((a, b) => { const x = a.split('|').map(Number), y = b.split('|').map(Number); return y[0] - x[0] || y[1] - x[1] || y[2] - x[2]; });
  if (!done.length) return null;
  const parts = done[0].split('|');
  const board = leaderboard_(pickVals, parts, roster);
  const w = board.map(b => ({ id: b.id, name: b.name, rec: b.week, wins: Number(b.week.split('-')[0]) }));
  const top = Math.max.apply(null, w.map(x => x.wins).concat([0]));
  if (!top) return null;
  let winners = w.filter(x => x.wins === top);
  let tb = null;
  if (winners.length > 1) {
    // Monday-night tiebreaker: closest guess at the week's last game's total points takes the crown.
    const wkGames = byWeek[done[0]].slice().sort((a, b) => b.kickoff - a.kickoff || b.id.localeCompare(a.id));
    const last = wkGames[0];
    const total = Number(last.awayScore || 0) + Number(last.homeScore || 0);
    const guesses = readTiebreaks_();
    const diff = x => { const v = guesses[done[0] + '|' + x.id]; return v === undefined ? Infinity : Math.abs(v - total); };
    const best = Math.min.apply(null, winners.map(diff));
    if (best !== Infinity) {
      const kept = winners.filter(x => diff(x) === best);
      if (kept.length < winners.length) tb = { total: total, game: last.awayAbbr + ' @ ' + last.homeAbbr, guess: guesses[done[0] + '|' + kept[0].id] };
      winners = kept;
    }
  }
  // 💩 fewest wins that week, among people who made at least one pick. None if everyone tied.
  const played = w.filter(x => x.rec !== '0-0');
  const low = Math.min.apply(null, played.map(x => x.wins));
  const losers = played.length > 1 && low < top ? played.filter(x => x.wins === low) : [];
  return {
    label: weekLabel_(parts[1], parts[2]), key: done[0], ids: winners.map(x => x.id), names: winners.map(x => x.name), rec: winners[0].rec, tb: tb,
    poop: losers.length ? { ids: losers.map(x => x.id), names: losers.map(x => x.name), rec: losers[0].rec } : null,
  };
}

// ======================================================================
// Tiebreaker, reactions, last week, milestones, archive
// ======================================================================

let TB_ = null;
function readTiebreaks_() {
  if (TB_) return TB_;
  const sh = ensureSheet_(TABS.TIEBREAK, ['Week', 'Player', 'Guess', 'At']);
  sh.getRange('A:A').setNumberFormat('@');
  const o = {};
  sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0] !== '') o[String(r[0]) + '|' + String(r[1]).toUpperCase()] = Number(r[2]); });
  TB_ = o;
  return o;
}

function tieGame_(weekGames) {
  return weekGames.slice().sort((a, b) => b.kickoff - a.kickoff || b.id.localeCompare(a.id))[0] || null;
}

/** The selected week's tiebreaker: your guess, and everyone's once that game kicks off. */
function tiebreakView_(weekGames, wk, me, roster, now) {
  const g = tieGame_(weekGames);
  if (!g) return null;
  const tb = readTiebreaks_();
  const locked = now >= g.kickoff.getTime();
  const mine = tb[wk + '|' + me.id];
  return {
    gameId: g.id, week: wk, locked: locked,
    mine: mine === undefined ? null : mine,
    total: g.status === 'final' ? Number(g.awayScore || 0) + Number(g.homeScore || 0) : null,
    count: roster.filter(p => tb[wk + '|' + p.id] !== undefined).length,
    all: locked ? roster.filter(p => tb[wk + '|' + p.id] !== undefined).map(p => ({ id: p.id, name: p.name, guess: tb[wk + '|' + p.id] })) : [],
  };
}

function setTiebreak_(key, wk, guess) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  const n = Math.round(Number(guess));
  if (!isFinite(n) || n < 0 || n > 200) throw new Error('Enter total points from 0 to 200.');
  const weekGames = readGames_().filter(g => weekKey_(g) === String(wk));
  const g = tieGame_(weekGames);
  if (!g) throw new Error('Week not found.');
  if (Date.now() >= g.kickoff.getTime()) throw new Error('The tiebreaker game has kicked off; guesses are locked.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(TABS.TIEBREAK, ['Week', 'Player', 'Guess', 'At']);
    sh.getRange('A:A').setNumberFormat('@');
    const vals = sh.getDataRange().getValues();
    let row = 0;
    for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === String(wk) && String(vals[i][1]).toUpperCase() === me.id) { row = i + 1; break; }
    if (row) sh.getRange(row, 3, 1, 2).setValues([[n, new Date()]]);
    else sh.appendRow([String(wk), me.id, n, new Date()]);
    TB_ = null;
  } finally {
    lock.releaseLock();
  }
  return tiebreakView_(weekGames, String(wk), me, players_().filter(p => p.active), Date.now());
}

let REACT_ = null;
function readReactions_() {
  if (REACT_) return REACT_;
  const sh = ensureSheet_(TABS.REACT, ['GameID', 'Player', 'Emoji', 'At']);
  sh.getRange('A:A').setNumberFormat('@');
  const o = {};
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (r[0] === '') return;
    const g = o[String(r[0])] || (o[String(r[0])] = {});
    (g[String(r[2])] = g[String(r[2])] || []).push(String(r[1]).toUpperCase());
  });
  REACT_ = o;
  return o;
}

/** Toggle your emoji reaction on a game (after kickoff). */
function react_(key, gameId, emoji) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  emoji = String(emoji || '');
  if (REACTIONS.indexOf(emoji) < 0) throw new Error('Unknown reaction.');
  const g = readGames_().find(x => x.id === String(gameId));
  if (!g) throw new Error('Game not found.');
  if (Date.now() < g.kickoff.getTime()) throw new Error('Reactions open at kickoff.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureSheet_(TABS.REACT, ['GameID', 'Player', 'Emoji', 'At']);
    sh.getRange('A:A').setNumberFormat('@');
    const vals = sh.getDataRange().getValues();
    let row = 0;
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === g.id && String(vals[i][1]).toUpperCase() === me.id && String(vals[i][2]) === emoji) { row = i + 1; break; }
    }
    if (row) sh.deleteRow(row); else sh.appendRow([g.id, me.id, emoji, new Date()]);
    REACT_ = null;
  } finally {
    lock.releaseLock();
  }
  return { gameId: g.id, react: readReactions_()[g.id] || {} };
}

/** Your record in the most recent fully final week (for confetti and the weekly recap). */
function lastWeekFor_(games, pickVals, me, roster) {
  const byWeek = {};
  games.forEach(g => { (byWeek[weekKey_(g)] = byWeek[weekKey_(g)] || []).push(g); });
  const done = Object.keys(byWeek).filter(k => byWeek[k].every(g => g.status === 'final'))
    .sort((a, b) => { const x = a.split('|').map(Number), y = b.split('|').map(Number); return y[0] - x[0] || y[1] - x[1] || y[2] - x[2]; });
  if (!done.length) return null;
  const parts = done[0].split('|');
  const b = leaderboard_(pickVals, parts, roster).find(x => x.id === me.id);
  const rec = b ? b.week : '0-0';
  const [w, l] = rec.split('-').map(Number);
  const n = byWeek[done[0]].length;
  return { key: done[0], label: weekLabel_(parts[1], parts[2]), rec: rec, games: n, perfect: l === 0 && w === n && n > 0 };
}

const ARCH_HEAD = { games: GAME_HEADERS, picks: PICK_HEADERS };
function readArchivedGames_() {
  const sh = ss_().getSheetByName(TABS.GAMES_ARCH);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getDataRange().getValues().slice(1).filter(r => r[G.GameID] !== '').map(r => ({
    id: String(r[G.GameID]), season: Number(r[G.Season]), seasonType: Number(r[G.SeasonType]), week: Number(r[G.Week]),
    kickoff: r[G.Kickoff] instanceof Date ? r[G.Kickoff] : new Date(r[G.Kickoff]),
    away: String(r[G.Away]), awayAbbr: String(r[G.AwayAbbr]), awayLogo: String(r[G.AwayLogo]), awayScore: r[G.AwayScore],
    home: String(r[G.Home]), homeAbbr: String(r[G.HomeAbbr]), homeLogo: String(r[G.HomeLogo]), homeScore: r[G.HomeScore],
    status: String(r[G.Status]), detail: String(r[G.Detail]), winner: String(r[G.Winner]),
    awayRecord: rec2_(r[G.AwayRecord]), homeRecord: rec2_(r[G.HomeRecord]),
    awayRecordPre: rec2_(r[G.AwayRecordPre]), homeRecordPre: rec2_(r[G.HomeRecordPre]),
  }));
}
function archivedPickRows_() {
  const sh = ss_().getSheetByName(TABS.PICKS_ARCH);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getDataRange().getValues().slice(1).filter(r => r[P.GameID] !== '');
}

/**
 * Each summer (June–August), moves finished older seasons from Games/Picks to the archive tabs so the app stays fast.
 * Stats, Wrapped and the Records tab still include them.
 */
function archiveOldSeasons_() {
  const now = new Date();
  const month = Number(Utilities.formatDate(now, TZ, 'M')), year = Number(Utilities.formatDate(now, TZ, 'yyyy'));
  if (month < 6 || month > 8) return;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ARCHIVED_YEAR') === String(year)) return;
  const gsh = sheet_(TABS.GAMES), psh = sheet_(TABS.PICKS);
  const gv = gsh.getDataRange().getValues(), pv = psh.getDataRange().getValues();
  const old = new Set();
  gv.slice(1).forEach(r => { if (r[G.GameID] !== '' && Number(r[G.Season]) < year && String(r[G.Status]) === 'final') old.add(String(r[G.GameID])); });
  if (old.size) {
    const ga = ensureSheet_(TABS.GAMES_ARCH, GAME_HEADERS), pa = ensureSheet_(TABS.PICKS_ARCH, PICK_HEADERS);
    ga.getRange('A:A').setNumberFormat('@'); pa.getRange('A:A').setNumberFormat('@');
    const moveG = gv.slice(1).filter(r => old.has(String(r[G.GameID]))).map(r => { const x = r.slice(0, GAME_HEADERS.length); while (x.length < GAME_HEADERS.length) x.push(''); return x; });
    const keepG = gv.slice(1).filter(r => !old.has(String(r[G.GameID])));
    const moveP = pv.slice(1).filter(r => old.has(String(r[P.GameID]))).map(r => r.slice(0, PICK_HEADERS.length));
    const keepP = pv.slice(1).filter(r => !old.has(String(r[P.GameID])));
    if (moveG.length) ga.getRange(ga.getLastRow() + 1, 1, moveG.length, GAME_HEADERS.length).setValues(moveG);
    if (moveP.length) pa.getRange(pa.getLastRow() + 1, 1, moveP.length, PICK_HEADERS.length).setValues(moveP);
    // rewrite the live tabs with only what's left
    gsh.getRange(2, 1, Math.max(gv.length - 1, 1), gv[0].length).clearContent();
    if (keepG.length) gsh.getRange(2, 1, keepG.length, gv[0].length).setValues(keepG);
    psh.getRange(2, 1, Math.max(pv.length - 1, 1), pv[0].length).clearContent();
    if (keepP.length) psh.getRange(2, 1, keepP.length, pv[0].length).setValues(keepP);
    Logger.log('Archived ' + moveG.length + ' games and ' + moveP.length + ' picks.');
  }
  props.setProperty('ARCHIVED_YEAR', String(year));
}

/** Milestones: 100th correct pick, first perfect week, 10/25/50 upsets in a season. Each sent once, ever. */
function milestones_(games, picks, roster, now) {
  const props = PropertiesService.getScriptProperties();
  let got = [];
  try { got = JSON.parse(props.getProperty('MILESTONES') || '[]'); } catch (e) { got = []; }
  const first = props.getProperty('MILESTONES_INIT') === null;   // first run: remember what's already earned, send nothing
  const gmap = {}; games.concat(readArchivedGames_()).forEach(g => { gmap[g.id] = g; });
  const rows = picks.vals.slice(1).concat(archivedPickRows_());
  const curSeason = Math.max.apply(null, games.map(g => g.season).concat([0]));
  const pct = r => { const m = String(r || '').split('-').map(Number); if (m.length < 2 || m.some(isNaN)) return null; const n = m[0] + m[1] + (m[2] || 0); return n ? (m[0] + (m[2] || 0) / 2) / n : null; };
  const stats = {};
  roster.forEach(p => { stats[p.id] = { W: 0, ups: 0, weeks: {} }; });
  rows.forEach(r => {
    const id = String(r[P.Player]).toUpperCase(), res = String(r[P.Result]), g = gmap[String(r[P.GameID])];
    const st = stats[id]; if (!st || !g || (res !== 'W' && res !== 'L')) return;
    const wk = weekKey_(g); const w = st.weeks[wk] || (st.weeks[wk] = { W: 0, n: 0 }); w.n++;
    if (res === 'W') {
      st.W++; w.W++;
      const pick = String(r[P.Pick]).toUpperCase();
      if (g.season === curSeason && pick !== 'TIE') {
        const a = pct(g.awayRecordPre), h = pct(g.homeRecordPre);
        if (a !== null && h !== null && a !== h && ((pick === g.awayAbbr && a < h) || (pick === g.homeAbbr && h < a))) st.ups++;
      }
    }
  });
  const weekSize = {}; games.concat(readArchivedGames_()).forEach(g => { weekSize[weekKey_(g)] = (weekSize[weekKey_(g)] || 0) + 1; });
  const sends = [];
  roster.forEach(p => {
    const st = stats[p.id];
    [50, 100, 250, 500, 1000].forEach(n => { if (st.W >= n) sends.push([p, 'M|' + p.id + '|W|' + n, n + ' correct picks! 🎯', 'You just hit ' + n + ' correct picks all-time in Greco Picks.']); });
    if (Object.keys(st.weeks).some(k => st.weeks[k].W === weekSize[k] && st.weeks[k].n === weekSize[k] && weekSize[k] >= 4))
      sends.push([p, 'M|' + p.id + '|perfect', 'Your first perfect week! 🎉', 'Every single pick right. Legendary.']);
    [10, 25, 50].forEach(n => { if (st.ups >= n) sends.push([p, 'M|' + p.id + '|U|' + curSeason + '|' + n, n + ' upsets called 🎯', 'You\'ve called ' + n + ' upsets this season.']); });
  });
  let changed = false;
  sends.forEach(([p, key, title, body]) => {
    if (got.indexOf(key) >= 0) return;
    got.push(key); changed = true;
    if (!first && p.results) sendPush_([p.id], title, body);
  });
  if (changed) props.setProperty('MILESTONES', JSON.stringify(got));
  if (first) props.setProperty('MILESTONES_INIT', '1');
}

/**
 * Admin: nudge one person (or 'ALL' = everyone with open picks) to make their picks for the current week.
 * Sent even if they turned pick reminders off (it's a personal nudge). Once per person per 10 minutes.
 */
function nudge_(key, playerId) {
  const me = requireAdmin_(key);
  if (!String(settings_().OneSignalKey || '').trim()) throw new Error('Notifications aren\'t set up (OneSignalKey).');
  const cur = PropertiesService.getScriptProperties().getProperty('CURRENT_WEEK') || '';
  const now = Date.now();
  const open = readGames_().filter(g => weekKey_(g) === cur && g.status === 'pre' && g.kickoff.getTime() > now)
    .sort((a, b) => a.kickoff - b.kickoff);
  if (!open.length) throw new Error('No open games left this week.');
  const picks = readPicks_();
  const cache = CacheService.getScriptCache();
  const parts = cur.split('|');
  const label = weekLabel_(parts[1], parts[2]);
  const targets = String(playerId || '').toUpperCase() === 'ALL'
    ? players_().filter(p => p.active && p.id !== me.id)
    : players_().filter(p => p.id === String(playerId || '').toUpperCase());
  if (!targets.length) throw new Error('Unknown player.');
  const sent = [], skipped = [];
  targets.forEach(p => {
    const mine = open.filter(g => !picks.map[g.id + '|' + p.id]);
    if (!mine.length) { skipped.push(p.name + ' (all picked)'); return; }
    if (cache.get('nudge_' + p.id)) { skipped.push(p.name + ' (nudged recently)'); return; }
    const first = fmtTime_(mine[0].kickoff);
    const what = mine.length === 1 ? mine[0].away + ' @ ' + mine[0].home : mine.length + ' open picks';
    const res = sendPush_([p.id], me.name + ' says: make your picks! 🏈',
      'You have ' + what + ' for ' + label + '. Next lock: ' + first.day + ' ' + first.time + '.');
    if (res.indexOf('"id":""') >= 0) { skipped.push(p.name + ' (notifications off)'); return; }
    cache.put('nudge_' + p.id, '1', 600);
    sent.push(p.name);
  });
  if (targets.length === 1 && !sent.length) throw new Error('Not sent: ' + skipped[0] + '.');
  return { sent: sent, skipped: skipped };
}

/** Tells the admin once, the moment every active player has picked every game of a week. */
function allInCheck_(wk, picks) {
  if (!String(settings_().OneSignalKey || '').trim()) return;
  const key = 'A|' + wk;
  const sent = sentSet_();
  if (sent.indexOf(key) >= 0) return;
  const games = readGames_().filter(g => weekKey_(g) === wk);
  const roster = players_().filter(p => p.active);
  if (!games.length || roster.length < 2) return;
  const done = roster.every(p => games.every(g => picks.map[g.id + '|' + p.id]));
  if (!done) return;
  const parts = wk.split('|');
  sendPush_([adminId_()], 'Everyone\'s in ✅', 'All ' + roster.length + ' players have picked every ' + weekLabel_(parts[1], parts[2]) + ' game.');
  markSent_(sent, key);
}

/** Editor: sends a test notification to you. Turn notifications on in the app first. */
function testPush() {
  const admin = (settings_().AdminPlayer || 'P1').toUpperCase();
  Logger.log(sendPush_([admin], 'Greco Picks', 'Test notification — it works! 🏈'));
}

function sentSet_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('NOTIFIED') || '[]'); } catch (e) { return []; }
}
function markSent_(list, key) {
  list.push(key);
  PropertiesService.getScriptProperties().setProperty('NOTIFIED', JSON.stringify(list.slice(-300)));
}

function fmtTime_(d) {
  const day = Utilities.formatDate(d, TZ, 'EEEE');
  const t = Utilities.formatDate(d, TZ, 'h:mm a').replace(':00', '');
  return { day: day, time: t };
}

/** Called after every refresh: sends whatever is due, once each. */
function notifications_(nowMs) {
  if (!String(settings_().OneSignalKey || '').trim()) return;   // not set up yet
  const now = nowMs || Date.now();
  const sent = sentSet_();
  const games = readGames_();
  const roster = players_().filter(p => p.active);
  const picks = readPicks_();

  // 1) Pick reminders: REMIND_HOURS before a kickoff, only to people with unpicked games, and at most
  //    ONE per person per day (it lists everything they still need that day).
  const dayOf = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const windows = {};
  games.forEach(g => {
    const t = g.kickoff.getTime();
    if (g.status === 'pre' && t > now && t - now <= REMIND_HOURS * 3600e3) (windows[t] = windows[t] || []).push(g);
  });
  Object.keys(windows).sort().forEach(t => {
    const key = 'R|' + t;
    if (sent.indexOf(key) >= 0) return;
    const ws = windows[t];
    const day = dayOf(new Date(Number(t)));
    const when = fmtTime_(new Date(Number(t)));
    roster.filter(p => p.remind).forEach(p => {
      const open = ws.filter(g => !picks.map[g.id + '|' + p.id]);
      if (!open.length) return;
      const dkey = 'RD|' + day + '|' + p.id;
      if (sent.indexOf(dkey) >= 0) return;   // already reminded today
      const today = games.filter(g => g.status === 'pre' && g.kickoff.getTime() > now && dayOf(g.kickoff) === day && !picks.map[g.id + '|' + p.id]);
      const label = today.length === 1 ? today[0].away + ' @ ' + today[0].home : today.length + ' picks today';
      sendPush_([p.id], 'Picks lock at ' + when.time, 'You still need ' + label + '. First one locks at ' + when.time + '.');
      markSent_(sent, dkey);
    });
    markSent_(sent, key);
  });

  // 1b) Night before: at 8 PM Eastern, people with open picks for tomorrow's games get one heads-up.
  const nowEt = new Date(now);
  const hourEt = Number(Utilities.formatDate(nowEt, TZ, 'H'));
  if (hourEt >= EVENING_HOUR) {
    const tomorrow = Utilities.formatDate(new Date(now + 86400e3), TZ, 'yyyy-MM-dd');
    const key = 'E|' + tomorrow;
    if (sent.indexOf(key) < 0) {
      const tg = games.filter(g => g.status === 'pre' && Utilities.formatDate(g.kickoff, TZ, 'yyyy-MM-dd') === tomorrow)
        .sort((a, b) => a.kickoff - b.kickoff);
      if (tg.length) {
        const first = fmtTime_(tg[0].kickoff);
        roster.filter(p => p.evening).forEach(p => {
          const open = tg.filter(g => !picks.map[g.id + '|' + p.id]);
          if (!open.length) return;
          const what = open.length === 1 ? open[0].away + ' @ ' + open[0].home : open.length + ' open picks';
          sendPush_([p.id], 'Picks for ' + first.day, 'You have ' + what + ' for ' + first.day + '. First kickoff is ' + first.time + '.');
        });
      }
      markSent_(sent, key);
    }
  }

  // 2) New week: when the current week changes and it still has games to pick.
  const cur = PropertiesService.getScriptProperties().getProperty('CURRENT_WEEK') || '';
  const lastWeek = PropertiesService.getScriptProperties().getProperty('LAST_NEW_WEEK');
  if (lastWeek === null) {
    PropertiesService.getScriptProperties().setProperty('LAST_NEW_WEEK', cur);   // first run: don't announce the current week
  } else if (cur && cur !== lastWeek) {
    const wkGames = games.filter(g => weekKey_(g) === cur);
    if (wkGames.some(g => g.status === 'pre')) {
      const [, type, week] = cur.split('|');
      const ids = roster.filter(p => p.newweek).map(p => p.id);
      sendPush_(ids, weekLabel_(type, week) + ' is up', 'The ' + weekLabel_(type, week) + ' games are ready. Make your picks!');
    }
    PropertiesService.getScriptProperties().setProperty('LAST_NEW_WEEK', cur);
  }

  // 3) Weekly results: once every game in a week is final (weeks finished in the last 3 days).
  const byWeek = {};
  games.forEach(g => { (byWeek[weekKey_(g)] = byWeek[weekKey_(g)] || []).push(g); });
  Object.keys(byWeek).forEach(wk => {
    const list = byWeek[wk];
    const key = 'X|' + wk;
    if (sent.indexOf(key) >= 0) return;
    if (!list.every(g => g.status === 'final')) return;
    const last = Math.max.apply(null, list.map(g => g.kickoff.getTime()));
    if (now - last > 3 * 86400e3) { markSent_(sent, key); return; }   // too old, skip quietly
    const [season, type, week] = wk.split('|');
    const board = leaderboard_(picks.vals, [season, type, week], roster);
    // Season lead within each person's own (first) group; the admin and people with no group get everyone.
    const leadFor = {};
    const leadLine = (group, p) => {
      if (!(group in leadFor)) {
        const gb = group ? leaderboard_(picks.vals, [season, type, week], roster.filter(x => inGroup_(x, group))) : board;
        const top = gb.filter(b => b.rank === 1);
        leadFor[group] = { top: top, label: group ? 'the ' + groupPhrase_(group) : 'the season' };
      }
      const L = leadFor[group];
      if (!L.top.length) return '';
      if (L.top.length > 1) return 'Tied for the lead in ' + L.label + ' at ' + L.top[0].season + '.';
      if (L.top[0].id === p.id) return 'You lead ' + L.label + ' at ' + L.top[0].season + '!';
      return L.top[0].name + ' leads ' + L.label + ' at ' + L.top[0].season + '.';
    };
    roster.filter(p => p.results).forEach(p => {
      const b = board.find(x => x.id === p.id);
      if (!b || b.week === '0-0') return;
      const group = p.id === adminId_() ? '' : (p.groups[0] || '');
      sendPush_([p.id], weekLabel_(type, week) + ' results', 'You went ' + b.week + '. ' + leadLine(group, p) + ' Your recap is in the app.');
    });
    markSent_(sent, key);
  });

  // 4) Upset roundup: in groups of 3+ (you included), a game where exactly one person picked the winner and 2+
  //    didn't. Sent once per game day, after that day's last game is final, as a single notification.
  const groupNames = groupNames_();
  const days = {};
  games.forEach(g => { const d = dayOf(g.kickoff); (days[d] = days[d] || []).push(g); });
  Object.keys(days).forEach(day => {
    const dg = days[day];
    const key = 'UD|' + day;
    if (sent.indexOf(key) >= 0) return;
    if (!dg.every(g => g.status === 'final')) return;
    const last = Math.max.apply(null, dg.map(g => g.kickoff.getTime()));
    if (now - last > 2 * 86400e3) { markSent_(sent, key); return; }   // too old, skip quietly
    const found = [];   // {game, grp, hero, team}
    dg.filter(g => g.winner).forEach(g => {
      groupNames.forEach(grp => {
        const members = roster.filter(p => inGroup_(p, grp));
        if (members.length < 3) return;
        const pickers = members.filter(p => { const x = picks.map[g.id + '|' + p.id]; return x && x.pick !== 'MISSED'; });
        const right = pickers.filter(p => picks.map[g.id + '|' + p.id].pick === g.winner);
        if (right.length === 1 && pickers.length - 1 >= 2) {
          found.push({ game: g, grp: grp, hero: right[0], team: g.winner === g.awayAbbr ? g.away : g.winner === g.homeAbbr ? g.home : 'the tie' });
        }
      });
    });
    if (found.length) {
      roster.filter(p => p.upsets).forEach(p => {
        const mine = groupsOf_(p);
        const items = [];
        const seen = {};
        found.forEach(f => {
          if (seen[f.game.id] || mine.indexOf(f.grp) < 0) return;   // one line per game, in their first group where it happened
          const grp = mine.find(x => found.some(y => y.game.id === f.game.id && y.grp === x));
          const it = found.find(y => y.game.id === f.game.id && y.grp === grp);
          seen[f.game.id] = 1;
          items.push(it);
        });
        if (!items.length) return;
        const oneGroup = items.every(x => x.grp === items[0].grp);
        const line = x => (x.hero.id === p.id ? 'You' : x.hero.name) + ' took ' + x.team + (oneGroup ? '' : ' (' + groupPhrase_(x.grp) + ')');
        let body;
        if (items.length === 1) {
          const x = items[0];
          body = x.hero.id === p.id
            ? 'You were the only one in the ' + groupPhrase_(x.grp) + ' to pick ' + x.team + '!'
            : 'Only ' + x.hero.name + ' picked ' + x.team + ' in the ' + groupPhrase_(x.grp) + '.';
        } else {
          body = items.length + ' upsets today' + (oneGroup ? ' in the ' + groupPhrase_(items[0].grp) : '') + ': ' + items.map(line).join('; ') + '.';
        }
        sendPush_([p.id], items.length === 1 ? 'Upset alert 👀' : 'Today\'s upsets 👀', body);
      });
    }
    markSent_(sent, key);
  });

  // 6) Milestones
  try { milestones_(games, picks, roster, now); } catch (e) { Logger.log('Milestones: ' + e.message); }

  // 5) Wrapped: once the Super Bowl is final, everyone hears their season recap is ready.
  games.filter(g => g.seasonType === 3 && g.week === 5 && g.status === 'final').forEach(g => {
    const key = 'W|' + g.season;
    if (sent.indexOf(key) >= 0) return;
    if (now - g.kickoff.getTime() < 3 * 86400e3) {
      sendPush_(roster.map(p => p.id), 'Your ' + g.season + ' Greco Picks Wrapped is ready 🎁', 'Open the app and tap Stats to relive your season.');
    }
    markSent_(sent, key);
  });
}
