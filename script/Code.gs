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

const SCRIPT_VERSION = '2.29.3';   // shown in the app (⚙). Use the same number in the deployment description.
const TZ = 'America/New_York';
const TABS = { SETTINGS: 'Settings', PLAYERS: 'Players', GAMES: 'Games', PICKS: 'Picks', RECORDS: 'Records', PREVIEW: 'Preview',
  TIEBREAK: 'Tiebreakers', GAMES_ARCH: 'Games Archive', PICKS_ARCH: 'Picks Archive' };
const PLAYER_HEADERS = ['ID', 'Name', 'Key', 'Link', 'Active', 'Remind', 'NewWeek', 'Results', 'Evening', 'Groups', 'Hidden', 'Upsets', 'LastSeen', 'AppVersion', 'Device', 'Notifications', 'NotifIssue', 'Avatar', 'Team'];
const EVENING_HOUR = 20;   // night-before reminder, 8 PM Eastern
const ONESIGNAL_APP_ID = 'e0fe6864-f2ca-4b41-8c87-00586a7634e4';
const REMIND_HOURS = 2;
const GAME_HEADERS = ['GameID', 'Season', 'SeasonType', 'Week', 'Kickoff', 'Away', 'AwayAbbr', 'AwayLogo', 'AwayScore',
  'Home', 'HomeAbbr', 'HomeLogo', 'HomeScore', 'Status', 'Detail', 'Winner', 'LastUpdated', 'AwayRecord', 'HomeRecord',
  'AwayRecordPre', 'HomeRecordPre'];
const PICK_HEADERS = ['GameID', 'Season', 'SeasonType', 'Week', 'Player', 'Pick', 'PickedAt', 'Override', 'Result', 'Locked'];
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
    const props = PropertiesService.getScriptProperties();
    const current = fetchScoreboard_();
    if (current.length) {
      upsertGames_(current);
      // Only move forward: once we've opened next week early, ESPN's "current" (still last week) doesn't pull us back.
      const espnWeek = weekKey_(current[0]), stored = props.getProperty('CURRENT_WEEK');
      if (!stored || weekOrder_(espnWeek) > weekOrder_(stored)) props.setProperty('CURRENT_WEEK', espnWeek);
    }
    const fetched = new Set(current.map(weekKey_));
    staleWeeks_().filter(k => !fetched.has(k)).slice(0, 3).forEach(k => {
      const [season, type, week] = k.split('|');
      upsertGames_(fetchScoreboard_(season, type, week));
    });
    try { openNextWeekEarly_(); } catch (e) { Logger.log('Next week not loaded yet: ' + e.message); }
    gradeAndRecord_();
    try { sbTick_(); } catch (e) { Logger.log('Super Bowl props: ' + e.message); }
    try { storeChamps_(); } catch (e) { Logger.log('Champs: ' + e.message); }
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

/**
 * Remembers when each person last opened the app, which app version they're on, their device and
 * notification status (Players tab: LastSeen, AppVersion, Device, Notifications). Written at most
 * every 10 minutes per person, or right away if something changed.
 */
function noteSeen_(key, v, d, n) {
  try {
    const me = playerForKey_(key);
    if (!me) return;
    v = String(v || '').slice(0, 20); d = String(d || '').slice(0, 30); n = String(n || '').slice(0, 20);
    const cache = CacheService.getScriptCache();
    const sig = [v, d, n].join('|');
    if (cache.get('seen_' + me.id) === sig) return;
    const sh = playersSheet_();
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim().toUpperCase() !== me.id) continue;
      sh.getRange(i + 1, 13, 1, 4).setValues([[new Date(), v || vals[i][13] || '', d || vals[i][14] || '', n || vals[i][15] || '']]);
      sh.getRange(i + 1, 13).setNumberFormat('m/d h:mm AM/PM');
      sh.getRange(i + 1, 14).setNumberFormat('@');
      break;
    }
    cache.put('seen_' + me.id, sig, 600);
    PLAYERS_ = null;
  } catch (e) { Logger.log('noteSeen: ' + e.message); }
}

/** JSON API used by the installable app: ?api=state|pick|admin&k=KEY&... */
function api_(p) {
  let out;
  try {
    let data;
    if (p.api === 'state') { noteSeen_(p.k, p.v, p.d, p.n); data = getState(p.k, p.wk || ''); }
    else if (p.api === 'pick') data = submitPick(p.k, p.g, p.p);
    else if (p.api === 'admin') data = adminSetPick(p.k, p.g, p.pl, p.p, p.wk || '');
    else if (p.api === 'history') data = history_(p.k, p.season || '');
    else if (p.api === 'prefs') data = setPref_(p.k, p.type, p.on);
    else if (p.api === 'testme') data = testMe_(p.k);
    else if (p.api === 'report') data = report_(p.k, p.text, p.info);
    else if (p.api === 'rename') data = rename_(p.k, p.name);
    else if (p.api === 'team') data = setTeam_(p.k, p.team);
    else if (p.api === 'adminname') data = adminRename_(p.k, p.pl, p.name, p.team);
    else if (p.api === 'people') data = people_(p.k);
    else if (p.api === 'addperson') data = addPerson_(p.k, p.name);
    else if (p.api === 'newlink') data = newLink_(p.k, p.pl);
    else if (p.api === 'groups') data = setGroups_(p.k, p.pl, p.groups);
    else if (p.api === 'hide') data = setHidden_(p.k, p.pl, p.on);
    else if (p.api === 'active') data = setActive_(p.k, p.pl, p.on);
    else if (p.api === 'delete') data = deletePerson_(p.k, p.pl);
    else if (p.api === 'nudge') data = nudge_(p.k, p.pl);
    else if (p.api === 'tiebreak') data = setTiebreak_(p.k, p.wk, p.guess);
    else if (p.api === 'lock') data = lockPicks_(p.k, p.g);
    else if (p.api === 'props') data = setProps_(p.k, p.a);
    else if (p.api === 'sbadmin') data = sbAdmin_(p.k, p.a);
    else if (p.api === 'sbroster') data = sbRoster_(p.k);
    else if (p.api === 'updnotice') data = updateNotice_(p.k);
    else if (p.api === 'pushok') data = pushFixed_(p.k);
    else if (p.api === 'avatars') data = avatars_(p.k);
    else if (p.api === 'avatar') data = setAvatar_(p.k, p.pl, p.av);
    else if (p.api === 'groupav') data = setGroupAv_(p.k, p.group, p.av);
    else if (p.api === 'invite') data = invite_(p.k, p.action, p.group, p.code);
    else if (p.api === 'invites') data = invitesList_(p.k);
    else if (p.api === 'joininfo') data = joinInfo_(p.code);
    else if (p.api === 'join') data = join_(p.code, p.name, p.team);
    else if (p.api === 'joinme') data = joinExisting_(p.k, p.code);
    else if (p.api === 'request') data = requestAdd_(p.k, p.name, p.note);
    else if (p.api === 'resolve') data = resolveRequest_(p.k, p.id, p.action, p.name, p.groups);
    else throw new Error('Unknown request.');
    out = { ok: true, data: data };
  } catch (err) {
    out = { ok: false, error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST version of the API, for requests too big for a web address (profile pictures).
 * The app sends a JSON body as plain text: {"api":"avatar","k":KEY,"av":"p:<base64>"}.
 */
function doPost(e) {
  let p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { p = {}; }
  if (p.api !== 'avatar' && p.api !== 'groupav') p = { api: 'bad' };   // only big requests come this way
  return api_(p);
}

// ---------- Profile pictures ----------
// Players tab column 18 (Avatar) holds one of:
//   p:<base64 JPEG>  a photo the person uploaded (the app shrinks it to 128x128 first, ~5-10 KB)
//   e:<emoji>        an emoji
//   t:<TEAM>         an NFL team logo (team abbreviation, e.g. PHI)
//   (blank)          none: the app shows a colored circle with their initial

/** Short fingerprint of a picture, so phones only download pictures that changed. */
function avStamp_(v) {
  if (!v) return '';
  if (v.indexOf('p:') !== 0) return v;   // emoji / team: the value itself is tiny
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, v);
  return 'p' + d.slice(0, 5).map(b => ((b + 256) % 256).toString(16)).join('');
}

// Group photos: Groups tab (Group, Avatar). Same formats as people's pictures. Only the admin sets them.
function groupAvs_() {
  const sh = ss_().getSheetByName('Groups');
  if (!sh || sh.getLastRow() < 2) return {};
  const out = {};
  sh.getDataRange().getValues().slice(1).forEach(r => { if (String(r[0]).trim() && String(r[1] || '')) out[String(r[0]).trim()] = String(r[1]); });
  return out;
}
function setGroupAv_(key, group, value) {
  requireAdmin_(key);
  group = String(group || '').trim();
  if (!group || groupNames_().map(g => g.toLowerCase()).indexOf(group.toLowerCase()) < 0) throw new Error('Couldn\'t find that group.');
  value = String(value || '').trim();
  if (value && !avValid_(value)) throw new Error('That picture didn\'t work. Try a different one.');
  const sh = ensureSheet_('Groups', ['Group', 'Avatar']), vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toLowerCase() === group.toLowerCase()) { sh.getRange(i + 1, 2).setNumberFormat('@').setValue(value); return groupAvs_(); }
  }
  sh.appendRow([group, '']); sh.getRange(sh.getLastRow(), 2).setNumberFormat('@').setValue(value);
  return groupAvs_();
}
function avValid_(value) {
  return (/^p:[A-Za-z0-9+/]+={0,2}$/.test(value) && value.length >= 200 && value.length <= 45000) ||
    (/^e:[^<>&"'\s]{1,16}$/.test(value) && !/[A-Za-z0-9]/.test(value.slice(2))) ||
    /^t:[A-Z]{2,3}$/.test(value);
}

// ---------- Privacy: who can see whom ----------
// The admin sees everyone. Everyone else only gets data for people who share a group with them
// (people with no group at all only see others with no group). Group names are limited to their own.
function myGroups_(me) {
  const meP = players_().find(p => p.id === me.id) || { groups: [] };
  return me.isAdmin ? groupNames_() : meP.groups;
}
function canSee_(me, p) {
  if (me.isAdmin || p.id === me.id) return true;
  const mine = myGroups_(me).map(g => g.toLowerCase());
  const theirs = groupsOf_(p);   // the admin counts as being in every group
  if (!mine.length) return p.id === adminId_() || !theirs.length;
  return theirs.some(g => mine.indexOf(g.toLowerCase()) >= 0);
}
function visibleGroupsOf_(me, p) {
  const all = groupsOf_(p);
  if (me.isAdmin) return all;
  const mine = myGroups_(me).map(g => g.toLowerCase());
  return all.filter(g => mine.indexOf(g.toLowerCase()) >= 0);
}

/** Pictures of everyone you can see. */
function avatars_(key) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  const out = {};
  players_().forEach(p => { if (p.avatar && canSee_(me, p)) out[p.id] = p.avatar; });
  return out;
}

/** Set your own picture, or (admin only) remove someone else's. Blank value = remove. */
function setAvatar_(key, playerId, value) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  const id = String(playerId || me.id).trim().toUpperCase();
  value = String(value || '').trim();
  if (id !== me.id) {
    if (!me.isAdmin) throw new Error('You can only change your own picture.');
    if (value) throw new Error('You can remove someone\'s picture, not set it.');
  }
  if (value) {
    const ok = (/^p:[A-Za-z0-9+/]+={0,2}$/.test(value) && value.length >= 200 && value.length <= 45000) ||
      (/^e:[^<>&"'\s]{1,16}$/.test(value) && !/[A-Za-z0-9]/.test(value.slice(2))) ||
      /^t:[A-Z]{2,3}$/.test(value);
    if (!ok) throw new Error('That picture didn\'t work. Try a different one.');
  }
  const sh = playersSheet_(), vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() !== id) continue;
    sh.getRange(i + 1, 18).setNumberFormat('@').setValue(value);
    PLAYERS_ = null;
    return { id: id, av: value };
  }
  throw new Error('Couldn\'t find that person.');
}

function getState(key, wk) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  const roster = players_().filter(p => p.active && canSee_(me, p));

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
  const weekGames = games.filter(g => weekKey_(g) === sel).sort((a, b) => a.kickoff - b.kickoff || a.id.localeCompare(b.id));
  const list = weekGames.map(g => gameView_(g, me, picks, now, roster));

  const parts = sel ? sel.split('|') : ['', '', ''];
  const board = leaderboard_(picks.vals, parts, roster);
  const mine = board.find(b => b.id === me.id) || { week: '0-0', season: '0-0' };

  const others = roster.filter(p => p.id !== me.id).sort((a, b) => a.name.localeCompare(b.name));
  const meP = players_().find(p => p.id === me.id) || { groups: [], hidden: [] };
  const groups = myGroups_(me);
  return {
    me: { id: me.id, name: me.name, isAdmin: me.isAdmin, pushId: 'gp_' + me.id, prefs: prefsFor_(me.id),
      groups: groupsOf_(meP), defaultGroup: me.isAdmin ? '' : (meP.groups[0] || ''), hidden: meP.hidden, pushIssue: meP.notifIssue || '' },
    groups: groups,
    groupAvs: (() => { const all = groupAvs_(), o = {}; groups.forEach(g => { if (all[g]) o[g] = all[g]; }); return o; })(),
    adminId: adminId_(),
    adminName: adminName_(),
    requests: me.isAdmin ? pendingRequests_().length : 0,
    my: { week: mine.week, season: mine.season },
    players: [meP].concat(others).map(p => ({ id: p.id, name: p.name, team: p.team, groups: visibleGroupsOf_(me, p), av: avStamp_(p.avatar) })),
    weeks: weeks,
    selected: sel,
    currentWeek: PropertiesService.getScriptProperties().getProperty('CURRENT_WEEK') || '',
    games: list,
    board: board,
    crown: crown_(games, picks.vals, roster),
    crowns: crowns_(games, picks.vals, roster, groups, me.isAdmin),
    tiebreak: tiebreakView_(weekGames, sel, me, roster, now),
    champs: champs_(games, picks.vals, roster, groups, me.isAdmin),
    reign: reignView_(games, groups, me.isAdmin),
    props: propsView_(games, me, roster, now),
    last: lastWeekFor_(games, picks.vals, me, roster),
    version: SCRIPT_VERSION,
  };
}

/**
 * One game as seen by one player. Other players' picks are only sent once you've
 * picked that game (or it has kicked off). Picks can be changed until kickoff.
 */
function gameView_(g, me, picks, now, roster) {
  const mine = picks.map[g.id + '|' + me.id];
  const locked = now >= g.kickoff.getTime();
  const others = roster.filter(p => p.id !== me.id).map(p => ({ p: p, pk: picks.map[g.id + '|' + p.id] }));
  const othersPicked = others.filter(x => x.pk).length;
  // Other people's picks show once you've LOCKED yours for this game, or once it kicks off.
  const reveal = (mine && mine.locked) || locked;
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
    mine: mine ? { pick: mine.pick, result: mine.result, override: mine.override, locked: !!mine.locked || locked } : null,
    othersPicked: othersPicked,
    othersTotal: others.length,
    others: reveal ? others.filter(x => x.pk).map(x => ({ name: x.p.name, pick: x.pk.pick, result: x.pk.result, override: x.pk.override })) : [],
  };
  out.all = {};
  roster.forEach(p => {
    const pk = picks.map[g.id + '|' + p.id];
    if (!pk) return;
    // before you lock: only whether they've locked (l) or are still picking, never the team
    out.all[p.id] = (reveal || p.id === me.id) ? { p: pk.pick, r: pk.result, o: pk.override, l: pk.locked ? 1 : 0 } : { h: 1, l: pk.locked ? 1 : 0 };
  });
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
    const had = picks.map[g.id + '|' + me.id];
    if (had && had.locked) throw new Error('You locked this pick in, so it can\'t be changed.');
    writePick_(g, me.id, pick, false, picks);
    const fresh = readPicks_();
    try { allInCheck_(weekKey_(g), fresh); } catch (e) { Logger.log('All-in check: ' + e.message); }
    return { game: gameView_(g, me, fresh, Date.now(), roster) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Lock in your picks for one or more games (comma-separated IDs). Locked picks can't be changed, and locking a game
 * reveals everyone else's picks for it. Unlocked picks still count; they lock on their own at kickoff.
 */
function lockPicks_(key, gameIds) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  const ids = String(gameIds || '').split(',').map(x => x.trim()).filter(Boolean);
  if (!ids.length) throw new Error('Nothing to lock.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(TABS.PICKS);
    if (String(sh.getRange(1, P.Locked + 1).getValue()) !== 'Locked') sh.getRange(1, P.Locked + 1).setValue('Locked').setFontWeight('bold');
    const games = {}; readGames_().forEach(g => { games[g.id] = g; });
    const picks = readPicks_();
    const now = new Date();
    let n = 0;
    ids.forEach(id => {
      const g = games[id], pk = picks.map[id + '|' + me.id];
      if (!g || !pk || pk.locked || now.getTime() >= g.kickoff.getTime()) return;
      sh.getRange(pk.row, P.Locked + 1).setValue(now);
      n++;
    });
    const fresh = readPicks_();
    const roster = players_().filter(p => p.active);
    return { locked: n, games: ids.filter(id => games[id]).map(id => gameView_(games[id], me, fresh, Date.now(), roster)) };
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
  const seeIds = {};
  players_().forEach(p => { if (canSee_(me, p)) seeIds[p.id] = true; });
  const pickRows = readPicks_().vals.slice(1).concat(archivedPickRows_())
    .filter(r => ids[String(r[P.GameID])] && seeIds[String(r[P.Player]).toUpperCase()])
    .map(r => [String(r[P.GameID]), String(r[P.Player]).toUpperCase(), String(r[P.Pick]).toUpperCase(), String(r[P.Result]),
      r[P.PickedAt] instanceof Date ? r[P.PickedAt].toISOString() : '']);
  return {
    me: me.id,
    season: want,
    seasons: seasons,
    players: players_().filter(p => seeIds[p.id]).map(p => ({ id: p.id, name: p.name, team: p.team, active: p.active, groups: visibleGroupsOf_(me, p) })),
    groups: myGroups_(me),
    hidden: (players_().find(p => p.id === me.id) || { hidden: [] }).hidden,
    games: gs.map(g => ({
      id: g.id, season: g.season, type: g.seasonType, week: g.week, kickoff: g.kickoff.toISOString(),
      away: g.awayAbbr, home: g.homeAbbr, awayName: g.away, homeName: g.home,
      awayScore: g.awayScore, homeScore: g.homeScore, status: g.status, winner: g.winner,
      awayPre: g.awayRecordPre, homePre: g.homeRecordPre, awayLogo: g.awayLogo, homeLogo: g.homeLogo,
    })),
    picks: pickRows,
    props: want === 'all' ? null : propsHistory_(games.filter(g => g.season === want), me),
    champs: want === 'all' ? null : champsHistory_(want, me),
    // every week's total game count (including games not played yet), so the app knows which weeks are finished
    weekGames: (() => { const o = {}; games.filter(g => want === 'all' || g.season === want).forEach(g => { const k = weekKey_(g); o[k] = (o[k] || 0) + 1; }); return o; })(),
  };
}

/** Ranked by season win % (late joiners aren't penalized), then total wins. Regular season and playoffs are counted separately. */
function leaderboard_(vals, parts, roster) {
  const t = {};
  roster.forEach(p => { t[p.id] = { id: p.id, name: p.name, wW: 0, wL: 0, sW: 0, sL: 0 }; });
  vals.slice(1).forEach(r => {
    const res = String(r[P.Result]);
    if (res !== 'W' && res !== 'L') return;
    const x = t[String(r[P.Player]).toUpperCase()];
    if (!x || String(r[P.Season]) !== parts[0]) return;
    if (parts[1] && String(r[P.SeasonType]) !== String(parts[1])) return;   // the playoffs are their own mini-season
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
    // skip playoff games whose teams aren't decided yet (ESPN lists them as TBD placeholders)
    const tbd = t => !t.abbreviation || /^TBD$/i.test(t.abbreviation) || /TBD|winner|seed/i.test(String(t.displayName || ''));
    if (tbd(home.team) || tbd(away.team)) return;

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
      lastSeen: r[12] instanceof Date ? r[12].toISOString() : (r[12] ? String(r[12]) : ''),
      appVersion: String(r[13] || ''), device: String(r[14] || ''), notifs: String(r[15] || ''),
      notifIssue: r[16] instanceof Date ? r[16].toISOString() : (r[16] ? String(r[16]) : ''),
      avatar: String(r[17] || ''),
      team: String(r[18] || '').trim(),
    }));
  return PLAYERS_;
}

function playerForKey_(key) {
  if (!key || !/^[a-f0-9]{32}$/.test(key)) return null;
  const p = players_().find(x => x.key === key && x.active);
  if (!p) return null;
  return { id: p.id, name: p.name, isAdmin: (settings_().AdminPlayer || 'P1').toUpperCase() === p.id };
}

function weekOrder_(k) { const [a, b, c] = String(k).split('|').map(Number); return a * 1000 + b * 100 + c; }

/** The week after season|type|week: regular season 1–18, then Wild Card, Divisional, Conference, Super Bowl. */
function nextWeekKey_(k) {
  const [season, type, week] = String(k).split('|').map(Number);
  if (type === 2) return week < 18 ? [season, 2, week + 1].join('|') : [season, 3, 1].join('|');
  if (type === 3) return week < 3 ? [season, 3, week + 1].join('|') : week === 3 ? [season, 3, 5].join('|') : '';
  return '';
}

/** As soon as every game of the current week is final, load next week's games and open picks for it. */
function openNextWeekEarly_() {
  const props = PropertiesService.getScriptProperties();
  const cur = props.getProperty('CURRENT_WEEK');
  if (!cur) return;
  const games = readGames_().filter(g => weekKey_(g) === cur);
  if (!games.length || !games.every(g => g.status === 'final')) return;
  const next = nextWeekKey_(cur);
  if (!next) return;
  const [season, type, week] = next.split('|');
  const ng = fetchScoreboard_(season, type, week).filter(g => weekKey_(g) === next);
  if (!ng.length) return;   // schedule not posted yet (e.g., playoff matchups still to be set); try again next refresh
  upsertGames_(ng);
  props.setProperty('CURRENT_WEEK', next);
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
      locked: !!r[P.Locked],
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
/** Sends one notification. `open` (optional) is added to the app address, e.g. '?recap=2026|2|7' opens that week's recap. */
function sendPush_(playerIds, title, body, open) {
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
  if (app) { const base = app.replace(/[?#].*$/, ''); payload.url = (base.endsWith('/') ? base : base + '/') + (open || ''); }
  const res = UrlFetchApp.fetch('https://api.onesignal.com/notifications?c=push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: (key.indexOf('os_v2_') === 0 ? 'Key ' : 'Basic ') + key },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const raw = res.getContentText();
  const text = res.getResponseCode() + ' ' + raw.slice(0, 300);
  logNotice_(playerIds.join(','), title + ' — ' + body, text);
  try { trackDelivery_(playerIds, res.getResponseCode(), raw); } catch (e) { Logger.log('trackDelivery: ' + e.message); }
  return text;
}

/**
 * Notices notifications that bounce. If OneSignal has no working sign-up for someone whose phone says notifications
 * are on, mark them (NotifIssue) so the app can offer them a one-tap fix. A later successful send clears it.
 */
function trackDelivery_(playerIds, code, raw) {
  if (code !== 200) return;
  let j = {}; try { j = JSON.parse(raw); } catch (e) { return; }
  const bad = ((j.errors && j.errors.invalid_aliases && j.errors.invalid_aliases.external_id) || []).map(x => String(x).replace(/^gp_/, '').toUpperCase());
  const failedAll = !j.id;   // nothing went out at all
  const ps = players_();
  const changes = [];
  playerIds.forEach(id => {
    const p = ps.find(x => x.id === id); if (!p) return;
    let bounced = bad.indexOf(id) >= 0 || (failedAll && playerIds.length === 1);
    // OneSignal's send reply can say "not linked" for people who DO get notifications. Double-check before flagging.
    if (bounced && !p.notifIssue && p.notifs === 'on') {
      const chk = pushLinked_(id);
      if (chk.code === 200 && chk.linked) bounced = false;
      else if (chk.code !== 200 && chk.code !== 404) bounced = false;   // couldn't ask: don't cry wolf
    }
    if (bounced && !p.notifIssue && p.notifs === 'on') changes.push([id, new Date()]);
    else if (!bounced && j.id && p.notifIssue) changes.push([id, '']);
  });
  if (!changes.length) return;
  const sh = playersSheet_(), vals = sh.getDataRange().getValues();
  changes.forEach(([id, v]) => { for (let i = 1; i < vals.length; i++) if (String(vals[i][0]).trim().toUpperCase() === id) { sh.getRange(i + 1, 17).setValue(v); break; } });
  PLAYERS_ = null;
}

/** Does OneSignal have a working push sign-up for this person? { linked, subs, code } (code 0 = couldn't ask). */
function pushLinked_(id) {
  const rest = String(settings_().OneSignalKey || '').trim();
  let subs = 0, code = 0;
  try {
    const res = UrlFetchApp.fetch('https://api.onesignal.com/apps/' + ONESIGNAL_APP_ID + '/users/by/external_id/' + encodeURIComponent('gp_' + id), {
      headers: rest ? { Authorization: (rest.indexOf('os_v2_') === 0 ? 'Key ' : 'Basic ') + rest } : {}, muteHttpExceptions: true });
    code = res.getResponseCode();
    if (code === 200) {
      const j = JSON.parse(res.getContentText() || '{}');
      subs = (j.subscriptions || []).filter(x => /push/i.test(String(x.type || '')) && x.enabled !== false &&
        !(typeof x.notification_types === 'number' && x.notification_types < 0)).length;
    }
  } catch (e) { Logger.log('pushLinked_: ' + e.message); }
  return { linked: subs > 0, subs: subs, code: code };
}

/**
 * App: "is my phone really linked to me?" Asks OneSignal whether this person (external id gp_<ID>) has a working
 * push sign-up. Only then is the "needs a fix" flag cleared. Returns { linked, subs }.
 */
function pushFixed_(key) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  const chk = pushLinked_(me.id), linked = chk.linked, subs = chk.subs, code = chk.code;
  if (linked) {
    const sh = playersSheet_(), vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) if (String(vals[i][0]).trim().toUpperCase() === me.id) { if (vals[i][16]) sh.getRange(i + 1, 17).setValue(''); break; }
    PLAYERS_ = null;
  }
  return { ok: true, linked: linked, subs: subs, code: code };
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

/** Optional team name (e.g. their fantasy team), shown under their name where there's room. Blank clears it. */
function cleanTeam_(t) {
  t = String(t || '').replace(/\s+/g, ' ').trim();
  if (t.length > 40) throw new Error('Keep team names to 40 characters or less.');
  if (/[<>]/.test(t)) throw new Error('Team names can\'t use < or >.');
  return t;
}
function setTeamCell_(id, team) {
  const sh = playersSheet_(), vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() === id) { sh.getRange(i + 1, 19).setNumberFormat('@').setValue(team); break; }
  }
  PLAYERS_ = null;
}
/** Admin: change anyone's name and team name. Everything is tied to their ID (P1, P2…), so picks, records and history follow. */
function adminRename_(key, playerId, name, team) {
  requireAdmin_(key);
  const id = String(playerId || '').trim().toUpperCase();
  if (!players_().some(p => p.id === id)) throw new Error('Couldn\'t find that person.');
  name = cleanName_(name);
  if (nameTaken_(name, id)) throw new Error('Someone already uses that name.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = playersSheet_(), vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) if (String(vals[i][0]).trim().toUpperCase() === id) { sh.getRange(i + 1, 2).setValue(name); break; }
    PLAYERS_ = null;
    if (team !== undefined) setTeamCell_(id, cleanTeam_(team));
    rebuildRecords_(readPicks_().vals);
  } finally { lock.releaseLock(); }
  return people_(key);
}
function setTeam_(key, team) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  setTeamCell_(me.id, cleanTeam_(team));
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
    people: players_().map(p => ({ id: p.id, name: p.name, team: p.team, active: p.active, groups: p.groups, link: base && p.key ? base + '?k=' + p.key : '',
      lastSeen: p.lastSeen, appVersion: p.appVersion, device: p.device, notifs: p.notifs, notifIssue: p.notifIssue, av: avStamp_(p.avatar) })),
    current: SCRIPT_VERSION,
    requests: pendingRequests_(),
  };
}

/** Admin: add a person and make their link. */
function addPerson_(key, name) {
  requireAdmin_(key);
  name = cleanName_(name);
  if (nameTaken_(name, '')) throw new Error('Someone already uses that name.');
  createPerson_(name, []);
  return people_(key);
}

/** Adds a row on the Players tab with a fresh link. Returns the new player ID. */
function createPerson_(name, groups) {
  if (!appBase_()) throw new Error('Put the app address in AppURL on the Settings tab first.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = playersSheet_();
    let maxId = 0;
    players_().forEach(p => { const m = /^P(\d+)$/.exec(p.id); if (m) maxId = Math.max(maxId, Number(m[1])); });
    maxId = Math.max(maxId, Number(PropertiesService.getScriptProperties().getProperty('MAX_ID') || 0));   // never reuse a deleted person's ID
    const id = 'P' + (maxId + 1), k = newKey_();
    sh.getRange('C:D').setNumberFormat('@');
    const row = [id, name, k, appBase_() + '?k=' + k, true, '', '', '', '', (groups || []).join(', ')];
    sh.appendRow(row);
    PLAYERS_ = null;
    return id;
  } finally {
    lock.releaseLock();
  }
}

// ---------- Invite links (admin shares one link per group; people join themselves) ----------
// Invites tab: Code, Group, Created, Expires, Active, Uses. One active link per group; each lasts 7 days.
const INV_HEADERS = ['Code', 'Group', 'Created', 'Expires', 'Active', 'Uses'];
const INVITE_DAYS = 7, INVITE_MAX_USES = 60;
function invSheet_() { const sh = ensureSheet_('Invites', INV_HEADERS); sh.getRange('A:B').setNumberFormat('@'); return sh; }

function invRows_() {
  const sh = invSheet_(), vals = sh.getDataRange().getValues(), now = Date.now();
  return vals.slice(1).map((r, i) => ({ row: i + 2, code: String(r[0]).toUpperCase(), group: String(r[1]),
    expires: r[3] instanceof Date ? r[3] : new Date(r[3]), active: r[4] === true || String(r[4]).toUpperCase() === 'TRUE',
    uses: Number(r[5]) || 0 })).map(x => Object.assign(x, { live: x.active && x.expires.getTime() > now && x.uses < INVITE_MAX_USES }));
}

function inviteOut_(x) { return { code: x.code, group: x.group, link: appBase_() + '?join=' + x.code, expires: x.expires.toISOString(), uses: x.uses }; }

/** Admin: action 'create' (new 7-day link for a group, replacing its old one) or 'off' (turn a link off). */
function invite_(key, action, group, code) {
  requireAdmin_(key);
  if (!appBase_()) throw new Error('Put the app address in AppURL on the Settings tab first.');
  const sh = invSheet_();
  if (action === 'off') {
    invRows_().filter(x => x.code === String(code || '').toUpperCase()).forEach(x => sh.getRange(x.row, 5).setValue(false));
    return invitesList_(key);
  }
  group = String(group || '').replace(/\s+/g, ' ').trim();
  if (!group) throw new Error('Pick a group for the link.');
  if (group.length > 30 || /[<>,]/.test(group)) throw new Error('Group names: 30 characters max, no commas.');
  invRows_().filter(x => x.active && x.group.toLowerCase() === group.toLowerCase()).forEach(x => sh.getRange(x.row, 5).setValue(false));
  const abc = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let c = ''; for (let i = 0; i < 6; i++) c += abc[Math.floor(Math.random() * abc.length)];
  sh.appendRow([c, group, new Date(), new Date(Date.now() + INVITE_DAYS * 864e5), true, 0]);
  return invitesList_(key);
}

function invitesList_(key) {
  requireAdmin_(key);
  return invRows_().filter(x => x.live).map(inviteOut_);
}

/** Public (no personal link needed): what the join screen shows. */
function joinInfo_(code) {
  const x = invRows_().find(r => r.code === String(code || '').toUpperCase());
  if (!x || !x.live) return { valid: false, admin: adminName_() };
  const count = players_().filter(p => p.active && p.id !== adminId_() && inGroup_(p, x.group)).length;
  return { valid: true, group: x.group, count: count, admin: adminName_(), groupAv: groupAvs_()[x.group] || '' };
}

/** Public: join with an invite link. Creates the person in that group and returns their personal key. */
function join_(code, name, team) {
  const cache = CacheService.getScriptCache(), ck = 'JOIN|' + String(code || '').toUpperCase();
  const n = Number(cache.get(ck) || 0);
  if (n >= 15) throw new Error('Lots of people are joining right now. Try again in a few minutes.');
  const x = invRows_().find(r => r.code === String(code || '').toUpperCase());
  if (!x || !x.live) throw new Error('This invite link isn\'t active anymore. Ask ' + adminName_() + ' for a new one.');
  name = cleanName_(name);
  team = cleanTeam_(team);
  if (nameTaken_(name, '')) throw new Error('Someone already goes by "' + name + '". Try adding a last initial.');
  cache.put(ck, String(n + 1), 600);
  const id = createPerson_(name, [x.group]);
  if (team) setTeamCell_(id, team);
  invSheet_().getRange(x.row, 6).setValue(x.uses + 1);
  const me = players_().find(p => p.id === id);
  try { sendPush_([adminId_()], '🎉 ' + name + ' joined', 'They joined the ' + x.group + ' group with your invite link.', ''); } catch (err) { /* never block a join */ }
  return { k: me.key, name: name, group: x.group };
}

// ---------- Requests to add someone (anyone asks, only the admin approves) ----------
const REQ_HEADERS = ['ID', 'Requested', 'From', 'Name', 'Note', 'Status', 'Resolved'];
function reqSheet_() { const sh = ensureSheet_('Requests', REQ_HEADERS); sh.getRange('A:A').setNumberFormat('@'); return sh; }

function pendingRequests_() {
  const sh = ss_().getSheetByName('Requests');
  if (!sh || sh.getLastRow() < 2) return [];
  const byId = {}; players_().forEach(p => { byId[p.id] = p; });
  return sh.getDataRange().getValues().slice(1).filter(r => r[0] !== '' && String(r[5]) === 'Pending').map(r => {
    const from = byId[String(r[2]).toUpperCase()] || { name: String(r[2]), groups: [] };
    return { id: String(r[0]), at: r[1] instanceof Date ? r[1].toISOString() : '', fromId: String(r[2]).toUpperCase(), fromName: from.name,
      name: String(r[3]), note: String(r[4]), suggestGroups: from.groups || [] };
  });
}

/** Anyone: ask the admin to add someone. Max 5 requests per person per day. */
function requestAdd_(key, name, note) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  name = cleanName_(name);
  if (nameTaken_(name, '')) throw new Error(name + ' is already in Greco Picks.');
  if (pendingRequests_().some(r => r.name.toLowerCase() === name.toLowerCase())) throw new Error('Someone already asked to add ' + name + '. ' + adminName_() + ' will take a look.');
  note = String(note || '').replace(/[<>]/g, '').trim().slice(0, 200);
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('req_' + me.id) || 0);
  if (n >= 5) throw new Error('That\'s a lot of requests for one day. Try again tomorrow.');
  cache.put('req_' + me.id, String(n + 1), 86400);
  const id = 'R' + Date.now();
  reqSheet_().appendRow([id, new Date(), me.id, name, note, 'Pending', '']);
  try {
    const admin = adminId_();
    if (admin !== me.id && String(settings_().OneSignalKey || '').trim()) {
      sendPush_([admin], 'New player request', me.name + ' wants to add ' + name + (note ? ': ' + note : '.') + ' Approve it in Settings → Manage people.');
    }
  } catch (e) { Logger.log('Request push failed: ' + e.message); }
  return { sent: true, name: name };
}

/** Admin: approve (adds them, optionally with an edited name and groups) or decline a request. */
function resolveRequest_(key, reqId, action, name, groups) {
  requireAdmin_(key);
  const sh = reqSheet_();
  const vals = sh.getDataRange().getValues();
  let row = 0, r = null;
  for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === String(reqId)) { row = i + 1; r = vals[i]; break; }
  if (!row || String(r[5]) !== 'Pending') throw new Error('That request was already handled.');
  const fromId = String(r[2]).toUpperCase();
  let newId = '';
  if (action === 'approve') {
    name = cleanName_(name || r[3]);
    if (nameTaken_(name, '')) throw new Error('Someone already uses that name. Change it before approving.');
    const known = groupNames_();
    const gl = splitList_(groups).map(g => known.find(k => k.toLowerCase() === g.toLowerCase()) || g.slice(0, 20));
    newId = createPerson_(name, gl);
    sh.getRange(row, 6, 1, 2).setValues([['Approved', new Date()]]);
    try {
      if (fromId !== adminId_() && String(settings_().OneSignalKey || '').trim())
        sendPush_([fromId], name + ' is in! 🎉', adminName_() + ' approved your request. They\'ll get their link soon.');
    } catch (e) { Logger.log('Approve push failed: ' + e.message); }
  } else if (action === 'decline') {
    sh.getRange(row, 6, 1, 2).setValues([['Declined', new Date()]]);
  } else throw new Error('Unknown action.');
  const out = people_(key);
  out.newId = newId;
  return out;
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

/** Someone already playing taps an invite link: add that group to their groups (no second account). */
function joinExisting_(key, code) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  const x = invRows_().find(r => r.code === String(code || '').toUpperCase());
  if (!x || !x.live) return { added: false, message: 'That invite link isn\'t active anymore. Ask ' + adminName_() + ' for a new one.' };
  if (me.isAdmin) return { already: true, group: x.group };
  const meP = players_().find(p => p.id === me.id);
  if (meP.groups.some(g => g.toLowerCase() === x.group.toLowerCase())) return { already: true, group: x.group };
  const sh = playersSheet_(), vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() !== me.id) continue;
    sh.getRange(i + 1, 10).setValue(meP.groups.concat([x.group]).join(', '));
    break;
  }
  PLAYERS_ = null;
  invSheet_().getRange(x.row, 6).setValue(x.uses + 1);
  try { sendPush_([adminId_()], '🎉 ' + meP.name + ' joined ' + x.group, 'They were already playing, so they were added to the ' + x.group + ' group (no new account).', ''); } catch (e) { }
  return { added: true, group: x.group };
}

function adminName_() { return (players_().find(p => p.id === adminId_()) || { name: 'Paul' }).name; }

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

/**
 * Admin: permanently delete a RETIRED person: their Players row, picks (current + archive) and tiebreaker guesses.
 * Their ID number is never reused (so old phones signed up under it can't receive someone else's notifications).
 */
function deletePerson_(key, playerId) {
  const me = requireAdmin_(key);
  playerId = String(playerId || '').toUpperCase();
  if (playerId === me.id) throw new Error('You can\'t delete yourself.');
  const p = players_().find(x => x.id === playerId);
  if (!p) throw new Error('Unknown player.');
  if (p.active) throw new Error('Retire them first, then delete.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    const n = Number((/^P(\d+)$/.exec(playerId) || [])[1] || 0);
    props.setProperty('MAX_ID', String(Math.max(n, Number(props.getProperty('MAX_ID') || 0))));
    const dropRows = (sh, col) => {
      if (!sh || sh.getLastRow() < 2) return 0;
      const vals = sh.getDataRange().getValues(); let gone = 0;
      for (let i = vals.length - 1; i >= 1; i--) if (String(vals[i][col]).trim().toUpperCase() === playerId) { sh.deleteRow(i + 1); gone++; }
      return gone;
    };
    const picks = dropRows(ss_().getSheetByName(TABS.PICKS), P.Player) + dropRows(ss_().getSheetByName(TABS.PICKS_ARCH), P.Player);
    dropRows(ss_().getSheetByName(TABS.TIEBREAK), 1);
    dropRows(ss_().getSheetByName('Props'), 1);
    dropRows(playersSheet_(), 0);
    PLAYERS_ = null;
    try { rebuildRecords_(readPicks_().vals); } catch (e) { Logger.log('rebuildRecords after delete: ' + e.message); }
    const out = people_(key); out.deleted = { name: p.name, picks: picks };
    return out;
  } finally { lock.releaseLock(); }
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
function crowns_(games, pickVals, roster, groups, withEveryone) {
  const out = withEveryone !== false ? { '': crown_(games, pickVals, roster) } : {};
  (groups || groupNames_()).forEach(g => { out[g] = crown_(games, pickVals, roster.filter(p => inGroup_(p, g))); });
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
  if (parts[1] === '3') return null;   // no 👑/💩/💯 in the playoffs
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
  const losers = played.length > 1 && low < top && byWeek[done[0]].length >= 3 ? played.filter(x => x.wins === low) : [];   // no 💩 in tiny playoff weeks
  return {
    label: weekLabel_(parts[1], parts[2]), key: done[0], ids: winners.map(x => x.id), names: winners.map(x => x.name), rec: winners[0].rec, tb: tb,
    poop: losers.length ? { ids: losers.map(x => x.id), names: losers.map(x => x.name), rec: losers[0].rec } : null,
  };
}

// ======================================================================
// Tiebreaker, last week, milestones, archive
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
  // Don't wait behind the 30-minute score update (that was the "frozen" Save): a guess only touches its own row.
  const lock = { releaseLock: () => {} };
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
  // "Between weeks": that week is final and no game after it has kicked off yet.
  const lastKick = Math.max.apply(null, byWeek[done[0]].map(g => g.kickoff.getTime()));
  const now = Date.now();
  const between = !games.some(g => weekKey_(g) !== done[0] && g.kickoff.getTime() > lastKick && g.kickoff.getTime() <= now);
  return { key: done[0], label: weekLabel_(parts[1], parts[2]), rec: rec, games: n, perfect: l === 0 && w === n && n > 0, between: between };
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
    if (Object.keys(st.weeks).some(k => st.weeks[k].W === weekSize[k] && st.weeks[k].n === weekSize[k] && weekSize[k] >= 4 && k.split('|')[1] === '2'))   // regular-season weeks only
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

/**
 * Admin: tell everyone on an older app version (or an unknown one) to update. Sent once per person per version.
 * Tapping the notification opens the app and loads the newest version.
 */
function updateNotice_(key) {
  requireAdmin_(key);
  if (!String(settings_().OneSignalKey || '').trim()) throw new Error('Notifications aren\'t set up (OneSignalKey).');
  const older = v => { if (!v) return true; const a = String(v).split('.').map(Number), b = SCRIPT_VERSION.split('.').map(Number);
    for (let i = 0; i < 3; i++) { if ((a[i] || 0) !== (b[i] || 0)) return (a[i] || 0) < (b[i] || 0); } return false; };
  const sent = sentSet_(), out = { sent: [], skipped: [] };
  players_().filter(p => p.active && older(p.appVersion)).forEach(p => {
    const k = 'UPD|' + SCRIPT_VERSION + '|' + p.id;
    if (sent.indexOf(k) >= 0) { out.skipped.push(p.name + ' (already sent)'); return; }
    const res = sendPush_([p.id], '✨ Greco Picks has an update', 'Tap to get the latest version.', '?update=1');
    if (res.indexOf('"id":""') >= 0 || res.indexOf('"id"') < 0) { out.skipped.push(p.name + (p.notifs === 'on' ? ' (phone\'s notifications need a fix; they\'ll be asked next time they open the app)' : ' (notifications off)')); return; }
    markSent_(sent, k);
    out.sent.push(p.name);
  });
  return out;
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
    const tbG = ws.find(g => tieGame_(games.filter(x => weekKey_(x) === weekKey_(g))) === g);
    const tbAll = tbG ? readTiebreaks_() : {};
    const isSb = tbG && tbG.seasonType === 3 && tbG.week === 5;
    const sbProps = isSb ? readProps_(tbG.season) : {};
    roster.filter(p => p.remind).forEach(p => {
      const open = ws.filter(g => !picks.map[g.id + '|' + p.id]);
      const noTb = tbG && (isSb ? !propsDone_(sbProps[p.id]) : tbAll[weekKey_(tbG) + '|' + p.id] === undefined);
      if (!open.length && noTb) {
        const dk = 'RT|' + weekKey_(tbG) + '|' + p.id;
        if (sent.indexOf(dk) < 0) {
          if (isSb) sendPush_([p.id], '🏈 Finish your Super Bowl props', 'Your props and final-score guess lock at kickoff (' + when.time + ').');
          else sendPush_([p.id], '🎯 Tiebreaker needed', 'Guess the total points for ' + tbG.away + ' @ ' + tbG.home + ' before ' + when.time + '.');
          markSent_(sent, dk);
        }
        return;
      }
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
  } else if (cur && cur !== lastWeek && hourEt >= 9 && hourEt < 21) {   // waits for a decent hour (Eastern)
    const wkGames = games.filter(g => weekKey_(g) === cur);
    if (wkGames.some(g => g.status === 'pre')) {
      const [, type, week] = cur.split('|');
      // only people who still have games to pick this week
      const ids = roster.filter(p => p.newweek && wkGames.some(g => g.status === 'pre' && !picks.map[g.id + '|' + p.id])).map(p => p.id);
      if (type === '3' && week === '5') sendPush_(ids, '🏈 Super Bowl picks and props are open', 'Make your pick, then answer 8 quick props and guess the final score. Everything locks at kickoff.');
      else sendPush_(ids, weekLabel_(type, week) + ' is up', 'The ' + weekLabel_(type, week) + ' games are ready. Make your picks!');
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
    const regOver = type === '2' && !games.some(g => g.seasonType === 2 && String(g.season) === season && g.status !== 'final');
    // Season lead within each person's own (first) group; the admin and people with no group get everyone.
    const leadFor = {};
    const leadLine = (group, p) => {
      if (!(group in leadFor)) {
        const gb = group ? leaderboard_(picks.vals, [season, type, week], roster.filter(x => inGroup_(x, group)))
          : group === null ? leaderboard_(picks.vals, [season, type, week], roster.filter(x => !x.groups.length)) : board;   // null = no group: only others with no group
        const top = gb.filter(b => b.rank === 1);
        leadFor[group] = { top: top, label: type === '3' ? (group ? 'the ' + groupPhrase_(group) + ' playoffs' : 'the playoffs') : (group ? 'the ' + groupPhrase_(group) : 'the season') };
        // last week of the regular season / the Super Bowl: name the champ instead of the leader
        if (regOver && type === '2') leadFor[group].champ = champFor_(games, picks.vals, group ? roster.filter(x => inGroup_(x, group)) : group === null ? roster.filter(x => !x.groups.length) : roster).reg;
        if (type === '3' && week === '5') leadFor[group].champ = champFor_(games, picks.vals, group ? roster.filter(x => inGroup_(x, group)) : group === null ? roster.filter(x => !x.groups.length) : roster).post;
      }
      const L = leadFor[group];
      if (L.champ) {
        const what = type === '3' ? 'playoff champ' : 'regular-season champ', where = group ? ' of the ' + groupPhrase_(group) : '';
        const icon = type === '3' ? '🥇 ' : '🏆 ';
        if (L.champ.ids.indexOf(p.id) >= 0) return icon + (L.champ.ids.length > 1 ? 'You\'re a co-' : 'You\'re the ') + what + where + '!';
        return icon + L.champ.names.join(' & ') + (L.champ.ids.length > 1 ? ' are co-' + what + 's' : ' is the ' + what) + where + '.';
      }
      if (!L.top.length) return '';
      if (L.top.length > 1) return 'Tied for the lead in ' + L.label + ' at ' + L.top[0].season + '.';
      if (L.top[0].id === p.id) return 'You lead ' + L.label + ' at ' + L.top[0].season + '!';
      return L.top[0].name + ' leads ' + L.label + ' at ' + L.top[0].season + '.';
    };
    roster.filter(p => p.results).forEach(p => {
      const b = board.find(x => x.id === p.id);
      if (!b || b.week === '0-0') return;
      // One leader line per group they're in (up to 3); the admin gets everyone; no group = others with no group.
      const gs = p.id === adminId_() ? [''] : (p.groups.length ? p.groups.slice(0, 3) : [null]);
      const lines = gs.map(g => leadLine(g, p)).filter(Boolean);
      sendPush_([p.id], weekLabel_(type, week) + ' results', 'You went ' + b.week + '. ' + lines.join(' ') + ' Tap to see your recap.', '?recap=' + encodeURIComponent(wk));
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
          // every group they're in where it happened (labeled), not just the first one
          const k = f.game.id + '|' + f.grp;
          if (seen[k] || mine.indexOf(f.grp) < 0) return;
          seen[k] = 1;
          items.push(f);
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
          const shown = items.slice(0, 4), more = items.length - shown.length;
          body = items.length + ' upsets today' + (oneGroup ? ' in the ' + groupPhrase_(items[0].grp) : '') + ': ' + shown.map(line).join('; ') + (more ? '; and ' + more + ' more' : '') + '.';
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


// ======================================================================
// Champions and Super Bowl props
// ======================================================================
// The regular season and the playoffs are two separate standings.
//   🏆 regular-season champ: best record once every regular-season game is final (ties share it).
//   🥇 playoff champ: best playoff record once the Super Bowl is final; ties go to Super Bowl props, then the final-score guess.

function champs_(games, pickVals, roster, groups, withEveryone) {
  const out = withEveryone !== false ? { '': champFor_(games, pickVals, roster) } : {};
  (groups || groupNames_()).forEach(g => { out[g] = champFor_(games, pickVals, roster.filter(p => inGroup_(p, g))); });
  return out;
}

function champFor_(games, pickVals, roster) {
  const reg = games.filter(g => g.seasonType === 2);
  if (!reg.length) return { reg: null, post: null };
  const season = Math.max.apply(null, reg.map(g => g.season));
  const out = { season: season, reg: null, post: null };
  const regG = reg.filter(g => g.season === season);
  if (regG.every(g => g.status === 'final')) {
    const b = leaderboard_(pickVals, [String(season), '2', '0'], roster).filter(x => x.rank === 1);
    if (b.length) out.reg = { ids: b.map(x => x.id), names: b.map(x => x.name), rec: b[0].season };
  }
  const sb = sbGame_(games.filter(g => g.season === season));
  if (sb && sb.status === 'final') {
    const rows = leaderboard_(pickVals, [String(season), '3', '0'], roster).filter(x => x.pct >= 0);
    if (rows.length) {
      const pv = propsBoard_(sb, roster);
      const top = Math.max.apply(null, rows.map(x => x.wins));
      let best = rows.filter(x => x.wins === top);
      const sc = id => (pv[id] || { score: -1 }).score, off = id => (pv[id] && pv[id].off != null) ? pv[id].off : Infinity;
      if (best.length > 1) { const m = Math.max.apply(null, best.map(x => sc(x.id))); best = best.filter(x => sc(x.id) === m); }
      if (best.length > 1) { const m = Math.min.apply(null, best.map(x => off(x.id))); if (m !== Infinity) best = best.filter(x => off(x.id) === m); }
      out.post = { ids: best.map(x => x.id), names: best.map(x => x.name), rec: best[0].season };
    }
  }
  return out;
}

// ---------- Super Bowl props ----------
// Props tab: Season | Player | Answers (JSON) | At. Answers use these keys:
//   coin H/T · first A/H (away/home scores first) · ftype TD/FG/SF · ou O/U (total points vs the line)
//   margin 1-3/4-7/8-14/15+ · half A/H/T (halftime leader or tied) · gator (Gatorade color) · mvp (ESPN player id or OTHER)
//   sa / sh: final-score guess (away / home). Their sum also goes on the Tiebreakers tab.
// Script properties: SB_CONF_<season> {line, mvp:[{id,n,pos,t}], ans:{...answers you set}}; SB_AUTO_<season> {...answers from ESPN}.
const SB_Q = ['coin', 'first', 'ftype', 'ou', 'margin', 'half', 'gator', 'mvp'];
const SB_OPTS = {
  coin: ['H', 'T'], first: ['A', 'H'], ftype: ['TD', 'FG', 'SF'], ou: ['O', 'U'], margin: ['1-3', '4-7', '8-14', '15+'],
  half: ['A', 'H', 'T'], gator: ['Orange', 'Yellow', 'Blue', 'Purple', 'Red', 'Green', 'Clear', 'None'],
};

function sbGame_(games) {
  return games.filter(g => g.seasonType === 3 && g.week === 5).sort((a, b) => b.season - a.season || a.kickoff - b.kickoff)[0] || null;
}
function sbJson_(name) {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(name) || '{}') || {}; } catch (e) { return {}; }
}
function sbConf_(season) { const c = sbJson_('SB_CONF_' + season); c.mvp = c.mvp || []; c.ans = c.ans || {}; return c; }
function sbAuto_(season) { return sbJson_('SB_AUTO_' + season); }
function sbSave_(name, obj) { PropertiesService.getScriptProperties().setProperty(name, JSON.stringify(obj)); }

let PROPS_ = {};
function readProps_(season) {
  if (PROPS_[season]) return PROPS_[season];
  const sh = ss_().getSheetByName('Props');
  const o = {};
  if (sh && sh.getLastRow() > 1) sh.getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[0]) !== String(season)) return;
    try { o[String(r[1]).toUpperCase()] = JSON.parse(String(r[2] || '{}')); } catch (e) { /* skip a bad row */ }
  });
  PROPS_[season] = o;
  return o;
}
function propsDone_(a) { return !!a && SB_Q.every(q => a[q] !== undefined && a[q] !== '') && a.sa !== undefined && a.sh !== undefined; }

/** The right answers so far: ESPN's (scored automatically), then anything the admin set (wins). */
function sbAnswers_(g) {
  const conf = sbConf_(g.season), auto = sbAuto_(g.season), out = {};
  ['first', 'ftype', 'half'].forEach(q => { if (auto[q]) out[q] = auto[q]; });
  const line = conf.line != null ? Number(conf.line) : (auto.line != null ? Number(auto.line) : null);
  if (g.status === 'final') {
    const a = Number(g.awayScore || 0), h = Number(g.homeScore || 0), total = a + h, m = Math.abs(a - h);
    if (line != null) out.ou = total > line ? 'O' : total < line ? 'U' : 'P';   // P = landed exactly on the line: nobody gets it
    out.margin = m <= 3 ? '1-3' : m <= 7 ? '4-7' : m <= 14 ? '8-14' : '15+';
  }
  Object.keys(conf.ans).forEach(q => { if (conf.ans[q] !== '' && conf.ans[q] != null) out[q] = conf.ans[q]; });
  return { answers: out, line: line };
}

/** Everyone's props score: {id: {score, graded, off}}. off = how far the final-score guess was (only once final). */
function propsBoard_(g, roster) {
  const all = readProps_(g.season), ans = sbAnswers_(g).answers, out = {};
  roster.forEach(p => {
    const a = all[p.id]; if (!a) return;
    let score = 0;
    SB_Q.forEach(q => { if (ans[q] !== undefined && a[q] === ans[q]) score++; });
    const off = g.status === 'final' && a.sa !== undefined && a.sh !== undefined
      ? Math.abs(Number(a.sa) - Number(g.awayScore || 0)) + Math.abs(Number(a.sh) - Number(g.homeScore || 0)) : null;
    out[p.id] = { score: score, off: off };
  });
  return out;
}

/** What the app needs for the Super Bowl props card (null when there's no Super Bowl on the schedule). */
function propsView_(games, me, roster, now) {
  const g = sbGame_(games);
  if (!g) return null;
  const conf = sbConf_(g.season), all = readProps_(g.season), A = sbAnswers_(g);
  const locked = now >= g.kickoff.getTime();
  const board = propsBoard_(g, roster);
  return {
    gameId: g.id, season: g.season, week: weekKey_(g), kickoff: g.kickoff.toISOString(), locked: locked, status: g.status,
    away: { abbr: g.awayAbbr, name: g.away, logo: g.awayLogo, score: g.awayScore }, home: { abbr: g.homeAbbr, name: g.home, logo: g.homeLogo, score: g.homeScore },
    line: A.line, mvp: conf.mvp, roster: sbJson_('SB_ROSTER_' + g.season).list || [],
    mine: all[me.id] || {},
    count: roster.filter(p => propsDone_(all[p.id])).length,
    answers: locked || me.isAdmin ? A.answers : {},
    auto: Object.keys(sbAuto_(g.season)).filter(k => ['first', 'ftype', 'half'].indexOf(k) >= 0),
    all: locked ? roster.filter(p => all[p.id]).map(p => ({ id: p.id, name: p.name, a: all[p.id], score: board[p.id].score, off: board[p.id].off })) : [],
    admin: me.isAdmin ? conf.ans : undefined,
  };
}

/** Save some or all of your props (JSON object). Only before kickoff. */
function setProps_(key, json) {
  const me = playerForKey_(key);
  if (!me) throw new Error('This link is no longer valid.');
  let inp;
  try { inp = JSON.parse(String(json || '{}')); } catch (e) { throw new Error('Couldn\'t read that.'); }
  const games = readGames_(), g = sbGame_(games);
  if (!g) throw new Error('There\'s no Super Bowl on the schedule yet.');
  if (Date.now() >= g.kickoff.getTime()) throw new Error('The Super Bowl has kicked off. Props are locked.');
  const conf = sbConf_(g.season);
  const clean = {};
  Object.keys(inp || {}).forEach(q => {
    const v = inp[q];
    if (v === '' || v == null) { clean[q] = ''; return; }
    if (SB_OPTS[q]) { if (SB_OPTS[q].indexOf(String(v)) < 0) throw new Error('That answer isn\'t one of the choices.'); clean[q] = String(v); }
    else if (q === 'mvp') { if (!conf.mvp.some(x => String(x.id) === String(v)) && !sbRosterIds_(g)[String(v)]) throw new Error('Pick a player from the list.'); clean[q] = String(v); }
    else if (q === 'sa' || q === 'sh') { const n = Math.round(Number(v)); if (!isFinite(n) || n < 0 || n > 99) throw new Error('Scores go from 0 to 99.'); clean[q] = n; }
  });
  const sh = ensureSheet_('Props', ['Season', 'Player', 'Answers', 'At']);
  const vals = sh.getDataRange().getValues();
  let row = 0, cur = {};
  for (let i = 1; i < vals.length; i++) if (String(vals[i][0]) === String(g.season) && String(vals[i][1]).toUpperCase() === me.id) { row = i + 1; try { cur = JSON.parse(String(vals[i][2] || '{}')); } catch (e) { cur = {}; } break; }
  Object.keys(clean).forEach(q => { if (clean[q] === '') delete cur[q]; else cur[q] = clean[q]; });
  if (row) sh.getRange(row, 3, 1, 2).setValues([[JSON.stringify(cur), new Date()]]);
  else sh.appendRow([g.season, me.id, JSON.stringify(cur), new Date()]);
  PROPS_ = {};
  // the final-score guess doubles as the week's tiebreaker (total points)
  if (cur.sa !== undefined && cur.sh !== undefined && ('sa' in clean || 'sh' in clean)) {
    try { setTiebreak_(key, weekKey_(g), Number(cur.sa) + Number(cur.sh)); } catch (e) { Logger.log('Props tiebreak: ' + e.message); }
  }
  return propsView_(games, me, players_().filter(p => p.active && canSee_(me, p)), Date.now());
}

/** Admin: set answers (coin, gator, mvp, or fix any other), the over/under line, or the MVP list. */
function sbAdmin_(key, json) {
  const me = requireAdmin_(key) || playerForKey_(key);
  let inp;
  try { inp = JSON.parse(String(json || '{}')); } catch (e) { throw new Error('Couldn\'t read that.'); }
  const games = readGames_(), g = sbGame_(games);
  if (!g) throw new Error('There\'s no Super Bowl on the schedule yet.');
  const conf = sbConf_(g.season);
  if (inp.ans) Object.keys(inp.ans).forEach(q => {
    const v = inp.ans[q];
    if (SB_Q.indexOf(q) < 0) return;
    if (v === '' || v == null) { delete conf.ans[q]; return; }
    if (SB_OPTS[q] && SB_OPTS[q].concat(q === 'ou' ? ['P'] : []).indexOf(String(v)) < 0) throw new Error('That answer isn\'t one of the choices.');
    if (q === 'mvp' && !conf.mvp.some(x => String(x.id) === String(v)) && !sbRosterIds_(g)[String(v)]) throw new Error('Pick a player from the list.');
    conf.ans[q] = String(v);
  });
  if ('line' in inp) {
    const n = Number(inp.line);
    if (inp.line === '' || inp.line == null) delete conf.line;
    else if (!isFinite(n) || n < 10 || n > 99) throw new Error('The line should be between 10 and 99.');
    else conf.line = Math.round(n * 2) / 2;
  }
  if (Array.isArray(inp.mvp)) {
    conf.mvp = inp.mvp.slice(0, 40).map(x => ({ id: String(x.id).replace(/\D/g, '').slice(0, 12), n: String(x.n || '').slice(0, 40), pos: String(x.pos || '').slice(0, 4), t: String(x.t || '').toUpperCase().slice(0, 3) }))
      .filter(x => x.id && x.n && (x.t === g.awayAbbr || x.t === g.homeAbbr));
  }
  sbSave_('SB_CONF_' + g.season, conf);
  return propsView_(games, me, players_().filter(p => p.active), Date.now());
}

function espnJson_(url) {
  const r = UrlFetchApp.fetch(url, FETCH_OPTS);
  if (r.getResponseCode() !== 200) throw new Error(r.getResponseCode() + ' ' + url);
  return JSON.parse(r.getContentText());
}
function sbSummary_(g) {
  const urls = ['https://site.api.espn.com/apis/site/v2/sports/football/nfl/summary?event=' + g.id,
    'https://site.web.api.espn.com/apis/site/v2/sports/football/nfl/summary?event=' + g.id];
  let err;
  for (const u of urls) { try { return espnJson_(u); } catch (e) { err = e; } }
  throw err;
}

/** Both teams' full rosters, saved once (SB_ROSTER_<season>) so anyone can pick any player for MVP. */
function sbRosterList_(g, refetch) {
  const have = sbJson_('SB_ROSTER_' + g.season);
  if (have.list && have.list.length && !refetch && have.teams === g.awayAbbr + g.homeAbbr) return have.list;
  const r = fetchRosters_(g);
  const list = r.away.concat(r.home).map(x => [x.id, x.n, x.pos, x.t]);
  if (list.length) sbSave_('SB_ROSTER_' + g.season, { teams: g.awayAbbr + g.homeAbbr, list: list });
  return list;
}
/** Admin: both teams' rosters from ESPN, for editing the MVP list. */
function sbRoster_(key) {
  requireAdmin_(key);
  const g = sbGame_(readGames_());
  if (!g) throw new Error('There\'s no Super Bowl on the schedule yet.');
  const list = sbRosterList_(g, true).map(x => ({ id: x[0], n: x[1], pos: x[2], t: x[3] }));
  return { away: list.filter(x => x.t === g.awayAbbr), home: list.filter(x => x.t === g.homeAbbr) };
}
function fetchRosters_(g) {
  const one = abbr => {
    const d = espnJson_('https://site.api.espn.com/apis/site/v2/sports/football/nfl/teams/' + abbr.toLowerCase() + '/roster');
    const out = [];
    (d.athletes || []).forEach(grp => (grp.items || []).forEach(a => out.push({ id: String(a.id), n: a.displayName || a.fullName || '', pos: ((a.position || {}).abbreviation) || '', t: abbr })));
    return out;
  };
  return { away: one(g.awayAbbr), home: one(g.homeAbbr) };
}
function sbRosterIds_(g) { const o = {}; (sbJson_('SB_ROSTER_' + g.season).list || []).forEach(x => { o[String(x[0])] = true; }); return o; }

/** Season leaders from ESPN, used to build the MVP shortlist (about 10 per team). */
function sbLeaders_(data, g) {
  const byTeam = {};
  (data.leaders || []).forEach(t => {
    const abbr = String(((t.team || {}).abbreviation) || '').toUpperCase();
    if (abbr !== g.awayAbbr && abbr !== g.homeAbbr) return;
    const list = byTeam[abbr] = byTeam[abbr] || [];
    (t.leaders || []).forEach(cat => (cat.leaders || []).slice(0, 2).forEach(l => {
      const a = l.athlete || {}; if (!a.id || list.some(x => x.id === String(a.id))) return;
      list.push({ id: String(a.id), n: a.displayName || a.shortName || '', pos: ((a.position || {}).abbreviation) || '', t: abbr });
    }));
  });
  const order = ['QB', 'RB', 'WR', 'TE'];
  const sortIt = l => l.sort((a, b) => { const x = order.indexOf(a.pos), y = order.indexOf(b.pos); return (x < 0 ? 9 : x) - (y < 0 ? 9 : y); }).slice(0, 10);
  return sortIt(byTeam[g.awayAbbr] || []).concat(sortIt(byTeam[g.homeAbbr] || []));
}

/** Runs with every refresh: fills the MVP list and line before the game, and scores the ESPN props during/after it. */
function sbTick_() {
  const g = sbGame_(readGames_());
  if (!g) return;
  const conf = sbConf_(g.season), auto = sbAuto_(g.season);
  const cache = CacheService.getScriptCache();
  if (g.status === 'pre') {
    const haveRoster = (sbJson_('SB_ROSTER_' + g.season).teams || '') === g.awayAbbr + g.homeAbbr;
    if ((conf.mvp.length && auto.line != null && haveRoster) || cache.get('sb_pre')) return;
    cache.put('sb_pre', '1', 3 * 3600);   // at most every 3 hours before the game
    if (!haveRoster) { try { sbRosterList_(g, true); } catch (e) { Logger.log('Rosters: ' + e.message); } }
    const d = sbSummary_(g);
    const pc = (d.pickcenter || [])[0] || {};
    if (pc.overUnder != null && isFinite(Number(pc.overUnder))) { auto.line = Number(pc.overUnder); sbSave_('SB_AUTO_' + g.season, auto); }
    if (!conf.mvp.length) { const l = sbLeaders_(d, g); if (l.length) { conf.mvp = l; sbSave_('SB_CONF_' + g.season, conf); } }
    return;
  }
  if (auto.complete) return;
  const d = sbSummary_(g);
  const sp = (d.scoringPlays || []).slice();
  if (sp.length) {
    const f = sp[0], a = Number(f.awayScore || 0), h = Number(f.homeScore || 0), pts = a + h;
    auto.first = a > 0 ? 'A' : 'H';
    const txt = String(((f.type || {}).text) || f.text || '');
    auto.ftype = pts === 2 || /safety/i.test(txt) ? 'SF' : pts === 3 || /field goal/i.test(txt) ? 'FG' : 'TD';
  }
  const comp = ((d.header || {}).competitions || [])[0] || {};
  const period = Number(((comp.status || {}).period) || 0);
  if (period >= 3 || g.status === 'final') {
    const half = side => { const c = (comp.competitors || []).find(x => x.homeAway === side) || {}; return (c.linescores || []).slice(0, 2).reduce((s, x) => s + Number(x.displayValue != null ? x.displayValue : (x.value || 0)), 0); };
    const A = half('away'), H = half('home');
    if (comp.competitors) auto.half = A > H ? 'A' : H > A ? 'H' : 'T';
  }
  if (g.status === 'final') auto.complete = true;
  sbSave_('SB_AUTO_' + g.season, auto);
}

// ---------- Reigning champ ----------
// Once the regular season is final, each group's champ is saved (CHAMPS_<season>), so it survives the summer archive.
// The reigning (most recent) champ gets a gold ring on their picture all the next season, and keeps the 🏆 badge
// until the next season's Week 1 is final. The playoff champ doesn't carry over.
function storeChamps_() {
  const games = readGames_(), reg = games.filter(g => g.seasonType === 2);
  if (!reg.length) return;
  const season = Math.max.apply(null, reg.map(g => g.season));
  if (!reg.filter(g => g.season === season).every(g => g.status === 'final')) return;
  const roster = players_().filter(p => p.active), picks = readPicks_();
  const out = { season: season, g: {} };
  [''].concat(groupNames_()).forEach(grp => {
    const both = champFor_(games, picks.vals, grp ? roster.filter(p => inGroup_(p, grp)) : roster), c = both.reg;
    if (c) out.g[grp] = { ids: c.ids, names: c.names, rec: c.rec };
    if (both.post) (out.post = out.post || {})[grp] = { ids: both.post.ids, names: both.post.names, rec: both.post.rec };
  });
  const key = 'CHAMPS_' + season, val = JSON.stringify(out);
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(key) !== val) props.setProperty(key, val);
}

function reignView_(games, groups, isAdmin) {
  const props = PropertiesService.getScriptProperties();
  const seasons = props.getKeys().map(k => (/^CHAMPS_(\d{4})$/.exec(k) || [])[1]).filter(Boolean).map(Number).sort((a, b) => b - a);
  if (!seasons.length) return null;
  let c; try { c = JSON.parse(props.getProperty('CHAMPS_' + seasons[0])); } catch (e) { return null; }
  // 🏆 stays on their picture until the next season's Week 1 is final
  const wk1 = games.filter(g => g.season === c.season + 1 && g.seasonType === 2 && g.week === 1);
  const badge = !(wk1.length && wk1.every(g => g.status === 'final'));
  const g = {};
  (isAdmin ? [''] : []).concat(groups).forEach(k => { if (c.g[k]) g[k] = { ids: c.g[k].ids, names: c.g[k].names, rec: c.g[k].rec }; });
  return { season: c.season, badge: badge, g: g };
}

/** A past (or current) season's Super Bowl props, for Stats: everyone you can see, their answers and scores. */
function propsHistory_(seasonGames, me) {
  const g = sbGame_(seasonGames);
  if (!g || Date.now() < g.kickoff.getTime()) return null;
  const all = readProps_(g.season);
  if (!Object.keys(all).length) return null;
  const roster = players_().filter(p => all[p.id] && canSee_(me, p));
  const board = propsBoard_(g, roster), conf = sbConf_(g.season), A = sbAnswers_(g);
  return {
    season: g.season, gameId: g.id, locked: true, status: g.status, line: A.line, answers: A.answers, mvp: conf.mvp,
    roster: sbJson_('SB_ROSTER_' + g.season).list || [],
    away: { abbr: g.awayAbbr, name: g.away, logo: g.awayLogo, score: g.awayScore }, home: { abbr: g.homeAbbr, name: g.home, logo: g.homeLogo, score: g.homeScore },
    all: roster.map(p => ({ id: p.id, name: p.name, a: all[p.id], score: board[p.id].score, off: board[p.id].off })),
  };
}
/** Saved champs for a season (regular season + playoffs), limited to your groups. */
function champsHistory_(season, me) {
  let c; try { c = JSON.parse(PropertiesService.getScriptProperties().getProperty('CHAMPS_' + season) || 'null'); } catch (e) { c = null; }
  if (!c) return null;
  const out = {};
  (me.isAdmin ? [''] : []).concat(myGroups_(me)).forEach(k => { out[k] = { season: c.season, reg: (c.g || {})[k] || null, post: (c.post || {})[k] || null }; });
  return out;
}
