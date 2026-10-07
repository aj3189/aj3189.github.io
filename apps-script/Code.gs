/**
 * NFL ATS picks backend. Bind this script to the picks spreadsheet and deploy
 * it as a web app (Execute as: Me, Who has access: Anyone). See README.md.
 *
 * Pick format: each player submits a number, the home team's spread as they
 * see it. Compared to the frozen DraftKings line (home_spread):
 *   guess < line  -> player takes HOME at the line
 *   guess > line  -> player takes AWAY at the line
 *   guess == line -> no pick
 */

// All tabs this script touches get this prefix. Other tabs are never modified.
// Change to '' once you're ready to go live with tabs named Games, Picks, etc.
var TAB_PREFIX = 'TEST_';

// Tried in order. site.api.espn.com returns 403 to Apps Script's user agent, so the
// site.web mirror (same JSON) comes first.
var ESPN_URLS = [
  'https://site.web.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard',
  'https://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard'
];
var SEASON_TYPE = 2; // regular season

var HEADERS = {
  Games: ['week', 'game_id', 'kickoff_utc', 'away_team', 'home_team', 'home_spread',
          'away_score', 'home_score', 'status', 'ats_winner'],
  Picks: ['submitted_at', 'user', 'week', 'game_id', 'guess', 'pick'],
  Users: ['user_slug', 'display_name'],
  Stats: ['user', 'display_name', 'scope', 'wins', 'losses', 'pushes', 'win_pct']
};

// ---------------------------------------------------------------------------
// Setup helpers (run manually from the Apps Script editor)
// ---------------------------------------------------------------------------

/** Creates the four tabs with headers if they don't exist. Never touches other tabs. */
function setupSheet() {
  Object.keys(HEADERS).forEach(function (name) {
    var sh = getSheet_(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
      sh.setFrozenRows(1);
    }
  });
  // Keep game ids and kickoff times as plain text so Sheets doesn't reformat them.
  getSheet_('Games').getRange('B:C').setNumberFormat('@');
  getSheet_('Picks').getRange('D:D').setNumberFormat('@');
}

/** Adds the players to the Users tab (skips any already present). Edit the list as needed. */
function seedUsers() {
  var users = [
    ['andrew', 'Andrew'],
    ['christian', 'Christian']
  ];
  var sh = getSheet_('Users');
  var existing = readTable_('Users').map(function (u) { return String(u.user_slug); });
  users.forEach(function (u) {
    if (existing.indexOf(u[0]) === -1) sh.appendRow(u);
  });
}

/** Replaces this project's triggers with the schedule below (all times Eastern). */
function setupTriggers() {
  var fns = ['fetchOdds', 'fetchResults'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (fns.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  var tz = 'America/New_York';
  // Lines: Wednesday 10am.
  ScriptApp.newTrigger('fetchOdds').timeBased()
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(10).inTimezone(tz).create();
  // Results: daily 9am...
  ScriptApp.newTrigger('fetchResults').timeBased()
    .everyDays(1).atHour(9).inTimezone(tz).create();
  // ...plus after Thursday, Sunday and Monday night games (1-2am the next morning).
  [ScriptApp.WeekDay.FRIDAY, ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY].forEach(function (day) {
    ScriptApp.newTrigger('fetchResults').timeBased()
      .onWeekDay(day).atHour(1).inTimezone(tz).create();
  });
}

// ---------------------------------------------------------------------------
// Web app
// ---------------------------------------------------------------------------

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'games') return json_(getGames_(p.week));
    if (p.action === 'picks') {
      if (!findUser_(p.user)) return json_({ error: 'Unknown user' });
      return json_({ picks: getPicks_(p.user, Number(p.week)) });
    }
    if (p.action === 'stats') return json_({ stats: readTable_('Stats') });
    return json_({ error: 'Unknown action' });
  } catch (err) {
    console.error(err);
    return json_({ error: String(err) });
  }
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ error: 'Invalid JSON' });
  }
  if (!findUser_(body.user)) return json_({ error: 'Unknown user' });
  if (!Array.isArray(body.picks)) return json_({ error: 'Missing picks' });

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ error: 'Busy, try again' });
  try {
    return json_(savePicks_(String(body.user), Number(body.week), body.picks));
  } catch (err) {
    console.error(err);
    return json_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getGames_(weekParam) {
  var games = readTable_('Games').map(normalizeGame_);
  var weeks = [];
  games.forEach(function (g) { if (weeks.indexOf(g.week) === -1) weeks.push(g.week); });
  weeks.sort(function (a, b) { return a - b; });

  var week = Number(weekParam);
  if (!week) {
    // Current week: earliest week with an unfinished game, else the latest week.
    var open = games.filter(function (g) { return !isFinal_(g.status); })
                    .map(function (g) { return g.week; });
    week = open.length ? Math.min.apply(null, open) : (weeks.length ? weeks[weeks.length - 1] : null);
  }
  var list = games.filter(function (g) { return g.week === week; })
                  .sort(function (a, b) { return a.kickoff_utc < b.kickoff_utc ? -1 : 1; });
  return { week: week, weeks: weeks, games: list };
}

function getPicks_(user, week) {
  var games = {};
  readTable_('Games').map(normalizeGame_).forEach(function (g) { games[g.game_id] = g; });
  return readTable_('Picks')
    .filter(function (r) { return String(r.user) === user && Number(r.week) === week; })
    .map(function (r) {
      var g = games[String(r.game_id)];
      var guess = Number(r.guess);
      var pick = g ? sideFor_(guess, g.home_spread) : String(r.pick);
      return {
        game_id: String(r.game_id),
        guess: guess,
        pick: pick,
        result: g ? grade_(pick, g.ats_winner) : ''
      };
    });
}

function savePicks_(user, week, picks) {
  var games = {};
  readTable_('Games').map(normalizeGame_).forEach(function (g) { games[g.game_id] = g; });

  var sh = getSheet_('Picks');
  var data = sh.getDataRange().getValues();
  var rowByGame = {}; // game_id -> 1-based sheet row for this user/week
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === user && Number(data[i][2]) === week) rowByGame[String(data[i][3])] = i + 1;
  }

  var now = new Date();
  var saved = [], rejected = [], toDelete = [];
  picks.forEach(function (p) {
    var id = String(p.game_id);
    var g = games[id];
    if (!g || g.week !== week) return rejected.push({ game_id: id, reason: 'Unknown game' });
    if (new Date(g.kickoff_utc) <= now) return rejected.push({ game_id: id, reason: 'Game has started' });
    if (g.home_spread === null) return rejected.push({ game_id: id, reason: 'No line yet' });

    var row = rowByGame[id];
    if (p.guess === '' || p.guess === null || p.guess === undefined) {
      if (row) toDelete.push(row);
      saved.push({ game_id: id, guess: null, pick: '' });
      return;
    }
    var guess = Number(p.guess);
    if (!isFinite(guess) || Math.abs(guess) > 60) return rejected.push({ game_id: id, reason: 'Invalid number' });

    var pick = sideFor_(guess, g.home_spread);
    var values = [[now.toISOString(), user, week, id, guess, pick]];
    if (row) sh.getRange(row, 1, 1, values[0].length).setValues(values);
    else sh.appendRow(values[0]);
    saved.push({ game_id: id, guess: guess, pick: pick });
  });
  toDelete.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  return { saved: saved, rejected: rejected };
}

// ---------------------------------------------------------------------------
// ESPN sync
// ---------------------------------------------------------------------------

/** Adds this week's (and next week's) games and freezes each line on first write. */
function fetchOdds() {
  var current = fetchScoreboard_();
  if (!current) return;
  var week = current.week && current.week.number;
  syncEvents_(current.events, week, { insert: true });
  var next = fetchScoreboard_(week + 1);
  if (next) syncEvents_(next.events, week + 1, { insert: true });
}

/**
 * Updates scores, status and ats_winner for games already in the sheet, then rebuilds Stats.
 * Also fills in a line that is still blank for a game that hasn't started; it never changes an existing line.
 */
function fetchResults() {
  var current = fetchScoreboard_();
  if (!current) return;
  var week = current.week && current.week.number;
  syncEvents_(current.events, week, { insert: false });
  if (week > 1) {
    var prev = fetchScoreboard_(week - 1);
    if (prev) syncEvents_(prev.events, week - 1, { insert: false });
  }
  rebuildStats();
}

function fetchScoreboard_(week) {
  var query = '?seasontype=' + SEASON_TYPE + (week ? '&week=' + week : '');
  for (var i = 0; i < ESPN_URLS.length; i++) {
    var url = ESPN_URLS[i] + query;
    try {
      var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) {
        console.error('ESPN ' + res.getResponseCode() + ' for ' + url);
        continue;
      }
      var data = JSON.parse(res.getContentText());
      if (!data || !Array.isArray(data.events)) {
        console.error('ESPN response missing events for ' + url);
        continue;
      }
      return data;
    } catch (err) {
      console.error('ESPN fetch failed for ' + url + ': ' + err);
    }
  }
  return null;
}

function syncEvents_(events, week, opts) {
  if (!week || !events.length) return;
  var sh = getSheet_('Games');
  var data = sh.getDataRange().getValues();
  var col = indexOf_(HEADERS.Games);
  var rowById = {};
  for (var i = 1; i < data.length; i++) rowById[String(data[i][col.game_id])] = i;

  var now = new Date();
  events.forEach(function (ev) {
    var parsed = parseEvent_(ev);
    if (!parsed) return;
    var i = rowById[parsed.game_id];
    var row;
    if (i === undefined) {
      if (!opts.insert) return;
      row = HEADERS.Games.map(function () { return ''; });
      row[col.week] = week;
      row[col.game_id] = parsed.game_id;
      row[col.away_team] = parsed.away_team;
      row[col.home_team] = parsed.home_team;
      data.push(row);
      rowById[parsed.game_id] = data.length - 1;
    } else {
      row = data[i];
    }
    row[col.kickoff_utc] = parsed.kickoff_utc;
    row[col.status] = parsed.status;
    // Freeze: only set the line if blank, and only before kickoff.
    if (row[col.home_spread] === '' && parsed.home_spread !== null && new Date(parsed.kickoff_utc) > now) {
      row[col.home_spread] = parsed.home_spread;
    }
    if (isFinal_(parsed.status) && parsed.home_score !== null && parsed.away_score !== null) {
      row[col.away_score] = parsed.away_score;
      row[col.home_score] = parsed.home_score;
      row[col.ats_winner] = row[col.home_spread] === '' ? ''
        : atsWinner_(parsed.home_score, parsed.away_score, Number(row[col.home_spread]));
    }
  });

  if (data.length > 1) {
    sh.getRange(2, 1, data.length - 1, HEADERS.Games.length).setValues(data.slice(1));
  }
}

function parseEvent_(ev) {
  try {
    var c = ev.competitions[0];
    var home = c.competitors.filter(function (t) { return t.homeAway === 'home'; })[0];
    var away = c.competitors.filter(function (t) { return t.homeAway === 'away'; })[0];
    var homeAbbr = home.team.abbreviation, awayAbbr = away.team.abbreviation;
    var status = c.status.type.name;
    return {
      game_id: String(ev.id),
      kickoff_utc: new Date(ev.date).toISOString(),
      away_team: awayAbbr,
      home_team: homeAbbr,
      home_spread: parseSpread_(c.odds, homeAbbr, awayAbbr),
      away_score: away.score === undefined || away.score === '' ? null : Number(away.score),
      home_score: home.score === undefined || home.score === '' ? null : Number(home.score),
      status: status
    };
  } catch (err) {
    console.error('Could not parse ESPN event ' + (ev && ev.id) + ': ' + err);
    return null;
  }
}

/** Home team's spread from ESPN odds, preferring DraftKings. Returns null if unavailable. */
function parseSpread_(odds, homeAbbr, awayAbbr) {
  if (!odds || !odds.length) return null;
  var o = odds.filter(function (x) { return x.provider && x.provider.name === 'DraftKings'; })[0] || odds[0];
  // details looks like "DAL -8.5" (favorite and its spread) or "EVEN".
  var m = /^([A-Z]+)\s+(-?\d+(\.\d+)?)$/.exec(String(o.details || '').trim());
  if (m) {
    var n = Math.abs(Number(m[2]));
    if (m[1] === homeAbbr) return -n;
    if (m[1] === awayAbbr) return n;
  }
  if (/^(EVEN|PK|PICK)/i.test(String(o.details || ''))) return 0;
  // ESPN's `spread` field is the home team's spread.
  return typeof o.spread === 'number' ? o.spread : null;
}

// ---------------------------------------------------------------------------
// Grading and stats
// ---------------------------------------------------------------------------

function atsWinner_(homeScore, awayScore, homeSpread) {
  var adj = homeScore - awayScore + homeSpread;
  return adj > 0 ? 'home' : adj < 0 ? 'away' : 'push';
}

function sideFor_(guess, homeSpread) {
  if (homeSpread === null || homeSpread === '' || !isFinite(guess)) return '';
  return guess < homeSpread ? 'home' : guess > homeSpread ? 'away' : 'none';
}

function grade_(pick, atsWinner) {
  if (!atsWinner || (pick !== 'home' && pick !== 'away')) return '';
  return atsWinner === 'push' ? 'P' : atsWinner === pick ? 'W' : 'L';
}

/** Rewrites the Stats tab: one season row per user plus one row per user per week. */
function rebuildStats() {
  var games = {};
  readTable_('Games').map(normalizeGame_).forEach(function (g) { games[g.game_id] = g; });
  var users = readTable_('Users');
  var tally = {}; // user -> scope -> {w,l,p}

  readTable_('Picks').forEach(function (r) {
    var g = games[String(r.game_id)];
    if (!g) return;
    var res = grade_(sideFor_(Number(r.guess), g.home_spread), g.ats_winner);
    if (!res) return;
    var u = String(r.user);
    ['season', 'week ' + g.week].forEach(function (scope) {
      tally[u] = tally[u] || {};
      var t = tally[u][scope] = tally[u][scope] || { W: 0, L: 0, P: 0 };
      t[res]++;
    });
  });

  var rows = [];
  users.forEach(function (u) {
    var slug = String(u.user_slug);
    var scopes = tally[slug] || { season: { W: 0, L: 0, P: 0 } };
    Object.keys(scopes).sort(function (a, b) {
      if (a === 'season') return -1;
      if (b === 'season') return 1;
      return Number(a.split(' ')[1]) - Number(b.split(' ')[1]);
    }).forEach(function (scope) {
      var t = scopes[scope];
      var pct = t.W + t.L ? Math.round(1000 * t.W / (t.W + t.L)) / 1000 : '';
      rows.push([slug, u.display_name, scope, t.W, t.L, t.P, pct]);
    });
  });

  var sh = getSheet_('Stats');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.Stats.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, HEADERS.Stats.length).setValues(rows);
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var full = TAB_PREFIX + name;
  return ss.getSheetByName(full) || ss.insertSheet(full);
}

function readTable_(name) {
  var values = getSheet_(name).getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0];
  return values.slice(1).map(function (row) {
    var o = {};
    head.forEach(function (h, i) { o[h] = row[i]; });
    return o;
  });
}

function normalizeGame_(g) {
  return {
    week: Number(g.week),
    game_id: String(g.game_id),
    kickoff_utc: g.kickoff_utc instanceof Date ? g.kickoff_utc.toISOString() : String(g.kickoff_utc),
    away_team: String(g.away_team),
    home_team: String(g.home_team),
    home_spread: g.home_spread === '' ? null : Number(g.home_spread),
    away_score: g.away_score === '' ? null : Number(g.away_score),
    home_score: g.home_score === '' ? null : Number(g.home_score),
    status: String(g.status),
    ats_winner: String(g.ats_winner)
  };
}

function findUser_(slug) {
  if (!slug) return null;
  return readTable_('Users').filter(function (u) { return String(u.user_slug) === String(slug); })[0] || null;
}

function isFinal_(status) {
  return /^STATUS_FINAL/.test(String(status));
}

function indexOf_(headers) {
  var o = {};
  headers.forEach(function (h, i) { o[h] = i; });
  return o;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
