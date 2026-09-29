/**
 * The Second Seat: sign-up backend.
 *
 * Lives inside John's Google Sheet (Extensions > Apps Script) and is deployed as a
 * web app. The page at PAGE_URL talks to it. The "Games" and "Requests" tabs are the
 * database, so anything the page can do, John can also do by editing those tabs.
 * The "Settings" tab holds the seat location, which is only sent to people with a code.
 *
 * @OnlyCurrentDoc
 */

var PAGE_URL = 'https://darkludus.github.io/second-seat/';

var GAMES_TAB = 'Games';
var REQUESTS_TAB = 'Requests';
var SETTINGS_TAB = 'Settings';
var GAME_COLS = ['id', 'date', 'kickoff', 'opponent', 'tv', 'price', 'note', 'seat_goes_to', 'paid'];
var REQUEST_COLS = ['id', 'game_id', 'name', 'note', 'added', 'token'];
var SETTING_COLS = ['key', 'value'];

var MAX_REQUESTS_PER_GAME = 12;
var MAX_NAME = 40;
var MAX_NOTE = 140;
var MAX_SEAT = 80;
var MAX_BODY = 2000;
var CREW_CODE_LENGTH = 8;
var HOST_CODE_LENGTH = 12;
var CODE_ALPHABET = 'abcdefghjkmnpqrstuvwxyz23456789';

// 2026 home slate. Price is face value per ticket.
var SEED_GAMES = [
  ['wk01', '2026-09-09', '5:20 PM', 'New England Patriots', 'NBC', '269', 'Kickoff Game · Super Bowl LX rematch', 'Dan', ''],
  ['wk04', '2026-10-04', '1:25 PM', 'Los Angeles Chargers', 'CBS', '143', '', '', ''],
  ['wk05', '2026-10-11', '1:25 PM', 'San Francisco 49ers', 'FOX', '206', 'NFC West', '', ''],
  ['wk07', '2026-10-25', '5:20 PM', 'Kansas City Chiefs', 'NBC', '251', 'Sunday Night Football', '', ''],
  ['wk08', '2026-11-02', '5:15 PM', 'Chicago Bears', 'ESPN', '188', 'Monday Night Football', '', ''],
  ['wk09', '2026-11-08', '1:25 PM', 'Arizona Cardinals', 'FOX', '125', 'NFC West', '', ''],
  ['wk13', '2026-12-07', '5:15 PM', 'Dallas Cowboys', 'ESPN', '268', 'Monday Night Football', '', ''],
  ['wk14', '2026-12-13', '1:25 PM', 'New York Giants', 'FOX', '125', '', '', ''],
  ['wk16', '2026-12-25', '5:15 PM', 'Los Angeles Rams', 'FOX', '161', 'Christmas Day · NFC West', '', '']
];

// ---------------------------------------------------------------- web app

function doPost(e) {
  var out;
  try {
    var body = (e && e.postData && e.postData.contents) || '{}';
    out = body.length > MAX_BODY ? fail('too_big') : handle(JSON.parse(body));
  } catch (err) {
    out = { ok: false, error: 'server_error', detail: String((err && err.message) || err) };
  }
  return json(out);
}

function doGet() {
  return json({ ok: true, service: 'second-seat' });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handle(req) {
  var props = PropertiesService.getScriptProperties();
  var crewCode = props.getProperty('CREW_CODE');
  var hostCode = props.getProperty('HOST_CODE');
  if (!crewCode || !hostCode) return fail('not_set_up');

  var isHost = false;
  if (req.host) {
    if (!sameCode(req.host, hostCode)) return fail('bad_host_code');
    isHost = true;
  } else if (!sameCode(req.crew, crewCode)) {
    return fail('bad_crew_code');
  }

  var action = String(req.action || 'state');
  if (action === 'state') return state(isHost);

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var result;
    if (action === 'request') result = addRequest(req);
    else if (action === 'withdraw') result = withdraw(req, isHost);
    else if (!isHost) return fail('host_only');
    else if (action === 'assign') result = assign(req);
    else if (action === 'release') result = setSeat(req.gameId, '', '');
    else if (action === 'paid') result = setPaid(req);
    else return fail('unknown_action');

    if (result && result.ok === false) return result;
    SpreadsheetApp.flush();
    var out = state(isHost);
    if (result && result.mine) out.mine = result.mine;
    return out;
  } finally {
    lock.releaseLock();
  }
}

function fail(code) {
  return { ok: false, error: code };
}

// Compares every character whatever the input, so timing gives nothing away.
function sameCode(given, real) {
  var guess = String(given == null ? '' : given);
  var diff = guess.length ^ real.length;
  for (var i = 0; i < real.length; i++) {
    diff |= (guess.charCodeAt(i) || 0) ^ real.charCodeAt(i);
  }
  return diff === 0;
}

// ---------------------------------------------------------------- reads

function state(isHost) {
  var games = readTable(GAMES_TAB, GAME_COLS).map(function (g) {
    return {
      id: g.id,
      date: isoDate(g.date),
      kickoff: g.kickoff,
      opponent: g.opponent,
      tv: g.tv,
      price: Number(g.price.replace(/[^0-9.]/g, '')) || 0,
      note: g.note,
      holder: g.seat_goes_to,
      paid: /^(y|yes|true|paid)$/i.test(g.paid)
    };
  });
  var requests = readTable(REQUESTS_TAB, REQUEST_COLS).map(function (r) {
    return { id: r.id, gameId: r.game_id, name: r.name, note: r.note, added: r.added };
  });
  return { ok: true, host: isHost, seat: seatLine(), games: games, requests: requests };
}

// Where the seats are. Kept in the Settings tab so it is only ever sent to
// someone who already has a code, and never sits in the public page.
function seatLine() {
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_TAB)) return '';
  var rows = readTable(SETTINGS_TAB, SETTING_COLS);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].key === 'seat') return clean(rows[i].value, MAX_SEAT);
  }
  return '';
}

function readTable(tab, cols) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tab);
  if (!sh) throw new Error('The "' + tab + '" tab is missing. Run Second Seat > Set up.');
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, cols.length).getDisplayValues();
  var rows = [];
  values.forEach(function (row, i) {
    var o = { _row: i + 2 };
    cols.forEach(function (c, j) { o[c] = String(row[j]).trim(); });
    if (o[cols[0]]) rows.push(o);
  });
  return rows;
}

function isoDate(text) {
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  var d = new Date(text);
  if (isNaN(d.getTime())) return '';
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function findGame(gameId) {
  var id = String(gameId || '');
  var games = readTable(GAMES_TAB, GAME_COLS);
  for (var i = 0; i < games.length; i++) if (games[i].id === id) return games[i];
  return null;
}

// ---------------------------------------------------------------- writes

function addRequest(req) {
  var game = findGame(req.gameId);
  if (!game) return fail('no_such_game');
  var name = clean(req.name, MAX_NAME);
  if (!name) return fail('name_required');
  var note = clean(req.note, MAX_NOTE);

  var sameGame = readTable(REQUESTS_TAB, REQUEST_COLS).filter(function (r) {
    return r.game_id === game.id;
  });
  if (sameGame.length >= MAX_REQUESTS_PER_GAME) return fail('list_full');
  var taken = sameGame.some(function (r) { return r.name.toLowerCase() === name.toLowerCase(); });
  if (taken) return fail('already_on_list');

  var id = Utilities.getUuid();
  var token = Utilities.getUuid();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUESTS_TAB);
  var row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, REQUEST_COLS.length).setNumberFormat('@')
    .setValues([[id, game.id, name, note, new Date().toISOString(), token]]);
  return { ok: true, mine: { id: id, token: token } };
}

function withdraw(req, isHost) {
  var id = String(req.requestId || '');
  var rows = readTable(REQUESTS_TAB, REQUEST_COLS);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].id !== id) continue;
    if (!isHost && rows[i].token !== String(req.token || '')) return fail('not_yours');
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUESTS_TAB).deleteRow(rows[i]._row);
    return { ok: true };
  }
  return { ok: true }; // already gone
}

function assign(req) {
  var name = clean(req.name, MAX_NAME);
  if (!name) return fail('name_required');
  return setSeat(req.gameId, name, '');
}

function setSeat(gameId, holder, paid) {
  var game = findGame(gameId);
  if (!game) return fail('no_such_game');
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GAMES_TAB);
  sh.getRange(game._row, GAME_COLS.indexOf('seat_goes_to') + 1, 1, 2).setNumberFormat('@')
    .setValues([[holder, paid]]);
  return { ok: true };
}

function setPaid(req) {
  var game = findGame(req.gameId);
  if (!game) return fail('no_such_game');
  if (!game.seat_goes_to) return fail('seat_is_open');
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GAMES_TAB);
  sh.getRange(game._row, GAME_COLS.indexOf('paid') + 1).setNumberFormat('@')
    .setValue(req.paid ? 'Yes' : '');
  return { ok: true };
}

// Trims, flattens whitespace, and drops a leading = + - @ so a name can never
// turn into a spreadsheet formula.
function clean(value, max) {
  return String(value == null ? '' : value)
    .replace(/[\u0000-\u001f\u007f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^[=+\-@\s]+/, '')
    .slice(0, max)
    .trim();
}

// ---------------------------------------------------------------- setup (run from the sheet)

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Second Seat')
    .addItem('Show links', 'showLinks')
    .addItem('Set up / repair tabs', 'setUp')
    .addItem('New codes (old links stop working)', 'newCodes')
    .addToUi();
}

function setUp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var firstRun = !ss.getSheetByName(GAMES_TAB);

  if (firstRun) {
    var games = makeTab(ss, GAMES_TAB, GAME_COLS);
    games.getRange(2, 1, SEED_GAMES.length, GAME_COLS.length).setValues(SEED_GAMES);
  }
  if (!ss.getSheetByName(REQUESTS_TAB)) makeTab(ss, REQUESTS_TAB, REQUEST_COLS);
  if (!ss.getSheetByName(SETTINGS_TAB)) {
    var settings = makeTab(ss, SETTINGS_TAB, SETTING_COLS);
    settings.getRange(2, 1, 1, 2).setValues([['seat', oldSeatLine(ss)]]);
  }
  if (firstRun) importOldSheet(ss);

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('CREW_CODE')) props.setProperty('CREW_CODE', randomCode(CREW_CODE_LENGTH));
  if (!props.getProperty('HOST_CODE')) props.setProperty('HOST_CODE', randomCode(HOST_CODE_LENGTH));
  showLinks();
}

// Finds the seat location in the old sign-up tab's heading and keeps only the
// section, row and seat parts.
function oldSeatLine(ss) {
  var ours = [GAMES_TAB, REQUESTS_TAB, SETTINGS_TAB];
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    if (ours.indexOf(sh.getName()) >= 0 || sh.getLastRow() < 1 || sh.getLastColumn() < 1) continue;
    var rows = Math.min(sh.getLastRow(), 5);
    var grid = sh.getRange(1, 1, rows, sh.getLastColumn()).getDisplayValues();
    for (var r = 0; r < grid.length; r++) {
      for (var c = 0; c < grid[r].length; c++) {
        var text = String(grid[r][c]);
        if (!/\bsec/i.test(text) || !/\brow\b/i.test(text)) continue;
        var parts = text.split(/[\u00b7|,]/).map(function (p) { return p.trim(); })
          .filter(function (p) { return /^(sec|row|seat)/i.test(p); });
        if (parts.length) return clean(parts.join(' \u00b7 '), MAX_SEAT);
      }
    }
  }
  return '';
}

function makeTab(ss, name, cols) {
  var sh = ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

// Carries names over from the old sign-up tab (the one with OPPONENT and
// WANT IT? ADD YOUR NAME columns), matching games by opponent.
function importOldSheet(ss) {
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    if ([GAMES_TAB, REQUESTS_TAB, SETTINGS_TAB].indexOf(sh.getName()) >= 0) continue;
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 2) continue;
    var grid = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();

    for (var r = 0; r < Math.min(grid.length, 15); r++) {
      var header = grid[r].map(function (v) { return String(v).trim().toUpperCase(); });
      var cOpp = header.indexOf('OPPONENT');
      var cWant = indexStarting(header, 'WANT IT');
      if (cOpp < 0 || cWant < 0) continue;
      var cSeat = header.indexOf('SEAT GOES TO');
      var cPaid = header.indexOf('PAID');
      importRows(grid.slice(r + 1), cOpp, cWant, cSeat, cPaid);
      return;
    }
  }
}

function indexStarting(header, prefix) {
  for (var i = 0; i < header.length; i++) if (header[i].indexOf(prefix) === 0) return i;
  return -1;
}

function importRows(rows, cOpp, cWant, cSeat, cPaid) {
  var games = readTable(GAMES_TAB, GAME_COLS);
  rows.forEach(function (row) {
    var opponent = String(row[cOpp]).trim().toLowerCase();
    if (!opponent) return;
    var game = null;
    games.forEach(function (g) { if (g.opponent.toLowerCase() === opponent) game = g; });
    if (!game) return;

    var holder = cSeat < 0 ? '' : clean(row[cSeat], MAX_NAME);
    var paid = cPaid >= 0 && /^yes$/i.test(String(row[cPaid]).trim());
    if (holder) setSeat(game.id, holder, paid ? 'Yes' : '');

    String(row[cWant]).split(/[,;\n\/]+/).forEach(function (piece) {
      if (clean(piece, MAX_NAME)) addRequest({ gameId: game.id, name: piece, note: '' });
    });
  });
}

function randomCode(length) {
  var out = '';
  for (var i = 0; i < length; i++) {
    out += CODE_ALPHABET.charAt(Math.floor(Math.random() * CODE_ALPHABET.length));
  }
  return out;
}

function newCodes() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('CREW_CODE', randomCode(CREW_CODE_LENGTH));
  props.setProperty('HOST_CODE', randomCode(HOST_CODE_LENGTH));
  showLinks();
}

function showLinks() {
  var props = PropertiesService.getScriptProperties();
  var crew = props.getProperty('CREW_CODE');
  var host = props.getProperty('HOST_CODE');
  if (!crew || !host) return setUp();
  var message =
    'Send your friends this link:\n' + PAGE_URL + '#' + crew +
    '\n\nYour own host link (keep it to yourself):\n' + PAGE_URL + '#' + crew + '.' + host +
    '\n\nBefore you send anything: press Share (top right) and make sure General access ' +
    'says Restricted. Anyone who can edit this sheet can read these codes.';
  Logger.log(message);
  try {
    SpreadsheetApp.getUi().alert('The Second Seat', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    // No sheet window to show it in (run from the editor). The log has it.
  }
}
