// Stand-ins for the Google Apps Script services Code.gs uses, so the real backend
// can run in a browser tab against a fake spreadsheet. Test use only.
(() => {
  'use strict';
  const KEY = 'secondseat.mockbook';

  const legacy = () => {
    const rows = [];
    rows.push(['THE SECOND SEAT']);
    rows.push(['LUMEN FIELD  \u00b7  SECTION 100  \u00b7  ROW A  \u00b7  SEATS 1-2  \u00b7  2026 HOME SLATE']);
    for (let i = 0; i < 3; i++) rows.push(['']);
    rows.push(['WEEK', 'DATE', 'DAY', 'KICKOFF (PT)', 'OPPONENT', 'TV', 'PRICE / TICKET', 'GAME NOTE',
      'STATUS', 'SEAT GOES TO', 'WANT IT? ADD YOUR NAME', 'PAID', 'NOTES']);
    rows.push(['WK 1', 'Sep 9', 'WED', '5:20 PM', 'NEW ENGLAND PATRIOTS', 'NBC', '$269', '', 'Locked in', 'Dan', '', 'Yes', '']);
    rows.push(['WK 4', 'Oct 4', 'SUN', '1:25 PM', 'LOS ANGELES CHARGERS', 'CBS', '$143', '', 'Requested', '', 'Mike, Sarah', '', '']);
    rows.push(['WK 5', 'Oct 11', 'SUN', '1:25 PM', 'SAN FRANCISCO 49ERS', 'FOX', '$206', '', 'Open', '', '=Pete', '', '']);
    rows.push(['WK 7', 'Oct 25', 'SUN', '5:20 PM', 'KANSAS CITY CHIEFS', 'NBC', '$251', '', 'Locked in', 'Priya', 'Priya; Tom', 'No', '']);
    return rows;
  };

  let book = null;
  try { book = JSON.parse(localStorage.getItem(KEY)); } catch (e) { book = null; }
  if (!book) book = { sheets: [{ name: 'Second Seat 2026', cells: legacy() }], props: {} };
  const save = () => { try { localStorage.setItem(KEY, JSON.stringify(book)); } catch (e) { /* fine */ } };

  const cell = (sheet, r, c) => {
    const row = sheet.cells[r - 1];
    const v = row ? row[c - 1] : '';
    return v == null ? '' : String(v);
  };
  const put = (sheet, r, c, v) => {
    while (sheet.cells.length < r) sheet.cells.push([]);
    const row = sheet.cells[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v == null ? '' : String(v);
  };

  function range(sheet, r, c, nr, nc) {
    nr = nr || 1; nc = nc || 1;
    const api = {
      getDisplayValues() {
        const out = [];
        for (let i = 0; i < nr; i++) {
          const row = [];
          for (let j = 0; j < nc; j++) row.push(cell(sheet, r + i, c + j));
          out.push(row);
        }
        return out;
      },
      setValues(values) {
        if (values.length !== nr || values.some(v => v.length !== nc)) {
          throw new Error('mock: setValues size does not match the range');
        }
        values.forEach((row, i) => row.forEach((v, j) => put(sheet, r + i, c + j, v)));
        save();
        return api;
      },
      setValue(v) { put(sheet, r, c, v); save(); return api; },
      setNumberFormat() { return api; },
      setFontWeight() { return api; }
    };
    return api;
  }

  function wrap(sheet) {
    return {
      getName: () => sheet.name,
      getMaxRows: () => 1000,
      getLastRow() {
        for (let i = sheet.cells.length; i > 0; i--) {
          if ((sheet.cells[i - 1] || []).some(v => v !== '' && v != null)) return i;
        }
        return 0;
      },
      getLastColumn: () => sheet.cells.reduce((m, row) => Math.max(m, row.length), 0),
      getRange: (r, c, nr, nc) => range(sheet, r, c, nr, nc),
      deleteRow(r) { sheet.cells.splice(r - 1, 1); save(); },
      setFrozenRows() {}
    };
  }

  const spreadsheet = {
    getSheetByName(name) {
      const s = book.sheets.find(x => x.name === name);
      return s ? wrap(s) : null;
    },
    getSheets: () => book.sheets.map(wrap),
    insertSheet(name) {
      const s = { name: name, cells: [] };
      book.sheets.push(s);
      save();
      return wrap(s);
    },
    getSpreadsheetTimeZone: () => 'America/Los_Angeles'
  };

  const menu = { addItem: () => menu, addToUi() {} };
  window.SpreadsheetApp = {
    getActiveSpreadsheet: () => spreadsheet,
    flush() {},
    getUi: () => ({ ButtonSet: { OK: 'OK' }, createMenu: () => menu,
      alert: (title, message) => console.log('[alert] ' + title + '\n' + message) })
  };
  window.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: k => (k in book.props ? book.props[k] : null),
      setProperty(k, v) { book.props[k] = String(v); save(); }
    })
  };
  window.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  window.ContentService = {
    MimeType: { JSON: 'application/json' },
    createTextOutput: text => ({ setMimeType() { return this; }, getContent: () => text })
  };
  window.Utilities = {
    getUuid: () => crypto.randomUUID(),
    formatDate: d => new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Los_Angeles',
      year: 'numeric', month: '2-digit', day: '2-digit' }).format(d)
  };
  window.Logger = { log: m => console.log('[log] ' + m) };

  window.__mock = {
    book: () => book,
    reset() { localStorage.removeItem(KEY); location.reload(); }
  };
})();
