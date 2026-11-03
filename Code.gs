/**
 * Turkey Day Throwdown — Google Apps Script backend.
 * Paste into Extensions > Apps Script of a Google Sheet.
 * Each locked-in bracket is one row on the "Brackets" tab.
 * Re-submitting from the same browser updates that student's row.
 */

const SHEET_NAME = 'Brackets';
const HEADERS = ['Last updated', 'Player ID', 'Name', 'Champion', 'Final Four', 'Picks (data)'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Turkey Day Throwdown')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // lets you embed it in Google Sites
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    const first = ss.getSheets()[0];
    if (first && first.getRange(1, 2).getValue() === 'Player ID') {
      first.setName(SHEET_NAME);
      return first;
    }
  }
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

/** Saves or updates one bracket, then returns everyone's results. */
function submitBracket(b) {
  if (!b || typeof b !== 'object' || typeof b.id !== 'string' || !b.picks || typeof b.picks !== 'object') {
    throw new Error('Invalid bracket');
  }
  const id = String(b.id).slice(0, 40);
  const name = String(b.name || '').slice(0, 40);
  const picks = {};
  Object.keys(b.picks).slice(0, 60).forEach(function (k) {
    picks[String(k).slice(0, 10)] = String(b.picks[k]).slice(0, 20);
  });
  const champion = String(b.championName || '').slice(0, 60);
  const finalFour = (b.finalFourNames || []).slice(0, 4).map(function (x) { return String(x).slice(0, 60); }).join(' · ');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const row = [new Date(), id, name, champion, finalFour, JSON.stringify(picks)];
    const last = sh.getLastRow();
    let target = 0;
    if (last > 1) {
      const ids = sh.getRange(2, 2, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === id) { target = i + 2; break; }
      }
    }
    if (target) sh.getRange(target, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return getBrackets();
}

/** Returns every bracket's picks (no names) for the Community Results tab. */
function getBrackets() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 6, last - 1, 1).getValues();
  const out = [];
  rows.forEach(function (r) {
    try {
      const picks = JSON.parse(r[0]);
      if (picks && picks.c) {
        out.push({ picks: picks, champion: picks.c, finalFour: [picks.e0, picks.e1, picks.e2, picks.e3] });
      }
    } catch (e) { /* skip rows that were edited by hand */ }
  });
  return out;
}
