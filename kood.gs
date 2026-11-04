// G1D viktoriin — Google Apps Script behind the game (Aivar's Google Sheet "G1D viktoriin").
// Paste into Extensions ▸ Apps Script of that sheet, then Deploy ▸ New deployment ▸ Web app,
// Execute as: Me, Who has access: Anyone. The /exec URL goes into config.js.
//
// POST  one finished game: {lesson, name, nick, score, correct, total, seconds, timeout}.
//       Only the first game per lesson and name counts; a repeat is answered ok + duplicate.
// GET   ?q=<lesson>: the leaderboard for that lesson, nicknames and scores only, never names.

const SHEET = "Tulemused";
const HEAD = ["Aeg", "Tund", "Nimi", "Hüüdnimi", "Punktid", "Õiged", "Küsimusi", "Sekundid", "Aeg otsas"];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean_(v, max) { return String(v == null ? "" : v).replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const lesson = clean_(d.lesson, 40), name = clean_(d.name, 60), nick = clean_(d.nick, 20);
    if (!lesson || name.length < 3 || nick.length < 2) return out_({ok: false, error: "missing fields"});
    const sh = sheet_();
    const rows = sh.getDataRange().getValues().slice(1);
    const key = name.toLowerCase();
    if (rows.some(r => r[1] === lesson && String(r[2]).toLowerCase() === key)) return out_({ok: true, duplicate: true});
    const num = (v, lo, hi) => Math.max(lo, Math.min(hi, Math.round(Number(v) || 0)));
    sh.appendRow([new Date(), lesson, name, nick, num(d.score, 0, 10000), num(d.correct, 0, 50),
                  num(d.total, 0, 50), num(d.seconds, 0, 3600), d.timeout ? "jah" : ""]);
    return out_({ok: true});
  } catch (err) {
    return out_({ok: false, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const lesson = clean_((e.parameter || {}).q, 40);
  const rows = sheet_().getDataRange().getValues().slice(1).filter(r => r[1] === lesson);
  const board = rows.map(r => ({nick: r[3], score: r[4], correct: r[5], total: r[6]}))
                    .sort((a, b) => b.score - a.score);
  return out_({ok: true, lesson: lesson, played: board.length, board: board.slice(0, 30)});
}
