// G1D tunnikontroll — Google Apps Script behind the game (Aivar's Google Sheet "G1D viktoriin").
// Paste into Extensions ▸ Apps Script of that sheet, then Deploy ▸ New deployment ▸ Web app,
// Execute as: Me, Who has access: Anyone. The /exec URL goes into config.js.
//
// One shared deadline per lesson (Aivar 4.11.2026): the projector page opens the lesson's quiz for
// everyone at once ({action: "open"} with the teacher key), and from then on there are `minutes`
// minutes. A game cannot start before or after, and a result that arrives after the deadline is
// refused, so a finished student cannot play again later for someone else.
// The teacher key is made on first use and sits in the sheet's "Seaded" tab, cell B1 — this file is
// public, so the key must never be written here.
//
// POST {action: "open", lesson, key, minutes}   open the quiz now, for `minutes` (1–15)
// POST {lesson, name, nick, score, correct, total, seconds, timeout}   one finished game; only the
//      first per lesson and name counts, a repeat is answered ok + duplicate
// GET  ?q=<lesson>   state {now, until} and the leaderboard: nicknames and scores only, never names

const SHEET = "Tulemused";
const HEAD = ["Aeg", "Tund", "Nimi", "Hüüdnimi", "Punktid", "Õiged", "Küsimusi", "Sekundid", "Aeg otsas"];
const GRACE = 30000;            // a result may arrive this long after the deadline (slow network)

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}

function key_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Seaded");
  if (!sh) {
    sh = ss.insertSheet("Seaded");
    sh.getRange("A1:B1").setValues([["Õpetaja võti", Utilities.getUuid().slice(0, 8)]]);
  }
  return String(sh.getRange("B1").getValue()).trim();
}

const until_ = lesson => Number(PropertiesService.getScriptProperties().getProperty("until_" + lesson) || 0);

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean_(v, max) { return String(v == null ? "" : v).replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const lesson = clean_(d.lesson, 40);
    if (!lesson) return out_({ok: false, error: "missing lesson"});
    if (d.action === "open") {
      if (!d.key || clean_(d.key, 40) !== key_()) return out_({ok: false, error: "wrong key"});
      const minutes = Math.max(1, Math.min(15, Number(d.minutes) || 5));
      const until = Date.now() + minutes * 60000;
      PropertiesService.getScriptProperties().setProperty("until_" + lesson, String(until));
      return out_({ok: true, now: Date.now(), until: until});
    }
    const until = until_(lesson);
    if (!until) return out_({ok: false, error: "not open"});
    if (Date.now() > until + GRACE) return out_({ok: false, error: "closed"});
    const name = clean_(d.name, 60), nick = clean_(d.nick, 20);
    if (name.length < 3 || nick.length < 2) return out_({ok: false, error: "missing fields"});
    const sh = sheet_();
    const rows = sh.getDataRange().getValues().slice(1);
    const k = name.toLowerCase();
    if (rows.some(r => r[1] === lesson && String(r[2]).toLowerCase() === k)) return out_({ok: true, duplicate: true});
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
  return out_({ok: true, lesson: lesson, now: Date.now(), until: until_(lesson),
               played: board.length, board: board.slice(0, 30)});
}
