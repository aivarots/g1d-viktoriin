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
// POST {lesson, name, nick, score, correct, total, seconds, timeout, away, awaySeconds}   one finished
//      game; only the first per lesson and name counts, a repeat is answered ok + duplicate.
//      away / awaySeconds (Aivar 6.11.2026): how often and how long the game page was in the
//      background — for Aivar to judge case by case, never a penalty by itself
//      picks (Aivar 10.11.2026): [[question, version, option], …], option 0 = right; written to the
//      "Vastused" tab WITHOUT name or nickname — only for finding the commonest error
// GET  ?q=<lesson>   state {now, until} and the leaderboard: nicknames and scores only, never names
// GET  ?stats=<lesson>   once the quiz is closed: counts per [question, version, option], no names
//      (read by skriptid/tunnikontroll_vead.py for the next lesson's opener)

const SHEET = "Tulemused";
const HEAD = ["Aeg", "Tund", "Nimi", "Hüüdnimi", "Punktid", "Õiged", "Küsimusi", "Sekundid", "Aeg otsas",
              "Lahkus lehelt", "Eemal (s)"];
const GRACE = 30000;            // a result may arrive this long after the deadline (slow network)

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  else if (sh.getRange(1, HEAD.length).getValue() !== HEAD[HEAD.length - 1])     // columns added 6.11.2026
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
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

// Test dates in "Seaded" (Aivar can change them): games before KT1 count for KT1, games from KT1 up
// to KT2 for KT2. Lessons "test" and "proov" never count.
const TESTS = [["KT1 kuupäev", new Date(2026, 10, 17)], ["KT2 kuupäev", new Date(2026, 11, 1)]];
const NOT_COUNTED = ["test", "proov"];
const BONUS_MIN = 4, BONUS_EACH = 3, BONUS_MAX = 15;   // Aivar 4.11.2026: +3 per game with ≥ 4/6, max +15

function dates_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Seaded") || (key_(), SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Seaded"));
  TESTS.forEach(([label, d], i) => {
    const r = sh.getRange(i + 2, 1, 1, 2);
    if (!r.getValues()[0][0]) r.setValues([[label, d]]);
  });
  return TESTS.map((_, i) => new Date(sh.getRange(i + 2, 2).getValue()));
}

// "Kokkuvõte": one row per student (by name), games and test bonus per KT block, all points.
// Rebuilt after every saved game, and by hand with Run ▸ kokkuvote.
function kokkuvote() {
  const [kt1, kt2] = dates_();
  const rows = sheet_().getDataRange().getValues().slice(1).filter(r => r[2] && !NOT_COUNTED.includes(r[1]));
  const by = {};
  rows.forEach(r => {
    const k = String(r[2]).trim().toLowerCase();
    const s = by[k] || (by[k] = {name: String(r[2]).trim(), n: [0, 0], good: [0, 0], points: 0, away: 0});
    const t = new Date(r[0]), block = t < kt1 ? 0 : t < kt2 ? 1 : -1;
    s.points += Number(r[4]) || 0;
    s.away += Number(r[9]) || 0;
    if (block < 0) return;
    s.n[block]++;
    if (Number(r[5]) >= BONUS_MIN) s.good[block]++;
  });
  const bonus = g => Math.min(BONUS_MAX, BONUS_EACH * g);
  const out = Object.values(by).sort((a, b) => a.name.localeCompare(b.name, "et"))
    .map(s => [s.name, s.n[0], s.good[0], bonus(s.good[0]), s.n[1], s.good[1], bonus(s.good[1]), s.points, s.away]);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Kokkuvõte") || ss.insertSheet("Kokkuvõte");
  sh.clearContents();
  sh.getRange(1, 1, 1, 9).setValues([["Nimi", "KT1: mänge", "KT1: ≥ 4/6", "KT1 boonus", "KT2: mänge", "KT2: ≥ 4/6", "KT2 boonus", "Punktid kokku", "Lahkus lehelt kokku"]])
    .setFontWeight("bold");
  sh.setFrozenRows(1);
  if (out.length) sh.getRange(2, 1, out.length, 9).setValues(out);
}

// One student's bonus in the current KT block, sent back after a saved game so the phone can show it.
function bonusFor_(name) {
  const [kt1, kt2] = dates_(), now = new Date();
  const block = now < kt1 ? 0 : now < kt2 ? 1 : -1;
  if (block < 0) return null;
  const k = String(name).trim().toLowerCase();
  const good = sheet_().getDataRange().getValues().slice(1).filter(r => {
    const t = new Date(r[0]);
    return String(r[2]).trim().toLowerCase() === k && !NOT_COUNTED.includes(r[1]) &&
           (block === 0 ? t < kt1 : t >= kt1 && t < kt2) && Number(r[5]) >= BONUS_MIN;
  }).length;
  return {test: block === 0 ? "KT1" : "KT2", good: good, bonus: Math.min(BONUS_MAX, BONUS_EACH * good), max: BONUS_MAX};
}

// Run once from the editor (Run ▸ seadista): asks for permissions, makes the tabs, shows the key.
function seadista() {
  sheet_();
  const k = key_();
  dates_();
  kokkuvote();
  Logger.log("Valmis. Õpetaja võti: " + k + " (sama on lehel Seaded, lahter B1)");
  return k;
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
    if (rows.some(r => r[1] === lesson && String(r[2]).toLowerCase() === k)) return out_({ok: true, duplicate: true, me: bonusFor_(name)});
    const num = (v, lo, hi) => Math.max(lo, Math.min(hi, Math.round(Number(v) || 0)));
    sh.appendRow([new Date(), lesson, name, nick, num(d.score, 0, 10000), num(d.correct, 0, 50),
                  num(d.total, 0, 50), num(d.seconds, 0, 3600), d.timeout ? "jah" : "",
                  d.away == null ? "" : num(d.away, 0, 999), d.awaySeconds == null ? "" : num(d.awaySeconds, 0, 3600)]);
    try { picks_(lesson, d.picks); } catch (err) {} // nor the per-answer record
    try { kokkuvote(); } catch (err) {}             // the summary must never cost a saved result
    let me = null;
    try { me = bonusFor_(name); } catch (err) {}
    return out_({ok: true, me: me});
  } catch (err) {
    return out_({ok: false, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}

const PICKS = "Vastused";
function picks_(lesson, p) {
  if (!Array.isArray(p) || !p.length) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PICKS);
  if (!sh) { sh = ss.insertSheet(PICKS); sh.appendRow(["Aeg", "Tund", "Küsimus", "Versioon", "Valik (0 = õige)"]); sh.setFrozenRows(1); }
  const n = v => Math.max(0, Math.min(20, Math.round(Number(v) || 0)));
  const now = new Date();
  const rows = p.slice(0, 20).filter(a => Array.isArray(a) && a.length === 3).map(a => [now, lesson, n(a[0]), n(a[1]), n(a[2])]);
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
}

function stats_(lesson) {
  const until = until_(lesson);
  if (until && Date.now() <= until + GRACE) return out_({ok: false, error: "still open"});
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PICKS);
  const count = {};
  if (sh) sh.getDataRange().getValues().slice(1).filter(r => r[1] === lesson).forEach(r => {
    const k = r[2] + "," + r[3] + "," + r[4]; count[k] = (count[k] || 0) + 1;
  });
  const games = sheet_().getDataRange().getValues().slice(1).filter(r => r[1] === lesson).length;
  return out_({ok: true, lesson: lesson, games: games,
               picks: Object.keys(count).map(k => k.split(",").map(Number).concat([count[k]]))});
}

function doGet(e) {
  if ((e.parameter || {}).stats) return stats_(clean_(e.parameter.stats, 40));
  const lesson = clean_((e.parameter || {}).q, 40);
  const rows = sheet_().getDataRange().getValues().slice(1).filter(r => r[1] === lesson);
  const board = rows.map(r => ({nick: r[3], score: r[4], correct: r[5], total: r[6]}))
                    .sort((a, b) => b.score - a.score);
  return out_({ok: true, lesson: lesson, now: Date.now(), until: until_(lesson),
               played: board.length, board: board.slice(0, 30)});
}
