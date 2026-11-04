# G1D viktoriin

Exit game for G1D maths lessons: QR code → name + nickname → a 5-minute quiz with four options
per question. Scores go to the teacher's Google Sheet through `kood.gs`; this repo holds only
the game and the questions, never names or scores.

- Game: `index.html?q=<quiz id>` (quizzes in `quizzes/`)
- Projector page with QR code, 5-minute clock and leaderboard: `tabel.html?q=<quiz id>`
- `config.js`: the Apps Script web app URL

Source of truth: `teaching-skills/G1D Lai matemaatika/vahendid/viktoriin/` (private repo);
publish with `avalda.sh`.
