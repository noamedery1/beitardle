/* ============================================================
   ביתרדל — שרת סטטיסטיקה
   ------------------------------------------------------------
   התקנה:
   1. sheets.new  → צור גיליון חדש בשם "ביתרדל סטטיסטיקה"
   2. תפריט Extensions → Apps Script
   3. מחק את הקוד שיש שם, הדבק את כל הקובץ הזה
   4. Deploy → New deployment → Type: Web app
        Execute as:      Me
        Who has access:  Anyone
   5. העתק את ה-URL שמתקבל (מסתיים ב-/exec)
      והדבק אותו ב-index.html בשדה ANALYTICS_URL

   לא נשמר שום מזהה אישי — רק מספר חידה, תוצאה, וחותמת זמן.
   ============================================================ */

const SHEET_EVENTS = 'events';
const SHEET_DAILY  = 'daily';

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- גיליון אירועים גולמי ---
    let sh = ss.getSheetByName(SHEET_EVENTS);
    if (!sh) {
      sh = ss.insertSheet(SHEET_EVENTS);
      sh.appendRow(['זמן', 'סוג', 'חידה', 'ניחושים', 'פוצח']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([
      new Date(),
      d.type || '',
      d.puzzle || '',
      d.guesses || '',
      d.won === true ? 'כן' : (d.won === false ? 'לא' : '')
    ]);

    rollup_(ss, d);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** סיכום יומי מתגלגל — כדי שלא תצטרך לחשב ידנית */
function rollup_(ss, d) {
  if (!d.puzzle) return;
  let sh = ss.getSheetByName(SHEET_DAILY);
  if (!sh) {
    sh = ss.insertSheet(SHEET_DAILY);
    sh.appendRow(['חידה', 'נכנסו', 'סיימו', 'פיצחו', 'אחוז הצלחה', 'ממוצע ניחושים']);
    sh.setFrozenRows(1);
  }

  const data = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < data.length; i++) if (data[i][0] == d.puzzle) { row = i + 1; break; }
  if (row === -1) {
    sh.appendRow([d.puzzle, 0, 0, 0, '', '']);
    row = sh.getLastRow();
  }

  const cur = sh.getRange(row, 1, 1, 6).getValues()[0];
  let [, views, done, wins] = cur;
  views = Number(views) || 0; done = Number(done) || 0; wins = Number(wins) || 0;

  // ממוצע הניחושים נשמר כסכום מצטבר בעמודה נסתרת
  const sumCell = sh.getRange(row, 7);
  let sum = Number(sumCell.getValue()) || 0;

  if (d.type === 'view') views++;
  if (d.type === 'done') {
    done++;
    if (d.won) { wins++; sum += Number(d.guesses) || 0; }
  }

  sumCell.setValue(sum);
  sh.getRange(row, 2, 1, 5).setValues([[
    views, done, wins,
    done ? Math.round(wins / done * 100) + '%' : '',
    wins ? (sum / wins).toFixed(1) : ''
  ]]);
}

/** בדיקה מהירה בדפדפן — פתח את ה-URL וראה שהוא עונה */
function doGet() {
  return json_({ ok: true, msg: 'beitardle analytics is running' });
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
