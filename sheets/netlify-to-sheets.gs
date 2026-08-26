/**
 * Netlify Forms → Google Sheets
 *
 * מקבל הרשמה מנטליפיי ומוסיף שורה בשתי לשוניות:
 *   1. "Form Integrations" — הקליטה הגולמית, כל ההרשמות
 *   2. "נרשמים ל<תאריך>"   — לשונית לכל מועד וובינר, נוצרת לבד אם חסרה
 *
 * התקנה:
 *   1. בגיליון: Extensions → Apps Script, להדביק את הקובץ הזה
 *   2. Deploy → New deployment → type: Web app
 *        Execute as:      Me
 *        Who has access:  Anyone
 *   3. להעתיק את ה-Web app URL
 *   4. בנטליפיי: Project configuration → Notifications →
 *      Form submission notifications → Add notification →
 *      HTTP POST request → להדביק את הכתובת
 */

var SHEET_RAW = 'Form Integrations';
var HEADERS = ['Submitted At', 'Name', 'Email', 'Replyto', 'טלפון', 'תאריך וובינר', 'Subject'];

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var p    = body.payload || body;
    var d    = p.data || {};

    var submittedAt = p.created_at || new Date().toISOString();
    var name    = d.name    || '';
    var email   = d.email   || '';
    var phone   = d.phone   || '';
    var webinar = d.webinar || '';           // '4.10' או 'יום ראשון, 4.10, 20:00'
    var subject = 'נר — ' + name;            // אותו דפוס שכבר בגיליון

    var row = [submittedAt, name, email, email, phone, webinar, subject];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    appendRow_(ss, SHEET_RAW, row);

    var dateTab = tabNameFor_(webinar);
    if (dateTab) appendRow_(ss, dateTab, row);

    return json_({ ok: true, sheet: dateTab || SHEET_RAW });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 'יום ראשון, 4.10, 20:00' → 'נרשמים ל4/10' */
function tabNameFor_(webinar) {
  var m = String(webinar).match(/(\d{1,2})[.\/](\d{1,2})/);
  if (!m) return null;
  return 'נרשמים ל' + Number(m[1]) + '/' + Number(m[2]);
}

function appendRow_(ss, tabName, row) {
  var sh = ss.getSheetByName(tabName);
  if (!sh) {
    sh = ss.insertSheet(tabName);
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  sh.appendRow(row);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}


/* ══════════════════════════════════════════════════════════════
   רשימת תפוצה למחזור חדש
   ──────────────────────────────────────────────────────────────
   מריצים ידנית מתוך העורך: בוחרים buildOutreachList ולוחצים Run.
   קורא את "Form Integrations", מנקה, ובונה לשונית "רשימת תפוצה".

   מה הוא עושה:
   · מסיר שורות בדיקה (טסט, בדיקה, וכתובות הצוות)
   · מאחד כפילויות לפי אימייל, ושומר את ההרשמה הראשונה
   · מנרמל טלפונים ל-+972... ובונה קישור וואטסאפ
   · מסמן טלפון שלא ניתן לפענח במקום לנחש אותו
   ══════════════════════════════════════════════════════════════ */

var TEST_PATTERNS = ['טסט', 'בדיקה', 'oren meyuhas', 'anna@meyuhas.tv',
                     'oren@meyuhas.tv', 'crip05@gmail.com', 'amir.meyuhas@gmail.com'];

function buildOutreachList() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName(SHEET_RAW);
  if (!src) throw new Error('לא נמצאה הלשונית ' + SHEET_RAW);

  var rows = src.getDataRange().getValues();
  rows.shift();                                   // כותרות

  var seen = {}, out = [];
  rows.forEach(function (r) {
    var name  = String(r[1] || '').trim();
    var email = String(r[2] || '').trim().toLowerCase();
    var phone = String(r[4] || '').trim();
    var when  = String(r[5] || '').trim();
    if (!email) return;

    var hay = (name + ' ' + email).toLowerCase();
    for (var i = 0; i < TEST_PATTERNS.length; i++) {
      if (hay.indexOf(TEST_PATTERNS[i].toLowerCase()) !== -1) return;
    }
    if (seen[email]) return;
    seen[email] = true;

    var p = normalizePhone_(phone);
    out.push([name, email, p.e164 || '', p.wa || '', p.note || '', when]);
  });

  var tab = ss.getSheetByName('רשימת תפוצה');
  if (tab) ss.deleteSheet(tab);
  tab = ss.insertSheet('רשימת תפוצה');
  tab.appendRow(['שם', 'אימייל', 'טלפון', 'וואטסאפ', 'הערה', 'וובינר קודם']);
  tab.getRange(1, 1, 1, 6).setFontWeight('bold');
  if (out.length) tab.getRange(2, 1, out.length, 6).setValues(out);
  tab.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    'נבנתה רשימת תפוצה: ' + out.length + ' אנשים ייחודיים.\n' +
    'שורות שסומנו בהערה דורשות בדיקה ידנית של הטלפון.');
}

/** מחזיר {e164, wa, note}. לא מנחש כשהמספר לא חד-משמעי. */
function normalizePhone_(raw) {
  var d = String(raw).replace(/\D/g, '');
  if (!d) return { note: 'אין טלפון' };

  if (d.indexOf('972') === 0) d = '0' + d.slice(3);      // 972541234567 → 0541234567
  if (d.length === 9 && d.charAt(0) !== '0') d = '0' + d; // 541234567    → 0541234567

  // נייד ישראלי: 05X ואחריו 7 ספרות
  if (/^05\d{8}$/.test(d)) {
    var e = '+972' + d.slice(1);
    return { e164: e, wa: 'https://wa.me/' + e.replace('+', '') };
  }
  // קווי ישראלי: 0X ואחריו 7 ספרות
  if (/^0[23489]\d{7}$/.test(d)) {
    var e2 = '+972' + d.slice(1);
    return { e164: e2, wa: 'https://wa.me/' + e2.replace('+', ''), note: 'קווי' };
  }
  // מספר זר שנראה שלם
  if (String(raw).charAt(0) === '+' && d.length >= 10) {
    return { e164: '+' + d, wa: 'https://wa.me/' + d, note: 'חו״ל' };
  }
  return { e164: String(raw), note: 'פורמט לא מזוהה — לבדוק ידנית' };
}
