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
