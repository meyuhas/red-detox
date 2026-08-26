var SHEET_RAW = 'Form Integrations';
var HEADERS = ['Submitted At', 'Name', 'Email', 'Replyto', 'טלפון', 'תאריך וובינר', 'Subject'];

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var d = extract_(e);

    var row = [
      d.created_at || new Date().toISOString(),
      d.name || '', d.email || '', d.email || '',
      d.phone || '', d.webinar || '',
      'נר — ' + (d.name || '')
    ];

    appendRow_(ss, SHEET_RAW, row);
    var tab = tabNameFor_(d.webinar || '');
    if (tab) appendRow_(ss, tab, row);

    return json_({ ok: true, sheet: tab || SHEET_RAW });
  } catch (err) {
    // כשל לא נעלם בשקט — נרשם ללשונית יומן עם מה שהתקבל בפועל
    try {
      var log = ss.getSheetByName('_log') || ss.insertSheet('_log');
      log.appendRow([new Date(), String(err),
                     e && e.postData ? e.postData.type : '',
                     e && e.postData ? String(e.postData.contents).slice(0, 400) : '',
                     JSON.stringify(e && e.parameter || {}).slice(0, 400)]);
    } catch (_) {}
    return json_({ ok: false, error: String(err) });
  }
}

/** נטליפיי עשויה לשלוח JSON או form-encoded. שניהם נתמכים. */
function extract_(e) {
  var raw = e && e.postData ? e.postData.contents : '';
  var body = null;
  try { body = JSON.parse(raw); } catch (_) { body = null; }

  if (!body && e && e.parameter && e.parameter.payload) {
    try { body = JSON.parse(e.parameter.payload); } catch (_) { body = null; }
  }
  if (body) {
    var p = body.payload || body;
    var d = p.data || p;
    return { created_at: p.created_at, name: d.name, email: d.email,
             phone: d.phone, webinar: d.webinar };
  }
  var q = (e && e.parameter) || {};
  if (q.name || q.email) {
    return { created_at: q.created_at, name: q.name, email: q.email,
             phone: q.phone, webinar: q.webinar };
  }
  throw new Error('גוף הבקשה לא נקרא. type=' + (e && e.postData ? e.postData.type : 'none'));
}

function tabNameFor_(webinar) {
  var m = String(webinar).match(/(\d{1,2})[.\/](\d{1,2})/);
  return m ? 'נרשמים ל' + Number(m[1]) + '/' + Number(m[2]) : null;
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
