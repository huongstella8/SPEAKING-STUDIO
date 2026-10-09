const SHEET_NAME = 'results';

const COLUMNS = [
  'received_at','session_id','student_name','student_email','part','module_set',
  'date_local','take_number','take_seconds','take_time','take_after_models',
  'recorded_before_models','models_opened','total_takes','total_practice_seconds',
  'audio_file','notes'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const payload = JSON.parse(e.postData.contents);
    const rows = payload.rows || [];
    const sheet = getSheet_();
    const now = new Date();

    const values = rows.map(function (r) {
      return COLUMNS.map(function (c) {
        return c === 'received_at' ? now : (r[c] !== undefined ? r[c] : '');
      });
    });

    if (values.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, values.length, COLUMNS.length)
           .setValues(values);
    }
    return json_({ ok: true, added: values.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, message: 'Endpoint is live. Post practice results here.' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
