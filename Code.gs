/**
 * Deepthi & Ravi Wedding — Guest Book collector
 *
 * 1. Create a Google Sheet.
 * 2. Open Extensions → Apps Script.
 * 3. Paste this entire file into Code.gs.
 * 4. Deploy → New deployment → Web app.
 * 5. Execute as: Me
 * 6. Who has access: Anyone
 * 7. Copy the Web app URL into index.html as GUESTBOOK_SCRIPT_URL.
 */

const SHEET_NAME = 'Guest Book';

function doPost(e) {
  try {
    const sheet = getOrCreateSheet_();

    const raw = e && e.postData && e.postData.contents
      ? e.postData.contents
      : '{}';

    const data = JSON.parse(raw);

    if (data.type !== 'guestbook') {
      return json_({ ok: false, error: 'Invalid submission type' });
    }

    const name = String(data.name || '').trim();
    const message = String(data.message || '').trim();
    const page = String(data.page || '').trim();

    if (!name || !message) {
      return json_({ ok: false, error: 'Name and message are required' });
    }

    sheet.appendRow([
      new Date(),
      name,
      message,
      page
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Guest Name', 'Message', 'Website Page']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }

  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
