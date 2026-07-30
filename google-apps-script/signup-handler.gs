const SPREADSHEET_ID = 'PASTE_SPREADSHEET_ID_HERE';
const SHEET_NAME = 'Signups';
const DEFAULT_NOTIFY_EMAIL = 'nnsh774@gmail.com';

function doPost(e) {
  const data = parseRequest_(e);
  const email = String(data.email || '').trim();
  const notifyEmail = String(data.notifyEmail || DEFAULT_NOTIFY_EMAIL).trim();

  if (!email || !email.includes('@')) {
    return json_({ ok: false, error: 'Invalid email' });
  }

  const submittedAt = data.submittedAt || new Date().toISOString();
  const source = data.source || 'life-alignment';
  const pageUrl = data.pageUrl || '';

  const sheet = getSheet_();
  sheet.appendRow([new Date(), email, source, pageUrl, submittedAt]);

  MailApp.sendEmail({
    to: notifyEmail,
    subject: 'Life Alignment: new email signup',
    body: [
      'Life Alignment に新しい登録がありました。',
      '',
      `Email: ${email}`,
      `Source: ${source}`,
      `Page: ${pageUrl}`,
      `Submitted at: ${submittedAt}`
    ].join('\n')
  });

  return json_({ ok: true });
}

function parseRequest_(e) {
  if (!e) return {};

  if (e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (error) {
      return e.parameter || {};
    }
  }

  return e.parameter || {};
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Received At', 'Email', 'Source', 'Page URL', 'Submitted At']);
  }

  return sheet;
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
