const ADMIN_EMAIL = 'dominioncitychurchajah1@gmail.com';
const ADMIN_CC_EMAILS = 'olayemisegunsolomon@gmail.com,oluwakeminkechi@gmail.com';

const FIELD_ORDER = [
  'fullName',
  'email',
  'whatsappNumber',
  'occupationStatus',
  'trackSelection',
  'linkedin',
  'referralSource',
  'programGoals',
  'termsAccepted',
  'newsletterOptIn'
];

const FIELD_LABELS = {
  fullName: 'Full Name',
  email: 'Email',
  whatsappNumber: 'WhatsApp Number',
  occupationStatus: 'Occupation / Status',
  trackSelection: 'Track Selection',
  linkedin: 'LinkedIn',
  referralSource: 'Referral Source',
  programGoals: 'Goals from Program',
  termsAccepted: 'Terms Accepted',
  newsletterOptIn: 'Newsletter Opt-In'
};

function doPost(e) {
  try {
    const data = parsePayload_(e);
    const sheet = getOrCreateSheet_();
    ensureHeaders_(sheet);

    const timestamp = new Date();
    const row = [timestamp].concat(FIELD_ORDER.map(function(field) {
      return data[field] !== undefined ? data[field] : '';
    }));
    sheet.appendRow(row);

    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      cc: ADMIN_CC_EMAILS,
      subject: 'New #TheTwenty Application: ' + displayName_(data),
      htmlBody: buildEmailHtml_(data, timestamp),
      replyTo: data.email || ADMIN_EMAIL
    });

    if (data.email) {
      MailApp.sendEmail({
        to: data.email,
        subject: 'Thank you for applying to TheTwenty',
        htmlBody: buildApplicantConfirmationHtml_(data),
        replyTo: ADMIN_EMAIL
      });
    }

    return jsonResponse_({ success: true });
  } catch (error) {
    return jsonResponse_({ success: false, error: error.message });
  }
}

function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Missing POST body.');
  }

  const body = e.postData.contents;
  try {
    return JSON.parse(body);
  } catch (error) {
    const params = {};
    body.split('&').forEach(function(pair) {
      const parts = pair.split('=');
      if (parts[0]) {
        params[decodeURIComponent(parts[0])] = decodeURIComponent((parts[1] || '').replace(/\+/g, ' '));
      }
    });
    return params;
  }
}

function getOrCreateSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('This script must be linked to a Google Sheet.');
  }

  const sheetName = 'Applications';
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function ensureHeaders_(sheet) {
  const headers = ['Timestamp'].concat(FIELD_ORDER.map(function(field) {
    return FIELD_LABELS[field];
  }));
  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = currentHeaders.some(function(value) { return value !== ''; });
  const headersMatch = headers.every(function(header, index) {
    return currentHeaders[index] === header;
  });

  if (!hasHeaders || !headersMatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function buildEmailHtml_(data, timestamp) {
  const rows = FIELD_ORDER.map(function(field) {
    return '<tr><td style="padding:8px 12px;border:1px solid #ddd;font-weight:bold;">' +
      escapeHtml_(FIELD_LABELS[field]) +
      '</td><td style="padding:8px 12px;border:1px solid #ddd;">' +
      escapeHtml_(String(data[field] !== undefined ? data[field] : '')) +
      '</td></tr>';
  }).join('');

  return '<h2>New #TheTwenty Application</h2>' +
    '<p><strong>Submitted:</strong> ' + escapeHtml_(timestamp.toString()) + '</p>' +
    '<table style="border-collapse:collapse;width:100%;max-width:760px;">' +
    rows +
    '</table>';
}

function buildApplicantConfirmationHtml_(data) {
  const firstName = data.fullName ? ' ' + escapeHtml_(data.fullName.split(' ')[0]) : '';

  return '<div style="font-family:Arial,sans-serif;line-height:1.6;color:#111827;max-width:640px;">' +
    '<h2 style="color:#0B0F17;">Thank you' + firstName + ' - your TheTwenty application is in.</h2>' +
    '<p>We have received your application for <strong>TheTwenty</strong>.</p>' +
    '<p>Our team reviews applications on a rolling basis. If your application is shortlisted, we will contact you with the next steps using the email or WhatsApp number you provided.</p>' +
    '<p>In the meantime, please keep an eye on your inbox for updates from TheTwenty.</p>' +
    '<p style="margin-top:28px;">With purpose,<br><strong>TheTwenty Team</strong></p>' +
    '</div>';
}

function displayName_(data) {
  const name = (data.fullName || '').trim();
  return name || 'New Applicant';
}

function escapeHtml_(value) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
