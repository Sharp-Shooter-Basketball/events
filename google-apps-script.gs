// Paste this into Extensions → Apps Script in your Google Sheet.
// Set both email addresses below, then Deploy → New deployment → Web app.
// After any edit: Deploy → Manage deployments → Edit → New version
//
// Sheet layout: one row per child (parent info repeated, grouped by Registration ID)

const NOTIFY_EMAILS = [
  'your-email@gmail.com',       // ← your first email
  'your-other-email@gmail.com'  // ← your second email
];

const SHEET_NAME = 'Registrations'; // new layout tab (created automatically)

const HEADERS = [
  'Timestamp',
  'Registration ID',
  'Parent 1',
  'Parent 2',
  'Phone',
  'Email',
  'Emergency Contact',
  'Emergency Phone',
  'Medical Info',
  'Special Requests',
  'Photo Consent',
  'Children in Registration',
  'Child #',
  'Child Name',
  'Gender',
  'Date of Birth',
  'Age',
  'School',
  'Experience Level'
];

function parseRequestData(e) {
  if (e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents);
  }
  if (e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }
  throw new Error('No registration data received. Check form is hosted online and URL is correct.');
}

function getTargetSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function ensureHeaders(sheet) {
  const firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell === 'Timestamp') {
    return;
  }

  if (sheet.getLastRow() > 0) {
    sheet.insertRowBefore(1);
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#2C3E50')
    .setFontColor('#ffffff')
    .setWrap(true);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);
}

function escapeHtml(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatDisplayDate(isoString) {
  if (!isoString) return 'N/A';
  try {
    return Utilities.formatDate(new Date(isoString), Session.getScriptTimeZone(), 'dd MMM yyyy, HH:mm');
  } catch (err) {
    return isoString;
  }
}

function formatDob(dob) {
  if (!dob) return 'N/A';
  try {
    return Utilities.formatDate(new Date(dob + 'T00:00:00'), Session.getScriptTimeZone(), 'dd MMM yyyy');
  } catch (err) {
    return dob;
  }
}

function displayValue(value, fallback) {
  if (value === null || value === undefined || String(value).trim() === '') {
    return fallback || 'N/A';
  }
  return String(value);
}

function buildPlainTextEmail(data, regId, children) {
  const childBlocks = children.map(function(child, i) {
    return [
      'Child ' + (i + 1) + ':',
      '  Name: ' + displayValue(child.name),
      '  Gender: ' + displayValue(child.gender),
      '  DOB: ' + formatDob(child.dob),
      '  Age: ' + displayValue(child.age),
      '  School: ' + displayValue(child.school),
      '  Experience: ' + displayValue(child.experience, 'Not specified')
    ].join('\n');
  }).join('\n\n');

  return [
    'New Sharp Shooter Basketball Camp registration!',
    'Registration ID: ' + regId,
    '',
    '--- Parent ---',
    'Parent 1: ' + displayValue(data.parent1Name),
    'Parent 2: ' + displayValue(data.parent2Name),
    'Phone: ' + displayValue(data.phoneNumber),
    'Email: ' + displayValue(data.email),
    '',
    '--- Emergency ---',
    'Contact: ' + displayValue(data.emergencyContact),
    'Phone: ' + displayValue(data.emergencyPhone),
    '',
    '--- Medical / Notes ---',
    displayValue(data.medicalInfo, 'None'),
    '',
    '--- Special Requests ---',
    displayValue(data.specialRequests, 'None'),
    '',
    'Photo consent: ' + displayValue(data.photoConsent),
    '',
    '--- Children (' + children.length + ') ---',
    childBlocks || 'None',
    '',
    'Submitted: ' + formatDisplayDate(data.submittedAt)
  ].join('\n');
}

function buildHtmlEmail(data, regId, children) {
  const childRows = children.map(function(child, i) {
    return '<tr>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #eee;font-weight:600;">' + (i + 1) + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #eee;">' + escapeHtml(child.name) + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #eee;">' + escapeHtml(child.gender) + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #eee;">' + escapeHtml(formatDob(child.dob)) + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #eee;">' + escapeHtml(child.age) + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #eee;">' + escapeHtml(child.school) + '</td>' +
      '<td style="padding:10px 12px;border-bottom:1px solid #eee;">' + escapeHtml(displayValue(child.experience, 'Not specified')) + '</td>' +
      '</tr>';
  }).join('');

  const photoBadge = data.photoConsent === 'Yes'
    ? '<span style="background:#d4edda;color:#155724;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:600;">Yes</span>'
    : '<span style="background:#f8d7da;color:#721c24;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:600;">No</span>';

  return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;">' +
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f4f6f8;padding:24px 12px;">' +
      '<tr><td align="center">' +
        '<table role="presentation" width="600" cellspacing="0" cellpadding="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 8px 24px rgba(0,0,0,0.08);">' +
          '<tr><td style="background:linear-gradient(135deg,#2C3E50,#34495E);padding:28px 32px;text-align:center;">' +
            '<div style="font-size:28px;line-height:1;margin-bottom:10px;">🏀</div>' +
            '<h1 style="margin:0;color:#ffffff;font-size:24px;">New Camp Registration</h1>' +
            '<p style="margin:8px 0 0;color:rgba(255,255,255,0.85);font-size:14px;">Sharp Shooter Basketball · Summer Camp 2026</p>' +
          '</td></tr>' +
          '<tr><td style="padding:24px 32px 8px;">' +
            '<p style="margin:0 0 16px;color:#666;font-size:14px;">Registration ID: <strong style="color:#2C3E50;">' + escapeHtml(regId) + '</strong> · ' + escapeHtml(formatDisplayDate(data.submittedAt)) + '</p>' +
            '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:20px;">' +
              '<tr><td style="padding:16px;background:#f8f9fa;border-radius:12px;border-left:4px solid #FF6B35;">' +
                '<h2 style="margin:0 0 12px;font-size:16px;color:#2C3E50;">Parent / Guardian</h2>' +
                '<p style="margin:0 0 6px;color:#444;font-size:14px;"><strong>Parent 1:</strong> ' + escapeHtml(displayValue(data.parent1Name)) + '</p>' +
                '<p style="margin:0 0 6px;color:#444;font-size:14px;"><strong>Parent 2:</strong> ' + escapeHtml(displayValue(data.parent2Name)) + '</p>' +
                '<p style="margin:0 0 6px;color:#444;font-size:14px;"><strong>Phone:</strong> <a href="tel:' + escapeHtml(data.phoneNumber) + '" style="color:#FF6B35;text-decoration:none;">' + escapeHtml(displayValue(data.phoneNumber)) + '</a></p>' +
                '<p style="margin:0;color:#444;font-size:14px;"><strong>Email:</strong> <a href="mailto:' + escapeHtml(data.email) + '" style="color:#FF6B35;text-decoration:none;">' + escapeHtml(displayValue(data.email)) + '</a></p>' +
              '</td></tr>' +
            '</table>' +
            '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:20px;">' +
              '<tr><td width="50%" style="padding-right:8px;vertical-align:top;">' +
                '<div style="padding:16px;background:#f8f9fa;border-radius:12px;height:100%;">' +
                  '<h2 style="margin:0 0 10px;font-size:15px;color:#2C3E50;">Emergency</h2>' +
                  '<p style="margin:0 0 6px;color:#444;font-size:13px;"><strong>Contact:</strong> ' + escapeHtml(displayValue(data.emergencyContact)) + '</p>' +
                  '<p style="margin:0;color:#444;font-size:13px;"><strong>Phone:</strong> ' + escapeHtml(displayValue(data.emergencyPhone)) + '</p>' +
                '</div>' +
              '</td>' +
              '<td width="50%" style="padding-left:8px;vertical-align:top;">' +
                '<div style="padding:16px;background:#f8f9fa;border-radius:12px;height:100%;">' +
                  '<h2 style="margin:0 0 10px;font-size:15px;color:#2C3E50;">Consent</h2>' +
                  '<p style="margin:0;color:#444;font-size:13px;"><strong>Photo permission:</strong> ' + photoBadge + '</p>' +
                '</div>' +
              '</td></tr>' +
            '</table>' +
            '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:20px;">' +
              '<tr><td style="padding:16px;background:#fff8f5;border-radius:12px;border:1px solid #ffd8c7;">' +
                '<h2 style="margin:0 0 8px;font-size:15px;color:#2C3E50;">Medical / Allergies</h2>' +
                '<p style="margin:0;color:#444;font-size:13px;line-height:1.5;">' + escapeHtml(displayValue(data.medicalInfo, 'None')) + '</p>' +
              '</td></tr>' +
            '</table>' +
            '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:24px;">' +
              '<tr><td style="padding:16px;background:#f8f9fa;border-radius:12px;">' +
                '<h2 style="margin:0 0 8px;font-size:15px;color:#2C3E50;">Special Requests</h2>' +
                '<p style="margin:0;color:#444;font-size:13px;line-height:1.5;">' + escapeHtml(displayValue(data.specialRequests, 'None')) + '</p>' +
              '</td></tr>' +
            '</table>' +
            '<h2 style="margin:0 0 12px;font-size:16px;color:#2C3E50;">Children (' + children.length + ')</h2>' +
            '<div style="overflow-x:auto;border:1px solid #eee;border-radius:12px;">' +
              '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;min-width:520px;">' +
                '<thead><tr style="background:#2C3E50;color:#ffffff;">' +
                  '<th style="padding:10px 12px;text-align:left;font-size:12px;">#</th>' +
                  '<th style="padding:10px 12px;text-align:left;font-size:12px;">Name</th>' +
                  '<th style="padding:10px 12px;text-align:left;font-size:12px;">Gender</th>' +
                  '<th style="padding:10px 12px;text-align:left;font-size:12px;">DOB</th>' +
                  '<th style="padding:10px 12px;text-align:left;font-size:12px;">Age</th>' +
                  '<th style="padding:10px 12px;text-align:left;font-size:12px;">School</th>' +
                  '<th style="padding:10px 12px;text-align:left;font-size:12px;">Experience</th>' +
                '</tr></thead>' +
                '<tbody>' + childRows + '</tbody>' +
              '</table>' +
            '</div>' +
          '</td></tr>' +
          '<tr><td style="background:#f8f9fa;padding:18px 32px;text-align:center;border-top:1px solid #eee;">' +
            '<p style="margin:0;color:#888;font-size:12px;">Sharp Shooter Basketball Camp · Nyamata Youth Center</p>' +
          '</td></tr>' +
        '</table>' +
      '</td></tr>' +
    '</table>' +
  '</body></html>';
}

function saveRegistrationRows(sheet, data, regId) {
  ensureHeaders(sheet);

  const timestamp = data.submittedAt || new Date().toISOString();
  const children = data.children && data.children.length ? data.children : [{}];
  const childCount = children.length;

  const rows = children.map(function(child, index) {
    return [
      timestamp,
      regId,
      data.parent1Name || '',
      data.parent2Name || '',
      data.phoneNumber || '',
      data.email || '',
      data.emergencyContact || '',
      data.emergencyPhone || '',
      data.medicalInfo || '',
      data.specialRequests || '',
      data.photoConsent || '',
      childCount,
      index + 1,
      child.name || '',
      child.gender || '',
      child.dob || '',
      child.age || '',
      child.school || '',
      child.experience || ''
    ];
  });

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
}

function doGet() {
  return ContentService
    .createTextOutput('Sharp Shooter camp registration endpoint is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const data = parseRequestData(e);
    const sheet = getTargetSheet();
    const regId = Utilities.getUuid().slice(0, 8).toUpperCase();
    const children = data.children || [];

    saveRegistrationRows(sheet, data, regId);

    const subject = '🏀 New Camp Registration: ' + (data.parent1Name || 'Unknown') +
      (children.length > 1 ? ' (' + children.length + ' children)' : '');
    const plainBody = buildPlainTextEmail(data, regId, children);
    const htmlBody = buildHtmlEmail(data, regId, children);

    MailApp.sendEmail({
      to: NOTIFY_EMAILS.join(','),
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, registrationId: regId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Run once to set up column headers (safe to run again)
function setupSheetHeaders() {
  const sheet = getTargetSheet();
  ensureHeaders(sheet);
  Logger.log('Headers set on sheet: ' + sheet.getName());
}

// Run once to confirm the script can write to your sheet
function testAppendRow() {
  const sheet = getTargetSheet();
  const testData = {
    submittedAt: new Date().toISOString(),
    parent1Name: 'TEST Parent',
    parent2Name: '',
    phoneNumber: '+250 781 686 414',
    email: 'test@example.com',
    emergencyContact: '',
    emergencyPhone: '',
    medicalInfo: '',
    specialRequests: '',
    photoConsent: 'No',
    children: [
      { name: 'Test Child One', gender: 'female', dob: '2012-05-10', age: 14, school: 'Test School', experience: 'beginner' },
      { name: 'Test Child Two', gender: 'male', dob: '2010-08-20', age: 16, school: 'Test School', experience: 'intermediate' }
    ]
  };
  saveRegistrationRows(sheet, testData, 'TEST1234');
  Logger.log('Test rows added to: ' + sheet.getName());
}
