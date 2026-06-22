// Paste this into Extensions → Apps Script in your Google Sheet.
// Set both email addresses below, then Deploy → New deployment → Web app.
// After any edit: Deploy → Manage deployments → Edit → New version

const NOTIFY_EMAILS = [
  'your-email@gmail.com',       // ← your first email
  'your-other-email@gmail.com'  // ← your second email
];

const SHEET_NAME = 'Sheet1'; // rename if your tab isn't "Sheet1"

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
    sheet = spreadsheet.getSheets()[0];
  }
  if (!sheet) {
    throw new Error('Could not find a sheet to write to.');
  }
  return sheet;
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

    const childrenText = (data.children || []).map(function(child, i) {
      return [
        'Child ' + (i + 1) + ':',
        '  Name: ' + child.name,
        '  Gender: ' + child.gender,
        '  DOB: ' + child.dob,
        '  Age: ' + child.age,
        '  School: ' + child.school,
        '  Experience: ' + (child.experience || 'Not specified')
      ].join('\n');
    }).join('\n\n');

    sheet.appendRow([
      data.submittedAt || new Date().toISOString(),
      data.parent1Name || '',
      data.parent2Name || '',
      data.phoneNumber || '',
      data.email || '',
      data.emergencyContact || '',
      data.emergencyPhone || '',
      data.medicalInfo || '',
      data.specialRequests || '',
      data.photoConsent || '',
      childrenText
    ]);

    const subject = '🏀 New Camp Registration: ' + (data.parent1Name || 'Unknown');
    const body = [
      'New Sharp Shooter Basketball Camp registration!',
      '',
      '--- Parent ---',
      'Parent 1: ' + (data.parent1Name || 'N/A'),
      'Parent 2: ' + (data.parent2Name || 'N/A'),
      'Phone: ' + (data.phoneNumber || 'N/A'),
      'Email: ' + (data.email || 'N/A'),
      '',
      '--- Emergency ---',
      'Contact: ' + (data.emergencyContact || 'N/A'),
      'Phone: ' + (data.emergencyPhone || 'N/A'),
      '',
      '--- Medical / Notes ---',
      data.medicalInfo || 'None',
      '',
      '--- Special Requests ---',
      data.specialRequests || 'None',
      '',
      'Photo consent: ' + (data.photoConsent || 'N/A'),
      '',
      '--- Children ---',
      childrenText || 'None',
      '',
      'Submitted: ' + (data.submittedAt || new Date().toISOString())
    ].join('\n');

    MailApp.sendEmail({
      to: NOTIFY_EMAILS.join(','),
      subject: subject,
      body: body
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Run once from the editor to confirm the script can write to your sheet
function testAppendRow() {
  const sheet = getTargetSheet();
  sheet.appendRow([new Date().toISOString(), 'TEST', '', '000', 'test@example.com', '', '', '', '', 'No', 'Test child']);
  Logger.log('Test row added to: ' + sheet.getName());
}
