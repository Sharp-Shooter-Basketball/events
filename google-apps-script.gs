// Paste this into Extensions → Apps Script in your Google Sheet.
// Set both email addresses below, then Deploy → New deployment → Web app.

const NOTIFY_EMAILS = [
  'your-email@gmail.com',       // ← your first email
  'your-other-email@gmail.com'    // ← your second email
];

const SHEET_NAME = 'Sheet1'; // rename if your tab isn't "Sheet1"

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

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
      data.parent1Name,
      data.parent2Name || '',
      data.phoneNumber,
      data.email,
      data.emergencyContact || '',
      data.emergencyPhone || '',
      data.medicalInfo || '',
      data.specialRequests || '',
      data.photoConsent,
      childrenText
    ]);

    const subject = '🏀 New Camp Registration: ' + data.parent1Name;
    const body = [
      'New Sharp Shooter Basketball Camp registration!',
      '',
      '--- Parent ---',
      'Parent 1: ' + data.parent1Name,
      'Parent 2: ' + (data.parent2Name || 'N/A'),
      'Phone: ' + data.phoneNumber,
      'Email: ' + data.email,
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
      'Photo consent: ' + data.photoConsent,
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
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
