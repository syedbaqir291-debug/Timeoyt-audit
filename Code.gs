/**
 * ============================================================
 * Time-Out Audit — Apps Script backend
 * ============================================================
 * SETUP:
 * 1. In your Google Sheet: Extensions > Apps Script
 * 2. Delete any placeholder code, paste this whole file in
 * 3. Click Deploy > New deployment
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Click Deploy, copy the Web App URL it gives you
 * 5. Paste that URL into APP_URL at the top of github.html
 *
 * Every submission from the HTML form lands as one new row in a
 * sheet called "Audit Data" (created automatically on first use).
 * ============================================================
 */

const SHEET_NAME = 'Audit Data';

const QUESTIONS = [
  '"TIME OUT" is conducted before start of the procedure in the location the procedure is be done',
  'All team members (physicians, assisting nurse) involved are present during the Time out',
  'The Team identifies the patient using TWO unique identifiers from patient and ID band.',
  'The team checks the consent for the correct procedure',
  'The correct side (if applicable) — Site marking done',
  'All relevant documents/images/studies are available, properly labelled and displayed',
  'All requested implants/special equipment are available and have been verified',
  'Are prophylactic antibiotics required for this procedure?',
  'PRE-SEDATION ASSESSMENT (if applicable)',
  'Name and Employee Code of the Physician and nurse documented.',
  'During timeout, there was no Interruptions/Disruptions observed'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet_();

    const answers = data.answers || [];
    const status = answers.some(a => a === 'Not Met') ? 'Not Met' : 'Met';

    const row = [
      new Date(),
      data.department || '',
      data.patientName || '',
      data.mrno || '',
      data.date || '',
      data.auditorName || '',
      data.speciality || '',
      ...QUESTIONS.map((_, i) => answers[i] || ''),
      status
    ];
    sheet.appendRow(row);

    return jsonOut_({ result: 'success' });
  } catch (err) {
    return jsonOut_({ result: 'error', message: err.message });
  }
}

function doGet(e) {
  return jsonOut_({ status: 'Time-Out Audit API is running' });
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      'Timestamp', 'Department', 'Name of Patient', 'MRNO', 'Date',
      'Auditor Name', 'Speciality'
    ];
    QUESTIONS.forEach((_, i) => headers.push('Q' + (i + 1)));
    headers.push('Patient Status');
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
