/**
 * OPEN.gs
 * Main entry point - creates menu and validates/creates required sheets
 */

/** ========================== MENU ============================ **/
function onOpen() {
  // First, ensure all required sheets and headers exist
  validateAndCreateSheets();

  // Then create the menu
  SpreadsheetApp.getUi()
    .createMenu('📧 Email Tools')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Create Drafts')
      .addItem('✉️ Create Drafts (no attachment)', 'createDraftsFromList')
      .addItem('📎 Create Drafts (with attachment)', 'createDraftsFromListWithAttachment'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Send Emails')
      .addItem('✉️ Send Emails (no attachment)', 'sendEmailsFromList')
      .addItem('📎 Send Emails (with attachment)', 'sendEmailsFromListWithAttachment'))
    .addSeparator()
    .addItem('📄 Import HTML from Google Doc', 'importHTMLFromGoogleDoc')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Find Contacts')
      .addItem('📇 Fill Emails from VCF File', 'fillEmailsFromVCF')
      .addItem('👤 Fill Emails from Google Contacts', 'fillEmailsFromGoogleContacts'))
    .addSeparator()
    .addItem('📑 Generate PDF Bundle & Labels', 'generatePDFBundleWithLabels')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Bounces')
      .addItem('🔎 Check for Bounces Now', 'checkForBounces')
      .addSeparator()
      .addItem('⏱️ Enable Auto-Check on Open', 'enableBounceAutoCheck')
      .addItem('🚫 Disable Auto-Check on Open', 'disableBounceAutoCheck'))
    .addToUi();
}

/** ========================== SHEET VALIDATION ================ **/
function validateAndCreateSheets() {
  const ss = SpreadsheetApp.getActive();

  // Validate or create "People" sheet with headers
  validatePeopleSheet(ss);

  // Validate or create "email details" sheet with labels
  validateEmailDetailsSheet(ss);

  // Validate or create "Bounced" sheet so the tab exists before the first scan
  getOrCreateBounceSheet(ss);
}

/**
 * Writes the expected header into a cell when it is blank OR holds something
 * else. The sheet is read positionally, so a renamed header is a silent
 * mismatch rather than a harmless label change.
 */
function enforceHeader(sheet, column, expected) {
  const cell = sheet.getRange(1, column);
  if (String(cell.getValue() || '').trim() !== expected) cell.setValue(expected);
}

function validatePeopleSheet(ss) {
  const SHEET_NAME = 'People';
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Enforce headers in row 1 (columns A-E)
  enforceHeader(sheet, 1, 'Name');
  enforceHeader(sheet, 2, 'PAC Names');
  enforceHeader(sheet, 3, 'Email');
  enforceHeader(sheet, 4, 'Phone');
  enforceHeader(sheet, 5, 'Address');

  // Format header row
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f3f3f3');

  // Resize columns for better readability
  sheet.setColumnWidth(1, 150); // Name
  sheet.setColumnWidth(2, 150); // PAC Names
  sheet.setColumnWidth(3, 200); // Email
  sheet.setColumnWidth(4, 120); // Phone
  sheet.setColumnWidth(5, 250); // Address
}

function validateEmailDetailsSheet(ss) {
  const SHEET_NAME = 'email details';
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Enforce labels in row 1
  enforceHeader(sheet, 1, 'Body Template');
  enforceHeader(sheet, 2, 'Subject Template');
  enforceHeader(sheet, 3, 'Drive URL or File ID');
  enforceHeader(sheet, 4, 'CC Emails');

  // Format label row
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f3f3f3');

  // Add helpful notes in row 2 (below labels)
  const noteStyle = SpreadsheetApp.newTextStyle().setFontSize(9).setForegroundColor('#666666').build();

  if (!sheet.getRange(2, 1).getValue()) {
    sheet.getRange(2, 1).setValue('Enter your email body here. Use [first name] or {{first name}} as placeholder.');
  }
  if (!sheet.getRange(2, 2).getValue()) {
    sheet.getRange(2, 2).setValue('Enter subject line here.');
  }
  if (!sheet.getRange(2, 3).getValue()) {
    sheet.getRange(2, 3).setValue('Optional: Google Drive file ID or URL for PDF attachment.');
  }
  if (!sheet.getRange(2, 4).getValue()) {
    sheet.getRange(2, 4).setValue('Optional: CC email addresses (one per row: D2, D3, D4, etc.)');
  }

  // Resize columns for better readability
  sheet.setColumnWidth(1, 400);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 250);
}
