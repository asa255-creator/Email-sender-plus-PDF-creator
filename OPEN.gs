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
    .addItem('Create Gmail Drafts', 'createDraftsFromList')
    .addItem('Send Emails with Attachment', 'sendEmailsFromListWithAttachment')
    .addSeparator()
    .addItem('Fill Emails from VCF File', 'fillEmailsFromVCF')
    .addItem('Fill Emails from Google Contacts', 'fillEmailsFromGoogleContacts')
    .addToUi();
}

/** ========================== SHEET VALIDATION ================ **/
function validateAndCreateSheets() {
  const ss = SpreadsheetApp.getActive();

  // Validate or create "People" sheet with headers
  validatePeopleSheet(ss);

  // Validate or create "email details" sheet with labels
  validateEmailDetailsSheet(ss);
}

function validatePeopleSheet(ss) {
  const SHEET_NAME = 'People';
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Check if headers exist in row 1
  const headers = sheet.getRange(1, 1, 1, 4).getValues()[0];

  // Set headers only if they're missing
  if (!headers[0]) sheet.getRange(1, 1).setValue('Name');
  if (!headers[2]) sheet.getRange(1, 3).setValue('Email');
  if (!headers[3]) sheet.getRange(1, 4).setValue('Phone');

  // Format header row
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f3f3f3');
}

function validateEmailDetailsSheet(ss) {
  const SHEET_NAME = 'email details';
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Check if labels exist in row 1
  const labels = sheet.getRange(1, 1, 1, 3).getValues()[0];

  // Set labels only if they're missing
  if (!labels[0]) sheet.getRange(1, 1).setValue('Body Template');
  if (!labels[1]) sheet.getRange(1, 2).setValue('Subject Template');
  if (!labels[2]) sheet.getRange(1, 3).setValue('Drive URL or File ID');

  // Format label row
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f3f3f3');

  // Add helpful notes in row 1 (below labels)
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

  // Resize columns for better readability
  sheet.setColumnWidth(1, 400);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 300);
}
