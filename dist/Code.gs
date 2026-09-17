////////////////////////////////////////////////////////////////////////////
// GENERATED FILE - DO NOT EDIT
////////////////////////////////////////////////////////////////////////////
//
// This is every file in src/ concatenated into one, so it can be pasted
// into the Apps Script editor in a single step.
//
// Edit the modules in src/ and run `node build.mjs` instead. CI regenerates
// this file on every push, so hand edits here will be overwritten.
//
// Modules: OPEN.gs, EmailSender.gs, BounceChecker.gs, EmailFinderGoogle.gs, EmailFinderVCF.gs, PDFBundler.gs, Utilities.gs
////////////////////////////////////////////////////////////////////////////
// END HEADER
////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////
// src/OPEN.gs
////////////////////////////////////////////////////////////////////////////

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


////////////////////////////////////////////////////////////////////////////
// src/EmailSender.gs
////////////////////////////////////////////////////////////////////////////

/**
 * EmailSender.gs
 * Functions for creating Gmail drafts and sending emails with attachments
 */

/** ========================== CONFIG ========================== **/
const LIST_SHEET = 'People';           // names in A, emails in C
const DETAILS_SHEET = 'email details'; // A2 = body template, B2 = subject, C2 = Drive URL or ID for PDF
const NAME_COL = 1;                    // A: Name
const PAC_COL = 2;                     // B: PAC Names
const EMAIL_COL = 3;                   // C: Email
const PHONE_COL = 4;                   // D: Phone
const ADDRESS_COL = 5;                 // E: Address
const USE_HTML = true;                 // create HTML drafts or HTML emails

/** ========================== CREATE DRAFTS =================== **/
function createDraftsFromList() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(LIST_SHEET) || ss.getActiveSheet();
  const detailsSh = ss.getSheetByName(DETAILS_SHEET);
  if (!detailsSh) throw new Error('Sheet "email details" not found.');

  // Use getValue so we keep raw HTML if present
  const bodyTemplate = String(detailsSh.getRange('A2').getValue() || '');
  const subjectTemplate = String(detailsSh.getRange('B2').getValue() || '');
  if (!bodyTemplate) throw new Error('Body template missing in email details A2.');
  if (!subjectTemplate) throw new Error('Subject missing in email details B2.');

  // Get CC addresses from column D (D2, D3, D4, etc.)
  const ccAddresses = getCCAddresses(detailsSh);

  const lastRow = listSh.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data rows found.');
    return;
  }

  const width = Math.max(NAME_COL, PAC_COL, EMAIL_COL, PHONE_COL, ADDRESS_COL);
  const values = listSh.getRange(2, 1, lastRow - 1, width).getDisplayValues();

  const signatureHtml = getDefaultSignatureHtml(); // may be ''

  let created = 0;
  values.forEach(row => {
    const fullName = String(row[NAME_COL - 1] || '').trim() || 'To Whom It May Concern';
    const pacName = String(row[PAC_COL - 1] || '').trim();
    const email = String(row[EMAIL_COL - 1] || '').trim();
    const phone = String(row[PHONE_COL - 1] || '').trim();
    const address = String(row[ADDRESS_COL - 1] || '').trim();

    if (!email) return; // Only skip if email is missing

    const firstName = fullName === 'To Whom It May Concern' ? fullName : extractFirstName(fullName);

    // Build person data object for placeholder replacement
    let personData = {
      fullName: fullName,
      firstName: firstName,
      pacName: pacName,
      email: email,
      phone: phone,
      address: address
    };

    // Normalize capitalization (ALL CAPS → Title Case)
    personData = normalizePersonData(personData);

    // Replace placeholders in subject and body
    const subject = replaceAllPlaceholders(subjectTemplate, personData);
    const bodyWithPlaceholders = replaceAllPlaceholders(bodyTemplate, personData);

    if (USE_HTML) {
      const bodyHtml = buildHtmlBodyFromTemplate(bodyWithPlaceholders, signatureHtml);
      const options = { htmlBody: bodyHtml };
      if (ccAddresses) options.cc = ccAddresses;
      GmailApp.createDraft(email, subject, '', options);
    } else {
      const bodyText = asPlainText(bodyWithPlaceholders);
      const bodyWithSig = bodyText + (signatureHtml ? '\n\n' + stripHtml(signatureHtml) : '');
      const options = {};
      if (ccAddresses) options.cc = ccAddresses;
      GmailApp.createDraft(email, subject, bodyWithSig, options);
    }

    created++;
  });

  SpreadsheetApp.getUi().alert('Drafts created: ' + created + (ccAddresses ? '\nCC: ' + ccAddresses : ''));
}

/** ============== CREATE DRAFTS WITH ATTACHMENT =============== **/
function createDraftsFromListWithAttachment() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(LIST_SHEET) || ss.getActiveSheet();
  const detailsSh = ss.getSheetByName(DETAILS_SHEET);
  if (!detailsSh) throw new Error('Sheet "email details" not found.');

  const bodyTemplate = String(detailsSh.getRange('A2').getValue() || '');
  const subjectTemplate = String(detailsSh.getRange('B2').getValue() || '');
  const attachmentRef = String(detailsSh.getRange('C2').getValue() || '').trim();
  if (!bodyTemplate) throw new Error('Body template missing in email details A2.');
  if (!subjectTemplate) throw new Error('Subject missing in email details B2.');

  // Get CC addresses from column D (D2, D3, D4, etc.)
  const ccAddresses = getCCAddresses(detailsSh);

  // Validate and confirm attachment file (same flow as PDF Bundle)
  const attachmentInfo = validateAndConfirmAttachment(ui, attachmentRef);
  if (attachmentInfo === null) return; // User cancelled or error

  const lastRow = listSh.getLastRow();
  if (lastRow < 2) {
    ui.alert('No data rows found.');
    return;
  }

  const width = Math.max(NAME_COL, PAC_COL, EMAIL_COL, PHONE_COL, ADDRESS_COL);
  const values = listSh.getRange(2, 1, lastRow - 1, width).getDisplayValues();

  const signatureHtml = getDefaultSignatureHtml(); // may be ''

  let created = 0;
  values.forEach(row => {
    const fullName = String(row[NAME_COL - 1] || '').trim() || 'To Whom It May Concern';
    const pacName = String(row[PAC_COL - 1] || '').trim();
    const email = String(row[EMAIL_COL - 1] || '').trim();
    const phone = String(row[PHONE_COL - 1] || '').trim();
    const address = String(row[ADDRESS_COL - 1] || '').trim();

    if (!email) return; // Only skip if email is missing

    const firstName = fullName === 'To Whom It May Concern' ? fullName : extractFirstName(fullName);

    // Build person data object for placeholder replacement
    let personData = {
      fullName: fullName,
      firstName: firstName,
      pacName: pacName,
      email: email,
      phone: phone,
      address: address
    };

    // Normalize capitalization (ALL CAPS → Title Case)
    personData = normalizePersonData(personData);

    // Replace placeholders in subject and body
    const subject = replaceAllPlaceholders(subjectTemplate, personData);
    const bodyWithPlaceholders = replaceAllPlaceholders(bodyTemplate, personData);

    // Build attachment blob: personalized PDF if Google Doc, otherwise static file
    const attachmentBlob = attachmentInfo ? buildAttachmentBlob(attachmentInfo, personData) : null;

    if (USE_HTML) {
      const bodyHtml = buildHtmlBodyFromTemplate(bodyWithPlaceholders, signatureHtml);
      const options = attachmentBlob
        ? { htmlBody: bodyHtml, attachments: [attachmentBlob] }
        : { htmlBody: bodyHtml };
      if (ccAddresses) options.cc = ccAddresses;
      GmailApp.createDraft(email, subject, '', options);
    } else {
      const bodyText = asPlainText(bodyWithPlaceholders);
      const bodyWithSig = bodyText + (signatureHtml ? '\n\n' + stripHtml(signatureHtml) : '');
      const options = attachmentBlob ? { attachments: [attachmentBlob] } : {};
      if (ccAddresses) options.cc = ccAddresses;
      GmailApp.createDraft(email, subject, bodyWithSig, options);
    }

    created++;
  });

  const attachLabel = attachmentInfo
    ? (attachmentInfo.isGoogleDoc ? ' (with personalized PDF attachment)' : ' (with attachment)')
    : ' (no attachment)';
  ui.alert('Drafts created: ' + created + attachLabel + (ccAddresses ? '\nCC: ' + ccAddresses : ''));
}

/** ==================== SEND WITHOUT ATTACHMENT =============== **/
function sendEmailsFromList() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(LIST_SHEET) || ss.getActiveSheet();
  const detailsSh = ss.getSheetByName(DETAILS_SHEET);
  if (!detailsSh) throw new Error('Sheet "email details" not found.');

  const bodyTemplate = String(detailsSh.getRange('A2').getValue() || '');
  const subjectTemplate = String(detailsSh.getRange('B2').getValue() || '');
  if (!bodyTemplate) throw new Error('Body template missing in email details A2.');
  if (!subjectTemplate) throw new Error('Subject missing in email details B2.');

  // Get CC addresses from column D (D2, D3, D4, etc.)
  const ccAddresses = getCCAddresses(detailsSh);

  const lastRow = listSh.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data rows found.');
    return;
  }

  const width = Math.max(NAME_COL, PAC_COL, EMAIL_COL, PHONE_COL, ADDRESS_COL);
  const values = listSh.getRange(2, 1, lastRow - 1, width).getDisplayValues();

  const signatureHtml = getDefaultSignatureHtml(); // may be ''

  let sent = 0;
  values.forEach(row => {
    const fullName = String(row[NAME_COL - 1] || '').trim() || 'To Whom It May Concern';
    const pacName = String(row[PAC_COL - 1] || '').trim();
    const email = String(row[EMAIL_COL - 1] || '').trim();
    const phone = String(row[PHONE_COL - 1] || '').trim();
    const address = String(row[ADDRESS_COL - 1] || '').trim();

    if (!email) return; // Only skip if email is missing

    const firstName = fullName === 'To Whom It May Concern' ? fullName : extractFirstName(fullName);

    // Build person data object for placeholder replacement
    let personData = {
      fullName: fullName,
      firstName: firstName,
      pacName: pacName,
      email: email,
      phone: phone,
      address: address
    };

    // Normalize capitalization (ALL CAPS → Title Case)
    personData = normalizePersonData(personData);

    // Replace placeholders in subject and body
    const subject = replaceAllPlaceholders(subjectTemplate, personData);
    const bodyWithPlaceholders = replaceAllPlaceholders(bodyTemplate, personData);

    if (USE_HTML) {
      const bodyHtml = buildHtmlBodyFromTemplate(bodyWithPlaceholders, signatureHtml);
      const options = { htmlBody: bodyHtml };
      if (ccAddresses) options.cc = ccAddresses;
      GmailApp.sendEmail(email, subject, stripHtml(bodyHtml) || ' ', options);
    } else {
      const bodyText = asPlainText(bodyWithPlaceholders);
      const bodyWithSig = bodyText + (signatureHtml ? '\n\n' + stripHtml(signatureHtml) : '');
      const options = {};
      if (ccAddresses) options.cc = ccAddresses;
      GmailApp.sendEmail(email, subject, bodyWithSig, options);
    }

    sent++;
  });

  SpreadsheetApp.getUi().alert('Emails sent: ' + sent + (ccAddresses ? '\nCC: ' + ccAddresses : ''));
}

/** ===================== SEND WITH ATTACHMENT ================= **/
function sendEmailsFromListWithAttachment() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(LIST_SHEET) || ss.getActiveSheet();
  const detailsSh = ss.getSheetByName(DETAILS_SHEET);
  if (!detailsSh) throw new Error('Sheet "email details" not found.');

  const bodyTemplate = String(detailsSh.getRange('A2').getValue() || '');
  const subjectTemplate = String(detailsSh.getRange('B2').getValue() || '');
  const attachmentRef = String(detailsSh.getRange('C2').getValue() || '').trim();
  if (!bodyTemplate) throw new Error('Body template missing in email details A2.');
  if (!subjectTemplate) throw new Error('Subject missing in email details B2.');

  // Get CC addresses from column D (D2, D3, D4, etc.)
  const ccAddresses = getCCAddresses(detailsSh);

  // Validate and confirm attachment file (same flow as PDF Bundle)
  const attachmentInfo = validateAndConfirmAttachment(ui, attachmentRef);
  if (attachmentInfo === null) return; // User cancelled or error

  const lastRow = listSh.getLastRow();
  if (lastRow < 2) {
    ui.alert('No data rows found.');
    return;
  }

  const width = Math.max(NAME_COL, PAC_COL, EMAIL_COL, PHONE_COL, ADDRESS_COL);
  const values = listSh.getRange(2, 1, lastRow - 1, width).getDisplayValues();

  const signatureHtml = getDefaultSignatureHtml(); // may be ''

  let sent = 0;
  values.forEach(row => {
    const fullName = String(row[NAME_COL - 1] || '').trim() || 'To Whom It May Concern';
    const pacName = String(row[PAC_COL - 1] || '').trim();
    const email = String(row[EMAIL_COL - 1] || '').trim();
    const phone = String(row[PHONE_COL - 1] || '').trim();
    const address = String(row[ADDRESS_COL - 1] || '').trim();

    if (!email) return; // Only skip if email is missing

    const firstName = fullName === 'To Whom It May Concern' ? fullName : extractFirstName(fullName);

    // Build person data object for placeholder replacement
    let personData = {
      fullName: fullName,
      firstName: firstName,
      pacName: pacName,
      email: email,
      phone: phone,
      address: address
    };

    // Normalize capitalization (ALL CAPS → Title Case)
    personData = normalizePersonData(personData);

    // Replace placeholders in subject and body
    const subject = replaceAllPlaceholders(subjectTemplate, personData);
    const bodyWithPlaceholders = replaceAllPlaceholders(bodyTemplate, personData);

    // Build attachment blob: personalized PDF if Google Doc, otherwise static file
    const attachmentBlob = attachmentInfo ? buildAttachmentBlob(attachmentInfo, personData) : null;

    if (USE_HTML) {
      const bodyHtml = buildHtmlBodyFromTemplate(bodyWithPlaceholders, signatureHtml);
      const options = attachmentBlob
        ? { htmlBody: bodyHtml, attachments: [attachmentBlob] }
        : { htmlBody: bodyHtml };
      if (ccAddresses) options.cc = ccAddresses;
      GmailApp.sendEmail(email, subject, stripHtml(bodyHtml) || ' ', options);
    } else {
      const bodyText = asPlainText(bodyWithPlaceholders);
      const bodyWithSig = bodyText + (signatureHtml ? '\n\n' + stripHtml(signatureHtml) : '');
      const options = attachmentBlob ? { attachments: [attachmentBlob] } : {};
      if (ccAddresses) options.cc = ccAddresses;
      GmailApp.sendEmail(email, subject, bodyWithSig, options);
    }

    sent++;
  });

  const attachLabel = attachmentInfo
    ? (attachmentInfo.isGoogleDoc ? ' (with personalized PDF attachment)' : ' (with attachment)')
    : ' (no attachment)';
  ui.alert('Emails sent: ' + sent + attachLabel + (ccAddresses ? '\nCC: ' + ccAddresses : ''));
}

/** ================== IMPORT HTML FROM GOOGLE DOC ============= **/
/**
 * Imports HTML content from a Google Doc URL and places it in cell A2 of "email details" sheet
 */
function importHTMLFromGoogleDoc() {
  const ui = SpreadsheetApp.getUi();

  // Prompt user for Google Doc URL
  const response = ui.prompt(
    'Import HTML from Google Doc',
    'Paste the Google Doc URL or file ID:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return; // User cancelled
  }

  const input = response.getResponseText().trim();
  if (!input) {
    ui.alert('No URL provided.');
    return;
  }

  // Extract file ID from URL or use input as ID
  const fileId = extractDriveId(input);
  if (!fileId) {
    ui.alert('Error: Could not extract file ID from input.\n\n' +
             'Please provide either:\n' +
             '• Full Google Doc URL\n' +
             '• Just the file ID');
    return;
  }

  try {
    // Open the Google Doc
    const doc = DocumentApp.openById(fileId);
    const body = doc.getBody();

    // Get the HTML content
    // Note: Apps Script doesn't have a direct "export as HTML" API
    // So we'll build HTML from the document structure
    const htmlContent = convertDocBodyToHTML(body);

    // Place in cell A2 of "email details" sheet
    const ss = SpreadsheetApp.getActive();
    const detailsSh = ss.getSheetByName(DETAILS_SHEET);
    if (!detailsSh) {
      ui.alert('Error: Sheet "email details" not found.');
      return;
    }

    detailsSh.getRange('A2').setValue(htmlContent);

    ui.alert('Success!\n\n' +
             'HTML content imported to cell A2 of "email details" sheet.\n\n' +
             'Document: ' + doc.getName());

  } catch (e) {
    ui.alert('Error importing HTML:\n\n' + e.message + '\n\n' +
             'Make sure you have access to the document and the URL/ID is correct.');
  }
}

/**
 * Converts Google Doc body to HTML
 */
function convertDocBodyToHTML(body) {
  let html = '<div>';

  const numChildren = body.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const element = body.getChild(i);
    const elementType = element.getType();

    if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      const para = element.asParagraph();
      const text = para.getText();

      if (text.trim() !== '') {
        // Get text attributes for basic formatting
        const textElement = para.editAsText();
        let paraHtml = '<p>';

        // For simplicity, we'll just add the text
        // More sophisticated version would handle bold, italic, etc.
        paraHtml += escapeHtml(text);
        paraHtml += '</p>';

        html += paraHtml;
      }
    } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
      const listItem = element.asListItem();
      const text = listItem.getText();
      html += '<li>' + escapeHtml(text) + '</li>';
    } else if (elementType === DocumentApp.ElementType.TABLE) {
      // Basic table support
      html += '<table border="1">';
      const table = element.asTable();
      const numRows = table.getNumRows();

      for (let r = 0; r < numRows; r++) {
        html += '<tr>';
        const row = table.getRow(r);
        const numCells = row.getNumCells();

        for (let c = 0; c < numCells; c++) {
          const cell = row.getCell(c);
          html += '<td>' + escapeHtml(cell.getText()) + '</td>';
        }
        html += '</tr>';
      }
      html += '</table>';
    }
  }

  html += '</div>';
  return html;
}

/** ========= ATTACHMENT VALIDATION & CONFIRMATION ============= **/

/**
 * Validates and confirms the attachment file from C2 with the same
 * step-by-step flow used by generatePDFBundleWithLabels().
 *
 * Returns:
 *   - null  → user cancelled or a hard error occurred (caller should return)
 *   - false → C2 was empty, no attachment (caller should proceed without one)
 *   - { fileId, fileName, file, isGoogleDoc, templateDoc }
 *             → confirmed file ready to use
 */
function validateAndConfirmAttachment(ui, attachmentRef) {
  if (!attachmentRef) return false; // No attachment specified

  // Step 1: show what's in C2 and attempt to extract the file ID
  ui.alert(
    'Step 1: Checking Attachment File\n\n' +
    'Value in "email details" sheet, cell C2:\n' +
    attachmentRef.substring(0, 100) + (attachmentRef.length > 100 ? '...' : '') + '\n\n' +
    'Extracting file ID...'
  );

  const fileId = extractDriveId(attachmentRef);
  if (!fileId) {
    ui.alert(
      'Error: Could not extract file ID from C2.\n\n' +
      'Value in C2: ' + attachmentRef + '\n\n' +
      'Please use one of these formats:\n' +
      '• Full Drive URL: https://drive.google.com/file/d/FILE_ID/view\n' +
      '• Google Doc URL: https://docs.google.com/document/d/FILE_ID/edit\n' +
      '• Just the file ID: 1a2b3c4d5e6f7g8h9i0j'
    );
    return null;
  }

  // Look up file name and type before asking for confirmation
  let fileName = '';
  let fileType = '';
  let driveFile;
  try {
    driveFile = DriveApp.getFileById(fileId);
    fileName = driveFile.getName();
    fileType = driveFile.getMimeType();
  } catch (e) {
    ui.alert(
      'Error: Cannot access file with ID: ' + fileId + '\n\n' +
      'Error: ' + e.message + '\n\n' +
      'Make sure you have access to this file.'
    );
    return null;
  }

  const isGoogleDoc = fileType === 'application/vnd.google-apps.document';
  const typeLabel = isGoogleDoc
    ? 'Google Doc (will generate a personalized PDF for each recipient)'
    : 'File type: ' + fileType + ' (will attach the same file to every email)';

  // Step 2: confirm with user
  const confirm = ui.alert(
    'Step 2: Confirm Attachment File',
    'Found this file:\n\n' +
    '📄 File Name: ' + fileName + '\n' +
    typeLabel + '\n\n' +
    'File ID: ' + fileId + '\n\n' +
    'Is "' + fileName + '" the correct attachment?',
    ui.ButtonSet.YES_NO
  );

  if (confirm === ui.Button.NO) {
    ui.alert(
      'Cancelled.\n\n' +
      'Please update cell C2 in "email details" sheet with the correct file URL.\n\n' +
      'Currently has: ' + fileName
    );
    return null;
  }

  // Open the Google Doc template if applicable
  let templateDoc = null;
  if (isGoogleDoc) {
    try {
      templateDoc = DocumentApp.openById(fileId);
    } catch (e) {
      ui.alert('Error: Cannot open Google Doc\n\nDocument: ' + fileName + '\nError: ' + e.message);
      return null;
    }
  }

  return { fileId, fileName, file: driveFile, isGoogleDoc, templateDoc };
}

/**
 * Builds the attachment blob for one recipient.
 * If the attachment is a Google Doc, generates a personalized PDF using
 * the same placeholder replacement as generatePDFBundleWithLabels().
 * Otherwise returns the file converted to PDF (static, same for all).
 */
function buildAttachmentBlob(attachmentInfo, personData) {
  if (!attachmentInfo) return null;

  if (attachmentInfo.isGoogleDoc) {
    // Personalized PDF – delegates to PDFBundler's createPersonalizedPDF()
    try {
      return createPersonalizedPDF(attachmentInfo.templateDoc, personData);
    } catch (e) {
      Logger.log('Could not generate personalized PDF for ' + personData.fullName + ': ' + e.message);
      return null;
    }
  }

  // Static file – convert to PDF and attach the same blob to every email
  try {
    return attachmentInfo.file.getAs(MimeType.PDF);
  } catch (e) {
    Logger.log('Could not convert attachment to PDF: ' + e.message);
    return null;
  }
}

/** =================== CC ADDRESS HELPER ====================== **/
/**
 * Gets CC email addresses from column D of email details sheet (D2, D3, D4, etc.)
 * Returns comma-separated string of addresses, or empty string if none found
 */
function getCCAddresses(detailsSheet) {
  const lastRow = detailsSheet.getLastRow();
  if (lastRow < 2) return ''; // No data rows

  // Read all values in column D starting from D2
  const ccValues = detailsSheet.getRange(2, 4, lastRow - 1, 1).getValues();

  // Filter out empty cells and trim
  const addresses = ccValues
    .map(row => String(row[0] || '').trim())
    .filter(addr => addr !== '' && addr.includes('@')); // Basic email validation

  if (addresses.length === 0) return '';

  // Join multiple addresses with comma
  return addresses.join(',');
}


////////////////////////////////////////////////////////////////////////////
// src/BounceChecker.gs
////////////////////////////////////////////////////////////////////////////

/**
 * BounceChecker.gs
 * Scans Gmail for bounce notifications (DSNs) and records them on a "Bounced" sheet.
 *
 * NOTE: a simple onOpen(e) trigger runs without authorization and cannot touch
 * Gmail. The on-open scan therefore runs through an INSTALLABLE trigger, which
 * the user opts into via the menu (enableBounceAutoCheck).
 */

/** ========================== CONFIG ========================== **/
const BOUNCE_SHEET = 'Bounced';
const BOUNCE_LAST_SCAN_KEY = 'lastBounceScanIso';
const BOUNCE_TRIGGER_FN = 'scanBouncesOnOpen';
const BOUNCE_FIRST_RUN_LOOKBACK_DAYS = 30;
const BOUNCE_MAX_THREADS = 200;

/** ========================== MENU ENTRY POINTS =============== **/
/**
 * Menu item: scan for bounces and report what was found.
 */
function checkForBounces() {
  const ui = SpreadsheetApp.getUi();
  const result = runBounceScan();

  if (result.error) {
    ui.alert('Bounce check failed:\n\n' + result.error);
    return;
  }

  const lines = [
    'Scanned ' + result.scanned + ' bounce notice(s) since ' + result.since + '.',
    '',
    'New bounces added: ' + result.added,
    'Already recorded: ' + result.duplicates
  ];
  if (result.unmatched) {
    lines.push('Ignored (not on your People sheet): ' + result.unmatched);
  }
  ui.alert(lines.join('\n'));
}

/**
 * Installable-trigger target. Runs silently so opening the sheet is not
 * interrupted by a dialog.
 */
function scanBouncesOnOpen() {
  try {
    runBounceScan();
  } catch (e) {
    console.error('Auto bounce scan failed: ' + e);
  }
}

/**
 * Menu item: install the on-open trigger.
 */
function enableBounceAutoCheck() {
  const ui = SpreadsheetApp.getUi();
  if (getBounceTrigger()) {
    ui.alert('Auto bounce check is already enabled.');
    return;
  }

  ScriptApp.newTrigger(BOUNCE_TRIGGER_FN)
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();

  ui.alert(
    'Auto bounce check enabled.\n\n' +
    'Bounces will be scanned each time this spreadsheet is opened. ' +
    'This adds a few seconds to load time — disable it from the same menu ' +
    'if that becomes annoying.'
  );
}

/**
 * Menu item: remove the on-open trigger.
 */
function disableBounceAutoCheck() {
  const ui = SpreadsheetApp.getUi();
  const trigger = getBounceTrigger();
  if (!trigger) {
    ui.alert('Auto bounce check is not currently enabled.');
    return;
  }
  ScriptApp.deleteTrigger(trigger);
  ui.alert('Auto bounce check disabled.');
}

/**
 * Finds this spreadsheet's installed bounce trigger, if any.
 */
function getBounceTrigger() {
  const triggers = ScriptApp.getUserTriggers(SpreadsheetApp.getActive());
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === BOUNCE_TRIGGER_FN) return triggers[i];
  }
  return null;
}

/** ========================== CORE SCAN ======================= **/
/**
 * Searches Gmail for bounce notices, matches them against the People sheet,
 * and appends any new ones to the Bounced sheet.
 *
 * Returns { scanned, added, duplicates, unmatched, since, error }.
 */
function runBounceScan() {
  const ss = SpreadsheetApp.getActive();

  let sinceDate;
  try {
    sinceDate = getBounceScanStart();
  } catch (e) {
    return { error: String(e) };
  }

  const query = buildBounceQuery(sinceDate);

  let threads;
  try {
    threads = GmailApp.search(query, 0, BOUNCE_MAX_THREADS);
  } catch (e) {
    return { error: 'Could not search Gmail: ' + e };
  }

  const recipients = buildRecipientMap(ss);
  const sheet = getOrCreateBounceSheet(ss);
  const known = buildKnownBounceSet(sheet);

  const rows = [];
  let scanned = 0;
  let duplicates = 0;
  let unmatched = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const parsed = parseBounceMessage(message);
      if (!parsed.email) return;

      scanned++;
      const key = parsed.email.toLowerCase();

      // Only report addresses that are actually on the People sheet — Gmail
      // returns unrelated mailer-daemon traffic too.
      if (!recipients.hasOwnProperty(key)) {
        unmatched++;
        return;
      }
      if (known[key]) {
        duplicates++;
        return;
      }

      known[key] = true;
      rows.push([
        parsed.email,
        recipients[key],
        message.getDate(),
        parsed.type,
        parsed.reason,
        'https://mail.google.com/mail/u/0/#inbox/' + thread.getId()
      ]);
    });
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    sheet.autoResizeColumns(1, 5);
  }

  setBounceScanStart(new Date());

  return {
    scanned: scanned,
    added: rows.length,
    duplicates: duplicates,
    unmatched: unmatched,
    since: Utilities.formatDate(sinceDate, Session.getScriptTimeZone(), 'MMM d, yyyy')
  };
}

/**
 * Gmail query covering the common bounce senders since the given date.
 */
function buildBounceQuery(sinceDate) {
  const after = Utilities.formatDate(sinceDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  return '(from:mailer-daemon OR from:postmaster OR ' +
         'subject:"Delivery Status Notification" OR ' +
         'subject:"Undelivered Mail Returned to Sender" OR ' +
         'subject:"Address not found") after:' + after;
}

/**
 * Start of the scan window: the last scan, or a default lookback on first run.
 */
function getBounceScanStart() {
  const stored = PropertiesService.getDocumentProperties().getProperty(BOUNCE_LAST_SCAN_KEY);
  if (stored) {
    const parsed = new Date(stored);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  const fallback = new Date();
  fallback.setDate(fallback.getDate() - BOUNCE_FIRST_RUN_LOOKBACK_DAYS);
  return fallback;
}

/**
 * Records when the last successful scan finished. Backs up one day so a
 * bounce that lands mid-scan is not skipped next time.
 */
function setBounceScanStart(date) {
  const safe = new Date(date.getTime());
  safe.setDate(safe.getDate() - 1);
  PropertiesService.getDocumentProperties()
    .setProperty(BOUNCE_LAST_SCAN_KEY, safe.toISOString());
}

/** ========================== PARSING ========================= **/
/**
 * Pulls the failed recipient, failure type and reason out of a bounce message.
 *
 * Prefers the machine-readable RFC 3464 delivery-status part, which is only
 * present in the raw MIME, and falls back to Gmail's human-readable wording.
 */
function parseBounceMessage(message) {
  const result = { email: '', type: '', reason: '' };

  let raw = '';
  try {
    raw = message.getRawContent() || '';
  } catch (e) {
    raw = '';
  }

  // Preferred: RFC 3464 delivery-status fields.
  const finalRecipient = raw.match(/Final-Recipient:\s*rfc822;\s*([^\s<>]+@[^\s<>]+)/i) ||
                         raw.match(/Original-Recipient:\s*rfc822;\s*([^\s<>]+@[^\s<>]+)/i);
  if (finalRecipient) result.email = cleanBounceAddress(finalRecipient[1]);

  const status = raw.match(/^Status:\s*([245])\.(\d+)\.(\d+)/mi);
  if (status) {
    result.type = status[1] === '4' ? 'Temporary (' + status[0].replace(/^Status:\s*/i, '') + ')'
                                    : 'Permanent (' + status[0].replace(/^Status:\s*/i, '') + ')';
  }

  const diagnostic = raw.match(/Diagnostic-Code:\s*smtp;\s*([^\r\n]+)/i);
  if (diagnostic) result.reason = diagnostic[1].trim();

  // Fallback: scrape Gmail's plain-text wording.
  if (!result.email || !result.reason) {
    let body = '';
    try {
      body = message.getPlainBody() || '';
    } catch (e) {
      body = '';
    }

    if (!result.email) {
      const wording = body.match(/(?:wasn't delivered to|was not delivered to|could not be delivered to)\s*[:\s]*([^\s<>]+@[^\s<>]+)/i);
      if (wording) result.email = cleanBounceAddress(wording[1]);
    }
    if (!result.reason) {
      const because = body.match(/because\s+([^\r\n.]{10,160})/i);
      if (because) result.reason = because[1].trim();
    }
  }

  if (!result.type && result.email) result.type = 'Unknown';
  if (!result.reason) result.reason = 'No reason reported';

  return result;
}

/**
 * Strips angle brackets and trailing punctuation from a parsed address.
 */
function cleanBounceAddress(value) {
  return String(value || '')
    .replace(/[<>]/g, '')
    .replace(/[.,;:]+$/, '')
    .trim();
}

/** ========================== SHEET HELPERS =================== **/
/**
 * Maps lowercased recipient email -> name, from the People sheet.
 */
function buildRecipientMap(ss) {
  const map = {};
  const sheet = ss.getSheetByName(LIST_SHEET);
  if (!sheet) return map;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  const values = sheet.getRange(2, 1, lastRow - 1, EMAIL_COL).getDisplayValues();
  values.forEach(row => {
    const email = String(row[EMAIL_COL - 1] || '').trim().toLowerCase();
    if (!email) return;
    map[email] = String(row[NAME_COL - 1] || '').trim();
  });
  return map;
}

/**
 * Set of emails already present on the Bounced sheet, for deduping.
 */
function buildKnownBounceSet(sheet) {
  const known = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return known;

  sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().forEach(row => {
    const email = String(row[0] || '').trim().toLowerCase();
    if (email) known[email] = true;
  });
  return known;
}

/**
 * Returns the Bounced sheet, creating and formatting it when absent.
 */
function getOrCreateBounceSheet(ss) {
  let sheet = ss.getSheetByName(BOUNCE_SHEET);
  if (!sheet) sheet = ss.insertSheet(BOUNCE_SHEET);

  const headers = ['Email', 'Name', 'Bounced On', 'Type', 'Reason', 'Message'];
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    if (!existing[i]) sheet.getRange(1, i + 1).setValue(headers[i]);
  }

  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 220); // Email
  sheet.setColumnWidth(2, 160); // Name
  sheet.setColumnWidth(3, 130); // Bounced On
  sheet.setColumnWidth(4, 150); // Type
  sheet.setColumnWidth(5, 320); // Reason
  sheet.setColumnWidth(6, 240); // Message

  return sheet;
}


////////////////////////////////////////////////////////////////////////////
// src/EmailFinderGoogle.gs
////////////////////////////////////////////////////////////////////////////

/**
 * EmailFinderGoogle.gs
 * Automatically finds missing emails by searching Google Contacts and Gmail history
 */

/** ========================== CONFIG ========================== **/
const GOOGLE_SHEET_NAME = 'People';  // Sheet to fill emails
const GOOGLE_NAME_COL = 1;           // Column A: Name
const GOOGLE_EMAIL_COL = 3;          // Column C: Email

/** ========================== MAIN FUNCTION =================== **/
/**
 * Fills missing emails by searching all available sources
 * Called from menu: Fill Emails from Google Contacts
 */
function fillEmailsFromGoogleContacts() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(GOOGLE_SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + GOOGLE_SHEET_NAME + '" not found');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data rows');
    return;
  }

  const width = Math.max(GOOGLE_NAME_COL, GOOGLE_EMAIL_COL);
  const values = sh.getRange(2, 1, lastRow - 1, width).getDisplayValues();

  let filled = 0;
  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    const fullName = String(values[i][GOOGLE_NAME_COL - 1] || '').trim();
    const currentEmail = String(values[i][GOOGLE_EMAIL_COL - 1] || '').trim();

    if (!fullName || currentEmail) continue;

    const email = findBestEmailByName(fullName);
    if (email) {
      sh.getRange(rowNumber, GOOGLE_EMAIL_COL).setValue(email);
      filled++;
    }
  }

  SpreadsheetApp.getUi().alert('Emails filled: ' + filled);
}

/** ========================== EMAIL LOOKUP ==================== **/
/**
 * Lookup flow: Contacts → Other Contacts → Legacy → Gmail history
 */
function findBestEmailByName(name) {
  const fromContacts = searchPeopleContacts(name);
  if (fromContacts) return fromContacts;

  const fromOther = searchOtherContacts(name);
  if (fromOther) return fromOther;

  const fromLegacy = searchLegacyContacts(name);
  if (fromLegacy) return fromLegacy;

  const fromGmail = searchGmailHistory(name);
  if (fromGmail) return fromGmail;

  return '';
}

/** ========================== PEOPLE API CONTACTS ============= **/
/**
 * Searches People API saved Contacts
 */
function searchPeopleContacts(name) {
  try {
    const resp = People.People.searchContacts({
      query: name,
      pageSize: 10,
      readMask: 'names,emailAddresses'
    });
    return pickBestCandidate(name, resp && resp.results);
  } catch (e) {}
  return '';
}

/**
 * Searches People API Other contacts
 */
function searchOtherContacts(name) {
  try {
    const resp = People.OtherContacts.search({
      query: name,
      pageSize: 10,
      readMask: 'names,emailAddresses'
    });
    return pickBestCandidate(name, resp && resp.results);
  } catch (e) {}
  return '';
}

/** ========================== LEGACY CONTACTS ================= **/
/**
 * Legacy ContactsApp fallback
 */
function searchLegacyContacts(name) {
  try {
    const matches = ContactsApp.getContactsByName(name) || [];
    for (const c of matches) {
      const emails = c.getEmails();
      if (emails && emails.length) {
        const primary = emails.find(e => e.isPrimary());
        const addr = (primary || emails[0]).getAddress();
        if (addr) return addr.trim();
      }
    }
  } catch (e) {}
  return '';
}

/** ========================== GMAIL HISTORY =================== **/
/**
 * Gmail history heuristic - searches recent messages
 */
function searchGmailHistory(name) {
  try {
    const query = `from:(${name}) OR to:(${name}) OR cc:(${name})`;
    const threads = GmailApp.search(query, 0, 30);
    const counts = {};
    const normName = normalizeName(name);

    for (const th of threads) {
      for (const m of th.getMessages()) {
        collectHeaderEmails(m.getFrom(), normName, counts);
        m.getTo().split(',').forEach(s => collectHeaderEmails(s, normName, counts));
        m.getCc().split(',').forEach(s => collectHeaderEmails(s, normName, counts));
        m.getBcc().split(',').forEach(s => collectHeaderEmails(s, normName, counts));
      }
    }

    const best = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
    return best ? best[0] : '';
  } catch (e) {}
  return '';
}

/** ========================== CANDIDATE SELECTION ============= **/
/**
 * Picks best candidate email from People API results
 */
function pickBestCandidate(queryName, results) {
  if (!results || !results.length) return '';
  const candidates = [];
  const q = normalizeName(queryName);

  results.forEach(r => {
    const p = r.person;
    if (!p || !p.emailAddresses) return;
    const display = bestDisplayName(p.names || []);
    const score = nameScore(q, normalizeName(display));
    p.emailAddresses.forEach(e => {
      const value = (e.value || '').trim();
      if (!value) return;
      const primary = e.metadata && e.metadata.primary ? 1 : 0;
      candidates.push({ value, score, primary });
    });
  });

  candidates.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    if (b.primary !== a.primary) return b.primary - a.primary;
    return a.value.localeCompare(b.value);
  });
  return candidates[0] ? candidates[0].value : '';
}

/**
 * Gets best display name from People API names array
 */
function bestDisplayName(names) {
  let best = '';
  for (const n of names) {
    if (n.metadata && n.metadata.primary && n.displayName) return n.displayName;
    if (!best && n.displayName) best = n.displayName;
  }
  return best;
}

/** ========================== SCORING & MATCHING ============== **/
/**
 * Scores how well two normalized names match
 */
function nameScore(q, d) {
  let s = 0;
  if (d.includes(q)) s += 2;
  const qFirst = q.split(' ')[0] || '';
  if (qFirst && new RegExp('\\b' + qFirst.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b').test(d)) s += 1;
  return s;
}

/**
 * Checks if two normalized names likely match
 */
function nameLikelyMatch(q, d) {
  if (d.includes(q)) return true;
  const qParts = q.split(' ').filter(Boolean);
  const dParts = d.split(' ').filter(Boolean);
  if (qParts.length >= 2) {
    const first = qParts[0];
    const last = qParts[qParts.length - 1];
    return dParts.includes(first) && dParts.includes(last);
  }
  return dParts.includes(qParts[0] || '');
}

/**
 * Collects email addresses from message headers
 */
function collectHeaderEmails(headerStr, normTargetName, counts) {
  const parts = String(headerStr || '').split(',');
  for (let raw of parts) {
    raw = raw.trim();
    if (!raw) continue;
    const m = raw.match(/^(.*)<([^>]+)>$/);
    let disp = '';
    let addr = '';
    if (m) {
      disp = m[1].trim();
      addr = m[2].trim();
    } else {
      addr = raw;
    }
    if (disp) {
      const nd = normalizeName(disp);
      if (!nd || !nameLikelyMatch(normTargetName, nd)) continue;
    }
    if (addr && addr.includes('@')) {
      counts[addr] = (counts[addr] || 0) + 1;
    }
  }
}


////////////////////////////////////////////////////////////////////////////
// src/EmailFinderVCF.gs
////////////////////////////////////////////////////////////////////////////

/**
 * EmailFinderVCF.gs
 * Fills missing emails and phone numbers from VCF files stored in Google Drive
 */

/** ========================== CONFIG ========================== **/
const VCF_SHEET_NAME = 'People';  // Sheet to fill emails/phones
const VCF_NAME_COL = 1;           // Column A: Name
const VCF_EMAIL_COL = 3;          // Column C: Email
const VCF_PHONE_COL = 4;          // Column D: Phone

/** ========================== MAIN FUNCTION =================== **/
/**
 * Prompts user for VCF file from Drive and fills missing emails/phones
 * Called from menu: Fill Emails from VCF File
 */
function fillEmailsFromVCF() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'VCF file',
    'Paste the Google Drive file URL or file ID for your .vcf and click OK.',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const input = String(resp.getResponseText() || '').trim();
  if (!input) { ui.alert('No input provided.'); return; }

  const fileId = extractDriveFileId(input);
  if (!fileId) {
    ui.alert('Could not extract a Drive file ID. Paste the full Drive URL or the file ID.');
    return;
  }

  let nameToEmail, nameToPhone;
  try {
    nameToEmail = buildVcfEmailMap(fileId); // normalized name -> email
    nameToPhone = buildVcfPhoneMap(fileId); // normalized name -> phone
  } catch (e) {
    ui.alert('Error reading VCF: ' + e.message);
    return;
  }

  if ((!nameToEmail || Object.keys(nameToEmail).length === 0) &&
      (!nameToPhone || Object.keys(nameToPhone).length === 0)) {
    ui.alert('No contacts with emails or phones found in that VCF.');
    return;
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(VCF_SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + VCF_SHEET_NAME + '" not found');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) { ui.alert('No data rows.'); return; }

  const width = Math.max(VCF_NAME_COL, VCF_EMAIL_COL, VCF_PHONE_COL);
  const values = sh.getRange(2, 1, lastRow - 1, width).getDisplayValues();

  let filledEmails = 0;
  let filledPhones = 0;
  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    const fullName = String(values[i][VCF_NAME_COL - 1] || '').trim();
    const currentEmail = String(values[i][VCF_EMAIL_COL - 1] || '').trim();
    const currentPhone = String(values[i][VCF_PHONE_COL - 1] || '').trim();
    if (!fullName) continue;

    // Fill email if blank
    if (!currentEmail) {
      const email = matchEmailFromMap(fullName, nameToEmail);
      if (email) {
        sh.getRange(rowNumber, VCF_EMAIL_COL).setValue(email);
        filledEmails++;
      }
    }

    // Fill phone if blank
    if (!currentPhone) {
      const phone = matchPhoneFromMap(fullName, nameToPhone);
      if (phone) {
        sh.getRange(rowNumber, VCF_PHONE_COL).setValue(phone);
        filledPhones++;
      }
    }
  }
  ui.alert('Filled from VCF\nEmails: ' + filledEmails + '\nPhones: ' + filledPhones);
}

/** ========================== VCF PARSING ===================== **/

/**
 * Builds map of normalized name -> email from VCF file
 */
function buildVcfEmailMap(vcfFileId) {
  const file = DriveApp.getFileById(vcfFileId);
  const text = file.getBlob().getDataAsString('UTF-8');
  const cards = text.split(/END:VCARD/i);
  const map = {};

  cards.forEach(blockRaw => {
    const block = unfoldVcardLines(blockRaw);
    const email = extractVcfEmail(block);
    if (!email) return;
    const name = extractVcfName(block);
    if (!name) return;

    nameVariants(name).forEach(v => {
      const key = norm(v);
      if (key && !map[key]) map[key] = email; // first email wins
    });
  });
  return map;
}

/**
 * Builds map of normalized name -> phone from VCF file
 */
function buildVcfPhoneMap(vcfFileId) {
  const file = DriveApp.getFileById(vcfFileId);
  const text = file.getBlob().getDataAsString('UTF-8');
  const cards = text.split(/END:VCARD/i);
  const map = {};

  cards.forEach(blockRaw => {
    const block = unfoldVcardLines(blockRaw);
    const phone = extractVcfBestPhone(block);
    if (!phone) return;
    const name = extractVcfName(block);
    if (!name) return;

    nameVariants(name).forEach(v => {
      const key = norm(v);
      if (key && !map[key]) map[key] = phone; // first phone wins
    });
  });
  return map;
}

/**
 * Unfolds VCard lines (joins continuation lines)
 */
function unfoldVcardLines(s) {
  return String(s || '')
    .replace(/\r\n/g, '\n')
    .replace(/\n[ \t]/g, ''); // join folded lines
}

/**
 * Extracts email from VCard block
 */
function extractVcfEmail(block) {
  const re = /^\s*EMAIL(?:;[^:]+)?:\s*([^ \t\r\n;]+)\s*$/gim;
  const m = re.exec(block);
  return m ? String(m[1]).trim() : '';
}

/**
 * Extracts best phone number from VCard block (prioritizes mobile)
 */
function extractVcfBestPhone(block) {
  const re = /^\s*TEL(?:;([^:]+))?:\s*([^\s]+)\s*$/gim;
  const found = [];
  let m;

  while ((m = re.exec(block)) !== null) {
    const params = (m[1] || '').toLowerCase();
    const value = (m[2] || '').trim();
    if (!value) continue;

    const score =
      (params.includes('cell') || params.includes('mobile') || params.includes('iphone') ? 3 : 0) +
      (params.includes('work') ? 1 : 0) +
      (params.includes('home') ? 0 : 0);
    found.push({ value, score });
  }

  if (!found.length) return '';
  found.sort((a, b) => b.score - a.score);
  return found[0].value;
}

/**
 * Extracts name from VCard block (FN or N field)
 */
function extractVcfName(block) {
  let m = /^\s*FN:\s*(.+?)\s*$/gim.exec(block);
  if (m) return m[1].trim();

  m = /^\s*N:\s*([^;\n\r]*);([^;\n\r]*)/gim.exec(block);
  if (m) {
    const family = (m[1] || '').trim();
    const given = (m[2] || '').trim();
    return [given, family].filter(Boolean).join(' ');
  }
  return '';
}

/** ========================== NAME MATCHING =================== **/

/**
 * Generates name variants for matching (First Last, Last First, etc.)
 */
function nameVariants(name) {
  const clean = cleanName(name);
  const parts = clean.split(/\s+/).filter(Boolean);
  const out = new Set();
  if (parts.length === 0) return [];

  out.add(parts.join(' ')); // full form
  if (parts.length >= 2) {
    out.add(parts[0] + ' ' + parts[parts.length - 1]);   // First Last
    out.add(parts[parts.length - 1] + ', ' + parts[0]);   // Last, First
  }
  out.add(parts[0]); // First only
  return Array.from(out);
}

/**
 * Normalizes name for map key (lowercase, alphanumeric only)
 */
function norm(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Matches email from map using name variants
 */
function matchEmailFromMap(fullName, map) {
  if (!map) return '';
  const variants = nameVariants(fullName).map(norm);
  for (const v of variants) {
    if (v && map[v]) return map[v];
  }
  return '';
}

/**
 * Matches phone from map using name variants
 */
function matchPhoneFromMap(fullName, map) {
  if (!map) return '';
  const variants = nameVariants(fullName).map(norm);
  for (const v of variants) {
    if (v && map[v]) return map[v];
  }
  return '';
}

/** ========================== DRIVE HELPERS =================== **/

/**
 * Extracts Drive file ID from URL or string
 */
function extractDriveFileId(s) {
  const m =
    s.match(/\/d\/([A-Za-z0-9_-]{20,})\//) ||
    s.match(/id=([A-Za-z0-9_-]{20,})/) ||
    s.match(/^([A-Za-z0-9_-]{20,})$/);
  return m ? m[1] : '';
}


////////////////////////////////////////////////////////////////////////////
// src/PDFBundler.gs
////////////////////////////////////////////////////////////////////////////

/**
 * PDFBundler.gs
 * Generates a folder with PDFs and a printable labels document for mailing
 */

/** ========================== CONFIG ========================== **/
const BUNDLE_SHEET_NAME = 'People';
const BUNDLE_DETAILS_SHEET = 'email details';
const BUNDLE_NAME_COL = 1;      // A: Name
const BUNDLE_PAC_COL = 2;       // B: PAC Names
const BUNDLE_EMAIL_COL = 3;     // C: Email (not used for PDF bundle)
const BUNDLE_PHONE_COL = 4;     // D: Phone (not used for PDF bundle)
const BUNDLE_ADDRESS_COL = 5;   // E: Address

// Avery 5160 label dimensions (30 labels per page, 3 columns x 10 rows)
const LABEL_WIDTH = 2.625;      // inches
const LABEL_HEIGHT = 1.0;       // inches
const LABELS_PER_ROW = 3;
const LABELS_PER_PAGE = 30;
const LABEL_MARGIN = 0.15;      // inches

/** ========================== MAIN FUNCTION =================== **/
/**
 * Creates a Drive folder with PDFs and generates printable mailing labels
 * Called from menu: Generate PDF Bundle & Labels
 */
function generatePDFBundleWithLabels() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  // Get the PDF template file
  const detailsSh = ss.getSheetByName(BUNDLE_DETAILS_SHEET);
  if (!detailsSh) {
    ui.alert('Error: Sheet "email details" not found.');
    return;
  }

  const attachmentRef = String(detailsSh.getRange('C2').getValue() || '').trim();
  if (!attachmentRef) {
    ui.alert('Error: No template document specified in email details C2.\n\n' +
             'Please add a Google Docs URL or file ID.\n' +
             'The document should contain placeholders like [FIRST NAME], [FULL NAME], [PAC NAME], etc.');
    return;
  }

  // Show what's actually in C2 (raw value)
  ui.alert('Step 1: Checking Template File\n\n' +
           'Value in "email details" sheet, cell C2:\n' +
           attachmentRef.substring(0, 100) + (attachmentRef.length > 100 ? '...' : '') + '\n\n' +
           'Extracting file ID...');

  Logger.log('Template reference from email details C2: ' + attachmentRef);

  // Extract the file ID
  const fileId = extractDriveId(attachmentRef);
  if (!fileId) {
    ui.alert('Error: Could not extract file ID from C2.\n\n' +
             'Value in C2: ' + attachmentRef + '\n\n' +
             'Please use one of these formats:\n' +
             '• Full Drive URL: https://docs.google.com/document/d/FILE_ID/edit\n' +
             '• Just the file ID: 1a2b3c4d5e6f7g8h9i0j');
    return;
  }

  Logger.log('Extracted file ID: ' + fileId);

  // Get the actual file name from Drive FIRST (before opening)
  let templateFileName = '';
  let fileType = '';
  try {
    const driveFile = DriveApp.getFileById(fileId);
    templateFileName = driveFile.getName();
    fileType = driveFile.getMimeType();
    Logger.log('Found file in Drive: ' + templateFileName + ' (type: ' + fileType + ')');
  } catch (e) {
    ui.alert('Error: Cannot access file with ID: ' + fileId + '\n\n' +
             'Error: ' + e.message + '\n\n' +
             'Make sure you have access to this file.');
    return;
  }

  // Check if it's actually a Google Doc
  if (fileType !== 'application/vnd.google-apps.document') {
    ui.alert('Error: Wrong file type!\n\n' +
             'File name: ' + templateFileName + '\n' +
             'File type: ' + fileType + '\n\n' +
             'This is NOT a Google Doc. Please use a Google Docs document as the template.');
    return;
  }

  // Now confirm with user using the actual document name
  const confirm = ui.alert(
    'Step 2: Confirm Template Document',
    'Found this Google Doc:\n\n' +
    '📄 Document Name: ' + templateFileName + '\n\n' +
    'File ID: ' + fileId + '\n\n' +
    'Is "' + templateFileName + '" the correct template?',
    ui.ButtonSet.YES_NO
  );

  if (confirm === ui.Button.NO) {
    ui.alert('Cancelled.\n\n' +
             'Please update cell C2 in "email details" sheet with the correct Google Doc URL.\n\n' +
             'Currently has: ' + templateFileName);
    return;
  }

  // Finally, open the document
  let templateDoc;
  try {
    templateDoc = DocumentApp.openById(fileId);
    Logger.log('Successfully opened template document: ' + templateFileName);
  } catch (e) {
    ui.alert('Error: Cannot open document\n\n' +
             'Document: ' + templateFileName + '\n' +
             'Error: ' + e.message);
    return;
  }

  // Get people data
  const listSh = ss.getSheetByName(BUNDLE_SHEET_NAME);
  if (!listSh) {
    ui.alert('Error: Sheet "People" not found.');
    return;
  }

  const lastRow = listSh.getLastRow();
  if (lastRow < 2) {
    ui.alert('No data rows found.');
    return;
  }

  const width = 5; // Read all 5 columns (A-E)
  const values = listSh.getRange(2, 1, lastRow - 1, width).getDisplayValues();

  // Filter valid rows (must have address; name falls back to "To Whom It May Concern")
  const people = [];
  values.forEach(row => {
    const fullName = String(row[BUNDLE_NAME_COL - 1] || '').trim() || 'To Whom It May Concern';
    const pacName = String(row[BUNDLE_PAC_COL - 1] || '').trim();
    const email = String(row[BUNDLE_EMAIL_COL - 1] || '').trim();
    const phone = String(row[BUNDLE_PHONE_COL - 1] || '').trim();
    const address = String(row[BUNDLE_ADDRESS_COL - 1] || '').trim();

    if (!address) return;

    const firstName = fullName === 'To Whom It May Concern' ? fullName : extractFirstName(fullName);

    // Build person data object
    let personData = {
      fullName: fullName,
      firstName: firstName,
      pacName: pacName,
      email: email,
      phone: phone,
      address: address
    };

    // Normalize capitalization (ALL CAPS → Title Case)
    personData = normalizePersonData(personData);

    people.push(personData);
  });

  if (people.length === 0) {
    ui.alert('No valid records found. Each row must have a Name (column A) and Address (column E).');
    return;
  }

  // Create folder with timestamp in the same location as the template
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmmss');
  const folderName = `PDF Bundle ${timestamp}`;

  // Get the parent folder of the template document
  const templateFile = DriveApp.getFileById(templateDoc.getId());
  const templateParents = templateFile.getParents();

  let folder;
  if (templateParents.hasNext()) {
    // Create the bundle folder in the same folder as the template
    const templateFolder = templateParents.next();
    folder = templateFolder.createFolder(folderName);
    Logger.log('Created PDF bundle folder in: ' + templateFolder.getName());
  } else {
    // Fallback: create in root if template has no parent (shouldn't happen)
    folder = DriveApp.createFolder(folderName);
    Logger.log('Created PDF bundle folder in Drive root (template has no parent folder)');
  }

  // Generate personalized PDFs
  let generatedCount = 0;
  people.forEach(person => {
    try {
      // Create personalized PDF for this person
      const pdfBlob = createPersonalizedPDF(templateDoc, person);
      const fileName = `${sanitizeFileName(person.fullName)}.pdf`;

      // Save PDF to folder
      folder.createFile(pdfBlob.setName(fileName));
      generatedCount++;
    } catch (e) {
      Logger.log('Error generating PDF for ' + person.fullName + ': ' + e.message);
    }
  });

  // Generate combined PDF with all letters
  const combinedPdfCreated = generateCombinedPDF(templateDoc, people, folderName, folder);

  // Generate labels PDF
  const labelsPdfCreated = generateLabelsPDF(people, folderName, folder);

  // Get template parent folder name for display (reuse templateFile from above)
  const parentFolderName = folder.getParents().hasNext() ? folder.getParents().next().getName() : 'Drive Root';

  // Show completion message
  if (labelsPdfCreated && combinedPdfCreated) {
    ui.alert(
      'PDF Bundle Created Successfully!\n\n' +
      '✓ Individual PDFs: ' + generatedCount + '\n' +
      '✓ Combined PDF: Created\n' +
      '✓ Labels PDF: Created\n' +
      '✓ Total people: ' + people.length + '\n\n' +
      'Created in same folder as template:\n' +
      '📁 ' + parentFolderName + '\n\n' +
      'Folder Name: ' + folderName + '\n' +
      'Folder URL: ' + folder.getUrl() + '\n\n' +
      'Files in folder:\n' +
      '• ' + generatedCount + ' personalized PDFs (one per person)\n' +
      '• Combined Letters.pdf (all letters in one file)\n' +
      '• Mailing Labels.pdf (print on Avery 5160 sheets)\n\n' +
      'Click the folder URL above to open it.'
    );
  } else {
    const warnings = [];
    if (!combinedPdfCreated) warnings.push('Combined PDF: FAILED');
    if (!labelsPdfCreated) warnings.push('Labels PDF: FAILED');

    ui.alert(
      'PDF Bundle Partially Created\n\n' +
      'Individual PDFs: ' + generatedCount + '\n' +
      (warnings.length > 0 ? warnings.join('\n') + '\n\n' : '') +
      'Created in: ' + parentFolderName + '\n' +
      'Folder: ' + folderName + '\n' +
      'Location: ' + folder.getUrl() + '\n\n' +
      'Check View → Logs for error details.'
    );
  }
}

/** ========================== DEBUG FUNCTION ================== **/

/**
 * DEBUG: Test replacement on ONE person and show document structure
 * Run this from Script Editor to diagnose the issue
 */
function debugReplacementForOnePerson() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  // Get template
  const detailsSh = ss.getSheetByName('email details');
  const attachmentRef = String(detailsSh.getRange('C2').getValue()).trim();
  const fileId = extractDriveId(attachmentRef);
  const templateDoc = DocumentApp.openById(fileId);

  // Get first person with ADDRESS LINE 2
  const listSh = ss.getSheetByName('People');
  const values = listSh.getRange(2, 1, listSh.getLastRow() - 1, 5).getDisplayValues();

  Logger.log('=== CHECKING ALL ADDRESSES ===');
  let testPerson = null;
  for (let i = 0; i < values.length; i++) {
    const fullName = String(values[i][0] || '').trim();
    const address = String(values[i][4] || '').trim();
    if (fullName && address) {
      const addressLines = parseAddress(address);
      Logger.log('\nPerson: ' + fullName);
      Logger.log('  Raw address: "' + address + '"');
      Logger.log('  Parsed line1: "' + addressLines.line1 + '"');
      Logger.log('  Parsed line2: "' + addressLines.line2 + '"');
      Logger.log('  Parsed cityStateZip: "' + addressLines.cityStateZip + '"');

      if (addressLines.line2 && addressLines.line2.trim() !== '') {
        // Found someone with ADDRESS LINE 2
        Logger.log('  ✓ HAS ADDRESS LINE 2 - Using this person for test');
        testPerson = {
          fullName: fullName,
          firstName: extractFirstName(fullName),
          pacName: String(values[i][1] || '').trim(),
          email: String(values[i][2] || '').trim(),
          phone: String(values[i][3] || '').trim(),
          address: address
        };
        testPerson = normalizePersonData(testPerson);
        break;
      }
    }
  }

  if (!testPerson) {
    ui.alert('No person found with ADDRESS LINE 2\n\nCheck View → Logs to see how all addresses were parsed.');
    return;
  }

  Logger.log('\n=== TESTING WITH: ' + testPerson.fullName + ' ===');
  Logger.log('Address: ' + testPerson.address);

  // Create temp doc
  const tempDocFile = DriveApp.getFileById(templateDoc.getId()).makeCopy('DEBUG_TEST_' + testPerson.fullName);
  const tempDoc = DocumentApp.openById(tempDocFile.getId());
  const body = tempDoc.getBody();

  Logger.log('\n=== BEFORE REPLACEMENT ===');
  const beforeParas = body.getParagraphs();
  for (let i = 0; i < beforeParas.length; i++) {
    Logger.log('Para ' + i + ': "' + beforeParas[i].getText() + '"');
  }

  // Do replacement
  replacePlaceholdersInDocument(body, testPerson);
  tempDoc.saveAndClose();

  // Reopen and show result
  const tempDoc2 = DocumentApp.openById(tempDocFile.getId());
  const body2 = tempDoc2.getBody();

  Logger.log('\n=== AFTER REPLACEMENT ===');
  const afterParas = body2.getParagraphs();
  for (let i = 0; i < afterParas.length; i++) {
    Logger.log('Para ' + i + ': "' + afterParas[i].getText() + '"');
  }

  tempDoc2.close();

  ui.alert('Debug complete!\n\n' +
           'Testing: ' + testPerson.fullName + '\n' +
           'Temp doc created: DEBUG_TEST_' + testPerson.fullName + '\n\n' +
           'Check View → Logs for paragraph-by-paragraph output.\n' +
           'Temp doc URL: ' + tempDocFile.getUrl() + '\n\n' +
           'The temp doc was NOT deleted so you can inspect it manually.');
}

/** ========================== PDF GENERATION =================== **/

/**
 * Creates a personalized PDF from template for one person
 */
function createPersonalizedPDF(templateDoc, personData) {
  let tempDocFile = null;

  try {
    // Log person data for debugging
    Logger.log('Creating PDF for: ' + personData.fullName);
    Logger.log('Person data: ' + JSON.stringify(personData));

    // Make a temporary copy of the template
    tempDocFile = DriveApp.getFileById(templateDoc.getId()).makeCopy(personData.fullName + (personData.pacName ? '-' + personData.pacName : ''));
    const tempDoc = DocumentApp.openById(tempDocFile.getId());
    const body = tempDoc.getBody();

    // Replace placeholders and handle empty values
    replacePlaceholdersInDocument(body, personData);
    tempDoc.saveAndClose();

    // Export as PDF
    const pdfBlob = tempDocFile.getAs('application/pdf');

    // Delete the temporary doc
    tempDocFile.setTrashed(true);

    Logger.log('Successfully created PDF for: ' + personData.fullName);
    return pdfBlob;

  } catch (e) {
    Logger.log('Error in createPersonalizedPDF for ' + personData.fullName + ': ' + e.message);
    Logger.log('Stack: ' + e.stack);

    // Clean up temp file if it exists
    if (tempDocFile) {
      try {
        tempDocFile.setTrashed(true);
      } catch (cleanupError) {
        Logger.log('Could not clean up temp file: ' + cleanupError.message);
      }
    }

    // Re-throw the error so it's caught by the caller
    throw e;
  }
}

/**
 * Replaces all placeholders in a document body
 * If placeholder value is empty, removes the entire paragraph
 */
function replacePlaceholdersInDocument(body, personData) {
  // Parse address into lines if provided
  const addressLines = parseAddress(personData.address || '');

  // Current date in multiple formats
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');
  const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy');

  Logger.log('Replacement values for ' + personData.fullName + ':');
  Logger.log('- ADDRESS LINE 1: "' + addressLines.line1 + '"');
  Logger.log('- ADDRESS LINE 2: "' + addressLines.line2 + '"');
  Logger.log('- CITY STATE ZIP: "' + addressLines.cityStateZip + '"');

  // Define all replacements with their values
  const replacements = {
    'FIRST NAME': personData.firstName || '',
    'FIRSTNAME': personData.firstName || '',
    'FULL NAME': personData.fullName || '',
    'FULLNAME': personData.fullName || '',
    'NAME': personData.fullName || '',
    'PAC NAME': personData.pacName || '',
    'PACNAME': personData.pacName || '',
    'PAC NAMES': personData.pacName || '',
    'ORGANIZATION NAME': personData.pacName || '',
    'ORGANIZATION': personData.pacName || '',
    'ORG': personData.pacName || '',
    'ADDRESS LINE 1': addressLines.line1,
    'ADDRESS LINE 2': addressLines.line2,
    'CITY STATE ZIP': addressLines.cityStateZip,
    'CITY, STATE ZIP': addressLines.cityStateZip,
    'CITYSTATEZIP': addressLines.cityStateZip,
    'ADDRESS': personData.address || '',
    'DATE': today,
    'TODAY': today,
    'MONTH DD, YYYY': todayFormatted,
    'Month DD, YYYY': todayFormatted
  };

  // Process each placeholder
  Object.keys(replacements).forEach(key => {
    let value = replacements[key];

    // Ensure value is a string
    if (value === null || value === undefined) {
      value = '';
    }
    value = String(value).trim();

    // Patterns to search for (different bracket types)
    const patterns = [
      '\\[' + key + '\\]',
      '\\[' + key.toLowerCase() + '\\]'
    ];

    patterns.forEach(pattern => {
      let searchResult = body.findText(pattern);

      while (searchResult !== null) {
        const element = searchResult.getElement();
        const para = element.getParent().asParagraph();

        if (value === '') {
          // Empty value - remove the entire paragraph
          try {
            if (body.getParagraphs().length > 1) {
              para.removeFromParent();
              Logger.log('Removed empty placeholder paragraph for: ' + key);
            } else {
              para.clear();
            }
          } catch (e) {
            Logger.log('Could not remove paragraph for ' + key + ': ' + e.message);
          }
          // Don't search for more - we removed the paragraph
          break;
        } else {
          // Has value - replace the placeholder text
          const escapedValue = value.replace(/\$/g, '$$$$');
          body.replaceText(pattern, escapedValue);
          // Continue searching for more occurrences
          searchResult = body.findText(pattern);
        }
      }
    });
  });
}

/** ========================== PDF OPERATIONS ================== **/

/**
 * Sanitizes filename by removing invalid characters
 */
function sanitizeFileName(name) {
  return name.replace(/[^a-zA-Z0-9_\- ]/g, '').trim().substring(0, 100);
}

/** ========================== LABELS GENERATION =============== **/

/**
 * Generates labels PDF and saves to folder
 */
function generateLabelsPDF(people, folderName, folder) {
  let tempDocFile = null;

  try {
    // Generate the labels document (created in root Drive initially)
    const labelsDoc = generateLabelsDocument(people, folderName);

    // Get the temporary doc file
    tempDocFile = DriveApp.getFileById(labelsDoc.getId());

    // Export as PDF blob
    const pdfBlob = tempDocFile.getAs('application/pdf');

    // Create the PDF in the target folder with proper name
    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setName('Mailing Labels.pdf');

    // Delete the temporary Google Doc from root
    tempDocFile.setTrashed(true);

    Logger.log('Labels PDF created successfully in folder: ' + folder.getName());
    return true;
  } catch (e) {
    Logger.log('Error generating labels PDF: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);

    // Try to clean up temp doc even if there was an error
    if (tempDocFile) {
      try {
        tempDocFile.setTrashed(true);
        Logger.log('Cleaned up temporary doc after error');
      } catch (cleanupError) {
        Logger.log('Could not clean up temp doc: ' + cleanupError.message);
      }
    }

    return false;
  }
}

/**
 * Generates a Google Doc with mailing labels in Avery 5160 format
 */
function generateLabelsDocument(people, folderName) {
  const doc = DocumentApp.create(`Temp Labels - ${folderName}`);
  const body = doc.getBody();

  // Set up document margins for Avery 5160 (standard letter size)
  body.setMarginTop(36);      // 0.5 inches
  body.setMarginBottom(36);   // 0.5 inches
  body.setMarginLeft(13);     // ~0.18 inches (Avery 5160 spec)
  body.setMarginRight(13);    // ~0.18 inches

  // Create table with 3 columns for labels
  const numRows = Math.ceil(people.length / LABELS_PER_ROW);
  const table = body.appendTable();

  let personIndex = 0;
  for (let row = 0; row < numRows; row++) {
    const tableRow = table.appendTableRow();

    // Set row height to 1 inch (72 points) for Avery 5160
    tableRow.setMinimumHeight(LABEL_HEIGHT * 72);

    for (let col = 0; col < LABELS_PER_ROW; col++) {
      if (personIndex >= people.length) {
        // Empty cell for remaining slots
        const cell = tableRow.appendTableCell('');
        formatLabelCell(cell);
      } else {
        const person = people[personIndex];
        const labelText = formatLabelText(person);
        const cell = tableRow.appendTableCell(labelText);
        formatLabelCell(cell);
        personIndex++;
      }
    }
  }

  // Format table - no borders for clean label printing
  table.setBorderWidth(0);

  doc.saveAndClose();
  return doc;
}

/**
 * Formats the text for a single label
 */
function formatLabelText(person) {
  let text = person.fullName;
  if (person.pacName) {
    text += '\n' + person.pacName;
  }
  text += '\n' + person.address;
  return text;
}

/**
 * Formats a label cell with proper dimensions and styling for Avery 5160
 */
function formatLabelCell(cell) {
  // Set cell dimensions (convert inches to points: 1 inch = 72 points)
  // Avery 5160: 2.625" x 1" labels
  cell.setWidth(LABEL_WIDTH * 72);  // 2.625 inches = 189 points

  // Reduced padding to maximize usable space on labels
  cell.setPaddingTop(8);     // ~0.11 inches
  cell.setPaddingBottom(8);  // ~0.11 inches
  cell.setPaddingLeft(10);   // ~0.14 inches
  cell.setPaddingRight(10);  // ~0.14 inches

  // Set vertical alignment to top
  cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);

  // Format text in cell
  const text = cell.editAsText();
  text.setFontSize(9);        // Slightly smaller for better fit
  text.setFontFamily('Arial');
  // Note: setLineSpacing() is not available on Text objects, only on Paragraphs
}

/** ========================== COMBINED PDF ==================== **/

/**
 * Generates a combined PDF with all personalized letters
 * Each letter starts on a new page
 * Uses the already-generated individual PDFs to ensure perfect page alignment
 */
function generateCombinedPDF(templateDoc, people, folderName, folder) {
  let tempDocFile = null;

  try {
    Logger.log('Creating combined PDF with ' + people.length + ' letters');

    // Create a new Google Doc for the combined letters
    const combinedDoc = DocumentApp.create('Temp Combined - ' + folderName);
    const combinedBody = combinedDoc.getBody();

    // Remove the default paragraph that Google Docs creates
    const initialParagraphs = combinedBody.getParagraphs();
    if (initialParagraphs.length > 0) {
      initialParagraphs[0].clear();
    }

    // Process each person and append their letter to the combined document
    for (let i = 0; i < people.length; i++) {
      const person = people[i];
      Logger.log('Adding letter ' + (i + 1) + ' of ' + people.length + ' for: ' + person.fullName);

      try {
        // Add page break BEFORE this letter (except for first letter)
        // This ensures each letter starts at the TOP of a new page
        if (i > 0) {
          combinedBody.appendPageBreak();
        }

        // Create a temporary copy of the template for this person
        tempDocFile = DriveApp.getFileById(templateDoc.getId()).makeCopy(person.fullName + (person.pacName ? '-' + person.pacName : ''));
        const tempDoc = DocumentApp.openById(tempDocFile.getId());
        const tempBody = tempDoc.getBody();

        // Replace placeholders (also removes empty ones)
        replacePlaceholdersInDocument(tempBody, person);
        tempDoc.saveAndClose();

        // Reopen to copy content to combined document
        const tempDoc2 = DocumentApp.openById(tempDocFile.getId());
        const finalBody = tempDoc2.getBody();

        // Copy all content from this letter to the combined document
        const numChildren = finalBody.getNumChildren();
        for (let j = 0; j < numChildren; j++) {
          const element = finalBody.getChild(j);
          const elementType = element.getType();

          // Copy the element to the combined document
          if (elementType === DocumentApp.ElementType.PARAGRAPH) {
            const para = element.asParagraph().copy();
            combinedBody.appendParagraph(para);
          } else if (elementType === DocumentApp.ElementType.TABLE) {
            const table = element.asTable().copy();
            combinedBody.appendTable(table);
          } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
            const listItem = element.asListItem().copy();
            combinedBody.appendListItem(listItem);
          } else if (elementType === DocumentApp.ElementType.PAGE_BREAK) {
            // Skip page breaks from the template - we're controlling them ourselves
            continue;
          }
        }

        tempDoc2.close();

        // Clean up temporary file
        tempDocFile.setTrashed(true);
        tempDocFile = null;

      } catch (e) {
        Logger.log('Error adding letter for ' + person.fullName + ': ' + e.message);
        // Clean up temp file if it exists
        if (tempDocFile) {
          try {
            tempDocFile.setTrashed(true);
          } catch (cleanupError) {
            // Ignore cleanup errors
          }
          tempDocFile = null;
        }
        // Continue with next person
      }
    }

    // Save the combined document
    combinedDoc.saveAndClose();

    // Get the file and export as PDF
    const combinedDocFile = DriveApp.getFileById(combinedDoc.getId());
    const pdfBlob = combinedDocFile.getAs('application/pdf');

    // Save PDF to folder
    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setName('Combined Letters.pdf');

    // Delete temporary combined document
    combinedDocFile.setTrashed(true);

    Logger.log('Combined PDF created successfully with ' + people.length + ' letters');
    return true;

  } catch (e) {
    Logger.log('Error generating combined PDF: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);

    // Clean up any remaining temp file
    if (tempDocFile) {
      try {
        tempDocFile.setTrashed(true);
      } catch (cleanupError) {
        Logger.log('Could not clean up temp file: ' + cleanupError.message);
      }
    }

    return false;
  }
}

/** ========================== HELPER MESSAGE ================== **/

/**
 * Information about label printing
 */
function showLabelPrintingHelp() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Label Printing Instructions',
    'The generated labels document is formatted for Avery 5160 labels:\n\n' +
    '• 30 labels per sheet\n' +
    '• 3 columns x 10 rows\n' +
    '• Standard 8.5" x 11" paper\n\n' +
    'To print:\n' +
    '1. Open the labels document from the Drive folder\n' +
    '2. Go to File → Print\n' +
    '3. Load Avery 5160 label sheets in your printer\n' +
    '4. Print normally\n\n' +
    'Each label shows:\n' +
    '• Name\n' +
    '• PAC Names (if provided)\n' +
    '• Address',
    ui.ButtonSet.OK
  );
}


////////////////////////////////////////////////////////////////////////////
// src/Utilities.gs
////////////////////////////////////////////////////////////////////////////

/**
 * Utilities.gs
 * Shared helper functions for name processing, HTML handling, and Drive operations
 */

/** ========================== TEXT NORMALIZATION ============== **/

/**
 * Normalizes text capitalization - converts ALL CAPS to Title Case
 */
function normalizeCapitalization(text) {
  if (!text) return '';

  const str = String(text).trim();

  // Check if text is all uppercase (accounting for spaces and punctuation)
  const lettersOnly = str.replace(/[^a-zA-Z]/g, '');
  if (lettersOnly.length > 0 && lettersOnly === lettersOnly.toUpperCase()) {
    // Text is all caps, convert to title case
    return toTitleCase(str);
  }

  // Text has mixed case, leave it alone
  return str;
}

/**
 * Converts string to Title Case
 */
function toTitleCase(str) {
  // Words that should stay lowercase (unless first word)
  const lowercase = ['a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'in', 'of', 'on', 'or', 'the', 'to', 'via'];

  // US State abbreviations (should stay uppercase)
  const stateAbbreviations = [
    'AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA',
    'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD',
    'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ',
    'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC',
    'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY',
    'DC', 'PR', 'VI', 'GU', 'AS', 'MP'
  ];

  return str.toLowerCase().replace(/\b\w+/g, function(word, index) {
    const upperWord = word.toUpperCase();

    // Keep state abbreviations uppercase
    if (stateAbbreviations.indexOf(upperWord) !== -1) {
      return upperWord;
    }

    // Always capitalize first word
    if (index === 0) {
      return word.charAt(0).toUpperCase() + word.slice(1);
    }

    // Keep certain words lowercase unless they're the first word
    if (lowercase.indexOf(word.toLowerCase()) !== -1) {
      return word.toLowerCase();
    }

    // Capitalize everything else
    return word.charAt(0).toUpperCase() + word.slice(1);
  });
}

/**
 * Normalizes person data - applies proper capitalization
 */
function normalizePersonData(personData) {
  return {
    fullName: normalizeCapitalization(personData.fullName || ''),
    firstName: normalizeCapitalization(personData.firstName || ''),
    pacName: normalizeCapitalization(personData.pacName || ''),
    email: personData.email || '', // Don't normalize emails
    phone: personData.phone || '', // Don't normalize phones
    address: normalizeCapitalization(personData.address || '')
  };
}

/** ========================== PLACEHOLDER REPLACEMENT ========== **/

/**
 * Replaces all placeholders in template with person data
 * Supports: [FIRST NAME], [FULL NAME], [PAC NAME], [ORGANIZATION NAME],
 *           [ADDRESS LINE 1], [ADDRESS LINE 2], [DATE]
 */
function replaceAllPlaceholders(template, personData) {
  if (!template) return '';

  let result = template;

  // Parse address into lines if provided
  const addressLines = parseAddress(personData.address || '');

  // Current date
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');

  // Define all replacements
  const replacements = {
    // Name variations
    'FIRST NAME': personData.firstName || '',
    'FIRSTNAME': personData.firstName || '',
    'FULL NAME': personData.fullName || '',
    'FULLNAME': personData.fullName || '',
    'NAME': personData.fullName || '',

    // Organization/PAC
    'PAC NAME': personData.pacName || '',
    'PACNAME': personData.pacName || '',
    'PAC NAMES': personData.pacName || '',
    'ORGANIZATION NAME': personData.pacName || '',
    'ORGANIZATION': personData.pacName || '',
    'ORG': personData.pacName || '',

    // Address (3-line format)
    'ADDRESS LINE 1': addressLines.line1,
    'ADDRESS LINE 2': addressLines.line2,
    'CITY STATE ZIP': addressLines.cityStateZip,
    'CITY, STATE ZIP': addressLines.cityStateZip,
    'CITYSTATEZIP': addressLines.cityStateZip,
    'ADDRESS': personData.address || '',

    // Date
    'DATE': today,
    'TODAY': today
  };

  // Replace all patterns: [PLACEHOLDER], <PLACEHOLDER>, {{PLACEHOLDER}}
  Object.keys(replacements).forEach(key => {
    const value = replacements[key];
    // [PLACEHOLDER] format (case insensitive)
    result = result.replace(new RegExp('\\[\\s*' + key + '\\s*\\]', 'gi'), value);
    // <PLACEHOLDER> format (case insensitive)
    result = result.replace(new RegExp('<\\s*' + key + '\\s*>', 'gi'), value);
    // {{PLACEHOLDER}} format (case insensitive)
    result = result.replace(new RegExp('\\{\\{\\s*' + key + '\\s*\\}\\}', 'gi'), value);
  });

  return result;
}

/**
 * Parses address into three lines
 * Example: "123 Main St, Apt 4B, Springfield, IL 62701"
 * -> Line 1: "123 Main St"
 * -> Line 2: "Apt 4B"
 * -> City/State/Zip: "Springfield, IL 62701"
 */
function parseAddress(address) {
  // Ensure we always return strings, never undefined
  const emptyResult = { line1: '', line2: '', cityStateZip: '' };

  if (!address) return emptyResult;

  const parts = address.split(',').map(p => p.trim()).filter(Boolean);

  if (parts.length === 0) {
    return { line1: '', line2: '', cityStateZip: '' };
  }

  if (parts.length === 1) {
    // No commas — try to split "123 Main St NW Washington DC 20010" style addresses.
    // Strategy 1: street ends at a known street-type word (optionally followed by a
    //   directional like NW, SE, etc.), then the remaining words are city + state + zip.
    const noCommaPattern = /^(.+?\b(?:rd|st|ave|blvd|dr|ln|ct|pl|way|cir|hwy|pkwy|loop|ter|terr|trl|run|row|path|pass|pike|pt)\b\.?\s*(?:(?:nw|ne|sw|se|north|south|east|west|n|s|e|w)\.?)?\s+)(.+?)\s+\b([A-Z]{2})\b\s+(\d{5}(?:-\d{4})?)\s*$/i;
    const m1 = parts[0].match(noCommaPattern);
    if (m1) {
      return {
        line1: m1[1].trim(),
        line2: '',
        cityStateZip: m1[2].trim() + ', ' + m1[3].toUpperCase() + ' ' + m1[4]
      };
    }

    // Strategy 2: no recognized street type — at least find STATE ZIP at the end
    //   and treat the last word before it as the city.
    const stateZipTail = parts[0].match(/\b([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\s*$/i);
    if (stateZipTail) {
      const before = parts[0].slice(0, parts[0].lastIndexOf(stateZipTail[0])).trim();
      const words = before.split(/\s+/);
      if (words.length >= 2) {
        return {
          line1: words.slice(0, -1).join(' '),
          line2: '',
          cityStateZip: words[words.length - 1] + ', ' + stateZipTail[1].toUpperCase() + ' ' + stateZipTail[2]
        };
      }
    }

    // Fallback: just one part with no recognizable pattern
    return { line1: parts[0], line2: '', cityStateZip: '' };
  }

  if (parts.length === 2) {
    // Two parts: Could be "Street City, State Zip" or "Street, City State Zip"
    // Check if second part is just "State Zip" (e.g., "KY 40513")
    const stateZipPattern = /^([A-Z]{2})\s+(\d{5}(-\d{4})?)$/i;
    const match = parts[1].match(stateZipPattern);

    if (match) {
      // Second part is just "State Zip", need to extract city from first part
      // Example: "1066 Wellington Way Lexington, KY 40513"
      const firstPart = parts[0];
      const words = firstPart.split(/\s+/);

      if (words.length >= 2) {
        // Assume last word(s) of first part is the city
        // Typically: "Street Number Street Name City"
        // Split: take last 1-2 words as city, rest as street
        const city = words[words.length - 1]; // Last word is likely city
        const street = words.slice(0, -1).join(' ');
        const cityStateZip = city + ', ' + parts[1];

        return { line1: street, line2: '', cityStateZip: cityStateZip };
      }
    }

    // Default: "Street, City State Zip"
    return { line1: parts[0], line2: '', cityStateZip: parts[1] };
  }

  // 3+ parts: "Street, Apt/Suite, City, State Zip" or similar
  // Last part is likely city/state/zip
  // Second to last might be part of city name (e.g., "New York, NY")

  // Check if last part looks like just a state + zip (e.g., "NY 10001")
  const lastPart = parts[parts.length - 1];
  const secondLastPart = parts[parts.length - 2];

  // Pattern: State abbreviation (2 letters) + space + 5 digits
  const stateZipPattern = /^[A-Z]{2}\s+\d{5}(-\d{4})?$/i;

  if (stateZipPattern.test(lastPart) && parts.length >= 3) {
    // Check if secondLastPart contains both suite/apt AND city
    // Example: "# 200 LOUISVILLE" or "Suite 263 Lexington"
    // Pattern: suite indicator + number/letter + city name
    const suiteAndCityPattern = /^(#|suite|apt|apartment|unit|ste|no\.?)\s*([a-z0-9-]+)\s+(.+)$/i;
    const suiteMatch = secondLastPart.match(suiteAndCityPattern);

    if (suiteMatch) {
      // Found suite/apt number AND city in same part
      const suiteIndicator = suiteMatch[1];  // "#" or "Suite" etc
      const suiteNumber = suiteMatch[2];      // "200" or "U-7" etc
      const city = suiteMatch[3];             // "LOUISVILLE" or "Lexington"

      const line2 = suiteIndicator + ' ' + suiteNumber;
      const cityStateZip = city + ', ' + lastPart;
      const addressParts = parts.slice(0, -2);

      if (addressParts.length === 1) {
        return { line1: addressParts[0], line2: line2, cityStateZip: cityStateZip };
      } else {
        return {
          line1: addressParts[0],
          line2: addressParts.slice(1).join(', ') + ', ' + line2,
          cityStateZip: cityStateZip
        };
      }
    }

    // Format: "..., City, ST ZIP" (no suite in secondLastPart)
    const cityStateZip = secondLastPart + ', ' + lastPart;
    const addressParts = parts.slice(0, -2);

    if (addressParts.length === 1) {
      return { line1: addressParts[0], line2: '', cityStateZip: cityStateZip };
    } else {
      return {
        line1: addressParts[0],
        line2: addressParts.slice(1).join(', '),
        cityStateZip: cityStateZip
      };
    }
  }

  // Default: last part is city/state/zip, everything before is address
  const cityStateZip = parts[parts.length - 1];
  const addressParts = parts.slice(0, -1);

  if (addressParts.length === 1) {
    return { line1: addressParts[0], line2: '', cityStateZip: cityStateZip };
  } else {
    return {
      line1: addressParts[0],
      line2: addressParts.slice(1).join(', '),
      cityStateZip: cityStateZip
    };
  }
}

/** ========================== NAME PROCESSING (Legacy) ========= **/
// These functions are kept for backward compatibility

/**
 * Extracts first name from full name, handling titles, quotes, and formats
 */
function extractFirstName(fullName) {
  let s = fullName.replace(/["']/g, '').replace(/\(.*?\)/g, ' ').replace(/\s+/g, ' ').trim();

  // Remove titles
  s = s.replace(/^(mr|mrs|ms|miss|mx|dr|prof)\.\s+/i, '');

  // Remove suffixes like Jr., Sr., II, III, IV, etc. (they come after commas)
  const suffixes = /,\s*(jr\.?|sr\.?|ii|iii|iv|v|esq\.?|phd\.?|md\.?)$/i;
  s = s.replace(suffixes, '');

  // Now handle "Last, First" format
  if (s.includes(',')) {
    const parts = s.split(',').map(t => t.trim()).filter(Boolean);
    if (parts.length > 1) s = parts[1]; // Take the part after comma (First name)
  }

  // Return first word of whatever remains
  return (s.split(/\s+/)[0] || '').trim();
}

/**
 * Replaces first name placeholders in template text
 */
function replaceFirstNamePlaceholders(template, firstName) {
  const patterns = [
    /\[\s*first\s*name\s*\]/ig,
    /<\s*first\s*name\s*>/ig,
    /\{\{\s*first\s*name\s*\}\}/ig,
    /\{\{\s*FirstName\s*\}\}/g
  ];
  let result = template;
  patterns.forEach(p => { result = result.replace(p, firstName); });
  return result;
}

/**
 * Fills first name in subject line
 */
function fillFirstNameInSubject(template, firstName) {
  if (!template) return '';
  return replaceFirstNamePlaceholders(template, firstName);
}

/**
 * Fills first name in plain text body
 */
function fillFirstNameInBody(templatePlainText, firstName) {
  if (!templatePlainText) return '';
  const replaced = replaceFirstNamePlaceholders(templatePlainText, firstName);
  if (replaced !== templatePlainText) return replaced;
  return `Hi ${firstName},\n\n` + templatePlainText;
}

/**
 * Fills first name in HTML body
 */
function fillFirstNameInHtml(templateHtml, firstName) {
  if (!templateHtml) return '';
  const replaced = replaceFirstNamePlaceholders(templateHtml, firstName);
  if (replaced !== templateHtml) return ensureHtmlContainer(replaced);
  // No placeholder found. Prepend a greeting paragraph.
  const greeting = `<p>Hi ${escapeHtml(firstName)},</p>`;
  return ensureHtmlContainer(greeting + templateHtml);
}

/** ========================== HTML PROCESSING ================== **/

/**
 * Build HTML body from template (after placeholders have been replaced)
 */
function buildHtmlBodyFromTemplate(templateWithReplacements, signatureHtml) {
  const looksHtml = isHtml(templateWithReplacements);
  let bodyHtml = looksHtml
    ? ensureHtmlContainer(templateWithReplacements)
    : textToHtml(asPlainText(templateWithReplacements));
  if (signatureHtml) bodyHtml += appendSignature(signatureHtml);
  return bodyHtml;
}

/**
 * Build HTML body, preserving lists and other tags if template is HTML (LEGACY)
 * @deprecated Use buildHtmlBodyFromTemplate with replaceAllPlaceholders instead
 */
function buildHtmlBody(templateFromA2, firstName, signatureHtml) {
  const looksHtml = isHtml(templateFromA2);
  let bodyHtml = looksHtml
    ? fillFirstNameInHtml(templateFromA2, firstName)
    : textToHtml(fillFirstNameInBody(asPlainText(templateFromA2), firstName));
  if (signatureHtml) bodyHtml += appendSignature(signatureHtml);
  return bodyHtml;
}

/**
 * Heuristics to detect if string contains HTML
 */
function isHtml(s) {
  if (!s) return false;
  const str = String(s).trim();
  if (str.indexOf('<') === -1 || str.indexOf('>') === -1) return false;
  // Require at least one common HTML tag to reduce false positives
  return /<\s*(p|div|br|ul|ol|li|a|strong|em|span|table|tbody|tr|td|h[1-6])\b/i.test(str);
}

/**
 * Ensures HTML is wrapped in a container element
 */
function ensureHtmlContainer(html) {
  if (/<\s*html\b|<\s*body\b/i.test(html)) return html;
  return `<div>${html}</div>`;
}

/**
 * Converts plain text to HTML, preserving structure
 */
function asPlainText(s) {
  // If someone pasted HTML, strip tags before treating as plain text
  return stripHtml(String(s || ''));
}

/**
 * Converts text to HTML with paragraph and link formatting
 */
function textToHtml(txt) {
  // Keep links and paragraphs, but do not invent bullets
  let html = escapeHtml(txt);
  html = html.replace(/(https?:\/\/[^\s]+)/g, '<a href="$1">$1</a>');
  html = html.replace(/\n{2,}/g, '</p><p>');
  html = '<p>' + html.replace(/\n/g, '<br>') + '</p>';
  return html;
}

/**
 * Escapes HTML special characters
 */
function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * Appends signature HTML to body
 */
function appendSignature(signatureHtml) {
  if (!signatureHtml) return '';
  return '<br><br>' + signatureHtml;
}

/**
 * Strips HTML tags from string
 */
function stripHtml(html) {
  return String(html).replace(/<[^>]*>/g, '').replace(/\s+\n/g, '\n').trim();
}

/**
 * Gets default Gmail signature as HTML
 */
function getDefaultSignatureHtml() {
  try {
    const res = Gmail.Users.Settings.SendAs.list('me');
    if (!res || !res.sendAs || !res.sendAs.length) return '';
    const primary = res.sendAs.find(s => s.isDefault) || res.sendAs[0];
    return primary.signature || '';
  } catch (e) {
    return '';
  }
}

/** ========================== GOOGLE DRIVE ===================== **/

/**
 * Gets a Drive file from URL or file ID
 */
function fileFromDriveLink(input) {
  const id = extractDriveId(input);
  if (!id) return null;
  try {
    return DriveApp.getFileById(id);
  } catch (e) {
    return null;
  }
}

/**
 * Extracts Google Drive file ID from URL or returns ID if already extracted
 */
function extractDriveId(s) {
  if (!s) return '';
  s = String(s).trim();

  if (/^[a-zA-Z0-9_-]{20,}$/.test(s) && s.indexOf('http') !== 0) return s;

  let m = s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];

  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];

  m = s.match(/\/uc\?[^#]*id=([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];

  m = s.match(/drive\.google\.com\/(?:file\/d\/|drive\/folders\/)?([a-zA-Z0-9_-]{20,})(?:\/|$)/);
  if (m && m[1]) return m[1];

  return '';
}

/** ========================== NAME NORMALIZATION =============== **/

/**
 * Normalizes name for comparison (used by email finders)
 */
function normalizeName(s) {
  return String(s || '').toLowerCase()
    .replace(/["']/g, '')
    .replace(/\(.*?\)/g, ' ')
    .replace(/\b(mr|mrs|ms|miss|mx|dr|prof)\.?\b/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Cleans name by removing titles and special characters
 */
function cleanName(s) {
  let t = String(s || '').replace(/["']/g, '').replace(/\(.*?\)/g, ' ');
  t = t.replace(/\b(mr|mrs|ms|miss|mx|dr|prof)\.?\b/i, '');
  return t.replace(/\s+/g, ' ').trim();
}
