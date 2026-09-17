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
