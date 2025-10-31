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
    const fullName = String(row[NAME_COL - 1] || '').trim();
    const pacName = String(row[PAC_COL - 1] || '').trim();
    const email = String(row[EMAIL_COL - 1] || '').trim();
    const phone = String(row[PHONE_COL - 1] || '').trim();
    const address = String(row[ADDRESS_COL - 1] || '').trim();

    if (!fullName || !email) return;

    const firstName = extractFirstName(fullName);

    // Build person data object for placeholder replacement
    const personData = {
      fullName: fullName,
      firstName: firstName,
      pacName: pacName,
      email: email,
      phone: phone,
      address: address
    };

    // Replace placeholders in subject and body
    const subject = replaceAllPlaceholders(subjectTemplate, personData);
    const bodyWithPlaceholders = replaceAllPlaceholders(bodyTemplate, personData);

    if (USE_HTML) {
      const bodyHtml = buildHtmlBodyFromTemplate(bodyWithPlaceholders, signatureHtml);
      GmailApp.createDraft(email, subject, '', { htmlBody: bodyHtml });
    } else {
      const bodyText = asPlainText(bodyWithPlaceholders);
      const bodyWithSig = bodyText + (signatureHtml ? '\n\n' + stripHtml(signatureHtml) : '');
      GmailApp.createDraft(email, subject, bodyWithSig);
    }

    created++;
  });

  SpreadsheetApp.getUi().alert('Drafts created: ' + created);
}

/** ===================== SEND WITH ATTACHMENT ================= **/
function sendEmailsFromListWithAttachment() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(LIST_SHEET) || ss.getActiveSheet();
  const detailsSh = ss.getSheetByName(DETAILS_SHEET);
  if (!detailsSh) throw new Error('Sheet "email details" not found.');

  const bodyTemplate = String(detailsSh.getRange('A2').getValue() || '');
  const subjectTemplate = String(detailsSh.getRange('B2').getValue() || '');
  const attachmentRef = String(detailsSh.getRange('C2').getValue() || ''); // Drive URL or file ID
  if (!bodyTemplate) throw new Error('Body template missing in email details A2.');
  if (!subjectTemplate) throw new Error('Subject missing in email details B2.');

  let file = null;
  if (attachmentRef) {
    file = fileFromDriveLink(attachmentRef);
    if (!file) throw new Error('Could not open the file from C2. Check that the link or ID is correct and you have access.');
  }

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
    const fullName = String(row[NAME_COL - 1] || '').trim();
    const pacName = String(row[PAC_COL - 1] || '').trim();
    const email = String(row[EMAIL_COL - 1] || '').trim();
    const phone = String(row[PHONE_COL - 1] || '').trim();
    const address = String(row[ADDRESS_COL - 1] || '').trim();

    if (!fullName || !email) return;

    const firstName = extractFirstName(fullName);

    // Build person data object for placeholder replacement
    const personData = {
      fullName: fullName,
      firstName: firstName,
      pacName: pacName,
      email: email,
      phone: phone,
      address: address
    };

    // Replace placeholders in subject and body
    const subject = replaceAllPlaceholders(subjectTemplate, personData);
    const bodyWithPlaceholders = replaceAllPlaceholders(bodyTemplate, personData);

    if (USE_HTML) {
      const bodyHtml = buildHtmlBodyFromTemplate(bodyWithPlaceholders, signatureHtml);
      const options = file
        ? { htmlBody: bodyHtml, attachments: [file.getAs(MimeType.PDF)] }
        : { htmlBody: bodyHtml };
      GmailApp.sendEmail(email, subject, stripHtml(bodyHtml) || ' ', options);
    } else {
      const bodyText = asPlainText(bodyWithPlaceholders);
      const bodyWithSig = bodyText + (signatureHtml ? '\n\n' + stripHtml(signatureHtml) : '');
      const options = file ? { attachments: [file.getAs(MimeType.PDF)] } : {};
      GmailApp.sendEmail(email, subject, bodyWithSig, options);
    }

    sent++;
  });

  SpreadsheetApp.getUi().alert('Emails sent: ' + sent + (file ? ' (with attachment)' : ' (no attachment)'));
}
