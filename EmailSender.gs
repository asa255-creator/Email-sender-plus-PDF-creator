/**
 * EmailSender.gs
 * Functions for creating Gmail drafts and sending emails with attachments
 */

/** ========================== CONFIG ========================== **/
const LIST_SHEET = 'People';           // names in A, emails in C
const DETAILS_SHEET = 'email details'; // A2 = body template, B2 = subject, C2 = Drive URL or ID for PDF
const NAME_COL = 1;                    // A
const EMAIL_COL = 3;                   // C
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

  const width = Math.max(NAME_COL, EMAIL_COL);
  const values = listSh.getRange(2, 1, lastRow - 1, width).getDisplayValues();

  const signatureHtml = getDefaultSignatureHtml(); // may be ''

  let created = 0;
  values.forEach(row => {
    const fullName = String(row[NAME_COL - 1] || '').trim();
    const email = String(row[EMAIL_COL - 1] || '').trim();
    if (!fullName || !email) return;

    const firstName = extractFirstName(fullName);
    const subject = fillFirstNameInSubject(subjectTemplate, firstName);

    if (USE_HTML) {
      const bodyHtml = buildHtmlBody(bodyTemplate, firstName, signatureHtml);
      GmailApp.createDraft(email, subject, '', { htmlBody: bodyHtml });
    } else {
      const bodyText = fillFirstNameInBody(asPlainText(bodyTemplate), firstName);
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

  const width = Math.max(NAME_COL, EMAIL_COL);
  const values = listSh.getRange(2, 1, lastRow - 1, width).getDisplayValues();

  const signatureHtml = getDefaultSignatureHtml(); // may be ''

  let sent = 0;
  values.forEach(row => {
    const fullName = String(row[NAME_COL - 1] || '').trim();
    const email = String(row[EMAIL_COL - 1] || '').trim();
    if (!fullName || !email) return;

    const firstName = extractFirstName(fullName);
    const subject = fillFirstNameInSubject(subjectTemplate, firstName);

    if (USE_HTML) {
      const bodyHtml = buildHtmlBody(bodyTemplate, firstName, signatureHtml);
      const options = file
        ? { htmlBody: bodyHtml, attachments: [file.getAs(MimeType.PDF)] }
        : { htmlBody: bodyHtml };
      GmailApp.sendEmail(email, subject, stripHtml(bodyHtml) || ' ', options);
    } else {
      const bodyText = fillFirstNameInBody(asPlainText(bodyTemplate), firstName);
      const bodyWithSig = bodyText + (signatureHtml ? '\n\n' + stripHtml(signatureHtml) : '');
      const options = file ? { attachments: [file.getAs(MimeType.PDF)] } : {};
      GmailApp.sendEmail(email, subject, bodyWithSig, options);
    }

    sent++;
  });

  SpreadsheetApp.getUi().alert('Emails sent: ' + sent + (file ? ' (with attachment)' : ' (no attachment)'));
}
