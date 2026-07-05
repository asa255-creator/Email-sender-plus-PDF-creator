/**
 * PlainEmailSender.gs
 * A simple, standalone "second system" for sending name-only emails.
 *
 * Unlike EmailSender.gs, this flow:
 *   - Uses ONLY the person's Name (column A) and Email (column C)
 *   - NEVER inserts an organization/PAC name, phone, or address
 *     (even if that data exists in the sheet, it is ignored)
 *   - NEVER attaches a PDF or any other file
 *
 * It still uses the Body/Subject templates from the "email details" sheet
 * (A2 / B2), personalizes by first/full name, and appends your Gmail
 * signature. Templates are read from the same place as the other tools,
 * so you can reuse them.
 */

/** ============== CREATE PLAIN DRAFTS (name only) ============== **/
function createPlainDraftsFromList() {
  runPlainEmail_(/* sendNow = */ false);
}

/** ================ SEND PLAIN EMAILS (name only) ============== **/
function sendPlainEmailsFromList() {
  runPlainEmail_(/* sendNow = */ true);
}

/**
 * Shared worker for the plain (name-only) draft/send flows.
 * @param {boolean} sendNow  true = send immediately, false = create drafts
 */
function runPlainEmail_(sendNow) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(LIST_SHEET) || ss.getActiveSheet();
  const detailsSh = ss.getSheetByName(DETAILS_SHEET);
  if (!detailsSh) throw new Error('Sheet "email details" not found.');

  // Use getValue so we keep raw HTML if present
  const bodyTemplate = String(detailsSh.getRange('A2').getValue() || '');
  const subjectTemplate = String(detailsSh.getRange('B2').getValue() || '');
  if (!bodyTemplate) throw new Error('Body template missing in email details A2.');
  if (!subjectTemplate) throw new Error('Subject missing in email details B2.');

  // CC addresses from column D (D2, D3, ...) are still supported
  const ccAddresses = getCCAddresses(detailsSh);

  const lastRow = listSh.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data rows found.');
    return;
  }

  // Only read the Name and Email columns — everything else is intentionally ignored
  const width = Math.max(NAME_COL, EMAIL_COL);
  const values = listSh.getRange(2, 1, lastRow - 1, width).getDisplayValues();

  const signatureHtml = getDefaultSignatureHtml(); // may be ''

  let count = 0;
  values.forEach(row => {
    const fullName = String(row[NAME_COL - 1] || '').trim() || 'To Whom It May Concern';
    const email = String(row[EMAIL_COL - 1] || '').trim();

    if (!email) return; // Only skip if email is missing

    const firstName = fullName === 'To Whom It May Concern' ? fullName : extractFirstName(fullName);

    // Name-only person data: organization/PAC, phone, and address are always
    // blank so they can never leak into a plain email.
    let personData = {
      fullName: fullName,
      firstName: firstName,
      pacName: '',
      email: email,
      phone: '',
      address: ''
    };

    // Normalize capitalization (ALL CAPS -> Title Case)
    personData = normalizePersonData(personData);

    // Replace placeholders in subject and body (name placeholders only will
    // resolve to values; org/address/phone placeholders resolve to empty)
    const subject = replaceAllPlaceholders(subjectTemplate, personData);
    const bodyWithPlaceholders = replaceAllPlaceholders(bodyTemplate, personData);

    if (USE_HTML) {
      const bodyHtml = buildHtmlBodyFromTemplate(bodyWithPlaceholders, signatureHtml);
      const options = { htmlBody: bodyHtml };
      if (ccAddresses) options.cc = ccAddresses;
      if (sendNow) {
        GmailApp.sendEmail(email, subject, stripHtml(bodyHtml) || ' ', options);
      } else {
        GmailApp.createDraft(email, subject, '', options);
      }
    } else {
      const bodyText = asPlainText(bodyWithPlaceholders);
      const bodyWithSig = bodyText + (signatureHtml ? '\n\n' + stripHtml(signatureHtml) : '');
      const options = {};
      if (ccAddresses) options.cc = ccAddresses;
      if (sendNow) {
        GmailApp.sendEmail(email, subject, bodyWithSig, options);
      } else {
        GmailApp.createDraft(email, subject, bodyWithSig, options);
      }
    }

    count++;
  });

  const verb = sendNow ? 'Plain emails sent: ' : 'Plain drafts created: ';
  SpreadsheetApp.getUi().alert(verb + count + '\n(name only — no attachment, no organization name)' +
    (ccAddresses ? '\nCC: ' + ccAddresses : ''));
}
