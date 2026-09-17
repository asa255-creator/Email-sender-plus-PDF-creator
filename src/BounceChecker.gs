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
