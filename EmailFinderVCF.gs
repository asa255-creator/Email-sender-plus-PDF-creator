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
