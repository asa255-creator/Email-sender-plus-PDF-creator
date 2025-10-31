/**
 * Utilities.gs
 * Shared helper functions for name processing, HTML handling, and Drive operations
 */

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

    // Address
    'ADDRESS LINE 1': addressLines.line1,
    'ADDRESS LINE 2': addressLines.line2,
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
 * Parses address into two lines
 * Example: "123 Main St, Apt 4B, City, ST 12345" -> Line 1: "123 Main St, Apt 4B", Line 2: "City, ST 12345"
 */
function parseAddress(address) {
  if (!address) return { line1: '', line2: '' };

  const parts = address.split(',').map(p => p.trim()).filter(Boolean);

  if (parts.length <= 1) {
    return { line1: address, line2: '' };
  }

  if (parts.length === 2) {
    return { line1: parts[0], line2: parts[1] };
  }

  // 3+ parts: assume last part is "City, ST ZIP", everything before is address
  const line2 = parts.slice(-2).join(', '); // Last two parts (City, State ZIP)
  const line1 = parts.slice(0, -2).join(', '); // Everything before

  return { line1, line2 };
}

/** ========================== NAME PROCESSING (Legacy) ========= **/
// These functions are kept for backward compatibility

/**
 * Extracts first name from full name, handling titles, quotes, and formats
 */
function extractFirstName(fullName) {
  let s = fullName.replace(/["']/g, '').replace(/\(.*?\)/g, ' ').replace(/\s+/g, ' ').trim();
  s = s.replace(/^(mr|mrs|ms|miss|mx|dr|prof)\.\s+/i, '');
  if (s.includes(',')) {
    const parts = s.split(',').map(t => t.trim()).filter(Boolean);
    if (parts.length > 1) s = parts[1];
  }
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
