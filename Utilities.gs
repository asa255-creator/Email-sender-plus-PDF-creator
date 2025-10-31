/**
 * Utilities.gs
 * Shared helper functions for name processing, HTML handling, and Drive operations
 */

/** ========================== NAME PROCESSING ================== **/

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
 * Build HTML body, preserving lists and other tags if template is HTML
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
