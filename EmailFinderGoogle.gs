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
