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
    ui.alert('Error: No PDF file specified in email details C2. Please add a Google Drive URL or file ID.');
    return;
  }

  // Debug: Extract and show the file ID
  const fileId = extractDriveId(attachmentRef);
  if (!fileId) {
    ui.alert('Error: Could not extract file ID from C2.\n\nValue in C2: ' + attachmentRef + '\n\nPlease use one of these formats:\n' +
             '• Full Drive URL: https://drive.google.com/file/d/FILE_ID/view\n' +
             '• Just the file ID: 1a2b3c4d5e6f7g8h9i0j');
    return;
  }

  let templateFile;
  try {
    templateFile = DriveApp.getFileById(fileId);
  } catch (e) {
    ui.alert('Error: Cannot access file with ID: ' + fileId + '\n\n' +
             'Error message: ' + e.message + '\n\n' +
             'Make sure:\n' +
             '1. The file exists in your Drive\n' +
             '2. You have access to the file\n' +
             '3. The file ID is correct');
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

  // Filter valid rows (must have name and address)
  const people = [];
  values.forEach(row => {
    const fullName = String(row[BUNDLE_NAME_COL - 1] || '').trim();
    const pacName = String(row[BUNDLE_PAC_COL - 1] || '').trim();
    const email = String(row[BUNDLE_EMAIL_COL - 1] || '').trim();
    const phone = String(row[BUNDLE_PHONE_COL - 1] || '').trim();
    const address = String(row[BUNDLE_ADDRESS_COL - 1] || '').trim();

    if (!fullName || !address) return;

    const firstName = extractFirstName(fullName);

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

  // Create folder with timestamp
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmmss');
  const folderName = `PDF Bundle ${timestamp}`;
  const folder = DriveApp.createFolder(folderName);

  // Copy PDFs to folder
  let copiedCount = 0;
  people.forEach(person => {
    try {
      const fileName = `${sanitizeFileName(person.fullName)}.pdf`;
      const copiedFile = templateFile.makeCopy(fileName, folder);
      copiedCount++;
    } catch (e) {
      Logger.log('Error copying PDF for ' + person.fullName + ': ' + e.message);
    }
  });

  // Generate labels document
  const labelsDoc = generateLabelsDocument(people, folderName);

  // Move labels doc to the same folder
  if (labelsDoc) {
    const labelsFile = DriveApp.getFileById(labelsDoc.getId());
    labelsFile.moveTo(folder);
  }

  // Show completion message
  ui.alert(
    'PDF Bundle Created!\n\n' +
    'PDFs copied: ' + copiedCount + '\n' +
    'Labels created: ' + people.length + '\n\n' +
    'Folder: ' + folderName + '\n' +
    'Location: ' + folder.getUrl()
  );
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
 * Generates a Google Doc with mailing labels in Avery 5160 format
 */
function generateLabelsDocument(people, folderName) {
  const doc = DocumentApp.create(`Mailing Labels - ${folderName}`);
  const body = doc.getBody();

  // Set up document
  body.setMarginTop(36);      // 0.5 inches
  body.setMarginBottom(36);   // 0.5 inches
  body.setMarginLeft(14);     // ~0.2 inches
  body.setMarginRight(14);    // ~0.2 inches

  // Create table with 3 columns for labels
  const numRows = Math.ceil(people.length / LABELS_PER_ROW);
  const table = body.appendTable();

  let personIndex = 0;
  for (let row = 0; row < numRows; row++) {
    const tableRow = table.appendTableRow();

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

  // Format table
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
 * Formats a label cell with proper dimensions and styling
 */
function formatLabelCell(cell) {
  // Set cell dimensions (convert inches to points: 1 inch = 72 points)
  cell.setWidth(LABEL_WIDTH * 72);
  cell.setPaddingTop(LABEL_MARGIN * 72);
  cell.setPaddingBottom(LABEL_MARGIN * 72);
  cell.setPaddingLeft(LABEL_MARGIN * 72);
  cell.setPaddingRight(LABEL_MARGIN * 72);

  // Set vertical alignment to top
  cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);

  // Format text in cell
  const text = cell.editAsText();
  text.setFontSize(10);
  text.setFontFamily('Arial');
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
