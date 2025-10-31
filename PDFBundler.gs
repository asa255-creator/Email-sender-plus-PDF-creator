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
    ui.alert('Error: No template document specified in email details C2.\n\n' +
             'Please add a Google Docs URL or file ID.\n' +
             'The document should contain placeholders like [FIRST NAME], [FULL NAME], [PAC NAME], etc.');
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

  let templateDoc;
  try {
    templateDoc = DocumentApp.openById(fileId);
  } catch (e) {
    ui.alert('Error: Cannot open document with ID: ' + fileId + '\n\n' +
             'Error message: ' + e.message + '\n\n' +
             'Make sure:\n' +
             '1. The file is a Google Doc (not PDF)\n' +
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

  // Generate personalized PDFs
  let generatedCount = 0;
  people.forEach(person => {
    try {
      // Create personalized PDF for this person
      const pdfBlob = createPersonalizedPDF(templateDoc, person);
      const fileName = `${sanitizeFileName(person.fullName)}.pdf`;

      // Save PDF to folder
      folder.createFile(pdfBlob.setName(fileName));
      generatedCount++;
    } catch (e) {
      Logger.log('Error generating PDF for ' + person.fullName + ': ' + e.message);
    }
  });

  // Generate labels PDF
  const labelsPdfCreated = generateLabelsPDF(people, folderName, folder);

  // Show completion message
  if (labelsPdfCreated) {
    ui.alert(
      'PDF Bundle Created Successfully!\n\n' +
      '✓ PDFs generated: ' + generatedCount + '\n' +
      '✓ Labels PDF: Created\n' +
      '✓ Total labels: ' + people.length + '\n\n' +
      'Folder Name: ' + folderName + '\n' +
      'Folder Location: ' + folder.getUrl() + '\n\n' +
      'Files in folder:\n' +
      '• ' + generatedCount + ' personalized PDFs (one per person)\n' +
      '• Mailing Labels.pdf (print on Avery 5160 sheets)\n\n' +
      'Click the folder URL above to open it.'
    );
  } else {
    ui.alert(
      'PDF Bundle Partially Created\n\n' +
      'PDFs generated: ' + generatedCount + '\n' +
      'Labels PDF: FAILED\n\n' +
      'Folder: ' + folderName + '\n' +
      'Location: ' + folder.getUrl() + '\n\n' +
      'Check View → Logs for error details.'
    );
  }
}

/** ========================== PDF GENERATION =================== **/

/**
 * Creates a personalized PDF from template for one person
 */
function createPersonalizedPDF(templateDoc, personData) {
  // Make a temporary copy of the template
  const tempDocFile = DriveApp.getFileById(templateDoc.getId()).makeCopy('temp_' + personData.fullName);
  const tempDoc = DocumentApp.openById(tempDocFile.getId());
  const body = tempDoc.getBody();

  // Replace placeholders using replaceText (preserves formatting)
  replacePlaceholdersInDocument(body, personData);

  // Save and close
  tempDoc.saveAndClose();

  // Export as PDF
  const pdfBlob = tempDocFile.getAs('application/pdf');

  // Delete the temporary doc
  tempDocFile.setTrashed(true);

  return pdfBlob;
}

/**
 * Replaces all placeholders in a document body
 */
function replacePlaceholdersInDocument(body, personData) {
  // Parse address into lines if provided
  const addressLines = parseAddress(personData.address || '');

  // Current date
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');

  // Define all replacements
  const replacements = {
    'FIRST NAME': personData.firstName || '',
    'FIRSTNAME': personData.firstName || '',
    'FULL NAME': personData.fullName || '',
    'FULLNAME': personData.fullName || '',
    'NAME': personData.fullName || '',
    'PAC NAME': personData.pacName || '',
    'PACNAME': personData.pacName || '',
    'PAC NAMES': personData.pacName || '',
    'ORGANIZATION NAME': personData.pacName || '',
    'ORGANIZATION': personData.pacName || '',
    'ADDRESS LINE 1': addressLines.line1,
    'ADDRESS LINE 2': addressLines.line2,
    'ADDRESS': personData.address || '',
    'DATE': today,
    'TODAY': today
  };

  // Replace all patterns: [PLACEHOLDER], <PLACEHOLDER>, {{PLACEHOLDER}}
  Object.keys(replacements).forEach(key => {
    const value = replacements[key];
    // [PLACEHOLDER] format (case insensitive)
    body.replaceText('\\[\\s*' + key + '\\s*\\]', value);
    body.replaceText('\\[\\s*' + key.toLowerCase() + '\\s*\\]', value);
    // <PLACEHOLDER> format (case insensitive)
    body.replaceText('<\\s*' + key + '\\s*>', value);
    body.replaceText('<\\s*' + key.toLowerCase() + '\\s*>', value);
    // {{PLACEHOLDER}} format (case insensitive)
    body.replaceText('\\{\\{\\s*' + key + '\\s*\\}\\}', value);
    body.replaceText('\\{\\{\\s*' + key.toLowerCase() + '\\s*\\}\\}', value);
  });
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
 * Generates labels PDF and saves to folder
 */
function generateLabelsPDF(people, folderName, folder) {
  let tempDocFile = null;

  try {
    // Generate the labels document (created in root Drive initially)
    const labelsDoc = generateLabelsDocument(people, folderName);

    // Get the temporary doc file
    tempDocFile = DriveApp.getFileById(labelsDoc.getId());

    // Export as PDF blob
    const pdfBlob = tempDocFile.getAs('application/pdf');

    // Create the PDF in the target folder with proper name
    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setName('Mailing Labels.pdf');

    // Delete the temporary Google Doc from root
    tempDocFile.setTrashed(true);

    Logger.log('Labels PDF created successfully in folder: ' + folder.getName());
    return true;
  } catch (e) {
    Logger.log('Error generating labels PDF: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);

    // Try to clean up temp doc even if there was an error
    if (tempDocFile) {
      try {
        tempDocFile.setTrashed(true);
        Logger.log('Cleaned up temporary doc after error');
      } catch (cleanupError) {
        Logger.log('Could not clean up temp doc: ' + cleanupError.message);
      }
    }

    return false;
  }
}

/**
 * Generates a Google Doc with mailing labels in Avery 5160 format
 */
function generateLabelsDocument(people, folderName) {
  const doc = DocumentApp.create(`Temp Labels - ${folderName}`);
  const body = doc.getBody();

  // Set up document margins for Avery 5160 (standard letter size)
  body.setMarginTop(36);      // 0.5 inches
  body.setMarginBottom(36);   // 0.5 inches
  body.setMarginLeft(13);     // ~0.18 inches (Avery 5160 spec)
  body.setMarginRight(13);    // ~0.18 inches

  // Create table with 3 columns for labels
  const numRows = Math.ceil(people.length / LABELS_PER_ROW);
  const table = body.appendTable();

  let personIndex = 0;
  for (let row = 0; row < numRows; row++) {
    const tableRow = table.appendTableRow();

    // Set row height to 1 inch (72 points) for Avery 5160
    tableRow.setMinimumHeight(LABEL_HEIGHT * 72);

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

  // Format table - no borders for clean label printing
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
 * Formats a label cell with proper dimensions and styling for Avery 5160
 */
function formatLabelCell(cell) {
  // Set cell dimensions (convert inches to points: 1 inch = 72 points)
  // Avery 5160: 2.625" x 1" labels
  cell.setWidth(LABEL_WIDTH * 72);  // 2.625 inches = 189 points

  // Reduced padding to maximize usable space on labels
  cell.setPaddingTop(8);     // ~0.11 inches
  cell.setPaddingBottom(8);  // ~0.11 inches
  cell.setPaddingLeft(10);   // ~0.14 inches
  cell.setPaddingRight(10);  // ~0.14 inches

  // Set vertical alignment to top
  cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);

  // Format text in cell
  const text = cell.editAsText();
  text.setFontSize(9);        // Slightly smaller for better fit
  text.setFontFamily('Arial');
  // Note: setLineSpacing() is not available on Text objects, only on Paragraphs
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
