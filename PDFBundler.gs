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

  // Show what's actually in C2 (raw value)
  ui.alert('Step 1: Checking Template File\n\n' +
           'Value in "email details" sheet, cell C2:\n' +
           attachmentRef.substring(0, 100) + (attachmentRef.length > 100 ? '...' : '') + '\n\n' +
           'Extracting file ID...');

  Logger.log('Template reference from email details C2: ' + attachmentRef);

  // Extract the file ID
  const fileId = extractDriveId(attachmentRef);
  if (!fileId) {
    ui.alert('Error: Could not extract file ID from C2.\n\n' +
             'Value in C2: ' + attachmentRef + '\n\n' +
             'Please use one of these formats:\n' +
             '• Full Drive URL: https://docs.google.com/document/d/FILE_ID/edit\n' +
             '• Just the file ID: 1a2b3c4d5e6f7g8h9i0j');
    return;
  }

  Logger.log('Extracted file ID: ' + fileId);

  // Get the actual file name from Drive FIRST (before opening)
  let templateFileName = '';
  let fileType = '';
  try {
    const driveFile = DriveApp.getFileById(fileId);
    templateFileName = driveFile.getName();
    fileType = driveFile.getMimeType();
    Logger.log('Found file in Drive: ' + templateFileName + ' (type: ' + fileType + ')');
  } catch (e) {
    ui.alert('Error: Cannot access file with ID: ' + fileId + '\n\n' +
             'Error: ' + e.message + '\n\n' +
             'Make sure you have access to this file.');
    return;
  }

  // Check if it's actually a Google Doc
  if (fileType !== 'application/vnd.google-apps.document') {
    ui.alert('Error: Wrong file type!\n\n' +
             'File name: ' + templateFileName + '\n' +
             'File type: ' + fileType + '\n\n' +
             'This is NOT a Google Doc. Please use a Google Docs document as the template.');
    return;
  }

  // Now confirm with user using the actual document name
  const confirm = ui.alert(
    'Step 2: Confirm Template Document',
    'Found this Google Doc:\n\n' +
    '📄 Document Name: ' + templateFileName + '\n\n' +
    'File ID: ' + fileId + '\n\n' +
    'Is "' + templateFileName + '" the correct template?',
    ui.ButtonSet.YES_NO
  );

  if (confirm === ui.Button.NO) {
    ui.alert('Cancelled.\n\n' +
             'Please update cell C2 in "email details" sheet with the correct Google Doc URL.\n\n' +
             'Currently has: ' + templateFileName);
    return;
  }

  // Finally, open the document
  let templateDoc;
  try {
    templateDoc = DocumentApp.openById(fileId);
    Logger.log('Successfully opened template document: ' + templateFileName);
  } catch (e) {
    ui.alert('Error: Cannot open document\n\n' +
             'Document: ' + templateFileName + '\n' +
             'Error: ' + e.message);
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

  // Create folder with timestamp in the same location as the template
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmmss');
  const folderName = `PDF Bundle ${timestamp}`;

  // Get the parent folder of the template document
  const templateFile = DriveApp.getFileById(templateDoc.getId());
  const templateParents = templateFile.getParents();

  let folder;
  if (templateParents.hasNext()) {
    // Create the bundle folder in the same folder as the template
    const templateFolder = templateParents.next();
    folder = templateFolder.createFolder(folderName);
    Logger.log('Created PDF bundle folder in: ' + templateFolder.getName());
  } else {
    // Fallback: create in root if template has no parent (shouldn't happen)
    folder = DriveApp.createFolder(folderName);
    Logger.log('Created PDF bundle folder in Drive root (template has no parent folder)');
  }

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

  // Generate combined PDF with all letters
  const combinedPdfCreated = generateCombinedPDF(templateDoc, people, folderName, folder);

  // Generate labels PDF
  const labelsPdfCreated = generateLabelsPDF(people, folderName, folder);

  // Get template parent folder name for display (reuse templateFile from above)
  const parentFolderName = folder.getParents().hasNext() ? folder.getParents().next().getName() : 'Drive Root';

  // Show completion message
  if (labelsPdfCreated && combinedPdfCreated) {
    ui.alert(
      'PDF Bundle Created Successfully!\n\n' +
      '✓ Individual PDFs: ' + generatedCount + '\n' +
      '✓ Combined PDF: Created\n' +
      '✓ Labels PDF: Created\n' +
      '✓ Total people: ' + people.length + '\n\n' +
      'Created in same folder as template:\n' +
      '📁 ' + parentFolderName + '\n\n' +
      'Folder Name: ' + folderName + '\n' +
      'Folder URL: ' + folder.getUrl() + '\n\n' +
      'Files in folder:\n' +
      '• ' + generatedCount + ' personalized PDFs (one per person)\n' +
      '• Combined Letters.pdf (all letters in one file)\n' +
      '• Mailing Labels.pdf (print on Avery 5160 sheets)\n\n' +
      'Click the folder URL above to open it.'
    );
  } else {
    const warnings = [];
    if (!combinedPdfCreated) warnings.push('Combined PDF: FAILED');
    if (!labelsPdfCreated) warnings.push('Labels PDF: FAILED');

    ui.alert(
      'PDF Bundle Partially Created\n\n' +
      'Individual PDFs: ' + generatedCount + '\n' +
      (warnings.length > 0 ? warnings.join('\n') + '\n\n' : '') +
      'Created in: ' + parentFolderName + '\n' +
      'Folder: ' + folderName + '\n' +
      'Location: ' + folder.getUrl() + '\n\n' +
      'Check View → Logs for error details.'
    );
  }
}

/** ========================== DEBUG FUNCTION ================== **/

/**
 * DEBUG: Test replacement on ONE person and show document structure
 * Run this from Script Editor to diagnose the issue
 */
function debugReplacementForOnePerson() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  // Get template
  const detailsSh = ss.getSheetByName('email details');
  const attachmentRef = String(detailsSh.getRange('C2').getValue()).trim();
  const fileId = extractDriveId(attachmentRef);
  const templateDoc = DocumentApp.openById(fileId);

  // Get first person with ADDRESS LINE 2
  const listSh = ss.getSheetByName('People');
  const values = listSh.getRange(2, 1, listSh.getLastRow() - 1, 5).getDisplayValues();

  let testPerson = null;
  for (let i = 0; i < values.length; i++) {
    const fullName = String(values[i][0] || '').trim();
    const address = String(values[i][4] || '').trim();
    if (fullName && address) {
      const addressLines = parseAddress(address);
      if (addressLines.line2) {
        // Found someone with ADDRESS LINE 2
        testPerson = {
          fullName: fullName,
          firstName: extractFirstName(fullName),
          pacName: String(values[i][1] || '').trim(),
          email: String(values[i][2] || '').trim(),
          phone: String(values[i][3] || '').trim(),
          address: address
        };
        testPerson = normalizePersonData(testPerson);
        break;
      }
    }
  }

  if (!testPerson) {
    ui.alert('No person found with ADDRESS LINE 2');
    return;
  }

  Logger.log('Testing with: ' + testPerson.fullName);
  Logger.log('Address: ' + testPerson.address);

  // Create temp doc
  const tempDocFile = DriveApp.getFileById(templateDoc.getId()).makeCopy('DEBUG_TEST_' + testPerson.fullName);
  const tempDoc = DocumentApp.openById(tempDocFile.getId());
  const body = tempDoc.getBody();

  Logger.log('\n=== BEFORE REPLACEMENT ===');
  const beforeParas = body.getParagraphs();
  for (let i = 0; i < beforeParas.length; i++) {
    Logger.log('Para ' + i + ': "' + beforeParas[i].getText() + '"');
  }

  // Do replacement
  replacePlaceholdersInDocument(body, testPerson);
  tempDoc.saveAndClose();

  // Reopen and show result
  const tempDoc2 = DocumentApp.openById(tempDocFile.getId());
  const body2 = tempDoc2.getBody();

  Logger.log('\n=== AFTER REPLACEMENT ===');
  const afterParas = body2.getParagraphs();
  for (let i = 0; i < afterParas.length; i++) {
    Logger.log('Para ' + i + ': "' + afterParas[i].getText() + '"');
  }

  tempDoc2.close();

  ui.alert('Debug complete!\n\n' +
           'Testing: ' + testPerson.fullName + '\n' +
           'Temp doc created: DEBUG_TEST_' + testPerson.fullName + '\n\n' +
           'Check View → Logs for paragraph-by-paragraph output.\n' +
           'Temp doc URL: ' + tempDocFile.getUrl() + '\n\n' +
           'The temp doc was NOT deleted so you can inspect it manually.');
}

/** ========================== PDF GENERATION =================== **/

/**
 * Creates a personalized PDF from template for one person
 */
function createPersonalizedPDF(templateDoc, personData) {
  let tempDocFile = null;

  try {
    // Log person data for debugging
    Logger.log('Creating PDF for: ' + personData.fullName);
    Logger.log('Person data: ' + JSON.stringify(personData));

    // Make a temporary copy of the template
    tempDocFile = DriveApp.getFileById(templateDoc.getId()).makeCopy('temp_' + personData.fullName);
    const tempDoc = DocumentApp.openById(tempDocFile.getId());
    const body = tempDoc.getBody();

    // Replace placeholders and handle empty values
    replacePlaceholdersInDocument(body, personData);
    tempDoc.saveAndClose();

    // Export as PDF
    const pdfBlob = tempDocFile.getAs('application/pdf');

    // Delete the temporary doc
    tempDocFile.setTrashed(true);

    Logger.log('Successfully created PDF for: ' + personData.fullName);
    return pdfBlob;

  } catch (e) {
    Logger.log('Error in createPersonalizedPDF for ' + personData.fullName + ': ' + e.message);
    Logger.log('Stack: ' + e.stack);

    // Clean up temp file if it exists
    if (tempDocFile) {
      try {
        tempDocFile.setTrashed(true);
      } catch (cleanupError) {
        Logger.log('Could not clean up temp file: ' + cleanupError.message);
      }
    }

    // Re-throw the error so it's caught by the caller
    throw e;
  }
}

/**
 * Replaces all placeholders in a document body
 * If placeholder value is empty, removes the entire paragraph
 */
function replacePlaceholdersInDocument(body, personData) {
  // Parse address into lines if provided
  const addressLines = parseAddress(personData.address || '');

  // Current date in multiple formats
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');
  const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy');

  Logger.log('Replacement values for ' + personData.fullName + ':');
  Logger.log('- ADDRESS LINE 1: "' + addressLines.line1 + '"');
  Logger.log('- ADDRESS LINE 2: "' + addressLines.line2 + '"');
  Logger.log('- CITY STATE ZIP: "' + addressLines.cityStateZip + '"');

  // Define all replacements with their values
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
    'CITY STATE ZIP': addressLines.cityStateZip,
    'CITY, STATE ZIP': addressLines.cityStateZip,
    'CITYSTATEZIP': addressLines.cityStateZip,
    'ADDRESS': personData.address || '',
    'DATE': today,
    'TODAY': today,
    'MONTH DD, YYYY': todayFormatted,
    'Month DD, YYYY': todayFormatted
  };

  // Process each placeholder in REVERSE order
  // This ensures later placeholders (like CITY STATE ZIP) are processed before earlier ones (ADDRESS LINE 2)
  // which may help preserve paragraph boundaries for consecutive placeholders
  const keys = Object.keys(replacements);
  keys.reverse();

  keys.forEach(key => {
    let value = replacements[key];

    // Ensure value is a string
    if (value === null || value === undefined) {
      value = '';
    }
    value = String(value).trim();

    // Patterns to search for (different bracket types)
    const patterns = [
      '\\[' + key + '\\]',
      '\\[' + key.toLowerCase() + '\\]'
    ];

    patterns.forEach(pattern => {
      let searchResult = body.findText(pattern);

      while (searchResult !== null) {
        const element = searchResult.getElement();
        const para = element.getParent().asParagraph();

        if (value === '') {
          // Empty value - remove the entire paragraph
          try {
            if (body.getParagraphs().length > 1) {
              para.removeFromParent();
              Logger.log('Removed empty placeholder paragraph for: ' + key);
            } else {
              para.clear();
            }
          } catch (e) {
            Logger.log('Could not remove paragraph for ' + key + ': ' + e.message);
          }
          // Don't search for more - we removed the paragraph
          break;
        } else {
          // Has value - replace the placeholder text
          let escapedValue = value.replace(/\$/g, '$$$$');

          // Special handling for ADDRESS LINE 2 - add newline to force separation
          if (key === 'ADDRESS LINE 2') {
            escapedValue = escapedValue + '\n';
          }

          body.replaceText(pattern, escapedValue);
          // Continue searching for more occurrences
          searchResult = body.findText(pattern);
        }
      }
    });
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

/** ========================== COMBINED PDF ==================== **/

/**
 * Generates a combined PDF with all personalized letters
 * Each letter starts on a new page
 * Uses the already-generated individual PDFs to ensure perfect page alignment
 */
function generateCombinedPDF(templateDoc, people, folderName, folder) {
  let tempDocFile = null;

  try {
    Logger.log('Creating combined PDF with ' + people.length + ' letters');

    // Create a new Google Doc for the combined letters
    const combinedDoc = DocumentApp.create('Temp Combined - ' + folderName);
    const combinedBody = combinedDoc.getBody();

    // Remove the default paragraph that Google Docs creates
    const initialParagraphs = combinedBody.getParagraphs();
    if (initialParagraphs.length > 0) {
      initialParagraphs[0].clear();
    }

    // Process each person and append their letter to the combined document
    for (let i = 0; i < people.length; i++) {
      const person = people[i];
      Logger.log('Adding letter ' + (i + 1) + ' of ' + people.length + ' for: ' + person.fullName);

      try {
        // Add page break BEFORE this letter (except for first letter)
        // This ensures each letter starts at the TOP of a new page
        if (i > 0) {
          combinedBody.appendPageBreak();
        }

        // Create a temporary copy of the template for this person
        tempDocFile = DriveApp.getFileById(templateDoc.getId()).makeCopy('temp_combined_' + person.fullName);
        const tempDoc = DocumentApp.openById(tempDocFile.getId());
        const tempBody = tempDoc.getBody();

        // Replace placeholders (also removes empty ones)
        replacePlaceholdersInDocument(tempBody, person);
        tempDoc.saveAndClose();

        // Reopen to copy content to combined document
        const tempDoc2 = DocumentApp.openById(tempDocFile.getId());
        const finalBody = tempDoc2.getBody();

        // Copy all content from this letter to the combined document
        const numChildren = finalBody.getNumChildren();
        for (let j = 0; j < numChildren; j++) {
          const element = finalBody.getChild(j);
          const elementType = element.getType();

          // Copy the element to the combined document
          if (elementType === DocumentApp.ElementType.PARAGRAPH) {
            const para = element.asParagraph().copy();
            combinedBody.appendParagraph(para);
          } else if (elementType === DocumentApp.ElementType.TABLE) {
            const table = element.asTable().copy();
            combinedBody.appendTable(table);
          } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
            const listItem = element.asListItem().copy();
            combinedBody.appendListItem(listItem);
          } else if (elementType === DocumentApp.ElementType.PAGE_BREAK) {
            // Skip page breaks from the template - we're controlling them ourselves
            continue;
          }
        }

        tempDoc2.close();

        // Clean up temporary file
        tempDocFile.setTrashed(true);
        tempDocFile = null;

      } catch (e) {
        Logger.log('Error adding letter for ' + person.fullName + ': ' + e.message);
        // Clean up temp file if it exists
        if (tempDocFile) {
          try {
            tempDocFile.setTrashed(true);
          } catch (cleanupError) {
            // Ignore cleanup errors
          }
          tempDocFile = null;
        }
        // Continue with next person
      }
    }

    // Save the combined document
    combinedDoc.saveAndClose();

    // Get the file and export as PDF
    const combinedDocFile = DriveApp.getFileById(combinedDoc.getId());
    const pdfBlob = combinedDocFile.getAs('application/pdf');

    // Save PDF to folder
    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setName('Combined Letters.pdf');

    // Delete temporary combined document
    combinedDocFile.setTrashed(true);

    Logger.log('Combined PDF created successfully with ' + people.length + ' letters');
    return true;

  } catch (e) {
    Logger.log('Error generating combined PDF: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);

    // Clean up any remaining temp file
    if (tempDocFile) {
      try {
        tempDocFile.setTrashed(true);
      } catch (cleanupError) {
        Logger.log('Could not clean up temp file: ' + cleanupError.message);
      }
    }

    return false;
  }
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
