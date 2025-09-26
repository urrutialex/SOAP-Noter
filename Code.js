// === Utility: Render SOAP Table ===
function renderSOAPTable(body, header, rows, headerColor, position) {
  const tableRows = [header, ...rows];
  // Accept an optional position argument for where to insert the table
  if (typeof position !== 'number') position = 0;
  const table = body.insertTable(position, tableRows);
  const headerRow = table.getRow(0);
  for (let i = 0; i < header.length; i++) {
    headerRow.getCell(i)
      .setBold(true)
      .setFontSize(13)
      .setUnderline(true);
    if (headerColor) {
      headerRow.getCell(i).setBackgroundColor(headerColor);
    }
  }
  const sectionColWidth = 160;
  const responseColWidth = 400;
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    row.getCell(0).setWidth(sectionColWidth);
    row.getCell(1).setWidth(responseColWidth);
  }
  for (let r = 1; r < table.getNumRows(); r++) {
    table.getRow(r).getCell(1).setBold(false);
  }
}
// === Note Type Metadata Map ===
const NOTE_TYPE_META = {
  'Direct Therapy': { color: '#cfe2f3', display: 'Direct Therapy' },
  'Supervision': { color: '#d9d2e9', display: 'Supervision' },
  'Parent Training': { color: '#d9ead3', display: 'Parent Training' },
  'Caregiver Readiness': { color: '#fff2cc', display: 'Caregiver Readiness' },
};
// === Utility: Centralized Date Formatting ===
function formatSessionDate(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'M/d/yyyy');
  }
  const parsed = new Date(val);
  if (!isNaN(parsed.getTime()) && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(val)) {
    return val;
  } else if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'M/d/yyyy');
  }
  return val;
}
// === Utility: Find Shared Drive by Name ===
function getDriveIdByName(name) {
  let pageToken;
  do {
    const drivesResponse = Drive.Drives.list({
      pageSize: 100,
      pageToken,
      useDomainAdminAccess: true
    });
    if (drivesResponse.drives) {
      for (let drive of drivesResponse.drives) {
        if (drive.name === name) {
          return drive.id;
        }
      }
    }
    pageToken = drivesResponse.nextPageToken;
  } while (pageToken);
  return null;
}

// === Utility: Find Folder by Name in Drive ===
function getFolderIdByName(folderName, driveId) {
  const folderSearch = Drive.Files.list({
    q: `name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false and '${driveId}' in parents`,
    corpora: 'drive',
    driveId,
    includeItemsFromAllDrives: true,
    supportsAllDrives: true
  });
  if (!folderSearch.files || folderSearch.files.length === 0) return null;
  return folderSearch.files[0].id;
}

// === Utility: Ensure Folder exists in Drive (create if missing) ===
function ensureFolderInDrive(folderName, driveId) {
  let folderId = getFolderIdByName(folderName, driveId);
  if (folderId) return folderId;
  try {
    const folder = Drive.Files.create({
      name: folderName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [driveId]
    }, null, { supportsAllDrives: true });
    Logger.log(`📁 Created folder '${folderName}' in Drive ID '${driveId}'.`);
    return folder.id;
  } catch (err) {
    Logger.log(`❌ Failed to create folder '${folderName}' in Drive ID '${driveId}': ${err}`);
    return null;
  }
}

// === Utility: Find Doc by Name in Folder ===
function getDocIdByPrefix(docNamePrefix, folderId, driveId) {
  // Search for any doc whose name contains the prefix; prefer most recently modified and true prefix matches
  const docSearch = Drive.Files.list({
    q: `name contains '${docNamePrefix}' and mimeType='application/vnd.google-apps.document' and trashed=false and '${folderId}' in parents`,
    corpora: 'drive',
    driveId,
    includeItemsFromAllDrives: true,
    supportsAllDrives: true,
    orderBy: 'modifiedTime desc',
    pageSize: 50
  });
  if (!docSearch.files || docSearch.files.length === 0) return null;
  const exact = docSearch.files.find(f => f.name && f.name.indexOf(docNamePrefix) === 0);
  const file = exact || docSearch.files[0];
  return file.id;
}

// === Utility: Derive two-letter doc prefix from Job Code/Drive Name ===
function deriveDocPrefix(jobCode) {
  const cleanInitials = s => (s || '').replace(/[^A-Za-z]/g, '').toUpperCase();
  if (!jobCode || typeof jobCode !== 'string') return null;
  // Take the part before the first parenthesis, e.g., "John S. (ABA)" -> "John S."
  const namePart = jobCode.split('(')[0].trim();
  if (!namePart) return null;

  const words = namePart.split(/\s+/).filter(Boolean);
  if (words.length >= 2) {
    const first = cleanInitials(words[0])[0];
    const last = cleanInitials(words[words.length - 1])[0];
    if (first && last) return first + last;
  } else if (words.length === 1) {
    const token = cleanInitials(words[0]);
    if (token.length >= 2) return token.slice(0, 2); // e.g., "CUSD" -> "CU"
    if (token.length === 1) return token + 'X';
  }
  return null;
}

// === Utility: Create New SOAP Log Document ===
function createSOAPLogDocument(jobCode, folderId, driveId) {
  try {
  // Derive two-letter prefix from job code
  const derived = deriveDocPrefix(jobCode);
  const docPrefix = derived || 'XX'; // Default fallback
    
    // Format today's date as mmddyy
    const today = new Date();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    const yy = String(today.getFullYear()).slice(-2);
    const docName = `${docPrefix}_SOAP_LOG_${mm}${dd}${yy}`;
    
    // Create the document
    const docFile = Drive.Files.create({
      name: docName,
      mimeType: 'application/vnd.google-apps.document',
      parents: [folderId]
    }, null, { supportsAllDrives: true });
    
  Logger.log(`✅ Created new SOAP log: ${docName} (ID: ${docFile.id}) in Session Notes (S.O.A.P.)`);
    return docFile.id;
    
  } catch (error) {
    Logger.log(`❌ Error creating SOAP log document: ${error.toString()}`);
    return null;
  }
}
// CONFIGURATION
const JOB_CODE_COLUMN_NAME = 'Job Code';
const TARGET_DOC_NAME = 'Client SOAP Notes';
const SOURCE_SHEET_ID = '11okaksQLR-sidlOxysZfsqgGhAWqkJMBtzarq1YGsH4';
const SHEET_NAME = 'Form Responses 1';
const USE_ACTIVE_SPREADSHEET = true;

// === Shared Spreadsheet Access Function ===
function getSpreadsheet() {
  try {
    if (USE_ACTIVE_SPREADSHEET) {
      Logger.log('Using active spreadsheet...');
      return SpreadsheetApp.getActiveSpreadsheet();
    } else {
      if (SOURCE_SHEET_ID === 'YOUR_SHEET_ID_HERE') {
        throw new Error('Update SOURCE_SHEET_ID or set USE_ACTIVE_SPREADSHEET = true');
      }
      Logger.log(`Opening spreadsheet by ID: ${SOURCE_SHEET_ID}`);
      return SpreadsheetApp.openById(SOURCE_SHEET_ID);
    }
  } catch (error) {
    Logger.log(`❌ Error accessing spreadsheet: ${error.toString()}`);
    throw error;
  }
}

// === Shared Row Processing Logic ===

function processLatestRowIfUnprocessed(sheet) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('❌ No data rows found.');
    return;
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log(`Headers found: ${headers.join(', ')}`);
  // Find the Timestamp column index
  const timestampColIdx = headers.indexOf('Timestamp');
  const responseIdColIdx = headers.findIndex(h => h && h.toLowerCase().includes('response') && h.toLowerCase().includes('id'));
  if (timestampColIdx === -1 || responseIdColIdx === -1) {
    Logger.log('❌ Required columns not found (Timestamp or Response ID).');
    return;
  }

  // Get all data rows
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Find the row with the most recent Timestamp
  let latestIdx = -1;
  let latestTimestamp = null;
  for (let i = 0; i < data.length; i++) {
    const ts = data[i][timestampColIdx];
    if (!ts) continue;
    const tsDate = (ts instanceof Date) ? ts : new Date(ts);
    if (!latestTimestamp || tsDate > latestTimestamp) {
      latestTimestamp = tsDate;
      latestIdx = i;
    }
  }

  if (latestIdx === -1) {
    Logger.log('No rows with Timestamp found.');
    return;
  }

  const rowData = data[latestIdx];
  let jobCode = null;
  let responseId = null;
  const responses = [];
  for (let i = 0; i < headers.length; i++) {
    const columnName = headers[i];
    const value = rowData[i];
    if (columnName === JOB_CODE_COLUMN_NAME) {
      jobCode = value;
    } else if (columnName && columnName.toLowerCase().includes('response') && columnName.toLowerCase().includes('id')) {
      responseId = value;
    } else if (columnName && value !== null && value !== '') {
      responses.push({ question: columnName, answer: value });
    }
  }

  if (!jobCode || !responseId) {
    Logger.log(`❌ Job Code or Response ID not found in row ${latestIdx + 2}.`);
    return;
  }

  // Check if already processed
  if (isNoteInDoc(jobCode, responseId)) {
    Logger.log(`Skipping row ${latestIdx + 2} - already processed in Doc.`);
    return;
  }

  processSOAPNote(jobCode, responses, responseId);
  Logger.log(`Row ${latestIdx + 2} (latest Timestamp) processed.`);
}

// === Core Processing Logic: Write SOAP Note to Doc ===
function processSOAPNote(jobCode, responses, responseId, rowNumber = null, uploadTimestampColIdx = -1) {
  try {
    // Use utility functions for Drive/Folder/Doc lookup
    const driveId = getDriveIdByName(jobCode);
    if (!driveId) {
      Logger.log(`❌ Shared Drive '${jobCode}' not found.`);
      return;
    }
    const folderId = ensureFolderInDrive('Session Notes (S.O.A.P.)', driveId);
    if (!folderId) {
      Logger.log(`❌ Could not ensure 'Session Notes (S.O.A.P.)' folder in Drive '${jobCode}'.`);
      return;
    }
    // Determine the doc name prefix using derived initials (handles single-word names like "CUSD (ABA)")
    const docPrefix = deriveDocPrefix(jobCode);
    const docNamePrefix = docPrefix ? `${docPrefix}_SOAP_LOG_` : TARGET_DOC_NAME;
    Logger.log(`Looking for doc. Job Code: '${jobCode}', Doc Name Prefix: '${docNamePrefix}'`);
    let docId = getDocIdByPrefix(docNamePrefix, folderId, driveId);
    
    // If no document found with the prefix, create a new one (only if we have a prefix)
    if (!docId && docPrefix) {
      Logger.log(`❌ No doc starting with '${docNamePrefix}' found. Creating new SOAP log...`);
      docId = createSOAPLogDocument(jobCode, folderId, driveId);
      if (!docId) {
        Logger.log(`❌ Failed to create new SOAP log document.`);
        return;
      }
    } else if (!docId) {
      Logger.log(`❌ No matching log and no prefix derived. Aborting to avoid misfile.`);
      return;
    }
    
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    // Insert horizontal rule at the top
    body.insertHorizontalRule(0);
    // Insert timestamp paragraph just below the horizontal rule, including Response ID
    body.insertParagraph(1, `SOAP Note Entry: ${new Date().toLocaleString()} [Response ID: ${responseId}]`).setBold(true);

    // Get the SOAP Note Type value for the header
    const noteTypeObj = responses.find(r => r.question === 'Select SOAP Note Type');
    const noteType = noteTypeObj ? String(noteTypeObj.answer) : '';

    // Use NOTE_TYPE_META for layout and color
    const meta = NOTE_TYPE_META[noteType] || { color: null, display: noteType };
    // Render the table just below the timestamp paragraph
    renderDefaultSOAP(body, responses, meta.display, meta.color, 2);

    // Insert a blank paragraph below the table for spacing
    body.insertParagraph(3, '');
    doc.saveAndClose();
    
    // Verify the note was written by checking if Response ID is in the Doc
    if (isNoteInDoc(jobCode, responseId)) {
      Logger.log('✅ Verification passed: Note found in Doc.');
      
      // Only mark as processed if verified
      if (rowNumber !== null && uploadTimestampColIdx !== -1) {
        const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
        if (sheet) {
          sheet.getRange(rowNumber, uploadTimestampColIdx + 1).setValue(new Date());
          Logger.log(`✅ Marked row ${rowNumber} as processed with timestamp.`);
        }
      }
      
      Logger.log('✅ SOAP note prepended to top of Google Doc.');
    } else {
      Logger.log('❌ Verification failed: Note not found in Doc after write. Not marking as processed.');
      // Optionally, mark as error for manual review
      if (rowNumber !== null && uploadTimestampColIdx !== -1) {
        const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
        if (sheet) {
          sheet.getRange(rowNumber, uploadTimestampColIdx + 1).setValue('ERROR: Verification failed');
        }
      }
    }
  } catch (error) {
    Logger.log("❌ Error in processSOAPNote: " + error.toString());
  }
}

// === SOAP Note Layouts ===

function renderDefaultSOAP(body, responses, noteType, headerColor, position) {
  const rows = responses
    .filter(r => r.question !== 'Timestamp' && r.question !== 'Select SOAP Note Type')
    .map(r => {
      let answer = r.answer;
      if (r.question === 'Session Date') {
        answer = formatSessionDate(answer);
      } else if (Array.isArray(answer)) {
        answer = answer.join(', ');
      }
      return [String(r.question), String(answer)];
    });
  if (rows.length > 0) {
    renderSOAPTable(body, ['Session Notes', noteType], rows, headerColor, typeof position === 'number' ? position : 0);
  }
}

// === Utility: Check if Note Exists in Doc ===
function isNoteInDoc(jobCode, responseId) {
  try {
    const driveId = getDriveIdByName(jobCode);
    if (!driveId) return false;
    
    const folderId = ensureFolderInDrive('Session Notes (S.O.A.P.)', driveId);
    if (!folderId) return false;
    
    const docPrefix = deriveDocPrefix(jobCode);
    const docNamePrefix = docPrefix ? `${docPrefix}_SOAP_LOG_` : TARGET_DOC_NAME;
    const docId = getDocIdByPrefix(docNamePrefix, folderId, driveId);
    if (!docId) return false; // No Doc found, so not processed
    
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const text = body.getText();
    
    // Search for Response ID for uniqueness
    Logger.log(`Searching for Response ID: ${responseId} in Doc text (length: ${text.length})`);
    if (text.includes(responseId.toString())) {
      Logger.log('Response ID found in Doc.');
      return true;
    } else {
      Logger.log('Response ID not found in Doc.');
      return false;
    }
  } catch (error) {
    Logger.log("Error checking doc: " + error.toString());
    return false; // Assume not processed on error
  }
}

// === Manual Test Helper ===
function testWithRow() {
  try {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log(`❌ Sheet '${SHEET_NAME}' not found.`);
      return;
    }
    Logger.log(`TestWithRow processing sheet: ${sheet.getName()}`);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const uploadTimestampColIdx = headers.indexOf('Upload Timestamp');
    if (uploadTimestampColIdx === -1) {
      Logger.log('❌ Upload Timestamp column not found. Add it as the last column.');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    
    // Collect all unprocessed rows (blank in Upload Timestamp) into an array
    const unprocessedRows = [];
    for (let r = 2; r <= lastRow; r++) {
      const uploadValue = sheet.getRange(r, uploadTimestampColIdx + 1).getValue();
      if (!uploadValue) {
        unprocessedRows.push(r);
      }
    }
    
    Logger.log(`Found ${unprocessedRows.length} unprocessed rows: ${unprocessedRows.join(', ')}`);
    
    // Process only the unprocessed rows
    for (const rowNumber of unprocessedRows) {
      const rowData = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      let jobCode = null;
      let timestamp = null;
      let responseId = null;
      const responses = [];
      
      for (let i = 0; i < headers.length; i++) {
        const columnName = headers[i];
        const value = rowData[i];
        if (columnName === JOB_CODE_COLUMN_NAME) {
          jobCode = value;
        } else if (columnName === 'Timestamp') {
          timestamp = value;
        } else if (columnName && columnName.toLowerCase().includes('response') && columnName.toLowerCase().includes('id')) {
          responseId = value;
        } else if (columnName && value !== null && value !== '' && columnName !== 'Upload Timestamp') {
          responses.push({ question: columnName, answer: value });
        }
      }
      
      if (!jobCode || !timestamp || !responseId) {
        Logger.log(`❌ Missing data in row ${rowNumber}. Skipping.`);
        continue;
      }
      
      // Check if already processed
      if (isNoteInDoc(jobCode, responseId)) {
        Logger.log(`Skipping row ${rowNumber} - already processed in Doc.`);
        // Mark as processed anyway to avoid re-checking
        sheet.getRange(rowNumber, uploadTimestampColIdx + 1).setValue(new Date());
        continue;
      }
      
      processSOAPNote(jobCode, responses, responseId, rowNumber, uploadTimestampColIdx);
    }
  } catch (error) {
    Logger.log("Error in testWithRow: " + error.toString());
  }
}

// === New Audit Function: Check and Re-Process Failed Rows ===
function auditAndReprocess() {
  try {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log(`❌ Sheet '${SHEET_NAME}' not found.`);
      return;
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const uploadTimestampColIdx = headers.indexOf('Upload Timestamp');
    if (uploadTimestampColIdx === -1) {
      Logger.log('❌ Upload Timestamp column not found.');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    
    Logger.log('Starting audit of processed rows...');
    let reprocessedCount = 0;
    
    for (let r = 2; r <= lastRow; r++) {
      const uploadValue = sheet.getRange(r, uploadTimestampColIdx + 1).getValue();
      if (!uploadValue || uploadValue.toString().startsWith('ERROR:')) {
        // Skip unprocessed or already errored rows
        continue;
      }
      
      // Extract data for this row
      const rowData = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getValues()[0];
      let jobCode = null;
      let timestamp = null;
      let responseId = null;
      const responses = [];
      
      for (let i = 0; i < headers.length; i++) {
        const columnName = headers[i];
        const value = rowData[i];
        if (columnName === JOB_CODE_COLUMN_NAME) {
          jobCode = value;
        } else if (columnName === 'Timestamp') {
          timestamp = value;
        } else if (columnName && columnName.toLowerCase().includes('response') && columnName.toLowerCase().includes('id')) {
          responseId = value;
        } else if (columnName && value !== null && value !== '' && columnName !== 'Upload Timestamp') {
          responses.push({ question: columnName, answer: value });
        }
      }
      
      if (!jobCode || !timestamp || !responseId) {
        Logger.log(`❌ Missing data in row ${r}. Skipping audit.`);
        continue;
      }
      
      // Check if the note is actually in the Doc
      if (!isNoteInDoc(jobCode, responseId)) {
        Logger.log(`❌ Audit failed for row ${r}: Note not found in Doc. Re-processing...`);
        // Re-process
        processSOAPNote(jobCode, responses, responseId, r, uploadTimestampColIdx);
        reprocessedCount++;
      } else {
        Logger.log(`✅ Audit passed for row ${r}: Note found in Doc.`);
      }
    }
    
    Logger.log(`Audit complete. Re-processed ${reprocessedCount} rows.`);
  } catch (error) {
    Logger.log("Error in auditAndReprocess: " + error.toString());
  }
}
