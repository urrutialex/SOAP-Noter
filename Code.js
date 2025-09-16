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

// === Entry Point for Form Submissions ===
function onFormSubmit(e) {
  try {
    Logger.log("=== onFormSubmit TRIGGER FIRED ===");
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log(`❌ Sheet '${SHEET_NAME}' not found.`);
      return;
    }
    processLatestRowIfUnprocessed(sheet);
  } catch (error) {
    Logger.log("Error in onFormSubmit: " + error.toString());
  }
}

// === Entry Point for Manual Sheet Changes ===
function onSheetChange(e) {
  try {
    Logger.log("=== onSheetChange TRIGGER FIRED ===");
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      Logger.log(`❌ Sheet '${SHEET_NAME}' not found.`);
      return;
    }
    processLatestRowIfUnprocessed(sheet);
  } catch (error) {
    Logger.log("Error in onSheetChange: " + error.toString());
  }
}

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
  // Find the Timestamp column index
  const timestampColIdx = headers.indexOf('Timestamp');
  if (timestampColIdx === -1) {
    Logger.log('❌ Timestamp column not found.');
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
  const responses = [];
  for (let i = 0; i < headers.length; i++) {
    const columnName = headers[i];
    const value = rowData[i];
    if (columnName === JOB_CODE_COLUMN_NAME) {
      jobCode = value;
    } else if (columnName && value !== null && value !== '') {
      responses.push({ question: columnName, answer: value });
    }
  }

  if (!jobCode) {
    Logger.log(`❌ Job Code not found in row ${latestIdx + 2}.`);
    return;
  }

  processSOAPNote(jobCode, responses);
  Logger.log(`Row ${latestIdx + 2} (latest Timestamp) processed.`);
}

// === Core Processing Logic: Write SOAP Note to Doc ===
function processSOAPNote(jobCode, responses) {
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
    // Insert timestamp paragraph just below the horizontal rule
    body.insertParagraph(1, `SOAP Note Entry: ${new Date().toLocaleString()}`).setBold(true);

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
    Logger.log('✅ SOAP note prepended to top of Google Doc.');
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

// === Manual Test Helper ===
function testWithRow(rowNumber) {
  try {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet || rowNumber <= 1 || rowNumber > sheet.getLastRow()) {
      Logger.log(`❌ Invalid sheet or row ${rowNumber}`);
      return;
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];

    let jobCode = null;
    const responses = [];

    for (let i = 0; i < headers.length; i++) {
      const columnName = headers[i];
      const value = rowData[i];
      if (columnName === JOB_CODE_COLUMN_NAME) {
        jobCode = value;
      } else if (columnName && value !== null && value !== '') {
        responses.push({ question: columnName, answer: value });
      }
    }

    if (!jobCode) {
      Logger.log('❌ Job Code not found.');
      return;
    }

    processSOAPNote(jobCode, responses);
  } catch (error) {
    Logger.log("Error in testWithRow: " + error.toString());
  }
}
