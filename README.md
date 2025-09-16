# SOAP Noter

## Overview

SOAP Noter is a Google Apps Script that automates the creation and management of SOAP (Subjective, Objective, Assessment, Plan) notes for therapy or counseling sessions. It integrates with Google Forms, Sheets, Docs, and Drive to streamline the documentation process.

## How It Works

1. **Form Submission**: A Google Form collects session data (e.g., client information, session notes, SOAP components).
2. **Sheet Storage**: Responses are stored in a Google Sheet.
3. **Automated Processing**: The script triggers on form submission or sheet changes, processing the latest entry.
4. **Document Creation/Update**: It finds or creates a Google Doc in a shared drive (based on the "Job Code"), then prepends a formatted SOAP note table to the document.

## Key Features

- **Automatic Triggers**: Processes new form submissions in real-time.
- **Drive Integration**: Organizes notes in shared drives and folders.
- **SOAP Formatting**: Renders notes in a structured table format with customizable colors for different note types.
- **Date Formatting**: Handles and formats session dates consistently.
- **Error Handling**: Logs errors and handles missing data gracefully.

## Configuration

- **SOURCE_SHEET_ID**: The ID of the Google Sheet containing form responses.
- **SHEET_NAME**: The name of the sheet (default: "Form Responses 1").
- **JOB_CODE_COLUMN_NAME**: The column name for the job code (default: "Job Code").
- **TARGET_DOC_NAME**: Fallback document name if prefix derivation fails.
- **USE_ACTIVE_SPREADSHEET**: Set to `true` to use the active spreadsheet instead of a specific ID.

## Note Types

The script supports different SOAP note types with associated colors:

- **Direct Therapy**: Light blue (#cfe2f3)
- **Supervision**: Light purple (#d9d2e9)
- **Parent Training**: Light green (#d9ead3)
- **Caregiver Readiness**: Light yellow (#fff2cc)

## Functions

- `onFormSubmit(e)`: Trigger for form submissions.
- `onSheetChange(e)`: Trigger for manual sheet edits.
- `processSOAPNote(jobCode, responses)`: Core logic to create/update SOAP notes.
- `createSOAPLogDocument(jobCode, folderId, driveId)`: Creates new SOAP log documents.
- `renderSOAPTable(body, header, rows, headerColor, position)`: Renders the SOAP table in the document.
- `testWithRow(rowNumber)`: Manual testing function for a specific row.

## Setup

1. Clone this script to your Google Apps Script project.
2. Update the configuration variables in `Code.js`.
3. Set up triggers for `onFormSubmit` and `onSheetChange`.
4. Ensure the script has access to Google Drive, Docs, and Sheets APIs.
5. Link your Google Form to the specified Sheet.

## Usage

- Submit data via the connected Google Form.
- The script will automatically process and add notes to the appropriate Google Doc.
- For manual testing, use `testWithRow(rowNumber)` in the Apps Script editor.

## Dependencies

- Google Apps Script runtime (V8)
- Advanced Services: Drive API (v3)
- Exception logging to Stackdriver

## Time Zone

Set to America/Los_Angeles (PST/PDT).

## Local Development

Use Clasp to develop locally:

- `clasp login` - Authenticate
- `clasp clone <scriptId>` - Clone the script
- `clasp push` - Push changes to Google
- `clasp pull` - Pull changes from Google

## Troubleshooting

- Check the execution logs in Google Apps Script for errors.
- Ensure shared drives are accessible and the script has necessary permissions.
- Verify column names in the Sheet match the script's expectations.
