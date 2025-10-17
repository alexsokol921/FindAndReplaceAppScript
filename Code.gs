/**
 * Generates personalized Google Docs from a .docx template and Google Sheet data.
 * Assumes Sheet headers are "Name" and "Courses".
 */
function generateDocsFromSheet() {
  // --- CONFIGURATION ---
  const SHEET_NAME = 'Sheet1'; // Change if your sheet name differs
  const TEMPLATE_FILE_ID = ''; // <-- Replace with the .docx file ID in Google Drive
  const OUTPUT_FOLDER_ID = ''; // <-- Replace with a Drive folder ID for the output files

 const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    Logger.log('No data rows.');
    return;
  }

  const headers = values[0].map(h => String(h).trim());
  const nameColIndex = headers.indexOf('Name');
  const coursesColIndex = headers.indexOf('Courses');
  if (nameColIndex === -1 || coursesColIndex === -1) {
    throw new Error('Columns "Name" and/or "Courses" not found in header row.');
  }

  const templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
  const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);

  // iterate rows 1..n (skip header row)
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name = String(row[nameColIndex] || '').trim() || `Row${r+1}`;
    const courses = String(row[coursesColIndex] || '').trim();

    try {
      Logger.log(`Processing row ${r+1}: ${name}`);

      // get the template as blob (keeps original .docx binary content)
      const blob = templateFile.getBlob();

      // upload & convert to Google Docs using Advanced Drive service with convert:true
      const title = `${name} - Document`;
      const resource = {
        title: title,
        parents: [{ id: OUTPUT_FOLDER_ID }]
      };

      // Drive.Files.insert(resource, mediaData, optionalArgs)
      const created = Drive.Files.insert(resource, blob, { convert: true });
      const newFileId = created.id;
      Logger.log(`Uploaded (id=${newFileId}). Waiting for Google Doc conversion...`);

      // Poll until the file's mimeType is google-apps.document (or timeout)
      const maxAttempts = 10;
      let attempt = 0;
      let metadata = null;
      while (attempt < maxAttempts) {
        try {
          metadata = Drive.Files.get(newFileId);
        } catch (e) {
          Logger.log(`Drive.Files.get failed attempt ${attempt+1}: ${e}`);
          metadata = null;
        }

        if (metadata && metadata.mimeType === 'application/vnd.google-apps.document') {
          Logger.log('Conversion confirmed: mimeType is Google Doc.');
          break;
        }

        attempt++;
        Utilities.sleep(1500); // wait 1.5 seconds then try again
      }

      if (!metadata || metadata.mimeType !== 'application/vnd.google-apps.document') {
        // conversion didn't complete in time — log and skip this row
        Logger.log(`⚠️ Conversion did not finish for ${title} (id=${newFileId}). Skipping.`);
        continue;
      }

      // Now safe to open the doc
      const doc = DocumentApp.openById(newFileId);
      const body = doc.getBody();
      // Replace placeholders — escape braces in regex
      body.replaceText('\\(Name\\)', name);
      body.replaceText('\\(Courses\\)', courses);
      doc.saveAndClose();

      Logger.log(`✅ Created and updated document: ${title} (id=${newFileId})`);
    } catch (err) {
      Logger.log(`Error processing row ${r+1} (${name}): ${err}`);
    }
  }
  Logger.log('Finished processing sheet.');
}
