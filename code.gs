// =================================================================
// GLOBAL CONFIGURATION
// =================================================================

// The ID of the Google Sheet that acts as the main database.
const MAIN_DB_ID = "16Oai_3c4H_E2wgC-CUkSk1Eez90_KdtlaqHnHJFclBQ"; 
// The ID of the Google Sheet used as a data source (e.g., for dropdowns).
const DATA_SOURCE_ID = "1zFIyODwGVHP2SbXpv6Aaj1mLDDV2AvdW1xjAMa4bnoQ"; 
// The ID of the Google Drive folder where uploaded files will be stored.
const UPLOAD_FOLDER_ID = "1wvyzaKUtYNLJhRPwXXpbIiKU3P1R2HQj"; 
// A list of file extensions that are allowed for upload.
const ALLOWED_EXTENSIONS = ["pdf", "doc", "docx", "xls", "xlsx", "csv", "ppt", "pptx"];

// =================================================================
// CORE APP SCRIPT FUNCTIONS
// =================================================================

/**
 * Serves the main HTML user interface.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('CHARGES TO BE BILLED TO TENANTS');
}

/**
 * Retrieves the names of all sheets from the data source spreadsheet.
 * @returns {string[]} An array of sheet names.
 */
function getSheetNames() {
  const ss = SpreadsheetApp.openById(DATA_SOURCE_ID);
  return ss.getSheets().map(sheet => sheet.getName());
}

/**
 * Handles file uploads from the client, saves them to Google Drive, and returns the file URL.
 * @param {object} formData The file data received from the client-side form.
 * @returns {object} A result object indicating success or failure, with the file URL or an error message.
 */
function uploadFile(formData) {
  try {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    const fileName = formData.fileName;
    const extension = fileName.split('.').pop().toLowerCase();
    
    if (!ALLOWED_EXTENSIONS.includes(extension)) {
      return { success: false, error: "Invalid file type. Only document and spreadsheet formats are allowed." };
    }

    const blob = Utilities.newBlob(Utilities.base64Decode(formData.fileData), formData.mimeType, fileName);
    const file = folder.createFile(blob);

    return { success: true, url: file.getUrl(), name: file.getName() };
  } catch (e) {
    Logger.log(`File Upload Error: ${e.toString()}`);
    return { success: false, error: "An error occurred during file upload: " + e.message };
  }
}

/**
 * Main function to process and submit form data. It uses LockService to handle concurrent submissions safely.
 * @param {object} data The complete form data from the client.
 * @returns {object} A result object indicating success or failure.
 */
function submitData(data) {
  const lock = LockService.getScriptLock();
  // Wait for up to 30 seconds for other processes to finish.
  lock.waitLock(30000); 

  try {
    const ss = SpreadsheetApp.openById(MAIN_DB_ID);
    const sheet = ss.getSheetByName(data.selectedSheet);
    const consoSheet = ss.getSheetByName("Conso_COG");

    if (!sheet || !consoSheet) {
      return { success: false, error: "Database sheet not found. Please contact the administrator." };
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const values = data.entries.map(entry => [
      timestamp, data.email, data.selectedSheet, data.note, data.startDate,
      data.endDate, entry.checkboxLabel, entry.file1 || "", entry.file2 || "",
      entry.file3 || "", entry.date || ""
    ]);

    // Append data to both the property-specific sheet and the consolidated sheet.
    appendDataToSheets(sheet, consoSheet, values);

    let emailMessage = "We included the database link and the person in charge of providing access to your email. Thanks";
    if (data.sendCopy) {
      try {
        sendConfirmationEmail(data, timestamp);
      } catch (emailError) {
        Logger.log(`Email Sending Error: ${emailError.toString()}`);
        // The main submission was successful, but the email failed. Inform the user.
        emailMessage = "Your submission was successful, but the confirmation email could not be sent due to an issue (e.g., daily quota reached). Please contact the administrator if you need a copy.";
      }
    }

    return { success: true, message: emailMessage };

  } catch (e) {
    Logger.log(`Data Submission Error: ${e.toString()}`);
    return { success: false, error: "A critical error occurred while saving your data. Please contact the administrator. Details: " + e.message };
  } finally {
    // Always release the lock to allow other users to submit.
    lock.releaseLock();
  }
}

// =================================================================
// HELPER FUNCTIONS
// =================================================================

/**
 * Appends rows of data to the specified sheets efficiently.
 * @param {SpreadsheetApp.Sheet} sheet The primary Google Sheet object.
 * @param {SpreadsheetApp.Sheet} consoSheet The consolidated Google Sheet object.
 * @param {Array<Array<string>>} values The 2D array of data to append.
 */
function appendDataToSheets(sheet, consoSheet, values) {
  // Find the next available row by checking for emptiness in columns A-K.
  const startRow = findFirstBlankRowInRange(sheet);
  const consoRow = findFirstBlankRowInRange(consoSheet);

  // Write the new data.
  sheet.getRange(startRow, 1, values.length, values[0].length).setValues(values);
  consoSheet.getRange(consoRow, 1, values.length, values[0].length).setValues(values);
}

/**
 * Finds the first row in a sheet where columns A through K are all empty.
 * @param {SpreadsheetApp.Sheet} sheet The sheet to inspect.
 * @returns {number} The 1-based index of the first blank row.
 */
function findFirstBlankRowInRange(sheet) {
  // We check data down to the last possible row that has content.
  const lastRow = sheet.getLastRow();
  // If the sheet is completely empty, start at row 1.
  if (lastRow === 0) {
    return 1;
  }
  
  // Get all values in the specified range (A:K) at once for efficiency.
  const range = sheet.getRange("A1:K" + lastRow);
  const values = range.getValues();

  // Iterate through the rows to find the first one where all cells (A-K) are empty.
  for (let i = 0; i < values.length; i++) {
    // .join('') is a fast way to concatenate all cell values into a single string.
    // If the trimmed result is empty, the row is blank.
    if (values[i].join('').trim() === '') {
      return i + 1; // Return the 1-based row number
    }
  }

  // If no blank row is found within the existing data, the next blank row is after the last row.
  return lastRow + 1;
}

/**
 * Composes and sends an HTML confirmation email to the user with links to their submitted files.
 * @param {object} data The form data object.
 * @param {string} timestamp The formatted timestamp of the submission.
 */
function sendConfirmationEmail(data, timestamp) {
  const fileLinks = [];
  data.entries.forEach(entry => {
    const files = [
      { label: `${entry.checkboxLabel} - Signed BTT`, url: entry.file1 },
      { label: `${entry.checkboxLabel} - Editable File`, url: entry.file2 },
      { label: `${entry.checkboxLabel} - Utility Billing`, url: entry.file3 }
    ];

    files.forEach(fileObj => {
      if (fileObj.url) {
        try {
          const fileId = extractFileIdFromUrl(fileObj.url);
          if (fileId) {
            const file = DriveApp.getFileById(fileId);
            fileLinks.push(`<li>${fileObj.label}: <a href="${file.getUrl()}" target="_blank">${file.getName()}</a></li>`);
          }
        } catch (e) {
          Logger.log(`Failed to retrieve file from URL: ${fileObj.url} — ${e.toString()}`);
          fileLinks.push(`<li>${fileObj.label}: Error retrieving file link.</li>`);
        }
      }
    });
  });

  const databaseUrl = `https://docs.google.com/spreadsheets/d/${MAIN_DB_ID}/edit`;
  const headerImageUrl = "https://ik.imagekit.io/nc8qwsrvq/Adobe%20Express%20-%20file.png?updatedAt=1749435937841";

  const htmlMessage = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto;">
      <img src="${headerImageUrl}" alt="Header" style="width: 100%; height: auto;">
      <p>Thank you for your submission.</p>
      <h3>Details:</h3>
      <ul>
        <li><strong>Timestamp:</strong> ${timestamp}</li>
        <li><strong>Email:</strong> ${data.email}</li>
        <li><strong>Property:</strong> ${data.selectedSheet}</li>
        <li><strong>Note:</strong> ${data.note}</li>
        <li><strong>Date Range:</strong> ${data.startDate} to ${data.endDate}</li>
      </ul>
      <h3>Entries:</h3>
      <ul>${data.entries.map(e => `<li>${e.checkboxLabel} (Billing Date: ${e.date || "N/A"})</li>`).join("")}</ul> 
      <h3>Files Submitted:</h3>
      <ul>${fileLinks.join("")}</ul>
      <hr>
      <p>To view the database, click this <a href="${databaseUrl}" target="_blank">LINK</a>.<br>
      To request access, please contact <a href="mailto:jecastro@megaworld-lifestyle.com">jecastro@megaworld-lifestyle.com</a>.</p>
    </div>`;

  MailApp.sendEmail({
    to: data.email,
    subject: "BTT Confirmation of Submission",
    htmlBody: htmlMessage,
    name: "MCD - Data & Process Oversight Unit"
  });
}

/**
 * Extracts the Google Drive file ID from a URL.
 * @param {string} url The Google Drive file URL.
 * @returns {string|null} The extracted file ID or null if not found.
 */
function extractFileIdFromUrl(url) {
  if (!url) return null;
  const regex = /[-\w]{25,}/;
  const match = url.match(regex);
  return match ? match[0] : null;
}


// =================================================================
// SESSION MANAGEMENT
// =================================================================

/**
 * Starts a user session by storing the start time in CacheService.
 */
function startSession() {
  CacheService.getUserCache().put("sessionStart", Date.now().toString(), 1800); // 30 min expiration
  return { success: true };
}

/**
 * Checks if the user's session is still valid.
 */
function checkSession() {
  const cache = CacheService.getUserCache();
  const start = cache.get("sessionStart");
  if (!start) return { expired: true };

  const now = Date.now();
  const duration = 30 * 60 * 1000;

  if ((now - parseInt(start, 10)) > duration) {
    cache.remove("sessionStart");
    return { expired: true };
  }

  return { expired: false };
}
