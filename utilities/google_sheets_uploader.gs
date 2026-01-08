/**
 * Google Apps Script - Image Uploader for gauntlet.gallery
 * Uploads images from Google Drive to cPanel via PHP endpoint
 *
 * Setup:
 * 1. Upload upload_image.php to your cPanel public_html folder
 * 2. Set your API_KEY below (must match the PHP file)
 * 3. Run uploadImagesFromDrive() or use the menu
 */

// Configuration - CHANGE THESE
const API_KEY = PropertiesService.getScriptProperties().getProperty('UPLOAD_API_KEY') || '';  // Must match PHP file
const UPLOAD_URL = "https://gauntlet.gallery/upload_image.php";
const DRIVE_FOLDER_ID = "YOUR_GOOGLE_DRIVE_FOLDER_ID";  // Folder containing images

/**
 * Upload a single image from Google Drive to cPanel
 * @param {File} file - Google Drive file object
 * @param {string} folder - 'products' or 'stock'
 * @param {string} newFilename - Optional new filename (uses original if not provided)
 */
function uploadImageToCPanel(file, folder, newFilename) {
  folder = folder || 'products';
  const filename = newFilename || file.getName();

  // Get file as base64
  const blob = file.getBlob();
  const base64Data = Utilities.base64Encode(blob.getBytes());

  // Send to PHP endpoint
  const response = UrlFetchApp.fetch(UPLOAD_URL, {
    method: 'post',
    payload: {
      api_key: API_KEY,
      action: 'upload',
      filename: filename,
      folder: folder,
      data: base64Data
    },
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());

  if (result.success) {
    Logger.log("Uploaded: " + filename + " -> " + result.url);
  } else {
    Logger.log("Failed: " + filename + " - " + result.error);
  }

  return result;
}

/**
 * Upload all images from a Google Drive folder to cPanel
 */
function uploadImagesFromDrive() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const files = folder.getFilesByType(MimeType.JPEG);

  let successCount = 0;
  let failCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const result = uploadImageToCPanel(file, 'products');

    if (result.success) {
      successCount++;
    } else {
      failCount++;
    }

    // Rate limiting - avoid hitting server too hard
    Utilities.sleep(500);
  }

  // Also check for PNG files
  const pngFiles = folder.getFilesByType(MimeType.PNG);
  while (pngFiles.hasNext()) {
    const file = pngFiles.next();
    const result = uploadImageToCPanel(file, 'products');

    if (result.success) {
      successCount++;
    } else {
      failCount++;
    }

    Utilities.sleep(500);
  }

  SpreadsheetApp.getUi().alert(
    "Upload Complete!\n\n" +
    "Uploaded: " + successCount + " files\n" +
    "Failed: " + failCount + " files"
  );
}

/**
 * Upload images for specific SKUs from the sheet
 * Reads SKU from Column A, finds matching image in Drive, uploads with full SKU name
 */
function uploadImagesForSKUs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PRODUCTS");
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  const lastRow = sheet.getLastRow();
  const skuValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  let uploaded = 0;
  let skipped = 0;

  for (let i = 0; i < skuValues.length; i++) {
    const fullSku = skuValues[i][0];
    if (!fullSku) continue;

    // Look for matching files in Drive
    const skuNum = fullSku.toString().split('_')[0];

    for (let cropNum = 1; cropNum <= 8; cropNum++) {
      // Try to find the crop file
      const searchName = skuNum + "_crop_" + cropNum + ".jpg";
      const newName = fullSku + "_crop_" + cropNum + ".jpg";

      const searchResults = folder.getFilesByName(searchName);
      if (searchResults.hasNext()) {
        const file = searchResults.next();
        const result = uploadImageToCPanel(file, 'products', newName);

        if (result.success) {
          uploaded++;
        }
        Utilities.sleep(300);
      } else {
        skipped++;
      }
    }
  }

  SpreadsheetApp.getUi().alert(
    "SKU Upload Complete!\n\n" +
    "Uploaded: " + uploaded + " images\n" +
    "Skipped (not found): " + skipped
  );
}

/**
 * List all files currently on cPanel
 */
function listCPanelFiles() {
  const response = UrlFetchApp.fetch(UPLOAD_URL, {
    method: 'post',
    payload: {
      api_key: API_KEY,
      action: 'list',
      folder: 'products'
    },
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());

  if (result.success) {
    Logger.log("Files on cPanel (" + result.files.length + "):");
    result.files.forEach(function(f) {
      Logger.log("  " + f.filename + " (" + f.size + " bytes)");
    });
  } else {
    Logger.log("Error: " + result.error);
  }

  return result;
}

/**
 * Delete a file from cPanel
 */
function deleteCPanelFile(filename, folder) {
  folder = folder || 'products';

  const response = UrlFetchApp.fetch(UPLOAD_URL, {
    method: 'post',
    payload: {
      api_key: API_KEY,
      action: 'delete',
      filename: filename,
      folder: folder
    },
    muteHttpExceptions: true
  });

  return JSON.parse(response.getContentText());
}

/**
 * Test the connection to cPanel
 */
function testConnection() {
  try {
    const result = listCPanelFiles();
    if (result.success) {
      SpreadsheetApp.getUi().alert(
        "Connection Successful!\n\n" +
        "Found " + result.files.length + " files in /images/products/"
      );
    } else {
      SpreadsheetApp.getUi().alert("Connection Failed: " + result.error);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Connection Error: " + e.message);
  }
}

/**
 * Add menu to Google Sheets
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Image Tools')
    .addItem('Populate Image URLs', 'populateAllImageUrls')
    .addSeparator()
    .addItem('Upload Images from Drive', 'uploadImagesFromDrive')
    .addItem('Upload Images for SKUs', 'uploadImagesForSKUs')
    .addItem('List cPanel Files', 'listCPanelFiles')
    .addItem('Test Connection', 'testConnection')
    .addToUi();
}
