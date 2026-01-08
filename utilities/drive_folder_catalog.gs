/**
 * Google Drive Folder Catalog Generator
 * Creates a well-formatted sheet with file name, thumbnail, link, and date
 *
 * Setup:
 * 1. Set FOLDER_ID to your Google Drive folder ID
 * 2. Run createCatalogSheet()
 */

// Configuration - CHANGE THIS
const CATALOG_FOLDER_ID = "YOUR_GOOGLE_DRIVE_FOLDER_ID";  // Get from Drive URL

/**
 * Main function - creates a new sheet with file catalog
 */
function createCatalogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create or get the catalog sheet
  let sheet = ss.getSheetByName("File Catalog");
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet("File Catalog");
  }

  // Set up headers
  const headers = ["Thumbnail", "File Name", "SKU #", "Date Created", "Size (KB)", "Google Drive Link"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#4285f4")
    .setFontColor("white");

  // Get folder and files
  const folder = DriveApp.getFolderById(CATALOG_FOLDER_ID);
  const files = folder.getFiles();

  const data = [];
  const thumbnails = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const fileId = file.getId();
    const dateCreated = file.getDateCreated();
    const fileSize = Math.round(file.getSize() / 1024); // KB
    const fileUrl = file.getUrl();

    // Extract SKU number from filename (e.g., "8_crop_1.jpg" -> "8")
    const skuMatch = fileName.match(/^(\d+)/);
    const skuNum = skuMatch ? skuMatch[1] : "";

    // Thumbnail URL (Google Drive thumbnail)
    const thumbnailUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w100";

    data.push({
      fileName: fileName,
      skuNum: skuNum,
      dateCreated: dateCreated,
      fileSize: fileSize,
      fileUrl: fileUrl,
      thumbnailUrl: thumbnailUrl,
      fileId: fileId
    });
  }

  // Sort by date (newest first)
  data.sort((a, b) => b.dateCreated - a.dateCreated);

  // Write data to sheet
  const rows = [];
  for (let i = 0; i < data.length; i++) {
    const d = data[i];
    rows.push([
      "",  // Thumbnail placeholder (will add IMAGE formula)
      d.fileName,
      d.skuNum,
      d.dateCreated,
      d.fileSize,
      d.fileUrl
    ]);
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    // Add IMAGE formulas for thumbnails
    for (let i = 0; i < data.length; i++) {
      const thumbnailFormula = '=IMAGE("' + data[i].thumbnailUrl + '", 1)';
      sheet.getRange(i + 2, 1).setFormula(thumbnailFormula);
    }

    // Add hyperlinks for Drive links
    for (let i = 0; i < data.length; i++) {
      const linkFormula = '=HYPERLINK("' + data[i].fileUrl + '", "Open")';
      sheet.getRange(i + 2, 6).setFormula(linkFormula);
    }
  }

  // Format the sheet
  sheet.setColumnWidth(1, 120);  // Thumbnail
  sheet.setColumnWidth(2, 350);  // File Name
  sheet.setColumnWidth(3, 60);   // SKU #
  sheet.setColumnWidth(4, 150);  // Date Created
  sheet.setColumnWidth(5, 80);   // Size
  sheet.setColumnWidth(6, 100);  // Link

  // Set row heights for thumbnails
  for (let i = 2; i <= data.length + 1; i++) {
    sheet.setRowHeight(i, 80);
  }

  // Format date column
  sheet.getRange(2, 4, data.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");

  // Freeze header row
  sheet.setFrozenRows(1);

  // Auto-filter
  sheet.getRange(1, 1, data.length + 1, headers.length).createFilter();

  SpreadsheetApp.getUi().alert(
    "Catalog Created!\n\n" +
    "Total files: " + data.length + "\n" +
    "Sorted by date (newest first)\n\n" +
    "You can filter by SKU # or sort by any column."
  );
}

/**
 * Create catalog for a specific subfolder by name
 */
function createCatalogForSubfolder(subfolderName) {
  const parentFolder = DriveApp.getFolderById(CATALOG_FOLDER_ID);
  const subfolders = parentFolder.getFoldersByName(subfolderName);

  if (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    createCatalogForFolder(subfolder.getId());
  } else {
    SpreadsheetApp.getUi().alert("Subfolder not found: " + subfolderName);
  }
}

/**
 * Catalog with additional cPanel status check
 * Shows which files exist on both Drive and cPanel
 */
function createCatalogWithCPanelStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // First get cPanel files
  const cpanelResponse = UrlFetchApp.fetch("https://gauntlet.gallery/upload_image.php", {
    method: 'post',
    payload: {
      api_key: "dwDDsaIqQPuIwVnYPDdfNR6B2OsDtZNl",
      action: 'list',
      folder: 'products'
    },
    muteHttpExceptions: true
  });

  const cpanelResult = JSON.parse(cpanelResponse.getContentText());
  const cpanelFiles = new Set(cpanelResult.files || []);

  // Create or get the catalog sheet
  let sheet = ss.getSheetByName("File Catalog");
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet("File Catalog");
  }

  // Set up headers (with cPanel status)
  const headers = ["Thumbnail", "File Name", "SKU #", "Date Created", "Size (KB)", "On cPanel?", "Drive Link"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#4285f4")
    .setFontColor("white");

  // Get folder and files
  const folder = DriveApp.getFolderById(CATALOG_FOLDER_ID);
  const files = folder.getFiles();

  const data = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const fileId = file.getId();
    const dateCreated = file.getDateCreated();
    const fileSize = Math.round(file.getSize() / 1024);
    const fileUrl = file.getUrl();

    const skuMatch = fileName.match(/^(\d+)/);
    const skuNum = skuMatch ? skuMatch[1] : "";

    const thumbnailUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w100";

    // Check if file exists on cPanel
    const onCPanel = cpanelFiles.has(fileName) ? "✅ Yes" : "❌ No";

    data.push({
      fileName: fileName,
      skuNum: skuNum,
      dateCreated: dateCreated,
      fileSize: fileSize,
      fileUrl: fileUrl,
      thumbnailUrl: thumbnailUrl,
      onCPanel: onCPanel
    });
  }

  // Sort by date (newest first)
  data.sort((a, b) => b.dateCreated - a.dateCreated);

  // Write data to sheet
  const rows = [];
  for (let i = 0; i < data.length; i++) {
    const d = data[i];
    rows.push([
      "",
      d.fileName,
      d.skuNum,
      d.dateCreated,
      d.fileSize,
      d.onCPanel,
      d.fileUrl
    ]);
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    // Add IMAGE formulas for thumbnails
    for (let i = 0; i < data.length; i++) {
      const thumbnailFormula = '=IMAGE("' + data[i].thumbnailUrl + '", 1)';
      sheet.getRange(i + 2, 1).setFormula(thumbnailFormula);
    }

    // Add hyperlinks
    for (let i = 0; i < data.length; i++) {
      const linkFormula = '=HYPERLINK("' + data[i].fileUrl + '", "Open")';
      sheet.getRange(i + 2, 7).setFormula(linkFormula);
    }
  }

  // Format
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 350);
  sheet.setColumnWidth(3, 60);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 100);

  for (let i = 2; i <= data.length + 1; i++) {
    sheet.setRowHeight(i, 80);
  }

  sheet.getRange(2, 4, data.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, data.length + 1, headers.length).createFilter();

  const onCPanelCount = data.filter(d => d.onCPanel.includes("Yes")).length;

  SpreadsheetApp.getUi().alert(
    "Catalog Created!\n\n" +
    "Total Drive files: " + data.length + "\n" +
    "On cPanel: " + onCPanelCount + "\n" +
    "Missing from cPanel: " + (data.length - onCPanelCount)
  );
}

/**
 * Add menu item
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Image Tools')
    .addItem('Populate Image URLs', 'populateAllImageUrls')
    .addSeparator()
    .addItem('Create File Catalog (Drive)', 'createCatalogSheet')
    .addItem('Create Catalog with cPanel Status', 'createCatalogWithCPanelStatus')
    .addSeparator()
    .addItem('Upload Images from Drive', 'uploadImagesFromDrive')
    .addItem('Test Connection', 'testConnection')
    .addToUi();
}
