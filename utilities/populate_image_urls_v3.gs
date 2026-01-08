/**
 * 3DSellers Image URL Populator v3
 * Populates Image 1-17 columns (W-AM) based on SKU in Column A
 * Uses FULL SKU names: {full_sku}_crop_{1-8}.jpg
 *
 * Image structure:
 * - Images 1-8: Product crops from gauntlet.gallery/images/products/
 * - Images 9-12: Empty (reserved for future mockups)
 * - Images 13-17: Stock images A-E from gauntlet.gallery/images/stock/
 */

// Configuration
const BASE_URL = "https://gauntlet.gallery/images/products/";
const STOCK_URL = "https://gauntlet.gallery/images/stock/";
const STOCK_IMAGES = ["A.png", "B.png", "C.png", "D.png", "E.png"];
const PRODUCTS_SHEET_NAME = "PRODUCTS";
const SKU_COLUMN = 1;        // Column A
const IMAGE_START_COLUMN = 23; // Column W (W=23)
const NUM_IMAGE_COLUMNS = 17;

/**
 * Get image URLs for a given full SKU
 * Returns array of 17 URLs (8 crops + 4 empty + 5 stock)
 */
function getImageUrls(fullSku) {
  const urls = [];

  // Add 8 crop images using full SKU naming: {fullSku}_crop_{1-8}.jpg
  for (let i = 1; i <= 8; i++) {
    urls.push(BASE_URL + fullSku + "_crop_" + i + ".jpg");
  }

  // Add 4 empty slots (reserved for mockups)
  for (let i = 0; i < 4; i++) {
    urls.push("");
  }

  // Add 5 stock images
  for (const stockImg of STOCK_IMAGES) {
    urls.push(STOCK_URL + stockImg);
  }

  return urls;
}

/**
 * Main function to populate all image URLs
 * Run this from the Script Editor
 */
function populateAllImageUrls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet '" + PRODUCTS_SHEET_NAME + "' not found!");
    return;
  }

  const lastRow = sheet.getLastRow();
  const skuRange = sheet.getRange(2, SKU_COLUMN, lastRow - 1, 1); // Skip header row
  const skuValues = skuRange.getValues();

  let updatedCount = 0;
  let skippedCount = 0;
  const skippedSkus = [];

  for (let i = 0; i < skuValues.length; i++) {
    const fullSku = skuValues[i][0];

    // Check if SKU exists and starts with a number followed by underscore
    if (fullSku && fullSku.toString().match(/^\d+_/)) {
      const imageUrls = getImageUrls(fullSku);
      const rowNum = i + 2; // +2 because we start at row 2 (skip header)

      // Write all 17 image URLs to columns W-AM
      sheet.getRange(rowNum, IMAGE_START_COLUMN, 1, NUM_IMAGE_COLUMNS).setValues([imageUrls]);
      updatedCount++;
    } else {
      skippedCount++;
      if (fullSku) skippedSkus.push(fullSku);
    }
  }

  let message = "Done!\n\n" +
    "Updated: " + updatedCount + " rows\n" +
    "Skipped: " + skippedCount + " rows";

  if (skippedSkus.length > 0 && skippedSkus.length <= 10) {
    message += "\n\nSkipped SKUs:\n" + skippedSkus.join("\n");
  }

  SpreadsheetApp.getUi().alert(message);
}

/**
 * Add menu item to run the script
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Image Tools')
    .addItem('Populate Image URLs', 'populateAllImageUrls')
    .addToUi();
}
