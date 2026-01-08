/**
 * SIMPLE FRAME VARIATIONS SCRIPT
 * Art Only = $90, Black Metal Frame = $150, Gold Metal Frame = $175
 *
 * Run: addFrameVariations()
 */

const SHEET_ID = '1qRRIA7rbW_JhDg20xXV1WrZ-yx8ndnlwQrLkapr65Hc';
const SHEET_NAME = 'Products';

const PRICES = {
  ART_ONLY: 90,
  BMF: 150,
  GMF: 175
};

function addFrameVariations() {
  Logger.log('=== STARTING FRAME VARIATIONS ===');

  // Open spreadsheet
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log('ERROR: Sheet not found');
    return;
  }

  Logger.log('Sheet found: ' + sheet.getName());

  // Get headers
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  Logger.log('Columns: ' + lastCol + ', Rows: ' + lastRow);
  Logger.log('Headers: ' + headers.join(', '));

  // Find column indexes
  let skuCol = -1, titleCol = -1, priceCol = -1;

  for (let i = 0; i < headers.length; i++) {
    const h = headers[i].toString().toLowerCase().trim();
    if (h === 'sku' || h.startsWith('sku')) skuCol = i;
    if (h === 'title' || h.startsWith('title')) titleCol = i;
    if (h === 'price' || h.startsWith('price')) priceCol = i;
  }

  Logger.log('SKU col: ' + skuCol + ', Title col: ' + titleCol + ', Price col: ' + priceCol);

  if (skuCol < 0) {
    Logger.log('ERROR: SKU column not found');
    return;
  }

  // Get all data
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  Logger.log('Data rows: ' + data.length);

  // Find base rows (no -BMF or -GMF)
  const baseRows = [];
  for (let i = 0; i < data.length; i++) {
    const sku = data[i][skuCol];
    if (sku && !sku.toString().endsWith('-BMF') && !sku.toString().endsWith('-GMF')) {
      baseRows.push({
        rowIndex: i + 2, // Sheet row number
        data: data[i],
        sku: sku.toString()
      });
    }
  }

  Logger.log('Found ' + baseRows.length + ' base products');

  if (baseRows.length === 0) {
    Logger.log('No base products to process');
    return;
  }

  // Update base row prices to Art Only
  let updatedCount = 0;
  if (priceCol >= 0) {
    Logger.log('Updating base prices to $' + PRICES.ART_ONLY);
    for (const base of baseRows) {
      sheet.getRange(base.rowIndex, priceCol + 1).setValue(PRICES.ART_ONLY);
      updatedCount++;
    }
    Logger.log('Updated ' + updatedCount + ' base prices');
  }

  // Create new rows for BMF and GMF
  const newRows = [];

  for (const base of baseRows) {
    // Black Metal Frame row
    const bmfRow = base.data.slice();
    bmfRow[skuCol] = base.sku + '-BMF';
    if (priceCol >= 0) bmfRow[priceCol] = PRICES.BMF;
    if (titleCol >= 0 && bmfRow[titleCol]) {
      bmfRow[titleCol] = bmfRow[titleCol].toString().replace(/\s*-\s*Art Only/gi, '') + ' - Black Metal Frame';
    }
    newRows.push(bmfRow);

    // Gold Metal Frame row
    const gmfRow = base.data.slice();
    gmfRow[skuCol] = base.sku + '-GMF';
    if (priceCol >= 0) gmfRow[priceCol] = PRICES.GMF;
    if (titleCol >= 0 && gmfRow[titleCol]) {
      gmfRow[titleCol] = gmfRow[titleCol].toString().replace(/\s*-\s*Art Only/gi, '') + ' - Gold Metal Frame';
    }
    newRows.push(gmfRow);
  }

  Logger.log('Created ' + newRows.length + ' new rows');

  // Append new rows ONE AT A TIME to avoid bulk write issues
  if (newRows.length > 0) {
    Logger.log('Appending rows one at a time...');
    let appendedCount = 0;

    for (const row of newRows) {
      try {
        sheet.appendRow(row);
        appendedCount++;
        if (appendedCount % 10 === 0) {
          Logger.log('Appended ' + appendedCount + ' rows...');
        }
      } catch (e) {
        Logger.log('Error appending row: ' + e.message);
      }
    }

    Logger.log('Successfully appended ' + appendedCount + ' rows');
  }

  Logger.log('=== COMPLETE ===');
  Logger.log('Updated ' + updatedCount + ' base products to $' + PRICES.ART_ONLY);
  Logger.log('Created ' + newRows.length + ' new framed variations');
}

// Quick test function
function testConnection() {
  Logger.log('Testing...');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log('Spreadsheet: ' + ss.getName());
  const sheet = ss.getSheetByName(SHEET_NAME);
  Logger.log('Sheet: ' + sheet.getName());
  Logger.log('Rows: ' + sheet.getLastRow());
  Logger.log('Columns: ' + sheet.getLastColumn());
}
