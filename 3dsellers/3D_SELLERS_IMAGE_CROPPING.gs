/**
 * 3D SELLERS IMAGE CROPPING
 * Google Apps Script for the image cropping sheet
 * Works with MASTER_IMAGE_CROPPING_SYSTEM.py
 */

// ============================================================================
// MENU
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('3D Sellers Cropping')
    .addItem('View Crop Stats', 'showCropStats')
    .addItem('Clear Processed Cache', 'clearProcessedCache')
    .addItem('Setup Sheet Headers', 'setupHeaders')
    .addItem('Test Menu', 'testMenu')
    .addToUi();
}

function testMenu() {
  SpreadsheetApp.getUi().alert('Menu working!');
}

// ============================================================================
// SETUP HEADERS
// ============================================================================

function setupHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var headers = [
    'SKU',                    // A
    'Artist',                 // B
    'Type',                   // C
    'Category',               // D
    'Orientation',            // E
    'Folder Link',            // F
    'Crop 1: BL Signature',   // G
    'Crop 2: BR Edition',     // H
    'Crop 3: TL Texture',     // I
    'Crop 4: TR Texture',     // J
    'Crop 5: Center Detail',  // K
    'Crop 6: Upper Subject',  // L
    'Processed Date',         // M
    'Source Folder'           // N
  ];

  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('white');

  // Set column widths
  sheet.setColumnWidth(1, 180);  // SKU
  sheet.setColumnWidth(2, 120);  // Artist
  sheet.setColumnWidth(3, 100);  // Type
  sheet.setColumnWidth(4, 120);  // Category
  sheet.setColumnWidth(5, 100);  // Orientation
  sheet.setColumnWidth(6, 200);  // Folder Link

  // Crop columns
  for (var i = 7; i <= 12; i++) {
    sheet.setColumnWidth(i, 150);
  }

  sheet.setColumnWidth(13, 150); // Date
  sheet.setColumnWidth(14, 150); // Source

  // Freeze header row
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('Headers setup complete!');
}

// ============================================================================
// STATS
// ============================================================================

function showCropStats() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('No data yet.');
    return;
  }

  var stats = {
    total: data.length - 1,
    byArtist: {},
    byType: {},
    byOrientation: { PORTRAIT: 0, LANDSCAPE: 0 }
  };

  for (var i = 1; i < data.length; i++) {
    var artist = data[i][1] || 'Unknown';
    var type = data[i][2] || 'Unknown';
    var orientation = data[i][4] || 'Unknown';

    stats.byArtist[artist] = (stats.byArtist[artist] || 0) + 1;
    stats.byType[type] = (stats.byType[type] || 0) + 1;

    if (orientation === 'PORTRAIT') stats.byOrientation.PORTRAIT++;
    else if (orientation === 'LANDSCAPE') stats.byOrientation.LANDSCAPE++;
  }

  var msg = 'CROP STATISTICS\n\n';
  msg += 'Total Images: ' + stats.total + '\n\n';

  msg += 'By Artist:\n';
  for (var artist in stats.byArtist) {
    msg += '  ' + artist + ': ' + stats.byArtist[artist] + '\n';
  }

  msg += '\nBy Type:\n';
  for (var type in stats.byType) {
    msg += '  ' + type + ': ' + stats.byType[type] + '\n';
  }

  msg += '\nBy Orientation:\n';
  msg += '  Portrait: ' + stats.byOrientation.PORTRAIT + '\n';
  msg += '  Landscape: ' + stats.byOrientation.LANDSCAPE + '\n';

  SpreadsheetApp.getUi().alert(msg);
}

// ============================================================================
// THUMBNAIL FORMULAS
// ============================================================================

function addThumbnailFormulas() {
  /**
   * Add IMAGE() formulas to show thumbnails for crop links
   * Run this after Python populates the sheet
   */
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = i + 1;

    // For each crop column (G-L), add IMAGE formula if there's a link
    for (var col = 7; col <= 12; col++) {
      var link = data[i][col - 1];
      if (link && link.toString().indexOf('drive.google.com') >= 0) {
        // Extract file ID and create IMAGE formula
        var formula = '=IMAGE("' + link + '",1)';
        sheet.getRange(row, col).setFormula(formula);
      }
    }
  }

  // Set row heights for thumbnails
  for (var i = 2; i <= data.length; i++) {
    sheet.setRowHeight(i, 80);
  }

  SpreadsheetApp.getUi().alert('Added thumbnail formulas for ' + (data.length - 1) + ' rows');
}

// ============================================================================
// CLEAR CACHE
// ============================================================================

function clearProcessedCache() {
  // This is managed by the Python script, but we can clear sheet-side cache
  PropertiesService.getScriptProperties().deleteProperty('LAST_PROCESSED');
  SpreadsheetApp.getUi().alert('Cache cleared. Python script manages main processed list.');
}

// ============================================================================
// TRIGGER FOR AUTO-PROCESSING (OPTIONAL)
// ============================================================================

function createTimeTrigger() {
  // Create trigger to run every 30 minutes
  ScriptApp.newTrigger('checkForNewCrops')
    .timeBased()
    .everyMinutes(30)
    .create();

  SpreadsheetApp.getUi().alert('Trigger created - will check every 30 minutes');
}

function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  SpreadsheetApp.getUi().alert('All triggers deleted');
}

function checkForNewCrops() {
  // Placeholder - the Python script handles the actual processing
  // This could be used to send notifications or update dashboards
  Logger.log('Check triggered at ' + new Date());
}

// ============================================================================
// EXPORT FUNCTIONS
// ============================================================================

function exportToCSV() {
  /**
   * Export sheet data to CSV format
   */
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var csv = '';
  for (var i = 0; i < data.length; i++) {
    var row = data[i].map(function(cell) {
      // Escape quotes and wrap in quotes
      return '"' + cell.toString().replace(/"/g, '""') + '"';
    });
    csv += row.join(',') + '\n';
  }

  // Create file in Drive
  var folder = DriveApp.getRootFolder();
  var filename = 'crop_export_' + new Date().toISOString().slice(0, 10) + '.csv';
  var file = folder.createFile(filename, csv, MimeType.CSV);

  SpreadsheetApp.getUi().alert('Exported to: ' + file.getUrl());
}

// ============================================================================
// LINK VALIDATION
// ============================================================================

function validateLinks() {
  /**
   * Check that all crop links are valid
   */
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var invalid = [];

  for (var i = 1; i < data.length; i++) {
    var sku = data[i][0];

    // Check crop columns (G-L)
    for (var col = 6; col <= 11; col++) {
      var link = data[i][col];
      if (!link || link.toString().trim() === '') {
        invalid.push(sku + ' - Column ' + (col + 1));
      }
    }
  }

  if (invalid.length === 0) {
    SpreadsheetApp.getUi().alert('All links valid!');
  } else {
    SpreadsheetApp.getUi().alert('Missing links (' + invalid.length + '):\n' + invalid.slice(0, 10).join('\n'));
  }
}
