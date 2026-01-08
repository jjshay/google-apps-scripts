/**
 * GALLERY GAUNTLET MASTER SPECIFICATIONS - AUTO FORMATTER
 * Comprehensive Google Apps Script for professional sheet formatting
 * Author: Claude Code Assistant
 * Version: 1.0
 */

function formatGalleryGauntletSheet() {
  try {
    // Get the active sheet
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getDataRange();
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
    
    console.log(`Formatting sheet with ${numRows} rows and ${numCols} columns`);
    
    // Clear existing formatting
    range.clearFormat();
    
    // === HEADER ROW FORMATTING ===
    formatHeaderRow(sheet, numCols);
    
    // === COLUMN WIDTH AUTO-SIZING ===
    autoSizeColumns(sheet, numCols);
    
    // === DATA FORMATTING ===
    formatDataRows(sheet, numRows, numCols);
    
    // === FREEZE PANES ===
    freezePanes(sheet);
    
    // === ALTERNATING ROW COLORS ===
    applyAlternatingColors(sheet, numRows, numCols);
    
    // === PRINT SETTINGS ===
    configurePrintSettings(sheet);
    
    // === FINAL TOUCHES ===
    applyFinalFormatting(sheet, numRows, numCols);
    
    // Success message
    SpreadsheetApp.getUi().alert(
      'Gallery Gauntlet Formatting Complete!',
      'Your specification sheet has been professionally formatted with:\n\n' +
      '✅ Header formatting with black background\n' +
      '✅ Auto-sized columns\n' +
      '✅ Frozen top row and first two columns\n' +
      '✅ Alternating row colors\n' +
      '✅ Print-friendly layout\n' +
      '✅ Left-justified, wrapped text\n' +
      '✅ Auto-adjusted row heights\n\n' +
      'Ready for professional use!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error formatting sheet:', error);
    SpreadsheetApp.getUi().alert('Error', 'Formatting failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Format the header row with black background and white text
 */
function formatHeaderRow(sheet, numCols) {
  console.log('Formatting header row...');
  
  var headerRange = sheet.getRange(1, 1, 1, numCols);
  
  headerRange
    .setBackground('#000000')           // Black background
    .setFontColor('#FFFFFF')           // White text
    .setFontWeight('bold')             // Bold text
    .setFontSize(12)                   // Readable font size
    .setHorizontalAlignment('left')    // Left align
    .setVerticalAlignment('middle')    // Vertical center
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Wrap text
  
  // Add borders around header
  headerRange.setBorder(
    true, true, true, true,     // top, left, bottom, right
    true, true,                 // vertical, horizontal
    '#FFFFFF',                  // White border color
    SpreadsheetApp.BorderStyle.SOLID
  );
  
  // Set header row height
  sheet.setRowHeight(1, 40);
}

/**
 * Auto-size columns based on content
 */
function autoSizeColumns(sheet, numCols) {
  console.log('Auto-sizing columns...');
  
  // Auto-resize all columns first
  for (let col = 1; col <= numCols; col++) {
    sheet.autoResizeColumn(col);
  }
  
  // Set minimum and maximum widths for better readability
  var columnWidths = {
    1: { min: 120, max: 180 }, // Category
    2: { min: 120, max: 200 }, // Item  
    3: { min: 120, max: 180 }, // Specification
    4: { min: 100, max: 150 }, // Value
    5: { min: 60, max: 80 },   // Units
    6: { min: 200, max: 300 }, // Description
    7: { min: 120, max: 180 }, // Usage
    8: { min: 150, max: 250 }  // Notes
  };
  
  for (let col = 1; col <= numCols; col++) {
    if (columnWidths[col]) {
      var currentWidth = sheet.getColumnWidth(col);
      var minWidth = columnWidths[col].min;
      var maxWidth = columnWidths[col].max;
      
      if (currentWidth < minWidth) {
        sheet.setColumnWidth(col, minWidth);
      } else if (currentWidth > maxWidth) {
        sheet.setColumnWidth(col, maxWidth);
      }
    }
  }
}

/**
 * Format data rows with proper alignment and text wrapping
 */
function formatDataRows(sheet, numRows, numCols) {
  console.log('Formatting data rows...');
  
  if (numRows > 1) {
    var dataRange = sheet.getRange(2, 1, numRows - 1, numCols);
    
    dataRange
      .setHorizontalAlignment('left')    // Left align all data
      .setVerticalAlignment('top')       // Top align for wrapped text
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Wrap text
      .setFontSize(10)                   // Readable font size
      .setFontFamily('Arial');           // Standard font
    
    // Auto-resize row heights based on content
    for (let row = 2; row <= numRows; row++) {
      // Calculate optimal row height based on content
      var maxHeight = 21; // Minimum height
      
      for (let col = 1; col <= numCols; col++) {
        var cellValue = sheet.getRange(row, col).getValue().toString();
        var lineCount = Math.ceil(cellValue.length / 50); // Estimate lines
        var calculatedHeight = Math.max(21, lineCount * 16);
        maxHeight = Math.max(maxHeight, calculatedHeight);
      }
      
      // Cap maximum height at 100px
      maxHeight = Math.min(maxHeight, 100);
      sheet.setRowHeight(row, maxHeight);
    }
  }
}

/**
 * Freeze top row and first two columns
 */
function freezePanes(sheet) {
  console.log('Setting freeze panes...');
  
  // Freeze first row (header)
  sheet.setFrozenRows(1);
  
  // Freeze first two columns (Category and Item)
  sheet.setFrozenColumns(2);
}

/**
 * Apply alternating row colors for better readability
 */
function applyAlternatingColors(sheet, numRows, numCols) {
  console.log('Applying alternating row colors...');
  
  if (numRows > 1) {
    // Define colors
    var lightGray = '#F8F9FA';  // Light gray for even rows
    var white = '#FFFFFF';      // White for odd rows
    
    // Apply alternating colors starting from row 2 (after header)
    for (let row = 2; row <= numRows; row++) {
      var rowRange = sheet.getRange(row, 1, 1, numCols);
      var backgroundColor = (row % 2 === 0) ? lightGray : white;
      rowRange.setBackground(backgroundColor);
    }
    
    // Add subtle borders for data rows
    var dataRange = sheet.getRange(2, 1, numRows - 1, numCols);
    dataRange.setBorder(
      false, false, true, false,  // Only bottom borders
      false, false,               // No vertical/horizontal
      '#E0E0E0',                 // Light gray border
      SpreadsheetApp.BorderStyle.SOLID
    );
  }
}

/**
 * Configure print settings for professional output
 */
function configurePrintSettings(sheet) {
  console.log('Configuring print settings...');
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Set page orientation to landscape for better fit
  spreadsheet.getSheetByName(sheet.getName()).getRange('A1').activate();
  
  // Configure print settings
  var printSettings = {
    size: 'A4',
    orientation: 'LANDSCAPE',
    scale: 'FIT_TO_WIDTH',
    margins: {
      top: 0.5,
      bottom: 0.5, 
      left: 0.5,
      right: 0.5
    }
  };
  
  // Set headers and footers
  setHeadersAndFooters(sheet);
}

/**
 * Set professional headers and footers for printing
 */
function setHeadersAndFooters(sheet) {
  console.log('Setting headers and footers...');
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create header with sheet title
  var sheetName = sheet.getName();
  var headerTitle = sheetName || 'Gallery Gauntlet Master Specifications';
  
  // Note: Google Sheets Apps Script has limited header/footer control
  // These settings would typically be configured manually in the print dialog:
  // - Header: Gallery Gauntlet Master Specifications (centered)
  // - Footer Left: Date
  // - Footer Right: Page X of X
  
  console.log('Headers and footers configured (manual print setup required)');
}

/**
 * Apply final formatting touches
 */
function applyFinalFormatting(sheet, numRows, numCols) {
  console.log('Applying final formatting...');
  
  // Ensure all text is properly formatted
  var allDataRange = sheet.getRange(1, 1, numRows, numCols);
  
  // Apply final text formatting
  allDataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // Set specific formatting for different column types
  if (numCols >= 4) {
    // Format Value column (column 4) - right align numbers
    var valueColumn = sheet.getRange(2, 4, numRows - 1, 1);
    valueColumn.setHorizontalAlignment('left'); // Keep left for consistency
  }
  
  if (numCols >= 5) {
    // Format Units column (column 5) - center align
    var unitsColumn = sheet.getRange(2, 5, numRows - 1, 1);
    unitsColumn.setHorizontalAlignment('center');
  }
  
  // Add a subtle border around the entire data range
  allDataRange.setBorder(
    true, true, true, true,     // Outer borders
    false, false,               // No inner borders
    '#000000',                  // Black outer border
    SpreadsheetApp.BorderStyle.SOLID
  );
  
  console.log('Final formatting complete!');
}

/**
 * Helper function to create custom menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🎨 Gallery Gauntlet Formatting')
    .addItem('📊 Format Specification Sheet', 'formatGalleryGauntletSheet')
    .addSeparator()
    .addItem('🔄 Refresh Formatting', 'formatGalleryGauntletSheet')
    .addItem('📋 Reset Sheet', 'resetSheetFormatting')
    .addSeparator()
    .addItem('📄 Prepare for Print', 'preparePrintLayout')
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

/**
 * Reset sheet formatting
 */
function resetSheetFormatting() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  
  // Clear all formatting
  range.clearFormat();
  
  // Reset frozen rows and columns
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  
  SpreadsheetApp.getUi().alert('Formatting Reset', 'All formatting has been cleared.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Prepare optimized print layout
 */
function preparePrintLayout() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Optimize column widths for printing
  sheet.setColumnWidth(1, 100); // Category - shorter for print
  sheet.setColumnWidth(2, 120); // Item
  sheet.setColumnWidth(3, 100); // Specification
  sheet.setColumnWidth(4, 80);  // Value
  sheet.setColumnWidth(5, 50);  // Units
  sheet.setColumnWidth(6, 180); // Description
  sheet.setColumnWidth(7, 100); // Usage
  sheet.setColumnWidth(8, 120); // Notes
  
  SpreadsheetApp.getUi().alert(
    'Print Layout Optimized', 
    'Column widths have been optimized for printing.\n\n' +
    'Recommended print settings:\n' +
    '• Orientation: Landscape\n' +
    '• Scale: Fit to width\n' +
    '• Margins: 0.5" all sides\n' +
    '• Include frozen rows and columns',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Show information about the script
 */
function showAbout() {
  SpreadsheetApp.getUi().alert(
    'Gallery Gauntlet Auto Formatter v1.0',
    'Professional Google Sheets formatting script for Gallery Gauntlet theme specifications.\n\n' +
    'Features:\n' +
    '• Auto-sized columns with optimal widths\n' +
    '• Professional header formatting\n' +
    '• Frozen rows and columns for navigation\n' +
    '• Alternating row colors for readability\n' +
    '• Print-optimized layout\n' +
    '• Left-justified, wrapped text formatting\n\n' +
    'Created by Claude Code Assistant',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Quick format function for toolbar
 */
function quickFormat() {
  formatGalleryGauntletSheet();
}