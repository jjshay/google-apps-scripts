/**
 * ============================================================================
 * CREATE CONFIG SHEET - RUN THIS ONCE
 * ============================================================================
 *
 * This script creates a CONFIG sheet in your spreadsheet with all variables.
 * Run this function ONCE, then use the main V3 script.
 *
 * HOW TO USE:
 * 1. Copy this entire script
 * 2. Paste into Google Apps Script (separate from main script)
 * 3. Run function: createConfigSheet()
 * 4. Check your spreadsheet for new "CONFIG" sheet
 * 5. Delete this script after running
 *
 * ============================================================================
 */

function createConfigSheet() {
  const ss = SpreadsheetApp.openById('1qRRIA7rbW_JhDg20xXV1WrZ-yx8ndnlwQrLkapr65Hc');

  // Check if CONFIG sheet exists
  let configSheet = ss.getSheetByName('CONFIG');
  if (configSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'CONFIG Sheet Exists',
      'CONFIG sheet already exists. Delete and recreate?',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      ss.deleteSheet(configSheet);
    } else {
      ui.alert('Cancelled', 'CONFIG sheet was not modified.', ui.ButtonSet.OK);
      return;
    }
  }

  // Create new CONFIG sheet
  configSheet = ss.insertSheet('CONFIG');

  // Set up formatting
  configSheet.setFrozenRows(1);
  configSheet.setColumnWidth(1, 250);
  configSheet.setColumnWidth(2, 250);
  configSheet.setColumnWidth(3, 250);
  configSheet.setColumnWidth(4, 150);
  configSheet.setColumnWidth(5, 150);

  // Define all configuration data
  const configData = [
    // HEADER ROW
    ['SETTING', 'VALUE', 'DESCRIPTION'],

    // SECTION: GOOGLE DRIVE & SHEET IDs
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['📁 GOOGLE DRIVE & SHEET IDs', '', ''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['INBOUND_FOLDER_ID', '15E4sMBfWACkv9EVUeRXj1j4uq6OoXbCu', 'Folder with artist subfolders (Shepard Fairey, Death NYC, etc.)'],
    ['IMAGES_FOLDER_ID', '1Aq6houE7SRLTwPYbrxpDg9aFZ9RcP6NJ', 'Folder with images 6-10 (generic product images)'],
    ['SHEET_ID', '1qRRIA7rbW_JhDg20xXV1WrZ-yx8ndnlwQrLkapr65Hc', 'This spreadsheet ID'],
    ['SHEET_TAB_NAME', 'PRODUCTS', 'Name of the products data sheet'],
    ['START_ROW', '2', 'First data row (row 1 = headers)'],
    ['', '', ''],

    // SECTION: API KEYS
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['🔑 API KEYS', '', ''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['CLAUDE_API_KEY', '', 'Claude API key for image analysis - Get from: https://console.anthropic.com'],
    ['OPENAI_API_KEY', '', 'OpenAI API key (optional) - Get from: https://platform.openai.com'],
    ['ENABLE_GPT4_VERIFICATION', 'FALSE', 'Enable GPT-4 verification? TRUE/FALSE'],
    ['', '', ''],

    // SECTION: ARTIST PRICING
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['💰 ARTIST PRICING', '', ''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['ARTIST_NAME', 'PRICE', 'SKU_PREFIX'],
    ['Shepard Fairey', '500', 'SF'],
    ['Death NYC', '200', 'DEATHNYC'],
    ['Death NYC Dollars', '150', 'DND'],
    ['DEFAULT', '200', 'ART'],
    ['', '', ''],

    // SECTION: ARTIST DIMENSIONS
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['📏 ARTIST DIMENSIONS', '', ''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['ARTIST_NAME', 'PORTRAIT_LENGTH', 'PORTRAIT_WIDTH', 'LANDSCAPE_LENGTH', 'LANDSCAPE_WIDTH'],
    ['Death NYC (Unframed)', '13"', '13"', '13"', '13"'],
    ['Death NYC (Framed)', '16"', '16"', '16"', '16"'],
    ['Death NYC Dollars', '6.14"', '2.61"', '6.14"', '2.61"'],
    ['Shepard Fairey', '24"', '18"', '18"', '24"'],
    ['', '', '', '', ''],

    // SECTION: PROCESSING SETTINGS
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['⚙️ PROCESSING SETTINGS', '', ''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['MAX_IMAGES_PER_RUN', '3', 'Process max N images per batch (prevents timeout)'],
    ['MAX_EXECUTION_TIME', '300000', 'Max execution time in milliseconds (5 minutes)'],
    ['AUTO_CONTINUE_DELAY_MS', '30000', 'Delay between auto-continue batches (30 seconds)'],
    ['MAX_RETRIES', '3', 'Max API retries on failure'],
    ['RETRY_DELAY_MS', '2000', 'Delay between retries (2 seconds)'],
    ['MAX_IMAGE_SIZE', '4194304', 'Max image size for Claude (4MB)'],
    ['', '', ''],

    // SECTION: DEFAULT VALUES
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['📦 DEFAULT VALUES', '', ''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['N_VALUE', '360', 'Column N value'],
    ['CATEGORY_2', 'Art', 'Column O - Category 2'],
    ['DEFAULT_QUANTITY', '1', 'Default quantity per item'],
    ['DEFAULT_CONDITION', 'New', 'Default condition'],
    ['DEFAULT_COUNTRY', 'US', 'Default country code'],
    ['DEFAULT_LOCATION', 'San Francisco, CA', 'Default location'],
    ['DEFAULT_POSTAL', '94518', 'Default postal code'],
    ['DEFAULT_PACKAGE', 'Rigid Pak', 'Default package type'],
    ['', '', ''],

    // SECTION: INSTRUCTIONS
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['📝 HOW TO USE THIS CONFIG', '', ''],
    ['━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━'],
    ['', '1. MODIFY VALUES in column B (VALUE column)', 'Do NOT change column A (SETTING names)'],
    ['', '2. ADD NEW ARTISTS in "ARTIST PRICING" section', 'Add new rows with: Artist Name | Price | SKU Prefix'],
    ['', '3. ADD NEW DIMENSIONS in "ARTIST DIMENSIONS" section', 'Add new rows with dimensions for each artist'],
    ['', '4. SAVE the sheet', 'Changes take effect on next script run'],
    ['', '5. DO NOT delete section headers or borders', 'Keep the structure intact'],
    ['', '', ''],
    ['', '⚠️ IMPORTANT: After modifying this CONFIG sheet,', 'Refresh your Google Sheet to reload the menu'],
  ];

  // Normalize all rows to 5 columns (pad with empty strings)
  const normalizedData = configData.map(row => {
    while (row.length < 5) {
      row.push('');
    }
    return row;
  });

  // Write all data to sheet (5 columns)
  configSheet.getRange(1, 1, normalizedData.length, 5).setValues(normalizedData);

  // Format header row
  configSheet.getRange(1, 1, 1, 5)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11);

  // Format section headers (rows with borders)
  for (let i = 2; i <= normalizedData.length; i++) {
    const cellValue = normalizedData[i-1][0];

    // Section separator rows
    if (cellValue && cellValue.startsWith('━━━')) {
      configSheet.getRange(i, 1, 1, 5)
        .setBackground('#E8EAED')
        .setFontWeight('bold');
    }

    // Section title rows
    if (cellValue && (cellValue.includes('📁') || cellValue.includes('🔑') ||
                      cellValue.includes('💰') || cellValue.includes('📏') ||
                      cellValue.includes('⚙️') || cellValue.includes('📦') ||
                      cellValue.includes('📝'))) {
      configSheet.getRange(i, 1, 1, 5)
        .setBackground('#F1F3F4')
        .setFontWeight('bold')
        .setFontSize(12);
    }

    // Artist pricing header row
    if (cellValue === 'ARTIST_NAME' && normalizedData[i-1][2] === 'SKU_PREFIX') {
      configSheet.getRange(i, 1, 1, 3)
        .setBackground('#D3E3FD')
        .setFontWeight('bold');
    }

    // Artist dimensions header row
    if (cellValue === 'ARTIST_NAME' && normalizedData[i-1][1] === 'PORTRAIT_LENGTH') {
      configSheet.getRange(i, 1, 1, 5)
        .setBackground('#D3E3FD')
        .setFontWeight('bold');
    }
  }

  // Add borders to data tables
  // Artist Pricing table (3 columns)
  const artistPricingStart = 20; // Adjust based on your data
  const artistPricingEnd = 23;
  configSheet.getRange(artistPricingStart, 1, artistPricingEnd - artistPricingStart + 1, 3)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // Artist Dimensions table (5 columns)
  const artistDimensionsStart = 29; // Adjust based on your data
  const artistDimensionsEnd = 32;
  configSheet.getRange(artistDimensionsStart, 1, artistDimensionsEnd - artistDimensionsStart + 1, 5)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // Success message
  SpreadsheetApp.getUi().alert(
    'CONFIG Sheet Created!',
    'CONFIG sheet has been created successfully!\n\n' +
    'You can now:\n' +
    '• Modify pricing for existing artists\n' +
    '• Add new artists with pricing\n' +
    '• Change default values\n' +
    '• Add new dimensions\n\n' +
    'Next step: Update the main script to read from this CONFIG sheet.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  Logger.log('✅ CONFIG sheet created successfully!');
}
