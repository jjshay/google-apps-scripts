/**
 * ============================================================================
 * UNIFIED 3DSELLERS SCRIPT V3 - COMPLETE WITH SKU & AUTO-CONTINUE
 * ============================================================================
 *
 * NEW IN V3:
 * 1. SKU Generation: #-ARTISTCODE-TITLE (e.g., 1-DNY-MICKEY, 2-OBEY-HOPE)
 *    - Sequential number for sorting
 *    - DNY = all Death NYC, OBEY = Shepard Fairey
 * 2. Auto-Continue: Processes all images in batches of 3 automatically
 * 3. Three Artist Folders: Shepard Fairey ($500), Death NYC ($200), Death NYC Dollars ($150)
 * 4. Column O = 28009 (Category ID 2)
 * 5. Images 6-10 in columns AE-AI (31-35) for 8 core product images
 * 6. 3Dsellers Export Functions
 *
 * FOLDER STRUCTURE:
 * INBOUND_FOLDER/
 *   Shepard Fairey/
 *   Death NYC/
 *   Death NYC Dollars/
 *
 * ============================================================================
 */

// ============================================================================
// UNIFIED CONFIGURATION
// ============================================================================

const UNIFIED_CONFIG = {
  // Google Drive and Sheet IDs
  INBOUND_FOLDER_ID: '15E4sMBfWACkv9EVUeRXj1j4uq6OoXbCu',
  IMAGES_FOLDER_ID: '',  // YOUR_FOLDER_ID_HERE - Get from Google Drive folder URL
  SHEET_ID: '',  // YOUR_SPREADSHEET_ID_HERE - Get from Google Sheets URL
  SHEET_TAB_NAME: 'PRODUCTS',
  START_ROW: 2,

  // API Keys - Get from respective services:
  // Claude: https://console.anthropic.com
  // OpenAI: https://platform.openai.com
  CLAUDE_API_KEY: '',  // YOUR_CLAUDE_KEY_HERE
  OPENAI_API_KEY: '',  // YOUR_OPENAI_KEY_HERE

  // Processing limits
  MAX_IMAGES_PER_RUN: 3, // Process max 3 images per run
  MAX_EXECUTION_TIME: 300000, // 5 minutes
  AUTO_CONTINUE_DELAY_MS: 30000, // 30 seconds between batches

  // Artist-based pricing
  ARTIST_PRICING: {
    'Shepard Fairey': 500,
    'Death NYC': 200,
    'Death NYC Dollars': 150,
    'DEFAULT': 200
  },

  // Artist abbreviations for SKU
  ARTIST_ABBREV: {
    'Shepard Fairey': 'OBEY',
    'Death NYC': 'DNY',
    'Death NYC Dollars': 'DNY',
    'DEFAULT': 'ART'
  },

  // Error handling
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 2000,
  MAX_IMAGE_SIZE: 4 * 1024 * 1024,

  // GPT-4 settings
  ENABLE_GPT4_VERIFICATION: false,

  // Column mapping
  COLUMNS: {
    SKU: 1, // Column A - NEW!
    ARTIST: 2,
    IMAGE_LINK: 3,
    ORIENTATION: 4,
    TITLE: 5,
    DESCRIPTION: 6,
    TAGS: 7,
    META_KEYWORDS: 8,
    META_DESCRIPTION: 9,
    MOBILE_DESC: 10,
    TYPE: 11,
    LENGTH: 12,
    WIDTH: 13,
    COL_N: 14,
    CATEGORY_2: 15,
    QUANTITY: 16,
    PRICE: 17,
    CONDITION: 18,
    COUNTRY_CODE: 19,
    LOCATION: 20,
    POSTAL_CODE: 21,
    PACKAGE_TYPE: 22,
    IMAGE_6: 31, // AE
    IMAGE_7: 32, // AF
    IMAGE_8: 33, // AG
    IMAGE_9: 34, // AH
    IMAGE_10: 35 // AI
  },

  // Default values
  DEFAULTS: {
    N_VALUE: '360',
    CATEGORY_2: '28009',
    QUANTITY: 1,
    CONDITION: 'New',
    COUNTRY: 'US',
    LOCATION: 'San Francisco, CA',
    POSTAL: '94518',
    PACKAGE: 'Rigid Pak'
  }
};

// ============================================================================
// CUSTOM MENU
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('3DSellers')
    // STEP 1: Process Images
    .addSubMenu(ui.createMenu('📸 Process Images')
      .addItem('✨ Process ALL Images (Auto-Continue)', 'runDescriptionsAutoContinue')
      .addItem('📝 Process Single Batch (3 images)', 'runDescriptions')
      .addItem('🛑 Stop Auto-Continue', 'stopAutoContinue'))

    // STEP 2: Fill Details
    .addItem('📋 Fill Product Details', 'runDetails')

    .addSeparator()

    // STEP 3: Export Functions
    .addSubMenu(ui.createMenu('📤 Export to 3DSellers')
      .addItem('1️⃣ Simple Export (1:1)', 'exportTo3DSellers')
      .addItem('2️⃣ Export with Variations', 'exportTo3DSellersWithVariations')
      .addItem('3️⃣ Generate 4 Variants (Max Inventory)', 'generate3DSellers4Variants'))

    .addSeparator()

    // Test Functions
    .addSubMenu(ui.createMenu('🧪 Test Functions')
      .addItem('Test Descriptions (1 image)', 'testDescriptions')
      .addItem('Test Details (1 row)', 'testDetails')
      .addItem('Test Export (Simple)', 'testExportSimple'))

    .addToUi();
}

// ============================================================================
// SKU GENERATION
// ============================================================================

/**
 * Generate SKU: #-ARTISTCODE-TITLE
 * Example: 1-DNY-MICKEY, 2-OBEY-HOPE, 3-DNY-SPRAYCAN
 * Sequential number allows sorting by order added
 */
function generateSKU(artist, title, sheet) {
  try {
    // Get artist abbreviation
    const artistCode = UNIFIED_CONFIG.ARTIST_ABBREV[artist] || UNIFIED_CONFIG.ARTIST_ABBREV['DEFAULT'];

    // Extract subject from title (remove common words)
    const titleClean = title
      .replace(/FINE ART/gi, '')
      .replace(/Limited Edition Print/gi, '')
      .replace(/Street Art/gi, '')
      .replace(/1 OF A KIND/gi, '')
      .replace(/DOLLAR ART/gi, '')
      .replace(/Death NYC/gi, '')
      .replace(/Shepard Fairey/gi, '')
      .replace(/\s+/g, ' ')
      .trim();

    // Get first meaningful words from title (up to 20 chars max)
    const titleParts = titleClean.split(' ').filter(word => word.length > 2);
    let titleCode = titleParts.slice(0, 3).join('').toUpperCase().replace(/[^A-Z0-9]/g, '').substring(0, 20);

    if (!titleCode) titleCode = 'ART';

    // Find next sequential number (global across ALL products)
    let sequentialNumber = 1;

    // Check existing SKUs in the sheet to find highest number
    const lastRow = sheet.getLastRow();
    if (lastRow >= UNIFIED_CONFIG.START_ROW) {
      const skuColumn = sheet.getRange(UNIFIED_CONFIG.START_ROW, UNIFIED_CONFIG.COLUMNS.SKU, lastRow - UNIFIED_CONFIG.START_ROW + 1, 1).getValues();

      const existingNumbers = skuColumn
        .map(row => row[0])
        .filter(sku => sku && sku.toString().trim() !== '')
        .map(sku => {
          // Extract leading number from format: #-ARTISTCODE-TITLE
          const match = sku.toString().match(/^(\d+)-/);
          return match ? parseInt(match[1]) : 0;
        });

      if (existingNumbers.length > 0) {
        sequentialNumber = Math.max(...existingNumbers) + 1;
      }
    }

    return `${sequentialNumber}-${artistCode}-${titleCode}`;

  } catch (error) {
    Logger.log(`⚠️ SKU generation error: ${error.message}`);
    return `999-ART-${new Date().getTime()}`;
  }
}

// ============================================================================
// AUTO-CONTINUE PROCESSING
// ============================================================================

/**
 * Menu button: Process All New Images with Auto-Continue
 */
function runDescriptionsAutoContinue() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '3DSELLERS AUTO-CONTINUE',
    `Process ALL new images automatically?\n\n` +
    `• Processes ${UNIFIED_CONFIG.MAX_IMAGES_PER_RUN} images at a time\n` +
    `• Automatically continues until all done\n` +
    `• 30-second pause between batches\n` +
    `• Artist detection from subfolder\n` +
    `• SKU auto-generation\n\n` +
    'Folders:\n' +
    '• Shepard Fairey ($500)\n' +
    '• Death NYC ($200)\n' +
    '• Death NYC Dollars ($150)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // Start the auto-continue process
  checkInboundFolderWithContinue();
}

/**
 * Check and process with auto-continue
 */
function checkInboundFolderWithContinue() {
  const result = checkInboundFolder();

  // If more images remain, schedule next batch
  if (result.remaining > 0) {
    Logger.log(`\n⏰ Scheduling next batch in 30 seconds...`);
    Logger.log(`📊 Progress: ${result.successful} processed, ${result.remaining} remaining`);

    // Create time-based trigger for next batch
    ScriptApp.newTrigger('checkInboundFolderWithContinue')
      .timeBased()
      .after(UNIFIED_CONFIG.AUTO_CONTINUE_DELAY_MS)
      .create();

    SpreadsheetApp.getUi().alert(
      'Batch Complete',
      `Processed ${result.successful}/${result.attempted} images.\n\n` +
      `${result.remaining} images remaining.\n` +
      `Next batch will start automatically in 30 seconds.\n\n` +
      `Click "Stop Auto-Continue" to cancel.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    Logger.log(`\n🎉 ALL IMAGES PROCESSED!`);
    cleanupTriggers(); // Remove any remaining triggers

    SpreadsheetApp.getUi().alert(
      'All Complete!',
      `Successfully processed all images!\n\n` +
      `Total: ${result.successful} images\n\n` +
      `Next step: Click "Fill Product Details"`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Stop auto-continue by removing triggers
 */
function stopAutoContinue() {
  cleanupTriggers();
  SpreadsheetApp.getUi().alert(
    'Stopped',
    'Auto-continue has been stopped.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Clean up all triggers for this function
 */
function cleanupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkInboundFolderWithContinue') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ============================================================================
// ARTIST DETECTION
// ============================================================================

/**
 * Get artist name from parent folder
 */
function getArtistFromFolder(file) {
  try {
    const parents = file.getParents();
    if (parents.hasNext()) {
      const parentFolder = parents.next();
      const folderName = parentFolder.getName();

      // Check against known artists
      if (UNIFIED_CONFIG.ARTIST_PRICING[folderName]) {
        return folderName;
      }

      return folderName;
    }
  } catch (error) {
    Logger.log(`  ⚠️ Could not get folder name: ${error.message}`);
  }

  return 'Unknown Artist';
}

/**
 * Get price based on artist
 */
function getPriceForArtist(artist) {
  return UNIFIED_CONFIG.ARTIST_PRICING[artist] || UNIFIED_CONFIG.ARTIST_PRICING['DEFAULT'];
}

// ============================================================================
// IMAGE PROCESSING
// ============================================================================

/**
 * Check INBOUND folder and process new images from subfolders
 */
function checkInboundFolder() {
  const startTime = new Date().getTime();
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('STEP 1: 3DSELLERS DESCRIPTIONS');
  Logger.log(`Max images per run: ${UNIFIED_CONFIG.MAX_IMAGES_PER_RUN}`);
  Logger.log(`Artist detection: FROM SUBFOLDER NAME`);
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  try {
    const rootFolder = DriveApp.getFolderById(UNIFIED_CONFIG.INBOUND_FOLDER_ID);
    const sheet = SpreadsheetApp.openById(UNIFIED_CONFIG.SHEET_ID).getSheetByName(UNIFIED_CONFIG.SHEET_TAB_NAME);

    if (!sheet) {
      Logger.log(`❌ Sheet "${UNIFIED_CONFIG.SHEET_TAB_NAME}" not found`);
      return { successful: 0, attempted: 0, remaining: 0 };
    }

    // Get already processed images
    const lastRow = sheet.getLastRow();
    const processedUrls = new Set();

    if (lastRow >= UNIFIED_CONFIG.START_ROW) {
      const urlColumn = sheet.getRange(UNIFIED_CONFIG.START_ROW, UNIFIED_CONFIG.COLUMNS.IMAGE_LINK, lastRow - UNIFIED_CONFIG.START_ROW + 1, 1);
      const formulas = urlColumn.getFormulas();

      formulas.forEach(row => {
        if (row[0]) {
          const match = row[0].match(/id=([^"]+)/);
          if (match) processedUrls.add(match[1]);
        }
      });
    }

    Logger.log(`📊 Already processed: ${processedUrls.size} images`);

    // Collect new images from root and subfolders
    const newImages = [];

    // Check root folder
    const rootFiles = rootFolder.getFiles();
    while (rootFiles.hasNext()) {
      const file = rootFiles.next();
      if (!file.getMimeType().startsWith('image/')) continue;
      if (processedUrls.has(file.getId())) continue;
      newImages.push(file);
    }

    // Check subfolders (artist folders)
    const subfolders = rootFolder.getFolders();
    while (subfolders.hasNext()) {
      const subfolder = subfolders.next();
      const folderName = subfolder.getName();
      Logger.log(`📁 Checking subfolder: ${folderName}`);

      const subfolderFiles = subfolder.getFiles();
      let folderCount = 0;
      while (subfolderFiles.hasNext()) {
        const file = subfolderFiles.next();
        if (!file.getMimeType().startsWith('image/')) continue;
        if (processedUrls.has(file.getId())) continue;
        newImages.push(file);
        folderCount++;
      }
      Logger.log(`  → Found ${folderCount} new images in ${folderName}`);
    }

    Logger.log(`\n📊 Total new images found: ${newImages.length}`);

    // Process limited batch
    let successCount = 0;
    let attemptCount = 0;
    const imagesToProcess = Math.min(newImages.length, UNIFIED_CONFIG.MAX_IMAGES_PER_RUN);

    for (let i = 0; i < imagesToProcess; i++) {
      // Check execution time
      const elapsed = new Date().getTime() - startTime;
      if (elapsed > UNIFIED_CONFIG.MAX_EXECUTION_TIME) {
        Logger.log(`⏱️ Approaching timeout limit, stopping at ${i} images`);
        break;
      }

      const file = newImages[i];
      const artist = getArtistFromFolder(file);
      Logger.log(`\n🆕 [${i + 1}/${imagesToProcess}] ${file.getName()}`);
      Logger.log(`  🎨 Artist: ${artist}`);

      const success = processNewImage(file, sheet, artist);
      attemptCount++;
      if (success) successCount++;

      Utilities.sleep(2000);
    }

    const remaining = newImages.length - attemptCount;

    Logger.log(`\n✅ BATCH COMPLETE: ${successCount}/${attemptCount} successful`);
    if (remaining > 0) {
      Logger.log(`📊 ${remaining} images remaining`);
    }

    return { successful: successCount, attempted: attemptCount, remaining: remaining };

  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    return { successful: 0, attempted: 0, remaining: 0 };
  }
}

/**
 * Process single image with AI analysis
 */
function processNewImage(file, sheet, artist) {
  try {
    Logger.log(`\n📸 Processing: ${file.getName()}`);

    const downloadUrl = `https://drive.google.com/uc?export=view&id=${file.getId()}`;

    // Step 1: Claude Vision
    Logger.log('  🤖 Claude Vision analysis...');
    const analysis = analyzeImageWithClaude(file, artist);
    if (!analysis) {
      Logger.log('  ❌ Claude analysis failed');
      return false;
    }
    Logger.log(`  ✓ Orientation: ${analysis.orientation}, Framed: ${analysis.framed}`);
    if (analysis.substrate === 'Dollar Bill Art') {
      Logger.log(`  💵 DOLLAR BILL ART DETECTED!`);
    }

    // Step 2: Generate listing
    Logger.log('  📝 Generating listing...');
    const listing = generateEbayListing(analysis, file.getName(), artist);

    // Step 3: Validate
    Logger.log('  🔍 Validating listing...');
    const validation = validateListingDetailed(listing);
    if (!validation.valid) {
      Logger.log(`  ❌ Validation failed: ${validation.errors.join(', ')}`);
      return false;
    }
    Logger.log('  ✓ Validation passed');

    // Step 4: Generate SKU
    Logger.log('  🔖 Generating SKU...');
    const sku = generateSKU(artist, listing.title, sheet);
    Logger.log(`  ✓ SKU: ${sku}`);

    // Step 5: Write to sheet (A-M)
    const nextRow = sheet.getLastRow() + 1;

    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.SKU).setValue(sku);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.ARTIST).setValue(artist);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.IMAGE_LINK).setFormula(`=IMAGE("${downloadUrl}", 1)`);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.ORIENTATION).setValue(listing.orientation);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.TITLE).setValue(listing.title);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.DESCRIPTION).setValue(listing.description);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.TAGS).setValue(listing.tags);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.META_KEYWORDS).setValue(listing.metaKeywords);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.META_DESCRIPTION).setValue(listing.metaDescription);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.MOBILE_DESC).setValue(listing.mobileDescription);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.TYPE).setValue(listing.type);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.LENGTH).setValue(listing.length);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.WIDTH).setValue(listing.width);

    sheet.setRowHeight(nextRow, 100);

    Logger.log(`  ✅ Added to row ${nextRow}`);
    return true;

  } catch (error) {
    Logger.log(`  ❌ Error: ${error.message}`);
    Logger.log(`  Stack: ${error.stack}`);
    return false;
  }
}

// ============================================================================
// AI ANALYSIS FUNCTIONS
// ============================================================================

/**
 * Analyze image with Claude Vision
 */
function analyzeImageWithClaude(file, artist) {
  try {
    const blob = getResizedImage(file);
    if (!blob) {
      Logger.log('  ❌ Could not resize image');
      return null;
    }

    const base64Image = Utilities.base64Encode(blob.getBytes());

    const prompt = `Analyze this ${artist} artwork:
1. Orientation: "portrait" or "landscape"
2. Is it framed? "Framed" or "Not Framed"
3. Is this printed on a DOLLAR BILL? "Dollar Bill Art" or "Standard Print"
4. Subject/theme of the artwork
5. Color palette (3-5 colors)
6. Art style
7. Key elements
8. Emotional tone

JSON format:
{
  "orientation": "portrait" or "landscape",
  "framed": "Framed" or "Not Framed",
  "substrate": "Dollar Bill Art" or "Standard Print",
  "subject": "description",
  "colors": ["color1", "color2"],
  "style": "style",
  "elements": ["element1", "element2"],
  "tone": "tone"
}`;

    const payload = {
      model: 'claude-3-opus-20240229',
      max_tokens: 2000,
      messages: [{
        role: 'user',
        content: [
          {
            type: 'image',
            source: {
              type: 'base64',
              media_type: blob.getContentType(),
              data: base64Image
            }
          },
          { type: 'text', text: prompt }
        ]
      }]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': UNIFIED_CONFIG.CLAUDE_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log(`  ❌ Claude API error: ${responseCode}`);
      return null;
    }

    const data = JSON.parse(response.getContentText());
    const analysisText = data.content[0].text;
    const jsonMatch = analysisText.match(/\{[\s\S]*\}/);

    if (!jsonMatch) {
      Logger.log(`  ❌ Could not find JSON in Claude response`);
      return null;
    }

    return JSON.parse(jsonMatch[0]);

  } catch (error) {
    Logger.log(`  ❌ Claude error: ${error.message}`);
    return null;
  }
}

/**
 * Resize image to fit under 5MB
 */
function getResizedImage(file) {
  try {
    const fileId = file.getId();
    const thumbnailSizes = [1600, 1200, 800];

    for (const size of thumbnailSizes) {
      try {
        const thumbnailUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w${size}`;
        const response = UrlFetchApp.fetch(thumbnailUrl, {
          headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
          muteHttpExceptions: true
        });

        if (response.getResponseCode() === 200) {
          const blob = response.getBlob();
          if (blob.getBytes().length < UNIFIED_CONFIG.MAX_IMAGE_SIZE) {
            return blob;
          }
        }
      } catch (e) {
        continue;
      }
    }

    return file.getBlob();
  } catch (error) {
    return null;
  }
}

/**
 * Generate eBay listing with artist-specific details
 */
function generateEbayListing(analysis, fileName, artist) {
  const subject = analysis.subject || 'Contemporary Art';
  const style = analysis.style || 'Pop Art';
  const isDollarBillArt = analysis.substrate === 'Dollar Bill Art';

  let title;
  if (isDollarBillArt) {
    title = `FINE ART ${subject} ${artist} 1 OF A KIND DOLLAR ART Street Art`.substring(0, 80);
  } else {
    title = `FINE ART ${subject} ${artist} Limited Edition Print`.substring(0, 80);
  }

  const description = `THE ARTWORK:
${analysis.subject ? 'This piece showcases ' + analysis.subject + '. ' : ''}${analysis.elements ? analysis.elements.join(', ') + '. ' : ''}The color palette features ${analysis.colors ? analysis.colors.join(', ') : 'vibrant colors'}.

THE ARTIST:
${artist} is a legendary contemporary urban artist, creating bold mashups of luxury brands, cartoon icons, and celebrity imagery.

CONDITION: High-quality art${isDollarBillArt ? ' on genuine dollar bill' : ' print'}. What you see is what you get.
SHIPPING: Gallery-level care. Protected with archival materials.
GUARANTEE: 30-day money-back guarantee.`;

  const tags = isDollarBillArt
    ? `${artist}, dollar bill art, currency art, money art, ${style.toLowerCase()}, street art, urban art, wall decor`
    : `${artist}, artwork, art, ${style.toLowerCase()}, wall decor`;

  const type = analysis.framed || 'Not Framed';
  const substrate = analysis.substrate || 'Standard Print';

  const length = calculateLength(artist, analysis.orientation, type, substrate);
  const width = calculateWidth(artist, analysis.orientation, type, substrate, length);

  return {
    title,
    description,
    tags,
    orientation: analysis.orientation,
    metaKeywords: tags,
    metaDescription: `${title} - ${subject}. Contemporary street art by ${artist}.`,
    mobileDescription: `FINE ART ${style} - ${subject}. ${artist} limited edition.`,
    type,
    length,
    width,
    substrate: substrate
  };
}

/**
 * Calculate length based on artwork type
 */
function calculateLength(artist, orientation, type, substrate) {
  if (substrate === 'Dollar Bill Art') {
    return '6.14"';
  }

  if (artist === 'Death NYC' || artist === 'Death NYC Dollars') {
    if (orientation === 'landscape' && type === 'Framed') return '16"';
    if (type === 'Not Framed') return '13"';
    return '16"';
  }

  if (artist === 'Shepard Fairey') {
    if (orientation === 'portrait') return '24"';
    return '18"';
  }

  return '13"';
}

/**
 * Calculate width based on artwork type
 */
function calculateWidth(artist, orientation, type, substrate, length) {
  if (substrate === 'Dollar Bill Art') {
    return '2.61"';
  }

  if (artist === 'Death NYC' || artist === 'Death NYC Dollars') {
    return length;
  }

  if (artist === 'Shepard Fairey') {
    if (orientation === 'portrait') return '18"';
    return '24"';
  }

  return length;
}

/**
 * Validate listing with detailed error reporting
 */
function validateListingDetailed(listing) {
  const errors = [];

  if (!listing) {
    return { valid: false, errors: ['Listing is null or undefined'] };
  }

  if (!listing.title) {
    errors.push('Title is missing');
  } else if (!listing.title.startsWith('FINE ART')) {
    errors.push(`Title does not start with "FINE ART"`);
  } else if (listing.title.length > 80) {
    errors.push(`Title too long: ${listing.title.length} chars`);
  }

  if (!listing.orientation) errors.push('Orientation is missing');
  if (!listing.type) errors.push('Type is missing');
  if (!listing.description) errors.push('Description is missing');

  return {
    valid: errors.length === 0,
    errors: errors
  };
}

// ============================================================================
// STEP 2: FILL PRODUCT DETAILS
// ============================================================================

/**
 * Menu button: Fill Product Details
 */
function runDetails() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '3DSELLERS DETAILS',
    'Fill product details for rows with descriptions?\n\n' +
    'This will fill:\n' +
    '• Column O (Category 2)\n' +
    '• Column Q (Price - artist-based)\n' +
    '• Columns M-V (Width, Condition, etc.)\n' +
    '• Columns AB-AF (Images 6-10)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  ui.alert('Processing...', 'Filling product details...', ui.ButtonSet.OK);

  const count = fillProductDetails();

  ui.alert('Complete!', `Filled details for ${count} rows.`, ui.ButtonSet.OK);
}

/**
 * Fill columns M onwards for rows with descriptions
 */
function fillProductDetails() {
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('STEP 2: 3DSELLERS DETAILS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  try {
    const sheet = SpreadsheetApp.openById(UNIFIED_CONFIG.SHEET_ID).getSheetByName(UNIFIED_CONFIG.SHEET_TAB_NAME);
    if (!sheet) return 0;

    Logger.log('📁 Getting images from Google Drive...');
    const imageUrls = getImageUrlsFromFolder();
    Logger.log(`✓ Found ${imageUrls.length} images`);

    const lastRow = sheet.getLastRow();
    if (lastRow < UNIFIED_CONFIG.START_ROW) return 0;

    const artistColumn = sheet.getRange(UNIFIED_CONFIG.START_ROW, UNIFIED_CONFIG.COLUMNS.ARTIST, lastRow - UNIFIED_CONFIG.START_ROW + 1, 1).getValues();
    const typeColumn = sheet.getRange(UNIFIED_CONFIG.START_ROW, UNIFIED_CONFIG.COLUMNS.TYPE, lastRow - UNIFIED_CONFIG.START_ROW + 1, 1).getValues();
    const lengthColumn = sheet.getRange(UNIFIED_CONFIG.START_ROW, UNIFIED_CONFIG.COLUMNS.LENGTH, lastRow - UNIFIED_CONFIG.START_ROW + 1, 1).getValues();

    let processedCount = 0;

    for (let i = 0; i < typeColumn.length; i++) {
      const rowNum = UNIFIED_CONFIG.START_ROW + i;
      const artist = artistColumn[i][0];
      const type = typeColumn[i][0];
      const length = lengthColumn[i][0];

      if (!type || type === '') continue;

      Logger.log(`\n📝 Row ${rowNum} - Artist: ${artist}`);

      const existingWidth = sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.WIDTH).getValue();
      const width = existingWidth || length;

      const price = getPriceForArtist(artist);
      Logger.log(`  💰 Price: $${price}`);

      const nValue = UNIFIED_CONFIG.DEFAULTS.N_VALUE;
      const category2 = UNIFIED_CONFIG.DEFAULTS.CATEGORY_2;
      const quantity = UNIFIED_CONFIG.DEFAULTS.QUANTITY;
      const condition = UNIFIED_CONFIG.DEFAULTS.CONDITION;
      const country = UNIFIED_CONFIG.DEFAULTS.COUNTRY;
      const location = UNIFIED_CONFIG.DEFAULTS.LOCATION;
      const postal = UNIFIED_CONFIG.DEFAULTS.POSTAL;
      const packageType = UNIFIED_CONFIG.DEFAULTS.PACKAGE;

      if (!existingWidth) {
        sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.WIDTH).setValue(width);
      }
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.COL_N).setValue(nValue);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.CATEGORY_2).setValue(category2);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.QUANTITY).setValue(quantity);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.PRICE).setValue(price);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.CONDITION).setValue(condition);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.COUNTRY_CODE).setValue(country);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.LOCATION).setValue(location);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.POSTAL_CODE).setValue(postal);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.PACKAGE_TYPE).setValue(packageType);

      if (imageUrls.length > 0) {
        if (imageUrls[0]) sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.IMAGE_6).setValue(imageUrls[0]);
        if (imageUrls[1]) sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.IMAGE_7).setValue(imageUrls[1]);
        if (imageUrls[2]) sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.IMAGE_8).setValue(imageUrls[2]);
        if (imageUrls[3]) sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.IMAGE_9).setValue(imageUrls[3]);
        if (imageUrls[4]) sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.IMAGE_10).setValue(imageUrls[4]);
      }

      Logger.log(`  ✓ Filled details and images`);
      processedCount++;
    }

    Logger.log(`\n✅ STEP 2 COMPLETE: Processed ${processedCount} rows`);
    return processedCount;

  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    return 0;
  }
}

/**
 * Get images from Google Drive folder
 */
function getImageUrlsFromFolder() {
  try {
    const folder = DriveApp.getFolderById(UNIFIED_CONFIG.IMAGES_FOLDER_ID);
    const files = folder.getFiles();
    const imageUrls = [];

    while (files.hasNext() && imageUrls.length < 5) {
      const file = files.next();
      if (file.getMimeType().startsWith('image/')) {
        imageUrls.push(`https://drive.google.com/uc?export=view&id=${file.getId()}`);
      }
    }

    return imageUrls;
  } catch (error) {
    Logger.log(`⚠️ Error accessing images: ${error.message}`);
    return [];
  }
}

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

/**
 * Test descriptions with one image
 */
function testDescriptions() {
  Logger.log('🧪 Testing DESCRIPTIONS with one image...');

  const rootFolder = DriveApp.getFolderById(UNIFIED_CONFIG.INBOUND_FOLDER_ID);
  let testFile = null;

  const rootFiles = rootFolder.getFiles();
  if (rootFiles.hasNext()) {
    testFile = rootFiles.next();
  }

  if (!testFile) {
    const subfolders = rootFolder.getFolders();
    while (subfolders.hasNext() && !testFile) {
      const subfolder = subfolders.next();
      const subfolderFiles = subfolder.getFiles();
      if (subfolderFiles.hasNext()) {
        testFile = subfolderFiles.next();
      }
    }
  }

  if (testFile) {
    const sheet = SpreadsheetApp.openById(UNIFIED_CONFIG.SHEET_ID).getSheetByName(UNIFIED_CONFIG.SHEET_TAB_NAME);
    const artist = getArtistFromFolder(testFile);
    Logger.log(`Testing with: ${testFile.getName()} (Artist: ${artist})`);
    const success = processNewImage(testFile, sheet, artist);
    Logger.log(success ? '✅ Test successful' : '❌ Test failed');
  } else {
    Logger.log('❌ No images found');
  }
}

/**
 * Test details with first row
 */
function testDetails() {
  Logger.log('🧪 Testing DETAILS with first row...');

  const sheet = SpreadsheetApp.openById(UNIFIED_CONFIG.SHEET_ID).getSheetByName(UNIFIED_CONFIG.SHEET_TAB_NAME);
  const imageUrls = getImageUrlsFromFolder();

  const rowNum = UNIFIED_CONFIG.START_ROW;
  const artist = sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.ARTIST).getValue();
  const type = sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.TYPE).getValue();
  const length = sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.LENGTH).getValue();

  if (!type) {
    Logger.log('❌ Row is empty');
    return;
  }

  const price = getPriceForArtist(artist);
  Logger.log(`Artist: ${artist}, Price: $${price}`);

  const width = length;
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.WIDTH).setValue(width);
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.COL_N).setValue(UNIFIED_CONFIG.DEFAULTS.N_VALUE);
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.CATEGORY_2).setValue(UNIFIED_CONFIG.DEFAULTS.CATEGORY_2);
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.QUANTITY).setValue(UNIFIED_CONFIG.DEFAULTS.QUANTITY);
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.PRICE).setValue(price);
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.CONDITION).setValue(UNIFIED_CONFIG.DEFAULTS.CONDITION);
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.COUNTRY_CODE).setValue(UNIFIED_CONFIG.DEFAULTS.COUNTRY);
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.LOCATION).setValue(UNIFIED_CONFIG.DEFAULTS.LOCATION);
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.POSTAL_CODE).setValue(UNIFIED_CONFIG.DEFAULTS.POSTAL);
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.PACKAGE_TYPE).setValue(UNIFIED_CONFIG.DEFAULTS.PACKAGE);

  if (imageUrls.length > 0) {
    if (imageUrls[0]) sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.IMAGE_6).setValue(imageUrls[0]);
    if (imageUrls[1]) sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.IMAGE_7).setValue(imageUrls[1]);
    if (imageUrls[2]) sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.IMAGE_8).setValue(imageUrls[2]);
    if (imageUrls[3]) sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.IMAGE_9).setValue(imageUrls[3]);
    if (imageUrls[4]) sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.IMAGE_10).setValue(imageUrls[4]);
  }

  Logger.log(`✅ Test complete for row ${rowNum}`);
}

// ============================================================================
// SINGLE BATCH PROCESSING (NO AUTO-CONTINUE)
// ============================================================================

/**
 * Process single batch of images (no auto-continue)
 */
function runDescriptions() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '3DSELLERS DESCRIPTIONS - SINGLE BATCH',
    `Process up to ${UNIFIED_CONFIG.MAX_IMAGES_PER_RUN} new images?\n\n` +
    `This will NOT auto-continue.\n` +
    `Use "Process ALL Images" for automatic processing.\n\n` +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const result = checkInboundFolder();

  ui.alert('Batch Complete!',
    `Processed ${result.successful}/${result.attempted} images.\n\n` +
    (result.remaining > 0 ? `${result.remaining} images remaining.\n\n` : '') +
    'Run again to process more, or use "Process ALL Images" for auto-continue.',
    ui.ButtonSet.OK);
}

// ============================================================================
// 3DSELLERS EXPORT FUNCTIONS
// ============================================================================

/**
 * Export Function 1: Simple 1:1 Export
 */
function exportTo3DSellers() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'EXPORT TO 3DSELLERS - SIMPLE',
    'Export PRODUCTS data 1:1 to 3DSellers format?\n\n' +
    '• No variations\n' +
    '• One row per product\n' +
    '• Creates sheet: 3DSELLERS_EXPORT\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const ss = SpreadsheetApp.openById(UNIFIED_CONFIG.SHEET_ID);
    const sourceSheet = ss.getSheetByName(UNIFIED_CONFIG.SHEET_TAB_NAME);

    if (!sourceSheet) {
      ui.alert('Error', 'PRODUCTS sheet not found', ui.ButtonSet.OK);
      return;
    }

    // Create or clear export sheet
    let exportSheet = ss.getSheetByName('3DSELLERS_EXPORT');
    if (exportSheet) {
      exportSheet.clear();
    } else {
      exportSheet = ss.insertSheet('3DSELLERS_EXPORT');
    }

    // Set headers
    const headers = ['SKU', 'ParentSKU', 'Title', 'Description', 'Brand', 'Manufacturer',
                     'Style', 'Material', 'Size', 'Quantity', 'Price', 'Condition',
                     'Country', 'Location', 'Postal', 'Package', 'Image1', 'Image2',
                     'Image3', 'Image4', 'Image5'];
    exportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Get source data
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < UNIFIED_CONFIG.START_ROW) {
      ui.alert('Complete', 'No data to export', ui.ButtonSet.OK);
      return;
    }

    const data = sourceSheet.getRange(UNIFIED_CONFIG.START_ROW, 1, lastRow - UNIFIED_CONFIG.START_ROW + 1, 32).getValues();
    const formulas = sourceSheet.getRange(UNIFIED_CONFIG.START_ROW, UNIFIED_CONFIG.COLUMNS.IMAGE_LINK, lastRow - UNIFIED_CONFIG.START_ROW + 1, 1).getFormulas();

    let exportRow = 2;

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const sku = row[0]; // Column A
      const artist = row[1]; // Column B
      const orientation = row[3]; // Column D
      const title = row[4]; // Column E
      const description = row[5]; // Column F
      const type = row[10]; // Column K
      const length = row[11]; // Column L
      const price = row[16]; // Column Q

      if (!title || !sku) continue;

      // Extract image URL from formula
      let imageUrl = '';
      if (formulas[i][0]) {
        const match = formulas[i][0].match(/https:\/\/drive\.google\.com\/uc\?export=view&id=([^"]+)/);
        if (match) {
          imageUrl = match[0];
        }
      }

      const exportData = [
        sku,                          // SKU
        '',                           // ParentSKU (empty for simple export)
        title,                        // Title
        description || '',            // Description
        artist || '',                 // Brand
        artist || '',                 // Manufacturer
        orientation || '',            // Style
        type || '',                   // Material
        length || '',                 // Size
        1,                            // Quantity
        price || 0,                   // Price
        'New',                        // Condition
        'US',                         // Country
        'San Francisco, CA',          // Location
        '94518',                      // Postal
        'Rigid Pak',                  // Package
        imageUrl,                     // Image1
        '', '', '', ''                // Image2-5
      ];

      exportSheet.getRange(exportRow, 1, 1, exportData.length).setValues([exportData]);
      exportRow++;
    }

    ui.alert('Export Complete!',
      `Exported ${exportRow - 2} products to sheet: 3DSELLERS_EXPORT\n\n` +
      'Download: File > Download > CSV',
      ui.ButtonSet.OK);

  } catch (error) {
    Logger.log(`Export error: ${error.message}`);
    ui.alert('Error', `Export failed: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Export Function 2: Export with Variations (Smart Grouping)
 */
function exportTo3DSellersWithVariations() {
  const ui = SpreadsheetApp.getUi();

  ui.alert('Coming Soon',
    'Export with Variations feature is under development.\n\n' +
    'For now, use "Simple Export" or "Generate 4 Variants".',
    ui.ButtonSet.OK);
}

/**
 * Export Function 3: Generate 4 Variants (Maximum Inventory)
 */
function generate3DSellers4Variants() {
  const ui = SpreadsheetApp.getUi();

  ui.alert('Coming Soon',
    'Generate 4 Variants feature is under development.\n\n' +
    'This will create:\n' +
    '• Portrait + Unframed\n' +
    '• Landscape + Unframed\n' +
    '• Portrait + Framed\n' +
    '• Landscape + Framed\n\n' +
    'For now, use "Simple Export".',
    ui.ButtonSet.OK);
}

/**
 * Test export function
 */
function testExportSimple() {
  Logger.log('🧪 Testing Simple Export...');

  const ss = SpreadsheetApp.openById(UNIFIED_CONFIG.SHEET_ID);
  const sourceSheet = ss.getSheetByName(UNIFIED_CONFIG.SHEET_TAB_NAME);

  if (!sourceSheet) {
    Logger.log('❌ PRODUCTS sheet not found');
    return;
  }

  const lastRow = sourceSheet.getLastRow();
  Logger.log(`Found ${lastRow - 1} rows in PRODUCTS sheet`);

  if (lastRow >= UNIFIED_CONFIG.START_ROW) {
    const firstRow = sourceSheet.getRange(UNIFIED_CONFIG.START_ROW, 1, 1, 20).getValues()[0];
    Logger.log(`First row SKU: ${firstRow[0]}`);
    Logger.log(`First row Artist: ${firstRow[1]}`);
    Logger.log(`First row Title: ${firstRow[4]}`);
    Logger.log('✅ Test complete - data looks good');
  } else {
    Logger.log('❌ No data to export');
  }
}
