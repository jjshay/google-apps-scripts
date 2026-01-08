/**
 * ============================================================================
 * FINAL UNIFIED 3DSELLERS SCRIPT - NO CONFLICTS
 * ============================================================================
 *
 * This combines BOTH workflows into ONE script:
 *   1. 3DSELLERS DESCRIPTIONS (Check for new images)
 *   2. 3DSELLERS DETAILS (Fill product details)
 *
 * CUSTOM MENU will appear in your Google Sheet after installation.
 *
 * ============================================================================
 */

// ============================================================================
// SINGLE UNIFIED CONFIGURATION
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

  // Error handling
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 2000,
  MAX_IMAGE_SIZE: 4 * 1024 * 1024,

  // Column mapping - ALL columns in one place
  COLUMNS: {
    // STEP 1: Descriptions columns (B-L)
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

    // STEP 2: Details columns (M-V, AB-AF)
    WIDTH: 13,
    COL_N: 14,
    QUANTITY: 16,
    PRICE: 17,
    CONDITION: 18,
    COUNTRY_CODE: 19,
    LOCATION: 20,
    POSTAL_CODE: 21,
    PACKAGE_TYPE: 22,
    IMAGE_6: 28,
    IMAGE_7: 29,
    IMAGE_8: 30,
    IMAGE_9: 31,
    IMAGE_10: 32
  },

  // Default values for STEP 2: Details
  DEFAULTS: {
    N_VALUE: '360',
    QUANTITY: 1,
    PRICE: '28009',
    CONDITION: 'New',
    COUNTRY: 'US',
    LOCATION: 'San Francisco, CA',
    POSTAL: '94518',
    PACKAGE: 'Rigid Pak'
  }
};

// ============================================================================
// CUSTOM MENU - Appears when sheet opens
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('3DSellers')
    .addItem('✨ Check for New Images (Descriptions)', 'runDescriptions')
    .addItem('📋 Fill Product Details', 'runDetails')
    .addSeparator()
    .addItem('🧪 Test Descriptions (1 image)', 'testDescriptions')
    .addItem('🧪 Test Details (1 row)', 'testDetails')
    .addToUi();
}

// ============================================================================
// STEP 1: 3DSELLERS DESCRIPTIONS
// ============================================================================

/**
 * Menu button: Check for New Images
 */
function runDescriptions() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '3DSELLERS DESCRIPTIONS',
    'Check INBOUND folder for new images?\n\n' +
    'This will analyze images with AI and fill:\n' +
    '• Columns B-L (Artist, Title, Description, etc.)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  ui.alert('Processing...', 'Checking for new images. This may take a few minutes.', ui.ButtonSet.OK);

  const count = checkInboundFolder();

  ui.alert('Complete!', `Processed ${count} new images.\n\nNext step: Click "Fill Product Details"`, ui.ButtonSet.OK);
}

/**
 * Check INBOUND folder and process new images
 */
function checkInboundFolder() {
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('STEP 1: 3DSELLERS DESCRIPTIONS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  try {
    const folder = DriveApp.getFolderById(UNIFIED_CONFIG.INBOUND_FOLDER_ID);
    const sheet = SpreadsheetApp.openById(UNIFIED_CONFIG.SHEET_ID).getSheetByName(UNIFIED_CONFIG.SHEET_TAB_NAME);

    if (!sheet) {
      Logger.log(`❌ Sheet "${UNIFIED_CONFIG.SHEET_TAB_NAME}" not found`);
      return 0;
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

    // Process new images
    const files = folder.getFiles();
    let newCount = 0;

    while (files.hasNext()) {
      const file = files.next();

      if (!file.getMimeType().startsWith('image/')) continue;
      if (processedUrls.has(file.getId())) continue;

      Logger.log(`\n🆕 New image: ${file.getName()}`);
      processNewImage(file, sheet);
      newCount++;

      Utilities.sleep(2000);
    }

    Logger.log(`\n✅ STEP 1 COMPLETE: Processed ${newCount} new images`);
    return newCount;

  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    return 0;
  }
}

/**
 * Process single image with AI analysis
 */
function processNewImage(file, sheet) {
  try {
    Logger.log(`\n📸 Processing: ${file.getName()}`);

    const downloadUrl = `https://drive.google.com/uc?export=view&id=${file.getId()}`;

    // Step 1: Claude Vision
    Logger.log('  🤖 Claude Vision analysis...');
    const analysis = analyzeImageWithClaude(file);
    if (!analysis) {
      Logger.log('  ❌ Analysis failed');
      return;
    }
    Logger.log(`  ✓ Orientation: ${analysis.orientation}, Framed: ${analysis.framed}`);
    if (analysis.substrate === 'Dollar Bill Art') {
      const president = analysis.billPresident || 'Unknown';
      const denomination = analysis.billDenomination || 'Unknown';
      Logger.log(`  💵 DOLLAR BILL ART DETECTED!`);
      Logger.log(`  💵 Bill: ${president}${denomination !== 'Unknown' ? ' (' + denomination + ')' : ''}`);
    }

    // Step 2: Generate listing
    Logger.log('  📝 Generating listing...');
    const listing = generateEbayListing(analysis, file.getName());

    // Step 3: GPT-4 verification
    Logger.log('  🔍 GPT-4 verification...');
    const verifiedListing = verifyWithGPT4(listing, analysis);

    // Step 4: Validate
    if (!validateListing(verifiedListing)) {
      Logger.log('  ❌ Validation failed');
      return;
    }

    // Step 5: Write to sheet (B-M - including WIDTH now)
    const nextRow = sheet.getLastRow() + 1;

    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.ARTIST).setValue('Death NYC');
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.IMAGE_LINK).setFormula(`=IMAGE("${downloadUrl}", 1)`);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.ORIENTATION).setValue(verifiedListing.orientation);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.TITLE).setValue(verifiedListing.title);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.DESCRIPTION).setValue(verifiedListing.description);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.TAGS).setValue(verifiedListing.tags);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.META_KEYWORDS).setValue(verifiedListing.metaKeywords);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.META_DESCRIPTION).setValue(verifiedListing.metaDescription);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.MOBILE_DESC).setValue(verifiedListing.mobileDescription);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.TYPE).setValue(verifiedListing.type);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.LENGTH).setValue(verifiedListing.length);
    sheet.getRange(nextRow, UNIFIED_CONFIG.COLUMNS.WIDTH).setValue(verifiedListing.width);

    sheet.setRowHeight(nextRow, 100);

    Logger.log(`  ✅ Added to row ${nextRow}`);

  } catch (error) {
    Logger.log(`  ❌ Error: ${error.message}`);
  }
}

// ============================================================================
// AI ANALYSIS FUNCTIONS
// ============================================================================

/**
 * Analyze image with Claude Vision
 */
function analyzeImageWithClaude(file) {
  try {
    const blob = getResizedImage(file);
    if (!blob) return null;

    const base64Image = Utilities.base64Encode(blob.getBytes());

    const prompt = `Analyze this Death NYC artwork:
1. Orientation: "portrait" or "landscape"
2. Is it framed? "Framed" or "Not Framed"
3. **IMPORTANT**: Is this printed/screenprinted on a DOLLAR BILL or currency?
   Look carefully for:
   - Dollar bill background or texture
   - Currency markings, serial numbers
   - Presidential portraits or imagery
   - "Federal Reserve Note" text
   - If YES: return "Dollar Bill Art"
   - If NO: return "Standard Print"

4. **IF DOLLAR BILL ART**: Identify which president/person is on the bill:
   - George Washington ($1 bill) = "Washington"
   - Thomas Jefferson ($2 bill) = "Jefferson"
   - Abraham Lincoln ($5 bill) = "Lincoln"
   - Alexander Hamilton ($10 bill) = "Hamilton"
   - Andrew Jackson ($20 bill) = "Jackson"
   - Ulysses Grant ($50 bill) = "Grant"
   - Benjamin Franklin ($100 bill) = "Franklin"
   - If you can see the denomination, note it
   - If unclear, return "Unknown"

5. Subject/theme of the artwork (what's printed ON the dollar)
6. Color palette (3-5 colors)
7. Art style
8. Key elements
9. Emotional tone

JSON format:
{
  "orientation": "portrait" or "landscape",
  "framed": "Framed" or "Not Framed",
  "substrate": "Dollar Bill Art" or "Standard Print",
  "billPresident": "Washington" or "Lincoln" or "Franklin" or "Unknown" (only if substrate is Dollar Bill Art),
  "billDenomination": "$1" or "$5" or "$100" or "Unknown" (only if substrate is Dollar Bill Art),
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
    if (response.getResponseCode() !== 200) return null;

    const data = JSON.parse(response.getContentText());
    const analysisText = data.content[0].text;
    const jsonMatch = analysisText.match(/\{[\s\S]*\}/);

    return jsonMatch ? JSON.parse(jsonMatch[0]) : null;

  } catch (error) {
    Logger.log(`Claude error: ${error.message}`);
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
 * Verify with GPT-4
 */
function verifyWithGPT4(listing, analysis) {
  try {
    const promptData = {
      analysis: analysis,
      listing: listing
    };

    const prompt = `Review and improve this eBay listing. Ensure title starts with "FINE ART" and description is compelling.

Data:
${JSON.stringify(promptData, null, 2)}

Return improved listing in same JSON format.`;

    const payload = {
      model: 'gpt-4',
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.7
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + UNIFIED_CONFIG.OPENAI_API_KEY },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    if (response.getResponseCode() !== 200) {
      return listing;
    }

    const data = JSON.parse(response.getContentText());
    const content = data.choices[0].message.content;
    const jsonMatch = content.match(/\{[\s\S]*\}/);

    return jsonMatch ? JSON.parse(jsonMatch[0]) : listing;

  } catch (error) {
    return listing;
  }
}

/**
 * Generate eBay listing
 */
function generateEbayListing(analysis, fileName) {
  const subject = analysis.subject || 'Contemporary Art';
  const style = analysis.style || 'Pop Art';
  const isDollarBillArt = analysis.substrate === 'Dollar Bill Art';

  // Create special title for dollar bill art - EMPHASIZE 1 OF A KIND
  let title;
  if (isDollarBillArt) {
    title = `FINE ART ${subject} Death NYC 1 OF A KIND DOLLAR ART Street Art`.substring(0, 80);
  } else {
    title = `FINE ART ${subject} Death NYC Limited Edition Print`.substring(0, 80);
  }

  // Create enhanced description with dollar bill details
  let artworkSection;
  if (isDollarBillArt) {
    // Get president and denomination info
    const president = analysis.billPresident || 'Unknown';
    const denomination = analysis.billDenomination || 'Unknown';
    const billDetails = president !== 'Unknown'
      ? `featuring ${president}${denomination !== 'Unknown' ? ' (' + denomination + ' bill)' : ''}`
      : 'authentic U.S. currency';

    artworkSection = `🔥 1 OF A KIND DOLLAR BILL ART 🔥

This is NOT just a print - this artwork is screenprinted on an ACTUAL REAL DOLLAR BILL ${billDetails}! This makes each piece absolutely one-of-a-kind currency art that cannot be replicated.

THE DOLLAR BILL:
The artwork is layered over genuine U.S. currency ${president !== 'Unknown' ? 'showing ' + president : ''}. You can see the original dollar bill ${president !== 'Unknown' ? 'portrait' : 'imagery'} beneath and around the artwork, creating a powerful statement about money, art, and consumerism.

THE ARTWORK:
${analysis.subject ? 'This piece showcases ' + analysis.subject + '. ' : ''}${analysis.elements ? analysis.elements.join(', ') + '. ' : ''}The color palette features ${analysis.colors ? analysis.colors.join(', ') : 'vibrant colors'}.

WHAT MAKES THIS 1 OF A KIND:
• Screenprinted on authentic ${denomination !== 'Unknown' ? denomination : 'U.S.'} dollar bill
• Each bill has unique serial numbers - truly one of a kind
• Real currency ${president !== 'Unknown' ? 'with ' + president + ' portrait visible' : 'visible beneath artwork'}
• Combines fine art with legal tender
• Commentary on money, commerce & pop culture
• Highly collectible street art medium
• Investment-quality contemporary art
• Cannot be duplicated - each serial number is unique`;
  } else {
    artworkSection = `This artwork showcases ${analysis.subject}. ${analysis.elements ? analysis.elements.join(', ') + '. ' : ''}The color palette features ${analysis.colors ? analysis.colors.join(', ') : 'vibrant colors'}.`;
  }

  const description = `THE ARTWORK:
${artworkSection}

THE ARTIST:
Death NYC is a legendary contemporary urban artist, creating bold mashups of luxury brands, cartoon icons, and celebrity imagery. ${isDollarBillArt ? 'His dollar bill series is particularly sought-after by collectors, as each piece transforms legal tender into provocative art.' : ''}

CONDITION: ${isDollarBillArt ? 'Authentic artwork on real currency. ' : ''}High-quality art${isDollarBillArt ? ' on genuine dollar bill' : ' print'}. What you see is what you get.
SHIPPING: Gallery-level care. Protected with archival materials.
GUARANTEE: 30-day money-back guarantee.`;

  // Enhanced tags for dollar bill art
  const tags = isDollarBillArt
    ? `Death NYC, dollar bill art, currency art, money art, ${style.toLowerCase()}, street art, urban art, wall decor`
    : `Death NYC, artwork, art, ${style.toLowerCase()}, wall decor`;

  const type = analysis.framed || 'Not Framed';
  const substrate = analysis.substrate || 'Standard Print';

  // Calculate dimensions based on substrate type
  const length = calculateLength('Death NYC', analysis.orientation, type, substrate);
  const width = calculateWidth('Death NYC', analysis.orientation, type, substrate, length);

  return {
    title,
    description,
    tags,
    orientation: analysis.orientation,
    metaKeywords: tags,
    metaDescription: isDollarBillArt
      ? `${title} - Unique currency art on real dollar bill. ${subject}. Collectible street art.`
      : `${title} - ${subject}. Contemporary street art by Death NYC.`,
    mobileDescription: isDollarBillArt
      ? `1 OF A KIND - ${subject} on real ${analysis.billDenomination || ''} dollar bill. Death NYC street art.`
      : `FINE ART ${style} - ${subject}. Death NYC limited edition.`,
    type,
    length,
    width,
    substrate: substrate
  };
}

/**
 * Calculate length based on artwork type
 * For dollar bill art: Use actual dollar bill dimensions (6.14" x 2.61")
 * For standard prints: Use Death NYC standard sizes
 */
function calculateLength(artist, orientation, type, substrate) {
  if (artist === 'Death NYC') {
    // Dollar bill art uses actual currency dimensions
    if (substrate === 'Dollar Bill Art') {
      // US dollar bills are 6.14" x 2.61" regardless of orientation
      // Length is always the longer dimension
      return '6.14"';
    }

    // Standard print dimensions
    if (orientation === 'landscape' && type === 'Framed') return '16"';
    if (type === 'Not Framed') return '13"';
  }
  return '';
}

/**
 * Calculate width based on artwork type
 * For dollar bill art: Use actual dollar bill dimensions
 * For standard prints: Same as length (square prints)
 */
function calculateWidth(artist, orientation, type, substrate, length) {
  if (artist === 'Death NYC') {
    // Dollar bill art uses actual currency dimensions
    if (substrate === 'Dollar Bill Art') {
      // Width is always the shorter dimension
      return '2.61"';
    }

    // Standard prints: width = length (square)
    return length;
  }
  return length;
}

/**
 * Validate listing
 */
function validateListing(listing) {
  if (!listing.title || !listing.title.startsWith('FINE ART')) return false;
  if (listing.title.length > 80) return false;
  if (!listing.orientation || !listing.type) return false;
  return true;
}

// ============================================================================
// STEP 2: 3DSELLERS DETAILS
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
    '• Columns M-V (Width, Price, Condition, etc.)\n' +
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

    // Get images from folder
    Logger.log('📁 Getting images from Google Drive...');
    const imageUrls = getImageUrlsFromFolder();
    Logger.log(`✓ Found ${imageUrls.length} images`);

    const lastRow = sheet.getLastRow();
    if (lastRow < UNIFIED_CONFIG.START_ROW) return 0;

    // Get Type and Length columns
    const typeColumn = sheet.getRange(UNIFIED_CONFIG.START_ROW, UNIFIED_CONFIG.COLUMNS.TYPE, lastRow - UNIFIED_CONFIG.START_ROW + 1, 1).getValues();
    const lengthColumn = sheet.getRange(UNIFIED_CONFIG.START_ROW, UNIFIED_CONFIG.COLUMNS.LENGTH, lastRow - UNIFIED_CONFIG.START_ROW + 1, 1).getValues();

    let processedCount = 0;

    for (let i = 0; i < typeColumn.length; i++) {
      const rowNum = UNIFIED_CONFIG.START_ROW + i;
      const type = typeColumn[i][0];
      const length = lengthColumn[i][0];

      // Skip rows without descriptions
      if (!type || type === '') continue;

      Logger.log(`\n📝 Row ${rowNum}`);

      // Check if width already exists (from descriptions workflow)
      const existingWidth = sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.WIDTH).getValue();
      const width = existingWidth || length; // Use existing width if present, otherwise use length

      const nValue = UNIFIED_CONFIG.DEFAULTS.N_VALUE;
      const quantity = UNIFIED_CONFIG.DEFAULTS.QUANTITY;
      const price = UNIFIED_CONFIG.DEFAULTS.PRICE;
      const condition = UNIFIED_CONFIG.DEFAULTS.CONDITION;
      const country = UNIFIED_CONFIG.DEFAULTS.COUNTRY;
      const location = UNIFIED_CONFIG.DEFAULTS.LOCATION;
      const postal = UNIFIED_CONFIG.DEFAULTS.POSTAL;
      const packageType = UNIFIED_CONFIG.DEFAULTS.PACKAGE;

      // Write details columns (only write width if not already set)
      if (!existingWidth) {
        sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.WIDTH).setValue(width);
      }
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.COL_N).setValue(nValue);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.QUANTITY).setValue(quantity);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.PRICE).setValue(price);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.CONDITION).setValue(condition);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.COUNTRY_CODE).setValue(country);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.LOCATION).setValue(location);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.POSTAL_CODE).setValue(postal);
      sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.PACKAGE_TYPE).setValue(packageType);

      // Write images
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

  const folder = DriveApp.getFolderById(UNIFIED_CONFIG.INBOUND_FOLDER_ID);
  const files = folder.getFiles();

  if (files.hasNext()) {
    const file = files.next();
    const sheet = SpreadsheetApp.openById(UNIFIED_CONFIG.SHEET_ID).getSheetByName(UNIFIED_CONFIG.SHEET_TAB_NAME);
    Logger.log(`Testing with: ${file.getName()}`);
    processNewImage(file, sheet);
  } else {
    Logger.log('❌ No images in INBOUND folder');
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
  const type = sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.TYPE).getValue();
  const length = sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.LENGTH).getValue();

  if (!type) {
    Logger.log('❌ Row is empty');
    return;
  }

  const width = length;
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.WIDTH).setValue(width);
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.COL_N).setValue(UNIFIED_CONFIG.DEFAULTS.N_VALUE);
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.QUANTITY).setValue(UNIFIED_CONFIG.DEFAULTS.QUANTITY);
  sheet.getRange(rowNum, UNIFIED_CONFIG.COLUMNS.PRICE).setValue(UNIFIED_CONFIG.DEFAULTS.PRICE);
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
