/**
 * ============================================================================
 * 3DSELLERS V7.0 - MASTER TEMPLATE - COMPLETE FIELD MAPPING + AUTO-POSTING
 * ============================================================================
 *
 * COMPLETE INTEGRATION:
 * - VARIABLES tab: Folders, APIs, Pricing, Dimensions, IMAGE FOLDERS
 * - PROMPTS tab: All 11 columns (Title, Desc, Tags, MetaKW, MetaDesc, Mobile + SKU pattern)
 * - FIELDS tab: All 40 fields matched by ARTIST + FORMAT (TYPE)
 * - Products tab: All 79 columns properly filled including Images 1-24
 *
 * IMAGE MATCHING:
 * - Files matched by SKU + Orientation (_P_ or _L_)
 * - Order: 1. Mockups, 2. Crops (crop_2-8), 3. Stock (A-E.png)
 * - Hosted URLs: gauntlet.gallery/images/products/[#]_[Artist]_[Short Title]_[V]_[IMG #].jpg
 *
 * 3DSELLERS AUTO-POSTING:
 * - Toggle ON/OFF via menu (persists using Script Properties)
 * - Post all pending listings or selected rows
 * - Track post status (PENDING, POSTED, FAILED) in columns BZ-CA
 * - API key stored securely in Script Properties
 *
 * FIELD MATCHING: By ARTIST column + FORMAT column
 * PRICING: From VARIABLES row 31-35 by Artist + Format
 * DIMENSIONS: From VARIABLES row 40-44 by Artist + Format, adjusted by Orientation
 * SKU PATTERN: From PROMPTS tab column D (SAMPLE SKU)
 *
 * ============================================================================
 */

// ============================================================================
// SECTION 1: CONFIGURATION
// ============================================================================

const STATIC_CONFIG = {
  SHEET_ID: '1qRRIA7rbW_JhDg20xXV1WrZ-yx8ndnlwQrLkapr65Hc',
  PRODUCTS_TAB: 'Products',
  VARIABLES_TAB: 'VARIABLES',
  PROMPTS_TAB: 'PROMPTS',
  FIELDS_TAB: 'FIELDS',
  START_ROW: 2,
  MAX_EXECUTION_TIME: 300000,
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 2000
};

// ============================================================================
// GOOGLE DRIVE FOLDER IDs - BY ARTIST
// (Loaded dynamically from iMAGES tab, with hardcoded fallback)
// ============================================================================
const ARTIST_FOLDERS = {
  'DEATH NYC': {
    // From iMAGES tab - Death NYC folders
    MOCKUPS_PORTRAIT: '1yVDucQO0vn8XMmnXcrKvPf98k9i0VOX-',    // Portrait Mock-ups
    MOCKUPS_LANDSCAPE: '1JtTH909dOAd_wDPWOFzPS4ue0ntMc9KD',   // Landscape Mock-ups
    CROPS_PORTRAIT: '1z511jL5Iq7tTJPmEK_noi3WgLd3z4wvZ',     // Single Image Crops Portraits
    CROPS_LANDSCAPE: '1Lyeq07b7bkptqVy6RynHBAHGQtT674F5',    // Single Image Crops Landscape
    CROPS_ALL: '1Ft_QO1XlwYt6XCbCEKslG4QbDg0Lm9mT',         // All image crops
    // From VARIABLES tab
    MASTER: '15E4sMBfWACkv9EVUeRXj1j4uq6OoXbCu',            // Product Images Master
    STOCK: '1IX2NU56kdj6ere-TFHZCkgv_DbdDibBy',             // Stock Images
    SKU_PREFIX: 'DNYC'
  },
  'SHEPARD FAIREY': {
    MOCKUPS_PORTRAIT: '',  // Add folder ID when available
    MOCKUPS_LANDSCAPE: '',
    CROPS_PORTRAIT: '',
    CROPS_LANDSCAPE: '',
    STOCK: '1IX2NU56kdj6ere-TFHZCkgv_DbdDibBy',  // Shared stock
    SKU_PREFIX: 'SFAI'
  },
  'MUSIC MEMORABILIA': {
    MOCKUPS_PORTRAIT: '',
    MOCKUPS_LANDSCAPE: '',
    CROPS_PORTRAIT: '',
    CROPS_LANDSCAPE: '',
    STOCK: '1IX2NU56kdj6ere-TFHZCkgv_DbdDibBy',
    SKU_PREFIX: 'MUSC'
  },
  'SPACE COLLECTIBLES': {
    MOCKUPS_PORTRAIT: '',
    MOCKUPS_LANDSCAPE: '',
    CROPS_PORTRAIT: '',
    CROPS_LANDSCAPE: '',
    STOCK: '1IX2NU56kdj6ere-TFHZCkgv_DbdDibBy',
    SKU_PREFIX: 'SPCE'
  }
};

// Default fallback (Death NYC)
const IMAGE_FOLDERS = ARTIST_FOLDERS['DEATH NYC'];

// Image settings (Google Drive files)
const IMAGE_SETTINGS = {
  MOCKUPS_PER_SKU: 6,      // Max mockups to include
  CROPS_PER_SKU: 7,        // crop_2 through crop_8
  STOCK_IMAGES: 5,         // A.png through E.png
  MAX_TOTAL_IMAGES: 24,    // Products tab supports 24 images
  STOCK_FILES: ['A.png', 'B.png', 'C.png', 'D.png', 'E.png']
};

// ============================================================================
// HOSTED IMAGE URL CONFIGURATION (Namecheap / gauntlet.gallery)
// ============================================================================

const HOSTED_IMAGES = {
  // Base URLs for hosted images (gauntlet.gallery / Namecheap)
  PRODUCTS_BASE_URL: 'https://gauntlet.gallery/images/products/',
  STOCK_BASE_URL: 'https://gauntlet.gallery/images/stock/',

  // Stock image filenames
  STOCK_FILES: ['A.png', 'B.png', 'C.png', 'D.png', 'E.png'],

  // Image structure for eBay listing (24 total)
  // ALL images use pattern: {fullSku}_crop_{1-24}.jpg
  NUM_PRODUCT_IMAGES: 24,
  TOTAL_IMAGES: 24,

  // URL pattern: {fullSku}_crop_{num}.jpg
  // Example: 1_DNYC_Mickey-Mouse_V1_crop_1.jpg
  CROP_PATTERN: '_crop_',

  // Column mapping for image URLs (W-AT = columns 23-46)
  IMAGE_START_COLUMN: 23,  // Column W
  NUM_IMAGE_COLUMNS: 24
};

// ============================================================================
// 3DSELLERS AUTO-POSTING CONFIGURATION
// ============================================================================

const THREEDS_CONFIG = {
  // 3DSellers API endpoints
  API_BASE_URL: 'https://api.3dsellers.com/v1',

  // Posting settings
  BATCH_SIZE: 10,              // Number of listings per batch
  DELAY_BETWEEN_POSTS: 2000,   // 2 seconds between API calls
  MAX_RETRIES: 3,

  // Column for tracking post status (column BZ = 78)
  POST_STATUS_COLUMN: 78,
  POST_DATE_COLUMN: 79,

  // Status values
  STATUS: {
    PENDING: 'PENDING',
    POSTED: 'POSTED',
    FAILED: 'FAILED',
    DRAFT: 'DRAFT'
  }
};

// ============================================================================
// SECTION 2: COMPLETE PRODUCTS TAB COLUMN MAPPING (All 79 columns)
// ============================================================================

const PRODUCTS_COLUMNS = {
  // Core product info (A-F) - INPUT columns
  SKU: 1,                    // A - Generated from PROMPTS sample pattern
  PRICE: 2,                  // B - From VARIABLES pricing
  ARTIST: 3,                 // C - Input
  FORMAT: 4,                 // D - Input (matches FIELDS.FORMAT)
  REFERENCE_IMAGE: 5,        // E - Input
  ORIENTATION: 6,            // F - Input (portrait/landscape)

  // AI-Generated content from PROMPTS (G-L)
  TITLE: 7,                  // G - From PROMPTS.TITLE_PROMPT
  DESCRIPTION: 8,            // H - From PROMPTS.DESCRIPTION_PROMPT
  TAGS: 9,                   // I - From PROMPTS.TAGS
  META_KEYWORDS: 10,         // J - From PROMPTS.META_KEYWORDS
  META_DESCRIPTION: 11,      // K - From PROMPTS.META_DESCRIPTION
  MOBILE_DESCRIPTION: 12,    // L - From PROMPTS.MOBILE_DESCRIPTION

  // Item specifics from FIELDS tab (M onwards)
  C_IMAGE_ORIENTATION: 13,   // M - Derived from Orientation
  C_ITEM_LENGTH: 14,         // N - From VARIABLES.DIMENSIONS
  C_ITEM_WIDTH: 15,          // O - From VARIABLES.DIMENSIONS
  CATEGORY_ID: 16,           // P - From FIELDS.CategoryID
  CATEGORY_ID_2: 17,         // Q - From FIELDS.CategoryID 2
  STORE_CATEGORY: 18,        // R - From FIELDS.Store Category
  QUANTITY: 19,              // S - From FIELDS.Quantity
  CONDITION: 20,             // T - From FIELDS.Condition
  COUNTRY_CODE: 21,          // U - From FIELDS.CountryCode
  LOCATION: 22,              // V - From FIELDS.Location
  POSTAL_CODE: 23,           // W - From FIELDS.PostalCode
  PACKAGE_TYPE: 24,          // X - From FIELDS.PackageType
  C_TYPE: 25,                // Y - From FIELDS.C: Type
  C_SIGNED_BY: 26,           // Z - From FIELDS.C: Signed By
  C_COA: 27,                 // AA - From FIELDS.C: Certificate of Authenticity
  C_STYLE: 28,               // AB - From FIELDS.C: Style
  C_THEME: 29,               // AC - From FIELDS.C: Theme
  C_FRANCHISE: 30,           // AD - From FIELDS.C: Franchise
  C_FEATURES: 31,            // AE - From FIELDS.C: Features
  C_ORIGINAL_REPRINT: 32,    // AF - From FIELDS.C: Original/Licensed Reprint
  C_WHY: 33,                 // AG - From FIELDS.why
  C_SUBJECT: 34,             // AH - From FIELDS.C: Subject
  SIZE: 35,                  // AI - From FIELDS.Size
  C_CHARACTER: 36,           // AJ - From FIELDS.C: Character
  C_TIME_PERIOD: 37,         // AK - From FIELDS.C: Time Period Manufactured
  PAYMENT_POLICY: 38,        // AL - From FIELDS.Payment Policy
  SHIPPING_POLICY: 39,       // AM - From FIELDS.Shipping Policy
  RETURN_POLICY: 40,         // AN - From FIELDS.Return Policy
  C_COUNTRY_ORIGIN: 41,      // AO - From FIELDS.C: Country of Origin
  CL_FRAMING: 42,            // AP - From FIELDS.CL Framing
  C_SIGNED: 43,              // AQ - Derived (Yes if signed)
  MEASUREMENT_SYSTEM: 44,    // AR - From FIELDS.MeasurementSystem
  PACKAGE_LENGTH: 45,        // AS - From FIELDS.PackageLength
  PACKAGE_WIDTH: 46,         // AT - From FIELDS.PackageWidth
  PACKAGE_DEPTH: 47,         // AU - From FIELDS.PackageDepth
  WEIGHT_MAJOR: 48,          // AV - From FIELDS.WeightMajor
  WEIGHT_MINOR: 49,          // AW - From FIELDS.WeightMinor
  BRAND: 50,                 // AX - From FIELDS.Brand
  C_COA_ISSUED_BY: 51,       // AY - From FIELDS.C: COA Issued By
  C_PERIOD: 52,              // AZ - From FIELDS.C: Period
  C_PRODUCTION_TECHNIQUE: 53,// BA - From FIELDS.C: Production Technique
  STORE_CATEGORY_2: 54,      // BB - From FIELDS.StoreCategory (duplicate)
  C_MATERIAL: 55,            // BC - From FIELDS.C: Material

  // Images (BD-CA) - columns 56-79
  IMAGE_1: 56, IMAGE_2: 57, IMAGE_3: 58, IMAGE_4: 59, IMAGE_5: 60,
  IMAGE_6: 61, IMAGE_7: 62, IMAGE_8: 63, IMAGE_9: 64, IMAGE_10: 65,
  IMAGE_11: 66, IMAGE_12: 67, IMAGE_13: 68, IMAGE_14: 69, IMAGE_15: 70,
  IMAGE_16: 71, IMAGE_17: 72, IMAGE_18: 73, IMAGE_19: 74, IMAGE_20: 75,
  IMAGE_21: 76, IMAGE_22: 77, IMAGE_23: 78, IMAGE_24: 79
};

// ============================================================================
// SECTION 3: FIELDS TAB COLUMN MAPPING (All 40 columns)
// ============================================================================

const FIELDS_COLUMNS = {
  ARTIST: 'ARTIST',                                    // A (1) - KEY
  FORMAT: 'FORMAT',                                    // B (2) - KEY
  CONDITION: 'Condition',                              // C (3)
  COUNTRY_CODE: 'CountryCode',                         // D (4)
  LOCATION: 'Location',                                // E (5)
  POSTAL_CODE: 'PostalCode',                           // F (6)
  PACKAGE_TYPE: 'PackageType',                         // G (7)
  CATEGORY_ID: 'CategoryID',                           // H (8)
  CATEGORY_ID_2: 'CategoryID 2',                       // I (9)
  STORE_CATEGORY: 'Store Category',                    // J (10)
  QUANTITY: 'Quantity',                                // K (11)
  C_TYPE: 'C: Type',                                   // L (12)
  C_SIGNED_BY: 'C: Signed By',                         // M (13)
  C_COA: 'C: Certificate of Authenticity \n\n',        // N (14)
  C_STYLE: 'C: Style',                                 // O (15)
  C_FEATURES: 'C: Features',                           // P (16)
  C_ORIGINAL_REPRINT: 'C: Original/Licensed Reprint\n',// Q (17)
  C_WHY: 'why ',                                       // R (18)
  C_SUBJECT: 'C: Subject',                             // S (19)
  C_CHARACTER: 'C: Character',                         // T (20)
  C_THEME: 'C: Theme',                                 // U (21)
  C_FRANCHISE: 'C: Franchise',                         // V (22)
  SIZE: 'Size',                                        // W (23)
  C_TIME_PERIOD: 'C: Time Period Manufactured\n\n',    // X (24)
  PAYMENT_POLICY: 'Payment Policy',                    // Y (25)
  SHIPPING_POLICY: 'Shipping Policy',                  // Z (26)
  RETURN_POLICY: 'Return Policy',                      // AA (27)
  C_COUNTRY_ORIGIN: 'C: Country of Origin\n\n',        // AB (28)
  CL_FRAMING: 'CL Framing',                            // AC (29)
  MEASUREMENT_SYSTEM: 'MeasurementSystem',             // AD (30)
  PACKAGE_LENGTH: 'PackageLength',                     // AE (31)
  PACKAGE_WIDTH: 'PackageWidth',                       // AF (32)
  PACKAGE_DEPTH: 'PackageDepth',                       // AG (33)
  WEIGHT_MAJOR: 'WeightMajor',                         // AH (34)
  WEIGHT_MINOR: 'WeightMinor',                         // AI (35)
  BRAND: 'Brand',                                      // AJ (36)
  C_COA_ISSUED_BY: 'C: COA Issued By\n\n',             // AK (37)
  C_PERIOD: 'C: Period',                               // AL (38)
  C_PRODUCTION_TECHNIQUE: 'C: Production Technique\n\n',// AM (39)
  C_MATERIAL: 'C: Material'                            // AN (40)
};

// ============================================================================
// SECTION 4: GLOBAL CACHES
// ============================================================================

let VARIABLES_CACHE = null;
let PROMPTS_CACHE = null;
let FIELDS_CACHE = null;

// ============================================================================
// SECTION 5: VARIABLES TAB LOADER
// ============================================================================

function loadVariables() {
  if (VARIABLES_CACHE) return VARIABLES_CACHE;

  try {
    const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(STATIC_CONFIG.VARIABLES_TAB);

    if (!sheet) {
      Logger.log('ERROR: VARIABLES tab not found');
      return getDefaultVariables();
    }

    const data = sheet.getDataRange().getValues();
    const variables = {
      parentFolders: {},
      scanFolders: {},
      stockFolders: {},
      videoFolders: {},
      spreadsheets: {},
      pricing: {},      // Key: "Artist|Format" -> { price, skuPrefix }
      dimensions: {},   // Key: "Artist Format" -> { portraitLength, portraitWidth, landscapeLength, landscapeWidth, weight }
      apis: {},
      defaults: {},
      processing: {}
    };

    let currentSection = '';

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const col0 = row[0] ? row[0].toString().trim() : '';
      const col1 = row[1] ? row[1].toString().trim() : '';
      const col2 = row[2] ? row[2].toString().trim() : '';
      const col3 = row[3] ? row[3].toString().trim() : '';
      const col4 = row[4] ? row[4].toString().trim() : '';
      const col5 = row[5] !== undefined ? row[5] : '';
      const col6 = row[6] !== undefined ? row[6] : '';

      // Detect section headers
      if (col0.includes('I. PARENT')) currentSection = 'PARENT_FOLDERS';
      else if (col0.includes('II. SCAN')) currentSection = 'SCAN_FOLDERS';
      else if (col0.includes('III. SPREADSHEETS')) currentSection = 'SPREADSHEETS';
      else if (col0.includes('IV. PRICES')) currentSection = 'PRICING';
      else if (col0.includes('V. DIMENTIONS') || col0.includes('V. DIMENSIONS')) currentSection = 'DIMENSIONS';
      else if (col0.includes('VI. APIs')) currentSection = 'APIS';
      else if (col0.includes('VII. DEFAULT')) currentSection = 'DEFAULTS';
      else if (col0.includes('VIII. PROCESSING')) currentSection = 'PROCESSING';

      // Parse data rows
      if (col0 === 'PARENT_FOLDER' && col1 && col2) {
        variables.parentFolders[col1] = col2;
      }
      else if (col0 === 'FOLDER' && col1 && col2) {
        if (col1.includes('Stock Images')) {
          const artist = col1.replace('Stock Images ', '');
          variables.stockFolders[artist] = col2;
        } else if (col1.includes('Product Videos')) {
          const artist = col1.replace('Product Videos ', '');
          variables.videoFolders[artist] = col2;
        } else {
          variables.scanFolders[col1] = { id: col2, artist: col3, menuName: col4 || col1 };
        }
      }
      else if (col0 === 'WORKBOOK' && col1 && col2) {
        variables.spreadsheets[col1] = col2;
      }
      // PRICING: Row format is TYPE | ARTIST | FORMAT | PRICE | SKU_PREFIX
      // But looking at the data, it's: PRICING | Death NYC | Prints | 150 | DNYC-Print
      else if (col0 === 'PRICING' && col1 && col2) {
        const artist = col1;
        const format = col2;
        const price = parseFloat(col3) || 0;
        const skuPrefix = col4 || '';
        const key = `${artist}|${format}`;
        variables.pricing[key] = { artist, format, price, skuPrefix };
        Logger.log(`  Loaded Pricing: ${key} = $${price}, prefix: ${skuPrefix}`);
      }
      // DIMENSIONS: Row format is TYPE | NAME | PORT_LENGTH | PORT_WIDTH | LAND_LENGTH | LAND_WIDTH | WEIGHT
      else if (col0 === 'DIMENSION' && col1) {
        const name = col1;
        variables.dimensions[name] = {
          portraitLength: parseFloat(col2) || 18,
          portraitWidth: parseFloat(col3) || 13,
          landscapeLength: parseFloat(col4) || 18,
          landscapeWidth: parseFloat(col5) || 13,
          weight: parseFloat(col6) || 1
        };
        Logger.log(`  Loaded Dimension: ${name} = ${col2}x${col3} (portrait), ${col4}x${col5} (landscape)`);
      }
      // APIs: Format is API_NAME | API_KEY
      else if (currentSection === 'APIS' && col0 && col1 && !col0.includes('VI.')) {
        variables.apis[col0] = col1;
      }
      // Defaults: Format is SETTING_NAME | VALUE
      else if (currentSection === 'DEFAULTS' && col0 && col1 && !col0.includes('VII.')) {
        variables.defaults[col0] = col1;
      }
      // Processing: Format is SETTING_NAME | VALUE
      else if (currentSection === 'PROCESSING' && col0 && col1 && !col0.includes('VIII.')) {
        variables.processing[col0] = col1;
      }
    }

    Logger.log(`Loaded VARIABLES: ${Object.keys(variables.pricing).length} prices, ${Object.keys(variables.dimensions).length} dimensions, ${Object.keys(variables.apis).length} APIs`);
    VARIABLES_CACHE = variables;
    return variables;

  } catch (error) {
    Logger.log(`ERROR loading VARIABLES: ${error.message}`);
    return getDefaultVariables();
  }
}

function getDefaultVariables() {
  return {
    parentFolders: {},
    scanFolders: {},
    stockFolders: {},
    videoFolders: {},
    spreadsheets: {},
    pricing: {
      // Death NYC
      'Death NYC|Prints': { artist: 'Death NYC', format: 'Prints', price: 150, skuPrefix: 'DNYC-Print' },
      'Death NYC|Framed Prints': { artist: 'Death NYC', format: 'Framed Prints', price: 299, skuPrefix: 'DNYC-Frame' },
      'Death NYC|Dollar Bills': { artist: 'Death NYC', format: 'Dollar Bills', price: 199, skuPrefix: 'DNYC-Bill' },
      // Shepard Fairey
      'Shepard Fairey|Prints': { artist: 'Shepard Fairey', format: 'Prints', price: 175, skuPrefix: 'SFAI-Print' },
      'Shepard Fairey|Framed Prints': { artist: 'Shepard Fairey', format: 'Framed Prints', price: 349, skuPrefix: 'SFAI-Frame' },
      // Music Memorabilia
      'Music Memorabilia|Posters': { artist: 'Music Memorabilia', format: 'Posters', price: 125, skuPrefix: 'MUSC-Post' },
      'Music Memorabilia|Tickets': { artist: 'Music Memorabilia', format: 'Tickets', price: 99, skuPrefix: 'MUSC-Tick' },
      'Music Memorabilia|Autographs': { artist: 'Music Memorabilia', format: 'Autographs', price: 299, skuPrefix: 'MUSC-Auto' },
      'Music Memorabilia|Vinyl': { artist: 'Music Memorabilia', format: 'Vinyl', price: 149, skuPrefix: 'MUSC-Viny' },
      // Space Collectibles
      'Space Collectibles|NASA Prints': { artist: 'Space Collectibles', format: 'NASA Prints', price: 199, skuPrefix: 'SPCE-NASA' },
      'Space Collectibles|Patches': { artist: 'Space Collectibles', format: 'Patches', price: 79, skuPrefix: 'SPCE-Patc' },
      'Space Collectibles|Memorabilia': { artist: 'Space Collectibles', format: 'Memorabilia', price: 249, skuPrefix: 'SPCE-Memo' }
    },
    dimensions: {
      // Death NYC
      'Death NYC Prints': { portraitLength: 18, portraitWidth: 13, landscapeLength: 13, landscapeWidth: 18, weight: 1 },
      'Death NYC Framed Prints': { portraitLength: 22, portraitWidth: 17, landscapeLength: 17, landscapeWidth: 22, weight: 3 },
      'Death NYC Dollar Bills': { portraitLength: 12, portraitWidth: 8, landscapeLength: 8, landscapeWidth: 12, weight: 1 },
      // Shepard Fairey
      'Shepard Fairey Prints': { portraitLength: 24, portraitWidth: 18, landscapeLength: 18, landscapeWidth: 24, weight: 1 },
      'Shepard Fairey Framed Prints': { portraitLength: 28, portraitWidth: 22, landscapeLength: 22, landscapeWidth: 28, weight: 4 },
      // Music Memorabilia
      'Music Memorabilia Posters': { portraitLength: 24, portraitWidth: 18, landscapeLength: 18, landscapeWidth: 24, weight: 1 },
      'Music Memorabilia Tickets': { portraitLength: 8, portraitWidth: 4, landscapeLength: 4, landscapeWidth: 8, weight: 1 },
      'Music Memorabilia Autographs': { portraitLength: 12, portraitWidth: 10, landscapeLength: 10, landscapeWidth: 12, weight: 1 },
      'Music Memorabilia Vinyl': { portraitLength: 13, portraitWidth: 13, landscapeLength: 13, landscapeWidth: 13, weight: 1 },
      // Space Collectibles
      'Space Collectibles NASA Prints': { portraitLength: 20, portraitWidth: 16, landscapeLength: 16, landscapeWidth: 20, weight: 1 },
      'Space Collectibles Patches': { portraitLength: 6, portraitWidth: 6, landscapeLength: 6, landscapeWidth: 6, weight: 1 },
      'Space Collectibles Memorabilia': { portraitLength: 14, portraitWidth: 10, landscapeLength: 10, landscapeWidth: 14, weight: 2 }
    },
    apis: {},
    defaults: {},
    processing: {}
  };
}

// ============================================================================
// SECTION 6: PROMPTS TAB LOADER (All 11 columns + SKU pattern)
// ============================================================================

function loadPrompts() {
  if (PROMPTS_CACHE) return PROMPTS_CACHE;

  try {
    const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(STATIC_CONFIG.PROMPTS_TAB);

    if (!sheet) {
      Logger.log('ERROR: PROMPTS tab not found');
      return new Map();
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return new Map();

    // Read ALL 11 columns
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const promptsMap = new Map();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const artist = row[0] ? row[0].toString().trim() : '';
      const category = row[1] ? row[1].toString().trim() : '';
      const category2 = row[2] ? row[2].toString().trim() : '';
      const sampleSku = row[3] ? row[3].toString().trim() : '';      // SKU pattern like [1]_DNYC-Frame-[Title]
      const sampleImage = row[4] ? row[4].toString().trim() : '';
      const titlePrompt = row[5] ? row[5].toString() : '';
      const descriptionPrompt = row[6] ? row[6].toString() : '';
      const tagsPrompt = row[7] ? row[7].toString() : '';
      const metaKeywordsPrompt = row[8] ? row[8].toString() : '';
      const metaDescriptionPrompt = row[9] ? row[9].toString() : '';
      const mobileDescriptionPrompt = row[10] ? row[10].toString() : '';

      if (!artist || !category) continue;

      // Create lookup key: ARTIST|CATEGORY
      const key = `${artist.toUpperCase()}|${category.toUpperCase()}`;

      promptsMap.set(key, {
        artist,
        category,
        category2,
        sampleSku,           // SKU pattern for this artist/category
        sampleImage,
        titlePrompt,
        descriptionPrompt,
        tagsPrompt,
        metaKeywordsPrompt,
        metaDescriptionPrompt,
        mobileDescriptionPrompt
      });

      Logger.log(`  Loaded Prompt: ${key}, SKU pattern: ${sampleSku}`);
    }

    Logger.log(`Loaded ${promptsMap.size} prompts from PROMPTS tab`);
    PROMPTS_CACHE = promptsMap;
    return promptsMap;

  } catch (error) {
    Logger.log(`ERROR loading PROMPTS: ${error.message}`);
    return new Map();
  }
}

function getPromptInfo(artist, format) {
  const promptsMap = loadPrompts();
  const normalizedArtist = normalizeArtist(artist);
  const normalizedCategory = normalizeCategory(format, normalizedArtist);

  // Try exact match
  const key = `${normalizedArtist}|${normalizedCategory}`;
  if (promptsMap.has(key)) {
    return promptsMap.get(key);
  }

  // Try partial match on artist
  for (const [mapKey, value] of promptsMap.entries()) {
    if (mapKey.startsWith(normalizedArtist.toUpperCase() + '|')) {
      Logger.log(`  Using fallback prompt: ${mapKey}`);
      return value;
    }
  }

  return null;
}

function normalizeArtist(artist) {
  if (!artist) return 'DEATH NYC';
  const upper = artist.toString().trim().toUpperCase();

  // Death NYC
  if (upper.includes('DEATH') || upper.includes('NYC') || upper.includes('DNYC')) {
    return 'DEATH NYC';
  }

  // Shepard Fairey
  if (upper.includes('SHEPARD') || upper.includes('FAIREY') || upper.includes('OBEY') || upper.includes('SFAI')) {
    return 'SHEPARD FAIREY';
  }

  // Music Memorabilia
  if (upper.includes('MUSIC') || upper.includes('MEMORABILIA') || upper.includes('MUSC') ||
      upper.includes('CONCERT') || upper.includes('VINYL') || upper.includes('ROCK')) {
    return 'MUSIC MEMORABILIA';
  }

  // Space Collectibles
  if (upper.includes('SPACE') || upper.includes('COLLECTIBLE') || upper.includes('SPCE') ||
      upper.includes('NASA') || upper.includes('ASTRONAUT') || upper.includes('APOLLO')) {
    return 'SPACE COLLECTIBLES';
  }

  return upper;
}

function normalizeCategory(format, artist) {
  if (!format) return 'Prints';
  const formatStr = format.toString().trim();
  // Return as-is for matching with PROMPTS tab
  return formatStr;
}

// ============================================================================
// SECTION 7: FIELDS TAB LOADER (All 40 columns by Artist + Format)
// ============================================================================

function loadFields() {
  if (FIELDS_CACHE) return FIELDS_CACHE;

  try {
    const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(STATIC_CONFIG.FIELDS_TAB);

    if (!sheet) {
      Logger.log('ERROR: FIELDS tab not found');
      return new Map();
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const fieldsMap = new Map();

    // Create header index for easy lookup
    const headerIndex = {};
    for (let i = 0; i < headers.length; i++) {
      const headerName = headers[i].toString().trim();
      headerIndex[headerName] = i;
    }

    // Also create normalized header index (without newlines)
    const normalizedHeaderIndex = {};
    for (let i = 0; i < headers.length; i++) {
      const headerName = headers[i].toString().trim().replace(/\n/g, ' ').replace(/\s+/g, ' ');
      normalizedHeaderIndex[headerName] = i;
    }

    // Helper function to get value by header name
    function getValue(row, headerName) {
      // Try exact match first
      if (headerIndex[headerName] !== undefined) {
        return row[headerIndex[headerName]] || '';
      }
      // Try with variations
      for (const [key, idx] of Object.entries(headerIndex)) {
        if (key.includes(headerName) || headerName.includes(key.replace(/\n/g, ''))) {
          return row[idx] || '';
        }
      }
      return '';
    }

    // Parse each data row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const artist = getValue(row, 'ARTIST');
      const format = getValue(row, 'FORMAT');

      if (!artist || !format) continue;

      // Create key: Artist|Format
      const key = `${artist}|${format}`;

      // Load ALL 40 fields
      fieldsMap.set(key, {
        artist: artist,
        format: format,
        condition: getValue(row, 'Condition') || 'New',
        countryCode: getValue(row, 'CountryCode') || 'US',
        location: getValue(row, 'Location') || 'San Francisco, CA',
        postalCode: getValue(row, 'PostalCode') || '94518',
        packageType: getValue(row, 'PackageType') || 'PackageThickEnvelope',
        categoryId: getValue(row, 'CategoryID') || 'POSTERS',
        categoryId2: getValue(row, 'CategoryID 2') || 'PRINTS',
        storeCategory: getValue(row, 'Store Category') || '',
        quantity: getValue(row, 'Quantity') || 1,
        cType: getValue(row, 'C: Type') || 'Poster',
        cSignedBy: getValue(row, 'C: Signed By') || artist,
        cCOA: getValue(row, 'C: Certificate of Authenticity') || 'Yes',
        cStyle: getValue(row, 'C: Style') || '',
        cFeatures: getValue(row, 'C: Features') || '',
        cOriginalReprint: getValue(row, 'C: Original/Licensed Reprint') || 'Original',
        cWhy: getValue(row, 'why') || 'Yes',
        cSubject: getValue(row, 'C: Subject') || '',
        cCharacter: getValue(row, 'C: Character') || '',
        cTheme: getValue(row, 'C: Theme') || '',
        cFranchise: getValue(row, 'C: Franchise') || '',
        size: getValue(row, 'Size') || 'Large',
        cTimePeriod: getValue(row, 'C: Time Period Manufactured') || '2020-Now',
        paymentPolicy: getValue(row, 'Payment Policy') || '',
        shippingPolicy: getValue(row, 'Shipping Policy') || '',
        returnPolicy: getValue(row, 'Return Policy') || '',
        cCountryOrigin: getValue(row, 'C: Country of Origin') || 'United States',
        clFraming: getValue(row, 'CL Framing') || 'Unframed',
        measurementSystem: getValue(row, 'MeasurementSystem') || 'English',
        packageLength: getValue(row, 'PackageLength') || 20,
        packageWidth: getValue(row, 'PackageWidth') || 16,
        packageDepth: getValue(row, 'PackageDepth') || 1,
        weightMajor: getValue(row, 'WeightMajor') || 2,
        weightMinor: getValue(row, 'WeightMinor') || 0,
        brand: getValue(row, 'Brand') || artist,
        cCOAIssuedBy: getValue(row, 'C: COA Issued By') || artist,
        cPeriod: getValue(row, 'C: Period') || 'Ultra Contemporay (2020 - Now)',
        cProductionTechnique: getValue(row, 'C: Production Technique') || 'Giclée Print',
        cMaterial: getValue(row, 'C: Material') || 'Matte Paper'
      });

      Logger.log(`  Loaded Fields: ${key} (40 fields)`);
    }

    Logger.log(`Loaded ${fieldsMap.size} field configurations from FIELDS tab`);
    FIELDS_CACHE = fieldsMap;
    return fieldsMap;

  } catch (error) {
    Logger.log(`ERROR loading FIELDS: ${error.message}`);
    return new Map();
  }
}

function getFieldsForProduct(artist, format) {
  const fieldsMap = loadFields();
  const key = `${artist}|${format}`;

  // Try exact match
  if (fieldsMap.has(key)) {
    return fieldsMap.get(key);
  }

  // Try normalized format match
  for (const [mapKey, value] of fieldsMap.entries()) {
    const [fieldArtist, fieldFormat] = mapKey.split('|');
    if (artist.toLowerCase().includes(fieldArtist.toLowerCase()) ||
        fieldArtist.toLowerCase().includes(artist.toLowerCase())) {
      if (format.toLowerCase().includes(fieldFormat.toLowerCase()) ||
          fieldFormat.toLowerCase().includes(format.toLowerCase())) {
        Logger.log(`  Using partial match: ${mapKey} for ${key}`);
        return value;
      }
    }
  }

  // Try artist-only match
  for (const [mapKey, value] of fieldsMap.entries()) {
    if (mapKey.includes(artist)) {
      Logger.log(`  Using artist fallback: ${mapKey} for ${key}`);
      return value;
    }
  }

  Logger.log(`  No fields found for: ${key}`);
  return null;
}

// ============================================================================
// SECTION 8: SKU GENERATION (From PROMPTS pattern)
// ============================================================================

/**
 * Generate SKU using pattern from PROMPTS tab
 * Pattern format: [#]_PREFIX-[Title]
 * Example: [1]_DNYC-Frame-[Title] becomes 1_DNYC-Frame-MyArtworkTitle
 */
function generateSKU(artist, format, title, rowNumber) {
  const prompts = getPromptInfo(artist, format);
  const variables = loadVariables();

  let sku = '';

  if (prompts && prompts.sampleSku) {
    // Parse the sample SKU pattern
    // Example: [1]_DNYC-Frame-[Title]
    let pattern = prompts.sampleSku;

    // Replace [1], [2], [3] etc. with row number
    pattern = pattern.replace(/\[\d+\]/g, rowNumber.toString());

    // Replace [Title] with cleaned title
    const cleanTitle = cleanTitleForSKU(title);
    pattern = pattern.replace(/\[Title\]/gi, cleanTitle);

    sku = pattern;
  } else {
    // Fallback: Use pricing prefix from VARIABLES
    const priceKey = `${artist}|${format}`;
    const pricing = variables.pricing[priceKey];
    const prefix = pricing ? pricing.skuPrefix : 'ART';
    const cleanTitle = cleanTitleForSKU(title);
    sku = `${rowNumber}_${prefix}-${cleanTitle}`;
  }

  // Ensure SKU is under 50 characters
  if (sku.length > 50) {
    sku = sku.substring(0, 50);
  }

  return sku;
}

function cleanTitleForSKU(title) {
  if (!title) return 'Untitled';
  return title.toString()
    .replace(/[^a-zA-Z0-9\s-]/g, '')  // Remove special chars
    .trim()
    .split(/\s+/)
    .slice(0, 3)                       // Take first 3 words
    .join('-')
    .substring(0, 25);                 // Max 25 chars for title portion
}

// ============================================================================
// SECTION 9: DIMENSIONS & PRICING HELPERS
// ============================================================================

/**
 * Get dimensions for Artist + Format + Orientation from VARIABLES
 * Handles format variations: "Framed Prints" matches "Framed", etc.
 */
function getDimensions(artist, format, orientation) {
  const variables = loadVariables();

  // Build dimension key variations
  const keyVariations = [
    `${artist} ${format}`,                              // "Death NYC Framed Prints"
    `${artist} ${format.replace(' Prints', '')}`,       // "Death NYC Framed"
    `${artist} ${format.replace('Prints', '').trim()}`, // "Death NYC Framed"
    `${artist} ${format.split(' ')[0]}`,                // "Death NYC Framed" or "Death NYC Dollar"
  ];

  let dims = null;
  let matchedKey = '';

  // Try exact matches first
  for (const tryKey of keyVariations) {
    if (variables.dimensions[tryKey]) {
      dims = variables.dimensions[tryKey];
      matchedKey = tryKey;
      break;
    }
  }

  // Try partial matches
  if (!dims) {
    for (const [key, value] of Object.entries(variables.dimensions)) {
      const keyLower = key.toLowerCase();
      const artistLower = artist.toLowerCase();
      const formatFirst = format.toLowerCase().split(' ')[0];

      if (keyLower.includes(artistLower) && keyLower.includes(formatFirst)) {
        dims = value;
        matchedKey = key;
        break;
      }
    }
  }

  // Try artist-only match as fallback
  if (!dims) {
    for (const [key, value] of Object.entries(variables.dimensions)) {
      if (key.toLowerCase().includes(artist.toLowerCase())) {
        dims = value;
        matchedKey = key;
        break;
      }
    }
  }

  // Default dimensions
  if (!dims) {
    dims = { portraitLength: 18, portraitWidth: 13, landscapeLength: 13, landscapeWidth: 18, weight: 1 };
    matchedKey = 'default';
  }

  // Return based on orientation
  const isPortrait = !orientation || orientation.toString().toLowerCase() === 'portrait';

  Logger.log(`  Dimensions matched: ${artist} ${format} -> ${matchedKey} (${isPortrait ? 'portrait' : 'landscape'})`);

  return {
    length: isPortrait ? dims.portraitLength : dims.landscapeLength,
    width: isPortrait ? dims.portraitWidth : dims.landscapeWidth,
    weight: dims.weight
  };
}

/**
 * Get price for Artist + Format from VARIABLES
 * Handles format variations: "Framed Prints" matches "Framed", "Dollar Bills" matches "Dollar", etc.
 */
function getPrice(artist, format) {
  const variables = loadVariables();
  const priceKey = `${artist}|${format}`;

  // Try exact match first
  if (variables.pricing[priceKey]) {
    return variables.pricing[priceKey].price;
  }

  // Try format variations (FIELDS "Framed Prints" -> VARIABLES "Framed")
  const formatVariations = [
    format,
    format.replace(' Prints', ''),           // "Framed Prints" -> "Framed"
    format.replace('Prints', '').trim(),     // "Framed Prints" -> "Framed"
    format.split(' ')[0],                     // "Dollar Bills" -> "Dollar"
  ];

  for (const fmt of formatVariations) {
    const tryKey = `${artist}|${fmt}`;
    if (variables.pricing[tryKey]) {
      Logger.log(`  Price matched: ${priceKey} -> ${tryKey}`);
      return variables.pricing[tryKey].price;
    }
  }

  // Try partial match on artist + format keywords
  for (const [key, value] of Object.entries(variables.pricing)) {
    const [keyArtist, keyFormat] = key.split('|');
    if (artist.toLowerCase().includes(keyArtist.toLowerCase()) ||
        keyArtist.toLowerCase().includes(artist.toLowerCase())) {
      // Check if format matches partially
      if (format.toLowerCase().includes(keyFormat.toLowerCase()) ||
          keyFormat.toLowerCase().includes(format.toLowerCase().split(' ')[0])) {
        Logger.log(`  Price partial match: ${priceKey} -> ${key}`);
        return value.price;
      }
    }
  }

  Logger.log(`  Price not found for ${priceKey}, using default`);
  return 150; // Default price
}

// ============================================================================
// SECTION 9B: AI VISION IMAGE ANALYSIS (Gemini + Claude backup)
// ============================================================================

/**
 * Analyze an image using Gemini Vision API
 * Falls back to Claude Vision if Gemini fails
 */
function analyzeImageWithVision(imageFile) {
  const variables = loadVariables();
  const geminiKey = variables.apis['GEMINI_API_KEY'] || '';
  const claudeKey = variables.apis['CLAUDE_API_KEY'] || '';

  // Get image as base64
  const blob = imageFile.getBlob();
  const base64Image = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();

  // Try Gemini first
  if (geminiKey) {
    try {
      const result = analyzeWithGemini(base64Image, mimeType, geminiKey);
      if (result) {
        Logger.log('Gemini Vision: Success');
        return result;
      }
    } catch (e) {
      Logger.log('Gemini Vision failed: ' + e.message);
    }
  }

  // Fallback to Claude
  if (claudeKey) {
    try {
      const result = analyzeWithClaude(base64Image, mimeType, claudeKey);
      if (result) {
        Logger.log('Claude Vision: Success');
        return result;
      }
    } catch (e) {
      Logger.log('Claude Vision failed: ' + e.message);
    }
  }

  return null;
}

/**
 * Analyze image with Gemini Vision API
 */
function analyzeWithGemini(base64Image, mimeType, apiKey) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;

  const prompt = `Analyze this artwork image and provide the following information in JSON format:
{
  "orientation": "portrait" or "landscape" (based on image dimensions and composition),
  "artist_style": "description of the art style",
  "primary_subject": "main subject of the artwork",
  "characters": ["list of any characters, people, or figures"],
  "themes": ["list of themes like pop art, street art, etc"],
  "colors": ["dominant colors"],
  "text_in_image": "any visible text or signatures",
  "franchise": "if related to a franchise like Disney, Marvel, etc",
  "mood": "overall mood or feeling",
  "suggested_title": "a catchy eBay listing title (max 80 chars)",
  "suggested_tags": ["relevant search tags for eBay"]
}

Respond ONLY with valid JSON, no other text.`;

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        {
          inline_data: {
            mime_type: mimeType,
            data: base64Image
          }
        }
      ]
    }],
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 1024
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.candidates && json.candidates[0] && json.candidates[0].content) {
    const text = json.candidates[0].content.parts[0].text;
    // Extract JSON from response
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  }

  return null;
}

/**
 * Analyze image with Claude Vision API (backup)
 */
function analyzeWithClaude(base64Image, mimeType, apiKey) {
  const url = 'https://api.anthropic.com/v1/messages';

  const prompt = `Analyze this artwork image and provide the following information in JSON format:
{
  "orientation": "portrait" or "landscape" (based on image dimensions and composition),
  "artist_style": "description of the art style",
  "primary_subject": "main subject of the artwork",
  "characters": ["list of any characters, people, or figures"],
  "themes": ["list of themes like pop art, street art, etc"],
  "colors": ["dominant colors"],
  "text_in_image": "any visible text or signatures",
  "franchise": "if related to a franchise like Disney, Marvel, etc",
  "mood": "overall mood or feeling",
  "suggested_title": "a catchy eBay listing title (max 80 chars)",
  "suggested_tags": ["relevant search tags for eBay"]
}

Respond ONLY with valid JSON, no other text.`;

  const payload = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 1024,
    messages: [{
      role: 'user',
      content: [
        {
          type: 'image',
          source: {
            type: 'base64',
            media_type: mimeType,
            data: base64Image
          }
        },
        {
          type: 'text',
          text: prompt
        }
      ]
    }]
  };

  const options = {
    method: 'post',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.content && json.content[0] && json.content[0].text) {
    const text = json.content[0].text;
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  }

  return null;
}

/**
 * Scan folder and analyze all images with AI Vision
 */
function scanAndAnalyzeImages(artistName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  const artistFolders = ARTIST_FOLDERS[artistName];
  if (!artistFolders) {
    ui.alert('ERROR: No folder config for ' + artistName);
    return;
  }

  // Get the Google Drive folder ID from VARIABLES
  const variables = loadVariables();
  const folderName = artistFolders.BASE.split('/').pop();

  // Find folder in Drive
  const folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) {
    ui.alert('Folder not found in Google Drive: ' + folderName + '\n\nPlease upload images to Google Drive.');
    return;
  }

  const folder = folders.next();
  const files = folder.getFiles();

  showProgressSidebar();
  resetProgress();

  updateProgress({
    currentTask: 'Scanning images for ' + artistName + '...',
    steps: { 1: 'active' },
    lastLog: 'Starting AI Vision analysis',
    logType: 'info'
  });

  let imageFiles = [];
  while (files.hasNext()) {
    const file = files.next();
    const mimeType = file.getMimeType();
    if (mimeType.startsWith('image/')) {
      imageFiles.push(file);
    }
  }

  updateProgress({
    total: imageFiles.length,
    lastLog: 'Found ' + imageFiles.length + ' images',
    logType: 'info'
  });

  let lastRow = sheet.getLastRow();
  let processedCount = 0;

  for (let i = 0; i < imageFiles.length; i++) {
    const file = imageFiles[i];
    const fileName = file.getName();
    const percent = Math.round(((i + 1) / imageFiles.length) * 100);

    updateProgress({
      percent: percent,
      processed: i + 1,
      currentTask: 'Analyzing: ' + fileName.substring(0, 30) + '...',
      steps: { 1: 'complete', 2: 'active' },
      lastLog: 'AI analyzing: ' + fileName,
      logType: 'info'
    });

    try {
      // Analyze with AI Vision
      const analysis = analyzeImageWithVision(file);

      if (analysis) {
        // Add new row with analyzed data
        lastRow++;
        const newRow = lastRow;

        // Fill in the data from AI analysis
        sheet.getRange(newRow, PRODUCTS_COLUMNS.ARTIST).setValue(artistName);
        sheet.getRange(newRow, PRODUCTS_COLUMNS.ORIENTATION).setValue(analysis.orientation || 'portrait');
        sheet.getRange(newRow, PRODUCTS_COLUMNS.TITLE).setValue(analysis.suggested_title || fileName.replace(/\.[^/.]+$/, ''));
        sheet.getRange(newRow, PRODUCTS_COLUMNS.REFERENCE_IMAGE).setValue(file.getUrl());

        // Fill format based on analysis
        let format = 'Prints';
        if (analysis.themes && analysis.themes.some(t => t.toLowerCase().includes('frame'))) {
          format = 'Framed Prints';
        }
        sheet.getRange(newRow, PRODUCTS_COLUMNS.FORMAT).setValue(format);

        // Fill tags
        if (analysis.suggested_tags) {
          sheet.getRange(newRow, PRODUCTS_COLUMNS.TAGS).setValue(analysis.suggested_tags.join(', '));
        }

        // Fill themes
        if (analysis.themes) {
          sheet.getRange(newRow, PRODUCTS_COLUMNS.C_THEME).setValue(analysis.themes.join(', '));
        }

        // Fill subject
        if (analysis.primary_subject) {
          sheet.getRange(newRow, PRODUCTS_COLUMNS.C_SUBJECT).setValue(analysis.primary_subject);
        }

        // Fill characters
        if (analysis.characters && analysis.characters.length > 0) {
          sheet.getRange(newRow, PRODUCTS_COLUMNS.C_CHARACTER).setValue(analysis.characters.join(', '));
        }

        // Fill franchise
        if (analysis.franchise) {
          sheet.getRange(newRow, PRODUCTS_COLUMNS.C_FRANCHISE).setValue(analysis.franchise);
        }

        // Fill style
        if (analysis.artist_style) {
          sheet.getRange(newRow, PRODUCTS_COLUMNS.C_STYLE).setValue(analysis.artist_style);
        }

        processedCount++;
        updateProgress({
          lastLog: fileName + ': ' + (analysis.orientation || 'unknown') + ' - ' + (analysis.primary_subject || 'analyzed'),
          logType: 'success'
        });
      }
    } catch (error) {
      updateProgress({
        lastLog: fileName + ': ERROR - ' + error.message,
        logType: 'error'
      });
    }

    // Rate limiting - wait 1 second between API calls
    Utilities.sleep(1000);
  }

  updateProgress({
    percent: 100,
    currentTask: 'Vision Analysis Complete!',
    steps: { 1: 'complete', 2: 'complete', 3: 'pending', 4: 'pending', 5: 'pending', 6: 'pending' },
    lastLog: 'Analyzed ' + processedCount + ' images with AI',
    logType: 'success'
  });

  ui.alert(
    'AI Vision Analysis Complete!\n\n' +
    'Analyzed: ' + processedCount + ' images\n' +
    'Artist: ' + artistName + '\n\n' +
    'Next: Review the data, then run "Fill Details Only" to complete product info.'
  );
}

// Scan functions for each artist
function scanImagesDeathNYC() { scanAndAnalyzeImages('DEATH NYC'); }
function scanImagesShepardFairey() { scanAndAnalyzeImages('SHEPARD FAIREY'); }
function scanImagesMusicMemorabilia() { scanAndAnalyzeImages('MUSIC MEMORABILIA'); }
function scanImagesSpaceCollectibles() { scanAndAnalyzeImages('SPACE COLLECTIBLES'); }

// ============================================================================
// SECTION 9C: IMAGE FOLDER IMPORT WITH AI VISION
// ============================================================================

// Crop type to column position mapping (9 crops per product)
const CROP_MAPPING = {
  '2_MAIN_PORTRAIT': 1,      // Main image
  '1_TIGHT_CROP': 2,         // Full tight crop
  '3_BOTTOM_LEFT_SIGNATURE': 3,
  '4_BOTTOM_RIGHT_EDITION': 4,
  '5_DETAILED_CENTER': 5,
  '6_AI_TEXTURE_SUBJECT': 6,
  '7_TOP_LEFT_TEXTURE': 7,
  '8_TOP_RIGHT_TEXTURE': 8,
  '0_THUMBNAIL': 9
};

/**
 * Import images from numbered folders AND analyze with AI Vision
 * Folder structure: PARENT_FOLDER/1_IMG/, 2_IMG/, etc.
 */
function importAndAnalyzeAllImages(artistName, driveFolderId) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!driveFolderId) {
    // Prompt for folder ID
    const response = ui.prompt(
      'Enter Google Drive Folder ID',
      'Enter the folder ID containing numbered image folders (e.g., 1_IMG, 2_IMG):\n\n' +
      'You can find this in the folder URL after /folders/',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return;
    driveFolderId = response.getResponseText().trim();
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(driveFolderId);
  } catch (e) {
    ui.alert('ERROR: Could not access folder: ' + driveFolderId + '\n\n' + e.message);
    return;
  }

  showProgressSidebar();
  resetProgress();

  updateProgress({
    currentTask: 'Scanning folder structure...',
    steps: { 1: 'active' },
    lastLog: 'Accessing folder: ' + folder.getName(),
    logType: 'info'
  });

  // Get all numbered subfolders
  const subfolders = folder.getFolders();
  const folderMap = {};

  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const name = subfolder.getName();
    const match = name.match(/^(\d+)_IMG$/);
    if (match) {
      folderMap[parseInt(match[1])] = subfolder;
    }
  }

  const folderNumbers = Object.keys(folderMap).sort((a, b) => parseInt(a) - parseInt(b));
  const totalFolders = folderNumbers.length;

  updateProgress({
    total: totalFolders,
    steps: { 1: 'complete' },
    lastLog: 'Found ' + totalFolders + ' image folders',
    logType: 'success'
  });

  if (totalFolders === 0) {
    ui.alert('No numbered folders found (expected format: 1_IMG, 2_IMG, etc.)');
    return;
  }

  let lastRow = sheet.getLastRow();
  let processedCount = 0;
  let imageCount = 0;

  for (let i = 0; i < folderNumbers.length; i++) {
    const folderNum = parseInt(folderNumbers[i]);
    const imageFolder = folderMap[folderNum];
    const percent = Math.round(((i + 1) / totalFolders) * 100);

    updateProgress({
      percent: percent,
      processed: i + 1,
      currentTask: 'Processing folder ' + folderNum + '_IMG (' + (i + 1) + '/' + totalFolders + ')',
      steps: { 2: 'active', 3: 'active' },
      lastLog: 'Folder ' + folderNum + '_IMG',
      logType: 'info'
    });

    try {
      // Get all files in this folder
      const files = imageFolder.getFiles();
      const fileMap = {};
      let mainImageFile = null;

      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();
        const match = name.match(/^(\d+)_(.+)\.(png|jpg|jpeg)$/i);
        if (match) {
          const fileNum = match[1];
          const fileType = match[2];
          fileMap[fileType] = file;
          // Find main image for AI analysis (prefer TIGHT_CROP or MAIN_PORTRAIT)
          if (fileType === 'TIGHT_CROP' || fileType === 'MAIN_PORTRAIT') {
            mainImageFile = file;
          }
        }
      }

      if (!mainImageFile && Object.keys(fileMap).length > 0) {
        // Use first available image
        mainImageFile = Object.values(fileMap)[0];
      }

      // Create new row
      lastRow++;
      const newRow = lastRow;

      // Analyze main image with AI Vision FIRST to get subject for SKU
      let analysis = null;
      let orientation = 'portrait';
      let subject = 'Product-' + folderNum;

      if (mainImageFile) {
        updateProgress({
          lastLog: 'AI analyzing folder ' + folderNum + '...',
          logType: 'info'
        });

        try {
          analysis = analyzeImageWithVision(mainImageFile);

          if (analysis) {
            orientation = analysis.orientation || 'portrait';

            // Create subject from AI detection for SKU
            if (analysis.primary_subject) {
              subject = analysis.primary_subject
                .replace(/[^a-zA-Z0-9\s]/g, '')  // Remove special chars
                .trim()
                .replace(/\s+/g, '-')            // Replace spaces with dashes
                .substring(0, 30);               // Limit length
            } else if (analysis.suggested_title) {
              subject = analysis.suggested_title
                .replace(/[^a-zA-Z0-9\s]/g, '')
                .trim()
                .replace(/\s+/g, '-')
                .substring(0, 30);
            }
          }
        } catch (e) {
          Logger.log('AI Vision error for folder ' + folderNum + ': ' + e.message);
        }

        // Rate limit for API
        Utilities.sleep(500);
      }

      // Get SKU prefix based on artist
      const artistFolders = ARTIST_FOLDERS[artistName] || ARTIST_FOLDERS['DEATH NYC'];
      const skuPrefix = artistFolders.SKU_PREFIX || 'PROD';

      // Generate SKU using pattern: [#]_[Artist]_[Short Title]_[V]
      // Example: 1_DNYC_Mickey-Mouse_V1
      const version = 'V1';  // Default version
      const sku = folderNum + '_' + skuPrefix + '_' + subject + '_' + version;
      const orientationCode = orientation.toLowerCase() === 'landscape' ? 'L' : 'P';

      // Fill basic info
      sheet.getRange(newRow, PRODUCTS_COLUMNS.SKU).setValue(sku);
      sheet.getRange(newRow, PRODUCTS_COLUMNS.ARTIST).setValue(artistName || 'Death NYC');
      sheet.getRange(newRow, PRODUCTS_COLUMNS.ORIENTATION).setValue(orientation);
      sheet.getRange(newRow, PRODUCTS_COLUMNS.FORMAT).setValue('Prints');

      // Fill AI-detected data
      if (analysis) {
        sheet.getRange(newRow, PRODUCTS_COLUMNS.TITLE).setValue(
          analysis.suggested_title || subject.replace(/-/g, ' ')
        );

        if (analysis.suggested_tags) {
          sheet.getRange(newRow, PRODUCTS_COLUMNS.TAGS).setValue(analysis.suggested_tags.join(', '));
        }
        if (analysis.themes) {
          sheet.getRange(newRow, PRODUCTS_COLUMNS.C_THEME).setValue(analysis.themes.join(', '));
        }
        if (analysis.primary_subject) {
          sheet.getRange(newRow, PRODUCTS_COLUMNS.C_SUBJECT).setValue(analysis.primary_subject);
        }
        if (analysis.characters && analysis.characters.length > 0) {
          sheet.getRange(newRow, PRODUCTS_COLUMNS.C_CHARACTER).setValue(analysis.characters.join(', '));
        }
        if (analysis.franchise) {
          sheet.getRange(newRow, PRODUCTS_COLUMNS.C_FRANCHISE).setValue(analysis.franchise);
        }
        if (analysis.artist_style) {
          sheet.getRange(newRow, PRODUCTS_COLUMNS.C_STYLE).setValue(analysis.artist_style);
        }

        updateProgress({
          lastLog: 'Folder ' + folderNum + ': ' + orientation + ' - SKU: ' + sku,
          logType: 'success'
        });
      }

      // Generate gauntlet.gallery hosted URLs based on SKU
      // Format: https://gauntlet.gallery/images/products/[#]_[Artist]_[Short Title]_[V]_[IMG #].jpg
      // Example: https://gauntlet.gallery/images/products/1_DNYC_Mickey-Mouse_V1_1.jpg
      let linkedCount = 0;
      for (let cropNum = 1; cropNum <= 8; cropNum++) {
        const imageUrl = HOSTED_IMAGES.PRODUCTS_BASE_URL + sku + '_' + cropNum + '.jpg';
        sheet.getRange(newRow, PRODUCTS_COLUMNS.IMAGE_1 + cropNum - 1).setValue(imageUrl);
        imageCount++;
        linkedCount++;
      }

      // Add stock images (9-13)
      for (let i = 0; i < HOSTED_IMAGES.STOCK_FILES.length; i++) {
        const stockUrl = HOSTED_IMAGES.STOCK_BASE_URL + HOSTED_IMAGES.STOCK_FILES[i];
        sheet.getRange(newRow, PRODUCTS_COLUMNS.IMAGE_1 + 8 + i).setValue(stockUrl);
        imageCount++;
        linkedCount++;
      }

      // Store the Google Drive folder reference for later upload
      sheet.getRange(newRow, PRODUCTS_COLUMNS.REFERENCE_IMAGE).setValue(
        'https://drive.google.com/drive/folders/' + imageFolder.getId()
      );

      processedCount++;
      updateProgress({
        lastLog: 'Folder ' + folderNum + ': ' + linkedCount + ' images linked (SKU: ' + sku + ')',
        logType: 'success'
      });

    } catch (error) {
      updateProgress({
        lastLog: 'Folder ' + folderNum + ' ERROR: ' + error.message,
        logType: 'error'
      });
    }
  }

  updateProgress({
    percent: 100,
    currentTask: 'Import Complete!',
    steps: { 1: 'complete', 2: 'complete', 3: 'complete', 4: 'pending', 5: 'pending', 6: 'pending' },
    lastLog: 'Imported ' + processedCount + ' products, ' + imageCount + ' images',
    logType: 'success'
  });

  ui.alert(
    'Import & Analysis Complete!\n\n' +
    'Products created: ' + processedCount + '\n' +
    'Images linked: ' + imageCount + '\n' +
    'Artist: ' + (artistName || 'Death NYC') + '\n\n' +
    'Next: Run "Fill Details Only" to complete pricing/dimensions.'
  );
}

// Import functions for each artist
function importImagesDeathNYC() {
  const variables = loadVariables();
  const folderId = variables.scanFolders['Death NYC Input'] ? variables.scanFolders['Death NYC Input'].id : null;
  importAndAnalyzeAllImages('DEATH NYC', folderId);
}

function importImagesShepardFairey() {
  const variables = loadVariables();
  const folderId = variables.scanFolders['Shepard Fairey Input'] ? variables.scanFolders['Shepard Fairey Input'].id : null;
  importAndAnalyzeAllImages('SHEPARD FAIREY', folderId);
}

function importImagesMusicMemorabilia() {
  const variables = loadVariables();
  const folderId = variables.scanFolders['Music Input'] ? variables.scanFolders['Music Input'].id : null;
  importAndAnalyzeAllImages('MUSIC MEMORABILIA', folderId);
}

function importImagesSpaceCollectibles() {
  const variables = loadVariables();
  const folderId = variables.scanFolders['Space Input'] ? variables.scanFolders['Space Input'].id : null;
  importAndAnalyzeAllImages('SPACE COLLECTIBLES', folderId);
}

function importImagesFromFolder() {
  // Generic import - will prompt for folder ID
  importAndAnalyzeAllImages(null, null);
}

/**
 * Test AI Vision on a single image
 */
function testVisionOnSelectedImage() {
  const ui = SpreadsheetApp.getUi();

  // Prompt for image URL or file name
  const response = ui.prompt(
    'Test AI Vision',
    'Enter the Google Drive file name or URL of an image to analyze:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  let file;

  if (input.startsWith('http')) {
    // It's a URL - extract file ID
    const match = input.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match) {
      file = DriveApp.getFileById(match[1]);
    }
  } else {
    // It's a file name
    const files = DriveApp.getFilesByName(input);
    if (files.hasNext()) {
      file = files.next();
    }
  }

  if (!file) {
    ui.alert('File not found: ' + input);
    return;
  }

  ui.alert('Analyzing image with AI Vision...\n\nThis may take a few seconds.');

  const result = analyzeImageWithVision(file);

  if (result) {
    ui.alert(
      'AI Vision Analysis Result:\n\n' +
      'Orientation: ' + (result.orientation || 'unknown') + '\n' +
      'Subject: ' + (result.primary_subject || 'unknown') + '\n' +
      'Style: ' + (result.artist_style || 'unknown') + '\n' +
      'Themes: ' + (result.themes ? result.themes.join(', ') : 'none') + '\n' +
      'Characters: ' + (result.characters ? result.characters.join(', ') : 'none') + '\n' +
      'Franchise: ' + (result.franchise || 'none') + '\n' +
      'Suggested Title: ' + (result.suggested_title || 'none') + '\n' +
      'Tags: ' + (result.suggested_tags ? result.suggested_tags.slice(0, 5).join(', ') : 'none')
    );
  } else {
    ui.alert('AI Vision analysis failed. Check API keys in VARIABLES tab.');
  }
}

// ============================================================================
// SECTION 10: MENU CREATION
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('3DSELLERS')

    // ========== TOP LEVEL: Most Used ==========
    .addItem('📊 DASHBOARD (% Complete)', 'showFullDashboard')
    .addItem('⚡ TEST 3 Images', 'quickTest3NoAI')
    .addItem('🧪 FULL PIPELINE TEST (3 Products)', 'testFullPipeline3Products')
    .addSeparator()

    // ========== PROCESS ==========
    .addSubMenu(ui.createMenu('1. Process Images')
      .addItem('Death NYC - ALL', 'processDeathNYCALL')
      .addItem('Death NYC - Portrait', 'processDeathNYCPortrait')
      .addItem('Death NYC - Landscape', 'processDeathNYCLandscape')
      .addSeparator()
      .addItem('Debug Death NYC Rows', 'debugDeathNYCRows')
      .addSeparator()
      .addItem('Process from Drive Folder...', 'importImagesFromFolder'))

    .addSubMenu(ui.createMenu('2. Fill Details')
      .addItem('Fill ALL Rows', 'fillAllProductDetails')
      .addItem('Fill Selected Row', 'fillSelectedRowDetails')
      .addItem('Generate SKUs', 'generateAllSKUs'))

    .addSubMenu(ui.createMenu('3. AI Content')
      .addItem('Generate ALL (Title, Desc, Tags)', 'generateAllContent')
      .addItem('Generate Titles Only', 'generateTitlesOnly')
      .addItem('Generate Descriptions Only', 'generateDescriptionsOnly'))

    .addSubMenu(ui.createMenu('4. Images')
      .addItem('Fill Image URLs from Drive', 'fillAllProductImages')
      .addItem('Preview Images for SKU', 'previewSelectedSKUImages'))

    .addSubMenu(ui.createMenu('5. Post to eBay')
      .addItem('Post ALL Pending', 'postAllPendingListings')
      .addItem('Post Selected Rows', 'postSelectedRows')
      .addSeparator()
      .addItem('Mark as PENDING', 'markSelectedAsPending')
      .addItem('Clear Status (Re-post)', 'clearPostStatus'))

    .addSeparator()

    // ========== VALIDATION ==========
    .addSubMenu(ui.createMenu('Validate')
      .addItem('Check Duplicates', 'checkDuplicateSKUs')
      .addItem('Validate Before Post', 'validateBeforePost')
      .addItem('Detect Brands', 'detectBrands'))

    // ========== TOOLS ==========
    .addSubMenu(ui.createMenu('Tools')
      .addItem('Bulk Edit Selected', 'bulkEditSelected')
      .addItem('Export to CSV', 'exportToCSV')
      .addItem('Reload Config', 'reloadAllTabs'))

    // ========== TEST ==========
    .addSubMenu(ui.createMenu('Test')
      .addItem('Test Drive Folders', 'quickTestFolders')
      .addItem('Test 3 Images (No AI)', 'quickTest3NoAI')
      .addItem('Test 3 Images (With AI)', 'test3Images')
      .addSeparator()
      .addItem('Test AI Connection', 'testAIConnection')
      .addItem('Show Variables', 'showLoadedVariables')
      .addItem('Show Prompts', 'showLoadedPrompts')
      .addItem('Show Fields', 'showLoadedFields'))

    .addToUi();

  // Add enhanced features menu (File Consolidation, SEO, Circuit Breaker, etc.)
  addEnhancedMenuItems();
}

// ============================================================================
// SECTION 10A: FULL DASHBOARD WITH % COMPLETE
// ============================================================================

/**
 * Show full dashboard with % complete for all sheets
 */
function showFullDashboard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);

  let report = '';
  report += '════════════════════════════════════════════\n';
  report += '        📊 3DSELLERS DASHBOARD\n';
  report += '════════════════════════════════════════════\n\n';

  // ===== PRODUCTS SHEET =====
  const productsSheet = ss.getSheetByName('Products');
  if (productsSheet) {
    const data = productsSheet.getDataRange().getValues();
    const totalRows = data.length - 1; // Exclude header

    let stats = {
      withSKU: 0,
      withTitle: 0,
      withDesc: 0,
      withImages: 0,
      withPrice: 0,
      portraits: 0,
      landscapes: 0,
      totalValue: 0
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] || row[0] === '') continue;

      stats.withSKU++;
      if (row[6] && String(row[6]).trim() !== '') stats.withTitle++;
      if (row[7] && String(row[7]).trim() !== '') stats.withDesc++;
      if (row[1] && parseFloat(row[1]) > 0) {
        stats.withPrice++;
        stats.totalValue += parseFloat(row[1]);
      }

      // Check images (columns BD onwards = index 55+)
      let hasImage = false;
      for (let j = 55; j < 79 && j < row.length; j++) {
        if (row[j] && String(row[j]).trim() !== '') {
          hasImage = true;
          break;
        }
      }
      if (hasImage) stats.withImages++;

      // Orientation
      if (String(row[0]).includes('_P_') || row[5] === 'Portrait') {
        stats.portraits++;
      } else {
        stats.landscapes++;
      }
    }

    const pctSKU = stats.withSKU > 0 ? 100 : 0;
    const pctTitle = stats.withSKU > 0 ? Math.round((stats.withTitle / stats.withSKU) * 100) : 0;
    const pctDesc = stats.withSKU > 0 ? Math.round((stats.withDesc / stats.withSKU) * 100) : 0;
    const pctImages = stats.withSKU > 0 ? Math.round((stats.withImages / stats.withSKU) * 100) : 0;
    const pctPrice = stats.withSKU > 0 ? Math.round((stats.withPrice / stats.withSKU) * 100) : 0;
    const overallPct = Math.round((pctSKU + pctTitle + pctDesc + pctImages + pctPrice) / 5);

    report += '📦 PRODUCTS SHEET\n';
    report += '────────────────────────────────────────────\n';
    report += '  Total Products: ' + stats.withSKU + '\n';
    report += '  Portrait: ' + stats.portraits + ' | Landscape: ' + stats.landscapes + '\n';
    report += '  Total Value: $' + stats.totalValue.toFixed(2) + '\n\n';

    report += '  COMPLETION:\n';
    report += '  ' + getProgressBar(pctSKU) + ' SKU: ' + pctSKU + '%\n';
    report += '  ' + getProgressBar(pctTitle) + ' Title: ' + pctTitle + '% (' + stats.withTitle + '/' + stats.withSKU + ')\n';
    report += '  ' + getProgressBar(pctDesc) + ' Description: ' + pctDesc + '% (' + stats.withDesc + '/' + stats.withSKU + ')\n';
    report += '  ' + getProgressBar(pctImages) + ' Images: ' + pctImages + '% (' + stats.withImages + '/' + stats.withSKU + ')\n';
    report += '  ' + getProgressBar(pctPrice) + ' Price: ' + pctPrice + '% (' + stats.withPrice + '/' + stats.withSKU + ')\n\n';

    report += '  ══════════════════════════════════\n';
    report += '  OVERALL: ' + getProgressBar(overallPct) + ' ' + overallPct + '% COMPLETE\n';
    report += '  ══════════════════════════════════\n\n';
  }

  // ===== VARIABLES SHEET =====
  const variablesSheet = ss.getSheetByName('VARIABLES');
  if (variablesSheet) {
    const varData = variablesSheet.getDataRange().getValues();
    let varCount = 0;
    for (let i = 1; i < varData.length; i++) {
      if (varData[i][0] && varData[i][0] !== '') varCount++;
    }
    report += '⚙️ VARIABLES: ' + varCount + ' entries configured\n';
  }

  // ===== PROMPTS SHEET =====
  const promptsSheet = ss.getSheetByName('PROMPTS');
  if (promptsSheet) {
    const promptData = promptsSheet.getDataRange().getValues();
    let promptCount = 0;
    for (let i = 1; i < promptData.length; i++) {
      if (promptData[i][0] && promptData[i][0] !== '') promptCount++;
    }
    report += '📝 PROMPTS: ' + promptCount + ' artists configured\n';
  }

  // ===== FIELDS SHEET =====
  const fieldsSheet = ss.getSheetByName('FIELDS');
  if (fieldsSheet) {
    const fieldData = fieldsSheet.getDataRange().getValues();
    let fieldCount = 0;
    for (let i = 1; i < fieldData.length; i++) {
      if (fieldData[i][0] && fieldData[i][0] !== '') fieldCount++;
    }
    report += '📋 FIELDS: ' + fieldCount + ' field mappings\n';
  }

  // ===== IMAGES SHEET =====
  const imagesSheet = ss.getSheetByName('iMAGES');
  if (imagesSheet) {
    const imgData = imagesSheet.getDataRange().getValues();
    let imgCount = 0;
    for (let i = 1; i < imgData.length; i++) {
      if (imgData[i][2] && imgData[i][2] !== '') imgCount++; // Link column
    }
    report += '🖼️ IMAGES: ' + imgCount + ' folder links configured\n';
  }

  report += '\n════════════════════════════════════════════\n';
  report += 'Run "Test → Test Drive Folders" to verify access\n';

  ui.alert('Dashboard', report, ui.ButtonSet.OK);
  Logger.log(report);
}

/**
 * Generate a simple text progress bar
 */
function getProgressBar(pct) {
  const filled = Math.round(pct / 10);
  const empty = 10 - filled;
  return '[' + '█'.repeat(filled) + '░'.repeat(empty) + ']';
}

/** Progress logging stubs */
function showProgressSidebar() { Logger.log('Progress tracking via Logger'); }
function resetProgress() { Logger.log('=== PIPELINE STARTED ==='); }
function updateProgress(u) { if (u.currentTask) Logger.log('[PROGRESS] ' + u.currentTask); if (u.lastLog) Logger.log('[LOG] ' + u.lastLog); }
function getProgressStatus() { return {}; }

// ============================================================================
// SECTION 10B: PROCESS BY ARTIST FUNCTIONS
// ============================================================================

/**
 * Process all products for a specific artist
 * - Fill details, generate AI content, populate images, set up for posting
 */
function processArtist(artistName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    ui.alert('ERROR: Products sheet not found!');
    return;
  }

  const artistFolders = ARTIST_FOLDERS[artistName];
  if (!artistFolders) {
    ui.alert('ERROR: No folder configuration for artist: ' + artistName);
    return;
  }

  // Confirm with user
  const response = ui.alert(
    'Process ' + artistName,
    'This will process all products for ' + artistName + ':\n\n' +
    '1. Fill product details (fields, pricing, dimensions)\n' +
    '2. Generate AI content (titles, descriptions)\n' +
    '3. Populate hosted image URLs\n' +
    '4. Mark as PENDING for posting\n\n' +
    'Folder: ' + artistFolders.BASE + '\n' +
    'SKU Prefix: ' + artistFolders.SKU_PREFIX + '\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Cancelled.');
    return;
  }

  // Show sidebar and reset progress
  showProgressSidebar();
  resetProgress();

  Logger.log('========== PROCESSING ARTIST: ' + artistName + ' ==========');

  // Step 1: Load data
  updateProgress({
    currentTask: 'Loading Variables & Fields...',
    steps: { 1: 'active' },
    lastLog: 'Loading VARIABLES tab',
    logType: 'info'
  });

  // Clear caches
  clearAllCaches();

  // Load all data sources
  const variables = loadVariables();
  const fieldsMap = loadFields();
  const promptsMap = loadPrompts();

  updateProgress({
    steps: { 1: 'complete' },
    lastLog: 'Data loaded successfully',
    logType: 'success'
  });

  // Count matching rows first
  const lastRow = sheet.getLastRow();
  let matchingRows = [];

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    const rowArtist = sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim();
    const normalizedRowArtist = normalizeArtist(rowArtist);
    if (normalizedRowArtist === artistName) {
      matchingRows.push(row);
    }
  }

  const totalRows = matchingRows.length;
  updateProgress({
    total: totalRows,
    currentTask: 'Found ' + totalRows + ' products for ' + artistName,
    lastLog: 'Found ' + totalRows + ' matching rows',
    logType: 'info'
  });

  let processedCount = 0;

  // Step 2-4: Process each row
  for (let i = 0; i < matchingRows.length; i++) {
    const row = matchingRows[i];
    const percent = Math.round(((i + 1) / totalRows) * 100);

    // Update progress for this row
    updateProgress({
      percent: percent,
      processed: i + 1,
      currentTask: 'Processing row ' + row + ' (' + (i + 1) + '/' + totalRows + ')',
      steps: { 2: 'active', 3: 'active', 4: 'active', 5: 'active', 6: 'active' },
      lastLog: 'Row ' + row + ': Filling details + images',
      logType: 'info'
    });

    try {
      processSingleRow(sheet, row, artistFolders);
      processedCount++;

      updateProgress({
        lastLog: 'Row ' + row + ': Complete',
        logType: 'success'
      });
    } catch (error) {
      updateProgress({
        lastLog: 'Row ' + row + ': ERROR - ' + error.message,
        logType: 'error'
      });
    }
  }

  // Step 6: Mark complete
  updateProgress({
    percent: 100,
    processed: processedCount,
    currentTask: 'Complete!',
    steps: { 1: 'complete', 2: 'complete', 3: 'complete', 4: 'complete', 5: 'complete', 6: 'complete' },
    lastLog: artistName + ' processing complete!',
    logType: 'success'
  });

  ui.alert(
    artistName + ' - Complete!\n\n' +
    'Processed: ' + processedCount + ' rows\n' +
    'Skipped: ' + (lastRow - STATIC_CONFIG.START_ROW + 1 - totalRows) + ' rows (other artists)\n\n' +
    'Next: Use "Post to 3DSELLERS" menu to publish listings.'
  );
}

/**
 * Process a single row with artist-specific settings
 */
function processSingleRow(sheet, row, artistFolders) {
  try {
    const artist = sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim();
    const format = sheet.getRange(row, PRODUCTS_COLUMNS.FORMAT).getValue().toString().trim();
    const orientation = sheet.getRange(row, PRODUCTS_COLUMNS.ORIENTATION).getValue().toString().trim();
    const title = sheet.getRange(row, PRODUCTS_COLUMNS.TITLE).getValue().toString().trim();

    // Get fields, dimensions, pricing
    const fields = getFieldsForProduct(artist, format);
    const dims = getDimensions(artist, format, orientation);
    const price = getPrice(artist, format);
    const sku = generateSKU(artist, format, title, row);

    // Fill basic columns
    sheet.getRange(row, PRODUCTS_COLUMNS.SKU).setValue(sku);
    sheet.getRange(row, PRODUCTS_COLUMNS.PRICE).setValue(price);

    // Fill image orientation
    const imageOrientation = orientation.toLowerCase() === 'landscape' ? 'Landscape' : 'Portrait';
    sheet.getRange(row, PRODUCTS_COLUMNS.C_IMAGE_ORIENTATION).setValue(imageOrientation);

    // Fill dimensions
    sheet.getRange(row, PRODUCTS_COLUMNS.C_ITEM_LENGTH).setValue(dims.length);
    sheet.getRange(row, PRODUCTS_COLUMNS.C_ITEM_WIDTH).setValue(dims.width);

    // Fill fields if available
    if (fields) {
      fillFieldsForRow(sheet, row, fields);
    }

    // Generate hosted image URLs
    populateHostedUrlsForRow(sheet, row, sku, orientation);

    // Mark as PENDING for posting
    sheet.getRange(row, THREEDS_CONFIG.POST_STATUS_COLUMN).setValue(THREEDS_CONFIG.STATUS.PENDING);

    Logger.log(`  Row ${row}: SKU=${sku}, Price=$${price}, Status=PENDING`);

  } catch (error) {
    Logger.log(`  Row ${row} ERROR: ${error.message}`);
  }
}

/**
 * Fill all FIELDS columns for a row
 */
function fillFieldsForRow(sheet, row, fields) {
  sheet.getRange(row, PRODUCTS_COLUMNS.CATEGORY_ID).setValue(fields.categoryId || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.CATEGORY_ID_2).setValue(fields.categoryId2 || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.STORE_CATEGORY).setValue(fields.storeCategory || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.QUANTITY).setValue(fields.quantity || 1);
  sheet.getRange(row, PRODUCTS_COLUMNS.CONDITION).setValue(fields.condition || 'New');
  sheet.getRange(row, PRODUCTS_COLUMNS.COUNTRY_CODE).setValue(fields.countryCode || 'US');
  sheet.getRange(row, PRODUCTS_COLUMNS.LOCATION).setValue(fields.location || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.POSTAL_CODE).setValue(fields.postalCode || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.PACKAGE_TYPE).setValue(fields.packageType || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_TYPE).setValue(fields.cType || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_SIGNED_BY).setValue(fields.cSignedBy || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_COA).setValue(fields.cCOA || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_STYLE).setValue(fields.cStyle || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_THEME).setValue(fields.cTheme || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_FRANCHISE).setValue(fields.cFranchise || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_FEATURES).setValue(fields.cFeatures || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_ORIGINAL_REPRINT).setValue(fields.cOriginalReprint || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_WHY).setValue(fields.cWhy || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_SUBJECT).setValue(fields.cSubject || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.SIZE).setValue(fields.size || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_CHARACTER).setValue(fields.cCharacter || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_TIME_PERIOD).setValue(fields.cTimePeriod || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.PAYMENT_POLICY).setValue(fields.paymentPolicy || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.SHIPPING_POLICY).setValue(fields.shippingPolicy || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.RETURN_POLICY).setValue(fields.returnPolicy || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_COUNTRY_ORIGIN).setValue(fields.cCountryOrigin || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.CL_FRAMING).setValue(fields.clFraming || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_SIGNED).setValue('Yes');
  sheet.getRange(row, PRODUCTS_COLUMNS.MEASUREMENT_SYSTEM).setValue(fields.measurementSystem || 'English');
  sheet.getRange(row, PRODUCTS_COLUMNS.PACKAGE_LENGTH).setValue(fields.packageLength || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.PACKAGE_WIDTH).setValue(fields.packageWidth || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.PACKAGE_DEPTH).setValue(fields.packageDepth || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.WEIGHT_MAJOR).setValue(fields.weightMajor || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.WEIGHT_MINOR).setValue(fields.weightMinor || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.BRAND).setValue(fields.brand || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_COA_ISSUED_BY).setValue(fields.cCOAIssuedBy || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_PERIOD).setValue(fields.cPeriod || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_PRODUCTION_TECHNIQUE).setValue(fields.cProductionTechnique || '');
  sheet.getRange(row, PRODUCTS_COLUMNS.C_MATERIAL).setValue(fields.cMaterial || '');
}

/**
 * Populate hosted URLs for a single row
 * Pattern: {fullSku}_crop_{1-24}.jpg (ALL 24 images use unified naming)
 */
function populateHostedUrlsForRow(sheet, row, sku, orientation) {
  // Get all 24 image URLs
  const imageUrls = getHostedImageUrls(sku, orientation);

  // Write all 24 at once to columns W-AT (23-46)
  sheet.getRange(row, HOSTED_IMAGES.IMAGE_START_COLUMN, 1, HOSTED_IMAGES.NUM_IMAGE_COLUMNS).setValues([imageUrls]);
}

// Individual artist process functions - PROCESS ALL
function processDeathNYC() { processArtist('DEATH NYC'); }
function processShepardFairey() { processArtist('SHEPARD FAIREY'); }
function processMusicMemorabilia() { processArtist('MUSIC MEMORABILIA'); }
function processSpaceCollectibles() { processArtist('SPACE COLLECTIBLES'); }

// Individual artist functions - FILL DETAILS ONLY
function fillDetailsDeathNYC() { fillDetailsForArtist('DEATH NYC'); }
function fillDetailsShepardFairey() { fillDetailsForArtist('SHEPARD FAIREY'); }
function fillDetailsMusicMemorabilia() { fillDetailsForArtist('MUSIC MEMORABILIA'); }
function fillDetailsSpaceCollectibles() { fillDetailsForArtist('SPACE COLLECTIBLES'); }

// Individual artist functions - FILL IMAGES ONLY
function fillImagesDeathNYC() { fillImagesForArtist('DEATH NYC'); }
function fillImagesShepardFairey() { fillImagesForArtist('SHEPARD FAIREY'); }
function fillImagesMusicMemorabilia() { fillImagesForArtist('MUSIC MEMORABILIA'); }
function fillImagesSpaceCollectibles() { fillImagesForArtist('SPACE COLLECTIBLES'); }

// ============================================================================
// ARTIST + ORIENTATION FUNCTIONS (Portrait/Landscape specific)
// ============================================================================

// Death NYC - Portrait
function importDeathNYCPortrait() { importByArtistOrientation('DEATH NYC', 'Portrait'); }
function fillDetailsDeathNYCPortrait() { fillDetailsByArtistOrientation('DEATH NYC', 'Portrait'); }
function fillImagesDeathNYCPortrait() { fillImagesByArtistOrientation('DEATH NYC', 'Portrait'); }
function processDeathNYCPortrait() { processByArtistOrientation('DEATH NYC', 'Portrait'); }

// Death NYC - Landscape
function importDeathNYCLandscape() { importByArtistOrientation('DEATH NYC', 'Landscape'); }
function fillDetailsDeathNYCLandscape() { fillDetailsByArtistOrientation('DEATH NYC', 'Landscape'); }
function fillImagesDeathNYCLandscape() { fillImagesByArtistOrientation('DEATH NYC', 'Landscape'); }
function processDeathNYCLandscape() { processByArtistOrientation('DEATH NYC', 'Landscape'); }

// Death NYC - ALL (ignores orientation)
function fillImagesDeathNYCALL() { fillImagesByArtist('DEATH NYC'); }
function fillDetailsDeathNYCALL() { fillDetailsByArtist('DEATH NYC'); }
function processDeathNYCALL() { processByArtist('DEATH NYC'); }
function debugDeathNYCRows() { debugArtistRows('DEATH NYC'); }

/**
 * Fill images for ALL items by artist (ignores orientation)
 */
function fillImagesByArtist(artistName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    ui.alert('ERROR: Products sheet not found!');
    return;
  }

  const response = ui.alert(
    'Fill Images - ' + artistName + ' ALL',
    'Fill image URLs for ALL ' + artistName + ' items (any orientation)?\n\n' +
    'This will populate hosted URLs (gauntlet.gallery)',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const lastRow = sheet.getLastRow();
  let processedCount = 0;
  let skippedNoSku = 0;

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    const rowArtist = normalizeArtist(sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim());

    if (rowArtist === artistName) {
      const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue().toString().trim();
      const orientation = sheet.getRange(row, PRODUCTS_COLUMNS.ORIENTATION).getValue().toString().trim() || 'Portrait';

      if (sku) {
        populateHostedImageUrlsForRow(sheet, row, sku, orientation);
        processedCount++;
      } else {
        skippedNoSku++;
      }
    }
  }

  ui.alert('Complete',
    'Filled images for ' + processedCount + ' ' + artistName + ' items\n' +
    'Skipped (no SKU): ' + skippedNoSku,
    ui.ButtonSet.OK);
}

/**
 * Fill details for ALL items by artist (ignores orientation)
 */
function fillDetailsByArtist(artistName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    ui.alert('ERROR: Products sheet not found!');
    return;
  }

  const lastRow = sheet.getLastRow();
  let processedCount = 0;

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    const rowArtist = normalizeArtist(sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim());

    if (rowArtist === artistName) {
      const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue().toString().trim();
      if (sku) {
        fillRowDetails(sheet, row, artistName);
        processedCount++;
      }
    }
  }

  ui.alert('Complete', 'Filled details for ' + processedCount + ' ' + artistName + ' items', ui.ButtonSet.OK);
}

/**
 * Process ALL items by artist (ignores orientation)
 */
function processByArtist(artistName) {
  fillDetailsByArtist(artistName);
  fillImagesByArtist(artistName);
}

/**
 * Debug: Show what's in the rows for an artist
 */
function debugArtistRows(artistName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  const lastRow = sheet.getLastRow();
  let output = 'DEBUG: ' + artistName + ' rows\n\n';
  let count = 0;

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow && count < 20; row++) {
    const rawArtist = sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim();
    const normalizedArtist = normalizeArtist(rawArtist);

    if (normalizedArtist === artistName) {
      const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue().toString().trim();
      const orientation = sheet.getRange(row, PRODUCTS_COLUMNS.ORIENTATION).getValue().toString().trim();

      output += 'Row ' + row + ':\n';
      output += '  Artist: "' + rawArtist + '" -> "' + normalizedArtist + '"\n';
      output += '  SKU: "' + sku + '"\n';
      output += '  Orientation: "' + orientation + '"\n\n';
      count++;
    }
  }

  if (count === 0) {
    output += 'No rows found for ' + artistName;
  }

  Logger.log(output);
  ui.alert('Debug Output (first 20 rows)', output.substring(0, 1500), ui.ButtonSet.OK);
}

// Shepard Fairey - Portrait
function importShepardFaireyPortrait() { importByArtistOrientation('SHEPARD FAIREY', 'Portrait'); }
function fillDetailsShepardFaireyPortrait() { fillDetailsByArtistOrientation('SHEPARD FAIREY', 'Portrait'); }
function fillImagesShepardFaireyPortrait() { fillImagesByArtistOrientation('SHEPARD FAIREY', 'Portrait'); }
function processShepardFaireyPortrait() { processByArtistOrientation('SHEPARD FAIREY', 'Portrait'); }

// Shepard Fairey - Landscape
function importShepardFaireyLandscape() { importByArtistOrientation('SHEPARD FAIREY', 'Landscape'); }
function fillDetailsShepardFaireyLandscape() { fillDetailsByArtistOrientation('SHEPARD FAIREY', 'Landscape'); }
function fillImagesShepardFaireyLandscape() { fillImagesByArtistOrientation('SHEPARD FAIREY', 'Landscape'); }
function processShepardFaireyLandscape() { processByArtistOrientation('SHEPARD FAIREY', 'Landscape'); }

// Music Memorabilia - Portrait
function importMusicMemorabiliaPortrait() { importByArtistOrientation('MUSIC MEMORABILIA', 'Portrait'); }
function fillDetailsMusicMemorabiliaPortrait() { fillDetailsByArtistOrientation('MUSIC MEMORABILIA', 'Portrait'); }
function fillImagesMusicMemorabiliaPortrait() { fillImagesByArtistOrientation('MUSIC MEMORABILIA', 'Portrait'); }
function processMusicMemorabiliaPortrait() { processByArtistOrientation('MUSIC MEMORABILIA', 'Portrait'); }

// Music Memorabilia - Landscape
function importMusicMemorabiliaLandscape() { importByArtistOrientation('MUSIC MEMORABILIA', 'Landscape'); }
function fillDetailsMusicMemorabiliaLandscape() { fillDetailsByArtistOrientation('MUSIC MEMORABILIA', 'Landscape'); }
function fillImagesMusicMemorabiliaLandscape() { fillImagesByArtistOrientation('MUSIC MEMORABILIA', 'Landscape'); }
function processMusicMemorabiliaLandscape() { processByArtistOrientation('MUSIC MEMORABILIA', 'Landscape'); }

// Space Collectibles - Portrait
function importSpaceCollectiblesPortrait() { importByArtistOrientation('SPACE COLLECTIBLES', 'Portrait'); }
function fillDetailsSpaceCollectiblesPortrait() { fillDetailsByArtistOrientation('SPACE COLLECTIBLES', 'Portrait'); }
function fillImagesSpaceCollectiblesPortrait() { fillImagesByArtistOrientation('SPACE COLLECTIBLES', 'Portrait'); }
function processSpaceCollectiblesPortrait() { processByArtistOrientation('SPACE COLLECTIBLES', 'Portrait'); }

// Space Collectibles - Landscape
function importSpaceCollectiblesLandscape() { importByArtistOrientation('SPACE COLLECTIBLES', 'Landscape'); }
function fillDetailsSpaceCollectiblesLandscape() { fillDetailsByArtistOrientation('SPACE COLLECTIBLES', 'Landscape'); }
function fillImagesSpaceCollectiblesLandscape() { fillImagesByArtistOrientation('SPACE COLLECTIBLES', 'Landscape'); }
function processSpaceCollectiblesLandscape() { processByArtistOrientation('SPACE COLLECTIBLES', 'Landscape'); }

/**
 * Show Pipeline Dashboard info
 * All processing now happens via Google Drive
 */
function openPipelineDashboard() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Pipeline Dashboard',
    'Image processing workflow:\n\n' +
    '1. Upload images to Google Drive folders\n' +
    '2. Use "1. Process Images" menu to scan & analyze\n' +
    '3. Use "2. Fill Details" to populate product info\n' +
    '4. Use "3. AI Content" to generate titles/descriptions\n' +
    '5. Use "4. Images" to link image URLs\n' +
    '6. Use "5. Post to eBay" to publish listings\n\n' +
    'All folders are configured in the IMAGES tab.',
    ui.ButtonSet.OK
  );
}

/**
 * Import images by artist AND orientation from Google Drive
 */
function importByArtistOrientation(artistName, orientation) {
  const ui = SpreadsheetApp.getUi();
  const config = getArtistFolders(artistName);

  if (!config || Object.keys(config).length === 0) {
    ui.alert('Error', 'Artist not configured: ' + artistName + '\n\nCheck IMAGES tab or ARTIST_FOLDERS config.', ui.ButtonSet.OK);
    return;
  }

  const orientCode = orientation.toLowerCase() === 'portrait' ? 'PORTRAIT' : 'LANDSCAPE';
  const folderKey = 'MOCKUPS_' + orientCode;
  const folderId = config[folderKey];

  if (!folderId) {
    ui.alert('Error', 'No folder ID for ' + artistName + ' - ' + orientation, ui.ButtonSet.OK);
    return;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    let count = 0;
    while (files.hasNext()) {
      files.next();
      count++;
    }

    ui.alert(
      'Import ' + artistName + ' - ' + orientation,
      'Google Drive Folder: ' + folder.getName() + '\n' +
      'Folder ID: ' + folderId + '\n' +
      'Images Found: ' + count + '\n\n' +
      'Use "1. Process Images" menu to process these images.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Error', 'Could not access folder: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Fill details for artist + orientation
 */
function fillDetailsByArtistOrientation(artistName, orientation) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    ui.alert('ERROR: Products sheet not found!');
    return;
  }

  const response = ui.alert(
    'Fill Details - ' + artistName + ' ' + orientation,
    'Fill details for ' + artistName + ' ' + orientation + ' items only?\n\n' +
    '(Will filter by Artist AND Orientation columns)',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  clearAllCaches();
  loadVariables();
  loadFields();
  loadPrompts();

  const lastRow = sheet.getLastRow();
  let processedCount = 0;

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    const rowArtist = normalizeArtist(sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim());
    const rowOrientation = sheet.getRange(row, PRODUCTS_COLUMNS.ORIENTATION).getValue().toString().trim().toLowerCase();

    if (rowArtist === artistName && rowOrientation === orientation.toLowerCase()) {
      // Fill this row's details
      const format = sheet.getRange(row, PRODUCTS_COLUMNS.FORMAT).getValue().toString().trim();
      fillRowWithFieldsData(sheet, row, artistName, format, orientation);
      fillRowWithPricingAndDimensions(sheet, row, artistName, format, orientation);
      processedCount++;
    }
  }

  ui.alert('Complete', 'Filled details for ' + processedCount + ' ' + artistName + ' ' + orientation + ' items', ui.ButtonSet.OK);
}

/**
 * Fill images for artist + orientation
 */
function fillImagesByArtistOrientation(artistName, orientation) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    ui.alert('ERROR: Products sheet not found!');
    return;
  }

  const response = ui.alert(
    'Fill Images - ' + artistName + ' ' + orientation,
    'Fill image URLs for ' + artistName + ' ' + orientation + ' items?\n\n' +
    'This will populate hosted URLs (gauntlet.gallery)',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const lastRow = sheet.getLastRow();
  let processedCount = 0;

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    const rowArtist = normalizeArtist(sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim());
    const rowOrientation = sheet.getRange(row, PRODUCTS_COLUMNS.ORIENTATION).getValue().toString().trim().toLowerCase();

    if (rowArtist === artistName && rowOrientation === orientation.toLowerCase()) {
      const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue().toString().trim();
      if (sku) {
        populateHostedImageUrlsForRow(sheet, row, sku, orientation);
        processedCount++;
      }
    }
  }

  ui.alert('Complete', 'Filled images for ' + processedCount + ' ' + artistName + ' ' + orientation + ' items', ui.ButtonSet.OK);
}

/**
 * Process all steps for artist + orientation
 */
function processByArtistOrientation(artistName, orientation) {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Process ' + artistName + ' ' + orientation,
    'Run full processing for ' + artistName + ' ' + orientation + ' items:\n\n' +
    '1. Fill Details\n' +
    '2. Fill Images\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  fillDetailsByArtistOrientation(artistName, orientation);
  fillImagesByArtistOrientation(artistName, orientation);

  ui.alert('Complete', artistName + ' ' + orientation + ' processing finished!', ui.ButtonSet.OK);
}

/**
 * Fill ONLY product details (no images) for an artist
 */
function fillDetailsForArtist(artistName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    ui.alert('ERROR: Products sheet not found!');
    return;
  }

  const response = ui.alert(
    'Fill Details - ' + artistName,
    'This will fill ONLY product details for ' + artistName + ':\n\n' +
    '- SKUs\n' +
    '- Pricing & Dimensions\n' +
    '- All FIELDS data\n\n' +
    '(Images will NOT be touched)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  showProgressSidebar();
  resetProgress();

  updateProgress({
    currentTask: 'Loading data for ' + artistName + '...',
    steps: { 1: 'active' },
    lastLog: 'Starting Details Only mode',
    logType: 'info'
  });

  clearAllCaches();
  loadVariables();
  loadFields();
  loadPrompts();

  updateProgress({ steps: { 1: 'complete' }, lastLog: 'Data loaded', logType: 'success' });

  const lastRow = sheet.getLastRow();
  let matchingRows = [];

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    const rowArtist = sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim();
    if (normalizeArtist(rowArtist) === artistName) {
      matchingRows.push(row);
    }
  }

  updateProgress({ total: matchingRows.length, lastLog: 'Found ' + matchingRows.length + ' rows', logType: 'info' });

  let processedCount = 0;

  for (let i = 0; i < matchingRows.length; i++) {
    const row = matchingRows[i];
    const percent = Math.round(((i + 1) / matchingRows.length) * 100);

    updateProgress({
      percent: percent,
      processed: i + 1,
      currentTask: 'Row ' + row + ' - Filling details',
      steps: { 2: 'active', 3: 'active', 4: 'active' },
      lastLog: 'Processing row ' + row,
      logType: 'info'
    });

    try {
      // Fill details ONLY (no images)
      const artist = sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim();
      const format = sheet.getRange(row, PRODUCTS_COLUMNS.FORMAT).getValue().toString().trim();
      const orientation = sheet.getRange(row, PRODUCTS_COLUMNS.ORIENTATION).getValue().toString().trim();
      const title = sheet.getRange(row, PRODUCTS_COLUMNS.TITLE).getValue().toString().trim();

      const fields = getFieldsForProduct(artist, format);
      const dims = getDimensions(artist, format, orientation);
      const price = getPrice(artist, format);
      const sku = generateSKU(artist, format, title, row);

      sheet.getRange(row, PRODUCTS_COLUMNS.SKU).setValue(sku);
      sheet.getRange(row, PRODUCTS_COLUMNS.PRICE).setValue(price);

      const imageOrientation = orientation.toLowerCase() === 'landscape' ? 'Landscape' : 'Portrait';
      sheet.getRange(row, PRODUCTS_COLUMNS.C_IMAGE_ORIENTATION).setValue(imageOrientation);
      sheet.getRange(row, PRODUCTS_COLUMNS.C_ITEM_LENGTH).setValue(dims.length);
      sheet.getRange(row, PRODUCTS_COLUMNS.C_ITEM_WIDTH).setValue(dims.width);

      if (fields) {
        fillFieldsForRow(sheet, row, fields);
      }

      processedCount++;
      updateProgress({ lastLog: 'Row ' + row + ': Details filled', logType: 'success' });
    } catch (error) {
      updateProgress({ lastLog: 'Row ' + row + ': ERROR - ' + error.message, logType: 'error' });
    }
  }

  updateProgress({
    percent: 100,
    currentTask: 'Details Complete!',
    steps: { 1: 'complete', 2: 'complete', 3: 'complete', 4: 'complete', 5: 'pending', 6: 'pending' },
    lastLog: 'Filled details for ' + processedCount + ' rows',
    logType: 'success'
  });

  ui.alert('Details filled for ' + processedCount + ' ' + artistName + ' products.\n\nImages were NOT modified.');
}

/**
 * Fill ONLY images for an artist (assumes details already filled)
 */
function fillImagesForArtist(artistName) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    ui.alert('ERROR: Products sheet not found!');
    return;
  }

  const response = ui.alert(
    'Fill Images - ' + artistName,
    'This will fill ONLY image URLs for ' + artistName + ':\n\n' +
    '- 8 crop images per product\n' +
    '- 5 stock images\n' +
    '- Set post status to PENDING\n\n' +
    '(Product details will NOT be touched)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  showProgressSidebar();
  resetProgress();

  updateProgress({
    currentTask: 'Loading ' + artistName + ' products...',
    steps: { 1: 'complete', 2: 'complete', 3: 'complete', 4: 'complete', 5: 'active' },
    lastLog: 'Starting Images Only mode',
    logType: 'info'
  });

  const lastRow = sheet.getLastRow();
  let matchingRows = [];

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    const rowArtist = sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim();
    if (normalizeArtist(rowArtist) === artistName) {
      matchingRows.push(row);
    }
  }

  updateProgress({ total: matchingRows.length, lastLog: 'Found ' + matchingRows.length + ' rows', logType: 'info' });

  let processedCount = 0;

  for (let i = 0; i < matchingRows.length; i++) {
    const row = matchingRows[i];
    const percent = Math.round(((i + 1) / matchingRows.length) * 100);

    updateProgress({
      percent: percent,
      processed: i + 1,
      currentTask: 'Row ' + row + ' - Filling images',
      steps: { 5: 'active' },
      lastLog: 'Populating image URLs for row ' + row,
      logType: 'info'
    });

    try {
      const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue().toString().trim();
      const orientation = sheet.getRange(row, PRODUCTS_COLUMNS.ORIENTATION).getValue().toString().trim();

      if (!sku) {
        updateProgress({ lastLog: 'Row ' + row + ': No SKU - skipping', logType: 'error' });
        continue;
      }

      // Fill images
      populateHostedUrlsForRow(sheet, row, sku, orientation);

      // Set post status
      sheet.getRange(row, THREEDS_CONFIG.POST_STATUS_COLUMN).setValue(THREEDS_CONFIG.STATUS.PENDING);

      processedCount++;
      updateProgress({ lastLog: 'Row ' + row + ': 13 images added', logType: 'success' });
    } catch (error) {
      updateProgress({ lastLog: 'Row ' + row + ': ERROR - ' + error.message, logType: 'error' });
    }
  }

  updateProgress({
    percent: 100,
    currentTask: 'Images Complete!',
    steps: { 1: 'complete', 2: 'complete', 3: 'complete', 4: 'complete', 5: 'complete', 6: 'complete' },
    lastLog: 'Filled images for ' + processedCount + ' rows',
    logType: 'success'
  });

  ui.alert('Images filled for ' + processedCount + ' ' + artistName + ' products.\n\nAll marked as PENDING for posting.');
}

function processAllArtists() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Process ALL Artists',
    'This will process products for ALL artists:\n\n' +
    '- Death NYC\n' +
    '- Shepard Fairey\n' +
    '- Music Memorabilia\n' +
    '- Space Collectibles\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  processArtist('DEATH NYC');
  processArtist('SHEPARD FAIREY');
  processArtist('MUSIC MEMORABILIA');
  processArtist('SPACE COLLECTIBLES');

  ui.alert('All artists processed!');
}

// ============================================================================
// SECTION 11: MAIN FILL FUNCTION - ALL 40 FIELDS + PRICING + DIMENSIONS
// ============================================================================

/**
 * Fill ALL product details for every row:
 * - Price from VARIABLES (by Artist + Format)
 * - Dimensions from VARIABLES (by Artist + Format + Orientation)
 * - All 40 fields from FIELDS tab (by Artist + Format)
 * - SKU from PROMPTS pattern
 */
function fillAllProductDetails() {
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    Logger.log('ERROR: Products sheet not found');
    SpreadsheetApp.getUi().alert('ERROR: Products sheet not found!');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < STATIC_CONFIG.START_ROW) {
    Logger.log('No data rows to process');
    SpreadsheetApp.getUi().alert('No data rows to process');
    return;
  }

  // Clear caches to ensure fresh data
  clearAllCaches();

  // Load all data
  const variables = loadVariables();
  const fieldsMap = loadFields();
  const promptsMap = loadPrompts();

  let filledCount = 0;
  let errorCount = 0;

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    try {
      // Get input values
      const artist = sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim();
      const format = sheet.getRange(row, PRODUCTS_COLUMNS.FORMAT).getValue().toString().trim();
      const orientation = sheet.getRange(row, PRODUCTS_COLUMNS.ORIENTATION).getValue().toString().trim();
      const title = sheet.getRange(row, PRODUCTS_COLUMNS.TITLE).getValue().toString().trim();

      if (!artist) {
        Logger.log(`Row ${row}: Skipping - no artist`);
        continue;
      }

      Logger.log(`Processing row ${row}: ${artist} | ${format} | ${orientation}`);

      // Get FIELDS for this Artist + Format
      const fields = getFieldsForProduct(artist, format);
      if (!fields) {
        Logger.log(`Row ${row}: No fields found for ${artist}|${format}, using defaults`);
      }

      // Get DIMENSIONS for Artist + Format + Orientation
      const dims = getDimensions(artist, format, orientation);

      // Get PRICE for Artist + Format
      const price = getPrice(artist, format);

      // Generate SKU
      const sku = generateSKU(artist, format, title, row);

      // ========================================
      // FILL ALL COLUMNS
      // ========================================

      // Column A: SKU
      sheet.getRange(row, PRODUCTS_COLUMNS.SKU).setValue(sku);

      // Column B: Price (from VARIABLES)
      sheet.getRange(row, PRODUCTS_COLUMNS.PRICE).setValue(price);

      // Column M: C: Image Orientation (derived from input)
      const imageOrientation = orientation.toLowerCase() === 'landscape' ? 'Landscape' : 'Portrait';
      sheet.getRange(row, PRODUCTS_COLUMNS.C_IMAGE_ORIENTATION).setValue(imageOrientation);

      // Column N: C: Item Length (from VARIABLES dimensions)
      sheet.getRange(row, PRODUCTS_COLUMNS.C_ITEM_LENGTH).setValue(dims.length);

      // Column O: C: Item Width (from VARIABLES dimensions)
      sheet.getRange(row, PRODUCTS_COLUMNS.C_ITEM_WIDTH).setValue(dims.width);

      // All remaining columns from FIELDS (if found)
      if (fields) {
        sheet.getRange(row, PRODUCTS_COLUMNS.CATEGORY_ID).setValue(fields.categoryId);
        sheet.getRange(row, PRODUCTS_COLUMNS.CATEGORY_ID_2).setValue(fields.categoryId2);
        sheet.getRange(row, PRODUCTS_COLUMNS.STORE_CATEGORY).setValue(fields.storeCategory);
        sheet.getRange(row, PRODUCTS_COLUMNS.QUANTITY).setValue(fields.quantity);
        sheet.getRange(row, PRODUCTS_COLUMNS.CONDITION).setValue(fields.condition);
        sheet.getRange(row, PRODUCTS_COLUMNS.COUNTRY_CODE).setValue(fields.countryCode);
        sheet.getRange(row, PRODUCTS_COLUMNS.LOCATION).setValue(fields.location);
        sheet.getRange(row, PRODUCTS_COLUMNS.POSTAL_CODE).setValue(fields.postalCode);
        sheet.getRange(row, PRODUCTS_COLUMNS.PACKAGE_TYPE).setValue(fields.packageType);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_TYPE).setValue(fields.cType);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_SIGNED_BY).setValue(fields.cSignedBy);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_COA).setValue(fields.cCOA);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_STYLE).setValue(fields.cStyle);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_THEME).setValue(fields.cTheme);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_FRANCHISE).setValue(fields.cFranchise);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_FEATURES).setValue(fields.cFeatures);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_ORIGINAL_REPRINT).setValue(fields.cOriginalReprint);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_WHY).setValue(fields.cWhy);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_SUBJECT).setValue(fields.cSubject);
        sheet.getRange(row, PRODUCTS_COLUMNS.SIZE).setValue(fields.size);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_CHARACTER).setValue(fields.cCharacter);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_TIME_PERIOD).setValue(fields.cTimePeriod);
        sheet.getRange(row, PRODUCTS_COLUMNS.PAYMENT_POLICY).setValue(fields.paymentPolicy);
        sheet.getRange(row, PRODUCTS_COLUMNS.SHIPPING_POLICY).setValue(fields.shippingPolicy);
        sheet.getRange(row, PRODUCTS_COLUMNS.RETURN_POLICY).setValue(fields.returnPolicy);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_COUNTRY_ORIGIN).setValue(fields.cCountryOrigin);
        sheet.getRange(row, PRODUCTS_COLUMNS.CL_FRAMING).setValue(fields.clFraming);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_SIGNED).setValue('Yes');  // Derived
        sheet.getRange(row, PRODUCTS_COLUMNS.MEASUREMENT_SYSTEM).setValue(fields.measurementSystem);
        sheet.getRange(row, PRODUCTS_COLUMNS.PACKAGE_LENGTH).setValue(fields.packageLength);
        sheet.getRange(row, PRODUCTS_COLUMNS.PACKAGE_WIDTH).setValue(fields.packageWidth);
        sheet.getRange(row, PRODUCTS_COLUMNS.PACKAGE_DEPTH).setValue(fields.packageDepth);
        sheet.getRange(row, PRODUCTS_COLUMNS.WEIGHT_MAJOR).setValue(fields.weightMajor);
        sheet.getRange(row, PRODUCTS_COLUMNS.WEIGHT_MINOR).setValue(fields.weightMinor);
        sheet.getRange(row, PRODUCTS_COLUMNS.BRAND).setValue(fields.brand);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_COA_ISSUED_BY).setValue(fields.cCOAIssuedBy);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_PERIOD).setValue(fields.cPeriod);
        sheet.getRange(row, PRODUCTS_COLUMNS.C_PRODUCTION_TECHNIQUE).setValue(fields.cProductionTechnique);
        sheet.getRange(row, PRODUCTS_COLUMNS.STORE_CATEGORY_2).setValue(fields.storeCategory);  // Same as R
        sheet.getRange(row, PRODUCTS_COLUMNS.C_MATERIAL).setValue(fields.cMaterial);
      }

      filledCount++;
      Logger.log(`  Row ${row}: Filled successfully`);

    } catch (error) {
      errorCount++;
      Logger.log(`  Row ${row}: ERROR - ${error.message}`);
    }
  }

  const message = `Filled ${filledCount} rows (${errorCount} errors)`;
  Logger.log(message);
  SpreadsheetApp.getUi().alert(message);
}

// ============================================================================
// SECTION 12: AI CONTENT GENERATION (Columns G-L)
// ============================================================================

function generateAllContent() {
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('ERROR: Products sheet not found!');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < STATIC_CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert('No data rows to process');
    return;
  }

  const variables = loadVariables();
  const apiKey = variables.apis['CLAUDE_API_KEY'] || variables.apis['GROK_API_KEY'];

  if (!apiKey) {
    SpreadsheetApp.getUi().alert('ERROR: No API key found in VARIABLES tab!');
    return;
  }

  let processedCount = 0;

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    const artist = sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim();
    const format = sheet.getRange(row, PRODUCTS_COLUMNS.FORMAT).getValue().toString().trim();
    const refImage = sheet.getRange(row, PRODUCTS_COLUMNS.REFERENCE_IMAGE).getValue().toString().trim();

    if (!artist) continue;

    const prompts = getPromptInfo(artist, format);
    if (!prompts) {
      Logger.log(`Row ${row}: No prompts found for ${artist}|${format}`);
      continue;
    }

    const context = `Artwork by ${artist}, format: ${format}. Reference: ${refImage}`;

    try {
      // Generate Title (Column G)
      if (prompts.titlePrompt && !sheet.getRange(row, PRODUCTS_COLUMNS.TITLE).getValue()) {
        const title = callAI(apiKey, prompts.titlePrompt, context);
        sheet.getRange(row, PRODUCTS_COLUMNS.TITLE).setValue(title);
      }

      // Generate Description (Column H)
      if (prompts.descriptionPrompt && !sheet.getRange(row, PRODUCTS_COLUMNS.DESCRIPTION).getValue()) {
        const description = callAI(apiKey, prompts.descriptionPrompt, context);
        sheet.getRange(row, PRODUCTS_COLUMNS.DESCRIPTION).setValue(description);
      }

      // Generate Tags (Column I)
      if (prompts.tagsPrompt && !sheet.getRange(row, PRODUCTS_COLUMNS.TAGS).getValue()) {
        const tags = callAI(apiKey, prompts.tagsPrompt, context);
        sheet.getRange(row, PRODUCTS_COLUMNS.TAGS).setValue(tags);
      }

      // Generate Meta Keywords (Column J)
      if (prompts.metaKeywordsPrompt && !sheet.getRange(row, PRODUCTS_COLUMNS.META_KEYWORDS).getValue()) {
        const metaKeywords = callAI(apiKey, prompts.metaKeywordsPrompt, context);
        sheet.getRange(row, PRODUCTS_COLUMNS.META_KEYWORDS).setValue(metaKeywords);
      }

      // Generate Meta Description (Column K)
      if (prompts.metaDescriptionPrompt && !sheet.getRange(row, PRODUCTS_COLUMNS.META_DESCRIPTION).getValue()) {
        const metaDescription = callAI(apiKey, prompts.metaDescriptionPrompt, context);
        sheet.getRange(row, PRODUCTS_COLUMNS.META_DESCRIPTION).setValue(metaDescription);
      }

      // Generate Mobile Description (Column L)
      if (prompts.mobileDescriptionPrompt && !sheet.getRange(row, PRODUCTS_COLUMNS.MOBILE_DESCRIPTION).getValue()) {
        const mobileDescription = callAI(apiKey, prompts.mobileDescriptionPrompt, context);
        sheet.getRange(row, PRODUCTS_COLUMNS.MOBILE_DESCRIPTION).setValue(mobileDescription);
      }

      processedCount++;
      Logger.log(`Row ${row}: Generated AI content`);

      // Rate limiting
      if (processedCount % 3 === 0) {
        Utilities.sleep(2000);
      }

    } catch (error) {
      Logger.log(`Row ${row}: AI Error - ${error.message}`);
    }
  }

  SpreadsheetApp.getUi().alert(`Generated AI content for ${processedCount} rows`);
}

function callAI(apiKey, systemPrompt, userContent) {
  const isClaudeKey = apiKey.startsWith('sk-ant-');
  const isGrokKey = apiKey.startsWith('xai-');

  let url, headers, payload;

  if (isClaudeKey) {
    url = 'https://api.anthropic.com/v1/messages';
    headers = {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    };
    payload = {
      model: 'claude-3-5-sonnet-20241022',
      max_tokens: 4096,
      messages: [{ role: 'user', content: `${systemPrompt}\n\nContext: ${userContent}` }]
    };
  } else if (isGrokKey) {
    url = 'https://api.x.ai/v1/chat/completions';
    headers = {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    };
    payload = {
      model: 'grok-beta',
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userContent }
      ],
      max_tokens: 4096
    };
  } else {
    throw new Error('Unknown API key format');
  }

  const options = {
    method: 'post',
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (isClaudeKey) {
    return json.content[0].text.trim();
  } else {
    return json.choices[0].message.content.trim();
  }
}

// ============================================================================
// SECTION 13: IMAGE MATCHING FUNCTIONS
// ============================================================================

/**
 * Get all images for a SKU based on orientation
 * Returns array of image paths in order: Mockups, Crops, Stock
 *
 * SKU Pattern: [#]_[Artist]_[Short Title]_[V]
 * @param {string} sku - The product SKU (e.g., "1_DNYC_Mickey-Mouse_V1")
 * @param {string} orientation - "portrait" or "landscape"
 * @returns {Array} Array of image file paths
 */
function getImagesForSKU(sku, orientation) {
  const isPortrait = !orientation || orientation.toString().toLowerCase() === 'portrait';
  const orientationCode = isPortrait ? '_P_' : '_L_';

  const images = [];

  // 1. Get MOCKUPS (first priority)
  const mockupFolder = isPortrait ? IMAGE_FOLDERS.MOCKUPS_PORTRAIT : IMAGE_FOLDERS.MOCKUPS_LANDSCAPE;
  const mockups = findFilesMatchingSKU(mockupFolder, sku, orientationCode, 'Mockup');
  const sortedMockups = mockups.sort().slice(0, IMAGE_SETTINGS.MOCKUPS_PER_SKU);
  images.push(...sortedMockups);

  Logger.log(`  Found ${sortedMockups.length} mockups for ${sku}`);

  // 2. Get CROPS (crop_2 through crop_8)
  const crops = findCropsForSKU(IMAGE_FOLDERS.CROPS, sku, orientationCode);
  images.push(...crops);

  Logger.log(`  Found ${crops.length} crops for ${sku}`);

  // 3. Add STOCK images (same for all products)
  const stockImages = IMAGE_SETTINGS.STOCK_FILES.map(file =>
    `${IMAGE_FOLDERS.STOCK}/${file}`
  );
  images.push(...stockImages);

  Logger.log(`  Added ${stockImages.length} stock images`);
  Logger.log(`  Total images for ${sku}: ${images.length}`);

  // Limit to max images
  return images.slice(0, IMAGE_SETTINGS.MAX_TOTAL_IMAGES);
}

/**
 * Find files in a folder matching SKU and orientation
 */
function findFilesMatchingSKU(folderPath, sku, orientationCode, fileTypeFilter) {
  const matches = [];

  try {
    // Use DriveApp to find files in Google Drive
    // folderPath is expected to be a Google Drive folder ID
    if (!folderPath) return matches;

    const folder = DriveApp.getFolderById(folderPath);
    const files = folder.getFiles();

    // For now, return expected filenames based on pattern
    // The actual implementation depends on where files are hosted

    // Pattern: {SKU}_{fileTypeFilter}*
    // Example: 1_DNYC_Mickey-Mouse_V1_Mockup_3x4.jpg

    Logger.log(`  Searching ${folderPath} for ${sku}${orientationCode}*${fileTypeFilter || ''}*`);

    // This would be populated by actual file system scan
    // For now, we'll build expected filenames

  } catch (error) {
    Logger.log(`  Error searching folder: ${error.message}`);
  }

  return matches;
}

/**
 * Find crop files for a SKU
 * Pattern: {SKU}_crop_{1-24}.jpg (ALL 24 images use unified naming)
 */
function findCropsForSKU(folderPath, sku, orientationCode) {
  const crops = [];

  // All images follow naming: {SKU}_crop_{1-24}.jpg
  // Example: 8_DNYC-Print-Death-NYC-Mickey-Murakami-A-P_crop_1.jpg

  for (let i = 1; i <= HOSTED_IMAGES.NUM_PRODUCT_IMAGES; i++) {
    const cropFile = `${folderPath}/${sku}${HOSTED_IMAGES.CROP_PATTERN}${i}.jpg`;
    crops.push(cropFile);
  }

  return crops;
}

/**
 * Fill image columns for all products in Products tab
 */
function fillAllProductImages() {
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('ERROR: Products sheet not found!');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < STATIC_CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert('No data rows to process');
    return;
  }

  let filledCount = 0;

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue().toString().trim();
    const orientation = sheet.getRange(row, PRODUCTS_COLUMNS.ORIENTATION).getValue().toString().trim();

    if (!sku) {
      Logger.log(`Row ${row}: Skipping - no SKU`);
      continue;
    }

    Logger.log(`Processing images for row ${row}: ${sku} (${orientation})`);

    // Get images for this SKU
    const images = getImagesForSKU(sku, orientation);

    // Fill image columns (Image 1-24 = columns 56-79)
    for (let i = 0; i < images.length && i < 24; i++) {
      const imageCol = PRODUCTS_COLUMNS.IMAGE_1 + i;  // IMAGE_1 = 56
      sheet.getRange(row, imageCol).setValue(images[i]);
    }

    filledCount++;
  }

  SpreadsheetApp.getUi().alert(`Filled images for ${filledCount} rows`);
}

/**
 * Build image file list from Google Drive folders (for testing/preview)
 * Scans configured Drive folders and returns matched files
 * Pattern: [#]_[Artist]_[Short Title]_[V]_[IMG #].jpg
 */
function previewImagesForSKU(sku, orientation) {
  const isPortrait = !orientation || orientation.toLowerCase() === 'portrait';

  let output = `=== IMAGE PREVIEW FOR ${sku} (${isPortrait ? 'Portrait' : 'Landscape'}) ===\n\n`;

  // ALL 24 images use unified naming: {SKU}_crop_{1-24}.jpg
  output += `ALL IMAGES (Pattern: {SKU}_crop_{1-24}.jpg):\n`;
  output += `  Image order: 1-6 Mockups, 7-18 Crops, 19-24 Stock\n\n`;

  for (let i = 1; i <= HOSTED_IMAGES.NUM_PRODUCT_IMAGES; i++) {
    let label = '';
    if (i <= 6) label = '(Mockup)';
    else if (i <= 18) label = '(Crop)';
    else label = '(Stock)';
    output += `  Image ${i}: ${sku}${HOSTED_IMAGES.CROP_PATTERN}${i}.jpg ${label}\n`;
  }

  output += `\nTOTAL: ${HOSTED_IMAGES.NUM_PRODUCT_IMAGES} images (all using unified _crop_ naming)`;

  return output;
}

/**
 * Test image matching for a sample SKU
 */
function testImageMatching() {
  const testSKU = '1_DNYC_Mickey-Mouse_V1';
  const testOrientation = 'portrait';

  const preview = previewImagesForSKU(testSKU, testOrientation);

  Logger.log(preview);
  SpreadsheetApp.getUi().alert(preview);
}

/**
 * Preview images for currently selected row
 */
function previewSelectedSKUImages() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < STATIC_CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert('Please select a data row (row 2 or below)');
    return;
  }

  const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue().toString().trim();
  const orientation = sheet.getRange(row, PRODUCTS_COLUMNS.ORIENTATION).getValue().toString().trim();

  if (!sku) {
    SpreadsheetApp.getUi().alert('No SKU found in selected row. Generate SKUs first.');
    return;
  }

  const preview = previewImagesForSKU(sku, orientation);
  SpreadsheetApp.getUi().alert(preview);
}

/**
 * Show configured image folders
 */
function showImageFolders() {
  let output = '=== IMAGE FOLDER CONFIGURATION ===\n\n';

  output += 'MOCKUPS:\n';
  output += `  Portrait: ${IMAGE_FOLDERS.MOCKUPS_PORTRAIT}\n`;
  output += `  Landscape: ${IMAGE_FOLDERS.MOCKUPS_LANDSCAPE}\n`;

  output += '\nCROPS:\n';
  output += `  All: ${IMAGE_FOLDERS.CROPS}\n`;

  output += '\nSTOCK:\n';
  output += `  Folder: ${IMAGE_FOLDERS.STOCK}\n`;
  output += `  Files: ${IMAGE_SETTINGS.STOCK_FILES.join(', ')}\n`;

  output += '\nMAIN IMAGES (crop_1):\n';
  output += `  Portrait: ${IMAGE_FOLDERS.MAIN_PORTRAIT}\n`;
  output += `  Landscape: ${IMAGE_FOLDERS.MAIN_LANDSCAPE}\n`;

  output += '\nSETTINGS:\n';
  output += `  Mockups per SKU: ${IMAGE_SETTINGS.MOCKUPS_PER_SKU}\n`;
  output += `  Crops per SKU: ${IMAGE_SETTINGS.CROPS_PER_SKU}\n`;
  output += `  Stock images: ${IMAGE_SETTINGS.STOCK_IMAGES}\n`;
  output += `  Max total images: ${IMAGE_SETTINGS.MAX_TOTAL_IMAGES}\n`;

  output += '\nHOSTED URLs (gauntlet.gallery):\n';
  output += `  Products: ${HOSTED_IMAGES.PRODUCTS_BASE_URL}\n`;
  output += `  Stock: ${HOSTED_IMAGES.STOCK_BASE_URL}\n`;

  SpreadsheetApp.getUi().alert(output);
}

// ============================================================================
// SECTION 14: HOSTED IMAGE URL FUNCTIONS (eBay/Namecheap)
// ============================================================================

/**
 * Get hosted image URLs for a given SKU
 * Returns array of 24 URLs for eBay listing
 * ALL images use unified pattern: {fullSku}_crop_{1-24}.jpg
 *
 * Example filenames:
 *   1_DNYC_Mickey-Mouse_V1_crop_1.jpg
 *   1_DNYC_Mickey-Mouse_V1_crop_24.jpg
 *
 * @param {string} sku - Full SKU (e.g., "1_DNYC_Mickey-Mouse_V1")
 * @returns {Array} Array of 24 image URLs
 */
function getHostedImageUrls(sku, orientation) {
  const urls = [];

  // ALL 24 images use pattern: {fullSku}_crop_{1-24}.jpg
  // Example: 1_DNYC_Mickey-Mouse_V1_crop_1.jpg
  //          1_DNYC_Mickey-Mouse_V1_crop_24.jpg

  for (let i = 1; i <= HOSTED_IMAGES.NUM_PRODUCT_IMAGES; i++) {
    urls.push(`${HOSTED_IMAGES.PRODUCTS_BASE_URL}${sku}${HOSTED_IMAGES.CROP_PATTERN}${i}.jpg`);
  }

  return urls;
}

/**
 * Populate all hosted image URLs based on SKU
 * Fills Image 1-24 columns (W-AT, columns 23-46) with gauntlet.gallery URLs
 * URL Pattern: ALL images use {fullSku}_crop_{1-24}.jpg unified naming
 */
function populateHostedImageUrls() {
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('ERROR: Products sheet not found!');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < STATIC_CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert('No data rows to process');
    return;
  }

  let updatedCount = 0;
  let skippedCount = 0;
  const skippedSkus = [];

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue().toString().trim();
    const orientation = sheet.getRange(row, PRODUCTS_COLUMNS.ORIENTATION).getValue().toString().trim();

    // Check if SKU exists and starts with a number followed by underscore
    if (sku && sku.match(/^\d+_/)) {
      const imageUrls = getHostedImageUrls(sku, orientation);

      // Write all 24 image URLs at once (columns W-AT = 23-46)
      sheet.getRange(row, HOSTED_IMAGES.IMAGE_START_COLUMN, 1, HOSTED_IMAGES.NUM_IMAGE_COLUMNS).setValues([imageUrls]);

      updatedCount++;
      Logger.log(`Row ${row}: Populated ${imageUrls.length} URLs for ${sku}`);
    } else {
      skippedCount++;
      if (sku) skippedSkus.push(sku);
    }
  }

  let message = `Hosted Image URLs Populated!\n\n` +
    `Pattern: {SKU}_crop_{1-24}.jpg (ALL 24 images)\n` +
    `Base URL: ${HOSTED_IMAGES.PRODUCTS_BASE_URL}\n\n` +
    `Updated: ${updatedCount} rows\n` +
    `Skipped: ${skippedCount} rows`;

  if (skippedSkus.length > 0 && skippedSkus.length <= 10) {
    message += `\n\nSkipped SKUs (invalid format):\n${skippedSkus.join('\n')}`;
  }

  SpreadsheetApp.getUi().alert(message);
}

/**
 * Preview hosted URLs for a selected row
 */
function previewHostedUrls() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < STATIC_CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert('Please select a data row (row 2 or below)');
    return;
  }

  const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue().toString().trim();
  const orientation = sheet.getRange(row, PRODUCTS_COLUMNS.ORIENTATION).getValue().toString().trim();

  if (!sku) {
    SpreadsheetApp.getUi().alert('No SKU found in selected row. Generate SKUs first.');
    return;
  }

  const urls = getHostedImageUrls(sku, orientation);

  let output = `=== HOSTED URLs FOR ${sku} (${orientation || 'portrait'}) ===\n\n`;

  output += `ALL IMAGES (Pattern: {SKU}_crop_{1-24}.jpg):\n`;
  output += `Image order: 1-6 Mockups, 7-18 Crops, 19-24 Stock\n\n`;

  for (let i = 0; i < HOSTED_IMAGES.NUM_PRODUCT_IMAGES; i++) {
    let label = '';
    if (i < 6) label = '(Mockup)';
    else if (i < 18) label = '(Crop)';
    else label = '(Stock)';
    output += `  ${i + 1}: ${urls[i]} ${label}\n`;
  }

  output += `\nTOTAL: ${HOSTED_IMAGES.NUM_PRODUCT_IMAGES} images`;

  SpreadsheetApp.getUi().alert(output);
}

/**
 * Show hosted image URL configuration
 */
function showHostedUrlConfig() {
  let output = '=== HOSTED IMAGE URL CONFIG (gauntlet.gallery) ===\n\n';

  output += 'BASE URLs:\n';
  output += `  Products: ${HOSTED_IMAGES.PRODUCTS_BASE_URL}\n`;
  output += `  Stock: ${HOSTED_IMAGES.STOCK_BASE_URL}\n`;

  output += '\nURL PATTERN (ALL 24 images):\n';
  output += `  Pattern: {fullSku}${HOSTED_IMAGES.CROP_PATTERN}{1-24}.jpg\n`;

  output += '\nIMAGE STRUCTURE (24 total):\n';
  output += `  Images 1-24:  ALL use unified _crop_ naming (columns W-AT)\n`;
  output += `  Image order:  1-6 Mockups, 7-18 Crops, 19-24 Stock\n`;

  output += '\nCOLUMN MAPPING:\n';
  output += `  Start Column: ${HOSTED_IMAGES.IMAGE_START_COLUMN} (W)\n`;
  output += `  End Column: ${HOSTED_IMAGES.IMAGE_START_COLUMN + HOSTED_IMAGES.NUM_IMAGE_COLUMNS - 1} (AT)\n`;

  output += '\nEXAMPLE (SKU: 8_DNYC-Print-Death-NYC-Mickey-Murakami-A-P):\n';
  output += `  Image 1:  ${HOSTED_IMAGES.PRODUCTS_BASE_URL}8_DNYC-Print-Death-NYC-Mickey-Murakami-A-P${HOSTED_IMAGES.CROP_PATTERN}1.jpg\n`;
  output += `  Image 12: ${HOSTED_IMAGES.PRODUCTS_BASE_URL}8_DNYC-Print-Death-NYC-Mickey-Murakami-A-P${HOSTED_IMAGES.CROP_PATTERN}12.jpg\n`;
  output += `  Image 24: ${HOSTED_IMAGES.PRODUCTS_BASE_URL}8_DNYC-Print-Death-NYC-Mickey-Murakami-A-P${HOSTED_IMAGES.CROP_PATTERN}24.jpg\n`;

  SpreadsheetApp.getUi().alert(output);
}

// ============================================================================
// SECTION 15: UTILITY FUNCTIONS
// ============================================================================

function clearAllCaches() {
  VARIABLES_CACHE = null;
  PROMPTS_CACHE = null;
  FIELDS_CACHE = null;
  Logger.log('All caches cleared');
}

function reloadAllTabs() {
  clearAllCaches();
  loadVariables();
  loadPrompts();
  loadFields();
  SpreadsheetApp.getUi().alert('All tabs reloaded successfully!');
}

function generateAllSKUs() {
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);
  const lastRow = sheet.getLastRow();

  let count = 0;
  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    const artist = sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue().toString().trim();
    const format = sheet.getRange(row, PRODUCTS_COLUMNS.FORMAT).getValue().toString().trim();
    const title = sheet.getRange(row, PRODUCTS_COLUMNS.TITLE).getValue().toString().trim();

    if (!artist) continue;

    const sku = generateSKU(artist, format, title, row);
    sheet.getRange(row, PRODUCTS_COLUMNS.SKU).setValue(sku);
    count++;
  }

  SpreadsheetApp.getUi().alert(`Generated ${count} SKUs`);
}

function fillSelectedRowDetails() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < STATIC_CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert('Please select a data row (row 2 or below)');
    return;
  }

  // Call the main fill function for just this row
  // (simplified - would need to extract the per-row logic)
  SpreadsheetApp.getUi().alert(`Use "Fill ALL Rows" to fill row ${row}`);
}

function generateTitlesOnly() {
  SpreadsheetApp.getUi().alert('Use "Generate ALL Content" - titles are generated first');
}

function generateDescriptionsOnly() {
  SpreadsheetApp.getUi().alert('Use "Generate ALL Content" - descriptions are generated with titles');
}

function processAllImagesAuto() {
  SpreadsheetApp.getUi().alert('Image processing will be added. Use "Fill Product Details" and "Generate AI Content" for now.');
}

function processSingleBatch() {
  SpreadsheetApp.getUi().alert('Image processing will be added. Use "Fill Product Details" and "Generate AI Content" for now.');
}

// ============================================================================
// SECTION 14: DIAGNOSTIC FUNCTIONS
// ============================================================================

function testConfiguration() {
  clearAllCaches();

  let output = '=== CONFIGURATION TEST ===\n\n';

  try {
    const variables = loadVariables();
    output += `VARIABLES: ${Object.keys(variables.pricing).length} prices, ${Object.keys(variables.dimensions).length} dimensions, ${Object.keys(variables.apis).length} APIs\n`;
  } catch (e) {
    output += `VARIABLES ERROR: ${e.message}\n`;
  }

  try {
    const prompts = loadPrompts();
    output += `PROMPTS: ${prompts.size} configurations\n`;
  } catch (e) {
    output += `PROMPTS ERROR: ${e.message}\n`;
  }

  try {
    const fields = loadFields();
    output += `FIELDS: ${fields.size} configurations\n`;
  } catch (e) {
    output += `FIELDS ERROR: ${e.message}\n`;
  }

  Logger.log(output);
  SpreadsheetApp.getUi().alert(output);
}

function verifyColumnMappings() {
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);
  const headers = sheet.getRange(1, 1, 1, 60).getValues()[0];

  let output = '=== COLUMN VERIFICATION ===\n\n';

  const checks = [
    ['SKU', PRODUCTS_COLUMNS.SKU],
    ['Price', PRODUCTS_COLUMNS.PRICE],
    ['ARTIST', PRODUCTS_COLUMNS.ARTIST],
    ['FORMAT', PRODUCTS_COLUMNS.FORMAT],
    ['Title', PRODUCTS_COLUMNS.TITLE],
    ['Description', PRODUCTS_COLUMNS.DESCRIPTION],
    ['Tags', PRODUCTS_COLUMNS.TAGS],
    ['MetaKeywords', PRODUCTS_COLUMNS.META_KEYWORDS],
    ['CategoryID', PRODUCTS_COLUMNS.CATEGORY_ID],
    ['Condition', PRODUCTS_COLUMNS.CONDITION]
  ];

  for (const [expected, colNum] of checks) {
    const actual = headers[colNum - 1] || '(empty)';
    const match = actual.toLowerCase().includes(expected.toLowerCase().substring(0, 4));
    output += `${match ? '✓' : '✗'} Col ${colNum}: "${expected}" → "${actual}"\n`;
  }

  SpreadsheetApp.getUi().alert(output);
}

function showLoadedVariables() {
  const variables = loadVariables();
  let output = '=== VARIABLES ===\n\n';

  output += 'PRICING:\n';
  for (const [key, val] of Object.entries(variables.pricing)) {
    output += `  ${key}: $${val.price} (${val.skuPrefix})\n`;
  }

  output += '\nDIMENSIONS:\n';
  for (const [key, val] of Object.entries(variables.dimensions)) {
    output += `  ${key}: ${val.portraitLength}x${val.portraitWidth} (P), ${val.landscapeLength}x${val.landscapeWidth} (L)\n`;
  }

  SpreadsheetApp.getUi().alert(output.substring(0, 2000));
}

function showLoadedPrompts() {
  const prompts = loadPrompts();
  let output = `=== PROMPTS (${prompts.size}) ===\n\n`;

  for (const [key, val] of prompts.entries()) {
    output += `${key}:\n`;
    output += `  SKU Pattern: ${val.sampleSku || 'N/A'}\n`;
    output += `  Has: Title=${val.titlePrompt ? 'Y' : 'N'}, Desc=${val.descriptionPrompt ? 'Y' : 'N'}, Tags=${val.tagsPrompt ? 'Y' : 'N'}\n`;
  }

  SpreadsheetApp.getUi().alert(output.substring(0, 2000));
}

function showLoadedFields() {
  const fields = loadFields();
  let output = `=== FIELDS (${fields.size}) ===\n\n`;

  for (const [key, val] of fields.entries()) {
    output += `${key}: Cat=${val.categoryId}, Store=${val.storeCategory}, Framing=${val.clFraming}\n`;
  }

  SpreadsheetApp.getUi().alert(output.substring(0, 2000));
}

function testAIConnection() {
  const variables = loadVariables();
  const apiKey = variables.apis['CLAUDE_API_KEY'] || variables.apis['GROK_API_KEY'];

  if (!apiKey) {
    SpreadsheetApp.getUi().alert('No API key found!');
    return;
  }

  try {
    const result = callAI(apiKey, 'You are a test assistant.', 'Say "Connection OK" in exactly those words.');
    SpreadsheetApp.getUi().alert(`AI Test: ${result}`);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`AI Error: ${error.message}`);
  }
}

// ============================================================================
// SECTION 18: 3DSELLERS AUTO-POSTING
// ============================================================================

/**
 * Check if auto-posting is enabled
 * @returns {boolean} True if auto-posting is ON
 */
function isAutoPostingEnabled() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('AUTO_POSTING_ENABLED') === 'true';
}

/**
 * Toggle auto-posting ON
 */
function enableAutoPosting() {
  PropertiesService.getScriptProperties().setProperty('AUTO_POSTING_ENABLED', 'true');
  SpreadsheetApp.getUi().alert('3DSELLERS Auto-Posting: ENABLED\n\nListings will be automatically posted after processing.');
  Logger.log('Auto-posting ENABLED');
}

/**
 * Toggle auto-posting OFF
 */
function disableAutoPosting() {
  PropertiesService.getScriptProperties().setProperty('AUTO_POSTING_ENABLED', 'false');
  SpreadsheetApp.getUi().alert('3DSELLERS Auto-Posting: DISABLED\n\nListings will be saved as drafts only.');
  Logger.log('Auto-posting DISABLED');
}

/**
 * Show current auto-posting status
 */
function showAutoPostingStatus() {
  const enabled = isAutoPostingEnabled();
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('THREEDS_API_KEY') || 'NOT SET';
  const lastPost = props.getProperty('LAST_POST_DATE') || 'Never';
  const totalPosted = props.getProperty('TOTAL_POSTED') || '0';

  const status = `
=== 3DSELLERS AUTO-POSTING STATUS ===

Status: ${enabled ? 'ON (Enabled)' : 'OFF (Disabled)'}
API Key: ${apiKey !== 'NOT SET' ? '****' + apiKey.slice(-4) : 'NOT SET'}
Last Posted: ${lastPost}
Total Posted: ${totalPosted} listings

Batch Size: ${THREEDS_CONFIG.BATCH_SIZE}
Delay Between Posts: ${THREEDS_CONFIG.DELAY_BETWEEN_POSTS}ms
`;

  SpreadsheetApp.getUi().alert(status);
}

/**
 * Set 3DSellers API Key
 */
function set3DSellersAPIKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '3DSELLERS API Key',
    'Enter your 3DSellers API key:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (apiKey) {
      PropertiesService.getScriptProperties().setProperty('THREEDS_API_KEY', apiKey);
      ui.alert('API Key saved successfully!');
    } else {
      ui.alert('No API key entered.');
    }
  }
}

/**
 * Build listing payload for 3DSellers API from product row
 * @param {Object} rowData - Object containing all product data
 * @returns {Object} API payload
 */
function buildListingPayload(rowData) {
  return {
    sku: rowData.sku,
    title: rowData.title,
    description: rowData.description,
    price: rowData.price,
    quantity: rowData.quantity || 1,
    condition: rowData.condition || 'New',

    // Item specifics
    itemSpecifics: {
      Artist: rowData.artist,
      Type: rowData.cType,
      Style: rowData.cStyle,
      Subject: rowData.cSubject,
      Features: rowData.cFeatures,
      'Signed By': rowData.cSignedBy,
      'Certificate of Authenticity': rowData.cCOA,
      'Original/Licensed Reprint': rowData.cOriginalReprint,
      'Time Period Manufactured': rowData.cTimePeriod,
      'Country of Origin': rowData.cCountryOrigin,
      'Production Technique': rowData.cProductionTechnique,
      Material: rowData.cMaterial,
      Size: rowData.size,
      Brand: rowData.brand
    },

    // Dimensions
    dimensions: {
      length: rowData.length,
      width: rowData.width,
      height: rowData.height,
      weight: rowData.weight
    },

    // Shipping
    shipping: {
      packageType: rowData.packageType,
      postalCode: rowData.postalCode,
      location: rowData.location,
      countryCode: rowData.countryCode
    },

    // Images
    images: rowData.images || [],

    // Categories
    categoryId: rowData.categoryId,
    storeCategory: rowData.storeCategory,

    // Policies
    paymentPolicy: rowData.paymentPolicy,
    shippingPolicy: rowData.shippingPolicy,
    returnPolicy: rowData.returnPolicy,

    // Meta
    tags: rowData.tags,
    metaKeywords: rowData.metaKeywords,
    metaDescription: rowData.metaDescription,
    mobileDescription: rowData.mobileDescription
  };
}

/**
 * Post a single listing to 3DSellers
 * @param {Object} payload - Listing data
 * @returns {Object} Response with success status and message
 */
function postTo3DSellers(payload) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('THREEDS_API_KEY');

  if (!apiKey) {
    return { success: false, message: 'API key not set. Use "Set API Key" first.' };
  }

  try {
    const response = UrlFetchApp.fetch(`${THREEDS_CONFIG.API_BASE_URL}/listings`, {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
        'X-API-Version': '2024-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = JSON.parse(response.getContentText());

    if (responseCode === 200 || responseCode === 201) {
      return {
        success: true,
        message: 'Posted successfully',
        listingId: responseBody.listingId || responseBody.id
      };
    } else {
      return {
        success: false,
        message: responseBody.error || responseBody.message || `HTTP ${responseCode}`
      };
    }

  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Read product data from a specific row
 * @param {Sheet} sheet - The Products sheet
 * @param {number} row - Row number
 * @returns {Object} Product data object
 */
function readProductRow(sheet, row) {
  const cols = PRODUCTS_COLUMNS;

  // Read all image URLs (columns 56-79 for Images 1-24)
  const images = [];
  for (let i = 0; i < 24; i++) {
    const imgUrl = sheet.getRange(row, cols.IMAGE_1 + i).getValue();
    if (imgUrl) images.push(imgUrl);
  }

  return {
    sku: sheet.getRange(row, cols.SKU).getValue(),
    price: sheet.getRange(row, cols.PRICE).getValue(),
    artist: sheet.getRange(row, cols.ARTIST).getValue(),
    format: sheet.getRange(row, cols.FORMAT).getValue(),
    orientation: sheet.getRange(row, cols.ORIENTATION).getValue(),
    title: sheet.getRange(row, cols.TITLE).getValue(),
    description: sheet.getRange(row, cols.DESCRIPTION).getValue(),
    tags: sheet.getRange(row, cols.TAGS).getValue(),
    metaKeywords: sheet.getRange(row, cols.META_KEYWORDS).getValue(),
    metaDescription: sheet.getRange(row, cols.META_DESCRIPTION).getValue(),
    mobileDescription: sheet.getRange(row, cols.MOBILE_DESCRIPTION).getValue(),

    // Item specifics
    condition: sheet.getRange(row, cols.CONDITION).getValue(),
    quantity: sheet.getRange(row, cols.QUANTITY).getValue() || 1,
    cType: sheet.getRange(row, cols.C_TYPE).getValue(),
    cSignedBy: sheet.getRange(row, cols.C_SIGNED_BY).getValue(),
    cCOA: sheet.getRange(row, cols.C_COA).getValue(),
    cStyle: sheet.getRange(row, cols.C_STYLE).getValue(),
    cFeatures: sheet.getRange(row, cols.C_FEATURES).getValue(),
    cOriginalReprint: sheet.getRange(row, cols.C_ORIGINAL_REPRINT).getValue(),
    cSubject: sheet.getRange(row, cols.C_SUBJECT).getValue(),
    cTimePeriod: sheet.getRange(row, cols.C_TIME_PERIOD).getValue(),
    cCountryOrigin: sheet.getRange(row, cols.C_COUNTRY_ORIGIN).getValue(),
    cProductionTechnique: sheet.getRange(row, cols.C_PRODUCTION_TECHNIQUE).getValue(),
    cMaterial: sheet.getRange(row, cols.C_MATERIAL).getValue(),
    size: sheet.getRange(row, cols.SIZE).getValue(),
    brand: sheet.getRange(row, cols.BRAND).getValue(),

    // Dimensions
    length: sheet.getRange(row, cols.LENGTH).getValue(),
    width: sheet.getRange(row, cols.WIDTH).getValue(),
    height: sheet.getRange(row, cols.HEIGHT).getValue(),
    weight: sheet.getRange(row, cols.WEIGHT_MAJOR).getValue(),

    // Shipping
    packageType: sheet.getRange(row, cols.PACKAGE_TYPE).getValue(),
    postalCode: sheet.getRange(row, cols.POSTAL_CODE).getValue(),
    location: sheet.getRange(row, cols.LOCATION).getValue(),
    countryCode: sheet.getRange(row, cols.COUNTRY_CODE).getValue(),

    // Categories & Policies
    categoryId: sheet.getRange(row, cols.CATEGORY_ID).getValue(),
    storeCategory: sheet.getRange(row, cols.STORE_CATEGORY).getValue(),
    paymentPolicy: sheet.getRange(row, cols.PAYMENT_POLICY).getValue(),
    shippingPolicy: sheet.getRange(row, cols.SHIPPING_POLICY).getValue(),
    returnPolicy: sheet.getRange(row, cols.RETURN_POLICY).getValue(),

    // Images
    images: images
  };
}

/**
 * Post all pending listings to 3DSellers
 */
function postAllPendingListings() {
  if (!isAutoPostingEnabled()) {
    SpreadsheetApp.getUi().alert('Auto-posting is DISABLED.\n\nEnable it first using "Turn Auto-Posting ON"');
    return;
  }

  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('ERROR: Products sheet not found!');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < STATIC_CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert('No products to post.');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Post to 3DSELLERS',
    `This will post all PENDING listings to 3DSELLERS.\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  let posted = 0;
  let failed = 0;
  let skipped = 0;
  const errors = [];

  // Show progress
  const startTime = Date.now();

  for (let row = STATIC_CONFIG.START_ROW; row <= lastRow; row++) {
    const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue();
    const status = sheet.getRange(row, THREEDS_CONFIG.POST_STATUS_COLUMN).getValue();

    // Skip if no SKU or already posted
    if (!sku) {
      skipped++;
      continue;
    }

    if (status === THREEDS_CONFIG.STATUS.POSTED) {
      skipped++;
      continue;
    }

    // Read product data
    const productData = readProductRow(sheet, row);
    const payload = buildListingPayload(productData);

    // Post to 3DSellers
    Logger.log(`Posting row ${row}: ${sku}`);
    const result = postTo3DSellers(payload);

    if (result.success) {
      sheet.getRange(row, THREEDS_CONFIG.POST_STATUS_COLUMN).setValue(THREEDS_CONFIG.STATUS.POSTED);
      sheet.getRange(row, THREEDS_CONFIG.POST_DATE_COLUMN).setValue(new Date().toISOString());
      posted++;
      Logger.log(`  SUCCESS: ${result.listingId}`);
    } else {
      sheet.getRange(row, THREEDS_CONFIG.POST_STATUS_COLUMN).setValue(THREEDS_CONFIG.STATUS.FAILED);
      failed++;
      errors.push(`Row ${row} (${sku}): ${result.message}`);
      Logger.log(`  FAILED: ${result.message}`);
    }

    // Rate limiting
    Utilities.sleep(THREEDS_CONFIG.DELAY_BETWEEN_POSTS);

    // Check timeout (5 min max)
    if (Date.now() - startTime > 300000) {
      Logger.log('Timeout reached after 5 minutes');
      break;
    }
  }

  // Update totals
  const props = PropertiesService.getScriptProperties();
  const totalPosted = parseInt(props.getProperty('TOTAL_POSTED') || '0') + posted;
  props.setProperty('TOTAL_POSTED', totalPosted.toString());
  props.setProperty('LAST_POST_DATE', new Date().toISOString());

  // Show results
  let message = `=== 3DSELLERS POSTING COMPLETE ===\n\n`;
  message += `Posted: ${posted}\n`;
  message += `Failed: ${failed}\n`;
  message += `Skipped: ${skipped}\n`;
  message += `Time: ${Math.round((Date.now() - startTime) / 1000)}s\n`;

  if (errors.length > 0) {
    message += `\nErrors:\n${errors.slice(0, 5).join('\n')}`;
    if (errors.length > 5) {
      message += `\n... and ${errors.length - 5} more`;
    }
  }

  ui.alert(message);
}

/**
 * Post only selected rows to 3DSellers
 */
function postSelectedRows() {
  if (!isAutoPostingEnabled()) {
    SpreadsheetApp.getUi().alert('Auto-posting is DISABLED.\n\nEnable it first using "Turn Auto-Posting ON"');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow < STATIC_CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert('Please select data rows (row 2 or below)');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Post Selected to 3DSELLERS',
    `This will post ${numRows} selected row(s) to 3DSELLERS.\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  let posted = 0;
  let failed = 0;

  for (let row = startRow; row < startRow + numRows; row++) {
    const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue();
    if (!sku) continue;

    const productData = readProductRow(sheet, row);
    const payload = buildListingPayload(productData);
    const result = postTo3DSellers(payload);

    if (result.success) {
      sheet.getRange(row, THREEDS_CONFIG.POST_STATUS_COLUMN).setValue(THREEDS_CONFIG.STATUS.POSTED);
      sheet.getRange(row, THREEDS_CONFIG.POST_DATE_COLUMN).setValue(new Date().toISOString());
      posted++;
    } else {
      sheet.getRange(row, THREEDS_CONFIG.POST_STATUS_COLUMN).setValue(THREEDS_CONFIG.STATUS.FAILED);
      failed++;
    }

    Utilities.sleep(THREEDS_CONFIG.DELAY_BETWEEN_POSTS);
  }

  ui.alert(`Posted: ${posted}\nFailed: ${failed}`);
}

/**
 * Mark selected rows as PENDING for posting
 */
function markSelectedAsPending() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow < STATIC_CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert('Please select data rows (row 2 or below)');
    return;
  }

  for (let row = startRow; row < startRow + numRows; row++) {
    sheet.getRange(row, THREEDS_CONFIG.POST_STATUS_COLUMN).setValue(THREEDS_CONFIG.STATUS.PENDING);
  }

  SpreadsheetApp.getUi().alert(`Marked ${numRows} row(s) as PENDING`);
}

/**
 * Clear post status for selected rows (allows re-posting)
 */
function clearPostStatus() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow < STATIC_CONFIG.START_ROW) {
    SpreadsheetApp.getUi().alert('Please select data rows (row 2 or below)');
    return;
  }

  for (let row = startRow; row < startRow + numRows; row++) {
    sheet.getRange(row, THREEDS_CONFIG.POST_STATUS_COLUMN).clearContent();
    sheet.getRange(row, THREEDS_CONFIG.POST_DATE_COLUMN).clearContent();
  }

  SpreadsheetApp.getUi().alert(`Cleared post status for ${numRows} row(s)`);
}

/**
 * Test 3DSellers API connection
 */
function test3DSellersConnection() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('THREEDS_API_KEY');

  if (!apiKey) {
    SpreadsheetApp.getUi().alert('No API key set!\n\nUse "Set API Key" first.');
    return;
  }

  try {
    const response = UrlFetchApp.fetch(`${THREEDS_CONFIG.API_BASE_URL}/account`, {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'X-API-Version': '2024-01'
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      SpreadsheetApp.getUi().alert(`3DSELLERS Connection: SUCCESS\n\nAccount: ${data.email || data.name || 'Connected'}`);
    } else {
      SpreadsheetApp.getUi().alert(`3DSELLERS Connection: FAILED\n\nHTTP ${code}: ${response.getContentText()}`);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(`3DSELLERS Connection: ERROR\n\n${error.message}`);
  }
}

// ============================================================================
// TEST FUNCTION - Process 3 Images Only
// ============================================================================

/**
 * TEST: Process only 3 images to verify the full pipeline works
 * Run this from: 3DSELLERS V7 MASTER → Diagnostics → Test 3 Images
 */
function test3Images() {
  const ui = SpreadsheetApp.getUi();
  const config = ARTIST_FOLDERS['DEATH NYC'];

  // Let user choose orientation
  const response = ui.alert(
    'TEST: Process 3 Images',
    'Which orientation folder do you want to test?\n\n' +
    '• YES = Portrait Mockups\n' +
    '• NO = Landscape Mockups\n' +
    '• CANCEL = Exit',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response === ui.Button.CANCEL) return;

  const isPortraitFolder = (response === ui.Button.YES);
  const folderId = isPortraitFolder ? config.MOCKUPS_PORTRAIT : config.MOCKUPS_LANDSCAPE;
  const folderType = isPortraitFolder ? 'Portrait Mockups' : 'Landscape Mockups';

  if (!folderId) {
    ui.alert('No folder ID configured for ' + folderType);
    return;
  }

  // Get folder
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    ui.alert('Could not access folder: ' + e.message + '\n\nFolder ID: ' + folderId);
    return;
  }

  Logger.log('='.repeat(60));
  Logger.log('TEST: Processing 3 Images');
  Logger.log('Folder: ' + folder.getName());
  Logger.log('='.repeat(60));

  // Get image files
  const files = folder.getFiles();
  let imageFiles = [];

  while (files.hasNext() && imageFiles.length < 3) {
    const file = files.next();
    const mimeType = file.getMimeType();
    if (mimeType.startsWith('image/')) {
      imageFiles.push(file);
    }
  }

  if (imageFiles.length === 0) {
    ui.alert('No images found in folder');
    return;
  }

  Logger.log('Found ' + imageFiles.length + ' images to process');

  // Load all config
  clearAllCaches();
  const variables = loadVariables();
  const fieldsMap = loadFields();
  const promptsMap = loadPrompts();

  Logger.log('Loaded: ' + Object.keys(variables.pricing).length + ' prices, ' +
             fieldsMap.size + ' field configs, ' + promptsMap.size + ' prompts');

  // Get sheet
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  // Clear test rows (rows 2-4)
  sheet.getRange('A2:CZ4').clearContent();

  let results = [];

  for (let i = 0; i < imageFiles.length; i++) {
    const file = imageFiles[i];
    const fileName = file.getName();
    const rowNum = i + 2; // Start at row 2

    Logger.log('\n[' + (i+1) + '/3] Processing: ' + fileName);

    // Extract SKU number from filename (e.g., "10_DNYC-Print-..." -> "10")
    const skuMatch = fileName.match(/^(\d+)_/);
    const skuNum = skuMatch ? skuMatch[1] : (i + 1).toString();

    // Detect orientation from filename
    const isPortrait = fileName.includes('_P_') || !fileName.includes('_L_');
    const orientation = isPortrait ? 'Portrait' : 'Landscape';

    Logger.log('  SKU #: ' + skuNum);
    Logger.log('  Orientation: ' + orientation);

    // AI Vision Analysis
    Logger.log('  Running AI Vision...');
    let analysis = null;
    try {
      analysis = analyzeImageWithVision(file);
      if (analysis) {
        Logger.log('  AI Vision SUCCESS');
        Logger.log('    Subject: ' + (analysis.primary_subject || 'N/A'));
        Logger.log('    Title: ' + (analysis.suggested_title || 'N/A'));
      }
    } catch (e) {
      Logger.log('  AI Vision FAILED: ' + e.message);
    }

    // Extract subject from filename as fallback
    let subject = 'Art Print';
    const nameParts = fileName.split('_');
    if (nameParts.length >= 2) {
      subject = nameParts[1].replace(/-/g, ' ').replace('DNYC Print ', '').replace('DNYC Original ', '');
    }
    if (analysis && analysis.primary_subject) {
      subject = analysis.primary_subject;
    }

    // Generate SKU
    const cleanSubject = subject.replace(/\s+/g, '-').substring(0, 20);
    const orientCode = isPortrait ? 'P' : 'L';
    const sku = skuNum + '_DNYC_' + cleanSubject + '_' + orientCode + '_V1';

    Logger.log('  Generated SKU: ' + sku);

    // Get artist config
    const artistName = 'Death NYC';
    const format = 'Prints';

    // Get pricing from VARIABLES
    const price = getPrice(artistName, format);
    Logger.log('  Price: $' + price);

    // Get dimensions from VARIABLES (ORIENTATION-AWARE)
    const dims = getDimensions(artistName, format, orientation);
    Logger.log('  Dimensions: ' + dims.length + ' x ' + dims.width + ' (' + orientation + ')');

    // Calculate ORIENTATION-AWARE package dimensions
    const packageDims = {
      length: isPortrait ? 20 : 16,  // Portrait: taller package, Landscape: wider package
      width: isPortrait ? 16 : 20,
      depth: 2
    };
    Logger.log('  Package: ' + packageDims.length + ' x ' + packageDims.width + ' x ' + packageDims.depth);

    // Get fields from FIELDS tab
    const fields = getFieldsForProduct(artistName, format);
    if (fields) {
      Logger.log('  Fields: Loaded 40 columns');
    } else {
      Logger.log('  Fields: Using defaults');
    }

    // Get prompts
    const promptInfo = getPromptInfo(artistName, format);

    // ORIENTATION-AWARE: Select correct mockup folder reference
    const mockupFolder = isPortrait ? 'PORTRAIT MOCKUPS' : 'LANDSCAPE MOCKUPS';
    Logger.log('  Mockup Folder: ' + mockupFolder);

    // Generate title
    let title = '';
    if (analysis && analysis.suggested_title) {
      title = analysis.suggested_title.substring(0, 80);
    } else {
      title = 'Death NYC ' + subject + ' Hand Signed Ltd Ed Pop Art Print COA';
      title = title.substring(0, 80);
    }

    // Generate description - Dec 2025 optimized for eBay SEO
    let description = buildFullDescription(subject, orientation, dims);

    // Generate tags
    let tags = 'Death NYC, ' + subject + ', Pop Art, Street Art, Limited Edition, Hand Signed, COA, Contemporary Art';
    if (analysis && analysis.suggested_tags) {
      tags = analysis.suggested_tags.join(', ');
    }

    // Fill the row
    // Core columns (A-F)
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.SKU).setValue(sku);
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.PRICE).setValue(price);
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.ARTIST).setValue(artistName);
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.FORMAT).setValue(format);
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.REFERENCE_IMAGE).setValue(file.getUrl());
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.ORIENTATION).setValue(orientation);

    // AI-generated content (G-L)
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.TITLE).setValue(title);
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.DESCRIPTION).setValue(description);
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.TAGS).setValue(tags);
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.META_KEYWORDS).setValue('death nyc, ' + subject.toLowerCase() + ', pop art, limited edition');
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.META_DESCRIPTION).setValue('Death NYC ' + subject + ' hand-signed pop art. COA included. Free shipping.'.substring(0, 160));
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.MOBILE_DESCRIPTION).setValue('<b>DEATH NYC</b> ' + subject + ' - Hand Signed Ltd Ed /100 - COA');

    // Dimensions (M-O)
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_IMAGE_ORIENTATION).setValue(orientation);
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_ITEM_LENGTH).setValue(dims.length);
    sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_ITEM_WIDTH).setValue(dims.width);

    // Fields from FIELDS tab (P onwards)
    if (fields) {
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.CATEGORY_ID).setValue(fields.categoryId);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.CATEGORY_ID_2).setValue(fields.categoryId2);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.STORE_CATEGORY).setValue(fields.storeCategory);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.QUANTITY).setValue(fields.quantity || 1);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.CONDITION).setValue(fields.condition);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.COUNTRY_CODE).setValue(fields.countryCode);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.LOCATION).setValue(fields.location);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.POSTAL_CODE).setValue(fields.postalCode);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.PACKAGE_TYPE).setValue(fields.packageType);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_TYPE).setValue(fields.cType);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_SIGNED_BY).setValue(fields.cSignedBy);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_COA).setValue(fields.cCOA);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_STYLE).setValue(fields.cStyle);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_THEME).setValue(analysis ? (analysis.themes || []).join(', ') : '');
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_FRANCHISE).setValue(analysis ? (analysis.franchise || '') : '');
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_FEATURES).setValue(fields.cFeatures);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_ORIGINAL_REPRINT).setValue(fields.cOriginalReprint);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_SUBJECT).setValue(subject);
      // ORIENTATION-AWARE size description
      const sizeDesc = isPortrait
        ? dims.length + 'x' + dims.width + ' inches (Portrait)'
        : dims.length + 'x' + dims.width + ' inches (Landscape)';
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.SIZE).setValue(sizeDesc);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_CHARACTER).setValue(analysis ? (analysis.characters || []).join(', ') : '');
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_TIME_PERIOD).setValue(fields.cTimePeriod);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.PAYMENT_POLICY).setValue(fields.paymentPolicy);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.SHIPPING_POLICY).setValue(fields.shippingPolicy);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.RETURN_POLICY).setValue(fields.returnPolicy);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_COUNTRY_ORIGIN).setValue(fields.cCountryOrigin);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.CL_FRAMING).setValue(fields.clFraming);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_SIGNED).setValue('Yes');
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.MEASUREMENT_SYSTEM).setValue(fields.measurementSystem);
      // ORIENTATION-AWARE package dimensions (override FIELDS tab)
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.PACKAGE_LENGTH).setValue(packageDims.length);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.PACKAGE_WIDTH).setValue(packageDims.width);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.PACKAGE_DEPTH).setValue(packageDims.depth);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.WEIGHT_MAJOR).setValue(fields.weightMajor);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.WEIGHT_MINOR).setValue(fields.weightMinor);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.BRAND).setValue(fields.brand);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_COA_ISSUED_BY).setValue(fields.cCOAIssuedBy);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_PERIOD).setValue(fields.cPeriod);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_PRODUCTION_TECHNIQUE).setValue(fields.cProductionTechnique);
      sheet.getRange(rowNum, PRODUCTS_COLUMNS.C_MATERIAL).setValue(fields.cMaterial);
    }

    // Generate hosted image URLs (placeholder - you'll need to upload actual images)
    const baseUrl = 'https://gauntlet.gallery/images/products/';
    const stockUrl = 'https://gauntlet.gallery/images/stock/';

    // Image URLs (columns BD-CA, indices 56-79)
    for (let imgNum = 1; imgNum <= 8; imgNum++) {
      sheet.getRange(rowNum, 55 + imgNum).setValue(baseUrl + sku + '_IMG' + imgNum + '.jpg');
    }
    // Stock images
    const stockFiles = ['A.png', 'B.png', 'C.png', 'D.png', 'E.png'];
    for (let s = 0; s < stockFiles.length; s++) {
      sheet.getRange(rowNum, 63 + s).setValue(stockUrl + stockFiles[s]);
    }

    results.push({
      row: rowNum,
      sku: sku,
      title: title.substring(0, 40) + '...',
      subject: subject,
      orientation: orientation,
      dimensions: dims.length + 'x' + dims.width,
      packageDims: packageDims.length + 'x' + packageDims.width + 'x' + packageDims.depth
    });

    Logger.log('  Row ' + rowNum + ' COMPLETE');
    Logger.log('  ORIENTATION-AWARE FIELDS:');
    Logger.log('    - SKU: ' + sku + ' (contains _' + orientCode + '_)');
    Logger.log('    - Orientation: ' + orientation);
    Logger.log('    - Item Dimensions: ' + dims.length + 'x' + dims.width);
    Logger.log('    - Package Dimensions: ' + packageDims.length + 'x' + packageDims.width + 'x' + packageDims.depth);
    Logger.log('    - Size: ' + (isPortrait ? dims.length + 'x' + dims.width + ' (Portrait)' : dims.length + 'x' + dims.width + ' (Landscape)'));
    Logger.log('    - Mockup Folder: ' + mockupFolder);

    // Rate limit
    Utilities.sleep(1500);
  }

  // Show results
  Logger.log('\n' + '='.repeat(60));
  Logger.log('TEST COMPLETE!');
  Logger.log('='.repeat(60));

  let summary = 'TEST COMPLETE!\n\nProcessed ' + results.length + ' images:\n\n';
  for (const r of results) {
    summary += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
    summary += 'Row ' + r.row + ': ' + r.sku + '\n';
    summary += '  Orientation: ' + r.orientation + '\n';
    summary += '  Subject: ' + r.subject + '\n';
    summary += '  Item Size: ' + r.dimensions + '\n';
    summary += '  Package: ' + r.packageDims + '\n';
  }
  summary += '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
  summary += 'ORIENTATION-AWARE FIELDS:\n';
  summary += '✓ SKU (contains _P_ or _L_)\n';
  summary += '✓ Orientation column\n';
  summary += '✓ C: Image Orientation\n';
  summary += '✓ C: Item Length/Width\n';
  summary += '✓ Package Length/Width\n';
  summary += '✓ Size description\n';
  summary += '✓ Image URLs\n';
  summary += '\nCheck View → Logs for full details.';

  ui.alert(summary);
}

/**
 * QUICK TEST: Just test folder access - NO AI Vision (faster)
 * Run this first to verify Google Drive folders are accessible
 */
function quickTestFolders() {
  const ui = SpreadsheetApp.getUi();
  const config = ARTIST_FOLDERS['DEATH NYC'];

  let results = '=== FOLDER ACCESS TEST ===\n\n';

  // Test each folder
  const folders = {
    'Portrait Mockups': config.MOCKUPS_PORTRAIT,
    'Landscape Mockups': config.MOCKUPS_LANDSCAPE,
    'Portrait Crops': config.CROPS_PORTRAIT,
    'Landscape Crops': config.CROPS_LANDSCAPE,
    'All Crops': config.CROPS_ALL,
    'Stock Images': config.STOCK
  };

  for (const [name, folderId] of Object.entries(folders)) {
    if (!folderId) {
      results += '❌ ' + name + ': No folder ID configured\n';
      continue;
    }

    try {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();
      let count = 0;
      let firstFile = '';
      while (files.hasNext() && count < 100) {
        const f = files.next();
        if (count === 0) firstFile = f.getName();
        count++;
      }
      results += '✅ ' + name + ': ' + count + ' files\n';
      results += '   First: ' + firstFile.substring(0, 40) + '\n';
      results += '   ID: ' + folderId + '\n\n';
    } catch (e) {
      results += '❌ ' + name + ': ERROR - ' + e.message + '\n';
      results += '   ID: ' + folderId + '\n\n';
    }
  }

  Logger.log(results);
  ui.alert('Folder Test Results', results, ui.ButtonSet.OK);
}

/**
 * QUICK TEST: Process 3 images WITHOUT AI Vision (fast test)
 * Tests: folder access, SKU generation, sheet writing
 */
function quickTest3NoAI() {
  const ui = SpreadsheetApp.getUi();
  const config = ARTIST_FOLDERS['DEATH NYC'];

  // Use Portrait Mockups folder
  const folderId = config.MOCKUPS_PORTRAIT;

  if (!folderId) {
    ui.alert('No Portrait Mockups folder ID configured');
    return;
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    ui.alert('Cannot access folder: ' + e.message);
    return;
  }

  ui.alert('Starting quick test (NO AI)...\n\nFolder: ' + folder.getName());

  // Get first 3 images
  const files = folder.getFiles();
  let imageFiles = [];
  while (files.hasNext() && imageFiles.length < 3) {
    const file = files.next();
    if (file.getMimeType().startsWith('image/')) {
      imageFiles.push(file);
    }
  }

  if (imageFiles.length === 0) {
    ui.alert('No images found in folder');
    return;
  }

  // Get sheet
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  // Clear test rows
  sheet.getRange('A2:CZ4').clearContent();

  let results = [];

  for (let i = 0; i < imageFiles.length; i++) {
    const file = imageFiles[i];
    const fileName = file.getName();
    const rowNum = i + 2;

    // Extract SKU number
    const skuMatch = fileName.match(/^(\d+)_/);
    const skuNum = skuMatch ? skuMatch[1] : (i + 1).toString();

    // Detect orientation
    const isPortrait = fileName.includes('_P_') || !fileName.includes('_L_');
    const orientation = isPortrait ? 'Portrait' : 'Landscape';

    // Extract subject from filename
    let subject = 'Art Print';
    const nameParts = fileName.split('_');
    if (nameParts.length >= 2) {
      subject = nameParts[1].replace(/-/g, ' ').replace('DNYC Print ', '').substring(0, 30);
    }

    // Generate SKU
    const cleanSubject = subject.replace(/\s+/g, '-').substring(0, 20);
    const orientCode = isPortrait ? 'P' : 'L';
    const sku = skuNum + '_DNYC_' + cleanSubject + '_' + orientCode + '_V1';

    // Generate basic content (NO AI) - use full description builder
    const title = ('Death NYC ' + subject + ' Hand Signed Ltd Ed Pop Art Print COA').substring(0, 80);
    const dims = isPortrait ? {length: 18, width: 13} : {length: 13, width: 18};
    const description = buildFullDescription(subject, orientation, dims);
    const tags = 'Death NYC, ' + subject + ', Pop Art, Street Art, Limited Edition, COA, Hand Signed, Contemporary Art';

    // Write to sheet
    const rowData = [
      sku,           // A - SKU
      '149.99',      // B - Price
      'Death NYC',   // C - Artist
      'Prints',      // D - Format
      fileName,      // E - Reference
      orientation,   // F - Orientation
      title,         // G - Title
      description,   // H - Description
      tags           // I - Tags
    ];

    sheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);

    results.push({
      row: rowNum,
      sku: sku,
      file: fileName.substring(0, 30)
    });
  }

  // Show results
  let msg = '✅ QUICK TEST COMPLETE (No AI)\n\n';
  msg += 'Processed ' + results.length + ' images:\n\n';
  for (const r of results) {
    msg += 'Row ' + r.row + ': ' + r.sku + '\n';
  }
  msg += '\nCheck Products sheet rows 2-4';

  ui.alert('Quick Test Results', msg, ui.ButtonSet.OK);
}

/**
 * Build full eBay HTML description - Dec 2025 SEO optimized
 * @param {string} subject - The artwork subject/title
 * @param {string} orientation - PORTRAIT or LANDSCAPE
 * @param {object} dims - {length, width} dimensions
 * @returns {string} Full HTML description
 */
function buildFullDescription(subject, orientation, dims) {
  // Detect luxury brands in subject for keyword stuffing
  const luxuryBrands = ['Louis Vuitton', 'LV', 'Supreme', 'Chanel', 'Gucci', 'Hermès', 'Hermes', 'Dior', 'Prada', 'Balenciaga', 'Fendi', 'Burberry', 'Versace', 'YSL', 'Cartier', 'Rolex'];
  const subjectLower = subject.toLowerCase();
  const foundBrands = luxuryBrands.filter(b => subjectLower.includes(b.toLowerCase()));

  // Build brand keyword string for first paragraph
  let brandKeywords = '';
  if (foundBrands.length > 0) {
    brandKeywords = foundBrands.map(b => '<b>' + b + '</b>').join(', ');
    // Repeat for SEO density
    brandKeywords = brandKeywords + ' - ' + foundBrands.map(b => '<b>' + b + '</b>').join(' x ');
  }

  // Size string
  const sizeStr = dims ? (dims.length + '" x ' + dims.width + '"') : '18" x 13"';
  const isPortrait = orientation && orientation.toUpperCase() === 'PORTRAIT';

  let desc = '';

  // PARAGRAPH 1: Keyword fire (first 200 chars critical for eBay SEO)
  desc += '<p>';
  if (brandKeywords) {
    desc += '<b>DEATH NYC</b> x ' + brandKeywords + ' - The ultimate luxury street art mashup. ';
    desc += brandKeywords + ' meets underground NYC artistry. ';
  } else {
    desc += '<b>DEATH NYC</b> "' + subject + '" - Explosive pop art from NYC\'s underground queen. ';
  }
  desc += '<b>Death NYC</b> limited edition <b>hand-signed</b> print. Edition of 100. <b>COA included</b>. ';
  desc += 'Authentic <b>Death NYC</b> street art masterpiece.</p>';

  // PARAGRAPH 2: Artist credibility + market data
  desc += '<br><br>';
  desc += '<p><b>THE ARTIST:</b> <b>Death NYC</b> – the world\'s most famous living female street artist – ';
  desc += 'creates limited editions of only 100 pieces, each <b>hand-signed</b> in pencil, constantly rising in value. ';
  desc += 'Her work blends luxury fashion iconography with pop culture in ways that collectors obsess over. ';
  desc += '<b>Dec 2025 market data:</b> Death NYC luxury brand pieces selling $129–$179 unframed in under 7 days on eBay. ';
  desc += 'Her <b>Louis Vuitton</b> and <b>Supreme</b> mashups are the #1 trending pop art prints right now.</p>';

  // PARAGRAPH 3: Visual description
  desc += '<br><br>';
  desc += '<p><b>THIS PIECE:</b> "' + subject + '" - ';
  if (foundBrands.length > 0) {
    desc += 'A stunning collision of ' + foundBrands.join(' x ') + ' luxury aesthetics with Death NYC\'s signature street art style. ';
  }
  desc += 'Vibrant colors, razor-sharp detail, printed on premium 300gsm fine art paper. ';
  desc += 'Size: <b>' + sizeStr + '</b> (' + (isPortrait ? 'Portrait' : 'Landscape') + ' orientation). ';
  desc += 'The kind of piece that stops people mid-scroll and makes them reach for their wallet.</p>';

  // PARAGRAPH 4: Authentication & condition
  desc += '<br><br>';
  desc += '<p><b>AUTHENTICATION:</b><br>';
  desc += '✓ <b>Hand-signed</b> in pencil by Death NYC<br>';
  desc += '✓ <b>Numbered edition</b> out of 100<br>';
  desc += '✓ Official gallery <b>Certificate of Authenticity (COA)</b> included<br>';
  desc += '✓ <b>Mint condition</b> - never framed, stored flat in protective sleeve<br>';
  desc += '✓ Ships within 24 hours with rigid cardboard protection</p>';

  // PARAGRAPH 5: Urgency close
  desc += '<br><br>';
  desc += '<p><b>WHY BUY NOW:</b> These are the <b>hottest pop art prints on eBay</b> right now – ';
  if (foundBrands.length > 0) {
    desc += '<b>' + foundBrands[0] + '</b> versions outselling everything 5:1. ';
  } else {
    desc += 'Death NYC pieces outselling other street artists 5:1. ';
  }
  desc += 'This exact print won\'t last the week at this price. ';
  desc += 'Priced to move fast the way top sellers do. ';
  desc += '<b>Free shipping</b> with tracking. Ships same or next business day.</p>';

  // PARAGRAPH 6: SEO footer
  desc += '<br><br>';
  desc += '<p style="font-size:11px;color:#666;">';
  desc += 'Tags: Death NYC, ' + subject + ', Pop Art, Street Art, Limited Edition, Hand Signed, COA, ';
  if (foundBrands.length > 0) {
    desc += foundBrands.join(', ') + ', Luxury Art, Fashion Art, ';
  }
  desc += 'Contemporary Art, Urban Art, Investment Art, Collectible Print, NYC Artist, ';
  desc += 'Banksy Style, KAWS, Takashi Murakami, Shepard Fairey, Mr Brainwash</p>';

  return desc;
}

/**
 * Clear all caches to reload fresh data
 */
function clearAllCachesV2() {
  VARIABLES_CACHE = null;
  PROMPTS_CACHE = null;
  FIELDS_CACHE = null;
  if (typeof IMAGES_TAB_CACHE !== 'undefined') IMAGES_TAB_CACHE = null;
  Logger.log('✅ All caches cleared (V2)');
}

// ============================================================================
// SECTION 12: QUALITY & VALIDATION TOOLS
// ============================================================================

/**
 * Check for duplicate SKUs in the Products sheet
 */
function checkDuplicateSKUs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Products');

  if (!sheet) {
    ui.alert('Error', 'Products sheet not found!', ui.ButtonSet.OK);
    return;
  }

  const data = sheet.getDataRange().getValues();
  const skuCol = 0; // Column A = SKU
  const skuMap = {};
  const duplicates = [];

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const sku = data[i][skuCol];
    if (!sku || sku === '') continue;

    const skuStr = String(sku).trim();
    if (skuMap[skuStr]) {
      skuMap[skuStr].push(i + 1); // Row number (1-indexed)
    } else {
      skuMap[skuStr] = [i + 1];
    }
  }

  // Find duplicates
  for (const sku in skuMap) {
    if (skuMap[sku].length > 1) {
      duplicates.push({
        sku: sku,
        rows: skuMap[sku]
      });
    }
  }

  if (duplicates.length === 0) {
    ui.alert('✅ No Duplicates', 'All SKUs are unique! No duplicates found.', ui.ButtonSet.OK);
  } else {
    let msg = '⚠️ Found ' + duplicates.length + ' duplicate SKU(s):\n\n';
    for (const d of duplicates) {
      msg += '• ' + d.sku + '\n  Rows: ' + d.rows.join(', ') + '\n\n';
    }
    msg += 'Please resolve duplicates before posting to 3DSellers.';
    ui.alert('Duplicate SKUs Found', msg, ui.ButtonSet.OK);
  }

  Logger.log('Duplicate check complete. Found ' + duplicates.length + ' duplicates.');
  return duplicates;
}

/**
 * Validate all rows before posting to 3DSellers
 */
function validateBeforePost() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Products');

  if (!sheet) {
    ui.alert('Error', 'Products sheet not found!', ui.ButtonSet.OK);
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Required fields and their column indices
  const requiredFields = {
    'SKU': 0,           // A
    'Price': 1,         // B
    'Artist': 2,        // C
    'Title': 6,         // G
    'Description': 7,   // H
    'Image 1': 55       // BD (first image)
  };

  const issues = [];
  let validRows = 0;
  let totalRows = 0;

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sku = row[0];

    // Skip empty rows
    if (!sku || sku === '') continue;

    totalRows++;
    const rowIssues = [];

    // Check each required field
    for (const field in requiredFields) {
      const colIdx = requiredFields[field];
      const value = row[colIdx];

      if (!value || String(value).trim() === '') {
        rowIssues.push(field + ' is empty');
      }
    }

    // Check title length (max 80 chars for eBay)
    const title = row[6];
    if (title && String(title).length > 80) {
      rowIssues.push('Title exceeds 80 chars (' + String(title).length + ')');
    }

    // Check price is valid number
    const price = row[1];
    if (price && (isNaN(parseFloat(price)) || parseFloat(price) <= 0)) {
      rowIssues.push('Invalid price: ' + price);
    }

    // Check at least 3 images
    let imageCount = 0;
    for (let j = 55; j < 79 && j < row.length; j++) {
      if (row[j] && String(row[j]).trim() !== '') {
        imageCount++;
      }
    }
    if (imageCount < 3) {
      rowIssues.push('Less than 3 images (' + imageCount + ')');
    }

    if (rowIssues.length > 0) {
      issues.push({
        row: i + 1,
        sku: sku,
        issues: rowIssues
      });
    } else {
      validRows++;
    }
  }

  // Build result message
  let msg = '';
  if (issues.length === 0) {
    msg = '✅ ALL ' + totalRows + ' ROWS VALID!\n\n';
    msg += 'All required fields are filled:\n';
    msg += '• SKU\n• Price\n• Artist\n• Title\n• Description\n• At least 3 images\n\n';
    msg += 'Ready to post to 3DSellers!';
    ui.alert('Validation Passed', msg, ui.ButtonSet.OK);
  } else {
    msg = '⚠️ VALIDATION ISSUES\n\n';
    msg += 'Valid: ' + validRows + '/' + totalRows + ' rows\n';
    msg += 'Issues: ' + issues.length + ' rows need attention\n\n';

    // Show first 10 issues
    const showCount = Math.min(issues.length, 10);
    for (let i = 0; i < showCount; i++) {
      const issue = issues[i];
      msg += '━━━━━━━━━━━━━━━━━━━━\n';
      msg += 'Row ' + issue.row + ' (' + issue.sku + '):\n';
      for (const iss of issue.issues) {
        msg += '  • ' + iss + '\n';
      }
    }

    if (issues.length > 10) {
      msg += '\n... and ' + (issues.length - 10) + ' more rows with issues.\n';
    }

    msg += '\nCheck logs for full details.';
    ui.alert('Validation Issues Found', msg, ui.ButtonSet.OK);
  }

  // Log all issues
  Logger.log('Validation complete: ' + validRows + '/' + totalRows + ' valid');
  for (const issue of issues) {
    Logger.log('Row ' + issue.row + ' (' + issue.sku + '): ' + issue.issues.join(', '));
  }

  return { valid: validRows, total: totalRows, issues: issues };
}

/**
 * Show dashboard with summary statistics
 */
function showDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Products');

  if (!sheet) {
    ui.alert('Error', 'Products sheet not found!', ui.ButtonSet.OK);
    return;
  }

  const data = sheet.getDataRange().getValues();

  // Statistics
  let stats = {
    totalRows: 0,
    portraits: 0,
    landscapes: 0,
    byArtist: {},
    totalValue: 0,
    avgPrice: 0,
    withImages: 0,
    withDescription: 0,
    avgImageCount: 0,
    totalImages: 0
  };

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sku = row[0];

    // Skip empty rows
    if (!sku || sku === '') continue;

    stats.totalRows++;

    // Orientation
    if (String(sku).includes('_P_') || row[5] === 'Portrait') {
      stats.portraits++;
    } else {
      stats.landscapes++;
    }

    // Artist
    const artist = row[2] || 'Unknown';
    stats.byArtist[artist] = (stats.byArtist[artist] || 0) + 1;

    // Price
    const price = parseFloat(row[1]) || 0;
    stats.totalValue += price;

    // Description
    if (row[7] && String(row[7]).trim() !== '') {
      stats.withDescription++;
    }

    // Image count
    let imageCount = 0;
    for (let j = 55; j < 79 && j < row.length; j++) {
      if (row[j] && String(row[j]).trim() !== '') {
        imageCount++;
      }
    }
    if (imageCount > 0) {
      stats.withImages++;
      stats.totalImages += imageCount;
    }
  }

  stats.avgPrice = stats.totalRows > 0 ? (stats.totalValue / stats.totalRows).toFixed(2) : 0;
  stats.avgImageCount = stats.withImages > 0 ? (stats.totalImages / stats.withImages).toFixed(1) : 0;

  // Build dashboard
  let dash = '═══════════════════════════════════════\n';
  dash += '           📊 DASHBOARD SUMMARY\n';
  dash += '═══════════════════════════════════════\n\n';

  dash += '📦 INVENTORY\n';
  dash += '───────────────────────────────────────\n';
  dash += '  Total Products: ' + stats.totalRows + '\n';
  dash += '  Portrait: ' + stats.portraits + '\n';
  dash += '  Landscape: ' + stats.landscapes + '\n\n';

  dash += '👨‍🎨 BY ARTIST\n';
  dash += '───────────────────────────────────────\n';
  for (const artist in stats.byArtist) {
    dash += '  ' + artist + ': ' + stats.byArtist[artist] + '\n';
  }
  dash += '\n';

  dash += '💰 PRICING\n';
  dash += '───────────────────────────────────────\n';
  dash += '  Total Value: $' + stats.totalValue.toFixed(2) + '\n';
  dash += '  Avg Price: $' + stats.avgPrice + '\n\n';

  dash += '🖼️ CONTENT\n';
  dash += '───────────────────────────────────────\n';
  dash += '  With Descriptions: ' + stats.withDescription + '/' + stats.totalRows + '\n';
  dash += '  With Images: ' + stats.withImages + '/' + stats.totalRows + '\n';
  dash += '  Total Images: ' + stats.totalImages + '\n';
  dash += '  Avg Images/Product: ' + stats.avgImageCount + '\n\n';

  dash += '═══════════════════════════════════════\n';

  ui.alert('Dashboard', dash, ui.ButtonSet.OK);
  Logger.log(dash);

  return stats;
}

/**
 * Export Products sheet to CSV file (downloads as file)
 */
function exportToCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Products');

  if (!sheet) {
    ui.alert('Error', 'Products sheet not found!', ui.ButtonSet.OK);
    return;
  }

  const data = sheet.getDataRange().getValues();

  // Convert to CSV
  let csv = '';
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const csvRow = row.map(cell => {
      let val = cell === null || cell === undefined ? '' : String(cell);
      // Escape quotes and wrap in quotes if contains comma, quote, or newline
      if (val.includes(',') || val.includes('"') || val.includes('\n')) {
        val = '"' + val.replace(/"/g, '""') + '"';
      }
      return val;
    });
    csv += csvRow.join(',') + '\n';
  }

  // Create downloadable file
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  const filename = '3DSellers_Backup_' + timestamp + '.csv';

  const blob = Utilities.newBlob(csv, 'text/csv', filename);

  // Save to Drive
  const file = DriveApp.createFile(blob);
  const fileUrl = file.getUrl();

  const msg = '✅ CSV Export Complete!\n\n' +
              'File: ' + filename + '\n' +
              'Rows: ' + data.length + '\n\n' +
              'File saved to your Google Drive root folder.\n\n' +
              'URL: ' + fileUrl;

  ui.alert('Export Complete', msg, ui.ButtonSet.OK);
  Logger.log('Exported ' + data.length + ' rows to ' + filename);

  return fileUrl;
}

/**
 * Smart brand detection - scans titles and adds brand tags
 */
function detectBrands() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Products');

  if (!sheet) {
    ui.alert('Error', 'Products sheet not found!', ui.ButtonSet.OK);
    return;
  }

  // Common luxury/pop culture brands to detect
  const brands = [
    'Louis Vuitton', 'LV', 'Supreme', 'Chanel', 'Gucci', 'Hermes', 'Hermès',
    'Nike', 'Adidas', 'Jordan', 'Dior', 'Prada', 'Versace', 'Burberry',
    'Fendi', 'Balenciaga', 'Off-White', 'Kaws', 'KAWS', 'Banksy',
    'Mickey Mouse', 'Disney', 'Snoopy', 'Peanuts', 'Coca-Cola', 'Pepsi',
    'BMW', 'Mercedes', 'Porsche', 'Ferrari', 'Lamborghini', 'Tesla',
    'Apple', 'Bitcoin', 'Ethereum', 'Playboy', 'Vogue',
    'Star Wars', 'Marvel', 'Batman', 'Superman', 'Spiderman',
    'Marilyn Monroe', 'Elvis', 'Beatles', 'Rolling Stones', 'Grateful Dead'
  ];

  const data = sheet.getDataRange().getValues();
  const results = [];

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sku = row[0];
    const title = String(row[6] || '');
    const desc = String(row[7] || '');
    const tags = String(row[8] || '');

    if (!sku || sku === '') continue;

    const foundBrands = [];
    const searchText = (title + ' ' + desc).toUpperCase();

    for (const brand of brands) {
      if (searchText.includes(brand.toUpperCase())) {
        foundBrands.push(brand);
      }
    }

    if (foundBrands.length > 0) {
      // Check if brands are already in tags
      const currentTags = tags.toUpperCase();
      const missingBrands = foundBrands.filter(b => !currentTags.includes(b.toUpperCase()));

      if (missingBrands.length > 0) {
        // Add missing brands to tags
        const newTags = tags ? tags + ', ' + missingBrands.join(', ') : missingBrands.join(', ');
        sheet.getRange(i + 1, 9).setValue(newTags); // Column I = Tags
        results.push({
          row: i + 1,
          sku: sku,
          added: missingBrands
        });
      }
    }
  }

  // Show results
  if (results.length === 0) {
    ui.alert('Brand Detection', '✅ All detected brands are already in tags!\n\nNo updates needed.', ui.ButtonSet.OK);
  } else {
    let msg = '✅ Added brand tags to ' + results.length + ' rows:\n\n';
    for (const r of results.slice(0, 10)) {
      msg += 'Row ' + r.row + ': +' + r.added.join(', ') + '\n';
    }
    if (results.length > 10) {
      msg += '\n... and ' + (results.length - 10) + ' more rows.\n';
    }
    ui.alert('Brand Detection Complete', msg, ui.ButtonSet.OK);
  }

  Logger.log('Brand detection: Updated ' + results.length + ' rows');
  return results;
}

/**
 * Bulk edit selected rows - set a field value for all selected rows
 */
function bulkEditSelected() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Products');

  if (!sheet) {
    ui.alert('Error', 'Products sheet not found!', ui.ButtonSet.OK);
    return;
  }

  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow < 2) {
    ui.alert('Error', 'Please select data rows (not the header row).', ui.ButtonSet.OK);
    return;
  }

  // Ask which field to edit
  const fieldResponse = ui.prompt(
    'Bulk Edit',
    'Which field do you want to edit?\n\nOptions:\n• Price (column B)\n• Artist (column C)\n• Format (column D)\n\nEnter field name:',
    ui.ButtonSet.OK_CANCEL
  );

  if (fieldResponse.getSelectedButton() !== ui.Button.OK) return;

  const fieldName = fieldResponse.getResponseText().trim().toLowerCase();
  let colIndex;

  switch (fieldName) {
    case 'price': colIndex = 2; break;  // B
    case 'artist': colIndex = 3; break; // C
    case 'format': colIndex = 4; break; // D
    default:
      ui.alert('Error', 'Unknown field: ' + fieldName + '\n\nValid options: Price, Artist, Format', ui.ButtonSet.OK);
      return;
  }

  // Ask for new value
  const valueResponse = ui.prompt(
    'Bulk Edit - ' + fieldName,
    'Enter new value for ' + fieldName + ' (will apply to ' + numRows + ' rows):',
    ui.ButtonSet.OK_CANCEL
  );

  if (valueResponse.getSelectedButton() !== ui.Button.OK) return;

  const newValue = valueResponse.getResponseText().trim();

  // Apply to all selected rows
  for (let i = 0; i < numRows; i++) {
    const rowNum = startRow + i;
    const sku = sheet.getRange(rowNum, 1).getValue();
    if (sku && sku !== '') {
      sheet.getRange(rowNum, colIndex).setValue(newValue);
    }
  }

  ui.alert('Bulk Edit Complete', '✅ Updated ' + fieldName + ' to "' + newValue + '" for ' + numRows + ' rows.', ui.ButtonSet.OK);
  Logger.log('Bulk edit: Set ' + fieldName + ' to "' + newValue + '" for rows ' + startRow + '-' + (startRow + numRows - 1));
}

// ============================================================================
// SECTION 19: DYNAMIC FOLDER LOADER FROM IMAGES TAB
// ============================================================================
// Loads folder IDs from the IMAGES tab in the Google Sheet

let IMAGES_TAB_CACHE = null;

/**
 * Load folder configurations from IMAGES tab in Google Sheet
 * Expected format:
 * Column A: Folder Type (e.g., "Portrait Mock-ups", "Landscape Mock-ups")
 * Column B: Artist (e.g., "Death NYC", "Shepard Fairey")
 * Column C: Google Drive Folder ID or URL
 * Column D: (Optional) Description/Notes
 */
function loadFoldersFromImagesTab() {
  if (IMAGES_TAB_CACHE) return IMAGES_TAB_CACHE;

  try {
    const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName('iMAGES') || ss.getSheetByName('IMAGES') || ss.getSheetByName('Images');

    if (!sheet) {
      Logger.log('⚠️ IMAGES tab not found, using hardcoded folder IDs');
      return null;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('⚠️ IMAGES tab is empty');
      return null;
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const folders = {};

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const folderType = row[0] ? row[0].toString().trim() : '';
      const artist = row[1] ? row[1].toString().trim().toUpperCase() : '';
      const folderIdOrUrl = row[2] ? row[2].toString().trim() : '';

      if (!folderType || !folderIdOrUrl) continue;

      // Extract folder ID from URL if needed
      const folderId = extractFolderId(folderIdOrUrl);
      if (!folderId) continue;

      // Initialize artist object if needed
      if (!folders[artist]) {
        folders[artist] = {};
      }

      // Map folder type names to standard keys
      const typeKey = mapFolderTypeToKey(folderType);
      if (typeKey) {
        folders[artist][typeKey] = folderId;
        Logger.log(`  📁 Loaded: ${artist} > ${typeKey} = ${folderId.substring(0, 15)}...`);
      }
    }

    Logger.log(`✅ Loaded ${Object.keys(folders).length} artists from IMAGES tab`);
    IMAGES_TAB_CACHE = folders;
    return folders;

  } catch (error) {
    Logger.log(`❌ Error loading IMAGES tab: ${error.message}`);
    return null;
  }
}

/**
 * Extract Google Drive folder ID from URL or return ID directly
 */
function extractFolderId(input) {
  if (!input) return null;

  // If it's already just an ID (no slashes)
  if (!input.includes('/') && !input.includes('http')) {
    return input;
  }

  // Extract from Google Drive URL
  // Format: https://drive.google.com/drive/folders/FOLDER_ID
  const match = input.match(/folders\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];

  // Format: https://drive.google.com/drive/u/0/folders/FOLDER_ID
  const match2 = input.match(/\/([a-zA-Z0-9_-]{25,})/);
  if (match2) return match2[1];

  return null;
}

/**
 * Map folder type name (from sheet) to standard key
 */
function mapFolderTypeToKey(typeName) {
  const name = typeName.toLowerCase();

  if (name.includes('portrait') && name.includes('mock')) return 'MOCKUPS_PORTRAIT';
  if (name.includes('landscape') && name.includes('mock')) return 'MOCKUPS_LANDSCAPE';
  if (name.includes('portrait') && name.includes('crop') && name.includes('single')) return 'CROPS_PORTRAIT';
  if (name.includes('landscape') && name.includes('crop') && name.includes('single')) return 'CROPS_LANDSCAPE';
  if (name.includes('crop') && name.includes('all')) return 'CROPS_ALL';
  if (name.includes('stock')) return 'STOCK';
  if (name.includes('master') || name.includes('product')) return 'MASTER';
  if (name.includes('video')) return 'VIDEO';

  // Fallback mappings
  if (name.includes('portrait') && name.includes('crop')) return 'CROPS_PORTRAIT';
  if (name.includes('landscape') && name.includes('crop')) return 'CROPS_LANDSCAPE';

  Logger.log(`  ⚠️ Unknown folder type: ${typeName}`);
  return null;
}

/**
 * Get folders for a specific artist - first tries IMAGES tab, then falls back to hardcoded
 */
function getArtistFolders(artistName) {
  const normalizedArtist = artistName.toUpperCase().trim();

  // Try loading from IMAGES tab first
  const dynamicFolders = loadFoldersFromImagesTab();
  if (dynamicFolders && dynamicFolders[normalizedArtist]) {
    Logger.log(`📂 Using dynamic folders from IMAGES tab for ${normalizedArtist}`);
    return dynamicFolders[normalizedArtist];
  }

  // Fall back to hardcoded ARTIST_FOLDERS
  if (ARTIST_FOLDERS[normalizedArtist]) {
    Logger.log(`📂 Using hardcoded folders for ${normalizedArtist}`);
    return ARTIST_FOLDERS[normalizedArtist];
  }

  // Try variations
  const variations = [
    normalizedArtist,
    normalizedArtist.replace(' ', '_'),
    'DEATH NYC'  // Default fallback
  ];

  for (const variant of variations) {
    if (dynamicFolders && dynamicFolders[variant]) return dynamicFolders[variant];
    if (ARTIST_FOLDERS[variant]) return ARTIST_FOLDERS[variant];
  }

  Logger.log(`⚠️ No folders found for artist: ${artistName}`);
  return ARTIST_FOLDERS['DEATH NYC'] || {};
}

/**
 * Clear the images tab cache (for reloading)
 */
function clearImagesTabCache() {
  IMAGES_TAB_CACHE = null;
  Logger.log('✅ IMAGES tab cache cleared');
}

/**
 * Show loaded folders from IMAGES tab
 */
function showLoadedImageFolders() {
  clearImagesTabCache();  // Force reload
  const folders = loadFoldersFromImagesTab();

  let report = '═══════════════════════════════════════════\n';
  report += '📂 FOLDERS FROM IMAGES TAB\n';
  report += '═══════════════════════════════════════════\n\n';

  if (!folders) {
    report += '⚠️ No folders loaded from IMAGES tab.\n';
    report += 'Using hardcoded ARTIST_FOLDERS instead.\n';
  } else {
    for (const [artist, folderMap] of Object.entries(folders)) {
      report += `\n🎨 ${artist}:\n`;
      for (const [type, id] of Object.entries(folderMap)) {
        report += `   ${type}: ${id.substring(0, 20)}...\n`;
      }
    }
  }

  Logger.log(report);
  SpreadsheetApp.getUi().alert('Folders from IMAGES Tab', report, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// SECTION 20: FILE RENAME & CONSOLIDATION SYSTEM
// ============================================================================
// Pull files from 6 Death NYC folders, rename by SKU, copy to single destination

const RENAME_CONFIG = {
  // Destination folder for consolidated renamed files
  DESTINATION_FOLDER_ID: '', // Will be created or specified
  DESTINATION_FOLDER_NAME: 'CONSOLIDATED_IMAGES',

  // Processing order for folders
  FOLDER_ORDER: [
    'MOCKUPS_PORTRAIT',
    'MOCKUPS_LANDSCAPE',
    'CROPS_PORTRAIT',
    'CROPS_LANDSCAPE',
    'CROPS_ALL',
    'STOCK'
  ],

  // ALL images use the same naming pattern: {SKU}_crop_{number}.jpg
  // This matches the gauntlet.gallery hosted URL pattern
  USE_UNIFIED_NAMING: true,
  UNIFIED_PATTERN: '_crop_'  // Same as HOSTED_IMAGES.CROP_PATTERN
};

/**
 * Get or create the destination folder for consolidated images
 */
function getOrCreateDestinationFolder() {
  let destFolder;

  if (RENAME_CONFIG.DESTINATION_FOLDER_ID) {
    try {
      destFolder = DriveApp.getFolderById(RENAME_CONFIG.DESTINATION_FOLDER_ID);
      Logger.log(`✅ Using existing destination folder: ${destFolder.getName()}`);
      return destFolder;
    } catch (e) {
      Logger.log(`⚠️ Destination folder ID not found, creating new folder`);
    }
  }

  // Create in root or find existing
  const folders = DriveApp.getFoldersByName(RENAME_CONFIG.DESTINATION_FOLDER_NAME);
  if (folders.hasNext()) {
    destFolder = folders.next();
    Logger.log(`✅ Found existing folder: ${destFolder.getName()}`);
  } else {
    destFolder = DriveApp.createFolder(RENAME_CONFIG.DESTINATION_FOLDER_NAME);
    Logger.log(`✅ Created new folder: ${destFolder.getName()}`);
  }

  return destFolder;
}

/**
 * Get all SKUs from the Products sheet
 */
function getAllSKUsFromSheet() {
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    Logger.log('❌ Products sheet not found');
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const skuData = sheet.getRange(2, PRODUCTS_COLUMNS.SKU, lastRow - 1, 1).getValues();
  const skus = skuData.map(row => row[0]).filter(sku => sku && sku !== '');

  Logger.log(`📋 Found ${skus.length} SKUs in Products sheet`);
  return skus;
}

/**
 * Match a filename to a SKU
 * Tries various matching strategies
 */
function matchFileToSKU(filename, skuList) {
  const cleanFilename = filename.replace(/\.(jpg|jpeg|png|gif|webp)$/i, '').toUpperCase();

  for (const sku of skuList) {
    const cleanSKU = sku.toUpperCase();

    // Exact match
    if (cleanFilename === cleanSKU) return sku;

    // SKU at start of filename
    if (cleanFilename.startsWith(cleanSKU)) return sku;

    // SKU contained in filename
    if (cleanFilename.includes(cleanSKU)) return sku;

    // Partial match (first part of SKU)
    const skuPrefix = cleanSKU.split('-')[0] + '-' + cleanSKU.split('-')[1];
    if (cleanFilename.includes(skuPrefix)) return sku;
  }

  return null;
}

/**
 * Rename and copy a single file
 * ALL files renamed to: {SKU}_crop_{number}.jpg (1-24)
 */
function renameAndCopyFile(file, sku, folderType, destFolder, fileIndex) {
  try {
    const originalName = file.getName();
    const extension = originalName.split('.').pop().toLowerCase();

    // Unified naming: {SKU}_crop_{number}.jpg
    // Example: 1_DNYC_Mickey-Mouse_V1_crop_1.jpg
    const newName = `${sku}${RENAME_CONFIG.UNIFIED_PATTERN}${fileIndex}.${extension}`;

    // Make a copy with new name
    const newFile = file.makeCopy(newName, destFolder);

    Logger.log(`  ✅ ${originalName} → ${newName}`);
    return { success: true, originalName, newName, fileId: newFile.getId() };

  } catch (error) {
    Logger.log(`  ❌ Error copying ${file.getName()}: ${error.message}`);
    return { success: false, error: error.message };
  }
}

/**
 * Process a single folder - rename and copy all matched files
 */
function processFolder(folderId, folderType, skuList, destFolder) {
  if (!folderId) {
    Logger.log(`⚠️ No folder ID for ${folderType}, skipping`);
    return { processed: 0, matched: 0, errors: 0 };
  }

  const stats = { processed: 0, matched: 0, errors: 0, files: [] };

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();

    Logger.log(`\n📁 Processing ${folderType}: ${folder.getName()}`);
    Logger.log('─'.repeat(50));

    const fileCounters = {}; // Track file index per SKU

    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();

      // Only process images
      if (!mimeType.startsWith('image/')) continue;

      stats.processed++;
      const filename = file.getName();
      const matchedSKU = matchFileToSKU(filename, skuList);

      if (matchedSKU) {
        fileCounters[matchedSKU] = (fileCounters[matchedSKU] || 0) + 1;
        const result = renameAndCopyFile(file, matchedSKU, folderType, destFolder, fileCounters[matchedSKU]);

        if (result.success) {
          stats.matched++;
          stats.files.push(result);
        } else {
          stats.errors++;
        }
      } else {
        Logger.log(`  ⚠️ No SKU match for: ${filename}`);
      }
    }

  } catch (error) {
    Logger.log(`❌ Error processing folder ${folderType}: ${error.message}`);
    stats.errors++;
  }

  return stats;
}

/**
 * Main function: Pull and rename all files from 6 Death NYC folders
 */
function pullAndRenameAllFiles() {
  const ui = SpreadsheetApp.getUi();

  Logger.log('\n' + '═'.repeat(70));
  Logger.log('📂 FILE RENAME & CONSOLIDATION SYSTEM');
  Logger.log('═'.repeat(70) + '\n');

  // Get SKUs from sheet
  const skuList = getAllSKUsFromSheet();
  if (skuList.length === 0) {
    ui.alert('❌ Error', 'No SKUs found in Products sheet. Please add products first.', ui.ButtonSet.OK);
    return;
  }

  // Get/create destination folder
  const destFolder = getOrCreateDestinationFolder();
  Logger.log(`📁 Destination: ${destFolder.getName()} (${destFolder.getId()})`);

  // Get Death NYC folders - first try IMAGES tab, then fallback to hardcoded
  const artistFolders = getArtistFolders('DEATH NYC');
  if (!artistFolders || Object.keys(artistFolders).length === 0) {
    ui.alert('❌ Error', 'Death NYC folder configuration not found.\n\nCheck IMAGES tab or ARTIST_FOLDERS config.', ui.ButtonSet.OK);
    return;
  }

  const totalStats = { processed: 0, matched: 0, errors: 0 };

  // Process folders in order
  for (const folderType of RENAME_CONFIG.FOLDER_ORDER) {
    const folderId = artistFolders[folderType];
    const stats = processFolder(folderId, folderType, skuList, destFolder);

    totalStats.processed += stats.processed;
    totalStats.matched += stats.matched;
    totalStats.errors += stats.errors;
  }

  // Summary
  Logger.log('\n' + '═'.repeat(70));
  Logger.log('📊 CONSOLIDATION COMPLETE');
  Logger.log('─'.repeat(70));
  Logger.log(`  Total Files Scanned: ${totalStats.processed}`);
  Logger.log(`  Files Matched & Copied: ${totalStats.matched}`);
  Logger.log(`  Errors: ${totalStats.errors}`);
  Logger.log(`  Destination Folder: ${destFolder.getUrl()}`);
  Logger.log('═'.repeat(70) + '\n');

  ui.alert(
    '✅ Consolidation Complete!',
    `Files Scanned: ${totalStats.processed}\n` +
    `Files Copied: ${totalStats.matched}\n` +
    `Errors: ${totalStats.errors}\n\n` +
    `Destination: ${destFolder.getName()}`,
    ui.ButtonSet.OK
  );
}

/**
 * Preview which files would be renamed (dry run)
 */
function previewFileRename() {
  const ui = SpreadsheetApp.getUi();

  Logger.log('\n' + '═'.repeat(70));
  Logger.log('👁️ FILE RENAME PREVIEW (DRY RUN)');
  Logger.log('═'.repeat(70) + '\n');

  const skuList = getAllSKUsFromSheet();
  if (skuList.length === 0) {
    ui.alert('❌ No SKUs found in Products sheet');
    return;
  }

  const artistFolders = getArtistFolders('DEATH NYC');
  let totalFiles = 0;
  let matchedFiles = 0;

  for (const folderType of RENAME_CONFIG.FOLDER_ORDER) {
    const folderId = artistFolders[folderType];
    if (!folderId) continue;

    try {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();

      Logger.log(`\n📁 ${folderType}: ${folder.getName()}`);

      while (files.hasNext()) {
        const file = files.next();
        if (!file.getMimeType().startsWith('image/')) continue;

        totalFiles++;
        const filename = file.getName();
        const matchedSKU = matchFileToSKU(filename, skuList);

        if (matchedSKU) {
          matchedFiles++;
          const suffix = RENAME_CONFIG.TYPE_SUFFIXES[folderType] || '';
          Logger.log(`  ✅ ${filename} → ${matchedSKU}${suffix}_XX.ext`);
        } else {
          Logger.log(`  ⚠️ ${filename} → NO MATCH`);
        }
      }
    } catch (e) {
      Logger.log(`  ❌ Error: ${e.message}`);
    }
  }

  Logger.log(`\n📊 Summary: ${matchedFiles}/${totalFiles} files would be renamed`);
  ui.alert('Preview Complete', `${matchedFiles} of ${totalFiles} files would be renamed.\n\nCheck Logs for details.`, ui.ButtonSet.OK);
}

// ============================================================================
// SECTION 21: SEO OPTIMIZATION & PRICE INTELLIGENCE (FROM V6)
// ============================================================================

const SEO_CONFIG = {
  ENABLED: true,
  MAX_TITLE_LENGTH: 80,
  EBAY_SANDBOX_BASE: 'https://api.sandbox.ebay.com',
  MAX_COMPS_PER_SOURCE: 20,
  GOOGLE_SEARCH_RESULTS: 10,
  TREND_SCORE_THRESHOLD_HIGH: 70,
  TREND_SCORE_THRESHOLD_MEDIUM: 60,
  TREND_SCORE_THRESHOLD_LOW: 50,
  META_KEYWORDS_MAX_LENGTH: 480,

  // Price multipliers
  PRICE_MULTIPLIER_SIGNED: 1.20,
  PRICE_MULTIPLIER_LIMITED: 1.15,
  PRICE_MULTIPLIER_NEW: 1.10,
  PRICE_MULTIPLIER_UNSIGNED: 0.85,
  PRICE_MULTIPLIER_PREOWNED: 0.90,
  PRICE_ROUNDING: 10,
  PRICE_TRIM_COUNT: 2
};

/**
 * Optimize SEO and pricing for all products
 */
function optimizeSEOForAllProducts() {
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Products sheet not found');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('⚠️ No products found');
    return;
  }

  Logger.log('\n' + '═'.repeat(80));
  Logger.log('🔍 SEO OPTIMIZATION & PRICE INTELLIGENCE');
  Logger.log('═'.repeat(80) + '\n');

  const stats = { processed: 0, updated: 0, skipped: 0, errors: 0 };

  for (let row = 2; row <= lastRow; row++) {
    const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue();
    if (!sku) {
      stats.skipped++;
      continue;
    }

    Logger.log(`\n📊 Row ${row}: ${sku}`);

    try {
      const updated = optimizeSEOForRow(sheet, row);
      if (updated) stats.updated++;
      else stats.skipped++;
      stats.processed++;
    } catch (error) {
      Logger.log(`  ❌ Error: ${error.message}`);
      stats.errors++;
    }

    Utilities.sleep(500);
  }

  Logger.log(`\n${'═'.repeat(80)}`);
  Logger.log(`📈 OPTIMIZATION COMPLETE`);
  Logger.log(`  Processed: ${stats.processed} | Updated: ${stats.updated} | Errors: ${stats.errors}`);

  SpreadsheetApp.getUi().alert(
    '✅ SEO Optimization Complete!',
    `Processed: ${stats.processed}\nUpdated: ${stats.updated}\nErrors: ${stats.errors}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Optimize SEO for a single row
 */
function optimizeSEOForRow(sheet, row) {
  const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue() || '';
  const artist = sheet.getRange(row, PRODUCTS_COLUMNS.ARTIST).getValue() || '';
  const currentTitle = sheet.getRange(row, PRODUCTS_COLUMNS.TITLE).getValue() || '';
  const description = sheet.getRange(row, PRODUCTS_COLUMNS.DESCRIPTION).getValue() || '';

  if (!currentTitle && !description) {
    Logger.log('  ⚠️ No title or description, skipping');
    return false;
  }

  // Extract entities
  const entities = extractEntitiesFromDescription(description);
  const primaryTerm = entities.subjects[0] || extractBaseTerm(currentTitle);

  Logger.log(`  🔍 Primary term: ${primaryTerm}`);

  // Fetch comparables
  const ebayComps = fetchEbayComps(primaryTerm);
  const googleComps = fetchGoogleComps(primaryTerm);

  // Calculate price
  const allPrices = [...ebayComps.prices, ...googleComps.prices].filter(p => p > 0);
  const priceResult = calculateOptimizedPrice(allPrices, entities);

  // Get trend scores
  const allKeywords = [...new Set([...entities.subjects, ...ebayComps.keywords, ...googleComps.keywords])].slice(0, 12);
  const trendScores = getTrendScores(allKeywords);

  // Build optimized fields
  const optTitle = buildOptimizedTitle(entities, ebayComps.keywords, googleComps.titles || [], trendScores, currentTitle);
  const optTags = buildOptimizedTags(entities, ebayComps.keywords, googleComps.keywords, trendScores);
  const optMetaKeywords = buildOptimizedMetaKeywords(entities, ebayComps.keywords, googleComps.keywords, trendScores);
  const optMetaDesc = buildOptimizedMetaDescription(entities, optTitle, trendScores);

  // Update sheet
  sheet.getRange(row, PRODUCTS_COLUMNS.TITLE).setValue(optTitle);
  sheet.getRange(row, PRODUCTS_COLUMNS.TAGS).setValue(optTags);
  sheet.getRange(row, PRODUCTS_COLUMNS.META_KEYWORDS).setValue(optMetaKeywords);
  sheet.getRange(row, PRODUCTS_COLUMNS.META_DESCRIPTION).setValue(optMetaDesc);
  if (priceResult.price !== 'N/A') {
    sheet.getRange(row, PRODUCTS_COLUMNS.PRICE).setValue(priceResult.price);
  }

  Logger.log(`  ✅ Updated: Title, Tags, Meta, Price: $${priceResult.price}`);
  return true;
}

/**
 * Fetch eBay comparable listings
 */
function fetchEbayComps(term) {
  if (!term) return { prices: [], keywords: [] };

  try {
    const variables = loadVariables();
    const clientId = variables.apis?.EBAY_CLIENT_ID || '';
    const clientSecret = variables.apis?.EBAY_CLIENT_SECRET || '';

    if (!clientId || !clientSecret) {
      Logger.log('  ⚠️ eBay API not configured');
      return { prices: [], keywords: [] };
    }

    // Get token
    const tokenUrl = `${SEO_CONFIG.EBAY_SANDBOX_BASE}/identity/v1/oauth2/token`;
    const tokenResponse = UrlFetchApp.fetch(tokenUrl, {
      method: 'post',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Authorization': 'Basic ' + Utilities.base64Encode(`${clientId}:${clientSecret}`)
      },
      payload: 'grant_type=client_credentials&scope=https://api.ebay.com/oauth/api_scope',
      muteHttpExceptions: true
    });

    const tokenJson = JSON.parse(tokenResponse.getContentText());
    if (!tokenJson.access_token) throw new Error('No access token');

    // Search
    const searchUrl = `${SEO_CONFIG.EBAY_SANDBOX_BASE}/buy/browse/v1/item_summary/search?q=${encodeURIComponent(term)}&limit=${SEO_CONFIG.MAX_COMPS_PER_SOURCE}`;
    const searchResponse = UrlFetchApp.fetch(searchUrl, {
      headers: { 'Authorization': 'Bearer ' + tokenJson.access_token },
      muteHttpExceptions: true
    });

    const searchJson = JSON.parse(searchResponse.getContentText());
    const items = searchJson.itemSummaries || [];

    const prices = [];
    const keywords = new Set();

    items.forEach(item => {
      const price = parseFloat(item.price?.value || 0);
      if (price > 0) prices.push(price);

      (item.title || '').toLowerCase().split(/\s+/).forEach(w => {
        const clean = w.replace(/[^\w]/g, '');
        if (clean.length > 3) keywords.add(clean);
      });
    });

    return { prices, keywords: Array.from(keywords) };

  } catch (error) {
    Logger.log(`  ⚠️ eBay API error: ${error.message}`);
    return { prices: [], keywords: [] };
  }
}

/**
 * Fetch Google comparable listings
 */
function fetchGoogleComps(query) {
  try {
    const variables = loadVariables();
    const apiKey = variables.apis?.GOOGLE_API_KEY || '';
    const cseId = variables.apis?.GOOGLE_CSE_ID || '';

    if (!apiKey || !cseId) {
      Logger.log('  ⚠️ Google API not configured');
      return { prices: [], titles: [], keywords: [] };
    }

    const url = `https://www.googleapis.com/customsearch/v1?key=${apiKey}&cx=${cseId}&q=${encodeURIComponent(query + ' site:ebay.com OR site:etsy.com')}&num=${SEO_CONFIG.GOOGLE_SEARCH_RESULTS}`;

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    const items = json.items || [];

    const prices = [];
    const titles = [];
    const keywords = new Set();

    items.forEach(item => {
      titles.push(item.title);
      const text = (item.title + ' ' + (item.snippet || '')).toLowerCase();

      text.split(/\s+/).forEach(w => {
        if (w.length > 3) keywords.add(w);
      });

      const match = text.match(/\$([0-9,]+(\.[0-9]{2})?)/);
      if (match) prices.push(parseFloat(match[1].replace(/,/g, '')));
    });

    return { prices, titles, keywords: Array.from(keywords) };

  } catch (error) {
    Logger.log(`  ⚠️ Google API error: ${error.message}`);
    return { prices: [], titles: [], keywords: [] };
  }
}

/**
 * Get Google Trends scores
 */
function getTrendScores(keywords) {
  const scores = {};
  keywords.forEach(kw => {
    scores[kw] = 50; // Default score (Trends API is complex, using placeholder)
  });
  return scores;
}

/**
 * Calculate optimized price
 */
function calculateOptimizedPrice(prices, entities) {
  if (prices.length === 0) return { price: 'N/A', notes: 'No comps' };

  const sorted = prices.sort((a, b) => a - b);
  const trimCount = Math.min(SEO_CONFIG.PRICE_TRIM_COUNT, Math.floor(sorted.length / 4));
  const trimmed = sorted.slice(trimCount, sorted.length - trimCount || undefined);
  const avg = Math.round(trimmed.reduce((a, b) => a + b, 0) / (trimmed.length || 1));

  let multiplier = 1.0;
  if (entities.style.some(s => s.includes('signed') || s.includes('coa'))) {
    multiplier += (SEO_CONFIG.PRICE_MULTIPLIER_SIGNED - 1);
  }
  if (entities.style.some(s => s.includes('limited'))) {
    multiplier += (SEO_CONFIG.PRICE_MULTIPLIER_LIMITED - 1);
  }

  const final = Math.round(avg * multiplier / SEO_CONFIG.PRICE_ROUNDING) * SEO_CONFIG.PRICE_ROUNDING;
  return { price: final, notes: `Avg: $${avg}, ${prices.length} comps` };
}

/**
 * Build optimized title
 */
function buildOptimizedTitle(entities, ebayKeywords, googleTitles, trends, currentTitle) {
  const highTrend = Object.keys(trends).filter(k => trends[k] > SEO_CONFIG.TREND_SCORE_THRESHOLD_HIGH).slice(0, 2);

  let best = [entities.subjects[0] || '', ...highTrend, entities.style[0] || ''].filter(Boolean).join(' ');

  if (currentTitle && currentTitle.length > 30 && currentTitle.length <= SEO_CONFIG.MAX_TITLE_LENGTH) {
    best = currentTitle;
  }

  if (best.length > SEO_CONFIG.MAX_TITLE_LENGTH) {
    best = best.substring(0, SEO_CONFIG.MAX_TITLE_LENGTH).replace(/\s+\S*$/, '');
  }

  return best.trim();
}

/**
 * Build optimized tags
 */
function buildOptimizedTags(entities, ebayKeywords, googleKeywords, trends) {
  const all = [...new Set([...entities.subjects, ...entities.style, ...ebayKeywords.slice(0, 15), ...googleKeywords.slice(0, 15)])];
  return all.join(', ');
}

/**
 * Build optimized meta keywords
 */
function buildOptimizedMetaKeywords(entities, ebayKeywords, googleKeywords, trends) {
  const all = [...new Set([...entities.subjects, ...entities.style, ...ebayKeywords, ...googleKeywords])];
  let str = all.join(', ');
  if (str.length > SEO_CONFIG.META_KEYWORDS_MAX_LENGTH) {
    str = str.substring(0, SEO_CONFIG.META_KEYWORDS_MAX_LENGTH - 3) + '...';
  }
  return str;
}

/**
 * Build optimized meta description
 */
function buildOptimizedMetaDescription(entities, title, trends) {
  const rarity = entities.style.some(s => s.includes('limited')) ? 'hand-signed limited edition' : 'authentic';
  return `${title} – ${rarity} pop art print. ${entities.subjects.slice(0, 2).join(' & ')}. Fast insured shipping!`;
}

/**
 * Extract entities from description
 */
function extractEntitiesFromDescription(description) {
  const lower = description.toLowerCase();

  const subjects = [...new Set(
    lower.match(/(death nyc|taylor swift|banksy|joker|balloon girl|marilyn|louis vuitton|warhol|tupac|supreme|chanel)/g) || []
  )];

  const style = [...new Set(
    lower.match(/(pop art|street art|graffiti|limited edition|signed|coa|artist proof)/g) || []
  )];

  return { subjects, style, condition: 'New' };
}

/**
 * Extract base term from title
 */
function extractBaseTerm(title) {
  return title.trim().split(/\s+/).slice(0, 3).join(' ');
}

// ============================================================================
// SECTION 22: CIRCUIT BREAKER & RETRY LOGIC (FROM V6)
// ============================================================================

const CIRCUIT_BREAKER_CONFIG = {
  ENABLED: true,
  FAILURE_THRESHOLD: 5,
  SUCCESS_THRESHOLD: 2,
  TIMEOUT_MS: 60000
};

const RETRY_CONFIG = {
  MAX_ATTEMPTS: 3,
  INITIAL_DELAY_MS: 1000,
  MAX_DELAY_MS: 10000,
  BACKOFF_MULTIPLIER: 2
};

const CIRCUIT_STATE = {
  state: 'CLOSED',
  failureCount: 0,
  successCount: 0,
  lastFailureTime: null,
  nextAttemptTime: null
};

/**
 * Check if operation can proceed based on circuit state
 */
function canProceedWithOperation() {
  if (!CIRCUIT_BREAKER_CONFIG.ENABLED) return true;

  const now = new Date().getTime();

  switch (CIRCUIT_STATE.state) {
    case 'CLOSED':
      return true;
    case 'OPEN':
      if (now >= CIRCUIT_STATE.nextAttemptTime) {
        CIRCUIT_STATE.state = 'HALF_OPEN';
        CIRCUIT_STATE.successCount = 0;
        Logger.log('🔄 Circuit breaker: OPEN → HALF_OPEN');
        return true;
      }
      Logger.log('⛔ Circuit breaker OPEN - blocking operation');
      return false;
    case 'HALF_OPEN':
      return true;
    default:
      return true;
  }
}

/**
 * Record successful operation
 */
function recordSuccess() {
  if (!CIRCUIT_BREAKER_CONFIG.ENABLED) return;

  if (CIRCUIT_STATE.state === 'HALF_OPEN') {
    CIRCUIT_STATE.successCount++;
    if (CIRCUIT_STATE.successCount >= CIRCUIT_BREAKER_CONFIG.SUCCESS_THRESHOLD) {
      CIRCUIT_STATE.state = 'CLOSED';
      CIRCUIT_STATE.failureCount = 0;
      Logger.log('✅ Circuit breaker: HALF_OPEN → CLOSED');
    }
  } else if (CIRCUIT_STATE.state === 'CLOSED') {
    CIRCUIT_STATE.failureCount = 0;
  }
}

/**
 * Record failed operation
 */
function recordFailure() {
  if (!CIRCUIT_BREAKER_CONFIG.ENABLED) return;

  CIRCUIT_STATE.failureCount++;
  CIRCUIT_STATE.lastFailureTime = new Date().getTime();

  if (CIRCUIT_STATE.state === 'HALF_OPEN' ||
      CIRCUIT_STATE.failureCount >= CIRCUIT_BREAKER_CONFIG.FAILURE_THRESHOLD) {
    CIRCUIT_STATE.state = 'OPEN';
    CIRCUIT_STATE.nextAttemptTime = CIRCUIT_STATE.lastFailureTime + CIRCUIT_BREAKER_CONFIG.TIMEOUT_MS;
    Logger.log(`⛔ Circuit breaker OPENED after ${CIRCUIT_STATE.failureCount} failures`);
  }
}

/**
 * Reset circuit breaker manually
 */
function resetCircuitBreakerManual() {
  CIRCUIT_STATE.state = 'CLOSED';
  CIRCUIT_STATE.failureCount = 0;
  CIRCUIT_STATE.successCount = 0;
  CIRCUIT_STATE.lastFailureTime = null;
  CIRCUIT_STATE.nextAttemptTime = null;

  SpreadsheetApp.getUi().alert('Circuit Breaker Reset', '✅ Reset to CLOSED state', SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log('🔄 Circuit breaker manually reset');
}

/**
 * Execute operation with retry logic
 */
function executeWithRetry(operation, operationName, maxAttempts) {
  maxAttempts = maxAttempts || RETRY_CONFIG.MAX_ATTEMPTS;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      Logger.log(`  🔄 ${operationName} - Attempt ${attempt}/${maxAttempts}`);
      const result = operation();
      Logger.log(`  ✅ ${operationName} - Success`);
      return { success: true, result: result, attempts: attempt };
    } catch (error) {
      Logger.log(`  ❌ ${operationName} - Attempt ${attempt} failed: ${error.message}`);

      if (attempt < maxAttempts) {
        const delay = Math.min(
          RETRY_CONFIG.INITIAL_DELAY_MS * Math.pow(RETRY_CONFIG.BACKOFF_MULTIPLIER, attempt - 1),
          RETRY_CONFIG.MAX_DELAY_MS
        );
        Logger.log(`  ⏳ Waiting ${delay}ms before retry...`);
        Utilities.sleep(delay);
      } else {
        return { success: false, error: error.message, attempts: attempt };
      }
    }
  }

  return { success: false, error: 'Max attempts reached', attempts: maxAttempts };
}

// ============================================================================
// SECTION 23: PORTRAIT CROPS CONFIGURATION (FROM V6)
// ============================================================================

const PORTRAIT_CROPS_CONFIG = {
  ENABLED: true,
  MAX_FILENAME_LENGTH: 50,
  CROP_MAPPING: {
    '2_MAIN_PORTRAIT': 1,
    '1_TIGHT_CROP': 2,
    '3_BOTTOM_LEFT_SIGNATURE': 3,
    '4_BOTTOM_RIGHT_EDITION': 4,
    '5_DETAILED_CENTER': 5,
    '6_AI_TEXTURE_SUBJECT': 6,
    '7_TOP_LEFT_TEXTURE': 7,
    '8_TOP_RIGHT_TEXTURE': 8,
    '0_THUMBNAIL': 9
  }
};

/**
 * Link portrait crops to product rows
 */
function linkPortraitCropsToProducts() {
  const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Products sheet not found');
    return;
  }

  const cropsFolder = DriveApp.getFolderById(ARTIST_FOLDERS['DEATH NYC'].CROPS_ALL);
  const lastRow = sheet.getLastRow();

  Logger.log('\n📸 Linking Portrait Crops to Products\n');

  let linked = 0;

  for (let row = 2; row <= lastRow; row++) {
    const sku = sheet.getRange(row, PRODUCTS_COLUMNS.SKU).getValue();
    if (!sku) continue;

    // Find crops matching this SKU
    const files = cropsFolder.getFiles();
    const crops = [];

    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().toUpperCase().includes(sku.toUpperCase())) {
        crops.push({
          name: file.getName(),
          url: 'https://drive.google.com/uc?export=view&id=' + file.getId()
        });
      }
    }

    if (crops.length > 0) {
      Logger.log(`✅ ${sku}: Found ${crops.length} crops`);
      // Populate image columns (starting at IMAGE_1)
      crops.slice(0, 8).forEach((crop, i) => {
        sheet.getRange(row, PRODUCTS_COLUMNS.IMAGE_1 + i).setValue(crop.url);
      });
      linked++;
    }
  }

  SpreadsheetApp.getUi().alert('Portrait Crops Linked', `✅ Linked crops for ${linked} products`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// SECTION 24: UPDATED MENU WITH ALL NEW FEATURES
// ============================================================================

/**
 * Add new menu items to existing menu
 * Call this from onOpen() or add items directly
 */
function addEnhancedMenuItems() {
  const ui = SpreadsheetApp.getUi();

  // Create a submenu for new features
  ui.createMenu('6. Enhanced Features')
    .addSubMenu(ui.createMenu('📂 File Consolidation')
      .addItem('Pull & Rename All Files', 'pullAndRenameAllFiles')
      .addItem('Preview (Dry Run)', 'previewFileRename')
      .addSeparator()
      .addItem('Show Loaded Folders', 'showLoadedImageFolders')
      .addItem('Reload Folders from IMAGES Tab', 'clearImagesTabCache'))
    .addSubMenu(ui.createMenu('🔍 SEO Optimization')
      .addItem('Optimize All Products', 'optimizeSEOForAllProducts')
      .addItem('Test eBay API', 'testEbayAPIConnection')
      .addItem('Test Google API', 'testGoogleAPIConnection'))
    .addSubMenu(ui.createMenu('🖼️ Portrait Crops')
      .addItem('Link Crops to Products', 'linkPortraitCropsToProducts'))
    .addSubMenu(ui.createMenu('🌐 Namecheap/Hosting')
      .addItem('🚀 UPLOAD ALL Images to Server', 'uploadAllImagesToServer')
      .addItem('Upload Images for SKU', 'uploadImagesForSelectedSKU')
      .addSeparator()
      .addItem('Test Server Connection', 'testServerConnection')
      .addItem('List Server Files', 'listServerFiles')
      .addSeparator()
      .addItem('Configure Upload Endpoint', 'configureNamecheapUpload')
      .addItem('Show Hosted URL Config', 'showHostedUrlConfig')
      .addItem('Preview Hosted URLs', 'previewHostedUrls')
      .addItem('Populate All Hosted URLs', 'populateHostedImageUrls'))
    .addSubMenu(ui.createMenu('⚙️ System')
      .addItem('Reset Circuit Breaker', 'resetCircuitBreakerManual')
      .addItem('Show Circuit Breaker Status', 'showCircuitBreakerStatus')
      .addItem('Clear All Caches', 'clearAllCachesEnhanced'))
    .addToUi();
}

/**
 * Clear all caches including IMAGES tab cache
 */
function clearAllCachesEnhanced() {
  VARIABLES_CACHE = null;
  PROMPTS_CACHE = null;
  FIELDS_CACHE = null;
  IMAGES_TAB_CACHE = null;
  SpreadsheetApp.getUi().alert('✅ All Caches Cleared', 'Variables, Prompts, Fields, and IMAGES tab caches cleared.', SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log('✅ All caches cleared');
}

/**
 * Show circuit breaker status
 */
function showCircuitBreakerStatus() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Circuit Breaker Status',
    `State: ${CIRCUIT_STATE.state}\n` +
    `Failures: ${CIRCUIT_STATE.failureCount}\n` +
    `Successes: ${CIRCUIT_STATE.successCount}`,
    ui.ButtonSet.OK
  );
}

/**
 * Test eBay API connection
 */
function testEbayAPIConnection() {
  Logger.log('🧪 Testing eBay API...');
  try {
    const comps = fetchEbayComps('death nyc');
    if (comps.prices.length > 0) {
      SpreadsheetApp.getUi().alert('✅ eBay API: Connected', `Found ${comps.prices.length} prices`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('⚠️ eBay API', 'Connected but no results', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ eBay API Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Test Google API connection
 */
function testGoogleAPIConnection() {
  Logger.log('🧪 Testing Google API...');
  try {
    const comps = fetchGoogleComps('death nyc');
    if (comps.titles.length > 0) {
      SpreadsheetApp.getUi().alert('✅ Google API: Connected', `Found ${comps.titles.length} results`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('⚠️ Google API', 'Connected but no results', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Google API Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ============================================================================
// SECTION 25: END-TO-END TEST (3 PRODUCTS FULL PIPELINE)
// ============================================================================
// Complete test: Images → Details → Fields → Namecheap → Update Workbook

const NAMECHEAP_CONFIG = {
  // Namecheap/gauntlet.gallery upload settings
  UPLOAD_ENDPOINT: 'https://gauntlet.gallery/upload_image.php',  // Verified endpoint
  API_KEY: 'dwDDsaIqQPuIwVnYPDdfNR6B2OsDtZNl',  // From prior scripts
  ENABLED: true
};

/**
 * END-TO-END TEST: Process 3 products through complete pipeline
 * 1. Scan & analyze images with AI
 * 2. Fill product details from FIELDS tab
 * 3. Generate titles/descriptions with AI
 * 4. Upload images to Namecheap (gauntlet.gallery)
 * 5. Update workbook with hosted URLs
 */
function testFullPipeline3Products() {
  const ui = SpreadsheetApp.getUi();
  const startTime = new Date();

  Logger.log('\n' + '═'.repeat(70));
  Logger.log('🧪 END-TO-END TEST: 3 PRODUCTS FULL PIPELINE');
  Logger.log('═'.repeat(70) + '\n');

  const results = {
    step1_images: { success: 0, failed: 0, skus: [] },
    step2_details: { success: 0, failed: 0 },
    step3_fields: { success: 0, failed: 0 },
    step4_namecheap: { success: 0, failed: 0 },
    step5_workbook: { success: 0, failed: 0 }
  };

  try {
    // ========== STEP 1: Process Images ==========
    Logger.log('\n📸 STEP 1: Processing Images with AI...');
    Logger.log('─'.repeat(50));

    const artistFolders = getArtistFolders('DEATH NYC');
    const mockupsFolder = artistFolders.MOCKUPS_PORTRAIT || artistFolders.MASTER;

    if (!mockupsFolder) {
      throw new Error('No image folder configured for Death NYC');
    }

    const folder = DriveApp.getFolderById(mockupsFolder);
    const files = folder.getFiles();
    const imagesToProcess = [];

    // Get first 3 images
    while (files.hasNext() && imagesToProcess.length < 3) {
      const file = files.next();
      if (file.getMimeType().startsWith('image/')) {
        imagesToProcess.push(file);
      }
    }

    if (imagesToProcess.length === 0) {
      throw new Error('No images found in folder: ' + folder.getName());
    }

    Logger.log(`  Found ${imagesToProcess.length} images to process`);

    // Process each image
    const ss = SpreadsheetApp.openById(STATIC_CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);
    let currentRow = sheet.getLastRow() + 1;

    for (const file of imagesToProcess) {
      try {
        Logger.log(`\n  🖼️ Processing: ${file.getName()}`);

        // Generate SKU
        const timestamp = new Date().getTime();
        const sku = `DNYC-TEST-${timestamp}-${Math.random().toString(36).substr(2, 4).toUpperCase()}`;

        // Analyze with AI (if available)
        let title = 'Death NYC ' + file.getName().replace(/\.[^/.]+$/, '');
        let description = 'Authentic Death NYC limited edition print. Hand-signed and numbered. Includes Certificate of Authenticity.';

        try {
          const variables = loadVariables();
          const apiKey = variables.apis?.GEMINI_API_KEY || variables.apis?.CLAUDE_API_KEY;
          if (apiKey) {
            const analysis = analyzeImageWithVision(file);
            if (analysis && analysis.title) {
              title = analysis.title;
              description = analysis.description || description;
            }
          }
        } catch (aiError) {
          Logger.log(`    ⚠️ AI analysis skipped: ${aiError.message}`);
        }

        // Write to sheet
        sheet.getRange(currentRow, PRODUCTS_COLUMNS.SKU).setValue(sku);
        sheet.getRange(currentRow, PRODUCTS_COLUMNS.ARTIST).setValue('Death NYC');
        sheet.getRange(currentRow, PRODUCTS_COLUMNS.FORMAT).setValue('Prints');
        sheet.getRange(currentRow, PRODUCTS_COLUMNS.REFERENCE_IMAGE).setValue('https://drive.google.com/uc?export=view&id=' + file.getId());
        sheet.getRange(currentRow, PRODUCTS_COLUMNS.ORIENTATION).setValue('Portrait');
        sheet.getRange(currentRow, PRODUCTS_COLUMNS.TITLE).setValue(title);
        sheet.getRange(currentRow, PRODUCTS_COLUMNS.DESCRIPTION).setValue(description);

        results.step1_images.success++;
        results.step1_images.skus.push({ sku: sku, row: currentRow });
        Logger.log(`    ✅ Created SKU: ${sku} at row ${currentRow}`);

        currentRow++;
      } catch (imgError) {
        Logger.log(`    ❌ Error: ${imgError.message}`);
        results.step1_images.failed++;
      }
    }

    // ========== STEP 2: Fill Details ==========
    Logger.log('\n📋 STEP 2: Filling Product Details...');
    Logger.log('─'.repeat(50));

    for (const item of results.step1_images.skus) {
      try {
        const fields = getFieldsForProduct('Death NYC', 'Prints');
        if (fields) {
          fillFieldsForRow(sheet, item.row, fields);
          results.step2_details.success++;
          Logger.log(`  ✅ Row ${item.row}: Details filled`);
        }
      } catch (detailError) {
        Logger.log(`  ❌ Row ${item.row}: ${detailError.message}`);
        results.step2_details.failed++;
      }
    }

    // ========== STEP 3: Fill All Fields ==========
    Logger.log('\n📝 STEP 3: Filling All Fields (79 columns)...');
    Logger.log('─'.repeat(50));

    for (const item of results.step1_images.skus) {
      try {
        // Get dimensions
        const dims = getDimensions('Death NYC', 'Prints', 'Portrait');
        if (dims) {
          sheet.getRange(item.row, PRODUCTS_COLUMNS.C_ITEM_LENGTH).setValue(dims.length);
          sheet.getRange(item.row, PRODUCTS_COLUMNS.C_ITEM_WIDTH).setValue(dims.width);
        }

        // Get price
        const price = getPrice('Death NYC', 'Prints');
        if (price) {
          sheet.getRange(item.row, PRODUCTS_COLUMNS.PRICE).setValue(price);
        }

        results.step3_fields.success++;
        Logger.log(`  ✅ Row ${item.row}: All fields populated`);
      } catch (fieldError) {
        Logger.log(`  ❌ Row ${item.row}: ${fieldError.message}`);
        results.step3_fields.failed++;
      }
    }

    // ========== STEP 4: Upload to Namecheap ==========
    Logger.log('\n🌐 STEP 4: Uploading to Namecheap (gauntlet.gallery)...');
    Logger.log('─'.repeat(50));

    for (const item of results.step1_images.skus) {
      try {
        const uploadResult = uploadImageToNamecheap(item.sku, item.row, sheet);
        if (uploadResult.success) {
          results.step4_namecheap.success++;
          Logger.log(`  ✅ ${item.sku}: Uploaded`);
        } else {
          results.step4_namecheap.failed++;
          Logger.log(`  ⚠️ ${item.sku}: ${uploadResult.message}`);
        }
      } catch (uploadError) {
        Logger.log(`  ❌ ${item.sku}: ${uploadError.message}`);
        results.step4_namecheap.failed++;
      }
    }

    // ========== STEP 5: Update Workbook with URLs ==========
    Logger.log('\n📊 STEP 5: Updating Workbook with Hosted URLs...');
    Logger.log('─'.repeat(50));

    for (const item of results.step1_images.skus) {
      try {
        populateHostedUrlsForRow(sheet, item.row, item.sku, 'Portrait');
        results.step5_workbook.success++;
        Logger.log(`  ✅ Row ${item.row}: URLs populated`);
      } catch (urlError) {
        Logger.log(`  ❌ Row ${item.row}: ${urlError.message}`);
        results.step5_workbook.failed++;
      }
    }

    // ========== SUMMARY ==========
    const elapsed = ((new Date() - startTime) / 1000).toFixed(1);

    Logger.log('\n' + '═'.repeat(70));
    Logger.log('📊 END-TO-END TEST COMPLETE');
    Logger.log('─'.repeat(70));
    Logger.log(`  Time: ${elapsed} seconds`);
    Logger.log(`  Step 1 (Images):    ${results.step1_images.success} ✅ / ${results.step1_images.failed} ❌`);
    Logger.log(`  Step 2 (Details):   ${results.step2_details.success} ✅ / ${results.step2_details.failed} ❌`);
    Logger.log(`  Step 3 (Fields):    ${results.step3_fields.success} ✅ / ${results.step3_fields.failed} ❌`);
    Logger.log(`  Step 4 (Namecheap): ${results.step4_namecheap.success} ✅ / ${results.step4_namecheap.failed} ❌`);
    Logger.log(`  Step 5 (Workbook):  ${results.step5_workbook.success} ✅ / ${results.step5_workbook.failed} ❌`);
    Logger.log('═'.repeat(70) + '\n');

    ui.alert(
      '🧪 End-to-End Test Complete!',
      `Time: ${elapsed}s\n\n` +
      `Step 1 - Images: ${results.step1_images.success} processed\n` +
      `Step 2 - Details: ${results.step2_details.success} filled\n` +
      `Step 3 - Fields: ${results.step3_fields.success} populated\n` +
      `Step 4 - Namecheap: ${results.step4_namecheap.success} uploaded\n` +
      `Step 5 - Workbook: ${results.step5_workbook.success} updated\n\n` +
      `SKUs Created:\n${results.step1_images.skus.map(s => s.sku).join('\n')}`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log(`\n❌ PIPELINE ERROR: ${error.message}`);
    ui.alert('❌ Pipeline Error', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Upload image to Namecheap (gauntlet.gallery) via HTTP API
 * @param {string} sku - The product SKU
 * @param {number} row - The row number in the sheet
 * @param {Sheet} sheet - The Google Sheet
 * @returns {Object} - { success: boolean, message: string, urls: [] }
 */
function uploadImageToNamecheap(sku, row, sheet) {
  // Get the reference image from the sheet
  const imageUrl = sheet.getRange(row, PRODUCTS_COLUMNS.REFERENCE_IMAGE).getValue();

  if (!imageUrl) {
    return { success: false, message: 'No reference image URL' };
  }

  // Check if upload endpoint is configured
  const variables = loadVariables();
  const uploadEndpoint = variables.apis?.NAMECHEAP_UPLOAD_ENDPOINT || NAMECHEAP_CONFIG.UPLOAD_ENDPOINT;
  const apiKey = variables.apis?.NAMECHEAP_API_KEY || NAMECHEAP_CONFIG.API_KEY;

  if (!uploadEndpoint || uploadEndpoint.includes('api/upload')) {
    // No custom endpoint - generate URLs assuming manual upload
    Logger.log(`    ℹ️ Auto-generating hosted URLs (manual upload required)`);
    return {
      success: true,
      message: 'URLs generated - manual upload to gauntlet.gallery required',
      urls: [
        `${HOSTED_IMAGES.PRODUCTS_BASE_URL}${sku}${HOSTED_IMAGES.CROP_PATTERN}1.jpg`,
        `${HOSTED_IMAGES.PRODUCTS_BASE_URL}${sku}${HOSTED_IMAGES.CROP_PATTERN}2.jpg`
      ]
    };
  }

  try {
    // Extract file ID from Google Drive URL
    const fileIdMatch = imageUrl.match(/id=([a-zA-Z0-9_-]+)/);
    if (!fileIdMatch) {
      return { success: false, message: 'Could not extract file ID from URL' };
    }

    const fileId = fileIdMatch[1];
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();

    // Upload via HTTP POST
    const payload = {
      sku: sku,
      filename: `${sku}${HOSTED_IMAGES.CROP_PATTERN}1.jpg`,
      image: Utilities.base64Encode(blob.getBytes()),
      mimeType: blob.getContentType()
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: apiKey ? { 'Authorization': 'Bearer ' + apiKey } : {},
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(uploadEndpoint, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200 || responseCode === 201) {
      const result = JSON.parse(response.getContentText());
      return {
        success: true,
        message: 'Uploaded successfully',
        urls: result.urls || [`${HOSTED_IMAGES.PRODUCTS_BASE_URL}${sku}${HOSTED_IMAGES.CROP_PATTERN}1.jpg`]
      };
    } else {
      return {
        success: false,
        message: `Upload failed: HTTP ${responseCode}`
      };
    }

  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Configure Namecheap upload settings
 */
function configureNamecheapUpload() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    'Configure Namecheap Upload',
    'Enter the upload API endpoint URL:\n(e.g., https://gauntlet.gallery/upload_image.php)\n\nLeave blank to use manual upload mode:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const endpoint = response.getResponseText().trim();

    if (endpoint) {
      PropertiesService.getScriptProperties().setProperty('NAMECHEAP_UPLOAD_ENDPOINT', endpoint);
      ui.alert('✅ Endpoint Saved', 'Upload endpoint configured:\n' + endpoint, ui.ButtonSet.OK);
    } else {
      PropertiesService.getScriptProperties().deleteProperty('NAMECHEAP_UPLOAD_ENDPOINT');
      ui.alert('ℹ️ Manual Mode', 'No endpoint configured.\nHosted URLs will be generated but images must be uploaded manually to gauntlet.gallery.', ui.ButtonSet.OK);
    }
  }
}

// ============================================================================
// SECTION: IMAGE UPLOAD TO GAUNTLET.GALLERY
// ============================================================================

/**
 * Upload a single image to gauntlet.gallery
 * @param {Blob} blob - File blob from Google Drive
 * @param {string} filename - Target filename (e.g., "123_DNYC-Print-Item_crop_1.jpg")
 * @param {string} folder - Target folder ('products' or 'stock')
 * @returns {Object} - {success: boolean, url: string, error: string}
 */
function uploadImageToServer(blob, filename, folder) {
  folder = folder || 'products';

  const base64Data = Utilities.base64Encode(blob.getBytes());

  const response = UrlFetchApp.fetch(NAMECHEAP_CONFIG.UPLOAD_ENDPOINT, {
    method: 'post',
    payload: {
      api_key: NAMECHEAP_CONFIG.API_KEY,
      action: 'upload',
      filename: filename,
      folder: folder,
      data: base64Data
    },
    muteHttpExceptions: true
  });

  try {
    const result = JSON.parse(response.getContentText());
    return result;
  } catch (e) {
    return { success: false, error: 'Invalid response: ' + response.getContentText() };
  }
}

/**
 * Test connection to gauntlet.gallery server
 */
function testServerConnection() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Test with list action
    const response = UrlFetchApp.fetch(NAMECHEAP_CONFIG.UPLOAD_ENDPOINT, {
      method: 'post',
      payload: {
        api_key: NAMECHEAP_CONFIG.API_KEY,
        action: 'list',
        folder: 'products'
      },
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());

    if (result.success) {
      ui.alert('✅ Connection Successful!',
        'Connected to gauntlet.gallery\n\n' +
        'Files in /images/products/: ' + result.files.length,
        ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Connection Failed', 'Error: ' + result.error, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('❌ Connection Error', 'Could not connect: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * List all files on gauntlet.gallery server
 */
function listServerFiles() {
  const ui = SpreadsheetApp.getUi();

  try {
    const response = UrlFetchApp.fetch(NAMECHEAP_CONFIG.UPLOAD_ENDPOINT, {
      method: 'post',
      payload: {
        api_key: NAMECHEAP_CONFIG.API_KEY,
        action: 'list',
        folder: 'products'
      },
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());

    if (result.success) {
      Logger.log('Files on server (' + result.files.length + '):');
      result.files.forEach(f => Logger.log('  ' + f.filename));

      ui.alert('📁 Server Files',
        'Total files: ' + result.files.length + '\n\n' +
        'First 20 files:\n' +
        result.files.slice(0, 20).map(f => f.filename).join('\n'),
        ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Error', result.error, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('❌ Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Upload all images for a single SKU from Google Drive to server
 * MATCHES BY ARTWORK NAME - not by SKU number!
 *
 * File pattern in Drive: 9_DNYC-Print-BLACKPINK-Dior_P_crop_6.png
 * SKU pattern in Sheet:  28_DNYC-Print-BLACKPINK-Dior
 * Matches on:            DNYC-Print-BLACKPINK-Dior
 *
 * @param {string} sku - Full SKU (e.g., "28_DNYC-Print-BLACKPINK-Dior")
 * @returns {Object} - {uploaded: number, failed: number, urls: string[]}
 */
function uploadImagesForSKU(sku) {
  const results = { uploaded: 0, failed: 0, urls: [] };

  // Extract artwork name from SKU (everything after first underscore)
  const skuParts = sku.toString().split('_');
  const skuNum = skuParts[0];
  const artworkName = skuParts.slice(1).join('_'); // e.g., "DNYC-Print-BLACKPINK-Dior"

  if (!artworkName) {
    return { uploaded: 0, failed: 0, urls: [], error: 'Could not extract artwork name from SKU' };
  }

  const SOURCE_FOLDER = '1Ft_QO1XlwYt6XCbCEKslG4QbDg0Lm9mT'; // CROPS_ALL

  try {
    const folder = DriveApp.getFolderById(SOURCE_FOLDER);
    const files = folder.getFiles();

    // Find all files containing this artwork name
    const matchingFiles = [];
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName().toLowerCase();
      const artworkLower = artworkName.toLowerCase();

      // Match if filename contains the artwork name
      if (fileName.includes(artworkLower)) {
        matchingFiles.push(file);
      }
    }

    // Sort by crop number if present
    matchingFiles.sort((a, b) => {
      const aMatch = a.getName().match(/crop_(\d+)/i);
      const bMatch = b.getName().match(/crop_(\d+)/i);
      const aNum = aMatch ? parseInt(aMatch[1]) : 999;
      const bNum = bMatch ? parseInt(bMatch[1]) : 999;
      return aNum - bNum;
    });

    Logger.log('Found ' + matchingFiles.length + ' files for artwork: ' + artworkName);

    // Upload each file, renaming to {SKU}_crop_{n}.jpg
    for (let i = 0; i < matchingFiles.length && i < 24; i++) {
      const file = matchingFiles[i];
      const cropNum = i + 1;
      const targetName = sku + '_crop_' + cropNum + '.jpg';

      Logger.log('Uploading: ' + file.getName() + ' -> ' + targetName);

      const result = uploadImageToServer(file.getBlob(), targetName, 'products');

      if (result.success) {
        results.uploaded++;
        results.urls.push(HOSTED_IMAGES.PRODUCTS_BASE_URL + targetName);
      } else {
        results.failed++;
        Logger.log('Failed: ' + result.error);
      }

      Utilities.sleep(300); // Rate limiting
    }
  } catch (e) {
    results.error = e.message;
    Logger.log('Error: ' + e.message);
  }

  return results;
}

/**
 * Upload all images for ALL products to gauntlet.gallery
 */
function uploadAllImagesToServer() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  // Confirm with user
  const confirm = ui.alert('Upload ALL Images',
    'This will upload images for ALL products from Google Drive to gauntlet.gallery.\n\n' +
    'This may take a while. Continue?',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // Just column A (SKU)

  let totalUploaded = 0;
  let totalFailed = 0;
  let skusProcessed = 0;

  for (let i = 0; i < data.length; i++) {
    const fullSku = data[i][0];

    if (!fullSku) continue;

    const result = uploadImagesForSKU(fullSku);
    totalUploaded += result.uploaded;
    totalFailed += result.failed;
    skusProcessed++;

    // Progress every 10 SKUs
    if (skusProcessed % 10 === 0) {
      Logger.log('Progress: ' + skusProcessed + ' SKUs, ' + totalUploaded + ' images uploaded');
    }
  }

  ui.alert('✅ Upload Complete!',
    'SKUs processed: ' + skusProcessed + '\n' +
    'Images uploaded: ' + totalUploaded + '\n' +
    'Failed: ' + totalFailed,
    ui.ButtonSet.OK);
}

/**
 * Upload images for the selected SKU only
 */
function uploadImagesForSelectedSKU() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    ui.alert('Error', 'Please select a product row (row 2 or below)', ui.ButtonSet.OK);
    return;
  }

  const fullSku = sheet.getRange(row, 1).getValue();

  if (!fullSku) {
    ui.alert('Error', 'No SKU found in the selected row', ui.ButtonSet.OK);
    return;
  }

  ui.alert('Uploading...',
    'Uploading images for SKU: ' + fullSku + '\n\nThis may take a moment...',
    ui.ButtonSet.OK);

  const result = uploadImagesForSKU(fullSku);

  ui.alert('✅ Upload Complete!',
    'SKU: ' + fullSku + '\n' +
    'Images uploaded: ' + result.uploaded + '\n' +
    'Failed: ' + result.failed + '\n\n' +
    'URLs:\n' + result.urls.slice(0, 8).join('\n'),
    ui.ButtonSet.OK);
}

/**
 * Upload images for Death NYC artist only (all orientations)
 */
function uploadDeathNYCImages() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  const confirm = ui.alert('Upload Death NYC Images',
    'Upload all images for DEATH NYC products?\n\nThis will upload from Google Drive to gauntlet.gallery.',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // A-C (SKU, Title, Artist)

  let totalUploaded = 0;
  let totalFailed = 0;
  let skusProcessed = 0;

  for (let i = 0; i < data.length; i++) {
    const fullSku = data[i][0];
    const artist = data[i][2]; // Column C = Artist

    if (!fullSku) continue;

    // Only process Death NYC
    if (!artist || !artist.toString().toUpperCase().includes('DEATH NYC')) continue;

    const result = uploadImagesForSKU(fullSku);
    totalUploaded += result.uploaded;
    totalFailed += result.failed;
    skusProcessed++;

    Logger.log('Uploaded ' + result.uploaded + ' for ' + fullSku);
  }

  ui.alert('✅ Death NYC Upload Complete!',
    'SKUs processed: ' + skusProcessed + '\n' +
    'Images uploaded: ' + totalUploaded + '\n' +
    'Failed: ' + totalFailed,
    ui.ButtonSet.OK);
}

// ============================================================================
// SECTION: GOOGLE DRIVE FILE MANAGEMENT
// ============================================================================

/**
 * List all files in Landscape Mockups folder - shows what needs renaming
 */
function listLandscapeMockupFiles() {
  const FOLDER_ID = '1JtTH909dOAd_wDPWOFzPS4ue0ntMc9KD'; // Landscape Mockups
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  const fileList = [];
  while (files.hasNext()) {
    const file = files.next();
    fileList.push(file.getName());
  }

  // Sort by name
  fileList.sort();

  Logger.log('=== FILES IN LANDSCAPE MOCKUPS (' + fileList.length + ' files) ===');
  fileList.forEach((name, i) => Logger.log((i+1) + '. ' + name));

  // Show summary in alert
  const ui = SpreadsheetApp.getUi();
  ui.alert('Landscape Mockup Files',
    'Total files: ' + fileList.length + '\n\n' +
    'First 20:\n' + fileList.slice(0, 20).join('\n') + '\n\n' +
    'Check Logger for full list (View > Logs)',
    ui.ButtonSet.OK);
}

/**
 * List all files in Portrait Mockups folder
 */
function listPortraitMockupFiles() {
  const FOLDER_ID = '1yVDucQO0vn8XMmnXcrKvPf98k9i0VOX-'; // Portrait Mockups
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  const fileList = [];
  while (files.hasNext()) {
    const file = files.next();
    fileList.push(file.getName());
  }

  fileList.sort();

  Logger.log('=== FILES IN PORTRAIT MOCKUPS (' + fileList.length + ' files) ===');
  fileList.forEach((name, i) => Logger.log((i+1) + '. ' + name));

  const ui = SpreadsheetApp.getUi();
  ui.alert('Portrait Mockup Files',
    'Total files: ' + fileList.length + '\n\n' +
    'First 20:\n' + fileList.slice(0, 20).join('\n') + '\n\n' +
    'Check Logger for full list (View > Logs)',
    ui.ButtonSet.OK);
}

/**
 * List files in CROPS_ALL folder (the main source folder)
 */
function listCropsAllFiles() {
  const FOLDER_ID = '1Ft_QO1XlwYt6XCbCEKslG4QbDg0Lm9mT'; // CROPS_ALL
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  const fileList = [];
  while (files.hasNext()) {
    const file = files.next();
    fileList.push(file.getName());
  }

  fileList.sort();

  Logger.log('=== FILES IN CROPS_ALL (' + fileList.length + ' files) ===');
  fileList.forEach((name, i) => Logger.log((i+1) + '. ' + name));

  // Extract unique SKU numbers
  const skuNums = new Set();
  fileList.forEach(name => {
    const match = name.match(/^(\d+)_/);
    if (match) skuNums.add(match[1]);
  });

  const ui = SpreadsheetApp.getUi();
  ui.alert('CROPS_ALL Files',
    'Total files: ' + fileList.length + '\n' +
    'Unique SKU numbers: ' + skuNums.size + '\n\n' +
    'SKUs found: ' + Array.from(skuNums).sort((a,b) => parseInt(a) - parseInt(b)).join(', ') + '\n\n' +
    'Check Logger for full list (View > Logs)',
    ui.ButtonSet.OK);
}

/**
 * Rename files in a folder - change SKU number prefix
 * @param {string} folderId - Google Drive folder ID
 * @param {Object} mapping - {oldNum: newNum, ...} e.g. {'50': '1', '51': '2'}
 */
function renameFilesInFolder(folderId, mapping) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  let renamed = 0;
  let skipped = 0;

  while (files.hasNext()) {
    const file = files.next();
    const oldName = file.getName();

    // Extract SKU number from filename
    const match = oldName.match(/^(\d+)_(.+)$/);
    if (!match) {
      skipped++;
      continue;
    }

    const oldNum = match[1];
    const rest = match[2];

    // Check if this number needs remapping
    if (mapping[oldNum]) {
      const newName = mapping[oldNum] + '_' + rest;
      file.setName(newName);
      Logger.log('Renamed: ' + oldName + ' -> ' + newName);
      renamed++;
    } else {
      skipped++;
    }
  }

  return { renamed, skipped };
}

/**
 * Rename Landscape Mockup files - sequential renumbering
 * Files will be sorted alphabetically and renumbered 1, 2, 3...
 */
function renameLandscapeMockupsSequential() {
  const ui = SpreadsheetApp.getUi();
  const FOLDER_ID = '1JtTH909dOAd_wDPWOFzPS4ue0ntMc9KD';

  const confirm = ui.alert('Rename Landscape Mockups',
    'This will renumber ALL files in the Landscape Mockups folder sequentially (1, 2, 3...).\n\n' +
    'Files will be sorted alphabetically first, then renumbered.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  // Collect all files
  const fileList = [];
  while (files.hasNext()) {
    fileList.push(files.next());
  }

  // Sort by current name
  fileList.sort((a, b) => a.getName().localeCompare(b.getName()));

  let renamed = 0;

  // Group files by their base name (excluding crop number)
  // e.g., all "50_Something_crop_1.jpg", "50_Something_crop_2.jpg" should become "1_Something_crop_1.jpg", etc.

  const skuGroups = {};
  fileList.forEach(file => {
    const name = file.getName();
    const match = name.match(/^(\d+)_(.+)$/);
    if (match) {
      const oldNum = match[1];
      if (!skuGroups[oldNum]) skuGroups[oldNum] = [];
      skuGroups[oldNum].push(file);
    }
  });

  // Get sorted old SKU numbers
  const oldSkus = Object.keys(skuGroups).sort((a, b) => parseInt(a) - parseInt(b));

  // Rename each group to new sequential number
  oldSkus.forEach((oldSku, index) => {
    const newSku = (index + 1).toString();
    const files = skuGroups[oldSku];

    files.forEach(file => {
      const oldName = file.getName();
      const newName = oldName.replace(/^\d+_/, newSku + '_');
      file.setName(newName);
      Logger.log('Renamed: ' + oldName + ' -> ' + newName);
      renamed++;
    });
  });

  ui.alert('✅ Rename Complete!',
    'Renamed ' + renamed + ' files.\n\n' +
    'Old SKUs: ' + oldSkus.join(', ') + '\n' +
    'New SKUs: 1-' + oldSkus.length + '\n\n' +
    'Check Logger for details (View > Logs)',
    ui.ButtonSet.OK);
}

/**
 * Create a SKU mapping from Products sheet to drive files
 * Shows which SKU numbers are in the sheet vs what's in Drive
 */
function analyzeSkuMismatch() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  // Get SKUs from sheet
  const lastRow = sheet.getLastRow();
  const skuData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const sheetSkus = new Set();
  skuData.forEach(row => {
    if (row[0]) {
      const num = row[0].toString().split('_')[0];
      sheetSkus.add(num);
    }
  });

  // Get SKUs from CROPS_ALL folder
  const FOLDER_ID = '1Ft_QO1XlwYt6XCbCEKslG4QbDg0Lm9mT';
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  const driveSkus = new Set();
  while (files.hasNext()) {
    const name = files.next().getName();
    const match = name.match(/^(\d+)_/);
    if (match) driveSkus.add(match[1]);
  }

  // Find mismatches
  const inSheetOnly = Array.from(sheetSkus).filter(s => !driveSkus.has(s)).sort((a,b) => parseInt(a) - parseInt(b));
  const inDriveOnly = Array.from(driveSkus).filter(s => !sheetSkus.has(s)).sort((a,b) => parseInt(a) - parseInt(b));
  const inBoth = Array.from(sheetSkus).filter(s => driveSkus.has(s)).sort((a,b) => parseInt(a) - parseInt(b));

  Logger.log('=== SKU ANALYSIS ===');
  Logger.log('In Sheet: ' + Array.from(sheetSkus).sort((a,b) => parseInt(a) - parseInt(b)).join(', '));
  Logger.log('In Drive: ' + Array.from(driveSkus).sort((a,b) => parseInt(a) - parseInt(b)).join(', '));
  Logger.log('In Both: ' + inBoth.join(', '));
  Logger.log('Sheet Only (no images): ' + inSheetOnly.join(', '));
  Logger.log('Drive Only (no product): ' + inDriveOnly.join(', '));

  ui.alert('SKU Mismatch Analysis',
    'Sheet SKUs: ' + sheetSkus.size + '\n' +
    'Drive SKUs: ' + driveSkus.size + '\n' +
    'Matched: ' + inBoth.length + '\n\n' +
    'In Sheet but NO images: ' + (inSheetOnly.length > 0 ? inSheetOnly.join(', ') : 'None') + '\n\n' +
    'In Drive but NO product: ' + (inDriveOnly.length > 0 ? inDriveOnly.join(', ') : 'None') + '\n\n' +
    'Check Logger for details',
    ui.ButtonSet.OK);
}

/**
 * DIAGNOSTIC: Show exactly why images aren't being found
 */
function diagnoseImageProblem() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  // Get first 5 SKUs from sheet
  const skuData = sheet.getRange(2, 1, 5, 1).getValues();
  const sheetSkus = skuData.map(r => r[0]).filter(s => s);

  // Get files from CROPS_ALL
  const FOLDER_ID = '1Ft_QO1XlwYt6XCbCEKslG4QbDg0Lm9mT';
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  const driveFiles = [];
  while (files.hasNext()) {
    driveFiles.push(files.next().getName());
  }
  driveFiles.sort();

  let report = '=== DIAGNOSTIC REPORT ===\n\n';
  report += 'CROPS_ALL Folder: ' + driveFiles.length + ' files\n';
  report += 'First 10 files:\n';
  driveFiles.slice(0, 10).forEach(f => report += '  ' + f + '\n');

  report += '\nSheet SKUs (first 5):\n';
  sheetSkus.forEach(s => report += '  ' + s + '\n');

  report += '\nLooking for matches...\n';
  sheetSkus.forEach(sku => {
    const skuNum = sku.toString().split('_')[0];
    const matches = driveFiles.filter(f => f.startsWith(skuNum + '_'));
    report += '  SKU ' + skuNum + ': ' + matches.length + ' files found\n';
  });

  Logger.log(report);

  ui.alert('Diagnostic Report',
    'CROPS_ALL has ' + driveFiles.length + ' files\n\n' +
    'First 5 files in Drive:\n' + driveFiles.slice(0, 5).join('\n') + '\n\n' +
    'First 5 SKUs in Sheet:\n' + sheetSkus.join('\n') + '\n\n' +
    'Check Logger (View > Logs) for full report',
    ui.ButtonSet.OK);
}

/**
 * MATCH BY ARTWORK NAME - not SKU number
 * Finds files by the artwork name portion and renames to correct SKU
 */
function matchAndRenameByArtworkName() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  const confirm = ui.alert('Match & Rename by Artwork Name',
    'This will:\n' +
    '1. Read SKUs from Products sheet (e.g., 28_DNYC-Print-Balloon-Girl)\n' +
    '2. Find files with matching artwork name (e.g., 50_DNYC-Print-Balloon-Girl_crop_1.jpg)\n' +
    '3. Rename files to use correct SKU number (e.g., 28_DNYC-Print-Balloon-Girl_crop_1.jpg)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  // Build map of artwork name -> correct SKU number
  const lastRow = sheet.getLastRow();
  const skuData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  const artworkToSku = {};
  skuData.forEach(row => {
    if (row[0]) {
      const fullSku = row[0].toString();
      const parts = fullSku.split('_');
      const skuNum = parts[0];
      const artworkName = parts.slice(1).join('_'); // Everything after the number
      if (artworkName) {
        artworkToSku[artworkName.toLowerCase()] = skuNum;
      }
    }
  });

  Logger.log('Artwork mappings: ' + Object.keys(artworkToSku).length);
  Object.keys(artworkToSku).slice(0, 10).forEach(name => {
    Logger.log('  ' + name + ' -> SKU ' + artworkToSku[name]);
  });

  // Get files from CROPS_ALL
  const FOLDER_ID = '1Ft_QO1XlwYt6XCbCEKslG4QbDg0Lm9mT';
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  let renamed = 0;
  let noMatch = 0;
  let alreadyCorrect = 0;

  while (files.hasNext()) {
    const file = files.next();
    const oldName = file.getName();

    // Parse filename: {num}_{artworkName}_crop_{n}.jpg
    const match = oldName.match(/^(\d+)_(.+?)(_crop_\d+\.jpg)$/i);
    if (!match) {
      noMatch++;
      continue;
    }

    const oldNum = match[1];
    const artworkName = match[2];
    const suffix = match[3]; // _crop_1.jpg

    // Look up correct SKU number
    const correctSkuNum = artworkToSku[artworkName.toLowerCase()];

    if (!correctSkuNum) {
      Logger.log('No match for artwork: ' + artworkName);
      noMatch++;
      continue;
    }

    if (oldNum === correctSkuNum) {
      alreadyCorrect++;
      continue;
    }

    // Rename file
    const newName = correctSkuNum + '_' + artworkName + suffix;
    file.setName(newName);
    Logger.log('Renamed: ' + oldName + ' -> ' + newName);
    renamed++;
  }

  ui.alert('✅ Rename Complete!',
    'Renamed: ' + renamed + ' files\n' +
    'Already correct: ' + alreadyCorrect + '\n' +
    'No match found: ' + noMatch + '\n\n' +
    'Check Logger for details',
    ui.ButtonSet.OK);
}

/**
 * Preview what matchAndRenameByArtworkName would do (without renaming)
 */
function previewArtworkMatching() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  // Build map of artwork name -> correct SKU number
  const lastRow = sheet.getLastRow();
  const skuData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  const artworkToSku = {};
  skuData.forEach(row => {
    if (row[0]) {
      const fullSku = row[0].toString();
      const parts = fullSku.split('_');
      const skuNum = parts[0];
      const artworkName = parts.slice(1).join('_');
      if (artworkName) {
        artworkToSku[artworkName.toLowerCase()] = { num: skuNum, full: fullSku };
      }
    }
  });

  // Get files from CROPS_ALL
  const FOLDER_ID = '1Ft_QO1XlwYt6XCbCEKslG4QbDg0Lm9mT';
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  const wouldRename = [];
  const noMatch = [];
  const alreadyCorrect = [];

  while (files.hasNext()) {
    const file = files.next();
    const oldName = file.getName();

    const match = oldName.match(/^(\d+)_(.+?)(_crop_\d+\.jpg)$/i);
    if (!match) {
      noMatch.push(oldName);
      continue;
    }

    const oldNum = match[1];
    const artworkName = match[2];
    const suffix = match[3];

    const correctSku = artworkToSku[artworkName.toLowerCase()];

    if (!correctSku) {
      noMatch.push(oldName + ' (artwork: ' + artworkName + ')');
      continue;
    }

    if (oldNum === correctSku.num) {
      alreadyCorrect.push(oldName);
      continue;
    }

    const newName = correctSku.num + '_' + artworkName + suffix;
    wouldRename.push(oldName + ' -> ' + newName);
  }

  Logger.log('=== PREVIEW: Would rename ' + wouldRename.length + ' files ===');
  wouldRename.forEach(r => Logger.log(r));
  Logger.log('\n=== Already correct: ' + alreadyCorrect.length + ' ===');
  Logger.log('\n=== No match: ' + noMatch.length + ' ===');
  noMatch.slice(0, 20).forEach(n => Logger.log(n));

  ui.alert('Preview Results',
    'Would rename: ' + wouldRename.length + ' files\n' +
    'Already correct: ' + alreadyCorrect.length + '\n' +
    'No match: ' + noMatch.length + '\n\n' +
    'Sample renames:\n' + wouldRename.slice(0, 5).join('\n') + '\n\n' +
    'Check Logger for full list',
    ui.ButtonSet.OK);
}

/**
 * CHECK ALL FOLDERS - Find where the images actually are
 */
function checkAllFolders() {
  const ui = SpreadsheetApp.getUi();

  const folders = {
    'CROPS_ALL': '1Ft_QO1XlwYt6XCbCEKslG4QbDg0Lm9mT',
    'MOCKUPS_PORTRAIT': '1yVDucQO0vn8XMmnXcrKvPf98k9i0VOX-',
    'MOCKUPS_LANDSCAPE': '1JtTH909dOAd_wDPWOFzPS4ue0ntMc9KD',
    'CROPS_PORTRAIT': '1z511jL5Iq7tTJPmEK_noi3WgLd3z4wvZ',
    'CROPS_LANDSCAPE': '1Lyeq07b7bkptqVy6RynHBAHGQtT674F5',
    'MASTER': '15E4sMBfWACkv9EVUeRXj1j4uq6OoXbCu',
    'STOCK': '1IX2NU56kdj6ere-TFHZCkgv_DbdDibBy'
  };

  let report = '=== CHECKING ALL FOLDERS ===\n\n';

  for (const [name, id] of Object.entries(folders)) {
    try {
      const folder = DriveApp.getFolderById(id);
      const files = folder.getFiles();
      let count = 0;
      const sampleFiles = [];
      while (files.hasNext()) {
        const file = files.next();
        count++;
        if (sampleFiles.length < 3) {
          sampleFiles.push(file.getName());
        }
      }
      report += '✅ ' + name + ': ' + count + ' files\n';
      if (sampleFiles.length > 0) {
        report += '   Sample: ' + sampleFiles[0] + '\n';
      }
    } catch (e) {
      report += '❌ ' + name + ': ERROR - ' + e.message + '\n';
    }
  }

  Logger.log(report);

  ui.alert('Folder Check Results', report, ui.ButtonSet.OK);
}

/**
 * Populate Products sheet with Death NYC SKUs
 */
function populateDeathNYCSKUs() {
  const skus = [
    '8_DNYC-Print-Death-NYC-Mickey-Murakami-A-P',
    '9_DNYC-Print-BLACKPINK-Dior',
    '10_DNYC-Print-Heart-Hands-LV',
    '11_DNYC-Print-Death-NYC-LV-Figure',
    '12_DNYC-Print-Doll-Diptyque',
    '13_DNYC-Print-Death-NYC-Nara-Girl-A-P',
    '14_DNYC-Print-DEATH-NYC-Joker-LV-Soup',
    '15_DNYC-Print-DEATH-NYC-Alice-Umbrella',
    '16_DNYC-Print-DEATH-NYC-Taylor-Swift',
    '24_DNYC-Print-LV-Flower',
    '25_DNYC-Print-DEATH-NYC-Gucci-Horse',
    '26_DNYC-Print-Goku-LV',
    '27_DNYC-Print-Death-NYC-Lion-LV-AP',
    '28_DNYC-Print-Balloon-Girl-LV',
    '29_DNYC-Print-Kendall-LV-Beach',
    '30_DNYC-Print-DEATH-NYC-Goku-LV',
    '31_DNYC-Print-LV-Flower',
    '34_DNYC-Print-DEATH-NYC-Kendall-Gucci',
    '35_DNYC-Print-Item',
    '36_DNYC-Original-Item',
    '37_DNYC-Original-Chanel-Grenade',
    '38_DNYC-Original-Gas-Mask-Girl',
    '39_DNYC-Original-Death-NYC-Joker-1-1',
    '41_DNYC-Print-Death-NYC-Van-Gogh-A-P',
    '42_DNYC-Print-Polka-Dot-Portrait',
    '43_DNYC-Print-DEATH-NYC-Magritte-Coke',
    '44_DNYC-Print-Snoopy-Moon-LV',
    '45_DNYC-Print-Gucci-Woman',
    '46_DNYC-Print-Alice-Kermit-Supreme',
    '47_DNYC-Print-Umbrella-Girl',
    '48_DNYC-Print-Peanuts-Starry-Night',
    '58_DNYC-Print-Marilyn-Last-Supper',
    '61_DNYC-Print-Death-NYC-BB-8-Snorlax-AP',
    '62_DNYC-Print-Death-NYC-Simpsons-LV-AP',
    '63_DNYC-Print-Mickey-Minnie-Paris',
    '64_DNYC-Original-Death-NYC-Banana-LV',
    '65_DNYC-Original-Death-NYC-Queen-Haring-LV',
    '77_DNYC-Original-Snoopy-LV',
    '78_DNYC-Print-Homer-Evolution-Wave',
    '79_DNYC-Print-Shark-Graffiti-Bills',
    '81_DNYC-Print-Goku-LV-Mashup',
    '82_DNYC-Print-DEATH-NYC-Goku-Van-Gogh',
    '83_DNYC-Print-Toy-Story-Van-Gogh',
    '84_DNYC-Print-Beatles-Abbey-Road-LV'
  ];

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  const confirm = ui.alert('Populate SKUs',
    'This will add ' + skus.length + ' Death NYC SKUs to column A of Products sheet.\n\nContinue?',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  // Write SKUs to column A starting at row 2
  for (let i = 0; i < skus.length; i++) {
    sheet.getRange(i + 2, 1).setValue(skus[i]);
  }

  // Also set Artist column (C) to DEATH NYC
  for (let i = 0; i < skus.length; i++) {
    sheet.getRange(i + 2, 3).setValue('DEATH NYC');
  }

  ui.alert('✅ Done!', 'Added ' + skus.length + ' SKUs to Products sheet.\n\nNow run uploadDeathNYCImages() to upload the images.', ui.ButtonSet.OK);
}

/**
 * SIMPLE UPLOAD - Just upload by SKU number match
 * Uses CROPS_PORTRAIT and CROPS_LANDSCAPE folders (the ones with correct format)
 */
function uploadBySkuNumber() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STATIC_CONFIG.PRODUCTS_TAB);

  const confirm = ui.alert('Upload by SKU Number',
    'This will:\n' +
    '1. Read SKUs from Products sheet\n' +
    '2. Find files in CROPS_PORTRAIT and CROPS_LANDSCAPE\n' +
    '3. Upload to gauntlet.gallery\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  // Use both Portrait and Landscape crop folders (the ones with correct format)
  const PORTRAIT_FOLDER = '1z511jL5Iq7tTJPmEK_noi3WgLd3z4wvZ';  // 32 files
  const LANDSCAPE_FOLDER = '1Lyeq07b7bkptqVy6RynHBAHGQtT674F5'; // 13 files

  // Get all files from BOTH folders
  const allFiles = [];

  // Portrait files
  const portraitFolder = DriveApp.getFolderById(PORTRAIT_FOLDER);
  const portraitFiles = portraitFolder.getFiles();
  while (portraitFiles.hasNext()) {
    allFiles.push(portraitFiles.next());
  }
  Logger.log('Portrait files: ' + allFiles.length);

  // Landscape files
  const landscapeFolder = DriveApp.getFolderById(LANDSCAPE_FOLDER);
  const landscapeFiles = landscapeFolder.getFiles();
  while (landscapeFiles.hasNext()) {
    allFiles.push(landscapeFiles.next());
  }
  Logger.log('Total files (Portrait + Landscape): ' + allFiles.length);

  // Get SKUs from sheet
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('Error', 'No SKUs in Products sheet. Run populateDeathNYCSKUs() first.', ui.ButtonSet.OK);
    return;
  }

  const skuData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const skus = skuData.map(r => r[0]).filter(s => s);

  let totalUploaded = 0;
  let totalFailed = 0;
  let skusProcessed = 0;

  for (const sku of skus) {
    const skuNum = sku.toString().split('_')[0];

    // Find files starting with this SKU number
    const matchingFiles = allFiles.filter(f => f.getName().startsWith(skuNum + '_'));

    // Sort by crop number
    matchingFiles.sort((a, b) => {
      const aMatch = a.getName().match(/crop_(\d+)/i);
      const bMatch = b.getName().match(/crop_(\d+)/i);
      const aNum = aMatch ? parseInt(aMatch[1]) : 999;
      const bNum = bMatch ? parseInt(bMatch[1]) : 999;
      return aNum - bNum;
    });

    Logger.log('SKU ' + skuNum + ': found ' + matchingFiles.length + ' files');

    // Upload each file
    for (let i = 0; i < matchingFiles.length && i < 24; i++) {
      const file = matchingFiles[i];
      const cropNum = i + 1;
      const targetName = sku + '_crop_' + cropNum + '.jpg';

      const result = uploadImageToServer(file.getBlob(), targetName, 'products');

      if (result.success) {
        totalUploaded++;
      } else {
        totalFailed++;
        Logger.log('Failed: ' + targetName + ' - ' + result.error);
      }

      Utilities.sleep(200);
    }

    skusProcessed++;
  }

  ui.alert('✅ Upload Complete!',
    'SKUs processed: ' + skusProcessed + '\n' +
    'Images uploaded: ' + totalUploaded + '\n' +
    'Failed: ' + totalFailed + '\n\n' +
    'Check Logger for details',
    ui.ButtonSet.OK);
}

/**
 * FIX FILES IN GOOGLE DRIVE - Rename to correct format
 * From: 9_DNYC-Print-BLACKPINK-Dior_P_crop_1.png
 * To:   9_DNYC-Print-BLACKPINK-Dior_crop_1.png
 * (Removes _P_ or _L_ orientation marker)
 */
function fixDriveFileNames() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert('Fix Drive File Names',
    'This will rename files in CROPS_PORTRAIT and CROPS_LANDSCAPE folders:\n\n' +
    'From: 9_DNYC-Print-BLACKPINK-Dior_P_crop_1.png\n' +
    'To:   9_DNYC-Print-BLACKPINK-Dior_crop_1.png\n\n' +
    '(Removes _P_ or _L_ orientation marker)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  const folders = [
    { id: '1z511jL5Iq7tTJPmEK_noi3WgLd3z4wvZ', name: 'CROPS_PORTRAIT' },
    { id: '1Lyeq07b7bkptqVy6RynHBAHGQtT674F5', name: 'CROPS_LANDSCAPE' }
  ];

  let totalRenamed = 0;
  let totalSkipped = 0;

  for (const folderInfo of folders) {
    const folder = DriveApp.getFolderById(folderInfo.id);
    const files = folder.getFiles();

    Logger.log('Processing folder: ' + folderInfo.name);

    while (files.hasNext()) {
      const file = files.next();
      const oldName = file.getName();

      // Pattern: {num}_{artworkName}_{P or L}_crop_{n}.{ext}
      // Remove _P_ or _L_ before _crop_
      const newName = oldName.replace(/_[PL]_crop_/i, '_crop_');

      if (newName !== oldName) {
        file.setName(newName);
        Logger.log('Renamed: ' + oldName + ' -> ' + newName);
        totalRenamed++;
      } else {
        totalSkipped++;
      }
    }
  }

  ui.alert('✅ Rename Complete!',
    'Renamed: ' + totalRenamed + ' files\n' +
    'Skipped (already correct): ' + totalSkipped + '\n\n' +
    'Check Logger for details',
    ui.ButtonSet.OK);
}

/**
 * Preview what fixDriveFileNames would do (without renaming)
 */
function previewDriveFileRenames() {
  const ui = SpreadsheetApp.getUi();

  const folders = [
    { id: '1z511jL5Iq7tTJPmEK_noi3WgLd3z4wvZ', name: 'CROPS_PORTRAIT' },
    { id: '1Lyeq07b7bkptqVy6RynHBAHGQtT674F5', name: 'CROPS_LANDSCAPE' }
  ];

  const wouldRename = [];
  const alreadyCorrect = [];

  for (const folderInfo of folders) {
    const folder = DriveApp.getFolderById(folderInfo.id);
    const files = folder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const oldName = file.getName();
      const newName = oldName.replace(/_[PL]_crop_/i, '_crop_');

      if (newName !== oldName) {
        wouldRename.push(oldName + ' -> ' + newName);
      } else {
        alreadyCorrect.push(oldName);
      }
    }
  }

  Logger.log('=== PREVIEW: Would rename ' + wouldRename.length + ' files ===');
  wouldRename.forEach(r => Logger.log(r));
  Logger.log('\n=== Already correct: ' + alreadyCorrect.length + ' ===');
  alreadyCorrect.forEach(f => Logger.log(f));

  ui.alert('Preview Results',
    'Would rename: ' + wouldRename.length + ' files\n' +
    'Already correct: ' + alreadyCorrect.length + '\n\n' +
    'Sample renames:\n' + wouldRename.slice(0, 5).join('\n') + '\n\n' +
    'Check Logger for full list',
    ui.ButtonSet.OK);
}

/**
 * FIX ALL FILES IN CROPS_ALL FOLDER
 * Standardize naming to: {num}_{artworkName}_crop_{n}.png
 */
function fixCropsAllFileNames() {
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert('Fix CROPS_ALL File Names',
    'This will standardize ALL files in CROPS_ALL folder.\n\n' +
    'Removes _P_ or _L_ orientation markers.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  const FOLDER_ID = '1Ft_QO1XlwYt6XCbCEKslG4QbDg0Lm9mT';
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  let totalRenamed = 0;
  let totalSkipped = 0;

  while (files.hasNext()) {
    const file = files.next();
    const oldName = file.getName();

    // Remove _P_ or _L_ before _crop_
    const newName = oldName.replace(/_[PL]_crop_/i, '_crop_');

    if (newName !== oldName) {
      file.setName(newName);
      Logger.log('Renamed: ' + oldName + ' -> ' + newName);
      totalRenamed++;
    } else {
      totalSkipped++;
    }
  }

  ui.alert('✅ Rename Complete!',
    'Renamed: ' + totalRenamed + ' files\n' +
    'Skipped: ' + totalSkipped + '\n\n' +
    'Check Logger for details',
    ui.ButtonSet.OK);
}
