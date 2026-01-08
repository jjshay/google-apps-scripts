/**
 * eBay Art Listing Automation Script
 * Automatically generates SKUs, Titles, and Descriptions for art listings
 * using prompts from the PROMPT tab and Claude AI integration
 */

// Configuration
const PRODUCTS_SHEET_NAME = 'PRODUCTS';
const PROMPT_SHEET_NAME = 'PROMPT';

// Column indices (0-based) for PRODUCTS sheet
const COL_SKU = 0;           // Column A
const COL_ARTIST = 1;        // Column B
const COL_TITLE = 4;         // Column E
const COL_DESCRIPTION = 5;   // Column F
const COL_TYPE = 10;         // Column K (Framed/Not Framed)

// Column indices (0-based) for PROMPT sheet
const PROMPT_COL_ARTIST = 0;      // Column A
const PROMPT_COL_CATEGORY = 1;    // Column B
const PROMPT_COL_CODE = 2;        // Column C
const PROMPT_COL_TITLE_PROMPT = 6; // Column G
const PROMPT_COL_DESC_PROMPT = 7;  // Column H

/**
 * Main function to process all products
 * Run this from: Extensions > Apps Script > Run > processAllProducts
 */
function processAllProducts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productsSheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);
  const promptSheet = ss.getSheetByName(PROMPT_SHEET_NAME);

  if (!productsSheet || !promptSheet) {
    SpreadsheetApp.getUi().alert('Error: Could not find PRODUCTS or PROMPT sheet');
    return;
  }

  // Get all data
  const productsData = productsSheet.getDataRange().getValues();
  const promptData = promptSheet.getDataRange().getValues();

  // Build lookup map from PROMPT sheet
  const promptMap = buildPromptMap(promptData);

  // Track SKU counters for each artist/category combination
  const skuCounters = {};

  // Process each row (skip header)
  for (let i = 1; i < productsData.length; i++) {
    const row = productsData[i];
    const artist = row[COL_ARTIST];
    const type = row[COL_TYPE]; // "Framed" or "Not Framed"
    const currentTitle = row[COL_TITLE];

    if (!artist) continue; // Skip empty rows

    // Determine category based on Type column
    const category = determineCategory(type);

    // Look up prompt information
    const promptInfo = promptMap[`${artist}|${category}`];

    if (!promptInfo) {
      Logger.log(`Warning: No prompt found for ${artist} - ${category} on row ${i + 1}`);
      continue;
    }

    // Generate SKU
    const sku = generateSKU(artist, category, currentTitle, promptInfo.code, skuCounters);

    // Update the sheet with new SKU
    productsSheet.getRange(i + 1, COL_SKU + 1).setValue(sku);

    Logger.log(`Processed row ${i + 1}: ${artist} - ${category} - SKU: ${sku}`);
  }

  SpreadsheetApp.getUi().alert('Processing complete! SKUs have been generated.\n\nNote: Titles and Descriptions require Claude AI integration.');
}

/**
 * Build a lookup map from PROMPT sheet data
 * Key format: "ARTIST|CATEGORY"
 */
function buildPromptMap(promptData) {
  const map = {};

  // Skip header row
  for (let i = 1; i < promptData.length; i++) {
    const row = promptData[i];
    const artist = row[PROMPT_COL_ARTIST];
    const category = row[PROMPT_COL_CATEGORY];
    const code = row[PROMPT_COL_CODE];
    const titlePrompt = row[PROMPT_COL_TITLE_PROMPT];
    const descPrompt = row[PROMPT_COL_DESC_PROMPT];

    if (!artist || !category) continue;

    // Normalize category (remove extra spaces)
    const normalizedCategory = category.trim().toUpperCase();

    const key = `${artist}|${normalizedCategory}`;
    map[key] = {
      code: code,
      titlePrompt: titlePrompt,
      descPrompt: descPrompt
    };
  }

  return map;
}

/**
 * Determine category based on Type field
 */
function determineCategory(type) {
  if (!type) return 'PRINT';

  const typeUpper = type.trim().toUpperCase();

  if (typeUpper.includes('FRAMED') || typeUpper === 'FRAMED') {
    return 'FRAMED';
  } else if (typeUpper.includes('NOT FRAMED') || typeUpper === 'NOT FRAMED') {
    return 'PRINT';
  } else if (typeUpper.includes('DOLLAR')) {
    return 'DOLLAR';
  } else if (typeUpper.includes('ORIGINAL')) {
    return 'ORIGINAL';
  } else if (typeUpper.includes('LETTERPRESS')) {
    return 'LETTERPRESS';
  }

  return 'PRINT'; // Default
}

/**
 * Generate SKU in format: #_CODE_TruncatedTitle
 * Must be under 50 characters total
 */
function generateSKU(artist, category, title, code, counters) {
  // Create counter key
  const counterKey = `${code}`;

  // Initialize or increment counter
  if (!counters[counterKey]) {
    counters[counterKey] = 1;
  } else {
    counters[counterKey]++;
  }

  const seqNumber = counters[counterKey];

  // Truncate title to fit within 50 character limit
  // Format: #_CODE_TruncatedTitle
  // Reserve space for: number (up to 3 digits) + underscore + code + underscore
  const reservedSpace = String(seqNumber).length + 1 + code.length + 1;
  const maxTitleLength = 50 - reservedSpace;

  // Clean and truncate title
  let cleanTitle = title
    .replace(/[^a-zA-Z0-9\s-]/g, '') // Remove special characters
    .replace(/\s+/g, '-')             // Replace spaces with hyphens
    .substring(0, maxTitleLength);

  // Remove trailing hyphen if present
  cleanTitle = cleanTitle.replace(/-+$/, '');

  return `${seqNumber}_${code}_${cleanTitle}`;
}

/**
 * Generate Titles using Claude AI
 * This requires Claude API integration
 */
function generateTitlesWithClaude() {
  SpreadsheetApp.getUi().alert(
    'Claude AI Integration Required\n\n' +
    'To generate titles, you need to:\n' +
    '1. Set up Claude API access\n' +
    '2. Add your API key to Script Properties\n' +
    '3. Run the generateTitlesWithClaudeAPI function\n\n' +
    'See the script comments for detailed instructions.'
  );
}

/**
 * Generate Descriptions using Claude AI
 * This requires Claude API integration
 */
function generateDescriptionsWithClaude() {
  SpreadsheetApp.getUi().alert(
    'Claude AI Integration Required\n\n' +
    'To generate descriptions, you need to:\n' +
    '1. Set up Claude API access\n' +
    '2. Add your API key to Script Properties\n' +
    '3. Run the generateDescriptionsWithClaudeAPI function\n\n' +
    'See the script comments for detailed instructions.'
  );
}

/**
 * ADVANCED: Generate titles with Claude API
 * Requires API key setup
 *
 * Setup Instructions:
 * 1. Get Claude API key from: https://console.anthropic.com/
 * 2. In Apps Script: File > Project properties > Script properties
 * 3. Add property: CLAUDE_API_KEY = your_api_key_here
 */
function generateTitlesWithClaudeAPI() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Please set CLAUDE_API_KEY in Script Properties');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productsSheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);
  const promptSheet = ss.getSheetByName(PROMPT_SHEET_NAME);

  const productsData = productsSheet.getDataRange().getValues();
  const promptData = promptSheet.getDataRange().getValues();
  const promptMap = buildPromptMap(promptData);

  // Process each row
  for (let i = 1; i < productsData.length; i++) {
    const row = productsData[i];
    const artist = row[COL_ARTIST];
    const type = row[COL_TYPE];
    const referenceImage = row[2]; // Column C - Reference Image

    if (!artist) continue;

    const category = determineCategory(type);
    const promptInfo = promptMap[`${artist}|${category}`];

    if (!promptInfo || !promptInfo.titlePrompt) continue;

    // Call Claude API to generate title
    try {
      const title = callClaudeAPI(apiKey, promptInfo.titlePrompt, referenceImage);
      productsSheet.getRange(i + 1, COL_TITLE + 1).setValue(title);
      Logger.log(`Generated title for row ${i + 1}`);

      // Rate limiting - wait 1 second between API calls
      Utilities.sleep(1000);

    } catch (error) {
      Logger.log(`Error generating title for row ${i + 1}: ${error}`);
    }
  }

  SpreadsheetApp.getUi().alert('Title generation complete!');
}

/**
 * ADVANCED: Generate descriptions with Claude API
 * Requires API key setup
 */
function generateDescriptionsWithClaudeAPI() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Please set CLAUDE_API_KEY in Script Properties');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productsSheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);
  const promptSheet = ss.getSheetByName(PROMPT_SHEET_NAME);

  const productsData = productsSheet.getDataRange().getValues();
  const promptData = promptSheet.getDataRange().getValues();
  const promptMap = buildPromptMap(promptData);

  // Process each row
  for (let i = 1; i < productsData.length; i++) {
    const row = productsData[i];
    const artist = row[COL_ARTIST];
    const type = row[COL_TYPE];
    const referenceImage = row[2]; // Column C - Reference Image

    if (!artist) continue;

    const category = determineCategory(type);
    const promptInfo = promptMap[`${artist}|${category}`];

    if (!promptInfo || !promptInfo.descPrompt) continue;

    // Call Claude API to generate description
    try {
      const description = callClaudeAPI(apiKey, promptInfo.descPrompt, referenceImage);
      productsSheet.getRange(i + 1, COL_DESCRIPTION + 1).setValue(description);
      Logger.log(`Generated description for row ${i + 1}`);

      // Rate limiting - wait 1 second between API calls
      Utilities.sleep(1000);

    } catch (error) {
      Logger.log(`Error generating description for row ${i + 1}: ${error}`);
    }
  }

  SpreadsheetApp.getUi().alert('Description generation complete!');
}

/**
 * Call Claude API with vision support
 */
function callClaudeAPI(apiKey, prompt, imageUrl) {
  const url = 'https://api.anthropic.com/v1/messages';

  // Build the message content
  let content = [];

  // Add image if provided
  if (imageUrl && imageUrl.trim() !== '') {
    // For Google Drive images, we need to fetch and encode them
    content.push({
      type: 'image',
      source: {
        type: 'url',
        url: imageUrl
      }
    });
  }

  // Add text prompt
  content.push({
    type: 'text',
    text: prompt
  });

  const payload = {
    model: 'claude-3-5-sonnet-20241022',
    max_tokens: 4096,
    messages: [{
      role: 'user',
      content: content
    }]
  };

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error(json.error.message);
  }

  return json.content[0].text;
}

/**
 * Create custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('eBay Automation')
    .addItem('Generate SKUs', 'processAllProducts')
    .addSeparator()
    .addItem('Generate Titles (Basic)', 'generateTitlesWithClaude')
    .addItem('Generate Descriptions (Basic)', 'generateDescriptionsWithClaude')
    .addSeparator()
    .addItem('Generate Titles with AI', 'generateTitlesWithClaudeAPI')
    .addItem('Generate Descriptions with AI', 'generateDescriptionsWithClaudeAPI')
    .addToUi();
}
