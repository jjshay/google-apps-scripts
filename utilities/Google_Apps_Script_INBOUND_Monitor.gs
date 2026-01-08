/**
 * Google Apps Script: INBOUND Folder Monitor
 *
 * This script runs on Google's servers and:
 * 1. Monitors the INBOUND folder for new images
 * 2. Analyzes each image with Claude Vision API
 * 3. Generates eBay listings
 * 4. Updates the Google Sheet automatically
 *
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1qRRIA7rbW_JhDg20xXV1WrZ-yx8ndnlwQrLkapr65Hc/
 * 2. Go to Extensions > Apps Script
 * 3. Copy and paste this entire script
 * 4. Click "Deploy" > "New Deployment" > "Web App"
 * 5. Set up a time-based trigger to run checkInboundFolder() every 5-60 minutes
 */

// ========================================
// CONFIGURATION
// ========================================

const CONFIG = {
  // INBOUND folder ID from Google Drive
  INBOUND_FOLDER_ID: '15E4sMBfWACkv9EVUeRXj1j4uq6OoXbCu',

  // Your Google Sheet ID
  SHEET_ID: '1qRRIA7rbW_JhDg20xXV1WrZ-yx8ndnlwQrLkapr65Hc',

  // Sheet tab name
  SHEET_TAB_NAME: 'Sheet1', // Change this to your actual tab name (e.g., "3D seller")

  // Claude API Key - Add your key here
  CLAUDE_API_KEY: 'YOUR_CLAUDE_API_KEY_HERE',

  // Starting row for data (row 2 if you have headers in row 1)
  START_ROW: 2,

  // Columns (C=3, D=4, E=5, etc.)
  COLUMNS: {
    IMAGE_LINK: 3,        // Column C
    ORIENTATION: 4,       // Column D
    TITLE: 5,            // Column E
    DESCRIPTION: 6,      // Column F
    TAGS: 7,             // Column G
    META_KEYWORDS: 8,    // Column H
    META_DESCRIPTION: 9, // Column I
    MOBILE_DESC: 10      // Column J
  }
};

// ========================================
// MAIN FUNCTIONS
// ========================================

/**
 * Main function to check INBOUND folder for new images
 * Set this to run every 5-60 minutes via Triggers
 */
function checkInboundFolder() {
  Logger.log('🔍 Checking INBOUND folder for new images...');

  try {
    const folder = DriveApp.getFolderById(CONFIG.INBOUND_FOLDER_ID);
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_TAB_NAME);

    if (!sheet) {
      Logger.log(`❌ Sheet "${CONFIG.SHEET_TAB_NAME}" not found!`);
      return;
    }

    // Get all processed image URLs from sheet
    const lastRow = sheet.getLastRow();
    const processedUrls = new Set();

    if (lastRow >= CONFIG.START_ROW) {
      const urlColumn = sheet.getRange(CONFIG.START_ROW, CONFIG.COLUMNS.IMAGE_LINK, lastRow - CONFIG.START_ROW + 1, 1).getValues();
      urlColumn.forEach(row => {
        if (row[0]) processedUrls.add(row[0]);
      });
    }

    Logger.log(`📊 Found ${processedUrls.size} already processed images`);

    // Get all images in INBOUND folder
    const files = folder.getFiles();
    let newCount = 0;

    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();

      // Check if it's an image
      if (!mimeType.startsWith('image/')) {
        continue;
      }

      const fileUrl = file.getUrl();

      // Skip if already processed
      if (processedUrls.has(fileUrl)) {
        continue;
      }

      Logger.log(`🆕 New image found: ${file.getName()}`);

      // Process the image
      processNewImage(file, sheet);
      newCount++;

      // Add small delay to avoid API rate limits
      Utilities.sleep(2000);
    }

    Logger.log(`✅ Processed ${newCount} new images`);

  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
  }
}

/**
 * Process a new image file
 */
function processNewImage(file, sheet) {
  try {
    Logger.log(`Processing: ${file.getName()}`);

    // Step 1: Get image URL
    const imageUrl = file.getUrl();
    const downloadUrl = `https://drive.google.com/uc?export=view&id=${file.getId()}`;

    // Step 2: Analyze image with Claude Vision
    const analysis = analyzeImageWithClaude(file);

    if (!analysis) {
      Logger.log(`⚠️ Failed to analyze ${file.getName()}`);
      return;
    }

    Logger.log(`✓ Analysis complete: ${analysis.orientation}`);

    // Step 3: Generate eBay listing
    const listing = generateEbayListing(analysis, file.getName());
    Logger.log(`✓ Listing generated: ${listing.title}`);

    // Step 4: Write to sheet
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, CONFIG.COLUMNS.IMAGE_LINK).setValue(downloadUrl);
    sheet.getRange(nextRow, CONFIG.COLUMNS.ORIENTATION).setValue(listing.orientation);
    sheet.getRange(nextRow, CONFIG.COLUMNS.TITLE).setValue(listing.title);
    sheet.getRange(nextRow, CONFIG.COLUMNS.DESCRIPTION).setValue(listing.description);
    sheet.getRange(nextRow, CONFIG.COLUMNS.TAGS).setValue(listing.tags);
    sheet.getRange(nextRow, CONFIG.COLUMNS.META_KEYWORDS).setValue(listing.metaKeywords);
    sheet.getRange(nextRow, CONFIG.COLUMNS.META_DESCRIPTION).setValue(listing.metaDescription);
    sheet.getRange(nextRow, CONFIG.COLUMNS.MOBILE_DESC).setValue(listing.mobileDescription);

    Logger.log(`✅ Added to sheet row ${nextRow}`);

  } catch (error) {
    Logger.log(`❌ Error processing ${file.getName()}: ${error.message}`);
  }
}

/**
 * Analyze image using Claude Vision API
 */
function analyzeImageWithClaude(file) {
  try {
    // Get image as blob and convert to base64
    const blob = file.getBlob();
    const base64Image = Utilities.base64Encode(blob.getBytes());

    const prompt = `Analyze this Death NYC artwork image and provide:

1. Orientation (portrait or landscape)
2. Main subject/theme
3. Color palette (3-5 dominant colors)
4. Art style (pop art, abstract, collage, etc.)
5. Key visual elements
6. Emotional tone

Respond in JSON format:
{
  "orientation": "portrait" or "landscape",
  "subject": "brief description",
  "colors": ["color1", "color2", "color3"],
  "style": "art style",
  "elements": ["element1", "element2"],
  "tone": "emotional description"
}`;

    const payload = {
      model: 'claude-3-5-sonnet-20241022',
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
          {
            type: 'text',
            text: prompt
          }
        ]
      }]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': CONFIG.CLAUDE_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log(`❌ Claude API error: ${responseCode} - ${response.getContentText()}`);
      return null;
    }

    const data = JSON.parse(response.getContentText());
    const analysisText = data.content[0].text;

    // Extract JSON from response (in case there's extra text)
    const jsonMatch = analysisText.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      Logger.log(`❌ Could not parse analysis JSON: ${analysisText}`);
      return null;
    }

    return JSON.parse(jsonMatch[0]);

  } catch (error) {
    Logger.log(`❌ Error calling Claude API: ${error.message}`);
    return null;
  }
}

/**
 * Generate eBay listing from analysis
 */
function generateEbayListing(analysis, fileName) {
  const subject = analysis.subject || 'Contemporary Art';
  const style = analysis.style || 'Pop Art';

  const title = `FINE ART ${subject} Death NYC Limited Edition Print`.substring(0, 80);

  const description = `THE ARTWORK:
${generateArtworkDescription(analysis)}

THE ARTIST:
Death NYC is a legendary figure in contemporary urban art, emerging from the underground street art movement to become one of the most collected artists in the pop culture mashup genre. Operating anonymously for years, Death NYC built a devoted following through bold appropriations of luxury fashion logos, iconic cartoon characters, and celebrity imagery.

What sets Death NYC apart is the fearless blend of high and low culture - Louis Vuitton meets Mickey Mouse, Chanel meets street graffiti. These works challenge notions of consumerism, brand worship, and pop culture iconography. Each piece is a commentary on our relationship with brands, nostalgia, and visual culture.

Death NYC's limited edition prints have appreciated significantly, with early works now commanding premium prices at auction. Collectors prize these pieces for their cultural relevance, visual impact, and the artist's growing reputation in contemporary art circles.

SUBJECT & STYLE:
${generateSubjectStyle(analysis)}

WHY THIS MATTERS:
• Authentic contemporary art with street art credibility and growing collector demand
• Museum-quality print reproduction capturing every detail of the original
• Conversation-starting imagery that bridges art history and pop culture
• Investment potential - works from these artists have shown consistent value appreciation
• Perfect scale and impact for residential, office, or gallery display
• Instantly recognizable aesthetic that signals sophisticated taste in contemporary art

CONDITION & AUTHENTICITY:
High-quality art print reproduction. The images shown are of the actual item - what you see is exactly what you'll receive. Please examine all photos carefully for condition assessment. This is a collectible print perfect for framing and display.

PROFESSIONAL SHIPPING:
We ship artwork with gallery-level care. Each print is protected with archival materials, reinforced corners, and shipped in a rigid container designed specifically for art transport. Your piece will arrive ready for framing. Combined shipping discounts available for multiple purchases.

COLLECTOR GUARANTEE:
30-day money-back guarantee, no questions asked. We're collectors ourselves and stand behind every piece. If you're not thrilled with your purchase, return it in original condition for a full refund.

Add this investment-quality contemporary art piece to your collection today. Questions? Message us - we typically respond within hours and love talking about art.`;

  const tags = `Death NYC, artwork, art, ${style.toLowerCase()}, wall decor, home decor, office art, gift`;

  const metaDescription = `${title} - ${analysis.subject}. Contemporary street art by Death NYC. ${analysis.tone} aesthetic with ${analysis.colors ? analysis.colors.join(', ') : 'vibrant'} palette.`;

  const mobileDescription = `FINE ART ${style} - ${analysis.subject}. Death NYC limited edition print. Museum-quality reproduction. Investment-grade contemporary art.`;

  return {
    title: title,
    description: description,
    tags: tags,
    orientation: analysis.orientation,
    metaKeywords: tags,
    metaDescription: metaDescription,
    mobileDescription: mobileDescription
  };
}

/**
 * Generate artwork description from analysis
 */
function generateArtworkDescription(analysis) {
  const colorPalette = analysis.colors && analysis.colors.length > 0
    ? `The color palette is striking with ${analysis.colors.join(', ')} creating a ${analysis.tone || 'dynamic'} visual impact.`
    : 'The color palette is striking with vibrant, eye-catching hues that demand attention.';

  const elements = analysis.elements && analysis.elements.length > 0
    ? `The composition features ${analysis.elements.join(', ')}. `
    : '';

  return `This artwork ${analysis.subject.toLowerCase()}. ${elements}${colorPalette}`;
}

/**
 * Generate subject & style section
 */
function generateSubjectStyle(analysis) {
  const styleDescriptions = {
    'pop art': 'The pop art influence is unmistakable - think Warhol, Lichtenstein, and Haring. Bold outlines, flat colors, and commercial imagery unite in a composition that feels both familiar and fresh. This is art that doesn\'t apologize for being accessible and visually striking.',
    'abstract': 'The layered, abstract composition rewards extended viewing. Multiple elements coexist in complex relationships, creating visual depth and conceptual richness.',
    'collage': 'The layered, collage-like composition rewards extended viewing. Multiple elements coexist in complex relationships, creating visual depth and conceptual richness. This technique reflects our fragmented media landscape where images, logos, and icons constantly overlap and compete for attention.',
    'graffiti': 'Rooted in street art tradition, this work carries the energy and immediacy of urban environments. The aesthetic recalls wheat-pasted posters, spray paint tags, and guerrilla art interventions - bringing that raw, authentic street credibility into a collectible format.'
  };

  const styleLower = analysis.style ? analysis.style.toLowerCase() : 'contemporary';

  return styleDescriptions[styleLower] ||
    `This ${analysis.style} piece demonstrates technical mastery and conceptual depth, creating a compelling visual experience that resonates with contemporary collectors.`;
}

// ========================================
// UTILITY FUNCTIONS
// ========================================

/**
 * Create a time-based trigger to check folder every hour
 * Run this once to set up automatic monitoring
 */
function createHourlyTrigger() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkInboundFolder') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger - runs every hour
  ScriptApp.newTrigger('checkInboundFolder')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('✅ Created hourly trigger for checkInboundFolder()');
}

/**
 * Manual test function - processes one random image
 */
function testProcessing() {
  Logger.log('🧪 Running test...');

  const folder = DriveApp.getFolderById(CONFIG.INBOUND_FOLDER_ID);
  const files = folder.getFiles();

  if (files.hasNext()) {
    const file = files.next();
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_TAB_NAME);

    Logger.log(`Testing with: ${file.getName()}`);
    processNewImage(file, sheet);
  } else {
    Logger.log('No images found in INBOUND folder');
  }
}

/**
 * List all files in INBOUND folder
 */
function listInboundFiles() {
  const folder = DriveApp.getFolderById(CONFIG.INBOUND_FOLDER_ID);
  const files = folder.getFiles();

  Logger.log('📁 Files in INBOUND folder:');
  let count = 0;

  while (files.hasNext()) {
    const file = files.next();
    count++;
    Logger.log(`${count}. ${file.getName()} (${file.getMimeType()})`);
  }

  Logger.log(`Total: ${count} files`);
}
