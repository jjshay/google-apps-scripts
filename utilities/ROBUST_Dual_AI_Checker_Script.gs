/**
 * ROBUST Google Apps Script: INBOUND Folder Monitor with Dual-AI Verification
 *
 * Features:
 * - Monitors INBOUND folder for new artwork images
 * - Claude Vision API for image analysis
 * - GPT-4 verification layer for quality assurance
 * - Automatic image resizing for large files
 * - Comprehensive error handling with retry logic
 * - Frame detection (Framed/Not Framed)
 * - Conditional length calculation for Death NYC artwork
 *
 * Setup:
 * 1. Go to: https://docs.google.com/spreadsheets/d/1qRRIA7rbW_JhDg20xXV1WrZ-yx8ndnlwQrLkapr65Hc/edit
 * 2. Click: Extensions > Apps Script
 * 3. Replace ALL code with this file
 * 4. Save and run "testProcessing"
 * 5. Grant permissions
 * 6. Run "createHourlyTrigger" for automatic monitoring
 */

// ========================================
// CONFIGURATION
// ========================================

const CONFIG = {
  INBOUND_FOLDER_ID: '',  // YOUR_FOLDER_ID_HERE - Get from Google Drive folder URL
  SHEET_ID: '',  // YOUR_SPREADSHEET_ID_HERE - Get from Google Sheets URL
  SHEET_TAB_NAME: 'PRODUCTS',
  // API Keys - Get from respective services:
  // Claude: https://console.anthropic.com
  // OpenAI: https://platform.openai.com
  CLAUDE_API_KEY: '',  // YOUR_CLAUDE_KEY_HERE
  OPENAI_API_KEY: '',  // YOUR_OPENAI_KEY_HERE
  START_ROW: 2,
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 2000,
  MAX_IMAGE_SIZE: 4 * 1024 * 1024, // 4MB to be safe

  COLUMNS: {
    ARTIST: 2,           // B
    IMAGE_LINK: 3,       // C
    ORIENTATION: 4,      // D
    TITLE: 5,            // E
    DESCRIPTION: 6,      // F
    TAGS: 7,             // G
    META_KEYWORDS: 8,    // H
    META_DESCRIPTION: 9, // I
    MOBILE_DESC: 10,     // J
    TYPE: 11,            // K - Framed or Not Framed
    LENGTH: 12           // L - Length (conditional)
  }
};

// ========================================
// MAIN FUNCTIONS
// ========================================

function checkInboundFolder() {
  Logger.log('🔍 Checking INBOUND folder for new images...');

  try {
    const folder = DriveApp.getFolderById(CONFIG.INBOUND_FOLDER_ID);
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_TAB_NAME);

    if (!sheet) {
      Logger.log(`❌ Sheet "${CONFIG.SHEET_TAB_NAME}" not found!`);
      return;
    }

    const lastRow = sheet.getLastRow();
    const processedUrls = new Set();

    if (lastRow >= CONFIG.START_ROW) {
      const urlColumn = sheet.getRange(CONFIG.START_ROW, CONFIG.COLUMNS.IMAGE_LINK, lastRow - CONFIG.START_ROW + 1, 1).getValues();
      urlColumn.forEach(row => {
        if (row[0]) processedUrls.add(row[0]);
      });
    }

    Logger.log(`📊 Found ${processedUrls.size} already processed images`);

    const files = folder.getFiles();
    let newCount = 0;

    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();

      if (!mimeType.startsWith('image/')) {
        continue;
      }

      const fileUrl = `https://drive.google.com/uc?export=view&id=${file.getId()}`;

      if (processedUrls.has(fileUrl)) {
        continue;
      }

      Logger.log(`🆕 New image found: ${file.getName()}`);
      processNewImage(file, sheet);
      newCount++;

      Utilities.sleep(2000);
    }

    Logger.log(`✅ Processed ${newCount} new images`);

  } catch (error) {
    Logger.log(`❌ Error in checkInboundFolder: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
  }
}

function processNewImage(file, sheet) {
  try {
    Logger.log(`Processing: ${file.getName()}`);

    const downloadUrl = `https://drive.google.com/uc?export=view&id=${file.getId()}`;

    // Step 1: Analyze image with Claude Vision (with retry logic)
    const analysis = analyzeImageWithClaude(file);

    if (!analysis) {
      Logger.log(`⚠️ Failed to analyze ${file.getName()} after all retries`);
      return;
    }

    Logger.log(`✓ Claude analysis complete: ${analysis.orientation}, ${analysis.framed}`);

    // Step 2: Generate eBay listing
    const listing = generateEbayListing(analysis, file.getName());
    Logger.log(`✓ Listing generated: ${listing.title}`);

    // Step 3: Verify and improve with GPT-4
    const verifiedListing = verifyWithGPT4(listing, analysis);
    Logger.log(`✓ GPT-4 verification complete`);

    // Step 4: Validate listing
    if (!validateListing(verifiedListing)) {
      Logger.log(`⚠️ Listing validation failed for ${file.getName()}`);
      return;
    }

    // Step 5: Write to sheet
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, CONFIG.COLUMNS.ARTIST).setValue('Death NYC');

    // Insert image as thumbnail using IMAGE formula
    sheet.getRange(nextRow, CONFIG.COLUMNS.IMAGE_LINK).setFormula(`=IMAGE("${downloadUrl}", 1)`);

    // Set row height for better thumbnail visibility (100 pixels)
    sheet.setRowHeight(nextRow, 100);

    sheet.getRange(nextRow, CONFIG.COLUMNS.ORIENTATION).setValue(verifiedListing.orientation);
    sheet.getRange(nextRow, CONFIG.COLUMNS.TITLE).setValue(verifiedListing.title);
    sheet.getRange(nextRow, CONFIG.COLUMNS.DESCRIPTION).setValue(verifiedListing.description);
    sheet.getRange(nextRow, CONFIG.COLUMNS.TAGS).setValue(verifiedListing.tags);
    sheet.getRange(nextRow, CONFIG.COLUMNS.META_KEYWORDS).setValue(verifiedListing.metaKeywords);
    sheet.getRange(nextRow, CONFIG.COLUMNS.META_DESCRIPTION).setValue(verifiedListing.metaDescription);
    sheet.getRange(nextRow, CONFIG.COLUMNS.MOBILE_DESC).setValue(verifiedListing.mobileDescription);
    sheet.getRange(nextRow, CONFIG.COLUMNS.TYPE).setValue(verifiedListing.type);
    sheet.getRange(nextRow, CONFIG.COLUMNS.LENGTH).setValue(verifiedListing.length);

    Logger.log(`✅ Added to sheet row ${nextRow}`);

  } catch (error) {
    Logger.log(`❌ Error processing ${file.getName()}: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
  }
}

// ========================================
// AI ANALYSIS FUNCTIONS
// ========================================

function analyzeImageWithClaude(file) {
  for (let attempt = 1; attempt <= CONFIG.MAX_RETRIES; attempt++) {
    try {
      Logger.log(`📏 Attempt ${attempt}/${CONFIG.MAX_RETRIES} - Original file size: ${(file.getSize() / 1024 / 1024).toFixed(2)} MB`);

      const blob = getResizedImage(file);

      if (!blob) {
        Logger.log(`❌ Could not resize image on attempt ${attempt}`);
        if (attempt < CONFIG.MAX_RETRIES) {
          Utilities.sleep(CONFIG.RETRY_DELAY_MS * attempt);
          continue;
        }
        return null;
      }

      const base64Image = Utilities.base64Encode(blob.getBytes());
      const sizeInMB = (blob.getBytes().length / 1024 / 1024).toFixed(2);
      Logger.log(`✓ Resized to: ${sizeInMB} MB`);

      const prompt = `Analyze this Death NYC artwork image and provide:

1. Orientation (portrait or landscape)
2. Main subject/theme
3. Color palette (3-5 dominant colors)
4. Art style (pop art, abstract, collage, etc.)
5. Key visual elements
6. Emotional tone
7. Framing status - Look carefully at the edges of the image. If there is a visible physical frame border around the artwork (wood, metal, or any frame material), respond "Framed". If the artwork extends to the edges without a frame border, respond "Not Framed".

Respond in JSON format:
{
  "orientation": "portrait" or "landscape",
  "subject": "brief description",
  "colors": ["color1", "color2", "color3"],
  "style": "art style",
  "elements": ["element1", "element2"],
  "tone": "emotional description",
  "framed": "Framed" or "Not Framed"
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
        Logger.log(`❌ Claude API error (attempt ${attempt}): ${responseCode} - ${response.getContentText()}`);
        if (attempt < CONFIG.MAX_RETRIES) {
          Utilities.sleep(CONFIG.RETRY_DELAY_MS * attempt);
          continue;
        }
        return null;
      }

      const data = JSON.parse(response.getContentText());
      const analysisText = data.content[0].text;

      const jsonMatch = analysisText.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        Logger.log(`❌ Could not parse analysis JSON (attempt ${attempt}): ${analysisText}`);
        if (attempt < CONFIG.MAX_RETRIES) {
          Utilities.sleep(CONFIG.RETRY_DELAY_MS * attempt);
          continue;
        }
        return null;
      }

      return JSON.parse(jsonMatch[0]);

    } catch (error) {
      Logger.log(`❌ Error calling Claude API (attempt ${attempt}): ${error.message}`);
      if (attempt < CONFIG.MAX_RETRIES) {
        Utilities.sleep(CONFIG.RETRY_DELAY_MS * attempt);
        continue;
      }
      return null;
    }
  }

  return null;
}

function getResizedImage(file) {
  try {
    const fileId = file.getId();
    const thumbnailSizes = [1600, 1200, 800];

    for (const size of thumbnailSizes) {
      try {
        const thumbnailUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w${size}`;

        const response = UrlFetchApp.fetch(thumbnailUrl, {
          headers: {
            'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
          },
          muteHttpExceptions: true
        });

        if (response.getResponseCode() === 200) {
          const blob = response.getBlob();
          const blobSize = blob.getBytes().length;

          Logger.log(`📐 Thumbnail size ${size}px: ${(blobSize / 1024 / 1024).toFixed(2)} MB`);

          if (blobSize < CONFIG.MAX_IMAGE_SIZE) {
            return blob;
          }
        }
      } catch (e) {
        Logger.log(`⚠️ Failed to get thumbnail at size ${size}: ${e.message}`);
      }
    }

    Logger.log(`⚠️ Using original image blob as fallback`);
    return file.getBlob();

  } catch (error) {
    Logger.log(`❌ Error resizing image: ${error.message}`);
    return null;
  }
}

function verifyWithGPT4(listing, analysis) {
  for (let attempt = 1; attempt <= CONFIG.MAX_RETRIES; attempt++) {
    try {
      Logger.log(`🤖 GPT-4 verification attempt ${attempt}/${CONFIG.MAX_RETRIES}`);

      // Build the prompt with properly escaped strings
      const promptData = {
        analysis: analysis,
        listing: {
          title: listing.title,
          description: listing.description,
          tags: listing.tags,
          metaDescription: listing.metaDescription,
          mobileDescription: listing.mobileDescription,
          orientation: listing.orientation,
          type: listing.type,
          length: listing.length
        },
        requirements: [
          "Title MUST start with 'FINE ART'",
          "Title maximum 80 characters",
          "Description should be compelling and SEO-optimized",
          "Verify accuracy based on the analysis",
          "Improve clarity and appeal if needed"
        ]
      };

      const prompt = `You are an eBay listing quality assurance expert. Review and improve this artwork listing if needed.

Here is the data in JSON format:
${JSON.stringify(promptData, null, 2)}

Please respond with an improved listing in the following JSON format:
{
  "title": "improved title starting with FINE ART",
  "description": "improved description",
  "tags": "improved tags",
  "metaDescription": "improved meta description",
  "mobileDescription": "improved mobile description",
  "orientation": "portrait or landscape",
  "type": "Framed or Not Framed",
  "length": "length value"
}

Return ONLY the JSON object, no other text.`;

      const payload = {
        model: 'gpt-4',
        messages: [{
          role: 'user',
          content: prompt
        }],
        temperature: 0.3,
        max_tokens: 2500
      };

      const options = {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'Authorization': 'Bearer ' + CONFIG.OPENAI_API_KEY
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
      const responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        Logger.log(`❌ GPT-4 API error (attempt ${attempt}): ${responseCode} - ${response.getContentText()}`);
        if (attempt < CONFIG.MAX_RETRIES) {
          Utilities.sleep(CONFIG.RETRY_DELAY_MS * attempt);
          continue;
        }
        // If GPT-4 fails, return original listing
        return listing;
      }

      const data = JSON.parse(response.getContentText());
      const gptResponse = data.choices[0].message.content;

      const jsonMatch = gptResponse.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        const improvedListing = JSON.parse(jsonMatch[0]);

        // Preserve original fields that shouldn't change
        improvedListing.metaKeywords = improvedListing.tags || listing.tags;

        Logger.log(`✓ GPT-4 verification successful`);
        return improvedListing;
      } else {
        Logger.log(`⚠️ Could not parse GPT-4 response, using original listing`);
        return listing;
      }

    } catch (error) {
      Logger.log(`❌ Error calling GPT-4 API (attempt ${attempt}): ${error.message}`);
      if (attempt < CONFIG.MAX_RETRIES) {
        Utilities.sleep(CONFIG.RETRY_DELAY_MS * attempt);
        continue;
      }
      return listing;
    }
  }

  return listing;
}

// ========================================
// LISTING GENERATION FUNCTIONS
// ========================================

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
30-day money-back guarantee, no questions asked. We're collectors themselves and stand behind every piece. If you're not thrilled with your purchase, return it in original condition for a full refund.

Add this investment-quality contemporary art piece to your collection today. Questions? Message us - we typically respond within hours and love talking about art.`;

  const tags = `Death NYC, artwork, art, ${style.toLowerCase()}, wall decor, home decor, office art, gift`;

  const metaDescription = `${title} - ${analysis.subject}. Contemporary street art by Death NYC. ${analysis.tone} aesthetic with ${analysis.colors ? analysis.colors.join(', ') : 'vibrant'} palette.`;

  const mobileDescription = `FINE ART ${style} - ${analysis.subject}. Death NYC limited edition print. Museum-quality reproduction. Investment-grade contemporary art.`;

  const type = analysis.framed || 'Not Framed';
  const length = calculateLength('Death NYC', analysis.orientation, type);

  return {
    title: title,
    description: description,
    tags: tags,
    orientation: analysis.orientation,
    metaKeywords: tags,
    metaDescription: metaDescription,
    mobileDescription: mobileDescription,
    type: type,
    length: length
  };
}

function calculateLength(artist, orientation, type) {
  // Death NYC specific sizing rules
  if (artist === 'Death NYC') {
    // Rule 1: Landscape + Framed = 16"
    if (orientation === 'landscape' && type === 'Framed') {
      return '16"';
    }
    // Rule 2: Not Framed = 13" (regardless of orientation)
    if (type === 'Not Framed') {
      return '13"';
    }
  }

  // Default if conditions not met
  return '';
}

function generateArtworkDescription(analysis) {
  const colorPalette = analysis.colors && analysis.colors.length > 0
    ? `The color palette is striking with ${analysis.colors.join(', ')} creating a ${analysis.tone || 'dynamic'} visual impact.`
    : 'The color palette is striking with vibrant, eye-catching hues that demand attention.';

  const elements = analysis.elements && analysis.elements.length > 0
    ? `The composition features ${analysis.elements.join(', ')}. `
    : '';

  return `This artwork ${analysis.subject.toLowerCase()}. ${elements}${colorPalette}`;
}

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
// VALIDATION FUNCTIONS
// ========================================

function validateListing(listing) {
  try {
    // Check title starts with "FINE ART"
    if (!listing.title.startsWith('FINE ART')) {
      Logger.log(`❌ Validation failed: Title does not start with "FINE ART"`);
      return false;
    }

    // Check title length
    if (listing.title.length > 80) {
      Logger.log(`❌ Validation failed: Title exceeds 80 characters (${listing.title.length})`);
      return false;
    }

    // Check required fields
    const requiredFields = ['title', 'description', 'tags', 'orientation', 'type'];
    for (const field of requiredFields) {
      if (!listing[field] || listing[field].trim() === '') {
        Logger.log(`❌ Validation failed: Missing required field "${field}"`);
        return false;
      }
    }

    // Check orientation value
    if (listing.orientation !== 'portrait' && listing.orientation !== 'landscape') {
      Logger.log(`❌ Validation failed: Invalid orientation "${listing.orientation}"`);
      return false;
    }

    // Check type value
    if (listing.type !== 'Framed' && listing.type !== 'Not Framed') {
      Logger.log(`❌ Validation failed: Invalid type "${listing.type}"`);
      return false;
    }

    Logger.log(`✓ Validation passed`);
    return true;

  } catch (error) {
    Logger.log(`❌ Validation error: ${error.message}`);
    return false;
  }
}

// ========================================
// UTILITY FUNCTIONS
// ========================================

function createHourlyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkInboundFolder') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('checkInboundFolder')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('✅ Created hourly trigger for checkInboundFolder()');
}

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
    Logger.log('❌ No images found in INBOUND folder');
  }
}

function listInboundFiles() {
  const folder = DriveApp.getFolderById(CONFIG.INBOUND_FOLDER_ID);
  const files = folder.getFiles();

  Logger.log('📁 Files in INBOUND folder:');
  let count = 0;

  while (files.hasNext()) {
    const file = files.next();
    count++;
    Logger.log(`${count}. ${file.getName()} (${file.getMimeType()}) - ${(file.getSize() / 1024 / 1024).toFixed(2)} MB`);
  }

  Logger.log(`Total: ${count} files`);
}

/**
 * Backfill Artist column (B) for all existing rows
 * If Title (E), Description (F), or Tags (G) contain "Death NYC", set Artist to "Death NYC"
 * RUN THIS ONCE to update existing data
 */
function backfillArtistColumn() {
  Logger.log('🔧 Starting artist column backfill...');

  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_TAB_NAME);

    if (!sheet) {
      Logger.log(`❌ Sheet "${CONFIG.SHEET_TAB_NAME}" not found!`);
      return;
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < CONFIG.START_ROW) {
      Logger.log('No data rows found to process');
      return;
    }

    Logger.log(`Processing rows ${CONFIG.START_ROW} to ${lastRow}...`);

    let updatedCount = 0;

    // Process each row
    for (let row = CONFIG.START_ROW; row <= lastRow; row++) {
      try {
        // Get current artist value
        const currentArtist = sheet.getRange(row, CONFIG.COLUMNS.ARTIST).getValue();

        // Skip if artist already populated
        if (currentArtist && currentArtist.trim() !== '') {
          continue;
        }

        // Get title, description, and tags
        const title = sheet.getRange(row, CONFIG.COLUMNS.TITLE).getValue() || '';
        const description = sheet.getRange(row, CONFIG.COLUMNS.DESCRIPTION).getValue() || '';
        const tags = sheet.getRange(row, CONFIG.COLUMNS.TAGS).getValue() || '';

        // Check if any field contains "Death NYC"
        const titleHasArtist = title.toLowerCase().includes('death nyc');
        const descriptionHasArtist = description.toLowerCase().includes('death nyc');
        const tagsHasArtist = tags.toLowerCase().includes('death nyc');

        if (titleHasArtist || descriptionHasArtist || tagsHasArtist) {
          sheet.getRange(row, CONFIG.COLUMNS.ARTIST).setValue('Death NYC');
          updatedCount++;
          Logger.log(`✓ Row ${row}: Set artist to "Death NYC"`);
        }

      } catch (error) {
        Logger.log(`⚠️ Error processing row ${row}: ${error.message}`);
      }
    }

    Logger.log(`✅ Backfill complete! Updated ${updatedCount} rows`);

  } catch (error) {
    Logger.log(`❌ Error in backfillArtistColumn: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
  }
}

/**
 * Convert existing image URLs in Column C to thumbnails
 * RUN THIS ONCE to update existing rows
 */
function convertUrlsToThumbnails() {
  Logger.log('🖼️ Converting URLs to thumbnails...');

  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_TAB_NAME);

    if (!sheet) {
      Logger.log(`❌ Sheet "${CONFIG.SHEET_TAB_NAME}" not found!`);
      return;
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < CONFIG.START_ROW) {
      Logger.log('No data rows found to process');
      return;
    }

    Logger.log(`Processing rows ${CONFIG.START_ROW} to ${lastRow}...`);

    let convertedCount = 0;

    // Process each row
    for (let row = CONFIG.START_ROW; row <= lastRow; row++) {
      try {
        const cell = sheet.getRange(row, CONFIG.COLUMNS.IMAGE_LINK);
        const currentValue = cell.getValue();
        const currentFormula = cell.getFormula();

        // Skip if already has IMAGE formula
        if (currentFormula && currentFormula.includes('IMAGE')) {
          continue;
        }

        // Skip if empty
        if (!currentValue) {
          continue;
        }

        // Check if it's a valid URL
        const urlString = currentValue.toString();
        if (urlString.startsWith('http')) {
          // Convert to IMAGE formula
          cell.setFormula(`=IMAGE("${urlString}", 1)`);

          // Set row height for thumbnail
          sheet.setRowHeight(row, 100);

          convertedCount++;
          Logger.log(`✓ Row ${row}: Converted URL to thumbnail`);
        }

      } catch (error) {
        Logger.log(`⚠️ Error processing row ${row}: ${error.message}`);
      }
    }

    Logger.log(`✅ Conversion complete! Updated ${convertedCount} rows`);

  } catch (error) {
    Logger.log(`❌ Error in convertUrlsToThumbnails: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
  }
}
