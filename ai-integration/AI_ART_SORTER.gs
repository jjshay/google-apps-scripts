/**
 * ============================================================================
 * AI-POWERED ART SORTER - CLAUDE VISION ARTIST IDENTIFICATION
 * ============================================================================
 *
 * This script analyzes art images using Claude Vision AI and automatically
 * sorts them into artist-specific subfolders based on visual style analysis.
 *
 * FEATURES:
 * 1. Identifies artist by visual style, technique, and signature elements
 * 2. Creates artist-specific subfolders automatically
 * 3. Confidence scoring for each identification
 * 4. Manual review queue for uncertain classifications
 * 5. Detailed logging of all decisions
 *
 * ============================================================================
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const ART_SORTER_CONFIG = {
  // Source folder containing mixed art images to sort
  SOURCE_FOLDER_ID: '15E4sMBfWACkv9EVUeRXj1j4uq6OoXbCu', // UPDATE THIS

  // Destination folder where artist subfolders will be created
  DESTINATION_FOLDER_ID: '15E4sMBfWACkv9EVUeRXj1j4uq6OoXbCu', // UPDATE THIS

  // Claude API - Get from Script Properties
  CLAUDE_API_KEY: PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY') || '',

  // Confidence threshold (0-100)
  // Images below this confidence will go to "NEEDS_REVIEW" folder
  CONFIDENCE_THRESHOLD: 70,

  // Processing limits
  MAX_IMAGES_PER_RUN: 10,
  MAX_IMAGE_SIZE: 4 * 1024 * 1024
};

// ============================================================================
// ARTIST DEFINITIONS - Visual Characteristics
// ============================================================================

const KNOWN_ARTISTS = {
  'Andy Warhol': {
    folder: 'Andy Warhol',
    characteristics: [
      'Pop art style with bold, vibrant colors',
      'Repeated images in grid format (like Campbell soup cans)',
      'Celebrity portraits with high contrast',
      'Screen printing aesthetic with flat colors',
      'Bright pinks, yellows, oranges, electric blues',
      'Commercial art influence'
    ],
    keywords: ['warhol', 'pop art', 'soup can', 'marilyn', 'screen print']
  },

  'Shepard Fairey': {
    folder: 'Shepard Fairey',
    characteristics: [
      'Bold graphic propaganda poster style',
      'Limited color palette: red, black, cream/beige',
      'Strong contrast and hard edges',
      'Political or social commentary themes',
      'OBEY branding or Andre the Giant imagery',
      'Street art aesthetic with stencil-like quality'
    ],
    keywords: ['fairey', 'obey', 'hope', 'propaganda', 'graphic']
  },

  'Keith Haring': {
    folder: 'Keith Haring',
    characteristics: [
      'Bold black outlines on bright backgrounds',
      'Stick figure-style characters in motion',
      'Radiant baby imagery',
      'Movement lines and energy radiating from figures',
      'Graffiti/subway art aesthetic',
      'Simple, iconic shapes and symbols',
      'Bright primary colors or neon backgrounds'
    ],
    keywords: ['haring', 'radiant baby', 'stick figure', 'movement lines']
  },

  'Mr. Brainwash': {
    folder: 'Mr Brainwash',
    characteristics: [
      'Pop art mashups with celebrity culture',
      'Spray paint and stencil aesthetic',
      'Warhol-influenced style with graffiti elements',
      'Mixed media collage look',
      'Bold text and graphics overlaid on images',
      'Street art meets commercial pop art'
    ],
    keywords: ['brainwash', 'mbw', 'thierry guetta', 'life is beautiful']
  },

  'Banksy': {
    folder: 'Banksy',
    characteristics: [
      'Stencil art with sharp, clean edges',
      'Often black and white with selective color',
      'Political or satirical commentary',
      'Street art/graffiti context',
      'Iconic imagery: girl with balloon, flower thrower, rats',
      'Dark humor or ironic messages'
    ],
    keywords: ['banksy', 'stencil', 'girl with balloon', 'rat', 'graffiti']
  },

  'Victor Vasarely': {
    folder: 'Victor Vasarely',
    characteristics: [
      'Op art (optical art) with geometric patterns',
      'Creates optical illusions and movement',
      'Precise geometric shapes: circles, squares, diamonds',
      'Black and white or vibrant color gradients',
      '3D illusion effects on flat surfaces',
      'Mathematical precision and symmetry'
    ],
    keywords: ['vasarely', 'op art', 'optical', 'geometric', 'zebra']
  },

  'Roy Lichtenstein': {
    folder: 'Roy Lichtenstein',
    characteristics: [
      'Comic book style with Ben-Day dots',
      'Bold black outlines',
      'Primary colors: red, yellow, blue',
      'Speech bubbles or thought bubbles',
      'Dramatic scenes or emotional faces',
      'Mechanical reproduction aesthetic',
      'Halftone dot patterns visible'
    ],
    keywords: ['lichtenstein', 'comic', 'ben-day dots', 'whaam', 'crying girl']
  },

  'Salvador Dali': {
    folder: 'Salvador Dali',
    characteristics: [
      'Surrealist imagery with dream-like quality',
      'Melting clocks or distorted objects',
      'Elephants with impossibly long legs',
      'Precise, detailed technical painting',
      'Desert or barren landscapes',
      'Bizarre juxtapositions of realistic objects',
      'Psychological or symbolic imagery'
    ],
    keywords: ['dali', 'surreal', 'melting clock', 'persistence of memory', 'elephant']
  },

  'Death NYC': {
    folder: 'Death NYC',
    characteristics: [
      'Pop culture mashups with urban aesthetic',
      'Designer brand logos mixed with pop icons',
      'Graffiti and street art influence',
      'Bold colors with contemporary subjects',
      'Cartoon characters mixed with luxury brands',
      'Urban contemporary pop art style'
    ],
    keywords: ['death nyc', 'dnyc', 'street art', 'pop culture']
  }
};

// ============================================================================
// CUSTOM MENU
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎨 Art Sorter')
    .addItem('🔍 Sort Images by Artist', 'sortArtByArtist')
    .addItem('📊 Test Single Image', 'testSingleImage')
    .addItem('📋 View Sort Results', 'viewSortResults')
    .addSeparator()
    .addItem('⚙️ Configure Folders', 'showConfigInstructions')
    .addToUi();
}

// ============================================================================
// MAIN SORTING FUNCTION
// ============================================================================

function sortArtByArtist() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'AI ART SORTER',
    `Analyze and sort images by artist?\n\n` +
    `• Uses Claude Vision AI\n` +
    `• Identifies artist by visual style\n` +
    `• Creates artist subfolders automatically\n` +
    `• Processes up to ${ART_SORTER_CONFIG.MAX_IMAGES_PER_RUN} images\n` +
    `• Confidence threshold: ${ART_SORTER_CONFIG.CONFIDENCE_THRESHOLD}%\n\n` +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const results = processArtImages();

  ui.alert(
    'Sorting Complete!',
    `Processed ${results.total} images\n\n` +
    `✅ Sorted: ${results.sorted}\n` +
    `⚠️ Needs Review: ${results.needsReview}\n` +
    `❌ Failed: ${results.failed}\n\n` +
    'Check Execution Log for details.',
    ui.ButtonSet.OK
  );
}

function processArtImages() {
  const startTime = new Date().getTime();
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('AI ART SORTER - CLAUDE VISION ANALYSIS');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  try {
    const sourceFolder = DriveApp.getFolderById(ART_SORTER_CONFIG.SOURCE_FOLDER_ID);
    const destFolder = DriveApp.getFolderById(ART_SORTER_CONFIG.DESTINATION_FOLDER_ID);

    // Get all image files from source folder
    const files = sourceFolder.getFiles();
    const imageFiles = [];

    while (files.hasNext() && imageFiles.length < ART_SORTER_CONFIG.MAX_IMAGES_PER_RUN) {
      const file = files.next();
      if (file.getMimeType().startsWith('image/')) {
        imageFiles.push(file);
      }
    }

    Logger.log(`📊 Found ${imageFiles.length} images to process\n`);

    let sorted = 0;
    let needsReview = 0;
    let failed = 0;

    for (let i = 0; i < imageFiles.length; i++) {
      const file = imageFiles[i];
      Logger.log(`\n[${i + 1}/${imageFiles.length}] 🖼️ ${file.getName()}`);

      const result = identifyAndSortArtwork(file, destFolder);

      if (result.success) {
        if (result.confidence >= ART_SORTER_CONFIG.CONFIDENCE_THRESHOLD) {
          sorted++;
          Logger.log(`  ✅ Sorted to: ${result.artist} (${result.confidence}% confidence)`);
        } else {
          needsReview++;
          Logger.log(`  ⚠️ Needs review: ${result.artist} (${result.confidence}% confidence)`);
        }
      } else {
        failed++;
        Logger.log(`  ❌ Failed: ${result.error}`);
      }

      Utilities.sleep(2000); // Rate limiting
    }

    Logger.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log(`✅ COMPLETE: ${sorted} sorted, ${needsReview} need review, ${failed} failed`);

    return { total: imageFiles.length, sorted, needsReview, failed };

  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    return { total: 0, sorted: 0, needsReview: 0, failed: 0 };
  }
}

// ============================================================================
// ARTIST IDENTIFICATION
// ============================================================================

function identifyAndSortArtwork(file, destFolder) {
  try {
    // Analyze image with Claude Vision
    const analysis = analyzeArtworkStyle(file);

    if (!analysis || !analysis.artist) {
      return { success: false, error: 'Could not identify artist' };
    }

    // Get or create artist folder
    const artistFolder = getOrCreateArtistFolder(destFolder, analysis.artist);

    // Determine if confidence is high enough
    const shouldSort = analysis.confidence >= ART_SORTER_CONFIG.CONFIDENCE_THRESHOLD;

    let targetFolder;
    if (shouldSort) {
      targetFolder = artistFolder;
    } else {
      // Low confidence - put in review folder
      targetFolder = getOrCreateReviewFolder(destFolder);
    }

    // Copy file to target folder
    file.makeCopy(file.getName(), targetFolder);

    return {
      success: true,
      artist: analysis.artist,
      confidence: analysis.confidence,
      reasoning: analysis.reasoning
    };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

function analyzeArtworkStyle(file) {
  try {
    const blob = getResizedImageForAnalysis(file);
    if (!blob) return null;

    const base64Image = Utilities.base64Encode(blob.getBytes());

    // Build comprehensive prompt with all artist characteristics
    let artistDescriptions = '';
    for (const [artistName, info] of Object.entries(KNOWN_ARTISTS)) {
      artistDescriptions += `\n${artistName}:\n`;
      info.characteristics.forEach(char => {
        artistDescriptions += `  • ${char}\n`;
      });
    }

    const prompt = `You are an art expert analyzing this image to identify the artist based on visual style.

KNOWN ARTISTS AND THEIR VISUAL CHARACTERISTICS:
${artistDescriptions}

ANALYSIS INSTRUCTIONS:
1. Carefully examine the visual style, technique, color palette, subject matter, and composition
2. Compare against the characteristics of each known artist
3. Identify the most likely artist
4. Provide a confidence score (0-100)
5. Explain your reasoning

If the artwork features a celebrity subject (like Taylor Swift, Ed Sheeran, Olivia Rodrigo), focus on the ARTISTIC STYLE, not the subject. For example, "Taylor Swift portrait in Andy Warhol's pop art style" should be categorized as Andy Warhol.

Return your analysis in this JSON format:
{
  "artist": "Artist Name",
  "confidence": 85,
  "reasoning": "Brief explanation of why you identified this artist",
  "subject": "What the artwork depicts",
  "styleElements": ["characteristic 1", "characteristic 2"]
}

If you cannot confidently identify the artist from the list, use "Unknown Artist" with a low confidence score.`;

    const payload = {
      model: 'claude-3-opus-20240229',
      max_tokens: 1000,
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: blob.getContentType(), data: base64Image } },
          { type: 'text', text: prompt }
        ]
      }]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': ART_SORTER_CONFIG.CLAUDE_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    const data = JSON.parse(response.getContentText());

    if (!data.content || !data.content[0]) {
      Logger.log('  ⚠️ No response from Claude');
      return null;
    }

    const analysisText = data.content[0].text;
    Logger.log(`  📝 Analysis: ${analysisText}`);

    const jsonMatch = analysisText.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      Logger.log('  ⚠️ Could not parse JSON response');
      return null;
    }

    const analysis = JSON.parse(jsonMatch[0]);

    // Validate artist is in our known list (or is "Unknown Artist")
    if (analysis.artist !== 'Unknown Artist' && !KNOWN_ARTISTS[analysis.artist]) {
      // Try to match artist name flexibly
      const matchedArtist = findClosestArtistMatch(analysis.artist);
      if (matchedArtist) {
        analysis.artist = matchedArtist;
      } else {
        analysis.artist = 'Unknown Artist';
        analysis.confidence = Math.min(analysis.confidence, 50);
      }
    }

    return analysis;

  } catch (error) {
    Logger.log(`  ❌ Claude Vision error: ${error.message}`);
    return null;
  }
}

function findClosestArtistMatch(artistName) {
  const lowerName = artistName.toLowerCase();

  for (const knownArtist of Object.keys(KNOWN_ARTISTS)) {
    const lowerKnown = knownArtist.toLowerCase();

    // Check if names are similar
    if (lowerKnown.includes(lowerName) || lowerName.includes(lowerKnown)) {
      return knownArtist;
    }

    // Check keywords
    const keywords = KNOWN_ARTISTS[knownArtist].keywords;
    for (const keyword of keywords) {
      if (lowerName.includes(keyword)) {
        return knownArtist;
      }
    }
  }

  return null;
}

function getResizedImageForAnalysis(file) {
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
          if (blob.getBytes().length < ART_SORTER_CONFIG.MAX_IMAGE_SIZE) {
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

// ============================================================================
// FOLDER MANAGEMENT
// ============================================================================

function getOrCreateArtistFolder(parentFolder, artistName) {
  const folderName = KNOWN_ARTISTS[artistName] ? KNOWN_ARTISTS[artistName].folder : artistName;

  const existingFolders = parentFolder.getFoldersByName(folderName);
  if (existingFolders.hasNext()) {
    return existingFolders.next();
  }

  return parentFolder.createFolder(folderName);
}

function getOrCreateReviewFolder(parentFolder) {
  const folderName = 'NEEDS_REVIEW';

  const existingFolders = parentFolder.getFoldersByName(folderName);
  if (existingFolders.hasNext()) {
    return existingFolders.next();
  }

  return parentFolder.createFolder(folderName);
}

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

function testSingleImage() {
  Logger.log('🧪 Testing with single image...');

  try {
    const sourceFolder = DriveApp.getFolderById(ART_SORTER_CONFIG.SOURCE_FOLDER_ID);
    const files = sourceFolder.getFiles();

    if (!files.hasNext()) {
      Logger.log('❌ No images found in source folder');
      return;
    }

    const testFile = files.next();
    Logger.log(`\n🖼️ Testing: ${testFile.getName()}`);

    const analysis = analyzeArtworkStyle(testFile);

    if (analysis) {
      Logger.log(`\n✅ RESULTS:`);
      Logger.log(`   Artist: ${analysis.artist}`);
      Logger.log(`   Confidence: ${analysis.confidence}%`);
      Logger.log(`   Subject: ${analysis.subject}`);
      Logger.log(`   Reasoning: ${analysis.reasoning}`);
      Logger.log(`   Style Elements: ${analysis.styleElements.join(', ')}`);
    } else {
      Logger.log('❌ Analysis failed');
    }

  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
  }
}

function showConfigInstructions() {
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    'Configuration',
    'To configure folder IDs:\n\n' +
    '1. Open Script Editor (Extensions > Apps Script)\n' +
    '2. Find AI_ART_SORTER.gs file\n' +
    '3. Update SOURCE_FOLDER_ID and DESTINATION_FOLDER_ID\n\n' +
    'To get folder IDs:\n' +
    '- Open folder in Google Drive\n' +
    '- Copy ID from URL after /folders/',
    ui.ButtonSet.OK
  );
}

function viewSortResults() {
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    'View Results',
    'Check the following:\n\n' +
    '1. Artist subfolders in destination folder\n' +
    '2. NEEDS_REVIEW folder for low-confidence matches\n' +
    '3. View > Execution Log for detailed analysis\n\n' +
    'Tip: Review images in NEEDS_REVIEW folder manually',
    ui.ButtonSet.OK
  );
}
