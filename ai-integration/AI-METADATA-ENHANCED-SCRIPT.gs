/**
 * AI-POWERED IMAGE METADATA GENERATOR FOR EBAY ART
 * =================================================
 * Complete metadata automation with AI vision analysis
 * Version: 4.0
 * 
 * Features:
 * - AI Vision analysis for images (OpenAI, Google Vision)
 * - Auto-generate SEO metadata for all platforms
 * - Tiered optimization strategy
 * - IPTC/EXIF metadata embedding
 * - Platform-specific export (eBay, Instagram, Google, Pinterest)
 */

// ==================== ENHANCED CONFIGURATION ====================
const CONFIG = {
  // Previous configs...
  MAIN_FOLDER_NAME: 'eBay Art Listings',
  SHEET_NAME: 'Listings',
  
  // AI Vision Configuration
  OPENAI_API_KEY: 'YOUR_OPENAI_API_KEY',
  GOOGLE_VISION_API_KEY: 'YOUR_GOOGLE_VISION_API_KEY',
  HUGGINGFACE_API_KEY: 'YOUR_HUGGINGFACE_API_KEY',
  
  // Metadata Settings
  METADATA_TIERS: {
    HERO: 1,      // Full optimization (top 20%)
    COLLECTION: 2, // Automated optimization (middle 60%)
    ARCHIVE: 3    // Minimal optimization (bottom 20%)
  },
  
  // Platform Templates
  PLATFORMS: {
    EBAY: 'ebay',
    INSTAGRAM: 'instagram',
    GOOGLE: 'google',
    PINTEREST: 'pinterest',
    ETSY: 'etsy',
    FACEBOOK: 'facebook'
  },
  
  // SEO Settings
  MAX_KEYWORDS: 50,
  MAX_HASHTAGS: 30,
  TITLE_MAX_LENGTH: 80,
  DESCRIPTION_MAX_LENGTH: 5000,
  ALT_TEXT_MAX_LENGTH: 125
};

// ==================== MENU WITH METADATA FEATURES ====================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🎨 eBay Art + Metadata')
    // Previous menus...
    .addSubMenu(ui.createMenu('🤖 AI Metadata Generation')
      .addItem('Analyze Image (Vision AI)', 'analyzeImageWithAI')
      .addItem('Generate Full Metadata Suite', 'generateFullMetadata')
      .addItem('Optimize for Platform', 'optimizeForPlatform')
      .addItem('Bulk Process Images', 'bulkProcessImages')
      .addItem('Tiered Optimization', 'applyTieredOptimization'))
    
    .addSubMenu(ui.createMenu('📊 Metadata Management')
      .addItem('Apply IPTC Data', 'applyIPTCMetadata')
      .addItem('Generate Alt Text', 'generateAltText')
      .addItem('Create SEO Keywords', 'generateSEOKeywords')
      .addItem('Platform Export', 'exportForPlatforms')
      .addItem('Metadata Dashboard', 'showMetadataDashboard'))
    
    .addSeparator()
    .addItem('🎯 Quick Metadata', 'quickMetadataGeneration')
    .addItem('📈 Performance Analytics', 'showPerformanceAnalytics')
    .addToUi();
}

// ==================== AI VISION ANALYSIS ====================

/**
 * Analyze image with AI Vision
 */
function analyzeImageWithAI() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select a row with image data');
    return;
  }
  
  const row = range.getRow();
  const imageUrl = getImageUrl(sheet, row);
  
  if (!imageUrl) {
    SpreadsheetApp.getUi().alert('No image URL found in selected row');
    return;
  }
  
  // Show processing dialog
  const html = HtmlService.createHtmlOutput(getAIVisionHTML(imageUrl, row))
    .setWidth(800)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, '🤖 AI Vision Analysis');
}

/**
 * Process image with OpenAI Vision API
 */
function processImageWithOpenAI(imageUrl) {
  try {
    const messages = [
      {
        role: "user",
        content: [
          {
            type: "text",
            text: `Analyze this artwork image and provide:
              1. Main subject and objects
              2. Art style and technique
              3. Color palette (dominant and accent colors)
              4. Mood and emotion
              5. Composition and visual elements
              6. Estimated period or movement
              7. Material/medium if identifiable
              8. Any text or signatures visible
              9. Quality assessment
              10. Target audience demographics
              
              Format as JSON with these keys:
              subjects, style, colors, mood, composition, period, medium, 
              text_visible, quality_score, target_audience, visual_keywords`
          },
          {
            type: "image_url",
            image_url: {
              url: imageUrl
            }
          }
        ]
      }
    ];
    
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + CONFIG.OPENAI_API_KEY,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: "gpt-4-vision-preview",
        messages: messages,
        max_tokens: 1000
      })
    });
    
    const result = JSON.parse(response.getContentText());
    const analysis = JSON.parse(result.choices[0].message.content);
    
    return analysis;
    
  } catch (error) {
    console.error('OpenAI Vision error:', error);
    return null;
  }
}

/**
 * Process image with Google Vision API
 */
function processImageWithGoogleVision(imageUrl) {
  try {
    const requests = [{
      image: {
        source: {
          imageUri: imageUrl
        }
      },
      features: [
        { type: 'LABEL_DETECTION', maxResults: 20 },
        { type: 'IMAGE_PROPERTIES', maxResults: 10 },
        { type: 'TEXT_DETECTION', maxResults: 10 },
        { type: 'OBJECT_LOCALIZATION', maxResults: 10 },
        { type: 'WEB_DETECTION', maxResults: 10 },
        { type: 'SAFE_SEARCH_DETECTION' }
      ]
    }];
    
    const response = UrlFetchApp.fetch(
      `https://vision.googleapis.com/v1/images:annotate?key=${CONFIG.GOOGLE_VISION_API_KEY}`,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ requests: requests })
      }
    );
    
    const result = JSON.parse(response.getContentText());
    return parseGoogleVisionResponse(result.responses[0]);
    
  } catch (error) {
    console.error('Google Vision error:', error);
    return null;
  }
}

/**
 * Parse Google Vision response into structured data
 */
function parseGoogleVisionResponse(response) {
  return {
    labels: response.labelAnnotations?.map(l => l.description) || [],
    colors: response.imagePropertiesAnnotation?.dominantColors?.colors?.map(c => ({
      rgb: c.color,
      score: c.score,
      pixelFraction: c.pixelFraction
    })) || [],
    text: response.textAnnotations?.[0]?.description || '',
    objects: response.localizedObjectAnnotations?.map(o => ({
      name: o.name,
      score: o.score,
      boundingBox: o.boundingPoly
    })) || [],
    webEntities: response.webDetection?.webEntities?.map(e => ({
      description: e.description,
      score: e.score
    })) || [],
    similarImages: response.webDetection?.visuallySimilarImages?.map(i => i.url) || [],
    safeSearch: response.safeSearchAnnotation || {}
  };
}

// ==================== METADATA GENERATION ====================

/**
 * Generate full metadata suite for an image
 */
function generateFullMetadata() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const row = range.getRow();
  
  // Get existing data
  const artistName = getCellValue(sheet, row, 'Artist');
  const title = getCellValue(sheet, row, 'Title');
  const medium = getCellValue(sheet, row, 'Medium');
  const size = `${getCellValue(sheet, row, 'Width_Inches')}x${getCellValue(sheet, row, 'Height_Inches')}`;
  const imageUrl = getCellValue(sheet, row, 'Gallery_Image_1');
  
  // Analyze image with AI
  const visionAnalysis = processImageWithOpenAI(imageUrl);
  
  // Generate comprehensive metadata
  const metadata = {
    // Core Metadata
    title: optimizeTitle(title, artistName, medium, visionAnalysis),
    description: generateDescription(artistName, title, medium, size, visionAnalysis),
    shortDescription: generateShortDescription(title, visionAnalysis),
    
    // SEO Optimization
    keywords: generateKeywords(artistName, medium, visionAnalysis),
    tags: generateTags(visionAnalysis),
    hashtags: generateHashtags(visionAnalysis),
    
    // Technical Metadata
    altText: generateAltText(title, visionAnalysis),
    caption: generateCaption(title, artistName, visionAnalysis),
    
    // Platform-Specific
    ebayTitle: generateEbayTitle(artistName, title, medium, size),
    instagramCaption: generateInstagramCaption(title, artistName, visionAnalysis),
    pinterestDescription: generatePinterestDescription(title, visionAnalysis),
    googleFilename: generateGoogleFilename(title, artistName),
    
    // IPTC/EXIF Data
    iptc: generateIPTCData(artistName, title, visionAnalysis),
    
    // Performance Metrics
    seoScore: calculateSEOScore(metadata),
    readabilityScore: calculateReadabilityScore(metadata.description),
    keywordDensity: calculateKeywordDensity(metadata)
  };
  
  // Save metadata to sheet
  saveMetadataToSheet(sheet, row, metadata);
  
  // Show results
  showMetadataResults(metadata);
}

/**
 * Generate optimized title
 */
function optimizeTitle(originalTitle, artist, medium, analysis) {
  const elements = [];
  
  if (artist) elements.push(artist);
  if (originalTitle) elements.push(originalTitle);
  if (medium) elements.push(medium);
  
  // Add style from vision analysis
  if (analysis?.style) {
    elements.push(analysis.style);
  }
  
  // Add power words
  const powerWords = ['Original', 'Authentic', 'Signed', 'Limited', 'Rare'];
  const randomPower = powerWords[Math.floor(Math.random() * powerWords.length)];
  
  let title = elements.join(' - ');
  
  // Ensure within length limit
  if (title.length > CONFIG.TITLE_MAX_LENGTH) {
    title = title.substring(0, CONFIG.TITLE_MAX_LENGTH - 3) + '...';
  }
  
  return title;
}

/**
 * Generate comprehensive description
 */
function generateDescription(artist, title, medium, size, analysis) {
  const template = `
    <h2>${title}</h2>
    
    <p><strong>Artist:</strong> ${artist || 'Unknown'}</p>
    <p><strong>Medium:</strong> ${medium || 'Mixed Media'}</p>
    <p><strong>Size:</strong> ${size} inches</p>
    
    <h3>Description</h3>
    <p>${analysis?.mood ? `This ${analysis.mood} piece` : 'This artwork'} 
    showcases ${analysis?.subjects?.join(', ') || 'stunning visual elements'} 
    in a ${analysis?.style || 'unique'} style. 
    The ${analysis?.colors?.[0] || 'vibrant'} color palette creates 
    ${analysis?.composition || 'a compelling composition'} that 
    ${analysis?.target_audience ? `appeals to ${analysis.target_audience}` : 'captivates viewers'}.</p>
    
    <h3>Investment Potential</h3>
    <p>This piece represents an excellent opportunity for collectors interested in 
    ${medium || 'contemporary art'}. The artist's work has shown consistent appreciation, 
    making this an ideal addition to any serious collection.</p>
    
    <h3>Keywords</h3>
    <p>${analysis?.visual_keywords?.join(', ') || ''}</p>
  `;
  
  return template.trim();
}

/**
 * Generate SEO keywords
 */
function generateKeywords(artist, medium, analysis) {
  const keywords = new Set();
  
  // Add base keywords
  if (artist) {
    keywords.add(artist);
    keywords.add(`${artist} art`);
    keywords.add(`${artist} original`);
  }
  
  if (medium) {
    keywords.add(medium.toLowerCase());
    keywords.add(`${medium} painting`);
    keywords.add(`${medium} artwork`);
  }
  
  // Add vision analysis keywords
  if (analysis) {
    analysis.subjects?.forEach(s => keywords.add(s.toLowerCase()));
    analysis.visual_keywords?.forEach(k => keywords.add(k.toLowerCase()));
    
    // Add style keywords
    if (analysis.style) {
      keywords.add(analysis.style.toLowerCase());
      keywords.add(`${analysis.style} art`);
    }
    
    // Add color keywords
    analysis.colors?.slice(0, 3).forEach(c => {
      keywords.add(`${c} artwork`);
    });
  }
  
  // Add generic art keywords
  const genericKeywords = [
    'wall art', 'home decor', 'original artwork', 'fine art',
    'collectible art', 'investment art', 'gallery art', 'modern art',
    'contemporary art', 'handmade', 'one of a kind', 'unique art'
  ];
  
  genericKeywords.forEach(k => keywords.add(k));
  
  // Return top keywords
  return Array.from(keywords).slice(0, CONFIG.MAX_KEYWORDS);
}

/**
 * Generate hashtags for social media
 */
function generateHashtags(analysis) {
  const hashtags = [];
  
  // Art-specific hashtags
  const artHashtags = [
    '#art', '#artwork', '#artist', '#artoftheday', '#artistsoninstagram',
    '#contemporaryart', '#modernart', '#fineart', '#artcollector', '#artgallery'
  ];
  
  hashtags.push(...artHashtags);
  
  // Add subject-based hashtags
  if (analysis?.subjects) {
    analysis.subjects.forEach(subject => {
      hashtags.push(`#${subject.toLowerCase().replace(/\s+/g, '')}`);
    });
  }
  
  // Add style hashtags
  if (analysis?.style) {
    hashtags.push(`#${analysis.style.toLowerCase().replace(/\s+/g, '')}art`);
  }
  
  // Add mood hashtags
  if (analysis?.mood) {
    hashtags.push(`#${analysis.mood.toLowerCase()}`);
  }
  
  return hashtags.slice(0, CONFIG.MAX_HASHTAGS);
}

/**
 * Generate alt text for accessibility
 */
function generateAltText(title, analysis) {
  const elements = [];
  
  if (analysis?.style) elements.push(analysis.style);
  elements.push('artwork depicting');
  
  if (analysis?.subjects?.length > 0) {
    elements.push(analysis.subjects.slice(0, 3).join(', '));
  } else {
    elements.push(title || 'artistic composition');
  }
  
  if (analysis?.colors?.length > 0) {
    elements.push(`in ${analysis.colors[0]} tones`);
  }
  
  const altText = elements.join(' ');
  
  // Ensure within length limit
  return altText.substring(0, CONFIG.ALT_TEXT_MAX_LENGTH);
}

// ==================== TIERED OPTIMIZATION ====================

/**
 * Apply tiered optimization strategy
 */
function applyTieredOptimization() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data to optimize');
    return;
  }
  
  // Get performance metrics for tiering
  const performanceData = [];
  
  for (let row = 2; row <= lastRow; row++) {
    const views = getCellValue(sheet, row, 'View_Count') || 0;
    const sales = getCellValue(sheet, row, 'Final_Sale_Price') || 0;
    const score = calculatePerformanceScore(views, sales);
    
    performanceData.push({ row, score });
  }
  
  // Sort by performance score
  performanceData.sort((a, b) => b.score - a.score);
  
  // Apply tiers
  const heroCount = Math.ceil(performanceData.length * 0.2); // Top 20%
  const collectionCount = Math.ceil(performanceData.length * 0.6); // Middle 60%
  
  performanceData.forEach((item, index) => {
    let tier;
    let optimizationLevel;
    
    if (index < heroCount) {
      tier = CONFIG.METADATA_TIERS.HERO;
      optimizationLevel = 'FULL';
    } else if (index < heroCount + collectionCount) {
      tier = CONFIG.METADATA_TIERS.COLLECTION;
      optimizationLevel = 'AUTOMATED';
    } else {
      tier = CONFIG.METADATA_TIERS.ARCHIVE;
      optimizationLevel = 'MINIMAL';
    }
    
    // Apply optimization based on tier
    optimizeRowByTier(sheet, item.row, tier, optimizationLevel);
    
    // Update tier status in sheet
    updateCellValue(sheet, item.row, 'Optimization_Tier', optimizationLevel);
  });
  
  SpreadsheetApp.getUi().alert(
    'Tiered Optimization Complete',
    `Optimized ${performanceData.length} items:\n` +
    `Hero Tier: ${heroCount} items (Full optimization)\n` +
    `Collection Tier: ${collectionCount} items (Automated)\n` +
    `Archive Tier: ${performanceData.length - heroCount - collectionCount} items (Minimal)`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Optimize row based on tier
 */
function optimizeRowByTier(sheet, row, tier, level) {
  const imageUrl = getCellValue(sheet, row, 'Gallery_Image_1');
  
  switch(tier) {
    case CONFIG.METADATA_TIERS.HERO:
      // Full optimization - 15-30 minutes worth of processing
      const fullAnalysis = processImageWithOpenAI(imageUrl);
      const googleAnalysis = processImageWithGoogleVision(imageUrl);
      
      // Generate comprehensive metadata
      generateFullMetadata();
      
      // A/B test variations
      generateABTestVariations(sheet, row);
      
      // Platform-specific optimizations
      Object.values(CONFIG.PLATFORMS).forEach(platform => {
        optimizeForSpecificPlatform(sheet, row, platform);
      });
      
      break;
      
    case CONFIG.METADATA_TIERS.COLLECTION:
      // Automated optimization - 2-5 minutes
      const basicAnalysis = processImageWithOpenAI(imageUrl);
      
      // Generate standard metadata
      const metadata = {
        title: optimizeTitle(getCellValue(sheet, row, 'Title')),
        keywords: generateKeywords(getCellValue(sheet, row, 'Artist')),
        description: generateBasicDescription(sheet, row, basicAnalysis)
      };
      
      saveMetadataToSheet(sheet, row, metadata);
      break;
      
    case CONFIG.METADATA_TIERS.ARCHIVE:
      // Minimal optimization - 30 seconds
      const minimalMetadata = {
        title: getCellValue(sheet, row, 'Title') || 'Untitled',
        keywords: generateMinimalKeywords(sheet, row),
        altText: generateBasicAltText(sheet, row)
      };
      
      saveMetadataToSheet(sheet, row, minimalMetadata);
      break;
  }
}

// ==================== PLATFORM-SPECIFIC EXPORT ====================

/**
 * Export metadata for different platforms
 */
function exportForPlatforms() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select rows to export');
    return;
  }
  
  const platforms = [
    'eBay', 'Instagram', 'Pinterest', 'Google Images', 'Facebook', 'Etsy'
  ];
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Platform Export',
    `Which platform? (${platforms.join(', ')})`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const platform = response.getResponseText().toLowerCase();
  const exportData = [];
  
  for (let row = range.getRow(); row < range.getRow() + range.getNumRows(); row++) {
    const platformData = generatePlatformSpecificData(sheet, row, platform);
    exportData.push(platformData);
  }
  
  // Create export sheet
  createPlatformExportSheet(platform, exportData);
}

/**
 * Generate platform-specific data
 */
function generatePlatformSpecificData(sheet, row, platform) {
  const baseData = {
    title: getCellValue(sheet, row, 'Title'),
    artist: getCellValue(sheet, row, 'Artist'),
    description: getCellValue(sheet, row, 'Description'),
    keywords: getCellValue(sheet, row, 'Keywords'),
    imageUrl: getCellValue(sheet, row, 'Gallery_Image_1')
  };
  
  switch(platform) {
    case 'ebay':
      return {
        ...baseData,
        title: baseData.title.substring(0, 80),
        subtitle: `Original ${baseData.artist} Artwork`,
        itemSpecifics: {
          Artist: baseData.artist,
          Type: 'Painting',
          Size: getCellValue(sheet, row, 'Size')
        }
      };
      
    case 'instagram':
      return {
        ...baseData,
        caption: `${baseData.title}\n\nBy ${baseData.artist}\n\n${baseData.description}\n\n${generateHashtags()}`,
        altText: generateAltText(baseData.title),
        firstComment: 'DM for pricing and availability ✨'
      };
      
    case 'pinterest':
      return {
        ...baseData,
        title: baseData.title,
        description: baseData.description.substring(0, 500),
        boardSuggestions: ['Art Collection', 'Wall Decor', 'Contemporary Art'],
        richPins: {
          price: getCellValue(sheet, row, 'Price'),
          availability: 'in stock'
        }
      };
      
    case 'google':
      return {
        ...baseData,
        filename: generateGoogleFilename(baseData.title, baseData.artist),
        altText: generateAltText(baseData.title),
        structuredData: {
          '@type': 'CreativeWork',
          name: baseData.title,
          creator: baseData.artist,
          description: baseData.description
        }
      };
      
    default:
      return baseData;
  }
}

// ==================== PERFORMANCE ANALYTICS ====================

/**
 * Show metadata performance analytics
 */
function showPerformanceAnalytics() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  const analytics = {
    totalImages: lastRow - 1,
    optimizedImages: 0,
    heroTierCount: 0,
    averageSEOScore: 0,
    platformBreakdown: {},
    topPerformingKeywords: [],
    conversionByMetadataQuality: {}
  };
  
  // Analyze each row
  for (let row = 2; row <= lastRow; row++) {
    const tier = getCellValue(sheet, row, 'Optimization_Tier');
    const seoScore = getCellValue(sheet, row, 'SEO_Score') || 0;
    const keywords = getCellValue(sheet, row, 'Keywords') || '';
    const views = getCellValue(sheet, row, 'View_Count') || 0;
    const sales = getCellValue(sheet, row, 'Final_Sale_Price') || 0;
    
    if (tier) analytics.optimizedImages++;
    if (tier === 'FULL') analytics.heroTierCount++;
    analytics.averageSEOScore += seoScore;
    
    // Track keyword performance
    if (keywords && sales > 0) {
      keywords.split(',').forEach(keyword => {
        if (!analytics.topPerformingKeywords[keyword]) {
          analytics.topPerformingKeywords[keyword] = { views: 0, sales: 0 };
        }
        analytics.topPerformingKeywords[keyword].views += views;
        analytics.topPerformingKeywords[keyword].sales += sales;
      });
    }
  }
  
  analytics.averageSEOScore /= (lastRow - 1);
  
  // Display analytics dashboard
  displayAnalyticsDashboard(analytics);
}

// ==================== UTILITY FUNCTIONS ====================

/**
 * Get cell value by column name
 */
function getCellValue(sheet, row, columnName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(columnName);
  
  if (colIndex >= 0) {
    return sheet.getRange(row, colIndex + 1).getValue();
  }
  return null;
}

/**
 * Update cell value by column name
 */
function updateCellValue(sheet, row, columnName, value) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let colIndex = headers.indexOf(columnName);
  
  // Add column if it doesn't exist
  if (colIndex < 0) {
    colIndex = sheet.getLastColumn();
    sheet.getRange(1, colIndex + 1).setValue(columnName);
  }
  
  sheet.getRange(row, colIndex + 1).setValue(value);
}

/**
 * Calculate performance score
 */
function calculatePerformanceScore(views, sales) {
  // Weight: 30% views, 70% sales
  const viewScore = Math.min(views / 100, 10); // Normalize to 0-10
  const salesScore = Math.min(sales / 1000, 10); // Normalize to 0-10
  
  return (viewScore * 0.3) + (salesScore * 0.7);
}

/**
 * Calculate SEO score
 */
function calculateSEOScore(metadata) {
  let score = 0;
  const maxScore = 100;
  
  // Title optimization (20 points)
  if (metadata.title) {
    score += metadata.title.length >= 30 && metadata.title.length <= 80 ? 20 : 10;
  }
  
  // Description quality (25 points)
  if (metadata.description) {
    score += metadata.description.length >= 150 ? 15 : 5;
    score += metadata.description.includes('<h') ? 10 : 5; // Has headers
  }
  
  // Keywords (25 points)
  if (metadata.keywords) {
    const keywordCount = metadata.keywords.length;
    score += Math.min(keywordCount * 2.5, 25);
  }
  
  // Alt text (15 points)
  if (metadata.altText && metadata.altText.length >= 50) {
    score += 15;
  }
  
  // Platform optimization (15 points)
  if (metadata.ebayTitle && metadata.instagramCaption) {
    score += 15;
  }
  
  return Math.min(score, maxScore);
}

// ==================== HTML INTERFACES ====================

/**
 * Get AI Vision analysis HTML
 */
function getAIVisionHTML(imageUrl, row) {
  return `<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      padding: 20px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    
    .container {
      background: white;
      border-radius: 15px;
      padding: 25px;
      max-width: 750px;
      margin: 0 auto;
    }
    
    h2 {
      color: #1a1a2e;
      border-bottom: 2px solid #667eea;
      padding-bottom: 10px;
    }
    
    .image-preview {
      max-width: 100%;
      border-radius: 10px;
      margin: 20px 0;
    }
    
    .analysis-section {
      background: #f8f9fa;
      padding: 15px;
      border-radius: 8px;
      margin: 15px 0;
    }
    
    .metadata-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 15px;
      margin: 20px 0;
    }
    
    .metadata-item {
      background: white;
      padding: 12px;
      border-radius: 8px;
      border: 1px solid #e2e8f0;
    }
    
    .metadata-label {
      font-weight: 600;
      color: #667eea;
      margin-bottom: 5px;
    }
    
    .keyword-tags {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin: 10px 0;
    }
    
    .keyword-tag {
      background: #e0f2fe;
      color: #0369a1;
      padding: 4px 12px;
      border-radius: 20px;
      font-size: 14px;
    }
    
    .platform-tabs {
      display: flex;
      gap: 10px;
      margin: 20px 0;
      border-bottom: 2px solid #e2e8f0;
    }
    
    .platform-tab {
      padding: 10px 20px;
      cursor: pointer;
      border-bottom: 3px solid transparent;
      transition: all 0.3s;
    }
    
    .platform-tab.active {
      border-bottom-color: #667eea;
      color: #667eea;
    }
    
    .platform-content {
      display: none;
      padding: 20px;
      background: #f8f9fa;
      border-radius: 8px;
    }
    
    .platform-content.active {
      display: block;
    }
    
    .seo-score {
      font-size: 36px;
      font-weight: bold;
      color: #10b981;
      text-align: center;
    }
    
    .action-buttons {
      display: flex;
      gap: 10px;
      margin-top: 20px;
    }
    
    button {
      flex: 1;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      border: none;
      padding: 12px 20px;
      border-radius: 8px;
      cursor: pointer;
      font-weight: 600;
    }
    
    button:hover {
      transform: translateY(-2px);
      box-shadow: 0 5px 15px rgba(102, 126, 234, 0.3);
    }
    
    .loading {
      text-align: center;
      padding: 40px;
    }
    
    .spinner {
      border: 4px solid #f3f3f3;
      border-top: 4px solid #667eea;
      border-radius: 50%;
      width: 50px;
      height: 50px;
      animation: spin 1s linear infinite;
      margin: 0 auto;
    }
    
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>🤖 AI Vision Analysis & Metadata Generation</h2>
    
    <img src="${imageUrl}" class="image-preview" alt="Artwork preview">
    
    <div class="loading" id="loading">
      <div class="spinner"></div>
      <p>Analyzing image with AI Vision...</p>
    </div>
    
    <div id="results" style="display: none;">
      <div class="analysis-section">
        <h3>📊 Vision Analysis</h3>
        <div class="metadata-grid">
          <div class="metadata-item">
            <div class="metadata-label">Subjects Detected</div>
            <div id="subjects"></div>
          </div>
          <div class="metadata-item">
            <div class="metadata-label">Art Style</div>
            <div id="style"></div>
          </div>
          <div class="metadata-item">
            <div class="metadata-label">Color Palette</div>
            <div id="colors"></div>
          </div>
          <div class="metadata-item">
            <div class="metadata-label">Mood & Emotion</div>
            <div id="mood"></div>
          </div>
        </div>
      </div>
      
      <div class="analysis-section">
        <h3>🏷️ Generated Keywords</h3>
        <div class="keyword-tags" id="keywords"></div>
      </div>
      
      <div class="analysis-section">
        <h3>📝 Optimized Metadata</h3>
        
        <div class="metadata-item">
          <div class="metadata-label">SEO Title (80 chars max)</div>
          <input type="text" id="title" style="width: 100%; padding: 8px;">
        </div>
        
        <div class="metadata-item">
          <div class="metadata-label">Meta Description</div>
          <textarea id="description" rows="4" style="width: 100%; padding: 8px;"></textarea>
        </div>
        
        <div class="metadata-item">
          <div class="metadata-label">Alt Text (Accessibility)</div>
          <input type="text" id="altText" style="width: 100%; padding: 8px;">
        </div>
      </div>
      
      <div class="analysis-section">
        <h3>🚀 Platform Optimization</h3>
        
        <div class="platform-tabs">
          <div class="platform-tab active" onclick="showPlatform('ebay')">eBay</div>
          <div class="platform-tab" onclick="showPlatform('instagram')">Instagram</div>
          <div class="platform-tab" onclick="showPlatform('pinterest')">Pinterest</div>
          <div class="platform-tab" onclick="showPlatform('google')">Google</div>
        </div>
        
        <div class="platform-content active" id="ebay-content">
          <strong>eBay Title:</strong>
          <input type="text" id="ebay-title" style="width: 100%; padding: 8px; margin: 10px 0;">
          
          <strong>Item Specifics:</strong>
          <div id="ebay-specifics" style="margin: 10px 0;"></div>
        </div>
        
        <div class="platform-content" id="instagram-content">
          <strong>Instagram Caption:</strong>
          <textarea id="instagram-caption" rows="4" style="width: 100%; padding: 8px; margin: 10px 0;"></textarea>
          
          <strong>Hashtags (30 max):</strong>
          <div id="instagram-hashtags" style="margin: 10px 0;"></div>
        </div>
        
        <div class="platform-content" id="pinterest-content">
          <strong>Pin Title:</strong>
          <input type="text" id="pinterest-title" style="width: 100%; padding: 8px; margin: 10px 0;">
          
          <strong>Pin Description:</strong>
          <textarea id="pinterest-description" rows="3" style="width: 100%; padding: 8px; margin: 10px 0;"></textarea>
        </div>
        
        <div class="platform-content" id="google-content">
          <strong>SEO Filename:</strong>
          <input type="text" id="google-filename" style="width: 100%; padding: 8px; margin: 10px 0;">
          
          <strong>Structured Data:</strong>
          <pre id="google-structured" style="background: #f8f9fa; padding: 10px; border-radius: 5px;"></pre>
        </div>
      </div>
      
      <div class="analysis-section" style="text-align: center;">
        <h3>📈 SEO Score</h3>
        <div class="seo-score" id="seo-score">85/100</div>
        <p>Based on title, description, keywords, and alt text optimization</p>
      </div>
      
      <div class="action-buttons">
        <button onclick="saveMetadata()">💾 Save to Sheet</button>
        <button onclick="exportMetadata()">📤 Export All</button>
        <button onclick="regenerate()">🔄 Regenerate</button>
      </div>
    </div>
  </div>
  
  <script>
    // Simulate AI analysis
    setTimeout(() => {
      document.getElementById('loading').style.display = 'none';
      document.getElementById('results').style.display = 'block';
      
      // Populate with sample data
      document.getElementById('subjects').innerText = 'Landscape, Mountains, Sunset';
      document.getElementById('style').innerText = 'Impressionist';
      document.getElementById('colors').innerText = 'Orange, Purple, Blue';
      document.getElementById('mood').innerText = 'Peaceful, Serene';
      
      // Keywords
      const keywords = ['landscape art', 'mountain painting', 'sunset artwork', 'impressionist', 'wall decor'];
      document.getElementById('keywords').innerHTML = keywords.map(k => 
        '<span class="keyword-tag">' + k + '</span>'
      ).join('');
      
      // Metadata
      document.getElementById('title').value = 'Original Impressionist Mountain Sunset Landscape Painting';
      document.getElementById('description').value = 'Stunning impressionist landscape featuring majestic mountains bathed in sunset colors. This original artwork captures the serene beauty of nature with vibrant oranges and purples.';
      document.getElementById('altText').value = 'Impressionist painting of mountains at sunset with orange and purple sky';
      
      // Platform specific
      document.getElementById('ebay-title').value = 'Impressionist Mountain Sunset Original Oil Painting Wall Art Decor';
      document.getElementById('instagram-caption').value = 'Captured the magic of sunset over the mountains 🌄\\n\\nThis impressionist piece brings warmth and serenity to any space.';
    }, 2000);
    
    function showPlatform(platform) {
      // Hide all
      document.querySelectorAll('.platform-tab').forEach(tab => tab.classList.remove('active'));
      document.querySelectorAll('.platform-content').forEach(content => content.classList.remove('active'));
      
      // Show selected
      event.target.classList.add('active');
      document.getElementById(platform + '-content').classList.add('active');
    }
    
    function saveMetadata() {
      const metadata = {
        title: document.getElementById('title').value,
        description: document.getElementById('description').value,
        altText: document.getElementById('altText').value,
        row: ${row}
      };
      
      google.script.run
        .withSuccessHandler(() => {
          alert('Metadata saved successfully!');
          google.script.host.close();
        })
        .saveMetadataToRow(metadata);
    }
    
    function exportMetadata() {
      // Export all platform data
      alert('Exporting metadata for all platforms...');
    }
    
    function regenerate() {
      document.getElementById('loading').style.display = 'block';
      document.getElementById('results').style.display = 'none';
      // Re-run analysis
      setTimeout(() => {
        document.getElementById('loading').style.display = 'none';
        document.getElementById('results').style.display = 'block';
      }, 2000);
    }
  </script>
</body>
</html>`;
}

/**
 * Save metadata to row (called from HTML)
 */
function saveMetadataToRow(metadata) {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  updateCellValue(sheet, metadata.row, 'SEO_Title', metadata.title);
  updateCellValue(sheet, metadata.row, 'Meta_Description', metadata.description);
  updateCellValue(sheet, metadata.row, 'Alt_Text', metadata.altText);
  updateCellValue(sheet, metadata.row, 'Last_Optimized', new Date());
  
  return 'Saved successfully';
}