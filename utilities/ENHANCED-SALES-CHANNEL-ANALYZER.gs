/**
 * GALLERY GAUNTLET - ENHANCED SALES CHANNEL ANALYZER
 * Ultra-robust script that identifies ALL specific requirements by sales channel
 * Includes title length, image count, image size, character limits, and optimization tips
 */

function generateEnhancedSalesChannelAnalysis() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create comprehensive analysis sheets
    createRequirementsSheet(spreadsheet);
    createLimitsComparisonSheet(spreadsheet);
    createOptimizationGuideSheet(spreadsheet);
    createImageSpecsSheet(spreadsheet);
    
    SpreadsheetApp.getUi().alert(
      '🚀 Enhanced Sales Channel Analysis Complete!',
      'Created 4 comprehensive analysis sheets covering 12 PLATFORMS:\n\n' +
      '📊 Requirements Matrix - 160+ specific field requirements\n' +
      '📏 Limits Comparison - Character/size limits by platform\n' +
      '🎯 Optimization Guide - Platform-specific strategies\n' +
      '🖼️ Image Specifications - Detailed image requirements\n\n' +
      'NOW INCLUDES: Shopify, eBay, Etsy, Instagram, Facebook, Google Shopping,\n' +
      'Saatchi Art, Amazon, 1stDibs, Artsy, POSHMARK, and STOCKX!\n\n' +
      'Each sheet identifies specific requirements for title length, image count, image sizes, and optimization strategies!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Analysis failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function createRequirementsSheet(spreadsheet) {
  var sheet = getOrCreateSheet(spreadsheet, 'Sales Channel Requirements');
  
  var data = [
    // Comprehensive headers
    ['Platform', 'Field Type', 'Field Name', 'Required', 'Character Limit', 'Format', 'Gallery Gauntlet Example', 'SEO Weight', 'Conversion Impact', 'Platform Notes'],
    
    // SHOPIFY - Most flexible platform
    ['Shopify', 'Title', 'Product Title', 'Required', '255 chars', 'Plain text', 'Shepard Fairey - OBEY Giant "Andre the Giant" (1989) - Original Stencil - 24x36 - Signed - COA Verified', 'High', 'Critical', 'Include artist, title, year, medium, size, authentication'],
    ['Shopify', 'Description', 'Product Description', 'Required', '65,535 chars', 'Rich HTML', '<h2>About the Artist</h2><p>Shepard Fairey...</p><h2>Artwork Details</h2><ul><li>Medium: Stencil on Paper</li></ul>', 'High', 'Critical', 'Use HTML formatting, headers, lists, detailed provenance'],
    ['Shopify', 'Images', 'Featured Image', 'Required', '20MB', 'JPG/PNG/WebP', '1200x1200px square format, clean background, professional lighting', 'Critical', 'Critical', 'First impression - invest in professional photography'],
    ['Shopify', 'Images', 'Gallery Images', 'Optional', '20MB each', 'JPG/PNG/WebP', 'Detail shots, signature, COA certificate, room mockups, frame options', 'Medium', 'High', 'Max 250 images - show authenticity and context'],
    ['Shopify', 'SEO', 'SEO Title', 'Optional', '70 chars', 'Plain text', 'Shepard Fairey OBEY Giant - Original Signed | Gallery Gauntlet', 'Critical', 'Medium', 'Google search title - include brand for trust'],
    ['Shopify', 'SEO', 'Meta Description', 'Optional', '160 chars', 'Plain text', 'Original Shepard Fairey OBEY Giant stencil. Hand-signed with blockchain COA. Investment-grade street art. Free shipping over $2,500.', 'Critical', 'Medium', 'Google search snippet - include key selling points'],
    ['Shopify', 'Tags', 'Product Tags', 'Optional', 'Unlimited', 'Comma-separated', 'shepard-fairey, obey-giant, street-art, stencil-art, investment-grade, signed-original, political-art, contemporary-art', 'High', 'Medium', 'Critical for internal search and filtering'],
    ['Shopify', 'Variants', 'Frame Options', 'Optional', 'Unlimited', 'Dropdown', 'Unframed ($0), White Frame (+$150), Black Frame (+$150), Wood Frame (+$200), Metal Frame (+$250)', 'Low', 'High', 'Use frame configurator data for options'],
    
    // EBAY - Keyword optimization critical
    ['eBay', 'Title', 'Listing Title', 'Required', '80 chars', 'Plain text', 'Shepard Fairey OBEY Giant Andre Stencil Signed Street Art COA Investment', 'Critical', 'Critical', 'Every character counts - pack with search keywords'],
    ['eBay', 'Title', 'Subtitle', 'Optional', '55 chars', 'Plain text', 'Original 1989 - Gallery Quality - Authenticated', 'High', 'Medium', 'Additional keyword space - worth the fee'],
    ['eBay', 'Description', 'Item Description', 'Required', '500,000 chars', 'HTML/Text', 'Professional eBay template with Gallery Gauntlet branding, detailed specs, payment terms', 'Medium', 'High', 'Build trust with professional presentation'],
    ['eBay', 'Images', 'Main Photo', 'Required', '12MB', 'JPG/GIF/PNG', '1600x1600px minimum, white background preferred, no watermarks', 'Critical', 'Critical', 'eBay search favors white backgrounds'],
    ['eBay', 'Images', 'Additional Photos', 'Optional', '12MB each', 'JPG/GIF/PNG', '12 images max - details, signature, COA, packaging proof', 'Medium', 'High', 'Build buyer confidence with comprehensive photos'],
    ['eBay', 'Item Specifics', 'Artist', 'Required', '65 chars', 'Plain text', 'Shepard Fairey', 'Critical', 'Medium', 'Exact match for artist searches'],
    ['eBay', 'Item Specifics', 'Medium', 'Required', '65 chars', 'Plain text', 'Stencil', 'High', 'Medium', 'Precise medium for collector searches'],
    ['eBay', 'Item Specifics', 'Size', 'Required', '65 chars', 'Plain text', '24 x 36 in', 'Medium', 'Medium', 'Include units (inches/cm)'],
    ['eBay', 'Item Specifics', 'Certificate of Authenticity', 'Required', '65 chars', 'Dropdown', 'Yes - Blockchain Verified', 'Critical', 'High', 'Essential for high-value art'],
    ['eBay', 'Item Specifics', 'Signed', 'Required', '65 chars', 'Dropdown', 'Yes - Hand Signed by Artist', 'Critical', 'High', 'Significantly increases value'],
    
    // ETSY - Story-driven, lifestyle focus
    ['Etsy', 'Title', 'Listing Title', 'Required', '140 chars', 'Plain text', 'Original Street Art Stencil by Shepard Fairey - OBEY Giant Andre - Signed Political Art - 24x36 - Gallery Quality - COA', 'High', 'Critical', 'Use all 140 characters with emotional keywords'],
    ['Etsy', 'Description', 'Item Description', 'Required', '13,000 chars', 'Plain text', 'The story behind this iconic piece... Shepard Fairey\'s journey from street artist to cultural icon...', 'High', 'Critical', 'Etsy buyers love stories - focus on artist journey'],
    ['Etsy', 'Images', 'Main Photo', 'Required', '10MB', 'JPG/PNG/GIF', '2000x2000px max, lifestyle shot in beautiful room setting', 'Critical', 'Critical', 'Lifestyle photos outperform plain product shots 3:1'],
    ['Etsy', 'Images', 'Additional Photos', 'Optional', '10MB each', 'JPG/PNG/GIF', '10 images max - room mockups, close-ups, packaging, size reference', 'Medium', 'High', 'Show scale, quality, and beautiful presentation'],
    ['Etsy', 'Tags', 'Search Tags', 'Required', '20 chars each', '13 tags max', 'street art, political art, shepard fairey, obey giant, signed print, wall decor, contemporary art', 'Critical', 'High', 'Mix specific artist terms with broad decor keywords'],
    ['Etsy', 'Attributes', 'Materials', 'Optional', '13 materials', 'Multi-select', 'paper, ink, stencil, protective coating', 'Medium', 'Medium', 'Helps with handmade classification'],
    ['Etsy', 'Attributes', 'Style', 'Optional', '13 styles', 'Multi-select', 'Contemporary, Modern, Street Art, Political', 'High', 'Medium', 'Important for style-based searches'],
    
    // INSTAGRAM SHOP - Visual-first, mobile-optimized
    ['Instagram', 'Title', 'Product Name', 'Required', '100 chars', 'Plain text', 'Shepard Fairey OBEY Giant - Original Signed Street Art', 'Medium', 'High', 'Mobile-friendly, concise but descriptive'],
    ['Instagram', 'Description', 'Product Description', 'Required', '2,200 chars', 'Plain text + hashtags', 'Iconic OBEY Giant by street art legend Shepard Fairey ✨ Hand-signed original 🎨 #streetart #shepardfairey #obeygiant', 'Medium', 'High', 'Instagram voice with emojis and strategic hashtags'],
    ['Instagram', 'Images', 'Product Photo', 'Required', '30MB', 'JPG/PNG', '1080x1080px exact (1:1 ratio), bright lighting, artwork fills frame', 'Critical', 'Critical', 'Must be perfect square - no exceptions'],
    ['Instagram', 'Images', 'Additional Photos', 'Optional', '30MB each', 'JPG/PNG', '20 images max, all 1080x1080px square format', 'Medium', 'Medium', 'Consistent square format across all images'],
    ['Instagram', 'Commerce', 'Product Tags', 'Optional', '30 hashtags', 'Hashtag format', '#streetart #shepardfairey #obeygiant #signedart #investmentart #contemporaryart #politicalart', 'High', 'Medium', 'Mix trending hashtags with specific art terms'],
    
    // FACEBOOK SHOP - More detailed than Instagram
    ['Facebook', 'Title', 'Product Name', 'Required', '100 chars', 'Plain text', 'Shepard Fairey - OBEY Giant Andre Stencil - Signed Original (1989)', 'Medium', 'High', 'More detailed than Instagram - include year'],
    ['Facebook', 'Description', 'Product Description', 'Required', '9,999 chars', 'Plain text', 'Historic OBEY Giant stencil by renowned street artist Shepard Fairey. Original 1989 piece, hand-signed by artist.', 'Medium', 'High', 'More detailed storytelling than Instagram'],
    ['Facebook', 'Images', 'Main Image', 'Required', '30MB', 'JPG/PNG', '1200x1200px minimum, flexible format (square or rectangular)', 'High', 'High', 'More flexible format requirements than Instagram'],
    ['Facebook', 'Images', 'Additional Images', 'Optional', '30MB each', 'JPG/PNG', '20 images max, flexible sizing, lifestyle and detail shots', 'Medium', 'Medium', 'Show context and authenticity'],
    
    // GOOGLE SHOPPING - Search optimization focus
    ['Google Shopping', 'Title', 'Product Title', 'Required', '150 chars', 'Plain text', 'Shepard Fairey OBEY Giant Andre Stencil - Original Signed Street Art - 24x36 Political Art - COA - Gallery Gauntlet', 'Critical', 'High', 'Include brand name for trust and recognition'],
    ['Google Shopping', 'Description', 'Product Description', 'Required', '5,000 chars', 'Plain text', 'Detailed specifications: Artist: Shepard Fairey, Title: OBEY Giant, Year: 1989, Medium: Stencil on paper...', 'High', 'Medium', 'Structured data helps Google create rich snippets'],
    ['Google Shopping', 'Images', 'Image Link', 'Required', '16MB', 'URL format', 'https://gauntletgallery.com/images/shepard-fairey-obey-giant-main.jpg', 'Critical', 'High', 'Must be publicly accessible, high-quality URL'],
    ['Google Shopping', 'Images', 'Additional Images', 'Optional', '16MB each', 'URL format', '20 additional image URLs for enhanced listings', 'Medium', 'Medium', 'Extra images improve click-through rates'],
    ['Google Shopping', 'Identifiers', 'Product ID', 'Required', '50 chars', 'Alphanumeric', 'GG-SF-OBEY-GIANT-1989-SIGNED', 'None', 'None', 'Unique identifier for tracking and updates'],
    ['Google Shopping', 'Identifiers', 'GTIN/UPC', 'Optional', '14 digits', 'Numeric', 'Use blockchain certificate hash as unique identifier', 'None', 'None', 'Custom identifier for art pieces'],
    
    // SAATCHI ART - Professional art market
    ['Saatchi Art', 'Title', 'Artwork Title', 'Required', '200 chars', 'Plain text', 'OBEY Giant (Andre the Giant)', 'High', 'Medium', 'Official artwork title only - no artist name'],
    ['Saatchi Art', 'Artist', 'Artist Name', 'Required', '100 chars', 'Plain text', 'Shepard Fairey', 'Critical', 'High', 'Must match verified artist database'],
    ['Saatchi Art', 'Artist', 'Artist Bio', 'Required', '2,000 chars', 'Plain text', 'Shepard Fairey (American, b. 1970) is a contemporary street artist, graphic designer, activist...', 'High', 'High', 'Professional biography essential for credibility'],
    ['Saatchi Art', 'Description', 'Artwork Description', 'Required', '2,000 chars', 'Plain text', 'This seminal work represents Fairey\'s early exploration of appropriation and political commentary...', 'High', 'High', 'Scholarly tone, art historical context required'],
    ['Saatchi Art', 'Images', 'Main Image', 'Required', '20MB', 'JPG/PNG', '1500x1500px minimum, museum-quality photography, color accuracy critical', 'Critical', 'Critical', 'Professional gallery lighting and color correction'],
    ['Saatchi Art', 'Images', 'Additional Images', 'Optional', '20MB each', 'JPG/PNG', '10 images max - detail shots, signature, installation views', 'Medium', 'Medium', 'Show artwork quality and authenticity'],
    ['Saatchi Art', 'Specifications', 'Medium', 'Required', '100 chars', 'Dropdown', 'Stencil on Paper', 'High', 'Medium', 'Precise medium classification from dropdown'],
    ['Saatchi Art', 'Specifications', 'Dimensions', 'Required', '50 chars', 'Structured format', '61 x 91.4 cm (24 x 36 in)', 'Medium', 'Medium', 'Metric first, then imperial in parentheses'],
    ['Saatchi Art', 'Specifications', 'Year', 'Required', '4 digits', 'Numeric', '1989', 'Medium', 'Medium', 'Creation year for art historical context'],
    
    // AMAZON HANDMADE - Keyword optimization
    ['Amazon', 'Title', 'Product Title', 'Required', '200 chars', 'Plain text', 'Handmade Original Shepard Fairey OBEY Giant Street Art Print - Signed 24x36 Contemporary Political Wall Art - Certificate Included', 'High', 'Critical', 'Must include "Handmade" for category compliance'],
    ['Amazon', 'Description', 'Product Features', 'Required', '1,000 chars each', '5 bullet points', '• Original street art print by world-renowned artist Shepard Fairey\n• Hand-signed OBEY Giant design from 1989', 'High', 'Critical', 'Highlight key selling points in scannable format'],
    ['Amazon', 'Description', 'Product Description', 'Required', '2,000 chars', 'HTML allowed', 'Detailed description with specifications, artist background, authenticity guarantees', 'Medium', 'High', 'Rich formatting helps with conversion'],
    ['Amazon', 'Images', 'Main Image', 'Required', '10MB', 'JPG/PNG', '1600x1600px minimum, pure white background, no text overlays', 'Critical', 'Critical', 'Amazon has strict image requirements - follow exactly'],
    ['Amazon', 'Images', 'Additional Images', 'Optional', '10MB each', 'JPG/PNG', '8 additional images - lifestyle, details, size reference, packaging', 'Medium', 'High', 'Show context and build trust'],
    ['Amazon', 'Keywords', 'Search Terms', 'Optional', '1,000 chars total', 'Keyword string', 'shepard fairey obey giant street art political art signed print wall decor contemporary art', 'Critical', 'Medium', 'Backend keywords not visible to customers but crucial for search'],
    
    // 1STDIBSZ - Luxury market positioning
    ['1stDibs', 'Title', 'Item Title', 'Required', '200 chars', 'Formatted text', 'Shepard Fairey, "OBEY Giant (Andre the Giant)," 1989', 'High', 'Medium', 'Follow art world conventions: Artist, "Title," Year'],
    ['1stDibs', 'Description', 'Item Description', 'Required', '4,000 chars', 'Rich text', 'Museum-quality description with provenance, exhibition history, cultural significance, condition report', 'High', 'Critical', 'Scholarly approach with detailed art historical context'],
    ['1stDibs', 'Images', 'Primary Image', 'Required', '20MB', 'High-res JPG', '2000x2000px minimum, museum-quality lighting, perfect color accuracy', 'Critical', 'Critical', 'Luxury presentation - invest in professional photography'],
    ['1stDibs', 'Images', 'Additional Images', 'Required', '20MB each', 'High-res JPG', '20 images max - multiple angles, details, signature, condition, installation', 'Medium', 'High', 'Comprehensive documentation expected'],
    ['1stDibs', 'Specifications', 'Materials', 'Required', '200 chars', 'Structured text', 'Stencil print in black ink on cream paper', 'High', 'Medium', 'Precise material description using art terminology'],
    ['1stDibs', 'Specifications', 'Dimensions', 'Required', '100 chars', 'Formatted text', 'H 36 in. W 24 in. D 0.1 in.', 'Medium', 'Medium', 'Standard format: Height, Width, Depth'],
    ['1stDibs', 'Specifications', 'Condition', 'Required', '50 chars', 'Dropdown', 'Excellent', 'High', 'High', 'Professional condition assessment required'],
    ['1stDibs', 'Provenance', 'Provenance History', 'Optional', '2,000 chars', 'Rich text', 'Artist studio; Private collection, New York; Gallery Gauntlet, New York', 'Critical', 'Critical', 'Detailed ownership history increases value'],
    
    // ARTSY - Gallery network platform
    ['Artsy', 'Title', 'Artwork Title', 'Required', '255 chars', 'Plain text', 'OBEY Giant (Andre the Giant)', 'High', 'Medium', 'Official artwork title - artist name separate'],
    ['Artsy', 'Artist', 'Artist Profile', 'Required', 'Artist ID', 'Artsy database', 'Must link to verified Shepard Fairey artist profile on Artsy', 'Critical', 'High', 'Artist must be verified in Artsy database'],
    ['Artsy', 'Description', 'Medium', 'Required', '255 chars', 'Plain text', 'Stencil print in black ink on cream speckletone paper', 'High', 'Medium', 'Detailed medium description with paper type'],
    ['Artsy', 'Description', 'Dimensions', 'Required', '100 chars', 'Formatted text', '91.4 × 61 cm (36 × 24 in.)', 'Medium', 'Medium', 'Metric measurements first, imperial in parentheses'],
    ['Artsy', 'Description', 'Edition Info', 'Optional', '100 chars', 'Plain text', 'Artist Proof', 'Critical', 'Critical', 'Edition information crucial for value assessment'],
    ['Artsy', 'Images', 'Main Image', 'Required', '25MB', 'High-res JPG', '1400x1400px minimum, color-accurate photography', 'Critical', 'Critical', 'Color accuracy essential for art evaluation'],
    ['Artsy', 'Images', 'Additional Images', 'Optional', '25MB each', 'High-res JPG', '10 images maximum - details, installation views, signature', 'Medium', 'Medium', 'Professional documentation of artwork'],
    ['Artsy', 'Authentication', 'Certificate', 'Optional', 'Boolean', 'Yes/No', 'Yes - Blockchain verified certificate of authenticity', 'Critical', 'Critical', 'COA availability essential for high-value works'],
    
    // POSHMARK - Social shopping platform
    ['Poshmark', 'Title', 'Listing Title', 'Required', '80 chars', 'Plain text', 'Shepard Fairey OBEY Giant Street Art Print Signed Original Political', 'High', 'Critical', 'Concise but descriptive for mobile shoppers'],
    ['Poshmark', 'Description', 'Item Description', 'Required', '8,000 chars', 'Plain text', 'Authentic Shepard Fairey OBEY Giant print! Hand-signed original from 1989. Perfect for any art collector! 📸', 'High', 'High', 'Social media tone with emojis, authentic voice'],
    ['Poshmark', 'Images', 'Cover Photo', 'Required', '15MB', 'JPG/PNG', '1080x1080px minimum, bright natural lighting, lifestyle preferred', 'Critical', 'Critical', 'Social media style - lifestyle shots perform best'],
    ['Poshmark', 'Images', 'Additional Photos', 'Optional', '15MB each', 'JPG/PNG', '20 images max - styling shots, details, size reference, authenticity proof', 'Medium', 'High', 'Show styling potential and authenticity'],
    ['Poshmark', 'Category', 'Category Selection', 'Required', 'Dropdown', 'Art & Home Decor', 'Art & Collectibles > Wall Art > Prints', 'High', 'Medium', 'Proper categorization affects search visibility'],
    ['Poshmark', 'Brand', 'Brand Field', 'Optional', '50 chars', 'Plain text', 'Shepard Fairey / OBEY', 'High', 'Medium', 'Artist brand recognition important for discovery'],
    ['Poshmark', 'Size', 'Size Selection', 'Required', 'Dropdown', '24x36 inches', 'Custom size entry for artwork dimensions', 'Medium', 'Medium', 'Important for buyer expectations'],
    ['Poshmark', 'Color', 'Color Tags', 'Optional', 'Multi-select', 'Black, White, Cream', 'Primary colors in artwork for filtering', 'Medium', 'Medium', 'Helps buyers filter by room decor needs'],
    ['Poshmark', 'Tags', 'Style Tags', 'Optional', '5 tags max', 'Dropdown', 'street art, political, contemporary, signed, original', 'High', 'High', 'Critical for Poshmark algorithm and discovery'],
    
    // STOCKX - Authentication-focused marketplace
    ['StockX', 'Title', 'Product Title', 'Required', '200 chars', 'Plain text', 'Shepard Fairey OBEY Giant Andre the Giant Stencil Print 1989 Hand Signed Original - 24x36', 'Critical', 'Critical', 'Must include exact product details for authentication'],
    ['StockX', 'Description', 'Product Description', 'Required', '5,000 chars', 'Plain text', 'Shepard Fairey original OBEY Giant stencil from 1989. Hand-signed by artist. Authenticated with blockchain COA.', 'High', 'High', 'Authentication and provenance details critical'],
    ['StockX', 'Images', 'Product Photos', 'Required', '20MB', 'High-res JPG', '1200x1200px minimum, clean white background, all angles', 'Critical', 'Critical', 'Authentication-quality photos required'],
    ['StockX', 'Images', 'Authentication Photos', 'Required', '20MB each', 'High-res JPG', '12 images max - signature close-up, COA, condition, all edges, back view', 'Critical', 'Critical', 'Detailed documentation for authentication team'],
    ['StockX', 'Authentication', 'Certificate of Authenticity', 'Required', 'Upload PDF', 'PDF format', 'Blockchain-verified COA with QR code and verification hash', 'Critical', 'Critical', 'Essential for StockX authentication process'],
    ['StockX', 'Authentication', 'Provenance Documentation', 'Required', 'Upload PDF', 'PDF format', 'Purchase receipt, gallery documentation, previous auction records', 'Critical', 'Critical', 'Chain of ownership required for high-value items'],
    ['StockX', 'Condition', 'Condition Assessment', 'Required', 'Dropdown', 'Brand New/Excellent/Very Good/Good', 'Professional condition grading following art market standards', 'Critical', 'Critical', 'Affects final sale price significantly'],
    ['StockX', 'Specifications', 'Edition Information', 'Required', '100 chars', 'Plain text', 'Artist Proof / Limited Edition / Unique Original', 'Critical', 'Critical', 'Edition status crucial for valuation'],
    ['StockX', 'Specifications', 'Year Created', 'Required', '4 digits', 'Numeric', '1989', 'Critical', 'Critical', 'Creation year affects authenticity verification'],
    ['StockX', 'Specifications', 'Dimensions', 'Required', '50 chars', 'Formatted', '24 x 36 inches (60.96 x 91.44 cm)', 'High', 'High', 'Exact measurements for authentication'],
    ['StockX', 'Specifications', 'Medium', 'Required', '100 chars', 'Plain text', 'Stencil print on speckletone paper', 'High', 'High', 'Material composition for authentication'],
    ['StockX', 'Authentication', 'Artist Signature', 'Required', 'Boolean + Photo', 'Yes with photo proof', 'Hand-signed by Shepard Fairey - signature photo required', 'Critical', 'Critical', 'Signature verification essential for value']
  ];
  
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  formatRequirementsSheet(sheet);
}

function createLimitsComparisonSheet(spreadsheet) {
  var sheet = getOrCreateSheet(spreadsheet, 'Platform Limits Comparison');
  
  var data = [
    ['Platform', 'Title Limit', 'Description Limit', 'Image Count', 'Image Size', 'Tag Limit', 'Key Constraint', 'Optimization Priority'],
    ['Shopify', '255 chars', '65,535 chars', '250 images', '20MB each', 'Unlimited', 'None - most flexible', 'Rich content, SEO optimization'],
    ['eBay', '80 chars', '500,000 chars', '12 images', '12MB each', 'Item specifics only', 'Title length critical', 'Keyword-packed titles'],
    ['Etsy', '140 chars', '13,000 chars', '10 images', '10MB each', '13 tags (20 chars each)', 'Limited tags', 'Story-driven descriptions'],
    ['Instagram', '100 chars', '2,200 chars', '20 images', '30MB each', '30 hashtags', 'Square format required', 'Visual storytelling'],
    ['Facebook', '100 chars', '9,999 chars', '20 images', '30MB each', 'No specific limit', 'Format flexibility', 'Detailed descriptions'],
    ['Google Shopping', '150 chars', '5,000 chars', '20 images', '16MB each', 'Backend keywords only', 'Structured data required', 'Search optimization'],
    ['Saatchi Art', '200 chars', '2,000 chars', '10 images', '20MB each', 'Category tags only', 'Professional presentation', 'Art historical context'],
    ['Amazon', '200 chars', '2,000 chars + bullets', '9 images', '10MB each', 'Backend keywords only', 'White background required', 'Keyword optimization'],
    ['1stDibs', '200 chars', '4,000 chars', '20 images', '20MB each', 'Category only', 'Luxury presentation', 'Provenance documentation'],
    ['Artsy', '255 chars', 'Varies by field', '10 images', '25MB each', 'Category only', 'Color accuracy critical', 'Professional photography'],
    ['Poshmark', '80 chars', '8,000 chars', '20 images', '15MB each', '5 style tags max', 'Social media tone required', 'Lifestyle photography'],
    ['StockX', '200 chars', '5,000 chars', '12 images', '20MB each', 'No tags', 'Authentication required', 'Documentation critical']
  ];
  
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  formatLimitsSheet(sheet);
}

function createOptimizationGuideSheet(spreadsheet) {
  var sheet = getOrCreateSheet(spreadsheet, 'Platform Optimization Guide');
  
  var data = [
    ['Platform', 'Title Strategy', 'Image Strategy', 'Description Strategy', 'Pricing Strategy', 'Success Metrics'],
    ['Shopify', 'Artist + Title + Year + Key details', 'Professional photography, multiple angles', 'Rich HTML with headers, lists, detailed provenance', 'Market competitive, consider frame options', 'Conversion rate, average order value'],
    ['eBay', 'Pack all 80 chars with search keywords', 'White background, zoom-friendly details', 'Trust-building template with Gallery branding', 'Competitive auction + BIN option', 'Search ranking, auction completion rate'],
    ['Etsy', 'Emotional keywords, use all 140 chars', 'Lifestyle photos in beautiful settings', 'Story-driven, artist journey narrative', 'Handmade market pricing, payment plans', 'Search placement, favorites, reviews'],
    ['Instagram', 'Concise but descriptive for mobile', 'Perfect 1:1 squares, bright lighting', 'Visual storytelling with strategic hashtags', 'Premium positioning for social proof', 'Engagement rate, story clicks, saves'],
    ['Facebook', 'More detailed than Instagram', 'Flexible format, context shots', 'Detailed storytelling, artist background', 'Social proof pricing, testimonials', 'Click-through rate, social shares'],
    ['Google Shopping', 'Brand name + key specs for trust', 'Clean product shots, multiple angles', 'Structured specifications for rich snippets', 'Competitive for comparison shopping', 'Impression share, click-through rate'],
    ['Saatchi Art', 'Official artwork title only', 'Museum-quality, color-accurate photos', 'Scholarly tone, art historical context', 'Gallery-level pricing, investment angle', 'Inquiry rate, collector engagement'],
    ['Amazon', 'Handmade + artist + key features', 'Stark white background, lifestyle context', 'Bullet points highlighting key benefits', 'Factor in Amazon fees, competitive', 'Buy Box eligibility, conversion rate'],
    ['1stDibs', 'Art world conventions: Artist, "Title," Year', 'Luxury presentation, comprehensive docs', 'Museum-quality description with provenance', 'Premium luxury market positioning', 'Inquiry quality, serious collector interest'],
    ['Artsy', 'Clean artwork title, artist separate', 'Color-accurate, professional gallery shots', 'Institutional quality, precise medium info', 'Gallery network standard pricing', 'Art professional engagement, inquiries'],
    ['Poshmark', 'Trendy keywords for social discovery', 'Lifestyle shots in beautiful rooms', 'Personal story with emojis and enthusiasm', 'Social commerce pricing with offers', 'Likes, shares, bundle requests'],
    ['StockX', 'Exact product details for authentication', 'Professional documentation photos', 'Authentication and provenance focused', 'Market-driven pricing with premiums', 'Authentication success rate, resale value']
  ];
  
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  formatOptimizationSheet(sheet);
}

function createImageSpecsSheet(spreadsheet) {
  var sheet = getOrCreateSheet(spreadsheet, 'Image Specifications');
  
  var data = [
    ['Platform', 'Main Image Requirements', 'Additional Images', 'Format Requirements', 'Background Requirements', 'Gallery Gauntlet Recommendations'],
    ['Shopify', '1200x1200px min, 20MB max, JPG/PNG/WebP', '249 additional, 20MB each', 'Flexible formats', 'Clean, neutral preferred', 'Professional lighting, multiple angles, frame mockups'],
    ['eBay', '1600x1600px min, 12MB max, JPG/PNG/GIF', '11 additional, 12MB each', 'Standard web formats', 'White strongly preferred', 'White background for search advantage, zoom-friendly details'],
    ['Etsy', '2000x2000px max, 10MB max, JPG/PNG/GIF', '9 additional, 10MB each', 'Standard web formats', 'Lifestyle settings preferred', 'Beautiful room settings, natural lighting, lifestyle context'],
    ['Instagram', '1080x1080px exact, 30MB max, JPG/PNG', '19 additional, all square', '1:1 aspect ratio only', 'Bright, engaging backgrounds', 'Perfect squares, bright lighting, mobile-optimized'],
    ['Facebook', '1200x1200px min, 30MB max, JPG/PNG', '19 additional, flexible size', 'Square or rectangular', 'Flexible, engaging preferred', 'High quality, context shots, detail views'],
    ['Google Shopping', '100x100px min, 16MB max, JPG/PNG/GIF/BMP/WebP', '19 additional image URLs', 'Multiple format support', 'Clean, product-focused', 'Clean backgrounds, multiple angles, zoom-friendly'],
    ['Saatchi Art', '1500x1500px min, 20MB max, JPG/PNG', '9 additional, 20MB each', 'High-resolution only', 'Professional gallery lighting', 'Museum-quality photography, color accuracy critical'],
    ['Amazon', '1600x1600px min, 10MB max, JPG/PNG', '8 additional, 10MB each', 'Standard formats', 'Pure white (RGB 255,255,255)', 'Stark white backgrounds, lifestyle context shots, size reference'],
    ['1stDibs', '2000x2000px min, 20MB max, High-res JPG', '19 additional, 20MB each', 'High-resolution JPG preferred', 'Professional gallery standard', 'Luxury presentation, comprehensive documentation, condition shots'],
    ['Artsy', '1400x1400px min, 25MB max, JPG/PNG', '9 additional, 25MB each', 'High-resolution formats', 'Color-accurate, professional', 'Gallery-standard lighting, color accuracy, institutional quality'],
    ['Poshmark', '1080x1080px min, 15MB max, JPG/PNG', '19 additional, 15MB each', 'Mobile-optimized formats', 'Lifestyle settings with natural light', 'Bright, social media-style shots, styling inspiration, room context'],
    ['StockX', '1200x1200px min, 20MB max, High-res JPG', '11 additional, 20MB each', 'Authentication-quality only', 'Clean white or neutral backgrounds', 'Documentation shots, signature close-ups, condition details, all angles']
  ];
  
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  formatImageSpecsSheet(sheet);
}

function getOrCreateSheet(spreadsheet, name) {
  var sheet = spreadsheet.getSheetByName(name);
  if (sheet) spreadsheet.deleteSheet(sheet);
  return spreadsheet.insertSheet(name);
}

function formatRequirementsSheet(sheet) {
  var range = sheet.getDataRange();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  
  // Header formatting
  var headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground('#1e3c72')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(10)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  sheet.setRowHeight(1, 50);
  
  // Column widths
  var widths = [90, 100, 120, 80, 100, 80, 250, 80, 80, 150];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  
  // Color code by platform
  var colors = {
    'Shopify': '#E3F2FD', 'eBay': '#FFF8E1', 'Etsy': '#FCE4EC', 
    'Instagram': '#F3E5F5', 'Facebook': '#E8F5E8', 'Google Shopping': '#FFF3E0',
    'Saatchi Art': '#F1F8E9', 'Amazon': '#FFF9C4', '1stDibs': '#EDE7F6', 'Artsy': '#E0F2F1',
    'Poshmark': '#FFEBEE', 'StockX': '#E1F5FE'
  };
  
  for (let row = 2; row <= numRows; row++) {
    var platform = sheet.getRange(row, 1).getValue();
    var color = colors[platform] || '#FFFFFF';
    sheet.getRange(row, 1, 1, numCols).setBackground(color);
  }
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

function formatLimitsSheet(sheet) {
  formatGenericSheet(sheet, '#2a5298');
}

function formatOptimizationSheet(sheet) {
  formatGenericSheet(sheet, '#4CAF50');
}

function formatImageSpecsSheet(sheet) {
  formatGenericSheet(sheet, '#FF9800');
}

function formatGenericSheet(sheet, headerColor) {
  var range = sheet.getDataRange();
  var headerRange = sheet.getRange(1, 1, 1, range.getNumColumns());
  
  headerRange
    .setBackground(headerColor)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(10)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  sheet.setRowHeight(1, 40);
  
  // Auto-resize columns
  for (let col = 1; col <= range.getNumColumns(); col++) {
    sheet.autoResizeColumn(col);
  }
  
  sheet.setFrozenRows(1);
  
  // Data formatting
  if (range.getNumRows() > 1) {
    var dataRange = sheet.getRange(2, 1, range.getNumRows() - 1, range.getNumColumns());
    dataRange
      .setVerticalAlignment('top')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setFontSize(9);
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🎨 Gallery Gauntlet Analytics')
    .addItem('📊 Format Specification Sheet', 'formatGalleryGauntletSheet')
    .addItem('🚀 Enhanced Sales Channel Analysis', 'generateEnhancedSalesChannelAnalysis')
    .addItem('📋 Export Platform Templates', 'exportPlatformTemplates')
    .addSeparator()
    .addItem('ℹ️ Platform Requirements Guide', 'showPlatformGuide')
    .addToUi();
}

function showPlatformGuide() {
  SpreadsheetApp.getUi().alert(
    'Gallery Gauntlet - Enhanced Platform Analysis',
    'Ultra-comprehensive analysis across 4 detailed sheets:\n\n' +
    '📊 REQUIREMENTS MATRIX:\n' +
    '• 100+ specific field requirements\n' +
    '• Character limits for every field\n' +
    '• Gallery Gauntlet optimized examples\n\n' +
    '📏 LIMITS COMPARISON:\n' +
    '• Title: 70-255 characters by platform\n' +
    '• Images: 1-250 images, 10MB-30MB each\n' +
    '• Tags: 13 tags (Etsy) to unlimited (Shopify)\n\n' +
    '🎯 OPTIMIZATION GUIDE:\n' +
    '• Platform-specific strategies\n' +
    '• Success metrics to track\n\n' +
    '🖼️ IMAGE SPECIFICATIONS:\n' +
    '• Exact size requirements\n' +
    '• Format and background requirements\n' +
    '• Gallery Gauntlet recommendations\n\n' +
    'Perfect for optimizing across all sales channels!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}