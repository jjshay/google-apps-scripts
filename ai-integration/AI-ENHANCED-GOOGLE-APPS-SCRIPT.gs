/**
 * AI-ENHANCED GOOGLE APPS SCRIPT FOR EBAY ART LISTINGS
 * =====================================================
 * Complete script with ChatGPT and Gemini integration
 * 
 * Features:
 * - Import JSON data from eBay Art Processor
 * - Auto-create Google Drive folders
 * - AI-powered description generation (ChatGPT/Gemini)
 * - Smart prompt templates with examples
 * - Target audience customization
 */

// ==================== CONFIGURATION ====================
const CONFIG = {
  MAIN_FOLDER_NAME: 'eBay Art Listings',
  IMAGE_FOLDER_PREFIX: 'Images_',
  SHEET_NAME: 'Listings',
  DRIVE_FOLDER_ID: '', // Optional: Set specific parent folder ID
  
  // AI Configuration (Add your API keys here)
  OPENAI_API_KEY: 'YOUR_OPENAI_API_KEY_HERE', // Get from: https://platform.openai.com/api-keys
  GEMINI_API_KEY: 'YOUR_GEMINI_API_KEY_HERE', // Get from: https://makersuite.google.com/app/apikey
  
  // Default AI Settings
  DEFAULT_MODEL: 'gpt-4', // or 'gpt-3.5-turbo' or 'gemini-pro'
  MAX_TOKENS: 500,
  TEMPERATURE: 0.7
};

// ==================== MENU & INITIALIZATION ====================
/**
 * Creates menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Main menu
  ui.createMenu('🎨 Art Import')
    .addItem('📥 Import from Clipboard', 'showImportDialog')
    .addItem('📁 Create Image Folders', 'createImageFolders')
    .addItem('🔗 Generate Drive Links', 'generateDriveLinks')
    .addSeparator()
    .addSubMenu(ui.createMenu('🤖 AI Tools')
      .addItem('✨ Generate Description (ChatGPT)', 'generateDescriptionChatGPT')
      .addItem('✨ Generate Description (Gemini)', 'generateDescriptionGemini')
      .addItem('🎯 Generate Tags', 'generateTags')
      .addItem('💰 Suggest Pricing', 'suggestPricing')
      .addItem('📝 Improve Title', 'improveTitle')
      .addSeparator()
      .addItem('📚 Show AI Examples', 'showAIExamples')
      .addItem('⚙️ AI Settings', 'showAISettings'))
    .addSeparator()
    .addItem('⚙️ Settings', 'showSettings')
    .addItem('ℹ️ Help', 'showHelp')
    .addToUi();
}

// ==================== AI FUNCTIONS ====================

/**
 * Generate description using ChatGPT
 */
function generateDescriptionChatGPT() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select cells first');
    return;
  }
  
  // Show prompt dialog
  const html = HtmlService.createHtmlOutput(getAIPromptHTML('chatgpt', 'description'))
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '🤖 Generate Description with ChatGPT');
}

/**
 * Generate description using Gemini
 */
function generateDescriptionGemini() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select cells first');
    return;
  }
  
  // Show prompt dialog
  const html = HtmlService.createHtmlOutput(getAIPromptHTML('gemini', 'description'))
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '✨ Generate Description with Gemini');
}

/**
 * Process AI request from dialog
 * @param {Object} params - Request parameters
 * @return {string} Generated content
 */
function processAIRequest(params) {
  try {
    const {
      model,
      task,
      data,
      targetAge,
      targetGender,
      targetLocation,
      style,
      customPrompt
    } = params;
    
    // Build context from selected cells
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    const values = range.getValues();
    
    // Get cell references
    const startRow = range.getRow();
    const startCol = range.getColumn();
    
    // Build data context
    let context = 'Selected Data:\n';
    values.forEach((row, i) => {
      row.forEach((cell, j) => {
        if (cell) {
          const cellRef = sheet.getRange(startRow + i, startCol + j).getA1Notation();
          context += `${cellRef}: ${cell}\n`;
        }
      });
    });
    
    // Build the prompt
    const prompt = buildAIPrompt({
      task,
      context,
      data,
      targetAge,
      targetGender,
      targetLocation,
      style,
      customPrompt
    });
    
    // Call appropriate AI service
    let result;
    if (model === 'chatgpt') {
      result = callChatGPT(prompt);
    } else if (model === 'gemini') {
      result = callGemini(prompt);
    } else {
      throw new Error('Invalid AI model selected');
    }
    
    // Store result in designated cell if specified
    if (params.targetCell) {
      const targetRange = sheet.getRange(params.targetCell);
      targetRange.setValue(result);
    }
    
    // Log AI usage
    logAIUsage(model, task, prompt.length, result.length);
    
    return result;
    
  } catch (error) {
    console.error('AI processing error:', error);
    throw new Error('AI generation failed: ' + error.toString());
  }
}

/**
 * Build AI prompt based on parameters
 * @param {Object} params - Prompt parameters
 * @return {string} Complete prompt
 */
function buildAIPrompt(params) {
  const {
    task,
    context,
    data,
    targetAge,
    targetGender,
    targetLocation,
    style,
    customPrompt
  } = params;
  
  let prompt = '';
  
  // Task-specific prompts
  switch(task) {
    case 'description':
      prompt = `Generate a compelling eBay listing description for an artwork.

Context:
${context}

Target Audience:
- Age: ${targetAge || 'All ages'}
- Gender: ${targetGender || 'All'}
- Location: ${targetLocation || 'Worldwide'}

Style: ${style || 'Professional and engaging'}

Requirements:
1. Create an engaging opening that captures attention
2. Describe the artwork's visual elements and artistic technique
3. Highlight what makes this piece special or unique
4. Include emotional appeal and investment potential
5. End with a clear call to action
6. Keep it between 150-300 words
7. Use HTML formatting (<p>, <strong>, <ul>, <li>) for eBay

${customPrompt ? 'Additional instructions: ' + customPrompt : ''}

Generate the description now:`;
      break;
      
    case 'tags':
      prompt = `Generate relevant tags for this artwork listing.

Context:
${context}

Requirements:
1. Generate 10-15 relevant tags
2. Include artist name, style, medium, subject matter
3. Add trending keywords for the target audience
4. Include investment/collectible related terms
5. Format as comma-separated list

Target Audience: ${targetAge} ${targetGender} in ${targetLocation}

Generate tags:`;
      break;
      
    case 'title':
      prompt = `Improve this eBay listing title for maximum visibility and appeal.

Current Title/Context:
${context}

Requirements:
1. Maximum 80 characters
2. Include artist name and key artwork details
3. Add compelling descriptors
4. Include relevant keywords for search
5. Make it stand out from competitors

Target Market: ${targetAge} ${targetGender} in ${targetLocation}

Improved title:`;
      break;
      
    case 'pricing':
      prompt = `Suggest optimal pricing strategy for this artwork.

Artwork Details:
${context}

Market Factors:
- Target Age: ${targetAge}
- Target Gender: ${targetGender}
- Target Location: ${targetLocation}

Consider:
1. Artist reputation and market value
2. Size and medium
3. Condition and rarity
4. Current market trends
5. Target audience purchasing power

Provide:
1. Suggested starting price
2. Buy it now price
3. Reserve price (if applicable)
4. Pricing justification

Pricing recommendation:`;
      break;
      
    default:
      prompt = customPrompt || 'Please provide instructions for the AI.';
  }
  
  return prompt;
}

/**
 * Call ChatGPT API
 * @param {string} prompt - The prompt
 * @return {string} Generated text
 */
function callChatGPT(prompt) {
  if (!CONFIG.OPENAI_API_KEY || CONFIG.OPENAI_API_KEY === 'YOUR_OPENAI_API_KEY_HERE') {
    throw new Error('Please add your OpenAI API key in the script configuration');
  }
  
  const url = 'https://api.openai.com/v1/chat/completions';
  
  const payload = {
    model: CONFIG.DEFAULT_MODEL,
    messages: [
      {
        role: 'system',
        content: 'You are an expert art dealer and eBay listing specialist with deep knowledge of contemporary art markets.'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    max_tokens: CONFIG.MAX_TOKENS,
    temperature: CONFIG.TEMPERATURE
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.OPENAI_API_KEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.error) {
      throw new Error(json.error.message);
    }
    
    return json.choices[0].message.content;
    
  } catch (error) {
    console.error('ChatGPT API error:', error);
    throw new Error('ChatGPT API call failed: ' + error.toString());
  }
}

/**
 * Call Gemini API
 * @param {string} prompt - The prompt
 * @return {string} Generated text
 */
function callGemini(prompt) {
  if (!CONFIG.GEMINI_API_KEY || CONFIG.GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY_HERE') {
    throw new Error('Please add your Gemini API key in the script configuration');
  }
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
  
  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: CONFIG.TEMPERATURE,
      maxOutputTokens: CONFIG.MAX_TOKENS
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.error) {
      throw new Error(json.error.message);
    }
    
    return json.candidates[0].content.parts[0].text;
    
  } catch (error) {
    console.error('Gemini API error:', error);
    throw new Error('Gemini API call failed: ' + error.toString());
  }
}

/**
 * Generate tags using AI
 */
function generateTags() {
  const params = {
    model: 'chatgpt',
    task: 'tags'
  };
  
  const result = processAIRequest(params);
  SpreadsheetApp.getUi().alert('Generated Tags', result, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Suggest pricing using AI
 */
function suggestPricing() {
  const params = {
    model: 'chatgpt',
    task: 'pricing'
  };
  
  const result = processAIRequest(params);
  SpreadsheetApp.getUi().alert('Pricing Suggestion', result, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Improve title using AI
 */
function improveTitle() {
  const params = {
    model: 'chatgpt',
    task: 'title'
  };
  
  const result = processAIRequest(params);
  SpreadsheetApp.getUi().alert('Improved Title', result, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Show AI examples
 */
function showAIExamples() {
  const html = HtmlService.createHtmlOutput(getAIExamplesHTML())
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '📚 AI Prompt Examples');
}

/**
 * Show AI settings
 */
function showAISettings() {
  const html = `
    <div style="padding: 20px; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;">
      <h3 style="color: #667eea;">🤖 AI Configuration</h3>
      
      <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 20px 0;">
        <h4>API Keys Status</h4>
        <p>🔑 OpenAI: ${CONFIG.OPENAI_API_KEY && CONFIG.OPENAI_API_KEY !== 'YOUR_OPENAI_API_KEY_HERE' ? '✅ Configured' : '❌ Not configured'}</p>
        <p>🔑 Gemini: ${CONFIG.GEMINI_API_KEY && CONFIG.GEMINI_API_KEY !== 'YOUR_GEMINI_API_KEY_HERE' ? '✅ Configured' : '❌ Not configured'}</p>
      </div>
      
      <div style="background: #fff3cd; padding: 15px; border-radius: 8px; margin: 20px 0;">
        <h4>⚙️ Current Settings</h4>
        <p><strong>Default Model:</strong> ${CONFIG.DEFAULT_MODEL}</p>
        <p><strong>Max Tokens:</strong> ${CONFIG.MAX_TOKENS}</p>
        <p><strong>Temperature:</strong> ${CONFIG.TEMPERATURE}</p>
      </div>
      
      <div style="background: #e3f2fd; padding: 15px; border-radius: 8px;">
        <h4>📝 How to Configure</h4>
        <ol>
          <li>Get OpenAI API key from: <a href="https://platform.openai.com/api-keys" target="_blank">OpenAI Platform</a></li>
          <li>Get Gemini API key from: <a href="https://makersuite.google.com/app/apikey" target="_blank">Google AI Studio</a></li>
          <li>Edit the CONFIG object in the script editor</li>
          <li>Save and reload the spreadsheet</li>
        </ol>
      </div>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '⚙️ AI Settings');
}

/**
 * Get AI prompt HTML dialog
 * @param {string} model - AI model
 * @param {string} task - Task type
 * @return {string} HTML content
 */
function getAIPromptHTML(model, task) {
  return `<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }
      
      body {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        padding: 20px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      }
      
      .container {
        background: white;
        border-radius: 15px;
        padding: 25px;
        box-shadow: 0 10px 40px rgba(0,0,0,0.2);
      }
      
      h2 {
        color: #1a1a2e;
        margin-bottom: 20px;
      }
      
      .selected-data {
        background: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        margin: 15px 0;
        max-height: 150px;
        overflow-y: auto;
      }
      
      .form-group {
        margin-bottom: 15px;
      }
      
      label {
        display: block;
        margin-bottom: 5px;
        font-weight: 600;
        color: #475569;
      }
      
      input, select, textarea {
        width: 100%;
        padding: 10px;
        border: 2px solid #e2e8f0;
        border-radius: 8px;
        font-size: 14px;
      }
      
      input:focus, select:focus, textarea:focus {
        outline: none;
        border-color: #667eea;
      }
      
      .target-fields {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 10px;
      }
      
      .example-prompts {
        background: #f0f9ff;
        padding: 15px;
        border-radius: 8px;
        margin: 15px 0;
      }
      
      .example-prompts h4 {
        color: #0369a1;
        margin-bottom: 10px;
      }
      
      .example-item {
        padding: 8px;
        background: white;
        border-radius: 5px;
        margin-bottom: 5px;
        cursor: pointer;
        transition: all 0.2s;
      }
      
      .example-item:hover {
        background: #e0f2fe;
        transform: translateX(5px);
      }
      
      button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 12px 24px;
        border-radius: 8px;
        font-size: 15px;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.3s;
      }
      
      button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(102, 126, 234, 0.3);
      }
      
      button:disabled {
        background: #cbd5e1;
        cursor: not-allowed;
      }
      
      .button-group {
        display: flex;
        gap: 10px;
        margin-top: 20px;
      }
      
      .result {
        margin-top: 20px;
        padding: 15px;
        background: #f0fdf4;
        border: 2px solid #10b981;
        border-radius: 8px;
        display: none;
      }
      
      .result.show {
        display: block;
      }
      
      .loading {
        text-align: center;
        padding: 20px;
      }
      
      .spinner {
        border: 3px solid #f3f3f3;
        border-top: 3px solid #667eea;
        border-radius: 50%;
        width: 40px;
        height: 40px;
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
      <h2>🤖 Generate ${task === 'description' ? 'Description' : task} with ${model === 'chatgpt' ? 'ChatGPT' : 'Gemini'}</h2>
      
      <div class="selected-data">
        <strong>Selected Data:</strong>
        <div id="selectedData">Loading...</div>
      </div>
      
      <div class="form-group">
        <label>🎯 Target Audience</label>
        <div class="target-fields">
          <input type="text" id="targetAge" placeholder="Age (e.g., 25-45)">
          <select id="targetGender">
            <option value="">All Genders</option>
            <option value="Male">Male</option>
            <option value="Female">Female</option>
            <option value="Non-binary">Non-binary</option>
          </select>
          <input type="text" id="targetLocation" placeholder="Location (e.g., USA)">
        </div>
      </div>
      
      <div class="form-group">
        <label>✍️ Writing Style</label>
        <select id="style">
          <option value="professional">Professional & Informative</option>
          <option value="casual">Casual & Friendly</option>
          <option value="luxury">Luxury & Exclusive</option>
          <option value="storytelling">Storytelling & Emotional</option>
          <option value="investment">Investment Focused</option>
        </select>
      </div>
      
      <div class="example-prompts">
        <h4>💡 Example Prompts (Click to use)</h4>
        <div class="example-item" onclick="useExample('modern')">
          "Focus on modern aesthetics and urban appeal"
        </div>
        <div class="example-item" onclick="useExample('investment')">
          "Emphasize investment potential and artist reputation"
        </div>
        <div class="example-item" onclick="useExample('emotional')">
          "Create emotional connection with the artwork's story"
        </div>
        <div class="example-item" onclick="useExample('technical')">
          "Highlight technical mastery and artistic techniques"
        </div>
      </div>
      
      <div class="form-group">
        <label>📝 Custom Instructions (Optional)</label>
        <textarea id="customPrompt" rows="3" placeholder="Add any specific instructions for the AI..."></textarea>
      </div>
      
      <div class="form-group">
        <label>📍 Output Cell (Optional)</label>
        <input type="text" id="targetCell" placeholder="e.g., D2">
      </div>
      
      <div class="button-group">
        <button onclick="generateContent()">✨ Generate</button>
        <button onclick="google.script.host.close()" style="background: #f1f5f9; color: #475569;">Cancel</button>
      </div>
      
      <div id="result" class="result">
        <div id="resultContent"></div>
      </div>
    </div>
    
    <script>
      // Load selected data
      google.script.run.withSuccessHandler(function(data) {
        document.getElementById('selectedData').innerHTML = data || 'No data selected';
      }).getSelectedDataPreview();
      
      function useExample(type) {
        const prompts = {
          modern: "Focus on modern aesthetics and urban appeal. Mention how this piece fits in contemporary spaces.",
          investment: "Emphasize investment potential, artist's market trajectory, and comparable sales data.",
          emotional: "Tell the story behind the artwork, create emotional connection with potential buyers.",
          technical: "Highlight the technical mastery, unique techniques used, and artistic innovation."
        };
        document.getElementById('customPrompt').value = prompts[type];
      }
      
      function generateContent() {
        const params = {
          model: '${model}',
          task: '${task}',
          targetAge: document.getElementById('targetAge').value,
          targetGender: document.getElementById('targetGender').value,
          targetLocation: document.getElementById('targetLocation').value,
          style: document.getElementById('style').value,
          customPrompt: document.getElementById('customPrompt').value,
          targetCell: document.getElementById('targetCell').value
        };
        
        const resultDiv = document.getElementById('result');
        const resultContent = document.getElementById('resultContent');
        
        resultDiv.className = 'result show';
        resultContent.innerHTML = '<div class="loading"><div class="spinner"></div><p>Generating content...</p></div>';
        
        google.script.run
          .withSuccessHandler(function(result) {
            resultContent.innerHTML = '<h4>✅ Generated Content:</h4><div style="margin-top: 10px; white-space: pre-wrap;">' + result + '</div>';
          })
          .withFailureHandler(function(error) {
            resultContent.innerHTML = '<h4>❌ Error:</h4><p>' + error + '</p>';
          })
          .processAIRequest(params);
      }
    </script>
  </body>
</html>`;
}

/**
 * Get AI examples HTML
 * @return {string} HTML content
 */
function getAIExamplesHTML() {
  return `<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      body {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        padding: 20px;
        background: #f8f9fa;
      }
      
      .container {
        max-width: 750px;
        margin: 0 auto;
        background: white;
        padding: 30px;
        border-radius: 15px;
        box-shadow: 0 5px 20px rgba(0,0,0,0.1);
      }
      
      h2 {
        color: #667eea;
        margin-bottom: 20px;
      }
      
      .example-section {
        margin-bottom: 30px;
        padding: 20px;
        background: #f8f9fa;
        border-radius: 10px;
      }
      
      .example-section h3 {
        color: #1a1a2e;
        margin-bottom: 15px;
      }
      
      .example-code {
        background: #1a1a2e;
        color: #00ff88;
        padding: 15px;
        border-radius: 8px;
        font-family: monospace;
        font-size: 13px;
        margin: 10px 0;
        overflow-x: auto;
      }
      
      .example-result {
        background: white;
        border: 2px solid #e2e8f0;
        padding: 15px;
        border-radius: 8px;
        margin: 10px 0;
      }
      
      .tip {
        background: #fef3c7;
        border-left: 4px solid #f59e0b;
        padding: 12px;
        margin: 15px 0;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <h2>📚 AI Prompt Examples & Best Practices</h2>
      
      <div class="example-section">
        <h3>🎨 Description Generation</h3>
        
        <strong>Setup:</strong>
        <p>Select cells containing: Title (A2), Artist (B2), Medium (E2), Size (F2)</p>
        
        <strong>Target Audience:</strong>
        <p>Age: 35-55, Gender: All, Location: USA</p>
        
        <div class="example-code">
Custom Prompt: "Emphasize the investment potential and limited availability. 
Mention this is perfect for modern home offices or living spaces."
        </div>
        
        <strong>Result Example:</strong>
        <div class="example-result">
          <p><strong>Elevate Your Space with Banksy's Iconic Street Art</strong></p>
          <p>This remarkable limited edition print captures the rebellious spirit and social commentary that has made Banksy one of the most sought-after contemporary artists of our time...</p>
        </div>
      </div>
      
      <div class="example-section">
        <h3>🏷️ Tag Generation</h3>
        
        <strong>Best Practices:</strong>
        <ul>
          <li>Select cells with: Title, Artist, Category, Medium</li>
          <li>Include target demographic info for trending tags</li>
          <li>Specify if you want seasonal or event-specific tags</li>
        </ul>
        
        <div class="example-code">
Custom Prompt: "Include Instagram-friendly hashtags and investment-related keywords"
        </div>
        
        <strong>Result Example:</strong>
        <div class="example-result">
          banksy, street art, urban art, contemporary, limited edition, investment art, 
          modern art, graffiti art, collectible, wall art, home decor, #banksyart, 
          #streetartcollector, #artinvestment
        </div>
      </div>
      
      <div class="example-section">
        <h3>💰 Pricing Suggestions</h3>
        
        <strong>Required Data:</strong>
        <ul>
          <li>Artist name and reputation level</li>
          <li>Artwork size and medium</li>
          <li>Edition size (if applicable)</li>
          <li>Condition and provenance</li>
        </ul>
        
        <div class="tip">
          💡 <strong>Tip:</strong> Include recent sale prices of similar works in your selected cells for more accurate pricing.
        </div>
      </div>
      
      <div class="example-section">
        <h3>✍️ Title Optimization</h3>
        
        <strong>Formula for Success:</strong>
        <div class="example-code">
[Artist] - [Artwork Title] - [Key Feature] - [Size/Medium] - [Unique Selling Point]

Example: "Banksy - Girl With Balloon - Signed Limited Edition - Screen Print - COA Included"
        </div>
        
        <strong>Character Limit:</strong> 80 characters max for eBay
      </div>
      
      <div class="example-section">
        <h3>🎯 Advanced Techniques</h3>
        
        <strong>Multi-Cell References:</strong>
        <p>You can select non-contiguous cells by holding Ctrl (Cmd on Mac)</p>
        
        <strong>Batch Processing:</strong>
        <p>Select an entire column to generate content for multiple items</p>
        
        <strong>Custom Formulas:</strong>
        <div class="example-code">
=CONCATENATE("Generate description for ", A2, " by ", B2, " targeting ", X2, " year old ", Y2, " in ", Z2)
        </div>
      </div>
      
      <div class="tip">
        🚀 <strong>Pro Tip:</strong> Save successful prompts in a separate sheet for reuse. 
        Create a "Prompt Library" tab with your best-performing prompts categorized by use case.
      </div>
    </div>
  </body>
</html>`;
}

/**
 * Get selected data preview
 * @return {string} Preview text
 */
function getSelectedDataPreview() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    
    if (!range) return 'No data selected';
    
    const values = range.getValues();
    let preview = '';
    
    values.forEach((row, i) => {
      row.forEach((cell, j) => {
        if (cell) {
          const cellRef = sheet.getRange(range.getRow() + i, range.getColumn() + j).getA1Notation();
          preview += `${cellRef}: ${cell}\n`;
        }
      });
    });
    
    return preview || 'Empty cells selected';
    
  } catch (error) {
    return 'Unable to read selected data';
  }
}

/**
 * Log AI usage for tracking
 * @param {string} model - AI model used
 * @param {string} task - Task performed
 * @param {number} promptLength - Prompt character count
 * @param {number} responseLength - Response character count
 */
function logAIUsage(model, task, promptLength, responseLength) {
  try {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AI Usage Log') 
                     || SpreadsheetApp.getActiveSpreadsheet().insertSheet('AI Usage Log');
    
    // Add headers if first time
    if (logSheet.getLastRow() === 0) {
      logSheet.getRange(1, 1, 1, 7).setValues([
        ['Timestamp', 'Model', 'Task', 'Prompt Length', 'Response Length', 'User', 'Cost Estimate']
      ]);
      logSheet.getRange(1, 1, 1, 7)
        .setBackground('#667eea')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
    
    // Estimate cost (rough estimates)
    let costEstimate = 0;
    if (model === 'chatgpt') {
      costEstimate = (promptLength / 1000 * 0.01) + (responseLength / 1000 * 0.03); // GPT-4 pricing
    } else if (model === 'gemini') {
      costEstimate = 0; // Gemini is free for now
    }
    
    const logEntry = [
      new Date(),
      model,
      task,
      promptLength,
      responseLength,
      Session.getActiveUser().getEmail(),
      '$' + costEstimate.toFixed(4)
    ];
    
    logSheet.appendRow(logEntry);
    
  } catch (error) {
    console.error('Failed to log AI usage:', error);
  }
}

// ==================== [INCLUDE ALL PREVIOUS IMPORT FUNCTIONS HERE] ====================
// Include all the import, folder creation, and other functions from the previous script
// (importFromClipboard, createImageFoldersForListings, formatSheet, etc.)
// ... [Previous functions continue here] ...