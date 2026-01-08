/**
 * NEWS ENGINE - FEEDLY & EMAIL TO SLIDE DECK
 *
 * Two input methods:
 * A) Save article to Feedly Board → Auto-generates slide deck
 * B) Forward email with article → Auto-generates slide deck
 *
 * ADD TO YOUR EXISTING GOOGLE APPS SCRIPT PROJECT
 */

// ═══════════════════════════════════════════════════════════════
// CONFIGURATION - UPDATE THESE VALUES
// ═══════════════════════════════════════════════════════════════

const FEEDLY_CONFIG = {
  // Get your token from: https://feedly.com/v3/auth/dev
  API_TOKEN: 'YOUR_FEEDLY_API_TOKEN',

  // Board to watch for new articles (the board ID from Feedly)
  // Find it: Open board in Feedly, look at URL for the board ID
  WATCH_BOARD_ID: 'YOUR_FEEDLY_BOARD_ID',

  // Or use board name (we'll look it up)
  WATCH_BOARD_NAME: 'TO POST',

  // How often to check (in minutes)
  POLL_INTERVAL: 5
};

const EMAIL_CONFIG = {
  // Gmail label to watch for forwarded articles
  WATCH_LABEL: 'NEWS-TO-POST',

  // Subject line trigger (articles forwarded with this in subject)
  SUBJECT_TRIGGER: '[POST]'
};

const SLIDES_CONFIG = {
  // Template slide deck ID (make a copy of your template)
  TEMPLATE_ID: 'YOUR_GOOGLE_SLIDES_TEMPLATE_ID',

  // Output folder for generated decks
  OUTPUT_FOLDER_ID: 'YOUR_GOOGLE_DRIVE_FOLDER_ID',

  // Slide placeholders (update to match your template)
  PLACEHOLDERS: {
    HEADLINE: '{{HEADLINE}}',
    SUMMARY: '{{SUMMARY}}',
    SO_WHAT: '{{SO_WHAT}}',
    MY_TAKE: '{{MY_TAKE}}',
    WHATS_NEXT: '{{WHATS_NEXT}}',
    SOURCE: '{{SOURCE}}',
    DATE: '{{DATE}}',
    TOPIC: '{{TOPIC}}'
  }
};

const NOTIFICATION_CONFIG = {
  // Email to send completed deck notification
  NOTIFY_EMAIL: 'YOUR_EMAIL@gmail.com',

  // Also add to NEWS IN sheet?
  ADD_TO_SHEET: true
};


// ═══════════════════════════════════════════════════════════════
// FEEDLY INTEGRATION
// ═══════════════════════════════════════════════════════════════

/**
 * Check Feedly board for new articles
 */
function checkFeedlyForNewArticles() {
  try {
    const processedIds = getProcessedArticleIds();
    const articles = fetchFeedlyBoardArticles();

    let newCount = 0;
    articles.forEach(article => {
      if (!processedIds.includes(article.id)) {
        processArticle(article, 'Feedly');
        markArticleAsProcessed(article.id);
        newCount++;
      }
    });

    if (newCount > 0) {
      Logger.log(`Processed ${newCount} new articles from Feedly`);
    }

    return newCount;
  } catch (error) {
    logError('checkFeedlyForNewArticles', error);
    return 0;
  }
}

/**
 * Fetch articles from Feedly board
 */
function fetchFeedlyBoardArticles() {
  const boardId = FEEDLY_CONFIG.WATCH_BOARD_ID || getBoardIdByName(FEEDLY_CONFIG.WATCH_BOARD_NAME);

  const url = `https://cloud.feedly.com/v3/streams/contents?streamId=${encodeURIComponent(boardId)}&count=20`;

  const options = {
    method: 'GET',
    headers: {
      'Authorization': `OAuth ${FEEDLY_CONFIG.API_TOKEN}`
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  if (data.items) {
    return data.items.map(item => ({
      id: item.id,
      title: item.title || 'Untitled',
      url: item.alternate?.[0]?.href || item.originId,
      summary: item.summary?.content || item.content?.content || '',
      source: item.origin?.title || 'Unknown',
      published: new Date(item.published),
      topics: item.keywords || []
    }));
  }

  return [];
}

/**
 * Get Feedly board ID by name
 */
function getBoardIdByName(boardName) {
  const url = 'https://cloud.feedly.com/v3/collections';

  const options = {
    method: 'GET',
    headers: {
      'Authorization': `OAuth ${FEEDLY_CONFIG.API_TOKEN}`
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const collections = JSON.parse(response.getContentText());

  const board = collections.find(c => c.label.toLowerCase() === boardName.toLowerCase());

  if (board) {
    return board.id;
  }

  throw new Error(`Board "${boardName}" not found in Feedly`);
}


// ═══════════════════════════════════════════════════════════════
// EMAIL INTEGRATION
// ═══════════════════════════════════════════════════════════════

/**
 * Check Gmail for forwarded articles
 */
function checkEmailForNewArticles() {
  try {
    const processedIds = getProcessedArticleIds();
    let newCount = 0;

    // Search for emails with trigger subject or label
    const query = EMAIL_CONFIG.WATCH_LABEL
      ? `label:${EMAIL_CONFIG.WATCH_LABEL} is:unread`
      : `subject:${EMAIL_CONFIG.SUBJECT_TRIGGER} is:unread`;

    const threads = GmailApp.search(query, 0, 10);

    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        const messageId = message.getId();

        if (!processedIds.includes(messageId)) {
          const article = extractArticleFromEmail(message);
          if (article) {
            processArticle(article, 'Email');
            markArticleAsProcessed(messageId);
            newCount++;
          }
          message.markRead();
        }
      });
    });

    if (newCount > 0) {
      Logger.log(`Processed ${newCount} new articles from Email`);
    }

    return newCount;
  } catch (error) {
    logError('checkEmailForNewArticles', error);
    return 0;
  }
}

/**
 * Extract article info from forwarded email
 */
function extractArticleFromEmail(message) {
  const subject = message.getSubject();
  const body = message.getPlainBody();
  const htmlBody = message.getBody();

  // Extract URL from email body
  const urlMatch = body.match(/https?:\/\/[^\s<>"{}|\\^`\[\]]+/);
  const url = urlMatch ? urlMatch[0] : null;

  if (!url) {
    Logger.log('No URL found in email: ' + subject);
    return null;
  }

  // Clean up subject (remove Fwd:, Re:, [POST], etc.)
  let title = subject
    .replace(/^(Fwd?:|Re:|Fw:|\[POST\])\s*/gi, '')
    .trim();

  // If title is empty, try to fetch from URL
  if (!title || title.length < 5) {
    title = fetchTitleFromUrl(url) || 'Article';
  }

  return {
    id: message.getId(),
    title: title,
    url: url,
    summary: extractSummaryFromBody(body),
    source: extractSourceFromUrl(url),
    published: message.getDate(),
    topics: []
  };
}

/**
 * Extract summary from email body
 */
function extractSummaryFromBody(body) {
  // Remove forwarded email headers
  const cleaned = body
    .replace(/---------- Forwarded message ---------[\s\S]*?Subject:.*?\n/gi, '')
    .replace(/From:.*?\n/gi, '')
    .replace(/Date:.*?\n/gi, '')
    .replace(/To:.*?\n/gi, '')
    .trim();

  // Take first 500 chars as summary
  return cleaned.substring(0, 500);
}

/**
 * Fetch page title from URL
 */
function fetchTitleFromUrl(url) {
  try {
    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true, followRedirects: true});
    const html = response.getContentText();
    const titleMatch = html.match(/<title[^>]*>([^<]+)<\/title>/i);
    return titleMatch ? titleMatch[1].trim() : null;
  } catch (e) {
    return null;
  }
}

/**
 * Extract source domain from URL
 */
function extractSourceFromUrl(url) {
  try {
    const match = url.match(/https?:\/\/(?:www\.)?([^\/]+)/);
    return match ? match[1] : 'Unknown';
  } catch (e) {
    return 'Unknown';
  }
}


// ═══════════════════════════════════════════════════════════════
// ARTICLE PROCESSING & CONTENT GENERATION
// ═══════════════════════════════════════════════════════════════

/**
 * Process a new article - generate content and create slide deck
 */
function processArticle(article, source) {
  Logger.log(`Processing article: ${article.title} (from ${source})`);

  // 1. Fetch full article content if needed
  const fullContent = fetchArticleContent(article.url);

  // 2. Generate AI content (So What, My Take, What's Next)
  const generatedContent = generateArticleContent(article.title, fullContent || article.summary);

  // 3. Detect topic
  const topic = detectTopic(article.title, fullContent);

  // 4. Create slide deck
  const deckUrl = createSlideDeck({
    headline: article.title,
    summary: generatedContent.summary || article.summary,
    soWhat: generatedContent.soWhat,
    myTake: generatedContent.myTake,
    whatsNext: generatedContent.whatsNext,
    source: article.source,
    url: article.url,
    date: article.published,
    topic: topic
  });

  // 5. Add to NEWS IN sheet if configured
  if (NOTIFICATION_CONFIG.ADD_TO_SHEET) {
    addToNewsSheet(article, generatedContent, topic, deckUrl);
  }

  // 6. Send notification
  sendNotification(article, deckUrl, source);

  return deckUrl;
}

/**
 * Fetch full article content from URL
 */
function fetchArticleContent(url) {
  try {
    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true, followRedirects: true});
    const html = response.getContentText();

    // Basic content extraction (remove scripts, styles, nav, etc.)
    let content = html
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
      .replace(/<nav[^>]*>[\s\S]*?<\/nav>/gi, '')
      .replace(/<header[^>]*>[\s\S]*?<\/header>/gi, '')
      .replace(/<footer[^>]*>[\s\S]*?<\/footer>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    // Limit to first 3000 chars
    return content.substring(0, 3000);
  } catch (e) {
    Logger.log('Error fetching article: ' + e.message);
    return null;
  }
}

/**
 * Generate AI content for article
 */
function generateArticleContent(title, content) {
  // Check if we have an AI service configured
  const aiService = getAIService();

  if (!aiService) {
    // Return placeholder content if no AI configured
    return {
      summary: content ? content.substring(0, 300) : '',
      soWhat: 'Why this matters for business leaders...',
      myTake: 'My perspective on this development...',
      whatsNext: 'What to watch for next...'
    };
  }

  try {
    const prompt = `Analyze this article and provide LinkedIn post content:

Title: ${title}

Content: ${content}

Provide:
1. SUMMARY (2-3 sentences, factual)
2. SO_WHAT (Why this matters for business/tech leaders - 2 sentences)
3. MY_TAKE (A thoughtful, slightly contrarian perspective - 2 sentences)
4. WHATS_NEXT (What to watch for, predictions - 1-2 sentences)

Format as JSON: {"summary": "...", "soWhat": "...", "myTake": "...", "whatsNext": "..."}`;

    const response = callAIService(aiService, prompt);
    return JSON.parse(response);
  } catch (e) {
    Logger.log('AI generation error: ' + e.message);
    return {
      summary: content ? content.substring(0, 300) : '',
      soWhat: '',
      myTake: '',
      whatsNext: ''
    };
  }
}

/**
 * Detect topic from article content
 */
function detectTopic(title, content) {
  const text = (title + ' ' + content).toLowerCase();

  const topicKeywords = {
    'AI': ['ai', 'artificial intelligence', 'machine learning', 'chatgpt', 'llm', 'openai', 'claude', 'gemini'],
    'Tech': ['software', 'startup', 'tech', 'silicon valley', 'app', 'platform'],
    'Business': ['ceo', 'company', 'revenue', 'acquisition', 'ipo', 'market'],
    'Finance': ['stock', 'investment', 'crypto', 'bitcoin', 'banking', 'fed'],
    'Leadership': ['leadership', 'management', 'culture', 'team', 'hiring'],
    'Product': ['product', 'launch', 'feature', 'user', 'design']
  };

  let bestTopic = 'General';
  let bestScore = 0;

  for (const [topic, keywords] of Object.entries(topicKeywords)) {
    const score = keywords.filter(kw => text.includes(kw)).length;
    if (score > bestScore) {
      bestScore = score;
      bestTopic = topic;
    }
  }

  return bestTopic;
}


// ═══════════════════════════════════════════════════════════════
// SLIDE DECK GENERATION
// ═══════════════════════════════════════════════════════════════

/**
 * Create slide deck from article content
 */
function createSlideDeck(content) {
  try {
    // Copy template
    const template = DriveApp.getFileById(SLIDES_CONFIG.TEMPLATE_ID);
    const folder = DriveApp.getFolderById(SLIDES_CONFIG.OUTPUT_FOLDER_ID);

    const deckName = `NEWS - ${content.topic} - ${formatDate(content.date)}`;
    const copy = template.makeCopy(deckName, folder);

    // Open the presentation
    const presentation = SlidesApp.openById(copy.getId());
    const slides = presentation.getSlides();

    // Replace placeholders in all slides
    slides.forEach(slide => {
      const shapes = slide.getShapes();
      shapes.forEach(shape => {
        if (shape.getText) {
          const textRange = shape.getText();
          const text = textRange.asString();

          // Replace each placeholder
          if (text.includes(SLIDES_CONFIG.PLACEHOLDERS.HEADLINE)) {
            textRange.replaceAllText(SLIDES_CONFIG.PLACEHOLDERS.HEADLINE, content.headline);
          }
          if (text.includes(SLIDES_CONFIG.PLACEHOLDERS.SUMMARY)) {
            textRange.replaceAllText(SLIDES_CONFIG.PLACEHOLDERS.SUMMARY, content.summary || '');
          }
          if (text.includes(SLIDES_CONFIG.PLACEHOLDERS.SO_WHAT)) {
            textRange.replaceAllText(SLIDES_CONFIG.PLACEHOLDERS.SO_WHAT, content.soWhat || '');
          }
          if (text.includes(SLIDES_CONFIG.PLACEHOLDERS.MY_TAKE)) {
            textRange.replaceAllText(SLIDES_CONFIG.PLACEHOLDERS.MY_TAKE, content.myTake || '');
          }
          if (text.includes(SLIDES_CONFIG.PLACEHOLDERS.WHATS_NEXT)) {
            textRange.replaceAllText(SLIDES_CONFIG.PLACEHOLDERS.WHATS_NEXT, content.whatsNext || '');
          }
          if (text.includes(SLIDES_CONFIG.PLACEHOLDERS.SOURCE)) {
            textRange.replaceAllText(SLIDES_CONFIG.PLACEHOLDERS.SOURCE, content.source || '');
          }
          if (text.includes(SLIDES_CONFIG.PLACEHOLDERS.DATE)) {
            textRange.replaceAllText(SLIDES_CONFIG.PLACEHOLDERS.DATE, formatDate(content.date));
          }
          if (text.includes(SLIDES_CONFIG.PLACEHOLDERS.TOPIC)) {
            textRange.replaceAllText(SLIDES_CONFIG.PLACEHOLDERS.TOPIC, content.topic || '');
          }
        }
      });
    });

    presentation.saveAndClose();

    const deckUrl = `https://docs.google.com/presentation/d/${copy.getId()}/edit`;
    Logger.log('Created slide deck: ' + deckUrl);

    return deckUrl;

  } catch (error) {
    logError('createSlideDeck', error);
    return null;
  }
}

/**
 * Format date for display
 */
function formatDate(date) {
  if (!date) return new Date().toLocaleDateString();
  const d = new Date(date);
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}


// ═══════════════════════════════════════════════════════════════
// NOTIFICATIONS & SHEET UPDATES
// ═══════════════════════════════════════════════════════════════

/**
 * Send notification email with deck link
 */
function sendNotification(article, deckUrl, source) {
  const subject = `[NEWS DECK READY] ${article.title}`;

  const body = `
New article processed from ${source}!

TITLE: ${article.title}
SOURCE: ${article.source}
URL: ${article.url}

SLIDE DECK: ${deckUrl}

---
Ready to post to LinkedIn!
  `.trim();

  GmailApp.sendEmail(NOTIFICATION_CONFIG.NOTIFY_EMAIL, subject, body);
}

/**
 * Add article to NEWS IN sheet
 */
function addToNewsSheet(article, generatedContent, topic, deckUrl) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);

    if (!sheet) {
      Logger.log('NEWS IN sheet not found');
      return;
    }

    // Add new row
    sheet.appendRow([
      topic,                    // Topic
      '',                       // Select
      '',                       // Score (will be calculated)
      article.published,        // Date
      article.title,            // Title
      article.url,              // Link
      generatedContent.summary, // Summary
      '',                       // AI 1
      '',                       // AI 2
      '',                       // AI 3
      '',                       // AI 4
      '',                       // AI 5
      '',                       // AI 6
      '',                       // AI Avg
      generatedContent.soWhat,  // So What
      generatedContent.myTake,  // My Take
      generatedContent.whatsNext, // What's Next
      deckUrl                   // Deck URL
    ]);

    Logger.log('Added article to NEWS IN sheet');
  } catch (error) {
    logError('addToNewsSheet', error);
  }
}


// ═══════════════════════════════════════════════════════════════
// TRACKING & UTILITIES
// ═══════════════════════════════════════════════════════════════

/**
 * Get list of already-processed article IDs
 */
function getProcessedArticleIds() {
  const props = PropertiesService.getScriptProperties();
  const idsJson = props.getProperty('PROCESSED_ARTICLE_IDS') || '[]';
  return JSON.parse(idsJson);
}

/**
 * Mark article as processed
 */
function markArticleAsProcessed(articleId) {
  const props = PropertiesService.getScriptProperties();
  const ids = getProcessedArticleIds();

  ids.push(articleId);

  // Keep only last 500 IDs
  if (ids.length > 500) {
    ids.splice(0, ids.length - 500);
  }

  props.setProperty('PROCESSED_ARTICLE_IDS', JSON.stringify(ids));
}

/**
 * Get AI service configuration (Gemini, OpenAI, or Claude)
 */
function getAIService() {
  const props = PropertiesService.getScriptProperties();

  // Prefer Gemini (free tier available)
  const geminiKey = props.getProperty('GEMINI_API_KEY');
  if (geminiKey) {
    return { type: 'gemini', key: geminiKey };
  }

  const openaiKey = props.getProperty('OPENAI_API_KEY');
  if (openaiKey) {
    return { type: 'openai', key: openaiKey };
  }

  const claudeKey = props.getProperty('CLAUDE_API_KEY');
  if (claudeKey) {
    return { type: 'claude', key: claudeKey };
  }

  return null;
}

/**
 * Call AI service
 */
function callAIService(service, prompt) {
  if (service.type === 'gemini') {
    return callGemini(service.key, prompt);
  } else if (service.type === 'openai') {
    return callOpenAI(service.key, prompt);
  } else if (service.type === 'claude') {
    return callClaude(service.key, prompt);
  }
  throw new Error('Unknown AI service type');
}

function callOpenAI(apiKey, prompt) {
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      model: 'gpt-4o-mini',
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.7
    })
  });

  const data = JSON.parse(response.getContentText());
  return data.choices[0].message.content;
}

function callClaude(apiKey, prompt) {
  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': apiKey,
      'Content-Type': 'application/json',
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: 'claude-3-haiku-20240307',
      max_tokens: 1024,
      messages: [{ role: 'user', content: prompt }]
    })
  });

  const data = JSON.parse(response.getContentText());
  return data.content[0].text;
}

/**
 * Call Google Gemini API
 */
function callGemini(apiKey, prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

  const response = UrlFetchApp.fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      contents: [{
        parts: [{ text: prompt }]
      }],
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 1024
      }
    }),
    muteHttpExceptions: true
  });

  const data = JSON.parse(response.getContentText());

  if (data.error) {
    throw new Error(data.error.message);
  }

  return data.candidates[0].content.parts[0].text;
}

/**
 * Test Gemini API Key
 */
function testGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();

  // Get key from properties or prompt
  let apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    const response = ui.prompt('Test Gemini API',
      'Enter your Gemini API key to test:',
      ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() !== ui.Button.OK) return;
    apiKey = response.getResponseText().trim();
  }

  try {
    ui.alert('Testing...', 'Calling Gemini API...', ui.ButtonSet.OK);

    const result = callGemini(apiKey, 'Say "Hello! Gemini API is working!" and nothing else.');

    ui.alert('SUCCESS!',
      'Gemini API is working!\n\nResponse: ' + result,
      ui.ButtonSet.OK);

    // Save the key since it works
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKey);

    return true;
  } catch (error) {
    ui.alert('ERROR',
      'Gemini API test failed:\n\n' + error.message +
      '\n\nCheck your API key at: https://aistudio.google.com/app/apikey',
      ui.ButtonSet.OK);
    return false;
  }
}

/**
 * Log errors
 */
function logError(functionName, error) {
  Logger.log(`ERROR in ${functionName}: ${error.message}`);
  console.error(`${functionName}: ${error.message}\n${error.stack}`);
}


// ═══════════════════════════════════════════════════════════════
// TRIGGERS & MENU
// ═══════════════════════════════════════════════════════════════

/**
 * Set up automated triggers
 */
function setupFeedlyEmailTriggers() {
  // Delete existing triggers for these functions
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    const fn = trigger.getHandlerFunction();
    if (fn === 'checkFeedlyForNewArticles' || fn === 'checkEmailForNewArticles') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Feedly check every 5 minutes
  ScriptApp.newTrigger('checkFeedlyForNewArticles')
    .timeBased()
    .everyMinutes(FEEDLY_CONFIG.POLL_INTERVAL)
    .create();

  // Email check every 5 minutes
  ScriptApp.newTrigger('checkEmailForNewArticles')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('Feedly & Email triggers set up successfully');
}

/**
 * Run full ingestion with UI details (the function you were looking for!)
 */
function runFullIngestionWithUIDetails() {
  const ui = SpreadsheetApp.getUi();

  ui.alert('Starting Ingestion',
    'Checking Feedly and Email for new articles...\n\n' +
    'This will:\n' +
    '1. Check your Feedly board\n' +
    '2. Check your Gmail for forwarded articles\n' +
    '3. Generate slide decks for new articles\n' +
    '4. Send you notification emails',
    ui.ButtonSet.OK);

  const feedlyCount = checkFeedlyForNewArticles();
  const emailCount = checkEmailForNewArticles();

  ui.alert('Ingestion Complete',
    `Processed:\n- ${feedlyCount} articles from Feedly\n- ${emailCount} articles from Email\n\n` +
    'Check your email for slide deck links!',
    ui.ButtonSet.OK);
}

/**
 * Manual test - process single URL
 */
function testProcessUrl() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Test Article Processing',
    'Enter an article URL to test:',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    const url = response.getResponseText();

    const article = {
      id: 'test-' + Date.now(),
      title: fetchTitleFromUrl(url) || 'Test Article',
      url: url,
      summary: '',
      source: extractSourceFromUrl(url),
      published: new Date(),
      topics: []
    };

    const deckUrl = processArticle(article, 'Manual Test');

    ui.alert('Test Complete',
      `Slide deck created!\n\n${deckUrl}`,
      ui.ButtonSet.OK);
  }
}

/**
 * Add menu items
 */
function onOpenFeedlyEmail() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('NEWS - Feedly/Email')
    .addItem('Run Full Ingestion', 'runFullIngestionWithUIDetails')
    .addItem('Check Feedly Now', 'checkFeedlyForNewArticles')
    .addItem('Check Email Now', 'checkEmailForNewArticles')
    .addSeparator()
    .addItem('Test Process URL', 'testProcessUrl')
    .addSeparator()
    .addItem('Test Gemini API Key', 'testGeminiApiKey')
    .addItem('Setup Triggers (Auto-check)', 'setupFeedlyEmailTriggers')
    .addItem('Configure API Keys', 'showApiKeySetup')
    .addToUi();
}

/**
 * Show API key setup dialog
 */
function showApiKeySetup() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 15px; font-weight: bold; }
      input { width: 100%; padding: 8px; margin-top: 5px; box-sizing: border-box; }
      button { margin-top: 20px; padding: 10px 20px; background: #4285f4; color: white; border: none; cursor: pointer; }
      .note { font-size: 12px; color: #666; margin-top: 5px; }
      .section { border-top: 1px solid #ddd; margin-top: 20px; padding-top: 10px; }
    </style>

    <h3>API Configuration</h3>

    <label>Feedly API Token</label>
    <input type="text" id="feedly" placeholder="Get from feedly.com/v3/auth/dev">
    <div class="note">Required for Feedly integration</div>

    <div class="section">
      <label>Google Gemini API Key (Recommended - Free!)</label>
      <input type="text" id="gemini" placeholder="Get from aistudio.google.com/app/apikey">
      <div class="note">Free tier available - for AI content generation</div>
    </div>

    <label>OpenAI API Key (optional, if not using Gemini)</label>
    <input type="text" id="openai" placeholder="sk-...">
    <div class="note">Alternative AI provider</div>

    <div class="section">
      <label>Google Slides Template ID</label>
      <input type="text" id="template" placeholder="From your template URL">
      <div class="note">The ID from your slide template URL</div>

      <label>Output Folder ID</label>
      <input type="text" id="folder" placeholder="From your Drive folder URL">
      <div class="note">Where to save generated decks</div>
    </div>

    <button onclick="save()">Save Configuration</button>

    <script>
      function save() {
        google.script.run.withSuccessHandler(() => {
          alert('Configuration saved!');
          google.script.host.close();
        }).saveApiKeys(
          document.getElementById('feedly').value,
          document.getElementById('gemini').value,
          document.getElementById('openai').value,
          document.getElementById('template').value,
          document.getElementById('folder').value
        );
      }
    </script>
  `)
  .setWidth(450)
  .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'API Setup');
}

/**
 * Save API keys to script properties
 */
function saveApiKeys(feedly, gemini, openai, template, folder) {
  const props = PropertiesService.getScriptProperties();

  if (feedly) props.setProperty('FEEDLY_API_TOKEN', feedly);
  if (gemini) props.setProperty('GEMINI_API_KEY', gemini);
  if (openai) props.setProperty('OPENAI_API_KEY', openai);
  if (template) props.setProperty('SLIDES_TEMPLATE_ID', template);
  if (folder) props.setProperty('OUTPUT_FOLDER_ID', folder);

  Logger.log('API keys saved');
}


// ═══════════════════════════════════════════════════════════════
// SLIDES TEMPLATE GENERATOR (Creates template via API)
// ═══════════════════════════════════════════════════════════════

/**
 * Create NEWS slide deck template with placeholders
 * Run this once to generate your template, then use the ID
 */
function createNewsSlideTemplate() {
  const ui = SpreadsheetApp.getUi();

  // Create new presentation
  const presentation = SlidesApp.create('NEWS TEMPLATE - Auto Generated');
  const presentationId = presentation.getId();

  // Get slides and remove default blank slide
  const slides = presentation.getSlides();
  if (slides.length > 0) {
    slides[0].remove();
  }

  // Color scheme
  const COLORS = {
    darkBg: '#1a1a2e',
    accentBlue: '#4361ee',
    accentOrange: '#ff6b35',
    white: '#ffffff',
    lightGray: '#e0e0e0',
    darkGray: '#2d2d44'
  };

  // ─────────────────────────────────────────────
  // SLIDE 1: Title/Headline Slide
  // ─────────────────────────────────────────────
  const slide1 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  // Background
  slide1.getBackground().setSolidFill(COLORS.darkBg);

  // Topic tag (top left)
  const topicBox = slide1.insertTextBox('{{TOPIC}}', 40, 30, 120, 35);
  topicBox.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(COLORS.white);
  topicBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  topicBox.getFill().setSolidFill(COLORS.accentBlue);

  // Date (top right)
  const dateBox = slide1.insertTextBox('{{DATE}}', 580, 35, 120, 25);
  dateBox.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setForegroundColor(COLORS.lightGray);

  // Main Headline
  const headlineBox = slide1.insertTextBox('{{HEADLINE}}', 40, 120, 640, 150);
  headlineBox.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(36)
    .setBold(true)
    .setForegroundColor(COLORS.white);

  // Summary
  const summaryBox = slide1.insertTextBox('{{SUMMARY}}', 40, 290, 640, 120);
  summaryBox.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(18)
    .setForegroundColor(COLORS.lightGray);

  // Source attribution (bottom)
  const sourceBox = slide1.insertTextBox('Source: {{SOURCE}}', 40, 480, 400, 25);
  sourceBox.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setItalic(true)
    .setForegroundColor(COLORS.lightGray);

  // ─────────────────────────────────────────────
  // SLIDE 2: So What? Slide
  // ─────────────────────────────────────────────
  const slide2 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide2.getBackground().setSolidFill(COLORS.darkBg);

  // Section header
  const soWhatHeader = slide2.insertTextBox('SO WHAT?', 40, 40, 200, 50);
  soWhatHeader.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(28)
    .setBold(true)
    .setForegroundColor(COLORS.accentOrange);

  // Subhead
  const soWhatSub = slide2.insertTextBox('Why This Matters', 40, 95, 300, 30);
  soWhatSub.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(14)
    .setForegroundColor(COLORS.lightGray);

  // Content
  const soWhatContent = slide2.insertTextBox('{{SO_WHAT}}', 40, 150, 640, 300);
  soWhatContent.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(22)
    .setForegroundColor(COLORS.white);

  // ─────────────────────────────────────────────
  // SLIDE 3: My Take Slide
  // ─────────────────────────────────────────────
  const slide3 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide3.getBackground().setSolidFill(COLORS.darkBg);

  // Section header
  const myTakeHeader = slide3.insertTextBox('MY TAKE', 40, 40, 200, 50);
  myTakeHeader.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(28)
    .setBold(true)
    .setForegroundColor(COLORS.accentBlue);

  // Subhead
  const myTakeSub = slide3.insertTextBox('Personal Perspective', 40, 95, 300, 30);
  myTakeSub.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(14)
    .setForegroundColor(COLORS.lightGray);

  // Content
  const myTakeContent = slide3.insertTextBox('{{MY_TAKE}}', 40, 150, 640, 300);
  myTakeContent.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(22)
    .setForegroundColor(COLORS.white);

  // ─────────────────────────────────────────────
  // SLIDE 4: What's Next Slide
  // ─────────────────────────────────────────────
  const slide4 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide4.getBackground().setSolidFill(COLORS.darkBg);

  // Section header
  const whatsNextHeader = slide4.insertTextBox("WHAT'S NEXT?", 40, 40, 280, 50);
  whatsNextHeader.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(28)
    .setBold(true)
    .setForegroundColor(COLORS.accentOrange);

  // Subhead
  const whatsNextSub = slide4.insertTextBox('Looking Ahead', 40, 95, 300, 30);
  whatsNextSub.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(14)
    .setForegroundColor(COLORS.lightGray);

  // Content
  const whatsNextContent = slide4.insertTextBox('{{WHATS_NEXT}}', 40, 150, 640, 250);
  whatsNextContent.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(22)
    .setForegroundColor(COLORS.white);

  // Call to action
  const ctaBox = slide4.insertTextBox('Read full article: {{SOURCE}}', 40, 450, 640, 30);
  ctaBox.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(14)
    .setForegroundColor(COLORS.accentBlue);

  // Save presentation
  presentation.saveAndClose();

  // Save template ID to properties
  PropertiesService.getScriptProperties().setProperty('SLIDES_TEMPLATE_ID', presentationId);

  const templateUrl = `https://docs.google.com/presentation/d/${presentationId}/edit`;

  ui.alert('Template Created!',
    `Your NEWS slide template has been created!\n\n` +
    `Template ID: ${presentationId}\n\n` +
    `URL: ${templateUrl}\n\n` +
    `The ID has been saved to your script properties.`,
    ui.ButtonSet.OK);

  Logger.log('Template created: ' + templateUrl);
  return presentationId;
}

/**
 * Create template with different theme options
 */
function createNewsSlideTemplateWithOptions() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt('Template Theme',
    'Choose a theme:\n\n' +
    '1 = Dark (default)\n' +
    '2 = Light\n' +
    '3 = Professional Blue\n\n' +
    'Enter number:',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const theme = response.getResponseText().trim();

  let colors;
  switch(theme) {
    case '2':
      colors = {
        bg: '#ffffff',
        text: '#1a1a2e',
        accent1: '#4361ee',
        accent2: '#ff6b35',
        subtle: '#666666'
      };
      break;
    case '3':
      colors = {
        bg: '#0a2463',
        text: '#ffffff',
        accent1: '#3e92cc',
        accent2: '#d8315b',
        subtle: '#a0c4e2'
      };
      break;
    default:
      colors = {
        bg: '#1a1a2e',
        text: '#ffffff',
        accent1: '#4361ee',
        accent2: '#ff6b35',
        subtle: '#e0e0e0'
      };
  }

  // Create with selected theme
  const presentation = SlidesApp.create('NEWS TEMPLATE - Theme ' + theme);
  const presentationId = presentation.getId();

  const slides = presentation.getSlides();
  if (slides.length > 0) slides[0].remove();

  // Slide 1: Headline
  const s1 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s1.getBackground().setSolidFill(colors.bg);

  s1.insertTextBox('{{TOPIC}}', 40, 30, 120, 35).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(14).setBold(true).setForegroundColor(colors.bg);
  s1.getShapes()[0].getFill().setSolidFill(colors.accent1);

  s1.insertTextBox('{{DATE}}', 580, 35, 120, 25).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(12).setForegroundColor(colors.subtle);

  s1.insertTextBox('{{HEADLINE}}', 40, 120, 640, 150).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(36).setBold(true).setForegroundColor(colors.text);

  s1.insertTextBox('{{SUMMARY}}', 40, 290, 640, 120).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(18).setForegroundColor(colors.subtle);

  s1.insertTextBox('Source: {{SOURCE}}', 40, 480, 400, 25).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(11).setItalic(true).setForegroundColor(colors.subtle);

  // Slide 2: So What
  const s2 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s2.getBackground().setSolidFill(colors.bg);
  s2.insertTextBox('SO WHAT?', 40, 40, 200, 50).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(28).setBold(true).setForegroundColor(colors.accent2);
  s2.insertTextBox('Why This Matters', 40, 95, 300, 30).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(14).setForegroundColor(colors.subtle);
  s2.insertTextBox('{{SO_WHAT}}', 40, 150, 640, 300).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(22).setForegroundColor(colors.text);

  // Slide 3: My Take
  const s3 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s3.getBackground().setSolidFill(colors.bg);
  s3.insertTextBox('MY TAKE', 40, 40, 200, 50).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(28).setBold(true).setForegroundColor(colors.accent1);
  s3.insertTextBox('Personal Perspective', 40, 95, 300, 30).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(14).setForegroundColor(colors.subtle);
  s3.insertTextBox('{{MY_TAKE}}', 40, 150, 640, 300).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(22).setForegroundColor(colors.text);

  // Slide 4: What's Next
  const s4 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s4.getBackground().setSolidFill(colors.bg);
  s4.insertTextBox("WHAT'S NEXT?", 40, 40, 280, 50).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(28).setBold(true).setForegroundColor(colors.accent2);
  s4.insertTextBox('Looking Ahead', 40, 95, 300, 30).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(14).setForegroundColor(colors.subtle);
  s4.insertTextBox('{{WHATS_NEXT}}', 40, 150, 640, 250).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(22).setForegroundColor(colors.text);
  s4.insertTextBox('Read full article: {{SOURCE}}', 40, 450, 640, 30).getText().getTextStyle()
    .setFontFamily('Arial').setFontSize(14).setForegroundColor(colors.accent1);

  presentation.saveAndClose();
  PropertiesService.getScriptProperties().setProperty('SLIDES_TEMPLATE_ID', presentationId);

  const templateUrl = `https://docs.google.com/presentation/d/${presentationId}/edit`;
  ui.alert('Template Created!', `URL: ${templateUrl}\n\nID saved to properties.`, ui.ButtonSet.OK);

  return presentationId;
}
