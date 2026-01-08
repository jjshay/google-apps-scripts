/**
 * NEWS ENGINE - AUDIENCE WEIGHTING & TARGETING
 * Auto-applies LinkedIn audience weights to incoming articles
 * Routes articles to optimal segments for distribution
 *
 * ADD TO YOUR EXISTING GOOGLE APPS SCRIPT PROJECT
 */

// ═══════════════════════════════════════════════════════════════
// CONFIGURATION
// ═══════════════════════════════════════════════════════════════

const CONFIG = {
  SPREADSHEET_ID: '',  // YOUR_SPREADSHEET_ID_HERE - Get from Google Sheets URL

  // Sheet names
  SHEETS: {
    NEWS_IN: 'NEWS IN',
    VARIABLES: 'VARIABLES',
    LINKEDIN_SEGMENTS: 'LinkedIn_Segments',
    DISTRIBUTION_QUEUE: 'Distribution_Queue',
    ENGAGEMENT_LOG: 'Engagement_Log',        // Track post performance
    PERFORMANCE_INSIGHTS: 'Performance_Insights', // What works analysis
    JJ_SCORE_MODEL: 'JJ_Score_Model'         // Your custom algorithm
  },

  // Column positions in NEWS IN (adjust if different)
  COLUMNS: {
    TOPIC: 1,
    SELECT: 2,
    SCORE: 3,
    DATE: 4,
    TITLE: 5,
    LINK: 6,
    AI_AVG: 14,
    // New columns we'll add
    INDUSTRY_WEIGHT: 23,
    FUNCTION_WEIGHT: 24,
    SENIORITY_WEIGHT: 25,
    COMBINED_WEIGHT: 26,
    FINAL_SCORE: 27,
    TARGET_SEGMENT: 28,
    REACH_RANK: 29,
    // Engagement tracking columns
    POST_URL: 42,           // LinkedIn post URL
    POSTED_DATE: 43,        // When you posted
    IMPRESSIONS: 44,        // Views
    LIKES: 45,              // Reactions
    COMMENTS: 46,           // Comments
    SHARES: 47,             // Reposts
    ENGAGEMENT_RATE: 48,    // (likes+comments+shares)/impressions
    JJ_SCORE: 49            // Your custom trained score
  }
};

// ═══════════════════════════════════════════════════════════════
// AUDIENCE WEIGHTS (from LinkedIn analysis)
// ═══════════════════════════════════════════════════════════════

const INDUSTRY_WEIGHTS = {
  'computer software': 1.50,
  'financial services': 1.42,
  'internet': 1.24,
  'information technology and services': 1.04,
  'hospital & health care': 0.84,
  'venture capital & private equity': 0.73,
  'management consulting': 0.70,
  'biotechnology': 0.69,
  'staffing and recruiting': 0.69,
  'real estate': 0.66,
  'marketing and advertising': 0.64,
  'default': 0.60
};

const FUNCTION_WEIGHTS = {
  'Executive': 1.06,
  'Engineering/Tech': 0.68,
  'Product': 0.64,
  'Sales': 0.63,
  'Data/AI': 0.61,
  'HR/People': 0.60,
  'Consulting': 0.60,
  'Marketing': 0.59,
  'Finance': 0.58,
  'Design': 0.57,
  'Operations': 0.56,
  'Legal': 0.54,
  'default': 0.55
};

const SENIORITY_WEIGHTS = {
  'C-Suite': 1.39,
  'Manager': 0.95,
  'Senior IC': 0.70,
  'Director': 0.59,
  'VP/SVP': 0.58,
  'Mid-Level': 0.57,
  'Entry': 0.51,
  'default': 0.60
};

// ═══════════════════════════════════════════════════════════════
// CONTENT QUALITY TAGS (3 Positive + 3 Negative)
// ═══════════════════════════════════════════════════════════════

const POSITIVE_TAGS = [
  'Verified',
  'Expert-backed',
  'Actionable'
];

const NEGATIVE_TAGS = [
  'Clickbait',
  'Speculative',
  'Misleading'
];

// All tags combined for LLM prompts
const ALL_TAGS = [...POSITIVE_TAGS, ...NEGATIVE_TAGS];

// LLM Tag columns (each LLM gets a column for their 3-5 tag picks)
const LLM_TAG_COLUMNS = {
  GPT_TAGS: 35,      // Column AI
  CLAUDE_TAGS: 36,   // Column AJ
  GEMINI_TAGS: 37,   // Column AK
  GROK_TAGS: 38,     // Column AL
  PERPLEXITY_TAGS: 39, // Column AM
  CONSENSUS_TAGS: 40,  // Column AN - tags picked by 3+ LLMs
  TAG_SENTIMENT: 41    // Column AO - overall +/-/neutral
};

/**
 * Simple prompt for LLM tagging (use with any AI)
 */
function getTagPrompt(title) {
  return `Tag: ${title}
Pick 1-3 from: Verified, Expert-backed, Actionable, Clickbait, Speculative, Misleading
Tags:`;
}

// ═══════════════════════════════════════════════════════════════
// "MY TAKE" INSIGHT PROMPTS - SO WHAT & WHAT'S NEXT
// ═══════════════════════════════════════════════════════════════

/**
 * Generate "So What" strategic insight prompt
 * Focus: Connecting dots, industry expertise, hidden implications
 */
function getSoWhatPrompt(title, summary, topic) {
  return `You are a senior industry analyst with deep expertise in ${topic || 'technology and business strategy'}.

ARTICLE: "${title}"
${summary ? `SUMMARY: ${summary}` : ''}

Write a "SO WHAT:" insight (2-3 sentences) that:
- Connects this news to broader industry trends others miss
- Reveals the second-order effects and hidden implications
- Shows pattern recognition across markets/sectors
- Demonstrates insider-level understanding of what this REALLY means
- Speaks to executives who need strategic context, not surface-level takes

Tone: Confident, analytical, forward-thinking. No hedging. No obvious observations.
Start with "So What:" then deliver the insight.`;
}

/**
 * Generate "What's Next" strategic outlook prompt
 * Focus: Actionable foresight, strategic positioning, concrete predictions
 */
function getWhatsNextPrompt(title, summary, topic) {
  return `You are a strategic advisor to Fortune 500 executives with expertise in ${topic || 'technology and business strategy'}.

ARTICLE: "${title}"
${summary ? `SUMMARY: ${summary}` : ''}

Write a "WHAT'S NEXT:" outlook (2-3 sentences) that:
- Predicts concrete next moves (who will do what, and why)
- Identifies opportunities others will miss until it's obvious
- Gives actionable positioning advice for smart operators
- Connects to timing - what should happen in weeks/months/quarters
- Shows you understand the chess game, not just the current move

Tone: Prescriptive, specific, high-conviction. Name names when relevant.
Start with "What's Next:" then deliver the strategic outlook.`;
}

/**
 * Combined prompt for full "My Take" section
 */
function getMyTakePrompt(title, summary, topic) {
  return `You are @jjshay - a senior technology and finance strategist known for connecting dots others miss and cutting through hype with sharp, actionable insights.

ARTICLE: "${title}"
${summary ? `SUMMARY: ${summary}` : ''}
TOPIC AREA: ${topic || 'Technology/Business'}

Write your "MY TAKE" with two sections:

**SO WHAT:** (2-3 sentences)
- What does this REALLY mean beyond the headline?
- Connect to broader patterns in tech, finance, or strategy
- Show the second-order effects and hidden implications
- Demonstrate expertise that makes readers think "I didn't see that angle"

**WHAT'S NEXT:** (2-3 sentences)
- Make a specific, high-conviction prediction
- Who moves next? What's the timeline? Who wins/loses?
- Give actionable insight for people positioning themselves
- Be concrete - no vague "time will tell" hedging

Voice: Sharp, confident, slightly contrarian. You've seen these patterns before. No fluff, no hedging, no corporate-speak. Write like you're advising a smart friend who respects your judgment.

Format exactly as:
So What: [your insight]

What's Next: [your prediction]`;
}

/**
 * Quick one-liner insight prompt (for social/comments)
 */
function getQuickTakePrompt(title) {
  return `Sharp one-liner reaction to: "${title}"

Write a single sentence that:
- Shows you understand the real story behind the headline
- Adds value beyond restating the obvious
- Sounds like an insider who's seen this movie before

Be direct. No emojis. No hashtags. Just insight.`;
}

// Topic to audience mapping
const TOPIC_AUDIENCE_MAP = {
  'ai frontier': { industry: 'computer software', function: 'Data/AI', seniority: 'Senior IC', segment: 'TECHNOLOGY' },
  'agents': { industry: 'computer software', function: 'Data/AI', seniority: 'Senior IC', segment: 'AI / EMERGING TECH' },
  'finance ai': { industry: 'financial services', function: 'Finance', seniority: 'Director', segment: 'FINANCE' },
  'fintech': { industry: 'financial services', function: 'Finance', seniority: 'Director', segment: 'FINANCE' },
  'compute/silicon': { industry: 'computer software', function: 'Engineering/Tech', seniority: 'Senior IC', segment: 'TECHNOLOGY' },
  'cloud/infrastructure': { industry: 'computer software', function: 'Engineering/Tech', seniority: 'Director', segment: 'TECHNOLOGY' },
  'robotics': { industry: 'computer software', function: 'Engineering/Tech', seniority: 'Senior IC', segment: 'AI / EMERGING TECH' },
  'biotech/healthcare': { industry: 'hospital & health care', function: 'Executive', seniority: 'Director', segment: 'SCIENCE & RESEARCH' },
  'leadership/strategy': { industry: 'management consulting', function: 'Executive', seniority: 'C-Suite', segment: 'FINANCE' },
  'enterprise ai': { industry: 'computer software', function: 'Data/AI', seniority: 'Director', segment: 'TECHNOLOGY' },
  'consumer tech': { industry: 'internet', function: 'Product', seniority: 'Manager', segment: 'CONSUMER / SMALL BIZ' },
  'legal ai': { industry: 'financial services', function: 'Legal', seniority: 'Director', segment: 'TECHNOLOGY' },
  'policy/governance': { industry: 'management consulting', function: 'Executive', seniority: 'Director', segment: 'FINANCE' },
  'default': { industry: 'computer software', function: 'Executive', seniority: 'Manager', segment: 'TECHNOLOGY' }
};

// Segment audience sizes (with email)
const SEGMENT_SIZES = {
  'TECHNOLOGY': 2040,
  'AI / EMERGING TECH': 1290,
  'FINANCE': 1085,
  'CONSUMER / SMALL BIZ': 573,
  'SCIENCE & RESEARCH': 415
};

// ═══════════════════════════════════════════════════════════════
// MAIN FUNCTIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Auto-apply audience weights to all articles in NEWS IN
 * Run this on a trigger or manually
 */
function applyAudienceWeights() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);
  const data = sheet.getDataRange().getValues();

  // Add headers if not present
  const headers = data[0];
  if (headers[CONFIG.COLUMNS.INDUSTRY_WEIGHT - 1] !== 'Industry Weight') {
    sheet.getRange(1, CONFIG.COLUMNS.INDUSTRY_WEIGHT).setValue('Industry Weight');
    sheet.getRange(1, CONFIG.COLUMNS.FUNCTION_WEIGHT).setValue('Function Weight');
    sheet.getRange(1, CONFIG.COLUMNS.SENIORITY_WEIGHT).setValue('Seniority Weight');
    sheet.getRange(1, CONFIG.COLUMNS.COMBINED_WEIGHT).setValue('Combined Weight');
    sheet.getRange(1, CONFIG.COLUMNS.FINAL_SCORE).setValue('Final Score');
    sheet.getRange(1, CONFIG.COLUMNS.TARGET_SEGMENT).setValue('Target Segment');
    sheet.getRange(1, CONFIG.COLUMNS.REACH_RANK).setValue('Reach Rank');
  }

  const scores = [];

  // Process each row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const topic = String(row[CONFIG.COLUMNS.TOPIC - 1] || '').toLowerCase().trim();
    const aiAvg = parseFloat(row[CONFIG.COLUMNS.AI_AVG - 1]) || 0;

    if (!topic || aiAvg === 0) continue;

    // Get audience profile
    const profile = TOPIC_AUDIENCE_MAP[topic] || TOPIC_AUDIENCE_MAP['default'];

    // Calculate weights
    const indWeight = INDUSTRY_WEIGHTS[profile.industry] || INDUSTRY_WEIGHTS['default'];
    const funcWeight = FUNCTION_WEIGHTS[profile.function] || FUNCTION_WEIGHTS['default'];
    const senWeight = SENIORITY_WEIGHTS[profile.seniority] || SENIORITY_WEIGHTS['default'];

    // Combined weight (geometric mean)
    const combinedWeight = Math.pow(indWeight * funcWeight * senWeight, 1/3);

    // Final score
    const finalScore = aiAvg * combinedWeight;

    // Write to sheet
    const rowNum = i + 1;
    sheet.getRange(rowNum, CONFIG.COLUMNS.INDUSTRY_WEIGHT).setValue(Math.round(indWeight * 100) / 100);
    sheet.getRange(rowNum, CONFIG.COLUMNS.FUNCTION_WEIGHT).setValue(Math.round(funcWeight * 100) / 100);
    sheet.getRange(rowNum, CONFIG.COLUMNS.SENIORITY_WEIGHT).setValue(Math.round(senWeight * 100) / 100);
    sheet.getRange(rowNum, CONFIG.COLUMNS.COMBINED_WEIGHT).setValue(Math.round(combinedWeight * 100) / 100);
    sheet.getRange(rowNum, CONFIG.COLUMNS.FINAL_SCORE).setValue(Math.round(finalScore * 10) / 10);
    sheet.getRange(rowNum, CONFIG.COLUMNS.TARGET_SEGMENT).setValue(profile.segment);

    scores.push({ row: rowNum, score: finalScore });
  }

  // Calculate ranks
  scores.sort((a, b) => b.score - a.score);
  scores.forEach((item, index) => {
    sheet.getRange(item.row, CONFIG.COLUMNS.REACH_RANK).setValue(index + 1);
  });

  Logger.log(`Applied audience weights to ${scores.length} articles`);
  return scores.length;
}

/**
 * Get top articles for a specific segment
 */
function getTopArticlesForSegment(segmentName, limit = 10) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);
  const data = sheet.getDataRange().getValues();

  const articles = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const segment = row[CONFIG.COLUMNS.TARGET_SEGMENT - 1];
    const finalScore = row[CONFIG.COLUMNS.FINAL_SCORE - 1];
    const title = row[CONFIG.COLUMNS.TITLE - 1];
    const link = row[CONFIG.COLUMNS.LINK - 1];

    if (segment === segmentName && finalScore > 0) {
      articles.push({
        title: title,
        link: link,
        score: finalScore,
        row: i + 1
      });
    }
  }

  // Sort by score and return top N
  articles.sort((a, b) => b.score - a.score);
  return articles.slice(0, limit);
}

/**
 * Create distribution queue - routes articles to segments
 */
function createDistributionQueue() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Create or get Distribution Queue sheet
  let queueSheet = ss.getSheetByName(CONFIG.SHEETS.DISTRIBUTION_QUEUE);
  if (!queueSheet) {
    queueSheet = ss.insertSheet(CONFIG.SHEETS.DISTRIBUTION_QUEUE);
  }
  queueSheet.clear();

  // Headers
  const headers = ['Segment', 'Audience Size', 'Article Title', 'Link', 'Final Score', 'Status', 'Scheduled'];
  queueSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  queueSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  let rowNum = 2;

  // Get top 5 articles for each segment
  for (const [segment, audienceSize] of Object.entries(SEGMENT_SIZES)) {
    const topArticles = getTopArticlesForSegment(segment, 5);

    for (const article of topArticles) {
      queueSheet.getRange(rowNum, 1, 1, 7).setValues([[
        segment,
        audienceSize,
        article.title,
        article.link,
        article.score,
        'PENDING',
        ''
      ]]);
      rowNum++;
    }
  }

  // Auto-resize columns
  queueSheet.autoResizeColumns(1, 7);

  Logger.log(`Created distribution queue with ${rowNum - 2} articles`);
  return rowNum - 2;
}

// ═══════════════════════════════════════════════════════════════
// RSS FEEDS CONFIGURATION
// ═══════════════════════════════════════════════════════════════

const RSS_FEEDS = [
  // Tech News
  { url: 'https://feeds.feedburner.com/TechCrunch/', name: 'TechCrunch', topic: 'AI Frontier' },
  { url: 'https://www.theverge.com/rss/index.xml', name: 'The Verge', topic: 'Consumer Tech' },
  { url: 'https://feeds.arstechnica.com/arstechnica/technology-lab', name: 'Ars Technica', topic: 'Cloud/Infrastructure' },

  // AI/ML
  { url: 'https://www.artificialintelligence-news.com/feed/', name: 'AI News', topic: 'AI Frontier' },
  { url: 'https://venturebeat.com/category/ai/feed/', name: 'VentureBeat AI', topic: 'AI Frontier' },

  // Finance/Fintech
  { url: 'https://www.finextra.com/rss/headlines.aspx', name: 'Finextra', topic: 'Finance AI' },

  // Biotech
  { url: 'https://www.fiercebiotech.com/rss/xml', name: 'FierceBiotech', topic: 'Biotech/Healthcare' },
];

// ═══════════════════════════════════════════════════════════════
// RSS FEED FETCHER
// ═══════════════════════════════════════════════════════════════

/**
 * Fetch articles from all configured RSS feeds
 */
function fetchAllRSSFeeds() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);

  // Get existing links to avoid duplicates
  const existingLinks = new Set();
  const data = sheet.getDataRange().getValues();
  data.forEach(row => {
    if (row[CONFIG.COLUMNS.LINK - 1]) {
      existingLinks.add(String(row[CONFIG.COLUMNS.LINK - 1]).trim());
    }
  });

  let totalAdded = 0;

  for (const feed of RSS_FEEDS) {
    try {
      const articles = fetchRSSFeed(feed.url, feed.name, feed.topic);

      for (const article of articles) {
        // Skip duplicates
        if (existingLinks.has(article.link)) continue;

        // Add to sheet
        const nextRow = sheet.getLastRow() + 1;
        sheet.getRange(nextRow, CONFIG.COLUMNS.TOPIC).setValue(article.topic);
        sheet.getRange(nextRow, CONFIG.COLUMNS.DATE).setValue(article.pubDate);
        sheet.getRange(nextRow, CONFIG.COLUMNS.TITLE).setValue(article.title);
        sheet.getRange(nextRow, CONFIG.COLUMNS.LINK).setValue(article.link);
        sheet.getRange(nextRow, 20).setValue(article.source); // Source column

        existingLinks.add(article.link);
        totalAdded++;
      }

      Logger.log(`Fetched ${articles.length} from ${feed.name}, added new: ${totalAdded}`);
    } catch (error) {
      Logger.log(`Error fetching ${feed.name}: ${error}`);
    }
  }

  Logger.log(`Total new articles from RSS: ${totalAdded}`);
  return totalAdded;
}

/**
 * Fetch single RSS feed
 */
function fetchRSSFeed(url, sourceName, defaultTopic) {
  const articles = [];

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const content = response.getContentText();
    const xml = XmlService.parse(content);
    const root = xml.getRootElement();

    // Handle both RSS and Atom formats
    let items = [];
    const namespace = root.getNamespace();

    if (root.getName() === 'rss') {
      const channel = root.getChild('channel');
      items = channel.getChildren('item');
    } else if (root.getName() === 'feed') {
      items = root.getChildren('entry', namespace);
    }

    for (let i = 0; i < Math.min(items.length, 15); i++) {
      const item = items[i];

      let title, link, pubDate;

      if (root.getName() === 'rss') {
        title = item.getChildText('title');
        link = item.getChildText('link');
        pubDate = item.getChildText('pubDate');
      } else {
        title = item.getChildText('title', namespace);
        const linkEl = item.getChild('link', namespace);
        link = linkEl ? linkEl.getAttribute('href').getValue() : '';
        pubDate = item.getChildText('published', namespace) || item.getChildText('updated', namespace);
      }

      // Auto-categorize based on title
      let topic = defaultTopic;
      const lowerTitle = (title || '').toLowerCase();

      if (lowerTitle.includes('fintech') || lowerTitle.includes('banking') || lowerTitle.includes('payment')) {
        topic = 'Finance AI';
      } else if (lowerTitle.includes('cyber') || lowerTitle.includes('security') || lowerTitle.includes('breach')) {
        topic = 'Cloud/Infrastructure';
      } else if (lowerTitle.includes('biotech') || lowerTitle.includes('pharma') || lowerTitle.includes('drug')) {
        topic = 'Biotech/Healthcare';
      } else if (lowerTitle.includes('robot') || lowerTitle.includes('automat')) {
        topic = 'Robotics';
      } else if (lowerTitle.includes('cloud') || lowerTitle.includes('aws') || lowerTitle.includes('azure')) {
        topic = 'Cloud/Infrastructure';
      } else if (lowerTitle.includes('gpt') || lowerTitle.includes('llm') || lowerTitle.includes('openai') || lowerTitle.includes('anthropic')) {
        topic = 'AI Frontier';
      } else if (lowerTitle.includes('agent') || lowerTitle.includes('autonomous')) {
        topic = 'Agents';
      }

      articles.push({
        title: title,
        link: link,
        pubDate: pubDate ? new Date(pubDate) : new Date(),
        source: sourceName,
        topic: topic
      });
    }
  } catch (error) {
    Logger.log(`RSS parse error for ${url}: ${error}`);
  }

  return articles;
}

// ═══════════════════════════════════════════════════════════════
// GOOGLE NEWS INTEGRATION
// ═══════════════════════════════════════════════════════════════

/**
 * Fetch news from Google News RSS feeds
 * FREE - No API key required!
 */
function fetchGoogleNews(query, maxResults = 20) {
  const encodedQuery = encodeURIComponent(query);
  const rssUrl = `https://news.google.com/rss/search?q=${encodedQuery}&hl=en-US&gl=US&ceid=US:en`;

  try {
    const response = UrlFetchApp.fetch(rssUrl);
    const xml = XmlService.parse(response.getContentText());
    const root = xml.getRootElement();
    const channel = root.getChild('channel');
    const items = channel.getChildren('item');

    const articles = [];

    for (let i = 0; i < Math.min(items.length, maxResults); i++) {
      const item = items[i];
      articles.push({
        title: item.getChildText('title'),
        link: item.getChildText('link'),
        pubDate: item.getChildText('pubDate'),
        source: 'Google News',
        description: item.getChildText('description') || ''
      });
    }

    return articles;
  } catch (error) {
    Logger.log('Error fetching Google News: ' + error);
    return [];
  }
}

/**
 * Fetch news for all topics and add to NEWS IN
 */
function fetchAllGoogleNews() {
  const topics = [
    'artificial intelligence business',
    'fintech startup funding',
    'cybersecurity enterprise',
    'cloud computing AWS Azure',
    'biotechnology pharma',
    'venture capital deals',
    'robotics automation',
    'blockchain crypto enterprise'
  ];

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);

  // Get existing links to avoid duplicates
  const existingLinks = new Set();
  const data = sheet.getDataRange().getValues();
  data.forEach(row => {
    if (row[CONFIG.COLUMNS.LINK - 1]) {
      existingLinks.add(row[CONFIG.COLUMNS.LINK - 1]);
    }
  });

  let addedCount = 0;

  for (const topic of topics) {
    const articles = fetchGoogleNews(topic, 10);

    for (const article of articles) {
      // Skip if already exists
      if (existingLinks.has(article.link)) continue;

      // Determine topic category
      let topicCategory = 'AI Frontier'; // default
      const lowerTitle = article.title.toLowerCase();

      if (lowerTitle.includes('fintech') || lowerTitle.includes('banking')) topicCategory = 'Finance AI';
      else if (lowerTitle.includes('cyber') || lowerTitle.includes('security')) topicCategory = 'Cloud/Infrastructure';
      else if (lowerTitle.includes('biotech') || lowerTitle.includes('pharma')) topicCategory = 'Biotech/Healthcare';
      else if (lowerTitle.includes('robot') || lowerTitle.includes('automat')) topicCategory = 'Robotics';
      else if (lowerTitle.includes('cloud') || lowerTitle.includes('aws') || lowerTitle.includes('azure')) topicCategory = 'Cloud/Infrastructure';
      else if (lowerTitle.includes('venture') || lowerTitle.includes('funding')) topicCategory = 'Finance AI';

      // Add to sheet
      const nextRow = sheet.getLastRow() + 1;
      sheet.getRange(nextRow, CONFIG.COLUMNS.TOPIC).setValue(topicCategory);
      sheet.getRange(nextRow, CONFIG.COLUMNS.DATE).setValue(new Date(article.pubDate));
      sheet.getRange(nextRow, CONFIG.COLUMNS.TITLE).setValue(article.title);
      sheet.getRange(nextRow, CONFIG.COLUMNS.LINK).setValue(article.link);
      sheet.getRange(nextRow, 20).setValue('Google News'); // Source column

      existingLinks.add(article.link);
      addedCount++;
    }
  }

  Logger.log(`Added ${addedCount} new articles from Google News`);
  return addedCount;
}

// ═══════════════════════════════════════════════════════════════
// TRIGGERS & AUTOMATION
// ═══════════════════════════════════════════════════════════════

/**
 * Set up automated triggers
 * Run this once to enable automation
 */
function setupTriggers() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // Fetch Google News every 4 hours
  ScriptApp.newTrigger('fetchAllGoogleNews')
    .timeBased()
    .everyHours(4)
    .create();

  // Apply audience weights every hour
  ScriptApp.newTrigger('applyAudienceWeights')
    .timeBased()
    .everyHours(1)
    .create();

  // Create distribution queue daily at 6 AM
  ScriptApp.newTrigger('createDistributionQueue')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();

  Logger.log('Triggers set up successfully');
}

/**
 * Manual run - do everything
 */
function runFullPipeline() {
  Logger.log('=== STARTING FULL PIPELINE ===');

  const rssAdded = fetchAllRSSFeeds();
  Logger.log(`Step 1: Added ${rssAdded} articles from RSS feeds`);

  const googleAdded = fetchAllGoogleNews();
  Logger.log(`Step 2: Added ${googleAdded} articles from Google News`);

  const weighted = applyAudienceWeights();
  Logger.log(`Step 3: Applied weights to ${weighted} articles`);

  const queued = createDistributionQueue();
  Logger.log(`Step 4: Created queue with ${queued} articles`);

  Logger.log('=== PIPELINE COMPLETE ===');
}

// ═══════════════════════════════════════════════════════════════
// MENU
// ═══════════════════════════════════════════════════════════════

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 Audience Targeting')
    .addItem('Apply Audience Weights', 'applyAudienceWeights')
    .addItem('Calculate Tag Consensus', 'calculateAllTagConsensus')
    .addItem('Create Quality Report', 'createQualityReport')
    .addItem('Fetch RSS Feeds', 'fetchAllRSSFeeds')
    .addItem('Fetch Google News', 'fetchAllGoogleNews')
    .addItem('Create Distribution Queue', 'createDistributionQueue')
    .addItem('Build Weekly Schedule', 'buildWeeklySchedule')
    .addSeparator()
    .addItem('Run Full Pipeline', 'runFullPipeline')
    .addItem('Run Safe Pipeline (with backup)', 'runSafePipeline')
    .addItem('Setup LLM Tag Columns', 'setupLLMTagColumns')
    .addItem('Setup Automation Triggers', 'setupTriggers')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 JJ Score & Engagement')
      .addItem('Log Post Engagement', 'showEngagementInputForm')
      .addItem('Setup Engagement Log', 'setupEngagementLog')
      .addItem('Train JJ Score Model', 'trainJJScoreModel')
      .addItem('Generate Performance Insights', 'generatePerformanceInsights')
      .addItem('View JJ Score Model', 'viewJJScoreModel'))
    .addSubMenu(ui.createMenu('🔗 Bit.ly & Links')
      .addItem('Setup Bit.ly Tracking', 'setupBitlyTracking')
      .addItem('Update All Bit.ly Stats', 'updateAllBitlyStats')
      .addItem('View Bit.ly Dashboard', 'viewBitlyDashboard'))
    .addSubMenu(ui.createMenu('💬 Comments')
      .addItem('Log Comment', 'showCommentInputForm')
      .addItem('Setup Comments Log', 'setupCommentsTracking')
      .addItem('Analyze Comments', 'analyzeComments')
      .addItem('View Comments Analysis', 'viewCommentsAnalysis'))
    .addSubMenu(ui.createMenu('🛡️ Backup & Recovery')
      .addItem('Create Backup Now', 'createManualBackup')
      .addItem('List All Backups', 'showBackupList')
      .addItem('Restore Last Backup', 'restoreLastBackup')
      .addItem('View Backup Index', 'viewBackupIndex'))
    .addSubMenu(ui.createMenu('🔧 Maintenance')
      .addItem('Run Health Check', 'runHealthCheck')
      .addItem('Validate Data', 'showValidationResults')
      .addItem('Cleanup & Optimize', 'cleanupAndOptimize')
      .addItem('View Error Log', 'viewErrorLog')
      .addItem('Clear Error Log', 'clearErrorLog'))
    .addToUi();
}

function viewJJScoreModel() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const modelSheet = ss.getSheetByName(CONFIG.SHEETS.JJ_SCORE_MODEL);
  if (modelSheet) {
    ss.setActiveSheet(modelSheet);
  } else {
    SpreadsheetApp.getUi().alert('No Model', 'JJ Score model not trained yet. Log 10+ posts first, then train the model.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function viewBitlyDashboard() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const bitlySheet = ss.getSheetByName('Bitly_Tracking');
  if (bitlySheet) {
    ss.setActiveSheet(bitlySheet);
  } else {
    SpreadsheetApp.getUi().alert('No Bit.ly Data', 'Bit.ly tracking not set up yet. Run Setup Bit.ly Tracking first.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function viewCommentsAnalysis() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const analysisSheet = ss.getSheetByName('Comments_Analysis');
  if (analysisSheet) {
    ss.setActiveSheet(analysisSheet);
  } else {
    // Try to generate it first
    const result = analyzeComments();
    if (result) {
      ss.setActiveSheet(ss.getSheetByName('Comments_Analysis'));
    } else {
      SpreadsheetApp.getUi().alert('No Comments', 'No comments logged yet. Log some comments first using the Comments sidebar.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

// ═══════════════════════════════════════════════════════════════
// MENU HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════

function createManualBackup() {
  const backupName = createBackup('manual');
  if (backupName) {
    SpreadsheetApp.getUi().alert('Backup Created', `Backup saved as: ${backupName}`, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('Backup Failed', 'Could not create backup. Check Error_Log for details.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function showBackupList() {
  const backups = listBackups();
  if (backups.length === 0) {
    SpreadsheetApp.getUi().alert('No Backups', 'No backups found. Run a backup first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const list = backups.map(b => `• ${b.name} (${b.rows} rows)`).join('\n');
  SpreadsheetApp.getUi().alert('Available Backups', list, SpreadsheetApp.getUi().ButtonSet.OK);
}

function restoreLastBackup() {
  const backups = listBackups();
  if (backups.length === 0) {
    SpreadsheetApp.getUi().alert('No Backups', 'No backups available to restore.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Get most recent backup (last in sorted list)
  const lastBackup = backups.sort((a, b) => b.name.localeCompare(a.name))[0];

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirm Restore',
    `This will restore from: ${lastBackup.name}\n\nCurrent data will be replaced. Continue?`,
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    const success = restoreFromBackup(lastBackup.name);
    if (success) {
      ui.alert('Restore Complete', `Successfully restored from ${lastBackup.name}`, ui.ButtonSet.OK);
    } else {
      ui.alert('Restore Failed', 'Could not restore backup. Check Error_Log for details.', ui.ButtonSet.OK);
    }
  }
}

function viewBackupIndex() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const backupIndex = ss.getSheetByName('Backup_Index');
  if (backupIndex) {
    ss.setActiveSheet(backupIndex);
  } else {
    SpreadsheetApp.getUi().alert('No Backups', 'No backups have been created yet.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function showValidationResults() {
  const validation = validateData();
  const ui = SpreadsheetApp.getUi();

  if (validation.valid) {
    ui.alert('Data Valid', 'All data validation checks passed!', ui.ButtonSet.OK);
  } else {
    const issues = validation.issues.join('\n• ');
    ui.alert('Validation Issues', `Found the following issues:\n\n• ${issues}`, ui.ButtonSet.OK);
  }
}

function viewErrorLog() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let errorSheet = ss.getSheetByName('Error_Log');
  if (errorSheet) {
    ss.setActiveSheet(errorSheet);
  } else {
    SpreadsheetApp.getUi().alert('No Errors', 'No errors have been logged yet.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function clearErrorLog() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Clear Error Log', 'This will delete all logged errors. Continue?', ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const errorSheet = ss.getSheetByName('Error_Log');
    if (errorSheet) {
      // Keep header, delete all data rows
      const lastRow = errorSheet.getLastRow();
      if (lastRow > 1) {
        errorSheet.deleteRows(2, lastRow - 1);
      }
      ui.alert('Cleared', 'Error log has been cleared.', ui.ButtonSet.OK);
    }
  }
}

// ═══════════════════════════════════════════════════════════════
// CONTENT QUALITY TAGGING SYSTEM
// ═══════════════════════════════════════════════════════════════

/**
 * Setup LLM tag columns with headers
 */
function setupLLMTagColumns() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);

  // Add headers for each LLM's tags
  sheet.getRange(1, LLM_TAG_COLUMNS.GPT_TAGS).setValue('GPT Tags');
  sheet.getRange(1, LLM_TAG_COLUMNS.CLAUDE_TAGS).setValue('Claude Tags');
  sheet.getRange(1, LLM_TAG_COLUMNS.GEMINI_TAGS).setValue('Gemini Tags');
  sheet.getRange(1, LLM_TAG_COLUMNS.GROK_TAGS).setValue('Grok Tags');
  sheet.getRange(1, LLM_TAG_COLUMNS.PERPLEXITY_TAGS).setValue('Perplexity Tags');
  sheet.getRange(1, LLM_TAG_COLUMNS.CONSENSUS_TAGS).setValue('Consensus Tags');
  sheet.getRange(1, LLM_TAG_COLUMNS.TAG_SENTIMENT).setValue('Tag Sentiment');

  // Color code headers
  sheet.getRange(1, LLM_TAG_COLUMNS.GPT_TAGS, 1, 5).setBackground('#e8f0fe').setFontWeight('bold');
  sheet.getRange(1, LLM_TAG_COLUMNS.CONSENSUS_TAGS).setBackground('#d9ead3').setFontWeight('bold');
  sheet.getRange(1, LLM_TAG_COLUMNS.TAG_SENTIMENT).setBackground('#fff2cc').setFontWeight('bold');

  Logger.log('LLM tag columns set up');
}

/**
 * Process LLM tag responses and calculate consensus
 * Call this after all LLMs have tagged an article
 */
function calculateTagConsensus(rowNum) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);

  // Get all LLM tags for this row
  const gptTags = String(sheet.getRange(rowNum, LLM_TAG_COLUMNS.GPT_TAGS).getValue() || '').split(',').map(t => t.trim()).filter(t => t);
  const claudeTags = String(sheet.getRange(rowNum, LLM_TAG_COLUMNS.CLAUDE_TAGS).getValue() || '').split(',').map(t => t.trim()).filter(t => t);
  const geminiTags = String(sheet.getRange(rowNum, LLM_TAG_COLUMNS.GEMINI_TAGS).getValue() || '').split(',').map(t => t.trim()).filter(t => t);
  const grokTags = String(sheet.getRange(rowNum, LLM_TAG_COLUMNS.GROK_TAGS).getValue() || '').split(',').map(t => t.trim()).filter(t => t);
  const perplexityTags = String(sheet.getRange(rowNum, LLM_TAG_COLUMNS.PERPLEXITY_TAGS).getValue() || '').split(',').map(t => t.trim()).filter(t => t);

  // Count tag occurrences across all LLMs
  const tagCounts = {};
  const allTags = [...gptTags, ...claudeTags, ...geminiTags, ...grokTags, ...perplexityTags];

  allTags.forEach(tag => {
    tagCounts[tag] = (tagCounts[tag] || 0) + 1;
  });

  // Find top 3 consensus tags (picked by 2+ LLMs)
  const consensusTags = Object.entries(tagCounts)
    .filter(([tag, count]) => count >= 2)
    .sort((a, b) => b[1] - a[1])
    .map(([tag]) => tag)
    .slice(0, 3);

  // Calculate sentiment: count positive vs negative tags
  let positiveCount = 0;
  let negativeCount = 0;

  for (const tag of Object.keys(tagCounts)) {
    if (POSITIVE_TAGS.includes(tag)) {
      positiveCount += tagCounts[tag];
    } else if (NEGATIVE_TAGS.includes(tag)) {
      negativeCount += tagCounts[tag];
    }
  }

  // Determine overall sentiment
  let sentiment = 'Neutral';
  let sentimentColor = '#fff2cc'; // Yellow

  if (positiveCount > negativeCount + 2) {
    sentiment = `+++ High Quality (${positiveCount}P/${negativeCount}N)`;
    sentimentColor = '#d9ead3'; // Green
  } else if (positiveCount > negativeCount) {
    sentiment = `+ Quality (${positiveCount}P/${negativeCount}N)`;
    sentimentColor = '#d9ead3';
  } else if (negativeCount > positiveCount + 2) {
    sentiment = `--- Low Quality (${positiveCount}P/${negativeCount}N)`;
    sentimentColor = '#f4cccc'; // Red
  } else if (negativeCount > positiveCount) {
    sentiment = `- Biased (${positiveCount}P/${negativeCount}N)`;
    sentimentColor = '#f4cccc';
  } else {
    sentiment = `= Mixed (${positiveCount}P/${negativeCount}N)`;
  }

  // Write results
  sheet.getRange(rowNum, LLM_TAG_COLUMNS.CONSENSUS_TAGS).setValue(consensusTags.join(', '));
  sheet.getRange(rowNum, LLM_TAG_COLUMNS.TAG_SENTIMENT).setValue(sentiment).setBackground(sentimentColor);

  return { consensusTags, sentiment, positiveCount, negativeCount };
}

/**
 * Calculate consensus for all articles that have LLM tags
 */
function calculateAllTagConsensus() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);
  const lastRow = sheet.getLastRow();

  let processed = 0;

  for (let i = 2; i <= lastRow; i++) {
    // Check if any LLM has tagged this row
    const gptTags = sheet.getRange(i, LLM_TAG_COLUMNS.GPT_TAGS).getValue();
    const claudeTags = sheet.getRange(i, LLM_TAG_COLUMNS.CLAUDE_TAGS).getValue();

    if (gptTags || claudeTags) {
      calculateTagConsensus(i);
      processed++;
    }
  }

  Logger.log(`Calculated consensus for ${processed} articles`);
  return processed;
}

/**
 * Get the tagging prompt for LLMs
 */
function getTagPromptForLLM(title, url, content) {
  return `Tag this article with 2-4 quality indicators.

POSITIVE: Data-backed, Expert quoted, Balanced, Original research, Actionable, Well-sourced, Timely, In-depth, Verified, Industry insight

NEGATIVE: Clickbait, Fear-mongering, Unverified, Cherry-picked, Oversimplified, Misleading, Sensational, Outdated, Speculative, Ad-driven

Title: ${title}
${content ? `Content: ${content.substring(0, 1500)}` : ''}

Return ONLY tags, comma-separated:`;
}

/**
 * Setup dropdown selectors for positive and negative tags
 */
function setupTagDropdowns() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);

  // Define columns for tags
  const POSITIVE_TAG_COL = 32; // Column AF
  const NEGATIVE_TAG_COL = 33; // Column AG
  const QUALITY_SCORE_COL = 34; // Column AH

  // Add headers
  sheet.getRange(1, POSITIVE_TAG_COL).setValue('Positive Tags');
  sheet.getRange(1, NEGATIVE_TAG_COL).setValue('Negative Tags');
  sheet.getRange(1, QUALITY_SCORE_COL).setValue('Quality Score');

  // Create data validation rules for dropdowns
  const positiveRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(POSITIVE_TAGS, true)
    .setAllowInvalid(false)
    .build();

  const negativeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(NEGATIVE_TAGS, true)
    .setAllowInvalid(false)
    .build();

  // Apply to rows 2-1000
  const lastRow = Math.max(sheet.getLastRow(), 1000);
  sheet.getRange(2, POSITIVE_TAG_COL, lastRow - 1, 1).setDataValidation(positiveRule);
  sheet.getRange(2, NEGATIVE_TAG_COL, lastRow - 1, 1).setDataValidation(negativeRule);

  // Color the header row
  sheet.getRange(1, POSITIVE_TAG_COL).setBackground('#d9ead3').setFontWeight('bold'); // Green
  sheet.getRange(1, NEGATIVE_TAG_COL).setBackground('#f4cccc').setFontWeight('bold'); // Red
  sheet.getRange(1, QUALITY_SCORE_COL).setBackground('#fff2cc').setFontWeight('bold'); // Yellow

  // Also create a Tags Reference sheet
  let tagSheet = ss.getSheetByName('Tags_Reference');
  if (!tagSheet) {
    tagSheet = ss.insertSheet('Tags_Reference');
  }
  tagSheet.clear();

  // Headers
  tagSheet.getRange(1, 1).setValue('POSITIVE TAGS (Quality Indicators)').setFontWeight('bold').setBackground('#d9ead3');
  tagSheet.getRange(1, 3).setValue('NEGATIVE TAGS (Bias Indicators)').setFontWeight('bold').setBackground('#f4cccc');

  // List all tags
  for (let i = 0; i < 20; i++) {
    tagSheet.getRange(i + 2, 1).setValue(i + 1);
    tagSheet.getRange(i + 2, 2).setValue(POSITIVE_TAGS[i]);
    tagSheet.getRange(i + 2, 3).setValue(i + 1);
    tagSheet.getRange(i + 2, 4).setValue(NEGATIVE_TAGS[i]);
  }

  // Color code
  tagSheet.getRange(2, 1, 20, 2).setBackground('#d9ead3');
  tagSheet.getRange(2, 3, 20, 2).setBackground('#f4cccc');
  tagSheet.autoResizeColumns(1, 4);

  Logger.log('Tag dropdowns set up successfully');
}

/**
 * Auto-tag articles based on title/description keywords
 */
function autoTagArticles() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);
  const data = sheet.getDataRange().getValues();

  const POSITIVE_TAG_COL = 32;
  const NEGATIVE_TAG_COL = 33;
  const QUALITY_SCORE_COL = 34;

  // Ensure headers exist
  if (data[0][POSITIVE_TAG_COL - 1] !== 'Positive Tags') {
    setupTagDropdowns();
  }

  let taggedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const title = String(row[CONFIG.COLUMNS.TITLE - 1] || '').toLowerCase();

    if (!title) continue;

    // Skip if already tagged
    if (row[POSITIVE_TAG_COL - 1] || row[NEGATIVE_TAG_COL - 1]) continue;

    const positiveTags = [];
    const negativeTags = [];

    // Check for tag keywords
    for (const [tag, keywords] of Object.entries(TAG_KEYWORDS)) {
      for (const keyword of keywords) {
        if (title.includes(keyword.toLowerCase())) {
          if (POSITIVE_TAGS.includes(tag)) {
            positiveTags.push(tag);
          } else if (NEGATIVE_TAGS.includes(tag)) {
            negativeTags.push(tag);
          }
          break; // Only add tag once
        }
      }
    }

    // Write tags (first match for dropdown, or leave for manual selection)
    const rowNum = i + 1;
    if (positiveTags.length > 0) {
      sheet.getRange(rowNum, POSITIVE_TAG_COL).setValue(positiveTags[0]);
    }
    if (negativeTags.length > 0) {
      sheet.getRange(rowNum, NEGATIVE_TAG_COL).setValue(negativeTags[0]);
    }

    // Calculate quality score: positive count - negative count
    const qualityScore = positiveTags.length - negativeTags.length;
    sheet.getRange(rowNum, QUALITY_SCORE_COL).setValue(qualityScore);

    // Color code based on score
    if (qualityScore > 0) {
      sheet.getRange(rowNum, QUALITY_SCORE_COL).setBackground('#d9ead3'); // Green
    } else if (qualityScore < 0) {
      sheet.getRange(rowNum, QUALITY_SCORE_COL).setBackground('#f4cccc'); // Red
    } else {
      sheet.getRange(rowNum, QUALITY_SCORE_COL).setBackground('#fff2cc'); // Yellow/neutral
    }

    taggedCount++;
  }

  Logger.log(`Auto-tagged ${taggedCount} articles`);
  return taggedCount;
}

/**
 * Get quality-filtered articles (only high-quality content)
 */
function getHighQualityArticles(minScore = 1) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);
  const data = sheet.getDataRange().getValues();

  const QUALITY_SCORE_COL = 34;
  const articles = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const qualityScore = parseInt(row[QUALITY_SCORE_COL - 1]) || 0;

    if (qualityScore >= minScore) {
      articles.push({
        title: row[CONFIG.COLUMNS.TITLE - 1],
        link: row[CONFIG.COLUMNS.LINK - 1],
        topic: row[CONFIG.COLUMNS.TOPIC - 1],
        qualityScore: qualityScore,
        positiveTag: row[31], // Column AF
        negativeTag: row[32], // Column AG
        row: i + 1
      });
    }
  }

  return articles.sort((a, b) => b.qualityScore - a.qualityScore);
}

/**
 * Create quality report showing tag distribution
 */
function createQualityReport() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);
  const data = sheet.getDataRange().getValues();

  // Create report sheet
  let reportSheet = ss.getSheetByName('Quality_Report');
  if (!reportSheet) {
    reportSheet = ss.insertSheet('Quality_Report');
  }
  reportSheet.clear();

  // Count tags
  const positiveCounts = {};
  const negativeCounts = {};
  POSITIVE_TAGS.forEach(t => positiveCounts[t] = 0);
  NEGATIVE_TAGS.forEach(t => negativeCounts[t] = 0);

  let totalPositive = 0;
  let totalNegative = 0;
  let totalNeutral = 0;

  for (let i = 1; i < data.length; i++) {
    const positiveTag = data[i][31];
    const negativeTag = data[i][32];
    const qualityScore = parseInt(data[i][33]) || 0;

    if (positiveTag && positiveCounts[positiveTag] !== undefined) {
      positiveCounts[positiveTag]++;
    }
    if (negativeTag && negativeCounts[negativeTag] !== undefined) {
      negativeCounts[negativeTag]++;
    }

    if (qualityScore > 0) totalPositive++;
    else if (qualityScore < 0) totalNegative++;
    else totalNeutral++;
  }

  // Write report
  reportSheet.getRange(1, 1).setValue('CONTENT QUALITY REPORT').setFontWeight('bold').setFontSize(14);
  reportSheet.getRange(2, 1).setValue(`Generated: ${new Date().toLocaleString()}`);

  // Summary
  reportSheet.getRange(4, 1).setValue('SUMMARY').setFontWeight('bold');
  reportSheet.getRange(5, 1).setValue('High Quality Articles:');
  reportSheet.getRange(5, 2).setValue(totalPositive).setBackground('#d9ead3');
  reportSheet.getRange(6, 1).setValue('Low Quality Articles:');
  reportSheet.getRange(6, 2).setValue(totalNegative).setBackground('#f4cccc');
  reportSheet.getRange(7, 1).setValue('Neutral Articles:');
  reportSheet.getRange(7, 2).setValue(totalNeutral).setBackground('#fff2cc');

  // Positive tag breakdown
  reportSheet.getRange(9, 1).setValue('POSITIVE TAGS').setFontWeight('bold').setBackground('#d9ead3');
  reportSheet.getRange(9, 2).setValue('Count').setFontWeight('bold');
  let row = 10;
  for (const [tag, count] of Object.entries(positiveCounts).sort((a, b) => b[1] - a[1])) {
    reportSheet.getRange(row, 1).setValue(tag);
    reportSheet.getRange(row, 2).setValue(count);
    row++;
  }

  // Negative tag breakdown
  reportSheet.getRange(9, 4).setValue('NEGATIVE TAGS').setFontWeight('bold').setBackground('#f4cccc');
  reportSheet.getRange(9, 5).setValue('Count').setFontWeight('bold');
  row = 10;
  for (const [tag, count] of Object.entries(negativeCounts).sort((a, b) => b[1] - a[1])) {
    reportSheet.getRange(row, 4).setValue(tag);
    reportSheet.getRange(row, 5).setValue(count);
    row++;
  }

  reportSheet.autoResizeColumns(1, 5);
  Logger.log('Quality report created');
}

// ═══════════════════════════════════════════════════════════════
// ADVANCED AUDIENCE TARGETING
// ═══════════════════════════════════════════════════════════════

/**
 * Granular sub-segment definitions (from LinkedIn analysis)
 */
const SUB_SEGMENTS = {
  'C-Suite': { size: 1359, titles: ['ceo', 'cfo', 'cto', 'coo', 'cmo', 'chief'], topics: ['Leadership/Strategy', 'Enterprise AI', 'Finance AI'] },
  'CTOs/Tech Leaders': { size: 1143, titles: ['cto', 'vp engineering', 'head of engineering', 'technical director'], topics: ['AI Frontier', 'Agents', 'Cloud/Infrastructure', 'Compute/Silicon'] },
  'VCs/Investors': { size: 930, titles: ['partner', 'principal', 'investor', 'venture', 'managing director'], topics: ['Finance AI', 'Fintech', 'Leadership/Strategy'] },
  'Founders/Entrepreneurs': { size: 1058, titles: ['founder', 'co-founder', 'entrepreneur', 'owner'], topics: ['AI Frontier', 'Fintech', 'Consumer Tech'] },
  'Directors': { size: 1969, titles: ['director', 'head of', 'lead'], topics: ['Enterprise AI', 'Cloud/Infrastructure', 'Policy/Governance'] },
  'Engineering/Tech': { size: 1850, titles: ['engineer', 'developer', 'architect', 'programmer'], topics: ['AI Frontier', 'Agents', 'Compute/Silicon', 'Cloud/Infrastructure'] },
  'Product': { size: 680, titles: ['product manager', 'product lead', 'product owner'], topics: ['Consumer Tech', 'Enterprise AI', 'AI Frontier'] },
  'Data/AI Specialists': { size: 920, titles: ['data scientist', 'ml engineer', 'ai', 'machine learning'], topics: ['AI Frontier', 'Agents', 'Enterprise AI'] },
  'Finance Professionals': { size: 785, titles: ['finance', 'financial', 'analyst', 'banking'], topics: ['Finance AI', 'Fintech', 'Leadership/Strategy'] }
};

/**
 * Optimal posting times by segment (LinkedIn engagement data)
 */
const POSTING_SCHEDULE = {
  'C-Suite': { bestDays: [2, 3, 4], bestHours: [7, 8, 17, 18] },           // Tue-Thu, early morning & evening
  'CTOs/Tech Leaders': { bestDays: [2, 3, 4], bestHours: [9, 10, 14, 15] },  // Tue-Thu, mid-morning & afternoon
  'VCs/Investors': { bestDays: [1, 2, 3], bestHours: [8, 9, 16, 17] },       // Mon-Wed, morning & late afternoon
  'Founders/Entrepreneurs': { bestDays: [1, 2, 4], bestHours: [7, 8, 20, 21] }, // Mon, Tue, Thu, early & evening
  'Directors': { bestDays: [2, 3, 4], bestHours: [9, 10, 14, 15] },         // Tue-Thu, business hours
  'Engineering/Tech': { bestDays: [2, 3, 4], bestHours: [10, 11, 15, 16] }, // Tue-Thu, late morning & afternoon
  'Product': { bestDays: [2, 3, 4], bestHours: [9, 10, 14, 15] },           // Tue-Thu, business hours
  'Data/AI Specialists': { bestDays: [2, 3, 4], bestHours: [10, 11, 14, 15] }, // Tue-Thu, late morning & afternoon
  'Finance Professionals': { bestDays: [1, 2, 3], bestHours: [7, 8, 16, 17] }  // Mon-Wed, market hours adjacent
};

/**
 * Route articles to optimal sub-segments based on content analysis
 */
function routeArticlesToSubSegments() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);
  const data = sheet.getDataRange().getValues();

  // Add Sub-Segment column header if not present
  const SUB_SEGMENT_COL = 30;
  if (data[0][SUB_SEGMENT_COL - 1] !== 'Target Sub-Segments') {
    sheet.getRange(1, SUB_SEGMENT_COL).setValue('Target Sub-Segments');
    sheet.getRange(1, SUB_SEGMENT_COL + 1).setValue('Est. Reach');
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const topic = String(row[CONFIG.COLUMNS.TOPIC - 1] || '').toLowerCase().trim();
    const title = String(row[CONFIG.COLUMNS.TITLE - 1] || '').toLowerCase();
    const finalScore = parseFloat(row[CONFIG.COLUMNS.FINAL_SCORE - 1]) || 0;

    if (!topic || finalScore === 0) continue;

    // Find matching sub-segments
    const matchingSegments = [];
    let totalReach = 0;

    for (const [segmentName, segmentData] of Object.entries(SUB_SEGMENTS)) {
      // Check if article topic matches segment interests
      const topicMatch = segmentData.topics.some(t =>
        topic.includes(t.toLowerCase()) || t.toLowerCase().includes(topic)
      );

      // Check title for additional relevance signals
      const titleSignals = [
        title.includes('ai') || title.includes('artificial intelligence'),
        title.includes('funding') || title.includes('raises') || title.includes('investment'),
        title.includes('enterprise') || title.includes('business'),
        title.includes('startup') || title.includes('launch'),
        title.includes('ceo') || title.includes('executive') || title.includes('leader')
      ];

      if (topicMatch) {
        matchingSegments.push(segmentName);
        totalReach += segmentData.size;
      }
    }

    // Write results
    const rowNum = i + 1;
    sheet.getRange(rowNum, SUB_SEGMENT_COL).setValue(matchingSegments.join(', ') || 'General');
    sheet.getRange(rowNum, SUB_SEGMENT_COL + 1).setValue(totalReach);
  }

  Logger.log('Routed articles to sub-segments');
}

/**
 * Build weekly distribution schedule optimized for each segment
 */
function buildWeeklySchedule() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Create or get Schedule sheet
  let scheduleSheet = ss.getSheetByName('Weekly_Schedule');
  if (!scheduleSheet) {
    scheduleSheet = ss.insertSheet('Weekly_Schedule');
  }
  scheduleSheet.clear();

  // Headers
  const headers = ['Day', 'Time', 'Target Segment', 'Article Title', 'Link', 'Score', 'Est. Reach', 'Status'];
  scheduleSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  scheduleSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');

  const newsSheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);
  const newsData = newsSheet.getDataRange().getValues();

  // Collect all scored articles
  const articles = [];
  for (let i = 1; i < newsData.length; i++) {
    const row = newsData[i];
    const finalScore = parseFloat(row[CONFIG.COLUMNS.FINAL_SCORE - 1]) || 0;
    if (finalScore >= 5) { // Only high-scoring articles
      articles.push({
        title: row[CONFIG.COLUMNS.TITLE - 1],
        link: row[CONFIG.COLUMNS.LINK - 1],
        topic: row[CONFIG.COLUMNS.TOPIC - 1],
        score: finalScore,
        segment: row[CONFIG.COLUMNS.TARGET_SEGMENT - 1] || 'TECHNOLOGY'
      });
    }
  }

  // Sort by score
  articles.sort((a, b) => b.score - a.score);

  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  let rowNum = 2;
  const scheduledArticles = new Set();

  // Schedule top articles for each segment
  for (const [segmentName, schedule] of Object.entries(POSTING_SCHEDULE)) {
    const segmentInfo = SUB_SEGMENTS[segmentName];
    if (!segmentInfo) continue;

    // Find articles matching this segment
    const segmentArticles = articles.filter(a => {
      const topicMatch = segmentInfo.topics.some(t =>
        String(a.topic).toLowerCase().includes(t.toLowerCase())
      );
      return topicMatch && !scheduledArticles.has(a.link);
    });

    // Schedule up to 3 articles per segment per week
    let scheduled = 0;
    for (const article of segmentArticles) {
      if (scheduled >= 3) break;

      const dayIndex = schedule.bestDays[scheduled % schedule.bestDays.length];
      const hour = schedule.bestHours[scheduled % schedule.bestHours.length];
      const timeStr = `${hour}:00 ${hour < 12 ? 'AM' : 'PM'}`;

      scheduleSheet.getRange(rowNum, 1, 1, 8).setValues([[
        days[dayIndex],
        timeStr,
        segmentName,
        article.title,
        article.link,
        article.score,
        segmentInfo.size,
        'SCHEDULED'
      ]]);

      scheduledArticles.add(article.link);
      scheduled++;
      rowNum++;
    }
  }

  // Color-code by day
  const dayColors = {
    'Monday': '#e8f0fe',
    'Tuesday': '#fce8e6',
    'Wednesday': '#e6f4ea',
    'Thursday': '#fef7e0',
    'Friday': '#f3e8fd'
  };

  const scheduleData = scheduleSheet.getDataRange().getValues();
  for (let i = 1; i < scheduleData.length; i++) {
    const day = scheduleData[i][0];
    if (dayColors[day]) {
      scheduleSheet.getRange(i + 1, 1, 1, 8).setBackground(dayColors[day]);
    }
  }

  // Auto-resize columns
  scheduleSheet.autoResizeColumns(1, 8);
  scheduleSheet.setColumnWidth(4, 400); // Article title column wider

  Logger.log(`Built weekly schedule with ${rowNum - 2} scheduled posts`);
  return rowNum - 2;
}

/**
 * Get engagement predictions based on historical patterns
 */
function predictEngagement(segmentName, dayOfWeek, hour) {
  const schedule = POSTING_SCHEDULE[segmentName];
  if (!schedule) return 0.5; // Default 50%

  const dayMatch = schedule.bestDays.includes(dayOfWeek);
  const hourMatch = schedule.bestHours.includes(hour);

  if (dayMatch && hourMatch) return 1.0;   // Optimal
  if (dayMatch || hourMatch) return 0.7;   // Partial match
  return 0.4;                               // Off-peak
}

/**
 * Create segment-specific content recommendations
 */
function generateSegmentRecommendations() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Create recommendations sheet
  let recSheet = ss.getSheetByName('Segment_Recommendations');
  if (!recSheet) {
    recSheet = ss.insertSheet('Segment_Recommendations');
  }
  recSheet.clear();

  const headers = ['Segment', 'Audience Size', 'Top Topic', 'Recommended Articles', 'Optimal Post Times', 'Weekly Reach Potential'];
  recSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  recSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#34a853').setFontColor('white');

  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  let rowNum = 2;

  for (const [segmentName, segmentData] of Object.entries(SUB_SEGMENTS)) {
    const schedule = POSTING_SCHEDULE[segmentName];
    const bestDays = schedule.bestDays.map(d => days[d]).join(', ');
    const bestTimes = schedule.bestHours.map(h => `${h}:00`).join(', ');

    // Get top articles for this segment
    const topArticles = [];
    const newsSheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);
    const newsData = newsSheet.getDataRange().getValues();

    for (let i = 1; i < Math.min(newsData.length, 100); i++) {
      const topic = String(newsData[i][CONFIG.COLUMNS.TOPIC - 1] || '').toLowerCase();
      const topicMatch = segmentData.topics.some(t => topic.includes(t.toLowerCase()));
      if (topicMatch) {
        topArticles.push(newsData[i][CONFIG.COLUMNS.TITLE - 1]);
        if (topArticles.length >= 3) break;
      }
    }

    // Weekly reach = audience size * 3 posts * estimated engagement
    const weeklyReach = Math.round(segmentData.size * 3 * 0.15); // 15% avg engagement

    recSheet.getRange(rowNum, 1, 1, 6).setValues([[
      segmentName,
      segmentData.size,
      segmentData.topics[0],
      topArticles.slice(0, 3).join('\n'),
      `${bestDays} at ${bestTimes}`,
      weeklyReach
    ]]);

    rowNum++;
  }

  // Format
  recSheet.autoResizeColumns(1, 6);
  recSheet.setColumnWidth(4, 400);
  recSheet.getRange(2, 6, rowNum - 2, 1).setNumberFormat('#,##0');

  Logger.log('Generated segment recommendations');
}

// ═══════════════════════════════════════════════════════════════
// ERROR HANDLING & BACKUP SYSTEM
// ═══════════════════════════════════════════════════════════════

/**
 * Log errors to dedicated Error_Log sheet
 */
function logError(functionName, error, context = '') {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let errorSheet = ss.getSheetByName('Error_Log');

    if (!errorSheet) {
      errorSheet = ss.insertSheet('Error_Log');
      errorSheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Function', 'Error', 'Context', 'Stack']]);
      errorSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#ea4335').setFontColor('white');
    }

    const nextRow = errorSheet.getLastRow() + 1;
    errorSheet.getRange(nextRow, 1, 1, 5).setValues([[
      new Date().toISOString(),
      functionName,
      String(error.message || error),
      context,
      String(error.stack || 'N/A').substring(0, 500)
    ]]);

    // Keep only last 500 errors
    const totalRows = errorSheet.getLastRow();
    if (totalRows > 501) {
      errorSheet.deleteRows(2, totalRows - 501);
    }

    Logger.log(`ERROR in ${functionName}: ${error.message || error}`);
  } catch (logErr) {
    Logger.log(`Failed to log error: ${logErr}`);
  }
}

/**
 * Create backup of NEWS IN sheet before major operations
 */
function createBackup(operationName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sourceSheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);

    if (!sourceSheet) {
      throw new Error('NEWS IN sheet not found');
    }

    // Create backup sheet name with timestamp
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    const backupName = `BACKUP_${timestamp}_${operationName}`;

    // Check if backup folder sheet exists, create if not
    let backupIndex = ss.getSheetByName('Backup_Index');
    if (!backupIndex) {
      backupIndex = ss.insertSheet('Backup_Index');
      backupIndex.getRange(1, 1, 1, 4).setValues([['Backup Name', 'Created', 'Operation', 'Rows Backed Up']]);
      backupIndex.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    }

    // Copy sheet as backup
    const backup = sourceSheet.copyTo(ss);
    backup.setName(backupName);

    // Move backup to end
    ss.setActiveSheet(backup);
    ss.moveActiveSheet(ss.getNumSheets());

    // Log backup in index
    const nextRow = backupIndex.getLastRow() + 1;
    backupIndex.getRange(nextRow, 1, 1, 4).setValues([[
      backupName,
      new Date().toISOString(),
      operationName,
      sourceSheet.getLastRow()
    ]]);

    // Keep only last 10 backups (delete oldest)
    const backupSheets = ss.getSheets().filter(s => s.getName().startsWith('BACKUP_'));
    if (backupSheets.length > 10) {
      // Sort by name (oldest first due to timestamp format)
      backupSheets.sort((a, b) => a.getName().localeCompare(b.getName()));
      for (let i = 0; i < backupSheets.length - 10; i++) {
        ss.deleteSheet(backupSheets[i]);
      }
    }

    Logger.log(`Backup created: ${backupName}`);
    return backupName;
  } catch (error) {
    logError('createBackup', error, operationName);
    return null;
  }
}

/**
 * Restore from a specific backup
 */
function restoreFromBackup(backupName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const backupSheet = ss.getSheetByName(backupName);

    if (!backupSheet) {
      throw new Error(`Backup "${backupName}" not found`);
    }

    // Get current NEWS IN sheet
    const currentSheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);

    // Rename current to _OLD
    const oldName = `${CONFIG.SHEETS.NEWS_IN}_OLD_${Date.now()}`;
    currentSheet.setName(oldName);

    // Copy backup as new NEWS IN
    const restored = backupSheet.copyTo(ss);
    restored.setName(CONFIG.SHEETS.NEWS_IN);

    // Delete old sheet
    ss.deleteSheet(ss.getSheetByName(oldName));

    Logger.log(`Restored from backup: ${backupName}`);
    return true;
  } catch (error) {
    logError('restoreFromBackup', error, backupName);
    return false;
  }
}

/**
 * List all available backups
 */
function listBackups() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const backupSheets = ss.getSheets().filter(s => s.getName().startsWith('BACKUP_'));

    const backups = backupSheets.map(s => ({
      name: s.getName(),
      rows: s.getLastRow()
    }));

    Logger.log(`Available backups: ${JSON.stringify(backups)}`);
    return backups;
  } catch (error) {
    logError('listBackups', error);
    return [];
  }
}

/**
 * Validate data integrity before operations
 */
function validateData() {
  const issues = [];

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);

    if (!sheet) {
      issues.push('NEWS IN sheet not found');
      return { valid: false, issues };
    }

    const data = sheet.getDataRange().getValues();

    // Check headers exist
    if (data.length < 1) {
      issues.push('Sheet is empty - no headers');
    }

    // Check for required columns
    const headers = data[0];
    const requiredHeaders = ['Topic', 'Title', 'Link'];
    for (const req of requiredHeaders) {
      if (!headers.some(h => String(h).toLowerCase().includes(req.toLowerCase()))) {
        issues.push(`Missing required column: ${req}`);
      }
    }

    // Check for duplicate links
    const links = new Set();
    let duplicates = 0;
    for (let i = 1; i < data.length; i++) {
      const link = data[i][CONFIG.COLUMNS.LINK - 1];
      if (link && links.has(link)) {
        duplicates++;
      }
      if (link) links.add(link);
    }
    if (duplicates > 10) {
      issues.push(`Found ${duplicates} duplicate article links`);
    }

    // Check for empty rows
    let emptyRows = 0;
    for (let i = 1; i < data.length; i++) {
      const title = data[i][CONFIG.COLUMNS.TITLE - 1];
      const link = data[i][CONFIG.COLUMNS.LINK - 1];
      if (!title && !link) {
        emptyRows++;
      }
    }
    if (emptyRows > 5) {
      issues.push(`Found ${emptyRows} empty rows`);
    }

    Logger.log(`Validation complete. Issues: ${issues.length}`);
    return { valid: issues.length === 0, issues };

  } catch (error) {
    logError('validateData', error);
    issues.push(`Validation error: ${error.message}`);
    return { valid: false, issues };
  }
}

/**
 * Safe wrapper for any function with error handling and optional backup
 */
function safeExecute(fn, fnName, createBackupFirst = false) {
  return function(...args) {
    try {
      // Create backup if requested
      if (createBackupFirst) {
        createBackup(fnName);
      }

      // Execute the function
      const result = fn.apply(this, args);
      Logger.log(`${fnName} completed successfully`);
      return result;

    } catch (error) {
      logError(fnName, error, JSON.stringify(args).substring(0, 200));

      // Return safe default based on expected return type
      if (fnName.includes('fetch') || fnName.includes('get')) {
        return [];
      }
      return 0;
    }
  };
}

/**
 * Clean up old data and optimize sheet
 */
function cleanupAndOptimize() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);

    // Create backup first
    createBackup('cleanup');

    const data = sheet.getDataRange().getValues();
    let deletedRows = 0;
    let fixedRows = 0;

    // Process from bottom to top (so row numbers stay valid)
    for (let i = data.length - 1; i >= 1; i--) {
      const title = data[i][CONFIG.COLUMNS.TITLE - 1];
      const link = data[i][CONFIG.COLUMNS.LINK - 1];
      const date = data[i][CONFIG.COLUMNS.DATE - 1];

      // Delete empty rows
      if (!title && !link) {
        sheet.deleteRow(i + 1);
        deletedRows++;
        continue;
      }

      // Fix missing dates
      if (!date && title) {
        sheet.getRange(i + 1, CONFIG.COLUMNS.DATE).setValue(new Date());
        fixedRows++;
      }
    }

    // Remove exact duplicate links (keep first occurrence)
    const seenLinks = new Set();
    const rowsToDelete = [];
    const freshData = sheet.getDataRange().getValues();

    for (let i = 1; i < freshData.length; i++) {
      const link = String(freshData[i][CONFIG.COLUMNS.LINK - 1] || '').trim();
      if (link && seenLinks.has(link)) {
        rowsToDelete.push(i + 1);
      } else if (link) {
        seenLinks.add(link);
      }
    }

    // Delete duplicates from bottom to top
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
      deletedRows++;
    }

    Logger.log(`Cleanup complete: ${deletedRows} rows deleted, ${fixedRows} rows fixed`);
    return { deleted: deletedRows, fixed: fixedRows };

  } catch (error) {
    logError('cleanupAndOptimize', error);
    return { deleted: 0, fixed: 0, error: error.message };
  }
}

/**
 * Health check - run to diagnose issues
 */
function runHealthCheck() {
  const report = {
    timestamp: new Date().toISOString(),
    status: 'OK',
    checks: []
  };

  try {
    // Check 1: Spreadsheet accessible
    try {
      const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      report.checks.push({ name: 'Spreadsheet Access', status: 'PASS' });
    } catch (e) {
      report.checks.push({ name: 'Spreadsheet Access', status: 'FAIL', error: e.message });
      report.status = 'ERROR';
    }

    // Check 2: Required sheets exist
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const requiredSheets = [CONFIG.SHEETS.NEWS_IN];
    for (const sheetName of requiredSheets) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        report.checks.push({ name: `Sheet: ${sheetName}`, status: 'PASS', rows: sheet.getLastRow() });
      } else {
        report.checks.push({ name: `Sheet: ${sheetName}`, status: 'FAIL', error: 'Not found' });
        report.status = 'WARNING';
      }
    }

    // Check 3: Data validation
    const validation = validateData();
    if (validation.valid) {
      report.checks.push({ name: 'Data Validation', status: 'PASS' });
    } else {
      report.checks.push({ name: 'Data Validation', status: 'WARNING', issues: validation.issues });
      if (report.status === 'OK') report.status = 'WARNING';
    }

    // Check 4: RSS feeds reachable
    let feedsOk = 0;
    let feedsFailed = 0;
    for (const feed of RSS_FEEDS.slice(0, 3)) { // Test first 3 only
      try {
        const response = UrlFetchApp.fetch(feed.url, { muteHttpExceptions: true });
        if (response.getResponseCode() === 200) {
          feedsOk++;
        } else {
          feedsFailed++;
        }
      } catch (e) {
        feedsFailed++;
      }
    }
    report.checks.push({ name: 'RSS Feeds', status: feedsFailed === 0 ? 'PASS' : 'WARNING', ok: feedsOk, failed: feedsFailed });

    // Check 5: Backups available
    const backups = listBackups();
    report.checks.push({ name: 'Backups', status: backups.length > 0 ? 'PASS' : 'WARNING', count: backups.length });

    // Check 6: Error log size
    const errorSheet = ss.getSheetByName('Error_Log');
    if (errorSheet) {
      const errorCount = errorSheet.getLastRow() - 1;
      report.checks.push({ name: 'Recent Errors', status: errorCount > 50 ? 'WARNING' : 'PASS', count: errorCount });
    }

  } catch (error) {
    report.status = 'ERROR';
    report.error = error.message;
    logError('runHealthCheck', error);
  }

  // Write report to Health_Report sheet
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let healthSheet = ss.getSheetByName('Health_Report');
    if (!healthSheet) {
      healthSheet = ss.insertSheet('Health_Report');
    }
    healthSheet.clear();

    healthSheet.getRange(1, 1).setValue('NEWS ENGINE HEALTH CHECK').setFontWeight('bold').setFontSize(14);
    healthSheet.getRange(2, 1).setValue(`Status: ${report.status}`).setFontWeight('bold');
    healthSheet.getRange(2, 1).setBackground(report.status === 'OK' ? '#d9ead3' : report.status === 'WARNING' ? '#fff2cc' : '#f4cccc');
    healthSheet.getRange(3, 1).setValue(`Generated: ${report.timestamp}`);

    healthSheet.getRange(5, 1, 1, 3).setValues([['Check', 'Status', 'Details']]).setFontWeight('bold');

    let row = 6;
    for (const check of report.checks) {
      const details = check.error || check.issues?.join(', ') || (check.count !== undefined ? `Count: ${check.count}` : '');
      healthSheet.getRange(row, 1, 1, 3).setValues([[check.name, check.status, details]]);
      healthSheet.getRange(row, 2).setBackground(check.status === 'PASS' ? '#d9ead3' : check.status === 'WARNING' ? '#fff2cc' : '#f4cccc');
      row++;
    }

    healthSheet.autoResizeColumns(1, 3);
  } catch (e) {
    Logger.log(`Could not write health report: ${e}`);
  }

  Logger.log(`Health check complete: ${report.status}`);
  return report;
}

// ═══════════════════════════════════════════════════════════════
// ENHANCED FULL PIPELINE (WITH ERROR HANDLING)
// ═══════════════════════════════════════════════════════════════

/**
 * Run complete pipeline with all enhancements and error handling
 */
function runEnhancedPipeline() {
  Logger.log('=== STARTING ENHANCED PIPELINE ===');

  const results = {
    started: new Date().toISOString(),
    steps: [],
    errors: []
  };

  // Create backup before starting
  const backupName = createBackup('pipeline');
  results.backup = backupName;

  // Step 1: Fetch RSS feeds
  try {
    const rssAdded = fetchAllRSSFeeds();
    results.steps.push({ step: '1a', name: 'RSS Feeds', added: rssAdded, status: 'OK' });
    Logger.log(`Step 1a: Added ${rssAdded} articles from RSS feeds`);
  } catch (error) {
    results.steps.push({ step: '1a', name: 'RSS Feeds', status: 'ERROR', error: error.message });
    results.errors.push({ step: '1a', error: error.message });
    logError('runEnhancedPipeline', error, 'Step 1a: RSS Feeds');
  }

  // Step 1b: Fetch Google News
  try {
    const googleAdded = fetchAllGoogleNews();
    results.steps.push({ step: '1b', name: 'Google News', added: googleAdded, status: 'OK' });
    Logger.log(`Step 1b: Added ${googleAdded} articles from Google News`);
  } catch (error) {
    results.steps.push({ step: '1b', name: 'Google News', status: 'ERROR', error: error.message });
    results.errors.push({ step: '1b', error: error.message });
    logError('runEnhancedPipeline', error, 'Step 1b: Google News');
  }

  // Step 2: Apply audience weights
  try {
    const weighted = applyAudienceWeights();
    results.steps.push({ step: '2', name: 'Audience Weights', processed: weighted, status: 'OK' });
    Logger.log(`Step 2: Applied weights to ${weighted} articles`);
  } catch (error) {
    results.steps.push({ step: '2', name: 'Audience Weights', status: 'ERROR', error: error.message });
    results.errors.push({ step: '2', error: error.message });
    logError('runEnhancedPipeline', error, 'Step 2: Audience Weights');
  }

  // Step 3: Route to sub-segments
  try {
    routeArticlesToSubSegments();
    results.steps.push({ step: '3', name: 'Sub-Segments', status: 'OK' });
    Logger.log('Step 3: Routed articles to sub-segments');
  } catch (error) {
    results.steps.push({ step: '3', name: 'Sub-Segments', status: 'ERROR', error: error.message });
    results.errors.push({ step: '3', error: error.message });
    logError('runEnhancedPipeline', error, 'Step 3: Sub-Segments');
  }

  // Step 4: Create distribution queue
  try {
    const queued = createDistributionQueue();
    results.steps.push({ step: '4', name: 'Distribution Queue', queued: queued, status: 'OK' });
    Logger.log(`Step 4: Created queue with ${queued} articles`);
  } catch (error) {
    results.steps.push({ step: '4', name: 'Distribution Queue', status: 'ERROR', error: error.message });
    results.errors.push({ step: '4', error: error.message });
    logError('runEnhancedPipeline', error, 'Step 4: Distribution Queue');
  }

  // Step 5: Build weekly schedule
  try {
    const scheduled = buildWeeklySchedule();
    results.steps.push({ step: '5', name: 'Weekly Schedule', scheduled: scheduled, status: 'OK' });
    Logger.log(`Step 5: Built schedule with ${scheduled} posts`);
  } catch (error) {
    results.steps.push({ step: '5', name: 'Weekly Schedule', status: 'ERROR', error: error.message });
    results.errors.push({ step: '5', error: error.message });
    logError('runEnhancedPipeline', error, 'Step 5: Weekly Schedule');
  }

  // Step 6: Generate recommendations
  try {
    generateSegmentRecommendations();
    results.steps.push({ step: '6', name: 'Recommendations', status: 'OK' });
    Logger.log('Step 6: Generated segment recommendations');
  } catch (error) {
    results.steps.push({ step: '6', name: 'Recommendations', status: 'ERROR', error: error.message });
    results.errors.push({ step: '6', error: error.message });
    logError('runEnhancedPipeline', error, 'Step 6: Recommendations');
  }

  // Final status
  results.completed = new Date().toISOString();
  results.overallStatus = results.errors.length === 0 ? 'SUCCESS' : 'PARTIAL';

  Logger.log(`=== PIPELINE ${results.overallStatus} (${results.errors.length} errors) ===`);

  return results;
}

/**
 * Safe version of full pipeline that won't crash on errors
 */
function runSafePipeline() {
  try {
    // Run health check first
    const health = runHealthCheck();
    if (health.status === 'ERROR') {
      Logger.log('Health check failed - aborting pipeline');
      return { status: 'ABORTED', reason: 'Health check failed', health };
    }

    // Run the enhanced pipeline
    return runEnhancedPipeline();

  } catch (error) {
    logError('runSafePipeline', error);
    return { status: 'CRASHED', error: error.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// ENGAGEMENT TRACKING & CUSTOM JJ SCORE SYSTEM
// ═══════════════════════════════════════════════════════════════
// This is YOUR moat - trained on YOUR audience's actual behavior

/**
 * Setup the Engagement Log sheet
 */
function setupEngagementLog() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  let engagementSheet = ss.getSheetByName(CONFIG.SHEETS.ENGAGEMENT_LOG);
  if (!engagementSheet) {
    engagementSheet = ss.insertSheet(CONFIG.SHEETS.ENGAGEMENT_LOG);
  }
  engagementSheet.clear();

  // Headers for tracking
  const headers = [
    'Post ID',           // Unique identifier
    'Article Title',     // What you posted about
    'Topic',             // Category
    'Posted Date',       // When posted
    'Day of Week',       // Mon-Sun
    'Hour Posted',       // 0-23
    'Post URL',          // LinkedIn URL
    'Impressions',       // Views
    'Likes',             // Reactions
    'Comments',          // Comment count
    'Shares',            // Reposts
    'Engagement Rate',   // (L+C+S)/Impressions %
    'Comment Quality',   // 1-5 rating (you rate)
    'Virality Score',    // Shares/Impressions
    'AI Score (Pre)',    // What the system predicted
    'Actual Performance',// 1-5 (you rate)
    'Tags Used',         // Positive/Negative tags
    'Hook Type',         // Question/Statement/Stat/Story
    'Content Length',    // Short/Medium/Long
    'Had Image',         // Yes/No
    'Had Link',          // Yes/No
    'Notes'              // Your observations
  ];

  engagementSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  engagementSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white')
    .setWrap(true);

  // Freeze header row
  engagementSheet.setFrozenRows(1);

  // Set column widths
  engagementSheet.setColumnWidth(2, 300); // Article title
  engagementSheet.setColumnWidth(22, 250); // Notes

  // Add data validation for ratings
  const performanceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1 - Flopped', '2 - Below Avg', '3 - Average', '4 - Good', '5 - Viral'], true)
    .build();
  engagementSheet.getRange(2, 16, 500, 1).setDataValidation(performanceRule);

  const hookTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Question', 'Bold Statement', 'Stat/Number', 'Story/Anecdote', 'Contrarian Take', 'How-To'], true)
    .build();
  engagementSheet.getRange(2, 18, 500, 1).setDataValidation(hookTypeRule);

  const lengthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Short (<100 words)', 'Medium (100-250)', 'Long (250+)'], true)
    .build();
  engagementSheet.getRange(2, 19, 500, 1).setDataValidation(lengthRule);

  const yesNoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .build();
  engagementSheet.getRange(2, 20, 500, 1).setDataValidation(yesNoRule);
  engagementSheet.getRange(2, 21, 500, 1).setDataValidation(yesNoRule);

  Logger.log('Engagement Log sheet created');
  return engagementSheet;
}

/**
 * Log a new post's engagement
 * Call this after posting to LinkedIn
 */
function logPostEngagement(articleTitle, topic, postUrl, impressions, likes, comments, shares) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let engagementSheet = ss.getSheetByName(CONFIG.SHEETS.ENGAGEMENT_LOG);

    if (!engagementSheet) {
      engagementSheet = setupEngagementLog();
    }

    const now = new Date();
    const dayOfWeek = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][now.getDay()];
    const hour = now.getHours();

    // Calculate engagement rate
    const engagementRate = impressions > 0
      ? ((likes + comments + shares) / impressions * 100).toFixed(2) + '%'
      : '0%';

    // Calculate virality score (shares are weighted higher)
    const viralityScore = impressions > 0
      ? (shares / impressions * 1000).toFixed(2)
      : '0';

    // Generate unique post ID
    const postId = `POST_${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;

    // Get AI predicted score from NEWS IN if available
    const newsSheet = ss.getSheetByName(CONFIG.SHEETS.NEWS_IN);
    let aiScore = '';
    if (newsSheet && articleTitle) {
      const newsData = newsSheet.getDataRange().getValues();
      for (let i = 1; i < newsData.length; i++) {
        if (newsData[i][CONFIG.COLUMNS.TITLE - 1] === articleTitle) {
          aiScore = newsData[i][CONFIG.COLUMNS.FINAL_SCORE - 1] || '';
          break;
        }
      }
    }

    const nextRow = engagementSheet.getLastRow() + 1;
    engagementSheet.getRange(nextRow, 1, 1, 15).setValues([[
      postId,
      articleTitle,
      topic,
      now,
      dayOfWeek,
      hour,
      postUrl,
      impressions,
      likes,
      comments,
      shares,
      engagementRate,
      '', // Comment quality - manual
      viralityScore,
      aiScore
    ]]);

    Logger.log(`Logged engagement for: ${articleTitle}`);
    return postId;

  } catch (error) {
    logError('logPostEngagement', error, articleTitle);
    return null;
  }
}

/**
 * Update engagement metrics for an existing post
 * Use this to update stats as they grow over 24-48 hours
 */
function updatePostEngagement(postId, impressions, likes, comments, shares) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const engagementSheet = ss.getSheetByName(CONFIG.SHEETS.ENGAGEMENT_LOG);

    if (!engagementSheet) {
      throw new Error('Engagement Log not found');
    }

    const data = engagementSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === postId) {
        const rowNum = i + 1;

        // Update metrics
        engagementSheet.getRange(rowNum, 8).setValue(impressions);
        engagementSheet.getRange(rowNum, 9).setValue(likes);
        engagementSheet.getRange(rowNum, 10).setValue(comments);
        engagementSheet.getRange(rowNum, 11).setValue(shares);

        // Recalculate engagement rate
        const engagementRate = impressions > 0
          ? ((likes + comments + shares) / impressions * 100).toFixed(2) + '%'
          : '0%';
        engagementSheet.getRange(rowNum, 12).setValue(engagementRate);

        // Recalculate virality
        const viralityScore = impressions > 0
          ? (shares / impressions * 1000).toFixed(2)
          : '0';
        engagementSheet.getRange(rowNum, 14).setValue(viralityScore);

        Logger.log(`Updated engagement for ${postId}`);
        return true;
      }
    }

    Logger.log(`Post ${postId} not found`);
    return false;

  } catch (error) {
    logError('updatePostEngagement', error, postId);
    return false;
  }
}

// ═══════════════════════════════════════════════════════════════
// JJ SCORE - YOUR CUSTOM ALGORITHM
// ═══════════════════════════════════════════════════════════════

/**
 * Initial weight factors (will be trained by your data)
 */
const JJ_SCORE_WEIGHTS = {
  // Topic multipliers (start equal, will diverge based on YOUR data)
  topics: {
    'AI Frontier': 1.0,
    'Agents': 1.0,
    'Finance AI': 1.0,
    'Fintech': 1.0,
    'Cloud/Infrastructure': 1.0,
    'Biotech/Healthcare': 1.0,
    'Leadership/Strategy': 1.0,
    'Consumer Tech': 1.0,
    'Robotics': 1.0,
    'Enterprise AI': 1.0,
    'default': 1.0
  },

  // Day of week multipliers
  dayOfWeek: {
    'Mon': 1.0,
    'Tue': 1.15,
    'Wed': 1.10,
    'Thu': 1.05,
    'Fri': 0.85,
    'Sat': 0.60,
    'Sun': 0.65
  },

  // Hour multipliers (will learn YOUR best times)
  hourBuckets: {
    'early_morning': 0.9,   // 5-8am
    'morning': 1.15,        // 8-11am
    'midday': 0.95,         // 11am-2pm
    'afternoon': 1.10,      // 2-5pm
    'evening': 0.85,        // 5-8pm
    'night': 0.60           // 8pm-5am
  },

  // Hook type multipliers
  hookTypes: {
    'Question': 1.0,
    'Bold Statement': 1.1,
    'Stat/Number': 1.15,
    'Story/Anecdote': 1.05,
    'Contrarian Take': 1.20,
    'How-To': 0.95
  },

  // Content characteristics
  hasImage: 1.25,
  hasLink: 0.90,  // LinkedIn deprioritizes external links

  // Length preferences
  contentLength: {
    'Short (<100 words)': 0.95,
    'Medium (100-250)': 1.10,
    'Long (250+)': 1.0
  }
};

/**
 * Calculate JJ Score for an article
 * This uses YOUR trained weights
 */
function calculateJJScore(title, topic, dayOfWeek, hour, hookType, hasImage, hasLink, contentLength) {
  // Base score from AI average (if available)
  let baseScore = 5; // Default middle score

  // Apply topic weight
  const topicWeight = JJ_SCORE_WEIGHTS.topics[topic] || JJ_SCORE_WEIGHTS.topics['default'];

  // Apply day weight
  const dayWeight = JJ_SCORE_WEIGHTS.dayOfWeek[dayOfWeek] || 1.0;

  // Apply hour weight
  let hourBucket = 'morning';
  if (hour >= 5 && hour < 8) hourBucket = 'early_morning';
  else if (hour >= 8 && hour < 11) hourBucket = 'morning';
  else if (hour >= 11 && hour < 14) hourBucket = 'midday';
  else if (hour >= 14 && hour < 17) hourBucket = 'afternoon';
  else if (hour >= 17 && hour < 20) hourBucket = 'evening';
  else hourBucket = 'night';
  const hourWeight = JJ_SCORE_WEIGHTS.hourBuckets[hourBucket];

  // Apply hook type weight
  const hookWeight = JJ_SCORE_WEIGHTS.hookTypes[hookType] || 1.0;

  // Apply content characteristics
  const imageWeight = hasImage ? JJ_SCORE_WEIGHTS.hasImage : 1.0;
  const linkWeight = hasLink ? JJ_SCORE_WEIGHTS.hasLink : 1.0;
  const lengthWeight = JJ_SCORE_WEIGHTS.contentLength[contentLength] || 1.0;

  // Calculate final JJ Score
  const jjScore = baseScore * topicWeight * dayWeight * hourWeight * hookWeight * imageWeight * linkWeight * lengthWeight;

  return Math.round(jjScore * 10) / 10;
}

/**
 * Train JJ Score weights from your actual engagement data
 * THIS IS THE MAGIC - learns what works for YOUR audience
 */
function trainJJScoreModel() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const engagementSheet = ss.getSheetByName(CONFIG.SHEETS.ENGAGEMENT_LOG);

    if (!engagementSheet) {
      Logger.log('No engagement data to train on');
      return null;
    }

    const data = engagementSheet.getDataRange().getValues();
    if (data.length < 11) { // Need at least 10 posts to train
      Logger.log('Need at least 10 posts to train the model');
      return null;
    }

    // Collect performance by category
    const topicPerformance = {};
    const dayPerformance = {};
    const hourPerformance = {};
    const hookPerformance = {};
    const lengthPerformance = {};
    const imagePerformance = { yes: [], no: [] };
    const linkPerformance = { yes: [], no: [] };

    // Process each post
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const topic = row[2];
      const dayOfWeek = row[4];
      const hour = row[5];
      const engagementRate = parseFloat(String(row[11]).replace('%', '')) || 0;
      const actualPerformance = parseInt(String(row[15]).charAt(0)) || 3;
      const hookType = row[17];
      const contentLength = row[18];
      const hadImage = row[19];
      const hadLink = row[20];

      // Performance score combines engagement rate and manual rating
      const perfScore = (engagementRate * 10) + (actualPerformance * 5);

      // Accumulate by category
      if (topic) {
        if (!topicPerformance[topic]) topicPerformance[topic] = [];
        topicPerformance[topic].push(perfScore);
      }

      if (dayOfWeek) {
        if (!dayPerformance[dayOfWeek]) dayPerformance[dayOfWeek] = [];
        dayPerformance[dayOfWeek].push(perfScore);
      }

      if (hour !== undefined) {
        let hourBucket = 'morning';
        if (hour >= 5 && hour < 8) hourBucket = 'early_morning';
        else if (hour >= 8 && hour < 11) hourBucket = 'morning';
        else if (hour >= 11 && hour < 14) hourBucket = 'midday';
        else if (hour >= 14 && hour < 17) hourBucket = 'afternoon';
        else if (hour >= 17 && hour < 20) hourBucket = 'evening';
        else hourBucket = 'night';

        if (!hourPerformance[hourBucket]) hourPerformance[hourBucket] = [];
        hourPerformance[hourBucket].push(perfScore);
      }

      if (hookType) {
        if (!hookPerformance[hookType]) hookPerformance[hookType] = [];
        hookPerformance[hookType].push(perfScore);
      }

      if (contentLength) {
        if (!lengthPerformance[contentLength]) lengthPerformance[contentLength] = [];
        lengthPerformance[contentLength].push(perfScore);
      }

      if (hadImage === 'Yes') imagePerformance.yes.push(perfScore);
      else if (hadImage === 'No') imagePerformance.no.push(perfScore);

      if (hadLink === 'Yes') linkPerformance.yes.push(perfScore);
      else if (hadLink === 'No') linkPerformance.no.push(perfScore);
    }

    // Calculate averages and convert to weights
    const avg = arr => arr.length > 0 ? arr.reduce((a, b) => a + b, 0) / arr.length : 0;
    const overallAvg = avg(Object.values(topicPerformance).flat());

    // Update topic weights
    const newWeights = JSON.parse(JSON.stringify(JJ_SCORE_WEIGHTS)); // Deep copy

    for (const [topic, scores] of Object.entries(topicPerformance)) {
      if (scores.length >= 3) { // Need 3+ samples
        newWeights.topics[topic] = avg(scores) / overallAvg;
      }
    }

    for (const [day, scores] of Object.entries(dayPerformance)) {
      if (scores.length >= 2) {
        newWeights.dayOfWeek[day] = avg(scores) / overallAvg;
      }
    }

    for (const [bucket, scores] of Object.entries(hourPerformance)) {
      if (scores.length >= 2) {
        newWeights.hourBuckets[bucket] = avg(scores) / overallAvg;
      }
    }

    for (const [hook, scores] of Object.entries(hookPerformance)) {
      if (scores.length >= 2) {
        newWeights.hookTypes[hook] = avg(scores) / overallAvg;
      }
    }

    for (const [length, scores] of Object.entries(lengthPerformance)) {
      if (scores.length >= 2) {
        newWeights.contentLength[length] = avg(scores) / overallAvg;
      }
    }

    // Image and link weights
    if (imagePerformance.yes.length >= 2 && imagePerformance.no.length >= 2) {
      newWeights.hasImage = avg(imagePerformance.yes) / avg(imagePerformance.no);
    }

    if (linkPerformance.yes.length >= 2 && linkPerformance.no.length >= 2) {
      newWeights.hasLink = avg(linkPerformance.yes) / avg(linkPerformance.no);
    }

    // Save trained model to sheet
    saveJJScoreModel(newWeights, data.length - 1);

    Logger.log(`JJ Score model trained on ${data.length - 1} posts`);
    return newWeights;

  } catch (error) {
    logError('trainJJScoreModel', error);
    return null;
  }
}

/**
 * Save the trained model to a sheet for inspection and persistence
 */
function saveJJScoreModel(weights, sampleSize) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  let modelSheet = ss.getSheetByName(CONFIG.SHEETS.JJ_SCORE_MODEL);
  if (!modelSheet) {
    modelSheet = ss.insertSheet(CONFIG.SHEETS.JJ_SCORE_MODEL);
  }
  modelSheet.clear();

  // Header
  modelSheet.getRange(1, 1).setValue('JJ SCORE MODEL - Trained on YOUR Audience').setFontWeight('bold').setFontSize(14);
  modelSheet.getRange(2, 1).setValue(`Last trained: ${new Date().toLocaleString()}`);
  modelSheet.getRange(3, 1).setValue(`Sample size: ${sampleSize} posts`);

  let row = 5;

  // Topics section
  modelSheet.getRange(row, 1).setValue('TOPIC WEIGHTS').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  modelSheet.getRange(row, 2).setValue('Weight').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  modelSheet.getRange(row, 3).setValue('Interpretation').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  row++;

  for (const [topic, weight] of Object.entries(weights.topics).sort((a, b) => b[1] - a[1])) {
    const interpretation = weight > 1.1 ? '🔥 High performer' : weight < 0.9 ? '⚠️ Underperforms' : '➡️ Average';
    modelSheet.getRange(row, 1).setValue(topic);
    modelSheet.getRange(row, 2).setValue(Math.round(weight * 100) / 100);
    modelSheet.getRange(row, 3).setValue(interpretation);
    if (weight > 1.1) modelSheet.getRange(row, 1, 1, 3).setBackground('#d9ead3');
    else if (weight < 0.9) modelSheet.getRange(row, 1, 1, 3).setBackground('#f4cccc');
    row++;
  }

  row += 2;

  // Day of week section
  modelSheet.getRange(row, 1).setValue('DAY OF WEEK').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
  modelSheet.getRange(row, 2).setValue('Weight').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
  row++;

  const dayOrder = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  for (const day of dayOrder) {
    const weight = weights.dayOfWeek[day] || 1.0;
    modelSheet.getRange(row, 1).setValue(day);
    modelSheet.getRange(row, 2).setValue(Math.round(weight * 100) / 100);
    if (weight > 1.05) modelSheet.getRange(row, 1, 1, 2).setBackground('#d9ead3');
    else if (weight < 0.85) modelSheet.getRange(row, 1, 1, 2).setBackground('#f4cccc');
    row++;
  }

  row += 2;

  // Hour buckets section
  modelSheet.getRange(row, 1).setValue('TIME OF DAY').setFontWeight('bold').setBackground('#ea4335').setFontColor('white');
  modelSheet.getRange(row, 2).setValue('Weight').setFontWeight('bold').setBackground('#ea4335').setFontColor('white');
  row++;

  for (const [bucket, weight] of Object.entries(weights.hourBuckets)) {
    modelSheet.getRange(row, 1).setValue(bucket);
    modelSheet.getRange(row, 2).setValue(Math.round(weight * 100) / 100);
    if (weight > 1.05) modelSheet.getRange(row, 1, 1, 2).setBackground('#d9ead3');
    else if (weight < 0.85) modelSheet.getRange(row, 1, 1, 2).setBackground('#f4cccc');
    row++;
  }

  row += 2;

  // Hook types section
  modelSheet.getRange(row, 1).setValue('HOOK TYPES').setFontWeight('bold').setBackground('#fbbc04').setFontColor('white');
  modelSheet.getRange(row, 2).setValue('Weight').setFontWeight('bold').setBackground('#fbbc04').setFontColor('white');
  row++;

  for (const [hook, weight] of Object.entries(weights.hookTypes).sort((a, b) => b[1] - a[1])) {
    modelSheet.getRange(row, 1).setValue(hook);
    modelSheet.getRange(row, 2).setValue(Math.round(weight * 100) / 100);
    if (weight > 1.1) modelSheet.getRange(row, 1, 1, 2).setBackground('#d9ead3');
    else if (weight < 0.95) modelSheet.getRange(row, 1, 1, 2).setBackground('#f4cccc');
    row++;
  }

  row += 2;

  // Content characteristics
  modelSheet.getRange(row, 1).setValue('CONTENT FACTORS').setFontWeight('bold').setBackground('#9c27b0').setFontColor('white');
  modelSheet.getRange(row, 2).setValue('Weight').setFontWeight('bold').setBackground('#9c27b0').setFontColor('white');
  row++;

  modelSheet.getRange(row, 1).setValue('Has Image');
  modelSheet.getRange(row, 2).setValue(Math.round(weights.hasImage * 100) / 100);
  row++;

  modelSheet.getRange(row, 1).setValue('Has Link');
  modelSheet.getRange(row, 2).setValue(Math.round(weights.hasLink * 100) / 100);
  row++;

  for (const [length, weight] of Object.entries(weights.contentLength)) {
    modelSheet.getRange(row, 1).setValue(length);
    modelSheet.getRange(row, 2).setValue(Math.round(weight * 100) / 100);
    row++;
  }

  modelSheet.autoResizeColumns(1, 3);
  Logger.log('JJ Score model saved');
}

// ═══════════════════════════════════════════════════════════════
// PERFORMANCE INSIGHTS - WHAT'S WORKING
// ═══════════════════════════════════════════════════════════════

/**
 * Generate insights report from your engagement data
 */
function generatePerformanceInsights() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const engagementSheet = ss.getSheetByName(CONFIG.SHEETS.ENGAGEMENT_LOG);

    if (!engagementSheet) {
      Logger.log('No engagement data found');
      return null;
    }

    const data = engagementSheet.getDataRange().getValues();
    if (data.length < 6) {
      Logger.log('Need at least 5 posts for insights');
      return null;
    }

    // Create insights sheet
    let insightsSheet = ss.getSheetByName(CONFIG.SHEETS.PERFORMANCE_INSIGHTS);
    if (!insightsSheet) {
      insightsSheet = ss.insertSheet(CONFIG.SHEETS.PERFORMANCE_INSIGHTS);
    }
    insightsSheet.clear();

    // Process data
    const posts = [];
    for (let i = 1; i < data.length; i++) {
      const engRate = parseFloat(String(data[i][11]).replace('%', '')) || 0;
      posts.push({
        title: data[i][1],
        topic: data[i][2],
        dayOfWeek: data[i][4],
        hour: data[i][5],
        impressions: data[i][7] || 0,
        likes: data[i][8] || 0,
        comments: data[i][9] || 0,
        shares: data[i][10] || 0,
        engagementRate: engRate,
        hookType: data[i][17],
        contentLength: data[i][18],
        hadImage: data[i][19],
        hadLink: data[i][20]
      });
    }

    // Sort by engagement
    const sortedByEngagement = [...posts].sort((a, b) => b.engagementRate - a.engagementRate);

    // Calculate averages
    const avgEngagement = posts.reduce((sum, p) => sum + p.engagementRate, 0) / posts.length;
    const avgImpressions = posts.reduce((sum, p) => sum + p.impressions, 0) / posts.length;
    const totalReach = posts.reduce((sum, p) => sum + p.impressions, 0);

    // Write report
    let row = 1;

    insightsSheet.getRange(row, 1).setValue('📊 PERFORMANCE INSIGHTS').setFontWeight('bold').setFontSize(16);
    row += 2;

    insightsSheet.getRange(row, 1).setValue(`Generated: ${new Date().toLocaleString()}`);
    insightsSheet.getRange(row, 3).setValue(`Based on ${posts.length} posts`);
    row += 2;

    // Key metrics
    insightsSheet.getRange(row, 1).setValue('KEY METRICS').setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    row++;
    insightsSheet.getRange(row, 1).setValue('Total Reach:');
    insightsSheet.getRange(row, 2).setValue(totalReach.toLocaleString()).setFontWeight('bold');
    row++;
    insightsSheet.getRange(row, 1).setValue('Avg Impressions:');
    insightsSheet.getRange(row, 2).setValue(Math.round(avgImpressions).toLocaleString());
    row++;
    insightsSheet.getRange(row, 1).setValue('Avg Engagement Rate:');
    insightsSheet.getRange(row, 2).setValue(avgEngagement.toFixed(2) + '%');
    row += 2;

    // Top performers
    insightsSheet.getRange(row, 1).setValue('🏆 TOP 5 PERFORMERS').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
    row++;
    for (let i = 0; i < Math.min(5, sortedByEngagement.length); i++) {
      const p = sortedByEngagement[i];
      insightsSheet.getRange(row, 1).setValue(`${i + 1}. ${p.title.substring(0, 50)}...`);
      insightsSheet.getRange(row, 2).setValue(`${p.engagementRate.toFixed(2)}%`).setFontWeight('bold');
      insightsSheet.getRange(row, 3).setValue(p.topic);
      insightsSheet.getRange(row, 4).setValue(`${p.dayOfWeek} ${p.hour}:00`);
      row++;
    }
    row += 2;

    // Best performing topics
    const topicStats = {};
    posts.forEach(p => {
      if (!topicStats[p.topic]) topicStats[p.topic] = { total: 0, count: 0 };
      topicStats[p.topic].total += p.engagementRate;
      topicStats[p.topic].count++;
    });

    const topicRanking = Object.entries(topicStats)
      .map(([topic, stats]) => ({ topic, avg: stats.total / stats.count, count: stats.count }))
      .filter(t => t.count >= 2)
      .sort((a, b) => b.avg - a.avg);

    if (topicRanking.length > 0) {
      insightsSheet.getRange(row, 1).setValue('📈 TOPIC PERFORMANCE').setFontWeight('bold').setBackground('#fbbc04').setFontColor('white');
      insightsSheet.getRange(row, 2).setValue('Avg Eng.').setFontWeight('bold').setBackground('#fbbc04').setFontColor('white');
      insightsSheet.getRange(row, 3).setValue('Posts').setFontWeight('bold').setBackground('#fbbc04').setFontColor('white');
      row++;

      for (const t of topicRanking) {
        insightsSheet.getRange(row, 1).setValue(t.topic);
        insightsSheet.getRange(row, 2).setValue(t.avg.toFixed(2) + '%');
        insightsSheet.getRange(row, 3).setValue(t.count);
        if (t.avg > avgEngagement * 1.2) insightsSheet.getRange(row, 1, 1, 3).setBackground('#d9ead3');
        else if (t.avg < avgEngagement * 0.8) insightsSheet.getRange(row, 1, 1, 3).setBackground('#f4cccc');
        row++;
      }
      row += 2;
    }

    // Best days
    const dayStats = {};
    posts.forEach(p => {
      if (!dayStats[p.dayOfWeek]) dayStats[p.dayOfWeek] = { total: 0, count: 0 };
      dayStats[p.dayOfWeek].total += p.engagementRate;
      dayStats[p.dayOfWeek].count++;
    });

    const dayRanking = Object.entries(dayStats)
      .map(([day, stats]) => ({ day, avg: stats.total / stats.count, count: stats.count }))
      .sort((a, b) => b.avg - a.avg);

    insightsSheet.getRange(row, 1).setValue('📅 BEST DAYS TO POST').setFontWeight('bold').setBackground('#9c27b0').setFontColor('white');
    row++;
    for (const d of dayRanking.slice(0, 3)) {
      insightsSheet.getRange(row, 1).setValue(d.day);
      insightsSheet.getRange(row, 2).setValue(d.avg.toFixed(2) + '%');
      insightsSheet.getRange(row, 3).setValue(`(${d.count} posts)`);
      row++;
    }
    row += 2;

    // Actionable recommendations
    insightsSheet.getRange(row, 1).setValue('💡 RECOMMENDATIONS').setFontWeight('bold').setBackground('#ea4335').setFontColor('white');
    row++;

    const recommendations = [];

    if (topicRanking.length > 0) {
      recommendations.push(`Post more about: ${topicRanking[0].topic} (${topicRanking[0].avg.toFixed(1)}% avg engagement)`);
    }

    if (dayRanking.length > 0) {
      recommendations.push(`Best posting day: ${dayRanking[0].day}`);
    }

    // Image analysis
    const withImage = posts.filter(p => p.hadImage === 'Yes');
    const withoutImage = posts.filter(p => p.hadImage === 'No');
    if (withImage.length >= 2 && withoutImage.length >= 2) {
      const imgAvg = withImage.reduce((s, p) => s + p.engagementRate, 0) / withImage.length;
      const noImgAvg = withoutImage.reduce((s, p) => s + p.engagementRate, 0) / withoutImage.length;
      if (imgAvg > noImgAvg * 1.2) {
        recommendations.push(`Images boost engagement by ${Math.round((imgAvg / noImgAvg - 1) * 100)}% - always include visuals`);
      }
    }

    // Link analysis
    const withLink = posts.filter(p => p.hadLink === 'Yes');
    const withoutLink = posts.filter(p => p.hadLink === 'No');
    if (withLink.length >= 2 && withoutLink.length >= 2) {
      const linkAvg = withLink.reduce((s, p) => s + p.engagementRate, 0) / withLink.length;
      const noLinkAvg = withoutLink.reduce((s, p) => s + p.engagementRate, 0) / withoutLink.length;
      if (noLinkAvg > linkAvg * 1.1) {
        recommendations.push(`Posts without links perform ${Math.round((noLinkAvg / linkAvg - 1) * 100)}% better - put links in comments`);
      }
    }

    for (const rec of recommendations) {
      insightsSheet.getRange(row, 1).setValue('→ ' + rec);
      row++;
    }

    insightsSheet.autoResizeColumns(1, 4);
    insightsSheet.setColumnWidth(1, 350);

    Logger.log('Performance insights generated');
    return { posts: posts.length, avgEngagement, recommendations };

  } catch (error) {
    logError('generatePerformanceInsights', error);
    return null;
  }
}

/**
 * Quick input form for logging engagement
 * Opens a sidebar for easy data entry
 */
function showEngagementInputForm() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      h3 { color: #1a73e8; }
      input, select { width: 100%; padding: 8px; margin: 5px 0 15px 0; border: 1px solid #ddd; border-radius: 4px; }
      button { background: #1a73e8; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; width: 100%; }
      button:hover { background: #1557b0; }
      label { font-weight: bold; color: #333; }
    </style>
    <h3>📊 Log Post Engagement</h3>
    <form id="engForm">
      <label>Article Title</label>
      <input type="text" id="title" placeholder="What article did you post about?">

      <label>Topic</label>
      <select id="topic">
        <option value="AI Frontier">AI Frontier</option>
        <option value="Finance AI">Finance AI</option>
        <option value="Fintech">Fintech</option>
        <option value="Cloud/Infrastructure">Cloud/Infrastructure</option>
        <option value="Biotech/Healthcare">Biotech/Healthcare</option>
        <option value="Leadership/Strategy">Leadership/Strategy</option>
        <option value="Consumer Tech">Consumer Tech</option>
        <option value="Enterprise AI">Enterprise AI</option>
      </select>

      <label>LinkedIn Post URL</label>
      <input type="text" id="postUrl" placeholder="https://linkedin.com/posts/...">

      <label>Impressions</label>
      <input type="number" id="impressions" placeholder="Views">

      <label>Likes</label>
      <input type="number" id="likes" placeholder="Reactions">

      <label>Comments</label>
      <input type="number" id="comments" placeholder="Comments">

      <label>Shares</label>
      <input type="number" id="shares" placeholder="Reposts">

      <button type="button" onclick="submitForm()">Log Engagement</button>
    </form>
    <div id="result" style="margin-top: 15px;"></div>

    <script>
      function submitForm() {
        const data = {
          title: document.getElementById('title').value,
          topic: document.getElementById('topic').value,
          postUrl: document.getElementById('postUrl').value,
          impressions: parseInt(document.getElementById('impressions').value) || 0,
          likes: parseInt(document.getElementById('likes').value) || 0,
          comments: parseInt(document.getElementById('comments').value) || 0,
          shares: parseInt(document.getElementById('shares').value) || 0
        };

        google.script.run
          .withSuccessHandler(function(result) {
            document.getElementById('result').innerHTML = '<p style="color: green;">✓ Logged successfully! ID: ' + result + '</p>';
            document.getElementById('engForm').reset();
          })
          .withFailureHandler(function(error) {
            document.getElementById('result').innerHTML = '<p style="color: red;">Error: ' + error + '</p>';
          })
          .logPostEngagement(data.title, data.topic, data.postUrl, data.impressions, data.likes, data.comments, data.shares);
      }
    </script>
  `)
    .setWidth(350)
    .setTitle('Log Engagement');

  SpreadsheetApp.getUi().showSidebar(html);
}

// ═══════════════════════════════════════════════════════════════
// BIT.LY INTEGRATION - TRACK LINK CLICKS
// ═══════════════════════════════════════════════════════════════

/**
 * Bit.ly API Configuration
 * Get your access token from: https://app.bitly.com/settings/api/
 */
const BITLY_CONFIG = {
  ACCESS_TOKEN: 'YOUR_BITLY_ACCESS_TOKEN', // Replace with your token
  GROUP_GUID: '', // Optional - leave blank to use default group
};

/**
 * Create a Bit.ly short link for an article
 */
function createBitlyLink(longUrl, title) {
  try {
    const payload = {
      long_url: longUrl,
      title: title
    };

    if (BITLY_CONFIG.GROUP_GUID) {
      payload.group_guid = BITLY_CONFIG.GROUP_GUID;
    }

    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${BITLY_CONFIG.ACCESS_TOKEN}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch('https://api-ssl.bitly.com/v4/shorten', options);
    const result = JSON.parse(response.getContentText());

    if (result.link) {
      Logger.log(`Created Bit.ly link: ${result.link} for ${longUrl}`);
      return {
        shortUrl: result.link,
        id: result.id,
        longUrl: longUrl
      };
    } else {
      Logger.log(`Bit.ly error: ${JSON.stringify(result)}`);
      return null;
    }
  } catch (error) {
    logError('createBitlyLink', error, longUrl);
    return null;
  }
}

/**
 * Get click stats for a Bit.ly link
 */
function getBitlyClicks(bitlyLink) {
  try {
    // Extract the bitlink ID (remove https://bit.ly/)
    const bitlinkId = bitlyLink.replace('https://', '').replace('http://', '');

    const options = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${BITLY_CONFIG.ACCESS_TOKEN}`
      },
      muteHttpExceptions: true
    };

    // Get total clicks
    const clicksResponse = UrlFetchApp.fetch(
      `https://api-ssl.bitly.com/v4/bitlinks/${bitlinkId}/clicks/summary`,
      options
    );
    const clicksData = JSON.parse(clicksResponse.getContentText());

    // Get clicks by country
    const countryResponse = UrlFetchApp.fetch(
      `https://api-ssl.bitly.com/v4/bitlinks/${bitlinkId}/countries`,
      options
    );
    const countryData = JSON.parse(countryResponse.getContentText());

    // Get clicks by referrer
    const referrerResponse = UrlFetchApp.fetch(
      `https://api-ssl.bitly.com/v4/bitlinks/${bitlinkId}/referrers`,
      options
    );
    const referrerData = JSON.parse(referrerResponse.getContentText());

    return {
      totalClicks: clicksData.total_clicks || 0,
      countries: countryData.metrics || [],
      referrers: referrerData.metrics || []
    };
  } catch (error) {
    logError('getBitlyClicks', error, bitlyLink);
    return { totalClicks: 0, countries: [], referrers: [] };
  }
}

/**
 * Setup Bit.ly tracking sheet
 */
function setupBitlyTracking() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  let bitlySheet = ss.getSheetByName('Bitly_Tracking');
  if (!bitlySheet) {
    bitlySheet = ss.insertSheet('Bitly_Tracking');
  }
  bitlySheet.clear();

  const headers = [
    'Post ID',           // Links to Engagement_Log
    'Article Title',
    'Original URL',
    'Bit.ly Link',
    'Created Date',
    'Total Clicks',
    'Last Updated',
    'Top Country',
    'Top Referrer',
    'LinkedIn Clicks',   // From referrer data
    'Click-Through Rate', // Clicks / Impressions
    'Notes'
  ];

  bitlySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  bitlySheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#ee6123')
    .setFontColor('white');

  bitlySheet.setFrozenRows(1);
  bitlySheet.setColumnWidth(2, 300);
  bitlySheet.setColumnWidth(3, 250);
  bitlySheet.setColumnWidth(4, 150);

  Logger.log('Bit.ly tracking sheet created');
  return bitlySheet;
}

/**
 * Log a new Bit.ly link and associate with a post
 */
function logBitlyLink(postId, articleTitle, originalUrl, bitlyLink) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let bitlySheet = ss.getSheetByName('Bitly_Tracking');

    if (!bitlySheet) {
      bitlySheet = setupBitlyTracking();
    }

    const nextRow = bitlySheet.getLastRow() + 1;
    bitlySheet.getRange(nextRow, 1, 1, 5).setValues([[
      postId,
      articleTitle,
      originalUrl,
      bitlyLink,
      new Date()
    ]]);

    // Get initial clicks
    const clicks = getBitlyClicks(bitlyLink);
    bitlySheet.getRange(nextRow, 6).setValue(clicks.totalClicks);
    bitlySheet.getRange(nextRow, 7).setValue(new Date());

    if (clicks.countries.length > 0) {
      bitlySheet.getRange(nextRow, 8).setValue(clicks.countries[0].value);
    }
    if (clicks.referrers.length > 0) {
      bitlySheet.getRange(nextRow, 9).setValue(clicks.referrers[0].value);
      // Check for LinkedIn clicks
      const linkedInRef = clicks.referrers.find(r =>
        r.value.toLowerCase().includes('linkedin')
      );
      if (linkedInRef) {
        bitlySheet.getRange(nextRow, 10).setValue(linkedInRef.clicks);
      }
    }

    Logger.log(`Logged Bit.ly link: ${bitlyLink}`);
    return true;
  } catch (error) {
    logError('logBitlyLink', error, bitlyLink);
    return false;
  }
}

/**
 * Update all Bit.ly click stats
 */
function updateAllBitlyStats() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const bitlySheet = ss.getSheetByName('Bitly_Tracking');
    const engagementSheet = ss.getSheetByName(CONFIG.SHEETS.ENGAGEMENT_LOG);

    if (!bitlySheet) {
      Logger.log('No Bit.ly tracking sheet found');
      return 0;
    }

    const data = bitlySheet.getDataRange().getValues();
    let updated = 0;

    for (let i = 1; i < data.length; i++) {
      const bitlyLink = data[i][3];
      const postId = data[i][0];

      if (!bitlyLink) continue;

      const clicks = getBitlyClicks(bitlyLink);
      const rowNum = i + 1;

      bitlySheet.getRange(rowNum, 6).setValue(clicks.totalClicks);
      bitlySheet.getRange(rowNum, 7).setValue(new Date());

      if (clicks.countries.length > 0) {
        bitlySheet.getRange(rowNum, 8).setValue(clicks.countries[0].value);
      }
      if (clicks.referrers.length > 0) {
        bitlySheet.getRange(rowNum, 9).setValue(clicks.referrers[0].value);
        const linkedInRef = clicks.referrers.find(r =>
          r.value.toLowerCase().includes('linkedin')
        );
        if (linkedInRef) {
          bitlySheet.getRange(rowNum, 10).setValue(linkedInRef.clicks);
        }
      }

      // Calculate CTR if we have impression data
      if (engagementSheet && postId) {
        const engData = engagementSheet.getDataRange().getValues();
        for (let j = 1; j < engData.length; j++) {
          if (engData[j][0] === postId) {
            const impressions = engData[j][7] || 0;
            if (impressions > 0) {
              const ctr = (clicks.totalClicks / impressions * 100).toFixed(2) + '%';
              bitlySheet.getRange(rowNum, 11).setValue(ctr);
            }
            break;
          }
        }
      }

      updated++;

      // Rate limit - Bit.ly allows ~1000 requests/hour
      Utilities.sleep(100);
    }

    Logger.log(`Updated ${updated} Bit.ly links`);
    return updated;
  } catch (error) {
    logError('updateAllBitlyStats', error);
    return 0;
  }
}

// ═══════════════════════════════════════════════════════════════
// COMMENTS TRACKING & ANALYSIS
// ═══════════════════════════════════════════════════════════════

/**
 * Setup Comments tracking sheet
 */
function setupCommentsTracking() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  let commentsSheet = ss.getSheetByName('Comments_Log');
  if (!commentsSheet) {
    commentsSheet = ss.insertSheet('Comments_Log');
  }
  commentsSheet.clear();

  const headers = [
    'Comment ID',
    'Post ID',           // Links to Engagement_Log
    'Article Title',
    'Commenter Name',
    'Commenter Title',   // Their job title
    'Commenter Company',
    'Comment Text',
    'Comment Date',
    'Sentiment',         // Positive/Neutral/Negative
    'Comment Type',      // Question/Agreement/Disagreement/Insight/Spam
    'Quality Score',     // 1-5 (how valuable)
    'Replied',           // Yes/No
    'Reply Text',
    'Led to Connection', // Yes/No
    'Led to DM',         // Yes/No
    'Influencer',        // Yes/No (>10K followers)
    'Notes'
  ];

  commentsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  commentsSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#0077b5')  // LinkedIn blue
    .setFontColor('white');

  commentsSheet.setFrozenRows(1);
  commentsSheet.setColumnWidth(4, 150);  // Commenter name
  commentsSheet.setColumnWidth(5, 200);  // Title
  commentsSheet.setColumnWidth(7, 400);  // Comment text
  commentsSheet.setColumnWidth(13, 300); // Reply text

  // Add data validation
  const sentimentRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Positive', 'Neutral', 'Negative'], true)
    .build();
  commentsSheet.getRange(2, 9, 500, 1).setDataValidation(sentimentRule);

  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Question', 'Agreement', 'Disagreement', 'Insight', 'Personal Story', 'Spam', 'Other'], true)
    .build();
  commentsSheet.getRange(2, 10, 500, 1).setDataValidation(typeRule);

  const qualityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1 - Low', '2 - Below Avg', '3 - Average', '4 - Good', '5 - High Value'], true)
    .build();
  commentsSheet.getRange(2, 11, 500, 1).setDataValidation(qualityRule);

  const yesNoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .build();
  commentsSheet.getRange(2, 12, 500, 1).setDataValidation(yesNoRule);
  commentsSheet.getRange(2, 14, 500, 1).setDataValidation(yesNoRule);
  commentsSheet.getRange(2, 15, 500, 1).setDataValidation(yesNoRule);
  commentsSheet.getRange(2, 16, 500, 1).setDataValidation(yesNoRule);

  Logger.log('Comments tracking sheet created');
  return commentsSheet;
}

/**
 * Log a comment
 */
function logComment(postId, articleTitle, commenterName, commenterTitle, commenterCompany, commentText, sentiment, commentType) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let commentsSheet = ss.getSheetByName('Comments_Log');

    if (!commentsSheet) {
      commentsSheet = setupCommentsTracking();
    }

    const commentId = `CMT_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}_${Math.random().toString(36).substr(2, 4)}`;

    const nextRow = commentsSheet.getLastRow() + 1;
    commentsSheet.getRange(nextRow, 1, 1, 10).setValues([[
      commentId,
      postId,
      articleTitle,
      commenterName,
      commenterTitle || '',
      commenterCompany || '',
      commentText,
      new Date(),
      sentiment || 'Neutral',
      commentType || 'Other'
    ]]);

    Logger.log(`Logged comment from ${commenterName}`);
    return commentId;
  } catch (error) {
    logError('logComment', error, commentText.substring(0, 50));
    return null;
  }
}

/**
 * Analyze comments for insights
 */
function analyzeComments() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const commentsSheet = ss.getSheetByName('Comments_Log');

    if (!commentsSheet || commentsSheet.getLastRow() < 2) {
      Logger.log('No comments to analyze');
      return null;
    }

    const data = commentsSheet.getDataRange().getValues();

    // Aggregate stats
    const stats = {
      total: data.length - 1,
      sentiment: { Positive: 0, Neutral: 0, Negative: 0 },
      types: {},
      replied: 0,
      connections: 0,
      dms: 0,
      influencers: 0,
      topCommenters: {},
      topCompanies: {}
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const sentiment = row[8];
      const type = row[9];
      const replied = row[11];
      const connection = row[13];
      const dm = row[14];
      const influencer = row[15];
      const commenter = row[3];
      const company = row[5];

      if (sentiment) stats.sentiment[sentiment] = (stats.sentiment[sentiment] || 0) + 1;
      if (type) stats.types[type] = (stats.types[type] || 0) + 1;
      if (replied === 'Yes') stats.replied++;
      if (connection === 'Yes') stats.connections++;
      if (dm === 'Yes') stats.dms++;
      if (influencer === 'Yes') stats.influencers++;

      if (commenter) {
        stats.topCommenters[commenter] = (stats.topCommenters[commenter] || 0) + 1;
      }
      if (company) {
        stats.topCompanies[company] = (stats.topCompanies[company] || 0) + 1;
      }
    }

    // Create analysis report
    let analysisSheet = ss.getSheetByName('Comments_Analysis');
    if (!analysisSheet) {
      analysisSheet = ss.insertSheet('Comments_Analysis');
    }
    analysisSheet.clear();

    let row = 1;
    analysisSheet.getRange(row, 1).setValue('💬 COMMENTS ANALYSIS').setFontWeight('bold').setFontSize(16);
    row += 2;

    analysisSheet.getRange(row, 1).setValue(`Total Comments: ${stats.total}`).setFontWeight('bold');
    analysisSheet.getRange(row, 3).setValue(`Generated: ${new Date().toLocaleString()}`);
    row += 2;

    // Sentiment breakdown
    analysisSheet.getRange(row, 1).setValue('SENTIMENT').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
    row++;
    analysisSheet.getRange(row, 1).setValue('Positive');
    analysisSheet.getRange(row, 2).setValue(stats.sentiment.Positive || 0).setBackground('#d9ead3');
    row++;
    analysisSheet.getRange(row, 1).setValue('Neutral');
    analysisSheet.getRange(row, 2).setValue(stats.sentiment.Neutral || 0).setBackground('#fff2cc');
    row++;
    analysisSheet.getRange(row, 1).setValue('Negative');
    analysisSheet.getRange(row, 2).setValue(stats.sentiment.Negative || 0).setBackground('#f4cccc');
    row += 2;

    // Engagement outcomes
    analysisSheet.getRange(row, 1).setValue('ENGAGEMENT OUTCOMES').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
    row++;
    analysisSheet.getRange(row, 1).setValue('Replied to');
    analysisSheet.getRange(row, 2).setValue(stats.replied);
    analysisSheet.getRange(row, 3).setValue(`${(stats.replied / stats.total * 100).toFixed(0)}% reply rate`);
    row++;
    analysisSheet.getRange(row, 1).setValue('Led to Connection');
    analysisSheet.getRange(row, 2).setValue(stats.connections);
    row++;
    analysisSheet.getRange(row, 1).setValue('Led to DM');
    analysisSheet.getRange(row, 2).setValue(stats.dms);
    row++;
    analysisSheet.getRange(row, 1).setValue('From Influencers');
    analysisSheet.getRange(row, 2).setValue(stats.influencers);
    row += 2;

    // Comment types
    analysisSheet.getRange(row, 1).setValue('COMMENT TYPES').setFontWeight('bold').setBackground('#fbbc04').setFontColor('white');
    row++;
    const sortedTypes = Object.entries(stats.types).sort((a, b) => b[1] - a[1]);
    for (const [type, count] of sortedTypes) {
      analysisSheet.getRange(row, 1).setValue(type);
      analysisSheet.getRange(row, 2).setValue(count);
      row++;
    }
    row++;

    // Top commenters (repeat engagers = potential advocates)
    analysisSheet.getRange(row, 1).setValue('TOP COMMENTERS (Potential Advocates)').setFontWeight('bold').setBackground('#9c27b0').setFontColor('white');
    row++;
    const sortedCommenters = Object.entries(stats.topCommenters)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10);
    for (const [name, count] of sortedCommenters) {
      if (count >= 2) {
        analysisSheet.getRange(row, 1).setValue(name);
        analysisSheet.getRange(row, 2).setValue(`${count} comments`);
        row++;
      }
    }
    row++;

    // Top companies
    analysisSheet.getRange(row, 1).setValue('TOP COMPANIES ENGAGING').setFontWeight('bold').setBackground('#ea4335').setFontColor('white');
    row++;
    const sortedCompanies = Object.entries(stats.topCompanies)
      .filter(([company]) => company && company.length > 0)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10);
    for (const [company, count] of sortedCompanies) {
      analysisSheet.getRange(row, 1).setValue(company);
      analysisSheet.getRange(row, 2).setValue(`${count} comments`);
      row++;
    }

    analysisSheet.autoResizeColumns(1, 3);

    Logger.log('Comments analysis complete');
    return stats;
  } catch (error) {
    logError('analyzeComments', error);
    return null;
  }
}

/**
 * Quick comment input form
 */
function showCommentInputForm() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      h3 { color: #0077b5; }
      input, select, textarea { width: 100%; padding: 8px; margin: 5px 0 15px 0; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
      textarea { height: 80px; }
      button { background: #0077b5; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; width: 100%; }
      button:hover { background: #005885; }
      label { font-weight: bold; color: #333; }
      .row { display: flex; gap: 10px; }
      .row > div { flex: 1; }
    </style>
    <h3>💬 Log Comment</h3>
    <form id="commentForm">
      <label>Post ID (from Engagement Log)</label>
      <input type="text" id="postId" placeholder="POST_20241214_...">

      <label>Article Title</label>
      <input type="text" id="articleTitle" placeholder="What post was this on?">

      <label>Commenter Name</label>
      <input type="text" id="commenterName" placeholder="Who commented?">

      <div class="row">
        <div>
          <label>Their Title</label>
          <input type="text" id="commenterTitle" placeholder="Job title">
        </div>
        <div>
          <label>Company</label>
          <input type="text" id="commenterCompany" placeholder="Company">
        </div>
      </div>

      <label>Comment Text</label>
      <textarea id="commentText" placeholder="What did they say?"></textarea>

      <div class="row">
        <div>
          <label>Sentiment</label>
          <select id="sentiment">
            <option value="Positive">Positive</option>
            <option value="Neutral" selected>Neutral</option>
            <option value="Negative">Negative</option>
          </select>
        </div>
        <div>
          <label>Type</label>
          <select id="commentType">
            <option value="Agreement">Agreement</option>
            <option value="Question">Question</option>
            <option value="Insight">Insight</option>
            <option value="Disagreement">Disagreement</option>
            <option value="Personal Story">Personal Story</option>
            <option value="Other">Other</option>
          </select>
        </div>
      </div>

      <button type="button" onclick="submitComment()">Log Comment</button>
    </form>
    <div id="result" style="margin-top: 15px;"></div>

    <script>
      function submitComment() {
        google.script.run
          .withSuccessHandler(function(result) {
            document.getElementById('result').innerHTML = '<p style="color: green;">✓ Comment logged! ID: ' + result + '</p>';
            // Clear just the comment-specific fields
            document.getElementById('commenterName').value = '';
            document.getElementById('commenterTitle').value = '';
            document.getElementById('commenterCompany').value = '';
            document.getElementById('commentText').value = '';
          })
          .withFailureHandler(function(error) {
            document.getElementById('result').innerHTML = '<p style="color: red;">Error: ' + error + '</p>';
          })
          .logComment(
            document.getElementById('postId').value,
            document.getElementById('articleTitle').value,
            document.getElementById('commenterName').value,
            document.getElementById('commenterTitle').value,
            document.getElementById('commenterCompany').value,
            document.getElementById('commentText').value,
            document.getElementById('sentiment').value,
            document.getElementById('commentType').value
          );
      }
    </script>
  `)
    .setWidth(400)
    .setTitle('Log Comment');

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Import comments from a CSV or another sheet
 * Expects columns: PostID, ArticleTitle, CommenterName, CommenterTitle, Company, CommentText, Date
 */
function importCommentsFromSheet(sourceSheetName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sourceSheet = ss.getSheetByName(sourceSheetName);

    if (!sourceSheet) {
      Logger.log(`Source sheet "${sourceSheetName}" not found`);
      return 0;
    }

    let commentsSheet = ss.getSheetByName('Comments_Log');
    if (!commentsSheet) {
      commentsSheet = setupCommentsTracking();
    }

    const sourceData = sourceSheet.getDataRange().getValues();
    let imported = 0;

    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];

      // Skip empty rows
      if (!row[2] || !row[5]) continue; // Need commenter name and comment text

      const commentId = `CMT_IMP_${i}_${Date.now()}`;
      const nextRow = commentsSheet.getLastRow() + 1;

      commentsSheet.getRange(nextRow, 1, 1, 8).setValues([[
        commentId,
        row[0] || '',        // PostID
        row[1] || '',        // ArticleTitle
        row[2] || '',        // CommenterName
        row[3] || '',        // CommenterTitle
        row[4] || '',        // Company
        row[5] || '',        // CommentText
        row[6] || new Date() // Date
      ]]);

      imported++;
    }

    Logger.log(`Imported ${imported} comments from ${sourceSheetName}`);
    return imported;
  } catch (error) {
    logError('importCommentsFromSheet', error, sourceSheetName);
    return 0;
  }
}
