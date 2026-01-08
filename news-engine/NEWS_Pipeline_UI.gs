/**
 * NEWS ENGINE - UNIFIED PIPELINE WITH PROGRESS UI
 * Consolidates all AI ratings into single "AI Analysis" step
 */

// ═══════════════════════════════════════════════════════════════
// PIPELINE STEPS CONFIGURATION
// ═══════════════════════════════════════════════════════════════

const PIPELINE_STEPS = [
  { id: 'feedly', name: 'Fetch Feedly', duration: 20, fn: 'stepFetchFeedly' },
  { id: 'dedupe', name: 'Dedupe & Classify', duration: 15, fn: 'stepDedupeClassify' },
  { id: 'ai', name: 'AI Analysis', duration: 45, fn: 'stepAIAnalysis' },
  { id: 'weights', name: 'Apply Weights', duration: 10, fn: 'stepApplyWeights' },
  { id: 'queue', name: 'Build Queue', duration: 10, fn: 'stepBuildQueue' },
  { id: 'reports', name: 'Save & Generate Reports', duration: 15, fn: 'stepGenerateReports' }
];

// ═══════════════════════════════════════════════════════════════
// MAIN PIPELINE WITH PROGRESS UI
// ═══════════════════════════════════════════════════════════════

/**
 * Run full pipeline with visual progress
 */
function runFullIngestionWithUIDetails() {
  const html = HtmlService.createHtmlOutput(getPipelineHTML())
    .setWidth(400)
    .setHeight(500);

  SpreadsheetApp.getUi().showModelessDialog(html, 'NEWS Pipeline');
}

/**
 * Get pipeline status (called from HTML)
 */
function getPipelineStatus() {
  const props = PropertiesService.getScriptProperties();
  return JSON.parse(props.getProperty('PIPELINE_STATUS') || '{}');
}

/**
 * Run pipeline step by step (called from HTML)
 */
function runPipelineStep(stepId) {
  const step = PIPELINE_STEPS.find(s => s.id === stepId);
  if (!step) return { error: 'Unknown step' };

  const props = PropertiesService.getScriptProperties();
  const status = JSON.parse(props.getProperty('PIPELINE_STATUS') || '{}');

  // Mark step as running
  status[stepId] = { status: 'running', startTime: Date.now() };
  props.setProperty('PIPELINE_STATUS', JSON.stringify(status));

  try {
    // Execute the step function
    const result = this[step.fn]();

    // Mark step as complete
    status[stepId] = {
      status: 'complete',
      result: result,
      duration: Math.round((Date.now() - status[stepId].startTime) / 1000)
    };
    props.setProperty('PIPELINE_STATUS', JSON.stringify(status));

    return { success: true, result: result };
  } catch (error) {
    status[stepId] = { status: 'error', error: error.message };
    props.setProperty('PIPELINE_STATUS', JSON.stringify(status));
    return { error: error.message };
  }
}

/**
 * Reset pipeline status
 */
function resetPipelineStatus() {
  PropertiesService.getScriptProperties().setProperty('PIPELINE_STATUS', '{}');
}

/**
 * Run entire pipeline sequentially
 */
function runEntirePipeline() {
  resetPipelineStatus();

  const results = {};
  for (const step of PIPELINE_STEPS) {
    const result = runPipelineStep(step.id);
    results[step.id] = result;
    if (result.error) break;
  }

  return results;
}


// ═══════════════════════════════════════════════════════════════
// INDIVIDUAL STEP FUNCTIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Step 1: Fetch from Feedly
 */
function stepFetchFeedly() {
  // Check for Feedly integration
  if (typeof checkFeedlyForNewArticles === 'function') {
    return checkFeedlyForNewArticles();
  }

  // Fallback: fetch RSS feeds
  if (typeof fetchAllRSSFeeds === 'function') {
    return fetchAllRSSFeeds();
  }

  return { articles: 0, message: 'Feedly not configured' };
}

/**
 * Step 2: Dedupe & Classify
 */
function stepDedupeClassify() {
  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch(e) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var sheet = ss.getSheetByName('NEWS IN') || ss.getSheets()[0];

  if (!sheet) return { deduped: 0, classified: 0 };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const titleCol = headers.indexOf('Title') !== -1 ? headers.indexOf('Title') : 4;
  const topicCol = headers.indexOf('Topic') !== -1 ? headers.indexOf('Topic') : 0;

  // Track seen titles for deduplication
  const seenTitles = new Set();
  const rowsToDelete = [];
  let classified = 0;

  for (let i = 1; i < data.length; i++) {
    const title = String(data[i][titleCol] || '').toLowerCase().trim();

    // Check for duplicate
    if (seenTitles.has(title)) {
      rowsToDelete.push(i + 1);
    } else {
      seenTitles.add(title);

      // Auto-classify if no topic
      if (!data[i][topicCol]) {
        const topic = autoClassifyTopic(data[i][titleCol]);
        if (topic) {
          sheet.getRange(i + 1, topicCol + 1).setValue(topic);
          classified++;
        }
      }
    }
  }

  // Delete duplicates (reverse order to preserve row numbers)
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  return { deduped: rowsToDelete.length, classified: classified };
}

/**
 * Auto-classify topic based on title keywords
 */
function autoClassifyTopic(title) {
  if (!title) return null;
  const text = title.toLowerCase();

  const topics = {
    'AI': ['ai', 'artificial intelligence', 'machine learning', 'chatgpt', 'llm', 'openai', 'claude', 'gemini', 'gpt'],
    'Tech': ['software', 'startup', 'tech', 'silicon valley', 'app', 'platform', 'saas', 'cloud'],
    'Business': ['ceo', 'company', 'revenue', 'acquisition', 'ipo', 'market', 'business', 'enterprise'],
    'Finance': ['stock', 'investment', 'crypto', 'bitcoin', 'banking', 'fed', 'interest rate'],
    'Leadership': ['leadership', 'management', 'culture', 'team', 'hiring', 'ceo'],
    'Product': ['product', 'launch', 'feature', 'user', 'design', 'ux']
  };

  for (const [topic, keywords] of Object.entries(topics)) {
    if (keywords.some(kw => text.includes(kw))) {
      return topic;
    }
  }

  return 'General';
}

/**
 * Step 3: AI Analysis (Consolidated)
 * Uses available AI (Gemini preferred) to rate and analyze articles
 */
function stepAIAnalysis() {
  var props = PropertiesService.getScriptProperties();
  var geminiKey = props.getProperty('GEMINI_API_KEY');

  if (!geminiKey) {
    return { analyzed: 0, message: 'No AI key configured' };
  }

  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch(e) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  var sheet = ss.getSheetByName('NEWS IN') || ss.getSheets()[0];

  if (!sheet) return { analyzed: 0 };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find relevant columns
  const titleCol = findColumn(headers, ['Title', 'TITLE']);
  const summaryCol = findColumn(headers, ['Summary', 'SUMMARY', 'Description']);
  const aiScoreCol = findColumn(headers, ['AI_AVG', 'AI Score', 'Score']);
  const soWhatCol = findColumn(headers, ['So What', 'SO_WHAT']);
  const myTakeCol = findColumn(headers, ['My Take', 'MY_TAKE']);

  let analyzed = 0;
  const batchSize = 10; // Process in batches to avoid timeout

  for (let i = 1; i < Math.min(data.length, batchSize + 1); i++) {
    const title = data[i][titleCol];
    const summary = data[i][summaryCol] || '';

    // Skip if already analyzed
    if (data[i][aiScoreCol]) continue;

    try {
      const analysis = analyzeArticleWithAI(geminiKey, title, summary);

      if (analysis) {
        if (aiScoreCol >= 0 && analysis.score) {
          sheet.getRange(i + 1, aiScoreCol + 1).setValue(analysis.score);
        }
        if (soWhatCol >= 0 && analysis.soWhat) {
          sheet.getRange(i + 1, soWhatCol + 1).setValue(analysis.soWhat);
        }
        if (myTakeCol >= 0 && analysis.myTake) {
          sheet.getRange(i + 1, myTakeCol + 1).setValue(analysis.myTake);
        }
        analyzed++;
      }
    } catch (e) {
      Logger.log('AI analysis error for row ' + i + ': ' + e.message);
    }

    // Rate limiting
    Utilities.sleep(1000);
  }

  return { analyzed: analyzed };
}

/**
 * Analyze single article with AI
 */
function analyzeArticleWithAI(apiKey, title, summary) {
  const prompt = `Analyze this news article for LinkedIn posting potential:

Title: ${title}
Summary: ${summary}

Respond in JSON format:
{
  "score": <1-10 rating for LinkedIn engagement potential>,
  "soWhat": "<2 sentences: why this matters to business/tech professionals>",
  "myTake": "<2 sentences: a thoughtful perspective on this news>",
  "topics": ["<relevant topic tags>"]
}`;

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.7, maxOutputTokens: 500 }
      }),
      muteHttpExceptions: true
    });

    const data = JSON.parse(response.getContentText());

    if (data.error) {
      throw new Error(data.error.message);
    }

    const text = data.candidates[0].content.parts[0].text;

    // Extract JSON from response
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }

    return null;
  } catch (e) {
    Logger.log('AI API error: ' + e.message);
    return null;
  }
}

/**
 * Helper: Find column by possible names
 */
function findColumn(headers, possibleNames) {
  for (const name of possibleNames) {
    const idx = headers.findIndex(h =>
      String(h).toLowerCase() === name.toLowerCase()
    );
    if (idx >= 0) return idx;
  }
  return -1;
}

/**
 * Step 4: Apply Weights
 */
function stepApplyWeights() {
  if (typeof applyAudienceWeights === 'function') {
    return applyAudienceWeights();
  }
  return { weighted: 0 };
}

/**
 * Step 5: Build Distribution Queue
 */
function stepBuildQueue() {
  if (typeof createDistributionQueue === 'function') {
    return createDistributionQueue();
  }
  return { queued: 0 };
}

/**
 * Step 6: Generate Reports
 */
function stepGenerateReports() {
  const results = {};

  // Create quality report if function exists
  if (typeof createQualityReport === 'function') {
    try {
      createQualityReport();
      results.qualityReport = true;
    } catch (e) {
      results.qualityReport = false;
    }
  }

  // Generate insights if function exists
  if (typeof generatePerformanceInsights === 'function') {
    try {
      generatePerformanceInsights();
      results.insights = true;
    } catch (e) {
      results.insights = false;
    }
  }

  return results;
}


// ═══════════════════════════════════════════════════════════════
// PROGRESS UI HTML
// ═══════════════════════════════════════════════════════════════

function getPipelineHTML() {
  var stepsJson = JSON.stringify(PIPELINE_STEPS);

  var html = '<!DOCTYPE html>' +
    '<html><head><base target="_top"><style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Segoe UI", Arial, sans-serif; background: #1a1a2e; color: #fff; padding: 20px; }' +
    'h2 { margin-bottom: 20px; color: #4361ee; font-size: 18px; }' +
    '.step { display: flex; align-items: center; padding: 12px 15px; margin-bottom: 8px; background: #2d2d44; border-radius: 8px; transition: all 0.3s ease; }' +
    '.step.running { background: #3d3d5c; border-left: 3px solid #4361ee; }' +
    '.step.complete { background: #1e3a2f; border-left: 3px solid #10b981; }' +
    '.step.error { background: #3a1e1e; border-left: 3px solid #ef4444; }' +
    '.step-icon { width: 24px; height: 24px; margin-right: 12px; display: flex; align-items: center; justify-content: center; font-size: 14px; }' +
    '.step-icon.pending { color: #666; } .step-icon.running { color: #4361ee; } .step-icon.complete { color: #10b981; } .step-icon.error { color: #ef4444; }' +
    '.spinner { width: 16px; height: 16px; border: 2px solid #4361ee; border-top-color: transparent; border-radius: 50%; animation: spin 1s linear infinite; }' +
    '@keyframes spin { to { transform: rotate(360deg); } }' +
    '.step-name { flex: 1; font-size: 14px; } .step-time { font-size: 12px; color: #888; } .step-result { font-size: 11px; color: #10b981; margin-left: 10px; }' +
    '.buttons { margin-top: 20px; display: flex; gap: 10px; }' +
    'button { flex: 1; padding: 12px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; transition: all 0.2s; }' +
    '.btn-primary { background: #4361ee; color: white; } .btn-primary:hover { background: #3651de; } .btn-primary:disabled { background: #666; cursor: not-allowed; }' +
    '.btn-secondary { background: #2d2d44; color: #fff; border: 1px solid #444; } .btn-secondary:hover { background: #3d3d5c; }' +
    '.status-bar { margin-top: 15px; padding: 10px; background: #2d2d44; border-radius: 6px; font-size: 12px; color: #888; }' +
    '</style></head><body>' +
    '<h2>NEWS Pipeline</h2>' +
    '<div id="steps"></div>' +
    '<div class="buttons">' +
    '<button class="btn-primary" id="runBtn" onclick="runPipeline()">Run Pipeline</button>' +
    '<button class="btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '</div>' +
    '<div class="status-bar" id="status">Ready to run</div>' +
    '<script>' +
    'var steps = ' + stepsJson + ';' +
    'var stepStatus = {};' +
    'var isRunning = false;' +
    'function renderSteps() {' +
    '  var container = document.getElementById("steps");' +
    '  container.innerHTML = steps.map(function(step) {' +
    '    var status = stepStatus[step.id] || {};' +
    '    var statusClass = status.status || "pending";' +
    '    var icon = "○";' +
    '    if (status.status === "running") icon = "<div class=\\"spinner\\"></div>";' +
    '    if (status.status === "complete") icon = "✓";' +
    '    if (status.status === "error") icon = "✗";' +
    '    var timeText = "~" + step.duration + "s";' +
    '    if (status.duration) timeText = status.duration + "s";' +
    '    var resultText = "";' +
    '    if (status.result && typeof status.result === "object") {' +
    '      var vals = Object.values(status.result).filter(function(v) { return typeof v === "number"; });' +
    '      if (vals.length) resultText = vals.join(", ");' +
    '    }' +
    '    return "<div class=\\"step " + statusClass + "\\">" +' +
    '      "<div class=\\"step-icon " + statusClass + "\\">" + icon + "</div>" +' +
    '      "<div class=\\"step-name\\">" + step.name + "</div>" +' +
    '      "<div class=\\"step-time\\">" + timeText + "</div>" +' +
    '      (resultText ? "<div class=\\"step-result\\">" + resultText + "</div>" : "") +' +
    '      "</div>";' +
    '  }).join("");' +
    '}' +
    'function runPipeline() {' +
    '  if (isRunning) return;' +
    '  isRunning = true;' +
    '  document.getElementById("runBtn").disabled = true;' +
    '  document.getElementById("status").textContent = "Running pipeline...";' +
    '  stepStatus = {};' +
    '  renderSteps();' +
    '  runNextStep(0);' +
    '}' +
    'function runNextStep(index) {' +
    '  if (index >= steps.length) {' +
    '    isRunning = false;' +
    '    document.getElementById("runBtn").disabled = false;' +
    '    document.getElementById("status").textContent = "Pipeline complete!";' +
    '    return;' +
    '  }' +
    '  var step = steps[index];' +
    '  stepStatus[step.id] = { status: "running" };' +
    '  renderSteps();' +
    '  document.getElementById("status").textContent = "Running: " + step.name;' +
    '  google.script.run' +
    '    .withSuccessHandler(function(result) {' +
    '      if (result.error) {' +
    '        stepStatus[step.id] = { status: "error", error: result.error };' +
    '        document.getElementById("status").textContent = "Error: " + result.error;' +
    '        isRunning = false;' +
    '        document.getElementById("runBtn").disabled = false;' +
    '      } else {' +
    '        stepStatus[step.id] = { status: "complete", result: result.result, duration: step.duration };' +
    '        renderSteps();' +
    '        runNextStep(index + 1);' +
    '      }' +
    '    })' +
    '    .withFailureHandler(function(e) {' +
    '      stepStatus[step.id] = { status: "error", error: e.message };' +
    '      document.getElementById("status").textContent = "Error: " + e.message;' +
    '      isRunning = false;' +
    '      document.getElementById("runBtn").disabled = false;' +
    '      renderSteps();' +
    '    })' +
    '    .runPipelineStep(step.id);' +
    '}' +
    'renderSteps();' +
    '</script></body></html>';

  return html;
}


// ═══════════════════════════════════════════════════════════════
// MENU INTEGRATION
// ═══════════════════════════════════════════════════════════════

/**
 * Add pipeline menu
 */
function onOpenPipeline() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('NEWS Pipeline')
    .addItem('Run Full Pipeline', 'runFullIngestionWithUIDetails')
    .addSeparator()
    .addItem('Fetch Feedly Only', 'stepFetchFeedly')
    .addItem('Dedupe & Classify Only', 'stepDedupeClassify')
    .addItem('AI Analysis Only', 'stepAIAnalysis')
    .addItem('Apply Weights Only', 'stepApplyWeights')
    .addItem('Build Queue Only', 'stepBuildQueue')
    .addItem('Generate Reports Only', 'stepGenerateReports')
    .addSeparator()
    .addItem('Setup Required Sheets', 'setupNewsSheets')
    .addItem('Configure AI Keys', 'showAIKeySetup')
    .addToUi();
}

/**
 * AI Key setup dialog
 */
function showAIKeySetup() {
  var htmlContent = '<style>' +
    'body { font-family: Arial; padding: 20px; }' +
    'label { display: block; margin-top: 15px; font-weight: bold; }' +
    'input { width: 100%; padding: 8px; margin-top: 5px; }' +
    'button { margin-top: 20px; padding: 10px 20px; background: #4361ee; color: white; border: none; cursor: pointer; }' +
    '.note { font-size: 12px; color: #666; margin-top: 5px; }' +
    '</style>' +
    '<h3>AI Configuration</h3>' +
    '<label>Gemini API Key</label>' +
    '<input type="text" id="gemini" placeholder="Get from aistudio.google.com/app/apikey">' +
    '<div class="note">Free tier: 15 requests/minute</div>' +
    '<button onclick="save()">Save</button>' +
    '<script>' +
    'function save() {' +
    '  google.script.run.withSuccessHandler(function() {' +
    '    alert("Saved!");' +
    '    google.script.host.close();' +
    '  }).saveGeminiKey(document.getElementById("gemini").value);' +
    '}' +
    '</script>';

  var html = HtmlService.createHtmlOutput(htmlContent).setWidth(400).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'AI Setup');
}

function saveGeminiKey(key) {
  if (key) {
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  }
}

/**
 * Setup required sheets if they don't exist
 */
function setupNewsSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create NEWS IN sheet if it doesn't exist
  var newsIn = ss.getSheetByName('NEWS IN');
  if (!newsIn) {
    newsIn = ss.insertSheet('NEWS IN');
    newsIn.getRange(1, 1, 1, 15).setValues([[
      'Topic', 'Select', 'Score', 'Date', 'Title', 'Link', 'Summary',
      'AI Score', 'So What', 'My Take', 'Whats Next', 'Source',
      'Industry Weight', 'Function Weight', 'Final Score'
    ]]);
    newsIn.getRange(1, 1, 1, 15).setFontWeight('bold').setBackground('#4361ee').setFontColor('white');
    newsIn.setFrozenRows(1);
  }

  // Create Distribution_Queue sheet
  var queue = ss.getSheetByName('Distribution_Queue');
  if (!queue) {
    queue = ss.insertSheet('Distribution_Queue');
    queue.getRange(1, 1, 1, 8).setValues([[
      'Priority', 'Title', 'Topic', 'Score', 'Target Segment', 'Scheduled', 'Status', 'Link'
    ]]);
    queue.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#10b981').setFontColor('white');
    queue.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert('Setup Complete', 'NEWS IN and Distribution_Queue sheets created!', SpreadsheetApp.getUi().ButtonSet.OK);
}
