# Google Apps Scripts Collection

**Powerful Google Sheets automation scripts - AI integration, data processing, inventory management.**

---

## Scripts Overview

### 3DSellers Integration (`/3dsellers/`)
Automated inventory and listing management for multi-channel e-commerce.

| Script | Purpose |
|--------|---------|
| `3DSELLERS_V7_MASTER.gs` | Complete inventory automation system |
| `3DSELLERS_V5_FIXED.gs` | Stable production version with fixes |
| `3dsellers-production-ready.gs` | Cleaned production deployment |
| `UNIFIED_3DSELLERS_V3_COMPLETE.gs` | Multi-channel sync and updates |
| `FINAL_UNIFIED_3DSELLERS.gs` | Production-ready listing manager |
| `3D_SELLERS_IMAGE_CROPPING.gs` | Image cropping automation |
| `CREATE_CONFIG_SHEET.gs` | Setup configuration sheet |

### AI Integration (`/ai-integration/`)
Scripts that connect Google Sheets to AI APIs.

| Script | Purpose |
|--------|---------|
| `AI-ENHANCED-GOOGLE-APPS-SCRIPT.gs` | AI-powered data processing |
| `AI-METADATA-ENHANCED-SCRIPT.gs` | AI metadata extraction and enhancement |
| `AI_ART_SORTER.gs` | AI-based artwork categorization |

### News Engine (`/news-engine/`)
AI-powered news curation and content processing.

| Script | Purpose |
|--------|---------|
| `NEWS_AudienceWeighting.gs` | Score articles by audience relevance |
| `NEWS_FeedlyEmail_SlideDeck.gs` | Convert news to presentations |
| `NEWS_Pipeline_UI.gs` | News processing interface |

### Utilities (`/utilities/`)
General-purpose Google Sheets automation.

| Script | Purpose |
|--------|---------|
| `EbayAutomation.gs` | eBay listing automation |
| `ENHANCED-SALES-CHANNEL-ANALYZER.gs` | Multi-channel sales analysis |
| `CREATIVE-AUTO-RENAMER.gs` | AI-powered file renaming |
| `GALLERY-GAUNTLET-AUTO-FORMAT-SCRIPT.gs` | Gallery formatting automation |
| `ROBUST_Dual_AI_Checker_Script.gs` | Validate data with multiple AI models |
| `drive_folder_catalog.gs` | Index Google Drive folders |
| `google_sheets_uploader.gs` | Batch upload data to Sheets |
| `FRAME_VARIATIONS_SIMPLE.gs` | Frame variation generator |
| `GoogleAppsScript_ChannelSeparator.gs` | Channel data separation |
| `Google_Apps_Script_INBOUND_Monitor.gs` | Inbound monitoring alerts |
| `populate_image_urls_v3.gs` | Populate image URLs in sheets |

---

## How to Use

### 1. Open Google Sheets
1. Create a new Google Sheet
2. Go to Extensions → Apps Script
3. Copy/paste the script code
4. Save and run

### 2. Set Up API Keys
Most scripts require API keys stored in a "Config" sheet:

```
| Row | Key Name        | Value              |
|-----|-----------------|--------------------|
| 1   | OPENAI_API_KEY  | sk-your-key-here   |
| 2   | CLAUDE_API_KEY  | sk-ant-your-key    |
| 3   | GEMINI_API_KEY  | AIza-your-key      |
```

### 3. Run the Script
- Click the play button in Apps Script editor
- Or use custom menu items added by the script

---

## Features

### AI Integration
All scripts support multiple AI models:
- **OpenAI GPT-4** - General processing
- **Claude** - Analysis and reasoning
- **Gemini** - Fast responses
- **Grok** - Alternative perspective

### Multi-Channel Support
3DSellers scripts work with:
- eBay listings
- Etsy products
- Shopify inventory
- Custom platforms

### Automation
- Scheduled triggers (hourly, daily)
- Webhook integrations
- Email notifications
- Slack/Discord alerts

---

## 3DSellers Master Script

The `3DSELLERS_V7_MASTER.gs` includes:

```
FEATURES:
├── Inventory Management
│   ├── SKU generation
│   ├── Price calculations
│   ├── Stock tracking
│   └── Multi-warehouse support
├── Listing Creation
│   ├── Title optimization
│   ├── Description generation
│   ├── SEO keywords
│   └── Category mapping
├── Image Processing
│   ├── URL validation
│   ├── Thumbnail generation
│   └── Gallery ordering
├── AI Analysis
│   ├── Product classification
│   ├── Condition assessment
│   └── Price recommendation
└── Reporting
    ├── Sales analytics
    ├── Inventory alerts
    └── Performance metrics
```

---

## News Engine

Process news articles through AI evaluation:

```
News Article → Fetch Content → AI Scoring → Audience Weighting → Output

Scoring criteria:
- Relevance (1-10)
- Timeliness (1-10)
- Credibility (1-10)
- Engagement potential (1-10)
```

---

## Setup Requirements

1. **Google Account** with Sheets access
2. **API Keys** for AI services (optional)
3. **Google Cloud Project** (for advanced features)

---

## Security Notes

- Never commit API keys to scripts
- Use Script Properties or Config sheets
- Limit script permissions appropriately
- Review OAuth scopes

---

## Installation

```bash
# Clone this repo
git clone https://github.com/jjshay/google-apps-scripts.git

# Open any .gs file
# Copy contents to Google Apps Script editor
# Configure API keys in Config sheet
# Run!
```

---

## License

MIT - Use freely in your Google Sheets!
