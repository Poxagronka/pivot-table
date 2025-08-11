# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Language Preference
**ВАЖНО: Всегда отвечай на русском языке пользователю. All responses to the user must be in Russian language.**

## Development Workflow

### Deployment Commands
- **Deploy to Google Apps Script**: `clasp push --force` (pushes all local changes to GAS)
- **Deploy and Sync**: `./sync_gas.sh` (interactive script that handles GAS deployment and git sync)
- **Open in Browser**: `clasp open` (opens the GAS project in web editor)

### Testing Commands
No automated tests - this is a Google Apps Script project. Testing is done manually through the Google Sheets interface.

## Architecture Overview

This is a Google Apps Script project that creates campaign reporting dashboards in Google Sheets. It fetches data from Appodeal GraphQL API and generates hierarchical pivot tables with week-over-week analytics across multiple advertising networks.

### Core Components

1. **Project Configuration System** (`01_Config.js`)
   - Manages 9 different advertising networks: TRICKY, MOLOCO, REGULAR, GOOGLE_ADS, APPLOVIN, MINTEGRAL, INCENT, INCENT_TRAFFIC, OVERALL
   - Each project has unique API filters, sheet names, and target eROAS values
   - Global `CURRENT_PROJECT` variable controls active context
   - Settings loaded dynamically from Google Sheets

2. **GraphQL API Integration** (`05_ApiClient.js`)
   - Single API endpoint: `https://app.appodeal.com/graphql`
   - Uses bearer token authentication (stored in settings sheet)
   - Project-specific filtering by attribution networks and campaign patterns
   - Different date dimensions: `INSTALL_DATE` vs `DATE` based on project type

3. **Data Processing Pipeline** (`15_TableBuilder.js`, `16_RowGrouping.js`)
   - Hierarchical grouping: App → Week → Campaign
   - Week-over-week (WoW) calculations with growth status categorization
   - Campaign name parsing and GEO extraction (project-specific logic)
   - Performance optimizations with caching layers

4. **Comment Persistence** (`02_CommentCache.js`)
   - Multi-level comment storage: week-level and campaign-level
   - Cached in separate Google Sheets to survive data refreshes
   - Automatic sync during report generation

5. **Sheet Formatting** (`07_SheetFormatting.js`)
   - Creates nested row groups with conditional formatting
   - Project-specific color coding based on eROAS targets
   - Growth status indicators with emoji categories

6. **Automation System** (`10_AutoFunctions.js`)
   - Daily triggers: cache comments at 3AM CET, update data at 5AM CET
   - Batch processing for all 9 projects
   - Error handling and status tracking

### Key Technical Details

#### Project Switching Pattern
```javascript
setCurrentProject('TRICKY');
const config = getCurrentConfig();
const apiConfig = getCurrentApiConfig();
```

#### Data Structure Hierarchy
```
API Response → processApiData() → Hierarchical Structure:
  └─ App Name
     └─ Week (Monday-Sunday)
        └─ Campaign
           └─ Metrics (spend, installs, eROAS, etc.)
```

#### Comment Storage Keys
- Week comments: `AppName|||WeekRange|||WEEK|||WEEK`
- Campaign comments: `AppName|||WeekRange|||CampaignId|||SourceApp`

#### Growth Status Algorithm
Uses configurable thresholds to categorize campaigns:
- 🟢 Healthy Growth, Efficiency Improvement  
- 🔴 Inefficient Growth
- 🟠 Declining Efficiency
- 🔵 Scaling Down (Efficient/Moderate/Problematic)
- 🟡 Moderate Growth/Decline, Minimal Growth
- ⚪ Stable

### File Execution Order
Files must load in sequence due to dependencies:
1. `01_Config.js` - Defines global constants and project configurations
2. `08_Utilities.js` - Date/string utilities used throughout
3. `05_ApiClient.js` - API integration (depends on Config)
4. `06_Analytics.js` - WoW calculations
5. All other files can load in any order

### Important Constraints

- **6-minute execution limit** for Google Apps Script functions
- **Bearer token required** - stored in separate Settings sheet for security
- **No package.json** - this is pure Google Apps Script, not Node.js
- **Sheet-based configuration** - all settings stored in Google Sheets, not files
- **Manual testing only** - no automated test framework

### Common Development Patterns

When adding new projects:
1. Add to `PROJECTS` object in `01_Config.js`
2. Add to `MENU_PROJECTS` array in `04_MenuFunctions.js`
3. Test campaign name parsing logic in your specific project context
4. Configure project-specific metrics and conditional formatting rules

When modifying API queries:
- Update `UNIFIED_MEASURES` array in Config
- Adjust data processing indices in `processApiData()`
- Update table headers and column widths in `TABLE_CONFIG`

### Debug Workflow
- Use `debugReportGeneration()` for full diagnostics
- Check "Debug" sheet tab for step-by-step API response analysis
- Campaign filtering verification shows matched/unmatched patterns
- Each project has individual debug functions available through menu