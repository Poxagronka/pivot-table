# Campaign Report Google Apps Script - Technical Documentation

## Architecture Overview

Multi-project campaign reporting system built on Google Apps Script that fetches data from Appodeal GraphQL API, processes it into hierarchical reports with week-over-week analytics, and maintains persistent comment system with automated caching.

## Core Design Principles

1. **Project Isolation**: Each project (Tricky/Moloco/Regular/Google_Ads/Applovin/Mintegral) operates independently with own sheet, cache, and configuration
2. **State Management**: Uses ScriptProperties for persistent settings and hidden sheets for comment storage
3. **Batch Processing**: API calls fetch all data in single request, then process client-side to minimize quota usage
4. **Progressive Enhancement**: Core functionality works without automation; triggers add convenience

## Technical Architecture

### Data Flow
```
API Request -> Raw JSON -> Processing -> Hierarchical Structure -> Sheet Rendering
    |              |            |                                      |
Bearer Token   Week Aggregation                           Group Creation
GraphQL Query  Campaign Filtering                    Conditional Formatting
              WoW Calculations                         Comment Restoration
```

### Project Configuration Structure
```javascript
PROJECTS = {
 TRICKY: {
   SHEET_NAME: 'Tricky',
   ATTRIBUTION_NETWORK_HID: ["234187180623265792"],
   ATTRIBUTION_CAMPAIGN_SEARCH: "/tricky/i"  // Regex inclusion
 },
 MOLOCO: {
   SHEET_NAME: 'Moloco', 
   ATTRIBUTION_NETWORK_HID: ["445856363109679104"],
   ATTRIBUTION_CAMPAIGN_SEARCH: null  // No filter - takes all
 },
 REGULAR: {
   SHEET_NAME: 'Regular',
   ATTRIBUTION_NETWORK_HID: ["234187180623265792"],
   ATTRIBUTION_CAMPAIGN_SEARCH: "!/tricky/i"  // Regex exclusion
 },
 GOOGLE_ADS: {
   SHEET_NAME: 'Google_Ads',
   ATTRIBUTION_NETWORK_HID: ["378302368699121664"],
   ATTRIBUTION_CAMPAIGN_SEARCH: "!/test_creo|creo_test|SL|TL|RnD|adq/i"
 },
 APPLOVIN: {
   SHEET_NAME: 'Applovin',
   ATTRIBUTION_NETWORK_HID: ["261208778387488768"],
   ATTRIBUTION_CAMPAIGN_SEARCH: "!/test_creo|creo_test|SL|TL|RnD|adq/i"
 },
 MINTEGRAL: {
   SHEET_NAME: 'Mintegral',
   ATTRIBUTION_NETWORK_HID: ["756604737398243328"],
   ATTRIBUTION_CAMPAIGN_SEARCH: null  // No filter - takes all
 }
}
```

## Key Technical Components

### 1. GraphQL API Integration

**Query Structure**: Single query fetches all metrics with grouping by INSTALL_DATE (or DATE for Google_Ads/Applovin), ATTRIBUTION_CAMPAIGN_HID, and APP. Uses RichStats operation with project-specific measures.

**Key Implementation Details**:
- Headers include trace-id for debugging and proper referrer for CORS
- Filters built dynamically based on project configuration
- Date range excludes current incomplete week
- Response parsing handles both UaCampaign and StatsValue structures
- Project-specific date dimensions (DATE vs INSTALL_DATE)

**Project-Specific Metrics**:
- **Traditional projects** (Tricky/Moloco/Regular/Mintegral): CPI, installs, IPM, spend, ROAS D-1, eARPU 365d, eROAS 365d, eProfit 730d
- **Google_Ads/Applovin**: CPI, installs, spend, retention rate D-1, ROAS D-1, retention rate D-7, eROAS 365d, eProfit 730d

### 2. Data Processing Pipeline

**processApiData() function**:
- Groups data by app → week → campaign hierarchy
- Calculates Monday-Sunday week boundaries
- Extracts GEO from campaign names using project-specific pattern matching
- Source app extraction differs by project (full name for Moloco/Mintegral, parsed for Tricky, campaign name for others)
- Filters out current week data as incomplete
- Filters campaigns with spend ≤ 0

**GEO Extraction Logic**:
- **Tricky/Regular**: Uses pipe-delimited patterns (| USA |, | MEX |, etc.)
- **Other projects**: Pattern matching for geo codes (US, RU, UK, etc.) with priority ordering

### 3. Week-over-Week Analytics

**calculateWoWMetrics()** generates two metric sets:
- **campaignWoW**: Campaign-level WoW changes keyed by campaignId_weekStart
- **appWeekWoW**: App-level weekly aggregates keyed by appName_weekStart

**Growth Status Algorithm** considers:
- Profit sign transitions (negative to positive = Healthy Growth)
- Spend vs profit correlation with configurable thresholds
- Multiple growth categories with emoji indicators:
  - 🟢 Healthy Growth, Efficiency Improvement
  - 🔴 Inefficient Growth
  - 🟠 Declining Efficiency
  - 🔵 Scaling Down (Efficient/Moderate/Problematic)
  - 🟡 Moderate Growth/Decline, Minimal Growth
  - ⚪ Stable

### 4. Comment Persistence System

**Two-level comment storage**:
- **Week-level**: Keyed by `AppName|||WeekRange|||WEEK|||WEEK`
- **Campaign-level**: Keyed by `AppName|||WeekRange|||CampaignId|||SourceApp`

**Cache sheet structure**: [AppName, WeekRange, CampaignId, SourceApp, Comment, LastUpdated]

**Optimized Sync Process**:
1. ~~Expand all groups to access hidden rows~~ **REMOVED** - `getDataRange()` reads all data including collapsed rows
2. Read comments from main sheet
3. Update cache only if new comment is longer (handles appending)
4. Collapse groups after caching (recursive method for reliability)

### 5. Sheet Formatting Engine

**Hierarchical grouping implementation**:
- Uses `shiftRowGroupDepth(1)` to create nested groups
- Groups created bottom-up: campaigns first, then weeks, then apps
- Collapse uses recursive approach to handle deep nesting

**Conditional formatting rules**:
- eROAS colored based on project-specific targets (stored in ScriptProperties)
- WoW percentages use positive/negative coloring
- Growth status uses emoji indicators with corresponding colors
- Project-specific column layouts and formatting

**Project-Specific Headers**:
- **Traditional**: Level, Week Range/Source App, ID, GEO, Spend, Spend WoW %, Installs, CPI, ROAS D-1, IPM, eARPU 365d, eROAS 365d, eProfit 730d, eProfit 730d WoW %, Growth Status, Comments
- **Google_Ads/Applovin**: Level, Week Range/Source App, ID, GEO, Spend, Spend WoW %, Installs, CPI, ROAS D-1, RR D-1, RR D-7, eROAS 365d, eProfit 730d, eProfit 730d WoW %, Growth Status, Comments

### 6. Automation System

**Time-based triggers**:
- `autoCacheAllProjects`: Daily at 3:00 AM CET
- `autoUpdateAllProjects`: Daily at 5:00 AM CET

**Auto-cache process**:
1. ~~Expand groups silently~~ **OPTIMIZED** - No longer needed
2. Sync comments to cache for all 6 projects
3. Collapse all groups recursively
4. No UI interruption

**Auto-update process**:
1. Find earliest date in existing data
2. Fetch from earliest to last complete Saturday
3. Clear and regenerate maintaining comments
4. Apply cached comments back to sheet

### 7. ~~Progress Management~~ **REMOVED**

~~ProgressManager class~~ **REMOVED** - All functions now work without progress indicators for improved performance and reliability.

## Critical Implementation Details

### Campaign Name Parsing

**Tricky campaigns**:
- Extract source app after "=" sign
- Remove "autobudget" suffix
- Handle multiple "subj" occurrences

**Moloco/Mintegral campaigns**:
- Keep full campaign name (APD_ prefix)
- No parsing applied

**Regular/Google_Ads/Applovin campaigns**:
- Return campaign name as-is
- No modification applied

### Date Handling

**Week boundaries**:
- Monday = start (day 1)
- Sunday = end (day 0 or 7)
- Current week always excluded from processing

**Date range calculation**:
- `getDateRange(days)`: Inclusive range from (today - days + 1) to today
- Custom ranges: Direct YYYY-MM-DD format validation
- Project-specific date dimensions (DATE vs INSTALL_DATE)

### Error Recovery

**API failures**:
- Retry logic not implemented (rely on manual retry)
- Error messages preserved in debug sheet
- Project-specific error handling

**Sheet operations**:
- Silent recreation on clear to avoid corruption
- Try-catch around group operations (groups may not exist)

**Comment sync**:
- Only update if comment is longer (prevents data loss)
- Hyperlink extraction from formula strings
- Optimized to work without expanding groups

### Performance Optimizations

**Batch operations**:
- Single API call per report generation
- Bulk sheet writes using `getRange().setValues()`
- Conditional formatting applied as rule sets

**Memory management**:
- Process data in streaming fashion where possible
- Clear references after use
- Limit debug output to prevent memory overflow

**Comment caching optimization**:
- Removed group expansion requirement
- Direct data reading from collapsed sheets
- Faster cache operations

### Security Considerations

**Token storage**:
- Shared Bearer token in code (necessary for automation)
- No user credentials stored
- Read-only API access

**Data isolation**:
- Each project has separate sheets
- No cross-project data leakage
- Hidden sheets for sensitive cache data

## File Structure

### Core Files
1. **01_Config.js** - Project configurations, API settings, target eROAS values
2. **02_CommentCache.js** - Comment persistence system with multi-level support
3. **~~03_ProgressManager.js~~** - **REMOVED**
4. **04_MenuFunctions.js** - UI menu system with smart wizards
5. **05_ApiClient.js** - GraphQL API integration with project-specific handling
6. **06_Analytics.js** - WoW calculations and growth status logic
7. **07_SheetFormatting.js** - Table creation and conditional formatting
8. **08_Utilities.js** - Date, string, and sheet utility functions
9. **09_debug.js** - Debugging tools and diagnostics
10. **10_AutoFunctions.js** - Automation triggers and scheduling
11. **11_SettingsFunctions.js** - Configuration management UI

### Dependencies
**Execution order matters**:
1. Config must load first (defines globals)
2. Utilities needed by most other files
3. Menu functions can reference all others
4. Auto functions depend on cache and analytics

**Cross-file references**:
- `CURRENT_PROJECT` global maintains state
- `setCurrentProject()` switches context
- Config getters provide dynamic values

## Menu System

### Smart Report Wizard
- **All Projects Together** - Generate reports for all 6 projects
- **Single Project** - Select specific project
- **Custom Selection** - Choose multiple projects
- **Period Options** - 30/60/90 days, custom days, date ranges

### Settings Hub
- **Target eROAS Settings** - Per-project eROAS targets
- **Growth Status Thresholds** - Customizable growth criteria
- **Project Overview** - Configuration summary
- **Comments Management** - Save/restore comments
- **Clear Data** - Project-specific or all data
- **API Health Check** - Connection testing
- **Debug Tools** - Troubleshooting utilities
- **System Status** - Automation monitoring

### GitHub Integration
- Direct link to repository in menu
- Version control and collaboration support

## Limitations and Constraints

### Google Apps Script limits
- 6-minute execution timeout
- URL fetch quota (important for large date ranges)
- Trigger timing precision (~15 minutes)

### Sheet limits
- Maximum groups depth affects nesting
- Conditional format rules have maximum count
- Hidden sheets still count against sheet limit

### API constraints
- GraphQL query complexity limits
- Rate limiting not documented but observed
- Date range affects response size significantly

## Maintenance Notes

### Adding new projects
1. Add to `PROJECTS` object in Config.js
2. Add to `MENU_PROJECTS` in MenuFunctions.js
3. Dynamic functions auto-generate
4. Test campaign name parsing logic
5. Configure project-specific metrics and headers

### Modifying metrics
- Update `API_CONFIG.MEASURES` array
- Adjust `processApiData()` indices
- Update table headers and formatting
- Consider project-specific differences

### Changing automation schedule
- Modify trigger creation in `enableAutoCache/Update`
- Consider timezone implications
- Test trigger removal/recreation

### Debug approach
- Use `debugReportGeneration()` for full diagnostics
- Check debug sheet for step-by-step analysis
- API response structure logged for inspection
- Filter verification shows matched/unmatched campaigns
- Project-specific debugging available

## Recent Optimizations

1. **Removed ProgressManager** - Eliminated visual progress indicators for better performance
2. **Optimized Comment Caching** - No longer requires group expansion, faster operations
3. **Enhanced Project Support** - Added Mintegral, improved Google_Ads/Applovin support
4. **Improved GEO Detection** - Project-specific geo extraction logic
5. **Better Error Handling** - More robust API and sheet operations
6. **Menu Enhancements** - Smart wizards and GitHub integration
7. **Daily Automation** - Updated triggers to run daily (cache 3AM, update 5AM CET)
