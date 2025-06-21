# Campaign Report Google Apps Script - Technical Documentation

## Architecture Overview

Multi-project campaign reporting system built on Google Apps Script that fetches data from Appodeal GraphQL API, processes it into hierarchical reports with week-over-week analytics, and maintains persistent comment system with automated caching.

## Core Design Principles

1. Project Isolation: Each project (Tricky/Moloco/Regular) operates independently with own sheet, cache, and configuration
2. State Management: Uses ScriptProperties for persistent settings and hidden sheets for comment storage
3. Batch Processing: API calls fetch all data in single request, then process client-side to minimize quota usage
4. Progressive Enhancement: Core functionality works without automation; triggers add convenience

## Technical Architecture

### Data Flow
API Request -> Raw JSON -> Processing -> Hierarchical Structure -> Sheet Rendering
    |              |            |                                      |
Bearer Token   Week Aggregation                           Group Creation
GraphQL Query  Campaign Filtering                    Conditional Formatting
              WoW Calculations                         Comment Restoration

### Project Configuration Structure
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
 }
}

## Key Technical Components

### 1. GraphQL API Integration

Query Structure: Single query fetches all metrics with grouping by INSTALL_DATE, ATTRIBUTION_CAMPAIGN_HID, and APP. Uses RichStats operation with measures including CPI, installs, spend, ROAS, and forecast metrics.

Key Implementation Details:
- Headers include trace-id for debugging and proper referrer for CORS
- Filters built dynamically based on project configuration
- Date range excludes current incomplete week
- Response parsing handles both UaCampaign (Tricky/Regular) and StatsValue (Moloco) structures

### 2. Data Processing Pipeline

processApiData() function:
- Groups data by app -> week -> campaign hierarchy
- Calculates Monday-Sunday week boundaries
- Extracts GEO from campaign names using pattern matching
- Source app extraction differs by project (full name for Moloco, parsed for others)
- Filters out current week data as incomplete

### 3. Week-over-Week Analytics

calculateWoWMetrics() generates two metric sets:
- sourceAppWoW: Tracks performance by source app across weeks
- appWeekWoW: Tracks app-level weekly aggregates

Growth status algorithm considers:
- Profit sign transitions (negative to positive = Healthy Growth)
- Spend vs profit correlation
- Threshold-based categorization (10% spend change, 5% profit change)

### 4. Comment Persistence System

Two-level comment storage:
- Week-level: Keyed by AppName|||WeekRange|||WEEK|||WEEK
- Campaign-level: Keyed by AppName|||WeekRange|||CampaignId|||SourceApp

Cache sheet structure: [AppName, WeekRange, CampaignId, SourceApp, Comment, LastUpdated]

Sync process:
1. Expand all groups to access hidden rows
2. Read comments from main sheet
3. Update cache only if new comment is longer (handles appending)
4. Collapse groups after caching (recursive method for reliability)

### 5. Sheet Formatting Engine

Hierarchical grouping implementation:
- Uses shiftRowGroupDepth(1) to create nested groups
- Groups created bottom-up: campaigns first, then weeks, then apps
- Collapse uses recursive approach to handle deep nesting

Conditional formatting rules:
- eROAS colored based on project-specific targets (stored in ScriptProperties)
- WoW percentages use positive/negative coloring
- Growth status uses emoji indicators with corresponding colors

### 6. Automation System

Time-based triggers:
- autoCacheAllProjects: Daily at 2AM
- autoUpdateAllProjects: Mondays at 5AM

Auto-cache process:
1. Expand groups silently
2. Sync comments to cache
3. Collapse all groups recursively
4. No UI interruption

Auto-update process:
1. Find earliest date in existing data
2. Fetch from earliest to last complete Saturday
3. Clear and regenerate maintaining comments

### 7. Progress Management

ProgressManager class provides visual feedback during long operations:
- Creates temporary sheet with status updates
- Tracks elapsed time and progress percentage
- Auto-cleanup on completion
- Batch operations show item count

## Critical Implementation Details

### Campaign Name Parsing

Tricky/Regular campaigns:
- Extract source app after "=" sign
- Remove "autobudget" suffix
- Handle multiple "subj" occurrences

Moloco campaigns:
- Keep full campaign name (APD_ prefix)
- No parsing applied

### Date Handling

Week boundaries:
- Monday = start (day 1)
- Sunday = end (day 0 or 7)
- Current week always excluded from processing

Date range calculation:
- getDateRange(days): Inclusive range from (today - days + 1) to today
- Custom ranges: Direct YYYY-MM-DD format validation

### Error Recovery

API failures:
- Retry logic not implemented (rely on manual retry)
- Error messages preserved in debug sheet

Sheet operations:
- Silent recreation on clear to avoid corruption
- Try-catch around group operations (groups may not exist)

Comment sync:
- Only update if comment is longer (prevents data loss)
- Hyperlink extraction from formula strings

### Performance Optimizations

Batch operations:
- Single API call per report generation
- Bulk sheet writes using getRange().setValues()
- Conditional formatting applied as rule sets

Memory management:
- Process data in streaming fashion where possible
- Clear references after use
- Limit debug output to prevent memory overflow

### Security Considerations

Token storage:
- Bearer token in code (not ideal but necessary for automation)
- No user credentials stored
- Read-only API access

Data isolation:
- Each project has separate sheets
- No cross-project data leakage
- Hidden sheets for sensitive cache data

## File Dependencies

Execution order matters:
1. Config must load first (defines globals)
2. Utilities needed by most other files
3. Menu functions can reference all others
4. Auto functions depend on cache and analytics

Cross-file references:
- CURRENT_PROJECT global maintains state
- setCurrentProject() switches context
- Config getters provide dynamic values

## Limitations and Constraints

Google Apps Script limits:
- 6-minute execution timeout
- URL fetch quota (important for large date ranges)
- Trigger timing precision (~15 minutes)

Sheet limits:
- Maximum groups depth affects nesting
- Conditional format rules have maximum count
- Hidden sheets still count against sheet limit

API constraints:
- GraphQL query complexity limits
- Rate limiting not documented but observed
- Date range affects response size significantly

## Maintenance Notes

Adding new projects:
1. Add to PROJECTS object in Config
2. Add to MENU_PROJECTS in MenuFunctions
3. Dynamic functions auto-generate
4. Test campaign name parsing logic

Modifying metrics:
- Update API_CONFIG.MEASURES array
- Adjust processApiData() indices
- Update table headers and formatting

Changing automation schedule:
- Modify trigger creation in enableAutoCache/Update
- Consider timezone implications
- Test trigger removal/recreation

Debug approach:
- Use debugReportGeneration() for full diagnostics
- Check debug sheet for step-by-step analysis
- API response structure logged for inspection
- Filter verification shows matched/unmatched campaigns
