/******************************************************************************
 * Google Ads Audit Script V1.0
 *
 * This script performs a comprehensive audit of a Google Ads account and
 * outputs detailed reports into a Google Sheet. It is designed to be
 * run from the Google Ads Scripts interface.
 *
 * All data retrieval is done using Google Ads Query Language (GAQL).
 * Core audit thresholds and settings are hardcoded in SECTION 1.0 for
 * user customization.
 *
 * Author: Hassan El-Sisi (linkedin.com/in/hassan-elsisi)
 *
 * Script Purpose:
 * To provide advertisers and agencies with an automated tool for
 * in-depth account analysis, identifying areas of strength, weakness,
 * and opportunities for optimization across various aspects of their
 * Google Ads campaigns.
 *
 * PLEASE NOTE: This script is very comprehensive and may exceed
 * Google Ads Scripts' 30-minute execution limit on large accounts
 * or with very long date ranges. Test with shorter date ranges first.
 * Consider running specific report sections if timeouts occur.
 *
 ******************************************************************************/

/*********************************************************************
 * SECTION 1: Defaults, Globals & Core Utility Functions
 *
 * This section contains:
 * 1.0: User-configurable script defaults and core audit benchmarks.
 * THIS IS THE PRIMARY SECTION FOR USERS TO CUSTOMIZE THE SCRIPT'S BEHAVIOR.
 * 1.1: Global constants derived automatically from the account or script settings.
 * (Generally, do not modify these directly).
 * 1.2: Core utility functions used throughout the script for common tasks.
 *********************************************************************/

/*──────── SECTION 1.0 – SCRIPT DEFAULTS & CORE BENCHMARKS ────────*/
// Modify the values in this section to tailor the audit to your specific needs and criteria.

// The ID of the Google Spreadsheet where the audit report will be generated.
// Replace 'YOUR_SPREADSHEET_ID_HERE' with your actual Spreadsheet ID.
var SPREADSHEET_ID          = 'YOUR_SPREADSHEET_ID_HERE'; 

// --- DATE RANGE (Hardcoded) ---
// Define the reporting period for the audit.
// Format: 'YYYY-MM-DD'
var DEFAULT_DATE_FROM         = '2025-01-01'; // Example: Start of the year
var DEFAULT_DATE_TO           = '2025-04-30'; // Example: End of April

// --- PERFORMANCE THRESHOLDS (Hardcoded) ---
// These values influence the recommendations and color-coding in the report.
var DEFAULT_CLICK_THRESHOLD   = 10;   // Minimum clicks before a keyword/search term with no conversions is flagged for review.
var DEFAULT_CTR_HIGH          = 0.05; // Click-Through Rate (CTR) at or above this value (e.g., 0.05 = 5%) is considered "High".
var DEFAULT_CTR_LOW           = 0.01; // CTR at or below this value (e.g., 0.01 = 1%) is considered "Low".
var DEFAULT_CPC_MULTIPLIER    = 1.3;  // If Avg. CPC is (Channel Avg. CPC * DEFAULT_CPC_MULTIPLIER) or higher, it's flagged as "High CPC". (e.g., 1.3 = 30% higher).
var DEFAULT_CPA_MULTIPLIER    = 1.5;  // If Cost Per Acquisition (CPA) is (Channel Avg. CPA * DEFAULT_CPA_MULTIPLIER) or higher, it's flagged as "High CPA". (e.g., 1.5 = 50% higher).
var DEFAULT_TARGET_ROAS       = 2.0;  // Target Return On Ad Spend (ROAS). Performance below this may be flagged (e.g., 2.0 means 200% or 2:1).
var DEFAULT_TARGET_CPA        = 50;   // Target Cost Per Acquisition (CPA) in your account's currency. Performance above this may be flagged.
var DEFAULT_QS_LOW_THRESHOLD  = 5;    // Quality Score at or below this value (1-10 scale) is considered "Low".
var DEFAULT_IS_LOW_THRESHOLD_PERCENT = 60; // Search Impression Share below this percentage (e.g., 60%) is considered "Low".

// --- BRAND PATTERNS (Hardcoded) ---
// Add your brand terms as case-insensitive strings. Use simple terms or valid regular expressions.
// These are used to identify brand vs. non-brand keywords and search terms.
// Example: ['my brand', 'mybrand variation', 'myproductname']
var DEFAULT_BRAND_PATTERNS    = [
  'yourbrand', 'your brand', // Replace with your actual brand terms
  'براندك',    // Example Arabic brand term
  'jouwmerk',   // Example Dutch brand term
  'markanız'    // Example Turkish brand term
];

/*──────── SECTION 1.1 – GLOBAL CONSTANTS (Do Not Modify These Directly) ────────*/
// These constants are automatically set by the script based on the account or above defaults.
var ACCOUNT_CURRENCY_CODE = AdsApp.currentAccount().getCurrencyCode();
var CURRENCY_SYMBOL       = getCurrencySymbol(ACCOUNT_CURRENCY_CODE);
var ACCOUNT_NAME          = AdsApp.currentAccount().getName();
var ACCOUNT_ID            = AdsApp.currentAccount().getCustomerId();

// Color codes for conditional formatting and sheet elements
var COLOR_GOOD      = '#c8e6c9';   // Light Green (Positive)
var COLOR_REVIEW    = '#ffe0b2';   // Light Orange (Needs Review)
var COLOR_BAD       = '#ffcdd2';   // Light Red (Action Suggested / Negative)
var COLOR_HIGHLIGHT = '#fff9c4';   // Light Yellow (Informational Highlight)
var COLOR_NEUTRAL   = '#f5f5f5';   // Lighter Grey (Headers, Neutral Info)
var COLOR_TITLE_BG  = '#e0e0e0';   // Grey (Main Sheet Titles)

// Tab Colors for sheet organization
var TAB_COLOR_SUMMARIES = "#AEC6CF"; // Pastel Blue (e.g., Executive Summary)
var TAB_COLOR_MASTER    = "#B2DFDB"; // Light Teal (e.g., Master Overview, Master Recs)
var TAB_COLOR_HEATMAP   = "#C5E1A5"; // Light Green (e.g., Account Heatmap)
var TAB_COLOR_CHANNELS  = "#C5E1A5"; // Light Green (e.g., Channel-specific tabs)
var TAB_COLOR_PRODUCTS  = "#D1C4E9"; // Light Purple (e.g., Shopping Product Performance, PMax Products)
var TAB_COLOR_DEEPDIVES = "#FFECB3"; // Pale Yellow (e.g., Search Terms, Keywords, QS, Ads, IS, Audience, Device, Budget Pacing)
var TAB_COLOR_ERRORS    = "#FF0000"; // Red (for ERRORS tab)

/*──────── SECTION 1.2 – CORE UTILITY FUNCTIONS ────────*/

/**
 * Retrieves a common currency symbol for a given ISO 4217 currency code.
 * Falls back to the code itself if no symbol is mapped.
 * @param {string} code The ISO 4217 currency code (e.g., "USD", "EUR", "EGP").
 * @return {string} The currency symbol or the original code.
 */
function getCurrencySymbol(code) {
  var map = { USD:'$', EUR:'€', GBP:'£', CAD:'C$', AUD:'A$', NZD:'NZ$', SGD:'S$', HKD:'HK$', AED:'د.إ', SAR:'﷼', EGP:'E£', KWD:'د.ك', QAR:'﷼', INR:'₹', JPY:'¥', CNY:'¥', CHF:'CHF', SEK:'kr', NOK:'kr', DKK:'kr', TRY:'₺'};
  return map[code] || code;
}

/**
 * Converts a currency value from micros (used by Google Ads API for cost) to standard currency units.
 * Handles null, undefined, or non-numeric inputs by returning 0.
 * @param {?(number|string)} micros The currency value in micros.
 * @return {number} The currency value in standard units.
 */
function microsToCurrency(micros) {
    if (micros === null || typeof micros === 'undefined' || isNaN(Number(micros))) return 0;
    return Number(micros) / 1e6;
}

/**
 * Applies a standard currency format to a Google Sheets range, prefixed with the account's currency code.
 * Example: "EGP1,234.56" for positive, "(EGP1,234.56)" for negative.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The range to format.
 */
function formatCurrency(range) {
  var positiveFormat = '"' + ACCOUNT_CURRENCY_CODE + '"#,##0.00';
  var negativeFormat = '("' + ACCOUNT_CURRENCY_CODE + '"#,##0.00)';
  range.setNumberFormat(positiveFormat + ';' + negativeFormat);
}

/**
 * Applies a standard percentage format (e.g., "12.34%") to a Google Sheets range.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The range to format.
 */
function formatPercent(range) { range.setNumberFormat('0.00%'); }

/**
 * Applies a standard integer format (e.g., "1,234") to a Google Sheets range.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The range to format.
 */
function formatInteger(range) { range.setNumberFormat('#,##0'); }

/**
 * Applies a gradient conditional formatting rule to a range.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The range to apply the gradient to.
 * @param {string=} minColor Optional hex color for the minimum value.
 * @param {string=} maxColor Optional hex color for the maximum value.
 * @param {string=} midColor Optional hex color for the midpoint value.
 */
function applyGradient(sheet, range, minColor, maxColor, midColor) {
  minColor = minColor || '#f44336'; maxColor = maxColor || '#4caf50';
  var ruleBuilder = SpreadsheetApp.newConditionalFormatRule().setRanges([range]);
  if (midColor) ruleBuilder.setGradientMidpointWithValue(midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50');
  var rule = ruleBuilder.setGradientMinpoint(minColor).setGradientMaxpoint(maxColor).build();
  var rules = sheet.getConditionalFormatRules(); rules.push(rule); sheet.setConditionalFormatRules(rules);
}

/**
 * Resets a sheet: clears content, formats, notes, and conditional formatting.
 * If the sheet doesn't exist, it creates a new one.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object.
 * @param {string} name The name of the sheet to reset or create.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The reset or newly created sheet object.
 */
function resetSheet(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) {
    sh.clearContents().clearFormats().clearNotes();
    sh.setConditionalFormatRules([]);
  } else {
    sh = ss.insertSheet(name);
  }
  return sh;
}

/**
 * Sets a standardized title row at the top of a sheet.
 * Includes formatting for font, background, and borders.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to add the title to.
 * @param {string} titleText The text for the title.
 */
function setSheetTitle(sheet, titleText) {
    try {
        sheet.insertRowBefore(1);
        var maxCols = sheet.getMaxColumns();
        var mergeLimit = maxCols > 0 ? Math.min(15, maxCols) : 1; // Limit merge to avoid excessive column creation
        var titleRange = sheet.getRange("A1:" + String.fromCharCode(64 + mergeLimit) + "1");
        if (mergeLimit > 1) titleRange.merge();
        titleRange.setValue(titleText)
                  .setFontSize(14)
                  .setFontWeight('bold')
                  .setHorizontalAlignment('center')
                  .setVerticalAlignment('middle')
                  .setBackground(COLOR_TITLE_BG)
                  .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        sheet.setRowHeight(1, 30);
        sheet.insertRowAfter(1); // Adds a spacer row below the title
    } catch(e) { Logger.log('Error in setSheetTitle for ' + sheet.getName() + ': ' + e); }
}

/**
 * Generates a unique placeholder sheet name to assist with sheet cleanup.
 * @return {string} A unique placeholder sheet name.
 */
function generatePlaceholderName() {
  var tz = AdsApp.currentAccount().getTimeZone();
  var ts = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmssSSS');
  return '_AUDITPLACEHOLDER_' + ts;
}

/**
 * Deletes all sheets in the spreadsheet except for essential ones and the placeholder.
 * Used to clean up sheets from previous script runs.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object.
 * @param {string} placeholderName The name of the placeholder sheet to preserve during cleanup.
 */
function wipeSheetsExcept(ss, placeholderName) {
  var essentialSheets = [placeholderName, 'CONFIG (Informational - Settings are Hardcoded in Script)', 'README', 'ERRORS', 'Executive Summary'];
  ss.getSheets().forEach(function(sh) {
    if (essentialSheets.indexOf(sh.getName()) === -1 && !sh.getName().startsWith('_AUDITPLACEHOLDER_')) {
      Logger.log('Deleting sheet: ' + sh.getName());
      ss.deleteSheet(sh);
    }
  });
}

/**
 * Logs an error to both the Google Ads Scripts logger and a dedicated 'ERRORS' tab in the spreadsheet.
 * @param {string} fnName The name of the function where the error occurred.
 * @param {Error} err The error object.
 * @param {string=} query Optional GAQL query string associated with the error.
 */
function logError(fnName, err, query) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName('ERRORS');
    if (!sh) {
        sh = ss.insertSheet('ERRORS');
        setSheetTitle(sh, 'SCRIPT EXECUTION ERRORS');
        sh.getRange(3, 1, 1, 5).setValues([['Timestamp','Function','Error Message', 'Query (If Applicable)', 'Stack Trace']])
                               .setFontWeight("bold").setBackground(COLOR_NEUTRAL);
        sh.setFrozenRows(3);
        if (TAB_COLOR_ERRORS) sh.setTabColor(TAB_COLOR_ERRORS);
    }
    var timestamp = Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    var errorMessage = err.message || err.toString();
    var stackTrace = err.stack || 'N/A';
    var queryText = query || 'N/A';
    sh.appendRow([timestamp, fnName, errorMessage, queryText, stackTrace]);
    Logger.log('ERROR in ' + fnName + ': ' + errorMessage + (queryText !== 'N/A' ? ' | Query: ' + queryText : '') + (stackTrace !== 'N/A' ? ' | Stack: ' + stackTrace : ''));
  } catch (e) {
    Logger.log('CRITICAL ERROR in logError: ' + e + ' | Original Error in ' + fnName + ': ' + (err.message || err.toString()));
  }
}

/**
 * Wrapper function to safely execute report-generating functions and log any errors.
 * @param {function} fn The function to execute.
 * @param {string} fnName The name of the function (for logging purposes).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object.
 * @param {object} cfg The configuration object.
 * @param {object} summaryData An object to collect summary data for the Executive Summary.
 */
function safeRun(fn, fnName, ss, cfg, summaryData) {
  try {
    Logger.log('Running: ' + fnName);
    fn(ss, cfg, summaryData);
    Logger.log('Finished: ' + fnName);
  } catch (e) {
    logError(fnName, e);
  }
}

/**
 * Automatically resizes all columns in a given sheet to fit their content.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 */
function autoResizeAllColumns(sheet) {
  try {
    var lastCol = sheet.getLastColumn();
    if (lastCol > 0) sheet.autoResizeColumns(1, lastCol);
  } catch(e){ logError('autoResizeAllColumns on ' + sheet.getName(), e); }
}

/**
 * Freezes header rows in a sheet. Assumes a title row, a spacer row, and data header row(s).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {number=} rowsToFreezeData The number of actual data header rows to freeze (default is 1).
 */
function freezeHeader(sheet, rowsToFreezeData) {
  rowsToFreezeData = rowsToFreezeData || 1;
  var actualRowsToFreeze = 2 + rowsToFreezeData; // 2 for title and spacer rows
  try { sheet.setFrozenRows(actualRowsToFreeze); } catch(e) { Logger.log('Error freezing header for ' + sheet.getName() + ': ' + e); }
}

/**
 * Applies standard formatting to a data header row in a sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {number=} dataHeaderRowIndex The 1-based index of the data header row (default is 3).
 */
function setHeaderRowFormatting(sheet, dataHeaderRowIndex) {
    dataHeaderRowIndex = dataHeaderRowIndex || 3; // Row 1: Title, Row 2: Spacer, Row 3: Headers
    var lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
        try {
            sheet.getRange(dataHeaderRowIndex, 1, 1, lastCol)
                 .setFontWeight("bold")
                 .setBackground(COLOR_NEUTRAL)
                 .setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
                 .setHorizontalAlignment('center');
        } catch(e) { Logger.log('Error setting header format for ' + sheet.getName() + ': ' + e); }
    }
}
// End of SECTION 1

/*********************************************************************
 * SECTION 2: CONFIG Sheet Builder & loadConfigObject()
 *
 * This section handles the script's configuration.
 * - buildConfigSheet: Creates a hidden, informational sheet displaying the
 * hardcoded default values used by the script.
 * - loadConfigObject: Loads the configuration from the hardcoded defaults
 * defined in SECTION 1.0.
 *********************************************************************/
function buildConfigSheet(ss) {
  var sheetName = 'CONFIG (Informational - Settings are Hardcoded in Script)';
  var sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName); else sh.clear();
  
  setSheetTitle(sh, 'SCRIPT CONFIGURATION SETTINGS (Informational)');
  ss.setActiveSheet(sh);

  var headers = ['Setting Group', 'Parameter', 'Value Used by Script', 'Description / How to Change'];
  var headerRowDataIndex = 3;
  sh.getRange(headerRowDataIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowDataIndex);

  var defaultsInfo = [
    ['Informational', 'Account Name', ACCOUNT_NAME, 'Detected from Google Ads account.'],
    ['Informational', 'Account ID', ACCOUNT_ID, 'Detected from Google Ads account.'],
    ['Informational', 'Account Currency', ACCOUNT_CURRENCY_CODE + ' (' + CURRENCY_SYMBOL + ')', 'Detected from Google Ads account.'],
    ['Date Range', 'Report Start Date', DEFAULT_DATE_FROM, 'Hardcoded in Script (SECTION 1.0). Edit DEFAULT_DATE_FROM there.'],
    ['Date Range', 'Report End Date', DEFAULT_DATE_TO, 'Hardcoded in Script (SECTION 1.0). Edit DEFAULT_DATE_TO there.'],
    ['Performance Thresholds', 'Click Threshold (for Negatives)', DEFAULT_CLICK_THRESHOLD, 'Hardcoded: DEFAULT_CLICK_THRESHOLD in Section 1.0'],
    ['Performance Thresholds', 'CTR - High', (DEFAULT_CTR_HIGH * 100).toFixed(2) + '%', 'Hardcoded: DEFAULT_CTR_HIGH in Section 1.0'],
    ['Performance Thresholds', 'CTR - Low', (DEFAULT_CTR_LOW * 100).toFixed(2) + '%', 'Hardcoded: DEFAULT_CTR_LOW in Section 1.0'],
    ['Performance Thresholds', 'CPC Multiplier (for High CPC)', DEFAULT_CPC_MULTIPLIER + 'x', 'Hardcoded: DEFAULT_CPC_MULTIPLIER in Section 1.0'],
    ['Performance Thresholds', 'CPA Multiplier (for High CPA)', DEFAULT_CPA_MULTIPLIER + 'x', 'Hardcoded: DEFAULT_CPA_MULTIPLIER in Section 1.0'],
    ['Performance Thresholds', 'Target ROAS', (DEFAULT_TARGET_ROAS * 100).toFixed(0) + '%', 'Hardcoded: DEFAULT_TARGET_ROAS in Section 1.0 (as decimal)'],
    ['Performance Thresholds', 'Target CPA', CURRENCY_SYMBOL + DEFAULT_TARGET_CPA.toFixed(2), 'Hardcoded: DEFAULT_TARGET_CPA in Section 1.0'],
    ['Performance Thresholds', 'Quality Score Low Threshold', DEFAULT_QS_LOW_THRESHOLD, 'Hardcoded: DEFAULT_QS_LOW_THRESHOLD in Section 1.0'],
    ['Performance Thresholds', 'Imp. Share Low Threshold', DEFAULT_IS_LOW_THRESHOLD_PERCENT + '%', 'Hardcoded: DEFAULT_IS_LOW_THRESHOLD_PERCENT in Section 1.0'],
    ['Patterns & Advanced', 'Brand Patterns (Regex)', DEFAULT_BRAND_PATTERNS.join(', '), 'Hardcoded: DEFAULT_BRAND_PATTERNS array in Section 1.0'],
    ['Patterns & Advanced', 'Change History Lookback', 'Uses Main Report Dates', 'Change History now uses the main report date range defined in SECTION 1.0.']
  ];
  
  var dataStartRow = headerRowDataIndex + 1;
  sh.getRange(dataStartRow, 1, defaultsInfo.length, headers.length).setValues(defaultsInfo);

  for (var i = 0; i < defaultsInfo.length; i++) {
      var currentRow = dataStartRow + i;
      if (defaultsInfo[i][0].startsWith('SEPARATOR_')) { 
          sh.getRange(currentRow, 1, 1, headers.length).setBackground(COLOR_NEUTRAL).setFontWeight('bold').mergeAcross();
      } else if (defaultsInfo[i][0].startsWith('INFO_')) {
           sh.getRange(currentRow, 1, 1, 1).setFontStyle('italic');
           sh.getRange(currentRow, 3, 1, 1).setFontWeight('bold').setFontStyle('italic'); 
      } else {
           sh.getRange(currentRow, 3, 1, 1).setFontWeight('bold'); 
      }
  }
  sh.getRange(dataStartRow, 3, defaultsInfo.length, 1).setNumberFormat('@'); // Value column as Plain Text
  freezeHeader(sh, 1);
  autoResizeAllColumns(sh);
  sh.hideSheet();
}

/**
 * Loads the script configuration directly from the hardcoded DEFAULT_ values in SECTION 1.0.
 * Processes brand patterns into RegExp objects.
 * @return {object} The configuration object used by the script.
 */
function loadConfigObject() {
  var patterns = DEFAULT_BRAND_PATTERNS.map(function(s) {
    try {
      var trimmedPattern = String(s).trim();
      if (trimmedPattern === '') return null;
      return new RegExp(trimmedPattern, 'i');
    } catch (e) { logError('loadConfigObject_BrandPattern', 'Invalid regex pattern in DEFAULT_BRAND_PATTERNS: ' + s, e); return null; }
  }).filter(function(p) { return p !== null; });

  var config = {
    DATE_FROM:       DEFAULT_DATE_FROM, DATE_TO:         DEFAULT_DATE_TO,
    CLICK_THRESHOLD: DEFAULT_CLICK_THRESHOLD,
    CTR_HIGH:        DEFAULT_CTR_HIGH, CTR_LOW:         DEFAULT_CTR_LOW,
    CPC_MULTIPLIER:  DEFAULT_CPC_MULTIPLIER, CPA_MULTIPLIER:  DEFAULT_CPA_MULTIPLIER,
    TARGET_ROAS:     DEFAULT_TARGET_ROAS, TARGET_CPA:      DEFAULT_TARGET_CPA,
    QS_LOW_THRESHOLD:DEFAULT_QS_LOW_THRESHOLD,
    IS_LOW_THRESHOLD_PERCENT: DEFAULT_IS_LOW_THRESHOLD_PERCENT,
    BRAND_PATTERNS:  patterns.length > 0 ? patterns : [new RegExp('^xanthophyll$', 'i')], // Fallback to a non-matching regex
  };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(config.DATE_FROM) || !/^\d{4}-\d{2}-\d{2}$/.test(config.DATE_TO)) {
      logError('loadConfigObject', 'Hardcoded DATE_FROM/TO (' + config.DATE_FROM + '/' + config.DATE_TO + ') is not in<y_bin_46>-MM-DD format. Script may fail.');
  }
  return config;
}
// End of SECTION 2

/*********************************************************************
 * SECTION 3: README Sheet Builder
 *
 * This function creates the README sheet with instructions, color legend,
 * and an overview of the generated tabs.
 *********************************************************************/
function buildReadmeSheet(ss, cfg) {
  var sh = resetSheet(ss, 'README');
  setSheetTitle(sh, 'GOOGLE ADS AUDIT SCRIPT - README & INFORMATION');
  if (sh.getTabColor() === null) sh.setTabColor(null); // Default tab color

  var rawRows = [
    ['Author:','=HYPERLINK("https://www.linkedin.com/in/hassan-elsisi","Hassan El-Sisi")', '', ''],
    ['Account Name:', ACCOUNT_NAME + ' (' + ACCOUNT_ID + ')', '', ''],
    ['Last Run:', Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd HH:mm:ss Z'), '', ''],
    ['Reporting Period:', (cfg ? cfg.DATE_FROM + ' to ' + cfg.DATE_TO : 'N/A - Check Script Defaults (Section 1)'), '', ''],
    ['','', '', ''],
    ['## How to Use','', '', ''],
    ['1. **SPREADSHEET_ID:** Ensure `SPREADSHEET_ID` in Section 1.0 of the script is correct.','', '', ''],
    ['2. **CORE SETTINGS (Dates, Thresholds, Brand Patterns):** These are now HARDCODED in "SECTION 1.0 – SCRIPT DEFAULTS & CORE BENCHMARKS" of the script code. Modify them directly in the script if needed.','', '', ''],
    ['3. **CONFIG Tab (Hidden):** This sheet is now for INFORMATIONAL purposes only. It displays the hardcoded values the script is using. To view it, go to View > Hidden sheets in Google Sheets.','', '', ''],
    ['4. **Run Script:** In Google Ads UI: Tools & Settings > Bulk Actions > Scripts. Find this script and click "▶ Run".','', '', ''],
    ['5. **Review Tabs:** After execution, review all generated tabs.','', '', ''],
    ['6. **ERRORS Tab:** Check `ERRORS` tab for any issues. If empty after a run, it will be auto-deleted.','', '', ''],
    ['7. **Execution Time:** For large accounts/long date ranges, script may time out (30 min limit).','', '', ''],
    ['','', '', ''],
    ['## Colour Legend','(Applied to "Recommendation" or relevant metric cells)', '', ''],
    ['Colour Code','Meaning (Examples, actual logic varies per tab)', '', ''],
    ['Green ('+COLOR_GOOD+')','Generally Positive Performance / Good Standing', '', ''],
    ['Orange ('+COLOR_REVIEW+')','Needs Review / Borderline Performance / Monitor', '', ''],
    ['Red ('+COLOR_BAD+')','Action Suggested / Negative Performance Concern', '', ''],
    ['Yellow ('+COLOR_HIGHLIGHT+')','Informational Highlight (e.g., Paused, Low Data with Cost)', '', ''],
    ['Grey ('+COLOR_NEUTRAL+')','Neutral Information / Headers', '', ''],
    ['','', '', ''],
    ['## Generated Tabs Overview','(Order may vary slightly)', '', ''],
    ['Tab Name','Description', '', ''],
    ['README', 'This sheet: Instructions, Legend, Tab List.', '', ''],
    ['CONFIG (Informational - Settings are Hardcoded in Script)','(Hidden) Informational display of hardcoded script settings.', '', ''],
    ['ERRORS','Logs of script errors. (Auto-deleted if no errors).', '', ''],
    ['Executive Summary', 'High-level summary of key findings (experimental).', '', ''],
    ['Master Overview','Performance summary by Advertising Channel Type.', '', ''],
    ['Master Recs','Automated recommendations for channels from Master Overview data.', '', ''],
    ['Account Heatmap','Hour × Day of Week performance heatmaps for key account metrics.', '', ''],
    ['Channel – Search','Summary & recommendations for Search campaigns.', '', ''],
    ['Channel – Display','Summary & recommendations for Display campaigns.', '', ''],
    ['Channel – YouTube','Summary & recommendations for Video (YouTube) campaigns.', '', ''],
    ['Channel – Shopping','Summary & recommendations for Shopping campaigns.', '', ''],
    ['Channel – PMax','Summary & recommendations for Performance Max campaigns.', '', ''],
    ['Channel – App Installs','Summary for App Install campaigns.', '', ''],
    ['Channel – App Engagement','Summary for App Engagement campaigns.', '', ''],
    ['Search Terms Report','Detailed list of search terms (Search campaigns) with brand flags & actions.', '', ''], 
    ['Keyword Deep Dive','Performance of all keywords (Search campaigns) with brand flags & actions.', '', ''],
    ['Quality Score (Search)','Analysis of Keyword Quality Score and its components for Search.', '', ''], 
    ['Search IS Analysis', 'Impression Share metrics for Search campaigns.', '', ''],
    ['Audience Insights', 'Performance breakdown by audience types used in campaigns.', '', ''],
    ['Device Performance', 'Performance breakdown by device (Desktop, Mobile, Tablet) at account level.', '', ''],
    ['Ads Performance (Search)','Performance of Search ads with recommendations.', '', ''], 
    ['Budget Pacing', 'Report on campaign budget utilization.', '', ''], 
    ['KW - {Campaign Name}','(Per active Search campaign) Keyword deep-dive for that specific campaign.', '', ''],
    ['PMax Asset Groups','Asset Group level performance within Performance Max campaigns.', '', ''],
    ['Shopping Product Performance','Product attribute level performance for Shopping campaigns (using GAQL).', '', ''],
    ['PMax Product Titles','Product Title level performance within PMax.', '', ''],
    ['','', '', ''],
    ['## Disclaimer & Notes','', '', ''],
    ['• This script provides automated analysis and recommendations based on the configured thresholds and common best practices.','', '', ''],
    ['• **ACTION REQUIRED:** All findings and recommendations should be critically reviewed by an experienced account manager before taking any action in the Google Ads account. Act as you see fit and best performing for your specific account goals and context.','', '', ''],
    ['• The script DOES NOT make any changes to your Google Ads account. It is for reporting and analysis only.','', '', ''],
    ['• Data accuracy depends on the Google Ads API. Metrics might differ slightly from the UI due to reporting lags or different attribution models.','', '', ''],
    ['• Ensure you have granted necessary permissions for the script to access Ads data and Google Sheets.','', '', '']
  ];

  var maxColsReadme = 4;
  var rows = rawRows.map(function(row) {
      var newRow = [];
      for (var i = 0; i < maxColsReadme; i++) {
          newRow.push(row[i] || '');
      }
      return newRow;
  });

  var dataStartRow = 3;
  sh.getRange(dataStartRow, 1, rows.length, maxColsReadme).setValues(rows);
  sh.getRange(dataStartRow, 2).setFormula('=HYPERLINK("https://www.linkedin.com/in/hassan-elsisi","Hassan El-Sisi")');

  freezeHeader(sh, 1);

  var legendDataStartRow = dataStartRow + 14;
  sh.getRange(legendDataStartRow, 1).setBackground(COLOR_GOOD);
  sh.getRange(legendDataStartRow + 1, 1).setBackground(COLOR_REVIEW);
  sh.getRange(legendDataStartRow + 2, 1).setBackground(COLOR_BAD);
  sh.getRange(legendDataStartRow + 3, 1).setBackground(COLOR_HIGHLIGHT);
  sh.getRange(legendDataStartRow + 4, 1).setBackground(COLOR_NEUTRAL);
  
  var boldHeadersInReadme = [dataStartRow + 5, dataStartRow + 13, dataStartRow + 20, dataStartRow + 48]; 
  boldHeadersInReadme.forEach(function(rowIndex){ sh.getRange(rowIndex, 1, 1, 1).setFontWeight("bold").setFontSize(11); });

  autoResizeAllColumns(sh);
}
// End of SECTION 3

/*********************************************************************
 * SECTION 4: MAIN DRIVER & Orchestration
 *
 * This is the main function that orchestrates the script's execution.
 * It initializes the spreadsheet, loads configuration, calls the
 * various report-building functions, and handles final sheet ordering.
 *********************************************************************/
function main() {
  var scriptStartTime = new Date();
  Logger.log('Starting Google Ads Audit Script V2.1.10 (Enhanced Budget Pacing) at ' + scriptStartTime);
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var placeholder = generatePlaceholderName();
  try { ss.insertSheet(placeholder, 0); } catch (e) { /* ignore if sheet already exists from a previously failed run */ }
  wipeSheetsExcept(ss, placeholder);

  var globalSummaryData = { findings: [], opportunities: [], warnings: [] };

  safeRun(buildConfigSheet, 'buildConfigSheet', ss, null, globalSummaryData);
  var cfg = loadConfigObject();
  if (!cfg || !cfg.DATE_FROM || !cfg.DATE_TO) {
      logError('main_loadConfig', 'Critical: Failed to load config. Aborting.');
      var phSheet = ss.getSheetByName(placeholder); if (phSheet) ss.deleteSheet(phSheet); return;
  }
  safeRun(buildReadmeSheet, 'buildReadmeSheet', ss, cfg, globalSummaryData);

  Logger.log('Config loaded. Using HARDCODED Date Range: ' + cfg.DATE_FROM + ' to ' + cfg.DATE_TO);

  // Call report functions in desired logical order
  safeRun(masterOverview_,   'Master Overview', ss, cfg, globalSummaryData);
  safeRun(masterRecs_,       'Master Recommendations', ss, cfg, globalSummaryData);
  safeRun(accountHeatmap_,   'Account Heatmap', ss, cfg, globalSummaryData);
  
  safeRun(buildChannelTabs,  'All Channel Tabs', ss, cfg, globalSummaryData); 
  
  safeRun(buildShoppingProductPerformance_, 'Shopping Product Performance (GAQL)', ss, cfg, globalSummaryData);
  safeRun(pmaxProducts_ProductTitle, 'PMax Product Titles Performance', ss, cfg, globalSummaryData);
  safeRun(deepPMax_AssetGroup,       'PMax Asset Groups Performance', ss, cfg, globalSummaryData);
  
  // Asset Health & Change History reports removed
  
  safeRun(buildSearchTerms_, 'Search Terms Report', ss, cfg, globalSummaryData);
  safeRun(deepSearch_KeywordDeepDive, 'Keyword Deep Dive (All Search)', ss, cfg, globalSummaryData);
  safeRun(buildAdsTab,       'Search Ads Performance', ss, cfg, globalSummaryData);
  safeRun(buildQualityScoreReport_, 'Quality Score Report (Search)', ss, cfg, globalSummaryData);
  safeRun(buildSearchISReport_,     'Search Impression Share Report', ss, cfg, globalSummaryData);
  safeRun(buildAudienceReport_,     'Audience Performance Report', ss, cfg, globalSummaryData);
  safeRun(buildDeviceReport_,       'Device Performance Report', ss, cfg, globalSummaryData);
  safeRun(buildBudgetPacingReport_, 'Budget Pacing Report', ss, cfg, globalSummaryData);
  
  safeRun(buildExecutiveSummary_, 'Executive Summary', ss, cfg, globalSummaryData);
  
  // Per-Campaign KW reports run last
  safeRun(buildCampaignKW_,  'Per-Campaign Keyword Reports (Search)', ss, cfg, globalSummaryData); 


  // Cleanup and Finalization
  var finalPlaceholderSheet = ss.getSheetByName(placeholder);
  if (finalPlaceholderSheet) ss.deleteSheet(finalPlaceholderSheet);
  
  // Sheet Ordering
  var errorsSheet = ss.getSheetByName('ERRORS');
  var errorsSheetExistsAndHasContent = errorsSheet && errorsSheet.getLastRow() > 3; 

  var configSheet = ss.getSheetByName('CONFIG (Informational - Settings are Hardcoded in Script)');

  if (errorsSheet && !errorsSheetExistsAndHasContent) {
      Logger.log('No errors logged. Deleting ERRORS sheet.');
      ss.deleteSheet(errorsSheet);
      errorsSheet = null; 
  } else if (errorsSheetExistsAndHasContent) {
      autoResizeAllColumns(errorsSheet);
  }
  
  var desiredOrder = ['README'];
  if (errorsSheet) desiredOrder.push('ERRORS');
  if (configSheet) desiredOrder.push(configSheet.getName()); 
  desiredOrder.push('Executive Summary');
  desiredOrder.push('Master Overview');
  desiredOrder.push('Master Recs');
  desiredOrder.push('Account Heatmap');
  
  var channelNames = [ 
    'Channel – Search', 'Channel – Shopping', 'Channel – PMax', 
    'Channel – Display', 'Channel – YouTube', 
    'Channel – App Installs', 'Channel – App Engagement'
  ];
  channelNames.forEach(function(cn) { if(ss.getSheetByName(cn)) desiredOrder.push(cn); });

  var productTabs = ['Shopping Product Performance', 'PMax Product Titles', 'PMax Asset Groups'];
  productTabs.forEach(function(pt) { if(ss.getSheetByName(pt)) desiredOrder.push(pt); });
  
  var otherReports = [
    'Search Terms Report', 'Keyword Deep Dive', 'Quality Score (Search)', 
    'Search IS Analysis', 'Ads Performance (Search)', 'Audience Insights', 
    'Device Performance', 'Budget Pacing'
  ];
  otherReports.forEach(function(or) { if(ss.getSheetByName(or)) desiredOrder.push(or); });

  var currentPos = 1;
  desiredOrder.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) { 
          try {
              if (sheet.getIndex() !== currentPos) {
                  ss.setActiveSheet(sheet);
                  ss.moveActiveSheet(currentPos);
              }
              currentPos++;
          } catch (e) {
              logError('main_sheetOrdering_Front', 'Error moving sheet: ' + sheetName + ' - ' + e.message);
          }
      }
  });


  var scriptEndTime = new Date();
  var durationMinutes = (scriptEndTime.getTime() - scriptStartTime.getTime()) / 60000;
  Logger.log('Google Ads Audit Script V2.1.10 (Enhanced Budget Pacing) finished successfully in ' + durationMinutes.toFixed(2) + ' minutes.');
  SpreadsheetApp.flush();
}
// End of SECTION 4

/*********************************************************************
 * SECTION 5: Master Overview & Master Recommendations (All GAQL)
 *
 * These functions provide a high-level summary of account performance
 * broken down by advertising channel type (Search, Display, etc.)
 * and then offer top-level recommendations based on this data.
 *********************************************************************/
function masterOverview_(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Master Overview');
  setSheetTitle(sh, 'MASTER OVERVIEW - Performance by Advertising Channel (GAQL)');
  if (TAB_COLOR_MASTER) sh.setTabColor(TAB_COLOR_MASTER);
  var headers = ['Channel', 'Impr', 'Clicks', 'Cost', 'Conv', 'Conv Value', 'CTR', 'Avg CPC', 'CPA', 'ROAS'];
  var headerRowIndex = 3;
  sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var gaql = 'SELECT campaign.advertising_channel_type, metrics.impressions, metrics.clicks, metrics.cost_micros, metrics.conversions, metrics.conversions_value ' +
             'FROM campaign WHERE segments.date BETWEEN \'' + dateFromGAQL + '\' AND \'' + dateToGAQL + '\' AND metrics.impressions > 0';
  Logger.log('Master Overview GAQL: ' + gaql);
  var reportIterator;
  try { reportIterator = AdsApp.report(gaql).rows(); } catch (e) { logError('masterOverview_GAQL', e, gaql); sh.getRange(headerRowIndex + 1, 1).setValue('Error fetching data.'); return; }

  var agg = {};
  while (reportIterator.hasNext()) {
    var r = reportIterator.next(); var chan = r['campaign.advertising_channel_type'];
    if (!agg[chan]) agg[chan] = { imp:0, clk:0, cost:0, conv:0, convVal:0 };
    agg[chan].imp += Number(r['metrics.impressions'] || 0); agg[chan].clk += Number(r['metrics.clicks'] || 0);
    agg[chan].cost += microsToCurrency(Number(r['metrics.cost_micros'] || 0)); agg[chan].conv += Number(r['metrics.conversions'] || 0);
    agg[chan].convVal += Number(r['metrics.conversions_value'] || 0); 
  }

  var outputRows = []; var totalCostAll = 0; var totalConvAll = 0;
  for (var chanKey in agg) {
    if (agg.hasOwnProperty(chanKey)) {
      var o = agg[chanKey];
      var ctr = o.imp > 0 ? Number(o.clk || 0) / Number(o.imp || 1) : 0;
      var avgCpc = o.clk > 0 ? Number(o.cost || 0) / Number(o.clk || 1) : 0;
      var cpa = o.conv > 0 ? Number(o.cost || 0) / Number(o.conv || 1) : 0;
      var roas = (typeof o.cost === 'number' && o.cost > 0 && typeof o.convVal === 'number') ? o.convVal / o.cost : 0;
      outputRows.push([chanKey, o.imp, o.clk, o.cost, o.conv, o.convVal, ctr, avgCpc, cpa, roas]);
      totalCostAll += o.cost; totalConvAll += o.conv;
      if (roas > cfg.TARGET_ROAS * 1.2) summaryData.findings.push(chanKey + ' has strong ROAS: ' + (roas*100).toFixed(0) + '%');
      if (cpa > 0 && cpa < cfg.TARGET_CPA * 0.8) summaryData.findings.push(chanKey + ' has good CPA: ' + CURRENCY_SYMBOL + cpa.toFixed(2));
    }
  }
  if (outputRows.length > 0) {
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow, 1, outputRows.length, headers.length).setValues(outputRows);
    var lr = outputRows.length;
    formatInteger(sh.getRange(dataStartRow, 2, lr, 1)); formatInteger(sh.getRange(dataStartRow, 3, lr, 1)); 
    formatCurrency(sh.getRange(dataStartRow, 4, lr, 1));
    formatInteger(sh.getRange(dataStartRow, 5, lr, 1)); formatCurrency(sh.getRange(dataStartRow, 6, lr, 1));
    formatPercent(sh.getRange(dataStartRow, 7, lr, 1)); formatCurrency(sh.getRange(dataStartRow, 8, lr, 1));
    formatCurrency(sh.getRange(dataStartRow, 9, lr, 1)); formatPercent(sh.getRange(dataStartRow, 10, lr, 1));
    autoResizeAllColumns(sh);
  } else { sh.getRange(headerRowIndex + 1, 1).setValue('No data found.'); }
  summaryData.accountAvgCpa = totalConvAll > 0 ? totalCostAll / totalConvAll : 0;
}

function masterRecs_(ss, cfg, summaryData) {
  var srcSheet = ss.getSheetByName('Master Overview');
  if (!srcSheet || srcSheet.getLastRow() < 4) { logError('masterRecs_', 'Master Overview sheet empty or no data.'); return; }
  var sh = resetSheet(ss, 'Master Recs');
  setSheetTitle(sh, 'MASTER RECOMMENDATIONS - Channel Level');
  if (TAB_COLOR_MASTER) sh.setTabColor(TAB_COLOR_MASTER);
  var headers = ['Channel','Impr','Clicks','CTR','CPA','ROAS','Advice'];
  var headerRowIndex = 3;
  sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var data = srcSheet.getRange(4, 1, srcSheet.getLastRow()-3, srcSheet.getLastColumn()).getValues();
  var accountAvgCpa = summaryData.accountAvgCpa || cfg.TARGET_CPA;

  var outputRows = [];
  data.forEach(function(r, index) {
    var channel = r[0], impr = Number(r[1]||0), clicks = Number(r[2]||0), ctr = Number(r[6]||0), cpa = Number(r[8]||0), roas = Number(r[9]||0), cost = Number(r[3]||0), conv = Number(r[4]||0);
    var tips = [], color = COLOR_NEUTRAL;

    if (impr === 0 && clicks === 0) { tips.push('No Impressions/Clicks'); color = COLOR_HIGHLIGHT; }
    else {
      if (ctr < cfg.CTR_LOW && clicks > 0) { tips.push('Low CTR ('+(ctr*100).toFixed(1)+'%)'); color = COLOR_BAD; }
      if (cpa > 0 && accountAvgCpa > 0 && cpa > (accountAvgCpa * cfg.CPA_MULTIPLIER)) { tips.push('High CPA ('+CURRENCY_SYMBOL+cpa.toFixed(0)+') vs Acct Avg'); color = COLOR_BAD; }
      else if (cpa > 0 && cpa > cfg.TARGET_CPA) { tips.push('CPA ('+CURRENCY_SYMBOL+cpa.toFixed(0)+') > Target'); color = COLOR_REVIEW; }
      if (conv === 0 && cost > 0 && clicks > cfg.CLICK_THRESHOLD) { tips.push('No Conv (Cost > 0)'); color = COLOR_BAD; }
      if (roas > 0 && roas < cfg.TARGET_ROAS) { tips.push('ROAS ('+(roas*100).toFixed(0)+'%) < Target'); color = roas < (cfg.TARGET_ROAS*0.5) ? COLOR_BAD : COLOR_REVIEW; }
      
      if (tips.length === 0) { tips.push('KPIs within basic thresholds. Review details.'); color = COLOR_GOOD; }
    }
    var adviceText = tips.join(' • ') || 'Review KPIs';
    outputRows.push([channel, impr, clicks, ctr, cpa, roas, adviceText]);
    sh.getRange(headerRowIndex + 1 + index, headers.length).setBackground(color);
  });

  if (outputRows.length > 0) {
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow, 1, outputRows.length, headers.length).setValues(outputRows);
    formatInteger(sh.getRange(dataStartRow,2,outputRows.length,1)); formatInteger(sh.getRange(dataStartRow,3,outputRows.length,1)); 
    formatPercent(sh.getRange(dataStartRow,4,outputRows.length,1));
    formatCurrency(sh.getRange(dataStartRow,5,outputRows.length,1)); formatPercent(sh.getRange(dataStartRow,6,outputRows.length,1));
    autoResizeAllColumns(sh);
  } else { sh.getRange(headerRowIndex + 1, 1).setValue('No data from Master Overview.'); }
}
// End of SECTION 5

/*********************************************************************
 * SECTION 6: Account Heatmap - GAQL
 *********************************************************************/
function accountHeatmap_(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Account Heatmap');
  setSheetTitle(sh, 'ACCOUNT PERFORMANCE HEATMAP - By Hour of Day & Day of Week');
  if(TAB_COLOR_HEATMAP) sh.setTabColor(TAB_COLOR_HEATMAP);
  sh.getRange(3,1).setValue('Data from: ' + cfg.DATE_FROM + ' to ' + cfg.DATE_TO + ' (Timezone: ' + AdsApp.currentAccount().getTimeZone() + ')').setFontStyle('italic');
  var dataHeaderBaseRow = 4;

  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var DAYS = ['MONDAY','TUESDAY','WEDNESDAY','THURSDAY','FRIDAY','SATURDAY','SUNDAY'];
  var METRIC_DEFS = [
    { name: 'Impressions', field: 'metrics.impressions', type: 'NUMBER', agg: 'SUM' }, { name: 'Clicks', field: 'metrics.clicks', type: 'NUMBER', agg: 'SUM' },
    { name: 'Cost', field: 'metrics.cost_micros', type: 'COST', agg: 'SUM' }, { name: 'Conversions', field: 'metrics.conversions', type: 'NUMBER', agg: 'SUM' },
    { name: 'CTR', num: 'metrics.clicks', den: 'metrics.impressions', type: 'PERCENT', agg: 'RATE'},
    { name: 'Avg CPC', num: 'metrics.cost_micros', den: 'metrics.clicks', type: 'CURRENCY', agg: 'RATE'},
    { name: 'CPA', num: 'metrics.cost_micros', den: 'metrics.conversions', type: 'CURRENCY', agg: 'RATE'}
  ];
  var baseFields = ['metrics.impressions', 'metrics.clicks', 'metrics.cost_micros', 'metrics.conversions'].filter(function(f,i,self){ return self.indexOf(f) === i; });
  var gaql = 'SELECT segments.day_of_week, segments.hour, ' + baseFields.join(', ') + ' FROM campaign WHERE segments.date BETWEEN \'' + dateFromGAQL + '\' AND \'' + dateToGAQL + '\'';
  Logger.log('Account Heatmap GAQL: ' + gaql);
  var hourlyData = {}; DAYS.forEach(function(d){ hourlyData[d] = {}; for(var hr=0; hr<24; hr++) { hourlyData[d][hr] = {}; baseFields.forEach(function(f){ hourlyData[d][hr][f] = 0;}); } });

  try {
    var iter = AdsApp.report(gaql).rows();
    while(iter.hasNext()){
        var r = iter.next(); var day = r['segments.day_of_week']; var hour = parseInt(r['segments.hour'], 10);
        if(hourlyData[day] && typeof hourlyData[day][hour] !== 'undefined'){ baseFields.forEach(function(f){ hourlyData[day][hour][f] += Number(r[f] || 0); }); }
    }
  } catch (e) { logError('accountHeatmap_GAQL', e, gaql); sh.getRange(dataHeaderBaseRow,1).setValue('Error fetching data.'); return; }

  var currentRow = dataHeaderBaseRow;
  METRIC_DEFS.forEach(function(mDef){
    sh.getRange(currentRow, 1).setValue(mDef.name).setFontWeight('bold').setFontSize(12); currentRow++;
    sh.getRange(currentRow, 1).setValue('Day/Hr'); for(var h=0; h<24; h++) { sh.getRange(currentRow, h+2).setValue(h); }
    setHeaderRowFormatting(sh, currentRow);
    var dataGrid = [];
    DAYS.forEach(function(day){
        var rVals = [day];
        for(var hr=0; hr<24; hr++){
            var cellVal; var hData = hourlyData[day][hr];
            if(mDef.agg === 'SUM'){ cellVal = hData[mDef.field]; if(mDef.type === 'COST') cellVal = microsToCurrency(cellVal); }
            else if (mDef.agg === 'RATE'){
                var num = hData[mDef.num], den = hData[mDef.den];
                cellVal = (den > 0) ? (num / den) : 0;
                if(mDef.type === 'CURRENCY' && mDef.num === 'metrics.cost_micros') cellVal = (den > 0) ? microsToCurrency(num) / den : 0;
            }
            rVals.push(cellVal !== undefined ? cellVal : 0);
        }
        dataGrid.push(rVals);
    });
    sh.getRange(currentRow+1, 1, DAYS.length, 25).setValues(dataGrid);
    var valRng = sh.getRange(currentRow+1, 2, DAYS.length, 24);
    if (mDef.type === 'PERCENT') formatPercent(valRng);
    else if (mDef.type === 'CURRENCY' || mDef.type === 'COST') formatCurrency(valRng);
    else formatInteger(valRng);
    applyGradient(sh, valRng);
    currentRow += DAYS.length + 2;
  });
  autoResizeAllColumns(sh);
}
// End of SECTION 6

/*********************************************************************
 * SECTION 7: Channel Performance Tabs - GAQL
 *********************************************************************/
function buildChannelTabs(ss, cfg, summaryData) {
  var channels = [
    { type: 'SEARCH', name: 'Channel – Search', color: TAB_COLOR_CHANNELS }, 
    { type: 'DISPLAY', name: 'Channel – Display', color: TAB_COLOR_CHANNELS },
    { type: 'VIDEO', name: 'Channel – YouTube', color: TAB_COLOR_CHANNELS }, 
    { type: 'SHOPPING', name: 'Channel – Shopping', color: TAB_COLOR_CHANNELS },
    { type: 'PERFORMANCE_MAX', name: 'Channel – PMax', color: TAB_COLOR_CHANNELS },
    { type: 'MULTI_CHANNEL', name: 'Channel – App Installs', subtypeFilter: "campaign.advertising_channel_sub_type = 'APP_CAMPAIGN'", color: TAB_COLOR_CHANNELS },
    { type: 'MULTI_CHANNEL', name: 'Channel – App Engagement', subtypeFilter: "campaign.advertising_channel_sub_type = 'APP_CAMPAIGN_FOR_ENGAGEMENT'", color: TAB_COLOR_CHANNELS }
  ];
  channels.forEach(function(ch) { safeRun(function() { channelTab_(ss, cfg, ch.type, ch.name, ch.subtypeFilter, summaryData, ch.color); }, 'channelTab_' + ch.name.replace(/[^A-Za-z0-9]/g, '_'), ss, cfg, summaryData); });
}

function channelTab_(ss, cfg, adChannelType, sheetName, subtypeFilter, summaryData, tabColor) {
  var sh = resetSheet(ss, sheetName);
  setSheetTitle(sh, sheetName.toUpperCase() + ' - CAMPAIGN PERFORMANCE');
  if(tabColor) sh.setTabColor(tabColor);
  var headers = ['Campaign', 'Status', 'Impr', 'Clicks', 'CTR', 'Cost', 'Conv', 'Conv Val', 'Avg CPC', 'CPA', 'ROAS', 'Recommendation'];
  var headerRowIndex = 3;
  sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var where = ["campaign.advertising_channel_type = '" + adChannelType + "'", "segments.date BETWEEN '" + dateFromGAQL + "' AND '" + dateToGAQL + "'", "metrics.impressions > 0"];
  if (subtypeFilter) where.push(subtypeFilter);
  var gaql = 'SELECT campaign.name, campaign.status, metrics.impressions, metrics.clicks, metrics.ctr, metrics.cost_micros, metrics.conversions, metrics.conversions_value, metrics.average_cpc FROM campaign WHERE ' + where.join(' AND ');
  Logger.log(sheetName + ' GAQL: ' + gaql);
  
  var campDataForAvg = [];
  try {
      var tempIter = AdsApp.report(gaql).rows();
      while(tempIter.hasNext()){
          var r = tempIter.next();
          campDataForAvg.push({cost: microsToCurrency(Number(r['metrics.cost_micros'] || 0)), conv: Number(r['metrics.conversions'] || 0)});
      }
  } catch (e) {
      logError('channelTab_PreCalc_' + sheetName, e, gaql);
  }
  var chanTotalCost = campDataForAvg.reduce((s,d)=>s+d.cost,0), chanTotalConv = campDataForAvg.reduce((s,d)=>s+d.conv,0);
  var chanAvgCpa = chanTotalConv > 0 ? chanTotalCost / chanTotalConv : (summaryData.accountAvgCpa || cfg.TARGET_CPA);

  var outRows = [];
  try {
      var iter = AdsApp.report(gaql).rows();
      while (iter.hasNext()) {
        var r = iter.next();
        var name = r['campaign.name'], status = r['campaign.status'], impr = Number(r['metrics.impressions'] || 0), clk = Number(r['metrics.clicks'] || 0);
        var ctr = Number(r['metrics.ctr'] || 0), cost = microsToCurrency(Number(r['metrics.cost_micros'] || 0)), conv = Number(r['metrics.conversions'] || 0);
        var convVal = Number(r['metrics.conversions_value'] || 0), avgCpc = microsToCurrency(Number(r['metrics.average_cpc'] || 0)); // Corrected convVal
        var cpa = conv > 0 ? cost/conv : 0, roas = cost > 0 ? convVal/cost : 0;
        var rec = '', color = COLOR_NEUTRAL;

        if (status !== 'ENABLED' && status !== 'PAUSED') { rec = 'Status: ' + status; color = COLOR_HIGHLIGHT; }
        else if (status === 'PAUSED') { rec = 'Paused - Review KPIs'; color = COLOR_HIGHLIGHT; }
        else {
          var issues = [];
          if (ctr < cfg.CTR_LOW && clk > 0) issues.push('Low CTR');
          if (cpa > (chanAvgCpa * cfg.CPA_MULTIPLIER) && chanAvgCpa > 0) issues.push('High CPA vs Channel');
          else if (cpa > cfg.TARGET_CPA && cfg.TARGET_CPA > 0) issues.push('CPA > Target');
          if (conv === 0 && cost > 0 && clk >= cfg.CLICK_THRESHOLD) issues.push('No Conv (Cost > 0)');
          if (roas > 0 && roas < cfg.TARGET_ROAS && cfg.TARGET_ROAS > 0) issues.push('ROAS < Target');
          
          if (issues.length > 0) { rec = issues.join('; '); color = COLOR_BAD; }
          else if (conv > 0) { rec = 'Good Performance - Monitor/Scale'; color = COLOR_GOOD; }
          else if (cost > 0) { rec = 'Monitor (No Conv yet)'; color = COLOR_REVIEW; }
          else { rec = 'Low Data/No Cost'; color = COLOR_HIGHLIGHT; }
        }
        outRows.push([name, status, impr, clk, ctr, cost, conv, convVal, avgCpc, cpa, roas, rec.trim() || 'Review KPIs']);
        if(outRows.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outRows.length, headers.length).setBackground(color); }
      }
  } catch (e) {
      logError('channelTab_' + sheetName, e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching data.'); return;
  }

  if (outRows.length > 0) {
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow, 1, outRows.length, headers.length).setValues(outRows);
    var lr = outRows.length;
    formatInteger(sh.getRange(dataStartRow,3,lr,1)); formatInteger(sh.getRange(dataStartRow,4,lr,1)); 
    formatPercent(sh.getRange(dataStartRow,5,lr,1)); formatCurrency(sh.getRange(dataStartRow,6,lr,1));
    formatInteger(sh.getRange(dataStartRow,7,lr,1)); formatCurrency(sh.getRange(dataStartRow,8,lr,1)); 
    formatCurrency(sh.getRange(dataStartRow,9,lr,1)); formatCurrency(sh.getRange(dataStartRow,10,lr,1)); 
    formatPercent(sh.getRange(dataStartRow,11,lr,1));
    autoResizeAllColumns(sh);
  } else { 
    sh.getRange(headerRowIndex+1,1).setValue('No campaigns found for this channel type or date range.');
    try {
        if (ss.getSheets().length > 1) {
          ss.deleteSheet(sh);
          Logger.log('Deleted empty sheet: ' + sheetName);
        }
    } catch (e) {
        logError('channelTab_deleteEmptySheet', 'Could not delete empty sheet: ' + sheetName + ' - ' + e.message);
    }
  }
}
// End of SECTION 7

/*********************************************************************
 * SECTION 8: Search Terms Tab - GAQL
 *********************************************************************/
function buildSearchTerms_(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Search Terms Report');
  setSheetTitle(sh, 'SEARCH TERMS ANALYSIS');
  if (TAB_COLOR_DEEPDIVES) sh.setTabColor(TAB_COLOR_DEEPDIVES);
  var headers = ['Search Term', 'ST Status', 'Campaign', 'Ad Group', 'Matched KW', 'KW Match', 'ST Match', 'Impr', 'Clicks', 'CTR', 'Cost', 'Avg CPC', 'Conv', 'Conv Rate', 'CPA', 'Conv Val', 'ROAS', 'Brand?', 'Action'];
  var headerRowIndex = 3;
  sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var gaql = 'SELECT search_term_view.search_term, search_term_view.status, campaign.name, ad_group.name, segments.keyword.info.text, segments.keyword.info.match_type, segments.search_term_match_type, metrics.impressions, metrics.clicks, metrics.ctr, metrics.cost_micros, metrics.average_cpc, metrics.conversions, metrics.conversions_value FROM search_term_view WHERE segments.date BETWEEN \'' + dateFromGAQL + '\' AND \'' + dateToGAQL + '\' AND metrics.clicks > 0 AND campaign.advertising_channel_type = \'SEARCH\'';
  Logger.log('Search Terms GAQL: ' + gaql);
  var iter; try { iter = AdsApp.report(gaql).rows(); } catch (e) { logError('buildSearchTerms_GAQL', e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching data.'); return; }
  
  var outRows = []; var wastedSpendPotential = 0;
  while (iter.hasNext()) {
    var r = iter.next();
    var term = r['search_term_view.search_term'], stStatus = r['search_term_view.status'], camp = r['campaign.name'], ag = r['ad_group.name'];
    var kw = r['segments.keyword.info.text']||'N/A', kwm = r['segments.keyword.info.match_type']||'N/A', stm = r['segments.search_term_match_type']||'N/A';
    var impr = Number(r['metrics.impressions'] || 0), clk = Number(r['metrics.clicks'] || 0), ctr = Number(r['metrics.ctr'] || 0);
    var cost = microsToCurrency(Number(r['metrics.cost_micros'] || 0)), avgCpc = microsToCurrency(Number(r['metrics.average_cpc'] || 0));
    var conv = Number(r['metrics.conversions'] || 0), convVal = Number(r['metrics.conversions_value'] || 0); // Corrected convVal
    var cr = clk > 0 ? conv / clk : 0;
    var cpa = conv>0?cost/conv:0, roas = cost>0?convVal/cost:0;
    var isBrand = cfg.BRAND_PATTERNS.some(function(rx){ return rx.test(term); });
    var action = '', color = COLOR_NEUTRAL;

    if (stStatus === 'ADDED' || stStatus === 'EXCLUDED') { action = 'Status: ' + stStatus; color = COLOR_HIGHLIGHT; }
    else {
      if (conv > 0) { action = 'Add as Positive'; color = COLOR_GOOD; }
      else if (cost > 0 && clk >= cfg.CLICK_THRESHOLD) { action = 'Add as Negative'; color = COLOR_BAD; wastedSpendPotential += cost; }
      else if (isBrand && clk > 0 && cost > 0) { action = 'Review (Brand ST)'; color = COLOR_REVIEW; }
      else if (clk > 0 && cost > 0) { action = 'Monitor'; color = COLOR_REVIEW; }
      else { action = 'Low Data'; }
    }
    outRows.push([term, stStatus, camp, ag, kw, kwm, stm, impr, clk, ctr, cost, avgCpc, conv, cr, cpa, convVal, roas, isBrand?'Yes':'No', action]);
    if(outRows.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outRows.length, headers.length).setBackground(color); }
  }
  if (outRows.length > 0) {
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow, 1, outRows.length, headers.length).setValues(outRows);
    var col=8; formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); // Impr, Clicks
    formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1));
    autoResizeAllColumns(sh);
    if (wastedSpendPotential > 0) summaryData.opportunities.push('Potential wasted spend of ' + CURRENCY_SYMBOL + wastedSpendPotential.toFixed(2) + ' on non-converting search terms. Review Search Terms tab.');
  } else { sh.getRange(headerRowIndex+1,1).setValue('No search terms with clicks found.'); }
}
// End of SECTION 8

/*********************************************************************
 * SECTION 9: Keyword Deep Dive - GAQL
 *********************************************************************/
function deepSearch_KeywordDeepDive(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Keyword Deep Dive');
  setSheetTitle(sh, 'ALL SEARCH KEYWORDS - DEEP DIVE ANALYSIS');
  if (TAB_COLOR_DEEPDIVES) sh.setTabColor(TAB_COLOR_DEEPDIVES);
  var headers = ['Campaign', 'Ad Group', 'Keyword', 'Match Type', 'Status', 'Impr', 'Clicks', 'CTR', 'Cost', 'Avg CPC', 'Conv', 'Conv Val', 'CPA', 'ROAS', 'Brand?', 'Action'];
  var headerRowIndex = 3;
  sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var gaql = 'SELECT campaign.name, ad_group.name, ad_group_criterion.keyword.text, ad_group_criterion.keyword.match_type, ad_group_criterion.status, metrics.impressions, metrics.clicks, metrics.ctr, metrics.cost_micros, metrics.average_cpc, metrics.conversions, metrics.conversions_value FROM keyword_view WHERE segments.date BETWEEN \'' + dateFromGAQL + '\' AND \'' + dateToGAQL + '\' AND metrics.impressions > 0 AND campaign.advertising_channel_type = \'SEARCH\'';
  Logger.log('Keyword Deep Dive GAQL: ' + gaql);
  var iter; try { iter = AdsApp.report(gaql).rows(); } catch (e) { logError('deepSearch_KeywordDeepDive_GAQL', e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching data.'); return; }
  
  var outRows = []; var nonConvSpend = 0;
  while (iter.hasNext()) {
    var r = iter.next();
    var camp = r['campaign.name'], ag = r['ad_group.name'], kw = r['ad_group_criterion.keyword.text'], mt = r['ad_group_criterion.keyword.match_type'], status = r['ad_group_criterion.status'];
    var impr = Number(r['metrics.impressions'] || 0), clk = Number(r['metrics.clicks'] || 0), ctr = Number(r['metrics.ctr'] || 0);
    var cost = microsToCurrency(Number(r['metrics.cost_micros'] || 0)), avgCpc = microsToCurrency(Number(r['metrics.average_cpc'] || 0));
    var conv = Number(r['metrics.conversions'] || 0), convVal = Number(r['metrics.conversions_value'] || 0); // Corrected convVal
    var cpa = conv>0?cost/conv:0, roas = cost>0?convVal/cost:0;
    var isBrand = cfg.BRAND_PATTERNS.some(function(rx){ return rx.test(kw); });
    var action = '', color = COLOR_NEUTRAL;

    if (status !== 'ENABLED') { action = 'Status: ' + status; color = COLOR_HIGHLIGHT; }
    else {
      if (conv > 0) {
        action = 'Keep'; color = COLOR_GOOD;
        if (roas > 0 && roas < cfg.TARGET_ROAS) { action += ' (ROAS < Target)'; color = COLOR_REVIEW;}
        if (cpa > 0 && cpa > cfg.TARGET_CPA) { action += ' (CPA > Target)'; color = COLOR_REVIEW;}
      } else if (cost > 0 && clk >= cfg.CLICK_THRESHOLD) { action = 'Pause/Review (No Conv)'; color = COLOR_BAD; nonConvSpend += cost; }
      else if (isBrand && clk > 0 && cost > 0) { action = 'Review (Brand KW)'; color = COLOR_REVIEW; }
      else if (cost > 0) { action = 'Monitor'; color = COLOR_REVIEW; }
      else { action = 'Low Data'; }
    }
    outRows.push([camp, ag, kw, mt, status, impr, clk, ctr, cost, avgCpc, conv, convVal, cpa, roas, isBrand?'Yes':'No', action]);
    if(outRows.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outRows.length, headers.length).setBackground(color); }
  }
  if (outRows.length > 0) {
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow, 1, outRows.length, headers.length).setValues(outRows);
    var col=6; formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); // Impr, Clicks
    formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1));
    autoResizeAllColumns(sh);
    if (nonConvSpend > 0) summaryData.opportunities.push('Keywords with ' + CURRENCY_SYMBOL + nonConvSpend.toFixed(2) + ' spend and no conversions. Review Keyword Deep Dive.');
  } else { sh.getRange(headerRowIndex+1,1).setValue('No keyword data found.'); }
}
// End of SECTION 9

/*********************************************************************
 * SECTION 10: Ads Performance & Per-Campaign Keywords (All GAQL)
 *********************************************************************/
function buildAdsTab(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Ads Performance (Search)');
  setSheetTitle(sh, 'SEARCH ADS PERFORMANCE ANALYSIS');
  if (TAB_COLOR_DEEPDIVES) sh.setTabColor(TAB_COLOR_DEEPDIVES);
  var headers = ['Campaign', 'Ad Group', 'Ad Type', 'Ad Status', 'Impr', 'Clicks', 'CTR', 'Cost', 'Conv', 'CPA', 'Avg CPC', 'Recommendation'];
  var headerRowIndex = 3; sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);
  
  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var gaql = 'SELECT campaign.name, ad_group.name, ad_group_ad.ad.type, ad_group_ad.status, metrics.impressions, metrics.clicks, metrics.ctr, metrics.cost_micros, metrics.conversions, metrics.average_cpc FROM ad_group_ad WHERE campaign.advertising_channel_type = \'SEARCH\' AND ad_group_ad.status IN (\'ENABLED\', \'PAUSED\') AND segments.date BETWEEN \''+dateFromGAQL+'\' AND \''+dateToGAQL+'\' AND metrics.impressions > 0';
  Logger.log('Ads Tab GAQL: ' + gaql);
  var iter; try { iter = AdsApp.report(gaql).rows(); } catch (e) { logError('buildAdsTab_GAQL', e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching data.'); return; }
  
  var outData = [], adsWithLowCTR = 0;
  while (iter.hasNext()) {
    var r = iter.next();
    var camp = r['campaign.name'], ag = r['ad_group.name'], type = r['ad_group_ad.ad.type'], status = r['ad_group_ad.status'];
    var impr = Number(r['metrics.impressions'] || 0), clk = Number(r['metrics.clicks'] || 0), ctr = Number(r['metrics.ctr'] || 0);
    var cost = microsToCurrency(Number(r['metrics.cost_micros'] || 0)), conv = Number(r['metrics.conversions'] || 0), avgCpc = microsToCurrency(Number(r['metrics.average_cpc'] || 0));
    var cpa = conv>0?cost/conv:0; var rec = '', color = COLOR_NEUTRAL;

    if (status !== 'ENABLED') { rec = 'Status: ' + status; color = COLOR_HIGHLIGHT; }
    else {
      if (conv > 0 && ctr >= cfg.CTR_LOW) { rec = 'Keep/Scale'; color = COLOR_GOOD; }
      else if (cost > 0 && clk >= cfg.CLICK_THRESHOLD && conv === 0) { rec = 'Pause/Rewrite (No Conv)'; color = COLOR_BAD; }
      else if (cost > 0 && ctr < cfg.CTR_LOW) { rec = 'Review (Low CTR)'; color = COLOR_REVIEW; adsWithLowCTR++; }
      else if (cost > 0) { rec = 'Monitor'; color = COLOR_REVIEW; }
      else { rec = 'Low Data'; }
    }
    outData.push([camp, ag, type, status, impr, clk, ctr, cost, conv, cpa, avgCpc, rec]);
    if(outData.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outData.length, headers.length).setBackground(color); }
  }
  if (outData.length > 0) {
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow,1,outData.length,headers.length).setValues(outData);
    var col=5; formatInteger(sh.getRange(dataStartRow,col++,outData.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outData.length,1)); // Impr, Clicks
    formatPercent(sh.getRange(dataStartRow,col++,outData.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outData.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outData.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outData.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outData.length,1));
    autoResizeAllColumns(sh);
    if(adsWithLowCTR > 5) summaryData.opportunities.push(adsWithLowCTR + ' Search ads have low CTR. Review Ads Performance tab.');
  } else { sh.getRange(headerRowIndex+1,1).setValue('No Ad data found.'); }
}

function buildCampaignKW_(ss, cfg, summaryData) {
  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var searchCamps = {};
  var campGaql = 'SELECT campaign.id, campaign.name FROM campaign WHERE campaign.status IN (\'ENABLED\', \'PAUSED\') AND campaign.advertising_channel_type = \'SEARCH\' AND metrics.cost_micros > 0 AND segments.date BETWEEN \'' + dateFromGAQL + '\' AND \'' + dateToGAQL + '\'';
  Logger.log('Campaign KW - Camp List GAQL: ' + campGaql);
  try { 
      var campRows = AdsApp.report(campGaql).rows();
      while(campRows.hasNext()){
          var r = campRows.next();
          searchCamps[r['campaign.id']] = r['campaign.name'];
      }
  }
  catch (e) { logError('buildCampaignKW_getCampaigns', e, campGaql); return; }

  var usedNames = {};
  for (var cid in searchCamps) {
    if (searchCamps.hasOwnProperty(cid)) {
      var campName = searchCamps[cid];
      var baseName = 'KW-' + campName.substring(0,35).replace(/[^A-Za-z0-9\-_]/g,'').trim();
      var sheetName = baseName, idx=1;
      while (ss.getSheetByName(sheetName) || usedNames[sheetName.toLowerCase()]) { sheetName = baseName.substring(0, 85-(String(idx).length+1)) + '_'+(idx++); }
      usedNames[sheetName.toLowerCase()] = true;

      var sh = resetSheet(ss, sheetName);
      setSheetTitle(sh, "KEYWORD ANALYSIS: " + campName.substring(0,50));
      var headers = ['Ad Group','Keyword','Match','Status','Impr','Clicks','CTR','Cost','AvgCPC','Conv','ConvVal','CPA','ROAS','Brand?','Action'];
      var headerRowIndex = 3; sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
      setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

      var kwGaql = 'SELECT ad_group.name, ad_group_criterion.keyword.text, ad_group_criterion.keyword.match_type, ad_group_criterion.status, metrics.impressions, metrics.clicks, metrics.ctr, metrics.cost_micros, metrics.average_cpc, metrics.conversions, metrics.conversions_value FROM keyword_view WHERE campaign.id = ' + cid + ' AND segments.date BETWEEN \'' + dateFromGAQL + '\' AND \'' + dateToGAQL + '\' AND metrics.impressions > 0';
      try {
        var kwRows = AdsApp.report(kwGaql).rows(); var data=[];
        while (kwRows.hasNext()) {
          var r = kwRows.next();
          var ag=r['ad_group.name'], kw=r['ad_group_criterion.keyword.text'], mt=r['ad_group_criterion.keyword.match_type'], st=r['ad_group_criterion.status'];
          var impr=Number(r['metrics.impressions'] || 0), clk=Number(r['metrics.clicks'] || 0), ctr=Number(r['metrics.ctr'] || 0);
          var cost=microsToCurrency(Number(r['metrics.cost_micros'] || 0)), avgCpc=microsToCurrency(Number(r['metrics.average_cpc'] || 0));
          var conv=Number(r['metrics.conversions'] || 0), val=Number(r['metrics.conversions_value'] || 0); // Corrected val
          var cpa=conv>0?cost/conv:0, roas=cost>0?val/cost:0;
          var isB=cfg.BRAND_PATTERNS.some(rx=>rx.test(kw)); var act='', color=COLOR_NEUTRAL;

          if(st.toLowerCase()!=='enabled'){act='Status: '+st;color=COLOR_HIGHLIGHT;}
          else if(conv>0){act='Keep';color=COLOR_GOOD;}
          else if(cost>0 && clk>=cfg.CLICK_THRESHOLD){act='Pause/Review (No Conv)';color=COLOR_BAD;}
          else if(isB && clk>0 && cost>0){act='Review (Brand)';color=COLOR_REVIEW;}
          else if(cost>0){act='Monitor';color=COLOR_REVIEW;} else{act='Low Data';}
          data.push([ag,kw,mt,st,impr,clk,ctr,cost,avgCpc,conv,val,cpa,roas,isB?'Y':'N',act]);
          if(data.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + data.length, headers.length).setBackground(color); }
        }
        if (!data.length) { sh.getRange(headerRowIndex+1,1).setValue("No keyword data found for this campaign."); continue; }
        
        var dataStartRow = headerRowIndex + 1;
        sh.getRange(dataStartRow,1,data.length,headers.length).setValues(data);
        var col=5; formatInteger(sh.getRange(dataStartRow,col++,data.length,1)); formatInteger(sh.getRange(dataStartRow,col++,data.length,1)); // Impr, Clicks
        formatPercent(sh.getRange(dataStartRow,col++,data.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,data.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,data.length,1)); formatInteger(sh.getRange(dataStartRow,col++,data.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,data.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,data.length,1)); formatPercent(sh.getRange(dataStartRow,col++,data.length,1));
        autoResizeAllColumns(sh);
      } catch (e) { logError('buildCampaignKW_report_' + campName, e, kwGaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching KWs.'); }
    }
  }
}
// End of SECTION 10

/*********************************************************************
 * SECTION 11: Quality Score Deep Dive - With Enhanced Formatting
 *********************************************************************/
function buildQualityScoreReport_(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Quality Score (Search)');
  setSheetTitle(sh, 'KEYWORD QUALITY SCORE ANALYSIS (SEARCH)');
  if (TAB_COLOR_DEEPDIVES) sh.setTabColor(TAB_COLOR_DEEPDIVES);
  var headers = ['Campaign', 'Ad Group', 'Keyword', 'Status', 'QS', 'Exp. CTR', 'Ad Rel.', 'LP Exp.', 'Impr.', 'Clicks', 'Cost', 'Recommendation'];
  var headerRowIndex = 3; sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var gaql = 'SELECT campaign.name, ad_group.name, ad_group_criterion.keyword.text, ad_group_criterion.status, ' +
             'ad_group_criterion.quality_info.quality_score, metrics.search_impression_share, ' +
             'ad_group_criterion.quality_info.creative_quality_score, ' +
             'ad_group_criterion.quality_info.post_click_quality_score, ' + // Corrected field name
             'ad_group_criterion.quality_info.search_predicted_ctr, ' +
             'metrics.impressions, metrics.clicks, metrics.cost_micros ' +
             'FROM keyword_view WHERE segments.date BETWEEN \'' + dateFromGAQL + '\' AND \'' + dateToGAQL + '\' ' +
             'AND campaign.advertising_channel_type = \'SEARCH\' AND ad_group_criterion.status = \'ENABLED\' ' +
             'AND metrics.impressions > 10 ORDER BY metrics.cost_micros DESC LIMIT 500';
  Logger.log('Quality Score GAQL: ' + gaql);
  var iter; try { iter = AdsApp.report(gaql).rows(); } catch (e) { logError('buildQualityScoreReport_GAQL', e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching QS data.'); return; }

  var outRows = [], lowQsHighSpendCount = 0;
  var qualityEnumMap = { 'BELOW_AVERAGE': 'Below Avg', 'AVERAGE': 'Avg', 'ABOVE_AVERAGE': 'Above Avg', 'UNKNOWN': 'Unknown', 'UNSPECIFIED': 'Unspec.'};

  while(iter.hasNext()){
    var r = iter.next();
    var camp = r['campaign.name'], ag = r['ad_group.name'], kw = r['ad_group_criterion.keyword.text'], status = r['ad_group_criterion.status'];
    var qs = r['ad_group_criterion.quality_info.quality_score'] !== null ? Number(r['ad_group_criterion.quality_info.quality_score']) : 'N/A';
    var expCtr = qualityEnumMap[r['ad_group_criterion.quality_info.search_predicted_ctr']] || r['ad_group_criterion.quality_info.search_predicted_ctr'] || 'N/A';
    var adRel = qualityEnumMap[r['ad_group_criterion.quality_info.creative_quality_score']] || r['ad_group_criterion.quality_info.creative_quality_score'] || 'N/A';
    var lpExp = qualityEnumMap[r['ad_group_criterion.quality_info.post_click_quality_score']] || r['ad_group_criterion.quality_info.post_click_quality_score'] || 'N/A';
    var impr = Number(r['metrics.impressions'] || 0), clk = Number(r['metrics.clicks'] || 0), cost = microsToCurrency(Number(r['metrics.cost_micros'] || 0));
    var searchIS = r['metrics.search_impression_share'] !== null ? (Number(r['metrics.search_impression_share'])*100).toFixed(1)+'%' : 'N/A';
    var rec = [], color = COLOR_NEUTRAL;

    if (qs !== 'N/A' && qs <= cfg.QS_LOW_THRESHOLD) {
        color = COLOR_BAD; rec.push('Low QS ('+qs+').');
        if (expCtr === 'Below Avg') rec.push('Improve CTR.');
        if (adRel === 'Below Avg') rec.push('Improve Ad Rel.');
        if (lpExp === 'Below Avg') rec.push('Improve LP Exp.');
        if (cost > 100 && searchIS !== 'N/A' && parseFloat(searchIS) > 30) lowQsHighSpendCount++;
    } else if (qs !== 'N/A' && qs >=8) { color = COLOR_GOOD; rec.push('Good QS ('+qs+').'); }
    else { color = COLOR_REVIEW; rec.push('Monitor QS.'); }
    if (rec.length === 0) rec.push('N/A');
    
    outRows.push([camp,ag,kw,status,qs,expCtr,adRel,lpExp,impr,clk,cost,rec.join(' ')]);
    if(outRows.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outRows.length, headers.length).setBackground(color); }
  }
  if(outRows.length > 0){
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow,1,outRows.length,headers.length).setValues(outRows);
    var col=5; formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); // QS
    col=9; formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); // Impr, Clicks
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); // Cost
    autoResizeAllColumns(sh);
    if(lowQsHighSpendCount > 0) summaryData.opportunities.push(lowQsHighSpendCount + ' keywords have Low QS & significant spend/IS. Review Quality Score tab.');
  } else { sh.getRange(headerRowIndex+1,1).setValue('No keyword data with QS found (min 10 impr, enabled status).'); }
}
// End of SECTION 11

/*********************************************************************
 * SECTION 12: Impression Share Analysis - With Enhanced Formatting
 *********************************************************************/
function buildSearchISReport_(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Search IS Analysis');
  setSheetTitle(sh, 'SEARCH CAMPAIGN IMPRESSION SHARE ANALYSIS');
  if (TAB_COLOR_DEEPDIVES) sh.setTabColor(TAB_COLOR_DEEPDIVES);
  var headers = ['Campaign', 'Status', 'Search IS %', 'Top IS %', 'Abs. Top IS %', 'Lost IS (Rank) %', 'Lost IS (Budget) %', 'Cost', 'Conv', 'ROAS', 'Recommendation'];
  var headerRowIndex = 3; sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var gaql = 'SELECT campaign.name, campaign.status, metrics.search_impression_share, metrics.search_top_impression_share, metrics.search_absolute_top_impression_share, metrics.search_budget_lost_impression_share, metrics.search_rank_lost_impression_share, metrics.cost_micros, metrics.conversions, metrics.conversions_value FROM campaign WHERE segments.date BETWEEN \'' + dateFromGAQL + '\' AND \'' + dateToGAQL + '\' AND campaign.advertising_channel_type = \'SEARCH\' AND metrics.impressions > 0 ORDER BY metrics.search_budget_lost_impression_share DESC, metrics.search_rank_lost_impression_share DESC';
  Logger.log('Search IS GAQL: ' + gaql);
  var iter; try { iter = AdsApp.report(gaql).rows(); } catch (e) { logError('buildSearchISReport_GAQL', e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching IS data.'); return; }

  var outRows = [], budgetConstrainedHighPerf = 0;
  while(iter.hasNext()){
    var r = iter.next();
    var camp = r['campaign.name'], status = r['campaign.status'];
    var searchIS = r['metrics.search_impression_share'] !== null ? Number(r['metrics.search_impression_share']) : null;
    var topIS = r['metrics.search_top_impression_share'] !== null ? Number(r['metrics.search_top_impression_share']) : null;
    var absTopIS = r['metrics.search_absolute_top_impression_share'] !== null ? Number(r['metrics.search_absolute_top_impression_share']) : null;
    var lostBudget = r['metrics.search_budget_lost_impression_share'] !== null ? Number(r['metrics.search_budget_lost_impression_share']) : null;
    var lostRank = r['metrics.search_rank_lost_impression_share'] !== null ? Number(r['metrics.search_rank_lost_impression_share']) : null;
    var cost = microsToCurrency(Number(r['metrics.cost_micros'] || 0)), conv = Number(r['metrics.conversions'] || 0), convVal = Number(r['metrics.conversions_value'] || 0); // Corrected convVal
    var roas = cost > 0 ? convVal/cost : 0;
    var cpa = conv > 0 ? cost/conv : 0;
    var rec = [], color = COLOR_NEUTRAL;

    if (status !== 'ENABLED') { rec.push('Status: ' + status); color = COLOR_HIGHLIGHT; }
    else {
        if (lostBudget !== null && lostBudget > 0.1) {
            rec.push('High Lost IS (Budget): ' + (lostBudget*100).toFixed(1) + '%. Consider budget.');
            color = COLOR_BAD;
            if (roas > cfg.TARGET_ROAS || (conv > 0 && cpa > 0 && cpa < cfg.TARGET_CPA) ) budgetConstrainedHighPerf++;
        }
        if (lostRank !== null && lostRank > 0.2) {
            rec.push('High Lost IS (Rank): ' + (lostRank*100).toFixed(1) + '%. Review Ad Rank.');
            color = color === COLOR_BAD ? COLOR_BAD : COLOR_REVIEW;
        }
        if (searchIS !== null && searchIS*100 < cfg.IS_LOW_THRESHOLD_PERCENT) {
             rec.push('Low Search IS: '+(searchIS*100).toFixed(1)+'%');
             color = color === COLOR_BAD ? COLOR_BAD : COLOR_REVIEW;
        }
    }
    if(rec.length === 0) { rec.push('IS appears stable.'); color = COLOR_GOOD; }
    outRows.push([camp, status, searchIS, topIS, absTopIS, lostRank, lostBudget, cost, conv, roas, rec.join(' ')]);
    if(outRows.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outRows.length, headers.length).setBackground(color); }
  }
  if(outRows.length > 0){
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow,1,outRows.length,headers.length).setValues(outRows);
    var col=3; formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1));
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1));
    autoResizeAllColumns(sh);
    if(budgetConstrainedHighPerf > 0) summaryData.opportunities.push(budgetConstrainedHighPerf + ' high-performing Search campaigns are budget-constrained. Review Search IS tab.');
  } else { sh.getRange(headerRowIndex+1,1).setValue('No Search campaign data with impressions found.'); }
}
// End of SECTION 12

/*********************************************************************
 * SECTION 13: Audience Performance Insights - Basic Implementation
 *********************************************************************/
function buildAudienceReport_(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Audience Insights');
  setSheetTitle(sh, 'AUDIENCE PERFORMANCE INSIGHTS');
  if (TAB_COLOR_DEEPDIVES) sh.setTabColor(TAB_COLOR_DEEPDIVES);
  var headers = ['Campaign', 'Ad Group', 'Audience Name', 'Type', 'Impr', 'Clicks', 'Cost', 'Conv', 'CPA', 'ROAS', 'Recommendation'];
  var headerRowIndex = 3; sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var gaql = 'SELECT campaign.name, ad_group.name, ad_group_criterion.display_name, ad_group_criterion.type, ' +
             'metrics.impressions, metrics.clicks, metrics.cost_micros, metrics.conversions, metrics.conversions_value ' +
             'FROM ad_group_audience_view ' +
             'WHERE segments.date BETWEEN \'' + dateFromGAQL + '\' AND \'' + dateToGAQL + '\' '+
             'AND ad_group_criterion.status = \'ENABLED\' AND metrics.impressions > 0 ORDER BY metrics.cost_micros DESC LIMIT 500';
  Logger.log('Audience Insights GAQL: ' + gaql);
  var iter; try { iter = AdsApp.report(gaql).rows(); } catch (e) { logError('buildAudienceReport_GAQL', e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching Audience data.'); return; }

  var outRows = [];
  while(iter.hasNext()){
    var r = iter.next();
    var camp = r['campaign.name'] || 'N/A';
    var ag = r['ad_group.name'] || 'N/A';
    var audName = r['ad_group_criterion.display_name'] || '(Not Set)';
    var audType = r['ad_group_criterion.type'] || 'N/A';
    
    var impr = Number(r['metrics.impressions'] || 0), clk = Number(r['metrics.clicks'] || 0);
    var cost = microsToCurrency(Number(r['metrics.cost_micros'] || 0));
    var conv = Number(r['metrics.conversions'] || 0), convVal = Number(r['metrics.conversions_value'] || 0); // Corrected convVal
    var cpa = conv > 0 ? cost/conv : 0;
    var roas = cost > 0 ? convVal/cost : 0;
    var rec = 'Review Performance', color = COLOR_REVIEW;

    if (conv > 0 && roas >= cfg.TARGET_ROAS) { rec = 'Good Performance'; color = COLOR_GOOD; }
    else if (cost > 0 && conv === 0 && clk > cfg.CLICK_THRESHOLD) { rec = 'No Conv - High Clicks'; color = COLOR_BAD; }
    
    outRows.push([camp, ag, audName, audType, impr, clk, cost, conv, cpa, roas, rec]);
    if(outRows.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outRows.length, headers.length).setBackground(color); }
  }

  if(outRows.length > 0){
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow,1,outRows.length,headers.length).setValues(outRows);
    var col=5; 
    formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    autoResizeAllColumns(sh);
  } else { sh.getRange(headerRowIndex+1,1).setValue('No Audience data found with impressions.');}
}
// End of SECTION 13

/*********************************************************************
 * SECTION 14: Device Performance Report - Basic Implementation
 *********************************************************************/
function buildDeviceReport_(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Device Performance');
  setSheetTitle(sh, 'DEVICE PERFORMANCE OVERVIEW');
  if (TAB_COLOR_DEEPDIVES) sh.setTabColor(TAB_COLOR_DEEPDIVES);
  var headers = ['Device', 'Impr', 'Clicks', 'CTR', 'Cost', 'Conv', 'CPA', 'ROAS', 'Recommendation'];
  var headerRowIndex = 3; sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var gaql = 'SELECT segments.device, metrics.impressions, metrics.clicks, metrics.ctr, metrics.cost_micros, metrics.conversions, metrics.conversions_value ' +
             'FROM campaign WHERE segments.date BETWEEN \'' + dateFromGAQL + '\' AND \'' + dateToGAQL + '\' AND metrics.impressions > 0';
  Logger.log('Device Performance GAQL: ' + gaql);
  var iter; try { iter = AdsApp.report(gaql).rows(); } catch (e) { logError('buildDeviceReport_GAQL', e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching Device data.'); return; }

  var agg = {};
  while(iter.hasNext()){
    var r = iter.next();
    var device = r['segments.device'] || 'UNKNOWN';
    if (!agg[device]) agg[device] = { impr:0, clk:0, cost:0, conv:0, convVal:0 };
    agg[device].impr += Number(r['metrics.impressions'] || 0);
    agg[device].clk += Number(r['metrics.clicks'] || 0);
    agg[device].cost += microsToCurrency(Number(r['metrics.cost_micros'] || 0));
    agg[device].conv += Number(r['metrics.conversions'] || 0);
    agg[device].convVal += Number(r['metrics.conversions_value'] || 0); // Corrected convVal
  }
  
  var outRows = [];
  for (var deviceKey in agg) {
    if (agg.hasOwnProperty(deviceKey)) {
        var o = agg[deviceKey];
        var ctr = o.impr > 0 ? o.clk / o.impr : 0;
        var cpa = o.conv > 0 ? o.cost / o.conv : 0;
        var roas = o.cost > 0 ? o.convVal / o.cost : 0;
        var rec = 'Review device performance. Consider bid adjustments if significant CPA/ROAS variance.';
        var color = COLOR_REVIEW;
        if (roas > cfg.TARGET_ROAS * 1.1) { rec = 'Good ROAS on ' + deviceKey; color = COLOR_GOOD; }
        else if (cpa > cfg.TARGET_CPA * 1.2 && cpa > 0) { rec = 'High CPA on ' + deviceKey; color = COLOR_BAD;}

        outRows.push([deviceKey, o.impr, o.clk, ctr, o.cost, o.conv, cpa, roas, rec]);
        if(outRows.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outRows.length, headers.length).setBackground(color); }
    }
  }

  if(outRows.length > 0){
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow,1,outRows.length,headers.length).setValues(outRows);
    var col=2; 
    formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    autoResizeAllColumns(sh);
  } else { sh.getRange(headerRowIndex+1,1).setValue('No Device data found with impressions.');}
}
// End of SECTION 14

/*********************************************************************
 * SECTION 15: Ad Extension (Asset) Health - Basic Implementation
 *********************************************************************/
function buildExtensionReport_(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Asset (Extension) Health');
  setSheetTitle(sh, 'ASSET (EXTENSION) HEALTH & PERFORMANCE');
  if (TAB_COLOR_DEEPDIVES) sh.setTabColor(TAB_COLOR_DEEPDIVES);
  var headers = ['Campaign', 'Ad Group', 'Asset Type', 'Asset Text/Details', 'Field Type', 'Performance Label', 'Impr', 'Clicks', 'Cost', 'Recommendation'];
  var headerRowIndex = 3; sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var gaql = 'SELECT campaign.name, ad_group.name, asset.type, asset.text_asset.text, asset.image_asset.full_size.url, ' +
             'ad_group_ad_asset_view.field_type, ad_group_ad_asset_view.performance_label, ' +
             'metrics.impressions, metrics.clicks, metrics.cost_micros ' +
             'FROM ad_group_ad_asset_view ' + 
             'WHERE segments.date BETWEEN \'' + dateFromGAQL + '\' AND \'' + dateToGAQL + '\' '+
             'AND ad_group_ad_asset_view.status = \'ENABLED\' AND metrics.impressions > 0 ORDER BY metrics.cost_micros DESC LIMIT 300';
  Logger.log('Asset Health GAQL: ' + gaql);
  var iter; try { iter = AdsApp.report(gaql).rows(); } catch (e) { logError('buildExtensionReport_GAQL', e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching Asset data.'); return; }
  
  var outRows = [];
  while(iter.hasNext()){
    var r = iter.next();
    var camp = r['campaign.name'] || 'N/A';
    var ag = r['ad_group.name'] || 'N/A';
    var assetType = r['asset.type'] || 'N/A';
    var assetText = r['asset.text_asset.text'] || (r['asset.image_asset.full_size.url'] ? 'Image Asset' : '(No Text)');
    var fieldType = r['ad_group_ad_asset_view.field_type'] || 'N/A';
    var perfLabel = r['ad_group_ad_asset_view.performance_label'] || 'N/A'; 
    
    var impr = Number(r['metrics.impressions'] || 0), clk = Number(r['metrics.clicks'] || 0);
    var cost = microsToCurrency(Number(r['metrics.cost_micros'] || 0));
    var rec = 'Review Performance Label: ' + perfLabel, color = COLOR_REVIEW;

    if (perfLabel === 'GOOD' || perfLabel === 'BEST' || perfLabel === 'PENDING') { color = COLOR_GOOD; }
    else if (perfLabel === 'LOW' || perfLabel === 'POOR' || perfLabel === 'LEARNING') { color = COLOR_BAD; rec = 'Low/Learning Performance Label: ' + perfLabel + '. Consider replacing/improving.'; }
    
    outRows.push([camp, ag, assetType, assetText.substring(0,100), fieldType, perfLabel, impr, clk, cost, rec]);
    if(outRows.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outRows.length, headers.length).setBackground(color); }
  }

  if(outRows.length > 0){
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow,1,outRows.length,headers.length).setValues(outRows);
    var col=7; 
    formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    autoResizeAllColumns(sh);
  } else { sh.getRange(headerRowIndex+1,1).setValue('No Asset data found with impressions.');}
}
// End of SECTION 15

/*********************************************************************
 * SECTION 16: Change History Summary - Basic Implementation
 *********************************************************************/
function buildChangeHistoryReport_(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Change History Summary');
  setSheetTitle(sh, 'RECENT ACCOUNT CHANGE HISTORY SUMMARY (Report Dates: ' + cfg.DATE_FROM + ' to ' + cfg.DATE_TO + ')');
  if (TAB_COLOR_DEEPDIVES) sh.setTabColor(TAB_COLOR_DEEPDIVES);
  var headers = ['Change Date Time', 'User Email', 'Resource Type', 'Resource Change Operation', 'Changed Resource Name', 'Campaign'];
  var headerRowIndex = 3; sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM; 
  var dateToGAQL = cfg.DATE_TO;

  var gaql = 'SELECT change_event.change_date_time, change_event.user_email, change_event.change_resource_type, ' +
             'change_event.resource_change_operation, change_event.changed_fields, change_event.old_resource, change_event.new_resource, ' +
             'campaign.name ' +
             'FROM change_event ' +
             'WHERE change_event.change_date_time BETWEEN \'' + dateFromGAQL + ' 00:00:00\' AND \'' + dateToGAQL + ' 23:59:59\' ' +
             'ORDER BY change_event.change_date_time DESC LIMIT 200';
  Logger.log('Change History GAQL: ' + gaql);
  var iter; try { iter = AdsApp.report(gaql).rows(); } catch (e) { logError('buildChangeHistoryReport_GAQL', e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching Change History data.'); return; }
  
  var outRows = [];
  while(iter.hasNext()){
    var r = iter.next();
    var changeDateTime = r['change_event.change_date_time'];
    var userEmail = r['change_event.user_email'];
    var resourceType = r['change_event.change_resource_type'];
    var changeOperation = r['change_event.resource_change_operation']; 
    var changedResourceName = r['change_event.new_resource'] || r['change_event.old_resource'] || 'N/A'; 
    if (typeof changedResourceName === 'object') changedResourceName = JSON.stringify(changedResourceName).substring(0,100);
    var campaignName = r['campaign.name'] || 'N/A';

    outRows.push([changeDateTime, userEmail, resourceType, changeOperation, changedResourceName, campaignName]);
  }

  if(outRows.length > 0){
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow,1,outRows.length,headers.length).setValues(outRows);
    autoResizeAllColumns(sh);
  } else { sh.getRange(headerRowIndex+1,1).setValue('No Change History data found for the period: ' + cfg.DATE_FROM + ' to ' + cfg.DATE_TO + '.');}
}
// End of SECTION 16

/*********************************************************************
 * SECTION 17: Budget Pacing Report - Basic Implementation
 *********************************************************************/
function buildBudgetPacingReport_(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Budget Pacing');
  setSheetTitle(sh, 'CAMPAIGN BUDGET PACING REPORT');
  if (TAB_COLOR_DEEPDIVES) sh.setTabColor(TAB_COLOR_DEEPDIVES);
  var headers = ['Campaign', 'Status', 'Budget Amount (Daily)', 'Days In Period', 'Planned Budget (Period)', 'Delivery Method', 'Cost (Period)', '% Spent (vs Period Budget)', 'Pacing Status'];
  var headerRowIndex = 3; sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  
  var campaignBudgets = {};
  var budgetGaql = 'SELECT campaign.id, campaign.name, campaign.status, campaign_budget.amount_micros, campaign_budget.delivery_method ' +
                   'FROM campaign ' +
                   'WHERE campaign.status IN (\'ENABLED\', \'PAUSED\') AND campaign_budget.amount_micros > 0';
  Logger.log('Budget Pacing - Budget GAQL: ' + budgetGaql);
  try {
    var budgetIter = AdsApp.report(budgetGaql).rows();
    while(budgetIter.hasNext()) {
      var r = budgetIter.next();
      campaignBudgets[r['campaign.id']] = {
        name: r['campaign.name'],
        status: r['campaign.status'],
        dailyBudget: microsToCurrency(Number(r['campaign_budget.amount_micros'] || 0)),
        deliveryMethod: r['campaign_budget.delivery_method']
      };
    }
  } catch(e) {
    logError('buildBudgetPacingReport_BudgetFetch', e, budgetGaql);
    sh.getRange(headerRowIndex+1,1).setValue('Error fetching campaign budget data.');
    return;
  }

  var campaignCosts = {};
  var costGaql = 'SELECT campaign.id, metrics.cost_micros ' +
                 'FROM campaign ' +
                 'WHERE segments.date BETWEEN \''+dateFromGAQL+'\' AND \''+dateToGAQL+'\' '+
                 'AND campaign.status IN (\'ENABLED\', \'PAUSED\')';
  Logger.log('Budget Pacing - Cost GAQL: ' + costGaql);
   try {
    var costIter = AdsApp.report(costGaql).rows();
    while(costIter.hasNext()) {
      var r = costIter.next();
      var campId = r['campaign.id'];
      if (!campaignCosts[campId]) {
        campaignCosts[campId] = 0;
      }
      campaignCosts[campId] += microsToCurrency(Number(r['metrics.cost_micros'] || 0));
    }
  } catch(e) {
    logError('buildBudgetPacingReport_CostFetch', e, costGaql);
  }

  var outRows = [];
  var daysInPeriod = (new Date(dateToGAQL).getTime() - new Date(dateFromGAQL).getTime())/(1000*60*60*24) + 1;

  for (var campId in campaignBudgets) {
    if (campaignBudgets.hasOwnProperty(campId)) {
      var campData = campaignBudgets[campId];
      var costPeriod = campaignCosts[campId] || 0;
      var dailyBudgetAmount = campData.dailyBudget;
      var deliveryMethod = campData.deliveryMethod;
      
      var plannedPeriodBudget = 0;
      if (deliveryMethod === 'DAILY' && dailyBudgetAmount > 0) {
          plannedPeriodBudget = dailyBudgetAmount * daysInPeriod;
      } else if (dailyBudgetAmount > 0) { 
          plannedPeriodBudget = dailyBudgetAmount; 
      }
      
      var percentSpent = (plannedPeriodBudget > 0) ? costPeriod / plannedPeriodBudget : 0;
      var pacingStatus = 'Review', color = COLOR_REVIEW;

      if (percentSpent > 1.1 && deliveryMethod === 'DAILY') { pacingStatus = 'Overspending Period'; color = COLOR_BAD; }
      else if (percentSpent > 1.3 && deliveryMethod === 'ACCELERATED') { pacingStatus = 'Significantly Overspending Period (Accelerated)'; color = COLOR_BAD; }
      else if (percentSpent < 0.7 && costPeriod > 0) { pacingStatus = 'Underspending Period'; color = COLOR_REVIEW; }
      else if (costPeriod === 0 && dailyBudgetAmount > 0) { pacingStatus = 'No Spend in Period'; color = COLOR_HIGHLIGHT; }
      else if (dailyBudgetAmount > 0) { pacingStatus = 'On Pace (Approx for Period)'; color = COLOR_GOOD; }

      outRows.push([campData.name, campData.status, dailyBudgetAmount, daysInPeriod, plannedPeriodBudget, deliveryMethod, costPeriod, percentSpent, pacingStatus]);
      if(outRows.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outRows.length, headers.length).setBackground(color); }
    }
  }


  if(outRows.length > 0){
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow,1,outRows.length,headers.length).setValues(outRows);
    var col=3; 
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); // Daily Budget
    formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1));  // Days In Period
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); // Planned Period Budget
    col++; // Skip Delivery Method
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); // Cost (Period)
    formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); // % Spent
    autoResizeAllColumns(sh);
  } else { sh.getRange(headerRowIndex+1,1).setValue('No Campaign Budget data found or no campaigns with daily budgets.');}
}
// End of SECTION 17

/*********************************************************************
 * SECTION 18: PMax & Shopping Deep Dives + Executive Summary - With Enhanced Formatting
 *********************************************************************/
function deepPMax_AssetGroup(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'PMax Asset Groups');
  setSheetTitle(sh, 'PERFORMANCE MAX - ASSET GROUP ANALYSIS');
  if (TAB_COLOR_PRODUCTS) sh.setTabColor(TAB_COLOR_PRODUCTS);
  var headers = ['Campaign','Asset Group', 'Status', 'Impr','Clicks','CTR', 'Cost','Conv','Conv Val','CPA','ROAS', 'Recommendation'];
  var headerRowIndex = 3; sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);
  
  var dateFromGAQL = cfg.DATE_FROM; var dateToGAQL = cfg.DATE_TO;
  var gaql = 'SELECT campaign.name, asset_group.name, asset_group.status, metrics.impressions, metrics.clicks, metrics.ctr, metrics.cost_micros, metrics.conversions, metrics.conversions_value ' +
             'FROM asset_group WHERE campaign.advertising_channel_type = \'PERFORMANCE_MAX\' ' +
             'AND segments.date BETWEEN \''+dateFromGAQL+'\' AND \''+dateToGAQL+'\' AND metrics.impressions > 0';
  Logger.log('PMax Asset Groups GAQL: ' + gaql);
  var iter; try { iter = AdsApp.report(gaql).rows(); } catch (e) { logError('deepPMax_AssetGroup_GAQL', e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching data.'); return; }
  
  var outRows = [];
  while(iter.hasNext()) {
    var r = iter.next();
    var c = r['campaign.name'], ag = r['asset_group.name'], status = r['asset_group.status'];
    var impr = Number(r['metrics.impressions'] || 0), clk = Number(r['metrics.clicks'] || 0), ctr = Number(r['metrics.ctr'] || 0);
    var cost = microsToCurrency(Number(r['metrics.cost_micros'] || 0));
    var conv = Number(r['metrics.conversions'] || 0);
    var convVal = Number(r['metrics.conversions_value'] || 0); // Corrected
    var cpa = conv>0? cost/conv : 0, roas = cost>0? convVal/cost : 0;
    var rec = '', color = COLOR_NEUTRAL;

    if (status !== 'ENABLED') { rec = 'Status: ' + status; color = COLOR_HIGHLIGHT; }
    else {
        if (conv > 0 && roas >= cfg.TARGET_ROAS) { rec = 'Good Performance'; color = COLOR_GOOD; }
        else if (conv > 0 && roas < cfg.TARGET_ROAS) { rec = 'ROAS < Target ('+(roas*100).toFixed(0)+'%)'; color = COLOR_REVIEW; }
        else if (cost > 0 && clk > cfg.CLICK_THRESHOLD) { rec = 'No Conv - Review'; color = COLOR_BAD; }
        else if (cost > 0) { rec = 'Monitor'; color = COLOR_REVIEW; }
        else { rec = 'Low Data'; }
    }
    outRows.push([c,ag,status,impr,clk,ctr,cost,conv,convVal,cpa,roas,rec]);
    if(outRows.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outRows.length, headers.length).setBackground(color); }
  }
  if (outRows.length > 0) {
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow,1,outRows.length, headers.length).setValues(outRows);
    var col=4; formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1));
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1));
    autoResizeAllColumns(sh);
  } else { sh.getRange(headerRowIndex+1,1).setValue('No PMax Asset Group data found.');}
}

function buildShoppingProductPerformance_(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'Shopping Product Performance');
  setSheetTitle(sh, 'SHOPPING - PRODUCT PERFORMANCE (GAQL)');
  if (TAB_COLOR_PRODUCTS) sh.setTabColor(TAB_COLOR_PRODUCTS);
  var headers = ['Campaign','Product Item ID','Product Title','Product Brand', 'Category L1', 'Impr','Clicks','CTR', 'Cost','Conv','CPA', 'Conv Val', 'ROAS', 'Recommendation'];
  var headerRowIndex = 3; sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM, dateToGAQL = cfg.DATE_TO;
  var gaql = 'SELECT campaign.name, campaign.advertising_channel_type, segments.product_item_id, segments.product_title, segments.product_brand, segments.product_category_level1, ' +
             'metrics.impressions, metrics.clicks, metrics.ctr, metrics.cost_micros, metrics.conversions, metrics.conversions_value ' +
             'FROM shopping_performance_view WHERE campaign.advertising_channel_type = \'SHOPPING\' ' +
             'AND segments.date BETWEEN \''+dateFromGAQL+'\' AND \''+dateToGAQL+'\' AND metrics.impressions > 0 ORDER BY metrics.cost_micros DESC LIMIT 1000';
  Logger.log('Shopping Product Performance GAQL: ' + gaql);
  var iter; try { iter = AdsApp.report(gaql).rows(); } catch (e) { logError('buildShoppingProductPerformance_GAQL', e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching data.'); return; }
  
  var outRows = [];
  while(iter.hasNext()){
    var r = iter.next();
    var camp = r['campaign.name'];
    var itemId = r['segments.product_item_id'] || '(Not Set)';
    var title = r['segments.product_title'] || '(Not Set)';
    var brand = r['segments.product_brand'] || '(Not Set)';
    var catL1 = r['segments.product_category_level1'] || '(Not Set)';
    var impr = Number(r['metrics.impressions'] || 0), clk = Number(r['metrics.clicks'] || 0), ctr = Number(r['metrics.ctr'] || 0);
    var cost = microsToCurrency(Number(r['metrics.cost_micros'] || 0)), conv = Number(r['metrics.conversions'] || 0), convVal = Number(r['metrics.conversions_value'] || 0); // Corrected convVal
    var cpa = conv>0? cost/conv : 0, roas = cost>0? convVal/cost : 0;
    var rec = '', color = COLOR_NEUTRAL;

    if (conv > 0 && roas >= cfg.TARGET_ROAS) { rec = 'Good Performance'; color = COLOR_GOOD; }
    else if (conv > 0 && roas < cfg.TARGET_ROAS) { rec = 'ROAS < Target ('+(roas*100).toFixed(0)+'%)'; color = COLOR_REVIEW; }
    else if (cost > 0 && clk > cfg.CLICK_THRESHOLD/2) { rec = 'No Conv - Review'; color = COLOR_BAD; }
    else if (cost > 0) { rec = 'Monitor'; color = COLOR_REVIEW; }
    else { rec = 'Low Data'; }

    outRows.push([camp,itemId,title,brand,catL1,impr,clk,ctr,cost,conv,cpa,convVal,roas,rec]);
    if(outRows.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outRows.length, headers.length).setBackground(color); }
  }
   if (outRows.length > 0) {
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow,1,outRows.length, headers.length).setValues(outRows);
    var col=6; formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); // Impr, Clicks
    formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1));
    autoResizeAllColumns(sh);
  } else { sh.getRange(headerRowIndex+1,1).setValue('No Shopping product performance data found.');}
}

function pmaxProducts_ProductTitle(ss, cfg, summaryData) {
  var sh = resetSheet(ss, 'PMax Product Titles');
  setSheetTitle(sh, 'PERFORMANCE MAX - PRODUCT TITLE ANALYSIS');
  if (TAB_COLOR_PRODUCTS) sh.setTabColor(TAB_COLOR_PRODUCTS);
  var headers = ['Campaign','Product Title','Product ID', 'Impr','Clicks','CTR', 'Cost','Conv','Conv Val','CPA','ROAS', 'Recommendation'];
  var headerRowIndex = 3; sh.getRange(headerRowIndex, 1, 1, headers.length).setValues([headers]);
  setHeaderRowFormatting(sh, headerRowIndex); freezeHeader(sh, 1);

  var dateFromGAQL = cfg.DATE_FROM, dateToGAQL = cfg.DATE_TO;
  var gaql = 'SELECT campaign.name, campaign.advertising_channel_type, segments.product_title, segments.product_item_id, metrics.impressions, metrics.clicks, metrics.ctr, metrics.cost_micros, metrics.conversions, metrics.conversions_value FROM shopping_performance_view WHERE campaign.advertising_channel_type = \'PERFORMANCE_MAX\' AND segments.date BETWEEN \''+dateFromGAQL+'\' AND \''+dateToGAQL+'\' AND metrics.impressions > 0 ORDER BY metrics.cost_micros DESC LIMIT 1000';
  Logger.log('PMax Product Titles GAQL: ' + gaql);
  var iter; try { iter = AdsApp.report(gaql).rows(); } catch (e) { logError('pmaxProducts_ProductTitle_GAQL', e, gaql); sh.getRange(headerRowIndex+1,1).setValue('Error fetching data.'); return; }
  
  var outRows = [];
  while(iter.hasNext()){
    var r = iter.next();
    var c = r['campaign.name'], title = r['segments.product_title']||'(Not Set)', itemId = r['segments.product_item_id']||'(Not Set)';
    var impr = Number(r['metrics.impressions'] || 0), clk = Number(r['metrics.clicks'] || 0), ctr = Number(r['metrics.ctr'] || 0);
    var cost = microsToCurrency(Number(r['metrics.cost_micros'] || 0));
    var conv = Number(r['metrics.conversions'] || 0);
    var convVal = Number(r['metrics.conversions_value'] || 0); // Corrected convVal
    var cpa = conv>0? cost/conv : 0, roas = cost>0? convVal/cost : 0;
    var rec = '', color = COLOR_NEUTRAL;

    if (title === '(Not Set)' && itemId === '(Not Set)') { rec = 'Data Missing Title/ID'; color = COLOR_HIGHLIGHT; }
    else if (conv > 0 && roas >= cfg.TARGET_ROAS) { rec = 'Good Performance'; color = COLOR_GOOD; }
    else if (conv > 0 && roas < cfg.TARGET_ROAS) { rec = 'ROAS < Target ('+(roas*100).toFixed(0)+'%)'; color = COLOR_REVIEW; }
    else if (cost > 0 && clk > cfg.CLICK_THRESHOLD/2) { rec = 'No Conv - Review'; color = COLOR_BAD; }
    else if (cost > 0) { rec = 'Monitor'; color = COLOR_REVIEW; }
    else { rec = 'Low Data'; }
    outRows.push([c,title,itemId,impr,clk,ctr,cost,conv,convVal,cpa,roas,rec]);
    if(outRows.length > 0 && color !== COLOR_NEUTRAL) { sh.getRange(headerRowIndex + outRows.length, headers.length).setBackground(color); }
  }
  if (outRows.length > 0) {
    var dataStartRow = headerRowIndex + 1;
    sh.getRange(dataStartRow,1,outRows.length, headers.length).setValues(outRows);
    var col=4; formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1)); // Impr, Clicks
    formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatInteger(sh.getRange(dataStartRow,col++,outRows.length,1));
    formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); formatCurrency(sh.getRange(dataStartRow,col++,outRows.length,1)); 
    formatPercent(sh.getRange(dataStartRow,col++,outRows.length,1));
    autoResizeAllColumns(sh);
  } else { sh.getRange(headerRowIndex+1,1).setValue('No PMax Product Title data found.');}
}

function buildExecutiveSummary_(ss, cfg, summaryData) {
    var sh = resetSheet(ss, 'Executive Summary');
    setSheetTitle(sh, 'EXECUTIVE SUMMARY - Google Ads Account Audit');
    if (TAB_COLOR_SUMMARIES) sh.setTabColor(TAB_COLOR_SUMMARIES);
    var dataStartRow = 3;
    sh.getRange(dataStartRow++, 1).setValue('Account: ' + ACCOUNT_NAME + ' (' + ACCOUNT_ID + ')').setFontWeight('bold');
    sh.getRange(dataStartRow++, 1).setValue('Reporting Period: ' + cfg.DATE_FROM + ' to ' + cfg.DATE_TO).setFontWeight('bold');
    sh.getRange(dataStartRow++, 1).setValue('Date Generated: ' + Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd HH:mm:ss')).setFontWeight('bold');
    dataStartRow++;

    sh.getRange(dataStartRow, 1).setValue('Key Findings (Highlights):').setFontWeight('bold').setFontSize(12).setBackground(COLOR_GOOD); dataStartRow++;
    if (summaryData.findings.length > 0) {
        summaryData.findings.slice(0,5).forEach(function(f){ sh.getRange(dataStartRow++, 1).setValue('• ' + f); });
    } else { sh.getRange(dataStartRow++, 1).setValue('No specific positive findings automatically flagged.'); }
    dataStartRow++;

    sh.getRange(dataStartRow, 1).setValue('Key Opportunities & Areas for Review:').setFontWeight('bold').setFontSize(12).setBackground(COLOR_REVIEW); dataStartRow++;
    if (summaryData.opportunities.length > 0) {
        summaryData.opportunities.slice(0,7).forEach(function(o){ sh.getRange(dataStartRow++, 1).setValue('• ' + o); });
    } else { sh.getRange(dataStartRow++, 1).setValue('No major opportunities automatically flagged.'); }
    dataStartRow++;
    
    sh.getRange(dataStartRow, 1).setValue('Potential Warnings/Critical Issues:').setFontWeight('bold').setFontSize(12).setBackground(COLOR_BAD); dataStartRow++;
     if (summaryData.warnings.length > 0) {
        summaryData.warnings.slice(0,5).forEach(function(w){ sh.getRange(dataStartRow++, 1).setValue('• ' + w); });
    } else { sh.getRange(dataStartRow++, 1).setValue('No critical warnings automatically flagged.'); }
    dataStartRow++;

    sh.getRange(dataStartRow, 1).setValue('Overall Recommendation:').setFontWeight('bold').setFontSize(12); dataStartRow++;
    sh.getRange(dataStartRow, 1, 3, 1).setValues([
        ['This automated audit provides a snapshot of account performance and structure based on predefined logic and thresholds.'],
        ['It is crucial to manually review the detailed tabs and apply strategic judgment before making any changes to the account.'],
        ['Focus on addressing "Opportunities" and "Warnings", and leverage "Findings" to scale successes.']
    ]).setWrap(true);
    
    autoResizeAllColumns(sh);
    try { if (sh.getIndex() > 4) { ss.setActiveSheet(sh); ss.moveActiveSheet(4); }} 
    catch(e) { logError('buildExecutiveSummary_move', e); }
}
// End of SECTION 18
