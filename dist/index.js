const LOG_SHEET_REQUIRED_TOKENS = ["KIN", "KAN"];
const LOG_SHEET_DAY_TOKENS = ["MON", "TUE", "WED", "THU", "FRI", "SAT"];
const INDEX_SHEET_NAME = "Index";
const INVENTORY_SHEET_NAME = "Inventory";
const INVENTORY_IMPORT_SOURCE_SPREADSHEET_ID = "1QYBk6N_RZygLDPWV8BjVpF2azXBCyvGNRuz9XvpakPE";
const INVENTORY_IMPORT_SOURCE_SHEET_NAME = "Inventory";
const INVENTORY_IMPORT_DESTINATION_COLUMN_COUNT = 7;
const INDEX_HEADERS = ["Sheet Name", "Go"];
const INDEX_SHEET_NAME_COLUMN = 1;
const INDEX_NAVIGATION_COLUMN = 2;
const INDEX_SHEET_NAME_MIN_WIDTH_PX = 180;
const INDEX_SHEET_NAME_MAX_WIDTH_PX = 900;
const INDEX_SHEET_NAME_CHAR_WIDTH_PX = 9;
const INDEX_SHEET_NAME_PADDING_PX = 36;
const INDEX_WEEK_PANEL_HEADER_ROW = 1;
const INDEX_WEEK_PANEL_START_ROW = 2;
const INDEX_WEEK_PANEL_LABEL_COLUMN = 4; // D
const INDEX_WEEK_PANEL_DATE_COLUMN = 5; // E
const INDEX_WEEK_PANEL_MONDAY_A1 = "F1";
const INDEX_DAY_LABELS_COLUMN = 6; // F
const INDEX_WEEK_DAY_LABELS = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday"
];
const INDEX_VISIBILITY_SYNC_TRIGGER_HANDLER = "syncIndexVisibilityFromTrigger";
const INVENTORY_AUTO_SYNC_TRIGGER_HANDLER = "syncInventoryFromTrigger";
const INVENTORY_AUTO_SYNC_ON_OPEN_TRIGGER_HANDLER = "syncInventoryFromOnOpenTrigger";
const INVENTORY_AUTO_SYNC_INTERVAL_MINUTES = 6;
const INVENTORY_AUTO_SYNC_TRIGGER_POLL_MINUTES = 1;
const INVENTORY_AUTO_SYNC_INTERVAL_MILLISECONDS = INVENTORY_AUTO_SYNC_INTERVAL_MINUTES * 60 * 1000;
const INVENTORY_AUTO_SYNC_LAST_RUN_PROPERTY_KEY = "inventoryAutoSyncLastRunMs";
const DAY_NUMBER_TO_DAY_TOKEN = {
    1: "MON",
    2: "TUE",
    3: "WED",
    4: "THU",
    5: "FRI",
    6: "SAT"
};
const INVENTORY_IMPORT_COLUMN_MAPPING = [
    { sourceHeader: "Title", destinationIndex: 0 },
    { sourceHeader: "Artist", destinationIndex: 1 },
    { sourceHeader: "Category", destinationIndex: 2 },
    { sourceHeader: "Number", destinationIndex: 3 },
    { sourceHeader: "LengthSeconds", destinationIndex: 4 },
    { sourceHeader: "StartDate", destinationIndex: 5 },
    { sourceHeader: "EndDate", destinationIndex: 6 }
];
const TITLE_COLUMN = 4; // D
const CATEGORY_COLUMN = 6; // F
const CART_ID_COLUMN = 7; // G
const LENGTH_COLUMN = 8; // H
const PICKER_COLUMN = 9; // I
const DEFAULT_TIME_COLUMN = 2; // B
const REQUIRED_BREAK_SECONDS = 60;
const HALF_BREAK_SECONDS = 30;
const BREAK_SECONDS_TOLERANCE = 1;
const VALID_CART_ID_FONT_COLOR = "#000000";
const INVALID_CART_ID_FONT_COLOR = "#d93025";
const FL5_CATEGORY_KEY = "FL5";
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui
        .createMenu("KIN KAN Tools")
        .addItem("Open Log Navigator", "showLogNavigatorDialog")
        .addSeparator()
        .addItem("Sync Inventory", "syncInventoryFromExternalWorkbook")
        .addSubMenu(ui
        .createMenu("Inventory Auto Sync")
        .addItem("Enable (On Open + Every 6 Min)", "enableInventoryAutoSyncTriggers")
        .addItem("Disable", "disableInventoryAutoSyncTriggers"))
        .addItem("Fill FL5 Break Gaps", "fillFl5BreakGapsFromInventory")
        .addItem("Remove FL5 Carts From Logs", "removeFl5CartsFromLogs")
        .addItem("Check Break Lengths", "checkActiveLogSheetBreakDurations")
        .addItem("Export to ASC", "exportActiveLogSheetToAsc")
        .addSeparator()
        .addItem("Open Index", "openIndexSheet")
        .addItem("Show Index Only", "showIndexOnly")
        .addItem("Refresh Index", "refreshIndexSheet")
        .addSubMenu(ui
        .createMenu("Visibility Sync")
        .addItem("Sync Visibility Now", "syncIndexVisibilityNow")
        .addItem("Enable 1-Min Visibility Sync", "enableIndexVisibilitySyncTrigger")
        .addItem("Disable Visibility Sync", "disableIndexVisibilitySyncTrigger"))
        .addSubMenu(ui
        .createMenu("Jump")
        .addItem("Go To Today Log", "jumpToTodayLogSheet")
        .addItem("Go To Next Log", "jumpToNextLogSheet")
        .addItem("Go To Previous Log", "jumpToPreviousLogSheet")
        .addItem("Go To Log By Name", "jumpToLogSheetByNamePrompt"))
        .addToUi();
}
function showLogNavigatorDialog() {
    const html = HtmlService.createHtmlOutputFromFile("LogNavigatorDialog")
        .setTitle("Log Navigator")
        .setWidth(440)
        .setHeight(560);
    SpreadsheetApp.getUi().showModelessDialog(html, "Log Navigator");
}
function getLogSheetNamesForDialog() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    return getLogSheetsForNavigation_(spreadsheet).map((sheet) => sheet.getName());
}
function openLogSheetFromDialog(sheetName) {
    const targetSheetName = String(sheetName !== null && sheetName !== void 0 ? sheetName : "").trim();
    if (!targetSheetName) {
        return { success: false, message: "Choose a log sheet first." };
    }
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = spreadsheet.getSheetByName(targetSheetName);
    if (!targetSheet || !isNavigableLogSheetName_(targetSheetName)) {
        return { success: false, message: `Sheet \"${targetSheetName}\" is not a valid log sheet.` };
    }
    activateSheet_(spreadsheet, targetSheet);
    setIndexSheetVisibilityFlag_(spreadsheet, targetSheetName, true);
    return { success: true, message: `Opened ${targetSheetName}.` };
}
function openIndexSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const indexSheet = getOrCreateIndexSheet_(spreadsheet);
    refreshIndexSheet();
    activateSheet_(spreadsheet, indexSheet);
    indexSheet.setActiveSelection("A1");
}
function showIndexOnly() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const indexSheet = getOrCreateIndexSheet_(spreadsheet);
    activateSheet_(spreadsheet, indexSheet);
    let hiddenCount = 0;
    for (const sheet of spreadsheet.getSheets()) {
        if (sheet.getSheetId() === indexSheet.getSheetId())
            continue;
        if (sheet.isSheetHidden())
            continue;
        sheet.hideSheet();
        hiddenCount++;
    }
    const lastRow = indexSheet.getLastRow();
    if (lastRow >= 2) {
        indexSheet
            .getRange(2, INDEX_NAVIGATION_COLUMN, lastRow - 1, 1)
            .setValue(false);
    }
    indexSheet.setActiveSelection("A1");
    spreadsheet.toast(`Index is now the only visible sheet. Hid ${hiddenCount} sheet(s).`, "Index", 4);
}
function syncIndexVisibilityNow() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const result = syncIndexVisibilityFlagsFromTabs_(spreadsheet);
    spreadsheet.toast(`Visibility synced. Updated ${result.updatedRowCount} row(s).`, "Index", 4);
}
function syncIndexVisibilityFromTrigger() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet)
        return;
    syncIndexVisibilityFlagsFromTabs_(spreadsheet);
}
function enableIndexVisibilitySyncTrigger() {
    deleteIndexVisibilitySyncTriggers_();
    ScriptApp.newTrigger(INDEX_VISIBILITY_SYNC_TRIGGER_HANDLER)
        .timeBased()
        .everyMinutes(1)
        .create();
    SpreadsheetApp
        .getActiveSpreadsheet()
        .toast("1-minute visibility sync enabled.", "Visibility Sync", 4);
}
function disableIndexVisibilitySyncTrigger() {
    const removedCount = deleteIndexVisibilitySyncTriggers_();
    SpreadsheetApp
        .getActiveSpreadsheet()
        .toast(`Removed ${removedCount} visibility sync trigger(s).`, "Visibility Sync", 4);
}
function deleteIndexVisibilitySyncTriggers_() {
    const triggers = ScriptApp.getProjectTriggers();
    let removedCount = 0;
    for (const trigger of triggers) {
        if (trigger.getHandlerFunction() !== INDEX_VISIBILITY_SYNC_TRIGGER_HANDLER)
            continue;
        ScriptApp.deleteTrigger(trigger);
        removedCount++;
    }
    return removedCount;
}
function syncIndexVisibilityFlagsFromTabs_(spreadsheet) {
    const indexSheet = spreadsheet.getSheetByName(INDEX_SHEET_NAME);
    if (!indexSheet)
        return { updatedRowCount: 0 };
    const lastRow = indexSheet.getLastRow();
    if (lastRow < 2)
        return { updatedRowCount: 0 };
    const rowCount = lastRow - 1;
    const nameValues = indexSheet
        .getRange(2, INDEX_SHEET_NAME_COLUMN, rowCount, 1)
        .getDisplayValues();
    const checkboxValues = indexSheet
        .getRange(2, INDEX_NAVIGATION_COLUMN, rowCount, 1)
        .getValues();
    const nextValues = [];
    let updatedRowCount = 0;
    for (let i = 0; i < rowCount; i++) {
        const sheetName = String(nameValues[i][0] || "").trim();
        let isVisible = false;
        if (sheetName) {
            const sheet = spreadsheet.getSheetByName(sheetName);
            if (sheet && isNavigableLogSheetName_(sheetName)) {
                isVisible = !sheet.isSheetHidden();
            }
        }
        const wasVisible = String(checkboxValues[i][0]).toUpperCase() === "TRUE";
        if (wasVisible !== isVisible) {
            updatedRowCount++;
        }
        nextValues.push([isVisible]);
    }
    if (updatedRowCount > 0) {
        indexSheet
            .getRange(2, INDEX_NAVIGATION_COLUMN, rowCount, 1)
            .setValues(nextValues);
    }
    return { updatedRowCount };
}
function refreshIndexSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const indexSheet = getOrCreateIndexSheet_(spreadsheet);
    const logSheets = getLogSheetsForNavigation_(spreadsheet);
    const longestSheetNameLength = getLongestSheetNameLength_(logSheets);
    const listArea = indexSheet.getRange(1, 1, indexSheet.getMaxRows(), INDEX_HEADERS.length);
    listArea.clearContent();
    listArea.clearFormat();
    listArea.clearDataValidations();
    indexSheet.getRange(1, 1, 1, INDEX_HEADERS.length).setValues([INDEX_HEADERS]);
    if (logSheets.length > 0) {
        const indexRows = logSheets.map((sheet) => {
            return [sheet.getName(), !sheet.isSheetHidden()];
        });
        indexSheet
            .getRange(2, 1, indexRows.length, INDEX_HEADERS.length)
            .setValues(indexRows);
        indexSheet
            .getRange(2, INDEX_SHEET_NAME_COLUMN, indexRows.length, 1)
            .setFontColor("#000000");
        formatIndexGoColumn_(indexSheet, indexRows.length);
    }
    const headerRange = indexSheet.getRange(1, 1, 1, INDEX_HEADERS.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#dbe9ff");
    indexSheet.setFrozenRows(1);
    autoFitIndexSheetNameColumn_(indexSheet, longestSheetNameLength);
    if (INDEX_HEADERS.length > 1) {
        indexSheet.autoResizeColumns(2, INDEX_HEADERS.length - 1);
    }
    ensureIndexWeekDatePanel_(indexSheet);
    indexSheet.setActiveSelection("A1");
    spreadsheet.toast(`Index refreshed with ${logSheets.length} log sheet(s).`, "Index", 4);
}
function ensureIndexWeekDatePanel_(indexSheet) {
    // Get the Monday date reference (default to most recent Monday if not set)
    const mondayCell = indexSheet.getRange(INDEX_WEEK_PANEL_MONDAY_A1);
    let mondayDate = mondayCell.getValue();
    if (!mondayDate || !(mondayDate instanceof Date)) {
        mondayDate = getMostRecentMonday_(new Date());
    }
    // Set Monday cell in F1
    mondayCell
        .setValue(mondayDate)
        .setNote("Enter the Monday date for this workbook week.")
        .setBackground("#ffffff")
        .setNumberFormat("ddd mmm d, yyyy");
    // Clear column E completely (including notes from E2)
    const maxRows = indexSheet.getMaxRows();
    const columnERange = indexSheet.getRange(INDEX_WEEK_PANEL_START_ROW, INDEX_WEEK_PANEL_DATE_COLUMN, maxRows - 1, 1);
    columnERange.clearContent();
    columnERange.clearFormat();
    columnERange.clearDataValidations();
    columnERange.clearNote();
    // Read all sheet names from column A (starting from row 2)
    // Get a large range to capture all potential sheets
    const sheetNameRange = indexSheet.getRange(INDEX_WEEK_PANEL_START_ROW, INDEX_SHEET_NAME_COLUMN, maxRows - 1, 1);
    const sheetNameValues = sheetNameRange.getValues();
    // Clear column D first to remove old values
    indexSheet
        .getRange(INDEX_WEEK_PANEL_START_ROW, INDEX_WEEK_PANEL_LABEL_COLUMN, maxRows - 1, 1)
        .clearContent();
    // Extract day abbreviations and calculate dates
    const dateValues = [];
    for (let i = 0; i < sheetNameValues.length; i++) {
        const sheetName = String(sheetNameValues[i][0]).trim();
        // Only process non-empty cells with valid sheet name format
        if (sheetName) {
            const dayAbbr = extractDayAbbrFromSheetName_(sheetName);
            // Calculate the date for this day based on offset from Monday
            if (dayAbbr !== "---" && mondayDate instanceof Date) {
                const dayOffset = getDayOffsetFromMonday_(dayAbbr);
                const cellDate = new Date(mondayDate);
                cellDate.setDate(cellDate.getDate() + dayOffset);
                dateValues.push([cellDate]);
            }
            else {
                dateValues.push([""]);
            }
        }
        else {
            dateValues.push([""]);
        }
    }
    // Write dates to column D
    if (dateValues.length > 0) {
        indexSheet
            .getRange(INDEX_WEEK_PANEL_START_ROW, INDEX_WEEK_PANEL_LABEL_COLUMN, dateValues.length, 1)
            .setValues(dateValues)
            .setNumberFormat("ddd mmm d, yyyy")
            .setBackground("#fff8e1");
    }
    indexSheet.setColumnWidth(INDEX_WEEK_PANEL_LABEL_COLUMN, 160);
}
function extractDayAbbrFromSheetName_(sheetName) {
    // Pattern: "KAN_1_MON" or "KIN_2_TUE" etc.
    // Extract the three-letter day abbreviation at the end
    const dayPattern = /_(MON|TUE|WED|THU|FRI|SAT|SUN)$/i;
    const match = sheetName.match(dayPattern);
    if (match) {
        return match[1].toUpperCase();
    }
    // Fallback if format doesn't match
    return "---";
}
function getDayOffsetFromMonday_(dayAbbr) {
    var _a;
    // Returns the number of days from Monday (Monday = 0, Tuesday = 1, etc.)
    const dayMap = {
        MON: 0,
        TUE: 1,
        WED: 2,
        THU: 3,
        FRI: 4,
        SAT: 5,
        SUN: 6
    };
    return (_a = dayMap[dayAbbr.toUpperCase()]) !== null && _a !== void 0 ? _a : 0;
}
function getMostRecentMonday_(today) {
    const monday = new Date(today);
    monday.setHours(0, 0, 0, 0);
    const day = monday.getDay();
    const dayOffset = (day + 6) % 7;
    monday.setDate(monday.getDate() - dayOffset);
    return monday;
}
function getLongestSheetNameLength_(logSheets) {
    let longestLength = INDEX_HEADERS[0].length;
    for (const sheet of logSheets) {
        const sheetNameLength = sheet.getName().length;
        if (sheetNameLength > longestLength) {
            longestLength = sheetNameLength;
        }
    }
    return longestLength;
}
function autoFitIndexSheetNameColumn_(indexSheet, longestSheetNameLength) {
    SpreadsheetApp.flush();
    indexSheet.autoResizeColumn(INDEX_SHEET_NAME_COLUMN);
    const estimatedWidth = Math.round(longestSheetNameLength * INDEX_SHEET_NAME_CHAR_WIDTH_PX + INDEX_SHEET_NAME_PADDING_PX);
    const constrainedWidth = Math.min(INDEX_SHEET_NAME_MAX_WIDTH_PX, Math.max(INDEX_SHEET_NAME_MIN_WIDTH_PX, estimatedWidth));
    if (indexSheet.getColumnWidth(INDEX_SHEET_NAME_COLUMN) >= constrainedWidth)
        return;
    indexSheet.setColumnWidth(INDEX_SHEET_NAME_COLUMN, constrainedWidth);
}
function formatIndexGoColumn_(indexSheet, rowCount) {
    if (rowCount < 1)
        return;
    const goRange = indexSheet.getRange(2, INDEX_NAVIGATION_COLUMN, rowCount, 1);
    goRange.insertCheckboxes();
    goRange.setHorizontalAlignment("center");
    goRange.setBackground("#e8f0fe");
    goRange.setFontWeight("bold");
}
function jumpToTodayLogSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const todayDayToken = getTodayDayToken_();
    if (!todayDayToken) {
        spreadsheet.toast("No log sheet mapping exists for today.", "Jump", 5);
        return;
    }
    const candidateSheets = getLogSheetsForNavigation_(spreadsheet).filter((sheet) => sheet.getName().toUpperCase().includes(todayDayToken));
    if (candidateSheets.length === 0) {
        spreadsheet.toast(`No ${todayDayToken} log sheet was found.`, "Jump", 5);
        return;
    }
    const targetSheet = candidateSheets[candidateSheets.length - 1];
    activateSheet_(spreadsheet, targetSheet);
    if (candidateSheets.length > 1) {
        spreadsheet.toast(`Multiple ${todayDayToken} sheets found. Opened ${targetSheet.getName()}.`, "Jump", 5);
    }
}
function jumpToNextLogSheet() {
    jumpToRelativeLogSheet_(1);
}
function jumpToPreviousLogSheet() {
    jumpToRelativeLogSheet_(-1);
}
function jumpToLogSheetByNamePrompt() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const logSheets = getLogSheetsForNavigation_(spreadsheet);
    if (logSheets.length === 0) {
        spreadsheet.toast("No log sheets were found.", "Jump", 5);
        return;
    }
    const previewNames = logSheets
        .slice(0, 8)
        .map((sheet) => sheet.getName())
        .join(", ");
    const hiddenCount = logSheets.length - Math.min(logSheets.length, 8);
    const promptMessage = `Enter a log sheet name. Examples: ${previewNames}` +
        (hiddenCount > 0 ? `, +${hiddenCount} more` : "");
    const response = ui.prompt("Go To Log Sheet", promptMessage, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK)
        return;
    const targetSheetName = response.getResponseText().trim();
    if (!targetSheetName)
        return;
    const targetSheet = spreadsheet.getSheetByName(targetSheetName);
    if (!targetSheet || !isNavigableLogSheetName_(targetSheetName)) {
        spreadsheet.toast(`Sheet \"${targetSheetName}\" is not a valid log sheet.`, "Jump", 5);
        return;
    }
    activateSheet_(spreadsheet, targetSheet);
}
function getOrCreateIndexSheet_(spreadsheet) {
    const existingIndexSheet = spreadsheet.getSheetByName(INDEX_SHEET_NAME);
    if (existingIndexSheet)
        return existingIndexSheet;
    const indexSheet = spreadsheet.insertSheet(INDEX_SHEET_NAME, 0);
    indexSheet.getRange(1, 1).setValue("Log sheet index");
    return indexSheet;
}
function activateSheet_(spreadsheet, targetSheet) {
    if (targetSheet.isSheetHidden()) {
        targetSheet.showSheet();
    }
    spreadsheet.setActiveSheet(targetSheet);
}
function setIndexSheetVisibilityFlag_(spreadsheet, sheetName, isVisible) {
    const indexSheet = spreadsheet.getSheetByName(INDEX_SHEET_NAME);
    if (!indexSheet)
        return;
    const lastRow = indexSheet.getLastRow();
    if (lastRow < 2)
        return;
    const names = indexSheet
        .getRange(2, INDEX_SHEET_NAME_COLUMN, lastRow - 1, 1)
        .getDisplayValues();
    for (let rowOffset = 0; rowOffset < names.length; rowOffset++) {
        if (String(names[rowOffset][0]).trim() !== sheetName)
            continue;
        indexSheet
            .getRange(rowOffset + 2, INDEX_NAVIGATION_COLUMN)
            .setValue(isVisible);
        return;
    }
}
function getLogSheetsForNavigation_(spreadsheet) {
    return spreadsheet
        .getSheets()
        .filter((sheet) => isNavigableLogSheetName_(sheet.getName()))
        .sort((left, right) => left.getName().localeCompare(right.getName()));
}
function isNavigableLogSheetName_(sheetName) {
    const normalizedName = sheetName.trim().toUpperCase();
    if (normalizedName === INVENTORY_SHEET_NAME.toUpperCase())
        return false;
    if (normalizedName === INDEX_SHEET_NAME.toUpperCase())
        return false;
    return LOG_SHEET_REQUIRED_TOKENS.some((token) => hasStandaloneSheetToken_(normalizedName, token));
}
function hasStandaloneSheetToken_(normalizedSheetName, token) {
    if (normalizedSheetName.startsWith(token))
        return true;
    const tokenRegex = new RegExp(`(^|[^A-Z])${token}([^A-Z]|$)`);
    return tokenRegex.test(normalizedSheetName);
}
function getTodayDayToken_() {
    var _a;
    const dayNumber = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "u"));
    return (_a = DAY_NUMBER_TO_DAY_TOKEN[dayNumber]) !== null && _a !== void 0 ? _a : "";
}
function jumpToRelativeLogSheet_(direction) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheets = getLogSheetsForNavigation_(spreadsheet);
    if (logSheets.length === 0) {
        spreadsheet.toast("No log sheets were found.", "Jump", 5);
        return;
    }
    const activeSheetName = spreadsheet.getActiveSheet().getName();
    const currentIndex = logSheets.findIndex((sheet) => sheet.getName() === activeSheetName);
    if (currentIndex < 0) {
        const fallbackIndex = direction >= 0 ? 0 : logSheets.length - 1;
        activateSheet_(spreadsheet, logSheets[fallbackIndex]);
        return;
    }
    const targetIndex = (currentIndex + direction + logSheets.length) % logSheets.length;
    activateSheet_(spreadsheet, logSheets[targetIndex]);
}
function enableInventoryAutoSyncTriggers() {
    deleteInventoryAutoSyncTriggers_();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    ScriptApp.newTrigger(INVENTORY_AUTO_SYNC_ON_OPEN_TRIGGER_HANDLER)
        .forSpreadsheet(spreadsheet)
        .onOpen()
        .create();
    ScriptApp.newTrigger(INVENTORY_AUTO_SYNC_TRIGGER_HANDLER)
        .timeBased()
        .everyMinutes(INVENTORY_AUTO_SYNC_TRIGGER_POLL_MINUTES)
        .create();
    spreadsheet.toast(`Inventory auto sync enabled (on open + every ${INVENTORY_AUTO_SYNC_INTERVAL_MINUTES} minutes).`, "Inventory Auto Sync", 5);
}
function disableInventoryAutoSyncTriggers() {
    const removedCount = deleteInventoryAutoSyncTriggers_();
    SpreadsheetApp
        .getActiveSpreadsheet()
        .toast(`Removed ${removedCount} inventory auto-sync trigger(s).`, "Inventory Auto Sync", 5);
}
function syncInventoryFromTrigger() {
    syncInventoryFromTrigger_(false);
}
function syncInventoryFromOnOpenTrigger() {
    syncInventoryFromTrigger_(true);
}
function syncInventoryFromTrigger_(forceSync) {
    const now = Date.now();
    const scriptProperties = PropertiesService.getScriptProperties();
    const lastRunValue = scriptProperties.getProperty(INVENTORY_AUTO_SYNC_LAST_RUN_PROPERTY_KEY);
    const lastRunMs = lastRunValue ? Number(lastRunValue) : 0;
    if (!forceSync && Number.isFinite(lastRunMs) && (now - lastRunMs) < INVENTORY_AUTO_SYNC_INTERVAL_MILLISECONDS) {
        return;
    }
    try {
        syncInventoryFromExternalWorkbook();
        scriptProperties.setProperty(INVENTORY_AUTO_SYNC_LAST_RUN_PROPERTY_KEY, String(now));
    }
    catch (error) {
        Logger.log(`Inventory auto sync failed: ${error}`);
    }
}
function deleteInventoryAutoSyncTriggers_() {
    const triggers = ScriptApp.getProjectTriggers();
    let removedCount = 0;
    for (const trigger of triggers) {
        const handlerFunction = trigger.getHandlerFunction();
        if (handlerFunction !== INVENTORY_AUTO_SYNC_TRIGGER_HANDLER
            && handlerFunction !== INVENTORY_AUTO_SYNC_ON_OPEN_TRIGGER_HANDLER) {
            continue;
        }
        ScriptApp.deleteTrigger(trigger);
        removedCount++;
    }
    return removedCount;
}
function syncInventoryFromExternalWorkbook() {
    const destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const destinationSheet = destinationSpreadsheet.getSheetByName(INVENTORY_SHEET_NAME);
    if (!destinationSheet) {
        throw new Error(`Destination sheet "${INVENTORY_SHEET_NAME}" was not found in the active spreadsheet.`);
    }
    const sourceSpreadsheet = SpreadsheetApp.openById(INVENTORY_IMPORT_SOURCE_SPREADSHEET_ID);
    const sourceSheet = sourceSpreadsheet.getSheetByName(INVENTORY_IMPORT_SOURCE_SHEET_NAME);
    if (!sourceSheet) {
        throw new Error(`Source sheet "${INVENTORY_IMPORT_SOURCE_SHEET_NAME}" was not found in the source spreadsheet.`);
    }
    const sourceLastColumn = sourceSheet.getLastColumn();
    if (sourceLastColumn < 1) {
        writeInventoryRows_(destinationSheet, []);
        return;
    }
    const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceLastColumn).getValues()[0];
    const sourceHeaderIndexByKey = buildHeaderIndexByKey_(sourceHeaders);
    validateInventoryImportHeaders_(sourceHeaderIndexByKey);
    const sourceLastRow = sourceSheet.getLastRow();
    const sourceRows = sourceLastRow > 1
        ? sourceSheet.getRange(2, 1, sourceLastRow - 1, sourceLastColumn).getValues()
        : [];
    const mappedRows = sourceRows
        .map((sourceRow) => mapSourceInventoryRow_(sourceRow, sourceHeaderIndexByKey))
        .filter((row) => !isInventoryRowBlank_(row));
    writeInventoryRows_(destinationSheet, mappedRows);
    Logger.log(`Inventory sync complete. Imported ${mappedRows.length} row(s).`);
}
function checkActiveLogSheetBreakDurations() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = spreadsheet.getActiveSheet();
    if (!isLogSheetName_(activeSheet.getName())) {
        spreadsheet.toast("Break length check works on KIN/KAN log sheets only.", "Duration Check", 5);
        return;
    }
    validateLogSheetBreakDurations_(activeSheet);
}
function fillFl5BreakGapsFromInventory() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const inventorySheet = spreadsheet.getSheetByName(INVENTORY_SHEET_NAME);
    if (!inventorySheet) {
        spreadsheet.toast("Inventory sheet not found.", "Fill FL5 Breaks", 5);
        return;
    }
    const fl5Inventory = getInventoryMatchesByCategory_(inventorySheet, FL5_CATEGORY_KEY);
    if (fl5Inventory.length === 0) {
        spreadsheet.toast("No FL5 inventory rows were found.", "Fill FL5 Breaks", 5);
        return;
    }
    const logSheets = getLogSheetsForNavigation_(spreadsheet);
    if (logSheets.length === 0) {
        spreadsheet.toast("No log sheets were found.", "Fill FL5 Breaks", 5);
        return;
    }
    let totalFilledRows = 0;
    let processedSheets = 0;
    let skippedForMissingDate = 0;
    let skippedForNoValidInventory = 0;
    for (const logSheet of logSheets) {
        const sheetDate = getIndexDateForSheetName_(spreadsheet, logSheet.getName())
            || getDateForSheetFromMondayCell_(spreadsheet, logSheet.getName());
        if (!(sheetDate instanceof Date) || Number.isNaN(sheetDate.getTime())) {
            skippedForMissingDate++;
            continue;
        }
        const monday = getMostRecentMonday_(sheetDate);
        const sunday = new Date(monday);
        sunday.setDate(sunday.getDate() + 6);
        const validFl5ForSheet = fl5Inventory.filter((match) => isInventoryMatchValidForSheetWeek_(match, sheetDate, monday, sunday));
        if (validFl5ForSheet.length === 0) {
            skippedForNoValidInventory++;
            continue;
        }
        totalFilledRows += fillFl5BreakGapsOnSheet_(logSheet, validFl5ForSheet);
        processedSheets++;
    }
    const message = `Filled ${totalFilledRows} break row(s) across ${processedSheets} sheet(s). ` +
        `Skipped ${skippedForMissingDate} sheet(s) with no index date and ` +
        `${skippedForNoValidInventory} with no valid FL5 carts for that week.`;
    spreadsheet.toast(message, "Fill FL5 Breaks", 8);
}
function removeFl5CartsFromLogs() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheets = getLogSheetsForNavigation_(spreadsheet);
    if (logSheets.length === 0) {
        spreadsheet.toast("No log sheets were found.", "Remove FL5 Carts", 5);
        return;
    }
    let totalRemovedRows = 0;
    let touchedSheets = 0;
    for (const logSheet of logSheets) {
        const removedRows = removeFl5RowsOnLogSheet_(logSheet);
        if (removedRows > 0) {
            touchedSheets++;
            totalRemovedRows += removedRows;
        }
    }
    spreadsheet.toast(`Removed ${totalRemovedRows} FL5 row(s) across ${touchedSheets} sheet(s).`, "Remove FL5 Carts", 8);
}
function removeFl5RowsOnLogSheet_(logSheet) {
    var _a;
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2)
        return 0;
    const rowCount = lastRow - 1;
    const categoryValues = logSheet.getRange(2, CATEGORY_COLUMN, rowCount, 1).getValues();
    const rowsToClear = [];
    for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
        const category = String((_a = categoryValues[rowOffset][0]) !== null && _a !== void 0 ? _a : "").trim().toUpperCase();
        if (category !== FL5_CATEGORY_KEY)
            continue;
        rowsToClear.push(rowOffset + 2);
    }
    if (rowsToClear.length === 0)
        return 0;
    for (const rowNumber of rowsToClear) {
        clearEntryRow_(logSheet, rowNumber);
        clearPicker_(logSheet, rowNumber);
    }
    validateLogSheetBreakDurations_(logSheet);
    return rowsToClear.length;
}
function onEdit(e) {
    if (!e)
        return;
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    if (sheetName === INDEX_SHEET_NAME) {
        handleIndexGoEdit_(sheet, e.range, e.value);
        return;
    }
    if (!isLogSheetName_(sheetName))
        return;
    const firstEditedRow = e.range.getRow();
    const lastEditedRow = firstEditedRow + e.range.getNumRows() - 1;
    if (lastEditedRow < 2)
        return;
    try {
        const row = e.range.getRow();
        const col = e.range.getColumn();
        if (row < 2)
            return;
        const singleCellEdit = isSingleCellEdit_(e.range);
        const cartIdRangeEdited = rangeIncludesColumn_(e.range, CART_ID_COLUMN);
        if (cartIdRangeEdited) {
            processCartIdRangeEdit_(sheet, e.range);
            if (!singleCellEdit || col === CART_ID_COLUMN)
                return;
        }
        if (!singleCellEdit)
            return;
        if (col === PICKER_COLUMN) {
            applyPickerSelection_(sheet, row, String(e.value || ""));
            return;
        }
        // Only respond to Title (D) here. Cart ID (G) is handled above.
        if (col !== TITLE_COLUMN)
            return;
        clearPicker_(sheet, row);
        const searchValue = String(e.value || "").trim().toLowerCase();
        if (!searchValue)
            return;
        const inventorySheet = SpreadsheetApp.getActive().getSheetByName(INVENTORY_SHEET_NAME);
        if (!inventorySheet)
            return;
        const activeInventory = getInventoryMatchesForLogSheetDate_(inventorySheet, sheet);
        const matches = findMatchesInInventory_(activeInventory, searchValue);
        if (matches.length === 0)
            return;
        if (matches.length === 1) {
            applyMatchToEntryRow_(sheet, row, matches[0]);
            return;
        }
        setPickerForMatches_(sheet, row, matches);
    }
    finally {
        validateLogSheetBreakDurations_(sheet, e.range);
    }
}
function handleIndexGoEdit_(indexSheet, editedRange, editedValue) {
    // Check if F1 (Monday date cell) was edited
    if (editedRange.getRow() === 1 && editedRange.getColumn() === 6) {
        updateIndexDatesFromMondayCell_(indexSheet);
        return;
    }
    // Only respond to edits that touch the navigation column
    if (!rangeIncludesColumn_(editedRange, INDEX_NAVIGATION_COLUMN))
        return;
    const lastRow = indexSheet.getLastRow();
    if (lastRow < 2)
        return;
    // Sheets only fires onEdit for the anchor cell when spacebar-toggling a
    // multi-cell checkbox selection. Work around this by always reading ALL
    // checkbox values from the sheet (which Sheets has already updated) and
    // syncing ALL sheet visibility to match. This handles single clicks,
    // rapid clicks, and bulk spacebar toggles correctly.
    const spreadsheet = indexSheet.getParent();
    const rowCount = lastRow - 1;
    const nameValues = indexSheet
        .getRange(2, INDEX_SHEET_NAME_COLUMN, rowCount, 1)
        .getDisplayValues();
    const checkboxValues = indexSheet
        .getRange(2, INDEX_NAVIGATION_COLUMN, rowCount, 1)
        .getValues();
    let needsIndexFallback = false;
    const activeSheet = spreadsheet.getActiveSheet();
    for (let i = 0; i < rowCount; i++) {
        const targetSheetName = String(nameValues[i][0] || "").trim();
        if (!targetSheetName)
            continue;
        const targetSheet = spreadsheet.getSheetByName(targetSheetName);
        if (!targetSheet || !isNavigableLogSheetName_(targetSheetName))
            continue;
        const isChecked = String(checkboxValues[i][0]).toUpperCase() === "TRUE";
        if (isChecked) {
            if (targetSheet.isSheetHidden()) {
                targetSheet.showSheet();
            }
        }
        else {
            if (!targetSheet.isSheetHidden()) {
                if (activeSheet.getSheetId() === targetSheet.getSheetId()) {
                    needsIndexFallback = true;
                }
                targetSheet.hideSheet();
            }
        }
    }
    if (needsIndexFallback) {
        const safeSheet = getOrCreateIndexSheet_(spreadsheet);
        activateSheet_(spreadsheet, safeSheet);
        safeSheet.setActiveSelection("A1");
    }
}
function updateIndexDatesFromMondayCell_(indexSheet) {
    // Get the Monday date from F1
    const mondayCell = indexSheet.getRange(INDEX_WEEK_PANEL_MONDAY_A1);
    const enteredValue = mondayCell.getValue();
    // Only update if a valid date was entered
    if (!enteredValue || !(enteredValue instanceof Date)) {
        return;
    }
    // Normalize to the Monday of whatever week the entered date falls in
    const mondayDate = getMostRecentMonday_(enteredValue);
    // Write back the normalized Monday date if it differs from what was entered
    if (mondayDate.getTime() !== enteredValue.getTime()) {
        mondayCell.setValue(mondayDate);
    }
    // Read all sheet names from column A
    const maxRows = indexSheet.getMaxRows();
    const sheetNameRange = indexSheet.getRange(INDEX_WEEK_PANEL_START_ROW, INDEX_SHEET_NAME_COLUMN, maxRows - 1, 1);
    const sheetNameValues = sheetNameRange.getValues();
    // Extract day abbreviations and calculate dates
    const dateValues = [];
    for (let i = 0; i < sheetNameValues.length; i++) {
        const sheetName = String(sheetNameValues[i][0]).trim();
        // Only process non-empty cells with valid sheet name format
        if (sheetName) {
            const dayAbbr = extractDayAbbrFromSheetName_(sheetName);
            // Calculate the date for this day based on offset from Monday
            if (dayAbbr !== "---") {
                const dayOffset = getDayOffsetFromMonday_(dayAbbr);
                const cellDate = new Date(mondayDate);
                cellDate.setDate(cellDate.getDate() + dayOffset);
                dateValues.push([cellDate]);
            }
            else {
                dateValues.push([""]);
            }
        }
        else {
            dateValues.push([""]);
        }
    }
    // Write dates to column D
    if (dateValues.length > 0) {
        indexSheet
            .getRange(INDEX_WEEK_PANEL_START_ROW, INDEX_WEEK_PANEL_LABEL_COLUMN, dateValues.length, 1)
            .setValues(dateValues)
            .setNumberFormat("ddd mmm d, yyyy")
            .setBackground("#fff8e1");
    }
}
function findInventoryMatches_(inventorySheet, searchValue) {
    const activeInventory = getActiveInventoryMatches_(inventorySheet);
    return findMatchesInInventory_(activeInventory, searchValue);
}
function getActiveInventoryMatches_(inventorySheet) {
    const today = new Date();
    return getActiveInventoryMatchesForDate_(inventorySheet, today);
}
function getInventoryMatchesForLogSheetDate_(inventorySheet, logSheet) {
    const spreadsheet = logSheet.getParent();
    const sheetDate = getIndexDateForSheetName_(spreadsheet, logSheet.getName())
        || getDateForSheetFromMondayCell_(spreadsheet, logSheet.getName())
        || new Date();
    return getActiveInventoryMatchesForDate_(inventorySheet, sheetDate);
}
function getActiveInventoryMatchesForDate_(inventorySheet, targetDate) {
    const lastRow = inventorySheet.getLastRow();
    if (lastRow < 2)
        return [];
    const data = inventorySheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const normalizedDate = new Date(targetDate);
    normalizedDate.setHours(0, 0, 0, 0);
    const activeInventory = [];
    for (const item of data) {
        const match = {
            title: String(item[0]),
            isci: item[1],
            category: item[2],
            cartId: normalizeCartId_(item[3]),
            length: item[4],
            startDate: item[5],
            endDate: item[6]
        };
        if (!isActiveToday_(normalizedDate, match.startDate, match.endDate))
            continue;
        activeInventory.push(match);
    }
    return activeInventory;
}
function getInventoryMatchesByCategory_(inventorySheet, categoryKey) {
    var _a, _b;
    const normalizedCategory = String(categoryKey !== null && categoryKey !== void 0 ? categoryKey : "").trim().toUpperCase();
    if (!normalizedCategory)
        return [];
    const lastRow = inventorySheet.getLastRow();
    if (lastRow < 2)
        return [];
    const data = inventorySheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const matches = [];
    for (const item of data) {
        const match = {
            title: String((_a = item[0]) !== null && _a !== void 0 ? _a : "").trim(),
            isci: item[1],
            category: item[2],
            cartId: normalizeCartId_(item[3]),
            length: item[4],
            startDate: item[5],
            endDate: item[6]
        };
        const matchCategory = String((_b = match.category) !== null && _b !== void 0 ? _b : "").trim().toUpperCase();
        if (matchCategory !== normalizedCategory)
            continue;
        if (!match.title || !match.cartId)
            continue;
        matches.push(match);
    }
    return matches;
}
function isInventoryMatchValidForSheetWeek_(match, sheetDate, monday, sunday) {
    const normalizedSheetDate = new Date(sheetDate);
    normalizedSheetDate.setHours(0, 0, 0, 0);
    const normalizedMonday = new Date(monday);
    normalizedMonday.setHours(0, 0, 0, 0);
    const normalizedSunday = new Date(sunday);
    normalizedSunday.setHours(0, 0, 0, 0);
    const start = toNormalizedDateOrNull_(match.startDate);
    const end = toNormalizedDateOrNull_(match.endDate);
    if (start && normalizedSheetDate < start)
        return false;
    if (end && normalizedSheetDate > end)
        return false;
    if (start && start > normalizedSunday)
        return false;
    if (end && end < normalizedMonday)
        return false;
    return true;
}
function toNormalizedDateOrNull_(value) {
    if (value === null ||
        value === "" ||
        !(value instanceof Date || typeof value === "string" || typeof value === "number")) {
        return null;
    }
    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime()))
        return null;
    parsed.setHours(0, 0, 0, 0);
    return parsed;
}
function fillFl5BreakGapsOnSheet_(logSheet, inventoryPool) {
    var _a;
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2)
        return 0;
    const rowCount = lastRow - 1;
    const timeColumn = findTimeColumnIndex_(logSheet);
    const timeValues = logSheet.getRange(2, timeColumn, rowCount, 1).getValues();
    const allData = logSheet.getRange(2, 1, rowCount, PICKER_COLUMN).getValues();
    const groups = {};
    for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
        const row = allData[rowOffset];
        const timeValue = timeValues[rowOffset][0];
        const timeKey = normalizeTimeSlotKey_(timeValue);
        if (!timeKey)
            continue;
        const titleValue = row[TITLE_COLUMN - 1];
        const categoryValue = row[CATEGORY_COLUMN - 1];
        const cartValue = row[CART_ID_COLUMN - 1];
        const lengthValue = row[LENGTH_COLUMN - 1];
        const group = (_a = groups[timeKey]) !== null && _a !== void 0 ? _a : { rows: [], totalSeconds: 0, hasAnyData: false };
        const parsedLength = parseLengthSeconds_(lengthValue);
        if (parsedLength !== null) {
            group.totalSeconds += parsedLength;
        }
        const hasAnyData = !isCellValueBlank_(titleValue) ||
            !isCellValueBlank_(categoryValue) ||
            !isCellValueBlank_(cartValue) ||
            !isCellValueBlank_(lengthValue);
        const isEmptyBreakRow = isCellValueBlank_(titleValue) &&
            isCellValueBlank_(categoryValue) &&
            isCellValueBlank_(cartValue) &&
            isCellValueBlank_(lengthValue);
        group.hasAnyData = group.hasAnyData || hasAnyData;
        group.rows.push({ rowNumber: rowOffset + 2, isEmptyBreakRow });
        groups[timeKey] = group;
    }
    const randomizedPool = shuffleInventoryMatches_(inventoryPool);
    let inventoryCursor = 0;
    let filledRows = 0;
    for (const group of Object.values(groups)) {
        let neededSeconds = 0;
        if (!group.hasAnyData) {
            neededSeconds = REQUIRED_BREAK_SECONDS;
        }
        else if (Math.abs(group.totalSeconds - HALF_BREAK_SECONDS) <= BREAK_SECONDS_TOLERANCE) {
            neededSeconds = HALF_BREAK_SECONDS;
        }
        if (neededSeconds === 0)
            continue;
        const targetRow = group.rows.find((row) => row.isEmptyBreakRow);
        if (!targetRow)
            continue;
        if (inventoryCursor >= randomizedPool.length) {
            const reshuffled = shuffleInventoryMatches_(randomizedPool);
            randomizedPool.splice(0, randomizedPool.length, ...reshuffled);
            inventoryCursor = 0;
        }
        const picked = randomizedPool[inventoryCursor++];
        logSheet.getRange(targetRow.rowNumber, TITLE_COLUMN, 1, 5).setValues([[
                picked.title,
                picked.isci,
                picked.category,
                picked.cartId,
                neededSeconds
            ]]);
        clearPicker_(logSheet, targetRow.rowNumber);
        markCartIdAsValid_(logSheet, targetRow.rowNumber);
        filledRows++;
    }
    if (filledRows > 0) {
        validateLogSheetBreakDurations_(logSheet);
    }
    return filledRows;
}
function shuffleInventoryMatches_(inventoryMatches) {
    const copy = inventoryMatches.slice();
    for (let i = copy.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        const swap = copy[i];
        copy[i] = copy[j];
        copy[j] = swap;
    }
    return copy;
}
function findMatchesInInventory_(inventory, searchValue) {
    const normalizedSearch = searchValue.trim().toLowerCase();
    if (!normalizedSearch)
        return [];
    const matches = [];
    for (const item of inventory) {
        const titleMatch = item.title.toLowerCase().includes(normalizedSearch);
        const cartMatch = item.cartId.toLowerCase().includes(normalizedSearch);
        if (!titleMatch && !cartMatch)
            continue;
        matches.push(item);
    }
    return matches;
}
function isActiveToday_(today, startDate, endDate) {
    if (startDate !== null &&
        startDate !== "" &&
        (startDate instanceof Date || typeof startDate === "string" || typeof startDate === "number")) {
        const start = new Date(startDate);
        if (!Number.isNaN(start.getTime())) {
            start.setHours(0, 0, 0, 0);
            if (today < start)
                return false;
        }
    }
    if (endDate !== null &&
        endDate !== "" &&
        (endDate instanceof Date || typeof endDate === "string" || typeof endDate === "number")) {
        const end = new Date(endDate);
        if (!Number.isNaN(end.getTime())) {
            end.setHours(0, 0, 0, 0);
            if (today > end)
                return false;
        }
    }
    return true;
}
function applyMatchToEntryRow_(entrySheet, row, match) {
    const normalizedLengthSeconds = normalizeInventoryLengthToBreakStandard_(match.length);
    const normalizedCartId = normalizeCartId_(match.cartId);
    entrySheet.getRange(row, TITLE_COLUMN, 1, 5).setValues([[
            match.title,
            match.isci,
            match.category,
            normalizedCartId,
            normalizedLengthSeconds
        ]]);
    markCartIdAsValid_(entrySheet, row);
}
function normalizeInventoryLengthToBreakStandard_(value) {
    const parsedSeconds = parseLengthSeconds_(value);
    if (parsedSeconds === null || !Number.isFinite(parsedSeconds)) {
        return 60;
    }
    return Math.abs(parsedSeconds - 30) <= Math.abs(parsedSeconds - 60) ? 30 : 60;
}
function setPickerForMatches_(entrySheet, row, matches) {
    const options = matches.map((match) => formatPickerOption_(match));
    const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(options, true)
        .setAllowInvalid(true)
        .build();
    const pickerCell = entrySheet.getRange(row, PICKER_COLUMN);
    pickerCell.clearContent();
    pickerCell.clearNote();
    pickerCell.setDataValidation(rule);
}
function applyPickerSelection_(entrySheet, row, pickerValue) {
    if (!pickerValue)
        return;
    const cartId = parsePickerCartId_(pickerValue);
    if (!cartId) {
        entrySheet.getRange(row, PICKER_COLUMN).clearContent();
        return;
    }
    const inventorySheet = SpreadsheetApp.getActive().getSheetByName(INVENTORY_SHEET_NAME);
    if (!inventorySheet)
        return;
    const activeInventory = getInventoryMatchesForLogSheetDate_(inventorySheet, entrySheet);
    const selectedMatch = activeInventory.find((match) => match.cartId.trim().toLowerCase() === cartId.toLowerCase());
    if (!selectedMatch)
        return;
    applyMatchToEntryRow_(entrySheet, row, selectedMatch);
    clearPicker_(entrySheet, row);
}
function parsePickerCartId_(value) {
    const cartIdMatch = value.match(/\(([^()]*)\)\s*$/);
    if (!cartIdMatch)
        return "";
    return cartIdMatch[1].trim();
}
function formatPickerOption_(match) {
    const maxTitleLength = 45;
    const shortTitle = match.title.length > maxTitleLength
        ? `${match.title.slice(0, maxTitleLength - 3)}...`
        : match.title;
    return `${shortTitle} (${match.cartId})`;
}
function clearPicker_(entrySheet, row) {
    const pickerCell = entrySheet.getRange(row, PICKER_COLUMN);
    pickerCell.clearContent();
    pickerCell.clearDataValidations();
    pickerCell.clearNote();
}
function clearEntryRow_(entrySheet, row) {
    const rowRange = entrySheet.getRange(row, TITLE_COLUMN, 1, 5);
    rowRange.clearContent();
    rowRange.clearDataValidations();
    rowRange.clearNote();
    markCartIdAsValid_(entrySheet, row);
}
function processCartIdRangeEdit_(entrySheet, editedRange) {
    if (!rangeIncludesColumn_(editedRange, CART_ID_COLUMN))
        return;
    const cartIdOffset = CART_ID_COLUMN - editedRange.getColumn();
    const values = editedRange.getValues();
    const inventorySheet = SpreadsheetApp.getActive().getSheetByName(INVENTORY_SHEET_NAME);
    const activeInventory = inventorySheet
        ? getInventoryMatchesForLogSheetDate_(inventorySheet, entrySheet)
        : [];
    const invalidCarts = [];
    for (let rowOffset = 0; rowOffset < values.length; rowOffset++) {
        const targetRow = editedRange.getRow() + rowOffset;
        if (targetRow < 2)
            continue;
        const cartIdValue = values[rowOffset][cartIdOffset];
        clearPicker_(entrySheet, targetRow);
        if (isCellCleared_(cartIdValue)) {
            clearEntryRow_(entrySheet, targetRow);
            continue;
        }
        if (!inventorySheet) {
            markCartIdAsValid_(entrySheet, targetRow);
            continue;
        }
        const searchValue = normalizeCartId_(cartIdValue).toLowerCase();
        const matches = findMatchesInInventory_(activeInventory, searchValue);
        if (matches.length === 0) {
            invalidCarts.push({ row: targetRow, cartId: normalizeCartId_(cartIdValue) });
            markCartIdAsInvalid_(entrySheet, targetRow);
            continue;
        }
        markCartIdAsValid_(entrySheet, targetRow);
        const exactMatch = matches.find((match) => match.cartId.toLowerCase() === searchValue);
        if (exactMatch) {
            applyMatchToEntryRow_(entrySheet, targetRow, exactMatch);
            continue;
        }
        if (matches.length === 1) {
            applyMatchToEntryRow_(entrySheet, targetRow, matches[0]);
            continue;
        }
        setPickerForMatches_(entrySheet, targetRow, matches);
    }
    if (invalidCarts.length > 0) {
        showInvalidCartToast_(entrySheet, invalidCarts);
    }
}
function showInvalidCartToast_(entrySheet, invalidCarts) {
    const spreadsheet = entrySheet.getParent();
    if (invalidCarts.length === 1) {
        const invalidCart = invalidCarts[0];
        const message = `Row ${invalidCart.row}: Cart ID "${invalidCart.cartId}" is invalid ` +
            "(missing or out of date).";
        spreadsheet.toast(message, "Invalid Cart ID", 6);
        return;
    }
    const preview = invalidCarts
        .slice(0, 3)
        .map((item) => `R${item.row}: ${item.cartId}`)
        .join(", ");
    const remainingCount = invalidCarts.length - Math.min(invalidCarts.length, 3);
    const suffix = remainingCount > 0 ? `, +${remainingCount} more` : "";
    const message = `${invalidCarts.length} invalid Cart IDs: ${preview}${suffix}.`;
    spreadsheet.toast(message, "Invalid Cart IDs", 8);
}
function validateLogSheetBreakDurations_(logSheet, editedRange) {
    var _a;
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) {
        if (!editedRange) {
            logSheet.getParent().toast("No log rows to validate.", "Duration Check", 4);
        }
        return;
    }
    const rowCount = lastRow - 1;
    const timeColumn = findTimeColumnIndex_(logSheet);
    const timeValues = logSheet.getRange(2, timeColumn, rowCount, 1).getValues();
    const titleValues = logSheet.getRange(2, TITLE_COLUMN, rowCount, 1).getValues();
    const cartValues = logSheet.getRange(2, CART_ID_COLUMN, rowCount, 1).getValues();
    const lengthValues = logSheet.getRange(2, LENGTH_COLUMN, rowCount, 1).getValues();
    const groups = {};
    for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
        const timeValue = timeValues[rowOffset][0];
        const timeKey = normalizeTimeSlotKey_(timeValue);
        if (!timeKey)
            continue;
        const rowNumber = rowOffset + 2;
        const group = (_a = groups[timeKey]) !== null && _a !== void 0 ? _a : {
            rows: [],
            totalSeconds: 0,
            hasCommercialData: false,
            displayTime: formatTimeSlotDisplay_(timeValue)
        };
        const lengthValue = lengthValues[rowOffset][0];
        const parsedLengthSeconds = parseLengthSeconds_(lengthValue);
        if (parsedLengthSeconds !== null) {
            group.totalSeconds += parsedLengthSeconds;
        }
        const hasCommercialData = !isCellValueBlank_(titleValues[rowOffset][0]) ||
            !isCellValueBlank_(cartValues[rowOffset][0]) ||
            !isCellValueBlank_(lengthValue);
        group.hasCommercialData = group.hasCommercialData || hasCommercialData;
        group.rows.push(rowNumber);
        groups[timeKey] = group;
    }
    const invalidRows = new Set();
    const invalidGroups = [];
    for (const group of Object.values(groups)) {
        if (!group.hasCommercialData)
            continue;
        if (Math.abs(group.totalSeconds - REQUIRED_BREAK_SECONDS) <= BREAK_SECONDS_TOLERANCE)
            continue;
        invalidGroups.push({
            displayTime: group.displayTime,
            totalSeconds: group.totalSeconds
        });
        for (const rowNumber of group.rows) {
            invalidRows.add(rowNumber);
        }
    }
    const timeFontColors = Array.from({ length: rowCount }, () => [VALID_CART_ID_FONT_COLOR]);
    const lengthFontColors = Array.from({ length: rowCount }, () => [VALID_CART_ID_FONT_COLOR]);
    for (const rowNumber of invalidRows) {
        const rowOffset = rowNumber - 2;
        if (rowOffset < 0 || rowOffset >= rowCount)
            continue;
        timeFontColors[rowOffset][0] = INVALID_CART_ID_FONT_COLOR;
        lengthFontColors[rowOffset][0] = INVALID_CART_ID_FONT_COLOR;
    }
    logSheet.getRange(2, timeColumn, rowCount, 1).setFontColors(timeFontColors);
    logSheet.getRange(2, LENGTH_COLUMN, rowCount, 1).setFontColors(lengthFontColors);
    if (invalidGroups.length === 0) {
        if (!editedRange) {
            logSheet
                .getParent()
                .toast(`All active break groups total ${REQUIRED_BREAK_SECONDS} seconds.`, "Duration Check", 4);
        }
        return;
    }
    if (editedRange) {
        const editedRows = getEditedDataRows_(editedRange);
        const intersectsInvalidRows = editedRows.some((rowNumber) => invalidRows.has(rowNumber));
        if (!intersectsInvalidRows)
            return;
    }
    showInvalidBreakDurationToast_(logSheet, invalidGroups);
}
function showInvalidBreakDurationToast_(logSheet, invalidGroups) {
    const preview = invalidGroups
        .slice(0, 3)
        .map((group) => `${group.displayTime}: ${formatBreakSeconds_(group.totalSeconds)}s`)
        .join(", ");
    const remainingCount = invalidGroups.length - Math.min(invalidGroups.length, 3);
    const suffix = remainingCount > 0 ? `, +${remainingCount} more` : "";
    const message = `${invalidGroups.length} time slot(s) are not ${REQUIRED_BREAK_SECONDS}s: ` +
        `${preview}${suffix}.`;
    logSheet.getParent().toast(message, "Duration Check", 8);
}
function findTimeColumnIndex_(logSheet) {
    const lastColumn = Math.max(logSheet.getLastColumn(), DEFAULT_TIME_COLUMN);
    if (lastColumn < 1)
        return DEFAULT_TIME_COLUMN;
    const headers = logSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const normalizedHeaders = headers.map((value) => normalizeHeaderKey_(value));
    const candidates = ["time", "logtime", "airtime", "starttime"];
    for (const candidate of candidates) {
        const headerIndex = normalizedHeaders.indexOf(candidate);
        if (headerIndex < 0)
            continue;
        return headerIndex + 1;
    }
    return DEFAULT_TIME_COLUMN;
}
function normalizeTimeSlotKey_(value) {
    if (value === null || value === "")
        return "";
    if (value instanceof Date) {
        return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm:ss");
    }
    if (typeof value === "number" && Number.isFinite(value)) {
        if (value >= 0 && value < 1) {
            const totalSeconds = Math.round(value * 24 * 60 * 60);
            return `time:${totalSeconds}`;
        }
        return String(value);
    }
    const textValue = String(value).trim();
    if (!textValue)
        return "";
    const parsedSeconds = extractTimeOfDaySeconds_(textValue);
    if (parsedSeconds !== null)
        return `time:${parsedSeconds}`;
    return textValue.toUpperCase();
}
function extractTimeOfDaySeconds_(value) {
    const normalizedValue = value.trim().toUpperCase();
    const match = normalizedValue.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?$/);
    if (!match)
        return null;
    const hours = Number(match[1]);
    const minutes = Number(match[2]);
    const seconds = Number(match[3] || "0");
    const meridiem = match[4] || "";
    if (!Number.isInteger(hours) || !Number.isInteger(minutes) || !Number.isInteger(seconds))
        return null;
    if (minutes < 0 || minutes > 59 || seconds < 0 || seconds > 59)
        return null;
    let normalizedHours = hours;
    if (meridiem) {
        if (hours < 1 || hours > 12)
            return null;
        normalizedHours = hours % 12;
        if (meridiem === "PM")
            normalizedHours += 12;
    }
    else if (hours < 0 || hours > 23) {
        return null;
    }
    return normalizedHours * 60 * 60 + minutes * 60 + seconds;
}
function formatTimeSlotDisplay_(value) {
    if (value instanceof Date) {
        return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm:ss");
    }
    const textValue = String(value !== null && value !== void 0 ? value : "").trim();
    if (!textValue)
        return "(blank time)";
    return textValue;
}
function parseLengthSeconds_(value) {
    if (value === null || value === "")
        return null;
    if (typeof value === "number" && Number.isFinite(value)) {
        if (value >= 0 && value < 1) {
            return value * 24 * 60 * 60;
        }
        return value;
    }
    if (value instanceof Date) {
        return value.getHours() * 60 * 60 + value.getMinutes() * 60 + value.getSeconds();
    }
    const textValue = String(value).trim();
    if (!textValue)
        return null;
    const timeMatch = textValue.match(/^(\d+):(\d{1,2})(?::(\d{1,2}))?$/);
    if (timeMatch) {
        const first = Number(timeMatch[1]);
        const second = Number(timeMatch[2]);
        const third = Number(timeMatch[3] || "0");
        if (!timeMatch[3]) {
            return first * 60 + second;
        }
        return first * 60 * 60 + second * 60 + third;
    }
    const numericValue = Number(textValue.replace(/\s*s$/i, ""));
    if (Number.isFinite(numericValue))
        return numericValue;
    return null;
}
function getEditedDataRows_(editedRange) {
    const rows = [];
    const firstRow = Math.max(editedRange.getRow(), 2);
    const lastRow = editedRange.getRow() + editedRange.getNumRows() - 1;
    for (let row = firstRow; row <= lastRow; row++) {
        rows.push(row);
    }
    return rows;
}
function formatBreakSeconds_(seconds) {
    const roundedSeconds = Math.round(seconds * 10) / 10;
    if (Number.isInteger(roundedSeconds))
        return String(roundedSeconds);
    return roundedSeconds.toFixed(1);
}
function isLogSheetName_(sheetName) {
    return isNavigableLogSheetName_(sheetName);
}
function markCartIdAsValid_(entrySheet, row) {
    entrySheet.getRange(row, CART_ID_COLUMN).setFontColor(VALID_CART_ID_FONT_COLOR);
}
function markCartIdAsInvalid_(entrySheet, row) {
    entrySheet.getRange(row, CART_ID_COLUMN).setFontColor(INVALID_CART_ID_FONT_COLOR);
}
function rangeIncludesColumn_(range, column) {
    const startColumn = range.getColumn();
    const endColumn = startColumn + range.getNumColumns() - 1;
    return column >= startColumn && column <= endColumn;
}
function isSingleCellEdit_(range) {
    return range.getNumRows() === 1 && range.getNumColumns() === 1;
}
function isCellCleared_(value) {
    if (value === undefined || value === null)
        return true;
    if (typeof value !== "string")
        return false;
    return value.trim() === "";
}
function buildHeaderIndexByKey_(headers) {
    const indexByHeader = {};
    for (let index = 0; index < headers.length; index++) {
        const headerKey = normalizeHeaderKey_(headers[index]);
        if (!headerKey)
            continue;
        indexByHeader[headerKey] = index;
    }
    return indexByHeader;
}
function validateInventoryImportHeaders_(sourceHeaderIndexByKey) {
    const missingHeaders = [];
    for (const mapping of INVENTORY_IMPORT_COLUMN_MAPPING) {
        const normalizedHeader = normalizeHeaderKey_(mapping.sourceHeader);
        if (sourceHeaderIndexByKey[normalizedHeader] !== undefined)
            continue;
        missingHeaders.push(mapping.sourceHeader);
    }
    if (missingHeaders.length === 0)
        return;
    throw new Error(`Source Inventory sheet is missing required column(s): ${missingHeaders.join(", ")}.`);
}
function mapSourceInventoryRow_(sourceRow, sourceHeaderIndexByKey) {
    const mappedRow = new Array(INVENTORY_IMPORT_DESTINATION_COLUMN_COUNT).fill("");
    for (const mapping of INVENTORY_IMPORT_COLUMN_MAPPING) {
        const sourceIndex = sourceHeaderIndexByKey[normalizeHeaderKey_(mapping.sourceHeader)];
        if (sourceIndex === undefined)
            continue;
        const sourceValue = sourceRow[sourceIndex];
        if (mapping.destinationIndex === 3) {
            mappedRow[mapping.destinationIndex] = normalizeCartId_(sourceValue);
            continue;
        }
        mappedRow[mapping.destinationIndex] = sourceValue !== null && sourceValue !== void 0 ? sourceValue : "";
    }
    return mappedRow;
}
function normalizeCartId_(value) {
    if (value === null || value === undefined)
        return "";
    if (typeof value === "number" && Number.isFinite(value)) {
        const integerValue = Math.trunc(value);
        if (integerValue >= 0 && integerValue < 10000 && integerValue === value) {
            return String(integerValue).padStart(4, "0");
        }
        return String(value).trim();
    }
    const textValue = String(value).trim();
    if (!textValue)
        return "";
    const numericLikeMatch = textValue.match(/^(\d+)(?:\.0+)?$/);
    if (numericLikeMatch) {
        const digits = numericLikeMatch[1];
        if (digits.length < 4)
            return digits.padStart(4, "0");
        return digits;
    }
    if (/^\d+$/.test(textValue) && textValue.length < 4) {
        return textValue.padStart(4, "0");
    }
    return textValue;
}
function writeInventoryRows_(destinationSheet, rows) {
    const existingRowCount = Math.max(destinationSheet.getLastRow() - 1, 0);
    const cartRowCount = Math.max(existingRowCount, rows.length);
    if (cartRowCount > 0) {
        destinationSheet
            .getRange(2, 4, cartRowCount, 1)
            .setNumberFormat("@");
    }
    if (existingRowCount > 0) {
        destinationSheet
            .getRange(2, 1, existingRowCount, INVENTORY_IMPORT_DESTINATION_COLUMN_COUNT)
            .clearContent();
    }
    if (rows.length === 0)
        return;
    destinationSheet
        .getRange(2, 1, rows.length, INVENTORY_IMPORT_DESTINATION_COLUMN_COUNT)
        .setValues(rows);
}
function isInventoryRowBlank_(row) {
    for (const value of row) {
        if (!isCellValueBlank_(value))
            return false;
    }
    return true;
}
function isCellValueBlank_(value) {
    if (value === null || value === "")
        return true;
    if (typeof value !== "string")
        return false;
    return value.trim() === "";
}
function normalizeHeaderKey_(value) {
    return String(value !== null && value !== void 0 ? value : "")
        .trim()
        .toLowerCase()
        .replace(/[\s_]+/g, "");
}
// ─── ASC Export ──────────────────────────────────────────────────────────────
const ASC_CART_ID_PREFIX = "DA";
const ASC_TIME_OFFSET_SECONDS = 120; // +2 minutes
const ASC_EXPORT_TARGET_FOLDER_ID = "1r6IPJ9_N9nQuVvdGi3Yfhzkfc0MmxMHE";
const ASC_EXPORT_DAY_TOKENS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"];
const ASC_EXPORT_DAY_LABELS = {
    MON: "Monday",
    TUE: "Tuesday",
    WED: "Wednesday",
    THU: "Thursday",
    FRI: "Friday",
    SAT: "Saturday",
    SUN: "Sunday"
};
function exportActiveLogSheetToAsc() {
    showAscExportDialog_();
}
function showAscExportDialog_() {
    const html = HtmlService.createHtmlOutputFromFile("AscExportDialog")
        .setTitle("KRN Log Export")
        .setWidth(420)
        .setHeight(680);
    SpreadsheetApp.getUi().showModalDialog(html, "KRN Log Export Screen");
}
function getAscExportDialogState() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheetName = spreadsheet.getActiveSheet().getName();
    const availableDayTokens = getAvailableAscDayTokens_(spreadsheet);
    const activeSheetDayToken = getLogSheetDayToken_(activeSheetName) || "";
    const suggestedDayToken = activeSheetDayToken || availableDayTokens[0] || "";
    return {
        availableDayTokens,
        suggestedDayToken,
        dayLabels: ASC_EXPORT_DAY_LABELS
    };
}
function exportAscForPreset(preset) {
    const normalizedPreset = String(preset !== null && preset !== void 0 ? preset : "").trim().toUpperCase();
    switch (normalizedPreset) {
        case "ALL_DAYS":
            return exportAscForDayTokens_(ASC_EXPORT_DAY_TOKENS);
        case "TUE_SAT":
            return exportAscForDayTokens_(["TUE", "WED", "THU", "FRI", "SAT"]);
        case "WED_SAT":
            return exportAscForDayTokens_(["WED", "THU", "FRI", "SAT"]);
        case "THU_SAT":
            return exportAscForDayTokens_(["THU", "FRI", "SAT"]);
        case "FRI_SAT":
            return exportAscForDayTokens_(["FRI", "SAT"]);
        default:
            throw new Error(`Unsupported export preset: ${preset}`);
    }
}
function exportAscForRange(startDayToken, endDayToken) {
    const normalizedStart = normalizeAscDayToken_(startDayToken);
    const normalizedEnd = normalizeAscDayToken_(endDayToken);
    if (!normalizedStart || !normalizedEnd) {
        throw new Error("Choose a valid start and end day.");
    }
    let startIndex = ASC_EXPORT_DAY_TOKENS.indexOf(normalizedStart);
    let endIndex = ASC_EXPORT_DAY_TOKENS.indexOf(normalizedEnd);
    if (startIndex > endIndex) {
        const swap = startIndex;
        startIndex = endIndex;
        endIndex = swap;
    }
    const dayTokens = ASC_EXPORT_DAY_TOKENS.slice(startIndex, endIndex + 1);
    return exportAscForDayTokens_(dayTokens);
}
function exportAscForSpecificDays(dayTokens) {
    return exportAscForDayTokens_(dayTokens);
}
function exportAscForDayTokens_(dayTokens) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const normalizedDayTokens = normalizeAscDayTokens_(dayTokens);
    if (normalizedDayTokens.length === 0) {
        throw new Error("Choose at least one day to export.");
    }
    const files = [];
    const skippedDayTokens = [];
    for (const dayToken of normalizedDayTokens) {
        const daySheets = getAscExportSheetsForDay_(spreadsheet, dayToken);
        if (daySheets.length === 0) {
            skippedDayTokens.push(dayToken);
            continue;
        }
        const content = buildAscContentFromSheets_(daySheets);
        if (!content) {
            skippedDayTokens.push(dayToken);
            continue;
        }
        const fileName = buildAscFileNameForLogSheet_(spreadsheet, daySheets[0]);
        const file = saveAscFileToDrive_(fileName, content);
        const downloadUrl = "https://drive.google.com/uc?export=download&id=" + file.getId();
        const lineCount = content.split(/\r\n/).filter((line) => line.trim() !== "").length;
        files.push({
            dayToken,
            dayLabel: ASC_EXPORT_DAY_LABELS[dayToken] || dayToken,
            fileName,
            downloadUrl,
            lineCount,
            sheetNames: daySheets.map((sheet) => sheet.getName())
        });
    }
    return {
        generatedCount: files.length,
        skippedDayTokens,
        files,
        message: buildAscExportResultMessage_(files, skippedDayTokens)
    };
}
function buildAscExportResultMessage_(files, skippedDayTokens) {
    if (files.length === 0) {
        if (skippedDayTokens.length === 0)
            return "No ASC files were generated.";
        return `No ASC files were generated. Missing data for: ${skippedDayTokens.join(", ")}.`;
    }
    if (skippedDayTokens.length === 0) {
        return `Generated ${files.length} ASC file(s).`;
    }
    return (`Generated ${files.length} ASC file(s). ` +
        `Skipped: ${skippedDayTokens.join(", ")}.`);
}
function getAvailableAscDayTokens_(spreadsheet) {
    const dayTokens = new Set();
    for (const sheet of spreadsheet.getSheets()) {
        if (!isLogSheetName_(sheet.getName()))
            continue;
        const dayToken = getLogSheetDayToken_(sheet.getName());
        if (!dayToken)
            continue;
        dayTokens.add(dayToken);
    }
    return ASC_EXPORT_DAY_TOKENS.filter((dayToken) => dayTokens.has(dayToken));
}
function normalizeAscDayTokens_(dayTokens) {
    const deduped = new Set();
    for (const dayToken of dayTokens) {
        const normalized = normalizeAscDayToken_(dayToken);
        if (!normalized)
            continue;
        deduped.add(normalized);
    }
    return ASC_EXPORT_DAY_TOKENS.filter((dayToken) => deduped.has(dayToken));
}
function normalizeAscDayToken_(dayToken) {
    const normalized = String(dayToken !== null && dayToken !== void 0 ? dayToken : "").trim().toUpperCase();
    if (ASC_EXPORT_DAY_TOKENS.indexOf(normalized) < 0)
        return "";
    return normalized;
}
function getLogSheetDayToken_(sheetName) {
    const dayToken = extractDayAbbrFromSheetName_(sheetName);
    if (ASC_EXPORT_DAY_TOKENS.indexOf(dayToken) < 0)
        return "";
    return dayToken;
}
function getAscExportSheetsForDay_(spreadsheet, dayToken) {
    const normalizedDayToken = normalizeAscDayToken_(dayToken);
    if (!normalizedDayToken)
        return [];
    return spreadsheet
        .getSheets()
        .filter((sheet) => isLogSheetName_(sheet.getName()))
        .filter((sheet) => getLogSheetDayToken_(sheet.getName()) === normalizedDayToken)
        .sort((left, right) => {
        const leftMergeKey = getLogSheetAscMergeKey_(left.getName()) || "";
        const rightMergeKey = getLogSheetAscMergeKey_(right.getName()) || "";
        const mergeKeyDiff = leftMergeKey.localeCompare(rightMergeKey);
        if (mergeKeyDiff !== 0)
            return mergeKeyDiff;
        const stationOrderDiff = getAscStationOrder_(left.getName()) - getAscStationOrder_(right.getName());
        if (stationOrderDiff !== 0)
            return stationOrderDiff;
        return left.getName().localeCompare(right.getName());
    });
}
function exportLegacySingleAscFromActiveSheet_() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = spreadsheet.getActiveSheet();
    if (!isLogSheetName_(activeSheet.getName())) {
        spreadsheet.toast("ASC export works on KIN/KAN log sheets only.", "Export to ASC", 5);
        return;
    }
    const logSheetsForExport = getLogSheetsForAscExport_(spreadsheet, activeSheet);
    const content = buildAscContentFromSheets_(logSheetsForExport);
    const fileName = buildAscFileNameForLogSheet_(spreadsheet, activeSheet);
    const file = saveAscFileToDrive_(fileName, content);
    const downloadUrl = "https://drive.google.com/uc?export=download&id=" + file.getId();
    const safeFileName = escapeHtmlAttr_(fileName);
    const safeUrl = escapeHtmlAttr_(downloadUrl);
    const html = HtmlService.createHtmlOutput("<style>body{font-family:Arial,sans-serif;padding:16px;margin:0;}" +
        "a{color:#1a73e8;}p{margin:8px 0;}</style>" +
        "<p>&#10003; <strong>" + safeFileName + "</strong> is ready.</p>" +
        "<p><a href=\"" + safeUrl + "\" target=\"_blank\">Click here to download</a></p>" +
        "<p style=\"color:#666;font-size:12px;\">The file was saved to your ASC export folder in Google Drive.</p>" +
        "<br><input type=\"button\" value=\"Close\" onclick=\"google.script.host.close();\">" +
        "")
        .setWidth(400)
        .setHeight(170);
    SpreadsheetApp.getUi().showModalDialog(html, "Export to ASC");
}
function getLogSheetsForAscExport_(spreadsheet, activeSheet) {
    const mergeKey = getLogSheetAscMergeKey_(activeSheet.getName());
    if (!mergeKey)
        return [activeSheet];
    const matchingSheets = spreadsheet
        .getSheets()
        .filter((sheet) => isLogSheetName_(sheet.getName()))
        .filter((sheet) => getLogSheetAscMergeKey_(sheet.getName()) === mergeKey)
        .sort((left, right) => {
        const stationOrderDiff = getAscStationOrder_(left.getName()) - getAscStationOrder_(right.getName());
        if (stationOrderDiff !== 0)
            return stationOrderDiff;
        return left.getName().localeCompare(right.getName());
    });
    if (matchingSheets.length > 0)
        return matchingSheets;
    return [activeSheet];
}
function getLogSheetAscMergeKey_(sheetName) {
    const normalizedName = sheetName.trim().toUpperCase();
    const match = normalizedName.match(/^(KIN|KAN)_(.+)$/);
    if (!match)
        return null;
    return match[2];
}
function getAscStationOrder_(sheetName) {
    const normalizedName = sheetName.trim().toUpperCase();
    if (normalizedName.startsWith("KIN_"))
        return 0;
    if (normalizedName.startsWith("KAN_"))
        return 1;
    return 2;
}
function buildAscFileNameForLogSheet_(spreadsheet, logSheet) {
    const scriptTimeZone = Session.getScriptTimeZone();
    const logSheetName = logSheet.getName();
    const indexDate = getIndexDateForSheetName_(spreadsheet, logSheetName);
    if (indexDate instanceof Date && !Number.isNaN(indexDate.getTime())) {
        return Utilities.formatDate(indexDate, scriptTimeZone, "yyMMdd") + ".asc";
    }
    const fallbackDate = getDateForSheetFromMondayCell_(spreadsheet, logSheetName);
    if (fallbackDate instanceof Date && !Number.isNaN(fallbackDate.getTime())) {
        return Utilities.formatDate(fallbackDate, scriptTimeZone, "yyMMdd") + ".asc";
    }
    return logSheetName + ".asc";
}
function getIndexDateForSheetName_(spreadsheet, sheetName) {
    var _a;
    const indexSheet = spreadsheet.getSheetByName(INDEX_SHEET_NAME);
    if (!indexSheet)
        return null;
    const lastRow = indexSheet.getLastRow();
    if (lastRow < 2)
        return null;
    const rowCount = lastRow - 1;
    const nameValues = indexSheet
        .getRange(2, INDEX_SHEET_NAME_COLUMN, rowCount, 1)
        .getDisplayValues();
    const dateValues = indexSheet
        .getRange(2, INDEX_WEEK_PANEL_LABEL_COLUMN, rowCount, 1)
        .getValues();
    for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
        const candidateName = String((_a = nameValues[rowOffset][0]) !== null && _a !== void 0 ? _a : "").trim();
        if (candidateName !== sheetName)
            continue;
        const candidateDate = dateValues[rowOffset][0];
        if (candidateDate instanceof Date && !Number.isNaN(candidateDate.getTime())) {
            return candidateDate;
        }
        return null;
    }
    return null;
}
function getDateForSheetFromMondayCell_(spreadsheet, sheetName) {
    const indexSheet = spreadsheet.getSheetByName(INDEX_SHEET_NAME);
    if (!indexSheet)
        return null;
    const mondayValue = indexSheet.getRange(INDEX_WEEK_PANEL_MONDAY_A1).getValue();
    if (!(mondayValue instanceof Date) || Number.isNaN(mondayValue.getTime())) {
        return null;
    }
    const dayAbbr = extractDayAbbrFromSheetName_(sheetName);
    if (dayAbbr === "---")
        return null;
    const dayOffset = getDayOffsetFromMonday_(dayAbbr);
    const sheetDate = new Date(mondayValue);
    sheetDate.setDate(sheetDate.getDate() + dayOffset);
    return sheetDate;
}
function buildAscContentFromSheets_(logSheets) {
    var _a, _b;
    const ascRows = [];
    let sequence = 0;
    for (const logSheet of logSheets) {
        const lastRow = logSheet.getLastRow();
        if (lastRow < 2)
            continue;
        const rowCount = lastRow - 1;
        // Read all needed columns in one API call (cols 1-8)
        const allData = logSheet.getRange(2, 1, rowCount, LENGTH_COLUMN).getValues();
        for (let i = 0; i < rowCount; i++) {
            const cartId = normalizeCartId_(allData[i][CART_ID_COLUMN - 1]);
            if (!cartId)
                continue;
            const rawTimeValue = allData[i][DEFAULT_TIME_COLUMN - 1];
            const time = formatTimeForAsc_(rawTimeValue, ASC_TIME_OFFSET_SECONDS);
            const category = String((_a = allData[i][CATEGORY_COLUMN - 1]) !== null && _a !== void 0 ? _a : "").trim();
            const title = String((_b = allData[i][TITLE_COLUMN - 1]) !== null && _b !== void 0 ? _b : "").trim();
            const length = formatLengthForAsc_(allData[i][LENGTH_COLUMN - 1]);
            const prefixedCartId = ASC_CART_ID_PREFIX + cartId;
            // Format: TIME,,CATEGORY,CART_ID,TITLE,,LENGTH,0
            const line = time + ",," + category + "," + prefixedCartId + "," + title + ",," + length + ",0";
            const sortSeconds = getAscTimeSortSeconds_(rawTimeValue, ASC_TIME_OFFSET_SECONDS);
            ascRows.push({ sortSeconds, sequence, line });
            sequence++;
        }
    }
    ascRows.sort((left, right) => {
        if (left.sortSeconds !== right.sortSeconds)
            return left.sortSeconds - right.sortSeconds;
        return left.sequence - right.sequence;
    });
    const lines = [];
    for (const row of ascRows) {
        lines.push(row.line);
    }
    return lines.join("\r\n");
}
function getAscTimeSortSeconds_(value, offsetSeconds) {
    const formattedTime = formatTimeForAsc_(value, offsetSeconds);
    const parsedSeconds = extractTimeOfDaySeconds_(formattedTime);
    if (parsedSeconds === null)
        return Number.POSITIVE_INFINITY;
    return parsedSeconds;
}
function saveAscFileToDrive_(fileName, content) {
    let folder;
    try {
        folder = DriveApp.getFolderById(ASC_EXPORT_TARGET_FOLDER_ID);
    }
    catch (error) {
        throw new Error(`ASC export folder could not be opened. Check folder access for ID ${ASC_EXPORT_TARGET_FOLDER_ID}. (${error})`);
    }
    // Trash any existing file with the same name before creating a fresh one
    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
        existingFiles.next().setTrashed(true);
    }
    return folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
}
function formatTimeForAsc_(value, offsetSeconds = 0) {
    let totalSeconds = null;
    if (value instanceof Date) {
        totalSeconds = value.getHours() * 3600 + value.getMinutes() * 60 + value.getSeconds();
    }
    else if (typeof value === "number" && Number.isFinite(value) && value >= 0 && value < 1) {
        totalSeconds = Math.round(value * 86400);
    }
    else {
        const text = String(value !== null && value !== void 0 ? value : "").trim();
        totalSeconds = extractTimeOfDaySeconds_(text);
        if (totalSeconds === null) {
            // Unknown format — return as-is without offset
            if (/^\d{1,2}:\d{2}:\d{2}$/.test(text))
                return text.replace(/^(\d):/, "0$1:");
            if (/^\d{1,2}:\d{2}$/.test(text))
                return text.replace(/^(\d):/, "0$1:") + ":00";
            return text;
        }
    }
    const shifted = ((totalSeconds + offsetSeconds) % 86400 + 86400) % 86400;
    const h = Math.floor(shifted / 3600);
    const m = Math.floor((shifted % 3600) / 60);
    const s = shifted % 60;
    return (String(h).padStart(2, "0") + ":" +
        String(m).padStart(2, "0") + ":" +
        String(s).padStart(2, "0"));
}
function formatLengthForAsc_(value) {
    const totalSeconds = parseLengthSeconds_(value);
    if (totalSeconds === null)
        return String(value !== null && value !== void 0 ? value : "").trim();
    const normalized = normalizeAscLengthSeconds_(Math.round(totalSeconds));
    const h = Math.floor(normalized / 3600);
    const m = Math.floor((normalized % 3600) / 60);
    const s = normalized % 60;
    if (h > 0) {
        return (String(h).padStart(2, "0") + ":" +
            String(m).padStart(2, "0") + ":" +
            String(s).padStart(2, "0"));
    }
    return String(m).padStart(2, "0") + ":" + String(s).padStart(2, "0");
}
function normalizeAscLengthSeconds_(seconds) {
    // Snap to the nearest of 30s or 60s
    return Math.abs(seconds - 30) <= Math.abs(seconds - 60) ? 30 : 60;
}
function escapeHtmlAttr_(value) {
    return value
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
}
