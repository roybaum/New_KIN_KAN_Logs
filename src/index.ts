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
const DAY_NUMBER_TO_DAY_TOKEN: Record<number, string> = {
  1: "MON",
  2: "TUE",
  3: "WED",
  4: "THU",
  5: "FRI",
  6: "SAT"
};

const INVENTORY_IMPORT_COLUMN_MAPPING: Array<{ sourceHeader: string; destinationIndex: number }> = [
  { sourceHeader: "Title", destinationIndex: 0 },
  { sourceHeader: "Artist", destinationIndex: 1 },
  { sourceHeader: "Category", destinationIndex: 2 },
  { sourceHeader: "Number", destinationIndex: 3 },
  { sourceHeader: "LengthSeconds", destinationIndex: 4 },
  { sourceHeader: "StartDate", destinationIndex: 5 },
  { sourceHeader: "EndDate", destinationIndex: 6 }
];

const TITLE_COLUMN = 4; // D
const CART_ID_COLUMN = 7; // G
const LENGTH_COLUMN = 8; // H
const PICKER_COLUMN = 9; // I
const DEFAULT_TIME_COLUMN = 2; // B
const REQUIRED_BREAK_SECONDS = 60;
const BREAK_SECONDS_TOLERANCE = 1;
const VALID_CART_ID_FONT_COLOR = "#000000";
const INVALID_CART_ID_FONT_COLOR = "#d93025";

type CellValue = string | number | boolean | Date | null;

interface InventoryMatch {
  title: string;
  isci: CellValue;
  category: CellValue;
  cartId: string;
  length: CellValue;
  startDate: CellValue;
  endDate: CellValue;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui
    .createMenu("KIN KAN Tools")
    .addItem("Open Log Navigator", "showLogNavigatorDialog")
    .addSeparator()
    .addItem("Sync Inventory", "syncInventoryFromExternalWorkbook")
    .addItem("Check Break Lengths", "checkActiveLogSheetBreakDurations")
    .addSeparator()
    .addItem("Open Index", "openIndexSheet")
    .addItem("Refresh Index", "refreshIndexSheet")
    .addSubMenu(
      ui
        .createMenu("Jump")
        .addItem("Go To Today Log", "jumpToTodayLogSheet")
        .addItem("Go To Next Log", "jumpToNextLogSheet")
        .addItem("Go To Previous Log", "jumpToPreviousLogSheet")
        .addItem("Go To Log By Name", "jumpToLogSheetByNamePrompt")
    )
    .addToUi();
}

function showLogNavigatorDialog() {
  const html = HtmlService.createHtmlOutputFromFile("LogNavigatorDialog")
    .setTitle("Log Navigator")
    .setWidth(440)
    .setHeight(560);

  SpreadsheetApp.getUi().showModelessDialog(html, "Log Navigator");
}

function getLogSheetNamesForDialog(): string[] {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return getLogSheetsForNavigation_(spreadsheet).map((sheet) => sheet.getName());
}

function openLogSheetFromDialog(sheetName: string): { success: boolean; message: string } {
  const targetSheetName = String(sheetName ?? "").trim();
  if (!targetSheetName) {
    return { success: false, message: "Choose a log sheet first." };
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = spreadsheet.getSheetByName(targetSheetName);
  if (!targetSheet || !isNavigableLogSheetName_(targetSheetName)) {
    return { success: false, message: `Sheet \"${targetSheetName}\" is not a valid log sheet.` };
  }

  spreadsheet.setActiveSheet(targetSheet);
  return { success: true, message: `Opened ${targetSheetName}.` };
}

function openIndexSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = getOrCreateIndexSheet_(spreadsheet);
  refreshIndexSheet();
  spreadsheet.setActiveSheet(indexSheet);
  indexSheet.setActiveSelection("A1");
}

function refreshIndexSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = getOrCreateIndexSheet_(spreadsheet);
  const logSheets = getLogSheetsForNavigation_(spreadsheet);
  const longestSheetNameLength = getLongestSheetNameLength_(logSheets);

  indexSheet.clear();
  indexSheet.getRange(1, 1, 1, INDEX_HEADERS.length).setValues([INDEX_HEADERS]);

  if (logSheets.length > 0) {
    const indexRows = logSheets.map((sheet) => {
      return [sheet.getName(), false];
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

  indexSheet.setActiveSelection("A1");
  spreadsheet.toast(`Index refreshed with ${logSheets.length} log sheet(s).`, "Index", 4);
}

function getLongestSheetNameLength_(
  logSheets: GoogleAppsScript.Spreadsheet.Sheet[]
): number {
  let longestLength = INDEX_HEADERS[0].length;

  for (const sheet of logSheets) {
    const sheetNameLength = sheet.getName().length;
    if (sheetNameLength > longestLength) {
      longestLength = sheetNameLength;
    }
  }

  return longestLength;
}

function autoFitIndexSheetNameColumn_(
  indexSheet: GoogleAppsScript.Spreadsheet.Sheet,
  longestSheetNameLength: number
): void {
  SpreadsheetApp.flush();
  indexSheet.autoResizeColumn(INDEX_SHEET_NAME_COLUMN);

  const estimatedWidth = Math.round(
    longestSheetNameLength * INDEX_SHEET_NAME_CHAR_WIDTH_PX + INDEX_SHEET_NAME_PADDING_PX
  );
  const constrainedWidth = Math.min(
    INDEX_SHEET_NAME_MAX_WIDTH_PX,
    Math.max(INDEX_SHEET_NAME_MIN_WIDTH_PX, estimatedWidth)
  );

  if (indexSheet.getColumnWidth(INDEX_SHEET_NAME_COLUMN) >= constrainedWidth) return;

  indexSheet.setColumnWidth(INDEX_SHEET_NAME_COLUMN, constrainedWidth);
}

function formatIndexGoColumn_(
  indexSheet: GoogleAppsScript.Spreadsheet.Sheet,
  rowCount: number
): void {
  if (rowCount < 1) return;

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

  const candidateSheets = getLogSheetsForNavigation_(spreadsheet).filter((sheet) =>
    sheet.getName().toUpperCase().includes(todayDayToken)
  );
  if (candidateSheets.length === 0) {
    spreadsheet.toast(`No ${todayDayToken} log sheet was found.`, "Jump", 5);
    return;
  }

  const targetSheet = candidateSheets[candidateSheets.length - 1];
  spreadsheet.setActiveSheet(targetSheet);

  if (candidateSheets.length > 1) {
    spreadsheet.toast(
      `Multiple ${todayDayToken} sheets found. Opened ${targetSheet.getName()}.`,
      "Jump",
      5
    );
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
  const promptMessage =
    `Enter a log sheet name. Examples: ${previewNames}` +
    (hiddenCount > 0 ? `, +${hiddenCount} more` : "");

  const response = ui.prompt("Go To Log Sheet", promptMessage, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const targetSheetName = response.getResponseText().trim();
  if (!targetSheetName) return;

  const targetSheet = spreadsheet.getSheetByName(targetSheetName);
  if (!targetSheet || !isNavigableLogSheetName_(targetSheetName)) {
    spreadsheet.toast(`Sheet \"${targetSheetName}\" is not a valid log sheet.`, "Jump", 5);
    return;
  }

  spreadsheet.setActiveSheet(targetSheet);
}

function getOrCreateIndexSheet_(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
): GoogleAppsScript.Spreadsheet.Sheet {
  const existingIndexSheet = spreadsheet.getSheetByName(INDEX_SHEET_NAME);
  if (existingIndexSheet) return existingIndexSheet;

  const indexSheet = spreadsheet.insertSheet(INDEX_SHEET_NAME, 0);
  indexSheet.getRange(1, 1).setValue("Log sheet index");
  return indexSheet;
}

function getLogSheetsForNavigation_(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
): GoogleAppsScript.Spreadsheet.Sheet[] {
  return spreadsheet
    .getSheets()
    .filter((sheet) => isNavigableLogSheetName_(sheet.getName()))
    .sort((left, right) => left.getName().localeCompare(right.getName()));
}

function isNavigableLogSheetName_(sheetName: string): boolean {
  const normalizedName = sheetName.trim().toUpperCase();
  if (normalizedName === INVENTORY_SHEET_NAME.toUpperCase()) return false;
  if (normalizedName === INDEX_SHEET_NAME.toUpperCase()) return false;

  return LOG_SHEET_REQUIRED_TOKENS.some((token) => hasStandaloneSheetToken_(normalizedName, token));
}

function hasStandaloneSheetToken_(normalizedSheetName: string, token: string): boolean {
  if (normalizedSheetName.startsWith(token)) return true;

  const tokenRegex = new RegExp(`(^|[^A-Z])${token}([^A-Z]|$)`);
  return tokenRegex.test(normalizedSheetName);
}

function getTodayDayToken_(): string {
  const dayNumber = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "u"));
  return DAY_NUMBER_TO_DAY_TOKEN[dayNumber] ?? "";
}

function jumpToRelativeLogSheet_(direction: number): void {
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
    spreadsheet.setActiveSheet(logSheets[fallbackIndex]);
    return;
  }

  const targetIndex = (currentIndex + direction + logSheets.length) % logSheets.length;
  spreadsheet.setActiveSheet(logSheets[targetIndex]);
}

function syncInventoryFromExternalWorkbook() {
  const destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const destinationSheet = destinationSpreadsheet.getSheetByName(INVENTORY_SHEET_NAME);
  if (!destinationSheet) {
    throw new Error(
      `Destination sheet "${INVENTORY_SHEET_NAME}" was not found in the active spreadsheet.`
    );
  }

  const sourceSpreadsheet = SpreadsheetApp.openById(INVENTORY_IMPORT_SOURCE_SPREADSHEET_ID);
  const sourceSheet = sourceSpreadsheet.getSheetByName(INVENTORY_IMPORT_SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    throw new Error(
      `Source sheet "${INVENTORY_IMPORT_SOURCE_SHEET_NAME}" was not found in the source spreadsheet.`
    );
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
  const sourceRows =
    sourceLastRow > 1
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

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit | undefined) {
  if (!e) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === INDEX_SHEET_NAME) {
    handleIndexGoEdit_(sheet, e.range, e.value);
    return;
  }

  if (!isLogSheetName_(sheetName)) return;

  const firstEditedRow = e.range.getRow();
  const lastEditedRow = firstEditedRow + e.range.getNumRows() - 1;
  if (lastEditedRow < 2) return;

  try {
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < 2) return;

    const singleCellEdit = isSingleCellEdit_(e.range);
    const cartIdRangeEdited = rangeIncludesColumn_(e.range, CART_ID_COLUMN);

    if (cartIdRangeEdited) {
      processCartIdRangeEdit_(sheet, e.range);
      if (!singleCellEdit || col === CART_ID_COLUMN) return;
    }

    if (!singleCellEdit) return;

    if (col === PICKER_COLUMN) {
      applyPickerSelection_(sheet, row, String(e.value || ""));
      return;
    }

    // Only respond to Title (D) here. Cart ID (G) is handled above.
    if (col !== TITLE_COLUMN) return;

    clearPicker_(sheet, row);

    const searchValue = String(e.value || "").trim().toLowerCase();
    if (!searchValue) return;

    const inventorySheet = SpreadsheetApp.getActive().getSheetByName(INVENTORY_SHEET_NAME);
    if (!inventorySheet) return;

    const activeInventory = getActiveInventoryMatches_(inventorySheet);
    const matches = findMatchesInInventory_(activeInventory, searchValue);
    if (matches.length === 0) return;

    if (matches.length === 1) {
      applyMatchToEntryRow_(sheet, row, matches[0]);
      return;
    }

    setPickerForMatches_(sheet, row, matches);
  } finally {
    validateLogSheetBreakDurations_(sheet, e.range);
  }
}

function handleIndexGoEdit_(
  indexSheet: GoogleAppsScript.Spreadsheet.Sheet,
  editedRange: GoogleAppsScript.Spreadsheet.Range,
  editedValue: string | undefined
): void {
  if (!isSingleCellEdit_(editedRange)) return;
  if (editedRange.getRow() < 2) return;
  if (editedRange.getColumn() !== INDEX_NAVIGATION_COLUMN) return;

  const isChecked = String(editedValue || "").toUpperCase() === "TRUE";
  if (!isChecked) return;

  enforceSingleIndexGoSelection_(indexSheet, editedRange.getRow());

  const targetSheetName = String(
    indexSheet.getRange(editedRange.getRow(), INDEX_SHEET_NAME_COLUMN).getDisplayValue()
  ).trim();

  if (!targetSheetName) {
    editedRange.setValue(false);
    return;
  }

  const spreadsheet = indexSheet.getParent();
  const targetSheet = spreadsheet.getSheetByName(targetSheetName);
  if (!targetSheet) {
    editedRange.setValue(false);
    spreadsheet.toast(`Sheet "${targetSheetName}" was not found.`, "Index", 5);
    return;
  }

  spreadsheet.setActiveSheet(targetSheet);
}

function enforceSingleIndexGoSelection_(
  indexSheet: GoogleAppsScript.Spreadsheet.Sheet,
  selectedRow: number
): void {
  const lastRow = indexSheet.getLastRow();
  if (lastRow < 2) return;

  if (selectedRow > 2) {
    indexSheet
      .getRange(2, INDEX_NAVIGATION_COLUMN, selectedRow - 2, 1)
      .setValue(false);
  }

  if (selectedRow < lastRow) {
    indexSheet
      .getRange(selectedRow + 1, INDEX_NAVIGATION_COLUMN, lastRow - selectedRow, 1)
      .setValue(false);
  }

  indexSheet.getRange(selectedRow, INDEX_NAVIGATION_COLUMN).setValue(true);
}

function findInventoryMatches_(
  inventorySheet: GoogleAppsScript.Spreadsheet.Sheet,
  searchValue: string
): InventoryMatch[] {
  const activeInventory = getActiveInventoryMatches_(inventorySheet);
  return findMatchesInInventory_(activeInventory, searchValue);
}

function getActiveInventoryMatches_(inventorySheet: GoogleAppsScript.Spreadsheet.Sheet): InventoryMatch[] {
  const lastRow = inventorySheet.getLastRow();
  if (lastRow < 2) return [];

  const data = inventorySheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const activeInventory: InventoryMatch[] = [];

  for (const item of data) {
    const match: InventoryMatch = {
      title: String(item[0]),
      isci: item[1],
      category: item[2],
      cartId: String(item[3]),
      length: item[4],
      startDate: item[5],
      endDate: item[6]
    };

    if (!isActiveToday_(today, match.startDate, match.endDate)) continue;

    activeInventory.push(match);
  }

  return activeInventory;
}

function findMatchesInInventory_(inventory: InventoryMatch[], searchValue: string): InventoryMatch[] {
  const normalizedSearch = searchValue.trim().toLowerCase();
  if (!normalizedSearch) return [];

  const matches: InventoryMatch[] = [];

  for (const item of inventory) {
    const titleMatch = item.title.toLowerCase().includes(normalizedSearch);
    const cartMatch = item.cartId.toLowerCase().includes(normalizedSearch);
    if (!titleMatch && !cartMatch) continue;
    matches.push(item);
  }

  return matches;
}

function isActiveToday_(today: Date, startDate: CellValue, endDate: CellValue): boolean {
  if (
    startDate !== null &&
    startDate !== "" &&
    (startDate instanceof Date || typeof startDate === "string" || typeof startDate === "number")
  ) {
    const start = new Date(startDate);
    if (!Number.isNaN(start.getTime())) {
      start.setHours(0, 0, 0, 0);
      if (today < start) return false;
    }
  }

  if (
    endDate !== null &&
    endDate !== "" &&
    (endDate instanceof Date || typeof endDate === "string" || typeof endDate === "number")
  ) {
    const end = new Date(endDate);
    if (!Number.isNaN(end.getTime())) {
      end.setHours(0, 0, 0, 0);
      if (today > end) return false;
    }
  }

  return true;
}

function applyMatchToEntryRow_(
  entrySheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  match: InventoryMatch
) {
  entrySheet.getRange(row, TITLE_COLUMN, 1, 5).setValues([[
    match.title,
    match.isci,
    match.category,
    match.cartId,
    match.length
  ]]);

  markCartIdAsValid_(entrySheet, row);
}

function setPickerForMatches_(
  entrySheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  matches: InventoryMatch[]
) {
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

function applyPickerSelection_(
  entrySheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  pickerValue: string
) {
  if (!pickerValue) return;

  const cartId = parsePickerCartId_(pickerValue);
  if (!cartId) {
    entrySheet.getRange(row, PICKER_COLUMN).clearContent();
    return;
  }

  const inventorySheet = SpreadsheetApp.getActive().getSheetByName(INVENTORY_SHEET_NAME);
  if (!inventorySheet) return;

  const activeInventory = getActiveInventoryMatches_(inventorySheet);
  const selectedMatch = activeInventory.find(
    (match) => match.cartId.trim().toLowerCase() === cartId.toLowerCase()
  );
  if (!selectedMatch) return;

  applyMatchToEntryRow_(entrySheet, row, selectedMatch);
  clearPicker_(entrySheet, row);
}

function parsePickerCartId_(value: string): string {
  const cartIdMatch = value.match(/\(([^()]*)\)\s*$/);
  if (!cartIdMatch) return "";
  return cartIdMatch[1].trim();
}

function formatPickerOption_(match: InventoryMatch): string {
  const maxTitleLength = 45;
  const shortTitle =
    match.title.length > maxTitleLength
      ? `${match.title.slice(0, maxTitleLength - 3)}...`
      : match.title;

  return `${shortTitle} (${match.cartId})`;
}

function clearPicker_(entrySheet: GoogleAppsScript.Spreadsheet.Sheet, row: number) {
  const pickerCell = entrySheet.getRange(row, PICKER_COLUMN);
  pickerCell.clearContent();
  pickerCell.clearDataValidations();
  pickerCell.clearNote();
}

function clearEntryRow_(entrySheet: GoogleAppsScript.Spreadsheet.Sheet, row: number) {
  const rowRange = entrySheet.getRange(row, TITLE_COLUMN, 1, 5);
  rowRange.clearContent();
  rowRange.clearDataValidations();
  rowRange.clearNote();
  markCartIdAsValid_(entrySheet, row);
}

function processCartIdRangeEdit_(
  entrySheet: GoogleAppsScript.Spreadsheet.Sheet,
  editedRange: GoogleAppsScript.Spreadsheet.Range
): void {
  if (!rangeIncludesColumn_(editedRange, CART_ID_COLUMN)) return;

  const cartIdOffset = CART_ID_COLUMN - editedRange.getColumn();
  const values = editedRange.getValues();
  const inventorySheet = SpreadsheetApp.getActive().getSheetByName(INVENTORY_SHEET_NAME);
  const activeInventory = inventorySheet ? getActiveInventoryMatches_(inventorySheet) : [];
  const invalidCarts: Array<{ row: number; cartId: string }> = [];

  for (let rowOffset = 0; rowOffset < values.length; rowOffset++) {
    const targetRow = editedRange.getRow() + rowOffset;
    if (targetRow < 2) continue;

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

    const searchValue = String(cartIdValue).trim().toLowerCase();
    const matches = findMatchesInInventory_(activeInventory, searchValue);
    if (matches.length === 0) {
      invalidCarts.push({ row: targetRow, cartId: String(cartIdValue).trim() });
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

function showInvalidCartToast_(
  entrySheet: GoogleAppsScript.Spreadsheet.Sheet,
  invalidCarts: Array<{ row: number; cartId: string }>
): void {
  const spreadsheet = entrySheet.getParent();

  if (invalidCarts.length === 1) {
    const invalidCart = invalidCarts[0];
    const message =
      `Row ${invalidCart.row}: Cart ID "${invalidCart.cartId}" is invalid ` +
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

function validateLogSheetBreakDurations_(
  logSheet: GoogleAppsScript.Spreadsheet.Sheet,
  editedRange?: GoogleAppsScript.Spreadsheet.Range
): void {
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

  const groups: Record<
    string,
    { rows: number[]; totalSeconds: number; hasCommercialData: boolean; displayTime: string }
  > = {};

  for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
    const timeValue = timeValues[rowOffset][0];
    const timeKey = normalizeTimeSlotKey_(timeValue);
    if (!timeKey) continue;

    const rowNumber = rowOffset + 2;
    const group =
      groups[timeKey] ??
      {
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

    const hasCommercialData =
      !isCellValueBlank_(titleValues[rowOffset][0] as CellValue) ||
      !isCellValueBlank_(cartValues[rowOffset][0] as CellValue) ||
      !isCellValueBlank_(lengthValue as CellValue);

    group.hasCommercialData = group.hasCommercialData || hasCommercialData;
    group.rows.push(rowNumber);
    groups[timeKey] = group;
  }

  const invalidRows = new Set<number>();
  const invalidGroups: Array<{ displayTime: string; totalSeconds: number }> = [];

  for (const group of Object.values(groups)) {
    if (!group.hasCommercialData) continue;

    if (Math.abs(group.totalSeconds - REQUIRED_BREAK_SECONDS) <= BREAK_SECONDS_TOLERANCE) continue;

    invalidGroups.push({
      displayTime: group.displayTime,
      totalSeconds: group.totalSeconds
    });

    for (const rowNumber of group.rows) {
      invalidRows.add(rowNumber);
    }
  }

  const timeFontColors: string[][] = Array.from({ length: rowCount }, () => [VALID_CART_ID_FONT_COLOR]);
  const lengthFontColors: string[][] = Array.from({ length: rowCount }, () => [VALID_CART_ID_FONT_COLOR]);

  for (const rowNumber of invalidRows) {
    const rowOffset = rowNumber - 2;
    if (rowOffset < 0 || rowOffset >= rowCount) continue;
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
    if (!intersectsInvalidRows) return;
  }

  showInvalidBreakDurationToast_(logSheet, invalidGroups);
}

function showInvalidBreakDurationToast_(
  logSheet: GoogleAppsScript.Spreadsheet.Sheet,
  invalidGroups: Array<{ displayTime: string; totalSeconds: number }>
): void {
  const preview = invalidGroups
    .slice(0, 3)
    .map((group) => `${group.displayTime}: ${formatBreakSeconds_(group.totalSeconds)}s`)
    .join(", ");
  const remainingCount = invalidGroups.length - Math.min(invalidGroups.length, 3);
  const suffix = remainingCount > 0 ? `, +${remainingCount} more` : "";
  const message =
    `${invalidGroups.length} time slot(s) are not ${REQUIRED_BREAK_SECONDS}s: ` +
    `${preview}${suffix}.`;

  logSheet.getParent().toast(message, "Duration Check", 8);
}

function findTimeColumnIndex_(logSheet: GoogleAppsScript.Spreadsheet.Sheet): number {
  const lastColumn = Math.max(logSheet.getLastColumn(), DEFAULT_TIME_COLUMN);
  if (lastColumn < 1) return DEFAULT_TIME_COLUMN;

  const headers = logSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const normalizedHeaders = headers.map((value) => normalizeHeaderKey_(value));
  const candidates = ["time", "logtime", "airtime", "starttime"];

  for (const candidate of candidates) {
    const headerIndex = normalizedHeaders.indexOf(candidate);
    if (headerIndex < 0) continue;
    return headerIndex + 1;
  }

  return DEFAULT_TIME_COLUMN;
}

function normalizeTimeSlotKey_(value: unknown): string {
  if (value === null || value === "") return "";

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
  if (!textValue) return "";

  const parsedSeconds = extractTimeOfDaySeconds_(textValue);
  if (parsedSeconds !== null) return `time:${parsedSeconds}`;

  return textValue.toUpperCase();
}

function extractTimeOfDaySeconds_(value: string): number | null {
  const normalizedValue = value.trim().toUpperCase();
  const match = normalizedValue.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?$/);
  if (!match) return null;

  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3] || "0");
  const meridiem = match[4] || "";

  if (!Number.isInteger(hours) || !Number.isInteger(minutes) || !Number.isInteger(seconds)) return null;
  if (minutes < 0 || minutes > 59 || seconds < 0 || seconds > 59) return null;

  let normalizedHours = hours;
  if (meridiem) {
    if (hours < 1 || hours > 12) return null;
    normalizedHours = hours % 12;
    if (meridiem === "PM") normalizedHours += 12;
  } else if (hours < 0 || hours > 23) {
    return null;
  }

  return normalizedHours * 60 * 60 + minutes * 60 + seconds;
}

function formatTimeSlotDisplay_(value: unknown): string {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm:ss");
  }

  const textValue = String(value ?? "").trim();
  if (!textValue) return "(blank time)";
  return textValue;
}

function parseLengthSeconds_(value: unknown): number | null {
  if (value === null || value === "") return null;

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
  if (!textValue) return null;

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
  if (Number.isFinite(numericValue)) return numericValue;

  return null;
}

function getEditedDataRows_(editedRange: GoogleAppsScript.Spreadsheet.Range): number[] {
  const rows: number[] = [];
  const firstRow = Math.max(editedRange.getRow(), 2);
  const lastRow = editedRange.getRow() + editedRange.getNumRows() - 1;

  for (let row = firstRow; row <= lastRow; row++) {
    rows.push(row);
  }

  return rows;
}

function formatBreakSeconds_(seconds: number): string {
  const roundedSeconds = Math.round(seconds * 10) / 10;
  if (Number.isInteger(roundedSeconds)) return String(roundedSeconds);
  return roundedSeconds.toFixed(1);
}

function isLogSheetName_(sheetName: string): boolean {
  return isNavigableLogSheetName_(sheetName);
}

function markCartIdAsValid_(entrySheet: GoogleAppsScript.Spreadsheet.Sheet, row: number): void {
  entrySheet.getRange(row, CART_ID_COLUMN).setFontColor(VALID_CART_ID_FONT_COLOR);
}

function markCartIdAsInvalid_(entrySheet: GoogleAppsScript.Spreadsheet.Sheet, row: number): void {
  entrySheet.getRange(row, CART_ID_COLUMN).setFontColor(INVALID_CART_ID_FONT_COLOR);
}

function rangeIncludesColumn_(
  range: GoogleAppsScript.Spreadsheet.Range,
  column: number
): boolean {
  const startColumn = range.getColumn();
  const endColumn = startColumn + range.getNumColumns() - 1;
  return column >= startColumn && column <= endColumn;
}

function isSingleCellEdit_(range: GoogleAppsScript.Spreadsheet.Range): boolean {
  return range.getNumRows() === 1 && range.getNumColumns() === 1;
}

function isCellCleared_(value: unknown): boolean {
  if (value === undefined || value === null) return true;
  if (typeof value !== "string") return false;
  return value.trim() === "";
}

function buildHeaderIndexByKey_(headers: unknown[]): Record<string, number> {
  const indexByHeader: Record<string, number> = {};

  for (let index = 0; index < headers.length; index++) {
    const headerKey = normalizeHeaderKey_(headers[index]);
    if (!headerKey) continue;
    indexByHeader[headerKey] = index;
  }

  return indexByHeader;
}

function validateInventoryImportHeaders_(sourceHeaderIndexByKey: Record<string, number>): void {
  const missingHeaders: string[] = [];

  for (const mapping of INVENTORY_IMPORT_COLUMN_MAPPING) {
    const normalizedHeader = normalizeHeaderKey_(mapping.sourceHeader);
    if (sourceHeaderIndexByKey[normalizedHeader] !== undefined) continue;
    missingHeaders.push(mapping.sourceHeader);
  }

  if (missingHeaders.length === 0) return;

  throw new Error(
    `Source Inventory sheet is missing required column(s): ${missingHeaders.join(", ")}.`
  );
}

function mapSourceInventoryRow_(
  sourceRow: unknown[],
  sourceHeaderIndexByKey: Record<string, number>
): CellValue[] {
  const mappedRow: CellValue[] = new Array(INVENTORY_IMPORT_DESTINATION_COLUMN_COUNT).fill("");

  for (const mapping of INVENTORY_IMPORT_COLUMN_MAPPING) {
    const sourceIndex = sourceHeaderIndexByKey[normalizeHeaderKey_(mapping.sourceHeader)];
    if (sourceIndex === undefined) continue;
    const sourceValue = sourceRow[sourceIndex] as CellValue;
    mappedRow[mapping.destinationIndex] = sourceValue ?? "";
  }

  return mappedRow;
}

function writeInventoryRows_(
  destinationSheet: GoogleAppsScript.Spreadsheet.Sheet,
  rows: CellValue[][]
): void {
  const existingRowCount = Math.max(destinationSheet.getLastRow() - 1, 0);
  if (existingRowCount > 0) {
    destinationSheet
      .getRange(2, 1, existingRowCount, INVENTORY_IMPORT_DESTINATION_COLUMN_COUNT)
      .clearContent();
  }

  if (rows.length === 0) return;

  destinationSheet
    .getRange(2, 1, rows.length, INVENTORY_IMPORT_DESTINATION_COLUMN_COUNT)
    .setValues(rows);
}

function isInventoryRowBlank_(row: CellValue[]): boolean {
  for (const value of row) {
    if (!isCellValueBlank_(value)) return false;
  }

  return true;
}

function isCellValueBlank_(value: CellValue): boolean {
  if (value === null || value === "") return true;
  if (typeof value !== "string") return false;
  return value.trim() === "";
}

function normalizeHeaderKey_(value: unknown): string {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[\s_]+/g, "");
}