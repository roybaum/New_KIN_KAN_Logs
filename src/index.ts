const LOG_SHEET_NAMES = ["KIN_MON", "KIN_TUE", "KIN_WED", "KIN_THU", "KIN_FRI", "KIN_SAT"];
const LOG_SHEET_NAME_PREFIX = "KIN_";
const INDEX_SHEET_NAME = "Index";
const INVENTORY_SHEET_NAME = "Inventory";
const INVENTORY_IMPORT_SOURCE_SPREADSHEET_ID = "1QYBk6N_RZygLDPWV8BjVpF2azXBCyvGNRuz9XvpakPE";
const INVENTORY_IMPORT_SOURCE_SHEET_NAME = "Inventory";
const INVENTORY_IMPORT_DESTINATION_COLUMN_COUNT = 7;
const INDEX_HEADERS = ["Sheet Name", "Open", "Data Rows", "Last Column", "Sheet ID"];
const DAY_NUMBER_TO_LOG_SHEET_NAME: Record<number, string> = {
  1: "KIN_MON",
  2: "KIN_TUE",
  3: "KIN_WED",
  4: "KIN_THU",
  5: "KIN_FRI",
  6: "KIN_SAT"
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
const PICKER_COLUMN = 9; // I
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
    .addItem("Sync Inventory", "syncInventoryFromExternalWorkbook")
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

function openIndexSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = getOrCreateIndexSheet_(spreadsheet);
  refreshIndexSheet();
  spreadsheet.setActiveSheet(indexSheet);
}

function refreshIndexSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = getOrCreateIndexSheet_(spreadsheet);
  const logSheets = getLogSheetsForNavigation_(spreadsheet);

  indexSheet.clear();
  indexSheet.getRange(1, 1, 1, INDEX_HEADERS.length).setValues([INDEX_HEADERS]);

  if (logSheets.length > 0) {
    const indexRows = logSheets.map((sheet) => {
      const sheetId = sheet.getSheetId();
      const openFormula = `=HYPERLINK("#gid=${sheetId}", "Open")`;
      const dataRows = Math.max(sheet.getLastRow() - 1, 0);

      return [
        sheet.getName(),
        openFormula,
        dataRows,
        sheet.getLastColumn(),
        sheetId
      ];
    });

    indexSheet
      .getRange(2, 1, indexRows.length, INDEX_HEADERS.length)
      .setValues(indexRows);
  }

  const headerRange = indexSheet.getRange(1, 1, 1, INDEX_HEADERS.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#dbe9ff");

  indexSheet.setFrozenRows(1);
  indexSheet.autoResizeColumns(1, INDEX_HEADERS.length);
  spreadsheet.toast(`Index refreshed with ${logSheets.length} log sheet(s).`, "Index", 4);
}

function jumpToTodayLogSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const todayLogSheetName = getTodayLogSheetName_();
  if (!todayLogSheetName) {
    spreadsheet.toast("No log sheet mapping exists for today.", "Jump", 5);
    return;
  }

  const targetSheet = spreadsheet.getSheetByName(todayLogSheetName);
  if (!targetSheet) {
    spreadsheet.toast(`Sheet \"${todayLogSheetName}\" was not found.`, "Jump", 5);
    return;
  }

  spreadsheet.setActiveSheet(targetSheet);
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
  if (sheetName === INVENTORY_SHEET_NAME) return false;
  if (sheetName === INDEX_SHEET_NAME) return false;
  return sheetName.startsWith(LOG_SHEET_NAME_PREFIX);
}

function getTodayLogSheetName_(): string {
  const dayNumber = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "u"));
  return DAY_NUMBER_TO_LOG_SHEET_NAME[dayNumber] ?? "";
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

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit | undefined) {
  if (!e) return;

  const sheet = e.range.getSheet();
  if (!isLogSheetName_(sheet.getName())) return;

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

function isLogSheetName_(sheetName: string): boolean {
  return LOG_SHEET_NAMES.includes(sheetName);
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