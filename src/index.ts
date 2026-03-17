const ENTRY_SHEET_NAME = "Entry";
const INVENTORY_SHEET_NAME = "Inventory";

const TITLE_COLUMN = 4; // D
const CART_ID_COLUMN = 7; // G
const PICKER_COLUMN = 9; // I

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

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit | undefined) {
  if (!e) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== ENTRY_SHEET_NAME) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  if (col === CART_ID_COLUMN && isCellCleared_(e.value)) {
    clearEntryRow_(sheet, row);
    return;
  }

  if (col === PICKER_COLUMN) {
    applyPickerSelection_(sheet, row, String(e.value || ""));
    return;
  }

  // Only respond to Title (D) or Cart ID (G)
  if (col !== TITLE_COLUMN && col !== CART_ID_COLUMN) return;

  clearPicker_(sheet, row);

  const searchValue = String(e.value || "").trim().toLowerCase();
  if (!searchValue) return;

  const inventorySheet = SpreadsheetApp.getActive().getSheetByName(INVENTORY_SHEET_NAME);
  if (!inventorySheet) return;

  const matches = findInventoryMatches_(inventorySheet, searchValue);
  if (matches.length === 0) return;

  if (matches.length === 1) {
    applyMatchToEntryRow_(sheet, row, matches[0]);
    return;
  }

  setPickerForMatches_(sheet, row, searchValue, matches);
}

function findInventoryMatches_(
  inventorySheet: GoogleAppsScript.Spreadsheet.Sheet,
  searchValue: string
): InventoryMatch[] {
  const lastRow = inventorySheet.getLastRow();
  if (lastRow < 2) return [];

  const data = inventorySheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const matches: InventoryMatch[] = [];

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

    const titleMatch = match.title.toLowerCase().includes(searchValue);
    const cartMatch = match.cartId.toLowerCase().includes(searchValue);
    if (!titleMatch && !cartMatch) continue;

    if (!isActiveToday_(today, match.startDate, match.endDate)) continue;

    matches.push(match);
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
}

function setPickerForMatches_(
  entrySheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  searchValue: string,
  matches: InventoryMatch[]
) {
  const options = matches.map((match, index) => formatPickerOption_(index, match));
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .setHelpText("Select the matching inventory item.")
    .build();

  const pickerCell = entrySheet.getRange(row, PICKER_COLUMN);
  pickerCell.clearContent();
  pickerCell.setDataValidation(rule);
  pickerCell.setNote(JSON.stringify({ searchValue }));
}

function applyPickerSelection_(
  entrySheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  pickerValue: string
) {
  if (!pickerValue) return;

  const pickerCell = entrySheet.getRange(row, PICKER_COLUMN);
  const note = pickerCell.getNote();
  if (!note) return;

  const matchIndex = parsePickerIndex_(pickerValue);
  if (matchIndex < 0) return;

  let metadata: { searchValue: string };
  try {
    metadata = JSON.parse(note) as { searchValue: string };
  } catch {
    return;
  }

  if (!metadata.searchValue) return;

  const inventorySheet = SpreadsheetApp.getActive().getSheetByName(INVENTORY_SHEET_NAME);
  if (!inventorySheet) return;

  const matches = findInventoryMatches_(inventorySheet, metadata.searchValue);
  if (matchIndex >= matches.length) return;

  applyMatchToEntryRow_(entrySheet, row, matches[matchIndex]);
  clearPicker_(entrySheet, row);
}

function parsePickerIndex_(value: string): number {
  const indexMatch = value.match(/^\[(\d+)\]/);
  if (!indexMatch) return -1;

  const zeroBasedIndex = Number(indexMatch[1]) - 1;
  if (!Number.isInteger(zeroBasedIndex) || zeroBasedIndex < 0) return -1;

  return zeroBasedIndex;
}

function formatPickerOption_(index: number, match: InventoryMatch): string {
  const maxTitleLength = 45;
  const shortTitle =
    match.title.length > maxTitleLength
      ? `${match.title.slice(0, maxTitleLength - 3)}...`
      : match.title;

  return `[${index + 1}] ${shortTitle} | ${match.cartId}`;
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
}

function isCellCleared_(value: unknown): boolean {
  if (value === undefined || value === null) return true;
  if (typeof value !== "string") return false;
  return value.trim() === "";
}