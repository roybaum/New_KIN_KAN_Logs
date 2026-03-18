const ENTRY_SHEET_NAME = "Entry";
const INVENTORY_SHEET_NAME = "Inventory";
const INVENTORY_IMPORT_SOURCE_SPREADSHEET_ID = "1QYBk6N_RZygLDPWV8BjVpF2azXBCyvGNRuz9XvpakPE";
const INVENTORY_IMPORT_SOURCE_SHEET_NAME = "Inventory";
const INVENTORY_IMPORT_DESTINATION_COLUMN_COUNT = 7;
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
const CART_ID_COLUMN = 7; // G
const PICKER_COLUMN = 9; // I
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu("KIN KAN Tools")
        .addItem("Sync Inventory", "syncInventoryFromExternalWorkbook")
        .addToUi();
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
function onEdit(e) {
    if (!e)
        return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== ENTRY_SHEET_NAME)
        return;
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
    const activeInventory = getActiveInventoryMatches_(inventorySheet);
    const matches = findMatchesInInventory_(activeInventory, searchValue);
    if (matches.length === 0)
        return;
    if (matches.length === 1) {
        applyMatchToEntryRow_(sheet, row, matches[0]);
        return;
    }
    setPickerForMatches_(sheet, row, matches);
}
function findInventoryMatches_(inventorySheet, searchValue) {
    const activeInventory = getActiveInventoryMatches_(inventorySheet);
    return findMatchesInInventory_(activeInventory, searchValue);
}
function getActiveInventoryMatches_(inventorySheet) {
    const lastRow = inventorySheet.getLastRow();
    if (lastRow < 2)
        return [];
    const data = inventorySheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const activeInventory = [];
    for (const item of data) {
        const match = {
            title: String(item[0]),
            isci: item[1],
            category: item[2],
            cartId: String(item[3]),
            length: item[4],
            startDate: item[5],
            endDate: item[6]
        };
        if (!isActiveToday_(today, match.startDate, match.endDate))
            continue;
        activeInventory.push(match);
    }
    return activeInventory;
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
    entrySheet.getRange(row, TITLE_COLUMN, 1, 5).setValues([[
            match.title,
            match.isci,
            match.category,
            match.cartId,
            match.length
        ]]);
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
    const activeInventory = getActiveInventoryMatches_(inventorySheet);
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
}
function processCartIdRangeEdit_(entrySheet, editedRange) {
    if (!rangeIncludesColumn_(editedRange, CART_ID_COLUMN))
        return;
    const cartIdOffset = CART_ID_COLUMN - editedRange.getColumn();
    const values = editedRange.getValues();
    const inventorySheet = SpreadsheetApp.getActive().getSheetByName(INVENTORY_SHEET_NAME);
    const activeInventory = inventorySheet ? getActiveInventoryMatches_(inventorySheet) : [];
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
        if (!inventorySheet)
            continue;
        const searchValue = String(cartIdValue).trim().toLowerCase();
        const matches = findMatchesInInventory_(activeInventory, searchValue);
        if (matches.length === 0)
            continue;
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
        mappedRow[mapping.destinationIndex] = sourceValue !== null && sourceValue !== void 0 ? sourceValue : "";
    }
    return mappedRow;
}
function writeInventoryRows_(destinationSheet, rows) {
    const existingRowCount = Math.max(destinationSheet.getLastRow() - 1, 0);
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
