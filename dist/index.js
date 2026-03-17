const ENTRY_SHEET_NAME = "Entry";
const INVENTORY_SHEET_NAME = "Inventory";
const TITLE_COLUMN = 4; // D
const CART_ID_COLUMN = 7; // G
const PICKER_COLUMN = 9; // I
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
    setPickerForMatches_(sheet, row, searchValue, matches);
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
function setPickerForMatches_(entrySheet, row, searchValue, matches) {
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
function applyPickerSelection_(entrySheet, row, pickerValue) {
    if (!pickerValue)
        return;
    const pickerCell = entrySheet.getRange(row, PICKER_COLUMN);
    const note = pickerCell.getNote();
    if (!note)
        return;
    const matchIndex = parsePickerIndex_(pickerValue);
    if (matchIndex < 0)
        return;
    let metadata;
    try {
        metadata = JSON.parse(note);
    }
    catch {
        return;
    }
    if (!metadata.searchValue)
        return;
    const inventorySheet = SpreadsheetApp.getActive().getSheetByName(INVENTORY_SHEET_NAME);
    if (!inventorySheet)
        return;
    const matches = findInventoryMatches_(inventorySheet, metadata.searchValue);
    if (matchIndex >= matches.length)
        return;
    applyMatchToEntryRow_(entrySheet, row, matches[matchIndex]);
    clearPicker_(entrySheet, row);
}
function parsePickerIndex_(value) {
    const indexMatch = value.match(/^\[(\d+)\]/);
    if (!indexMatch)
        return -1;
    const zeroBasedIndex = Number(indexMatch[1]) - 1;
    if (!Number.isInteger(zeroBasedIndex) || zeroBasedIndex < 0)
        return -1;
    return zeroBasedIndex;
}
function formatPickerOption_(index, match) {
    const maxTitleLength = 45;
    const shortTitle = match.title.length > maxTitleLength
        ? `${match.title.slice(0, maxTitleLength - 3)}...`
        : match.title;
    return `[${index + 1}] ${shortTitle} | ${match.cartId}`;
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
        setPickerForMatches_(entrySheet, targetRow, searchValue, matches);
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
