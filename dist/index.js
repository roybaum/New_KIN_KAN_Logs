function onEdit(e) {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== "Entry")
        return;
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < 2)
        return;
    // Only respond to Title (D) or Cart ID (G)
    if (col !== 4 && col !== 7)
        return;
    const searchValue = String(e.value || "").toLowerCase();
    if (!searchValue)
        return;
    const ss = SpreadsheetApp.getActive();
    const inventorySheet = ss.getSheetByName("Inventory");
    if (!inventorySheet)
        return;
    const data = inventorySheet
        .getRange(2, 1, inventorySheet.getLastRow() - 1, 7)
        .getValues();
    const now = new Date();
    for (const item of data) {
        const title = String(item[0]);
        const isci = item[1];
        const category = item[2];
        const cartId = String(item[3]);
        const length = item[4];
        const startDate = item[5];
        const endDate = item[6];
        const titleMatch = title.toLowerCase().includes(searchValue);
        const cartMatch = cartId.includes(searchValue);
        if (!titleMatch && !cartMatch)
            continue;
        // Validate dates
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        if (startDate) {
            const start = new Date(startDate);
            start.setHours(0, 0, 0, 0);
            if (today < start)
                continue;
        }
        if (endDate) {
            const end = new Date(endDate);
            end.setHours(0, 0, 0, 0);
            if (today > end)
                continue;
        }
        sheet.getRange(row, 4, 1, 5).setValues([[
                title,
                isci,
                category,
                cartId,
                length
            ]]);
        return;
    }
}
