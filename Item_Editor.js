// @ts-nocheck
/*****************************************************
 EEI - GRID BASED ITEM CONTROL
 Add / Update / Delete directly from items grid
*****************************************************/

function saveQuotationItemsGrid() {

    const sh = QFORM.SHEET();

    const qID = String(sh.getRange("B7").getDisplayValue()).trim();
    const revision = String(sh.getRange("E2").getDisplayValue()).trim();

    if (!qID) {
        SpreadsheetApp.getUi().alert("Load quotation first.");
        return;
    }

    if (!revision) {
        SpreadsheetApp.getUi().alert("Select revision first.");
        return;
    }

    const grid = sh.getRange("A22:L90").getValues();

    let added = 0;
    let updated = 0;
    let deleted = 0;

    grid.forEach(function (row) {

        const lineNo = Number(row[0]) || 0;
        const description = String(row[1] || "").trim();
        const transformerType = String(row[2] || "").trim();
        const powerKVA = String(row[3] || "").trim();
        const voltage = String(row[4] || "").trim();
        const quantity = Number(row[5]) || 0;
        const unitPrice = Number(row[6]) || 0;
        const deliveryTime = String(row[8] || "").trim();
        const warranty = String(row[9] || "").trim();
        const notes = String(row[10] || "").trim();
        const action = String(row[11] || "").trim();

        const isBlank =
            !lineNo &&
            !description &&
            !quantity &&
            !unitPrice;

        if (isBlank) return;

        if (lineNo && action === "Delete") {
            softDeleteQuotationItemByLine_(qID, revision, lineNo);
            deleted++;
            return;
        }

        if (!description) {
            throw new Error("Description is required at grid line: " + (lineNo || "new"));
        }

        if (quantity <= 0) {
            throw new Error("Quantity must be greater than zero at grid line: " + (lineNo || "new"));
        }

        if (unitPrice <= 0) {
            throw new Error("Unit price must be greater than zero at grid line: " + (lineNo || "new"));
        }

        const itemData = {
            qID: qID,
            revisionNo: revision,
            description: description,
            transformerType: transformerType,
            powerKVA: powerKVA,
            voltage: voltage,
            quantity: quantity,
            unitPrice: unitPrice,
            currency: String(sh.getRange("B11").getDisplayValue()).trim(),
            deliveryTime: deliveryTime,
            warranty: warranty,
            notes: notes
        };

        if (!lineNo) {
            addQuotationItem(itemData);
            added++;
            return;
        }

        const existingItem = getQuotationItemByLine_(qID, revision, lineNo);

        if (!existingItem) {
            addQuotationItem(itemData);
            added++;
            return;
        }

        updateQuotationItemByLine_(qID, revision, lineNo, itemData);
        updated++;

    });

    updateQuotationTotals(qID, revision);
    loadQuotationItemsToForm_(qID, revision);
    refreshQuotationKPIsFromForm();

    SpreadsheetApp.getUi().alert(
        "Items saved ✅\nAdded: " + added +
        "\nUpdated: " + updated +
        "\nDeleted: " + deleted
    );
}


function updateQuotationItemByLine_(qID, revision, lineNo, itemData) {

    const item = getQuotationItemByLine_(qID, revision, lineNo);

    if (!item) {
        throw new Error("Item line not found: " + lineNo);
    }

    if (isDuplicateQuotationItemExcludingLine_(qID, revision, lineNo, itemData)) {
        throw new Error("Duplicate item at line: " + lineNo);
    }

    const sheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_ITEMS);

    const qty = Number(itemData.quantity);
    const unitPrice = Number(itemData.unitPrice);
    const total = qty * unitPrice;

    sheet.getRange(item.row, 5, 1, 11).setValues([[
        itemData.description,
        itemData.transformerType,
        itemData.powerKVA,
        itemData.voltage,
        qty,
        unitPrice,
        total,
        itemData.currency,
        itemData.deliveryTime,
        itemData.warranty,
        itemData.notes
    ]]);

    sheet.getRange(item.row, 18).setValue(new Date());
    sheet.getRange(item.row, 19).setValue(getCurrentUserName());
    sheet.getRange(item.row, 21).setValue(getCurrentUserName());
    sheet.getRange(item.row, 22).setValue(new Date());
}


function softDeleteQuotationItemByLine_(qID, revision, lineNo) {

    const item = getQuotationItemByLine_(qID, revision, lineNo);

    if (!item) {
        throw new Error("Item line not found: " + lineNo);
    }

    const sheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_ITEMS);

    sheet.getRange(item.row, 20).setValue("Deleted");
    sheet.getRange(item.row, 23).setValue(getCurrentUserName());
    sheet.getRange(item.row, 24).setValue(new Date());
}


function getQuotationItemByLine_(qID, revision, lineNo) {

    const sheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_ITEMS);

    if (sheet.getLastRow() < 2) return null;

    const data = sheet
        .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
        .getValues();

    for (let i = 0; i < data.length; i++) {

        const row = data[i];
        const status = row[19] || "Active";

        if (
            row[1] === qID &&
            row[2] === revision &&
            Number(row[3]) === Number(lineNo) &&
            status !== "Deleted"
        ) {
            return {
                row: i + 2,
                data: row
            };
        }
    }

    return null;
}


function isDuplicateQuotationItemExcludingLine_(qID, revision, lineNo, itemData) {

    const sheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_ITEMS);

    if (sheet.getLastRow() < 2) return false;

    const data = sheet
        .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
        .getValues();

    const newType = String(itemData.transformerType || "")
        .toLowerCase()
        .replace(/\s+/g, "")
        .trim();

    if (!newType) return false;

    for (let i = 0; i < data.length; i++) {

        const row = data[i];
        const status = row[19] || "Active";

        if (status === "Deleted") continue;

        const rowQID = row[1];
        const rowRevision = row[2];
        const rowLineNo = Number(row[3]) || 0;
        const oldType = String(row[5] || "")
            .toLowerCase()
            .replace(/\s+/g, "")
            .trim();

        if (
            rowQID === qID &&
            rowRevision === revision &&
            rowLineNo !== Number(lineNo) &&
            oldType === newType
        ) {
            return true;
        }
    }

    return false;
}


function archiveQuotationFromForm() {

    const sh = QFORM.SHEET();
    const qID = String(sh.getRange("B7").getDisplayValue()).trim();

    if (!qID) {
        SpreadsheetApp.getUi().alert("Load quotation first.");
        return;
    }

    archiveQuotation_(qID);

    SpreadsheetApp.getUi().alert("Quotation archived ✅");
}