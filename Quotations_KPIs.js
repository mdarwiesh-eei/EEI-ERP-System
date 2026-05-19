function refreshQuotationKPIsFromForm() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!sheet) return;

  const qID = sheet.getRange("B7").getValue().toString().trim();
  const revisionNo = sheet.getRange("E2").getValue().toString().trim();
  const status = sheet.getRange("E7").getValue().toString().trim();
  const rfqDate = sheet.getRange("B9").getValue();

  let totalItems = 0;
  let totalQty = 0;

  if (qID && revisionNo) {

    const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.QUOTATION_ITEMS);

    if (itemsSheet && itemsSheet.getLastRow() >= 2) {

      const data = itemsSheet
        .getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn())
        .getValues();

      data.forEach(function(row) {

        const itemQID = row[1] ? row[1].toString().trim() : "";
        const itemRevision = row[2] ? row[2].toString().trim() : "";
        const qty = Number(row[8]) || 0;
        const itemStatus = row[19] ? row[19].toString().trim() : "Active";

        if (
          itemQID === qID &&
          itemRevision === revisionNo &&
          itemStatus !== "Deleted"
        ) {
          totalItems++;
          totalQty += qty;
        }

      });
    }
  }

  let rfqAge = "";

  if (rfqDate instanceof Date) {
    const today = new Date();
    const diffDays = Math.floor(
      (today - rfqDate) / (1000 * 60 * 60 * 24)
    );

    rfqAge = diffDays + " Days";
  }

  const lockStatus = canEditQuotation_(status) ? "Editable" : "Locked";

  sheet.getRange("I5:I9").setValues([
    [totalItems],
    [totalQty],
    [revisionNo],
    [rfqAge],
    [lockStatus]
  ]);
}