function refreshQuotationKPIsFromForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);
  if (!sheet) return;

  const headerValues = sheet.getRange("B6:B8").getValues();

  const qID = headerValues[0][0] ? headerValues[0][0].toString().trim() : "";
  const revisionNo = headerValues[1][0] ? headerValues[1][0].toString().trim() : "";
  const status = headerValues[2][0] ? headerValues[2][0].toString().trim() : "";
  const rfqDate = sheet.getRange("E4").getValue();

  let totalItems = 0;
  let totalQty = 0;

  if (qID && revisionNo) {
    const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.QUOTATION_ITEMS);

    if (itemsSheet && itemsSheet.getLastRow() >= 2) {
      const data = itemsSheet
        .getRange(2, 1, itemsSheet.getLastRow() - 1, 19)
        .getValues();

      data.forEach(function(row) {
        const itemQID = row[1] ? row[1].toString().trim() : "";
        const itemRevision = row[2] ? row[2].toString().trim() : "";

        if (itemQID === qID && itemRevision === revisionNo) {
          totalItems++;
          totalQty += Number(row[8]) || 0;
        }
      });
    }
  }

  let rfqAge = "";

  if (rfqDate instanceof Date) {
    const today = new Date();
    const diffDays = Math.floor((today - rfqDate) / (1000 * 60 * 60 * 24));
    rfqAge = diffDays + " Days";
  }

  const lockStatus = canEditQuotation_(status) ? "Editable" : "Locked";

  sheet.getRange("H4:H8").setValues([
    [totalItems],
    [totalQty],
    [revisionNo],
    [rfqAge],
    [lockStatus]
  ]);
}