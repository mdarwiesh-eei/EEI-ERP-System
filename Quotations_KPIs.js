function setupQuotationKPISection() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Quotation_Form sheet not found.");
    return;
  }

  sheet.getRange("J10:L10")
    .merge()
    .setValue("QUOTATION KPIs")
    .setBackground("#A61C00")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  const kpis = [
    ["J11", "Total Items", "K11"],
    ["J12", "Total Qty", "K12"],
    ["J13", "Current Revision", "K13"],
    ["J14", "RFQ Age", "K14"],
    ["J15", "Status Lock", "K15"]
  ];

  kpis.forEach(function(row) {

    sheet.getRange(row[0])
      .setValue(row[1])
      .setBackground("#D6EAF8")
      .setFontWeight("bold")
      .setFontColor("#1F3A5F")
      .setBorder(true, true, true, true, true, true);

    sheet.getRange(row[2])
      .setBackground("#E5E7E9")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBorder(true, true, true, true, true, true);

  });

  SpreadsheetApp.getUi().alert("Quotation KPI Section Created ✅");
}






function refreshQuotationKPIsFromForm() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!sheet) return;

  const qID =
    sheet.getRange("B6").getValue();

  const revisionNo =
    sheet.getRange("B7").getValue();

  const status =
    sheet.getRange("B8").getValue();

  const rfqDate =
    sheet.getRange("E4").getValue();

  let totalItems = 0;
  let totalQty = 0;

  if (qID && revisionNo) {

    const itemsSheet =
      getRequiredSheet_(CONFIG.SHEETS.QUOTATION_ITEMS);

    if (itemsSheet.getLastRow() >= 2) {

      const data =
        itemsSheet
          .getRange(
            2,
            1,
            itemsSheet.getLastRow() - 1,
            11
          )
          .getValues();

      data.forEach(function(row) {

        if (
          row[1] === qID &&
          row[2] === revisionNo
        ) {

          totalItems++;

          totalQty += Number(row[8]) || 0;

        }

      });
    }
  }

  let rfqAge = "";

  if (rfqDate instanceof Date) {

    const today = new Date();

    const diffDays =
      Math.floor(
        (today - rfqDate) /
        (1000 * 60 * 60 * 24)
      );

    rfqAge =
      diffDays + " Days";
  }

  const lockStatus =
    canEditQuotation_(status)
      ? "Editable"
      : "Locked";

  sheet.getRange("H4").setValue(totalItems);
  sheet.getRange("H5").setValue(totalQty);
  sheet.getRange("H6").setValue(revisionNo || "");
  sheet.getRange("H7").setValue(rfqAge);
  sheet.getRange("H8").setValue(lockStatus);

  applyQuotationStatusColor_();
}




function applyQuotationStatusColor_() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!sheet) return;

  const status =
    sheet.getRange("B8").getValue();

  let color = "#E5E7E9";

  if (status === "Draft") color = "#E5E7EB";
  if (status === "Under Review") color = "#FDE68A";
  if (status === "Approved") color = "#BFDBFE";
  if (status === "Sent") color = "#BBF7D0";
  if (status === "Negotiation") color = "#DDD6FE";
  if (status === "Revised") color = "#FDEBD0";
  if (status === "Won") color = "#86EFAC";
  if (status === "Lost") color = "#FCA5A5";
  if (status === "Cancelled") color = "#D1D5DB";
  if (status === "Superseded") color = "#CBD5E1";

  sheet.getRange("B8")
    .setBackground(color)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.getRange("H8")
    .setBackground(
      canEditQuotation_(status)
        ? "#D5F5E3"
        : "#FADBD8"
    )
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}









