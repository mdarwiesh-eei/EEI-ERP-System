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

  kpis.forEach(function (row) {

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
    sheet.getRange("B6").getValue().toString().trim();

  const revisionNo =
    sheet.getRange("B7").getValue().toString().trim();

  const status =
    sheet.getRange("B8").getValue().toString().trim();

  const rfqDate =
    sheet.getRange("E4").getValue();

  let totalItems = 0;
  let totalQty = 0;

  if (qID && revisionNo) {

    const itemsSheet =
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName(CONFIG.SHEETS.QUOTATION_ITEMS);

    if (itemsSheet && itemsSheet.getLastRow() >= 2) {

      const data =
        itemsSheet
          .getRange(
            2,
            1,
            itemsSheet.getLastRow() - 1,
            19
          )
          .getValues();

      data.forEach(function (row) {

        const itemQID =
          row[1] ? row[1].toString().trim() : "";

        const itemRevision =
          row[2] ? row[2].toString().trim() : "";

        const qty =
          Number(row[8]) || 0;

        if (
          itemQID === qID &&
          itemRevision === revisionNo
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

    const diffDays =
      Math.floor(
        (today - rfqDate) /
        (1000 * 60 * 60 * 24)
      );

    rfqAge =
      diffDays + " Days";

  }

  const lockedStatuses = [
    "Sent",
    "Won",
    "Lost",
    "Cancelled",
    "Superseded"
  ];

  const lockStatus =
    lockedStatuses.includes(status)
      ? "Locked"
      : "Editable";

  sheet.getRange("H4").setValue(totalItems);
  sheet.getRange("H5").setValue(totalQty);
  sheet.getRange("H6").setValue(revisionNo);
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
    sheet.getRange("B8").getValue().toString().trim();

  let statusColor = "#E5E7E9";

  if (status === "Draft") statusColor = "#E5E7EB";
  if (status === "Under Review") statusColor = "#FDE68A";
  if (status === "Approved") statusColor = "#BFDBFE";
  if (status === "Sent") statusColor = "#BBF7D0";
  if (status === "Negotiation") statusColor = "#DDD6FE";
  if (status === "Revised") statusColor = "#FDEBD0";
  if (status === "Won") statusColor = "#86EFAC";
  if (status === "Lost") statusColor = "#FCA5A5";
  if (status === "Cancelled") statusColor = "#D1D5DB";
  if (status === "Superseded") statusColor = "#CBD5E1";

  sheet.getRange("B8")
    .setBackground(statusColor)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  const lockStatus =
    sheet.getRange("H8").getValue().toString().trim();

  sheet.getRange("H8")
    .setBackground(
      lockStatus === "Editable"
        ? "#D5F5E3"
        : "#FADBD8"
    )
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}









