// @ts-nocheck

function setupQuotationBackendSheets() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createOrResetSheet_(
    ss,
    CONFIG.SHEETS.QUOTATIONS,
    [
      "Q_ID",
      "Customer_ID",
      "Customer_Name",
      "Project_Name",
      "Current_Revision",
      "Status",
      "RFQ_Date",
      "Created_Date",
      "Created_By",
      "Assigned_To",
      "Currency",
      "Sub_Total",
      "Discount_Percent",
      "VAT_Percent",
      "Grand_Total",
      "Quotation_Folder_Link",
      "Customer_RFQ_Link",
      "Notes",
      "Last_Updated",
      "Last_Updated_By"
    ]
  );

  createOrResetSheet_(
    ss,
    CONFIG.SHEETS.QUOTATION_REVISIONS,
    [
      "Revision_ID",
      "Q_ID",
      "Revision_No",
      "Revision_Status",
      "Revision_Reason",
      "Created_Date",
      "Created_By",
      "Technical_Changes",
      "Commercial_Changes",
      "Sub_Total",
      "Discount_Percent",
      "VAT_Percent",
      "Grand_Total",
      "PDF_Link",
      "Approved_By",
      "Approval_Date",
      "Approval_Notes",
      "Sent_Date",
      "Is_Current",
      "Notes"
    ]
  );

  createOrResetSheet_(
    ss,
    CONFIG.SHEETS.QUOTATION_ITEMS,
    [
      "Item_ID",
      "Q_ID",
      "Revision_No",
      "Line_No",
      "Item_Description",
      "Transformer_Type",
      "Power_kVA",
      "Voltage",
      "Quantity",
      "Unit_Price",
      "Total_Price",
      "Currency",
      "Delivery_Time",
      "Warranty",
      "Notes",
      "Created_Date",
      "Created_By",
      "Last_Updated",
      "Last_Updated_By"
    ]
  );

  createOrResetSheet_(
    ss,
    CONFIG.SHEETS.QUOTATION_STATUS_LOG,
    [
      "Log_ID",
      "Q_ID",
      "Revision_No",
      "Old_Status",
      "New_Status",
      "Changed_By",
      "Changed_Date",
      "Reason",
      "Notes"
    ]
  );

  applyQuotationBackendFormatting_();

  SpreadsheetApp.getUi().alert(
    "Quotation backend sheets created successfully ✅"
  );
}



function createOrResetSheet_(ss, sheetName, headers) {

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {

    sheet = ss.insertSheet(sheetName);

  } else {

    sheet.clear();

  }

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  sheet.setFrozenRows(1);

  sheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#1F3A5F")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");

  sheet.autoResizeColumns(1, headers.length);
}



function applyQuotationBackendFormatting_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = [
    CONFIG.SHEETS.QUOTATIONS,
    CONFIG.SHEETS.QUOTATION_REVISIONS,
    CONFIG.SHEETS.QUOTATION_ITEMS,
    CONFIG.SHEETS.QUOTATION_STATUS_LOG
  ];

  sheets.forEach(function(sheetName) {

    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) return;

    const lastCol = sheet.getLastColumn();

    sheet
      .getRange(1, 1, 1, lastCol)
      .setWrap(true);

    sheet
      .getDataRange()
      .setVerticalAlignment("middle");

    sheet
      .getRange(1, 1, sheet.getMaxRows(), lastCol)
      .createFilter();

  });




  //FORMATING

  function createQuotationSheet_PRO_(ss, sheetName, headers) {

    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
      sheet.clearFormats();
    }

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]);

    sheet.setFrozenRows(1);

    const headerRange =
      sheet.getRange(1, 1, 1, headers.length);

    headerRange
      .setBackground("#1F3A5F")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    sheet.setRowHeight(1, 40);

    sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length)
      .setBackground("#F8FAFC")
      .setVerticalAlignment("middle");

    sheet.getRange(1, 1, sheet.getMaxRows(), headers.length)
      .setFontFamily("Arial")
      .setFontSize(10);

    sheet.autoResizeColumns(1, headers.length);

    for (let c = 1; c <= headers.length; c++) {
      sheet.setColumnWidth(c, Math.max(sheet.getColumnWidth(c), 130));
    }

    if (!sheet.getFilter()) {
      sheet.getRange(1, 1, sheet.getMaxRows(), headers.length)
        .createFilter();
    }

    protectHeader_PRO_(sheet);
  }



  function protectHeader_PRO_(sheet) {

    const protection =
      sheet.getRange(1, 1, 1, sheet.getLastColumn())
        .protect();

    protection.setDescription(
      "Protected ERP Header"
    );

    protection.setWarningOnly(true);
  }

  function applyQuotationDataValidation_PRO_() {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const statusList = [
      CONFIG.QUOTATION_STATUS.DRAFT,
      CONFIG.QUOTATION_STATUS.UNDER_REVIEW,
      CONFIG.QUOTATION_STATUS.APPROVED,
      CONFIG.QUOTATION_STATUS.SENT,
      CONFIG.QUOTATION_STATUS.NEGOTIATION,
      CONFIG.QUOTATION_STATUS.REVISED,
      CONFIG.QUOTATION_STATUS.WON,
      CONFIG.QUOTATION_STATUS.LOST,
      CONFIG.QUOTATION_STATUS.CANCELLED,
      CONFIG.QUOTATION_STATUS.SUPERSEDED
    ];

    const currencyList = [
      CONFIG.CURRENCY.EGP,
      CONFIG.CURRENCY.USD,
      CONFIG.CURRENCY.EUR
    ];

    const statusRule =
      SpreadsheetApp.newDataValidation()
        .requireValueInList(statusList, true)
        .setAllowInvalid(false)
        .build();

    const currencyRule =
      SpreadsheetApp.newDataValidation()
        .requireValueInList(currencyList, true)
        .setAllowInvalid(false)
        .build();

    const qSheet =
      ss.getSheetByName(CONFIG.SHEETS.QUOTATIONS);

    const revSheet =
      ss.getSheetByName(CONFIG.SHEETS.QUOTATION_REVISIONS);

    const itemsSheet =
      ss.getSheetByName(CONFIG.SHEETS.QUOTATION_ITEMS);

    const logSheet =
      ss.getSheetByName(CONFIG.SHEETS.QUOTATION_STATUS_LOG);

    if (qSheet) {

      qSheet.getRange("F2:F")
        .setDataValidation(statusRule);

      qSheet.getRange("K2:K")
        .setDataValidation(currencyRule);

      qSheet.getRange("G2:H")
        .setNumberFormat("dd/mm/yyyy");

      qSheet.getRange("L2:L")
        .setNumberFormat("#,##0.00");

      qSheet.getRange("M2:N")
        .setNumberFormat("0.00%");

      qSheet.getRange("O2:O")
        .setNumberFormat("#,##0.00");

      qSheet.getRange("S2:S")
        .setNumberFormat("dd/mm/yyyy hh:mm");
    }

    if (revSheet) {

      revSheet.getRange("D2:D")
        .setDataValidation(statusRule);

      revSheet.getRange("F2:F")
        .setNumberFormat("dd/mm/yyyy");

      revSheet.getRange("J2:J")
        .setNumberFormat("#,##0.00");

      revSheet.getRange("K2:L")
        .setNumberFormat("0.00%");

      revSheet.getRange("M2:M")
        .setNumberFormat("#,##0.00");

      revSheet.getRange("P2:P")
        .setNumberFormat("dd/mm/yyyy");

      revSheet.getRange("R2:R")
        .setNumberFormat("dd/mm/yyyy");
    }

    if (itemsSheet) {

      itemsSheet.getRange("L2:L")
        .setDataValidation(currencyRule);

      itemsSheet.getRange("I2:K")
        .setNumberFormat("#,##0.00");

      itemsSheet.getRange("P2:P")
        .setNumberFormat("dd/mm/yyyy");

      itemsSheet.getRange("R2:R")
        .setNumberFormat("dd/mm/yyyy hh:mm");
    }

    if (logSheet) {

      logSheet.getRange("G2:G")
        .setNumberFormat("dd/mm/yyyy hh:mm");
    }

    applyStatusConditionalFormatting_PRO_();
  }





  function applyStatusConditionalFormatting_PRO_() {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const targetSheets = [
      {
        sheet: ss.getSheetByName(CONFIG.SHEETS.QUOTATIONS),
        range: "F2:F"
      },
      {
        sheet: ss.getSheetByName(CONFIG.SHEETS.QUOTATION_REVISIONS),
        range: "D2:D"
      }
    ];

    targetSheets.forEach(function(obj) {

      if (!obj.sheet) return;

      const range =
        obj.sheet.getRange(obj.range);

      const rules = [];

      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Draft")
          .setBackground("#E5E7EB")
          .setRanges([range])
          .build()
      );

      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Under Review")
          .setBackground("#FDE68A")
          .setRanges([range])
          .build()
      );

      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Approved")
          .setBackground("#BFDBFE")
          .setRanges([range])
          .build()
      );

      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Sent")
          .setBackground("#BBF7D0")
          .setRanges([range])
          .build()
      );

      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Negotiation")
          .setBackground("#DDD6FE")
          .setRanges([range])
          .build()
      );

      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Won")
          .setBackground("#86EFAC")
          .setRanges([range])
          .build()
      );

      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Lost")
          .setBackground("#FCA5A5")
          .setRanges([range])
          .build()
      );

      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Cancelled")
          .setBackground("#D1D5DB")
          .setRanges([range])
          .build()
      );

      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("Superseded")
          .setBackground("#CBD5E1")
          .setRanges([range])
          .build()
      );

      obj.sheet.setConditionalFormatRules(rules);
    });
  }


}



// @ts-nocheck

function setupQuotationTermsSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const headers = [
    "Terms_ID",
    "Q_ID",
    "Revision_No",
    "Currency",
    "Sub_Total",
    "Discount_Percent",
    "VAT_Percent",
    "Advance_Percent",
    "Advance_Amount",
    "Remaining_Amount",
    "Grand_Total",
    "Delivery_Terms",
    "Payment_Terms",
    "Validity_Days",
    "Created_Date",
    "Created_By",
    "Last_Updated",
    "Last_Updated_By",
    "Notes"
  ];

  let sheet = ss.getSheetByName(CONFIG.SHEETS.QUOTATION_TERMS);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.QUOTATION_TERMS);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  formatQuotationTermsSheet_(sheet, headers.length);
  applyQuotationTermsValidation_(sheet);

  SpreadsheetApp.getUi().alert("Quotation_Terms sheet is ready ✅");
}


function formatQuotationTermsSheet_(sheet, lastCol) {

  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, 1, lastCol)
    .setBackground("#1F3A5F")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  sheet.setRowHeight(1, 42);

  sheet.getRange(1, 1, sheet.getMaxRows(), lastCol)
    .setFontFamily("Arial")
    .setFontSize(10)
    .setVerticalAlignment("middle");

  sheet.getRange(2, 1, sheet.getMaxRows() - 1, lastCol)
    .setBackground("#F8FAFC");

  sheet.autoResizeColumns(1, lastCol);

  for (let c = 1; c <= lastCol; c++) {
    sheet.setColumnWidth(c, Math.max(sheet.getColumnWidth(c), 130));
  }

  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, sheet.getMaxRows(), lastCol).createFilter();
  }

  sheet.getRange("E2:E").setNumberFormat("#,##0.00");
  sheet.getRange("F2:H").setNumberFormat("0.00%");
  sheet.getRange("I2:K").setNumberFormat("#,##0.00");
  sheet.getRange("N2:N").setNumberFormat("0");
  sheet.getRange("O2:O").setNumberFormat("dd/mm/yyyy hh:mm");
  sheet.getRange("Q2:Q").setNumberFormat("dd/mm/yyyy hh:mm");

  const protection = sheet.getRange(1, 1, 1, lastCol).protect();
  protection.setDescription("Protected ERP Header - Quotation Terms");
  protection.setWarningOnly(true);
}




function applyQuotationTermsValidation_(sheet) {

  const currencyList = [
    CONFIG.CURRENCY.EGP,
    CONFIG.CURRENCY.USD,
    CONFIG.CURRENCY.EUR
  ];

  const currencyRule =
    SpreadsheetApp.newDataValidation()
      .requireValueInList(currencyList, true)
      .setAllowInvalid(false)
      .build();

  sheet.getRange("D2:D").setDataValidation(currencyRule);
}



function testQuotationWorkflow() {

  submitQuotationForReview("Q-0001");

  approveQuotation(
    "Q-0001",
    "Approved by management"
  );

  sendQuotation(
    "Q-0001",
    "Sent by email"
  );

  moveQuotationToNegotiation(
    "Q-0001",
    "Customer requested discount"
  );

  markQuotationWon(
    "Q-0001",
    "PO received"
  );

}


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



