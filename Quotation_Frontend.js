function createQuotationFromForm() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert("Quotation_Form sheet not found.");
    return;
  }

  try {

    const customerValue = sheet.getRange("B4").getValue().toString().trim();
    const projectName = sheet.getRange("B5").getValue().toString().trim();

    if (!customerValue || !customerValue.includes("|")) {
      ui.alert("Please select a customer.");
      return;
    }

    if (!projectName) {
      ui.alert("Please enter project name.");
      return;
    }

    const customerID = customerValue.split("|")[0].trim();

    const result = createQuotation({

      customerID: customerID,
      projectName: projectName,
      rfqDate: sheet.getRange("E4").getValue(),
      assignedTo: sheet.getRange("E5").getValue(),
      currency: sheet.getRange("E6").getValue() || CONFIG.CURRENCY.EGP,
      customerRFQLink: sheet.getRange("E7").getValue(),
      notes: sheet.getRange("E8").getValue()

    });

    if (!result || !result.success) return;

    sheet.getRange("B6").setValue(result.qID);
    sheet.getRange("B7").setValue(result.revisionNo);
    sheet.getRange("B8").setValue(CONFIG.QUOTATION_STATUS.DRAFT);

    saveCommercialTermsFromForm();
    refreshQuotationForm();

    ui.alert("Quotation Created Successfully ✅");

  } catch (err) {

    ui.alert("Create Quotation Error: " + err.message);
    Logger.log(err);

  }
}




function saveCommercialTermsFromForm() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!sheet) return;

  const qID = sheet.getRange("B6").getValue().toString().trim();
  const revisionNo = sheet.getRange("B7").getValue().toString().trim();

  if (!qID || !revisionNo) return;

  updateQuotationCommercialTerms({

    qID: qID,
    revisionNo: revisionNo,
    currency: sheet.getRange("E6").getValue() || CONFIG.CURRENCY.EGP,
    discountPercent: sheet.getRange("B11").getValue(),
    vatPercent: sheet.getRange("B12").getValue(),
    advancePercent: sheet.getRange("B13").getValue(),
    deliveryTerms: sheet.getRange("B14").getValue(),
    paymentTerms: sheet.getRange("E11").getValue(),
    validityDays: sheet.getRange("E12").getValue(),
    notes: "Updated from Quotation_Form"

  });
}



function refreshQuotationForm() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!sheet) return;

  const qID = sheet.getRange("B6").getValue().toString().trim();
  const revisionNo = sheet.getRange("B7").getValue().toString().trim();

  if (!qID || !revisionNo) return;

  updateQuotationTotals(qID, revisionNo);

  const terms = getQuotationTerms_(qID, revisionNo);

  if (terms) {

    const vatAmount =
      terms.grandTotal -
      terms.subTotal +
      (terms.subTotal * terms.discountPercent / 100);

    sheet.getRange("H11").setValue(terms.subTotal);
    sheet.getRange("H12").setValue(vatAmount);
    sheet.getRange("H13").setValue(terms.advanceAmount);
    sheet.getRange("H14").setValue(terms.grandTotal);
    sheet.getRange("H15").setValue(terms.remainingAmount);

  }

  if (typeof refreshQuotationKPIsFromForm === "function") {
    refreshQuotationKPIsFromForm();
  }

  if (typeof applyQuotationStatusColor_ === "function") {
    applyQuotationStatusColor_();
  }
}



function repairQuotationFormButtons() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!sheet) return;

  const buttonCells = [
    "K4",
    "K5",
    "K6",
    "K7",
    "K8",
    "K9",
    "K10",
    "K11",
    "K12",
    "K13",
    "K14",
    "K15"
  ];

  buttonCells.forEach(function (cell) {

    sheet.getRange(cell)
      .clearContent()
      .clearDataValidations()
      .insertCheckboxes()
      .setValue(false)
      .setHorizontalAlignment("center");

  });

  const labels = [
    ["L4", "➕ Create Quotation"],
    ["L5", "💾 Add Items"],
    ["L6", "📤 Submit Review"],
    ["L7", "✅ Approve"],
    ["L8", "📧 Send"],
    ["L9", "🔁 Create Revision"],
    ["L10", "🏆 Mark Won"],
    ["L11", "❌ Mark Lost"],
    ["L12", "🚫 Cancel"],
    ["L13", "🔄 Refresh"],
    ["L14", "🧹 Clear Form"],
    ["L15", "📂 Open Folder"]
  ];

  labels.forEach(function (row) {

    sheet.getRange(row[0])
      .setValue(row[1])
      .setFontWeight("bold");

  });

  refreshQuotationKPIsFromForm();

  SpreadsheetApp.getUi().alert("Quotation buttons repaired ✅");
}





function addQuotationItemsFromGrid() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert("Quotation_Form sheet not found.");
    return;
  }

  const qID =
    sheet.getRange("B6").getValue().toString().trim();

  const revisionNo =
    sheet.getRange("B7").getValue().toString().trim();

  if (!qID || !revisionNo) {
    ui.alert("Create or load quotation first.");
    return;
  }

  const rows =
    sheet.getRange("A19:L35").getValues();

  let addedCount = 0;

  rows.forEach(function (row) {

    const description = row[1];
    const transformerType = row[2];
    const powerKVA = row[3];
    const voltage = row[4];
    const qty = row[5];
    const unitPrice = row[6];
    const delivery = row[8];
    const warranty = row[9];
    const notes = row[10];

    if (!description) return;

    addQuotationItem({

      qID: qID,
      revisionNo: revisionNo,

      description: description,
      transformerType: transformerType,
      powerKVA: powerKVA,
      voltage: voltage,

      quantity: qty,
      unitPrice: unitPrice,

      currency:
        sheet.getRange("E6").getValue()
        || CONFIG.CURRENCY.EGP,

      deliveryTime: delivery,
      warranty: warranty,
      notes: notes

    });

    addedCount++;

  });

  refreshQuotationForm();

  ui.alert(
    addedCount + " item(s) added successfully ✅"
  );
}



function createQuotationRevisionFromForm() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert("Quotation_Form sheet not found.");
    return;
  }

  const qID =
    sheet.getRange("B6").getValue().toString().trim();

  if (!qID) {
    ui.alert("Load quotation first.");
    return;
  }

  const response =
    ui.prompt(
      "Create Revision",
      "Please enter revision reason:",
      ui.ButtonSet.OK_CANCEL
    );

  if (
    response.getSelectedButton()
    !== ui.Button.OK
  ) {
    return;
  }

  const reason =
    response.getResponseText().trim();

  if (!reason) {
    ui.alert("Revision reason is required.");
    return;
  }

  if (typeof createQuotationRevision !== "function") {
    ui.alert("createQuotationRevision backend function is not available yet.");
    return;
  }

  const result =
    createQuotationRevision(
      qID,
      reason
    );

  if (!result || !result.success) return;

  sheet.getRange("B7")
    .setValue(result.revisionNo);

  sheet.getRange("B8")
    .setValue(CONFIG.QUOTATION_STATUS.REVISED);

  buildRevisionSelectorForForm(qID);

  refreshQuotationForm();

  ui.alert("Revision created successfully ✅");
}





function clearQuotationForm() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!sheet) return;

  const clearRanges = [

    "B4:B8",
    "E4:E8",

    "B11:B14",
    "E11:E12",

    "H4:H8",
    "H11:H15",

    "A20:L35"

  ];

  clearRanges.forEach(function (range) {

    sheet.getRange(range)
      .clearContent();

  });

  const buttonCells = [
    "K4",
    "K5",
    "K6",
    "K7",
    "K8",
    "K9",
    "K10",
    "K11",
    "K12",
    "K13",
    "K14",
    "K15"
  ];

  buttonCells.forEach(function (cell) {

    sheet.getRange(cell)
      .setValue(false);

  });

  if (typeof buildQuotationCustomerSelector === "function") {
    buildQuotationCustomerSelector();
  }

  if (typeof buildQuotationCurrencySelector === "function") {
    buildQuotationCurrencySelector();
  }

  sheet.getRange("B4").activate();

  SpreadsheetApp
    .getUi()
    .alert("Quotation Form cleared ✅");
}





function openQuotationFolderFromForm() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert("Quotation_Form sheet not found.");
    return;
  }

  const qID =
    sheet.getRange("B6").getValue().toString().trim();

  if (!qID) {
    ui.alert("Load quotation first.");
    return;
  }

  const quotation =
    getQuotationById_(qID);

  if (!quotation || !quotation.folderLink) {
    ui.alert("Quotation folder link not found.");
    return;
  }

  const html =
    HtmlService
      .createHtmlOutput(
        '<script>' +
        'window.open("' + quotation.folderLink + '", "_blank");' +
        'google.script.host.close();' +
        '</script>'
      )
      .setWidth(100)
      .setHeight(50);

  ui.showModalDialog(
    html,
    "Opening Folder..."
  );
}





function openRFQFolderFromForm() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  const ui =
    SpreadsheetApp.getUi();

  if (!sheet) return;

  const qID =
    sheet.getRange("B6")
      .getValue()
      .toString()
      .trim();

  const revisionNo =
    sheet.getRange("B7")
      .getValue()
      .toString()
      .trim();

  if (!qID || !revisionNo) {
    ui.alert("Load quotation first.");
    return;
  }

  const quotation =
    getQuotationById_(qID);

  if (!quotation || !quotation.folderLink) {
    ui.alert("Quotation folder not found.");
    return;
  }

  const quotationFolderId =
    extractIdFromUrl(quotation.folderLink);

  const quotationFolder =
    DriveApp.getFolderById(quotationFolderId);

  const revisionFolder =
    findChildFolderByName_(
      quotationFolder,
      revisionNo
    );

  if (!revisionFolder) {
    ui.alert("Revision folder not found.");
    return;
  }

  const rfqFolder =
    findChildFolderByName_(
      revisionFolder,
      "01 RFQ"
    );

  if (!rfqFolder) {
    ui.alert("RFQ folder not found.");
    return;
  }

  openUrl_(rfqFolder.getUrl(), "Opening RFQ Folder...");
}






function loadQuotationToForm() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  const ui =
    SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert("Quotation_Form sheet not found.");
    return;
  }

  const selected =
    sheet.getRange("B2")
      .getDisplayValue()
      .toString()
      .trim();

  if (!selected || !selected.includes("|")) {
    ui.alert("Please select a valid quotation.");
    return;
  }

  const qID =
    selected.split("|")[0]
      .toString()
      .trim();

  const quotation =
    getQuotationById_(qID);

  if (!quotation) {
    ui.alert("Quotation not found.");
    return;
  }

  const revisionNo =
    quotation.currentRevision;

  buildRevisionSelectorForForm(qID);

  const terms =
    getQuotationTerms_(
      qID,
      revisionNo
    );

  // ===============================
  // HEADER
  // ===============================

  sheet.getRange("B4").setValue(
    quotation.customerID +
    " | " +
    quotation.customerName
  );

  sheet.getRange("B5")
    .setValue(quotation.projectName);

  sheet.getRange("B6")
    .setValue(quotation.qID);

  sheet.getRange("B7")
    .setValue(quotation.currentRevision);

  sheet.getRange("B8")
    .setValue(quotation.status);

  sheet.getRange("E4")
    .setValue(quotation.rfqDate || "");

  sheet.getRange("E5")
    .setValue(quotation.assignedTo || "");

  sheet.getRange("E6")
    .setValue(
      quotation.currency ||
      CONFIG.CURRENCY.EGP
    );

  sheet.getRange("E7")
    .setValue(quotation.customerRFQLink || "");

  sheet.getRange("E8")
    .setValue(quotation.notes || "");

  // ===============================
  // COMMERCIAL TERMS
  // ===============================

  if (terms) {

    sheet.getRange("B11")
      .setValue(terms.discountPercent / 100);

    sheet.getRange("B12")
      .setValue(terms.vatPercent / 100);

    sheet.getRange("B13")
      .setValue(terms.advancePercent / 100);

    sheet.getRange("B14")
      .setValue(terms.deliveryTerms);

    sheet.getRange("E11")
      .setValue(terms.paymentTerms);

    sheet.getRange("E12")
      .setValue(terms.validityDays);
  }

  // ===============================
  // ITEMS
  // ===============================

  loadQuotationItemsToGrid_(
    qID,
    revisionNo
  );

  // ===============================
  // DIRECT TOTALS + KPIs UPDATE
  // ===============================

  updateQuotationTotals(qID, revisionNo);

  const updatedTerms =
    getQuotationTerms_(
      qID,
      revisionNo
    );

  if (updatedTerms) {

    const vatAmount =
      updatedTerms.grandTotal -
      updatedTerms.subTotal +
      (
        updatedTerms.subTotal *
        updatedTerms.discountPercent /
        100
      );

    sheet.getRange("H11")
      .setValue(updatedTerms.subTotal);

    sheet.getRange("H12")
      .setValue(vatAmount);

    sheet.getRange("H13")
      .setValue(updatedTerms.advanceAmount);

    sheet.getRange("H14")
      .setValue(updatedTerms.grandTotal);

    sheet.getRange("H15")
      .setValue(updatedTerms.remainingAmount);
  }

  // ===============================
  // DIRECT KPIs UPDATE
  // ===============================

  let totalItems = 0;
  let totalQty = 0;

  const itemsSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.QUOTATION_ITEMS
      );

  if (
    itemsSheet &&
    itemsSheet.getLastRow() >= 2
  ) {

    const itemsData =
      itemsSheet
        .getRange(
          2,
          1,
          itemsSheet.getLastRow() - 1,
          19
        )
        .getValues();

    itemsData.forEach(function (row) {

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

  let rfqAge = "";

  const rfqDate =
    sheet.getRange("E4").getValue();

  if (rfqDate instanceof Date) {

    const today =
      new Date();

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
    lockedStatuses.includes(
      quotation.status
    )
      ? "Locked"
      : "Editable";

  sheet.getRange("H4").setValue(totalItems);
  sheet.getRange("H5").setValue(totalQty);
  sheet.getRange("H6").setValue(revisionNo);
  sheet.getRange("H7").setValue(rfqAge);
  sheet.getRange("H8").setValue(lockStatus);

  if (typeof applyQuotationStatusColor_ === "function") {
    applyQuotationStatusColor_();
  }

  if (typeof applyQuotationReadOnlyUI_ === "function") {
    applyQuotationReadOnlyUI_();
  }

  SpreadsheetApp.flush();

  ui.alert("Quotation loaded ✅");


  // ===============================
  // REVISION SELECTOR
  // ===============================

  if (
    typeof buildRevisionSelectorForForm
    === "function"
  ) {
    buildRevisionSelectorForForm(qID);
  }

  // ===============================
  // REFRESH
  // ===============================

  SpreadsheetApp.flush();


  applyQuotationStatusColor_();

  applyQuotationReadOnlyUI_();

  SpreadsheetApp.flush();

  ui.alert("Quotation loaded ✅");
}




function loadQuotationItemsToGrid_(
  qID,
  revisionNo
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  const itemsSheet =
    getRequiredSheet_(CONFIG.SHEETS.QUOTATION_ITEMS);

  sheet.getRange("A20:L35").clearContent();

  if (itemsSheet.getLastRow() < 2) return;

  const data =
    itemsSheet
      .getRange(
        2,
        1,
        itemsSheet.getLastRow() - 1,
        20
      )
      .getValues();

  const rows = [];

  data.forEach(function (row) {

    if (
      row[1] === qID &&
      row[2] === revisionNo
    ) {

      rows.push([
        row[3],
        row[4],
        row[5],
        row[6],
        row[7],
        row[8],
        row[9],
        row[10],
        row[12],
        row[13],
        row[14],
        ""
      ]);

    }

    applyQuotationGridBorders_()

  });

  if (rows.length) {

    sheet
      .getRange(
        20,
        1,
        rows.length,
        12
      )
      .setValues(rows);
  }
}







function applyRFQDatePicker() {

  const form =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!form) return;

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build();

  form.getRange("E4")
    .clearDataValidations()
    .setDataValidation(rule)
    .setNumberFormat("dd/mm/yyyy");
}



function applyQuotationReadOnlyUI_() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!sheet) return;

  const status =
    sheet.getRange("B8")
      .getValue()
      .toString()
      .trim();

  const editable =
    canEditQuotation_(status);

  sheet.getRange("H8")
    .setValue(
      editable
        ? "Editable"
        : "Locked"
    );
}


function applyQuotationGridBorders_() {

  return; // Disable grid borders for now as they cause performance issues with larger item lists. Can be optimized later if needed.      

  
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.QUOTATION_FORM
      );

  if (!sheet) return;

  const startRow = 20;
  const endRow = 35;

  const range =
    sheet.getRange(
      startRow,
      1,
      endRow - startRow + 1,
      12
    );

  const data =
    range.getValues();

  for (
    let i = 0;
    i < data.length;
    i++
  ) {

    const rowData =
      data[i];

    const hasData =
      rowData.some(
        cell =>
          cell !== "" &&
          cell !== null
      );

    const rowRange =
      sheet.getRange(
        startRow + i,
        1,
        1,
        12
      );

    if (hasData) {

      rowRange.setBorder(
        true,
        true,
        true,
        true,
        true,
        true
      );

    } else {

      rowRange.setBorder(
        false,
        false,
        false,
        false,
        false,
        false
      );

    }

  }

}











