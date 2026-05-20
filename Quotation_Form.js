// @ts-nocheck
/*****************************************************
 EEI - QUOTATION FORM CONTROLLER
*****************************************************/

const QFORM = {

  SHEET() {
    return SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);
  },

  getData() {

    const sh = this.SHEET();

    return {

      quotationSelector: sh.getRange("B2").getValue(),
      revision: sh.getRange("E2").getValue(),

      customer: sh.getRange("B5").getValue(),
      projectName: sh.getRange("E5").getValue(),

      quotationID: sh.getRange("B7").getValue(),
      status: sh.getRange("E7").getValue(),

      rfqDate: sh.getRange("B9").getValue(),
      assignedTo: sh.getRange("E9").getValue(),

      currency: sh.getRange("B11").getValue(),
      rfqLink: sh.getRange("E11").getValue(),

      notes: sh.getRange("B13").getValue()

    };

  }

};


/*****************************************************
 CLEAR FORM
*****************************************************/

function clearQuotationForm() {

  const sh = QFORM.SHEET();

  sh.getRangeList([
    "B2:C2",
    "E2:F2",
    "B5:C5",
    "E5:F5",
    "B7:C7",
    "E7:F7",
    "B9:C9",
    "E9:F9",
    "B11:C11",
    "E11:F11",
    "B13:F14",
    "I5:J9",
    "A19:H19",
    "A23:L90"
  ]).clearContent();

}


/*****************************************************
 REFRESH FORM
*****************************************************/

function refreshQuotationForm() {

  buildQuotationSelectorForForm();
  buildQuotationCustomerSelector();
  buildAssignedUserSelector();

  const data = QFORM.getData();

  if (data.quotationID) {
    buildRevisionSelectorForForm(data.quotationID);
    refreshQuotationKPIsFromForm();
  }

  SpreadsheetApp
    .getUi()
    .alert("Quotation form refreshed ✅");
}


/*****************************************************
 LOAD SELECTED QUOTATION
*****************************************************/

function loadQuotationToForm() {

  const sh = QFORM.SHEET();

  const selectorValue = sh.getRange("B2").getValue();
  const selectedRevision = sh.getRange("E2").getValue();

  if (!selectorValue) {
    SpreadsheetApp.getUi().alert("Select quotation first.");
    return;
  }

  const qID = extractQuotationIDFromSelector_(selectorValue);

  if (!qID) {
    SpreadsheetApp.getUi().alert("Could not read quotation ID from selector.");
    return;
  }

  const quotation = getQuotationById_(qID);

  if (!quotation) {
    SpreadsheetApp.getUi().alert("Quotation not found: " + qID);
    return;
  }

  let revisionNo = selectedRevision;

  if (!revisionNo || revisionNo === "No QID") {
    revisionNo = quotation.currentRevision || "R00";
  }

  // Clear old loaded data first
  sh.getRangeList([
    "B5:C5",
    "E5:F5",
    "B7:C7",
    "E7:F7",
    "B9:C9",
    "E9:F9",
    "B11:C11",
    "E11:F11",
    "B13:F14",
    "I5:J9",
    "A19:H19",
    "A23:L90"
  ]).clearContent();

  // Header
  sh.getRange("B7").setValue(quotation.qID);
  sh.getRange("B5").setValue(quotation.customerName || "");
  sh.getRange("E5").setValue(quotation.projectName || "");
  sh.getRange("E7").setValue(quotation.status || "");
  sh.getRange("B9").setValue(quotation.rfqDate || "");
  sh.getRange("E9").setValue(quotation.assignedTo || "");
  sh.getRange("B11").setValue(quotation.currency || "");
  sh.getRange("E11").setValue(quotation.customerRFQLink || "");
  sh.getRange("B13").setValue(quotation.notes || "");

  // Rebuild revision selector then keep chosen revision
  buildRevisionSelectorForForm(quotation.qID);
  sh.getRange("E2").setValue(revisionNo);

  // Load selected revision data
  loadQuotationItemsToForm_(quotation.qID, revisionNo);

  if (typeof loadQuotationTermsToForm_ === "function") {
    loadQuotationTermsToForm_(quotation.qID, revisionNo);
  }

  refreshQuotationKPIsFromForm();

  SpreadsheetApp
    .getUi()
    .alert("Quotation loaded ✅ " + quotation.qID + " / " + revisionNo);
}


/*****************************************************
 CREATE QUOTATION FROM FORM
*****************************************************/

function createQuotationFromForm() {

  const data = QFORM.getData();

  if (!data.customer) {
    SpreadsheetApp.getUi().alert("Customer is required.");
    return;
  }

  if (!data.projectName) {
    SpreadsheetApp.getUi().alert("Project name is required.");
    return;
  }

  const customerID = extractCustomerIDFromSelector_(data.customer);

  const result = createQuotation({
    customerID: customerID,
    projectName: data.projectName,
    rfqDate: data.rfqDate,
    customerRFQLink: data.rfqLink,
    assignedTo: data.assignedTo,
    currency: data.currency,
    notes: data.notes
  });

  if (result && result.success) {
    buildQuotationSelectorForForm();
    SpreadsheetApp.getUi().alert("Quotation created: " + result.qID);
  }

}


/*****************************************************
 ADD ITEMS FROM GRID
*****************************************************/

function addQuotationItemsFromGrid() {

  const sh = QFORM.SHEET();
  const data = QFORM.getData();

  const qID = data.quotationID;
  const revisionNo = data.revision;

  if (!qID) {
    SpreadsheetApp.getUi().alert("Load quotation first.");
    return;
  }

  const rows = sh.getRange("A23:L90").getValues();

  const items = [];

  rows.forEach(function(row) {

    const description = row[1];
    const transformerType = row[2];
    const powerKVA = row[3];
    const voltage = row[4];
    const quantity = row[5];
    const unitPrice = row[6];
    const deliveryTime = row[8];
    const warranty = row[9];
    const notes = row[10];

    if (description || quantity || unitPrice) {

      items.push({
        description: description,
        transformerType: transformerType,
        powerKVA: powerKVA,
        voltage: voltage,
        quantity: quantity,
        unitPrice: unitPrice,
        currency: data.currency,
        deliveryTime: deliveryTime,
        warranty: warranty,
        notes: notes
      });

    }

  });

  if (!items.length) {
    SpreadsheetApp.getUi().alert("No items to add.");
    return;
  }

  const result = addQuotationItemsBatch({
    qID: qID,
    revisionNo: revisionNo,
    items: items
  });

  if (result && result.success) {
    loadQuotationItemsToForm_(qID, revisionNo);
    refreshQuotationKPIsFromForm();
    SpreadsheetApp.getUi().alert(result.addedCount + " item(s) added ✅");
  }

}


/*****************************************************
 LOAD ITEMS TO FORM
*****************************************************/

function loadQuotationItemsToForm_(qID, revisionNo) {

  const sh = QFORM.SHEET();
  const itemsSheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_ITEMS);

  sh.getRange("A23:L90").clearContent();

  if (itemsSheet.getLastRow() < 2) return;

  const data = itemsSheet
    .getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn())
    .getValues();

  const output = [];

  data.forEach(function(row) {

    const itemQID = row[1];
    const itemRevision = row[2];
    const itemStatus = row[19] || "Active";

    if (
      itemQID === qID &&
      itemRevision === revisionNo &&
      itemStatus !== "Deleted"
    ) {

      output.push([
        row[3],   // Line
        row[4],   // Description
        row[5],   // Type
        row[6],   // Power
        row[7],   // Voltage
        row[8],   // Qty
        row[9],   // Unit Price
        row[10],  // Total
        row[12],  // Delivery
        row[13],  // Warranty
        row[14],  // Notes
        ""
      ]);

    }

  });

  if (output.length) {
    sh.getRange(23, 1, output.length, 12).setValues(output);
  }

}


/*****************************************************
 REVISION
*****************************************************/

function createRevisionFromForm() {

  const data = QFORM.getData();

  if (!data.quotationID) {
    SpreadsheetApp.getUi().alert("Load quotation first.");
    return;
  }

  createQuotationRevision(data.quotationID);

  buildRevisionSelectorForForm(data.quotationID);

  SpreadsheetApp
    .getUi()
    .alert("Revision created ✅");
}

function createQuotationRevisionFromForm() {
  createRevisionFromForm();
}


/*****************************************************
 OPEN FOLDERS
*****************************************************/

function openQuotationFolder() {

  const data = QFORM.getData();

  if (!data.quotationID) {
    SpreadsheetApp.getUi().alert("Load quotation first.");
    return;
  }

  const quotation = getQuotationById_(data.quotationID);

  if (!quotation || !quotation.folderLink) {
    SpreadsheetApp.getUi().alert("Quotation folder not found.");
    return;
  }

  openUrl_(quotation.folderLink);
}

function openQuotationFolderFromForm() {
  openQuotationFolder();
}

function openRFQFolderFromForm() {
  openQuotationFolder();
}


/*****************************************************
 HELPERS
*****************************************************/

function extractQuotationIDFromSelector_(selector) {

  if (!selector) return "";

  const text = selector.toString();

  const match = text.match(/Q-\d+/);

  return match ? match[0] : "";
}

function extractCustomerIDFromSelector_(selector) {

  if (!selector) return "";

  const text = selector.toString();

  const match = text.match(/C-\d+/);

  return match ? match[0] : "";
}

function openUrl_(url) {

  const html = HtmlService.createHtmlOutput(
    '<script>' +
    'window.open("' + url + '", "_blank");' +
    'google.script.host.close();' +
    '</script>'
  );

  SpreadsheetApp
    .getUi()
    .showModalDialog(html, "Opening...");
}