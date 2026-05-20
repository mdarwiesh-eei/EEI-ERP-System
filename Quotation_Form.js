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
  let selectedRevision = sh.getRange("E2").getValue();

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

  if (!selectedRevision || selectedRevision === "No QID" || selectedRevision === "No Revisions") {
    buildRevisionSelectorForForm(qID);
    selectedRevision = sh.getRange("E2").getValue();
  }

  if (!selectedRevision || selectedRevision === "No QID" || selectedRevision === "No Revisions") {
    SpreadsheetApp.getUi().alert("Select revision first.");
    return;
  }

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

  sh.getRange("B7").setValue(quotation.qID);
  sh.getRange("B5").setValue(quotation.customerName || "");
  sh.getRange("E5").setValue(quotation.projectName || "");
  sh.getRange("E7").setValue(quotation.status || "");
  sh.getRange("B9").setValue(quotation.rfqDate || "");
  sh.getRange("E9").setValue(quotation.assignedTo || "");
  sh.getRange("B11").setValue(quotation.currency || "");
  sh.getRange("E11").setValue(quotation.customerRFQLink || "");
  sh.getRange("B13").setValue(quotation.notes || "");

  sh.getRange("E2").setValue(selectedRevision);

  loadQuotationItemsToForm_(quotation.qID, selectedRevision);

  if (typeof loadQuotationTermsToForm_ === "function") {
    loadQuotationTermsToForm_(quotation.qID, selectedRevision);
  }

  refreshQuotationKPIsFromForm();

  // Reset selectors after load
  sh.getRange("B2").clearContent();
  sh.getRange("E2").clearContent();

  SpreadsheetApp
    .getUi()
    .alert("Quotation loaded ✅ " + quotation.qID + " - " + selectedRevision);
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

  const qID =
    sh.getRange("B7").getValue();

  const revision =
    sh.getRange("E2").getValue();

  if (!qID) {

    SpreadsheetApp
      .getUi()
      .alert(
        "Load quotation first"
      );

    return;
  }

  if (!revision) {

    SpreadsheetApp
      .getUi()
      .alert(
        "Select revision first"
      );

    return;
  }

  // Selected line reset
  sh.getRange("B20").clearContent();

  // Editor reset
  sh.getRangeList([
    "B21",
    "F21",
    "B22",
    "F22",
    "B23",
    "D23",
    "F23",
    "H23",
    "B24:H25"
  ]).clearContent();

  // Next line suggestion
  const nextLine =
    getNextQuotationLineNo_(
      qID,
      revision
    );

  sh.getRange("B20")
    .setValue(nextLine);

  SpreadsheetApp
    .getUi()
    .alert(
      "New item ready → Line "
      + nextLine
    );

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

  data.forEach(function (row) {

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



function showCreateQuotationDialog() {

  const html = HtmlService
    .createHtmlOutputFromFile("CreateQuotationDialog")
    .setWidth(520)
    .setHeight(620);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      "Create New Quotation"
    );
}


function getCreateQuotationDialogData() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const customersSheet =
    ss.getSheetByName(CONFIG.SHEETS.CUSTOMERS);

  const usersSheet =
    ss.getSheetByName(CONFIG.SHEETS.USERS);

  const customers = [];
  const users = [];

  if (
    customersSheet &&
    customersSheet.getLastRow() >= 3
  ) {

    const data =
      customersSheet
        .getRange(
          3,
          1,
          customersSheet.getLastRow() - 2,
          11
        )
        .getValues();

    data.forEach(function(row) {

      if (row[0] && row[1] && row[10] === "Active") {
        customers.push({
          id: row[0],
          name: row[1]
        });
      }

    });
  }

  if (
    usersSheet &&
    usersSheet.getLastRow() >= 2
  ) {

    const data =
      usersSheet
        .getRange(
          2,
          2,
          usersSheet.getLastRow() - 1,
          2
        )
        .getValues();

    data.forEach(function(row) {

      if (row[0] && row[1]) {
        users.push({
          email: row[0],
          name: row[1]
        });
      }

    });
  }

  const today =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

  return {
    customers: customers,
    users: users,
    today: today
  };
}


function submitCreateQuotationDialog(payload) {

  try {

    if (!payload.customerID) {
      throw new Error("Customer is required.");
    }

    if (!payload.projectName) {
      throw new Error("Project name is required.");
    }

    const result =
      createQuotation({
        customerID: payload.customerID,
        projectName: payload.projectName,
        rfqDate: payload.rfqDate
          ? new Date(payload.rfqDate)
          : "",
        assignedTo: payload.assignedTo,
        currency: payload.currency,
        customerRFQLink: payload.customerRFQLink,
        notes: payload.notes
      });

    if (!result || !result.success) {
      throw new Error(
        result && result.error
          ? result.error
          : "Create quotation failed."
      );
    }

    buildQuotationSelectorForForm();

    const sh = QFORM.SHEET();

    sh.getRange("B2")
      .setValue(result.qID);

    buildRevisionSelectorForForm(result.qID);

    sh.getRange("E2")
      .setValue(result.revisionNo);

    loadQuotationToForm();

    return result;

  } catch (err) {

    return {
      success: false,
      error: err.message
    };

  }
}



