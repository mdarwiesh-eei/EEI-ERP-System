
function buildCustomerSelector() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = ss.getSheetByName(CONFIG.SHEETS.CUSTOMERS);
  const profile = ss.getSheetByName("Customer_Profile");

  if (!db || !profile || db.getLastRow() < 3) return;

  const data = db.getRange(3, 1, db.getLastRow() - 2, 11).getValues();

  const list = data
    .filter(function (r) {
      return r[0] && r[1] && r[10] === "Active";
    })
    .map(function (r) {
      return r[0] + " | " + r[1];
    });

  if (list.length === 0) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();

  profile.getRange("B3").setDataValidation(rule);
}





function buildQuotationCustomerSelector() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const customers = ss.getSheetByName(CONFIG.SHEETS.CUSTOMERS);
  const form = ss.getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!customers || !form) return;
  if (customers.getLastRow() < 3) return;

  const data = customers
    .getRange(3, 1, customers.getLastRow() - 2, 11)
    .getValues();

  const list = [];

  data.forEach(function (row) {

    if (
      row[0] &&
      row[1] &&
      row[10] === "Active"
    ) {
      list.push(row[0] + " | " + row[1]);
    }

  });

  if (!list.length) return;

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();

  form.getRange("B4")
    .clearDataValidations()
    .setDataValidation(rule);
}



function buildQuotationSelectorForForm() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const quotations =
    ss.getSheetByName(CONFIG.SHEETS.QUOTATIONS);

  const form =
    ss.getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!quotations || !form) return;

  if (quotations.getLastRow() < 2) return;

  const data =
    quotations
      .getRange(
        2,
        1,
        quotations.getLastRow() - 1,
        6
      )
      .getValues();

  const list = [];

  data.forEach(function (row) {

    const qID =
      row[0] ? row[0].toString().trim() : "";

    const project =
      row[3] ? row[3].toString().trim() : "";

    const status =
      row[5] ? row[5].toString().trim() : "";

    if (qID && project && status) {

      list.push(
        qID +
        " | " +
        project +
        " | " +
        status
      );

    }

  });

  if (!list.length) return;

  const cell =
    form.getRange("B2");

  const currentValue =
    cell.getDisplayValue()
      .toString()
      .trim();

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(list, true)
      .setAllowInvalid(true)
      .build();

  cell.clearDataValidations();
  cell.setDataValidation(rule);

  if (currentValue && list.includes(currentValue)) {
    cell.setValue(currentValue);
  }
}



function buildAssignedUserSelector() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = ss.getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!form) return;

  form.getRange("E9:F9").clearDataValidations();
}



function buildRevisionSelectorForForm(qID) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = ss.getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);
  const revisionsSheet = ss.getSheetByName(CONFIG.SHEETS.QUOTATION_REVISIONS);

  if (!form || !revisionsSheet) return;

  const revisionCell = form.getRange("E2");

  revisionCell.clearContent();
  revisionCell.clearDataValidations();

  if (!qID) {
    const selectorValue = form.getRange("B2").getValue();
    qID = extractQuotationIDFromSelector_(selectorValue);
  }

  if (!qID) {
    revisionCell.setValue("No QID");
    return;
  }

  if (revisionsSheet.getLastRow() < 2) {
    revisionCell.setValue("No Revisions");
    return;
  }

  const data = revisionsSheet
    .getRange(2, 1, revisionsSheet.getLastRow() - 1, revisionsSheet.getLastColumn())
    .getValues();

  const revisions = [];

  data.forEach(function(row) {
    const rowQID = row[1];
    const revisionNo = row[2];

    if (rowQID === qID && revisionNo) {
      revisions.push(revisionNo.toString());
    }
  });

  if (!revisions.length) {
    revisionCell.setValue("No Revisions");
    return;
  }

  revisions.sort();

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(revisions, true)
    .setAllowInvalid(false)
    .build();

  revisionCell.setDataValidation(rule);

  const quotation = getQuotationById_(qID);
  const defaultRevision = quotation && quotation.currentRevision
    ? quotation.currentRevision
    : revisions[revisions.length - 1];

  revisionCell.setValue(defaultRevision);
}


