
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

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const users =
    ss.getSheetByName(CONFIG.SHEETS.USERS);

  const form =
    ss.getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!users || !form) return;
  if (users.getLastRow() < 2) return;

  const data =
    users
      .getRange(
        2,
        1,
        users.getLastRow() - 1,
        6
      )
      .getValues();

  const allowedRoles = [
    "Sales",
    "Engineering",
    "Admin"
  ];

  const list = [];

  data.forEach(function (row) {

    const email = row[1];
    const name = row[2];
    const role = row[3];
    const status = row[5];

    if (
      email &&
      name &&
      allowedRoles.includes(role) &&
      status === "Active"
    ) {

      list.push(
        name + " | " + email
      );

    }

  });

  if (!list.length) return;

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(list, true)
      .setAllowInvalid(false)
      .build();

  form.getRange("E5")
    .clearDataValidations()
    .setDataValidation(rule);
}




function buildRevisionSelectorForForm(qID) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const revisions =
    ss.getSheetByName(CONFIG.SHEETS.QUOTATION_REVISIONS);

  const form =
    ss.getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!form) return;

  const cell = form.getRange("E2");

  cell.clearDataValidations();

  if (!revisions) {
    cell.setValue("Revision sheet not found");
    return;
  }

  if (!qID) {
    cell.setValue("No QID");
    return;
  }

  qID = qID.toString().trim();

  if (revisions.getLastRow() < 2) {
    cell.setValue("No revisions data");
    return;
  }

  const data = revisions
    .getRange(2, 1, revisions.getLastRow() - 1, 20)
    .getValues();

  const list = [];

  data.forEach(function (row) {

    const rowQID =
      row[1] ? row[1].toString().trim() : "";

    const revisionNo =
      row[2] ? row[2].toString().trim() : "";

    const status =
      row[3] ? row[3].toString().trim() : "";

    const reason =
      row[4] ? row[4].toString().trim() : "";

    if (rowQID === qID && revisionNo) {

      list.push(
        revisionNo +
        " | " +
        status +
        (reason ? " | " + reason : "")
      );

    }

  });

  if (!list.length) {
    cell.setValue("No revision for " + qID);
    return;
  }

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(true)
    .build();

  cell.setDataValidation(rule);
  cell.setValue(list[0]);
}



