// ===============================
// CUSTOMER MODULE
// ===============================


// ===============================
// CUSTOMER FORM ACTIONS
// ===============================


function registerCustomer() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const form =
      ss.getSheetByName(
        CONFIG.SHEETS.REGISTER
      );

    const db =
      ss.getSheetByName(
        CONFIG.SHEETS.CUSTOMERS
      );

    const ui =
      SpreadsheetApp.getUi();

    // ===============================
    // READ FORM DATA
    // ===============================

    var name =
      form.getRange("B4")
        .getValue()
        .toString()
        .trim();

    var type =
      form.getRange("B5")
        .getValue();

    var city =
      form.getRange("B6")
        .getValue();

    var address =
      form.getRange("B7")
        .getValue();

    var referredBy =
      form.getRange("B8")
        .getValue();

    var contact =
      form.getRange("B11")
        .getValue();

    var phone =
      form.getRange("B12")
        .getValue()
        .toString()
        .trim();

    var email =
      form.getRange("B13")
        .getValue()
        .toString()
        .trim();

    var notes =
      form.getRange("B15")
        .getValue();

    name =
      name.replace(/\s+/g, " ");

    // ===============================
    // VALIDATION
    // ===============================

    if (!name) {

      ui.alert(
        "Please enter customer name."
      );

      return;
    }

    if (!isValidName(name)) {

      ui.alert(
        "Invalid customer name."
      );

      return;
    }

    if (phone && !isValidPhone(phone)) {

      ui.alert(
        "Invalid phone number."
      );

      return;
    }

    if (email && !isValidEmail(email)) {

      ui.alert(
        "Invalid email format."
      );

      return;
    }

    // ===============================
    // DUPLICATE CHECK
    // ===============================



    if (db.getLastRow() >= 3) {
      const existing = db.getRange(
        3,
        2,
        db.getLastRow() - 2,
        1
      ).getValues();

      // تعديل هذا الجزء لاستخدام function(row) بدلاً من r =>
      const exists = existing.some(function(row) {
        return row[0] && row[0].toString().trim().toLowerCase() === name.toLowerCase();
      });

      if (exists) {
        ui.alert("Customer already exists.");
        return;
      }
    }
        // ===============================
        // GENERATE CUSTOMER ID
        // ===============================

        const lastRowDb =
          db.getLastRow();

        const nextRow =
          lastRowDb < 3
            ? 3
            : lastRowDb + 1;

        const customerNumber =
          nextRow - 2;

        const customerID =
          "C-" +
          ("0000" + customerNumber)
            .slice(-4);

        // ===============================
        // CREATE FOLDER
        // ===============================

        let folderLink = "";

        if (
          typeof createCustomerFolder
          === "function"
        ) {

          folderLink =
            createCustomerFolder(
              customerID,
              name
            );
        }

        // ===============================
        // SAVE DATA
        // ===============================

        const rowData = [[

          customerID,
          name,
          type,
          city,
          address,
          contact,
          referredBy,
          phone,
          email,
          folderLink,
          "Active",
          notes,
          new Date(),
          getCurrentUserName()

        ]];

        db.getRange(
          nextRow,
          1,
          1,
          14
        ).setValues(rowData);

        // ===============================
        // REFRESH SELECTORS
        // ===============================

      

        if (
          typeof buildCustomerSelector
          === "function"
        ) {

          buildCustomerSelector();
        }

          // ===============================
          // AUDIT LOG
          // ===============================

          if (
            typeof logAction
            === "function"
          ) {

            logAction({

              module: "Customers",

              action: "CREATE",

              recordID: customerID,

              field: "Customer",

              oldValue: "",

              newValue: name,

              notes:
                "New customer created"

            });
          }

          // ===============================
          // CLEAR FORM
          // ===============================

          clearCustomerForm();

          ui.alert(
            "Customer Created ✅"
          );

        }

        catch(error) {

          SpreadsheetApp.getUi()
            .alert(
              "Error: " + error.message
            );

          Logger.log(error);

        }

        finally {

          lock.releaseLock();

        }
}


function updateCustomer() {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const form =
      ss.getSheetByName("Register_Customer");

    const db =
      ss.getSheetByName("Customers");

    const ui =
      SpreadsheetApp.getUi();

    const customerID =
      form.getRange("D4").getValue();

    if (!customerID) {

      ui.alert("Load customer first");

      return;
    }

    const data =
      db.getRange(
        3,
        1,
        db.getLastRow() - 2,
        14
      ).getValues();

    for (let i = 0; i < data.length; i++) {

      if (data[i][0] === customerID) {

        const row = i + 3;

        // ===============================
        // OLD VALUES
        // ===============================

        const oldName = data[i][1];

        // ===============================
        // NEW VALUES
        // ===============================

        const newName =
          form.getRange("B4").getValue();

        const type =
          form.getRange("B5").getValue();

        const city =
          form.getRange("B6").getValue();

        const address =
          form.getRange("B7").getValue();

        const contact =
          form.getRange("B11").getValue();

        const referredBy =
          form.getRange("B8").getValue();

        const phone =
          form.getRange("B12").getValue();

        const email =
          form.getRange("B13").getValue();

        const notes =
          form.getRange("B15").getValue();

        // ===============================
        // UPDATE DATABASE
        // ===============================

        db.getRange(row, 2).setValue(newName);

        db.getRange(row, 3).setValue(type);

        db.getRange(row, 4).setValue(city);

        db.getRange(row, 5).setValue(address);

        db.getRange(row, 6).setValue(contact);

        db.getRange(row, 7).setValue(referredBy);

        db.getRange(row, 8).setValue(phone);

        db.getRange(row, 9).setValue(email);

        db.getRange(row, 12).setValue(notes);

        // ===============================
        // AUDIT LOG
        // ===============================

        logAction({

          module: "Customers",

          action: "UPDATE",

          recordID: customerID,

          field: "Customer",

          oldValue: oldName,

          newValue: newName,

          notes: "Customer data updated"

        });

        ui.alert("Customer Updated ✅");

        return;
      }
    }

  } finally {

    lock.releaseLock();

  }
}


function clearCustomerForm() {

  const form =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Register_Customer");

  form.getRangeList([
    "B4","B5","B6","B7","B8",
    "B11","B12","B13","B15","D4"
  ]).clearContent();
}


function deleteCustomer() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const form =
    ss.getSheetByName("Register_Customer");

  const db =
    ss.getSheetByName("Customers");

  const ui =
    SpreadsheetApp.getUi();

  const customerID =
    form.getRange("D4").getValue();

  if (!customerID) {

    ui.alert("Load customer first");

    return;
  }

  const confirm =
    ui.alert(

      "Delete Customer",

      "Are you sure?",

      ui.ButtonSet.YES_NO

    );

  if (confirm !== ui.Button.YES) return;

  const data =
    db.getRange(
      3,
      1,
      db.getLastRow() - 2,
      14
    ).getValues();

  for (let i = 0; i < data.length; i++) {

    if (data[i][0] === customerID) {

      const row = i + 3;

      db.getRange(row,11)
        .setValue("Deleted");

      logAction({

        module: "Customers",

        action: "DELETE",

        recordID: customerID,

        oldValue: data[i][1],

        notes: "Customer soft deleted"

      });

      ui.alert("Customer Deleted ✅");

      clearCustomerForm();

      buildCustomerSelector();

      return;
    }
  }
}