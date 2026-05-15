function onOpen() {

  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🏭 EEI System")

    .addSubMenu(
  ui.createMenu("👤 Customer Management")

    .addItem(
      "➕ Add Customer",
      "openCustomerForm"
    )

    .addItem(
      "📋 View Customer Profile",
      "openCustomerProfile"
    )

    .addItem(
      "🔄 Refresh Selectors",
      "refreshFromMenu"
    )
  ) 

    .addToUi();
}





function openCustomerProfile() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      "Customer_Profile"
    );

  if (!sheet) {

    SpreadsheetApp.getUi()
      .alert(
        "Customer_Profile sheet not found."
      );

    return;
  }

  ss.setActiveSheet(sheet);

  sheet.getRange("B3").activate();
}





function openCustomerForm() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      "Register_Customer"
    );

  if (!sheet) {

    SpreadsheetApp.getUi()
      .alert(
        "Register_Customer sheet not found."
      );

    return;
  }

  ss.setActiveSheet(sheet);

  sheet.getRange("B4").activate();
}


