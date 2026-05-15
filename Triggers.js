function installedOnEdit(e) {

  if (!e) return;

  const ss = e.source;

  const sheet = e.range.getSheet();

  const name = sheet.getName();

  const cell = e.range.getA1Notation();

  // ====================================
  // CUSTOMER PROFILE
  // ====================================

  if (name === "Customer_Profile") {

    if (cell === "B3") {

      if (e.value && e.value.includes("|")) {

        loadCustomerProfile();

      }

      return;
    }

    if (cell === "D3" && e.range.getValue() === true) {

      loadCustomerProfile();

      e.range.setValue(false);

      return;
    }

    if (cell === "F9" && e.range.getValue() === true) {

      editCustomerFromProfile();

      e.range.setValue(false);

      return;
    }

    if (cell === "F10" && e.range.getValue() === true) {

      clearCustomerProfile();

      e.range.setValue(false);

      return;
    }


  }
  
    refreshAll();

  // ====================================
  // REGISTER CUSTOMER
  // ====================================

  if (name === "Register_Customer") {

    if (cell === "F16" && e.range.getValue() === true) {

      registerCustomer();

      e.range.setValue(false);

      return;
    }

    if (cell === "F17" && e.range.getValue() === true) {

      updateCustomer();

      e.range.setValue(false);

      return;
    }

    if (cell === "F18" && e.range.getValue() === true) {

      clearCustomerForm();

      e.range.setValue(false);

      return;
    }
  }





//====================
// Quotation Management Form
//====================

if (name === CONFIG.SHEETS.QUOTATION_FORM) {

  const qID = sheet.getRange("B6").getValue();

  // K4 = Create Quotation
  if (cell === "K3" && e.range.getValue() === true) {
    createQuotationFromForm();
    e.range.setValue(false);
    return;
  }

  // K5 = Add Items
  if (cell === "K4" && e.range.getValue() === true) {
    addQuotationItemsFromGrid();
    refreshQuotationForm();
    e.range.setValue(false);
    return;
  }

  // K6 = Submit Review
  if (cell === "K5" && e.range.getValue() === true) {
    submitQuotationForReview(qID);
    refreshQuotationForm();
    e.range.setValue(false);
    return;
  }

  // K7 = Approve
  if (cell === "K6" && e.range.getValue() === true) {
    approveQuotation(qID, "Approved from Quotation Form");
    refreshQuotationForm();
    e.range.setValue(false);
    return;
  }

  // K8 = Send
  if (cell === "K7" && e.range.getValue() === true) {
    sendQuotation(qID, "Sent from Quotation Form");
    refreshQuotationForm();
    e.range.setValue(false);
    return;
  }

  // K9 = Create Revision
  if (cell === "K8" && e.range.getValue() === true) {
    createQuotationRevisionFromForm();
    refreshQuotationForm();
    e.range.setValue(false);
    return;
  }

  // K10 = Mark Won
  if (cell === "K9" && e.range.getValue() === true) {
    markQuotationWon(qID, "Marked won from Quotation Form");
    refreshQuotationForm();
    e.range.setValue(false);
    return;
  }

  // K11 = Mark Lost
  if (cell === "K10" && e.range.getValue() === true) {
    const reason = SpreadsheetApp.getUi().prompt(
      "Lost Reason",
      "Please enter lost reason:",
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );

    if (reason.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
      markQuotationLost(
        qID,
        reason.getResponseText(),
        "Marked lost from Quotation Form"
      );
      refreshQuotationForm();
    }

    e.range.setValue(false);
    return;
  }

  // K12 = Cancel
  if (cell === "K11" && e.range.getValue() === true) {
    const reason = SpreadsheetApp.getUi().prompt(
      "Cancel Reason",
      "Please enter cancellation reason:",
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );

    if (reason.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
      cancelQuotation(
        qID,
        reason.getResponseText(),
        "Cancelled from Quotation Form"
      );
      refreshQuotationForm();
    }

    e.range.setValue(false);
    return;
  }

  // K13 = Refresh
  if (cell === "K12" && e.range.getValue() === true) {
    refreshQuotationForm();
    e.range.setValue(false);
    return;
  }

  // K14 = Clear Form
  if (cell === "K13" && e.range.getValue() === true) {
    clearQuotationForm();
    e.range.setValue(false);
    return;
  }

  // K15 = Open Folder
  if (cell === "K14" && e.range.getValue() === true) {
    openQuotationFolderFromForm();
    e.range.setValue(false);
    return;
  }



  // K2 = Load Quotation
  if (cell === "K2" && e.range.getValue() === true) {
    loadQuotationToForm();
    e.range.setValue(false);
    return;
  }

  // K16 = Open RFQ Folder
  if (cell === "K16" && e.range.getValue() === true) {
    openRFQFolderFromForm();
    e.range.setValue(false);
    return;
  }




}
}










