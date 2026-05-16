function installedOnEdit(e) {

  if (!e) return;

  const sheet = e.range.getSheet();
  const name = sheet.getName();
  const cell = e.range.getA1Notation();

  // =====================
  // CUSTOMER PROFILE
  // =====================
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

  // =====================
  // REGISTER CUSTOMER
  // =====================
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

  // =====================
  // QUOTATION FORM
  // =====================
  if (name === CONFIG.SHEETS.QUOTATION_FORM) {

    const qID = sheet.getRange("B6").getValue();

    if (cell === "K2" && e.range.getValue() === true) {
      e.range.setValue(false);
      SpreadsheetApp.flush();

      loadQuotationToForm();
      refreshQuotationForm();


      return;
    }

    if (cell === "K4" && e.range.getValue() === true) {
      createQuotationFromForm();
      e.range.setValue(false);
      return;
    }

    if (cell === "K5" && e.range.getValue() === true) {
      addQuotationItemsFromGrid();
      refreshQuotationForm();
      e.range.setValue(false);
      return;
    }

    if (cell === "K6" && e.range.getValue() === true) {
      submitQuotationForReview(qID);
      refreshQuotationForm();
      e.range.setValue(false);
      return;
    }

    if (cell === "K7" && e.range.getValue() === true) {
      approveQuotation(qID, "Approved from Quotation Form");
      refreshQuotationForm();
      e.range.setValue(false);
      return;
    }

    if (cell === "K8" && e.range.getValue() === true) {
      sendQuotation(qID, "Sent from Quotation Form");
      refreshQuotationForm();
      e.range.setValue(false);
      return;
    }

    if (cell === "K9" && e.range.getValue() === true) {
      createQuotationRevisionFromForm();
      refreshQuotationForm();
      e.range.setValue(false);
      return;
    }

    if (cell === "K10" && e.range.getValue() === true) {
      markQuotationWon(qID, "Marked won from Quotation Form");
      refreshQuotationForm();
      e.range.setValue(false);
      return;
    }

    if (cell === "K11" && e.range.getValue() === true) {
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

    if (cell === "K12" && e.range.getValue() === true) {
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

    if (cell === "K13" && e.range.getValue() === true) {
      refreshQuotationForm();
      e.range.setValue(false);
      return;
    }

    if (cell === "K14" && e.range.getValue() === true) {
      clearQuotationForm();
      e.range.setValue(false);
      return;
    }

    if (cell === "K15" && e.range.getValue() === true) {
      openQuotationFolderFromForm();
      e.range.setValue(false);
      return;
    }

    if (cell === "K16" && e.range.getValue() === true) {
      openRFQFolderFromForm();
      e.range.setValue(false);
      return;
    }
  }
}