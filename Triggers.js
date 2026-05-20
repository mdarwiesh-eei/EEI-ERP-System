// @ts-nocheck
/*****************************************************
 EEI GLOBAL TRIGGERS
*****************************************************/

function onEdit(e) {

  if (!e) return;

  const sheet = e.source.getActiveSheet();

  if (
    sheet.getName() !==
    CONFIG.SHEETS.QUOTATION_FORM
  ) {
    return;
  }

  const cell =
    e.range.getA1Notation();

  // When quotation selector changes, rebuild revision selector
  if (cell === "B2") {

    const selectorValue = e.range.getValue();
    const qID = extractQuotationIDFromSelector_(selectorValue);

    if (qID) {
      buildRevisionSelectorForForm(qID);
    }

    return;
  }

  const value =
    e.range.getValue();

  if (value !== true) return;

  try {

    switch (cell) {

      case "K4":
        createQuotationFromForm();
        break;

      case "K5":
        loadQuotationToForm();
        break;

      case "K6":
        addQuotationItemsFromGrid();
        break;

      case "K7":
        approveQuotationFromForm_();
        break;

      case "K8":
        sendQuotationFromForm_();
        break;

      case "K9":
        createQuotationRevisionFromForm();
        break;

      case "K10":
        markQuotationWon_();
        break;

      case "K11":
        markQuotationLost_();
        break;

      case "K12":
        cancelQuotation_();
        break;

      case "K13":
        refreshQuotationForm();
        break;

      case "K14":
        clearQuotationForm();
        break;

      case "K15":
        openRFQFolderFromForm();
        break;

      case "K16":
        openQuotationFolderFromForm();
        break;

      case "K17":
        loadSelectedItemToEditor();
        break;

      case "K18":
        saveSelectedItemChanges();
        break;

      case "K19":
        deleteSelectedItemSoft();
        break;

      case "K20":
        archiveQuotationFromForm();
        break;
    }

  } catch (err) {

    SpreadsheetApp
      .getUi()
      .alert(
        "Trigger Error:\n" +
        err.message
      );

    Logger.log(err);

  } finally {

    e.range.setValue(false);

  }
}


/*****************************************************
 WORKFLOW WRAPPERS
*****************************************************/

function approveQuotationFromForm_(){

SpreadsheetApp
.getUi()
.alert(
"Approve workflow next"
);

}


function sendQuotationFromForm_(){

SpreadsheetApp
.getUi()
.alert(
"Send workflow next"
);

}


function markQuotationWon_(){

SpreadsheetApp
.getUi()
.alert(
"Won workflow next"
);

}


function markQuotationLost_(){

SpreadsheetApp
.getUi()
.alert(
"Lost workflow next"
);

}


function cancelQuotation_(){

SpreadsheetApp
.getUi()
.alert(
"Cancel workflow next"
);

}