// @ts-nocheck
/*****************************************************
 EEI GLOBAL TRIGGERS
*****************************************************/

// @ts-nocheck
/*****************************************************
 EEI GLOBAL TRIGGERS
*****************************************************/

function handleQuotationEdit(e) {

  if (!e) return;

  const sheet = e.source.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEETS.QUOTATION_FORM) {
    return;
  }

  const cell = e.range.getA1Notation();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  /*****************************************************
   When quotation selector changes, rebuild revision selector
  *****************************************************/
  if (cell === "B2") {

    const selectorValue = e.range.getValue();
    const qID = extractQuotationIDFromSelector_(selectorValue);

    buildRevisionSelectorForForm(qID);
    return;
  }

  /*****************************************************
   Auto fill item data when Type is selected
   Type column = C
   Grid rows = 22:90
  *****************************************************/
  if (col === 3 && row >= 22 && row <= 90) {

    fillItemDataFromType_(row);
    return;
  }

  const value = e.range.getValue();

  if (value !== true) return;

  try {

    switch (cell) {

      case "K4":
        showCreateQuotationDialog();
        break;

      case "K5":
        loadQuotationToForm();
        break;

      case "K6":
        saveQuotationItemsGrid();
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

      case "K20":
        archiveQuotationFromForm();
        break;
    }

  } catch (err) {

    SpreadsheetApp
      .getUi()
      .alert("Trigger Error:\n" + err.message);

    Logger.log(err);

  } finally {

    e.range.setValue(false);

  }
}


/*****************************************************
 AUTO FILL ITEM DATA FROM TRANSFORMER_TYPES
*****************************************************/

function fillItemDataFromType_(targetRow) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = ss.getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);
  const typesSheet = ss.getSheetByName("Transformer_Types");

  if (!form || !typesSheet) return;

  const selectedType = String(form.getRange(targetRow, 3).getDisplayValue()).trim();

  if (!selectedType) {

    form.getRange(targetRow, 2).clearContent(); // Description
    form.getRange(targetRow, 4).clearContent(); // Power
    form.getRange(targetRow, 5).clearContent(); // Ratio
    form.getRange(targetRow, 9).clearContent(); // Delivery
    form.getRange(targetRow, 10).clearContent(); // Warranty

    return;
  }

  if (typesSheet.getLastRow() < 2) return;

  const data = typesSheet
    .getRange(2, 1, typesSheet.getLastRow() - 1, typesSheet.getLastColumn())
    .getValues();

  const headers = typesSheet
    .getRange(1, 1, 1, typesSheet.getLastColumn())
    .getValues()[0];

  const idx = getTransformerTypeIndexes_(headers);

  for (let i = 0; i < data.length; i++) {

    const row = data[i];
    const typeValue = String(row[idx.type] || "").trim();

    if (typeValue === selectedType) {

      form.getRange(targetRow, 2).setValue(row[idx.description] || selectedType);
      form.getRange(targetRow, 4).setValue(row[idx.power] || "");
      form.getRange(targetRow, 5).setValue(row[idx.ratio] || "");
      form.getRange(targetRow, 9).setValue(row[idx.delivery] || "");
      form.getRange(targetRow, 10).setValue(row[idx.warranty] || "");

      return;
    }
  }
}


/*****************************************************
 TRANSFORMER_TYPES HEADER MAPPING
*****************************************************/

function getTransformerTypeIndexes_(headers) {

  const normalized = headers.map(function (h) {
    return String(h || "")
      .toLowerCase()
      .replace(/\s+/g, "")
      .replace(/_/g, "");
  });

  function findIndex(possibleNames, fallbackIndex) {

    for (let i = 0; i < possibleNames.length; i++) {

      const name = possibleNames[i]
        .toLowerCase()
        .replace(/\s+/g, "")
        .replace(/_/g, "");

      const index = normalized.indexOf(name);

      if (index !== -1) {
        return index;
      }
    }

    return fallbackIndex;
  }

  return {
    type: findIndex(["Type", "TransformerType", "Model", "ItemType"], 0),
    description: findIndex(["Description", "ItemDescription"], 1),
    power: findIndex(["Power", "PowerKVA", "kVA", "Power[kVA]"], 2),
    ratio: findIndex(["Ratio", "Voltage", "RatioKV", "Ratio[kV]"], 3),
    delivery: findIndex(["Delivery", "DeliveryTime"], 4),
    warranty: findIndex(["Warranty"], 5)
  };
}

/*****************************************************
 WORKFLOW WRAPPERS
*****************************************************/

function approveQuotationFromForm_() {

  SpreadsheetApp
    .getUi()
    .alert(
      "Approve workflow next"
    );

}


function sendQuotationFromForm_() {

  SpreadsheetApp
    .getUi()
    .alert(
      "Send workflow next"
    );

}


function markQuotationWon_() {

  SpreadsheetApp
    .getUi()
    .alert(
      "Won workflow next"
    );

}


function markQuotationLost_() {

  SpreadsheetApp
    .getUi()
    .alert(
      "Lost workflow next"
    );

}


function cancelQuotation_() {

  SpreadsheetApp
    .getUi()
    .alert(
      "Cancel workflow next"
    );

}



function setupInstallableOnEditTrigger() {

  const ss = SpreadsheetApp.getActive();

  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "handleQuotationEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp
    .newTrigger("handleQuotationEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp
    .getUi()
    .alert("Installable onEdit trigger created ✅");
}