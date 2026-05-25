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
        submitQuotationForReviewFromForm_();
        break;

      case "K8":
        approveQuotationFromForm_();
        break;

      case "K9":
        sendQuotationFromForm_();
        break;

      case "K10":
        createQuotationRevisionFromForm();
        break;

      case "K11":
        markQuotationWon_();
        break;

      case "K12":
        markQuotationLost_();
        break;

      case "K13":
        cancelQuotation_();
        break;

      case "K14":
        refreshQuotationForm();
        break;

      case "K15":
        clearQuotationForm();
        break;

      case "K16":
        openRFQFolderFromForm();
        break;

      case "K17":
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

function getLoadedQuotationContext_() {

  const sh = QFORM.SHEET();

  const qID = String(
    sh.getRange("B7").getDisplayValue()
  ).trim();

  const loadedRevision = String(
    sh.getRange("I7").getDisplayValue()
  ).trim();

  const loadedStatus = String(
    sh.getRange("E7").getDisplayValue()
  ).trim();

  if (!qID) {
    throw new Error("Load quotation first.");
  }

  if (!loadedRevision) {
    throw new Error("Loaded revision is missing.");
  }

  if (!loadedStatus) {
    throw new Error("Loaded revision status is missing.");
  }

  const quotation = getQuotationById_(qID);

  if (!quotation) {
    throw new Error("Quotation not found.");
  }

  if (String(quotation.currentRevision) !== String(loadedRevision)) {
    throw new Error(
      "Workflow actions are allowed only on the current revision.\n\n" +
      "Loaded revision: " + loadedRevision + "\n" +
      "Current revision: " + quotation.currentRevision
    );
  }

  return {
    qID: qID,
    revision: loadedRevision,
    status: loadedStatus,
    quotation: quotation
  };
}


function reloadLoadedQuotation_(qID) {

  const sh = QFORM.SHEET();

  const quotation = getQuotationById_(qID);

  if (!quotation) {
    throw new Error("Quotation not found.");
  }

  sh.getRange("B2").setValue(qID);

  buildRevisionSelectorForForm(qID);

  sh.getRange("E2").setValue(quotation.currentRevision);

  loadQuotationToForm();
}


function submitQuotationForReviewFromForm_() {

  const ctx = getLoadedQuotationContext_();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    "Submit for Review",
    "Submit quotation " + ctx.qID + " / " + ctx.revision + " for review?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const result = submitQuotationForReview(ctx.qID);

  if (result && result.success) {
    reloadLoadedQuotation_(ctx.qID);
    ui.alert("Quotation submitted for review ✅");
  }
}


function approveQuotationFromForm_() {

  const ctx = getLoadedQuotationContext_();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    "Approve Quotation",
    "Approve quotation " + ctx.qID + " / " + ctx.revision + "?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const result = approveQuotation(
    ctx.qID,
    "Approved from Quotation Management Form"
  );

  if (result && result.success) {
    reloadLoadedQuotation_(ctx.qID);
    ui.alert("Quotation approved ✅");
  }
}


function sendQuotationFromForm_() {

  const ctx = getLoadedQuotationContext_();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    "Send Quotation",
    "Mark quotation " + ctx.qID + " / " + ctx.revision + " as Sent?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const result = sendQuotation(
    ctx.qID,
    "Sent from Quotation Management Form"
  );

  if (result && result.success) {
    reloadLoadedQuotation_(ctx.qID);
    ui.alert("Quotation sent ✅");
  }
}


function markQuotationWon_() {

  const ctx = getLoadedQuotationContext_();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    "Mark Quotation Won",
    "Mark quotation " + ctx.qID + " / " + ctx.revision + " as Won?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const result = markQuotationWon(
    ctx.qID,
    "Marked won from Quotation Management Form"
  );

  if (result && result.success) {
    reloadLoadedQuotation_(ctx.qID);
    ui.alert("Quotation marked as Won ✅");
  }
}


function markQuotationLost_() {

  const ctx = getLoadedQuotationContext_();
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    "Mark Quotation Lost",
    "Enter lost reason for " + ctx.qID + " / " + ctx.revision + ":",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const reason = String(response.getResponseText() || "").trim();

  if (!reason) {
    ui.alert("Lost reason is required.");
    return;
  }

  const result = markQuotationLost(
    ctx.qID,
    reason,
    "Marked lost from Quotation Management Form"
  );

  if (result && result.success) {
    reloadLoadedQuotation_(ctx.qID);
    ui.alert("Quotation marked as Lost ✅");
  }
}


function cancelQuotation_() {

  const ctx = getLoadedQuotationContext_();
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    "Cancel Quotation",
    "Enter cancellation reason for " + ctx.qID + " / " + ctx.revision + ":",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const reason = String(response.getResponseText() || "").trim();

  if (!reason) {
    ui.alert("Cancellation reason is required.");
    return;
  }

  const result = cancelQuotation(
    ctx.qID,
    reason,
    "Cancelled from Quotation Management Form"
  );

  if (result && result.success) {
    reloadLoadedQuotation_(ctx.qID);
    ui.alert("Quotation cancelled ✅");
  }
}



function setupInstallableOnEditTrigger() {

  const ss = SpreadsheetApp.getActive();

  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function (trigger) {
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