// @ts-nocheck
/*****************************************************
 EEI REVISION CONTROL
 Revision isolation:
 - Copy Active items only
 - Copy 24 columns
 - Old revision = Superseded
 - New revision = Draft
 - Quotations status = current revision status
*****************************************************/

function createQuotationRevision(qID) {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const quotation = getQuotationById_(qID);

    if (!quotation) {
      throw new Error("Quotation not found.");
    }

    if (quotation.recordStatus === "Archived") {
      throw new Error(
        "Archived quotation cannot create revisions."
      );
    }

    const oldRevision =
      quotation.currentRevision;

    const newRevision =
      generateNextRevisionNo_(oldRevision);

    const revisionsSheet =
      getRequiredSheet_(
        CONFIG.SHEETS.QUOTATION_REVISIONS
      );

    /*****************************************************
     DUPLICATE REVISION PROTECTION
    *****************************************************/

    const revisionID =
      qID + "-" + newRevision;

    if (revisionExists_(revisionID)) {

      SpreadsheetApp
        .getUi()
        .alert(
          "Revision already exists: " +
          revisionID
        );

      return {
        success: false,
        error: "Duplicate revision"
      };
    }

    /*****************************************************
     CREATE NEW REVISION
    *****************************************************/

    const now = new Date();

    const user =
      getCurrentUserName();

    const newStatus =
      CONFIG.QUOTATION_STATUS.DRAFT;

    /*****************************************************
     LOCK OLD REVISION
    *****************************************************/

    markOldRevisionSuperseded_(
      qID,
      oldRevision
    );

    /*****************************************************
     COPY DATA
    *****************************************************/

    copyRevisionItems_(
      qID,
      oldRevision,
      newRevision
    );

    copyRevisionTerms_(
      qID,
      oldRevision,
      newRevision
    );

    /*****************************************************
     CREATE REVISION HEADER
    *****************************************************/

    revisionsSheet
      .getRange(
        getNextDataRow_(revisionsSheet),
        1,
        1,
        20
      )
      .setValues([[
        revisionID,
        qID,
        newRevision,
        newStatus,
        "Revision created from " + oldRevision,
        now,
        user,
        oldRevision,
        "",
        0,
        0,
        0,
        0,
        "",
        "",
        "",
        "",
        "",
        true,
        ""
      ]]);

    /*****************************************************
     UPDATE MASTER QUOTATION
    *****************************************************/

    updateQuotationCurrentRevision_(
      quotation.row,
      newRevision,
      newStatus,
      now,
      user
    );

    /*****************************************************
     LOG
    *****************************************************/

    logQuotationStatus_(
      qID,
      newRevision,
      oldRevision,
      newStatus,
      "Revision",
      "New revision created from " +
      oldRevision
    );

    if (typeof logAction === "function") {

      logAction({
        module: "Quotation Revision",
        action: "CREATE REVISION",
        recordID: revisionID,
        field: "Revision",
        oldValue: oldRevision,
        newValue: newRevision,
        notes:
          "New revision created"
      });
    }

    SpreadsheetApp
      .getUi()
      .alert(
        "Revision created: " +
        newRevision
      );

    return {
      success: true,
      qID: qID,
      revision: newRevision
    };

  } catch (err) {

    SpreadsheetApp
      .getUi()
      .alert(
        "Create Revision Error: " +
        err.message
      );

    Logger.log(err);

    return {
      success: false,
      error: err.message
    };

  } finally {

    lock.releaseLock();

  }
}


/*****************************************************
 REVISION EXISTS
*****************************************************/

function revisionExists_(revisionID) {

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.QUOTATION_REVISIONS
    );

  if (sheet.getLastRow() < 2) {
    return false;
  }

  const ids =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        1
      )
      .getValues()
      .flat();

  return ids.includes(revisionID);
}


/*****************************************************
 MARK OLD REVISION SUPERSEDED
*****************************************************/

function markOldRevisionSuperseded_(qID, oldRevision) {

  const revisions = getRequiredSheet_(
    CONFIG.SHEETS.QUOTATION_REVISIONS
  );

  if (revisions.getLastRow() < 2) return;

  const data = revisions
    .getRange(2, 1, revisions.getLastRow() - 1, revisions.getLastColumn())
    .getValues();

  for (let i = 0; i < data.length; i++) {

    const row = data[i];

    if (
      row[1] === qID &&
      row[2] === oldRevision
    ) {

      const sheetRow = i + 2;

      revisions.getRange(sheetRow, 4)
        .setValue(CONFIG.QUOTATION_STATUS.SUPERSEDED);

      revisions.getRange(sheetRow, 19)
        .setValue(false);

      return;
    }
  }
}


/*****************************************************
 COPY ACTIVE ITEMS ONLY - 24 COLUMNS
*****************************************************/

function copyRevisionItems_(qID, oldRevision, newRevision) {

  const sheet = getRequiredSheet_(
    CONFIG.SHEETS.QUOTATION_ITEMS
  );

  if (sheet.getLastRow() < 2) return;

  const data = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 24)
    .getValues();

  const now = new Date();
  const user = getCurrentUserName();

  const rows = [];

  data.forEach(function(row) {

    const itemQID = row[1];
    const itemRevision = row[2];
    const itemStatus = row[19] || "Active";

    if (
      itemQID === qID &&
      itemRevision === oldRevision &&
      itemStatus !== "Deleted"
    ) {

      const lineNo = Number(row[3]) || rows.length + 1;

      const newRow = row.slice();

      newRow[0] =
        qID + "-" + newRevision + "-L" + ("00" + lineNo).slice(-2);

      newRow[2] = newRevision;
      newRow[15] = now;
      newRow[16] = user;
      newRow[17] = now;
      newRow[18] = user;
      newRow[19] = "Active";
      newRow[20] = "";
      newRow[21] = "";
      newRow[22] = "";
      newRow[23] = "";

      rows.push(newRow);
    }
  });

  if (rows.length) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, rows.length, 24)
      .setValues(rows);
  }
}


/*****************************************************
 COPY TERMS
*****************************************************/

function copyRevisionTerms_(qID, oldRevision, newRevision) {

  const terms = getRequiredSheet_(
    CONFIG.SHEETS.QUOTATION_TERMS
  );

  if (terms.getLastRow() < 2) return;

  const data = terms
    .getRange(2, 1, terms.getLastRow() - 1, terms.getLastColumn())
    .getValues();

  const now = new Date();
  const user = getCurrentUserName();

  for (let i = 0; i < data.length; i++) {

    const row = data[i];

    if (
      row[1] === qID &&
      row[2] === oldRevision
    ) {

      const newRow = row.slice();

      newRow[0] = qID + "-" + newRevision + "-TERMS";
      newRow[2] = newRevision;

      if (newRow.length >= 18) {
        newRow[16] = now;
        newRow[17] = user;
      }

      terms
        .getRange(terms.getLastRow() + 1, 1, 1, newRow.length)
        .setValues([newRow]);

      return;
    }
  }
}


/*****************************************************
 CREATE REVISION HEADER
*****************************************************/

function createRevisionHeader_(
  qID,
  newRevision,
  oldRevision,
  status,
  now,
  user
) {

  const revisions = getRequiredSheet_(
    CONFIG.SHEETS.QUOTATION_REVISIONS
  );

  const revisionID = qID + "-" + newRevision;

  revisions
    .getRange(getNextDataRow_(revisions), 1, 1, 20)
    .setValues([[
      revisionID,
      qID,
      newRevision,
      status,
      "Revision created from " + oldRevision,
      now,
      user,
      oldRevision,
      "",
      0,
      0,
      0,
      0,
      "",
      "",
      "",
      "",
      "",
      true,
      ""
    ]]);
}


/*****************************************************
 UPDATE QUOTATION MASTER
*****************************************************/

function updateQuotationCurrentRevision_(
  quotationRow,
  newRevision,
  newStatus,
  now,
  user
) {

  const quotations = getRequiredSheet_(
    CONFIG.SHEETS.QUOTATIONS
  );

  quotations.getRange(quotationRow, 5).setValue(newRevision);
  quotations.getRange(quotationRow, 6).setValue(newStatus);
  quotations.getRange(quotationRow, 19).setValue(now);
  quotations.getRange(quotationRow, 20).setValue(user);
}


/*****************************************************
 REVISION NUMBER
*****************************************************/

function generateNextRevisionNo_(currentRevision) {

  const n = parseInt(
    String(currentRevision || "R00").replace("R", ""),
    10
  ) || 0;

  return "R" + ("00" + (n + 1)).slice(-2);
}




function getCurrentRevisionNo_(qID) {

  const quotation = getQuotationById_(qID);

  if (!quotation) {
    throw new Error("Quotation not found.");
  }

  return quotation.currentRevision;
}


function isCurrentRevision_(qID, revisionNo) {

  const currentRevision =
    getCurrentRevisionNo_(qID);

  return String(currentRevision) === String(revisionNo);
}


function ensureCurrentRevisionEditable_(qID, revisionNo) {

  const quotation = getQuotationById_(qID);

  if (!quotation) {
    throw new Error("Quotation not found.");
  }

  if (quotation.recordStatus === "Archived") {
    throw new Error("Archived quotation cannot be edited.");
  }

  if (!isCurrentRevision_(qID, revisionNo)) {
    throw new Error(
      "This revision is superseded and cannot be edited."
    );
  }

  if (!canEditQuotation_(quotation.status)) {
    throw new Error(
      "Editing is not allowed for status: " +
      quotation.status
    );
  }

  return true;
}


function ensureCurrentRevisionForWorkflow_(qID, revisionNo) {

  if (!isCurrentRevision_(qID, revisionNo)) {
    throw new Error(
      "Workflow actions are allowed only on the current revision."
    );
  }

  return true;
}