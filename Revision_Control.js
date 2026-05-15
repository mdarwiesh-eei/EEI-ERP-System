
//--------------------------------//
//-----UPDATE REVISIONS TOTALS----//
//--------------------------------//

function updateRevisionTotals_(
  qID,
  revisionNo,
  subTotal,
  discountPercent,
  vatPercent,
  grandTotal
) {

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.QUOTATION_REVISIONS
    );

  if (sheet.getLastRow() < 2) return;

  const data =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        20
      )
      .getValues();

  for (let i = 0; i < data.length; i++) {

    if (
      data[i][1] === qID &&
      data[i][2] === revisionNo
    ) {

      const row = i + 2;

      sheet.getRange(row, 10)
        .setValue(subTotal);

      sheet.getRange(row, 11)
        .setValue(discountPercent);

      sheet.getRange(row, 12)
        .setValue(vatPercent);

      sheet.getRange(row, 13)
        .setValue(grandTotal);

      return;
    }
  }
}




function createQuotationRevision(qID, revisionReason) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    if (!qID) {
      throw new Error("Quotation ID is required.");
    }

    if (!revisionReason) {
      throw new Error("Revision reason is required.");
    }

    const quotation =
      getQuotationById_(qID);

    if (!quotation) {
      throw new Error("Quotation not found.");
    }

    const allowedStatuses = [
      CONFIG.QUOTATION_STATUS.SENT,
      CONFIG.QUOTATION_STATUS.NEGOTIATION
    ];

    if (!allowedStatuses.includes(quotation.status)) {
      throw new Error(
        "Revision can be created only from Sent or Negotiation status."
      );
    }

    const oldRevision =
      quotation.currentRevision;

    const newRevision =
      getNextRevisionNo_(qID);

    const newRevisionID =
      qID + "-" + newRevision;

    const now =
      new Date();

    const user =
      getCurrentUserName();

    const quotationsSheet =
      getRequiredSheet_(CONFIG.SHEETS.QUOTATIONS);

    const revisionsSheet =
      getRequiredSheet_(CONFIG.SHEETS.QUOTATION_REVISIONS);

    // ===============================
    // CREATE REVISION FOLDER
    // ===============================

    const revisionFolderLink =
      createRevisionFolder_(
        quotation.folderLink,
        newRevision
      );

    // ===============================
    // MARK OLD REVISIONS NOT CURRENT
    // ===============================

    markOldRevisionsAsNotCurrent_(qID);

    updateRevisionStatus_(
      qID,
      oldRevision,
      CONFIG.QUOTATION_STATUS.SUPERSEDED
    );

    // ===============================
    // COPY OLD TERMS
    // ===============================

    const oldTerms =
      getQuotationTerms_(
        qID,
        oldRevision
      );

    // ===============================
    // CREATE NEW REVISION RECORD
    // ===============================

    const newRevisionRow = [[

      newRevisionID,
      qID,
      newRevision,
      CONFIG.QUOTATION_STATUS.REVISED,
      revisionReason,
      now,
      user,
      "",
      "",
      oldTerms ? oldTerms.subTotal : 0,
      oldTerms ? oldTerms.discountPercent / 100 : 0,
      oldTerms ? oldTerms.vatPercent / 100 : 0,
      oldTerms ? oldTerms.grandTotal : 0,
      revisionFolderLink,
      "",
      "",
      "",
      "",
      true,
      "Created from " + oldRevision

    ]];

    const nextRevRow =
      revisionsSheet.getLastRow() < 2
        ? 2
        : revisionsSheet.getLastRow() + 1;

    revisionsSheet
      .getRange(
        nextRevRow,
        1,
        1,
        20
      )
      .setValues(newRevisionRow);

    // ===============================
    // COPY ITEMS FROM OLD REVISION
    // ===============================

    copyQuotationItemsToNewRevision_(
      qID,
      oldRevision,
      newRevision
    );

    // ===============================
    // COPY TERMS FROM OLD REVISION
    // ===============================

    copyQuotationTermsToNewRevision_(
      qID,
      oldRevision,
      newRevision
    );

    // ===============================
    // UPDATE MASTER
    // ===============================

    quotationsSheet
      .getRange(
        quotation.row,
        5
      )
      .setValue(newRevision);

    quotationsSheet
      .getRange(
        quotation.row,
        6
      )
      .setValue(CONFIG.QUOTATION_STATUS.REVISED);

    quotationsSheet
      .getRange(
        quotation.row,
        19
      )
      .setValue(now);

    quotationsSheet
      .getRange(
        quotation.row,
        20
      )
      .setValue(user);

    // ===============================
    // RECALCULATE TOTALS
    // ===============================

    updateQuotationTotals(
      qID,
      newRevision
    );

    // ===============================
    // STATUS LOG
    // ===============================

    logQuotationStatus_(
      qID,
      newRevision,
      quotation.status,
      CONFIG.QUOTATION_STATUS.REVISED,
      revisionReason,
      "New revision created from " + oldRevision
    );

    // ===============================
    // AUDIT LOG
    // ===============================

    if (typeof logAction === "function") {

      logAction({

        module: "Quotations",

        action: "CREATE_REVISION",

        recordID: qID,

        field: "Revision",

        oldValue: oldRevision,

        newValue: newRevision,

        notes:
          revisionReason +
          " | Folder: " +
          revisionFolderLink

      });

    }

    return {

      success: true,

      qID: qID,

      oldRevision: oldRevision,

      revisionNo: newRevision,

      revisionFolderLink: revisionFolderLink

    };

  }

  catch (err) {

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

  }

  finally {

    lock.releaseLock();

  }
}





function createRevisionFolder_(
  quotationFolderLink,
  revisionNo
) {

  if (!quotationFolderLink) {
    throw new Error("Quotation folder link is missing.");
  }

  const quotationFolderId =
    extractIdFromUrl(quotationFolderLink);

  if (!quotationFolderId) {
    throw new Error("Invalid quotation folder link.");
  }

  const quotationFolder =
    DriveApp.getFolderById(
      quotationFolderId
    );

  let revisionFolder =
    findChildFolderByName_(
      quotationFolder,
      revisionNo
    );

  if (!revisionFolder) {

    revisionFolder =
      quotationFolder.createFolder(
        revisionNo
      );

  }

  const subFolders = [
    "01 RFQ",
    "02 Technical Offer",
    "03 Commercial Offer",
    "04 Client Comments",
    "05 Revisions"
  ];

  subFolders.forEach(function(folderName) {

    let child =
      findChildFolderByName_(
        revisionFolder,
        folderName
      );

    if (!child) {
      revisionFolder.createFolder(folderName);
    }

  });

  return revisionFolder.getUrl();
}





function getNextRevisionNo_(qID) {

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.QUOTATION_REVISIONS
    );

  if (sheet.getLastRow() < 2) {
    return CONFIG.REVISION.INITIAL_REVISION;
  }

  const data =
    sheet
      .getRange(
        2,
        2,
        sheet.getLastRow() - 1,
        2
      )
      .getValues();

  let maxRev = -1;

  data.forEach(function(row) {

    const existingQID =
      row[0];

    const revisionNo =
      row[1];

    if (existingQID === qID) {

      const match =
        revisionNo
          .toString()
          .match(/R(\d+)/);

      if (match) {

        maxRev =
          Math.max(
            maxRev,
            Number(match[1])
          );

      }
    }
  });

  const next =
    maxRev + 1;

  return "R" +
    ("00" + next).slice(-2);
}







function markOldRevisionsAsNotCurrent_(qID) {

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.QUOTATION_REVISIONS
    );

  if (sheet.getLastRow() < 2) return;

  const data =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        20
      )
      .getValues();

  for (let i = 0; i < data.length; i++) {

    if (data[i][1] === qID) {

      const row =
        i + 2;

      sheet
        .getRange(row, 19)
        .setValue(false);
    }
  }
}





function copyQuotationItemsToNewRevision_(
  qID,
  oldRevision,
  newRevision
) {

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.QUOTATION_ITEMS
    );

  if (sheet.getLastRow() < 2) return;

  const data =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        19
      )
      .getValues();

  const rowsToAdd = [];

  data.forEach(function(row) {

    if (
      row[1] === qID &&
      row[2] === oldRevision
    ) {

      const lineNo =
        row[3];

      const newItemID =
        qID +
        "-" +
        newRevision +
        "-L" +
        ("00" + lineNo).slice(-2);

      rowsToAdd.push([

        newItemID,
        qID,
        newRevision,
        lineNo,
        row[4],
        row[5],
        row[6],
        row[7],
        row[8],
        row[9],
        row[10],
        row[11],
        row[12],
        row[13],
        row[14],
        new Date(),
        getCurrentUserName(),
        new Date(),
        getCurrentUserName()

      ]);
    }
  });

  if (rowsToAdd.length) {

    const nextRow =
      sheet.getLastRow() + 1;

    sheet
      .getRange(
        nextRow,
        1,
        rowsToAdd.length,
        19
      )
      .setValues(rowsToAdd);
  }
}





function copyQuotationTermsToNewRevision_(
  qID,
  oldRevision,
  newRevision
) {

  const sheet =
    getRequiredSheet_(
      CONFIG.SHEETS.QUOTATION_TERMS
    );

  const oldTerms =
    getQuotationTerms_(
      qID,
      oldRevision
    );

  if (!oldTerms) {

    createInitialQuotationTerms(
      qID,
      newRevision,
      CONFIG.CURRENCY.EGP
    );

    return;
  }

  const now =
    new Date();

  const user =
    getCurrentUserName();

  const newTermsID =
    qID +
    "-" +
    newRevision +
    "-TERMS";

  const rowData = [[

    newTermsID,
    qID,
    newRevision,
    oldTerms.currency,
    oldTerms.subTotal,
    oldTerms.discountPercent / 100,
    oldTerms.vatPercent / 100,
    oldTerms.advancePercent / 100,
    oldTerms.advanceAmount,
    oldTerms.remainingAmount,
    oldTerms.grandTotal,
    oldTerms.deliveryTerms,
    oldTerms.paymentTerms,
    oldTerms.validityDays,
    now,
    user,
    now,
    user,
    "Copied from " + oldRevision

  ]];

  const nextRow =
    sheet.getLastRow() < 2
      ? 2
      : sheet.getLastRow() + 1;

  sheet
    .getRange(
      nextRow,
      1,
      1,
      19
    )
    .setValues(rowData);
}





