// @ts-nocheck

const QUOTATION_WORKFLOW = {

  Draft: [
    "Under Review",
    "Cancelled"
  ],

  "Under Review": [
    "Approved",
    "Draft",
    "Cancelled"
  ],

  Approved: [
    "Sent",
    "Cancelled"
  ],

  Sent: [
    "Negotiation",
    "Won",
    "Lost",
    "Cancelled"
  ],

  Negotiation: [
    "Revised",
    "Won",
    "Lost",
    "Cancelled"
  ],
: [
    "Under Review",
    "Cancelled"
  ],

  Won: [],

  Lost: [],

  Cancelled: [],

  Superseded: []

};




//Status Transition Validation


function validateStatusTransition_(
  currentStatus,
  newStatus
) {

  const allowed =
    QUOTATION_WORKFLOW[currentStatus];

  if (!allowed) {

    throw new Error(
      "Invalid current status: " +
      currentStatus
    );
  }

  if (!allowed.includes(newStatus)) {

    throw new Error(
      "Transition not allowed: " +
      currentStatus +
      " → " +
      newStatus
    );
  }

  return true;
}





//Permission Engine

function canChangeQuotationStatus_(
  currentStatus,
  newStatus
) {

  const role =
    getCurrentUserRole
      ? getCurrentUserRole()
      : "Admin";

  // ===============================
  // ADMIN
  // ===============================

  if (role === "Admin") {
    return true;
  }

  // ===============================
  // MANAGER
  // ===============================

  if (role === "Manager") {

    const allowed = [
      "Under Review",
      "Approved",
      "Sent",
      "Negotiation",
      "Won",
      "Lost"
    ];

    return allowed.includes(newStatus);
  }

  // ===============================
  // SALES
  // ===============================

  if (role === "Sales") {

    const allowed = [
      "Draft",
      "Under Review",
      "Negotiation"
    ];

    return allowed.includes(newStatus);
  }

  return false;
}






//Central Workflow Engine

function changeQuotationStatus_(
  qID,
  newStatus,
  reason,
  notes
) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const quotations =
      getRequiredSheet_(
        CONFIG.SHEETS.QUOTATIONS
      );

    const quotation =
      getQuotationById_(qID);

    if (!quotation) {
      throw new Error("Quotation not found.");
    }

    const currentStatus =
      quotation.status;

    refreshQuotationKPIsFromForm();

    // ===============================
    // VALIDATE TRANSITION
    // ===============================

    validateStatusTransition_(
      currentStatus,
      newStatus
    );

    // ===============================
    // PERMISSION CHECK
    // ===============================

    if (
      !canChangeQuotationStatus_(
        currentStatus,
        newStatus
      )
    ) {

      throw new Error(
        "You do not have permission to change quotation status."
      );
    }

    // ===============================
    // MANDATORY REASONS
    // ===============================

    if (
      newStatus === "Lost" &&
      !reason
    ) {

      throw new Error(
        "Lost reason is required."
      );
    }

    if (
      newStatus === "Cancelled" &&
      !reason
    ) {

      throw new Error(
        "Cancellation reason is required."
      );
    }

    // ===============================
    // UPDATE MASTER STATUS
    // ===============================

    quotations
      .getRange(
        quotation.row,
        6
      )
      .setValue(newStatus);

    quotations
      .getRange(
        quotation.row,
        19
      )
      .setValue(new Date());

    quotations
      .getRange(
        quotation.row,
        20
      )
      .setValue(getCurrentUserName());

    refreshQuotationKPIsFromForm();

    // ===============================
    // UPDATE CURRENT REVISION STATUS
    // ===============================

    updateRevisionStatus_(
      qID,
      quotation.currentRevision,
      newStatus
    );

    syncQuotationStatusWithCurrentRevision_(qID);

    refreshQuotationKPIsFromForm();

    // ===============================
    // STATUS LOG
    // ===============================

    logQuotationStatus_(
      qID,
      quotation.currentRevision,
      currentStatus,
      newStatus,
      reason || "",
      notes || ""
    );

    // ===============================
    // AUDIT LOG
    // ===============================

    if (typeof logAction === "function") {

      logAction({

        module: "Quotations",

        action: "STATUS_CHANGE",

        recordID: qID,

        field: "Status",

        oldValue: currentStatus,

        newValue: newStatus,

        notes:
          "Reason: " +
          (reason || "") +
          " | Notes: " +
          (notes || "")

      });

    }

    return {
      success: true,
      qID: qID,
      oldStatus: currentStatus,
      newStatus: newStatus
    };

  } catch (err) {

    SpreadsheetApp.getUi().alert(
      "Status Change Error: " +
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






//Update Revision Status

function updateRevisionStatus_(
  qID,
  revisionNo,
  newStatus
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

      sheet
        .getRange(row, 4)
        .setValue(newStatus);

      return;
    }
  }

  refreshQuotationKPIsFromForm();

}




//Wrapper Functions
//Submit forreview

function submitQuotationForReview(
  qID
) {

  return changeQuotationStatus_(
    qID,
    "Under Review",
    "",
    "Submitted for technical/commercial review"
  );

  refreshQuotationKPIsFromForm();

}



//Approve

function approveQuotation(
  qID,
  approvalNotes
) {

  return changeQuotationStatus_(
    qID,
    "Approved",
    "Quotation approved",
    approvalNotes || ""
  );

  refreshQuotationKPIsFromForm();

}





//Reject Back To Draft

function rejectQuotation(
  qID,
  rejectionReason
) {

  return changeQuotationStatus_(
    qID,
    "Draft",
    rejectionReason || "Rejected during review",
    ""
  );

  refreshQuotationKPIsFromForm();

}




//send

function sendQuotation(
  qID,
  notes
) {

  return changeQuotationStatus_(
    qID,
    "Sent",
    "Quotation sent to customer",
    notes || ""
  );
}



//Move To Negotiation

function moveQuotationToNegotiation(
  qID,
  notes
) {

  return changeQuotationStatus_(
    qID,
    "Negotiation",
    "Customer negotiation started",
    notes || ""
  );

  refreshQuotationKPIsFromForm();

}





//Mark Won

function markQuotationWon(
  qID,
  notes
) {

  return changeQuotationStatus_(
    qID,
    "Won",
    "Quotation awarded",
    notes || ""
  );

  refreshQuotationKPIsFromForm();

}



//Mark Lost

function markQuotationLost(
  qID,
  reason,
  notes
) {

  return changeQuotationStatus_(
    qID,
    "Lost",
    reason,
    notes || ""
  );

  refreshQuotationKPIsFromForm();

}





//Cancel

function cancelQuotation(
  qID,
  reason,
  notes
) {

  return changeQuotationStatus_(
    qID,
    "Cancelled",
    reason,
    notes || ""
  );

  refreshQuotationKPIsFromForm();

}





function syncQuotationStatusWithCurrentRevision_(qID) {

  const quotations =
    getRequiredSheet_(
      CONFIG.SHEETS.QUOTATIONS
    );

  const revisions =
    getRequiredSheet_(
      CONFIG.SHEETS.QUOTATION_REVISIONS
    );

  const quotation =
    getQuotationById_(qID);

  if (!quotation) {
    throw new Error("Quotation not found.");
  }

  if (revisions.getLastRow() < 2) {
    return;
  }

  const data =
    revisions
      .getRange(
        2,
        1,
        revisions.getLastRow() - 1,
        revisions.getLastColumn()
      )
      .getValues();

  let currentRevisionStatus = "";

  data.forEach(function (row) {

    const rowQID = row[1];
    const revisionNo = row[2];
    const revisionStatus = row[3];
    const isCurrent = row[18];

    if (
      rowQID === qID &&
      revisionNo === quotation.currentRevision &&
      isCurrent === true
    ) {
      currentRevisionStatus = revisionStatus;
    }

  });

  if (!currentRevisionStatus) {
    return;
  }

  quotations
    .getRange(
      quotation.row,
      6
    )
    .setValue(
      currentRevisionStatus
    );

  quotations
    .getRange(
      quotation.row,
      19
    )
    .setValue(
      new Date()
    );

  quotations
    .getRange(
      quotation.row,
      20
    )
    .setValue(
      getCurrentUserName()
    );
}