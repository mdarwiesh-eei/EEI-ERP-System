// =====================================
// AUDIT LOG ENGINE
// =====================================

function logAction(data) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    let sheet =
      ss.getSheetByName("Audit_Log");

    // =====================================
    // CREATE SHEET IF NOT EXISTS
    // =====================================

    if (!sheet) {

      sheet = ss.insertSheet("Audit_Log");

      sheet.appendRow([

        "Timestamp",
        "User",
        "Module",
        "Action",
        "Record ID",
        "Field",
        "Old Value",
        "New Value",
        "Notes"

      ]);

    }

    // =====================================
    // INSERT LOG
    // =====================================

    sheet.appendRow([

      new Date(),

      getCurrentUserName(),

      data.module || "",

      data.action || "",

      data.recordID || "",

      data.field || "",

      data.oldValue || "",

      data.newValue || "",

      data.notes || ""

    ]);

  }

  catch(err){

    Logger.log(err);

  }

  finally {

    lock.releaseLock();

  }
}




  // =====================================
  // HELPERS
  // =====================================



function auditCreate(
  module,
  recordID,
  value
  )
  {

  logAction({

    module: module,

    action: "CREATE",

    recordID: recordID,

    newValue: value

  });

}

