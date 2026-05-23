// @ts-nocheck
/*****************************************************
 EEI - QUOTATION BACKEND
 Master Quotation + Items + Totals
 Updated for:
 - Item_Status
 - Soft delete readiness
 - Duplicate item prevention
 - Record_Status in Quotations
*****************************************************/


/*****************************************************
 CREATE QUOTATION
*****************************************************/

function createQuotation(data) {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    validateCreateQuotationInput_(data);

    const quotations = getRequiredSheet_(CONFIG.SHEETS.QUOTATIONS);
    const revisions = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_REVISIONS);

    const customer = getCustomerById_(data.customerID);

    if (!customer) {
      throw new Error("Customer not found.");
    }

    const qID = generateQuotationId_();
    const revisionNo = CONFIG.REVISION.INITIAL_REVISION;
    const revisionID = qID + "-" + revisionNo;
    const status = CONFIG.QUOTATION_STATUS.DRAFT;

    const folderLink = createQuotationProjectFolder_(
      customer.folderLink,
      qID,
      data.projectName
    );

    const now = new Date();
    const user =
      data.createdBy ||
      data.assignedTo ||
      getCurrentUserName();

    const quotationRow = [[
      qID,
      data.customerID,
      customer.name,
      data.projectName,
      revisionNo,
      status,
      data.rfqDate || "",
      now,
      user,
      data.assignedTo || user,
      data.currency || CONFIG.CURRENCY.EGP,
      0,
      0,
      0,
      0,
      folderLink,
      data.customerRFQLink || "",
      data.notes || "",
      now,
      user,
      "Active"
    ]];

    quotations
      .getRange(getNextDataRow_(quotations), 1, 1, 21)
      .setValues(quotationRow);

    const revisionRow = [[
      revisionID,
      qID,
      revisionNo,
      status,
      "Initial quotation revision",
      now,
      user,
      "",
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
      data.notes || ""
    ]];

    revisions
      .getRange(getNextDataRow_(revisions), 1, 1, 20)
      .setValues(revisionRow);

    createInitialQuotationTerms(
      qID,
      revisionNo,
      data.currency || CONFIG.CURRENCY.EGP
    );

    logQuotationStatus_(
      qID,
      revisionNo,
      "",
      status,
      "Quotation created",
      "Initial draft created"
    );

    if (typeof logAction === "function") {
      logAction({
        module: "Quotations",
        action: "CREATE",
        recordID: qID,
        field: "Quotation",
        oldValue: "",
        newValue: data.projectName,
        notes: "Quotation created with initial revision " + revisionNo
      });
    }

    return {
      success: true,
      qID: qID,
      revisionNo: revisionNo,
      folderLink: folderLink
    };

  } catch (err) {

    SpreadsheetApp.getUi().alert("Create Quotation Error: " + err.message);
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
 REQUIRED SHEET
*****************************************************/

function getRequiredSheet_(sheetName) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetName);

  if (!sheet) {
    throw new Error("Required sheet not found: " + sheetName);
  }

  return sheet;
}


function getNextDataRow_(sheet) {
  return sheet.getLastRow() < 2 ? 2 : sheet.getLastRow() + 1;
}


/*****************************************************
 GENERATE QUOTATION ID
*****************************************************/

function generateQuotationId_() {

  const sheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATIONS);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return "Q-0001";
  }

  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .filter(String);

  let maxNo = 0;

  ids.forEach(function (id) {

    const match = id.toString().match(/Q-(\d+)/);

    if (match) {
      maxNo = Math.max(maxNo, Number(match[1]));
    }

  });

  return "Q-" + ("0000" + (maxNo + 1)).slice(-4);
}


/*****************************************************
 GET CUSTOMER BY ID
*****************************************************/

function getCustomerById_(customerID) {

  const sheet = getRequiredSheet_(CONFIG.SHEETS.CUSTOMERS);

  if (sheet.getLastRow() < 3) return null;

  const data = sheet
    .getRange(3, 1, sheet.getLastRow() - 2, 14)
    .getValues();

  for (let i = 0; i < data.length; i++) {

    if (
      data[i][0] === customerID &&
      data[i][10] === "Active"
    ) {

      return {
        id: data[i][0],
        name: data[i][1],
        folderLink: data[i][9],
        status: data[i][10]
      };

    }
  }

  return null;
}


/*****************************************************
 CREATE QUOTATION PROJECT FOLDER
*****************************************************/

function createQuotationProjectFolder_(customerFolderLink, qID, projectName) {

  if (!customerFolderLink) {
    throw new Error("Customer folder link is missing.");
  }

  const customerFolderId = extractIdFromUrl(customerFolderLink);

  if (!customerFolderId) {
    throw new Error("Invalid customer folder link.");
  }

  const customerFolder = DriveApp.getFolderById(customerFolderId);

  let quotationsFolder = findChildFolderByName_(
    customerFolder,
    "01 Quotations"
  );

  if (!quotationsFolder) {
    quotationsFolder = customerFolder.createFolder("01 Quotations");
  }

  const quotationFolderName = qID + " - " + projectName;
  const quotationFolder = quotationsFolder.createFolder(quotationFolderName);

  const revisionFolder = quotationFolder.createFolder("R00");

  revisionFolder.createFolder("01 RFQ");
  revisionFolder.createFolder("02 Technical Offer");
  revisionFolder.createFolder("03 Commercial Offer");
  revisionFolder.createFolder("04 Client Comments");
  revisionFolder.createFolder("05 Revisions");

  return quotationFolder.getUrl();
}


function findChildFolderByName_(parentFolder, folderName) {

  const folders = parentFolder.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return null;
}


/*****************************************************
 STATUS LOG
*****************************************************/

function logQuotationStatus_(qID, revisionNo, oldStatus, newStatus, reason, notes) {

  const sheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_STATUS_LOG);
  const nextRow = getNextDataRow_(sheet);
  const logID = "QSL-" + ("0000" + (nextRow - 1)).slice(-4);

  sheet
    .getRange(nextRow, 1, 1, 9)
    .setValues([[
      logID,
      qID,
      revisionNo,
      oldStatus,
      newStatus,
      getCurrentUserName(),
      new Date(),
      reason || "",
      notes || ""
    ]]);
}


/*****************************************************
 ADD SINGLE QUOTATION ITEM
*****************************************************/

function addQuotationItem(data) {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    if (typeof validateQuotationDatabaseStructure_ === "function") {
      validateQuotationDatabaseStructure_();
    }

    validateAddQuotationItemInput_(data);

    const itemsSheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_ITEMS);
    const quotation = getQuotationById_(data.qID);

    if (!quotation) {
      throw new Error("Quotation not found.");
    }

    if (!canEditQuotation_(quotation.status)) {
      throw new Error("Editing is not allowed for status: " + quotation.status);
    }

    if (quotation.recordStatus && quotation.recordStatus === "Archived") {
      throw new Error("Archived quotation cannot be edited.");
    }

    const revisionNo = data.revisionNo || quotation.currentRevision;

    if (isDuplicateQuotationItem_(data.qID, revisionNo, data)) {
      throw new Error("This item already exists in this quotation revision.");
    }

    const lineNo = getNextQuotationLineNo_(data.qID, revisionNo);
    const itemID =
      data.qID + "-" + revisionNo + "-L" + ("00" + lineNo).slice(-2);

    const quantity = Number(data.quantity);
    const unitPrice = Number(data.unitPrice);
    const totalPrice = quantity * unitPrice;

    const now = new Date();
    const user = getCurrentUserName();

    const rowData = [[
      itemID,
      data.qID,
      revisionNo,
      lineNo,
      data.description,
      data.transformerType || "",
      data.powerKVA || "",
      data.voltage || "",
      quantity,
      unitPrice,
      totalPrice,
      data.currency || quotation.currency || CONFIG.CURRENCY.EGP,
      data.deliveryTime || "",
      data.warranty || "",
      data.notes || "",
      now,
      user,
      now,
      user,
      "Active",
      "",
      "",
      "",
      ""
    ]];

    const nextRow = getNextDataRow_(itemsSheet);

    itemsSheet
      .getRange(nextRow, 1, 1, 24)
      .setValues(rowData);

    updateQuotationTotals(data.qID, revisionNo);

    if (typeof logAction === "function") {
      logAction({
        module: "Quotations",
        action: "ADD_ITEM",
        recordID: itemID,
        field: "Quotation Item",
        oldValue: "",
        newValue:
          data.description +
          " | Qty: " +
          quantity +
          " | Unit Price: " +
          unitPrice +
          " | Total: " +
          totalPrice,
        notes: "Item added to " + data.qID + " revision " + revisionNo
      });
    }

    return {
      success: true,
      itemID: itemID,
      qID: data.qID,
      revisionNo: revisionNo
    };

  } catch (err) {

    SpreadsheetApp
      .getUi()
      .alert("Add Quotation Item Error: " + err.message);

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
 ADD ITEM VALIDATION
*****************************************************/

function validateAddQuotationItemInput_(data) {

  if (!data) {
    throw new Error("Missing item data.");
  }

  if (!data.qID) {
    throw new Error("Quotation ID is required.");
  }

  if (!data.description) {
    throw new Error("Item description is required.");
  }

  if (!data.quantity || isNaN(Number(data.quantity))) {
    throw new Error("Valid quantity is required.");
  }

  if (Number(data.quantity) <= 0) {
    throw new Error("Quantity must be greater than zero.");
  }

  if (!data.unitPrice || isNaN(Number(data.unitPrice))) {
    throw new Error("Valid unit price is required.");
  }

  if (Number(data.unitPrice) <= 0) {
    throw new Error("Unit price must be greater than zero.");
  }
}


/*****************************************************
 ADD ITEMS BATCH
*****************************************************/

function addQuotationItemsBatch(data) {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    if (typeof validateQuotationDatabaseStructure_ === "function") {
      validateQuotationDatabaseStructure_();
    }

    validateAddQuotationItemsBatchInput_(data);

    const itemsSheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_ITEMS);
    const quotation = getQuotationById_(data.qID);

    if (!quotation) {
      throw new Error("Quotation not found.");
    }

    if (!canEditQuotation_(quotation.status)) {
      throw new Error("Editing is not allowed for status: " + quotation.status);
    }

    if (quotation.recordStatus && quotation.recordStatus === "Archived") {
      throw new Error("Archived quotation cannot be edited.");
    }

    const revisionNo = data.revisionNo || quotation.currentRevision;

    const now = new Date();
    const user = getCurrentUserName();

    let maxLine = getCurrentMaxQuotationLineNo_(data.qID, revisionNo);

    const rowsToInsert = [];
    const batchKeys = {};

    data.items.forEach(function (item, index) {

      const quantity = Number(item.quantity);
      const unitPrice = Number(item.unitPrice);

      if (
        !item.description ||
        isNaN(quantity) ||
        quantity <= 0 ||
        isNaN(unitPrice) ||
        unitPrice <= 0
      ) {
        throw new Error("Invalid item data at line " + (index + 1));
      }

      const itemKey = buildQuotationItemKey_(
        data.qID,
        revisionNo,
        item
      );

      if (batchKeys[itemKey]) {
        throw new Error("Duplicate item inside current batch: " + item.description);
      }

      batchKeys[itemKey] = true;

      if (isDuplicateQuotationItem_(data.qID, revisionNo, item)) {
        throw new Error("Duplicate item already exists: " + item.description);
      }

      const lineNo = maxLine + rowsToInsert.length + 1;
      const itemID =
        data.qID + "-" + revisionNo + "-L" + ("00" + lineNo).slice(-2);

      const totalPrice = quantity * unitPrice;

      rowsToInsert.push([
        itemID,
        data.qID,
        revisionNo,
        lineNo,
        item.description,
        item.transformerType || "",
        item.powerKVA || "",
        item.voltage || "",
        quantity,
        unitPrice,
        totalPrice,
        item.currency || quotation.currency || CONFIG.CURRENCY.EGP,
        item.deliveryTime || "",
        item.warranty || "",
        item.notes || "",
        now,
        user,
        now,
        user,
        "Active",
        "",
        "",
        "",
        ""
      ]);

    });

    if (!rowsToInsert.length) {
      throw new Error("No valid items to add.");
    }

    const nextRow = getNextDataRow_(itemsSheet);

    itemsSheet
      .getRange(nextRow, 1, rowsToInsert.length, 24)
      .setValues(rowsToInsert);

    updateQuotationTotals(data.qID, revisionNo);

    if (typeof logAction === "function") {
      logAction({
        module: "Quotations",
        action: "ADD_ITEMS_BATCH",
        recordID: data.qID + "-" + revisionNo,
        field: "Quotation Items",
        oldValue: "",
        newValue: rowsToInsert.length + " item(s) added",
        notes: "Batch add items to " + data.qID + " revision " + revisionNo
      });
    }

    return {
      success: true,
      qID: data.qID,
      revisionNo: revisionNo,
      addedCount: rowsToInsert.length
    };

  } catch (err) {

    SpreadsheetApp
      .getUi()
      .alert("Add Quotation Items Batch Error: " + err.message);

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
 ADD ITEMS BATCH VALIDATION
*****************************************************/

function validateAddQuotationItemsBatchInput_(data) {

  if (!data) {
    throw new Error("Missing batch item data.");
  }

  if (!data.qID) {
    throw new Error("Quotation ID is required.");
  }

  if (
    !data.items ||
    !Array.isArray(data.items) ||
    data.items.length === 0
  ) {
    throw new Error("No items provided.");
  }
}


/*****************************************************
 DUPLICATE ITEM CHECK
*****************************************************/

function isDuplicateQuotationItem_(qID, revisionNo, item) {

  const sheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_ITEMS);

  if (sheet.getLastRow() < 2) return false;

  const data = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .getValues();

  const newKey = buildQuotationItemKey_(qID, revisionNo, item);

  for (let i = 0; i < data.length; i++) {

    const row = data[i];

    const itemStatus = row[19] ? row[19].toString().trim() : "Active";

    if (itemStatus === "Deleted") continue;

    const existingItem = {
      transformerType: row[5] || "",
      powerKVA: row[6] || "",
      voltage: row[7] || ""
    };

    const existingKey = buildQuotationItemKey_(
      row[1],
      row[2],
      existingItem
    );

    if (existingKey === newKey) {
      return true;
    }
  }

  return false;
}


function buildQuotationItemKey_(qID, revisionNo, item) {

  return [
    qID || "",
    revisionNo || "",
    item.transformerType || ""
  ]
    .join("|")
    .toString()
    .toLowerCase()
    .replace(/\s+/g, "")
    .trim();
}


/*****************************************************
 GET QUOTATION BY ID
*****************************************************/

function getQuotationById_(qID) {

  const sheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATIONS);

  if (!qID) return null;
  if (sheet.getLastRow() < 2) return null;

  const lastColumn = Math.max(sheet.getLastColumn(), 21);

  const data = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, lastColumn)
    .getValues();

  for (let i = 0; i < data.length; i++) {

    if (data[i][0] === qID) {

      return {
        row: i + 2,
        qID: data[i][0],
        customerID: data[i][1],
        customerName: data[i][2],
        projectName: data[i][3],
        currentRevision: data[i][4],
        status: data[i][5],
        rfqDate: data[i][6],
        createdDate: data[i][7],
        createdBy: data[i][8],
        assignedTo: data[i][9],
        currency: data[i][10],
        subTotal: data[i][11],
        discountPercent: data[i][12],
        vatPercent: data[i][13],
        grandTotal: data[i][14],
        folderLink: data[i][15],
        customerRFQLink: data[i][16],
        notes: data[i][17],
        lastUpdated: data[i][18],
        lastUpdatedBy: data[i][19],
        recordStatus: data[i][20] || "Active"
      };

    }
  }

  return null;
}


/*****************************************************
 EDIT PERMISSION BY STATUS
*****************************************************/

function canEditQuotation_(status) {

  const lockedStatuses = [
    CONFIG.QUOTATION_STATUS.SENT,
    CONFIG.QUOTATION_STATUS.WON,
    CONFIG.QUOTATION_STATUS.LOST,
    CONFIG.QUOTATION_STATUS.CANCELLED,
    CONFIG.QUOTATION_STATUS.SUPERSEDED
  ];

  return !lockedStatuses.includes(status);
}


/*****************************************************
 LINE NUMBER
*****************************************************/

function getNextQuotationLineNo_(qID, revisionNo) {
  return getCurrentMaxQuotationLineNo_(qID, revisionNo) + 1;
}


function getCurrentMaxQuotationLineNo_(qID, revisionNo) {

  const sheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_ITEMS);

  if (sheet.getLastRow() < 2) {
    return 0;
  }

  const data = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .getValues();

  let maxLine = 0;

  data.forEach(function (row) {

    const existingQID = row[1];
    const existingRevision = row[2];
    const lineNo = Number(row[3]) || 0;

    if (
      existingQID === qID &&
      existingRevision === revisionNo
    ) {
      maxLine = Math.max(maxLine, lineNo);
    }

  });

  return maxLine;
}


/*****************************************************
 UPDATE QUOTATION TOTALS
*****************************************************/

function updateQuotationTotals(qID, revisionNo) {

  const quotations = getRequiredSheet_(CONFIG.SHEETS.QUOTATIONS);
  const termsSheet = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_TERMS);
  const items = getRequiredSheet_(CONFIG.SHEETS.QUOTATION_ITEMS);

  const quotation = getQuotationById_(qID);

  if (!quotation) {
    throw new Error("Quotation not found.");
  }

  revisionNo = revisionNo || quotation.currentRevision;

  let terms = getQuotationTerms_(qID, revisionNo);

  if (!terms) {
    createInitialQuotationTerms(
      qID,
      revisionNo,
      quotation.currency || CONFIG.CURRENCY.EGP
    );

    terms = getQuotationTerms_(qID, revisionNo);
  }

  let subTotal = 0;

  if (items.getLastRow() >= 2) {

    const itemsData = items
      .getRange(2, 1, items.getLastRow() - 1, items.getLastColumn())
      .getValues();

    itemsData.forEach(function (row) {

      const itemQID = row[1];
      const itemRevision = row[2];
      const totalPrice = Number(row[10]) || 0;
      const itemStatus = row[19] ? row[19].toString().trim() : "Active";

      if (
        itemQID === qID &&
        itemRevision === revisionNo &&
        itemStatus !== "Deleted"
      ) {
        subTotal += totalPrice;
      }

    });
  }

  const discountPercent = terms.discountPercent || 0;
  const vatPercent = terms.vatPercent || 0;
  const advancePercent = terms.advancePercent || 0;

  const discountAmount = subTotal * discountPercent / 100;
  const afterDiscount = subTotal - discountAmount;
  const vatAmount = afterDiscount * vatPercent / 100;
  const grandTotal = afterDiscount + vatAmount;

  const advanceAmount = grandTotal * advancePercent / 100;
  const remainingAmount = grandTotal - advanceAmount;

  const now = new Date();
  const user = getCurrentUserName();

  termsSheet.getRange(terms.row, 5).setValue(subTotal);
  termsSheet.getRange(terms.row, 9).setValue(advanceAmount);
  termsSheet.getRange(terms.row, 10).setValue(remainingAmount);
  termsSheet.getRange(terms.row, 11).setValue(grandTotal);
  termsSheet.getRange(terms.row, 17).setValue(now);
  termsSheet.getRange(terms.row, 18).setValue(user);

  quotations.getRange(quotation.row, 11).setValue(terms.currency);
  quotations.getRange(quotation.row, 12).setValue(subTotal);
  quotations.getRange(quotation.row, 13).setValue(discountPercent / 100);
  quotations.getRange(quotation.row, 14).setValue(vatPercent / 100);
  quotations.getRange(quotation.row, 15).setValue(grandTotal);
  quotations.getRange(quotation.row, 19).setValue(now);
  quotations.getRange(quotation.row, 20).setValue(user);

  updateRevisionTotals_(
    qID,
    revisionNo,
    subTotal,
    discountPercent,
    vatPercent,
    grandTotal
  );

  return {
    subTotal: subTotal,
    discountPercent: discountPercent,
    vatPercent: vatPercent,
    advancePercent: advancePercent,
    advanceAmount: advanceAmount,
    remainingAmount: remainingAmount,
    grandTotal: grandTotal
  };
}


/*****************************************************
 ARCHIVE QUOTATION
*****************************************************/

function archiveQuotation_(qID) {

  const quotations = getRequiredSheet_(CONFIG.SHEETS.QUOTATIONS);
  const quotation = getQuotationById_(qID);

  if (!quotation) {
    throw new Error("Quotation not found.");
  }

  quotations.getRange(quotation.row, 21).setValue("Archived");
  quotations.getRange(quotation.row, 19).setValue(new Date());
  quotations.getRange(quotation.row, 20).setValue(getCurrentUserName());

  if (typeof logAction === "function") {
    logAction({
      module: "Quotations",
      action: "ARCHIVE",
      recordID: qID,
      field: "Record_Status",
      oldValue: quotation.recordStatus || "Active",
      newValue: "Archived",
      notes: "Quotation archived"
    });
  }

  return {
    success: true,
    qID: qID
  };
}


/*****************************************************
 TEST FUNCTIONS
*****************************************************/

function testCreateQuotation() {

  const result = createQuotation({
    customerID: "C-0001",
    projectName: "Test Transformer Project",
    rfqDate: new Date(),
    customerRFQLink: "",
    assignedTo: getCurrentUserName(),
    currency: CONFIG.CURRENCY.EGP,
    notes: "Backend test quotation"
  });

  Logger.log(result);
}


function testAddQuotationItem() {

  const result = addQuotationItem({
    qID: "Q-0001",
    description: "Oil Immersed Distribution Transformer",
    transformerType: "Distribution Transformer",
    powerKVA: "500",
    voltage: "11/0.4 kV",
    quantity: 2,
    unitPrice: 250000,
    currency: CONFIG.CURRENCY.EGP,
    deliveryTime: "8 Weeks",
    warranty: "12 Months",
    notes: "Backend item test"
  });

  Logger.log(result);
}




function updateRevisionTotals_(
  qID,
  revisionNo,
  subTotal,
  discountPercent,
  vatPercent,
  grandTotal
) {

  const revisions =
    getRequiredSheet_(
      CONFIG.SHEETS.QUOTATION_REVISIONS
    );

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

  for (let i = 0; i < data.length; i++) {

    const row = data[i];

    if (
      row[1] === qID &&
      row[2] === revisionNo
    ) {

      const sheetRow = i + 2;

      revisions.getRange(sheetRow, 10).setValue(subTotal);
      revisions.getRange(sheetRow, 11).setValue(discountPercent);
      revisions.getRange(sheetRow, 12).setValue(vatPercent);
      revisions.getRange(sheetRow, 13).setValue(grandTotal);

      return;
    }
  }
}



function validateCreateQuotationInput_(data) {

  if (!data) {
    throw new Error("Missing quotation data.");
  }

  if (!data.customerID) {
    throw new Error("Customer is required.");
  }

  if (!data.projectName) {
    throw new Error("Project name is required.");
  }

  if (data.currency) {

    const allowed = [
      CONFIG.CURRENCY.EGP,
      CONFIG.CURRENCY.USD,
      CONFIG.CURRENCY.EUR
    ];

    if (!allowed.includes(data.currency)) {
      throw new Error("Invalid currency.");
    }
  }
}

