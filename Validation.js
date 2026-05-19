// Validate customer name (Windows safe)
function isValidName(name) {
  return !(/[\\\/:*?"<>|]/.test(name));
}

function isValidPhone(phone) {

  if (!phone) return true;

  return /^[0-9]+$/.test(phone);
}

function isValidEmail(email) {

  if (!email) return true;

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}


// @ts-nocheck

function validateQuotationDatabaseStructure_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const itemsSheet = ss.getSheetByName(CONFIG.SHEETS.QUOTATION_ITEMS);
  const quotationsSheet = ss.getSheetByName(CONFIG.SHEETS.QUOTATIONS);

  if (!itemsSheet) {
    throw new Error("Missing sheet: Quotation_Items");
  }

  if (!quotationsSheet) {
    throw new Error("Missing sheet: Quotations");
  }

  const itemsHeaders = itemsSheet
    .getRange(1, 1, 1, itemsSheet.getLastColumn())
    .getValues()[0];

  const quotationHeaders = quotationsSheet
    .getRange(1, 1, 1, quotationsSheet.getLastColumn())
    .getValues()[0];

  const requiredItemHeaders = [
    "Item_Status",
    "Modified_By",
    "Modified_Date",
    "Deleted_By",
    "Deleted_Date"
  ];

  const requiredQuotationHeaders = [
    "Record_Status"
  ];

  requiredItemHeaders.forEach(function(header) {
    if (!itemsHeaders.includes(header)) {
      throw new Error("Missing column in Quotation_Items: " + header);
    }
  });

  requiredQuotationHeaders.forEach(function(header) {
    if (!quotationHeaders.includes(header)) {
      throw new Error("Missing column in Quotations: " + header);
    }
  });

  return true;
}

