
function clearQuotationFormData(){

  const form = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Quotation_Form");

  form.getRangeList([
    "B3",
    "B6",
    "B7",
    "B8",
    "B9",
    "B10",
    "B11",
    "D6",
    "D9",
    "D10",
    "H6",
    "H7",
    "H8"
  ]).clearContent();

  form.getRange("A15:F100").clearContent();
  form.getRange("H15:L100").clearContent();
}

