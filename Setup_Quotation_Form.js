function styleEditableCell_(range) {

  range
    .setBackground("#FFFDE7")
    .setFontColor("#000000")
    .setBorder(true, true, true, true, true, true);
}

function styleSystemCell_(range) {

  range
    .setBackground("#E5E7E9")
    .setFontWeight("bold")
    .setFontColor("#1B2631")
    .setBorder(true, true, true, true, true, true);
}

function styleSectionHeader_(range) {

  range
    .setBackground("#A61C00")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}

function applyQuotationFormUIStyles() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEETS.QUOTATION_FORM);

  if (!sheet) {
    SpreadsheetApp.getUi().alert("Quotation_Form sheet not found.");
    return;
  }

  const editableCells = [
    "B4",
    "B5",
    "E4",
    "E5",
    "E6",
    "E7",
    "E8",
    "B11",
    "B12",
    "B13",
    "B14",
    "E11",
    "E12",
    "B19:K35"
  ];

  editableCells.forEach(function(cell) {
    styleEditableCell_(sheet.getRange(cell));
  });

  const systemCells = [
    "B6",
    "B7",
    "B8",
    "H11",
    "H12",
    "H13",
    "H14",
    "H15"
  ];

  systemCells.forEach(function(cell) {
    styleSystemCell_(sheet.getRange(cell));
  });

  const sectionHeaders = [
    "A1:L1",
    "A3:L3",
    "A10:F10",
    "G10:L10",
    "A17:L17",
    "A38:L38"
  ];

  sectionHeaders.forEach(function(range) {
    styleSectionHeader_(sheet.getRange(range));
  });

  SpreadsheetApp.getUi().alert("Quotation Form UI Styles Applied ✅");
}