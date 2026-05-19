// @ts-nocheck
/*****************************************************
 EEI - QUOTATION FORM CONTROLLER
 Main controller for quotation management form
*****************************************************/

const QFORM = {

  SHEET() {
    return SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Quotation_Form");
  },

  getData() {

    const sh = this.SHEET();

    return {

      quotationSelector:
        sh.getRange("B2").getValue(),

      customer:
        sh.getRange("B4").getValue(),

      quotationID:
        sh.getRange("B6").getValue(),

      revision:
        sh.getRange("E2").getValue(),

      status:
        sh.getRange("B8").getValue(),

      rfqDate:
        sh.getRange("D6").getValue(),

      assignedTo:
        sh.getRange("E5").getValue(),

      currency:
        sh.getRange("D9").getValue(),

      rfqLink:
        sh.getRange("D10").getValue(),

      notes:
        sh.getRange("B11").getValue()

    };

  }

};


/*****************************************************
 CLEAR FORM
*****************************************************/

function clearQuotationForm() {

  const sh = QFORM.SHEET();

  sh.getRangeList([

    "B2",
    "B4",
    "B6",
    "B8",
    "E2",
    "E5",
    "D6",
    "D9",
    "D10",
    "B11"

  ]).clearContent();

  sh.getRange("A20:F100")
    .clearContent();

  refreshQuotationKPIsFromForm();

}


/*****************************************************
 LOAD SELECTED QUOTATION
*****************************************************/

function loadQuotationToForm() {

  const data = QFORM.getData();

  if (!data.quotationSelector) {

    SpreadsheetApp
      .getUi()
      .alert("Select quotation");

    return;

  }

  loadQuotationData_(
    data.quotationSelector
  );

  refreshQuotationKPIsFromForm();

}


/*****************************************************
 CREATE QUOTATION
*****************************************************/

function createQuotationFromForm() {

  createQuotation();

  refreshQuotationKPIsFromForm();

}


/*****************************************************
 ADD ITEMS
*****************************************************/

function addQuotationItemsFromGrid() {

  addQuotationItems();

  refreshQuotationKPIsFromForm();

}


/*****************************************************
 CREATE REVISION
*****************************************************/

function createRevisionFromForm() {

  createQuotationRevision();

  refreshQuotationKPIsFromForm();

}


/*****************************************************
 OPEN FOLDER
*****************************************************/

function openQuotationFolder() {

  const sh = QFORM.SHEET();

  const link =
    sh.getRange("D10")
      .getValue();

  if (link) {

    const html =
      HtmlService
        .createHtmlOutput(
          `<script>
      window.open("${link}");
      google.script.host.close();
      </script>`);

    SpreadsheetApp
      .getUi()
      .showModalDialog(
        html,
        "Opening..."
      );

  }

}




function refreshQuotationForm() {
  loadQuotationToForm();
}

function createQuotationRevisionFromForm() {
  createRevisionFromForm();
}

function openQuotationFolderFromForm() {
  openQuotationFolder();
}

function openRFQFolderFromForm() {
  openQuotationFolder();
}