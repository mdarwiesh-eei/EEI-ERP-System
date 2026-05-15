function refreshAll(){

  buildCustomerSelector();

  buildQuotationCustomerSelector();

  buildAssignedUserSelector();

  buildQuotationSelectorForForm();

  buildRevisionSelectorForForm();

}




function refreshFromMenu(){

  
  buildQuotationSelectorForForm();
  
  buildQuotationCustomerSelector();
  
  buildAssignedUserSelector();
  
  
  applyRFQDatePicker();

  buildRevisionSelectorForForm();

  SpreadsheetApp
    .getUi()
    .alert("Selectors are ready ✅");
}





