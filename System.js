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
  
  buildRevisionSelectorForForm();

  SpreadsheetApp
    .getUi()
    .alert("Selectors are ready ✅");
}





