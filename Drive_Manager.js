// ===============================
// DRIVE MANAGER MODULE
// ===============================




// ===============================
// CUSTOMER FOLDERS
// ===============================

function createCustomerFolder(customerID, customerName) {

  const parentFolder =
    DriveApp.getFolderById(
      CONFIG.PARENT_FOLDER_ID
    );

  const customerFolder =
    parentFolder.createFolder(
      `${customerID} - ${customerName}`
    );

  const subFolders = [

    "00 Documents",

    "01 Quotations",

    "02 Projects",

    "03 Drawings",

    "04 Test Reports",

    "05 Warranty"

  ];

  subFolders.forEach(name => {

    customerFolder.createFolder(name);

  });

  return customerFolder.getUrl();
}