//==================================
// USER ROLE MANAGEMENT
//==================================

function getCurrentUserRole() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const users =
    ss.getSheetByName(CONFIG.SHEETS.USERS);

  if (!users) {
    return "User";
  }

  const currentEmail =
    Session
      .getActiveUser()
      .getEmail()
      .toString()
      .trim()
      .toLowerCase();

  if (!currentEmail) {
    return "Admin";
  }

  if (users.getLastRow() < 2) {
    return "User";
  }

  const data =
    users
      .getRange(
        2,
        1,
        users.getLastRow() - 1,
        6
      )
      .getValues();

  for (let i = 0; i < data.length; i++) {

    const email =
      data[i][1]
        .toString()
        .trim()
        .toLowerCase();

    const role =
      data[i][3]
        .toString()
        .trim();

    const status =
      data[i][5]
        .toString()
        .trim();

    if (
      email === currentEmail &&
      status === "Active"
    ) {
      return role || "User";
    }
  }

  return "User";
}
