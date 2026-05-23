// Extract Google Drive Folder ID
function extractIdFromUrl(url) {
  if (!url) return null;
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

// Get logged user name
function getCurrentUserName() {

  let userEmail = "";

  try {
    userEmail = Session.getActiveUser().getEmail();
  } catch (e) {}

  if (!userEmail) {
    try {
      userEmail = Session.getEffectiveUser().getEmail();
    } catch (e) {}
  }

  if (!userEmail) {
    return "Unknown_User";
  }

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);

    if (!usersSheet || usersSheet.getLastRow() < 2) {
      return userEmail;
    }

    const data = usersSheet
      .getRange("B2:C" + usersSheet.getLastRow())
      .getValues();

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === userEmail.toLowerCase()) {
        return data[i][1] || userEmail;
      }
    }

    return userEmail;

  } catch (e) {
    return userEmail;
  }
}


function getSheet(name){

  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(name);

}


function sheetExists(sheet) {

  return sheet && sheet.getLastRow() > 0;

}




function openUrl_(url, title) {

  const html =
    HtmlService
      .createHtmlOutput(
        '<script>' +
        'window.open("' + url + '", "_blank");' +
        'google.script.host.close();' +
        '</script>'
      )
      .setWidth(100)
      .setHeight(50);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      title || "Opening..."
    );
}





