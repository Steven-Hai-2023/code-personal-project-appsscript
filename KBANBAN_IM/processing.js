function select_Machine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Machine_Group_In_IM');
  const data = sheet.getRange('B2:B' + sheet.getLastRow()).setValue('True');
  console.log(sheet.getLastRow())
}

function clearAll_Machine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Machine_Group_In_IM');
  const data = sheet.getRange('B2:B' + sheet.getLastRow()).clearContent();
}