// Màu tiếng Việt
// Cước S, M => M đậm hơn S để dễ phân biệt




function select_Machine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Machine_Group_In_TE');
  const data = sheet.getRange('B2:B' + sheet.getLastRow()).setValue('True');
  console.log(sheet.getLastRow())
}

function clearAll_Machine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Machine_Group_In_TE');
  const data = sheet.getRange('B2:B' + sheet.getLastRow()).clearContent();
}










function clearDateShift(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data_Kanban');
  const data = sheet.getRange('G1:H1').clearContent();

}