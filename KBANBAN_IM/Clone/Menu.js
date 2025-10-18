function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Tạo thẻ kanban')
    .addItem('Copy Kế hoạch & Master Data', 'copy_Plan_MasterData')
    .addItem('Tạo thẻ Kanban', 'showDialog')
    .addToUi();
}

function copy_Plan_MasterData() {
  copyPlan();
  copyMasterDataLE();
  copyMasterDataMAKT();
  copyMasterDataZMNU();
}







