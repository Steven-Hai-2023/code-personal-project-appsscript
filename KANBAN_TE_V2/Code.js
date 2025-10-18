const TEMPLATE_ID = "ID"//
const OUTPUT_FOLDER = "ID";
const FONT_ID = "1jsRm-ID-_"; 




function showDialog() {
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(HtmlService.createHtmlOutputFromFile('kanban').setWidth(500)
    .setHeight(400), "Tạo thẻ kanban");
}

function getKanbanTemplate() {
  const file = DriveApp.getFileById(TEMPLATE_ID);
  const bytes = file.getBlob().getBytes();
  return bytes;
}

function getFont() {
  const fontFile = DriveApp.getFileById(FONT_ID);
  return fontFile.getBlob().getBytes();
}


function nhap() {
  console.log(getDate_Sheet())
}







function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data_Kanban');
  const data = sheet.getRange('A3:I' + sheet.getLastRow()).getValues();
  const list = data.map(row => {
    const [machine, material, ske, output, bundle, color, resin, quantity, multiply] = row;
    return { machine, material, ske, output, bundle, color, resin, quantity, multiply };
  });

  const newList = [];
  for (const item of list) {
    for (let i = 0; i < item.multiply; i++) {
      // Tạo một bản sao của item để không làm thay đổi bản gốc
      const newItem = { ...item, cardNumber: i + 1 };
      if (i === (item.multiply - 1)) {  //  if (i === (item.multiply - 1)) {
        newItem.quantity = ''; // Chỉ xóa quantity ở 1 thẻ cuối cùng
      }
      newList.push(newItem); // Thêm bản sao vào newList
   
    // Thêm thẻ PDF trống sau mỗi dòng dữ liệu
    //   if (i === item.multiply - 1) {
    //     newList.push({ ...item, quantity: '1' }); // Thẻ PDF trống
    //   }
    }
    
  }

  return newList;
}


function createPdf(base64Content, name) {
  const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER);
  const decoded = Utilities.base64Decode(base64Content);

  const blob = Utilities.newBlob(decoded, "application/pdf", name || Date.now().toString());
  const file = outputFolder.createFile(blob);
  return file.getUrl();

}





//-------------------Copy lịch sử in : ngày và user ----------------//

function getUser() {
  return Session.getActiveUser().getEmail();
}




function copyHis() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data_Kanban');
  const data = sheet.getRange('A3:I' + sheet.getLastRow()).getValues().filter(x=>x[0] !== "");

  const date_start = sheet.getRange("B1").getDisplayValue();
  const date_Actual = new Date(date_start)
  const formatter = new Intl.DateTimeFormat('en-US');
  const formattedDate = formatter.format(date_Actual);

  const s_History = ss.getSheetByName('History');
  const lr_s_History = getLastRow('History', 'A:A')

  //Lấy thông tin người dùng & Ngày ca in thẻ
  const user = getUser();
  const now = new Date();
  const ca = getBatch_Shift();
  const ngay = formattedDate;

  // Tạo mảng
  const combineData = data.map(row => {
    return [...row, ngay, ca, user, now];
  });

  //Ghi data
  s_History.getRange(lr_s_History + 1, 1, combineData.length, combineData[0].length).setValues(combineData);

}





//------------ Clear all files PDF in Folder ------------//



function emptyFolder() {

    const folder = DriveApp.getFolderById(OUTPUT_FOLDER);
  
    while (folder.getFiles().hasNext()) {
      const file = folder.getFiles().next();
      Logger.log('Moving file to trash: ', file);
      file.setTrashed(true);
  }
}




//-------------------Copy Batch ngày & Ca ----------------//


//lấy Batch chứa ngày
function getBatch_Date(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_KanBan = ss.getSheetByName('Data_Kanban');
  const Batch_Date = sheet_KanBan.getRange("E1").getValue().toString();
  console.log(Batch_Date)
  return Batch_Date
}

//lấy Batch chứa ca
function getBatch_Shift(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_KanBan = ss.getSheetByName('Data_Kanban');
  const Batch_Shift = sheet_KanBan.getRange("D1").getValue().toString();
  console.log(Batch_Shift)
  return Batch_Shift // "C" + 
}


//lấy ngày kế hoạch in thẻ
function getDate_Kanban(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_KanBan = ss.getSheetByName('Data_Kanban');
  const date_Kanban = sheet_KanBan.getRange("B1").getValue();
  // const date_Actual = new Date(date_Kanban)
  const ngay_KeHoach = date_Kanban.toLocaleDateString('en-GB')
  const time_now = (new Date).toLocaleTimeString();

  console.log(ngay_KeHoach, time_now)
  return [ngay_KeHoach, time_now]
}



//-------------------Copy Kế hoạch & Master Data ----------------//



//Lấy kế hoạch về file 
function copyPlan() {
  // try{

    const ID_Plan = 'ID_Plan'
    const ss_Id = SpreadsheetApp.openById(ID_Plan);
    const sheet_Plan = ss_Id.getSheetByName('Planning input');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet_KanBan = ss.getSheetByName('Data_Kanban');
    const date_start = sheet_KanBan.getRange("B1").getDisplayValue();
    const date_Actual = new Date(date_start)
    const shift_Actual = sheet_KanBan.getRange("D1").getDisplayValue().toString();

    // chỉ lấy các dòng có ngày và khác N/A ở cột ngày
    const Plan = sheet_Plan.getRange('B2:I' + sheet_Plan.getLastRow()).getValues().filter(x => (x[0] !== '') && (x[0] !== '#N/A'));

    console.log(Plan)

    const new_Data = Plan.filter(row => {
        // Destructuring để code rõ ràng hơn
      const [dateStr, , area, , , , shift, output] = row;
      const datePlan = new Date(dateStr);
      const shiftPlan = shift.toString();
      const areaPlan = area.toString();

      // Kiểm tra ngày hợp lệ
      if (isNaN(datePlan.getTime())) {
        console.warn(`Invalid date found: ${row[0]}`);
        // console.log(row)

        return false; // Bỏ qua phần tử không hợp lệ
      }

      const planDateString = datePlan.toISOString().split('T')[0]; // Lấy chuỗi 'YYYY-MM-DD'
      const actualDateString = date_Actual.toISOString().split('T')[0];

      return (planDateString === actualDateString 
                && shiftPlan === shift_Actual 
                && areaPlan.toString().includes('TE')  // 'TE'
                && output > 0);
      });
    //Lấy dữ liệu ngày, ca cần in
    console.log(new_Data.length)

    const sheet_Des_Plan = ss.getSheetByName('dữ liệu kế hoạch Planning');
    sheet_Des_Plan.getRange("A3:H").clearContent();
    sheet_Des_Plan.getRange(3, 1, new_Data.length, new_Data[0].length).setValues(new_Data);
  // }catch(e){
  //   console.log('Lỗi kế hoạch không tồn tại!')
  // }

}



//Copy số lượng cán / thùng LE

function copyMasterDataLE() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ID_LE = 'ID_LE'
  const ss_Id = SpreadsheetApp.openById(ID_LE);
  const sheet_LE = ss_Id.getSheetByName('MLGN_LE quantity');

    // chỉ lấy các dòng có ngày và khác N/A ở cột ngày
  const LE = sheet_LE.getRange('B2:C' + sheet_LE.getLastRow()).getValues().filter(x => x[0].toString().startsWith('C'));
  console.log(LE.length)

  const sheet_Des_MasterData = ss.getSheetByName('Master_Data');
  sheet_Des_MasterData.getRange("A2:B").clearContent();
  sheet_Des_MasterData.getRange(2, 1, LE.length, LE[0].length).setValues(LE);

}


//Copy C, Bundle màu
function copyMasterDataMAKT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ID_MAKT = 'ID_MAKT'
  const ss_Id = SpreadsheetApp.openById(ID_MAKT);
  const sheet_MAKT = ss_Id.getSheetByName('MAKT_NAME');

    // chỉ lấy các dòng có  Col4 Like 'IM' AND Col5 <> 'C5'
  const MAKT = sheet_MAKT.getRange('A2:F' + sheet_MAKT.getLastRow()).getValues()
  .filter(x => x[3].toString() === 'TE' && x[4].toString() !== 'C5')
  .map(row => [row[0],row[1]]);
  console.log(MAKT.length)

  const sheet_Des_MasterData = ss.getSheetByName('Master_Data');
  sheet_Des_MasterData.getRange("F2:G").clearContent();
  sheet_Des_MasterData.getRange(2, 6, MAKT.length, MAKT[0].length).setValues(MAKT);
    // timBundleAndColor(sheet_Input, range_Data)

    console.log(MAKT)
}





//Xử lý tên Bundle
function timBundleAndColor(sheet_Input, range_Data)
{

const rangee_1 = sheet_Master_Data.getRange(range_Data).getValues()
const rangee = rangee_1.filter(x => x[0] !== '')
// const rangee =  updateData_MAKT_NAME_Ggsheet();

let arr = []
let search_S = '_S '
let search_M = '_M '
let search_H = '_H '
for(let i=0;i<rangee.length;i++)
{
  let name = rangee[i][1].toString()    
  if(name.search('_S ')>-1 ||name.search(' S ')>-1 )
  {
    let bundle = name.slice(0,name.search(search_S))
    let color = name.slice(name.search(search_S)+3,name.length+1)
        arr.push([rangee[i][0],rangee[i][1],bundle,'Soft',color])
  }
  else if(name.search('_M ')>-1 ||name.search(' M ')>-1 )
  {
    let bundle = name.slice(0,name.search(search_M))
    let color = name.slice(name.search(search_M)+3,name.length+1)
        arr.push([rangee[i][0],rangee[i][1],bundle,'Medium',color])
  }
  else if(name.search('_H ')>-1 ||name.search(' H ')>-1 )
  {
    let bundle = name.slice(0,name.search(search_H))
    let color = name.slice(name.search(search_H)+3,name.length+1)
        arr.push([rangee[i][0],rangee[i][1],bundle,'Firm',color])
  }
}

console.log(arr)
//Xóa và ghi lại data từ cột F
sheet_Input.getRange(2,6,3000,arr[0].length).clearContent()
sheet_Input.getRange(2,6,arr.length,arr[0].length).setValues(arr)
 
}




















// ---------------- Return Function -------///
function getLastRow(sheetName, rangeColumn) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName)
  const range = sheet.getRange(rangeColumn).getValues()

  let lr = false
  let row = 1
  for (let i = 0; i < range.length; i++) {
    if (range[i][0] !== '') {
      lr = true
      row = i + 1
    } else {
      lr = false
    }
  }
  return row
}

