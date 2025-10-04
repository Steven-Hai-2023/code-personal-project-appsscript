//Reply email in a week (Monday-Sunday)
function replyAllEmails() {
  const now = new Date();
  const year = now.getFullYear();
  // Lấy tất cả các thư trong hộp thư đến
  //Sử dụng Try để khi code bên trong lỗi tức là tiêu đề email không tìm thấy sẽ gửi email theo tiêu đề mới
  try {
     //Chủ đề email 
    const subject1 = '.. PRODUCTION OUTPUT REPORT WEEK: ' + (weekNumber()) + " - " + year
    //Tìm chủ đề email
    let threads = GmailApp.search('subject:"' + subject1 + '"');
    //let threads = GmailApp.search('subject:"REPLY TO ALL CHECK FIRST HAI WEEK 1"');
    //var threads = GmailApp.search(`from:${emailto} `);
    console.log(threads.length, '1function replyAllEmails')
    if (threads.length == 0) {

          //lấy email
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("Email Report")
      
      let columnToCheckA = sheet.getRange("A:A").getValues();
      let lastRowA = getLastRosheetpecial(columnToCheckA); // Last row
      const rangeA = sheet.getRange("A3:A" + lastRowA).getValues();
      //Với a là email đúng, b là email sai
      let [a, b] = kiemTraEmailReportTE(rangeA)
      console.log(b, b.length,rangeA, '2function replyAllEmails')

      //Nếu b không có giá trị sẽ bỏ qua
      try{
          sheet.getRange("D3:D").clearContent()
          sheet.getRange("D3:D" + (b.length + 2)).setValues(b)
      }catch(a){
        console.log("Range không có data" + 'Try Catch 3function replyAllEmails')
      }
      
      // Danh sách email từ sheet Email Report cột A.
      //Chỉ lấy các email đúng format, còn email sai format đại diện biến b
      const email = a

      //Nếu chưa có tiêu đề email được gửi thì gửi email mới với tiêu đề mới
      // const email = ['', '' ]
      //Nếu có email mới cần add thêm vào đây

      //Nếu có email bị sai sẽ thông báo ra màn hình
      if(b.length>0){
 SpreadsheetApp.getUi().alert("Email đã được gửi thành công.\n ***Email bên dưới bị sai cần sửa lại:\n \t" + b.join("\n\t"))
      }

      MailApp.sendEmail({
        //email cần gửi, chủ đề đã gửi, htmlBody cần gửi
        to: email.join(","),
        //cc: ccEmail,
        subject: ".. PRODUCTION OUTPUT REPORT WEEK: " + (weekNumber()) + " - " + year,
        htmlBody: sendEmailOutput()
      });

    } else {
      //Ngược lại nếu tên tiêu đề email đã có thì chỉ cần reply email cuối cùng đã nhận
      //lấy email cuối cùng trong chuỗi sau đó reply email
      //htmlBody: gồm sendEmailOutput() là output ..., Output_...(): output CAP
      threads[threads.length - 1].replyAll("", { htmlBody: sendEmailOutput() })
      // gửi nối tiếp theo email cuối cùng
    }

  }
  //Nếu code phía trong Try có lỗi sẽ không thực thi mà sẽ thực thi code bên trong catch(e){ code thực thi}
  catch (e) {
 SpreadsheetApp.getUi().alert("Không thể thực thi code.\n\n Bạn hãy kiểm tra lại - Lỗi: " + e + 'Nếu số lượng email > 50 email vui lòng xóa bớt hoặc gửi group mail')
  }
}








//Code gửi email output hàng ca
//Giải thuật
//Tạo menu khi click menu sẽ gửi email output vs đk gửi theo loop email theo tuần
function sendEmailOutput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pilot");
  let columnToCheckA = sheet.getRange("A:A").getValues();
  let lastRowA = getLastRosheetpecial(columnToCheckA); // Last row

  //tiêu đề email cần gửi + ca, ngày, tháng, năm.
  const tieudeemail = "Please kindly update the .. output for " + getDayMonthYear() + ":";
  let dear = "Dear Team,";
  let ca = sheet.getRange("L2").getDisplayValues();
  let ngay = sheet.getRange("K2").getDisplayValues();

  //Header1 - các thông tin phía trên bảng output
  let nhom = sheet.getRange("C3").getDisplayValues();
  let HCs = sheet.getRange("E4").getDisplayValues();
  let tongSanLuong = sheet.getRange("F4").getDisplayValues();
  let chenhLech = convertStringToNumber(sheet.getRange("H4").getDisplayValues()[0][0]);
  let nangSuat = convertStringToNumber(sheet.getRange("I4").getDisplayValues()[0][0]);
  let title_nangSuat = sheet.getRange("I3").getDisplayValues();
  let phePham = sheet.getRange("J4").getDisplayValues();
  let tiLePhePham = sheet.getRange("K4").getDisplayValues();

  // const lrSheet = sheet.getLastRow();

  //--------------------------Khai báo bảng output để gửi HTML - 1---------------------
  //Gán tiêu đề bảng Sự cố máy trực tiếp vào HTML ko cần khai báo biến

  /**
 * Thêm phần bỏ cell 8 ra
 */

      /**
      //Khi có thêm xóa dòng ở sheet Pilot thì chỉnh sửa ở đây
      Tính cả dòng tổng cuối cùng
      *** Thay đổi dòng ở đây tongSoDong
      */  
    let tongSoDong = 55;
  const tablerangeValueOrign = sheet.getRange(6, 1, tongSoDong, 12).getValues();
  const tablerangeValue = tablerangeValueOrign.filter(row => (row[0] !== '0801...' &&  row[0] !== '0802....'));

  //Phần thân của bảng
  // const tablerangeValue = sheet.getRange(6, 1, 47, 12).getValues();
  // nếu nguyên nhân thiếu sản lượng được nhập >1 : nhập G và H, Hoặc G, H, I thì nối chuỗi này lại và truyền vào mảng tablerangeValue[i][6]
  for (i = 0; i < tablerangeValue.length; i++) {
    if (!(tablerangeValue[i][6] == "" && tablerangeValue[i][7] == "" && tablerangeValue[i][8] == "")) {
      tablerangeValue[i][6] = `${tablerangeValue[i][6]}    .  ${tablerangeValue[i][7]}    ${tablerangeValue[i][8]}`
    }
  }

  const htmlTemplate = HtmlService.createTemplateFromFile('html báo cáo output cuối ca');
  htmlTemplate.dear = dear;
  htmlTemplate.tieudeemail = tieudeemail;
  htmlTemplate.ca = ca;
  htmlTemplate.ngay = ngay;

  //----- Bảng phía trên report Output
  htmlTemplate.nhom = nhom;
  htmlTemplate.HCs = HCs;
  htmlTemplate.tongSanLuong = tongSanLuong;
  htmlTemplate.chenhLech = chenhLech;
  htmlTemplate.title_nangSuat = title_nangSuat;
  htmlTemplate.nangSuat = nangSuat;
  htmlTemplate.phePham = phePham;
  htmlTemplate.tiLePhePham = tiLePhePham;
  htmlTemplate.tablerangeValue = tablerangeValue;// Nội dung bảng Output


  //--------------------------Khai báo bảng Sự cố máy để gửi HTML - 2---------------------
  //Gán tiêu đề bảng Sự cố máy trực tiếp vào HTML ko cần khai báo biến

  //Phần thân của bảng
  let tablerangeValueBang2 = []; // Khai báo biến và gán giá trị mặc định
   try {
      /**Khi thêm xóa dòng cần chỉnh lại ở đây
       * Dòng bắt đầu dữ liệu
       * LRA - dòng chứa tiêu đề
       * 
       * 
       * 
       * *** Thay đổi dòng ở đây dongBD
       */
      let dongBD = 69;
      tablerangeValueBang2 = sheet.getRange(dongBD, 1, (lastRowA - (dongBD-1)), 12).getValues();
      tablerangeValueBang2.sort(function (a, b) {
        return a[2] - b[2];
     });
   console.log("Values = try")
  }catch(e){
      tablerangeValueBang2 = []
      console.log("Values = catch")
  }



  htmlTemplate.tablerangeValueBang2 = tablerangeValueBang2;// Nội dung bảng Output
  console.log(tablerangeValueBang2)

  //-------------------------------------------------

  const htmlForEmail = htmlTemplate.evaluate().getContent();

  let toEmail = 'mail.com'
  // let ccEmail = 'mail.com, ...'// nhiều email phân cách nhau bởi dấu ,

  // MailApp.sendEmail({
  //   to: toEmail,
  //   //cc: ccEmail,
  //   subject: "THỬ NGHIỆM  ...  PRODUCTION OUTPUT REPORT - WEEK " + (getWeekNumber() + 1) + ".V2",
  //   htmlBody: htmlForEmail
  // });

  return htmlForEmail;
}





//Lấy tuần hiện tại
// năm 2023 cần +1 mới đúng tuần hiện tại
// năm 2024 cần check lại
function getWeekNumber() {

  //var now = new Date("2023-1-8");
  //lấy ngày giờ hiện tại
  let now = (new Date());
  console.log(now)
  //lấy ngày đầu tiên của năm
  let onejan = new Date(now.getFullYear(), 0, 1);  //now.getFullYear()
  //trả về tuần hiện tại
  //now-onejan : ngày hiện tại trừ ngày đầu tiên của năm
  // / 86400000 là số milisecond giây trong 1 ngày, tính ra số ngày hiện tại
  //onejan.getDay()-1 ngày đầu tuần là chủ nhật, trừ đi 1 tức là ngày thứ 2 đầu tuần
  return Math.floor((((now - onejan) / 86400000) + onejan.getDay() - 1) / 7) + 1;
}



///Sau khi cập nhật số tuần hiện tại
//// Nếu giờ hiện tại nhỏ thua 8h ngày thứ 2 tức là ca 3 ngày chủ nhật => vẫn tính tuần cũ
// If hiện tại <8h vào ngày thứ 2 tức là tuần cũ, ngược lại tuần mới (sẽ cộng thêm 1);

function weekNumber() {

  //lấy ngày giờ hiện tại
  const now = new Date();
  const hours = now.getHours()
  const day = now.getDay();//Ngày trong 1 tuần từ 0-6 trng đó 0: Chủ Nhật, 1: thứ 2, 6: thứ 7

  //gán tuần hiện tại từ function getWeekNumber 
  let numberWeek = getWeekNumber();
  console.log(numberWeek)

  //Nếu ngày hiện tại là thứ 2 và chưa tới 9h sáng thì vẫn tính tuần cũ, ngược lại tuần mới
  if (hours < 9 && day == 1) {

    //Tuần cũ = tuần hiện tại -1
    return numberWeek - 1;
    console.log(numberWeek, "IF")

  } else {
    //Tuần hiện tại
    return numberWeek;
    console.log(hours, day)
    console.log(numberWeek, "Ngoài If")

  }
}

//function này dùng để nháp có thể xóa hoặc thay đổi nội dung mong muốn
function nhap() {

  console.log(weekNumber())
  console.log("VVVVVVV" + getDayMonthYear())

  //lấy email
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Email Report")
  // const rangeA = sheet.getRange("A3:A").getValues();
        let columnToCheckA = sheet.getRange("A:A").getValues();
      let lastRowA = getLastRosheetpecial(columnToCheckA); // Last row
      const rangeA = sheet.getRange("A3:A" + lastRowA).getValues();
  console.log(rangeA[0])

  // MailApp.sendEmail({
  // to: rangeA.join("\n"),
  // subject: "Logos",
  // htmlBody: "inline Google Logo<img src='cid:googleLogo'> images! <br>" +
  //           "inline YouTube Logo <img src='cid:youtubeLogo'>",



  // })


  let emails = [["email1@example.com"], ["email1@example"], ["email1@example.com"], ["email1@example.com"], ["email1@example.com"]];

  let validEmails = [];

  for (let i = 0; i < emails.length; i++) {
    let emailArr = emails[i];
    for (let j = 0; j < emailArr.length; j++) {
      let email = emailArr[j];
      if (isValidEmail(email)) {
        validEmails.push(email);

      }
    }
  }
  console.log(validEmails,)




  let [a, b] = kiemTraEmailReportTE(rangeA)
  console.log(b, b.length , a, a.length, "-----")

        try{
          sheet.getRange("D3:D").clearContent()
          sheet.getRange("D3:D" + (b.length + 2)).setValues(b)
      }catch(a){
        console.log("Range không có data")
      }
      if(b.length>0){

 SpreadsheetApp.getUi().alert("Email đã được gửi thành công.\n ***Email bên dưới bị sai cần sửa lại:\n \t" + b.join("\n\t"))

      }
  
const c = a;
console.log(c)
}


//Check email gửi, nếu email nào lỗi sẽ trả về ở cột D,
//Lưu ý cần chỉnh sửa lại email đúng để những lần sau gửi có email đó nhận được report
function kiemTraEmailReportTE(emails) {

  let validEmails = [];//email đúng
  let inValidEmails = []; //email sai
  
  for (let i = 0; i < emails.length; i++) {
    let emailArr = emails[i];
    for (let j = 0; j < emailArr.length; j++) {
      let email = emailArr[j];
      //Kiểm tra format email có đúng ko?
      if (isValidEmail(email)) {
        validEmails.push(email);

      } else {
      //Email sai format được đẩy vào mảng mới.
        inValidEmails.push([email]);
      }
    }
  }

  return [validEmails, inValidEmails];


}


//kiểm tra định dạng email
function isValidEmail(email) {
  // Kiểm tra định dạng email
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}



///Không xóa funtion này
///Dùng để lấy thông tin tiêu đề khi gửi email output PKG & CAP
////Lấy ca và ngày , tháng , năm
function getDayMonthYear() {

  const now = new Date();
  const hour = now.getHours();
  let date = now.getDate();
  //const month = now.getMonth().toLocaleString('default', {month: 'short'});
  //format tháng :02 == Feb
  const month = new Date(now.getFullYear(), now.getMonth()).toLocaleString('default', { month: 'short' });
  const year = now.getFullYear();
  let subject = "";
  let shift = 0;

  //Nếu giờ 6-13h59 ca1, 14-22h ca 2, còn lại là ca 3
  //Nếu ALs gửi email qua ca 1-2 tiếng thì cần chỉnh lại hour<(14+2) && hour>=(6+2), tương tự cho 2 ca còn lại
  //giờ này được tính để gán vào tiêu đề, nên người gửi phải gửi trước hoặc trong khung giờ này

  if (hour < 15 && hour >= 7) {
    shift = 1;
  } else if (hour < 23 && hour >= 15) {
    shift = 2;
  } else shift = 3;

  //Nếu ca 3 thì ngày trừ 1 (ngày hôm trước)
  //còn ngược lại thì ngày hiện tại
  if (shift === 3) {
    date = date - 1;
  } else date

  // trả về ca, này, tháng, năm 
  return subject = "shift " + shift + ", " + date + "th " + month + " " + year
  //console.log(hour,date,month,year, shift, subject)

}


// Không xóa function này
//function lấy dòng cuối chứa dữ liệu với điều kiện dòng cuối không được chứa dữ liệu mà phải trống
function getLastRosheetpecial(range) {
  var rowNum = 0;
  var blank = false;
  for (var row = 0; row < range.length; row++) {

    if (range[row][0] === "" && !blank) {
      rowNum = row;
      blank = true;
    } else if (range[row][0] !== "") {
      blank = false;
    };
  };
  return rowNum;

  //let columnToCheck = sheet.getRange(1,i,sheet.getMaxRosheet(),1).getValues();
  //let lastRowI = getLastRosheetpecial(columnToCheck); // Last row

}



