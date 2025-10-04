/**
 * @OnlyCurrentDoc
 * Đây là đoạn mã để tự động gửi thông báo sự cố chất lượng từ Google Sheet
 * đến một chatbot (ví dụ: Google Chat) khi có dữ liệu mới được thêm vào.
 */

// BIẾN TOÀN CỤC
const CHAT_WEBHOOK_URL = 'https://chat.googleapis.com/v1/spaces/';
const SHEET_NAME = 'NC';

function triggerOnDataAdd(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`Không tìm thấy sheet với tên '${SHEET_NAME}'`);
    return;
  }

  const lastRow = sheet.getLastRow();

  // Dùng DocumentProperties để gắn với file hiện tại (container-bound)
  const props = PropertiesService.getDocumentProperties();
  let lastNotifiedRow = parseInt(props.getProperty('LAST_NOTIFIED_ROW') || '0', 10);

  // Nếu đã xóa bớt dòng khiến lastRow < lastNotifiedRow, tự hiệu chỉnh để tránh kẹt
  if (lastNotifiedRow > lastRow) {
    Logger.log(`LAST_NOTIFIED_ROW (${lastNotifiedRow}) > lastRow hiện tại (${lastRow}). Có thể do xóa/dọn dữ liệu. Tự hiệu chỉnh về ${lastRow}.`);
    lastNotifiedRow = lastRow;
    props.setProperty('LAST_NOTIFIED_ROW', String(lastRow));
  }

  if (lastRow > lastNotifiedRow) {
    Logger.log(`Phát hiện dòng mới. Gửi thông báo cho dòng cuối cùng: ${lastRow}.`);

    const newDataRow = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    if (newDataRow.join("").trim().length > 0) {
      sendNotification(newDataRow);
      props.setProperty('LAST_NOTIFIED_ROW', String(lastRow));
      Logger.log(`Đã cập nhật dòng cuối cùng đã thông báo thành: ${lastRow}.`);
    } else {
      Logger.log(`Dòng cuối cùng (${lastRow}) trống. Bỏ qua thông báo.`);
    }
  } else {
    Logger.log(`Không có dòng mới để thông báo. Dòng cuối cùng: ${lastRow}, Dòng đã thông báo: ${lastNotifiedRow}.`);
  }
}


/**
 * Hàm để định dạng và gửi thông báo đến webhook.
 * @param {Array} rowData Mảng chứa dữ liệu của một dòng.
 */
function sendNotification(rowData) {
  // Ánh xạ dữ liệu từ các cột vào các biến để dễ đọc
  const dateValue = rowData[0] ? new Date(rowData[0]).toLocaleDateString('vi-VN') : 'N/A';
  const shift = rowData[1] || 'N/A';
  const department = rowData[2] || 'N/A';
  const refNumber = rowData[3] || 'N/A';
  const description = rowData[4] || 'N/A';
  const productCode = rowData[5] || 'N/A';
  const quantity = rowData[7] || '0';
  const unit = rowData[8] || '';
  const machine = rowData[11] || 'N/A';

  // Tạo nội dung tin nhắn theo định dạng Card V2 của Google Chat
  const payload = {
    'cardsV2': [{
      'cardId': 'quality-incident-card',
      'card': {
        'header': {
          'title': 'SỰ CỐ CHẤT LƯỢNG KHU VỰC ...',
          'subtitle': 'Phát hiện sự cố mới cần chú ý.',
          'imageUrl': 'https://img.icons8.com/color/48/FA5252/error--v1.png',
          'imageType': 'CIRCLE'
        },
        'sections': [
          {
            'widgets': [
              {
                'decoratedText': {
                  'icon': { 'iconUrl': 'https://img.icons8.com/fluency-systems-filled/48/FA5252/rules.png' },
                  'topLabel': 'Mô tả sự cố',
                  'text': `<b>${description}</b>`,
                  'wrapText': true
                }
              },
              { 'decoratedText': { 'topLabel': 'Ngày', 'text': dateValue, 'icon': { 'iconUrl': 'https://img.icons8.com/material-outlined/48/757575/calendar.png' } } },
              { 'decoratedText': { 'topLabel': 'Ca', 'text': shift.toString(), 'icon': { 'iconUrl': 'https://img.icons8.com/material-outlined/48/757575/time.png' } } },
              { 'decoratedText': { 'topLabel': 'Bộ phận phát hiện', 'text': department, 'icon': { 'iconUrl': 'https://img.icons8.com/material-outlined/48/757575/building.png' } } },
              { 'decoratedText': { 'topLabel': 'Mã tham chiếu', 'text': refNumber, 'icon': { 'iconUrl': 'https://img.icons8.com/material-outlined/48/757575/hashtag.png' } } },
              { 'decoratedText': { 'topLabel': 'Mã sản phẩm', 'text': productCode, 'icon': { 'iconUrl': 'https://img.icons8.com/material-outlined/48/757575/barcode.png' } } },
              { 'decoratedText': { 'topLabel': 'Số lượng', 'text': `${quantity} ${unit}`, 'icon': { 'iconUrl': 'https://img.icons8.com/material-outlined/48/757575/stack.png' } } },
              { 'decoratedText': { 'topLabel': 'Máy', 'text': machine, 'icon': { 'iconUrl': 'https://img.icons8.com/material-outlined/48/757575/engine.png' } } }
            ]
          }
        ]
      }
    }]
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json; charset=UTF-8',
    'payload': JSON.stringify(payload)
  };

  try {
    UrlFetchApp.fetch(CHAT_WEBHOOK_URL, options);
    Logger.log('Gửi thông báo thành công.');
  } catch (e) {
    Logger.log('Lỗi khi gửi thông báo: ' + e.toString());
  }
}


//data: Date	Shift	Department	ReferenceNumber	DescriptionVn	ProductCode	Batch	RefQuantity	Unit	AreaNc	GroupNc	MachineGroup	Defect Cost	Total Cost	Tháng	Year
//2025-01-07	2	PD	NC-001-25	Bao bì hư	A50307	0301252C	32	cs	AMG	A	Máy A4	1.101.204	1.763.627	2025-01	2025
