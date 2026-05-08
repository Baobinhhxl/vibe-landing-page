/**
 * IDEMY — Waitlist Form → Google Sheets
 * ═══════════════════════════════════════════════════════════════
 *
 * HƯỚNG DẪN SETUP (5 bước):
 *
 * 1. Tạo Google Sheets mới tại sheets.google.com
 *    - Đặt tên sheet tab: "Danh sách chờ"
 *
 * 2. Vào menu: Extensions → Apps Script
 *
 * 3. Xóa code mặc định, paste toàn bộ code trong file này vào
 *
 * 4. Deploy:
 *    - Click "Deploy" → "New deployment"
 *    - Type: "Web app"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 *    - Click "Deploy" → Copy URL
 *
 * 5. Mở index.html, tìm dòng:
 *       const ENDPOINT = 'YOUR_APPS_SCRIPT_URL_HERE';
 *    Thay bằng URL vừa copy.
 *
 * BẢO MẬT:
 *    - TOKEN phải khớp với TOKEN trong index.html
 *    - Honeypot field 'website' phải trống (bot thường điền vào)
 *    - Nếu muốn đổi token: sửa cả 2 chỗ (đây và index.html) rồi deploy lại
 *
 * ═══════════════════════════════════════════════════════════════
 */

// Phải khớp với TOKEN trong index.html
var SECRET_TOKEN = 'idm-wl-Kx9#mP2025';

function doPost(e) {
  try {
    var params = e.parameter;

    // 1. Kiểm tra secret token
    if (params.token !== SECRET_TOKEN) {
      return ContentService
        .createTextOutput(JSON.stringify({ result: 'error', message: 'Unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Honeypot — nếu field 'website' có giá trị thì là bot
    if (params.website && params.website.trim() !== '') {
      return ContentService
        .createTextOutput(JSON.stringify({ result: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Danh sách chờ') || ss.getActiveSheet();

    // Thêm header nếu sheet trống
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Thời gian', 'Họ tên', 'Email', 'Số điện thoại']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1654a8').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    var name  = params.name  || '';
    var email = params.email || '';
    var phone = params.phone || '';

    // Validate cơ bản phía server
    if (!name || !email || !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) {
      return ContentService
        .createTextOutput(JSON.stringify({ result: 'error', message: 'Invalid data' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    sheet.appendRow([new Date(), name, email, phone]);

    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat('dd/mm/yyyy HH:mm:ss');

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', row: lastRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test function — chạy trong Apps Script editor để kiểm tra
function testDoPost() {
  var mockEvent = {
    parameter: {
      token:   'idm-wl-Kx9#mP2025',
      website: '',
      name:    'Nguyễn Văn Test',
      email:   'test@example.com',
      phone:   '0912345678'
    }
  };
  var result = doPost(mockEvent);
  Logger.log(result.getContent());
}
