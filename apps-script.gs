// Google Apps Script — deploy as Web App (Execute as: Me, Who has access: Anyone)
// วิธีติดตั้ง:
//   1. เปิด Google Sheet → Extensions → Apps Script
//   2. วางโค้ดนี้แทนที่โค้ดเดิมทั้งหมด
//   3. Deploy → New deployment → Web app
//      - Execute as: Me
//      - Who has access: Anyone
//   4. Copy the Web App URL แล้วนำไปวางใน index.html ที่ APPS_SCRIPT_URL

function doGet() {
  const ss = SpreadsheetApp.openById('1erWakOsUhujs82wLaci1qoZojT8450Fd9Ouw2lgmpeI');
  const sheet = ss.getSheets()[0];
  const values = sheet.getDataRange().getValues();

  // Row 0 = headers: [empty, size1, size2, ...]
  // Row 1+ = [serviceName, price1, price2, ...]
  const sizes = values[0].slice(1).map(String).filter(Boolean);
  const services = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const name = String(row[0]).trim();
    if (!name) continue;
    services[name] = {};
    for (let j = 0; j < sizes.length; j++) {
      services[name][sizes[j]] = Number(row[j + 1]) || 0;
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ sizes, services }))
    .setMimeType(ContentService.MimeType.JSON);
}
