/**
 * Tranquila Elderly Care — ระบบบันทึก Log การเสนอราคา (Google Apps Script)
 * -------------------------------------------------------------
 * วิธีติดตั้ง
 * 1) สร้าง Google Sheet ใหม่ (ชื่ออะไรก็ได้ เช่น "Tranquila Sales Log")
 * 2) เมนู ส่วนขยาย (Extensions) > Apps Script
 * 3) ลบโค้ดเดิมทั้งหมด แล้ววางไฟล์นี้ทั้งไฟล์แทน
 * 4) กด Deploy > New deployment
 *    - Select type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5) กด Deploy แล้วคัดลอก "Web app URL" ที่ได้ ไปวางในหน้าแอป (แท็บ "ตั้งค่า")
 * 6) ทุกครั้งที่แก้โค้ดนี้ ต้องกด Deploy > Manage deployments > แก้ไข (ไอคอนดินสอ) > Version: New version > Deploy ใหม่
 */

var SHEET_NAME = 'Log';
var HEADERS = ['เวลาที่บันทึก','เจ้าหน้าที่','ชื่อลูกค้า','เบอร์โทร/LINE','กลุ่มผู้รับบริการ','โซน','ประเภทห้อง','ห้องพักจริง','ราคา/วัน','ราคา/เดือน','รูปแบบ','จำนวน','ยอดรวม','มีเจ้าหน้าที่ประกบ24ชม.','หมายเหตุ'];

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet_();
    sheet.appendRow([
      new Date(),
      data.staff || '',
      data.customerName || '',
      data.customerContact || '',
      data.group || '',
      data.zone || '',
      data.roomType || '',
      data.roomChoice || '',
      data.priceDay || '',
      data.priceMonth || '',
      data.durationType || '',
      data.durationValue || '',
      data.total || '',
      data.escort ? 'มี' : 'ไม่มี',
      data.notes || ''
    ]);
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var sheet = getOrCreateSheet_();
    var values = sheet.getDataRange().getValues();
    values.shift(); // remove header row
    var records = values.map(function (r) {
      return {
        timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        staff: r[1],
        customerName: r[2],
        customerContact: r[3],
        group: r[4],
        zone: r[5],
        roomType: r[6],
        roomChoice: r[7],
        priceDay: r[8],
        priceMonth: r[9],
        durationType: r[10],
        durationValue: r[11],
        total: r[12],
        escort: r[13],
        notes: r[14]
      };
    }).reverse();
    return ContentService.createTextOutput(JSON.stringify(records))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
