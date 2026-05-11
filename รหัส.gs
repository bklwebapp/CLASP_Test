/**
 * ฟังก์ชันหลักสำหรับแสดงผลหน้าเว็บ
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ระบบบริหารจัดการข้อมูลผู้ป่วย')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ฟังก์ชันสำหรับบันทึกข้อมูลลง Google Sheet แบบ Dynamic Column
 * ระบบจะเช็คหัวตาราง ถ้าไม่มีคอลัมน์นั้นจะสร้างเพิ่มให้
 */
function saveData(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Database');
  
  // 1. ถ้ายังไม่มี Sheet ให้สร้างใหม่
  if (!sheet) {
    sheet = ss.insertSheet('Database');
    // สร้างหัวตารางเริ่มต้น (Core Fields)
    sheet.appendRow(['Timestamp', 'date', 'month', 'hn', 'name', 'status', 'type']);
  }
  
  // 2. อ่านหัวตารางปัจจุบัน
  const lastCol = sheet.getLastColumn();
  // ถ้าไม่มีข้อมูลเลย ให้ถือว่ามี 0 คอลัมน์
  const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  
  // เตรียมข้อมูลแถวใหม่ (Row) ให้ตรงกับหัวตาราง
  const newRow = [];
  const timestamp = new Date();
  
  // ใส่ Timestamp เป็นค่าแรกเสมอ (ถ้ามีหัวข้อนี้)
  const tsIndex = headers.indexOf('Timestamp');
  if (tsIndex > -1) newRow[tsIndex] = timestamp;
  else if (headers.length === 0) { // กรณีสร้างชีทใหม่แต่ยังไม่ได้ appendRow (กันพลาด)
     newRow[0] = timestamp;
  }

  // 3. วนลูปข้อมูลที่ส่งมา เพื่อเช็คกับหัวตาราง
  Object.keys(formData).forEach(key => {
    let colIndex = headers.indexOf(key);
    
    if (colIndex === -1) {
      // ถ้ายังไม่มีหัวข้อนี้ใน Sheet ให้เพิ่มคอลัมน์ใหม่ต่อท้าย
      colIndex = headers.length; 
      sheet.getRange(1, colIndex + 1).setValue(key); // สร้างหัวคอลัมน์ใหม่
      headers.push(key); // อัปเดตตัวแปร headers
    }
    
    // ใส่ข้อมูลให้ตรงช่อง
    newRow[colIndex] = formData[key];
  });
  
  // 4. บันทึกลง Sheet (หาแถวสุดท้ายแล้วต่อท้าย)
  // ต้องมั่นใจว่า newRow มีความยาวเท่ากับ headers ปัจจุบัน (เติมช่องว่างถ้าขาด)
  while(newRow.length < headers.length) {
    newRow.push("");
  }

  sheet.appendRow(newRow);
  
  return { success: true };
}

/**
 * ฟังก์ชันดึงข้อมูลล่าสุด (สำหรับตารางแสดงผล)
 * ปรับให้ map ตามชื่อหัวคอลัมน์จริง
 */
function getRecentData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Database');
  
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return []; 
  
  // อ่านหัวตาราง
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // ดึงข้อมูล 20 แถวล่าสุด
  const startRow = Math.max(2, lastRow - 20);
  const numRows = lastRow - startRow + 1;
  const data = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  
  // แปลงข้อมูลเป็น Object โดยใช้ key ตาม header
  const formattedData = data.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      // แปลงวันที่ให้อ่านง่าย
      if (header === 'date' && row[index] instanceof Date) {
         obj[header] = Utilities.formatDate(row[index], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
         obj[header] = row[index];
      }
    });
    return obj;
  });
  
  return formattedData;
}