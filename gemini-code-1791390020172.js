function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  
  // ตรวจสอบว่ามีหัวข้อคอลัมน์หรือยัง หากยังไม่มีให้สร้างให้อัตโนมัติ
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "วันที่-เวลา", 
      "รหัสนักศึกษา (เลขใบเสร็จ)", 
      "ชื่อ-นามสกุล", 
      "เซค (Sec)", 
      "ประเภทเสื้อโปโล", 
      "ไซส์เสื้อโปโล", 
      "ราคาเสื้อโปโล", 
      "เสื้อแจ็คเก็ต", 
      "ไซส์แจ็คเก็ต", 
      "ราคาเสื้อแจ็คเก็ต", 
      "รูปการชำระเงิน", 
      "ราคารวมทั้งสิ้น (บาท)"
    ]);
  }
  
  // บันทึกข้อมูล
  sheet.appendRow([
    new Date(),
    data.studentId,
    data.fullName,
    data.sec,
    data.isMuslim ? "ชาวมุสลิม (แขนยาว)" : "ทั่วไป",
    data.poloSize,
    data.poloPrice,
    data.buyJacket ? "ซื้อแจ็คเก็ต" : "ไม่รับ",
    data.buyJacket ? data.jacketSize : "-",
    data.jacketPrice,
    data.isInstallment ? "ผ่อนชำระ" : "ชำระเต็มจำนวน",
    data.totalPrice
  ]);
  
  return ContentService.createTextOutput(JSON.stringify({"result": "success"}))
    .setMimeType(ContentService.MimeType.JSON);
}