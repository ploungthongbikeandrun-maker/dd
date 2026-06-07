function doGet(e) {
  var output = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);
  
  try {
    var p = e.parameter;
    var action = p.action;
    var userId = p.userId;

    if (!userId) {
      return output.setContent(JSON.stringify({ status: "error", message: "ไม่พบข้อมูล User ID" }));
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. ส่วนบันทึกการลา (คงเดิมไว้)
    if (action === "submitLeave") {
      var sheetLeave = ss.getSheetByName("บันทึกการลา");
      if (!sheetLeave) {
        sheetLeave = ss.insertSheet("บันทึกการลา");
        sheetLeave.appendRow(["วัน-เวลาที่ยื่น", "LINE User ID", "ชื่อ-นามสกุล", "ประเภทการลา", "วันที่เริ่มลา", "ถึงวันที่", "เหตุผลการลา"]);
        SpreadsheetApp.flush();
      }
      sheetLeave.appendRow([new Date(), p.userId, p.name || "-", p.leaveType || "-", p.startDate || "-", p.endDate || "-", p.reason || "-"]);
      return output.setContent(JSON.stringify({ status: "success", message: "บันทึกใบลาสำเร็จ" }));
    }

    // 2. ส่วนดึงข้อมูลบุคคล (ปรับมาใช้ getDisplayValues ดึงค่าตามที่ตาเห็น)
    var sheetEmp = ss.getSheetByName("ข้อมูลบุคคล"); 
    
    var range = sheetEmp.getDataRange();
    var dataValues = range.getValues();         // ดึงค่าดิบ (ใช้เช็ก User ID ขาแรก)
    var displayValues = range.getDisplayValues(); // 🌟 ดึงค่าตามที่แสดงผลหน้าจอ (ขจัดปัญหา Timezone เพี้ยน)
    
    var headers = dataValues[0]; // ดึงหัวตารางมาทำ Key

    for (var i = 1; i < dataValues.length; i++) {
      if (dataValues[i][0] == userId) { // ถ้าพบ LINE ID ตรงกัน
        var empData = {};
        
        for (var j = 0; j < headers.length; j++) {
          // 🌟 เลือกดึงค่าที่เป็นข้อความสวยงามจาก displayValues ตรง ๆ เลย
          // ทำให้วันที่ "01/09/2565" จะส่งไปหน้าบ้านเป็น "01/09/2565" เสมอ ไม่กลายเป็น 2022-08-31...
          empData[headers[j]] = displayValues[i][j]; 
        }
        
        return output.setContent(JSON.stringify({ status: "success", data: empData }));
      }
    }
    
    return output.setContent(JSON.stringify({ status: "error", message: "ไม่พบข้อมูลพนักงานในระบบ" }));

  } catch (error) {
    return output.setContent(JSON.stringify({ status: "error", message: error.toString() }));
  }
}
