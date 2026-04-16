/**
 * ============================================================================
 * ระบบขอรับสิทธิและเบิกเงินค่าเช่าบ้านข้าราชการ (Google Apps Script Backend)
 * ============================================================================
 */

// ตัวแปรตั้งค่า (Config) - ใส่ ID ของ Sheet, Folder และ Template
const CONFIG = {
  SPREADSHEET_ID: "1QQzdZX8fizK0QZqNl1SRUfP1NJgAzPCRVxvB2Bppqno", // แนะนำ: นำ ID ของ Google Sheet มาใส่ตรงนี้ (หรือปล่อยว่าง "" ถ้ารันสคริปต์จาก Sheet โดยตรง)
  UPLOAD_FOLDER_ID: "1cnwFjfW3SuoPVKJJrlIsfdThEXzPX0sc", // ID โฟลเดอร์ Google Drive หลักสำหรับเก็บไฟล์/แฟ้มบุคคล ที่คุณแจ้งไว้
  
  // ID ของ Google Docs Template ที่เตรียมไว้
  TEMPLATE_6005_ID: "1aH7xknzFnjxM3xR3GGly5ck8he-lJhrY1ZCTiGiTL_g",
  TEMPLATE_COMMITTEE_ID: "1W2_1Szlt1uuFDTzmHo9SRKSKtI0tXmQIsUjEO1zidAk",
  TEMPLATE_REPORT_ID: "1zCB-TuGFp4xNRdWofbsyR3AUrUkZitOc",
  TEMPLATE_6006_ID: "133W_ZVGsCM35p4TAIoeKNO7wvQCNVYmXMHNW2HB4GMs"
};

/**
 * 1. ฟังก์ชัน doGet(e) - สำหรับ Render หน้าเว็บ HTML
 * รองรับการทำ Multi-page ผ่าน Query Parameter (let page = 'form_6005';)
 */
function doGet(e) {
  // ตรวจสอบตัวแปร e ป้องกัน Error กรณีที่กดปุ่ม "เรียกใช้" โดยตรงจาก Editor
  let page = 'admin_management'; // ตั้งค่าหน้าเริ่มต้น
  if (e && e.parameter && e.parameter.page) {
    page = e.parameter.page;
  }
  
  try {
    let template = HtmlService.createTemplateFromFile(page);
    return template.evaluate()
      .setTitle('ระบบเบิกจ่ายค่าเช่าบ้านข้าราชการ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    // กรณีหาไฟล์ HTML ไม่เจอ ให้แสดงหน้า Error
    return HtmlService.createHtmlOutput('<h2><center>ไม่พบหน้าเว็บที่คุณต้องการ (Error 404)</center></h2><p><center>' + error.message + '</center></p>');
  }
}

/**
 * 2. ฟังก์ชัน include(filename) 
 * สำหรับเรียกใช้ไฟล์ CSS/JS ย่อยเข้ามาในหน้า HTML หลัก (หากมีการแยกไฟล์)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ============================================================================
 * API สำหรับการจัดการข้อมูล (Database Operations - Google Sheets)
 * ============================================================================
 */

/**
 * 3. บันทึกข้อมูลลง Google Sheets (รองรับฟอร์ม 6005, 6006, รายงาน ฯลฯ)
 */
function saveDataToSheet(dataObject, sheetName) {
  try {
    let ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    // หากไม่มีชีตนี้ ให้สร้างใหม่พร้อมกำหนด Header อัตโนมัติ
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      let headers = Object.keys(dataObject);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#D0E2F3");
    }
    
    // ดึง Header ปัจจุบันเพื่อจับคู่ Key ของ Object ให้ตรงคอลัมน์
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let rowData = [];
    
    // บันทึก Timestamp อัตโนมัติถ้ามีคอลัมน์ Timestamp
    dataObject["Timestamp"] = new Date();
    if (headers.indexOf("Timestamp") === -1) {
      headers.unshift("Timestamp"); // แทรกคอลัมน์ Timestamp ไปหน้าสุด
      sheet.insertColumnBefore(1);
      sheet.getRange(1, 1).setValue("Timestamp").setFontWeight("bold").setBackground("#D0E2F3");
    }

    // จัดเรียงข้อมูลให้ตรงกับ Header
    for (let i = 0; i < headers.length; i++) {
      let key = headers[i];
      let value = dataObject[key] !== undefined ? dataObject[key] : "";
      
      // ถ้าข้อมูลเป็น Array (เช่น Checkbox หลายอัน) ให้แปลงเป็น String ขั้นด้วย Comma
      if (Array.isArray(value)) {
        value = value.join(", ");
      }
      rowData.push(value);
    }
    
    sheet.appendRow(rowData);
    return { status: 'success', message: 'บันทึกข้อมูลลงฐานข้อมูลเรียบร้อยแล้ว' };
    
  } catch (error) {
    console.error("Error in saveDataToSheet: " + error.message);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการบันทึกข้อมูล: ' + error.message };
  }
}

/**
 * ============================================================================
 * API สำหรับการสร้างเอกสารจากเทมเพลต (Document & PDF Generation)
 * ============================================================================
 */

/**
 * 4. สร้างไฟล์ PDF จาก Google Docs Template โดยแทนที่ตัวแปร <<...>>
 * @param {String} formType - ประเภทฟอร์ม ("6005", "COMMITTEE", "REPORT", "6006")
 * @param {Object} dataObject - ข้อมูล JSON ที่ต้องการนำไปแทนที่ในเอกสาร
 * @param {String} subFolderName - ชื่อโฟลเดอร์แฟ้มบุคคล (เช่น ชื่อ-สกุล ผู้ยื่น)
 */
function generatePDFFromTemplate(formType, dataObject, subFolderName) {
  try {
    let templateId = "";
    let fileNamePrefix = "";
    
    // 4.1 เลือก Template ID ให้ตรงกับประเภทเอกสาร
    if (formType === "6005") {
      templateId = CONFIG.TEMPLATE_6005_ID;
      fileNamePrefix = "แบบขอรับสิทธิ_6005_";
    } else if (formType === "COMMITTEE") {
      templateId = CONFIG.TEMPLATE_COMMITTEE_ID;
      fileNamePrefix = "รายงานตรวจสอบ_";
    } else if (formType === "REPORT") {
      templateId = CONFIG.TEMPLATE_REPORT_ID;
      fileNamePrefix = "แบบรายงานข้อมูล_";
    } else if (formType === "6006") {
      templateId = CONFIG.TEMPLATE_6006_ID;
      fileNamePrefix = "แบบขอเบิกเงิน_6006_";
    } else {
      throw new Error("ไม่พบประเภทแบบฟอร์มที่ต้องการสร้างเอกสาร");
    }

    if (!templateId) throw new Error("ยังไม่ได้กำหนด Template ID สำหรับฟอร์มนี้ในระบบ");

    // 4.2 ค้นหาหรือสร้างโฟลเดอร์บุคคล
    let mainFolder = CONFIG.UPLOAD_FOLDER_ID ? DriveApp.getFolderById(CONFIG.UPLOAD_FOLDER_ID) : DriveApp.getRootFolder();
    let targetFolder = mainFolder;
    
    if (subFolderName) {
      let folderIterator = mainFolder.getFoldersByName(subFolderName);
      if (folderIterator.hasNext()) {
        targetFolder = folderIterator.next();
      } else {
        targetFolder = mainFolder.createFolder(subFolderName);
      }
    }

    // 4.3 สร้างสำเนา (Copy) จาก Template ต้นฉบับ
    let timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd_HHmmss");
    let newFileName = fileNamePrefix + (dataObject.fullName || "Unknown") + "_" + timestamp;
    let copiedFile = DriveApp.getFileById(templateId).makeCopy(newFileName, targetFolder);
    let copiedDocId = copiedFile.getId();
    
    // 4.4 เปิดไฟล์สำเนาเพื่อเขียนข้อมูลลงไป
    let doc = DocumentApp.openById(copiedDocId);
    let body = doc.getBody();

    // 4.5 วนลูปนำ Data Object ไปแทนที่ใน <<key>>
    for (let key in dataObject) {
      let placeholder = "<<" + key + ">>";
      let value = dataObject[key];
      
      // จัดการกรณีค่าว่าง หรือเป็น Array
      if (value === undefined || value === null) value = "";
      if (Array.isArray(value)) value = value.join(", ");
      
      // สั่งให้ Google Docs แทนที่คำ
      body.replaceText(placeholder, value);
    }

    // ต้อง Save และ Close ก่อนเพื่อล้าง Buffer ให้ข้อมูลอัปเดตครบถ้วน
    doc.saveAndClose();

    // 4.6 แปลงไฟล์ Docs นั้นเป็น PDF
    let pdfBlob = copiedFile.getAs("application/pdf");
    pdfBlob.setName(newFileName + ".pdf");
    let pdfFile = targetFolder.createFile(pdfBlob);

    // *หมายเหตุ: ถ้าต้องการลบไฟล์ .doc ชั่วคราวออกเพื่อประหยัดพื้นที่ สามารถเปิดคอมเมนต์ด้านล่างได้
    // copiedFile.setTrashed(true); 

    return { 
      status: 'success', 
      message: 'สร้างเอกสารแบบฟอร์มสำเร็จ',
      pdfUrl: pdfFile.getUrl(),
      docUrl: copiedFile.getUrl()
    };

  } catch (error) {
    console.error("Error in generatePDFFromTemplate: " + error.message);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการสร้างเอกสาร PDF: ' + error.message };
  }
}

/**
 * ============================================================================
 * API สำหรับการจัดการไฟล์แนบ (File Uploads - Google Drive)
 * ============================================================================
 */

/**
 * 5. อัปโหลดไฟล์ไปเก็บบน Google Drive (รองรับ Base64 ขนาดสูงสุดเกือบ 50MB)
 */
function uploadFileToDrive(base64Data, fileName, subFolderName) {
  try {
    let splitBase = base64Data.split(',');
    let mimeType = splitBase[0].split(';')[0].replace('data:', '');
    let byteCharacters = Utilities.base64Decode(splitBase[1]);
    let blob = Utilities.newBlob(byteCharacters, mimeType, fileName);
    
    let mainFolder = CONFIG.UPLOAD_FOLDER_ID ? DriveApp.getFolderById(CONFIG.UPLOAD_FOLDER_ID) : DriveApp.getRootFolder();
    let targetFolder = mainFolder;

    if (subFolderName) {
      let folderIterator = mainFolder.getFoldersByName(subFolderName);
      if (folderIterator.hasNext()) {
        targetFolder = folderIterator.next();
      } else {
        targetFolder = mainFolder.createFolder(subFolderName);
      }
    }

    let file = targetFolder.createFile(blob);
    
    return { 
      status: 'success', 
      url: file.getUrl(), 
      id: file.getId(),
      name: file.getName()
    };
    
  } catch (error) {
    console.error("Error in uploadFileToDrive: " + error.message);
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการอัปโหลดไฟล์แนบ: ' + error.message };
  }
}

/**
 * ============================================================================
 * API สำหรับดึงข้อมูลและจัดการเลขที่เอกสาร
 * ============================================================================
 */

/**
 * 6. ค้นหาและดึงข้อมูลจากแบบ 6005 อัตโนมัติ (เพื่อใช้ในหน้า 6006)
 */
function fetch6005DataByApprovalNo(approvalNo) {
  try {
    let ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Form6005_Approved");
    
    if (!sheet) throw new Error("ไม่พบฐานข้อมูลการอนุมัติ 6005");

    let data = sheet.getDataRange().getValues();
    let headers = data[0];
    let approvalNoIndex = headers.indexOf("ApprovalNumber");
    
    if(approvalNoIndex === -1) throw new Error("ไม่พบคอลัมน์เลขที่อนุมัติในระบบ");

    for (let i = 1; i < data.length; i++) {
      if (data[i][approvalNoIndex] === approvalNo) {
        return {
          status: 'success',
          data: {
            fullName: data[i][headers.indexOf("fullName")] || "",
            position: data[i][headers.indexOf("position")] || "",
            department: data[i][headers.indexOf("department")] || "",
            salary: data[i][headers.indexOf("salary")] || "",
            maxAllowance: data[i][headers.indexOf("allowanceAmountApprove")] || "",
            refNo: approvalNo
          }
        };
      }
    }
    
    return { status: 'not_found', message: 'ไม่พบเลขที่อนุมัตินี้ในระบบ' };

  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

/**
 * 7. สร้างและรันเลขที่คำขออัตโนมัติ (ป้องกันปัญหา Concurrency เลขซ้ำด้วย LockService)
 */
function generateRunningNumber(prefix, type) {
  let lock = LockService.getScriptLock();
  lock.waitLock(5000); 
  
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("SystemSettings") || ss.insertSheet("SystemSettings");
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Key", "Value"]);
      sheet.appendRow(["CurrentYear", new Date().getFullYear() + 543]);
      sheet.appendRow(["LastRunning_R", 0]);
      sheet.appendRow(["LastRunning_W", 0]);
    }
    
    let data = sheet.getDataRange().getValues();
    let currentYear = "";
    let runningIndexRow = -1;
    let currentRunning = 0;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === "CurrentYear") currentYear = data[i][1];
      if (data[i][0] === ("LastRunning_" + type)) {
        runningIndexRow = i + 1; 
        currentRunning = parseInt(data[i][1]);
      }
    }

    currentRunning += 1;
    sheet.getRange(runningIndexRow, 2).setValue(currentRunning);
    
    let formattedRunning = ("00000" + currentRunning).slice(-5);
    let finalNumber = `${prefix}-FMD${currentYear}-${type}${formattedRunning}`;
    
    return { status: 'success', refNumber: finalNumber };

  } catch (error) {
    return { status: 'error', message: error.message };
  } finally {
    lock.releaseLock();
  }
}
