/**
 * ============================================================================
 * ระบบขอรับสิทธิและเบิกเงินค่าเช่าบ้านข้าราชการ (Google Apps Script Backend)
 * ============================================================================
 */

// ตัวแปรตั้งค่า (Config) - ใส่ ID ของ Sheet, Folder และ Template
const CONFIG = {
  SPREADSHEET_ID: "1QQzdZX8fizK0QZqNl1SRUfP1NJgAzPCRVxvB2Bppqno", 
  UPLOAD_FOLDER_ID: "1cnwFjfW3SuoPVKJJrlIsfdThEXzPX0sc", 
  
  // ID ของ Google Docs Template ที่เตรียมไว้
  TEMPLATE_6005_ID: "1aH7xknzFnjxM3xR3GGly5ck8he-lJhrY1ZCTiGiTL_g",
  TEMPLATE_COMMITTEE_ID: "1W2_1Szlt1uuFDTzmHo9SRKSKtI0tXmQIsUjEO1zidAk",
  TEMPLATE_REPORT_ID: "1zCB-TuGFp4xNRdWofbsyR3AUrUkZitOc",
  TEMPLATE_6006_ID: "133W_ZVGsCM35p4TAIoeKNO7wvQCNVYmXMHNW2HB4GMs"
};

/**
 * 1. ฟังก์ชัน doGet(e) - บังคับโหลด index.html เสมอ เพื่อเป็นโครงสร้างหลักของ SPA
 */
function doGet(e) {
  try {
    // บังคับโหลดไฟล์ index.html เท่านั้น ห้ามโหลดไฟล์ย่อยตรงๆ
    let template = HtmlService.createTemplateFromFile('index');
    return template.evaluate()
      .setTitle('ระบบเบิกจ่ายค่าเช่าบ้านข้าราชการ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput('<h2><center>เกิดข้อผิดพลาดในการโหลดหน้าเว็บ</center></h2><p><center>' + error.message + '</center></p>');
  }
}

/**
 * 2. ฟังก์ชัน include(filename) 
 * ดึงไฟล์ย่อยมาแสดงตรงกลางหน้า index พร้อมระบบป้องกันเว็บล่ม
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    return '<div style="padding: 20px; background-color: #fee2e2; color: #dc2626; border: 2px dashed #f87171; border-radius: 8px; text-align: center; margin: 20px;">' +
           '<h4><strong>⚠️ ไม่พบไฟล์: ' + filename + '.html</strong></h4>' +
           '<p>กรุณาตรวจสอบว่าสร้างไฟล์นี้ไว้ที่แถบซ้ายมือแล้วหรือยัง</p></div>';
  }
}

/**
 * ============================================================================
 * API สำหรับการจัดการข้อมูล (Database Operations - Google Sheets)
 * ============================================================================
 */

/**
 * 3. บันทึกข้อมูลลง Google Sheets
 */
function saveDataToSheet(dataObject, sheetName) {
  try {
    let ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      let headers = Object.keys(dataObject);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#D0E2F3");
    }
    
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let rowData = [];
    
    dataObject["Timestamp"] = new Date();
    if (headers.indexOf("Timestamp") === -1) {
      headers.unshift("Timestamp"); 
      sheet.insertColumnBefore(1);
      sheet.getRange(1, 1).setValue("Timestamp").setFontWeight("bold").setBackground("#D0E2F3");
    }

    for (let i = 0; i < headers.length; i++) {
      let key = headers[i];
      let value = dataObject[key] !== undefined ? dataObject[key] : "";
      
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

function generatePDFFromTemplate(formType, dataObject, subFolderName) {
  try {
    let templateId = "";
    let fileNamePrefix = "";
    
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

    let timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd_HHmmss");
    let newFileName = fileNamePrefix + (dataObject.fullName || "Unknown") + "_" + timestamp;
    let copiedFile = DriveApp.getFileById(templateId).makeCopy(newFileName, targetFolder);
    let copiedDocId = copiedFile.getId();
    
    let doc = DocumentApp.openById(copiedDocId);
    let body = doc.getBody();

    for (let key in dataObject) {
      let placeholder = "<<" + key + ">>";
      let value = dataObject[key];
      
      if (value === undefined || value === null) value = "";
      if (Array.isArray(value)) value = value.join(", ");
      
      body.replaceText(placeholder, value);
    }

    doc.saveAndClose();

    let pdfBlob = copiedFile.getAs("application/pdf");
    pdfBlob.setName(newFileName + ".pdf");
    let pdfFile = targetFolder.createFile(pdfBlob);

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
