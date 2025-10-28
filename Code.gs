// --- CONFIGURATION ---
const PARENT_FOLDER_ID = "1cocqOqYVWffWc7BvL-9p5Haq3-LLY690";
const SHEET_DATA_NAME = "ALLData_841";
const SHEET_ANS_NAME = "AnsForm";

// --- MAIN FUNCTIONS ---

function doGet(e) {
  let template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle("แบบฟอร์มสำรวจการรื้อถอนฯ เนตชายขอบ Zone C+")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- DATA FETCHING FUNCTIONS --- (เหมือน V5.2)

function getSheetData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DATA_NAME);
    if (!sheet) throw new Error(`ไม่พบชีตชื่อ: ${SHEET_DATA_NAME}`);
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const numRows = lastRow - 1;
    // (V5.3) ปรับปรุงการอ่าน getLastColumn ให้ปลอดภัยขึ้น
    const lastCol = sheet.getLastColumn();
    const numCols = lastCol >= 14 ? 14 : lastCol; // อ่านไม่เกินคอลัมน์ N
    if (numCols === 0) return []; // ชีตไม่มีคอลัมน์

    const dataRange = sheet.getRange(2, 1, numRows, numCols);
    const data = dataRange.getValues();
    
    // A=0, C=2, D=3, K=10, L=11, M=12, N=13
    const necessaryData = data.map(row => {
      // (V5.3) เพิ่มการตรวจสอบความยาวของ row ก่อนเข้าถึง index
      return [
        row[0] || null, // A
        row.length > 2 ? row[2] : null,  // C
        row.length > 3 ? row[3] : null,  // D
        row.length > 10 ? row[10] : null, // K
        row.length > 11 ? row[11] : null, // L
        row.length > 12 ? row[12] : null, // M
        row.length > 13 ? row[13] : null  // N
      ];
    });
    return necessaryData;
    
  } catch (error) {
    Logger.log(`Error in getSheetData: ${error.message}\n${error.stack}`); // (V5.3) Log stack trace
    return { error: true, message: `getSheetData: ${error.message}` };
  }
}

function getAllSavedData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ANS_NAME);
    if (!sheet) throw new Error(`ไม่พบชีตชื่อ: ${SHEET_ANS_NAME}`);

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return {};
    
    const numRows = lastRow - 2;
    // อ่าน A ถึง O (15 คอลัมน์)
    const data = sheet.getRange(3, 1, numRows, 15).getValues(); 

    const savedDataMap = {};
    
    data.forEach((row, index) => {
      // A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10, L=11, M=12, N=13, O=14
      const reqId = row[4]; // คอลัมน์ E
      if (reqId) {
        savedDataMap[reqId] = {
          rowNum: index + 3,
          seq: row[0],
          recorderName: row[2], 
          recorderId:   row[3], 
          snCPE:        row[6], 
          imgCPE:       row[7], 
          snLNB:        row[8], 
          imgLNB:       row[9], 
          imgPole:      row[10], 
          imgFeed:      row[11], 
          imgBase:      row[12], 
          imgSolar:     row[13], 
          notes:        row[14]  
        };
      }
    });
    return savedDataMap;

  } catch (error) {
    Logger.log(`Error in getAllSavedData: ${error.message}\n${error.stack}`); // (V5.3) Log stack trace
    return { error: true, message: `getAllSavedData: ${error.message}` };
  }
}

// --- DATA SAVING FUNCTIONS ---

function saveData(formData) {
  let savedFileUrls = {}; 
  let uploadErrorOccurred = false; 
  let specificUploadErrors = []; // (V5.3) เก็บชื่อไฟล์ที่ Error

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ansSheet = ss.getSheetByName(SHEET_ANS_NAME);
    
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    let provinceFolder;
    const folderIter = parentFolder.getFoldersByName(formData.province);
    
    if (folderIter.hasNext()) provinceFolder = folderIter.next();
    else provinceFolder = parentFolder.createFolder(formData.province);
    
    const timestamp = new Date();
    
    // 2. อัปโหลดไฟล์ (Option A)
    function uploadFileOptionA(fileData, oldLink, deviceKey) {
      if (fileData && fileData.data) {
        const uploadResult = uploadFileToDrive( 
          provinceFolder, fileData, formData.reqId, formData.tambon,
          formData.province, deviceKey, timestamp
        );
        if (uploadResult === null) {
          uploadErrorOccurred = true; 
          specificUploadErrors.push(deviceKey + (fileData.name ? ` (${fileData.name})` : '')); // (V5.3) เก็บชื่อไฟล์ที่ Error
          Logger.log(`Upload failed for ${deviceKey}, using old link: ${oldLink}`);
          return oldLink; // ถ้า Error ใช้อันเก่า
        }
        return uploadResult; 
      }
      return oldLink; 
    }

    savedFileUrls = { 
      fileCPE:   uploadFileOptionA(formData.fileCPE,   formData.oldLinkCPE,   'fileCPE'),
      fileLNB:   uploadFileOptionA(formData.fileLNB,   formData.oldLinkLNB,   'fileLNB'),
      filePole:  uploadFileOptionA(formData.filePole,  formData.oldLinkPole,  'filePole'),
      fileFeed:  uploadFileOptionA(formData.fileFeed,  formData.oldLinkFeed,  'fileFeed'),
      fileBase:  uploadFileOptionA(formData.fileBase,  formData.oldLinkBase,  'fileBase'),
      fileSolar: uploadFileOptionA(formData.fileSolar, formData.oldLinkSolar, 'fileSolar')
    };
    
    // 3. เตรียมข้อมูล
    const timestampStr = Utilities.formatDate(timestamp, "GMT+7", "yyyy-MM-dd HH:mm:ss");
    // (V5.3) ตรวจสอบค่า null ก่อนบันทึก
    const newRowData = [
      timestampStr,             // B
      formData.recorderName || "", // C
      formData.recorderId || "",   // D
      formData.reqId || "",        // E
      formData.province || "",     // F
      formData.snCPE || "",        // G
      savedFileUrls.fileCPE || null, // H (ใช้ null ถ้าไม่มี URL)
      formData.snLNB || "",        // I
      savedFileUrls.fileLNB || null, // J
      savedFileUrls.filePole || null, // K
      savedFileUrls.fileFeed || null, // L
      savedFileUrls.fileBase || null, // M
      savedFileUrls.fileSolar || null, // N
      formData.notes || ""         // O
    ];
    
    // 4. ค้นหาแถวเดิมเพื่อ "Update" หรือ "Insert"
    let rowNumToUpdate = -1;
    let currentSeq = null; 

    if (formData.existingRowNum && !isNaN(parseInt(formData.existingRowNum, 10))) { // (V5.3) ตรวจสอบค่าก่อน parseInt
       rowNumToUpdate = parseInt(formData.existingRowNum, 10);
       if (rowNumToUpdate >= 3) {
           try { // (V5.3) เพิ่ม try-catch ตอนอ่าน seq
             currentSeq = ansSheet.getRange(rowNumToUpdate, 1).getValue();
           } catch(e) { Logger.log(`Could not read seq at row ${rowNumToUpdate}: ${e.message}`); }
       }
    } else {
       const lastRow = ansSheet.getLastRow();
       if (lastRow >= 3) {
         const reqIds = ansSheet.getRange(3, 5, lastRow - 2, 1).getValues();
         for (let i = 0; i < reqIds.length; i++) {
           if (reqIds[i][0] == formData.reqId) {
             rowNumToUpdate = i + 3;
             try { // (V5.3) เพิ่ม try-catch ตอนอ่าน seq
               currentSeq = ansSheet.getRange(rowNumToUpdate, 1).getValue(); 
             } catch(e) { Logger.log(`Could not read seq at row ${rowNumToUpdate}: ${e.message}`); }
             break;
           }
         }
       }
    }

    if (rowNumToUpdate > -1 && rowNumToUpdate >= 3) { // (V5.3) เช็ค >= 3 ด้วย
      // --- UPDATE ---
      ansSheet.getRange(rowNumToUpdate, 2, 1, 14).setValues([newRowData]);
      Logger.log(`Updated row ${rowNumToUpdate} for reqId ${formData.reqId}`);
    } else {
      // --- INSERT ---
      const lastRow = ansSheet.getLastRow();
      const newSeq = lastRow < 3 ? 1 : (ansSheet.getRange(lastRow, 1).getValue() || 0) + 1;
      ansSheet.appendRow([newSeq, ...newRowData]);
      rowNumToUpdate = lastRow + 1;
      currentSeq = newSeq; 
      Logger.log(`Appended new row ${rowNumToUpdate} for reqId ${formData.reqId}`);
    }
    
    // 5. คืนค่าข้อมูลที่อัปเดตแล้ว
    const savedDataMap = {
      [formData.reqId]: {
        rowNum: rowNumToUpdate,
        seq: currentSeq, 
        recorderName: formData.recorderName,
        recorderId: formData.recorderId,
        snCPE: formData.snCPE,
        imgCPE: savedFileUrls.fileCPE, // ใช้ URL ที่สำเร็จ หรือ ลิงก์เก่า หรือ null
        snLNB: formData.snLNB,
        imgLNB: savedFileUrls.fileLNB,
        imgPole: savedFileUrls.filePole,
        imgFeed: savedFileUrls.fileFeed,
        imgBase: savedFileUrls.fileBase,
        imgSolar: savedFileUrls.fileSolar,
        notes: formData.notes
      }
    };

    if (uploadErrorOccurred) {
        const errorFiles = specificUploadErrors.join(', ');
        return { success: true, message: `บันทึกข้อมูล ${formData.reqId} แล้ว แต่เกิดข้อผิดพลาดในการอัปโหลดไฟล์: ${errorFiles}`, savedDataMap: savedDataMap, warning: true };
    } else {
        return { success: true, message: `บันทึกข้อมูล ${formData.reqId} เรียบร้อยแล้ว`, savedDataMap: savedDataMap };
    }

  } catch (error) {
    Logger.log(`Error in saveData: ${error.message}\n${error.stack}`); // (V5.3) Log stack trace
    return { success: false, message: `saveData: ${error.message}` };
  }
}

/**
 * (--- V5.3 UPDATED ---)
 * ฟังก์ชันย่อยสำหรับอัปโหลดไฟล์ (คืนค่า null ถ้า Error)
 * เพิ่มการตรวจสอบ name และ mimeType, เพิ่ม try-catch ครอบ createFile
 */
function uploadFileToDrive(folder, fileData, reqId, tambon, province, deviceKey, timestamp) {
  let fileUrl = null; 
  try {
    // 1. ตรวจสอบข้อมูลเบื้องต้น
    const mimeType = fileData.mimeType || 'application/octet-stream';
    // (V5.3) ตรวจสอบ name ให้ละเอียดขึ้น
    const originalName = fileData.name && fileData.name.trim() !== "" ? fileData.name.trim() : `upload_${deviceKey}.bin`; 
    
    if (!fileData || !fileData.data) {
        Logger.log(`Upload skipped for ${deviceKey}: Missing data.`);
        return null; 
    }

    const base64Data = fileData.data.split(',')[1];
    if (!base64Data) {
        Logger.log(`Upload skipped for ${deviceKey}: Invalid base64 data.`);
        return null; 
    }
    
    const decodedData = Utilities.base64Decode(base64Data);
    
    // (V5.3) ใช้ชื่อที่ตรวจสอบแล้ว
    const blob = Utilities.newBlob(decodedData, mimeType, originalName); 
    
    // 2. ตั้งชื่อไฟล์ใหม่
    const timestampStr = Utilities.formatDate(timestamp, "GMT+7", "yyyyMMdd_HHmmss");
    const shortName = getShortFileName(deviceKey);
    // (V5.3) Sanitize ชื่อไฟล์เล็กน้อย (แทนที่อักขระที่ไม่ปลอดภัย)
    const safeOriginalName = originalName.replace(/[^a-zA-Z0-9.\-_]/g, '_'); 
    const finalFileName = `${reqId}_${tambon}_${province}_${shortName}_${timestampStr}_${safeOriginalName}`; 
    
    // 3. สร้างไฟล์ (เพิ่ม try-catch เฉพาะส่วนนี้)
    let file;
    try {
        file = folder.createFile(blob);
        file.setName(finalFileName);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (createError) {
        Logger.log(`ERROR during createFile/setName for ${deviceKey} (${originalName}): ${createError.message}\n${createError.stack}`);
        return null; // คืนค่า null หากเกิด Error ตอนสร้างไฟล์
    }
    
    // 4. ถ้าสำเร็จ ให้สร้าง URL
    fileUrl = "https://drive.google.com/uc?id=" + file.getId() + "&export=download";
    Logger.log(`Successfully uploaded ${deviceKey}: ${finalFileName} -> ${fileUrl}`);

  } catch (error) {
    // จับ Error อื่นๆ ที่อาจเกิดขึ้นก่อน createFile
    Logger.log(`Unexpected error in uploadFileToDrive for ${deviceKey}: ${error.message}\n${error.stack}`);
    return null; // คืนค่า null ถ้ามี Error อื่น
  }
  return fileUrl; // คืนค่า URL หรือ null
}


/**
 * ฟังก์ชันช่วยตั้งชื่อย่อของไฟล์ (V.3)
 */
function getShortFileName(deviceKey) {
  switch (deviceKey) {
    case "fileCPE": return "1.1CPE";
    case "fileLNB": return "1.2LNB-BUC";
    case "filePole": return "1.3ขาจานดาวเทียม";
    case "fileFeed": return "2.1FeedCable";
    case "fileBase": return "2.2Base";
    case "fileSolar": return "2.3SolarCell";
    default: return "file"; // Fallback
  }
}
