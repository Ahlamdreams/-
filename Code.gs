const SETTINGS = {};

/**
 * Ensures a file or folder in Google Drive is publicly accessible for viewing.
 * @param {GoogleAppsScript.Drive.File | GoogleAppsScript.Drive.Folder} driveObject
 */
function makePubliclyViewable(driveObject) {
  try {
    driveObject.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log("Failed to set public sharing for file/folder: " + e.message);
  }
}

/**
 * Loads all settings from the "الإعدادات" sheet.
 */
function loadSettings() {
  if (Object.keys(SETTINGS).length > 0 && SETTINGS.SIGNATURE_FOLDER_ID) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("الإعدادات");
  if (!settingsSheet) throw new Error("Sheet named 'الإعدادات' not found.");
  const settingsData = settingsSheet.getRange("A2:B" + settingsSheet.getLastRow()).getValues();
  settingsData.forEach(row => {
    if (row[0]) SETTINGS[row[0]] = row[1];
  });
  if (!SETTINGS.SIGNATURE_FOLDER_NAME) throw new Error("SIGNATURE_FOLDER_NAME is not defined in الإعدادات sheet.");
  const folders = DriveApp.getFoldersByName(SETTINGS.SIGNATURE_FOLDER_NAME);
  if (folders.hasNext()) {
    const folder = folders.next();
    SETTINGS.SIGNATURE_FOLDER_ID = folder.getId();
    makePubliclyViewable(folder);
  } else {
    const newFolder = DriveApp.createFolder(SETTINGS.SIGNATURE_FOLDER_NAME);
    SETTINGS.SIGNATURE_FOLDER_ID = newFolder.getId();
    makePubliclyViewable(newFolder);
  }
}

/**
 * Serves the main HTML page.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('نظام سجل الاحتياط الذكي')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Gets all the necessary data for the app to start in one single, reliable call from the backend.
 */
function getInitialData() {
  try {
    loadSettings();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const getValuesFromSheet = (sheetName) => {
      const sheet = ss.getSheetByName(sheetName.trim());
      if (!sheet || sheet.getLastRow() < 2) return [];
      return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String).filter(Boolean);
    };
    
    const teacherSheet = ss.getSheetByName('ورقة المعلمات');
    let teacherPhoneMap = {}, teacherSubjectMap = {};
    if (teacherSheet && teacherSheet.getLastRow() > 1) {
        const teacherData = teacherSheet.getRange('A2:C' + teacherSheet.getLastRow()).getValues();
        teacherData.forEach(row => {
            const teacher = String(row[0]).trim();
            if (teacher) {
                teacherPhoneMap[teacher] = String(row[1]).trim();
                teacherSubjectMap[teacher] = String(row[2]).trim();
            }
        });
    }
    const dropdowns = {
      absentTeachers: getValuesFromSheet('المعلمة الغائبة'),
      substituteTeachers: getValuesFromSheet('المعلمة البديلة'),
      periods: getValuesFromSheet('الحصة'),
      classes: getValuesFromSheet('الصف'),
      teacherPhoneMap: teacherPhoneMap,
      teacherSubjectMap: teacherSubjectMap
    };

    const logSheet = ss.getSheetByName(SETTINGS.LOG_SHEET_NAME);
    let allSubstitutes = [];
    if (logSheet && logSheet.getLastRow() > 1) {
      allSubstitutes = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 9).getValues().map(row => {
        const signatureData = typeof row[8] === 'string' ? row[8].trim() : "";
        return {
          date: new Date(row[0]).toISOString().slice(0, 10),
          day: row[1], period: row[2], class: row[3], subject: row[4],
          absentTeacher: row[5], substituteTeacher: row[6], phone: row[7],
          signatureData: signatureData
        };
      });
    }
    
    const period7Counts = {};
    allSubstitutes.filter(s => s.period.toString().includes('7')).forEach(s => {
      if (s.substituteTeacher) {
        period7Counts[s.substituteTeacher] = (period7Counts[s.substituteTeacher] || 0) + 1;
      }
    });
    const mostFrequentPeriod7Teacher = Object.entries(period7Counts).sort((a, b) => b[1] - a[1])[0] || ['لا يوجد', 0];
    
    const monthlyStats = {
      mostFrequentPeriod7Teacher: mostFrequentPeriod7Teacher[0],
      period7Count: mostFrequentPeriod7Teacher[1]
    };
    
    return {
      dropdowns: dropdowns,
      allSubstitutes: allSubstitutes,
      monthlyStats: monthlyStats
    };
  } catch(e) {
    Logger.log("Error in getInitialData: " + e.message);
    return { error: `فشل تحميل البيانات من الخلفية. تأكد من أن أسماء الأوراق صحيحة. الخطأ المسجل: ${e.message}` };
  }
}

function getImageAsBase64() {
  const fileId = "1hIiEd1NAXdOKcMsgvmvXYKl0JJbhi_B9";
  try {
    const image = DriveApp.getFileById(fileId);
    makePubliclyViewable(image);
    const blob = image.getBlob();
    return `data:${blob.getMimeType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
  } catch (e) { return null; }
}

function saveSignatureToDrive(base64Data, teacherName) {
  loadSettings();
  const decoded = Utilities.base64Decode(base64Data.split(',')[1]);
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const fileName = `توقيع-${teacherName}-${timestamp}.png`;
  const blob = Utilities.newBlob(decoded, 'image/png', fileName);
  const folder = DriveApp.getFolderById(SETTINGS.SIGNATURE_FOLDER_ID);
  const file = folder.createFile(blob);
  makePubliclyViewable(file);
  return file.getId();
}

function processForm(formData) {
  try {
    loadSettings();
    const signatureFileId = saveSignatureToDrive(formData.signature, formData.substituteTeacher);
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS.LOG_SHEET_NAME);
    if (!logSheet) throw new Error(`Sheet named "${SETTINGS.LOG_SHEET_NAME}" was not found.`);
    logSheet.appendRow([
      new Date(formData.date), formData.day, formData.period, formData.class,
      formData.subject, formData.absentTeacher, formData.substituteTeacher,
      formData.phone, formData.signature
    ]);
    
    try {
      sendWhatsAppMessage(formData);
    } catch (e) {
      Logger.log("Failed to send WhatsApp notification. Error: " + e.message);
    }
    
    return { status: 'success', message: 'تم تسجيل الاحتياط بنجاح!', signatureFileId: signatureFileId };
  } catch (error) {
    Logger.log(`ERROR in processForm: ${error.toString()}`);
    return { status: 'error', message: `فشل الحفظ: ${error.message}` };
  }
}

function sendWhatsAppMessage(data) {
  loadSettings();
  
  if (!SETTINGS.TWILIO_ACCOUNT_SID || !SETTINGS.TWILIO_AUTH_TOKEN || !SETTINGS.TWILIO_FROM_NUMBER) {
    Logger.log("Twilio settings are incomplete. Skipping WhatsApp notification.");
    return;
  }

  let phoneNumber = String(data.phone).trim();
  const OMAN_COUNTRY_CODE = "+968";

  if (!phoneNumber.startsWith("+")) {
    phoneNumber = OMAN_COUNTRY_CODE + phoneNumber;
  }
  
  const recipientNumber = `whatsapp:${phoneNumber}`;
  const messageBody = `*🔔 إشعار حصة احتياط*\n\nمرحباً أ/${data.substituteTeacher}،\n\nتم إسناد حصة احتياط لكِ بالتفاصيل التالية:\n\n*الحصة:* ${data.period}\n*الصف:* ${data.class}\n*المادة:* ${data.subject}\n\nعطاؤكِ يصنع الفرق. شكراً لتعاونكِ!`;
  
  const payload = { 
    'To': recipientNumber, 
    'From': SETTINGS.TWILIO_FROM_NUMBER, 
    'Body': messageBody 
  };

  const options = { 
    'method': 'post', 
    'payload': payload, 
    'headers': { 
      'Authorization': 'Basic ' + Utilities.base64Encode(`${SETTINGS.TWILIO_ACCOUNT_SID}:${SETTINGS.TWILIO_AUTH_TOKEN}`) 
    } 
  };
  
  const url = `https://api.twilio.com/2010-04-01/Accounts/${SETTINGS.TWILIO_ACCOUNT_SID}/Messages.json`;
  
  UrlFetchApp.fetch(url, options);
  Logger.log(`WhatsApp notification sent to ${recipientNumber}`);
}

function generateAndSavePdfReport(reportType, folderId) {
  try {
    loadSettings();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(SETTINGS.LOG_SHEET_NAME);
    if (!logSheet) throw new Error('Log sheet not found.');

    let data;
    let title;
    let filename;
    
    if (logSheet.getLastRow() < 2) {
      return { status: 'error', message: 'لا توجد بيانات لإنشاء التقرير.' };
    }

    if (reportType === 'today') {
      const today = new Date().toISOString().slice(0, 10);
      data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 9).getValues().filter(row => new Date(row[0]).toISOString().slice(0, 10) === today);
      title = `تقرير حصص الاحتياط لليوم: ${new Date().toLocaleDateString('ar-EG', { year: 'numeric', month: 'long', day: 'numeric' })}`;
      filename = `تقرير_اليوم_${today}.pdf`;
    } else if (reportType === 'month') {
      const thisMonth = new Date().toISOString().slice(0, 7);
      data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 9).getValues().filter(row => new Date(row[0]).toISOString().slice(0, 7) === thisMonth);
      title = `تقرير حصص الاحتياط الشهري: ${new Date().toLocaleDateString('ar-EG', { month: 'long', year: 'numeric' })}`;
      filename = `تقرير_الشهر_${thisMonth}.pdf`;
    } else {
      throw new Error('Invalid report type.');
    }
    
    if (data.length === 0) {
      return { status: 'error', message: 'لا توجد بيانات لإنشاء التقرير.' };
    }

    const htmlTemplate = HtmlService.createTemplateFromFile('report_template');
    const processedData = data.map(row => {
      const signatureData = typeof row[8] === 'string' ? row[8].trim() : "";
      return {
        date: new Date(row[0]).toLocaleDateString('ar-EG'),
        period: row[2],
        class: row[3],
        absentTeacher: row[5],
        substituteTeacher: row[6],
        signatureData: signatureData
      };
    });
    htmlTemplate.data = processedData;
    htmlTemplate.title = title;
    
    const htmlOutput = htmlTemplate.evaluate().getContent();
    const blob = Utilities.newBlob(htmlOutput, MimeType.HTML, filename).getAs('application/pdf');

    const folder = DriveApp.getFolderById(folderId);
    const pdfFile = folder.createFile(blob);
    makePubliclyViewable(pdfFile);
    
    return { status: 'success', url: pdfFile.getUrl() };
  } catch(e) {
    Logger.log("Error in generateAndSavePdfReport: " + e.message);
    return { status: 'error', message: `فشل إنشاء التقرير: ${e.message}` };
  }
}

function saveDailyReportToDrive() {
  const folderId = "1ZWeBdHUCbOpbmyYFYwb3U7IFP8lRuvCh";
  const result = generateAndSavePdfReport('today', folderId);
  Logger.log(result.message);
}

function getTeachersStats() {
  try {
    loadSettings();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(SETTINGS.LOG_SHEET_NAME);
    if (!logSheet || logSheet.getLastRow() < 2) {
      return { error: 'لا توجد بيانات في سجل الاحتياط.' };
    }

    const data = logSheet.getRange(2, 7, logSheet.getLastRow() - 1, 1).getValues().flat().map(String);
    const stats = {};
    data.forEach(teacher => {
      stats[teacher] = (stats[teacher] || 0) + 1;
    });

    const sortedStats = Object.entries(stats).sort(([,a],[,b]) => b - a).map(([teacher, count]) => ({ teacher, count }));
    
    return { stats: sortedStats };
  } catch(e) {
    Logger.log("Error in getTeachersStats: " + e.message);
    return { error: `فشل تحميل إحصائيات المعلمات: ${e.message}` };
  }
}
