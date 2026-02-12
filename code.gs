/**
 * ==================================================================
 * ส่วนที่ 1: การตั้งค่า (CONFIGURATION)
 * ==================================================================
 */
const CONFIG = {
  SHEET_NAME: "New Form",        // ชื่อ Sheet ที่เก็บข้อมูล
  COUPON_IMAGE_URL: "https://hba-th.org/wp-content/uploads/2026/01/coupon-1000.jpg", // รูปคูปอง
  ADMIN_EMAIL: "your-email@example.com", // อีเมลแอดมิน
  EMAIL_NOTIFY: false               // แจ้งเตือนเมลแอดมิน (true/false)
};

const FIELD_EMAIL = "อีเมล";                
const FIELD_NAME = "ชื่อ - นามสกุล";  
const FIELD_PHONE = "เบอร์โทรศัพท์";          

const FIELD_MAP = {
  PHONE: ["เบอร์โทรศัพท์", "Tel", "phone", "Telephone", "Mobile"],
  NAME: ["ชื่อ - นามสกุล", "กรุณากรอกชื่อ - นามสกุล", "Name", "Full Name"],
  EMAIL: ["อีเมล", "Email"],
  DATE: ["Date", "วันที่", "Time", "Timestamp"],
  AGE: ["อายุ", "Age", "Age Group"],
  AREA: ["พื้นที่ใช้สอย (ตร.ม)", "พื้นที่ใช้สอย", "Area"]
};

let postedData = [];
const EXCLUDE_PROPERTY = 'e_gs_exclude';
const ORDER_PROPERTY = 'e_gs_order';
const SHEET_NAME_PROPERTY = 'e_gs_SheetName';

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// ฟังก์ชันสำหรับดึงไฟล์ HTML อื่นๆ มาแสดง (เช่น header, css)
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

function doGet(e) {
  var page = e.parameter.page || 'dashboard';
  var template;

  if (page == 'checkin') {
    template = HtmlService.createTemplateFromFile('page-checkin');
  } else if (page == 'redeem') { 
    // เพิ่ม Route ใหม่สำหรับหน้ารับออเดอร์/Redeem
    template = HtmlService.createTemplateFromFile('page-redeem');
  } else {
    template = HtmlService.createTemplateFromFile('index');
  }

  template.url = getScriptUrl();
  return template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('Home Focus 2026 Command Center')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    let params = JSON.stringify(e.parameter);
    params = JSON.parse(params);
    postedData = params;
    insertToSheet(params); 
    return HtmlService.createHtmlOutput("Success");
  } catch (f) {
    return HtmlService.createHtmlOutput("Error: " + f.toString());
  }
}

/**
 * ==================================================================
 * ส่วนที่ 3: ระบบสแกนและตรวจสอบ (CORE SCANNER LOGIC)
 * ==================================================================
 */
function processCheckIn(scannedData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
  } catch (e) {
    return { success: false, message: "ระบบกำลังทำงานหนัก กรุณาลองใหม่อีกครั้ง" };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0]; 

  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) {
    lock.releaseLock();
    return { success: false, message: "ไม่พบฐานข้อมูลผู้ลงทะเบียน" };
  }

  const headers = data[0];
  const phoneIdx = findColumnIndex(headers, FIELD_MAP.PHONE);
  const nameIdx = findColumnIndex(headers, FIELD_MAP.NAME);
  
  let statusIdx = headers.indexOf("CheckInStatus");
  let timeIdx = headers.indexOf("CheckInTime");

  if (statusIdx === -1) {
    statusIdx = headers.length;
    timeIdx = headers.length + 1;
    sheet.getRange(1, statusIdx + 1).setValue("CheckInStatus").setFontWeight("bold");
    sheet.getRange(1, timeIdx + 1).setValue("CheckInTime").setFontWeight("bold");
  }

  if (phoneIdx === -1) {
    lock.releaseLock();
    return { success: false, message: "System Error: ไม่พบคอลัมน์เบอร์โทรศัพท์ใน Sheet" };
  }

  const cleanScanned = scannedData.toString().replace(/[^0-9]/g, "");
  let foundRows = [];
  
  for (let i = 1; i < data.length; i++) {
    let rowPhone = data[i][phoneIdx].toString().replace(/[^0-9]/g, "");
    if (rowPhone === cleanScanned || (rowPhone.length > 5 && cleanScanned.endsWith(rowPhone))) {
       foundRows.push(i);
    }
  }

  if (foundRows.length === 0) {
    lock.releaseLock();
    return { success: false, message: "❌ ไม่พบข้อมูลการลงทะเบียน" };
  }

  let targetRowIndex = -1;
  let alreadyCheckedInTime = "";
  let userName = "ผู้เข้าร่วมงาน";

  for (let idx of foundRows) {
    const status = data[idx][statusIdx];
    if (!status || status === "") {
      targetRowIndex = idx;
      userName = (nameIdx !== -1 && data[idx][nameIdx]) ? data[idx][nameIdx] : userName;
      break;
    } else {
      if (!alreadyCheckedInTime) alreadyCheckedInTime = data[idx][timeIdx];
    }
  }

  if (targetRowIndex === -1) {
    lock.releaseLock();
    return { success: false, message: "⛔ ท่านนี้เข้างานไปแล้วเมื่อ:\n" + alreadyCheckedInTime };
  }

  const timestamp = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });
  sheet.getRange(targetRowIndex + 1, statusIdx + 1).setValue("Checked-in"); 
  sheet.getRange(targetRowIndex + 1, timeIdx + 1).setValue(timestamp);

  lock.releaseLock();
  return { success: true, message: "✅ ลงทะเบียนสำเร็จ", name: userName, couponUrl: CONFIG.COUPON_IMAGE_URL };
}

/**
 * ==================================================================
 * ส่วนที่ 4: ฟังก์ชันบันทึกข้อมูลและส่งอีเมล (INSERT & EMAIL)
 * ==================================================================
 */
const insertToSheet = (data) => {
  const flat = flattenObject(data);
  const keys = Object.keys(flat);
  const sheetName = getSheetName(data) || CONFIG.SHEET_NAME; 
  const formSheet = getFormSheet(sheetName);
  const headers = getHeaders(formSheet, keys);
  const values = getValues(headers, flat);
  setHeaders(formSheet, headers);
  setValues(formSheet, values);
  if (CONFIG.EMAIL_NOTIFY) sendNotification(data, getSheetURL());
  sendUserConfirmationWithQR(flat); 
}

const sendUserConfirmationWithQR = (flatData) => {
  try {
    const userEmail = flatData[FIELD_EMAIL];
    const userName = flatData[FIELD_NAME] || "ผู้เข้าร่วมงาน";
    const qrData = flatData[FIELD_PHONE] || "NoData"; 

    if (!userEmail) return;

    const today = new Date();
    const thaiDate = today.toLocaleDateString('th-TH', {
      year: 'numeric', 
      month: 'short', 
      day: 'numeric',
      timeZone: 'Asia/Bangkok'
    }); 

    // ปรับปรุงจุดที่ 2: ปรับ URL QR Code ให้รองรับภาษาไทยใน Caption ได้เสถียรขึ้น
    const qrUrl = "https://quickchart.io/qr?text=" + encodeURIComponent(qrData) + 
                  "&size=300&margin=2" + 
                  "&caption=" + encodeURIComponent(userName);

    const logoUrl = "https://hba-th.org/wp-content/uploads/2025/12/3179-269x300-1.jpg";
    const bgPosterUrl = "https://hba-th.org/wp-content/uploads/2026/01/AW-Poster-Focus26_CS6-1.png";

    const subject = "QR Code : เข้าร่วมงาน งานรับสร้างบ้าน Focus 2026";

    // คงข้อความและดีไซน์เดิมของคุณไว้ทุกประการ
    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600&family=Sarabun:wght@300;400;600&display=swap');
          body, table, td, a { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }
          table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; }
          img { -ms-interpolation-mode: bicubic; }
          body { height: 100% !important; margin: 0 !important; padding: 0 !important; width: 100% !important; }
        </style>
      </head>
      <body style="margin: 0; padding: 0; background-color: #222222; font-family: 'Sarabun', sans-serif;">
        
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="
            width: 100%;
            height: 100%;
            min-height: 100vh;
            background-image: url('${bgPosterUrl}');
            background-repeat: no-repeat;
            background-size: cover;
            background-position: center center;
            background-color: #222222;">
          <tr>
            <td align="center" valign="top" style="padding: 50px 10px;">
              
              <table border="0" cellpadding="0" cellspacing="0" width="600" style="
                  background-color: #ffffff; 
                  border-radius: 12px; 
                  overflow: hidden; 
                  box-shadow: 0 15px 35px rgba(0,0,0,0.4);
                  max-width: 600px;
                  width: 100%;">
                
                <tr>
                  <td style="padding: 40px 40px 10px 40px;">
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                      <tr>
                        <td width="80">
                          <img src="${logoUrl}" alt="Logo" width="70" style="display: block; border-radius: 4px;" />
                        </td>
                        <td style="padding-left: 15px; font-family: 'Prompt', sans-serif; font-size: 14px; color: #666;">
                          Home Builder Association<br>
                          ${thaiDate} 
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>

                <tr>
                  <td style="padding: 20px 40px 10px 40px;">
                    <h1 style="font-family: 'Prompt', sans-serif; font-size: 24px; color: #1a1a1a; margin: 0; font-weight: 600; line-height: 1.3;">
                      QR Code : เข้าร่วมงานรับสร้างบ้าน Focus 2026
                    </h1>
                  </td>
                </tr>

                <tr>
                  <td style="padding: 10px 40px;">
                    <div style="height: 1px; width: 100%; background-color: #f0f0f0;"></div>
                  </td>
                </tr>

                <tr>
                  <td style="padding: 20px 40px; font-family: 'Sarabun', sans-serif; font-size: 16px; color: #444; line-height: 1.8;">
                    <p style="margin-bottom: 20px;">เรียน คุณ <strong style="color: #000; font-family: 'Prompt', sans-serif;">${userName}</strong> ที่เคารพ</p>
                    
                    <p>ขอขอบคุณสำหรับการลงทะเบียนเข้าร่วมงาน "งานรับสร้างบ้าน Focus 2026" ระหว่างวันที่ 18-22 มีนาคม 2569 ณ อิมแพ็ค ฮอลล์ 6 เมืองทองธานี จัดโดย สมาคมธุรกิจรับสร้างบ้าน</p>
                    
                    <p>งาน "งานรับสร้างบ้าน Focus 2026" งานเดียวที่รวมบริษัทรับสร้างบ้านชั้นนำไว้มากที่สุด จัดเต็มสิทธิประโยชน์ ส่วนลดและของแถม จากบริษัทรับสร้างบ้านชั้นนํา พร้อมสินเชื่ออัตราดอกเบี้ยพิเศษ เฉพาะในงานนี้เท่านั้น</p>
                    
                    <p>เพื่ออำนวยความสะดวกในการเข้าชมงาน ทางสมาคมธุรกิจรับสร้างบ้าน ได้สร้าง QR Code สำหรับ Check-in เข้าร่วมงาน</p>
                  </td>
                </tr>

                <tr>
                  <td align="center" style="padding: 10px 40px 30px 40px;">
                    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 1px solid #e5e5e5; border-radius: 12px; overflow: hidden;">
                      <tr>
                        <td width="35%" bgcolor="#f9f9f9" style="padding: 25px; border-right: 1px dashed #cccccc; text-align: center;">
                          <p style="font-size: 12px; color: #888; margin: 0; font-family: 'Prompt', sans-serif; text-transform: uppercase;"></p>
                          <p style="font-size: 15px; font-weight: 600; color: #333; margin: 5px 0;"></p>
                        </td>
                        <td width="65%" style="padding: 25px; text-align: center;">
                          <p style="font-family: 'Prompt', sans-serif; font-weight: 600; color: #0056b3; margin: 0 0 10px 0; font-size: 14px; letter-spacing: 1px;">E-TICKET FOR ENTRY</p>
                          <img src="${qrUrl}" alt="QR Code" width="160" style="display: block; margin: 0 auto; border: 4px solid #fff;" />
                          <p style="font-size: 12px; color: #999; margin-top: 10px;">Ref ID: ${qrData}</p>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>

                <tr>
                  <td style="padding: 0 40px 40px 40px; font-family: 'Sarabun', sans-serif; font-size: 15px; color: #444; line-height: 1.8;">
                    <p>แล้วพบกันในงาน "งานรับสร้างบ้าน Focus 2026" ระหว่างวันที่ 18-22 มีนาคม 2569 ณ อิมแพ็ค ฮอลล์ 6 เมืองทองธานี สอบถามรายละเอียดเพิ่มเติมได้ที่ <a href="https://line.me/R/ti/p/@055cuumo?oat_content=url" style="color: #0056b3; font-weight: 600; text-decoration: none;">LINE</a></p>
                    
                    <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 20px; font-size: 14px; color: #666;">
                      จึงเรียนมาเพื่อโปรดพิจารณาและขอแสดงความนับถือ<br><br>
                      <strong>สมาคมธุรกิจรับสร้างบ้าน</strong><br>
                      www.hba-th.org | โทร : 0 2570 0153, 0 2940 2744
                    </div>
                  </td>
                </tr>

              </table>
              
              <p style="font-family: 'Prompt', sans-serif; color: rgba(255,255,255,0.6); font-size: 11px; margin-top: 30px; letter-spacing: 1px;">
                &copy; HOME BUILDER ASSOCIATION | ALL RIGHTS RESERVED
              </p>

            </td>
          </tr>
        </table>
        
      </body>
      </html>
    `;

    MailApp.sendEmail({ to: userEmail, subject: subject, htmlBody: htmlBody, name: "Home Builder Association" });
  } catch (err) { Logger.log("Email Error: " + err.toString()); }
};


/**
 * ==================================================================
 * ส่วนที่ 5: HELPER FUNCTIONS
 * ==================================================================
 */
const flattenObject = (ob) => {
  let toReturn = {};
  for (let i in ob) {
    if (!ob.hasOwnProperty(i)) continue;
    if ((typeof ob[i]) === 'object' && ob[i] !== null) {
      let flatObject = flattenObject(ob[i]);
      for (let x in flatObject) { toReturn[i + '.' + x] = flatObject[x]; }
    } else { toReturn[i] = ob[i]; }
  }
  return toReturn;
};

const getHeaders = (formSheet, keys) => {
  let headers = [];
  try { headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0]; } catch(e) {}
  const newHeaders = keys.filter(h => !headers.includes(h));
  return [...headers.filter(h => h !== ""), ...newHeaders];
};

const getValues = (headers, flat) => headers.map(h => flat[h] || "");

const setHeaders = (sheet, values) => {
  sheet.getRange(1, 1, 1, values.length).setValues([values]).setFontWeight("bold").setHorizontalAlignment("center");
};

const setValues = (sheet, values) => {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(lastRow + 1, 1, 1, values.length).setValues([values]).setHorizontalAlignment("center");
};

const getFormSheet = (sheetName) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
};

const getSheetName = (data) => data[SHEET_NAME_PROPERTY] || data["form_name"] || CONFIG.SHEET_NAME;
const getSheetURL = () => SpreadsheetApp.getActiveSpreadsheet().getUrl();
const sendNotification = (data, url) => {
  MailApp.sendEmail(CONFIG.ADMIN_EMAIL, "New Entry: " + data['form_name'], `New submission in: ${url}`);
};

function findColumnIndex(headers, possibleNames) {
  for (let name of possibleNames) {
    for (let i = 0; i < headers.length; i++) {
      if (headers[i].toString().trim() === name) return i;
    }
  }
  return -1;
}
function getDashboardData(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];

  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  
  // ... (ส่วนประกาศ cols ให้คงเดิมเหมือนใน code ของคุณ) ...
  const cols = {
    status: headers.indexOf("CheckInStatus"),
    phone: findColumnIndex(headers, FIELD_MAP.PHONE),
    budget: findColumnIndex(headers, ["งบประมาณการก่อสร้าง", "Budget"]),
    region: findColumnIndex(headers, ["ภูมิภาคที่ต้องการสร้างบ้าน", "Region"]), 
    decision: findColumnIndex(headers, ["ระยะเวลาการตัดสินใจ", "Decision"]),
    purpose: findColumnIndex(headers, ["วัตถุประสงค์ในการมางาน", "Purpose"]),
    name: findColumnIndex(headers, FIELD_MAP.NAME),
    date: findColumnIndex(headers, FIELD_MAP.DATE),
    age: findColumnIndex(headers, FIELD_MAP.AGE),
    area: findColumnIndex(headers, FIELD_MAP.AREA),
    channel: findColumnIndex(headers, ["No Label channel_checkbox", "channel_checkbox", "ช่องทาง"])
  };

  let stats = {
    total_entries: 0, checked_in: 0, budget: {}, region: {}, 
    decision: {}, purpose: {}, daily_trend: {}, demographics: {}, 
    project_size: {}, channels: {}, registrant_list: [] 
  };

  // --- ปรับปรุง Logic การ Filter ตรงนี้ ---
  const startFilter = (filters && filters.startDate) ? new Date(filters.startDate) : null;
  const endFilter = (filters && filters.endDate) ? new Date(filters.endDate) : null;
  
  // เซ็ตเวลาให้ครอบคลุมทั้งวัน
  if (startFilter) startFilter.setHours(0,0,0,0);
  if (endFilter) endFilter.setHours(23,59,59,999);

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const rowDate = new Date(row[cols.date]);
    
    if (!isNaN(rowDate.getTime())) {
      // กรณี 1: เลือกทั้งเริ่มและสิ้นสุด
      if (startFilter && endFilter) {
        if (rowDate < startFilter || rowDate > endFilter) continue;
      } 
      // กรณี 2: เลือกเฉพาะวันเริ่มต้นวันเดียว (ให้แสดงข้อมูลตั้งแต่วันนั้นเป็นต้นไป หรือเฉพาะวันนั้น)
      // *หากต้องการให้แสดง "เฉพาะวันนั้นวันเดียว" เมื่อกรอกช่องเดียว ให้ปรับ Logic ด้านล่าง*
      else if (startFilter && !endFilter) {
        let onlyStart = new Date(startFilter);
        let onlyEnd = new Date(startFilter);
        onlyEnd.setHours(23,59,59,999);
        if (rowDate < onlyStart || rowDate > onlyEnd) continue;
      }
      // กรณี 3: เลือกเฉพาะวันสิ้นสุด (แสดงข้อมูลทุกอย่างที่เกิดขึ้น "ก่อนหรือภายใน" วันนั้น)
      else if (!startFilter && endFilter) {
        if (rowDate > endFilter) continue;
      }
    }

    // ... (ส่วนประมวลผล stats.total_entries++, stats.checked_in++ ฯลฯ ให้คงเดิม) ...
    // [Copy Code ส่วนประมวลผลเดิมของคุณมาวางต่อที่นี่]
    const phone = row[cols.phone] ? row[cols.phone].replace(/[^0-9]/g, "") : "";
    if (!phone) continue;
    stats.total_entries++;
    if (cols.status !== -1 && row[cols.status] === "Checked-in") stats.checked_in++;
    if (cols.channel !== -1) {
      let chVal = row[cols.channel] || "ไม่ระบุ";
      chVal.split(',').forEach(ch => {
        let cleanCh = ch.trim();
        if (cleanCh && cleanCh !== "on") stats.channels[cleanCh] = (stats.channels[cleanCh] || 0) + 1;
      });
    }
    if (cols.purpose !== -1) { let v = row[cols.purpose] || "ไม่ระบุ"; stats.purpose[v] = (stats.purpose[v] || 0) + 1; }
    if (cols.date !== -1) {
      let dateStr = rowDate.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
      stats.daily_trend[dateStr] = (stats.daily_trend[dateStr] || 0) + 1;
    }
    if (cols.budget !== -1) { let v = row[cols.budget] || "ไม่ระบุ"; stats.budget[v] = (stats.budget[v] || 0) + 1; }
    if (cols.decision !== -1) { let v = row[cols.decision] || "ไม่ระบุ"; stats.decision[v] = (stats.decision[v] || 0) + 1; }
    if (cols.area !== -1) { let v = row[cols.area] || "ไม่ระบุ"; stats.project_size[v] = (stats.project_size[v] || 0) + 1; }
    if (cols.age !== -1) { let v = row[cols.age] || "ไม่ระบุ"; stats.demographics[v] = (stats.demographics[v] || 0) + 1; }

    stats.registrant_list.push({
      name: (cols.name !== -1 && row[cols.name]) ? row[cols.name] : "ไม่ระบุชื่อ",
      phone: phone.substring(0, 3) + "xxx" + phone.substring(phone.length - 3),
      status: (cols.status !== -1 && row[cols.status] === "Checked-in") ? "✅ มาแล้ว" : "รอสแกน",
      timestamp: row[cols.date]
    });
  }
  return stats;
}

function findUserByQr(qrData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME); // Sheet "New Form"
  if (!sheet) return { found: false };

  const data = sheet.getDataRange().getDisplayValues();
  const phoneIdx = findColumnIndex(data[0], FIELD_MAP.PHONE);
  const nameIdx = findColumnIndex(data[0], FIELD_MAP.NAME);

  // ทำความสะอาด QR Code (เอาเฉพาะตัวเลข)
  const cleanQr = qrData.toString().replace(/[^0-9]/g, "");

  for (let i = 1; i < data.length; i++) {
    let rowPhone = data[i][phoneIdx].toString().replace(/[^0-9]/g, "");
    // ตรวจสอบเบอร์โทร (ตรงกัน หรือ ลงท้ายด้วย)
    if (rowPhone === cleanQr || (rowPhone.length > 5 && cleanQr.endsWith(rowPhone))) {
       return {
         found: true,
         name: data[i][nameIdx],
         phone: data[i][phoneIdx]
       };
    }
  }
  return { found: false };
}


// --- ฟังก์ชันใหม่: บันทึกการเยี่ยมชมบูท (Booth Visit) ---
function recordBoothVisit(data) {
  const sheetName = "Booth_Visits"; // ชื่อ Sheet ใหม่
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  // ถ้าไม่มี Sheet ให้สร้างใหม่
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["Timestamp", "Name", "Phone", "Activity"]);
  }
  
  // ค้นหาข้อมูลลูกค้าก่อน
  const user = findUserByQr(data.qr);
  const name = user.found ? user.name : "Unknown/Walk-in";
  const phone = user.found ? user.phone : data.qr;
  
  sheet.appendRow([new Date(), name, phone, "Visit Booth"]);
  return { success: true, name: name };
}

// --- ฟังก์ชันใหม่: บันทึก Order ---
function saveOrder(orderData) {
  const sheetName = "Orders"; // ชื่อ Sheet เก็บออเดอร์
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["Timestamp", "Customer Name", "Phone", "Item", "Value", "Discount", "Net Value"]);
  }

  sheet.appendRow([
    new Date(),
    orderData.name,
    orderData.phone,
    orderData.item,
    orderData.value,
    orderData.discount,
    orderData.net
  ]);
  
  return { success: true };
}
