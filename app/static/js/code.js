const DRIVE_FOLDER_ID = "11As49KFkPT-AcQPt4y6FDcodkzyPVsp7";
const CALL_REPORT_HEADERS = ["RowIndex", "ExpertName", "ReportText", "Date", "Time", "Timestamp"];

const CRM_HEADERS = [
  "ردیف", "نام کارشناس", "تلفن همراه", "نام و نام خانوادگی",
  "کد ملی", "نوع مشتری",
  "محصول مورد نیاز",
  "تاریخ ثبت نام", "تماس", "نتیجه تماس", "وضعیت کارت هوشمند", "سن",
  "وضعیت رتبه اعتباری", "وضعیت تریلی فرسوده", "وضعیت تضامین",
  "چک صیادی", "وضعیت اسقاط", "واریز پیش پرداخت", "وضعیت ترهین تضامین",
  "وضعیت قرار داد", "وضعیت تحویل", "دلایل انصراف",
  "فایل کارت ملی", "فایل استعلام سخا", "فایل فیش واریزی",
  "برند انتخابی", "محور انتخابی", "کاربری انتخابی", "توضیحات محصول",
  "روش پرداخت", "قیمت نهایی",
  "Customer_Paid_Prepayment", "Target_Prepayment",
  "Assigned_By"
];

// Headers for Messages sheet
const MESSAGES_HEADERS = [
  "MessageId",
  "ThreadId",
  "Type",                // expert_to_supervisor | supervisor_to_expert
  "FromUser",
  "ToUser",
  "ExpertUsername",
  "SupervisorUsername",
  "RelatedRowIndex",
  "Text",
  "Timestamp",
  "ReadBySupervisor",
  "ReadByExpert"
];

const VALID_ROLES = ['admin', 'supervisor', 'expert', 'monitoring', 'agency'];

function normalizeRole(roleValue, fallback = 'expert') {
  const normalized = String(roleValue || '').toLowerCase().trim();
  if (VALID_ROLES.includes(normalized)) return normalized;
  const fallbackRole = VALID_ROLES.includes(fallback) ? fallback : 'expert';
  return fallbackRole;
}

function normalizeUserContext(input) {
  if (typeof input === 'string') {
    const username = String(input || '').toLowerCase().trim();
    const role = username === 'admin' ? 'admin' : 'expert';
    return { username, role };
  }

  const username = String(input && input.username ? input.username : '').toLowerCase().trim();
  let role = String(input && input.role ? input.role : '').toLowerCase().trim();

  if (!role && username === 'admin') role = 'admin';
  if (!role) role = 'expert';

  return { username, role };
}

/**
 * Normalize ExpertName / username cell content safely.
 * - Converts to string
 * - Removes zero‑width chars and non‑breaking spaces
 * - Trims and lowercases
 */
function normalizeExpertCell(value) {
  const raw = String(value || "");
  // Remove zero-width space, zero-width joiner, BOM, etc.
  const withoutZeroWidth = raw.replace(/[\u200B-\u200D\uFEFF]/g, "");
  // Convert non‑breaking space to normal space, then trim
  const cleaned = withoutZeroWidth.replace(/\u00A0/g, " ").trim();
  return cleaned.toLowerCase();
}

/**
 * Normalize username safely (lowercase + trim).
 */
function normalizeUsername(value) {
  return String(value || "").toLowerCase().trim();
}

/**
 * Get Messages sheet, create if not exists, ensure headers.
 */
function getMessagesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Messages");
  if (!sheet) {
    sheet = ss.insertSheet("Messages");
    sheet.appendRow(MESSAGES_HEADERS);
  } else {
    const lastCol = sheet.getLastColumn();
    if (lastCol < MESSAGES_HEADERS.length) {
      sheet.getRange(1, 1, 1, MESSAGES_HEADERS.length).setValues([MESSAGES_HEADERS]);
    }
  }
  return sheet;
}

/**
 * Helper: get supervisor username for an expert from Users sheet.
 * @param {string} expertUsername
 * @return {string} supervisor username (normalized) or empty string
 */
function getSupervisorForExpert(expertUsername) {
  const username = normalizeUsername(expertUsername);
  if (!username) return "";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return "";

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return "";

  for (let i = 1; i < data.length; i++) {
    const rowUser = normalizeUsername(data[i][0]);
    if (rowUser === username) {
      const supervisor = normalizeUsername(data[i][4] || "");
      return supervisor;
    }
  }
  return "";
}


function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('index');
  tpl.initialRole = (e && e.parameter && e.parameter.role) ? String(e.parameter.role).toLowerCase() : '';
  return tpl.evaluate()
    .setTitle('CRM دنیای ماموت')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getPricingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Pricing");
  if (!sheet) {
    sheet = ss.insertSheet("Pricing");
    sheet.appendRow(["Brand", "AxleType", "UsageType", "Description", "Price"]);
    sheet.appendRow(["دنیای ماموت", "۳ محور", "یخچالی", "با یونیت", 4500000000]);
  }
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data.map(r => ({
    brand: String(r[0]).trim(),
    axle: String(r[1]).trim(),
    usage: String(r[2]).trim(),
    desc: String(r[3]).trim(),
    price: Number(r[4]) || 0
  }));
}

function getUsersList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data.map(r => ({
    username: r[0],
    name: r[2],
    role: normalizeRole(r[3], 'expert'),
    supervisor: r[4] || ''
  }));
}

/**
 * افزودن کاربر جدید با کنترل دسترسی نقش
 * @param {Object} payload - { username, password, name, role, supervisor, currentUser: { username, role } }
 */
function addUser(payload) {
  const username = payload && payload.username;
  const password = payload && payload.password;
  const name = payload && payload.name;
  const requestedRole = payload && payload.role;
  const supervisor = payload && payload.supervisor;
  const currentUserCtx = normalizeUserContext(payload && payload.currentUser);

  // فقط ادمین یا سرپرست می‌توانند کاربر اضافه کنند
  if (currentUserCtx.role !== 'admin' && currentUserCtx.role !== 'supervisor') {
    return { success: false, message: "مجوز کافی ندارید." };
  }

  const normalizedRole = normalizeRole(requestedRole, 'expert');

  // سرپرست فقط می‌تواند کارشناس اضافه کند
  if (currentUserCtx.role === 'supervisor' && normalizedRole !== 'expert') {
    return { success: false, message: "سرپرست فقط مجاز به افزودن کارشناس است." };
  }

  // تعیین سرپرست مالک کارشناس
  let ownerSupervisor = '';
  if (normalizedRole === 'expert') {
    ownerSupervisor = supervisor ? String(supervisor).toLowerCase().trim() : currentUserCtx.username;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "شیت Users یافت نشد" };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(username).toLowerCase()) {
      return { success: false, message: "نام کاربری تکراری است." };
    }
  }
  sheet.appendRow([username, password, name, normalizedRole, ownerSupervisor]);
  return { success: true };
}

/**
 * حذف کاربر با کنترل نقش
 * @param {Object} payload - { username, currentUser: { username, role } }
 */
function deleteUser(payload) {
  const targetUsername = payload && payload.username;
  const ctx = normalizeUserContext(payload && payload.currentUser);

  if (!targetUsername) return { success: false, message: "کاربر معتبر نیست." };

  // فقط ادمین یا سرپرست مجاز
  if (ctx.role !== 'admin' && ctx.role !== 'supervisor') {
    return { success: false, message: "مجوز کافی ندارید." };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowUser = String(data[i][0] || '').toLowerCase();
    const rowRole = normalizeRole(data[i][3], 'expert');
    const rowSupervisor = String(data[i][4] || '').toLowerCase();

    if (rowUser === String(targetUsername).toLowerCase()) {
      // سرپرست فقط می‌تواند کارشناسان زیرمجموعه خود را حذف کند
      if (ctx.role === 'supervisor') {
        if (rowRole !== 'expert' || rowSupervisor !== ctx.username) {
          return { success: false, message: "مجوز حذف این کاربر را ندارید." };
        }
      }
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: "کاربر یافت نشد." };
}

/**
 * ذخیره گزارش تماس
 * @param {Object} reportData - { rowIndex, expertName, reportText }
 */
function saveCallReport(reportData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("CallReports");
    if (!sheet) {
      sheet = ss.insertSheet("CallReports");
      sheet.appendRow(CALL_REPORT_HEADERS);
    } else {
      // Ensure headers exist
      const lastCol = sheet.getLastColumn();
      if (lastCol < CALL_REPORT_HEADERS.length) {
        sheet.getRange(1, 1, 1, CALL_REPORT_HEADERS.length).setValues([CALL_REPORT_HEADERS]);
      }
    }

    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm");
    const ts = now.getTime();

    sheet.appendRow([
      reportData.rowIndex,
      reportData.expertName,
      reportData.reportText,
      dateStr,
      timeStr,
      ts
    ]);

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * دریافت گزارش‌های تماس برای یک مشتری
 * @param {number} rowIndex
 * @return {Array<Object>}
 */
function getCallReports(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CallReports");
  if (!sheet) return { success: true, reports: [] };

  let data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, reports: [] };

  data.shift(); // remove headers

  const list = data
    .filter(r => Number(r[0]) === Number(rowIndex))
    .map(r => ({
      rowIndex: r[0],
      expertName: r[1],
      reportText: r[2],
      date: formatReportDate(r[3]),
      time: formatReportTime(r[4]),
      timestamp: Number(r[5]) || 0
    }))
    .sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));

  return { success: true, reports: list };
}

function formatReportDate(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy/MM/dd");
  }
  return value || '';
}

function formatReportTime(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }
  return value || '';
}

function loginUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  if (!userSheet) return { success: false, message: "خطا: شیت Users یافت نشد." };
  const data = userSheet.getDataRange().getValues();
  data.shift();
  const normalizedUser = String(username).toLowerCase().trim();
  const normalizedPass = String(password).trim();
  for (let i = 0; i < data.length; i++) {
    const rowUser = String(data[i][0]).toLowerCase().trim();
    const rowPass = String(data[i][1]).trim();
    const rowName = data[i][2];
    const rowRole = normalizeRole(data[i][3], rowUser === 'admin' ? 'admin' : 'expert');
    const rowSupervisor = normalizeUsername(data[i][4] || "");
    if (rowUser === normalizedUser && rowPass === normalizedPass) {
      return {
        success: true,
        username: rowUser,
        fullName: rowName,
        role: rowRole,
        supervisor: rowSupervisor
      };
    }
  }
  return { success: false, message: "نام کاربری یا رمز عبور اشتباه است." };
}

/**
 * Generate a unique MessageId.
 */
function generateMessageId() {
  const ts = new Date().getTime();
  const randomPart = Math.floor(Math.random() * 1000000);
  return "MSG-" + ts + "-" + randomPart;
}

/**
 * Send a message between expert and supervisor.
 * @param {Object} payload
 *  {
 *    type: "expert_to_supervisor" | "supervisor_to_expert",
 *    fromUser,
 *    toUser,
 *    expertUsername,
 *    supervisorUsername,
 *    relatedRowIndex,
 *    text,
 *    threadId
 *  }
 */
function sendMessage(payload) {
  try {
    const type = String(payload && payload.type || "").toLowerCase().trim();
    const allowedTypes = ["expert_to_supervisor", "supervisor_to_expert", "supervisor_to_admin"];
    if (!allowedTypes.includes(type)) {
      return { success: false, message: "نوع پیام نامعتبر است." };
    }

    const fromUser = normalizeUsername(payload && payload.fromUser);
    const toUser = normalizeUsername(payload && payload.toUser);
    const expertUsername = normalizeUsername(payload && payload.expertUsername);
    const supervisorUsername = normalizeUsername(payload && payload.supervisorUsername);
    const relatedRowIndex = Number(payload && payload.relatedRowIndex) || 0;
    const text = String(payload && payload.text || "").trim();
    const clientThreadId = String(payload && payload.threadId || "").trim();

    if (!fromUser || !toUser || !expertUsername || !supervisorUsername || !text) {
      return { success: false, message: "اطلاعات پیام کامل نیست." };
    }

    // Access rules:
    // - expert_to_supervisor: fromUser must equal expertUsername و supervisor مطابق Users باشد.
    // - supervisor_to_expert: fromUser must equal supervisorUsername.
    // - supervisor_to_admin: fromUser must equal supervisorUsername و toUser باید admin باشد.
    if (type === "expert_to_supervisor") {
      if (fromUser !== expertUsername) {
        return { success: false, message: "مجوز ارسال این پیام را ندارید." };
      }
      const mappedSupervisor = getSupervisorForExpert(expertUsername);
      if (mappedSupervisor && mappedSupervisor !== supervisorUsername) {
        return { success: false, message: "سرپرست معتبر برای این کارشناس یافت نشد." };
      }
    } else if (type === "supervisor_to_expert") {
      if (fromUser !== supervisorUsername) {
        return { success: false, message: "مجوز ارسال این پیام را ندارید." };
      }
    } else if (type === "supervisor_to_admin") {
      if (fromUser !== supervisorUsername) {
        return { success: false, message: "مجوز ارسال این پیام را ندارید." };
      }
      if (toUser !== "admin") {
        return { success: false, message: "گیرنده پیام باید ادمین باشد." };
      }
    }

    const sheet = getMessagesSheet();
    const messageId = generateMessageId();
    const threadId = clientThreadId || messageId;
    const now = new Date().getTime();

    const readBySupervisor = type === "supervisor_to_expert"; // در این نوع، سرپرست ارسال‌کننده است
    const readByExpert = type === "expert_to_supervisor";     // در این نوع، کارشناس ارسال‌کننده است

    const row = [
      messageId,
      threadId,
      type,
      fromUser,
      toUser,
      expertUsername,
      supervisorUsername,
      relatedRowIndex,
      text,
      now,
      readBySupervisor,
      readByExpert
    ];

    sheet.appendRow(row);

    return {
      success: true,
      messageId: messageId,
      threadId: threadId
    };
  } catch (e) {
    return { success: false, message: e.message || "خطای غیرمنتظره در ارسال پیام." };
  }
}

/**
 * Get unread counts per expert for a supervisor.
 * @param {string} supervisorUsername
 * @return {Object} { success, counts: Array<{ expertUsername, unreadCount }> }
 */
function getUnreadCountsForSupervisor(supervisorUsername) {
  try {
    const supervisor = normalizeUsername(supervisorUsername);
    if (!supervisor) {
      return { success: true, counts: [] };
    }

    const sheet = getMessagesSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: true, counts: [] };
    }

    const header = data[0];
    const idxType = header.indexOf("Type");
    const idxSupervisor = header.indexOf("SupervisorUsername");
    const idxExpert = header.indexOf("ExpertUsername");
    const idxReadBySupervisor = header.indexOf("ReadBySupervisor");

    const countsMap = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const type = String(row[idxType] || "").toLowerCase().trim();
      const rowSupervisor = normalizeUsername(row[idxSupervisor] || "");
      const rowExpert = normalizeUsername(row[idxExpert] || "");
      const readBySupervisor = row[idxReadBySupervisor] === true;

      if (type === "expert_to_supervisor" && rowSupervisor === supervisor && !readBySupervisor && rowExpert) {
        if (!countsMap[rowExpert]) countsMap[rowExpert] = 0;
        countsMap[rowExpert]++;
      }
    }

    const counts = Object.keys(countsMap).map(expert => ({
      expertUsername: expert,
      unreadCount: countsMap[expert]
    }));

    return { success: true, counts: counts };
  } catch (e) {
    return { success: false, counts: [], message: e.message || "خطا در خواندن پیام‌ها." };
  }
}

/**
 * Get messages (from expert to supervisor) for a supervisor per expert and status.
 * @param {string} supervisorUsername
 * @param {string} expertUsername
 * @param {string} status - 'unread' | 'read' | 'all'
 * @return {Object} { success, messages: [] }
 */
function getMessagesForSupervisor(supervisorUsername, expertUsername, status) {
  try {
    const supervisor = normalizeUsername(supervisorUsername);
    const expert = normalizeUsername(expertUsername);
    const normalizedStatus = String(status || "").toLowerCase().trim() || "unread";

    if (!supervisor || !expert) {
      return { success: true, messages: [] };
    }

    const sheet = getMessagesSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: true, messages: [] };
    }

    const header = data[0];
    const idxId = header.indexOf("MessageId");
    const idxThread = header.indexOf("ThreadId");
    const idxType = header.indexOf("Type");
    const idxFrom = header.indexOf("FromUser");
    const idxTo = header.indexOf("ToUser");
    const idxExpert = header.indexOf("ExpertUsername");
    const idxSupervisor = header.indexOf("SupervisorUsername");
    const idxRowIndex = header.indexOf("RelatedRowIndex");
    const idxText = header.indexOf("Text");
    const idxTs = header.indexOf("Timestamp");
    const idxReadBySupervisor = header.indexOf("ReadBySupervisor");
    const idxReadByExpert = header.indexOf("ReadByExpert");

    const list = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const type = String(row[idxType] || "").toLowerCase().trim();
      if (type !== "expert_to_supervisor") continue;

      const rowExpert = normalizeUsername(row[idxExpert] || "");
      const rowSupervisor = normalizeUsername(row[idxSupervisor] || "");
      if (rowExpert !== expert || rowSupervisor !== supervisor) continue;

      const readBySupervisor = row[idxReadBySupervisor] === true;

      if (normalizedStatus === "unread" && readBySupervisor) continue;
      if (normalizedStatus === "read" && !readBySupervisor) continue;

      list.push({
        messageId: row[idxId],
        threadId: row[idxThread],
        type: type,
        fromUser: row[idxFrom],
        toUser: row[idxTo],
        expertUsername: row[idxExpert],
        supervisorUsername: row[idxSupervisor],
        relatedRowIndex: Number(row[idxRowIndex]) || 0,
        text: row[idxText] || "",
        timestamp: Number(row[idxTs]) || 0,
        readBySupervisor: readBySupervisor,
        readByExpert: row[idxReadByExpert] === true
      });
    }

    list.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));

    return { success: true, messages: list };
  } catch (e) {
    return { success: false, messages: [], message: e.message || "خطا در خواندن پیام‌ها." };
  }
}

/**
 * Mark a message as read by supervisor or expert.
 * @param {string} messageId
 * @param {string} readerRole - 'supervisor' | 'expert'
 * @return {Object} { success, updated?: object }
 */
function markMessageRead(messageId, readerRole) {
  try {
    const id = String(messageId || "").trim();
    const role = String(readerRole || "").toLowerCase().trim();
    if (!id || (role !== "supervisor" && role !== "expert")) {
      return { success: false, message: "ورودی نامعتبر است." };
    }

    const sheet = getMessagesSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: false, message: "پیام یافت نشد." };
    }

    const header = data[0];
    const idxId = header.indexOf("MessageId");
    const idxReadBySupervisor = header.indexOf("ReadBySupervisor");
    const idxReadByExpert = header.indexOf("ReadByExpert");

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[idxId] || "").trim() === id) {
        if (role === "supervisor") {
          row[idxReadBySupervisor] = true;
        } else if (role === "expert") {
          row[idxReadByExpert] = true;
        }
        sheet.getRange(i + 1, 1, 1, header.length).setValues([row]);
        return { success: true, updated: { messageId: id } };
      }
    }

    return { success: false, message: "پیام یافت نشد." };
  } catch (e) {
    return { success: false, message: e.message || "خطا در بروزرسانی وضعیت پیام." };
  }
}

/**
 * Get supervisor replies for an expert.
 * @param {string} expertUsername
 * @param {string} status - 'unread' | 'read' | 'all'
 * @return {Object} { success, messages: [] }
 */
function getSupervisorRepliesForExpert(expertUsername, status) {
  try {
    const expert = normalizeUsername(expertUsername);
    const normalizedStatus = String(status || "").toLowerCase().trim() || "unread";

    if (!expert) {
      return { success: true, messages: [] };
    }

    const sheet = getMessagesSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: true, messages: [] };
    }

    const header = data[0];
    const idxId = header.indexOf("MessageId");
    const idxThread = header.indexOf("ThreadId");
    const idxType = header.indexOf("Type");
    const idxFrom = header.indexOf("FromUser");
    const idxTo = header.indexOf("ToUser");
    const idxExpert = header.indexOf("ExpertUsername");
    const idxSupervisor = header.indexOf("SupervisorUsername");
    const idxRowIndex = header.indexOf("RelatedRowIndex");
    const idxText = header.indexOf("Text");
    const idxTs = header.indexOf("Timestamp");
    const idxReadBySupervisor = header.indexOf("ReadBySupervisor");
    const idxReadByExpert = header.indexOf("ReadByExpert");

    const list = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const type = String(row[idxType] || "").toLowerCase().trim();
      if (type !== "supervisor_to_expert") continue;

      const rowExpert = normalizeUsername(row[idxExpert] || "");
      if (rowExpert !== expert) continue;

      const readByExpert = row[idxReadByExpert] === true;

      if (normalizedStatus === "unread" && readByExpert) continue;
      if (normalizedStatus === "read" && !readByExpert) continue;

      list.push({
        messageId: row[idxId],
        threadId: row[idxThread],
        type: type,
        fromUser: row[idxFrom],
        toUser: row[idxTo],
        expertUsername: row[idxExpert],
        supervisorUsername: row[idxSupervisor],
        relatedRowIndex: Number(row[idxRowIndex]) || 0,
        text: row[idxText] || "",
        timestamp: Number(row[idxTs]) || 0,
        readBySupervisor: row[idxReadBySupervisor] === true,
        readByExpert: readByExpert
      });
    }

    list.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));

    return { success: true, messages: list };
  } catch (e) {
    return { success: false, messages: [], message: e.message || "خطا در خواندن پیام‌های سرپرست." };
  }
}

/**
 * دریافت مشتری‌های دارای کارشناس برای سرپرست جاری (یا ادمین)
 * @param {Object|string} userContext - { username, role }
 * @return {Object[]} داده‌های فیلترشده (فقط ستون‌های موردنیاز) به همراه SheetRowIndex
 */
function getLeadsForSupervisor(userContext) {
  const sheet = getCRMSheet(); // CRM_Leads
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return [];

  const { username, role } = normalizeUserContext(userContext);
  const displayName = (userContext && typeof userContext === 'object') ? String(userContext.displayName || '').toLowerCase().trim() : '';

  const normalizedSupervisor = String(username || '').toLowerCase();
  const isAdmin = role === 'admin';
  const isSupervisor = role === 'supervisor';

  if (!isAdmin && !isSupervisor) return [];

  const headerRow = values[0];
  const headersNormalized = headerRow.map(h => normalizeExpertCell(h));
  const expertNameIndex = headersNormalized.indexOf(normalizeExpertCell("نام کارشناس"));
  const assignedByIndex = headersNormalized.indexOf(normalizeExpertCell("Assigned_By"));

  const result = [];

  // Start from index 1 to skip header row
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sheetRowIndex = i + 1; // 1-based index (Row 1 is values[0])

    // Skip truly empty rows
    if (!row || !row.join('').trim()) continue;

    const rowExpert = normalizeExpertCell(row[expertNameIndex]);
    const rowAssignedBy = normalizeExpertCell(row[assignedByIndex]);

    // فقط ردیف‌هایی با کارشناس و زیرمجموعه همان سرپرست (یا ادمین)
    if (
      !isAdmin &&
      isSupervisor &&
      normalizedSupervisor &&
      rowAssignedBy !== normalizedSupervisor &&
      (!displayName || rowAssignedBy !== normalizeExpertCell(displayName)) &&
      rowExpert !== normalizedSupervisor &&
      (!displayName || rowExpert !== normalizeExpertCell(displayName))
    ) continue;

    result.push({
      rowIndex: sheetRowIndex,
      expertName: row[expertNameIndex] || "",
      phone: String(row[2] || '').trim(),
      fullName: String(row[3] || '').trim(),
      product: String(row[6] || '').trim(),
      status: String(row[9] || '').trim(),
      assignedBy: row[assignedByIndex] || ""
    });
  }

  return result;
}

/**
 * به‌روزرسانی کارشناس یک مشتری در شیت CRM_Leads
 * @param {number} sheetRowIndex - ایندکس ردیف در شیت (۱-بیس، شامل هدر)
 * @param {string} newExpert - نام کارشناس جدید
 * @param {string} assignedBy - نام سرپرست فعلی
 * @return {Object} نتیجه
 */
function updateLeadExpert(sheetRowIndex, newExpert, assignedBy) {
  if (!sheetRowIndex || !newExpert) {
    return { success: false, message: "مقادیر ورودی معتبر نیست." };
  }

  try {
    const sheet = getCRMSheet(); // CRM_Leads
    const expertCol = 2; // "نام کارشناس" ستون دوم (۱-بیس)
    const assignedByCol = CRM_HEADERS.indexOf("Assigned_By") + 1;

    sheet.getRange(sheetRowIndex, expertCol).setValue(newExpert);
    if (assignedByCol > 0) {
      sheet.getRange(sheetRowIndex, assignedByCol).setValue(assignedBy || '');
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getCRMSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let crmSheet = ss.getSheetByName("CRM_Leads");
  if (!crmSheet) {
    const sheets = ss.getSheets();
    if (sheets.length > 0 && sheets[0].getName() === 'Sheet1' && sheets[0].getLastRow() < 2) {
      crmSheet = sheets[0];
      crmSheet.setName("CRM_Leads");
    } else {
      crmSheet = ss.insertSheet("CRM_Leads");
    }
    crmSheet.getRange(1, 1, 1, CRM_HEADERS.length).setValues([CRM_HEADERS]);
  }
  const currentLastCol = crmSheet.getLastColumn();
  if (currentLastCol < CRM_HEADERS.length) {
    crmSheet.getRange(1, 1, 1, CRM_HEADERS.length).setValues([CRM_HEADERS]);
  }
  return crmSheet;
}



function getData(userContext) {
  const sheet = getCRMSheet();
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return { headers: CRM_HEADERS, data: [] };

  const { username, role } = normalizeUserContext(userContext);
  const displayName = (userContext && typeof userContext === 'object') ? String(userContext.displayName || '').toLowerCase().trim() : '';
  const normalizedUsername = String(username || '').toLowerCase();

  const headerRow = values[0];
  const headersNormalized = headerRow.map(h => normalizeExpertCell(h));

  const expertNameIndex = headersNormalized.indexOf(normalizeExpertCell("نام کارشناس"));
  const assignedByIndex = headersNormalized.indexOf(normalizeExpertCell("Assigned_By"));

  const sheetDataWithIndex = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sheetRowIndex = i + 1;

    if (!row || !row.join('').trim()) continue;

    // Pad row for consistency
    while (row.length < headerRow.length) {
      row.push("");
    }

    const rowExpert = normalizeExpertCell(row[expertNameIndex]);
    const rowAssignedBy = normalizeExpertCell(row[assignedByIndex]);

    const includeRow =
      role === 'admin' ||
      role === 'monitoring' ||
      (role === 'supervisor' && normalizedUsername && (rowAssignedBy === normalizedUsername || (displayName && rowAssignedBy === normalizeExpertCell(displayName)) || rowExpert === normalizedUsername || (displayName && rowExpert === normalizeExpertCell(displayName)))) ||
      (role === 'expert' && normalizedUsername && (rowExpert === normalizedUsername || (displayName && rowExpert === normalizeExpertCell(displayName))));

    if (includeRow) {
      // Date formatting
      if (row[7] instanceof Date) {
        row[7] = Utilities.formatDate(
          row[7],
          Session.getScriptTimeZone(),
          "yyyy/MM/dd HH:mm"
        );
      }

      // Add sheetRowIndex at the end
      sheetDataWithIndex.push([...row, sheetRowIndex]);
    }
  }

  return { headers: CRM_HEADERS, data: sheetDataWithIndex };
}

// Monitoring dashboard optimized data (cached, single pass, required columns only)
function getMonitoringDashboardData(startDate, endDate) {
  const cache = CacheService.getDocumentCache();
  const normDigits = (s) => String(s || '').replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d)).replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));
  const cacheKey = 'MONITORING_DASHBOARD_' + normDigits(startDate || '').replace(/[^0-9]/g, '') + '_' + normDigits(endDate || '').replace(/[^0-9]/g, '');
  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      const parsed = JSON.parse(cached);
      parsed.cached = true;
      return parsed;
    }
  } catch (e) {
    // ignore cache read errors
  }

  const sheet = getCRMSheet();
  const lastRow = sheet.getLastRow();
  const responseTemplate = {
    success: true,
    cached: false,
    generatedAt: new Date().toISOString(),
    summary: { total: 0, contract: 0, delivered: 0, cancelled: 0, waiting: 0 },
    brands: [],
    usages: [],
    brandUsage: [],
    cancelByUsage: []
  };

  if (lastRow < 2) return responseTemplate;

  // Columns: contract(20), delivery(21), cancellation(22), brand(26), usage(28), paid(32), target(33), assigned_by(34)
  const rows = sheet.getRange(2, 20, lastRow - 1, 15).getValues();
  // Column 2: ExpertName (single additional read to avoid full sheet read)
  const expertCol = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  // Column 8: Registration Date (for trend / filters)
  const dateCol = sheet.getRange(2, 8, lastRow - 1, 1).getValues();

  const tz = Session.getScriptTimeZone();
  const normStart = normDigits(startDate || '');
  const normEnd = normDigits(endDate || '');
  const startTs = normStart ? new Date(`${normStart}T00:00:00`) : null;
  const endTs = normEnd ? new Date(`${normEnd}T23:59:59`) : null;

  const parseSheetDate = (val) => {
    if (val instanceof Date) return new Date(val.getTime());
    if (typeof val === 'number' && !isNaN(val)) {
      // Google Sheets serial date to JS Date
      const ms = Math.round((val - 25569) * 86400 * 1000);
      return new Date(ms);
    }
    const str = normDigits(String(val || '').trim());
    if (!str) return null;
    const tryParse = (s) => {
      const d = new Date(s);
      return isNaN(d.getTime()) ? null : d;
    };
    const normalized = str.replace(/-/g, '/');
    let d = tryParse(normalized);
    if (!d) d = tryParse(str);
    return d;
  };

  const brandStats = {};
  const usageStats = {};
  const brandUsageStats = {};
  const cancelMap = {};
  const expertStats = {};
  const trendMap = {};

  let total = 0, contract = 0, delivered = 0, cancelled = 0;
  let paidSum = 0, targetSum = 0;

  rows.forEach((row, idx) => {
    const contractVal = String(row[0] || '').toLowerCase();
    const deliveryVal = String(row[1] || '').toLowerCase();
    const cancellation = String(row[2] || '').trim();
    const brand = String(row[6] || '').trim() || 'نامشخص';
    const usage = String(row[8] || '').trim() || 'نامشخص';
    const numFromCell = (v) => {
      const s = normDigits(String(v || '').replace(/,/g, '').trim());
      const n = parseFloat(s);
      return isNaN(n) ? 0 : n;
    };
    const paid = numFromCell(row[12]);
    const target = numFromCell(row[13]);
    const expertName = String(expertCol[idx] && expertCol[idx][0] || '').trim() || 'نامشخص';
    const rawDate = dateCol[idx] && dateCol[idx][0];
    const parsedDate = parseSheetDate(rawDate);
    if (startTs && parsedDate && parsedDate < startTs) return;
    if (endTs && parsedDate && parsedDate > endTs) return;
    const dateKey = parsedDate
      ? Utilities.formatDate(parsedDate, tz, 'yyyy/MM/dd')
      : (String(rawDate || '').trim() || 'نامشخص');

    if (!brandStats[brand]) brandStats[brand] = { leads: 0, sold: 0 };
    if (!usageStats[usage]) usageStats[usage] = { leads: 0, sold: 0 };
    const buKey = `${brand}__${usage}`;
    if (!brandUsageStats[buKey]) brandUsageStats[buKey] = { brand, usage, leads: 0, sold: 0 };
    if (!cancelMap[usage]) cancelMap[usage] = {};
    if (!expertStats[expertName]) expertStats[expertName] = { leads: 0, contracts: 0, deliveries: 0, cancellations: 0, sold: 0 };

    total++;
    paidSum += paid;
    targetSum += target;
    brandStats[brand].leads++;
    usageStats[usage].leads++;
    brandUsageStats[buKey].leads++;
    expertStats[expertName].leads++;
    if (dateKey) trendMap[dateKey] = (trendMap[dateKey] || 0) + 1;

    const isContract = contractVal.indexOf('انجام') !== -1;
    const isDelivered = deliveryVal.indexOf('شد') !== -1;
    const isSold = isContract || isDelivered;

    if (isContract) { contract++; expertStats[expertName].contracts++; }
    if (isDelivered) { delivered++; expertStats[expertName].deliveries++; }

    if (isSold) {
      brandStats[brand].sold++;
      usageStats[usage].sold++;
      brandUsageStats[buKey].sold++;
      expertStats[expertName].sold++;
    }

    if (cancellation) {
      cancelled++;
      cancelMap[usage][cancellation] = (cancelMap[usage][cancellation] || 0) + 1;
      expertStats[expertName].cancellations++;
    }
  });

  const waiting = Math.max(0, total - (contract + delivered + cancelled));

  const brands = Object.keys(brandStats).map(b => ({
    name: b,
    leads: brandStats[b].leads,
    sold: brandStats[b].sold,
    conversion: brandStats[b].leads ? (brandStats[b].sold / brandStats[b].leads) * 100 : 0
  })).sort((a, b) => b.leads - a.leads);

  const usages = Object.keys(usageStats).map(u => ({
    name: u,
    leads: usageStats[u].leads,
    sold: usageStats[u].sold,
    conversion: usageStats[u].leads ? (usageStats[u].sold / usageStats[u].leads) * 100 : 0
  })).sort((a, b) => b.leads - a.leads);

  const brandUsage = Object.values(brandUsageStats).map(item => ({
    brand: item.brand,
    usage: item.usage,
    leads: item.leads,
    sold: item.sold,
    conversion: item.leads ? (item.sold / item.leads) * 100 : 0
  })).sort((a, b) => b.leads - a.leads);

  const cancelByUsage = [];
  const cancellationReasons = {};
  Object.keys(cancelMap).forEach(usage => {
    const reasons = cancelMap[usage];
    const totalCancel = Object.values(reasons).reduce((s, v) => s + v, 0) || 1;
    Object.entries(reasons).forEach(([reason, cnt]) => {
      cancellationReasons[reason] = (cancellationReasons[reason] || 0) + cnt;
      cancelByUsage.push({
        usage,
        reason,
        count: cnt,
        pct: (cnt / totalCancel) * 100
      });
    });
  });

  const trend = Object.keys(trendMap).map(key => ({
    date: key,
    count: trendMap[key]
  })).sort((a, b) => a.date.localeCompare(b.date));

  const experts = Object.keys(expertStats).map(name => {
    const e = expertStats[name];
    return {
      name,
      leads: e.leads,
      contracts: e.contracts,
      deliveries: e.deliveries,
      cancellations: e.cancellations,
      conversion: e.leads ? (e.sold / e.leads) * 100 : 0
    };
  }).sort((a, b) => b.leads - a.leads);

  const cancellationReasonsArr = Object.entries(cancellationReasons).map(([reason, count]) => ({ reason, count }))
    .sort((a, b) => b.count - a.count);

  const insights = [];
  if (brands.length) insights.push(`برند برتر از نظر سرنخ: ${brands[0].name} (${brands[0].leads} مورد)`);
  if (usages.length) insights.push(`کاربری برتر از نظر فروش: ${usages[0].name}`);
  if (cancellationReasonsArr.length) insights.push(`بیشترین دلیل لغو: ${cancellationReasonsArr[0].reason}`);
  if (contract > 0 || delivered > 0) {
    insights.push(`قرارداد/تحویل: ${contract} / ${delivered}`);
  }

  const response = {
    success: true,
    cached: false,
    generatedAt: new Date().toISOString(),
    summary: { total, contract, delivered, cancelled, waiting },
    brands,
    usages,
    brandUsage,
    cancelByUsage,
    financial: { paidSum, targetSum, unpaid: Math.max(targetSum - paidSum, 0) },
    experts,
    cancellationReasons: cancellationReasonsArr,
    insights,
    trend
  };

  try {
    cache.put(cacheKey, JSON.stringify(response), 60);
  } catch (e) {
    // ignore cache write errors
  }

  return response;
}


function updateLeadData(sheetRowIndex, data) {
  const sheet = getCRMSheet();
  const range = sheet.getRange(sheetRowIndex, 1, 1, data.length);
  range.setValues([data]);
  return `ردیف ${sheetRowIndex} با موفقیت ذخیره شد`;
}

function registerNewClient(formData, userContext) {
  const sheet = getCRMSheet();
  const date = new Date();

  let newRow = new Array(CRM_HEADERS.length).fill("");

  newRow[0] = sheet.getLastRow();
  newRow[1] = formData.expertName;
  newRow[2] = formData.phone;
  newRow[3] = `${formData.name} ${formData.lastName}` + "\u200B";
  newRow[4] = formData.nationalId;
  newRow[5] = formData.clientType;
  newRow[6] = formData.requiredProduct;
  newRow[7] = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy/MM/dd HH:mm"
  );

  const actualHeadersNormalized = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => normalizeExpertCell(h));
  const assignedByIndex = actualHeadersNormalized.indexOf(normalizeExpertCell("Assigned_By"));

  if (assignedByIndex !== -1) {
    const ctx = normalizeUserContext(userContext);
    const displayName = (userContext && typeof userContext === 'object') ? String(userContext.displayName || '').trim() : '';
    const assignedByFromPayload = String(formData && formData.assignedBy || '').trim();
    const resolvedAssignedBy = assignedByFromPayload || displayName || ctx.username || "";

    // Only set if not already provided to avoid overriding explicit values
    if (!newRow[assignedByIndex]) {
      newRow[assignedByIndex] = resolvedAssignedBy;
    }
  }

  sheet.appendRow(newRow);

  return `سرنخ جدید (${formData.name} ${formData.lastName}) با موفقیت ثبت شد.`;
}


function deleteData(sheetRowIndex) {
  const sheet = getCRMSheet();
  try {
    if (sheetRowIndex > 1) {
      sheet.deleteRow(sheetRowIndex);
      return { success: true, message: `ردیف ${sheetRowIndex} حذف شد.` };
    }
    return { success: false, message: "امکان حذف وجود ندارد." };
  } catch (e) {
    return { success: false, message: `خطا: ${e.message}` };
  }
}
function handleFileUpload(fileData, sheetRowIndex, targetColumnIndex) {
  if (!DRIVE_FOLDER_ID) return { success: false, message: "شناسه درایو تنظیم نشده." };

  try {
    const mainFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const sheet = getCRMSheet();
    const customerName = sheet.getRange(sheetRowIndex, 4).getValue();

    const folderName = `مشتری ردیف ${sheetRowIndex} - ${customerName}`;

    let targetFolder;
    const folders = mainFolder.getFoldersByName(folderName);
    targetFolder = folders.hasNext() ? folders.next() : mainFolder.createFolder(folderName);

    const bytes = fileData.bytes.map(b => (b < 0 ? b + 256 : b));
    const blob = Utilities.newBlob(bytes, fileData.mimeType, fileData.fileName);
    const file = targetFolder.createFile(blob);
    const fileId = file.getId();

    // 🔥 ایجاد لینک مستقیم قابل کلیک در شیت
    const fileUrl = `https://drive.google.com/file/d/${fileId}/view?usp=sharing`;

    // 🔥 قالب نهایی ذخیره در شیت
    const newEntry = `${fileUrl}|${fileData.fileName}`;

    const excelColNum = targetColumnIndex + 1;
    const range = sheet.getRange(sheetRowIndex, excelColNum);
    const currentVal = String(range.getValue() || "");

    let finalVal = newEntry;
    if (currentVal.trim() !== "") {
      finalVal = currentVal + "," + newEntry;
    }

    range.setValue(finalVal);

    return {
      success: true,
      fileId: fileId,
      fileName: fileData.fileName,
      fullCellValue: finalVal,
      message: "آپلود موفق بود."
    };

  } catch (e) {
    return { success: false, message: `خطا در آپلود: ${e.message}` };
  }
}


// === تابع جدید: حذف فایل از لیست ===
function deleteFileFromSheet(sheetRowIndex, targetColumnIndex, fileIdToRemove) {
  try {
    const sheet = getCRMSheet();
    const excelColNum = targetColumnIndex + 1;
    const range = sheet.getRange(sheetRowIndex, excelColNum);
    const currentVal = String(range.getValue() || "");

    if (!currentVal) return { success: false, message: "فایلی یافت نشد." };

    const parts = currentVal.split(',');
    // فیلتر کردن فایلی که آیدی‌اش باید حذف شود
    const newParts = parts.filter(part => {
      const id = part.split('|')[0];
      return id !== fileIdToRemove;
    });

    const newVal = newParts.join(',');
    range.setValue(newVal);

    return { success: true, fullCellValue: newVal };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * اطلاعات انتقال فرآیند (Process Transfer) را در شیت Process_Transfer ذخیره می‌کند.
 * @param {Object} data - شامل fullName, nationalId, phone, refName
 * @return {boolean} - true در صورت موفقیت
 */
function saveProcessTransfer(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Process_Transfer");

  const HEADERS = ["RowIndex", "FullName", "NationalId", "Phone", "RefName", "Timestamp"];

  if (!sheet) {
    sheet = ss.insertSheet("Process_Transfer");
    sheet.appendRow(HEADERS);
  } else {
    const lastCol = sheet.getLastColumn();
    if (lastCol < HEADERS.length) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }
  }

  const now = new Date().getTime();
  sheet.appendRow([
    data.rowIndex,
    data.fullName,
    data.nationalId,
    data.phone,
    data.refName,
    now
  ]);

  return { success: true };
}

/**
 * برگرداندن داده‌های انتقال فرایند برای لیست rowIndex ها
 * @param {number[]} rowIndices
 * @return {Object} map: { [rowIndex]: { fullName, nationalId, phone, refName, timestamp } }
 */
function getProcessTransfers(rowIndices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Process_Transfer");
  const map = {};
  if (!sheet) return { success: true, transfers: map };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, transfers: map };

  const header = data.shift();
  const idxRow = header.indexOf("RowIndex");
  const idxFull = header.indexOf("FullName");
  const idxNat = header.indexOf("NationalId");
  const idxPhone = header.indexOf("Phone");
  const idxRef = header.indexOf("RefName");
  const idxTs = header.indexOf("Timestamp");

  const requested = new Set((rowIndices || []).map(r => Number(r)));

  data.forEach(r => {
    const rIdx = Number(r[idxRow]);
    if (requested.size === 0 || requested.has(rIdx)) {
      map[rIdx] = {
        rowIndex: rIdx,
        fullName: r[idxFull] || "",
        nationalId: r[idxNat] || "",
        phone: r[idxPhone] || "",
        refName: r[idxRef] || "",
        timestamp: Number(r[idxTs]) || 0
      };
    }
  });

  return { success: true, transfers: map };
}

/**
 * Get or create the agency sheet with proper headers
 * @return {Sheet} The agency sheet
 */
function getAgencySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let agencySheet = ss.getSheetByName("agency");

  if (!agencySheet) {
    agencySheet = ss.insertSheet("agency");
    // Copy headers from CRM_HEADERS
    agencySheet.getRange(1, 1, 1, CRM_HEADERS.length).setValues([CRM_HEADERS]);
  } else {
    // Ensure headers exist
    const lastCol = agencySheet.getLastColumn();
    if (lastCol < CRM_HEADERS.length) {
      agencySheet.getRange(1, 1, 1, CRM_HEADERS.length).setValues([CRM_HEADERS]);
    }
  }

  return agencySheet;
}

/**
 * دریافت لیست یکتای شهرها از شیت AgencyData
 */
function getAgencyCities() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AgencyData");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift(); // remove header if exists
  const cities = [...new Set(data.map(r => String(r[0] || '').trim()).filter(Boolean))];
  return cities;
}

/**
 * دریافت لیست نمایندگی‌های یک شهر از شیت AgencyData
 * @param {string} city
 */
function getAgenciesByCity(city) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AgencyData");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data
    .filter(r => String(r[0] || '').trim() === String(city || '').trim())
    .map(r => String(r[1] || '').trim())
    .filter(Boolean);
}

/**
 * ثبت نمایندگی انتخاب‌شده برای یک ردیف در شیت agency (ستون AI = 35)
 * @param {number} sheetRowIndex - ردیف مشتری در شیت CRM_Leads (۱-بیس)
 * @param {string} city - شهر انتخاب‌شده
 * @param {string} agencyName - نام نمایندگی انتخاب‌شده
 * @return {Object} نتیجه
 */
function transferToAgency(sheetRowIndex, city, agencyName) {
  if (!sheetRowIndex || !agencyName) {
    return { success: false, message: "شهر و نمایندگی انتخاب نشده است." };
  }
  try {
    const crmSheet = getCRMSheet();
    const agencySheet = getAgencySheet();
    const targetCol = 35; // ستون AI

    const crmLastRow = crmSheet.getLastRow();
    if (sheetRowIndex > crmLastRow) {
      return { success: false, message: "ردیف در شیت CRM_Leads یافت نشد." };
    }

    // خواندن ردیف از CRM_Leads
    const rowData = crmSheet.getRange(sheetRowIndex, 1, 1, CRM_HEADERS.length).getValues()[0];
    while (rowData.length < CRM_HEADERS.length) rowData.push("");

    // مقدار ستون نمایندگی (AI) را با نام نمایندگی تنظیم می‌کنیم
    rowData[targetCol - 1] = agencyName;
    // در صورت نیاز می‌توان ستون شهر را هم در همان خانه یا ستون مجزا ذخیره کرد
    // در حال حاضر فقط نام نمایندگی طبق نیاز نوشته می‌شود.

    const agencyLastRow = agencySheet.getLastRow();
    // اگر ردیف معادل در شیت agency وجود ندارد، ردیف جدید اضافه می‌کنیم
    if (sheetRowIndex > agencyLastRow) {
      agencySheet.appendRow(rowData);
    } else {
      agencySheet.getRange(sheetRowIndex, 1, 1, CRM_HEADERS.length).setValues([rowData]);
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Transfer a lead from CRM_Leads to agency sheet
 * @param {number} sheetRowIndex - The row index in CRM_Leads sheet (1-based, including header)
 * @return {Object} { success: boolean, message: string }
 */
function transferLeadToAgency(sheetRowIndex) {
  try {
    const crmSheet = getCRMSheet();
    const agencySheet = getAgencySheet();

    // Read the row from CRM_Leads
    const rowData = crmSheet.getRange(sheetRowIndex, 1, 1, CRM_HEADERS.length).getValues()[0];

    // Ensure row has all columns (pad if needed)
    while (rowData.length < CRM_HEADERS.length) {
      rowData.push("");
    }

    // Append to agency sheet
    agencySheet.appendRow(rowData);

    // Delete from CRM_Leads
    crmSheet.deleteRow(sheetRowIndex);

    return { success: true, message: "مشتری با موفقیت به نمایندگی منتقل شد." };
  } catch (e) {
    return { success: false, message: `خطا: ${e.message}` };
  }
}

/**
 * Get data from agency sheet
 * @return {Object} { headers: Array, data: Array }
 */
function getAgencyData() {
  const agencySheet = getAgencySheet();
  const lastRow = agencySheet.getLastRow();

  if (lastRow < 2) {
    return { headers: CRM_HEADERS, data: [] };
  }

  let values = agencySheet.getDataRange().getValues();

  // Remove empty rows
  values = values.filter(r => r.join('').trim() !== '');

  // Remove header
  const headers = values.shift();

  // Add sheet row index to each row
  const dataWithIndex = values.map((row, index) => {
    const sheetRowIndex = index + 2; // Row 1 is header
    // Pad row if needed
    while (row.length < CRM_HEADERS.length) {
      row.push("");
    }
    // Format date if present
    if (row[7] instanceof Date) {
      row[7] = Utilities.formatDate(
        row[7],
        Session.getScriptTimeZone(),
        "yyyy/MM/dd HH:mm"
      );
    }
    return [...row, sheetRowIndex];
  });

  return { headers: CRM_HEADERS, data: dataWithIndex };
}

/**
 * Assign a lead to an expert (supervisor only)
 * @param {number} sheetRowIndex - The row index in CRM_Leads sheet (1-based, including header)
 * @param {string} expertName - The name of the expert to assign to
 * @param {string} assignedBy - The username of the supervisor assigning
 * @return {Object} { success: boolean, message: string }
 */
function assignLeadToExpert(sheetRowIndex, expertName, assignedBy) {
  try {
    const sheet = getCRMSheet();

    // Read current row
    const rowData = sheet.getRange(sheetRowIndex, 1, 1, CRM_HEADERS.length).getValues()[0];

    // Ensure row has all columns (pad if needed)
    while (rowData.length < CRM_HEADERS.length) {
      rowData.push("");
    }

    // Update expert name (column 1, index 1)
    rowData[1] = expertName;

    // Update Assigned_By column (last column, index CRM_HEADERS.length - 1)
    const assignedByIndex = CRM_HEADERS.length - 1;
    rowData[assignedByIndex] = assignedBy;

    // Write back to sheet
    sheet.getRange(sheetRowIndex, 1, 1, CRM_HEADERS.length).setValues([rowData]);

    return { success: true, message: `مشتری با موفقیت به ${expertName} تخصیص داده شد.` };
  } catch (e) {
    return { success: false, message: `خطا: ${e.message}` };
  }
}

/**
 * Get unassigned leads (ExpertName empty or "admin")
 * Filtered by supervisor ownership if user is a supervisor.
 */
function getUnassignedLeadsInfo(userContext) {
  try {
    const sheet = getCRMSheet();
    const values = sheet.getDataRange().getValues();

    if (values.length < 2) {
      return { count: 0, rows: [] };
    }

    const ctx = normalizeUserContext(userContext);
    const role = ctx.role;
    const username = ctx.username; // lowercase from normalizeUserContext
    const displayName = (userContext && userContext.displayName) ? normalizeExpertCell(userContext.displayName) : '';

    // Normalize headers from the actual sheet to find indexes robustly
    const actualHeadersNormalized = values[0].map(h => normalizeExpertCell(h));

    let expertNameIndex = actualHeadersNormalized.indexOf(normalizeExpertCell("نام کارشناس"));
    if (expertNameIndex === -1) expertNameIndex = 1;

    let assignedByIndex = actualHeadersNormalized.indexOf(normalizeExpertCell("Assigned_By"));
    if (assignedByIndex === -1) assignedByIndex = 15; // Fallback to index 15 from CRM_HEADERS

    const unassignedRows = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const sheetRowIndex = i + 1;

      if (!row || !row.join('').trim()) continue;

      const normalizedExpert = normalizeExpertCell(row[expertNameIndex]);
      const normalizedAssignedBy = normalizeExpertCell(row[assignedByIndex]);

      let includeRow = false;

      if (role === 'admin') {
        // Admin sees everything that is "unassigned" (empty or admin)
        if (normalizedExpert === "" || normalizedExpert === "admin") {
          includeRow = true;
        }
      } else if (role === 'supervisor') {
        // Supervisor sees leads where:
        // 1. ExpertName is THEIR name (they are the current holder)
        const isCurrentHolder = (username && normalizedExpert === username) || (displayName && normalizedExpert === displayName);

        // 2. OR (ExpertName is empty/admin AND Assigned_By is THEIR name)
        const isOwnerOfUnassigned = (normalizedExpert === "" || normalizedExpert === "admin") &&
          ((username && normalizedAssignedBy === username) || (displayName && normalizedAssignedBy === displayName));

        if (isCurrentHolder || isOwnerOfUnassigned) {
          includeRow = true;
        }
      }

      if (includeRow) {
        unassignedRows.push({
          sheetRowIndex: sheetRowIndex,
          rowData: row
        });
      }
    }

    // Sort by row index
    unassignedRows.sort((a, b) => a.sheetRowIndex - b.sheetRowIndex);

    return {
      count: unassignedRows.length,
      rows: unassignedRows
    };
  } catch (e) {
    return { count: 0, rows: [] };
  }
}

/**
 * اطلاعات تخصیص برای ادمین:
 * تعداد ردیف‌هایی که ExpertName آن‌ها admin است.
 */
function getAdminAssignmentInfo(userContext) {
  try {
    const ctx = normalizeUserContext(userContext);
    if (ctx.role !== 'admin') {
      return { success: false, total: 0, message: 'فقط ادمین مجاز است.' };
    }

    const sheet = getCRMSheet();
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      return { success: true, total: 0 };
    }

    const header = values[0];
    const headersNormalized = header.map(h => normalizeExpertCell(h));
    let expertNameIndex = headersNormalized.indexOf(normalizeExpertCell("نام کارشناس"));
    if (expertNameIndex === -1) expertNameIndex = 1;

    let total = 0;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row || !row.join('').trim()) continue;
      const expertName = normalizeUsername(row[expertNameIndex] || '');
      if (expertName === 'admin') total++;
    }

    return { success: true, total: total };
  } catch (e) {
    return { success: false, total: 0, message: e.message || 'خطا در محاسبه تعداد.' };
  }
}

/**
 * اطلاعات تخصیص برای سرپرست:
 * تعداد ردیف‌هایی که ExpertName آن‌ها برابر نام سرپرست فعلی است.
 */
function getSupervisorAssignmentInfo(userContext) {
  try {
    const ctx = normalizeUserContext(userContext);
    if (ctx.role !== 'supervisor') {
      return { success: false, total: 0, message: 'فقط سرپرست مجاز است.' };
    }

    const supervisor = ctx.username;
    const sheet = getCRMSheet();
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      return { success: true, total: 0 };
    }

    const header = values[0];
    const headersNormalized = header.map(h => normalizeExpertCell(h));
    let expertNameIndex = headersNormalized.indexOf(normalizeExpertCell("نام کارشناس"));
    if (expertNameIndex === -1) expertNameIndex = 1;

    let total = 0;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row || !row.join('').trim()) continue;
      const expertName = normalizeUsername(row[expertNameIndex] || '');
      if (expertName === supervisor) total++;
    }

    return { success: true, total: total };
  } catch (e) {
    return { success: false, total: 0, message: e.message || 'خطا در محاسبه تعداد.' };
  }
}

/**
 * ادمین: تخصیص گروهی از admin به سرپرست‌ها.
 * assignments: [{ supervisorUsername, count }]
 */
function adminAssignToSupervisors(assignments, userContext) {
  try {
    const ctx = normalizeUserContext(userContext);
    if (ctx.role !== 'admin') {
      return { success: false, message: 'فقط ادمین مجاز است.' };
    }

    const sheet = getCRMSheet();
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      return { success: false, message: 'مشتری قابل تخصیصی یافت نشد.' };
    }

    const header = values[0];
    const headersNormalized = header.map(h => normalizeExpertCell(h));
    let expertNameIndex = headersNormalized.indexOf(normalizeExpertCell("نام کارشناس"));
    if (expertNameIndex === -1) expertNameIndex = 1;
    let assignedByIndex = headersNormalized.indexOf(normalizeExpertCell("Assigned_By"));
    if (assignedByIndex === -1) assignedByIndex = CRM_HEADERS.indexOf("Assigned_By");

    const adminRows = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row || !row.join('').trim()) continue;
      const expertName = normalizeUsername(row[expertNameIndex] || '');
      if (expertName === 'admin') {
        adminRows.push(i + 1); // sheet row index
      }
    }

    if (adminRows.length === 0) {
      return { success: false, message: 'ردیفی با نام کارشناس admin یافت نشد.' };
    }

    const safeAssignments = Array.isArray(assignments) ? assignments : [];
    let totalRequested = 0;
    safeAssignments.forEach(a => {
      const c = Number(a && a.count) || 0;
      if (c > 0) totalRequested += c;
    });

    if (totalRequested === 0) {
      return { success: false, message: 'هیچ تخصیصی ثبت نشده است.' };
    }
    if (totalRequested > adminRows.length) {
      return {
        success: false,
        message: 'تعداد درخواستی بیشتر از تعداد ردیف‌های موجود با نام کارشناس admin است.'
      };
    }

    const resultMap = {};
    let currentIndex = 0;

    safeAssignments.forEach(a => {
      const supervisorUsername = normalizeUsername(a && a.supervisorUsername);
      const count = Number(a && a.count) || 0;
      if (!supervisorUsername || count <= 0) return;
      if (!resultMap[supervisorUsername]) resultMap[supervisorUsername] = 0;

      for (let i = 0; i < count && currentIndex < adminRows.length; i++) {
        const rowIndex = adminRows[currentIndex];
        const rowValues = sheet.getRange(rowIndex, 1, 1, header.length).getValues()[0];

        rowValues[expertNameIndex] = supervisorUsername;
        if (assignedByIndex >= 0) {
          rowValues[assignedByIndex] = supervisorUsername;
        }

        sheet.getRange(rowIndex, 1, 1, header.length).setValues([rowValues]);
        resultMap[supervisorUsername]++;
        currentIndex++;
      }
    });

    return { success: true, assigned: resultMap };
  } catch (e) {
    return { success: false, message: e.message || 'خطا در تخصیص به سرپرست‌ها.' };
  }
}

/**
 * سرپرست: تخصیص گروهی از خودش به کارشناسان زیرمجموعه.
 * assignments: [{ expertUsername, count }]
 */
function supervisorAssignToExperts(assignments, userContext) {
  try {
    const ctx = normalizeUserContext(userContext);
    if (ctx.role !== 'supervisor') {
      return { success: false, message: 'فقط سرپرست مجاز است.' };
    }
    const supervisor = ctx.username;

    const sheet = getCRMSheet();
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      return { success: false, message: 'مشتری قابل تخصیصی یافت نشد.' };
    }

    const header = values[0];
    const headersNormalized = header.map(h => normalizeExpertCell(h));
    let expertNameIndex = headersNormalized.indexOf(normalizeExpertCell("نام کارشناس"));
    if (expertNameIndex === -1) expertNameIndex = 1;
    let assignedByIndex = headersNormalized.indexOf(normalizeExpertCell("Assigned_By"));
    if (assignedByIndex === -1) assignedByIndex = CRM_HEADERS.indexOf("Assigned_By");

    const sourceRows = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row || !row.join('').trim()) continue;
      const expertName = normalizeUsername(row[expertNameIndex] || '');
      if (expertName === supervisor) {
        sourceRows.push(i + 1); // sheet row index
      }
    }

    if (sourceRows.length === 0) {
      return { success: false, message: 'هیچ ردیفی با نام کارشناس برابر سرپرست فعلی یافت نشد.' };
    }

    const safeAssignments = Array.isArray(assignments) ? assignments : [];
    let totalRequested = 0;
    safeAssignments.forEach(a => {
      const c = Number(a && a.count) || 0;
      if (c > 0) totalRequested += c;
    });

    if (totalRequested === 0) {
      return { success: false, message: 'هیچ تخصیصی ثبت نشده است.' };
    }
    if (totalRequested > sourceRows.length) {
      return {
        success: false,
        message: 'تعداد درخواستی بیشتر از تعداد ردیف‌های متعلق به این سرپرست است.'
      };
    }

    const resultMap = {};
    let currentIndex = 0;

    safeAssignments.forEach(a => {
      const expertUsername = normalizeUsername(a && a.expertUsername);
      const count = Number(a && a.count) || 0;
      if (!expertUsername || count <= 0) return;
      if (!resultMap[expertUsername]) resultMap[expertUsername] = 0;

      for (let i = 0; i < count && currentIndex < sourceRows.length; i++) {
        const rowIndex = sourceRows[currentIndex];
        const rowValues = sheet.getRange(rowIndex, 1, 1, header.length).getValues()[0];

        rowValues[expertNameIndex] = expertUsername;
        if (assignedByIndex >= 0) {
          rowValues[assignedByIndex] = supervisor;
        }

        sheet.getRange(rowIndex, 1, 1, header.length).setValues([rowValues]);
        resultMap[expertUsername]++;
        currentIndex++;
      }
    });

    return { success: true, assigned: resultMap };
  } catch (e) {
    return { success: false, message: e.message || 'خطا در تخصیص به کارشناسان.' };
  }
}

/**
 * Bulk assign leads to experts
 */
function bulkAssignLeads(assignments, userContext) {
  try {
    const ctx = normalizeUserContext(userContext);
    const assignedBy = ctx.username;

    const unassignedInfo = getUnassignedLeadsInfo(userContext);

    if (unassignedInfo.count === 0) {
      return { success: false, message: "مشتری تخصیص‌نشده‌ای برای شما یافت نشد." };
    }

    const totalRequested = assignments.reduce((sum, a) => sum + (Number(a.count) || 0), 0);
    if (totalRequested > unassignedInfo.count) {
      return {
        success: false,
        message: `تعداد درخواستی (${totalRequested}) بیشتر از تعداد مشتری‌های موجود (${unassignedInfo.count}) است.`
      };
    }

    if (totalRequested === 0) {
      return { success: false, message: "لطفاً حداقل یک تخصیص را وارد کنید." };
    }

    const sheet = getCRMSheet();
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headersNormalized = headerRow.map(h => normalizeExpertCell(h));

    let expertNameIndex = headersNormalized.indexOf(normalizeExpertCell("نام کارشناس"));
    if (expertNameIndex === -1) expertNameIndex = 1;

    let assignedByIndex = headersNormalized.indexOf(normalizeExpertCell("Assigned_By"));
    if (assignedByIndex === -1) assignedByIndex = 15;

    const assigned = {};
    let currentIndex = 0;
    const updates = [];

    for (const assignment of assignments) {
      const expertName = String(assignment.expertName || "").trim();
      const count = Number(assignment.count) || 0;
      if (!expertName || count <= 0) continue;

      assigned[expertName] = 0;
      for (let i = 0; i < count && currentIndex < unassignedInfo.rows.length; i++) {
        const lead = unassignedInfo.rows[currentIndex];
        updates.push({ rowIndex: lead.sheetRowIndex, expertName: expertName });
        assigned[expertName]++;
        currentIndex++;
      }
    }

    if (updates.length > 0) {
      updates.sort((a, b) => a.rowIndex - b.rowIndex);
      let batchStart = 0;
      for (let i = 0; i < updates.length; i++) {
        const isLast = (i === updates.length - 1);
        const isContiguous = !isLast && (updates[i + 1].rowIndex === updates[i].rowIndex + 1);
        const sameExpert = !isLast && (updates[i + 1].expertName === updates[i].expertName);

        if (!isContiguous || !sameExpert || isLast) {
          const batchSize = i - batchStart + 1;
          const currentBatch = updates.slice(batchStart, i + 1);

          // Update Expert Name (Column index 1-based)
          const expertNames = currentBatch.map(u => [u.expertName]);
          sheet.getRange(updates[batchStart].rowIndex, expertNameIndex + 1, batchSize, 1).setValues(expertNames);

          // Update Assigned_By (Column index 1-based)
          if (assignedByIndex !== -1) {
            const assignedByValues = Array(batchSize).fill([assignedBy]);
            sheet.getRange(updates[batchStart].rowIndex, assignedByIndex + 1, batchSize, 1).setValues(assignedByValues);
          }
          batchStart = i + 1;
        }
      }
    }

    const remaining = unassignedInfo.count - currentIndex;
    const summaryParts = [];
    for (const expertName in assigned) {
      if (assigned[expertName] > 0) summaryParts.push(`${assigned[expertName]} مشتری به ${expertName}`);
    }

    return {
      success: true,
      assigned: assigned,
      remaining: remaining,
      message: summaryParts.length > 0 ? summaryParts.join("، ") + " تخصیص داده شد." : "تخصیص انجام شد."
    };
  } catch (e) {
    return { success: false, message: `خطا: ${e.message}` };
  }
}

