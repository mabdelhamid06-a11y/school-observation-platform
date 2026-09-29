/* منصة المراجعة الداخلية - Google Apps Script - نسخة متعددة المدارس */

const USERS_SHEET = 'Users';
const RECORDS_SHEET = 'Records';
const AUDIT_SHEET = 'AuditLog';
const MAX_PAYLOAD_BYTES = 900000;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('منصة المراجعة الداخلية الذكية')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const user = requireUser_();
    if (body.action === 'health' || body.action === 'context') {
      return json_({ok:true, email:user.email, schoolId:user.schoolId, schoolName:user.schoolName, role:user.role});
    }
    if (body.action === 'getState') return getState_(user);
    if (body.action === 'saveState') return saveState_(body.state, user);
    throw new Error('الإجراء غير معروف.');
  } catch (err) {
    return json_({ok:false, error:String(err.message || err)});
  }
}

/* شغّل هذه الدالة مرة واحدة فقط لإنشاء الأوراق تلقائيًا. */
function setupDatabase() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, USERS_SHEET, ['email','schoolId','schoolName','role','active','createdAt']);
  ensureSheet_(ss, RECORDS_SHEET, ['recordId','schoolId','schoolName','updatedAt','updatedBy','stateJson']);
  ensureSheet_(ss, AUDIT_SHEET, ['timestamp','email','action','schoolId','details']);
  Logger.log('تم إنشاء أوراق Users وRecords وAuditLog بنجاح.');
}

function getState_(user) {
  const sheet = getSpreadsheet_().getSheetByName(RECORDS_SHEET);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]) === user.schoolId) {
      return json_({
        ok:true,
        found:true,
        schoolId:user.schoolId,
        schoolName:user.schoolName,
        state:JSON.parse(values[i][5] || '{}'),
        updatedAt:values[i][3],
        updatedBy:values[i][4]
      });
    }
  }
  return json_({ok:true, found:false, schoolId:user.schoolId, schoolName:user.schoolName, state:null});
}

function saveState_(incomingState, user) {
  if (!incomingState || typeof incomingState !== 'object') throw new Error('بيانات التقييم غير صالحة.');
  const state = JSON.parse(JSON.stringify(incomingState));
  state.settings = state.settings || {};
  state.settings.school = user.schoolName;
  state.settings.schoolId = user.schoolId;
  const serialized = JSON.stringify(state);
  if (Utilities.newBlob(serialized).getBytes().length > MAX_PAYLOAD_BYTES) throw new Error('حجم البيانات أكبر من الحد المسموح.');

  const sheet = getSpreadsheet_().getSheetByName(RECORDS_SHEET);
  const values = sheet.getDataRange().getValues();
  const now = new Date();
  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]) === user.schoolId) { row = i + 1; break; }
  }
  const data = [user.schoolId, user.schoolId, user.schoolName, now, user.email, serialized];
  if (row < 0) sheet.appendRow(data);
  else sheet.getRange(row, 1, 1, data.length).setValues([data]);
  audit_(user, 'saveState', 'تم حفظ بيانات المدرسة');
  return json_({ok:true, schoolId:user.schoolId, schoolName:user.schoolName, updatedAt:now, updatedBy:user.email});
}

function requireUser_() {
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('افتح المنصة بحساب Google مسجل الدخول.');
  const sheet = getSpreadsheet_().getSheetByName(USERS_SHEET);
  if (!sheet) throw new Error('شغّل الدالة setupDatabase مرة واحدة أولًا.');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const rowEmail = String(rows[i][0] || '').toLowerCase().trim();
    const active = String(rows[i][4]).toLowerCase() !== 'false' && rows[i][4] !== false;
    if (rowEmail === email && active) {
      if (!rows[i][1] || !rows[i][2]) throw new Error('حسابك غير مرتبط بمدرسة في ورقة Users.');
      return {
        email:email,
        schoolId:String(rows[i][1]),
        schoolName:String(rows[i][2]),
        role:String(rows[i][3] || 'editor')
      };
    }
  }
  throw new Error('حساب Google هذا غير مسجل في ورقة Users.');
}

function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('يجب فتح Apps Script من داخل Google Sheet نفسه.');
  return ss;
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
}

function audit_(user, action, details) {
  const sheet = getSpreadsheet_().getSheetByName(AUDIT_SHEET);
  if (sheet) sheet.appendRow([new Date(), user.email, action, user.schoolId, details]);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
