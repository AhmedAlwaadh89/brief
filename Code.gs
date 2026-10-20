/**
 * استقبال طلبات استمارة العملاء
 * ------------------------------------------------------------
 * 1) افتح Google Sheet جديداً ← الإضافات ← Apps Script، والصق هذا الملف كاملاً.
 * 2) اختر الدالة setup من القائمة العلوية واضغط تشغيل، ووافق على الصلاحيات.
 * 3) نشر ← نشر جديد ← تطبيق ويب:
 *      التنفيذ باسم: أنا      |      من يمكنه الوصول: أي شخص
 *    انسخ رابط تطبيق الويب (ينتهي بـ /exec) وأرسله لـ Claude.
 */

const CONFIG = {
  NOTIFY_EMAIL: '',          // اتركه فارغاً ليصل الإشعار إلى بريد حسابك
  CC_EMAILS: '',             // بريد إضافي يصله الإشعار، وللأكثر من بريد افصل بفاصلة
  SHEET_NAME: 'الطلبات',
  FOLDER_NAME: 'طلبات العملاء',
  MAX_FILE_MB: 10
};

const HEADERS = ['رقم الطلب', 'تاريخ الوصول', 'الحالة', 'الجهة', 'المسؤول', 'واتساب العميل', 'الخدمات', 'الهدف', 'الميزانية', 'موعد التسليم', 'الملفات', 'الطلب الكامل', 'ملاحظات المتابعة', 'آخر تحديث'];
const COL = { code: 1, date: 2, status: 3, phone: 6, files: 11, full: 12, notes: 13, updated: 14 };
const STATUSES = ['جديد', 'قيد التحضير', 'بانتظار موافقة العميل', 'قيد التنفيذ', 'تم التسليم', 'ملغي'];
const STATUS_COLORS = {
  'جديد': '#FDE7B0',
  'قيد التحضير': '#DCE8F7',
  'بانتظار موافقة العميل': '#F6DCC6',
  'قيد التنفيذ': '#D8EEDD',
  'تم التسليم': '#C3E3CC',
  'ملغي': '#E4E4E4'
};

/* ---------- الإعداد (مرة واحدة) ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_NAME, 0);
  sh.setRightToLeft(true);

  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#13232A').setFontColor('#FFFFFF')
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  sh.setRowHeight(1, 36);
  [130, 135, 175, 170, 130, 135, 170, 140, 120, 120, 105, 360, 260, 135].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const rows = sh.getMaxRows() - 1;
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
  sh.getRange(2, COL.status, rows, 1).setDataValidation(statusRule);
  sh.getRange(2, COL.full, rows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sh.getRange(2, 1, rows, HEADERS.length).setVerticalAlignment('top');

  const area = sh.getRange(2, 1, rows, HEADERS.length);
  sh.setConditionalFormatRules(STATUSES.map(s =>
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$C2="' + s + '"')
      .setBackground(STATUS_COLORS[s]).setRanges([area]).build()));

  ['Sheet1', 'الورقة1', 'ورقة1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  getRootFolder_();
  MailApp.getRemainingDailyQuota(); // يطلب صلاحية إرسال البريد أثناء الإعداد
  Logger.log('تم الإعداد. الآن انشر التطبيق كتطبيق ويب.');
}

/* ---------- استقبال الطلب ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const p = JSON.parse(e.postData.contents);
    if (p.website) return json_({ ok: true });            // حقل فخ للرسائل الآلية
    const s = p.summary || {};
    if (!s.company || !s.phone) return json_({ ok: false, error: 'missing_fields' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sh) { setup(); sh = ss.getSheetByName(CONFIG.SHEET_NAME); }

    const code = String(p.code || ('REQ-' + Date.now())).replace(/[^\w-]/g, '').slice(0, 40);
    const last = sh.getLastRow();
    if (last > 1) {
      const codes = sh.getRange(2, COL.code, last - 1, 1).getValues().map(r => String(r[0]));
      if (codes.indexOf(code) !== -1) return json_({ ok: true, code: code, duplicate: true });
    }

    let folderUrl = '';
    const files = Array.isArray(p.files) ? p.files.slice(0, 12) : [];
    if (files.length) {
      const folder = getRootFolder_().createFolder(code + ' - ' + fileName_(s.company));
      files.forEach((f, i) => {
        if (!f || !f.b64) return;
        const bytes = Utilities.base64Decode(f.b64);
        if (bytes.length > CONFIG.MAX_FILE_MB * 1024 * 1024) return;
        const mime = /^(image\/[\w.+-]+|application\/pdf)$/.test(f.mime || '') ? f.mime : 'application/octet-stream';
        folder.createFile(Utilities.newBlob(bytes, mime, fileName_(f.name) || ('file-' + (i + 1))));
      });
      folderUrl = folder.getUrl();
    }

    const now = new Date();
    sh.appendRow([
      code, now, 'جديد',
      cell_(s.company), cell_(s.contact), "'" + digits_(s.phone),
      cell_(s.services), cell_(s.goal), cell_(s.budget), cell_(s.when),
      folderUrl ? '=HYPERLINK("' + folderUrl + '","فتح الملفات")' : '',
      cell_(String(p.message || '').slice(0, 45000)), '', now
    ]);
    const row = sh.getLastRow();
    sh.getRange(row, COL.date).setNumberFormat('yyyy/mm/dd hh:mm');
    sh.getRange(row, COL.updated).setNumberFormat('yyyy/mm/dd hh:mm');

    notify_(p, code, folderUrl, ss, sh, row);
    return json_({ ok: true, code: code });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function doGet() {
  return ContentService.createTextOutput('استقبال الطلبات يعمل');
}

/* تحديث "آخر تحديث" تلقائياً عند تغيير الحالة أو الملاحظات */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== CONFIG.SHEET_NAME || e.range.getRow() < 2) return;
  const c = e.range.getColumn();
  if (c === COL.status || c === COL.notes) {
    sh.getRange(e.range.getRow(), COL.updated).setValue(new Date()).setNumberFormat('yyyy/mm/dd hh:mm');
  }
}

/* ---------- الإشعار ---------- */
function notify_(p, code, folderUrl, ss, sh, row) {
  const s = p.summary || {};
  const to = CONFIG.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  const sheetUrl = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=A' + row;
  const wa = waLink_(s.phone);

  const btn = (href, label, bg, fg) =>
    '<a href="' + href + '" style="display:inline-block;margin:0 0 8px 8px;padding:10px 16px;border-radius:8px;background:' + bg + ';color:' + fg + ';text-decoration:none;font-weight:bold">' + label + '</a>';

  const sections = (p.sections || []).map(sec =>
    '<h3 style="margin:20px 0 6px;font-size:15px;color:#13232A;border-bottom:1px solid #D7DFDE;padding-bottom:4px">' + esc_(sec.title) + '</h3>' +
    '<table style="border-collapse:collapse;width:100%">' +
    (sec.rows || []).map(r =>
      '<tr><td style="padding:4px 0 4px 14px;color:#5A6B71;vertical-align:top;width:34%">' + esc_(r.label) + '</td>' +
      '<td style="padding:4px 0;vertical-align:top;white-space:pre-line">' + esc_(r.value) + '</td></tr>').join('') +
    '</table>').join('');

  const html =
    '<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;font-size:14px;line-height:1.7;color:#13232A;max-width:640px">' +
    '<div style="font-size:13px;color:#5A6B71">طلب جديد · ' + esc_(code) + '</div>' +
    '<h2 style="margin:4px 0 6px;font-size:21px">' + esc_(s.company) + '</h2>' +
    '<p style="margin:0 0 14px">' + esc_(s.contact) + (s.services ? ' · ' + esc_(s.services) : '') + '</p>' +
    '<div>' +
      btn(sheetUrl, 'فتح في جدول المتابعة', '#13232A', '#FFFFFF') +
      (wa ? btn(wa, 'مراسلة العميل على واتساب', '#E8A62A', '#1C1404') : '') +
      (folderUrl ? btn(folderUrl, 'ملفات العميل', '#E9EEED', '#13232A') : '') +
    '</div>' + sections + '</div>';

  const mail = {
    to: to,
    subject: 'طلب جديد: ' + (s.company || '') + (s.services ? ' — ' + s.services : ''),
    htmlBody: html,
    body: String(p.message || ''),
    name: 'استمارة العملاء'
  };
  if (CONFIG.CC_EMAILS) mail.cc = CONFIG.CC_EMAILS;
  MailApp.sendEmail(mail);
}

/* ---------- أدوات ---------- */
function getRootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(CONFIG.FOLDER_NAME);
  props.setProperty('ROOT_FOLDER_ID', f.getId());
  return f;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function digits_(v) {
  return String(v || '').replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[^\d+]/g, '');
}
function waLink_(phone) {
  let d = digits_(phone).replace(/^\+/, '');
  if (!d) return '';
  if (d.indexOf('00') === 0) d = d.slice(2);
  if (d.indexOf('0') === 0) d = '964' + d.slice(1);
  else if (d.length === 10 && d.charAt(0) === '7') d = '964' + d;
  return 'https://wa.me/' + d;
}
function cell_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function fileName_(v) {
  return String(v || '').replace(/[\\\/:*?"<>|]/g, ' ').trim().slice(0, 80);
}
function esc_(v) {
  return String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
