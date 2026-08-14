/**
 * نظام إدارة العمل التطوعي | جامعة تبوك
 *
 * ترتيب الأعمدة:
 * A = طابع زمني
 * B = اسم مقدم/منسق الفرصة
 * C = مسمى الفرصة
 * D = البريد الإلكتروني
 * E = الحالة
 * F = تاريخ إنهاء الفرصة (يُسجل تلقائياً عند إرسال المتطلبات)
 * G = رابط التسجيل (يُدخل عند القبول)
 * H = تاريخ انتهاء الفرصة (يُدخل عند القبول بصيغة YYYY-MM-DD)
 */

const CONFIG = {
  SPREADSHEET_ID: '1tJ0te9q7cpGVI5_drEvbmJINg9OVa_wxeCtTiOvkvsQ',
  SHEET_GID: 884978884,
  MANAGER_EMAIL: 'aa6vdd@gmail.com',

  // بعد أن ترسل ملف Word: ارفعه إلى Google Drive وضع ID الملف هنا.
  REQUIREMENTS_FILE_ID: '',

  COL: {
    TIMESTAMP: 1,
    COORDINATOR: 2,
    OPPORTUNITY: 3,
    EMAIL: 4,
    STATUS: 5,
    COMPLETED_AT: 6,
    REGISTRATION_URL: 7,
    END_DATE: 8
  }
};

/* ========================= WEB APP ========================= */

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = String(p.action || 'ping');
    let result;

    switch (action) {
      case 'ping':
        result = { message: 'Volunteer API is running' };
        break;
      case 'list':
        result = listRequests_();
        break;
      case 'setStatus':
        result = setStatus_(
          Number(p.rowNumber),
          String(p.status || ''),
          String(p.registrationUrl || ''),
          String(p.endDate || '')
        );
        break;
      case 'complete':
        result = complete_(Number(p.rowNumber), false);
        break;
      case 'debug':
        result = debugSheet_();
        break;
      default:
        throw new Error('إجراء غير معروف.');
    }

    return json_({ ok: true, ...result });
  } catch (err) {
    return json_({
      ok: false,
      error: err && err.message ? err.message : String(err)
    });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fakeEvent = { parameter: {} };
    Object.keys(body).forEach(function(key) {
      fakeEvent.parameter[key] = String(body[key] == null ? '' : body[key]);
    });
    return doGet(fakeEvent);
  } catch (err) {
    return json_({ ok: false, error: err && err.message ? err.message : String(err) });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ========================= GOOGLE SHEETS ========================= */

function getSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheets().find(function(s) {
    return Number(s.getSheetId()) === Number(CONFIG.SHEET_GID);
  });
  if (!sheet) throw new Error('تعذر العثور على ورقة Google Sheets المطلوبة.');
  return sheet;
}

function ensureSystemColumns_(sheet) {
  const headers = [
    [CONFIG.COL.STATUS, 'الحالة'],
    [CONFIG.COL.COMPLETED_AT, 'تاريخ إنهاء الفرصة'],
    [CONFIG.COL.REGISTRATION_URL, 'رابط التسجيل'],
    [CONFIG.COL.END_DATE, 'تاريخ انتهاء الفرصة']
  ];

  headers.forEach(function(item) {
    const col = item[0];
    const title = item[1];
    const current = String(sheet.getRange(1, col).getDisplayValue() || '').trim();
    if (!current) sheet.getRange(1, col).setValue(title);
  });
}

function listRequests_() {
  const sheet = getSheet_();
  ensureSystemColumns_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return { count: 0, rows: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
  const rows = [];

  values.forEach(function(row, index) {
    const hasData = row.some(function(cell) { return String(cell || '').trim() !== ''; });
    if (!hasData) return;

    rows.push({
      rowNumber: index + 2,
      coordinatorName: row[1] || '',
      opportunityName: row[2] || '',
      email: row[3] || '',
      timestamp: row[0] || '',
      status: row[4] || 'قيد المراجعة',
      completedAt: row[5] || '',
      registrationUrl: row[6] || '',
      endDate: row[7] || '',
      fields: {
        'طابع زمني': row[0] || '',
        'اسم الفرصة': row[1] || '',
        'مسمى الفرصة': row[2] || '',
        'البريد الإلكتروني': row[3] || '',
        'الحالة': row[4] || '',
        'تاريخ إنهاء الفرصة': row[5] || '',
        'رابط التسجيل': row[6] || '',
        'تاريخ انتهاء الفرصة': row[7] || ''
      }
    });
  });

  rows.reverse();
  return { count: rows.length, rows: rows };
}

/* ========================= قبول / رفض ========================= */

function setStatus_(rowNumber, status, registrationUrl, endDate) {
  if (!rowNumber || rowNumber < 2) throw new Error('رقم الصف غير صالح.');
  if (!['مقبول', 'مرفوض', 'قيد المراجعة'].includes(status)) throw new Error('الحالة غير معتمدة.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet_();
    ensureSystemColumns_(sheet);

    const coordinator = sheet.getRange(rowNumber, CONFIG.COL.COORDINATOR).getDisplayValue();
    const opportunity = sheet.getRange(rowNumber, CONFIG.COL.OPPORTUNITY).getDisplayValue();
    const email = sheet.getRange(rowNumber, CONFIG.COL.EMAIL).getDisplayValue();
    const currentStatus = String(sheet.getRange(rowNumber, CONFIG.COL.STATUS).getDisplayValue() || '').trim();

    if (status === 'مقبول') {
      if (!email) throw new Error('لا يوجد بريد إلكتروني لهذا الطلب.');
      if (!registrationUrl) throw new Error('أدخل رابط التسجيل قبل قبول الفرصة.');
      if (!/^https?:\/\//i.test(registrationUrl)) throw new Error('رابط التسجيل يجب أن يبدأ بـ http:// أو https://');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(endDate)) throw new Error('حدد تاريخ انتهاء الفرصة قبل القبول.');

      // نخزن البيانات قبل الإرسال حتى تكون محفوظة مع الطلب.
      sheet.getRange(rowNumber, CONFIG.COL.REGISTRATION_URL).setValue(registrationUrl);
      sheet.getRange(rowNumber, CONFIG.COL.END_DATE).setNumberFormat('@').setValue(endDate);

      // يمنع تكرار إيميل القبول إذا ضُغط الزر مرتين.
      if (currentStatus !== 'مقبول') {
        sendAcceptedEmail_(email, coordinator, opportunity, registrationUrl, endDate);
      }
    }

    if (status === 'مرفوض') {
      if (!email) throw new Error('لا يوجد بريد إلكتروني لهذا الطلب.');
      if (currentStatus !== 'مرفوض') {
        sendRejectedEmail_(email, coordinator, opportunity);
      }
    }

    sheet.getRange(rowNumber, CONFIG.COL.STATUS).setValue(status);
    SpreadsheetApp.flush();

    if (status === 'مقبول') {
      return { message: 'تم قبول الفرصة وإرسال رابط التسجيل بالبريد الإلكتروني.' };
    }
    if (status === 'مرفوض') {
      return { message: 'تم رفض الفرصة وإرسال إشعار الرفض بالبريد الإلكتروني.' };
    }
    return { message: 'تم إعادة الطلب إلى قيد المراجعة.' };
  } finally {
    lock.releaseLock();
  }
}

/* ========================= إنهاء وإرسال ملف Word ========================= */

function complete_(rowNumber, automatic) {
  if (!rowNumber || rowNumber < 2) throw new Error('رقم الصف غير صالح.');
  if (!CONFIG.REQUIREMENTS_FILE_ID) throw new Error('ملف Word الخاص بمتطلبات إنهاء الفرصة لم تتم إضافته بعد.');

  const sheet = getSheet_();
  ensureSystemColumns_(sheet);

  const coordinator = sheet.getRange(rowNumber, CONFIG.COL.COORDINATOR).getDisplayValue();
  const opportunity = sheet.getRange(rowNumber, CONFIG.COL.OPPORTUNITY).getDisplayValue();
  const email = sheet.getRange(rowNumber, CONFIG.COL.EMAIL).getDisplayValue();
  const currentStatus = String(sheet.getRange(rowNumber, CONFIG.COL.STATUS).getDisplayValue() || '').trim();

  if (!email) throw new Error('لا يوجد بريد إلكتروني لهذا الطلب.');
  if (['منتهية', 'منتهي', 'منتهٍ'].includes(currentStatus)) {
    return { message: 'الفرصة مسجلة كمنتهية مسبقاً.' };
  }

  const file = DriveApp.getFileById(CONFIG.REQUIREMENTS_FILE_ID);

  MailApp.sendEmail({
    to: email,
    subject: 'متطلبات استكمال الفرصة التطوعية — جامعة تبوك',
    body:
      'السلام عليكم ورحمة الله وبركاته\n\n' +
      'انتهت الفرصة التطوعية: ' + opportunity + '\n\n' +
      'نأمل تعبئة ملف المتطلبات المرفق واستكمال البيانات المطلوبة.\n\n' +
      'إدارة العمل التطوعي - جامعة تبوك',
    htmlBody: emailTemplate_(coordinator,
      `تم تسجيل انتهاء الفرصة التطوعية <b>${escapeHtml_(opportunity)}</b>.<br><br>` +
      'نأمل تعبئة ملف المتطلبات المرفق واستكمال البيانات المطلوبة، وسيتم استكمال إجراءات الفرصة بعد استلام المتطلبات.'),
    attachments: [file.getBlob()],
    name: 'إدارة العمل التطوعي - جامعة تبوك'
  });

  sheet.getRange(rowNumber, CONFIG.COL.STATUS).setValue('منتهية');
  sheet.getRange(rowNumber, CONFIG.COL.COMPLETED_AT).setValue(new Date());
  SpreadsheetApp.flush();

  return {
    message: automatic
      ? 'تم إرسال ملف المتطلبات تلقائياً وتسجيل الفرصة كمنتهية.'
      : 'تم إنهاء الفرصة وإرسال ملف المتطلبات.'
  };
}

/**
 * هذه الدالة يشغلها Trigger كل ساعة.
 * ترسل ملف المتطلبات بعد انقضاء تاريخ انتهاء الفرصة.
 * مثال: إذا تاريخ الانتهاء 2026-08-15، يتم الإرسال بعد دخول يوم 2026-08-16.
 */
function checkEndedOpportunities() {
  // قبل إضافة ملف Word لا نفعل شيئاً حتى لا تنتج أخطاء متكررة.
  if (!CONFIG.REQUIREMENTS_FILE_ID) return;

  const sheet = getSheet_();
  ensureSystemColumns_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
  const tz = Session.getScriptTimeZone() || 'Asia/Riyadh';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  values.forEach(function(row, i) {
    const rowNumber = i + 2;
    const status = String(row[4] || '').trim();
    const endDate = String(row[7] || '').trim();

    if (status === 'مقبول' && /^\d{4}-\d{2}-\d{2}$/.test(endDate) && today > endDate) {
      try {
        complete_(rowNumber, true);
      } catch (err) {
        console.error('تعذر إنهاء الصف ' + rowNumber + ': ' + err.message);
      }
    }
  });
}

/**
 * شغّل هذه الدالة مرة واحدة يدوياً بعد نشر السكربت.
 * تنشئ Trigger يفحص الفرص كل ساعة.
 */
function setupAutomationTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkEndedOpportunities') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('checkEndedOpportunities')
    .timeBased()
    .everyHours(1)
    .create();

  return 'تم إنشاء التحقق التلقائي كل ساعة.';
}

/* ========================= رسائل البريد ========================= */

function sendAcceptedEmail_(to, coordinator, opportunity, registrationUrl, endDate) {
  const safeUrl = escapeHtml_(registrationUrl);
  const linkButton = `<div style="text-align:center;margin:26px 0">
    <a href="${safeUrl}" target="_blank" style="display:inline-block;background:#006c43;color:#fff;text-decoration:none;padding:13px 26px;border-radius:10px;font-weight:bold">التسجيل في الفرصة</a>
  </div>`;

  MailApp.sendEmail({
    to: to,
    subject: 'تم قبول الفرصة التطوعية — جامعة تبوك',
    body:
      'السلام عليكم ورحمة الله وبركاته\n\n' +
      'تم قبول الفرصة التطوعية: ' + opportunity + '\n\n' +
      'رابط التسجيل: ' + registrationUrl + '\n' +
      'تاريخ انتهاء الفرصة: ' + endDate + '\n\n' +
      'إدارة العمل التطوعي - جامعة تبوك',
    htmlBody: emailTemplate_(coordinator,
      `يسرنا إشعاركم بأنه تم <b>قبول الفرصة التطوعية</b>${opportunity ? `: <b>${escapeHtml_(opportunity)}</b>` : ''}.<br><br>` +
      'يمكنكم البدء باستكمال إجراءات تنفيذ الفرصة من خلال رابط التسجيل أدناه.' +
      linkButton +
      `<div style="background:#f4f8f6;border-radius:10px;padding:12px 16px">تاريخ انتهاء الفرصة المسجل: <b>${escapeHtml_(endDate)}</b></div>`),
    name: 'إدارة العمل التطوعي - جامعة تبوك'
  });
}

function sendRejectedEmail_(to, coordinator, opportunity) {
  MailApp.sendEmail({
    to: to,
    subject: 'إشعار بشأن الفرصة التطوعية — جامعة تبوك',
    body:
      'السلام عليكم ورحمة الله وبركاته\n\n' +
      'نشكر لكم تقديم الفرصة التطوعية: ' + opportunity + '\n\n' +
      'نود إشعاركم بأنه لم تتم الموافقة على الفرصة في الوقت الحالي.\n\n' +
      'إدارة العمل التطوعي - جامعة تبوك',
    htmlBody: emailTemplate_(coordinator,
      `نشكر لكم تقديم الفرصة التطوعية ${opportunity ? `<b>${escapeHtml_(opportunity)}</b>` : ''}.<br><br>` +
      'نود إشعاركم بأنه <b>لم تتم الموافقة على الفرصة في الوقت الحالي</b>.<br><br>' +
      'نقدر لكم اهتمامكم ومبادرتكم في مجال العمل التطوعي.'),
    name: 'إدارة العمل التطوعي - جامعة تبوك'
  });
}

function emailTemplate_(name, body) {
  return `<div dir="rtl" style="font-family:Tahoma,Arial;line-height:2;color:#17211c;max-width:650px;margin:auto;border:1px solid #e5ebe7;border-radius:16px;overflow:hidden">
    <div style="background:#006c43;color:#fff;padding:18px 24px;font-size:20px;font-weight:bold">جامعة تبوك | إدارة العمل التطوعي</div>
    <div style="padding:24px">
      <p>السلام عليكم ورحمة الله وبركاته،</p>
      ${name ? `<p>الأخ/الأخت <b>${escapeHtml_(name)}</b>،</p>` : ''}
      <p>${body}</p>
      <p>مع خالص التحية،<br><b>إدارة العمل التطوعي — جامعة تبوك</b></p>
    </div>
  </div>`;
}

/* إشعار المدير عند وصول فرصة جديدة — يحتاج Trigger منفصل إذا أردته. */
function onFormSubmitNotifyManager(e) {
  if (!CONFIG.MANAGER_EMAIL) return;
  const named = e && e.namedValues ? e.namedValues : {};
  const details = Object.keys(named)
    .map(function(k) { return `<b>${escapeHtml_(k)}</b>: ${escapeHtml_((named[k] || []).join(', '))}`; })
    .join('<br>');

  MailApp.sendEmail({
    to: CONFIG.MANAGER_EMAIL,
    subject: 'فرصة تطوعية جديدة — جامعة تبوك',
    body: 'وصلت فرصة تطوعية جديدة عبر النموذج.',
    htmlBody: emailTemplate_('', `وصلت فرصة تطوعية جديدة عبر النموذج.<br><br>${details}`),
    name: 'نظام العمل التطوعي - جامعة تبوك'
  });
}

function debugSheet_() {
  const sheet = getSheet_();
  ensureSystemColumns_(sheet);
  return {
    sheetName: sheet.getName(),
    sheetId: sheet.getSheetId(),
    lastRow: sheet.getLastRow(),
    headers: sheet.getRange(1, 1, 1, 8).getDisplayValues()[0]
  };
}

function escapeHtml_(value) {
  return String(value || '').replace(/[&<>"']/g, function(char) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'
    }[char];
  });
}
