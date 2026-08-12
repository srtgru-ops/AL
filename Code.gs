/**
 * Backend — نظام إدارة العمل التطوعي | جامعة تبوك
 * اربط هذا السكربت بملف Google Sheets ثم انشره كـ Web App.
 */
const CONFIG = {
  SPREADSHEET_ID: '1tJ0te9q7cpGVI5_drEvbmJINg9OVa_wxeCtTiOvkvsQ',
  SHEET_GID: 884978884,
  MANAGER_EMAIL: 'aa6vdd@gmil.com', // تم إدخاله كما أرسله المستخدم — راجع نطاق البريد قبل التشغيل النهائي.
  REQUIREMENTS_FILE_ID: '', // يوضع لاحقاً بعد رفع ملف Word على Google Drive.
  STATUS_HEADER: 'الحالة',
  COMPLETED_AT_HEADER: 'تاريخ إنهاء الفرصة'
};

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ok:true,message:'Volunteer API is running'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    let result;
    switch (body.action) {
      case 'list': result = listRequests_(); break;
      case 'setStatus': result = setStatus_(Number(body.rowNumber), String(body.status || '')); break;
      case 'complete': result = complete_(Number(body.rowNumber)); break;
      default: throw new Error('إجراء غير معروف.');
    }
    return json_({ok:true,...result});
  } catch (err) {
    return json_({ok:false,error:err.message || String(err)});
  }
}

function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(){
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheets().find(s => s.getSheetId() === CONFIG.SHEET_GID);
  if (!sheet) throw new Error('تعذر العثور على ورقة Google Sheets المطلوبة.');
  return sheet;
}

function ensureSystemColumns_(sheet){
  let lastCol = Math.max(sheet.getLastColumn(),1);
  let headers = sheet.getRange(1,1,1,lastCol).getDisplayValues()[0];
  [CONFIG.STATUS_HEADER, CONFIG.COMPLETED_AT_HEADER].forEach(name => {
    if (!headers.includes(name)) {
      lastCol++;
      sheet.getRange(1,lastCol).setValue(name);
      headers.push(name);
    }
  });
  return headers;
}

function listRequests_(){
  const sheet = getSheet_();
  const headers = ensureSystemColumns_(sheet);
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return {rows:[]};
  const values = sheet.getRange(2,1,lastRow-1,headers.length).getDisplayValues();

  const emailHeader = findHeader_(headers, ['البريد الإلكتروني','البريد الالكتروني','البريد','Email','E-mail','عنوان البريد الإلكتروني']);
  const applicantHeader = findHeader_(headers, ['الاسم','اسم الطالب','اسم المتقدم','الاسم الكامل','اسم مقدم الفرصة','مقدم الفرصة']);
  const opportunityHeader = findHeader_(headers, ['اسم الفرصة','اسم الفرصة التطوعية','الفرصة التطوعية','عنوان الفرصة','اسم المبادرة']);
  const timestampHeader = findHeader_(headers, ['الطابع الزمني','Timestamp','وقت الإرسال','تاريخ التقديم']);
  const statusIndex = headers.indexOf(CONFIG.STATUS_HEADER);

  const rows = values.map((row,i) => {
    const fields = {};
    headers.forEach((h,idx)=>fields[h || ('حقل '+(idx+1))] = row[idx]);
    return {
      rowNumber:i+2,
      applicantName:valueByHeader_(headers,row,applicantHeader),
      opportunityName:valueByHeader_(headers,row,opportunityHeader),
      email:valueByHeader_(headers,row,emailHeader),
      timestamp:valueByHeader_(headers,row,timestampHeader),
      status:row[statusIndex] || 'قيد المراجعة',
      fields
    };
  }).reverse();
  return {rows};
}

function setStatus_(rowNumber,status){
  if(!rowNumber || rowNumber < 2) throw new Error('رقم الصف غير صالح.');
  if(!['مقبول','مرفوض','قيد المراجعة'].includes(status)) throw new Error('الحالة غير معتمدة.');
  const sheet = getSheet_();
  const headers = ensureSystemColumns_(sheet);
  const statusCol = headers.indexOf(CONFIG.STATUS_HEADER)+1;
  sheet.getRange(rowNumber,statusCol).setValue(status);

  const row = sheet.getRange(rowNumber,1,1,headers.length).getDisplayValues()[0];
  const email = smartValue_(headers,row,['البريد الإلكتروني','البريد الالكتروني','البريد','Email','E-mail','عنوان البريد الإلكتروني']);
  const name = smartValue_(headers,row,['الاسم','اسم الطالب','اسم المتقدم','الاسم الكامل','اسم مقدم الفرصة','مقدم الفرصة']);
  const opportunity = smartValue_(headers,row,['اسم الفرصة','اسم الفرصة التطوعية','الفرصة التطوعية','عنوان الفرصة','اسم المبادرة']);

  if(status === 'مقبول'){
    if(!email) throw new Error('تم تغيير الحالة، لكن لم يتم العثور على بريد المتقدم لإرسال رسالة القبول.');
    sendAcceptedEmail_(email,name,opportunity);
    return {message:'تم قبول الطلب وإرسال رسالة القبول للمتقدم.'};
  }
  return {message:'تم تحديث حالة الطلب إلى «'+status+'».'};
}

function complete_(rowNumber){
  if(!rowNumber || rowNumber < 2) throw new Error('رقم الصف غير صالح.');
  const sheet = getSheet_();
  const headers = ensureSystemColumns_(sheet);
  const row = sheet.getRange(rowNumber,1,1,headers.length).getDisplayValues()[0];
  const email = smartValue_(headers,row,['البريد الإلكتروني','البريد الالكتروني','البريد','Email','E-mail','عنوان البريد الإلكتروني']);
  const name = smartValue_(headers,row,['الاسم','اسم الطالب','اسم المتقدم','الاسم الكامل','اسم مقدم الفرصة','مقدم الفرصة']);
  const opportunity = smartValue_(headers,row,['اسم الفرصة','اسم الفرصة التطوعية','الفرصة التطوعية','عنوان الفرصة','اسم المبادرة']);
  if(!email) throw new Error('لا يوجد بريد إلكتروني للمتقدم.');
  if(!CONFIG.REQUIREMENTS_FILE_ID) throw new Error('لم يتم إضافة ملف Word الخاص بمتطلبات إنهاء الفرصة بعد.');

  const file = DriveApp.getFileById(CONFIG.REQUIREMENTS_FILE_ID);
  MailApp.sendEmail({
    to: email,
    subject: 'متطلبات استكمال الفرصة التطوعية — جامعة تبوك',
    htmlBody: emailTemplate_(name,`تم تسجيل انتهاء الفرصة التطوعية <b>${escapeHtml_(opportunity || '')}</b>.<br><br>نأمل تعبئة الملف المرفق واستكمال المتطلبات المطلوبة، وسيتم استكمال إجراءات الفرصة بعد استلام المتطلبات.`),
    attachments:[file.getBlob()],
    name:'إدارة العمل التطوعي - جامعة تبوك'
  });

  sheet.getRange(rowNumber,headers.indexOf(CONFIG.STATUS_HEADER)+1).setValue('منتهٍ');
  sheet.getRange(rowNumber,headers.indexOf(CONFIG.COMPLETED_AT_HEADER)+1).setValue(new Date());
  return {message:'تم إنهاء الفرصة وإرسال ملف المتطلبات للمتقدم.'};
}

function sendAcceptedEmail_(to,name,opportunity){
  MailApp.sendEmail({
    to,
    subject:'تم قبول طلب الفرصة التطوعية — جامعة تبوك',
    htmlBody:emailTemplate_(name,`يسرنا إشعاركم بأنه تم <b>قبول طلبكم</b>${opportunity?` للفرصة التطوعية: <b>${escapeHtml_(opportunity)}</b>`:''}.<br><br>سيتم التواصل معكم بشأن تفاصيل التنفيذ حسب الإجراءات المعتمدة.`),
    name:'إدارة العمل التطوعي - جامعة تبوك'
  });
}

function emailTemplate_(name,body){
  return `<div dir="rtl" style="font-family:Tahoma,Arial;line-height:2;color:#17211c;max-width:650px;margin:auto;border:1px solid #e5ebe7;border-radius:16px;overflow:hidden">
    <div style="background:#006c43;color:#fff;padding:18px 24px;font-size:20px;font-weight:bold">جامعة تبوك | إدارة العمل التطوعي</div>
    <div style="padding:24px"><p>السلام عليكم ورحمة الله وبركاته،</p><p>${name?`الأخ/الأخت <b>${escapeHtml_(name)}</b>،`:''}</p><p>${body}</p><p>مع خالص التحية،<br><b>إدارة العمل التطوعي — جامعة تبوك</b></p></div>
  </div>`;
}

/** إشعار اختياري عند وصول رد جديد من Google Form.
 * بعد التأكد من البريد، أنشئ Trigger من Apps Script: From spreadsheet > On form submit.
 */
function onFormSubmitNotifyManager(e){
  if(!CONFIG.MANAGER_EMAIL) return;
  const named = e && e.namedValues ? e.namedValues : {};
  const body = Object.keys(named).map(k=>`<b>${escapeHtml_(k)}</b>: ${escapeHtml_((named[k]||[]).join(', '))}`).join('<br>');
  MailApp.sendEmail({to:CONFIG.MANAGER_EMAIL,subject:'طلب تطوعي جديد — جامعة تبوك',htmlBody:emailTemplate_('',`وصل طلب جديد عبر نموذج العمل التطوعي.<br><br>${body}`),name:'نظام العمل التطوعي - جامعة تبوك'});
}

function findHeader_(headers,candidates){
  const normalized = headers.map(h=>normalize_(h));
  for(const c of candidates){ const i=normalized.indexOf(normalize_(c)); if(i!==-1) return headers[i]; }
  return '';
}
function valueByHeader_(headers,row,header){ return header ? row[headers.indexOf(header)] || '' : ''; }
function smartValue_(headers,row,candidates){ const h=findHeader_(headers,candidates); return valueByHeader_(headers,row,h); }
function normalize_(s){ return String(s||'').trim().toLowerCase().replace(/[أإآ]/g,'ا').replace(/ة/g,'ه').replace(/ى/g,'ي').replace(/\s+/g,' '); }
function escapeHtml_(s){ return String(s||'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m])); }
