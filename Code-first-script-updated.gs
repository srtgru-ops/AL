/**
 * نظام إدارة الفرص التطوعية | جامعة تبوك
 *
 * A  = طابع زمني
 * B  = الاسم
 * D  = مسمى الفرصة
 * BS = الإيميل
 */

const CONFIG = {

  SPREADSHEET_ID:
    '1giQB70mJ-Vh69fdClCYAZttVFCiY1zb-jPjy1GCit3Y',

  SHEET_GID:
    531971439,

  MANAGER_EMAIL:
    'ao.alsharif@ut.edu.sa',

  // ملف Word الجديد الخاص بمتطلبات إنهاء الفرصة
  REQUIREMENTS_FILE_ID:
    '1KUTdig_i4oaF8sQmjZ73-H9-UVgCxrnP',

  // فورم رفع ملف الإنهاء بعد التعبئة
  COMPLETION_UPLOAD_FORM_URL:
    'https://docs.google.com/forms/d/e/1FAIpQLSc4yw8qitK87juLn9alDY29r4KjSQh209NrnQeGb8nLUrEknQ/viewform?usp=dialog',

  MAIN_COLUMNS: {
    TIMESTAMP: 1,     // A
    NAME: 2,          // B
    OPPORTUNITY: 4,   // D
    EMAIL: 71         // BS
  },

  SYSTEM_HEADERS: {
    STATUS: 'حالة النظام',
    REGISTRATION_LINK: 'رابط التسجيل',
    REJECTION_REASON: 'سبب الرفض',
    PARTICIPANTS: 'أسماء المشاركين'
  }

};


/* =========================================================
   WEB APP
========================================================= */

function doGet(e) {

  try {

    const p =
      (e && e.parameter) || {};

    const action =
      String(p.action || 'ping');

    let result;


    switch (action) {

      case 'ping':

        result = {
          message: 'Volunteer API is running'
        };

        break;


      case 'list':

        result =
          listRequests_();

        break;


      case 'setStatus':

        result =
          setStatus_(

            Number(p.rowNumber),

            String(
              p.status || ''
            ),

            String(
              p.registrationLink || ''
            ),

            String(
              p.rejectionReason || p.reason || ''
            )

          );

        break;


      case 'sendRequirements':

        result =
          sendRequirements_(

            Number(
              p.rowNumber
            ),

            String(
              p.participants || ''
            )

          );

        break;


      case 'finalize':

        result =
          finalizeOpportunity_(

            Number(
              p.rowNumber
            )

          );

        break;


      case 'debug':

        result =
          debugSheet_();

        break;


      default:

        throw new Error(
          'إجراء غير معروف.'
        );

    }


    return json_({

      ok: true,

      ...result

    });


  } catch (err) {


    return json_({

      ok: false,

      error:
        err && err.message
          ? err.message
          : String(err)

    });

  }

}


/* =========================================================
   POST
========================================================= */

function doPost(e) {

  try {

    const body =
      JSON.parse(

        (
          e &&
          e.postData &&
          e.postData.contents
        ) || '{}'

      );


    const fakeEvent = {
      parameter: {}
    };


    Object
      .keys(body)
      .forEach(function(key) {

        fakeEvent.parameter[key] =
          String(
            body[key] ?? ''
          );

      });


    return doGet(
      fakeEvent
    );


  } catch (err) {


    return json_({

      ok: false,

      error:
        err && err.message
          ? err.message
          : String(err)

    });

  }

}


/* =========================================================
   JSON
========================================================= */

function json_(obj) {

  return ContentService
    .createTextOutput(
      JSON.stringify(obj)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}


/* =========================================================
   GOOGLE SHEETS
========================================================= */

function getSheet_() {

  const ss =
    SpreadsheetApp
      .openById(
        CONFIG.SPREADSHEET_ID
      );


  const sheet =
    ss.getSheets().find(
      function(s) {

        return Number(
          s.getSheetId()
        ) === Number(
          CONFIG.SHEET_GID
        );

      }
    );


  if (!sheet) {

    throw new Error(
      'تعذر العثور على ورقة Google Sheets المطلوبة.'
    );

  }


  return sheet;

}


function findHeaderColumn_(
  headers,
  wanted
) {

  const target =
    String(
      wanted || ''
    ).trim();


  for (
    let i = 0;
    i < headers.length;
    i++
  ) {

    if (

      String(
        headers[i] || ''
      ).trim()
      === target

    ) {

      return i + 1;

    }

  }


  return -1;

}


/* =========================================================
   إنشاء أعمدة النظام
========================================================= */

function ensureSystemColumns_(
  sheet
) {

  let lastCol =
    Math.max(
      sheet.getLastColumn(),
      1
    );

  let headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastCol
      )
      .getDisplayValues()[0];

  function ensureColumn_(headerName) {
    let col = findHeaderColumn_(headers, headerName);

    if (col === -1) {
      lastCol++;
      sheet
        .getRange(1, lastCol)
        .setValue(headerName);

      headers.push(headerName);
      col = lastCol;
    }

    return col;
  }

  const statusCol =
    ensureColumn_(
      CONFIG.SYSTEM_HEADERS.STATUS
    );

  const registrationLinkCol =
    ensureColumn_(
      CONFIG.SYSTEM_HEADERS.REGISTRATION_LINK
    );

  const rejectionReasonCol =
    ensureColumn_(
      CONFIG.SYSTEM_HEADERS.REJECTION_REASON
    );

  const participantsCol =
    ensureColumn_(
      CONFIG.SYSTEM_HEADERS.PARTICIPANTS
    );

  return {
    headers: headers,
    statusCol: statusCol,
    registrationLinkCol: registrationLinkCol,
    rejectionReasonCol: rejectionReasonCol,
    participantsCol: participantsCol
  };

}


/* =========================================================
   قراءة الطلبات
========================================================= */

function listRequests_() {

  const sheet =
    getSheet_();


  const system =
    ensureSystemColumns_(
      sheet
    );


  const headers =
    system.headers;


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return {
      count: 0,
      rows: []
    };

  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getDisplayValues();


  const rows = [];


  values.forEach(
    function(row, index) {


      const hasData =
        row.some(
          function(cell, i) {

            if (
              i + 1 === system.statusCol
              ||
              i + 1 === system.registrationLinkCol
              ||
              i + 1 === system.rejectionReasonCol
              ||
              i + 1 === system.participantsCol
            ) {
              return false;
            }


            return String(
              cell || ''
            ).trim() !== '';

          }
        );


      if (!hasData) {
        return;
      }


      const fields = {};


      headers.forEach(
        function(header, i) {

          const title =
            String(
              header || ''
            ).trim();


          if (!title) {
            return;
          }


          if (
            title === CONFIG.SYSTEM_HEADERS.STATUS
            ||
            title === CONFIG.SYSTEM_HEADERS.REGISTRATION_LINK
            ||
            title === CONFIG.SYSTEM_HEADERS.REJECTION_REASON
            ||
            title === CONFIG.SYSTEM_HEADERS.PARTICIPANTS
          ) {
            return;
          }


          let key =
            title;


          if (
            Object.prototype.hasOwnProperty.call(
              fields,
              key
            )
          ) {

            key =
              title +
              ' - العمود ' +
              columnToLetter_(i + 1);

          }


          fields[key] =
            row[i] || '';

        }
      );


      rows.push({

        rowNumber:
          index + 2,

        coordinatorName:
          row[
            CONFIG.MAIN_COLUMNS.NAME - 1
          ] || '',

        opportunityName:
          row[
            CONFIG.MAIN_COLUMNS.OPPORTUNITY - 1
          ] || '',

        email:
          row[
            CONFIG.MAIN_COLUMNS.EMAIL - 1
          ] || '',

        timestamp:
          row[
            CONFIG.MAIN_COLUMNS.TIMESTAMP - 1
          ] || '',

        status:
          row[
            system.statusCol - 1
          ] || 'قيد المراجعة',

        registrationLink:
          row[
            system.registrationLinkCol - 1
          ] || '',

        rejectionReason:
          row[
            system.rejectionReasonCol - 1
          ] || '',

        participants:
          row[
            system.participantsCol - 1
          ] || '',

        fields:
          fields

      });

    }
  );


  rows.reverse();


  return {
    count: rows.length,
    rows: rows
  };

}


/* =========================================================
   قبول / رفض
========================================================= */

function setStatus_(
  rowNumber,
  status,
  registrationLink,
  rejectionReason
) {

  if (!rowNumber || rowNumber < 2) {
    throw new Error('رقم الصف غير صالح.');
  }

  const allowed = [
    'مقبول',
    'مرفوض',
    'قيد المراجعة'
  ];

  if (!allowed.includes(status)) {
    throw new Error('الحالة غير معتمدة.');
  }

  const sheet =
    getSheet_();

  const system =
    ensureSystemColumns_(
      sheet
    );

  const name =
    getCellValue_(
      sheet,
      rowNumber,
      CONFIG.MAIN_COLUMNS.NAME
    );

  const opportunity =
    getCellValue_(
      sheet,
      rowNumber,
      CONFIG.MAIN_COLUMNS.OPPORTUNITY
    );

  const email =
    getCellValue_(
      sheet,
      rowNumber,
      CONFIG.MAIN_COLUMNS.EMAIL
    ).trim();

  if (status === 'مقبول') {

    if (!email) {
      throw new Error('لا يوجد بريد إلكتروني في هذا الطلب.');
    }

    registrationLink =
      String(
        registrationLink || ''
      ).trim();

    if (!registrationLink) {
      throw new Error('أدخل رابط التسجيل أولاً قبل قبول الفرصة.');
    }

    if (!/^https?:\/\//i.test(registrationLink)) {
      throw new Error('رابط التسجيل غير صحيح.');
    }

    sendAcceptedEmail_(
      email,
      name,
      opportunity,
      registrationLink
    );

    sheet
      .getRange(
        rowNumber,
        system.registrationLinkCol
      )
      .setValue(
        registrationLink
      );

    sheet
      .getRange(
        rowNumber,
        system.rejectionReasonCol
      )
      .clearContent();
  }

  if (status === 'مرفوض') {

    if (!email) {
      throw new Error('لا يوجد بريد إلكتروني في هذا الطلب.');
    }

    rejectionReason =
      String(
        rejectionReason || ''
      ).trim();

    if (!rejectionReason) {
      throw new Error('يجب كتابة سبب الرفض قبل رفض الطلب.');
    }

    sendRejectedEmail_(
      email,
      name,
      opportunity,
      rejectionReason
    );

    sheet
      .getRange(
        rowNumber,
        system.rejectionReasonCol
      )
      .setValue(
        rejectionReason
      );
  }

  if (status === 'قيد المراجعة') {
    sheet
      .getRange(
        rowNumber,
        system.rejectionReasonCol
      )
      .clearContent();
  }

  sheet
    .getRange(
      rowNumber,
      system.statusCol
    )
    .setValue(
      status
    );

  SpreadsheetApp.flush();

  if (status === 'مقبول') {
    return {
      message:
        'تم قبول الفرصة وإرسال رابط التسجيل.'
    };
  }

  if (status === 'مرفوض') {
    return {
      message:
        'تم رفض الطلب وإرسال سبب الرفض إلى مقدم الطلب.'
    };
  }

  return {
    message:
      'تم إعادة الطلب إلى قيد المراجعة.'
  };

}


/* =========================================================
   إرسال ملف الإنهاء + رابط فورم الرفع
========================================================= */

function sendRequirements_(
  rowNumber,
  participants
) {

  if (
    !rowNumber ||
    rowNumber < 2
  ) {

    throw new Error(
      'رقم الصف غير صالح.'
    );

  }


  const sheet =
    getSheet_();


  const system =
    ensureSystemColumns_(
      sheet
    );


  const currentStatus =
    getCellValue_(
      sheet,
      rowNumber,
      system.statusCol
    ).trim();


  if (
    currentStatus !== 'مقبول'
  ) {

    throw new Error(
      'يمكن إرسال ملف الإنهاء للفرص المقبولة فقط.'
    );

  }


  const name =
    getCellValue_(
      sheet,
      rowNumber,
      CONFIG.MAIN_COLUMNS.NAME
    );


  const opportunity =
    getCellValue_(
      sheet,
      rowNumber,
      CONFIG.MAIN_COLUMNS.OPPORTUNITY
    );


  const email =
    getCellValue_(
      sheet,
      rowNumber,
      CONFIG.MAIN_COLUMNS.EMAIL
    ).trim();


  participants =
    String(
      participants || ''
    )
    .split(/\r?\n/)
    .map(function(name) {
      return name.trim();
    })
    .filter(function(name) {
      return !!name;
    })
    .join('\n');


  if (!participants) {
    throw new Error(
      'يجب إدخال أسماء المشاركين قبل إرسال ملف الإنهاء.'
    );
  }


  if (!email) {
    throw new Error(
      'لا يوجد بريد إلكتروني لهذا الطلب.'
    );
  }


  const attachment =
    getRequirementsAttachment_();


  const uploadFormUrl =
    CONFIG.COMPLETION_UPLOAD_FORM_URL;


  MailApp.sendEmail({

    to:
      email,

    subject:
      'متطلبات إنهاء الفرصة التطوعية — جامعة تبوك',

    body:

      'السلام عليكم ورحمة الله وبركاته\n\n'

      +

      'نفيدكم بأنه تم الوصول إلى مرحلة إنهاء الفرصة التطوعية: '

      +

      opportunity

      +

      '\n\n'

      +

      'أسماء المشاركين في الفرصة:\n'

      +

      participants

      +

      '\n\n'

      +

      'نأمل تعبئة ملف متطلبات إنهاء الفرصة المرفق.'

      +

      '\n\n'

      +

      'بعد تعبئة الملف يرجى رفعه من خلال الرابط التالي:\n'

      +

      uploadFormUrl

      +

      '\n\n'

      +

      'لن يتم اعتماد إنهاء الفرصة إلا بعد رفع الملف واستكمال المتطلبات.'

      +

      '\n\n'

      +

      'مع خالص التحية\n'

      +

      'إدارة العمل التطوعي - جامعة تبوك',


    htmlBody:

      emailTemplate_(

        name,

        `نفيدكم بأنه تم الوصول إلى مرحلة إنهاء الفرصة التطوعية

        ${
          opportunity
            ? '<b>'
              +
              escapeHtml_(
                opportunity
              )
              +
              '</b>'
            : ''
        }.

        <br><br>

        <div
          style="
            background:#f1f8f4;
            border:1px solid #cfe3d7;
            border-radius:10px;
            padding:14px;
          "
        >
          <b>أسماء المشاركين في الفرصة:</b>
          <br><br>
          ${escapeHtml_(participants).replace(/\n/g, '<br>')}
        </div>

        <br><br>

        نأمل تعبئة
        <b>
          ملف متطلبات إنهاء الفرصة
        </b>
        المرفق في هذه الرسالة،
        واستكمال جميع البيانات المطلوبة.

        <br><br>

        بعد الانتهاء من تعبئة الملف،
        يرجى رفع الملف من خلال الزر التالي:

        <br><br>

        <div style="text-align:center">

          <a
            href="${escapeHtml_(uploadFormUrl)}"
            target="_blank"
            style="
              display:inline-block;
              background:#006c43;
              color:#ffffff;
              text-decoration:none;
              padding:14px 30px;
              border-radius:9px;
              font-size:16px;
              font-weight:bold;
            "
          >
            رفع ملف الإنهاء بعد التعبئة
          </a>

        </div>

        <br><br>

        <div
          style="
            background:#fff6dc;
            border:1px solid #efd891;
            border-radius:10px;
            padding:14px;
          "
        >
          لن يتم اعتماد انتهاء الفرصة
          إلا بعد تعبئة الملف
          ورفعه من خلال النموذج.
        </div>`

      ),


    attachments: [
      attachment
    ],


    name:
      'إدارة العمل التطوعي - جامعة تبوك'

  });


  sheet
    .getRange(
      rowNumber,
      system.participantsCol
    )
    .setValue(
      participants
    );


  sheet
    .getRange(
      rowNumber,
      system.statusCol
    )
    .setValue(
      'بانتظار متطلبات الإنهاء'
    );


  SpreadsheetApp.flush();


  return {
    message:
      'تم إرسال ملف الإنهاء ورابط رفع الملف إلى مقدم الفرصة.'
  };

}


/* =========================================================
   اعتماد الإنهاء
========================================================= */

function finalizeOpportunity_(
  rowNumber
) {

  if (
    !rowNumber ||
    rowNumber < 2
  ) {
    throw new Error(
      'رقم الصف غير صالح.'
    );
  }


  const sheet =
    getSheet_();


  const system =
    ensureSystemColumns_(
      sheet
    );


  sheet
    .getRange(
      rowNumber,
      system.statusCol
    )
    .setValue(
      'منتهية'
    );


  SpreadsheetApp.flush();


  return {
    message:
      'تم اعتماد انتهاء الفرصة بنجاح.'
  };

}


/* =========================================================
   جلب ملف Word
========================================================= */

function getRequirementsAttachment_() {

  if (
    !CONFIG.REQUIREMENTS_FILE_ID
  ) {

    throw new Error(
      'لم يتم تحديد ملف متطلبات الإنهاء.'
    );

  }


  const file =
    DriveApp.getFileById(
      CONFIG.REQUIREMENTS_FILE_ID
    );


  const mimeType =
    file.getMimeType();


  if (
    mimeType ===
    MimeType.GOOGLE_DOCS
  ) {

    const url =

      'https://docs.google.com/document/d/'

      +

      CONFIG.REQUIREMENTS_FILE_ID

      +

      '/export?format=docx';


    const response =
      UrlFetchApp.fetch(
        url,
        {
          headers: {
            Authorization:
              'Bearer ' +
              ScriptApp.getOAuthToken()
          },
          muteHttpExceptions: true
        }
      );


    if (
      response.getResponseCode() !== 200
    ) {

      throw new Error(
        'تعذر تصدير ملف المتطلبات بصيغة Word.'
      );

    }


    return response
      .getBlob()
      .setName(
        'متطلبات إنهاء الفرصة التطوعية.docx'
      );

  }


  return file
    .getBlob()
    .setName(
      'متطلبات إنهاء الفرصة التطوعية.docx'
    );

}


/* =========================================================
   إيميل القبول
========================================================= */

function sendAcceptedEmail_(
  to,
  name,
  opportunity,
  registrationLink
) {

  const safeLink =
    escapeHtml_(
      registrationLink
    );


  MailApp.sendEmail({

    to:
      to,

    subject:
      'تم قبول الفرصة التطوعية — جامعة تبوك',

    body:
      'السلام عليكم ورحمة الله وبركاته\n\n' +
      'تم قبول الفرصة التطوعية: ' +
      opportunity +
      '\n\n' +
      'رابط التسجيل:\n' +
      registrationLink +
      '\n\n' +
      'إدارة العمل التطوعي - جامعة تبوك',

    htmlBody:

      emailTemplate_(

        name,

        `يسرنا إشعاركم بأنه تم
        <b>قبول الفرصة التطوعية</b>

        ${
          opportunity
            ? ': <b>'
              +
              escapeHtml_(
                opportunity
              )
              +
              '</b>'
            : ''
        }.

        <br><br>

        يمكنكم التسجيل من خلال الزر التالي:

        <br><br>

        <div style="text-align:center">

          <a
            href="${safeLink}"
            target="_blank"
            style="
              display:inline-block;
              background:#006c43;
              color:#ffffff;
              text-decoration:none;
              padding:13px 28px;
              border-radius:9px;
              font-weight:bold;
            "
          >
            التسجيل في الفرصة
          </a>

        </div>`

      ),

    name:
      'إدارة العمل التطوعي - جامعة تبوك'

  });

}


/* =========================================================
   إيميل الرفض
========================================================= */

function sendRejectedEmail_(
  to,
  name,
  opportunity,
  rejectionReason
) {

  const safeReason =
    escapeHtml_(
      rejectionReason
    )
    .replace(
      /\n/g,
      '<br>'
    );

  MailApp.sendEmail({

    to:
      to,

    subject:
      'إشعار بشأن الفرصة التطوعية — جامعة تبوك',

    body:
      'السلام عليكم ورحمة الله وبركاته،\n\n'
      +
      'الأخ/الأخت '
      +
      (name || '')
      +
      ' المحترم/ة،\n\n'
      +
      'نشكر لكم تقديم الفرصة التطوعية'
      +
      (
        opportunity
          ? ': ' + opportunity
          : ''
      )
      +
      '.\n\n'
      +
      'نود إشعاركم بأنه لم تتم الموافقة على الفرصة في الوقت الحالي.\n\n'
      +
      'سبب الرفض:\n'
      +
      rejectionReason
      +
      '\n\n'
      +
      'مع خالص التحية والتقدير،\n'
      +
      'إدارة العمل التطوعي\n'
      +
      'جامعة تبوك',

    htmlBody:
      emailTemplate_(
        name,
        `نشكر لكم تقديم الفرصة التطوعية
        ${
          opportunity
            ? '<b>' + escapeHtml_(opportunity) + '</b>'
            : ''
        }.

        <br><br>

        نود إشعاركم بأنه
        <b>لم تتم الموافقة على الفرصة في الوقت الحالي.</b>

        <br><br>

        <div
          style="
            background:#fff3f3;
            border:1px solid #efcccc;
            border-radius:12px;
            padding:16px;
          "
        >
          <b style="color:#b02d2d">سبب الرفض:</b>
          <br><br>
          ${safeReason}
        </div>`
      ),

    name:
      'إدارة العمل التطوعي - جامعة تبوك'

  });

}


/* =========================================================
   قالب البريد
========================================================= */

function emailTemplate_(
  name,
  body
) {

  return `

  <div
    dir="rtl"
    style="
      font-family:Tahoma,Arial,sans-serif;
      line-height:2;
      color:#17211c;
      max-width:650px;
      margin:auto;
      border:1px solid #e5ebe7;
      border-radius:16px;
      overflow:hidden;
    "
  >

    <div
      style="
        background:#006c43;
        color:#ffffff;
        padding:18px 24px;
        font-size:20px;
        font-weight:bold;
      "
    >
      جامعة تبوك |
      إدارة العمل التطوعي
    </div>


    <div
      style="
        padding:24px;
      "
    >

      <p>
        السلام عليكم ورحمة الله وبركاته،
      </p>


      ${
        name
          ? `
          <p>
            الأخ/الأخت
            <b>
              ${escapeHtml_(name)}
            </b>،
          </p>
          `
          : ''
      }


      <div>
        ${body}
      </div>


      <p>
        مع خالص التحية،
        <br>

        <b>
          إدارة العمل التطوعي
          — جامعة تبوك
        </b>
      </p>

    </div>

  </div>

  `;

}


/* =========================================================
   قراءة خلية
========================================================= */

function getCellValue_(
  sheet,
  row,
  column
) {

  return String(

    sheet
      .getRange(
        row,
        column
      )
      .getDisplayValue()

    || ''

  );

}


/* =========================================================
   DEBUG
========================================================= */

function debugSheet_() {

  const sheet =
    getSheet_();


  const system =
    ensureSystemColumns_(
      sheet
    );


  return {

    sheetName:
      sheet.getName(),

    sheetId:
      sheet.getSheetId(),

    lastRow:
      sheet.getLastRow(),

    lastColumn:
      sheet.getLastColumn(),

    emailColumn:
      'BS',

    statusColumn:
      columnToLetter_(
        system.statusCol
      ),

    registrationLinkColumn:
      columnToLetter_(
        system.registrationLinkCol
      ),

    rejectionReasonColumn:
      columnToLetter_(
        system.rejectionReasonCol
      ),

    participantsColumn:
      columnToLetter_(
        system.participantsCol
      ),

    requirementsFileId:
      CONFIG.REQUIREMENTS_FILE_ID

  };

}


/* =========================================================
   رقم العمود إلى حرف
========================================================= */

function columnToLetter_(
  column
) {

  let temp = '';
  let letter = '';


  while (
    column > 0
  ) {

    temp =
      (column - 1) % 26;


    letter =
      String.fromCharCode(
        temp + 65
      ) + letter;


    column =
      (
        column -
        temp -
        1
      ) / 26;

  }


  return letter;

}


/* =========================================================
   حماية HTML
========================================================= */

function escapeHtml_(
  value
) {

  return String(
    value || ''
  )
    .replace(

      /[&<>"']/g,

      function(char) {

        return {
          '&': '&amp;',
          '<': '&lt;',
          '>': '&gt;',
          '"': '&quot;',
          "'": '&#039;'
        }[char];

      }

    );

}


/* =========================================================
   اختبار صلاحية Drive
========================================================= */

function testDriveAccess() {

  const file =
    DriveApp.getFileById(
      CONFIG.REQUIREMENTS_FILE_ID
    );


  Logger.log(
    'تم الوصول للملف بنجاح: ' +
    file.getName()
  );


  return file.getName();

}