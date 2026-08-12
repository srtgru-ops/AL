نظام إدارة العمل التطوعي — جامعة تبوك
=====================================

الملفات:
1) index.html                 واجهة HTML/CSS/JavaScript
2) Code.gs                    Backend لـ Google Apps Script
3) university-tabuk-logo.png شعار الجامعة المستخدم في الواجهة

تم تجهيز الربط على ملف Google Sheets التالي:
Spreadsheet ID: 1tJ0te9q7cpGVI5_drEvbmJINg9OVa_wxeCtTiOvkvsQ
Sheet GID: 884978884

خطوات التشغيل:
1. افتح Google Sheet.
2. من Extensions > Apps Script.
3. الصق محتوى Code.gs واحفظ.
4. Deploy > New deployment > Web app.
5. Execute as: Me.
6. Who has access: اختر النطاق المناسب لسياسة الجامعة (ويُفضّل تقييده بحسابات الجامعة إن أمكن).
7. انسخ رابط Web App.
8. افتح index.html وابحث عن:
   const API_URL = 'PUT_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
   واستبدل القيمة برابط Web App.
9. ارفع index.html + university-tabuk-logo.png على الاستضافة التي تستخدمها.

ما يعمل الآن:
- قراءة ردود Google Form من Google Sheet.
- عرض جميع حقول النموذج في نافذة التفاصيل حتى لو تغيّرت أسماء/عدد الحقول.
- حالات: قيد المراجعة / مقبول / مرفوض / منتهٍ.
- زر قبول: يحدّث Google Sheet ويرسل إيميل قبول تلقائي للمتقدم.
- زر رفض: يحدّث الحالة.
- زر إنهاء الفرصة: مبرمج لإرسال ملف Word كـ attachment ثم تغيير الحالة إلى منتهٍ.
- بحث وفلترة وإحصاءات.

قبل تشغيل الإيميلات:
- البريد الذي تم تزويدي به هو: aa6vdd@gmil.com
- النطاق "gmil.com" يبدو مختلفاً عن "gmail.com"؛ تأكد منه قبل تفعيل إشعار المدير.

عند إرسال ملف Word لاحقاً:
- ارفعه إلى Google Drive.
- انسخ File ID.
- ضعه في Code.gs داخل CONFIG.REQUIREMENTS_FILE_ID.

ملاحظة مهمة:
- Apps Script يرسل البريد من حساب Google الذي يشغّل السكربت، وليس من بريد آخر بشكل تلقائي.
- إذا أردت أن تكون المنصة داخلية فقط، لا تجعل واجهة الإدارة عامة بدون حماية دخول.
