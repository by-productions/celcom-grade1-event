/****************************************************************
 *  סלקום · אירוע כיתות א' — חיבור טופס ל-Google Sheets + Drive
 *  ----------------------------------------------------------------
 *  הקוד הזה מקבל את הטופס מהאתר, שומר את התמונה ל-Drive,
 *  וכותב שורה חדשה ל-Google Sheet עם כל הפרטים + קישור לתמונה.
 *
 *  אין צורך לשנות כלום — ה-ID של הגיליון והתיקייה כבר מוגדרים.
 ****************************************************************/

// הגיליון שלך (Google Sheet)
const SHEET_ID  = '12QAZxngk4gaiMUVnHyTASRZW-5PhM6namaSnHZRJYyc';

// תיקיית האחסון שלך ב-Google Drive
const FOLDER_ID = '1dyitJA7VKGed4M0bw44_g3QwWEMS11N4';

// כותרות העמודות בגיליון
const HEADERS = ['תאריך ושעה','שם הילד/ה','שם משפחה','חוזקות','ברכה אישית','חלומות','שם ההורה','טלפון','אימייל','תמונה','מספר עובד'];

// פרטי האירוע (מופיעים במייל האישור)
const EVENT = {
  date:     'יום שני · 3 באוגוסט 2026',
  place:    'הספארי',
  schedule: [
    ['16:00','קבלת פנים'],
    ['18:30','ארוחת ערב'],
    ['19:00','מופע'],
    ['20:00','חלוקת מתנות וסיום משוער']
  ]
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // ----- 1) שמירת התמונה ל-Drive -----
    let photoUrl = '';
    if (data.photoData && data.photoData.indexOf(',') > -1) {
      try {
        const parts       = data.photoData.split(',');
        const contentType = (parts[0].match(/data:(.*?);/) || [])[1] || 'image/jpeg';
        const ext         = (contentType.split('/')[1] || 'jpg').replace('jpeg','jpg');
        const bytes       = Utilities.base64Decode(parts[1]);
        const fileName    = (data.photoName || 'registration') + '.' + ext;
        const blob        = Utilities.newBlob(bytes, contentType, fileName);

        const folder = DriveApp.getFolderById(FOLDER_ID);
        const file   = folder.createFile(blob);
        // ניסיון לשתף בקישור ציבורי; אם הדומיין חוסם שיתוף — ממשיכים בלי להיכשל
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (shareErr) {}
        photoUrl = file.getUrl();
      } catch (photoErr) {
        // גם אם שמירת התמונה נכשלה — לא חוסמים את הכתיבה לגיליון
        photoUrl = 'שגיאה בשמירת תמונה';
      }
    }

    // ----- 2) כתיבת שורה לגיליון -----
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }
    sheet.appendRow([
      new Date(),
      data.fname     || '',
      data.lname     || '',
      data.strengths || '',
      data.blessing  || '',
      data.dreams    || '',
      data.parent    || '',
      "'" + (data.phone || ''),   // הגרש שומר על אפס מוביל בטלפון
      data.email     || '',
      photoUrl,
      "'" + (data.employeeId || '')   // הגרש שומר על אפס מוביל במספר עובד
    ]);

    // ----- 3) מייל אישור הרשמה להורה (Resend) -----
    // עטוף ב-try/catch — כשל במייל לעולם לא מפיל את ההרשמה עצמה
    let emailSent = false;
    try {
      emailSent = sendConfirmationEmail(data);
    } catch (mailErr) {
      Logger.log('שגיאה בשליחת מייל אישור: ' + mailErr);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, photoUrl: photoUrl, emailSent: emailSent }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// בדיקה מהירה שהשרת חי (אם נכנסים לכתובת בדפדפן)
function doGet() {
  return ContentService.createTextOutput('✓ נקודת הקצה של אירוע סלקום פעילה');
}

/****************************************************************
 *  📧 מייל אישור הרשמה — נשלח דרך Resend
 *  ----------------------------------------------------------------
 *  ההגדרות נשמרות ב-Script Properties (לא בקוד!):
 *  Project Settings ⚙️ ← Script Properties ← Add script property
 *
 *    RESEND_API_KEY  = re_xxxxxxxx        (מ-resend.com/api-keys)
 *    MAIL_FROM       = סיפור חדש מתחיל <events@הדומיין-המאומת-שלך>
 *    MAIL_REPLY_TO   = 2byproduction@gmail.com   (אופציונלי)
 *
 *  אם אין RESEND_API_KEY — נופלים אוטומטית לשליחה דרך Gmail
 *  (MailApp), כדי שההורים תמיד יקבלו אישור.
 ****************************************************************/
function sendConfirmationEmail(data) {
  if (!data.email) return false;

  const props   = PropertiesService.getScriptProperties();
  const apiKey  = props.getProperty('RESEND_API_KEY');
  const from    = props.getProperty('MAIL_FROM') || 'סיפור חדש מתחיל <onboarding@resend.dev>';
  const replyTo = props.getProperty('MAIL_REPLY_TO');

  const subject = '✨ ההרשמה של ' + (data.fname || 'הילד/ה') + ' התקבלה — סיפור חדש מתחיל!';
  const html    = buildEmailHtml(data);

  // --- שליחה דרך Resend ---
  if (apiKey) {
    const payload = { from: from, to: [data.email], subject: subject, html: html };
    if (replyTo) payload.reply_to = replyTo;

    const res = UrlFetchApp.fetch('https://api.resend.com/emails', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() < 300) return true;
    Logger.log('Resend החזיר שגיאה (' + res.getResponseCode() + '): ' + res.getContentText());
    // ממשיכים לניסיון גיבוי דרך Gmail
  }

  // --- גיבוי: Gmail ---
  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    htmlBody: html,
    name: 'סיפור חדש מתחיל · סלקום',
    replyTo: replyTo || undefined
  });
  return true;
}

/* בניית גוף המייל — HTML ממותג (עיצוב inline כדי שיעבוד בכל תוכנות המייל) */
function buildEmailHtml(data) {
  const kid    = (data.fname || '') + (data.lname ? ' ' + data.lname : '');
  const parent = data.parent || 'הורים יקרים';

  // פונטים של המותג + פולבק נקי לתוכנות מייל שלא טוענות פונטי רשת (Gmail)
  const FONT      = "'Rubik','Segoe UI','Helvetica Neue',Arial,sans-serif";
  const FONT_HEAD = "'Varela Round','Rubik','Segoe UI','Helvetica Neue',Arial,sans-serif";

  const scheduleRows = EVENT.schedule.map(function(row) {
    return '<tr>' +
      '<td style="font-family:' + FONT_HEAD + ';padding:7px 0;font-weight:700;color:#E6189C;white-space:nowrap;width:56px;font-size:15px">' + row[0] + '</td>' +
      '<td style="font-family:' + FONT + ';padding:7px 12px 7px 0;color:#3A1670;font-size:15px">' + row[1] + '</td>' +
    '</tr>';
  }).join('');

  return '' +
  '<style>@import url(https://fonts.googleapis.com/css2?family=Varela+Round&family=Rubik:wght@400;500;700;900&display=swap);</style>' +
  '<div dir="rtl" style="margin:0;padding:0;background:#F4ECFC">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#F4ECFC;padding:28px 12px">' +
      '<tr><td align="center">' +
        '<table role="presentation" width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background:#FFFFFF;border-radius:22px;overflow:hidden;box-shadow:0 18px 40px -18px rgba(74,30,140,.35)">' +

          /* ---- header ---- */
          '<tr><td align="center" bgcolor="#4A1E8C" style="background:linear-gradient(135deg,#4A1E8C 0%,#7B3FE4 100%);padding:34px 24px 30px">' +
            '<div style="font-family:' + FONT + ';font-size:15px;color:#E9D8FF;letter-spacing:1px;margin-bottom:6px">✦ ✦ ✦</div>' +
            '<div style="font-family:' + FONT_HEAD + ';font-weight:900;font-size:30px;line-height:1.25;color:#FFFFFF">סיפור חדש מתחיל</div>' +
            '<div style="font-family:' + FONT + ';font-size:15px;color:#FFD9F0;margin-top:6px">חגיגת העלייה לכיתה א\' · ארגון עובדי קבוצת סלקום</div>' +
          '</td></tr>' +

          /* ---- body ---- */
          '<tr><td style="padding:30px 30px 8px;font-family:' + FONT + '">' +
            '<div style="font-family:' + FONT_HEAD + ';font-size:20px;font-weight:800;color:#6322D6;margin-bottom:12px">היי ' + parent + ', ההרשמה התקבלה! 🎉</div>' +
            '<div style="font-size:15.5px;line-height:1.8;color:#4a3a6e">' +
              'הפרק הראשון בסיפור של <b style="color:#E6189C">' + kid + '</b> נכתב בהצלחה!<br>' +
              'קיבלנו את הפרטים, התמונה והברכה שלכם — והכול מוכן לקראת החגיגה הגדולה.' +
            '</div>' +
          '</td></tr>' +

          /* ---- event card ---- */
          '<tr><td style="padding:18px 30px 6px">' +
            '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#FBF3FF;border:2px solid #E3D2F7;border-radius:16px">' +
              '<tr><td style="padding:20px 22px;font-family:' + FONT + '">' +
                '<div style="font-family:' + FONT_HEAD + ';font-size:17px;font-weight:800;color:#6322D6;margin-bottom:10px">📅 ' + EVENT.date + ' · 📍 ' + EVENT.place + '</div>' +
                '<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%">' + scheduleRows + '</table>' +
              '</td></tr>' +
            '</table>' +
          '</td></tr>' +

          /* ---- footer note ---- */
          '<tr><td style="padding:18px 30px 28px;font-family:' + FONT + '">' +
            '<div style="font-size:14px;line-height:1.75;color:#6A4D9E">' +
              'נשלח אליכם עדכון נוסף לקראת האירוע.<br>' +
              'יש שאלה? פשוט השיבו למייל הזה ונחזור אליכם.' +
            '</div>' +
            '<div style="font-size:13px;color:#B7A3DC;margin-top:18px;border-top:1px dashed #E3D2F7;padding-top:14px;text-align:center">' +
              'נתראה בספארי! 💜 ארגון עובדי קבוצת סלקום' +
            '</div>' +
          '</td></tr>' +

        '</table>' +
      '</td></tr>' +
    '</table>' +
  '</div>';
}

/****************************************************************
 *  📧 בדיקת מייל — הריצי פעם אחת כדי לוודא שהחיבור ל-Resend עובד
 *  (עדכני קודם את כתובת המייל לבדיקה)
 ****************************************************************/
function testEmail() {
  const ok = sendConfirmationEmail({
    fname: 'יהלי', lname: 'לוי',
    parent: 'בדיקה',
    email: '2byproduction@gmail.com'   // ← כתובת לבדיקה
  });
  Logger.log(ok ? '✓ מייל הבדיקה נשלח!' : '✗ המייל לא נשלח — בדקי את ה-Script Properties');
}

/****************************************************************
 *  ⭐ הריצי את הפונקציה הזו פעם אחת (כפתור Run למעלה)
 *  היא תבקש את כל ההרשאות (כולל Drive) ותיצור את כותרות הגיליון.
 *  אחרי שתאשרי — הטופס יעבוד מלא.
 ****************************************************************/
function setup() {
  // נגיעה ב-Drive: מכריחה את גוגל לבקש הרשאת Drive מלאה
  const folder = DriveApp.getFolderById(FOLDER_ID);

  // יצירת כותרות בגיליון (אם ריק)
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }

  Logger.log('✓ הכל מחובר! תיקיית Drive: "' + folder.getName() + '" | גיליון: "' + sheet.getName() + '"');
}
