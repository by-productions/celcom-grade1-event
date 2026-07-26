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
const HEADERS = ['תאריך ושעה','שם הילד/ה','שם משפחה','חוזקות','ברכה אישית','חלומות','שם ההורה','טלפון','אימייל','תמונה','מספר עובד','רגישויות','כשרויות מיוחדות','אופן הגעה','תחנת הסעה','כמות בהסעה'];

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

// גיליון הנתונים = הלשונית הראשונה שאיננה "סיכום" (יציב גם אחרי שנוצרת לשונית הסיכום)
function getDataSheet() {
  const sheets = SpreadsheetApp.openById(SHEET_ID).getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== 'סיכום') return sheets[i];
  }
  return sheets[0];
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // ----- 1) שמירת התמונה ל-Drive -----
    let photoUrl = '';
    // חסם גודל (הגנה מפני ניצול לרעה/מילוי Drive) — ~10MB בקירוב
    if (data.photoData && data.photoData.length > 14000000) {
      photoUrl = 'תמונה גדולה מדי';
    } else if (data.photoData && data.photoData.indexOf(',') > -1) {
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
    const sheet = getDataSheet();
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
      "'" + (data.employeeId || ''),   // הגרש שומר על אפס מוביל במספר עובד
      data.sensitivities || '',
      data.kosher        || '',
      data.transport     || '',
      data.busStation    || '',
      data.busCount      || ''
    ]);

    // ----- 3) עדכון לשונית "סיכום" (הסעות/כמויות) — לא חוסם את ההרשמה -----
    try { updateSummary(); } catch (sumErr) { Logger.log('שגיאה בעדכון סיכום: ' + sumErr); }

    // ----- 4) מייל אישור הרשמה להורה (Resend) -----
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
  // אימות בסיסי של כתובת המייל — לא שולחים לכתובות לא-תקינות (הגנה מפני ניצול לרעה)
  if (!data.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(data.email)) || String(data.email).length > 254) return false;

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
  // בריחה מ-HTML כדי למנוע הזרקת תגיות/קישורים דרך שדות שהמשתמש ממלא
  const esc = function (s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  };
  const kid    = esc((data.fname || '') + (data.lname ? ' ' + data.lname : ''));
  const parent = esc(data.parent || 'הורים יקרים');

  // פונטים של המותג + פולבק נקי לתוכנות מייל שלא טוענות פונטי רשת (Gmail)
  const FONT      = "'Rubik','Segoe UI','Helvetica Neue',Arial,sans-serif";
  const FONT_HEAD = "'Rubik','Segoe UI','Helvetica Neue',Arial,sans-serif";

  const scheduleRows = EVENT.schedule.map(function(row) {
    return '<tr>' +
      '<td style="font-family:' + FONT_HEAD + ';padding:7px 0;font-weight:700;color:#E6189C;white-space:nowrap;width:56px;font-size:15px">' + row[0] + '</td>' +
      '<td style="font-family:' + FONT + ';padding:7px 12px 7px 0;color:#3A1670;font-size:15px">' + row[1] + '</td>' +
    '</tr>';
  }).join('');

  return '' +
  '<style>@import url(https://fonts.googleapis.com/css2?family=Rubik:wght@400;500;600;700;800;900&display=swap);</style>' +
  '<div dir="rtl" style="margin:0;padding:0;background:#F4ECFC">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#F4ECFC;padding:28px 12px">' +
      '<tr><td align="center">' +
        '<table role="presentation" width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background:#FFFFFF;border-radius:22px;overflow:hidden;box-shadow:0 18px 40px -18px rgba(74,30,140,.35)">' +

          /* ---- header ---- */
          '<tr><td align="center" bgcolor="#4A1E8C" style="background:linear-gradient(135deg,#4A1E8C 0%,#7B3FE4 100%);padding:34px 24px 30px">' +
            '<div style="font-family:' + FONT + ';font-size:15px;color:#E9D8FF;letter-spacing:1px;margin-bottom:6px">✦ ✦ ✦</div>' +
            '<div style="font-family:' + FONT_HEAD + ';font-weight:900;font-size:30px;line-height:1.25;color:#FFFFFF">סיפור חדש מתחיל</div>' +
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
              'יש שאלה? כתבו לנו ל־<a href="mailto:RevahaMashotafat@cellcom.co.il" style="color:#7B3FE4;font-weight:700;text-decoration:none">RevahaMashotafat@cellcom.co.il</a>' +
            '</div>' +
            '<div style="font-size:13px;color:#B7A3DC;margin-top:18px;border-top:1px dashed #E3D2F7;padding-top:14px;text-align:center">' +
              'נתראה בספארי! 💜' +
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

  // כתיבה/עדכון של שורת הכותרות תמיד (שורה 1) — בלי לפגוע בנתונים שכבר קיימים למטה
  updateHeaders();

  Logger.log('✓ הכל מחובר! תיקיית Drive: "' + folder.getName() + '" | גיליון: "' + getDataSheet().getName() + '"');
}

/****************************************************************
 *  📋 עדכון שורת הכותרות בלבד — הריצי אם הוספתָ עמודות חדשות
 *  (כותב מחדש את שורה 1 עם כל הכותרות; הנתונים בשורות 2+ נשמרים)
 ****************************************************************/
function updateHeaders() {
  const sheet = getDataSheet();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  Logger.log('✓ הכותרות עודכנו (' + HEADERS.length + ' עמודות): ' + HEADERS.join(' | '));
}

/****************************************************************
 *  📊 לשונית "סיכום" — כמה מגיעים בהסעה, מאיפה וכמה משתתפים
 *  מתעדכנת אוטומטית בכל הרשמה; אפשר גם להריץ ידנית.
 ****************************************************************/
function updateSummary() {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const data = getDataSheet();
  const values = data.getDataRange().getValues();
  if (values.length < 1) return;

  const header = values[0];
  const iTransport = header.indexOf('אופן הגעה');
  const iStation   = header.indexOf('תחנת הסעה');
  const iCount     = header.indexOf('כמות בהסעה');
  const rows = values.slice(1).filter(function(r){ return r.join('') !== ''; });

  let self = 0, bus = 0, busPeople = 0;
  const byStation = {};   // תחנה -> {families, people}
  rows.forEach(function(r){
    const t = iTransport > -1 ? r[iTransport] : '';
    if (t === 'עם ההסעה') {
      bus++;
      const st = (iStation > -1 && r[iStation]) ? r[iStation] : '(ללא תחנה)';
      const c  = iCount > -1 ? (parseInt(String(r[iCount]).replace(/\D/g, ''), 10) || 0) : 0;
      busPeople += c;
      if (!byStation[st]) byStation[st] = { families: 0, people: 0 };
      byStation[st].families++;
      byStation[st].people += c;
    } else if (t === 'עצמאית') {
      self++;
    }
  });

  const out = [
    ['סיכום הרשמות', '', ''],
    ['סה"כ נרשמים', rows.length, ''],
    ['מגיעים עצמאית', self, ''],
    ['מגיעים עם ההסעה', bus, ''],
    ['סה"כ משתתפים בהסעות', busPeople, ''],
    ['', '', ''],
    ['תחנת הסעה', 'מס\' רישומים', 'סה"כ משתתפים']
  ];
  Object.keys(byStation).sort().forEach(function(st){
    out.push([st, byStation[st].families, byStation[st].people]);
  });

  let sum = ss.getSheetByName('סיכום');
  if (!sum) sum = ss.insertSheet('סיכום');   // נוספת בסוף — לא מזיזה את גיליון הנתונים
  sum.clear();
  sum.getRange(1, 1, out.length, 3).setValues(out);
  sum.getRange(1, 1, 1, 3).setFontWeight('bold').setFontSize(14);
  sum.getRange(7, 1, 1, 3).setFontWeight('bold').setBackground('#EDE7FB');
  sum.setColumnWidth(1, 340);
  Logger.log('✓ לשונית "סיכום" עודכנה. הסעה: ' + bus + ' | עצמאית: ' + self + ' | משתתפים בהסעה: ' + busPeople);
}

/****************************************************************
 *  🩺 אבחון גיליון — כמה לשוניות/שורות/עמודות יש בפועל
 *  שימושי כשהגיליון נטען לאט: מריצים כאן (לא בדפדפן) ורואים ב-Execution log.
 ****************************************************************/
function diagnose() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheets = ss.getSheets();
  let out = 'מספר לשוניות: ' + sheets.length + '\n';
  sheets.forEach(function (s) {
    out += '• "' + s.getName() + '" — נתונים: ' + s.getLastRow() + ' שורות × ' + s.getLastColumn() +
           ' עמודות | גודל הגיליון בפועל: ' + s.getMaxRows() + ' × ' + s.getMaxColumns() + '\n';
  });
  Logger.log(out);
  return out;
}

/****************************************************************
 *  ✂️ ניקוי שורות/עמודות ריקות מיותרות — מאיץ טעינה של גיליון כבד
 *  (משאיר מרווח של 200 שורות פנויות להרשמות חדשות)
 ****************************************************************/
function trimSheet() {
  const s = getDataSheet();
  const lastRow = Math.max(s.getLastRow(), 1);
  const lastCol = Math.max(s.getLastColumn(), 1);
  const keepRows = lastRow + 200;
  let removedRows = 0, removedCols = 0;
  if (s.getMaxRows() > keepRows) {
    removedRows = s.getMaxRows() - keepRows;
    s.deleteRows(keepRows + 1, removedRows);
  }
  if (s.getMaxColumns() > lastCol) {
    removedCols = s.getMaxColumns() - lastCol;
    s.deleteColumns(lastCol + 1, removedCols);
  }
  Logger.log('✓ "' + s.getName() + '": נמחקו ' + removedRows + ' שורות ריקות ו-' + removedCols + ' עמודות ריקות');
}

/****************************************************************
 *  🧹 איפוס נתוני בדיקה — מוחק את כל שורות הנתונים (הכותרות נשמרות)
 *  ⚠️ פעולה בלתי הפיכה. הריצי רק כשרוצים להתחיל נקי לפני האירוע.
 ****************************************************************/
function resetData() {
  const sheet = getDataSheet();
  const last  = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).clearContent();
  try { updateSummary(); } catch (e) {}
  Logger.log('✓ נמחקו ' + Math.max(0, last - 1) + ' שורות נתונים (הכותרות נשמרו). התמונות ב-Drive לא נמחקו.');
}

/****************************************************************
 *  🖼️ מחולל מצגת Google Slides — כל התעודות במצגת אחת, ניתנת לעריכה
 *  ----------------------------------------------------------------
 *  buildCertificatesSlides()   — מוסיף רק נרשמים חדשים (שומר עריכות ידניות שכבר עשית).
 *  rebuildCertificatesSlides() — בונה מצגת חדשה מאפס לכל הנרשמים.
 *
 *  כל ילד = 2 שקופיות (כריכה + פנים). הטקסט אמיתי וניתן לעריכה ישירות ב-Slides.
 *  התמונה מוכנסת מאחורי "חור הענן" — אפשר להזיז/לחתוך אותה ב-Slides והיא נשארת בצורת ענן.
 *  הרצה ראשונה תבקש הרשאות ל-Slides + Drive.
 ****************************************************************/
const SLIDES_ASSETS = 'https://raw.githubusercontent.com/by-productions/celcom-grade1-event/main/certificates/assets/';

// מיקומים ביחס לתבנית (שברים של 2000×1414) — כאן מכווננים אם צריך
const CERT_L = {
  photo:  { l: 0.545, t: 0.1485, w: 0.3735, h: 0.7376 },
  banner: { l: 0.150, t: 0.082,  w: 0.225,  h: 0.078 },
  sec1:   { l: 0.121, t: 0.250,  w: 0.278,  h: 0.155 },
  sec2:   { l: 0.121, t: 0.430,  w: 0.278,  h: 0.155 },
  sec3:   { l: 0.121, t: 0.610,  w: 0.278,  h: 0.165 },
  letter: { l: 0.121, t: 0.232,  w: 0.286,  h: 0.460 },
  emp:    { l: 0.194, t: 0.826,  w: 0.142,  h: 0.080 }
};
const CERT_C = { pink: '#E6189C', purple: '#7B3FE4', purple2: '#6322D6', body: '#5B4A8E', white: '#FFFFFF', letter: '#4A1E8C', lav: '#EFE8FB' };

// מכתב הכריכה האחורית (זהה לזה שבמחולל ה-HTML)
const COVER_LETTER = [
  'אנחנו רגילים לחבר בין אנשים, אבל היום אנחנו חוגגים חיבור מיוחד במינו – את החיבור שלך לעולם חדש של למידה, סקרנות, חברים, חוויות וגילויים.',
  'המעבר לכיתה א\' הוא צעד גדול ומרגש. זהו הרגע שבו את/ה מתחיל/ה לכתוב פרק חדש בסיפור שלך – פרק מלא בהרפתקאות, הצלחות, למידה וחלומות שמתחילים להתגשם.',
  'בתוך הספרון הזה תמצא/י את כל הדברים שהופכים אותך למיוחד/ת: החוזקות שלך, החלומות שלך והמילים שהאנשים שאוהבים אותך יותר מכל בחרו לכתוב לך כדי ללוות אותך בדרך.',
  'כשתפתח/י את הספרון הזה בעוד שנים, אנחנו מקווים שתיזכר/י ביום המיוחד שבו יצאת למסע החדש שלך – עם חיוך, אומץ, סקרנות והמון אמונה בעצמך.',
  'אנחנו מאחלים לך שנה מלאה בחוויות, חברים טובים, הצלחות קטנות וגדולות, והרבה רגעים של שמחה וגאווה.'
];

function rebuildCertificatesSlides() { return _certSlides(true); }
function buildCertificatesSlides()   { return _certSlides(false); }

function _certSlides(fresh) {
  const props  = PropertiesService.getScriptProperties();
  const data   = getDataSheet();
  const values = data.getDataRange().getValues();
  const header = values[0];
  const idx = function (n) { return header.indexOf(n); };
  const iF = idx('שם הילד/ה'), iS = idx('חוזקות'), iD = idx('חלומות'),
        iB = idx('ברכה אישית'), iP = idx('תמונה'), iE = idx('מספר עובד');

  let pres, startRow;
  const savedId = props.getProperty('SLIDES_ID');
  if (!fresh && savedId) { try { pres = SlidesApp.openById(savedId); } catch (e) { pres = null; } }
  if (fresh || !pres) {
    pres = SlidesApp.create('תעודות · סיפור חדש מתחיל');
    try { pres.getSlides()[0].remove(); } catch (e) {}
    props.setProperty('SLIDES_ID', pres.getId());
    startRow = 1;
  } else {
    startRow = parseInt(props.getProperty('SLIDES_LAST_ROW') || '1', 10);
  }

  const coverBlob = UrlFetchApp.fetch(SLIDES_ASSETS + 'cover.jpg').getBlob();
  const holeBlob  = UrlFetchApp.fetch(SLIDES_ASSETS + 'inside-hole.png').getBlob();
  const fit = _certFit(pres);

  let added = 0;
  for (let r = Math.max(startRow, 1); r < values.length; r++) {
    const row = values[r];
    if (String(row.join('')).trim() === '') continue;
    const rec = {
      fname:     (row[iF] || '').toString().trim(),
      strengths: iS > -1 ? (row[iS] || '').toString().trim() : '',
      dreams:    iD > -1 ? (row[iD] || '').toString().trim() : '',
      blessing:  iB > -1 ? (row[iB] || '').toString().trim() : '',
      photo:     iP > -1 ? (row[iP] || '').toString().trim() : '',
      emp:       iE > -1 ? (row[iE] || '').toString().replace(/^'/, '').trim() : ''
    };
    if (!rec.fname && !rec.strengths && !rec.blessing) continue;
    _certCover(pres, fit, coverBlob, rec);
    _certInside(pres, fit, holeBlob, rec);
    added++;
  }
  props.setProperty('SLIDES_LAST_ROW', String(values.length));
  const msg = '✓ המצגת מוכנה (' + added + ' תעודות חדשות). קישור: ' + pres.getUrl();
  Logger.log(msg);
  return msg;
}

// מלבן התבנית בתוך השקופית (שומר יחס 2000:1414, ממורכז)
function _certFit(pres) {
  const W = pres.getPageWidth(), H = pres.getPageHeight(), R = 2000 / 1414;
  let w = W, h = W / R;
  if (h > H) { h = H; w = H * R; }
  return { x: (W - w) / 2, y: (H - h) / 2, w: w, h: h, W: W, H: H };
}
function _px(fit, box) { return { L: fit.x + box.l * fit.w, T: fit.y + box.t * fit.h, W: box.w * fit.w, H: box.h * fit.h }; }

// תיבת טקסט (ללא מסגרת), מיושרת לימין כברירת מחדל
function _certText(slide, fit, box, opts) {
  opts = opts || {};
  const p = _px(fit, box);
  const sh = slide.insertTextBox(opts.text || ' ', p.L, p.T, p.W, p.H);   // לעולם לא ריק
  try { sh.getBorder().setTransparent(); } catch (e) {}
  const t = sh.getText();
  if (t.asString().length > 0) {
    t.getParagraphs().forEach(function (par) {
      par.getRange().getParagraphStyle().setParagraphAlignment(opts.align || SlidesApp.ParagraphAlignment.END);
    });
    const st = t.getTextStyle();
    st.setFontFamily(opts.font || 'Rubik').setForegroundColor(opts.color || CERT_C.body).setFontSize(opts.size || (0.024 * fit.h));
    if (opts.bold) st.setBold(true);
  }
  try { sh.setContentAlignment(opts.valign || SlidesApp.ContentAlignment.TOP); } catch (e) {}
  return sh;
}

function _certCover(pres, fit, coverBlob, rec) {
  const s = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s.insertImage(coverBlob, fit.x, fit.y, fit.w, fit.h);
  const letterText = 'הפרק החדש שלך מתחיל כאן\n' + COVER_LETTER.join('\n\n') + '\n\nבהצלחה בכיתה א׳ – אנחנו מאמינים בך!';
  const box = _certText(s, fit, CERT_L.letter, { text: letterText, color: CERT_C.letter, size: 0.017 * fit.h });
  const t = box.getText();
  const paras = t.getParagraphs();
  paras[0].getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  paras[0].getRange().getTextStyle().setBold(true).setForegroundColor(CERT_C.pink).setFontSize(0.023 * fit.h).setFontFamily('Varela Round');
  const lastP = paras[paras.length - 1];
  lastP.getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  lastP.getRange().getTextStyle().setBold(true).setForegroundColor(CERT_C.purple2);
  if (rec.emp) {
    const e = _certText(s, fit, CERT_L.emp, { text: 'מספר עובד\n' + rec.emp, color: CERT_C.purple2, size: 0.014 * fit.h,
      align: SlidesApp.ParagraphAlignment.CENTER, valign: SlidesApp.ContentAlignment.MIDDLE });
    e.getText().getParagraphs()[1].getRange().getTextStyle().setBold(true).setFontSize(0.017 * fit.h);
  }
}

function _certInside(pres, fit, holeBlob, rec) {
  const s = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const cb = _px(fit, CERT_L.photo);
  const bg = s.insertShape(SlidesApp.ShapeType.RECTANGLE, cb.L, cb.T, cb.W, cb.H);
  bg.getFill().setSolidFill(CERT_C.lav); try { bg.getBorder().setTransparent(); } catch (e) {}
  const blob = _certPhotoBlob(rec.photo);
  if (blob) {
    try {
      const img = s.insertImage(blob);
      const nW = img.getWidth(), nH = img.getHeight();
      const sc = Math.max(cb.W / nW, cb.H / nH);
      const w = nW * sc, h = nH * sc;
      img.setWidth(w).setHeight(h).setLeft(cb.L + (cb.W - w) / 2).setTop(cb.T + (cb.H - h) * 0.28);
    } catch (e) { Logger.log('תמונה נכשלה: ' + e); }
  }
  s.insertImage(holeBlob, fit.x, fit.y, fit.w, fit.h);
  _certText(s, fit, CERT_L.banner, { text: 'על ' + rec.fname, color: CERT_C.white, font: 'Varela Round',
    bold: true, size: 0.037 * fit.h, align: SlidesApp.ParagraphAlignment.CENTER, valign: SlidesApp.ContentAlignment.MIDDLE });
  _certSection(s, fit, CERT_L.sec1, 'החוזקות של ' + rec.fname, rec.strengths, CERT_C.pink);
  _certSection(s, fit, CERT_L.sec2, 'החלומות של ' + rec.fname, rec.dreams,   CERT_C.purple);
  _certSection(s, fit, CERT_L.sec3, 'הברכה שלנו',              rec.blessing,  CERT_C.pink);
}

function _certSection(slide, fit, box, head, body, headColor) {
  const sh = _certText(slide, fit, box, { text: head + '\n' + body, color: CERT_C.body, size: 0.023 * fit.h });
  const paras = sh.getText().getParagraphs();
  paras[0].getRange().getTextStyle().setBold(true).setForegroundColor(headColor).setFontSize(0.030 * fit.h).setFontFamily('Varela Round');
  return sh;
}

function _certPhotoBlob(url) {
  if (!url) return null;
  const m = String(url).match(/[-\w]{25,}/);
  if (!m) return null;
  try { return DriveApp.getFileById(m[0]).getBlob(); } catch (e) { Logger.log('Drive: ' + e); return null; }
}
