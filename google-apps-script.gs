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
const HEADERS = ['תאריך ושעה','שם הילד/ה','שם משפחה','חוזקות','ברכה אישית','חלומות','שם ההורה','טלפון','אימייל','תמונה'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // ----- 1) שמירת התמונה ל-Drive -----
    let photoUrl = '';
    if (data.photoData && data.photoData.indexOf(',') > -1) {
      const parts       = data.photoData.split(',');
      const contentType = (parts[0].match(/data:(.*?);/) || [])[1] || 'image/jpeg';
      const ext         = (contentType.split('/')[1] || 'jpg').replace('jpeg','jpg');
      const bytes       = Utilities.base64Decode(parts[1]);
      const fileName    = (data.photoName || 'registration') + '.' + ext;
      const blob        = Utilities.newBlob(bytes, contentType, fileName);

      const folder = DriveApp.getFolderById(FOLDER_ID);
      const file   = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = file.getUrl();
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
      photoUrl
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, photoUrl: photoUrl }))
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
