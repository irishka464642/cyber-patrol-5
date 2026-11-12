/**
 * Google Apps Script receiver for "Спілкування онлайн" (5 клас).
 * Deploy as a web app: Execute as Me, access Anyone.
 * This code ONLY receives explicitly submitted quiz results.
 */
const TARGETS = Object.freeze({
  "5-А": Object.freeze({
    "1": "1o0ozxUICmgg4hSIQ382sfbjFFxXSh0TAdOznEzpYLp4",
    "2": "1k95-4TbeKOj3pMehAUhZ_VtMD45G_lNo93f7JPBRhEM"
  }),
  "5-Б": Object.freeze({
    "1": "1GLqm-9fv-HppbXYHniOGL-D8VUjxPr8dQghlGtbC034",
    "2": "1DxgJx952AEMguMpA9Okj0ADNWvbeGMapWBqgFNqmYR0"
  }),
  "5-В": Object.freeze({
    "1": "1HlPRx3bM5KYanQPJ4t3AeKK5VFWgbkv2aH55VfUIj-E",
    "2": "1xjVHN2EMfyCq0-vzrogKe_jQjY1dvICGupz_-cmg3og"
  })
});
const HEADERS = [
  "Дата отримання", "Клас", "Група", "Прізвище", "Ім’я",
  "Вправа", "Бали", "Максимум", "Спроба", "Час (с)"
];
const EXERCISES = [
  "Етикет у чаті",
  "Як відповісти?",
  "Halloween Decor Awards",
  "Месенджер",
  "Цифрові налаштування"
];
function doGet() {
  return ContentService.createTextOutput("Приймач результатів готовий.").setMimeType(ContentService.MimeType.TEXT);
}
function safeText(v) {
  return String(v == null ? "" : v).trim().replace(/[\r\n\t]/g, " ").slice(0, 65);
}
function safeCell(v) {
  const text = safeText(v);
  return /^[=+\-@]/.test(text) ? "'" + text : text; // prevent formula injection
}
function doPost(e) {
  try {
    const data = JSON.parse(e.postData && e.postData.contents || "{}");
    const schoolClass = safeText(data.schoolClass), group = safeText(data.group);
    const surname = safeText(data.surname), firstname = safeText(data.firstname);
    const scores = data.scores;
    const times = data.times;
    const attempt = Number(data.attempt);
    if (!TARGETS[schoolClass] || !TARGETS[schoolClass][group]) throw Error("Некоректний клас або група");
    if (!surname || !firstname || surname.length > 50 || firstname.length > 50) throw Error("Не вказано учня");
    if (!Array.isArray(scores) || scores.length !== 5 || scores.some(n=>!Number.isInteger(n)||n<0||n>3)) throw Error("Помилка балів");
    if (!Array.isArray(times) || times.length !== 5 || times.some(n=>!Number.isInteger(n)||n<0||n>3600)) throw Error("Помилка часу");
    if (!Number.isInteger(attempt) || attempt < 1 || attempt > 100) throw Error("Помилка спроби");
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const ss = SpreadsheetApp.openById(TARGETS[schoolClass][group]);
      const sheet = ss.getSheetByName("Результати");
      if (!sheet) throw Error("Немає аркуша Результати");
      const values = sheet.getRange(1,1,1,10).getValues()[0];
      if (HEADERS.some((h,i)=>values[i]!==h)) throw Error("Заголовки таблиці відрізняються");
      const date = new Date();
      const rows = EXERCISES.map((exercise,i)=>[
        date,schoolClass,group,safeCell(surname),safeCell(firstname),exercise,scores[i],3,attempt,times[i]
      ]);
      const startRow = sheet.getLastRow()+1;
      sheet.getRange(startRow,1,rows.length,HEADERS.length).setValues(rows);
      sheet.getRange(startRow,1,rows.length,1).setNumberFormat("dd.MM.yyyy HH:mm:ss");
      SpreadsheetApp.flush();
    } finally { lock.releaseLock(); }
    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
  } catch(err) {
    console.error(err);
    return ContentService.createTextOutput("ERROR: "+err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}
