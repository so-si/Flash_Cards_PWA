const SPREADSHEET_ID = '1IYyx3Peb6Jkiq8miyZaTOpfhoHKR5rGQ1qhwhLrdDXw';
const FLASH_SHEET_NAME = 'Flash';
const FLASH_HISTORY_SHEET_NAME = 'FlashHistory';
const JST_TIMEZONE = 'Asia/Tokyo';
const FLASH_DAILY_GOAL = 50;

function doGet(e) {
  const parameters = e && e.parameter ? e.parameter : {};
  if (!parameters.action) {
    const template = HtmlService.createTemplateFromFile('Index');
    template.apiUrl = ScriptApp.getService().getUrl();
    return template.evaluate().setTitle('Flash Cards')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  }
  const callback = normalizeCallback_(parameters.callback);
  try {
    verifyApiToken_(parameters.token);
    const payload = parameters.payload ? JSON.parse(parameters.payload) : {};
    let response;
    switch (parameters.action) {
      case 'getFlashCards': response = { success: true, cards: getFlashCards() }; break;
      case 'saveFlashResult':
        response = saveFlashResult(payload.cardNo, payload.subject, payload.userAnswer,
          payload.correctAnswer, payload.isCorrect, payload.eventId, payload.timestamp); break;
      case 'getFlashStats': response = { success: true, stats: getFlashStats() }; break;
      case 'getFlashDailyStatus': response = { success: true, status: getFlashDailyStatus() }; break;
      case 'getFlashQuestionStats': response = { success: true, stats: getFlashQuestionStats(payload.cardNo) }; break;
      case 'getFlashStatsBundle': response = { success: true, bundle: getFlashStatsBundle() }; break;
      default: throw new Error('不明なAPIアクションです。');
    }
    return createApiOutput_(callback, response);
  } catch (error) {
    return createApiOutput_(callback, { success: false, error: error && error.message ? error.message : String(error) });
  }
}

function normalizeCallback_(callback) {
  const value = normalizeText(callback);
  if (!/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(value)) throw new Error('callbackが不正です。');
  return value;
}
function verifyApiToken_(providedToken) {
  const expectedToken = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  if (expectedToken && normalizeText(providedToken) !== expectedToken) throw new Error('APIトークンが不正です。');
}
function createApiOutput_(callback, data) {
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(data) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function getSpreadsheet_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function normalizeText(value) { return value === null || value === undefined ? '' : String(value).trim(); }
function formatJstDate_(date) { return Utilities.formatDate(date, JST_TIMEZONE, 'yyyy-MM-dd'); }
function normalizeStudyDate_(value, fallbackValue) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return formatJstDate_(value);
  const text = normalizeText(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  if (text) { const parsed = new Date(text); if (!Number.isNaN(parsed.getTime())) return formatJstDate_(parsed); }
  if (fallbackValue) { const d = fallbackValue instanceof Date ? fallbackValue : new Date(fallbackValue); if (!Number.isNaN(d.getTime())) return formatJstDate_(d); }
  return '';
}
function getFlashSheet_() {
  const sheet = getSpreadsheet_().getSheetByName(FLASH_SHEET_NAME);
  if (!sheet) throw new Error('Flash シートが見つかりません。');
  return sheet;
}
function getFlashHistorySheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(FLASH_HISTORY_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(FLASH_HISTORY_SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, 8).setValues([['日時','問','科目','回答','正答','正誤','EventID','学習日(JST)']]);
  return sheet;
}

function buildDetailUrl_(precedent, source, question) {
  const q = ['site:law-lib.jp', normalizeText(precedent), normalizeText(source), normalizeText(question).slice(0, 60)]
    .filter(Boolean).join(' ');
  return q ? 'https://www.google.com/search?q=' + encodeURIComponent(q) : '';
}
function fallbackExplanation_(answer) {
  return answer === '○'
    ? '○。この肢は判例・法令の結論に沿います。詳しい根拠は「詳しい解説を見る」で確認してください。'
    : '×。この肢は判例・法令の結論と一致しません。誤りとなるポイントは「詳しい解説を見る」で確認してください。';
}

function getFlashCards() {
  const sheet = getFlashSheet_();
  const lastRow = sheet.getLastRow(), lastColumn = sheet.getLastColumn();
  if (lastRow < 2) return [];
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = values[0].map(normalizeText);
  const index = {
    no: headers.indexOf('問'), subject: headers.indexOf('科目'), precedent: headers.indexOf('判例名'),
    question: headers.indexOf('質問'), answer: headers.indexOf('正答'), source: headers.indexOf('出典'),
    explanation: headers.indexOf('解説'), detailUrl: headers.indexOf('詳細リンク')
  };
  ['no','subject','precedent','question','answer','source'].forEach(key => {
    if (index[key] < 0) throw new Error('Flash シートの列が不足しています: ' + key);
  });
  return values.slice(1).filter(row => normalizeText(row[index.question])).map(row => {
    const answer = normalizeFlashAnswer_(row[index.answer]);
    const precedent = normalizeText(row[index.precedent]);
    const source = normalizeText(row[index.source]);
    const question = normalizeText(row[index.question]);
    const explanation = index.explanation >= 0 ? normalizeText(row[index.explanation]) : '';
    const detailUrl = index.detailUrl >= 0 ? normalizeText(row[index.detailUrl]) : '';
    return {
      no: normalizeText(row[index.no]), subject: normalizeText(row[index.subject]), precedent,
      question, answer, source,
      explanation: explanation || fallbackExplanation_(answer),
      detailUrl: detailUrl || buildDetailUrl_(precedent, source, question)
    };
  });
}
function normalizeFlashAnswer_(value) {
  const text = normalizeText(value);
  if (text === '〇' || text === '○') return '○';
  if (text === '✕' || text === '×' || text.toLowerCase() === 'x') return '×';
  return text;
}

function saveFlashResult(cardNo, subject, userAnswer, correctAnswer, isCorrect, eventId, eventTimestamp) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const no = normalizeText(cardNo), normalizedSubject = normalizeText(subject);
    const answer = normalizeFlashAnswer_(userAnswer), correct = normalizeFlashAnswer_(correctAnswer);
    const normalizedEventId = normalizeText(eventId);
    if (!no || !answer || !correct) throw new Error('短答回答データが不正です。');
    const historySheet = getFlashHistorySheet_();
    if (normalizedEventId && flashEventIdExists_(historySheet, normalizedEventId)) return { success: true, duplicate: true };
    let eventDate = new Date();
    if (eventTimestamp) { const parsed = new Date(eventTimestamp); if (!Number.isNaN(parsed.getTime())) eventDate = parsed; }
    const correctFlag = isCorrect === true || String(isCorrect).toLowerCase() === 'true';
    historySheet.appendRow([eventDate,no,normalizedSubject,answer,correct,correctFlag,normalizedEventId,formatJstDate_(eventDate)]);
    return { success: true, duplicate: false };
  } finally { lock.releaseLock(); }
}
function flashEventIdExists_(sheet, eventId) {
  if (!eventId || sheet.getLastRow() < 2) return false;
  return Boolean(sheet.getRange(2,7,sheet.getLastRow()-1,1).createTextFinder(eventId).matchEntireCell(true).findNext());
}
function getFlashStats() {
  const b = getFlashStatsBundle().stats;
  return b;
}
function getFlashStatsBundle() {
  const today = formatJstDate_(new Date()), sheet = getFlashHistorySheet_(), lastRow = sheet.getLastRow();
  let total = 0, correct = 0, todayCount = 0; const questions = Object.create(null);
  if (lastRow >= 2) {
    sheet.getRange(2,1,lastRow-1,8).getValues().forEach(row => {
      const cardNo = normalizeText(row[1]); if (!cardNo) return;
      const ok = row[5] === true || String(row[5]).toLowerCase() === 'true';
      total++; if (ok) correct++;
      if (!questions[cardNo]) questions[cardNo] = { total:0, correct:0 };
      questions[cardNo].total++; if (ok) questions[cardNo].correct++;
      if (normalizeStudyDate_(row[7], row[0]) === today) todayCount++;
    });
  }
  return {
    stats:{total,correct,accuracy:total?correct/total:0},
    daily:{date:today,goal:FLASH_DAILY_GOAL,count:todayCount,remaining:Math.max(FLASH_DAILY_GOAL-todayCount,0),achieved:todayCount>=FLASH_DAILY_GOAL},
    questions
  };
}
function getFlashQuestionStats(cardNo) {
  const targetNo = normalizeText(cardNo); if (!targetNo) throw new Error('問番号が不正です。');
  const s = getFlashStatsBundle().questions[targetNo] || {total:0,correct:0};
  return { cardNo:targetNo,total:s.total,correct:s.correct,accuracy:s.total?s.correct/s.total:0 };
}
function getFlashDailyStatus() { return getFlashStatsBundle().daily; }
