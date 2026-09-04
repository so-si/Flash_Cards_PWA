const SPREADSHEET_ID = '1IYyx3Peb6Jkiq8miyZaTOpfhoHKR5rGQ1qhwhLrdDXw';
const FLASH_SHEET_NAME = 'Flash';
const FLASH_HISTORY_SHEET_NAME = 'FlashHistory';
const JST_TIMEZONE = 'Asia/Tokyo';
const FLASH_DAILY_GOAL = 50;

function doGet(e) {
  const parameters = e && e.parameter ? e.parameter : {};

  // action が無い場合は Flash Cards 本体を表示する。
  if (!parameters.action) {
    const template = HtmlService.createTemplateFromFile('Index');
    template.apiUrl = ScriptApp.getService().getUrl();
    return template
      .evaluate()
      .setTitle('Flash Cards')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  }

  const callback = normalizeCallback_(parameters.callback);

  try {
    verifyApiToken_(parameters.token);
    const payload = parameters.payload ? JSON.parse(parameters.payload) : {};
    let response;

    switch (parameters.action) {
      case 'getFlashCards':
        response = { success: true, cards: getFlashCards() };
        break;
      case 'saveFlashResult':
        response = saveFlashResult(
          payload.cardNo,
          payload.subject,
          payload.userAnswer,
          payload.correctAnswer,
          payload.isCorrect,
          payload.eventId,
          payload.timestamp
        );
        break;
      case 'getFlashStats':
        response = { success: true, stats: getFlashStats() };
        break;
      case 'getFlashDailyStatus':
        response = { success: true, status: getFlashDailyStatus() };
        break;
      case 'getFlashQuestionStats':
        response = { success: true, stats: getFlashQuestionStats(payload.cardNo) };
        break;
      case 'getFlashStatsBundle':
        response = { success: true, bundle: getFlashStatsBundle() };
        break;
      default:
        throw new Error('不明なAPIアクションです。');
    }

    return createApiOutput_(callback, response);
  } catch (error) {
    return createApiOutput_(callback, {
      success: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}

function normalizeCallback_(callback) {
  const value = normalizeText(callback);
  if (!/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(value)) {
    throw new Error('callbackが不正です。');
  }
  return value;
}

function verifyApiToken_(providedToken) {
  const expectedToken = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
  if (expectedToken && normalizeText(providedToken) !== expectedToken) {
    throw new Error('APIトークンが不正です。');
  }
}

function createApiOutput_(callback, data) {
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(data) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function normalizeText(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function formatJstDate_(date) {
  return Utilities.formatDate(date, JST_TIMEZONE, 'yyyy-MM-dd');
}

function normalizeStudyDate_(value, fallbackValue) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatJstDate_(value);
  }

  const text = normalizeText(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text;
  }

  if (text) {
    const parsed = new Date(text);
    if (!Number.isNaN(parsed.getTime())) {
      return formatJstDate_(parsed);
    }
  }

  if (fallbackValue) {
    const fallbackDate = fallbackValue instanceof Date
      ? fallbackValue
      : new Date(fallbackValue);
    if (!Number.isNaN(fallbackDate.getTime())) {
      return formatJstDate_(fallbackDate);
    }
  }

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

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 8).setValues([[
      '日時', '問', '科目', '回答', '正答', '正誤', 'EventID', '学習日(JST)'
    ]]);
  }
  return sheet;
}

function getFlashCards() {
  const sheet = getFlashSheet_();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = values[0].map(normalizeText);

  const index = {
    no: headers.indexOf('問'),
    subject: headers.indexOf('科目'),
    precedent: headers.indexOf('判例名'),
    question: headers.indexOf('質問'),
    answer: headers.indexOf('正答'),
    source: headers.indexOf('出典')
  };

  ['no', 'subject', 'precedent', 'question', 'answer', 'source'].forEach(key => {
    if (index[key] < 0) throw new Error('Flash シートの列が不足しています: ' + key);
  });

  return values.slice(1)
    .filter(row => normalizeText(row[index.question]))
    .map(row => ({
      no: normalizeText(row[index.no]),
      subject: normalizeText(row[index.subject]),
      precedent: normalizeText(row[index.precedent]),
      question: normalizeText(row[index.question]),
      answer: normalizeFlashAnswer_(row[index.answer]),
      source: normalizeText(row[index.source])
    }));
}

function normalizeFlashAnswer_(value) {
  const text = normalizeText(value);
  if (text === '〇' || text === '○') return '○';
  if (text === '✕' || text === '×' || text.toLowerCase() === 'x') return '×';
  return text;
}

function saveFlashResult(cardNo, subject, userAnswer, correctAnswer, isCorrect, eventId, eventTimestamp) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const no = normalizeText(cardNo);
    const normalizedSubject = normalizeText(subject);
    const answer = normalizeFlashAnswer_(userAnswer);
    const correct = normalizeFlashAnswer_(correctAnswer);
    const normalizedEventId = normalizeText(eventId);

    if (!no || !answer || !correct) {
      throw new Error('短答回答データが不正です。');
    }

    const historySheet = getFlashHistorySheet_();

    if (normalizedEventId && flashEventIdExists_(historySheet, normalizedEventId)) {
      return { success: true, duplicate: true };
    }

    let eventDate = new Date();
    if (eventTimestamp) {
      const parsed = new Date(eventTimestamp);
      if (!Number.isNaN(parsed.getTime())) eventDate = parsed;
    }

    const correctFlag = isCorrect === true || String(isCorrect).toLowerCase() === 'true';
    const studyDate = formatJstDate_(eventDate);

    historySheet.appendRow([
      eventDate,
      no,
      normalizedSubject,
      answer,
      correct,
      correctFlag,
      normalizedEventId,
      studyDate
    ]);

    return { success: true, duplicate: false };
  } finally {
    lock.releaseLock();
  }
}

function flashEventIdExists_(sheet, eventId) {
  if (!eventId) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const found = sheet
    .getRange(2, 7, lastRow - 1, 1)
    .createTextFinder(eventId)
    .matchEntireCell(true)
    .findNext();

  return Boolean(found);
}

function getFlashStats() {
  const sheet = getFlashHistorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { total: 0, correct: 0, accuracy: 0 };

  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  let total = 0;
  let correct = 0;

  values.forEach(row => {
    if (!row[1]) return;
    total += 1;
    if (row[5] === true || String(row[5]).toLowerCase() === 'true') correct += 1;
  });

  return {
    total: total,
    correct: correct,
    accuracy: total ? correct / total : 0
  };
}

function getFlashStatsBundle() {
  const today = formatJstDate_(new Date());
  const sheet = getFlashHistorySheet_();
  const lastRow = sheet.getLastRow();
  let total = 0;
  let correct = 0;
  let todayCount = 0;
  const questions = Object.create(null);

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

    values.forEach(row => {
      const cardNo = normalizeText(row[1]);
      if (!cardNo) return;

      const isCorrect = row[5] === true || String(row[5]).toLowerCase() === 'true';

      total += 1;
      if (isCorrect) correct += 1;

      if (!questions[cardNo]) questions[cardNo] = { total: 0, correct: 0 };
      questions[cardNo].total += 1;
      if (isCorrect) questions[cardNo].correct += 1;

      const studyDate = normalizeStudyDate_(row[7], row[0]);
      if (studyDate === today) todayCount += 1;
    });
  }

  return {
    stats: {
      total: total,
      correct: correct,
      accuracy: total ? correct / total : 0
    },
    daily: {
      date: today,
      goal: FLASH_DAILY_GOAL,
      count: todayCount,
      remaining: Math.max(FLASH_DAILY_GOAL - todayCount, 0),
      achieved: todayCount >= FLASH_DAILY_GOAL
    },
    questions: questions
  };
}

function getFlashQuestionStats(cardNo) {
  const targetNo = normalizeText(cardNo);
  if (!targetNo) throw new Error('問番号が不正です。');

  const stats = getFlashStatsBundle().questions[targetNo] || {
    total: 0,
    correct: 0
  };

  return {
    cardNo: targetNo,
    total: stats.total,
    correct: stats.correct,
    accuracy: stats.total ? stats.correct / stats.total : 0
  };
}

function getFlashDailyStatus() {
  return getFlashStatsBundle().daily;
}
