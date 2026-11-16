// Сайт @ii_vedet на Google Apps Script: страница + форма заданий + Google Таблица с ответами.
const RUBRICS = ['Деньги', 'Работа', 'Быт', 'Путешествия', 'Сложный разговор'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ИИ ведёт: дай задание')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Запустите один раз вручную: создаст таблицу и покажет ссылку на неё в «Журнале выполнения».
function setup() {
  Logger.log('Таблица с заданиями: ' + getSheet_().getParent().getUrl());
}

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id).getSheets()[0];
  const ss = SpreadsheetApp.create('Задания для @ii_vedet');
  props.setProperty('SHEET_ID', ss.getId());
  const sh = ss.getSheets()[0];
  sh.appendRow(['Время', 'Ник в TikTok', 'Рубрика', 'Задание', 'Можно показать ник']);
  sh.setFrozenRows(1);
  return sh;
}

function clean_(s, max) {
  s = String(s || '').replace(/\s+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s; // защита от формул в таблице
}

// Вызывается со страницы через google.script.run
function submitTask(data) {
  if (!data || data.website) return { ok: true };            // скрытое поле-ловушка для ботов
  const task = clean_(data.task, 600);
  if (task.length < 10) return { ok: false, error: 'Опиши задание подробнее (минимум 10 символов).' };
  const rubric = RUBRICS.indexOf(data.rubric) >= 0 ? data.rubric : 'Другое';
  const lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    getSheet_().appendRow([new Date(), clean_(data.nick, 60), rubric, task, data.showNick ? 'да' : 'нет']);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}
