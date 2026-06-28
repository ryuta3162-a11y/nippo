/**
 * JOYFIT24経堂 — 日報メール
 *
 * createShopDailyReportDraft … 手動：下書き作成（シートのボタン用）
 * sendShopDailyReportSilent_ … 自動送信（トリガーから）
 * sendDailyReportAt20_ / sendDailyReportAt21_ … 土日祝20時 / 平日21時
 */

const NIPPO_MAIL_RECIPIENTS = [
  "m-harada@okamoto-group.co.jp",
  "mito-sato@okamoto-group.co.jp",
  "r-kusaka@okamoto-group.co.jp",
  "h-nakata@okamoto-group.co.jp",
  "s-kurokawa@okamoto-group.co.jp",
  "ka-yoshida@okamoto-group.co.jp",
  "m-tokushige@okamoto-group.co.jp",
  "k-moriyasu@okamoto-group.co.jp",
  "k-ishibashi@okamoto-group.co.jp",
  "m-osari@okamoto-group.co.jp",
  "yuka-hachiya@okamoto-group.co.jp"
];

const NIPPO_MAIL_FROM = "jf-kyoudou@okamoto-group.co.jp";
const NIPPO_SHEET_NAME = "日報";
const NIPPO_RANGE_START_ROW = 8;
const NIPPO_RANGE_START_COL = 2;
const NIPPO_RANGE_ROWS = 33;
const NIPPO_RANGE_COLS = 8;

const JAPAN_HOLIDAY_CALENDAR_ID = "ja.japanese#holiday@group.v8.calendar.google.com";
const SENT_REPORT_PROP_PREFIX = "dailyReportSent_";

/** 手動：下書き作成 */
function createShopDailyReportDraft() {
  const mail = buildShopDailyReportMail_();
  if (!mail) return;
  GmailApp.createDraft(mail.recipient, mail.subject, "", mail.options);
  Logger.log("店舗日報の下書きを作成しました: " + mail.subject);
}

/** 土日祝 20:00 トリガー用 */
function sendDailyReportAt20_() {
  const now = new Date();
  if (!isWeekendOrJapaneseHoliday_(now)) return;
  sendShopDailyReportSilent_();
}

/** 平日 21:00 トリガー用 */
function sendDailyReportAt21_() {
  const now = new Date();
  if (isWeekendOrJapaneseHoliday_(now)) return;
  sendShopDailyReportSilent_();
}

/** 自動送信（ダイアログなし・同日二重送信防止） */
function sendShopDailyReportSilent_() {
  const tz = Session.getScriptTimeZone();
  const todayKey = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  if (wasDailyReportSentToday_(todayKey)) {
    Logger.log("日報メールは本日送信済みのためスキップ: " + todayKey);
    return;
  }

  const mail = buildShopDailyReportMail_();
  if (!mail) return;

  GmailApp.sendEmail(mail.recipient, mail.subject, "", mail.options);
  markDailyReportSentToday_(todayKey);
  Logger.log("店舗日報を送信しました: " + mail.subject);
}

function buildShopDailyReportMail_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log("スプレッドシートを取得できませんでした");
    return null;
  }

  const sheet = ss.getSheetByName(NIPPO_SHEET_NAME);
  if (!sheet) {
    Logger.log("「日報」シートが見つかりません");
    return null;
  }

  const startRow = NIPPO_RANGE_START_ROW;
  const startCol = NIPPO_RANGE_START_COL;
  const numRows = NIPPO_RANGE_ROWS;
  const numCols = NIPPO_RANGE_COLS;

  const formattedDate = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), "yyyy.MM.dd"
  );
  const subject = "JF24経堂　報連相 " + formattedDate;

  const range = sheet.getRange(startRow, startCol, numRows, numCols);
  const values = range.getDisplayValues();
  const fontColors = range.getFontColors();
  const backgrounds = range.getBackgrounds();
  const hAligns = range.getHorizontalAlignments();
  const vAligns = range.getVerticalAlignments();

  const mergeMap = buildMergeMap_(range, startRow, startCol, numRows, numCols);
  const htmlBody = buildNippoHtmlBody_(values, fontColors, backgrounds, hAligns, vAligns, mergeMap, numRows, numCols);

  const options = { htmlBody: htmlBody, cc: "" };
  if (NIPPO_MAIL_FROM) options.from = NIPPO_MAIL_FROM;

  return {
    recipient: NIPPO_MAIL_RECIPIENTS.join(","),
    subject: subject,
    options: options
  };
}

function buildMergeMap_(range, startRow, startCol, numRows, numCols) {
  const mergeMap = [];
  for (let r = 0; r < numRows; r++) {
    mergeMap[r] = [];
    for (let c = 0; c < numCols; c++) {
      mergeMap[r][c] = { skip: false, rowSpan: 1, colSpan: 1 };
    }
  }

  const merges = range.getMergedRanges();
  for (let i = 0; i < merges.length; i++) {
    const rng = merges[i];
    const relativeRow = rng.getRow() - startRow;
    const relativeCol = rng.getColumn() - startCol;
    const mNumRows = rng.getNumRows();
    const mNumCols = rng.getNumColumns();

    if (relativeRow < 0 || relativeRow >= numRows || relativeCol < 0 || relativeCol >= numCols) continue;

    mergeMap[relativeRow][relativeCol].rowSpan = mNumRows;
    mergeMap[relativeRow][relativeCol].colSpan = mNumCols;
    for (let r = 0; r < mNumRows; r++) {
      for (let c = 0; c < mNumCols; c++) {
        if (r === 0 && c === 0) continue;
        if (relativeRow + r < numRows && relativeCol + c < numCols) {
          mergeMap[relativeRow + r][relativeCol + c].skip = true;
        }
      }
    }
  }
  return mergeMap;
}

function buildNippoHtmlBody_(values, fontColors, backgrounds, hAligns, vAligns, mergeMap, numRows, numCols) {
  let htmlBody = [
    '<div style="font-family: sans-serif; font-size: 10pt; color: #333;">',
    "<p>お元気様です。<br>本日の業務日報です。<br>よろしくお願いします。</p>",
    '<table border="0" cellspacing="0" cellpadding="0" style="',
    "border-collapse: collapse; border: none; font-size: 10pt;",
    "line-height: 1.3; font-weight: bold; margin-bottom: 20px;\">"
  ].join("");

  for (let i = 0; i < numRows; i++) {
    htmlBody += "<tr>";
    for (let j = 0; j < numCols; j++) {
      if (mergeMap[i][j].skip) continue;

      let val = values[i][j];
      if (val === "") val = "&nbsp;";

      let bg = backgrounds[i][j];
      let color = fontColors[i][j];
      let hAlign = hAligns[i][j];
      const vAlign = vAligns[i][j];
      if (j === 0) hAlign = "left";
      if (i === 0 && bg === "#ffffff") {
        bg = "#444444";
        color = "#ffffff";
      }

      let spanAttr = "";
      if (mergeMap[i][j].rowSpan > 1) spanAttr += ' rowspan="' + mergeMap[i][j].rowSpan + '"';
      if (mergeMap[i][j].colSpan > 1) spanAttr += ' colspan="' + mergeMap[i][j].colSpan + '"';

      htmlBody += "<td" + spanAttr + ' style="border: none; padding: 4px 6px; background-color: ' + bg +
        "; color: " + color + "; font-weight: bold; text-align: " + hAlign +
        "; vertical-align: " + vAlign + '; font-size: 10pt; white-space: nowrap;">' + val + "</td>";
    }
    htmlBody += "</tr>";
  }

  htmlBody += [
    "</table><br><div>",
    "☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆<br>",
    "株式会社ヤマウチ<br>",
    "スポーツクラブJOYFIT24経堂<br>",
    "〒156-0052<br>",
    "東京都世田谷区経堂5-23-13<br>",
    "TEL：03-6804-4100<br>",
    "スタッフ常駐時間：平日　10：00～21：00<br>",
    "　　　　　　　　　土日祝　12：00～20：00<br>",
    "※毎週月曜日・木曜日は終日スタッフ不在でございます。<br>",
    "☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆〇☆<br>",
    "</div></div>"
  ].join("");

  return htmlBody;
}

function isWeekendOrJapaneseHoliday_(date) {
  const day = date.getDay();
  if (day === 0 || day === 6) return true;
  return isJapanesePublicHoliday_(date);
}

function isJapanesePublicHoliday_(date) {
  try {
    const cal = CalendarApp.getCalendarById(JAPAN_HOLIDAY_CALENDAR_ID);
    if (!cal) return false;
    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const end = new Date(start.getTime() + 86400000);
    return cal.getEvents(start, end).length > 0;
  } catch (err) {
    Logger.log("祝日カレンダー参照エラー（土日のみ判定）: " + err.message);
    return false;
  }
}

function wasDailyReportSentToday_(todayKey) {
  return PropertiesService.getScriptProperties().getProperty(SENT_REPORT_PROP_PREFIX + todayKey) === "1";
}

function markDailyReportSentToday_(todayKey) {
  PropertiesService.getScriptProperties().setProperty(SENT_REPORT_PROP_PREFIX + todayKey, "1");
}

/** 日報送信トリガーを installOptionDailyTrigger から呼ぶ */
function installDailyReportSendTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === "sendDailyReportAt20_" || fn === "sendDailyReportAt21_") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("sendDailyReportAt20_")
    .timeBased()
    .everyDays(1)
    .atHour(20)
    .nearMinute(0)
    .create();

  ScriptApp.newTrigger("sendDailyReportAt21_")
    .timeBased()
    .everyDays(1)
    .atHour(21)
    .nearMinute(0)
    .create();
}

function removeDailyReportSendTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === "sendDailyReportAt20_" || fn === "sendDailyReportAt21_") {
      ScriptApp.deleteTrigger(t);
    }
  });
}
