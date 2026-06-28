/**
 * JOYFIT24経堂 — オプション契約メール集計（メイン）
 *
 * スプレッドシート:
 * https://docs.google.com/spreadsheets/d/14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w/
 *
 * 【自動更新】毎日19:30 → smartUpdateDataSilent（OP・入会退会・数値反映）
 * 【自動送信】土日祝20:00 / 平日21:00 → sendDailyReportAt20_ / sendDailyReportAt21_
 * 【手動メニュー】JOYFIT → 入会・退会を更新する（入会・退会.gs）
 */

const OPTION_SPREADSHEET_ID = "14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w";

const SEARCH_QUERY =
  'subject:("【JOYFIT24経堂】オプションご契約につきまして" OR "【JOYFIT24経堂】ご入会ありがとうございます")';

const SHEET_NAME_LOG = "OPデータ";
const SHEET_NAME_SUMMARY = "集計";

const OP_LOG_HEADERS = [
  "受信日時", "氏名", "区分",
  "オプション名(メール記載)", "オプション名(集計用)", "メールID"
];

const OPTION_LIST = [
  "安心サポート", "安心サポートVIP", "水素水", "オンラインレッスン",
  "体組成計", "契約ロッカー1,500", "レンタルマット", "プロテイン12杯",
  "プロテイン無制限", "プロテイン＋水素水", "レンタルタオル", "タンニング",
  "セルフエステ", "ホットスタジオ", "ヨガロッカー", "ピラティスリフォーマー"
];

// ── メニュー・イベント ──

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("JOYFIT")
    .addItem("入会・退会を更新", "updateMembershipSheet")
    .addSeparator()
    .addItem("6ヶ月割メールを取込", "kyodoImportEnrollmentSmart_")
    .addItem("キャンペーン対象を整理", "kyodoPrepareCampaignTargets_")
    .addItem("確認メールを送る", "kyodoSendCampaignMailMenu_")
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  if (handleTaikaiSheetEdit_(e)) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() === SHEET_NAME_SUMMARY && e.range.getA1Notation() === "B1") {
    const ss = e.source;
    const logSheet = ss.getSheetByName(SHEET_NAME_LOG);
    if (!logSheet) return;
    const ym = resolveSummaryTargetYearMonth_(ss);
    syncSummarySheetFromOpData_(ss, ym.year, ym.month, logSheet);
  }
}

// ── トリガー（初回のみ installOptionDailyTrigger を実行）──

function installOptionDailyTrigger() {
  removeOptionDailyTriggers_();
  removeDailyReportSendTriggers_();

  ScriptApp.newTrigger("smartUpdateDataSilent")
    .timeBased()
    .everyDays(1)
    .atHour(19)
    .nearMinute(30)
    .create();

  installDailyReportSendTriggers_();

  SpreadsheetApp.getUi().alert(
    "トリガー設定",
    "毎日の自動処理を設定しました。\n\n" +
      "19:30 … OP・集計・日報数値・入会退会を更新\n" +
      "20:00 … 日報メール送信（土日祝）\n" +
      "21:00 … 日報メール送信（平日）\n\n" +
      "※送信は数値更新のあとです。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function removeOptionDailyTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === "smartUpdateDataSilent" || fn === "smartUpdateData" ||
        fn === "refreshNippoCurrentMonthFromLog") {
      ScriptApp.deleteTrigger(t);
    }
  });
  removeDailyReportSendTriggers_();
}

/** 時間トリガーから呼ばれる（ダイアログなし） */
function smartUpdateDataSilent() {
  smartUpdateDataCore_(true);
}

/** 手動実行用（ダイアログあり） */
function smartUpdateData() {
  smartUpdateDataCore_(false);
}

function smartUpdateDataCore_(silent) {
  setupSpreadsheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAME_LOG);

  if (logSheet.getLastRow() <= 1) {
    if (silent) {
      executeFetchAll_(ss, logSheet, true);
    } else {
      const ui = SpreadsheetApp.getUi();
      const res = ui.alert(
        "初回セットアップ",
        "OPデータが空です。過去すべてのメールを一括取得しますか？\n（はい推奨・数分かかる場合あり）",
        ui.ButtonSet.YES_NO
      );
      if (res === ui.Button.YES) {
        executeFetchAll_(ss, logSheet, false);
      } else {
        executeFetchMonth_(ss, logSheet, false);
      }
    }
    return;
  }

  executeFetchMonth_(ss, logSheet, silent);
}

// ── OPシートセットアップ ──

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let logSheet = ss.getSheetByName(SHEET_NAME_LOG);
  if (!logSheet) logSheet = ss.insertSheet(SHEET_NAME_LOG);
  logSheet.getRange(1, 1, 1, OP_LOG_HEADERS.length).setValues([OP_LOG_HEADERS]);
  logSheet.getRange(1, 1, 1, OP_LOG_HEADERS.length).setFontWeight("bold").setBackground("#f3f3f3");
  logSheet.setFrozenRows(1);

  let summarySheet = ss.getSheetByName(SHEET_NAME_SUMMARY);
  if (!summarySheet) summarySheet = ss.insertSheet(SHEET_NAME_SUMMARY);

  summarySheet.getRange("1:2").clear();
  summarySheet.getRange("A1").setValue("対象月を選択 ➡").setFontWeight("bold").setHorizontalAlignment("right");

  const monthCell = summarySheet.getRange("B1");
  monthCell.clearDataValidations();
  monthCell.clearFormat();
  monthCell.setNumberFormat("@");

  const monthList = buildOpMonthList_(ss);
  if (monthList.length > 0) {
    monthCell.setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(monthList, true).build()
    );
    const val = String(monthCell.getValue() || "").trim();
    if (!monthList.includes(val)) monthCell.setValue(monthList[0]);
  } else {
    const today = new Date();
    monthCell.setValue(today.getFullYear() + "年" + (today.getMonth() + 1) + "月");
  }
  monthCell.setBackground("#fff2cc").setFontWeight("bold").setHorizontalAlignment("center");

  summarySheet.getRange(2, 1, 1, 6).setValues([[
    "オプション名", "利用開始(新規入会)", "利用開始(OP追加)",
    "利用開始合計", "利用停止数", "翌月の±"
  ]]);
  summarySheet.getRange("A2:F2").setFontWeight("bold").setBackground("#e2efda");
  summarySheet.setFrozenRows(2);

  const summaryNames = OPTION_LIST.map(function (n) { return [n, "", "", "", "", ""]; });
  summarySheet.getRange(3, 1, summaryNames.length, 6).setValues(summaryNames);
  ensureSummaryFormulas_(summarySheet);
}

/** 集計B1: OPデータにある月 ＋ 当月（データがなくても当月は選べる） */
function buildOpMonthList_(ss) {
  const keys = {};
  const logSheet = ss.getSheetByName(SHEET_NAME_LOG);
  if (logSheet && logSheet.getLastRow() > 1) {
    const data = logSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const d = parseOpLogDate_(data[i][0]);
      if (!d) continue;
      const label = d.getFullYear() + "年" + (d.getMonth() + 1) + "月";
      keys[label] = d.getFullYear() * 100 + (d.getMonth() + 1);
    }
  }
  const today = new Date();
  const cur = today.getFullYear() + "年" + (today.getMonth() + 1) + "月";
  keys[cur] = today.getFullYear() * 100 + (today.getMonth() + 1);

  return Object.keys(keys).sort(function (a, b) { return keys[b] - keys[a]; });
}

// ── OPメール取得 ──

function executeFetchMonth_(ss, logSheet, silent) {
  const summarySheet = ss.getSheetByName(SHEET_NAME_SUMMARY);
  let targetStr = String(summarySheet.getRange("B1").getValue() || "").trim();
  let match = targetStr.match(/^(\d{4})年(\d{1,2})月$/);
  if (!match) {
    const today = new Date();
    targetStr = today.getFullYear() + "年" + (today.getMonth() + 1) + "月";
    summarySheet.getRange("B1").setValue(targetStr);
    match = [null, today.getFullYear(), today.getMonth() + 1];
  }

  const targetYear = parseInt(match[1], 10);
  const targetMonth = parseInt(match[2], 10) - 1;
  const targetMonthStr = String(targetMonth + 1).padStart(2, "0");

  const logData = logSheet.getDataRange().getValues();
  const header = logData.shift();
  const retained = logData.filter(function (row) {
    const d = parseOpLogDate_(row[0]);
    if (!d) return true;
    return !(d.getFullYear() === targetYear && d.getMonth() === targetMonth);
  });

  logSheet.clearContents();
  logSheet.appendRow(header);
  if (retained.length > 0) {
    logSheet.getRange(2, 1, retained.length, retained[0].length).setValues(retained);
  }

  const firstDay = targetYear + "/" + targetMonthStr + "/01";
  const nextMonth = new Date(targetYear, targetMonth + 1, 1);
  const nextFirst = nextMonth.getFullYear() + "/" + String(nextMonth.getMonth() + 1).padStart(2, "0") + "/01";

  const threads = searchGmailAllThreads_(SEARCH_QUERY + " after:" + firstDay + " before:" + nextFirst);
  const newData = extractDataFromThreads(threads, targetYear, targetMonth);

  if (newData.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, newData.length, newData[0].length).setValues(newData);
  }

  syncSummarySheetFromOpData_(ss, targetYear, targetMonth, logSheet);

  const nippoYm = resolveNippoTargetYearMonth_(ss);
  const nippoUpdated = nippoYm.year === targetYear && nippoYm.month === targetMonth;
  if (nippoUpdated) {
    updateNippoSheetForMonth(ss, targetYear, targetMonth, logSheet);
  }

  try {
    fetchMembershipEmailsSilent();
  } catch (err) {
    Logger.log("入会・退会取得エラー: " + err.message);
  }

  try {
    syncMembershipDailyCountsSilent_();
  } catch (err) {
    Logger.log("日別入会・退会の反映エラー: " + err.message);
  }

  if (!silent) {
    let msg = (targetMonth + 1) + "月分を取り込み、集計を更新しました。";
    msg += nippoUpdated
      ? "\n日報C21:C36も更新しました。"
      : "\n日報は未更新（日報B1の月が集計B1と異なります）。";
    SpreadsheetApp.getUi().alert("OP更新", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function executeFetchAll_(ss, logSheet, silent) {
  logSheet.clearContents();
  logSheet.getRange(1, 1, 1, OP_LOG_HEADERS.length).setValues([OP_LOG_HEADERS]);

  const threads = searchGmailAllThreads_(SEARCH_QUERY);
  const newData = extractDataFromThreads(threads, null, null);

  if (newData.length > 0) {
    newData.sort(function (a, b) { return a[0].getTime() - b[0].getTime(); });
    logSheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
  }

  setupSpreadsheet();

  const summaryYm = resolveSummaryTargetYearMonth_(ss);
  syncSummarySheetFromOpData_(ss, summaryYm.year, summaryYm.month, logSheet);

  const nippoYm = resolveNippoTargetYearMonth_(ss);
  updateNippoSheetForMonth(ss, nippoYm.year, nippoYm.month, logSheet);

  try {
    fetchMembershipEmailsSilent();
  } catch (err) {
    Logger.log("入会・退会取得エラー: " + err.message);
  }

  try {
    syncMembershipDailyCountsSilent_();
  } catch (err) {
    Logger.log("日別入会・退会の反映エラー: " + err.message);
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      "OP一括取得",
      "合計 " + newData.length + " 件を取得しました。",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/** 日報ボタン用: Gmail取得なしで日報C列だけ更新 */
function refreshNippoCurrentMonthFromLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAME_LOG);
  if (!logSheet || logSheet.getLastRow() < 2) return;

  const ym = resolveNippoTargetYearMonth_(ss);
  const summaryYm = resolveSummaryTargetYearMonth_(ss);
  if (summaryYm.year === ym.year && summaryYm.month === ym.month) {
    syncSummarySheetFromOpData_(ss, summaryYm.year, summaryYm.month, logSheet);
  }
  updateNippoSheetForMonth(ss, ym.year, ym.month, logSheet);
}

// ── Gmail共通 ──

function searchGmailAllThreads_(query) {
  const out = [];
  let start = 0;
  while (start < 500) {
    const batch = GmailApp.search(query, start, 100);
    if (!batch.length) break;
    for (let i = 0; i < batch.length; i++) out.push(batch[i]);
    start += batch.length;
    if (batch.length < 100) break;
  }
  return out;
}

function getMessageBodyText_(message) {
  const plain = message.getPlainBody();
  if (plain && String(plain).trim()) return String(plain);
  const html = message.getBody();
  if (!html) return "";
  return String(html)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .trim();
}

// ── OPメール解析 ──

function extractDataFromThreads(threads, targetYear, targetMonth) {
  const newData = [];
  const seen = {};

  for (let t = 0; t < threads.length; t++) {
    const messages = threads[t].getMessages();
    for (let m = 0; m < messages.length; m++) {
      const message = messages[m];
      const date = message.getDate();
      if (targetYear !== null && targetMonth !== null) {
        if (date.getFullYear() !== targetYear || date.getMonth() !== targetMonth) continue;
      }

      const subject = String(message.getSubject() || "");
      const body = getMessageBodyText_(message);
      const msgId = message.getId();

      try {
        let parsed = [];
        if (subject.indexOf("オプションご契約につきまして") !== -1) {
          parsed = parseOptionContractEmail(date, body, msgId);
        } else if (subject.indexOf("ご入会ありがとうございます") !== -1) {
          parsed = parseSignupEmail(date, body, msgId);
        }
        for (let i = 0; i < parsed.length; i++) {
          const row = padLogDataRow_(parsed[i], msgId);
          const key = row[5] + "|" + row[2] + "|" + row[4];
          if (seen[key]) continue;
          seen[key] = true;
          newData.push(row);
        }
      } catch (e) {
        Logger.log("OP解析エラー: " + subject + " / " + e.message);
      }
    }
  }
  return newData;
}

function padLogDataRow_(row, msgId) {
  const r = row.slice();
  while (r.length < 5) r.push("");
  return [r[0], r[1], r[2], r[3], normalizeOptionName(String(r[4] || r[3] || "")), msgId || r[5] || ""];
}

function parseSignupEmail(date, body, msgId) {
  const name = extractSignupName_(body);
  const breakdown = parseSignupBreakdown_(date, body, name, msgId);
  const got = {};
  for (let i = 0; i < breakdown.length; i++) got[breakdown[i][4]] = true;
  return breakdown.concat(parseSignupLooseLines_(date, body, name, msgId, got));
}

function extractSignupName_(body) {
  const m1 = body.match(/お名前\s*[:：]\s*(.+?)\s*様/);
  if (m1) return String(m1[1]).replace(/^[>\s]+/, "").trim();
  const lines = body.split(/\r?\n/);
  for (let i = 0; i < Math.min(lines.length, 12); i++) {
    const line = String(lines[i] || "").trim();
    if (!line || line.indexOf("JOYFIT") !== -1) continue;
    const m2 = line.match(/^(.{1,30}?)\s*様\s*$/);
    if (m2) return String(m2[1]).replace(/^[>\s]+/, "").trim();
  }
  return "";
}

function isBreakdownNoiseLine_(line) {
  const s = String(line || "").trim();
  if (!s || s.indexOf("----") !== -1) return true;
  if (s.indexOf("小計") !== -1 || s.indexOf("合計") !== -1) return true;
  if (s.indexOf("初期費用") !== -1 || s.indexOf("ナショナル会員") !== -1) return true;
  return false;
}

function parseSignupBreakdown_(date, body, name, msgId) {
  const results = [];
  const lines = body.split(/\r?\n/);
  let inBreakdown = false;
  const got = {};
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.indexOf("月会費の内訳") !== -1) { inBreakdown = true; continue; }
    if (inBreakdown && (line.indexOf("APP登録方法") !== -1 || line.indexOf("ご契約中のオプション") !== -1)) break;
    if (!inBreakdown || isBreakdownNoiseLine_(line) || line.indexOf("円") === -1) continue;
    const norm = normalizeOptionName(line);
    if (!OPTION_LIST.includes(norm) || got[norm]) continue;
    got[norm] = true;
    results.push([date, name, "利用開始(新規入会)",
      line.replace(/\(\d+月分\).*/, "").replace(/（\d+月分）.*/, "").trim(), norm, msgId]);
  }
  return results;
}

function parseSignupLooseLines_(date, body, name, msgId, alreadyGot) {
  const results = [];
  const got = alreadyGot || {};
  const lines = body.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = String(lines[i] || "").trim();
    if (!line || line.indexOf("月分") === -1 || isBreakdownNoiseLine_(line)) continue;
    const norm = normalizeOptionName(line);
    if (!OPTION_LIST.includes(norm) || got[norm]) continue;
    got[norm] = true;
    results.push([date, name, "利用開始(新規入会)", line, norm, msgId]);
  }
  return results;
}

function parseOptionContractEmail(date, body, msgId) {
  const results = [];
  const nameMatch = body.match(/([^\r\n]{1,40}?)\s*様/);
  const name = nameMatch ? String(nameMatch[1]).replace(/^[>\s]+/, "").trim() : "";
  const re = /[（(](利用開始|利用停止)[）)]\s*([^\r\n]+)/g;
  let m;
  while ((m = re.exec(body)) !== null) {
    const status = m[1] === "利用開始" ? "利用開始(OP追加)" : m[1];
    const raw = String(m[2]).trim();
    if (!raw) continue;
    results.push([date, name, status, raw, normalizeOptionName(raw), msgId]);
  }
  return results;
}

function normalizeOptionName(rawName) {
  const s = String(rawName || "");
  if (s.includes("水素水") && s.includes("プロテイン")) return "プロテイン＋水素水";
  if (s.includes("VIP") && (s.includes("あんしん") || s.includes("安心"))) return "安心サポートVIP";
  if (s.includes("安心サポート") || s.includes("あんしんサポート")) return "安心サポート";
  if (s.includes("ボディプランナー") || s.includes("ボディープランナー")) return "体組成計";
  if (s.includes("マットレンタル") || s.includes("レンタルマット")) return "レンタルマット";
  if (s.includes("ピラティス")) return "ピラティスリフォーマー";
  for (let i = 0; i < OPTION_LIST.length; i++) {
    if (s.includes(OPTION_LIST[i])) return OPTION_LIST[i];
  }
  return s;
}

// ── 集計・日報 ──

function parseOpLogDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "number" && v > 30000) {
    const base = new Date(1899, 11, 30);
    const d = new Date(base.getTime() + Math.floor(v) * 86400000);
    return isNaN(d.getTime()) ? null : d;
  }
  const s = String(v || "").trim();
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function resolveNippoTargetYearMonth_(ss) {
  const nippo = ss.getSheetByName("日報");
  if (nippo) {
    const b1 = String(nippo.getRange("B1").getDisplayValue() || nippo.getRange("B1").getValue() || "").trim();
    const yyMM = b1.match(/^(\d{2})(\d{2})$/);
    if (yyMM) {
      const y = 2000 + parseInt(yyMM[1], 10);
      const m = parseInt(yyMM[2], 10) - 1;
      if (m >= 0 && m <= 11) return { year: y, month: m, label: y + "年" + (m + 1) + "月" };
    }
  }
  return resolveSummaryTargetYearMonth_(ss);
}

function resolveSummaryTargetYearMonth_(ss) {
  const summary = ss.getSheetByName(SHEET_NAME_SUMMARY);
  if (summary) {
    const s = String(summary.getRange("B1").getDisplayValue() || summary.getRange("B1").getValue() || "").trim();
    const ym = s.match(/^(\d{4})年(\d{1,2})月$/);
    if (ym) return { year: parseInt(ym[1], 10), month: parseInt(ym[2], 10) - 1, label: s };
  }
  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth(), label: now.getFullYear() + "年" + (now.getMonth() + 1) + "月" };
}

function countOpDataForMonth_(logSheet, targetYear, targetMonth) {
  const data = logSheet.getDataRange().getValues();
  const counts = OPTION_LIST.map(function () { return { newSignup: 0, opAdd: 0, stop: 0 }; });
  const idxByOpt = {};
  for (let o = 0; o < OPTION_LIST.length; o++) idxByOpt[OPTION_LIST[o]] = o;

  for (let i = 1; i < data.length; i++) {
    const d = parseOpLogDate_(data[i][0]);
    if (!d || d.getFullYear() !== targetYear || d.getMonth() !== targetMonth) continue;
    const cat = String(data[i][2] || "").trim();
    const opt = String(data[i][4] || "").trim();
    const idx = idxByOpt[opt];
    if (idx === undefined) continue;
    if (cat === "利用開始(新規入会)") counts[idx].newSignup++;
    else if (cat === "利用開始(OP追加)") counts[idx].opAdd++;
    else if (cat === "利用停止") counts[idx].stop++;
  }
  return { counts: counts };
}

function syncSummarySheetFromOpData_(ss, targetYear, targetMonth, logSheet) {
  const summarySheet = ss.getSheetByName(SHEET_NAME_SUMMARY);
  if (!summarySheet) return;
  SpreadsheetApp.flush();
  const result = countOpDataForMonth_(logSheet, targetYear, targetMonth);
  const bCol = [], cCol = [], dCol = [], eCol = [];
  for (let i = 0; i < OPTION_LIST.length; i++) {
    const start = result.counts[i].newSignup + result.counts[i].opAdd;
    bCol.push([result.counts[i].newSignup]);
    cCol.push([result.counts[i].opAdd]);
    dCol.push([start]);
    eCol.push([result.counts[i].stop]);
  }
  summarySheet.getRange(3, 2, OPTION_LIST.length, 1).setValues(bCol);
  summarySheet.getRange(3, 3, OPTION_LIST.length, 1).setValues(cCol);
  summarySheet.getRange(3, 4, OPTION_LIST.length, 1).setValues(dCol);
  summarySheet.getRange(3, 5, OPTION_LIST.length, 1).setValues(eCol);
  ensureSummaryFormulas_(summarySheet);
}

function ensureSummaryFormulas_(summarySheet) {
  for (let i = 0; i < OPTION_LIST.length; i++) {
    summarySheet.getRange(i + 3, 6).setFormula("=D" + (i + 3) + "-E" + (i + 3));
  }
}

function updateNippoSheetForMonth(ss, targetYear, targetMonth, logSheet) {
  const nippoSheet = ss.getSheetByName("日報");
  if (!nippoSheet) return;
  const result = countOpDataForMonth_(logSheet, targetYear, targetMonth);
  const cCol = result.counts.map(function (c) { return [c.newSignup + c.opAdd]; });
  nippoSheet.getRange(21, 3, OPTION_LIST.length, 1).setValues(cCol);
}
