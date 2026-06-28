/**
 * JOYFIT24経堂 — 入会・退会メール集計
 *
 * 【入会・退会】     見る用（A〜D=入会, F〜J=退会, 3行目から）
 * 【入会・退会_データ】非表示（全メール保管）
 *
 * 入会区分: 6ヶ月割 / 法人会員
 * 退会の J列「退会キャンセル」ON → 退会数から除外（K列にメールID・非表示）
 */

const SHEET_NAME_MEMBERS = "入会・退会";
const SHEET_NAME_DATA = "入会・退会_データ";

const DISPLAY_HEADER_ROW = 2;
const DISPLAY_START_ROW = 3;
const DISPLAY_COLS = 11;

const DATA_ENROLL_COLS = 5;
const DATA_WITHDRAW_COLS = 6;
const DISPLAY_MAX_SCAN = 300;

const DATA_HEADER_ROW = 1;
const DATA_START_ROW = 2;
const DATA_ENROLL_COL = 1;
const DATA_WITHDRAW_COL = 7;

const NYUKAI_SEARCH_QUERY =
  'from:info@joyfit-service.jp subject:(ご入会ありがとうございます)';
const TAIKAI_SEARCH_QUERY =
  'from:info@joyfit-service.jp subject:("ご退会のお手続きについて" OR "【JOYFIT24経堂】ご退会のお手続きについて")';

const LABEL_NYUKAI_GENERAL = "入会メール/一般会員";
const LABEL_NYUKAI_CORPORATE = "入会メール/法人会員";
const LABEL_TAIKAI_GENERAL = "退会メール/一般会員";
const LABEL_TAIKAI_CORPORATE = "退会メール/法人会員";

const ENROLL_HEADERS = ["タイムスタンプ", "氏名", "入会月", "区分", "メールID"];
const WITHDRAW_HEADERS = ["タイムスタンプ", "氏名", "退会月", "区分", "メールID", "退会キャンセル"];
/** 入会・退会シート退会側の表示ヘッダー（F〜J列。メールIDはK列非表示） */
const DISPLAY_WITHDRAW_HEADERS = ["タイムスタンプ", "氏名", "退会月", "区分", "退会キャンセル"];

const CATEGORY_SIX_MONTH = "6ヶ月割";
const CATEGORY_CORPORATE = "法人会員";

/** 月次日報シート（例: 2606）の日別入力列（値のみ書き込み・書式は変更しない） */
const DAILY_ENROLL_GENERAL_COL = 5;    // E列 … 6ヶ月割入会
const DAILY_ENROLL_CORPORATE_COL = 7;  // G列 … 法人会員入会
const DAILY_WITHDRAW_COL = 5;          // E列 … 退会
const DAILY_ENROLL_ROW_START = 5;    // 1日 → 5行目、31日 → 35行目
const DAILY_WITHDRAW_ROW_START = 40;   // 1日 → 40行目、31日 → 70行目
const DAILY_COUNT_DAYS = 31;

// ── メニュー ──

function updateMembershipSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  setupDataSheet_(ss);
  migrateLegacyLayouts_(ss);
  setupDisplaySheet_(ss);
  installMembershipLabelsSilent_();
  repairMonthValuesInData_(ss);
  dedupeStoredMemberData_(ss);

  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  const enrollCount = readEnrollData_(dataSheet).length;
  const withdrawCount = readWithdrawData_(dataSheet).length;
  const isEmpty = enrollCount === 0 && withdrawCount === 0;

  let enrollAdded = 0;
  let withdrawAdded = 0;
  let enrollTotal = enrollCount;
  let withdrawTotal = withdrawCount;

  if (isEmpty) {
    const res = ui.alert(
      "初回取得",
      "まだデータがありません。\n過去のメールをすべて取得します（数分かかる場合があります）。\n\n実行しますか？",
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) {
      refreshMembershipDisplay_(ss);
      return;
    }
    const full = fetchAllMembershipEmailsCore_(ss);
    enrollTotal = full.enrollTotal;
    withdrawTotal = full.withdrawTotal;
  } else {
    const inc = fetchMembershipEmailsIncremental_(ss, 2);
    enrollAdded = inc.enrollAdded;
    withdrawAdded = inc.withdrawAdded;
    enrollTotal = readEnrollData_(dataSheet).length;
    withdrawTotal = readWithdrawData_(dataSheet).length;
  }

  const view = refreshMembershipDisplay_(ss);

  try {
    syncMembershipDailyCountsSilent_();
  } catch (err) {
    Logger.log("日別入会・退会の反映エラー: " + err.message);
  }

  ui.alert(
    "入会・退会を更新",
    "完了しました。\n\n" +
      (isEmpty
        ? "取得 … 入会 " + enrollTotal + " 件 / 退会 " + withdrawTotal + " 件（初回・全件）\n"
        : "新規 … 入会 +" + enrollAdded + " 件 / 退会 +" + withdrawAdded + " 件（直近2か月を確認）\n" +
          "合計 … 入会 " + enrollTotal + " 件 / 退会 " + withdrawTotal + " 件\n") +
      "表示 … 入会 " + view.enrollLabel + "：" + view.enrollCount + " 件 / " +
      "退会 " + view.withdrawLabel + "：" + view.withdrawCount + " 件" +
      (view.withdrawCancelCount ? "（キャンセル " + view.withdrawCancelCount + " 件除く）" : "") + "\n\n" +
      "※OP・日報は毎日19:30に自動更新されます。\n" +
      "※当月シート（例: 6月なら2606）の日別欄にも反映されます。",
    ui.ButtonSet.OK
  );
}

/** 過去メールをすべて取得（初回のみ） */
function fetchAllMembershipEmailsCore_(ss) {
  clearAllMemberData_(ss);

  const enrollThreads = searchGmailAllThreads_(NYUKAI_SEARCH_QUERY);
  const withdrawThreads = searchGmailAllThreads_(TAIKAI_SEARCH_QUERY);
  const enrollRows = extractNyukaiRowsFromThreads_(enrollThreads);
  const withdrawRows = extractTaikaiRowsFromThreads_(withdrawThreads);

  writeEnrollData_(ss, enrollRows);
  writeWithdrawData_(ss, withdrawRows);
  applyNyukaiLabelsToThreads_(enrollThreads);
  applyTaikaiLabelsToThreads_(withdrawThreads);

  return { enrollTotal: enrollRows.length, withdrawTotal: withdrawRows.length };
}

/** 直近Nか月分だけGmailを確認し、新着のみ追加（手動更新用・高速） */
function fetchMembershipEmailsIncremental_(ss, monthsBack) {
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  const existingEnrollIds = loadMessageIds_(readEnrollData_(dataSheet));
  const existingWithdrawIds = loadMessageIds_(readWithdrawData_(dataSheet));
  const enrollSeen = loadMemberDedupeState_(readEnrollData_(dataSheet));
  const withdrawSeen = loadMemberDedupeState_(readWithdrawData_(dataSheet));
  const afterQuery = " after:" + gmailAfterDate_(monthsBack || 2);

  const newEnroll = [];
  const newWithdraw = [];

  searchGmailAllThreads_(NYUKAI_SEARCH_QUERY + afterQuery).forEach(function (thread) {
    const messages = thread.getMessages();
    for (let i = 0; i < messages.length; i++) {
      const message = messages[i];
      const msgId = message.getId();
      if (existingEnrollIds[msgId]) continue;
      const row = parseNyukaiMessage_(message);
      if (!row) continue;
      if (!registerMembershipRow_(row, enrollSeen, newEnroll, DATA_ENROLL_COLS)) continue;
      existingEnrollIds[msgId] = true;
      break;
    }
  });

  searchGmailAllThreads_(TAIKAI_SEARCH_QUERY + afterQuery).forEach(function (thread) {
    const messages = thread.getMessages();
    for (let i = 0; i < messages.length; i++) {
      const message = messages[i];
      const msgId = message.getId();
      if (existingWithdrawIds[msgId]) continue;
      const row = parseTaikaiMessage_(message);
      if (!row) continue;
      if (!registerMembershipRow_(row, withdrawSeen, newWithdraw, DATA_WITHDRAW_COLS)) continue;
      existingWithdrawIds[msgId] = true;
      break;
    }
  });

  if (newEnroll.length > 0) {
    writeEnrollData_(ss, readEnrollData_(dataSheet).concat(newEnroll));
  }
  if (newWithdraw.length > 0) {
    writeWithdrawData_(ss, readWithdrawData_(dataSheet).concat(newWithdraw));
  }

  return { enrollAdded: newEnroll.length, withdrawAdded: newWithdraw.length };
}

function gmailAfterDate_(monthsBack) {
  const d = new Date();
  d.setDate(1);
  d.setMonth(d.getMonth() - (monthsBack || 0));
  return d.getFullYear() + "/" + String(d.getMonth() + 1).padStart(2, "0") + "/01";
}

/** 毎日19:30のOP自動更新から呼ばれる（当月の新着のみ追加） */
function fetchMembershipEmailsSilent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupDataSheet_(ss);
  setupDisplaySheet_(ss);
  installMembershipLabelsSilent_();

  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  const existingEnrollIds = loadMessageIds_(readEnrollData_(dataSheet));
  const existingWithdrawIds = loadMessageIds_(readWithdrawData_(dataSheet));
  const enrollSeen = loadMemberDedupeState_(readEnrollData_(dataSheet));
  const withdrawSeen = loadMemberDedupeState_(readWithdrawData_(dataSheet));

  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const afterQuery = " after:" + year + "/" + month + "/01";

  const newEnroll = [];
  const newWithdraw = [];

  const enrollThreads = searchGmailAllThreads_(NYUKAI_SEARCH_QUERY + afterQuery);
  enrollThreads.forEach(function (thread) {
    const messages = thread.getMessages();
    for (let i = 0; i < messages.length; i++) {
      const message = messages[i];
      const msgId = message.getId();
      if (existingEnrollIds[msgId]) continue;
      const date = message.getDate();
      if (date.getFullYear() !== year || date.getMonth() !== now.getMonth()) continue;
      const row = parseNyukaiMessage_(message);
      if (!row) continue;
      if (!registerMembershipRow_(row, enrollSeen, newEnroll, DATA_ENROLL_COLS)) continue;
      existingEnrollIds[msgId] = true;
      break;
    }
  });

  const withdrawThreads = searchGmailAllThreads_(TAIKAI_SEARCH_QUERY + afterQuery);
  withdrawThreads.forEach(function (thread) {
    const messages = thread.getMessages();
    for (let i = 0; i < messages.length; i++) {
      const message = messages[i];
      const msgId = message.getId();
      if (existingWithdrawIds[msgId]) continue;
      const date = message.getDate();
      if (date.getFullYear() !== year || date.getMonth() !== now.getMonth()) continue;
      const row = parseTaikaiMessage_(message);
      if (!row) continue;
      if (!registerMembershipRow_(row, withdrawSeen, newWithdraw, DATA_WITHDRAW_COLS)) continue;
      existingWithdrawIds[msgId] = true;
      break;
    }
  });

  if (newEnroll.length > 0) {
    writeEnrollData_(ss, readEnrollData_(dataSheet).concat(newEnroll));
  }
  if (newWithdraw.length > 0) {
    writeWithdrawData_(ss, readWithdrawData_(dataSheet).concat(newWithdraw));
  }

  applyNyukaiLabelsToThreads_(enrollThreads);
  applyTaikaiLabelsToThreads_(withdrawThreads);
  refreshMembershipDisplay_(ss);
}

function handleTaikaiSheetEdit_(e) {
  if (!e || !e.range) return false;
  if (e.range.getSheet().getName() !== SHEET_NAME_MEMBERS) return false;
  const cell = e.range.getA1Notation();
  if (cell === "B1" || cell === "G1") {
    refreshMembershipDisplay_(e.source);
    return true;
  }
  if (e.range.getColumn() === 10 && e.range.getRow() >= DISPLAY_START_ROW) {
    persistWithdrawCancelFromDisplay_(e);
    try {
      syncMembershipDailyCountsSilent_();
    } catch (err) {
      Logger.log("日別退会の再反映エラー: " + err.message);
    }
    return true;
  }
  return false;
}

function persistWithdrawCancelFromDisplay_(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const mailId = String(sheet.getRange(row, 11).getValue() || "").trim();
  const checked = e.range.getValue() === true;
  const ss = e.source;
  const rows = readWithdrawData_(ss.getSheetByName(SHEET_NAME_DATA));
  let updated = false;

  for (let i = 0; i < rows.length; i++) {
    if (mailId && String(rows[i][4] || "") === mailId) {
      rows[i][5] = checked;
      updated = true;
      break;
    }
  }
  if (!updated && !mailId) {
    const name = String(sheet.getRange(row, 7).getValue() || "").trim();
    const month = normalizeYearMonthLabel_(sheet.getRange(row, 8).getDisplayValue() || sheet.getRange(row, 8).getValue());
    for (let j = 0; j < rows.length; j++) {
      if (String(rows[j][1] || "").trim() === name && normalizeYearMonthLabel_(rows[j][2]) === month) {
        rows[j][5] = checked;
        updated = true;
        break;
      }
    }
  }
  if (updated) writeWithdrawData_(ss, rows);
}

// ── シート構造 ──

function setupDataSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME_DATA);
  sheet.hideSheet();
  sheet.getRange(DATA_HEADER_ROW, DATA_ENROLL_COL, 1, DATA_ENROLL_COLS).setValues([ENROLL_HEADERS]);
  sheet.getRange(DATA_HEADER_ROW, DATA_WITHDRAW_COL, 1, DATA_WITHDRAW_COLS).setValues([WITHDRAW_HEADERS]);
  sheet.getRange(DATA_HEADER_ROW, DATA_ENROLL_COL, 1, DATA_ENROLL_COLS).setFontWeight("bold").setBackground("#e2efda");
  sheet.getRange(DATA_HEADER_ROW, DATA_WITHDRAW_COL, 1, DATA_WITHDRAW_COLS).setFontWeight("bold").setBackground("#fce4d6");
  sheet.setFrozenRows(1);
}

function setupDisplaySheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME_MEMBERS);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME_MEMBERS);

  removeSheetFilter_(sheet);

  sheet.getRange("A1").setValue("入会年月を選択 ➡").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("F1").setValue("退会年月を選択 ➡").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange(DISPLAY_HEADER_ROW, 1, 1, 4).setValues([ENROLL_HEADERS.slice(0, 4)]);
  sheet.getRange(DISPLAY_HEADER_ROW, 6, 1, DISPLAY_WITHDRAW_HEADERS.length).setValues([DISPLAY_WITHDRAW_HEADERS]);
  sheet.getRange(DISPLAY_HEADER_ROW, 1, 1, 4).setFontWeight("bold").setBackground("#e2efda");
  sheet.getRange(DISPLAY_HEADER_ROW, 6, 1, DISPLAY_WITHDRAW_HEADERS.length).setFontWeight("bold").setBackground("#fce4d6");
  sheet.setFrozenRows(DISPLAY_START_ROW - 1);
  sheet.setColumnWidth(1, 155);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 95);
  sheet.setColumnWidth(4, 85);
  sheet.setColumnWidth(6, 155);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 95);
  sheet.setColumnWidth(9, 85);
  sheet.setColumnWidth(10, 95);
  sheet.hideColumns(5);
  sheet.hideColumns(11);
  sheet.getRange("B1").setNumberFormat("@");
  sheet.getRange("G1").setNumberFormat("@");
  sheet.getRange("C:C").setNumberFormat("@");
  sheet.getRange("H:H").setNumberFormat("@");
}

function refreshMembershipDisplay_(ss) {
  setupDisplaySheet_(ss);
  const displaySheet = ss.getSheetByName(SHEET_NAME_MEMBERS);
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  const allEnroll = readEnrollData_(dataSheet);
  const allWithdraw = readWithdrawData_(dataSheet);

  setupMonthSelectorCell_(displaySheet.getRange("B1"), extractMonthListFromRows_(allEnroll), true);
  setupMonthSelectorCell_(displaySheet.getRange("G1"), extractMonthListFromRows_(allWithdraw), true);

  const enrollLabel = normalizeYearMonthLabel_(
    displaySheet.getRange("B1").getDisplayValue() || displaySheet.getRange("B1").getValue()
  );
  const withdrawLabel = normalizeYearMonthLabel_(
    displaySheet.getRange("G1").getDisplayValue() || displaySheet.getRange("G1").getValue()
  );

  const filteredEnroll = allEnroll.filter(function (r) {
    return normalizeYearMonthLabel_(r[2]) === enrollLabel;
  });
  const filteredWithdraw = allWithdraw.filter(function (r) {
    return normalizeYearMonthLabel_(r[2]) === withdrawLabel;
  });

  sortRowsNewestFirst_(filteredEnroll);
  sortRowsNewestFirst_(filteredWithdraw);
  writeDisplayRows_(displaySheet, filteredEnroll, filteredWithdraw);

  return {
    enrollCount: filteredEnroll.length,
    withdrawCount: filteredWithdraw.filter(function (r) { return !isWithdrawalCancelled_(r); }).length,
    withdrawCancelCount: filteredWithdraw.filter(function (r) { return isWithdrawalCancelled_(r); }).length,
    enrollLabel: enrollLabel,
    withdrawLabel: withdrawLabel
  };
}

function writeDisplayRows_(sheet, enrollRows, withdrawRows) {
  removeSheetFilter_(sheet);

  const rowCount = Math.max(enrollRows.length, withdrawRows.length);
  const prevRows = countDisplayDataRows_(sheet);
  const clearRows = Math.max(rowCount, prevRows);

  if (clearRows > 0) {
    sheet.getRange(DISPLAY_START_ROW, 1, clearRows, DISPLAY_COLS).clearContent();
    sheet.getRange(DISPLAY_START_ROW, 10, clearRows, 1).removeCheckboxes();
  }
  if (rowCount === 0) return;

  const withdrawCount = withdrawRows.length;
  const displayRows = [];
  for (let i = 0; i < rowCount; i++) {
    const row = new Array(DISPLAY_COLS).fill("");
    if (i < enrollRows.length) {
      row[0] = enrollRows[i][0];
      row[1] = enrollRows[i][1];
      row[2] = normalizeYearMonthLabel_(enrollRows[i][2]);
      row[3] = enrollRows[i][3];
    }
    if (i < withdrawCount) {
      row[5] = withdrawRows[i][0];
      row[6] = withdrawRows[i][1];
      row[7] = normalizeYearMonthLabel_(withdrawRows[i][2]);
      row[8] = withdrawRows[i][3];
      row[9] = isWithdrawalCancelled_(withdrawRows[i]);
      row[10] = withdrawRows[i][4];
    }
    displayRows.push(row);
  }

  sheet.getRange(DISPLAY_START_ROW, 1, displayRows.length, DISPLAY_COLS).setValues(displayRows);
  sheet.getRange(DISPLAY_START_ROW, 1, displayRows.length, 1).setNumberFormat("yyyy/mm/dd hh:mm");
  sheet.getRange(DISPLAY_START_ROW, 6, displayRows.length, 1).setNumberFormat("yyyy/mm/dd hh:mm");
  if (withdrawCount > 0) {
    sheet.getRange(DISPLAY_START_ROW, 10, withdrawCount, 1).insertCheckboxes();
  }

  const enrollMonths = [];
  const withdrawMonths = [];
  for (let i = 0; i < rowCount; i++) {
    enrollMonths.push(i < enrollRows.length ? normalizeYearMonthLabel_(enrollRows[i][2]) : "");
    withdrawMonths.push(i < withdrawRows.length ? normalizeYearMonthLabel_(withdrawRows[i][2]) : "");
  }
  applyMonthTextToColumn_(sheet, 3, DISPLAY_START_ROW, enrollMonths);
  applyMonthTextToColumn_(sheet, 8, DISPLAY_START_ROW, withdrawMonths);
}

/** 3行目から連続する表示行のみ数える（900行目以降の残骸は無視） */
function countDisplayDataRows_(sheet) {
  const maxRow = Math.min(sheet.getLastRow(), DISPLAY_START_ROW + DISPLAY_MAX_SCAN - 1);
  if (maxRow < DISPLAY_START_ROW) return 0;

  let count = 0;
  let emptyStreak = 0;
  for (let r = DISPLAY_START_ROW; r <= maxRow; r++) {
    const row = sheet.getRange(r, 1, 1, DISPLAY_COLS).getValues()[0];
    const hasData = row[0] || row[1] || row[2] || row[5] || row[6] || row[7];
    if (hasData) {
      count = r - DISPLAY_START_ROW + 1;
      emptyStreak = 0;
    } else {
      emptyStreak++;
      if (emptyStreak >= 10 && count > 0) break;
    }
  }
  return count;
}

function setupMonthSelectorCell_(cell, monthList, keepCurrent) {
  cell.clearDataValidations();
  cell.clearFormat();
  cell.setBackground("#fff2cc").setFontWeight("bold").setHorizontalAlignment("center");

  if (!monthList.length) {
    setMonthTextCell_(cell, "（データなし）");
    return;
  }

  const val = normalizeYearMonthLabel_(cell.getDisplayValue() || cell.getValue());
  const pick = (keepCurrent && monthList.indexOf(val) >= 0) ? val : monthList[0];
  setMonthTextCell_(cell, pick);

  cell.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(monthList, true).build()
  );
}

function extractMonthListFromRows_(rows) {
  const keys = {};
  rows.forEach(function (r) {
    const label = normalizeYearMonthLabel_(r[2]);
    if (label) keys[label] = yearMonthSortKey_(label);
  });
  return Object.keys(keys).sort(function (a, b) { return keys[b] - keys[a]; });
}

function yearMonthSortKey_(label) {
  const m = String(label).match(/^(\d{4})年(\d{1,2})月$/);
  return m ? parseInt(m[1], 10) * 100 + parseInt(m[2], 10) : 0;
}

function removeSheetFilter_(sheet) {
  if (!sheet) return;
  const f = sheet.getFilter();
  if (f) f.remove();
}

// ── データ読み書き ──

function readEnrollData_(dataSheet) {
  return readDataBlock_(dataSheet, DATA_ENROLL_COL, DATA_ENROLL_COLS);
}

function readWithdrawData_(dataSheet) {
  return readDataBlock_(dataSheet, DATA_WITHDRAW_COL, DATA_WITHDRAW_COLS);
}

function readDataBlock_(dataSheet, startCol, colCount) {
  if (!dataSheet) return [];
  const lastRow = dataSheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];
  const numRows = lastRow - DATA_START_ROW + 1;
  return dataSheet.getRange(DATA_START_ROW, startCol, numRows, colCount)
    .getValues()
    .filter(function (r) { return r[0] || r[1]; })
    .map(function (r) { return normalizeMemberRow_(r, colCount); });
}

function writeEnrollData_(ss, rows) {
  writeDataBlock_(ss, DATA_ENROLL_COL, dedupeMemberRows_(rows, DATA_ENROLL_COLS), DATA_ENROLL_COLS);
}

function writeWithdrawData_(ss, rows) {
  writeDataBlock_(ss, DATA_WITHDRAW_COL, dedupeMemberRows_(rows, DATA_WITHDRAW_COLS), DATA_WITHDRAW_COLS);
}

function writeDataBlock_(ss, startCol, rows, colCount) {
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!dataSheet) return;

  const normalized = rows.map(function (r) { return normalizeMemberRow_(r, colCount); });
  sortRowsNewestFirst_(normalized);

  const lastRow = dataSheet.getLastRow();
  if (lastRow >= DATA_START_ROW) {
    dataSheet.getRange(DATA_START_ROW, startCol, lastRow - DATA_START_ROW + 1, colCount).clearContent();
  }
  if (!normalized.length) return;

  const monthCol = startCol + 2;
  dataSheet.getRange(DATA_START_ROW, monthCol, normalized.length, 1).setNumberFormat("@");
  dataSheet.getRange(DATA_START_ROW, startCol, normalized.length, colCount).setValues(normalized);
  dataSheet.getRange(DATA_START_ROW, startCol, normalized.length, 1).setNumberFormat("yyyy/mm/dd hh:mm");

  const monthLabels = normalized.map(function (r) { return r[2]; });
  applyMonthTextToColumn_(dataSheet, monthCol, DATA_START_ROW, monthLabels);
}

/** 裏方データの「入会月」「退会月」を文字列に直す */
function repairMonthValuesInData_(ss) {
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!dataSheet) return;

  const enroll = readEnrollData_(dataSheet);
  const withdraw = readWithdrawData_(dataSheet);
  let changed = false;

  enroll.forEach(function (r) {
    const fixed = normalizeYearMonthLabel_(r[2]);
    if (r[2] instanceof Date || /Mon |GMT/.test(String(r[2])) || fixed !== r[2]) {
      r[2] = fixed;
      changed = true;
    }
  });
  withdraw.forEach(function (r) {
    const fixed = normalizeYearMonthLabel_(r[2]);
    if (r[2] instanceof Date || /Mon |GMT/.test(String(r[2])) || fixed !== r[2]) {
      r[2] = fixed;
      changed = true;
    }
  });

  if (changed) {
    writeEnrollData_(ss, enroll);
    writeWithdrawData_(ss, withdraw);
  }
}

function clearAllMemberData_(ss) {
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (dataSheet && dataSheet.getLastRow() >= DATA_START_ROW) {
    dataSheet.getRange(DATA_START_ROW, 1, dataSheet.getLastRow() - DATA_START_ROW + 1, 12).clearContent();
  }
}

function loadMessageIds_(rows) {
  const ids = {};
  rows.forEach(function (r) {
    const id = String(r[4] || "").trim();
    if (id) ids[id] = true;
  });
  return ids;
}

function normalizeMemberRow_(row, colCount) {
  const n = colCount || DATA_ENROLL_COLS;
  const out = row.slice(0, n);
  while (out.length < n) out.push("");
  out[2] = normalizeYearMonthLabel_(out[2]);
  if (n >= DATA_WITHDRAW_COLS && out[3] === "一般会員") out[3] = CATEGORY_SIX_MONTH;
  if (n >= DATA_WITHDRAW_COLS) out[5] = isWithdrawalCancelled_(out);
  if (n === DATA_ENROLL_COLS && out[3] === "一般会員") out[3] = CATEGORY_SIX_MONTH;
  return out;
}

function isWithdrawalCancelled_(row) {
  if (!row || row.length < 6) return false;
  const v = row[5];
  return v === true || String(v).toUpperCase() === "TRUE";
}

function normalizeYearMonthLabel_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return formatYearMonth_(value.getFullYear(), value.getMonth() + 1);
  }
  const s = String(value || "").trim();
  if (!s || s === "（データなし）") return s;

  let m = s.match(/^(\d{4})年(\d{1,2})月$/);
  if (m) return formatYearMonth_(parseInt(m[1], 10), parseInt(m[2], 10));

  m = s.match(/^(\d{4})\/(\d{1,2})$/);
  if (m) return formatYearMonth_(parseInt(m[1], 10), parseInt(m[2], 10));

  if (/Mon |Tue |Wed |Thu |Fri |Sat |Sun |GMT/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return formatYearMonth_(d.getFullYear(), d.getMonth() + 1);
  }

  return s;
}

function formatYearMonth_(year, month) {
  return year + "年" + month + "月";
}

/** 「2026年6月」を日付にされないようテキストとして書き込む */
function setMonthTextCell_(cell, text) {
  const label = normalizeYearMonthLabel_(text);
  cell.setNumberFormat("@");
  if (!label || label === "（データなし）") {
    cell.setValue(label || "");
    return;
  }
  cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(label).build());
}

function applyMonthTextToColumn_(sheet, col, startRow, labels) {
  for (let i = 0; i < labels.length; i++) {
    if (!labels[i]) continue;
    setMonthTextCell_(sheet.getRange(startRow + i, col), labels[i]);
  }
}

function sortRowsNewestFirst_(rows) {
  rows.sort(function (a, b) {
    const da = a[0] instanceof Date ? a[0].getTime() : new Date(a[0]).getTime();
    const db = b[0] instanceof Date ? b[0].getTime() : new Date(b[0]).getTime();
    return db - da;
  });
}

function dedupeStoredMemberData_(ss) {
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!dataSheet) return;

  const enroll = dedupeMemberRows_(readEnrollData_(dataSheet), DATA_ENROLL_COLS);
  const withdraw = dedupeMemberRows_(readWithdrawData_(dataSheet), DATA_WITHDRAW_COLS);
  const rawEnroll = readEnrollData_(dataSheet);
  const rawWithdraw = readWithdrawData_(dataSheet);
  if (enroll.length !== rawEnroll.length || withdraw.length !== rawWithdraw.length) {
    writeEnrollData_(ss, enroll);
    writeWithdrawData_(ss, withdraw);
  }
}

function createMemberDedupeState_() {
  return { msgId: {}, composite: {} };
}

function loadMemberDedupeState_(rows) {
  const seen = createMemberDedupeState_();
  rows.forEach(function (r) {
    rememberMemberRowKeys_(r, seen);
  });
  return seen;
}

function rememberMemberRowKeys_(row, seen) {
  const colCount = row.length >= DATA_WITHDRAW_COLS ? DATA_WITHDRAW_COLS : DATA_ENROLL_COLS;
  const normalized = normalizeMemberRow_(row, colCount);
  const msgId = String(normalized[4] || "").trim();
  if (msgId) seen.msgId[msgId] = true;
  seen.composite[buildMemberRowDedupeKey_(normalized)] = true;
}

function canRegisterMembershipRow_(row, seen, colCount) {
  const normalized = normalizeMemberRow_(row, colCount);
  const msgId = String(normalized[4] || "").trim();
  if (msgId && seen.msgId[msgId]) return false;
  const composite = buildMemberRowDedupeKey_(normalized);
  if (seen.composite[composite]) return false;
  return true;
}

function registerMembershipRow_(row, seen, rows, colCount) {
  if (!canRegisterMembershipRow_(row, seen, colCount)) return false;
  const normalized = normalizeMemberRow_(row, colCount);
  rememberMemberRowKeys_(normalized, seen);
  rows.push(normalized);
  return true;
}

function buildMemberRowDedupeKey_(row) {
  const ts = parseMemberTimestamp_(row[0]);
  const minute = ts
    ? Utilities.formatDate(truncateToMinute_(ts), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm")
    : String(row[0] || "").trim();
  const name = normalizeMemberName_(row[1]);
  const month = normalizeYearMonthLabel_(row[2]);
  return name + "\x1f" + minute + "\x1f" + month;
}

function normalizeMemberName_(name) {
  return String(name || "").replace(/\s+/g, "").trim();
}

function truncateToMinute_(date) {
  const d = date instanceof Date ? new Date(date.getTime()) : new Date(date);
  d.setSeconds(0, 0);
  return d;
}

function dedupeMemberRows_(rows, colCount) {
  const seen = createMemberDedupeState_();
  const out = [];
  const n = colCount || DATA_ENROLL_COLS;
  rows.forEach(function (r) {
    if (!canRegisterMembershipRow_(r, seen, n)) return;
    registerMembershipRow_(r, seen, out, n);
  });
  return out;
}

function dedupeRows_(rows, colCount) {
  return dedupeMemberRows_(rows, colCount);
}

// ── 月次日報シート（2606など）への日別反映 ──

/**
 * 当月の月次シート（例: 6月なら 2606）へ、タイムスタンプ日付ごとの入会・退会件数を書き込む。
 * 6ヶ月割入会: E5〜E35 / 法人会員入会: G5〜G35 / 退会: E40〜E70（退会キャンセル除く）
 * ※対象セルに値だけ入れ、書式・他セルは変更しない
 */
function syncMembershipDailyCountsSilent_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ym = resolveCurrentMonthYm_();
  syncMembershipDailyCountsForMonth_(ss, ym.year, ym.month);
}

function syncMembershipDailyCountsForMonth_(ss, year, month) {
  const sheetName = formatMonthlySheetName_(year, month);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("月次シートが見つかりません: " + sheetName);
    return { sheetName: sheetName, updated: false };
  }

  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!dataSheet) return { sheetName: sheetName, updated: false };

  const enrollRows = readEnrollData_(dataSheet);
  const withdrawRows = readWithdrawData_(dataSheet);

  const enrollGeneralByDay = countMembersByDay_(enrollRows, year, month, CATEGORY_SIX_MONTH);
  const enrollCorporateByDay = countMembersByDay_(enrollRows, year, month, CATEGORY_CORPORATE);
  const withdrawByDay = countWithdrawalsByDay_(withdrawRows, year, month);

  writeDailyCountColumn_(sheet, DAILY_ENROLL_ROW_START, DAILY_ENROLL_GENERAL_COL, enrollGeneralByDay, "");
  writeDailyCountColumn_(sheet, DAILY_ENROLL_ROW_START, DAILY_ENROLL_CORPORATE_COL, enrollCorporateByDay, "");
  writeDailyCountColumn_(sheet, DAILY_WITHDRAW_ROW_START, DAILY_WITHDRAW_COL, withdrawByDay, "");

  return { sheetName: sheetName, updated: true };
}

/** 指定列の日別件数だけを書き込む（書式・装飾は触らない） */
function writeDailyCountColumn_(sheet, startRow, col, countsByDay, zeroValue) {
  const values = [];
  for (let day = 1; day <= DAILY_COUNT_DAYS; day++) {
    const n = countsByDay[day] || 0;
    values.push([n > 0 ? n : zeroValue]);
  }
  sheet.getRange(startRow, col, DAILY_COUNT_DAYS, 1).setValues(values);
}

function resolveCurrentMonthYm_() {
  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth() };
}

function formatMonthlySheetName_(year, month) {
  const yy = String(year % 100).padStart(2, "0");
  const mm = String(month + 1).padStart(2, "0");
  return yy + mm;
}

function countMembersByDay_(rows, year, month, category) {
  const counts = {};
  for (let d = 1; d <= DAILY_COUNT_DAYS; d++) counts[d] = 0;
  rows.forEach(function (r) {
    if (category && !matchesEnrollCategory_(r[3], category)) return;
    const ts = parseMemberTimestamp_(r[0]);
    if (!ts) return;
    if (ts.getFullYear() !== year || ts.getMonth() !== month) return;
    const day = ts.getDate();
    if (day >= 1 && day <= DAILY_COUNT_DAYS) counts[day]++;
  });
  return counts;
}

function matchesEnrollCategory_(value, category) {
  const cat = String(value || "").trim();
  if (cat === category) return true;
  if (category === CATEGORY_SIX_MONTH && cat === "一般会員") return true;
  if (category === CATEGORY_CORPORATE && cat === "法人会員") return true;
  return false;
}

/** 退会は「退会月」で月を判定（法人・翌月末は翌月に計上） */
function countWithdrawalsByDay_(rows, year, month) {
  const counts = {};
  for (let d = 1; d <= DAILY_COUNT_DAYS; d++) counts[d] = 0;
  rows.forEach(function (r) {
    if (isWithdrawalCancelled_(r)) return;
    const ym = parseYearMonthFromLabel_(r[2]);
    if (!ym) return;
    if (ym.year !== year || ym.month !== month) return;
    const ts = parseMemberTimestamp_(r[0]);
    if (!ts) return;
    const day = ts.getDate();
    if (day >= 1 && day <= DAILY_COUNT_DAYS) counts[day]++;
  });
  return counts;
}

function parseYearMonthFromLabel_(label) {
  const normalized = normalizeYearMonthLabel_(label);
  const m = String(normalized).match(/^(\d{4})年(\d{1,2})月$/);
  if (!m) return null;
  const month = parseInt(m[2], 10) - 1;
  if (month < 0 || month > 11) return null;
  return { year: parseInt(m[1], 10), month: month };
}

function parseMemberTimestamp_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === "number" && value > 30000) {
    const base = new Date(1899, 11, 30);
    const d = new Date(base.getTime() + Math.floor(value) * 86400000);
    return isNaN(d.getTime()) ? null : d;
  }
  const s = String(value || "").trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    const d = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// ── 旧レイアウト移行 ──

function migrateLegacyLayouts_(ss) {
  setupDataSheet_(ss);
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (readEnrollData_(dataSheet).length > 0 || readWithdrawData_(dataSheet).length > 0) return;

  const display = ss.getSheetByName(SHEET_NAME_MEMBERS);
  if (!display) return;

  const scanEnd = Math.min(display.getLastRow(), DISPLAY_START_ROW + DISPLAY_MAX_SCAN - 1);
  if (scanEnd < DISPLAY_START_ROW) return;

  const numRows = scanEnd - DISPLAY_START_ROW + 1;
  const read5 = function (col) {
    return display.getRange(DISPLAY_START_ROW, col, numRows, 5).getValues()
      .filter(function (r) { return r[0] || r[1]; })
      .map(function (r) { return normalizeMemberRow_(r, DATA_ENROLL_COLS); });
  };

  const enroll = read5(11).length > 0 ? read5(11) : read5(1);
  let withdraw = read5(16).length > 0 ? read5(16) : read5(6);
  if (!withdraw.length && enroll.length && !read5(16).length && !read5(6).length) {
    withdraw = enroll;
    enroll.length = 0;
  }

  if (enroll.length) writeEnrollData_(ss, dedupeRows_(enroll, DATA_ENROLL_COLS));
  if (withdraw.length) writeWithdrawData_(ss, dedupeRows_(withdraw, DATA_WITHDRAW_COLS));

  removeSheetFilter_(display);
  display.getRange(DISPLAY_START_ROW, 1, numRows, 20).clearContent();

  ["退会", "退会データ"].forEach(function (name) {
    const legacy = ss.getSheetByName(name);
    if (legacy) ss.deleteSheet(legacy);
  });
}

// ── 入会メール解析 ──

function parseNyukaiMessage_(message) {
  const subject = String(message.getSubject() || "");
  if (subject.indexOf("ご入会") === -1) return null;

  const body = getMessageBodyText_(message);
  const date = message.getDate();
  const msgId = message.getId();
  const name = extractNyukaiName_(body);
  const ym = calcNyukaiYm_(body, date);
  if (!name || !ym) return null;

  return [date, name, ym.label, detectNyukaiCategory_(subject, body), msgId];
}

function extractNyukaiName_(body) {
  const text = String(body || "");
  const m1 = text.match(/お名前\s*[:：]\s*(.+?)\s*様/);
  if (m1) return String(m1[1]).replace(/^[>\s]+/, "").trim();

  const lines = text.split(/\r?\n/);
  for (let i = 0; i < Math.min(lines.length, 20); i++) {
    const line = String(lines[i] || "").trim();
    if (!line) continue;
    if (/JOYFIT|受付番号|この度は|お支払い|ご利用開始|ご入会内容|会員情報/.test(line)) continue;
    const m = line.match(/^(.{1,30}?)\s*様\s*$/);
    if (m) return String(m[1]).replace(/^[>\s]+/, "").trim();
  }
  return "";
}

function calcNyukaiYm_(body, emailDate) {
  const startMatch = String(body || "").match(/ご利用開始日[の]?\s*(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (startMatch) {
    return { label: normalizeYearMonthLabel_(startMatch[1] + "年" + parseInt(startMatch[2], 10) + "月") };
  }
  const monthMatch = String(body || "").match(/[（(](\d{1,2})月分[）)]/);
  if (monthMatch) {
    return { label: normalizeYearMonthLabel_(emailDate.getFullYear() + "年" + monthMatch[1] + "月") };
  }
  return { label: normalizeYearMonthLabel_(emailDate.getFullYear() + "年" + (emailDate.getMonth() + 1) + "月") };
}

function detectNyukaiCategory_(subject, body) {
  if (String(body || "").indexOf("法人") !== -1) return CATEGORY_CORPORATE;
  return CATEGORY_SIX_MONTH;
}

function extractNyukaiRowsFromThreads_(threads) {
  const rows = [];
  const seen = createMemberDedupeState_();
  threads.forEach(function (thread) {
    const messages = thread.getMessages();
    for (let i = 0; i < messages.length; i++) {
      const row = parseNyukaiMessage_(messages[i]);
      if (!row) continue;
      if (registerMembershipRow_(row, seen, rows, DATA_ENROLL_COLS)) break;
    }
  });
  return rows;
}

// ── 退会メール解析 ──

function parseTaikaiMessage_(message) {
  const subject = String(message.getSubject() || "");
  if (subject.indexOf("ご退会") === -1) return null;

  const body = getMessageBodyText_(message);
  const date = message.getDate();
  const msgId = message.getId();
  const name = extractTaikaiName_(body);
  const finalMatch = body.match(/【最終ご利用日】[：:]\s*([^\r\n]+)/);
  const ym = calcTaikaiWithdrawalYm_(date, finalMatch ? String(finalMatch[1]).trim() : "");
  if (!ym || !name) return null;

  return [date, name, ym.label, detectTaikaiCategory_(subject, body), msgId];
}

function extractTaikaiName_(body) {
  const text = String(body || "");
  const m1 = text.match(/お名前\s*[:：]\s*(.+?)\s*様/);
  if (m1) return String(m1[1]).replace(/^[>\s]+/, "").trim();

  const lines = text.split(/\r?\n/);
  for (let i = 0; i < Math.min(lines.length, 15); i++) {
    const line = String(lines[i] || "").trim();
    if (!line || /JOYFIT|ご退会|日頃より|以下情報|会員情報/.test(line)) continue;
    const m = line.match(/^(.{1,30}?)\s*様\s*$/);
    if (m) return String(m[1]).replace(/^[>\s]+/, "").trim();
  }
  return "";
}

function calcTaikaiWithdrawalYm_(emailDate, finalDateText) {
  let y = emailDate.getFullYear();
  let m = emailDate.getMonth();
  if (finalDateText.indexOf("翌月末") !== -1) {
    m += 1;
    if (m > 11) { m = 0; y += 1; }
  } else if (finalDateText.indexOf("当月末") !== -1) {
    // 受信月が退会月
  } else {
    return null;
  }
  return { label: normalizeYearMonthLabel_(y + "年" + (m + 1) + "月") };
}

function detectTaikaiCategory_(subject, body) {
  if (subject.indexOf("【JOYFIT24経堂】") !== -1) return "一般会員";
  if (body.indexOf("ご退会申請の受付") !== -1) return "一般会員";
  if (body.indexOf("ご退会のお手続きが完了") !== -1) return "法人会員";
  return "不明";
}

function extractTaikaiRowsFromThreads_(threads) {
  const rows = [];
  const seen = createMemberDedupeState_();
  threads.forEach(function (thread) {
    const messages = thread.getMessages();
    for (let i = 0; i < messages.length; i++) {
      const row = parseTaikaiMessage_(messages[i]);
      if (!row) continue;
      if (registerMembershipRow_(row, seen, rows, DATA_WITHDRAW_COLS)) break;
    }
  });
  return rows;
}

// ── Gmailラベル ──

function installMembershipLabelsSilent_() {
  [LABEL_NYUKAI_GENERAL, LABEL_NYUKAI_CORPORATE, LABEL_TAIKAI_GENERAL, LABEL_TAIKAI_CORPORATE]
    .forEach(function (name) { getOrCreateGmailLabel_(name); });
}

function applyNyukaiLabelsToThreads_(threads) {
  const labelGeneral = getOrCreateGmailLabel_(LABEL_NYUKAI_GENERAL);
  const labelCorporate = getOrCreateGmailLabel_(LABEL_NYUKAI_CORPORATE);
  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (message) {
      if (String(message.getSubject() || "").indexOf("ご入会") === -1) return;
      const cat = detectNyukaiCategory_(message.getSubject(), getMessageBodyText_(message));
      thread.addLabel(cat === CATEGORY_CORPORATE ? labelCorporate : labelGeneral);
    });
  });
}

function applyTaikaiLabelsToThreads_(threads) {
  const labelGeneral = getOrCreateGmailLabel_(LABEL_TAIKAI_GENERAL);
  const labelCorporate = getOrCreateGmailLabel_(LABEL_TAIKAI_CORPORATE);
  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (message) {
      if (String(message.getSubject() || "").indexOf("ご退会") === -1) return;
      const cat = detectTaikaiCategory_(message.getSubject(), getMessageBodyText_(message));
      if (cat === "一般会員") thread.addLabel(labelGeneral);
      else if (cat === "法人会員") thread.addLabel(labelCorporate);
    });
  });
}

function getOrCreateGmailLabel_(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) label = GmailApp.createLabel(name);
  return label;
}
