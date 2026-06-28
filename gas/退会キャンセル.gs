/**
 * JOYFIT24経堂 — 6ヶ月割キャンペーン管理
 *
 * シート: 6ヶ月割_入会確認（入会メール台帳）/ 照合 / 在籍期間（要連絡者・一斉送信用）
 * メニュー: オプション出力.gs の JOYFIT から実行
 */

var KYODO_ENROLL_SUBJECT = "【JOYFIT24経堂】ご入会ありがとうございます。※必ず一読ください";
var KYODO_ENROLL_FROM = "info@joyfit-service.jp";
var KYODO_ENROLL_SHEET_NAME = "6ヶ月割_入会確認";
var KYODO_MATCH_SHEET_NAME = "照合";
var KYODO_TEST_SHEET_NAME = "テスト";
var KYODO_ENROLL_LOG_SHEET_NAME = "_取込ログ";
var KYODO_ENROLL_PAGE_SIZE = 100;
var KYODO_ENROLL_RANGE_START = new Date(2025, 11, 1); // 2025/12/1
var MY_TEST_EMAIL = "r-kusaka@okamoto-group.co.jp"; // テスト用下書きアドレス

var KYODO_WITHDRAW_SUBJECT = "【JOYFIT24経堂】ご退会のお手続きについて";
var KYODO_TENURE_SHEET_NAME = "在籍期間";
var KYODO_TENURE_MIN_MONTHS = 6;
var KYODO_TENURE_SEND_TARGET_HEADER = "送信対象";

/** メニューは オプション出力.gs の onOpen（JOYFIT）に統合済み */

/** メニュー: 空なら全件確認、それ以外は当月。照合シートに行があれば自動反映 */
function kyodoImportEnrollmentSmart_() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = kyodoGetOutputSheet_(ss);
    var isEmpty = !sh || sh.getLastRow() < 2;
    var rangeStart;
    var modeLabel;

    if (isEmpty) {
      var res = ui.alert(
        "6ヶ月割メール取込",
        "「" + KYODO_ENROLL_SHEET_NAME + "」が空です。\n2025/12/1 からの全件を取り込みますか？\n\nはい … 全件\nいいえ … 当月のみ",
        ui.ButtonSet.YES_NO_CANCEL
      );
      if (res === ui.Button.CANCEL) return;
      if (res === ui.Button.YES) {
        rangeStart = KYODO_ENROLL_RANGE_START;
        modeLabel = "全件";
      } else {
        var now = new Date();
        rangeStart = new Date(now.getFullYear(), now.getMonth(), 1);
        modeLabel = "当月";
      }
    } else {
      var now2 = new Date();
      rangeStart = new Date(now2.getFullYear(), now2.getMonth(), 1);
      modeLabel = "当月（追加分）";
    }

    var result = kyodoImportEnrollmentEmailsCore_(rangeStart, kyodoEndOfToday_());
    var lines = [
      modeLabel + "の取込完了",
      "期間: " + result.rangeLabel,
      "総数: " + result.count + " 件（入会メール1通＝1行）",
      "Gmailメッセージ: " + result.messageCount + " 通",
      "名前未取得: " + result.nameMissing + " 件",
      "スレッド返信など除外: " + result.skippedDuplicate + " 件"
    ];
    if (result.hitPageLimit) {
      lines.push("※ 一部チャンクで件数上限。再実行するか _取込ログ を確認");
    }

    var matchSh = ss.getSheetByName(KYODO_MATCH_SHEET_NAME);
    if (matchSh && matchSh.getLastRow() >= 2) {
      try {
        var hook = kyodoHookupEmailsCore_();
        lines.push("");
        lines.push("照合シートへ自動反映: 一致 " + hook.matched + " / 不一致 " + hook.unmatched + " / 要確認 " + hook.ambiguous);
      } catch (hookErr) {
        lines.push("");
        lines.push("照合の自動反映をスキップ: " + hookErr.message);
      }
    }

    ui.alert("6ヶ月割メール取込", lines.join("\n"), ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(e);
    ui.alert("6ヶ月割メール取込", "エラー: " + e.message, ui.ButtonSet.OK);
  }
}

/** メニュー: 要連絡者を「在籍期間」シートに一覧 */
function kyodoPrepareCampaignTargets_() {
  var ui = SpreadsheetApp.getUi();
  try {
    ui.alert(
      "キャンペーン対象を整理",
      "「" + KYODO_TENURE_SHEET_NAME + "」に、次の条件の方だけを一覧します。\n\n" +
        "・6ヶ月割入会（法人除く）\n" +
        "・2026年に退会手続き\n" +
        "・在籍月数が満6ヶ月未満（要連絡の方）\n\n" +
        "このシートの「" + KYODO_TENURE_SEND_TARGET_HEADER + "」にチェックを付けた人だけに送れます。",
      ui.ButtonSet.OK
    );

    var tenure = kyodoCheckWithdrawalTenure2026Core_();

    ui.alert(
      "完了",
      "キャンペーン要連絡: " + tenure.targetCount + " 件\n" +
        "（法人退会を除外: " + tenure.excludedCorporate + " 件 / " +
        "満6ヶ月以上を除外: " + tenure.excludedFullTenure + " 件）\n\n" +
        "「" + KYODO_TENURE_SHEET_NAME + "」をご確認ください。",
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log(e);
    ui.alert("キャンペーン対象を整理", "エラー: " + e.message, ui.ButtonSet.OK);
  }
}

/** メニュー: テスト下書き or 一斉送信 */
function kyodoSendCampaignMailMenu_() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert(
    "確認メール",
    "どちらにしますか？\n\n" +
      "はい … テスト（文面を編集 → " + MY_TEST_EMAIL + " に下書きのみ）\n" +
      "いいえ … 本番（文面を編集 → 送信対象チェックONの人のみ送信）",
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (res === ui.Button.YES) kyodoCreateDraftTest();
  else if (res === ui.Button.NO) kyodoSendCampaignMailToRed();
}

/** 在籍期間シートの送信対象件数（プレビュー用・チェックONかつ未送信） */
function kyodoGetTenureSendPreview_() {
  var list = kyodoLoadTenureSheetRecipients_();
  var pending = 0;
  var skipped = 0;
  var unchecked = 0;
  for (var i = 0; i < list.length; i++) {
    if (!list[i].sendTarget) {
      unchecked++;
      continue;
    }
    if (list[i].alreadySent) {
      skipped++;
    } else if (list[i].email) {
      pending++;
    }
  }
  return { total: list.length, pending: pending, skipped: skipped, unchecked: unchecked };
}

function kyodoImportEnrollmentEmails() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = kyodoImportEnrollmentEmailsCore_(KYODO_ENROLL_RANGE_START, kyodoEndOfToday_());
    var lines = [
      "全件取込完了",
      "期間: " + result.rangeLabel,
      "総数: " + result.count + " 件（入会メール1通＝1行）",
      "Gmailメッセージ: " + result.messageCount + " 通",
      "名前未取得: " + result.nameMissing + " 件",
      "スレッド返信など除外: " + result.skippedDuplicate + " 件"
    ];
    if (result.hitPageLimit) {
      lines.push("※ 一部チャンクで件数上限。再実行するか _取込ログ を確認");
    }
    ui.alert("入会メール取込（全件）", lines.join("\n"), ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(e);
    ui.alert("入会メール取込（全件）", "エラー: " + e.message, ui.ButtonSet.OK);
  }
}

function kyodoImportEnrollmentEmailsCurrentMonth() {
  var ui = SpreadsheetApp.getUi();
  try {
    var now = new Date();
    var startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    var endOfToday = kyodoEndOfToday_();
    
    var result = kyodoImportEnrollmentEmailsCore_(startOfMonth, endOfToday);
    var lines = [
      "当月分の取込完了",
      "期間: " + result.rangeLabel,
      "総数: " + result.count + " 件（入会メール1通＝1行）",
      "Gmailメッセージ: " + result.messageCount + " 通",
      "名前未取得: " + result.nameMissing + " 件",
      "スレッド返信など除外: " + result.skippedDuplicate + " 件"
    ];
    if (result.hitPageLimit) {
      lines.push("※ 一部チャンクで件数上限。再実行するか _取込ログ を確認");
    }
    ui.alert("入会メール取込（当月分）", lines.join("\n"), ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(e);
    ui.alert("入会メール取込（当月分）", "エラー: " + e.message, ui.ButtonSet.OK);
  }
}

function kyodoHookupEmailsByNameAndDatetime() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = kyodoHookupEmailsCore_();
    ui.alert(
      "照合",
      "反映完了\n一致: " + result.matched + " 件\n不一致: " + result.unmatched + " 件\n要確認: " + result.ambiguous + " 件",
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log(e);
    ui.alert("照合", "エラー: " + e.message, ui.ButtonSet.OK);
  }
}

function kyodoImportEnrollmentEmailsCore_(rangeStart, rangeEnd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = kyodoGetOutputSheet_(ss);
  if (!sh) throw new Error("シートが見つかりません。");

  var start = rangeStart || KYODO_ENROLL_RANGE_START;
  var end = rangeEnd || kyodoEndOfToday_();
  var fetchResult = kyodoFetchAllEnrollmentRows_(start, end);
  var rows = fetchResult.rows;

  rows.sort(function (a, b) {
    return b.date.getTime() - a.date.getTime();
  });

  var tz = Session.getScriptTimeZone();
  sh.clear();
  sh.getRange(1, 1, 1, 5).setValues([
    ["タイムスタンプ", "名前", "メールアドレス", "受付番号", "備考"]
  ]);
  sh.getRange(1, 1, 1, 5).setFontWeight("bold");

  var nameMissing = 0;
  if (rows.length) {
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      if (!rows[i].name) nameMissing++;
      out.push([
        Utilities.formatDate(rows[i].date, tz, "yyyy/MM/dd HH:mm"),
        rows[i].name,
        rows[i].email,
        rows[i].receptionNo,
        rows[i].note
      ]);
    }
    sh.getRange(2, 1, out.length, 5).setValues(out);
    sh.autoResizeColumns(1, 5);
  }

  kyodoWriteImportLog_(ss, fetchResult, rows.length, nameMissing, tz, start, end);

  return {
    count: rows.length,
    sheetName: sh.getName(),
    rangeLabel: kyodoFormatRangeLabel_(start, end, tz),
    messageCount: fetchResult.messageCount,
    nameMissing: nameMissing,
    skippedDuplicate: fetchResult.skippedDuplicate,
    hitPageLimit: fetchResult.hitPageLimit
  };
}

/** 月別チャンク × ページ送りで Gmail 全探索 */
function kyodoFetchAllEnrollmentRows_(rangeStart, rangeEnd) {
  var chunks = kyodoBuildMonthChunks_(rangeStart, rangeEnd);
  var seenMessageId = {};
  var seenEnrollmentKey = {};
  var rows = [];
  var messageCount = 0;
  var skippedDuplicate = 0;
  var hitPageLimit = false;
  var threadsScanned = 0;

  for (var c = 0; c < chunks.length; c++) {
    var chunk = chunks[c];
    var start = 0;
    while (true) {
      var query = kyodoBuildGmailSearchQuery_(chunk.after, chunk.before);
      var threads = GmailApp.search(query, start, KYODO_ENROLL_PAGE_SIZE);
      if (!threads.length) break;

      threadsScanned += threads.length;

      for (var t = 0; t < threads.length; t++) {
        var messages = threads[t].getMessages();
        for (var m = 0; m < messages.length; m++) {
          var msg = messages[m];
          if (!kyodoIsSixMonthEnrollmentMessage_(msg)) continue;

          var date = msg.getDate();
          if (!kyodoIsInEnrollRange_(date, rangeStart, rangeEnd)) continue;

          var msgId = msg.getId();
          if (seenMessageId[msgId]) continue;
          seenMessageId[msgId] = true;
          messageCount++;

          var email = kyodoGetRecipientEmail_(msg);
          if (!email) continue;

          var minuteKey =
            email +
            "\x1f" +
            Utilities.formatDate(
              kyodoTruncateToMinute_(date),
              Session.getScriptTimeZone(),
              "yyyy/MM/dd HH:mm"
            );
          if (seenEnrollmentKey[minuteKey]) {
            skippedDuplicate++;
            continue;
          }
          seenEnrollmentKey[minuteKey] = true;

          var name = kyodoExtractMemberName_(msg);
          var receptionNo = kyodoExtractReceptionNo_(msg);
          var note = "";
          if (!name) note = "名前未取得（要確認）";

          rows.push({
            date: date,
            name: name,
            email: email,
            receptionNo: receptionNo,
            note: note
          });
        }
      }

      start += threads.length;
      if (threads.length < KYODO_ENROLL_PAGE_SIZE) break;
      if (start >= 500) {
        hitPageLimit = true;
        Logger.log("Gmail search page limit at chunk " + chunk.label + " start=" + start);
        break;
      }
      Utilities.sleep(150);
    }
    Utilities.sleep(100);
  }

  return {
    rows: rows,
    messageCount: messageCount,
    skippedDuplicate: skippedDuplicate,
    hitPageLimit: hitPageLimit,
    threadsScanned: threadsScanned,
    chunks: chunks.length
  };
}

function kyodoBuildGmailSearchQuery_(afterDate, beforeDate) {
  var afterStr = Utilities.formatDate(afterDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
  var beforeStr = Utilities.formatDate(beforeDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
  return (
    "from:" +
    KYODO_ENROLL_FROM +
    ' subject:"' +
    KYODO_ENROLL_SUBJECT +
    '" after:' +
    afterStr +
    " before:" +
    beforeStr
  );
}

function kyodoBuildMonthChunks_(rangeStart, rangeEnd) {
  var chunks = [];
  var cur = new Date(rangeStart.getFullYear(), rangeStart.getMonth(), 1);
  var endLimit = new Date(rangeEnd.getFullYear(), rangeEnd.getMonth() + 1, 1);

  while (cur.getTime() < endLimit.getTime()) {
    var next = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
    var chunkStart = cur.getTime() < rangeStart.getTime() ? rangeStart : cur;
    var chunkEnd = next.getTime() > endLimit.getTime() ? endLimit : next;
    chunks.push({
      after: chunkStart,
      before: chunkEnd,
      label: Utilities.formatDate(cur, Session.getScriptTimeZone(), "yyyy/MM")
    });
    cur = next;
  }
  return chunks;
}

function kyodoIsInEnrollRange_(date, rangeStart, rangeEnd) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return false;
  var t = date.getTime();
  return t >= rangeStart.getTime() && t <= rangeEnd.getTime();
}

function kyodoEndOfToday_() {
  var now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
}

function kyodoFormatRangeLabel_(start, end, tz) {
  return (
    Utilities.formatDate(start, tz, "yyyy/MM/dd") +
    " ～ " +
    Utilities.formatDate(end, tz, "yyyy/MM/dd")
  );
}

function kyodoWriteImportLog_(ss, fetchResult, rowCount, nameMissing, tz, start, end) {
  var sh = ss.getSheetByName(KYODO_ENROLL_LOG_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(KYODO_ENROLL_LOG_SHEET_NAME);
  sh.clear();
  var now = new Date();
  var rangeStart = start || KYODO_ENROLL_RANGE_START;
  var rangeEnd = end || kyodoEndOfToday_();
  sh.getRange(1, 1, 8, 2).setValues([
    ["最終取込", Utilities.formatDate(now, tz, "yyyy/MM/dd HH:mm:ss")],
    ["対象期間", kyodoFormatRangeLabel_(rangeStart, rangeEnd, tz)],
    ["一覧行数（総数）", rowCount],
    ["Gmailメッセージ数", fetchResult.messageCount],
    ["名前未取得", nameMissing],
    ["重複除外", fetchResult.skippedDuplicate],
    ["探索スレッド数", fetchResult.threadsScanned],
    ["月別チャンク数", fetchResult.chunks]
  ]);
  try {
    sh.hideSheet();
  } catch (e) {}
}

function kyodoHookupEmailsCore_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var enrollSh = kyodoGetOutputSheet_(ss);
  if (!enrollSh || enrollSh.getLastRow() < 2) {
    throw new Error("先に「Gmailから取込（全件）」で「" + KYODO_ENROLL_SHEET_NAME + "」にデータを入れてください。");
  }

  var matchSh = kyodoEnsureMatchSheet_(ss);
  var lastRow = matchSh.getLastRow();
  if (lastRow < 2) {
    throw new Error("「" + KYODO_MATCH_SHEET_NAME + "」に名前と入会日時を入力してください。");
  }

  var index = kyodoBuildEnrollmentIndex_(enrollSh);
  var tz = Session.getScriptTimeZone();
  var height = lastRow - 1;
  var vals = matchSh.getRange(2, 1, height, 2).getValues();
  var emails = [];
  var statuses = [];
  var matched = 0;
  var unmatched = 0;
  var ambiguous = 0;

  for (var r = 0; r < vals.length; r++) {
    var name = kyodoNormalizePersonName_(vals[r][0]);
    var dt = kyodoParseSheetDateTime_(vals[r][1]);
    var emailOut = "";
    var status = "";

    if (!name) {
      status = "名前が空";
      unmatched++;
    } else if (!dt) {
      status = "入会日時が読めない";
      unmatched++;
    } else {
      var key = kyodoMakeLookupKey_(name, dt, tz);
      var hit = index[key];
      if (!hit) {
        var nameOnly = index["name:" + name];
        if (nameOnly && nameOnly.count > 0) {
          status = "日時不一致（同名は別日時に存在）";
          ambiguous++;
        } else {
          status = "該当なし";
          unmatched++;
        }
      } else if (hit.ambiguous) {
        status = "複数候補（要手動確認）";
        emailOut = hit.email || "";
        ambiguous++;
      } else {
        emailOut = hit.email;
        status = "一致";
        matched++;
      }
    }

    emails.push([emailOut]);
    statuses.push([status]);
  }

  matchSh.getRange(2, 3, height, 1).setValues(emails);
  matchSh.getRange(2, 4, height, 1).setValues(statuses);
  matchSh.autoResizeColumns(1, 4);

  return { matched: matched, unmatched: unmatched, ambiguous: ambiguous };
}

function kyodoBuildEnrollmentIndex_(enrollSh) {
  var tz = Session.getScriptTimeZone();
  var data = enrollSh.getDataRange().getValues();
  var index = {};

  for (var i = 1; i < data.length; i++) {
    var dt = kyodoParseSheetDateTime_(data[i][0]);
    var name = kyodoNormalizePersonName_(data[i][1]);
    var email = String(data[i][2] || "").trim().toLowerCase();
    if (!name || !dt || !email) continue;

    var key = kyodoMakeLookupKey_(name, dt, tz);
    if (!index[key]) {
      index[key] = { email: email, count: 0, ambiguous: false };
    }
    index[key].count++;
    if (index[key].count === 1) {
      index[key].email = email;
    } else if (index[key].email !== email) {
      index[key].ambiguous = true;
    }

    if (!index["name:" + name]) {
      index["name:" + name] = { count: 0 };
    }
    index["name:" + name].count++;
  }

  return index;
}

function kyodoMakeLookupKey_(name, date, tz) {
  return name + "\x1f" + Utilities.formatDate(kyodoTruncateToMinute_(date), tz, "yyyy/MM/dd HH:mm");
}

function kyodoEnsureMatchSheet_(ss) {
  var sh = ss.getSheetByName(KYODO_MATCH_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(KYODO_MATCH_SHEET_NAME);
  }
  var header = sh.getRange(1, 1, 1, 4).getValues()[0];
  if (!header[0]) {
    sh.getRange(1, 1, 1, 4).setValues([["名前", "入会日時", "メールアドレス", "照合結果"]]);
    sh.getRange(1, 1, 1, 4).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function kyodoParseSheetDateTime_(v) {
  if (v === "" || v === null || v === undefined) return null;
  if (v instanceof Date && !isNaN(v.getTime())) {
    return kyodoTruncateToMinute_(v);
  }
  var s = String(v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})[ T]?(\d{1,2})?:?(\d{1,2})?/);
  if (m) {
    var hh = m[4] !== undefined && m[4] !== "" ? parseInt(m[4], 10) : 0;
    var mm = m[5] !== undefined && m[5] !== "" ? parseInt(m[5], 10) : 0;
    return kyodoTruncateToMinute_(
      new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10), hh, mm)
    );
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) return kyodoTruncateToMinute_(d);
  return null;
}

function kyodoTruncateToMinute_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes());
}

function kyodoGetOutputSheet_(ss) {
  return (
    ss.getSheetByName(KYODO_ENROLL_SHEET_NAME) ||
    ss.getSheetByName("シート1") ||
    ss.getSheetByName("Sheet1") ||
    ss.getSheets()[0]
  );
}

function kyodoIsEnrollmentMessage_(msg) {
  var subj = String(msg.getSubject() || "").trim();
  if (subj !== KYODO_ENROLL_SUBJECT) return false;
  var from = kyodoParseEmailAddress_(msg.getFrom());
  if (from === KYODO_ENROLL_FROM.toLowerCase()) return true;
  return String(msg.getFrom() || "").indexOf("joyfit-service") !== -1;
}

/** 6ヶ月割入会メール（法人を除く） */
function kyodoIsSixMonthEnrollmentMessage_(msg) {
  if (!kyodoIsEnrollmentMessage_(msg)) return false;
  var body = kyodoGetMessageBodyText_(msg);
  if (body.indexOf("法人") !== -1) return false;
  return true;
}

/** 一般会員の退会メール（法人退会を除く） */
function kyodoIsGeneralWithdrawalMessage_(msg) {
  var body = kyodoGetMessageBodyText_(msg);
  if (body.indexOf("ご退会のお手続きが完了") !== -1) return false;
  return true;
}

function kyodoGetRecipientEmail_(msg) {
  var fields = [msg.getTo(), msg.getCc(), msg.getBcc()];
  for (var i = 0; i < fields.length; i++) {
    var emails = kyodoParseAllEmails_(fields[i]);
    for (var j = 0; j < emails.length; j++) {
      if (emails[j].indexOf("joyfit") === -1 && emails[j].indexOf("okamoto-group") === -1) {
        return emails[j];
      }
    }
  }
  var toOnly = kyodoParseEmailAddress_(msg.getTo());
  return toOnly || "";
}

function kyodoParseEmailAddress_(raw) {
  var list = kyodoParseAllEmails_(raw);
  return list.length ? list[0] : "";
}

function kyodoParseAllEmails_(raw) {
  var s = String(raw || "").trim();
  if (!s) return [];
  var out = [];
  var re = /[\w.+-]+@[\w.-]+\.\w+/g;
  var m;
  while ((m = re.exec(s)) !== null) {
    out.push(m[0].toLowerCase());
  }
  return out;
}

function kyodoGetMessageBodyText_(msg) {
  var plain = msg.getPlainBody();
  if (plain && String(plain).trim()) return String(plain);
  var html = msg.getBody();
  if (!html) return "";
  return String(html)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .trim();
}

function kyodoExtractReceptionNo_(msg) {
  var body = kyodoGetMessageBodyText_(msg);
  var m = body.match(/受付番号\s*[:：]\s*(\d+)/);
  return m ? m[1] : "";
}

function kyodoExtractMemberName_(msg) {
  var body = kyodoGetMessageBodyText_(msg);

  var patterns = [
    /お名前\s*[:：]\s*([^\n\r]+)/,
    /氏名\s*[:：]\s*([^\n\r]+)/,
    /会員名\s*[:：]\s*([^\n\r]+)/
  ];
  for (var p = 0; p < patterns.length; p++) {
    var m = body.match(patterns[p]);
    if (m) {
      var n = kyodoNormalizePersonName_(m[1]);
      if (n) return n;
    }
  }

  var lines = body.split(/\r?\n/);
  for (var i = 0; i < Math.min(lines.length, 15); i++) {
    var line = String(lines[i] || "").trim();
    if (!line) continue;
    if (/^(https?|ー|－|－－|【|■|●|この度|ご利用|JOYFIT)/.test(line)) continue;
    var mTop = line.match(/^(.{1,40}?)\s*様\s*$/);
    if (mTop) {
      var n2 = kyodoNormalizePersonName_(mTop[1]);
      if (n2 && !/[@＠]/.test(n2)) return n2;
    }
  }

  return "";
}

function kyodoNormalizePersonName_(raw) {
  var s = String(raw || "")
    .replace(/\s*様\s*$/g, "")
    .replace(/[ \u3000]+/g, " ")
    .replace(/^[0-9０-９\s]+/, "")
    .trim();
  if (!s) return "";
  if (/[@＠]/.test(s) || /^https?:/i.test(s)) return "";

  var parts = s.split(" ");
  if (parts.length >= 2 && parts.length % 2 === 0) {
    var half = parts.length / 2;
    var a = parts.slice(0, half).join(" ");
    var b = parts.slice(half).join(" ");
    if (a === b) return a;
  }
  return s;
}

// --- 一斉送信・下書き作成（ダイアログ編集機能つき） ---

var KYODO_DEFAULT_SUBJECT = "【重要】JOYFIT24経堂より：キャンペーン在籍期間に関するご確認";
var KYODO_CAMPAIGN_SEND_LOG_HEADER = "送信日時";
var KYODO_CAMPAIGN_SEND_DELAY_MS = 1200;

/** テスト送信（自分のアドレス宛に下書きを作成） */
function kyodoCreateDraftTest() {
  kyodoShowMailEditorDialog_('test');
}

/** 「在籍期間」シートの全員にHTMLメール一斉送信 */
function kyodoSendCampaignMailToRed() {
  kyodoShowMailEditorDialog_('red');
}

/** 編集ダイアログを表示する関数 */
function kyodoShowMailEditorDialog_(targetMode) {
  var subject = KYODO_DEFAULT_SUBJECT;
  var body = kyodoGetDefaultMailBody_();
  var title, btnText, countNote = "";

  if (targetMode === 'test') {
    title = 'テスト下書きの作成';
    btnText = '下書きを作成する';
    countNote = "📬 テスト送信先: " + MY_TEST_EMAIL + "（下書きのみ・会員には送りません）";
  } else if (targetMode === 'red') {
    title = '一斉送信の確認と編集';
    btnText = '送信内容を確定する';
    try {
      var preview = kyodoGetTenureSendPreview_();
      countNote =
        "📬 送信対象（チェックON）: " + preview.pending + " 名" +
        "（全 " + preview.total + " 名 / 未チェック " + preview.unchecked + " / 送信済み " + preview.skipped + "）";
    } catch (e) {
      countNote = "📬 送信対象: 「" + KYODO_TENURE_SHEET_NAME + "」シートを確認してください";
    }
  }
  
  var html = `<!DOCTYPE html>
  <html>
  <head>
    <base target="_top">
    <style>
      body { font-family: 'Meiryo', sans-serif; padding: 15px; color: #333; }
      label { font-weight: bold; display: block; margin-bottom: 5px; font-size: 14px; }
      input[type="text"], textarea { width: 100%; box-sizing: border-box; padding: 10px; border: 1px solid #ccc; border-radius: 4px; margin-bottom: 15px; font-family: inherit; font-size: 14px; }
      textarea { height: 280px; resize: vertical; line-height: 1.5; }
      .note { font-size: 12px; color: #555; margin-bottom: 15px; background: #e8f0fe; padding: 10px; border-radius: 4px; border-left: 4px solid #1a73e8; }
      .note code { background: #fff; padding: 2px 4px; border-radius: 2px; font-weight: bold; color: #c62828; }
      .actions { text-align: right; margin-top: 20px; }
      button { padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: bold; }
      .btn-submit { background-color: #d32f2f; color: white; margin-left: 10px; }
      .btn-cancel { background-color: #f1f3f4; color: #333; }
      button:disabled { background-color: #ccc; cursor: not-allowed; }
    </style>
  </head>
  <body>
    <div class="note">
      ${kyodoEscapeHtml_(countNote)}<br><br>
      💡 本文中の <code>{name}</code> は、送信時に各会員のお名前に置き換わります。
    </div>
    
    <label>件名</label>
    <input type="text" id="subject" value="${kyodoEscapeHtml_(subject)}">
    
    <label>本文</label>
    <textarea id="body">${kyodoEscapeHtml_(body)}</textarea>
    
    <div class="actions">
      <button class="btn-cancel" onclick="google.script.host.close()">キャンセル</button>
      <button id="submitBtn" class="btn-submit" onclick="submitForm()">${kyodoEscapeHtml_(btnText)}</button>
    </div>
    
    <script>
      function submitForm() {
        var btn = document.getElementById('submitBtn');
        btn.disabled = true;
        btn.innerText = "処理中...（画面を閉じないで）";
        
        var subj = document.getElementById('subject').value;
        var body = document.getElementById('body').value;
        
        google.script.run
          .withSuccessHandler(function(res) {
            alert(res.message);
            google.script.host.close();
          })
          .withFailureHandler(function(err) {
            alert("エラーが発生しました:\\n" + err.message);
            btn.disabled = false;
            btn.innerText = "${kyodoEscapeHtml_(btnText)}";
          })
          .kyodoExecuteSendMail('${kyodoEscapeHtml_(targetMode)}', subj, body);
      }
    </script>
  </body>
  </html>`;
  
  var uiOutput = HtmlService.createHtmlOutput(html)
      .setWidth(650)
      .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(uiOutput, title);
}

/** ダイアログから呼び出されて実際の送信を行う関数 */
function kyodoExecuteSendMail(targetMode, subject, bodyTemplate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (targetMode === 'test') {
    var to = MY_TEST_EMAIL;
    var name = "テスト 太郎";
    kyodoSendOneCampaignMail_(to, name, true, subject, bodyTemplate);
    return { message: "以下の宛先でGmailの「下書き」を作成しました:\n" + to + "\n\nご自身のGmailの「下書き」フォルダを確認してください。" };
    
  } else if (targetMode === 'red') {
    var list = kyodoLoadTenureSheetRecipients_();
    if (!list.length) throw new Error("「" + KYODO_TENURE_SHEET_NAME + "」に送信先がありません。");

    var pending = [];
    for (var i = 0; i < list.length; i++) {
      if (!list[i].sendTarget) continue;
      if (!list[i].alreadySent && list[i].email) pending.push(list[i]);
    }
    if (!pending.length) {
      throw new Error("送信対象にチェックが付いている未送信の行がありません。\n「" + KYODO_TENURE_SHEET_NAME + "」の「" + KYODO_TENURE_SEND_TARGET_HEADER + "」列を確認してください。");
    }

    var ui = SpreadsheetApp.getUi();
    var confirm = ui.alert(
      "送信の最終確認",
      "件名・本文の内容で、チェックONの未送信 " + pending.length + " 名にメールを送ります。\n\n本当に送信しますか？",
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) {
      return { message: "送信をキャンセルしました。" };
    }

    var result = kyodoSendCampaignMailBatch_(ss, list, KYODO_TENURE_SHEET_NAME, {
      isTest: false,
      skipAlreadySent: true,
      requireSendTarget: true,
      subject: subject,
      bodyTemplate: bodyTemplate
    });
    
    return { message: "一斉送信が完了しました。\n\n対象: " + result.sent + " 件\n失敗: " + result.failed + " 件\nスキップ: " + result.skipped + " 件" };
  }
}

function kyodoSendCampaignMailBatch_(ss, list, sheetName, options) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("シート「" + sheetName + "」が見つかりません。");

  var opts = options || {};
  var isTest = !!opts.isTest;
  var skipAlreadySent = opts.skipAlreadySent !== false;
  var requireSendTarget = !!opts.requireSendTarget;
  var subject = opts.subject || KYODO_DEFAULT_SUBJECT;
  var bodyTemplate = opts.bodyTemplate || kyodoGetDefaultMailBody_();

  var sent = 0;
  var failed = 0;
  var skipped = 0;
  var logCol = kyodoEnsureCampaignSendLogColumn_(sh);
  var tz = Session.getScriptTimeZone();

  for (var i = 0; i < list.length; i++) {
    var row = list[i];
    if (requireSendTarget && !row.sendTarget) {
      skipped++;
      continue;
    }
    if (skipAlreadySent && row.alreadySent) {
      skipped++;
      continue;
    }
    if (!row.email) {
      failed++;
      if (logCol > 0) {
        sh.getRange(row.sheetRow, logCol).setValue("失敗: メールなし");
      }
      continue;
    }
    try {
      kyodoSendOneCampaignMail_(row.email, row.name, isTest, subject, bodyTemplate);
      sent++;
      if (logCol > 0) {
        sh.getRange(row.sheetRow, logCol).setValue(
          Utilities.formatDate(new Date(), tz, "yyyy/MM/dd HH:mm") + (isTest ? "（下書き作成）" : "")
        );
      }
      Utilities.sleep(KYODO_CAMPAIGN_SEND_DELAY_MS);
    } catch (err) {
      failed++;
      Logger.log(err);
      if (logCol > 0) {
        sh.getRange(row.sheetRow, logCol).setValue("失敗: " + String(err.message || err));
      }
    }
  }
  return { sent: sent, failed: failed, skipped: skipped };
}

function kyodoSendOneCampaignMail_(to, name, isDraft, subject, bodyTemplate) {
  var displayName = String(name || "").trim() || "お客";
  
  // テンプレート内の {name} を実際の名前に置換
  var replacedBody = bodyTemplate.replace(/{name}/g, displayName);
  
  // HTML用のリッチな枠に流し込む
  var html = kyodoWrapHtmlEmail_(replacedBody);
  
  var finalSubject = subject;
  if (isDraft) {
    finalSubject = "【テスト下書き】" + subject;
    GmailApp.createDraft(to, finalSubject, replacedBody, {
      htmlBody: html
    });
  } else {
    MailApp.sendEmail({
      to: to,
      subject: finalSubject,
      body: replacedBody,
      htmlBody: html
    });
  }
}

/** ダイアログを開いたときの「初期テキスト」 */
function kyodoGetDefaultMailBody_() {
  var lines = [];
  lines.push("{name} 様");
  lines.push("");
  lines.push("いつもJOYFIT24経堂をご利用いただき、誠にありがとうございます。");
  lines.push("");
  lines.push("本日は、ご入会時に適用されました「満6ヶ月キャンペーン」の在籍条件につきまして、大切なお願いがありご連絡いたしました。");
  lines.push("適用させていただいたキャンペーンは「満6ヶ月間のご在籍」が必須条件となっております。");
  lines.push("今回は期間途中でのご解約となりますため、恐れ入りますが下記【A】または【B】のいずれかをご選択いただきたく存じます。");
  lines.push("お手数ですが、どちらがご希望かお教えください。");
  lines.push("");
  lines.push("【A】退会を確定し、解約金を支払う");
  lines.push("キャンペーン時の値引き相当額（途中解約金）をお支払いいただきます。");
  lines.push("");
  lines.push("【B】退会をキャンセルし、必須期間まで会員を継続する");
  lines.push("一度退会をキャンセルし、必須在籍期間まで会員としてご継続いただきます。");
  lines.push("");
  lines.push("お手数ですが、本メール受信より5日以内にご返信いただけますと幸いです。");
  lines.push("");
  lines.push("■ ご返信のお願い");
  lines.push("・【A】または【B】のどちらがご希望か、本メールへご返信ください。");
  lines.push("・ご不明点がございましたら、同じく本メールよりお気軽にご連絡ください。");
  lines.push("");
  lines.push("お客様にはお手数をおかけいたしますが、何卒ご理解とご協力のほどよろしくお願い申し上げます。");
  lines.push("");
  lines.push("JOYFIT24経堂");
  return lines.join("\n");
}

/** 編集されたテキストを自動でリッチなHTMLの枠（ヘッダー付き）にはめ込む */
function kyodoWrapHtmlEmail_(bodyText) {
  // 改行を <br> に変換
  var escapedBody = kyodoEscapeHtml_(bodyText).replace(/\n/g, "<br>");
  
  return (
    "<div style=\"background:#f5f5f5;padding:24px 0;font-family:'Meiryo','Hiragino Kaku Gothic ProN',Arial,sans-serif;color:#333;line-height:1.75;\">" +
    "<div style=\"max-width:640px;margin:0 auto;background:#fff;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;\">" +
    "<div style=\"background:linear-gradient(90deg,#c62828 0%,#e91e63 100%);padding:16px 20px;color:#fff;font-size:17px;font-weight:bold;\">" +
    "JOYFIT24経堂" +
    "</div>" +
    "<div style=\"padding:24px 20px;font-size:14px;\">" +
    escapedBody +
    "</div></div></div>"
  );
}

function kyodoEscapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function kyodoLoadTenureSheetRecipients_() {
  return kyodoLoadSheetRecipients_(KYODO_TENURE_SHEET_NAME, "「" + KYODO_TENURE_SHEET_NAME + "」にデータがありません。");
}

/** 指定シートから名前・メールを読み込む（2行目以降・メールがある行のみ） */
function kyodoLoadSheetRecipients_(sheetName, emptyError, options) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) {
    throw new Error(emptyError || "シートにデータがありません。");
  }

  var opts = options || {};
  var data = sh.getDataRange().getValues();
  var header = data[0];
  var nameCol = kyodoFindHeaderColIndex_(header, ["名前"]);
  var emailCol = kyodoFindHeaderColIndex_(header, ["メールアドレス", "メール"]);
  var sentCol = kyodoFindHeaderColIndex_(header, [KYODO_CAMPAIGN_SEND_LOG_HEADER, "送信済"]);
  var sendTargetCol = kyodoFindHeaderColIndex_(header, [KYODO_TENURE_SEND_TARGET_HEADER]);
  if (nameCol < 0) nameCol = 1;
  if (emailCol < 0) emailCol = 2;

  var list = [];
  for (var r = 1; r < data.length; r++) {
    var name = kyodoNormalizePersonName_(data[r][nameCol]);
    var email = kyodoParseEmailAddress_(data[r][emailCol]);
    if (!email) continue;

    var sendTarget = sendTargetCol >= 0 ? data[r][sendTargetCol] === true : false;

    var alreadySent = false;
    if (sentCol >= 0) {
      var logVal = String(data[r][sentCol] || "").trim();
      if (logVal && logVal.indexOf("失敗") !== 0) {
        alreadySent = true;
      }
    }
    list.push({
      sheetRow: r + 1,
      name: name || "お客様",
      email: email,
      sendTarget: sendTarget,
      alreadySent: alreadySent
    });
  }
  return list;
}

/** 再整理時に送信対象・送信日時をメールアドレスで引き継ぐ */
function kyodoReadTenureSheetState_(sh) {
  var state = { byEmail: {} };
  if (!sh || sh.getLastRow() < 2) return state;

  var data = sh.getDataRange().getValues();
  var header = data[0];
  var emailCol = kyodoFindHeaderColIndex_(header, ["メールアドレス", "メール"]);
  var targetCol = kyodoFindHeaderColIndex_(header, [KYODO_TENURE_SEND_TARGET_HEADER]);
  var sentCol = kyodoFindHeaderColIndex_(header, [KYODO_CAMPAIGN_SEND_LOG_HEADER, "送信済"]);
  if (emailCol < 0) emailCol = 2;

  for (var r = 1; r < data.length; r++) {
    var email = String(data[r][emailCol] || "").trim().toLowerCase();
    if (!email) continue;
    state.byEmail[email] = {
      sendTarget: targetCol >= 0 ? data[r][targetCol] === true : false,
      sendLog: sentCol >= 0 ? String(data[r][sentCol] || "") : ""
    };
  }
  return state;
}

function kyodoFindHeaderColIndex_(headerRow, labels) {
  for (var c = 0; c < headerRow.length; c++) {
    var h = String(headerRow[c] || "").trim();
    for (var i = 0; i < labels.length; i++) {
      if (h === labels[i]) return c;
    }
  }
  return -1;
}

function kyodoEnsureCampaignSendLogColumn_(sh) {
  var lastCol = sh.getLastColumn();
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = kyodoFindHeaderColIndex_(header, [KYODO_CAMPAIGN_SEND_LOG_HEADER]);
  if (idx >= 0) return idx + 1;
  var newCol = lastCol + 1;
  sh.getRange(1, newCol).setValue(KYODO_CAMPAIGN_SEND_LOG_HEADER);
  sh.getRange(1, newCol).setFontWeight("bold");
  return newCol;
}

// --- 以下、退会者在籍期間チェック（2026年）用機能 ---

/** 2026年の入会メールと退会メールを照合し、在籍期間を計算してシートに出力する */
function kyodoCheckWithdrawalTenure2026() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var enrollListSh = ss.getSheetByName("入会者");
    var startMsg = "6ヶ月割の退会者のうち、満6ヶ月未満の方だけを一覧します。\n（法人退会は含みません）\n";
    if (!enrollListSh) {
      startMsg += "\n※「入会者」シートが見つからないため、会員番号は空欄になります。\n";
    }
    startMsg += "\n※メール件数が多い場合、数分かかることがあります。";
    ui.alert("処理開始", startMsg, ui.ButtonSet.OK);

    var result = kyodoCheckWithdrawalTenure2026Core_();
    ui.alert(
      "完了",
      "キャンペーン要連絡: " + result.targetCount + " 件\n\n" +
        "シート「" + KYODO_TENURE_SHEET_NAME + "」をご確認ください。\n" +
        "（6ヶ月割・満6ヶ月未満・法人除く）",
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log(e);
    ui.alert("エラー", e.message, ui.ButtonSet.OK);
  }
}

/** 在籍期間チェック本体（UIなし） */
function kyodoCheckWithdrawalTenure2026Core_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var start2026 = new Date(2026, 0, 1);
  var end2026 = new Date(2026, 11, 31, 23, 59, 59);
  var enrollStart = KYODO_ENROLL_RANGE_START;
  var enrollEnd = kyodoEndOfToday_();

  var memberNoMap = {};
  var enrollListSh = ss.getSheetByName("入会者");
  if (enrollListSh) {
    var enrollData = enrollListSh.getDataRange().getValues();
    if (enrollData.length > 1) {
      var header = enrollData[0];
      var noCol = kyodoFindHeaderColIndex_(header, ["会員番号"]);
      var nameCol = kyodoFindHeaderColIndex_(header, ["会員氏名", "氏名", "名前"]);
      var emailCol = kyodoFindHeaderColIndex_(header, ["メールアドレス", "メール", "E-Mail", "Email"]);
      if (noCol < 0) noCol = 5;
      if (nameCol < 0) nameCol = 6;

      for (var r = 1; r < enrollData.length; r++) {
        var mName = kyodoNormalizePersonName_(enrollData[r][nameCol]);
        var mNo = String(enrollData[r][noCol] || "").trim();
        var mEmail = emailCol >= 0 ? String(enrollData[r][emailCol] || "").trim().toLowerCase() : "";

        if (mNo) {
          if (mName) {
            var keyName = mName.replace(/\s+/g, "");
            memberNoMap["NAME:" + keyName] = mNo;
          }
          if (mEmail) {
            memberNoMap["EMAIL:" + mEmail] = mNo;
          }
        }
      }
    }
  }

  var enrollResult = kyodoFetchAllEnrollmentRows_(enrollStart, enrollEnd);
  var enrollRows = enrollResult.rows;
  var withdrawRows = kyodoFetchWithdrawalRows_(start2026, end2026);

  var enrollMap = {};
  for (var i = 0; i < enrollRows.length; i++) {
    var email = enrollRows[i].email.toLowerCase();
    if (!email) continue;
    if (!enrollMap[email]) {
      enrollMap[email] = [];
    }
    enrollMap[email].push(enrollRows[i]);
  }

  var resultData = [];
  var excludedCorporate = 0;
  var excludedFullTenure = 0;

    for (var j = 0; j < withdrawRows.length; j++) {
      var wRow = withdrawRows[j];
      if (wRow.isCorporate) {
        excludedCorporate++;
        continue;
      }
      var email = wRow.email.toLowerCase();
      if (!email || !enrollMap[email]) continue;

      var validEnrolls = [];
      for (var k = 0; k < enrollMap[email].length; k++) {
        if (enrollMap[email][k].date <= wRow.date) {
          validEnrolls.push(enrollMap[email][k]);
        }
      }

      if (validEnrolls.length > 0) {
        validEnrolls.sort(function(a, b) {
          return b.date.getTime() - a.date.getTime();
        });
        
        var eRow = validEnrolls[0];
        var months = (wRow.date.getFullYear() - eRow.date.getFullYear()) * 12 + (wRow.date.getMonth() - eRow.date.getMonth());

        if (months >= KYODO_TENURE_MIN_MONTHS) {
          excludedFullTenure++;
          continue;
        }

        var displayName = eRow.name || wRow.name;
        var normName = kyodoNormalizePersonName_(displayName);
        var keyName = normName.replace(/\s+/g, "");
        var memberNo = memberNoMap["EMAIL:" + email] || memberNoMap["NAME:" + keyName] || "";

        resultData.push([
          memberNo,
          displayName,
          email,
          Utilities.formatDate(eRow.date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm"),
          Utilities.formatDate(wRow.date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm"),
          months
        ]);
      }
    }

    var sh = ss.getSheetByName(KYODO_TENURE_SHEET_NAME);
    var prevState = kyodoReadTenureSheetState_(sh);
    if (!sh) {
      sh = ss.insertSheet(KYODO_TENURE_SHEET_NAME);
    } else {
      sh.clear();
    }

    sh.getRange(1, 1, 1, 8).setValues([[
      "会員番号", "名前", "メールアドレス", "入会日", "退会手続き日",
      "在籍月数", KYODO_TENURE_SEND_TARGET_HEADER, KYODO_CAMPAIGN_SEND_LOG_HEADER
    ]]);
    sh.getRange(1, 1, 1, 8).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.getRange(1, 9).setValue("※G列にチェック＝送信する人").setFontColor("#666666").setFontSize(9);

    if (resultData.length > 0) {
      resultData.sort(function(a, b) {
        return a[4] < b[4] ? 1 : -1;
      });
      var outRows = resultData.map(function(row) {
        var email = String(row[2] || "").trim().toLowerCase();
        var prev = prevState.byEmail[email];
        return row.concat([
          prev && prev.sendTarget === true,
          prev ? prev.sendLog : ""
        ]);
      });
      sh.getRange(2, 1, outRows.length, 8).setValues(outRows);
      sh.getRange(2, 7, outRows.length, 1).insertCheckboxes();
      sh.autoResizeColumns(1, 8);
    }

  return {
    targetCount: resultData.length,
    matchedCount: resultData.length,
    excludedCorporate: excludedCorporate,
    excludedFullTenure: excludedFullTenure,
    targets: resultData
  };
}

/** 月別チャンク × ページ送りで Gmail から退会メールのみを取得 */
function kyodoFetchWithdrawalRows_(rangeStart, rangeEnd) {
  var chunks = kyodoBuildMonthChunks_(rangeStart, rangeEnd);
  var seenMessageId = {};
  var rows = [];

  for (var c = 0; c < chunks.length; c++) {
    var chunk = chunks[c];
    var start = 0;
    while (true) {
      var afterStr = Utilities.formatDate(chunk.after, Session.getScriptTimeZone(), "yyyy/MM/dd");
      var beforeStr = Utilities.formatDate(chunk.before, Session.getScriptTimeZone(), "yyyy/MM/dd");
      
      // 退会メール検索クエリ（宛先ではなく差出人が店舗の想定）
      var query = "from:" + KYODO_ENROLL_FROM + ' subject:"' + KYODO_WITHDRAW_SUBJECT + '" after:' + afterStr + " before:" + beforeStr;
      var threads = GmailApp.search(query, start, KYODO_ENROLL_PAGE_SIZE);
      if (!threads.length) break;

      for (var t = 0; t < threads.length; t++) {
        var messages = threads[t].getMessages();
        for (var m = 0; m < messages.length; m++) {
          var msg = messages[m];
          var msgId = msg.getId();
          if (seenMessageId[msgId]) continue;
          seenMessageId[msgId] = true;

          var date = msg.getDate();
          if (!kyodoIsInEnrollRange_(date, rangeStart, rangeEnd)) continue;

          var isCorporate = !kyodoIsGeneralWithdrawalMessage_(msg);
          var email = kyodoGetRecipientEmail_(msg);
          if (!email) continue;

          var name = kyodoExtractMemberName_(msg);

          rows.push({
            date: date,
            name: name,
            email: email,
            isCorporate: isCorporate
          });
        }
      }

      start += threads.length;
      if (threads.length < KYODO_ENROLL_PAGE_SIZE) break;
      if (start >= 500) break; // 500件上限
      Utilities.sleep(150);
    }
    Utilities.sleep(100);
  }
  return rows;
}