/**
 * 8 大科學原理常態班 報名 → Google 試算表 接收程式
 * 使用方式見同資料夾「設定說明.md」
 */

const SHEET_NAME = '報名資料';
// 試算表 ID（網址 /d/ 與 /edit 中間那串）。填了就固定寫進這份試算表，不會寫錯檔；
// 留空字串則寫進「這個 Apps Script 所附屬的試算表」（從試算表的 擴充功能 → Apps Script 建立時）
const SHEET_ID = '';
// 早鳥以伺服器時間判斷（不信任訪客電腦的時間）；要改截止日，網站 index.html 的 CONFIG 也要一起改
const EARLY_BIRD_DEADLINE = new Date('2026-11-30T23:59:59+08:00');
const PRICE_ORIGINAL = 4800;
const PRICE_EARLY = 4200;
const CAPACITY = 12; // 每班人數上限；超過的整期報名自動標「候補」
// 有人報名時寄通知信給這個信箱；不需要就留空字串 ''
const NOTIFY_EMAIL = '';

const HEADERS = [
  '收到時間', '報名編號', '報名類型', '課程', '學生姓名', '年級', '就讀學校',
  '家長姓名', '家長電話', '緊急聯絡人', '緊急聯絡電話', '付款方式', '應繳金額', '早鳥',
  '得知管道', '備註', '瀏覽風格', '個資同意', '整期報名序', '處理狀態', '管理備註'
];
const KEYS = [
  null, 'refCode', 'type', 'course', 'studentName', 'grade', 'school',
  'parentName', 'parentPhone', 'emergencyName', 'emergencyPhone', 'payment', 'amount', 'earlyBird',
  'source', 'note', 'pageStyle', 'consent', null, null, null
];

function getSheet_() {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#06182F').setFontColor('#2FE6FF');
    sh.setFrozenRows(1);
    // 全部設為純文字，避免電話開頭的 0 被吃掉
    sh.getRange(2, 1, sh.getMaxRows() - 1, HEADERS.length).setNumberFormat('@');
  }
  return sh;
}

// 開頭加 ' 讓試算表一律當文字存：
// 1. 以 = + - @ 開頭的內容不會被當成公式執行
// 2. 0 開頭的數字（電話）不會被吃掉開頭的 0
function safe_(v) {
  const s = String(v == null ? '' : v).slice(0, 200);
  return /^[=+\-@]/.test(s) || /^0\d/.test(s) ? "'" + s : s;
}

// 統計整期報名：seats＝佔名額的人數，waiting＝還在排候補的人數
// 處理狀態含「取消」「測試」的列不算；處理狀態以「候補（」開頭的算候補，
// 你把候補的人改成「已錄取」等其他狀態後，他就改算佔名額
function countFull_(sh) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return { seats: 0, waiting: 0 };
  const typeCol = HEADERS.indexOf('報名類型'), statusCol = HEADERS.indexOf('處理狀態');
  let seats = 0, waiting = 0;
  sh.getRange(2, 1, n, HEADERS.length).getValues().forEach(r => {
    const st = String(r[statusCol]);
    if (r[typeCol] !== '報名整期' || /取消|測試/.test(st)) return;
    if (st.indexOf('候補（') === 0) waiting++; else seats++;
  });
  return { seats, waiting };
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (!data.studentName || !data.parentPhone) throw new Error('缺少必要欄位');
    const sh = getSheet_();
    let seq = '', status = '待聯繫';
    if (data.type === '報名整期') {
      const early = new Date() <= EARLY_BIRD_DEADLINE;
      data.earlyBird = early ? '是' : '否';
      data.amount = String(early ? PRICE_EARLY : PRICE_ORIGINAL);
      const c = countFull_(sh);
      seq = c.seats + c.waiting + 1;
      // 有空位而且沒有人在排候補，才直接給名額；否則排到候補最後面
      if (c.seats >= CAPACITY || c.waiting > 0) status = `候補（第 ${c.waiting + 1} 位）`;
    }
    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    const row = HEADERS.map((h, i) => {
      if (i === 0) return now;
      if (h === '整期報名序') return String(seq);
      if (h === '處理狀態') return status;
      return KEYS[i] ? safe_(data[KEYS[i]]) : '';
    });
    sh.appendRow(row);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL,
        `【新報名】8 大科學原理常態班｜${data.type}｜${data.studentName}${status !== '待聯繫' ? '｜' + status : ''}`,
        HEADERS.map((h, i) => `${h}：${row[i]}`).join('\n'));
    }
    return json_({ ok: true, refCode: data.refCode });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 用瀏覽器開啟部署網址時顯示，方便確認有部署成功
function doGet() {
  return json_({ ok: true, message: '8 大科學原理常態班報名接收程式運作中' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
