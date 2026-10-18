/**
 * 8 大科學原理常態班 報名 → Google 試算表 接收程式
 * 使用方式見同資料夾「設定說明.md」
 *
 * 網站會分兩次送資料進來：
 *   1. 報名表送出（action 空白）→ 新增一列，回傳是否候補
 *   2a. 報名整期：家長轉帳後回填末 5 碼（action = 'payment'）→ 寄「報名成功」給家長、寄通知給管理信箱
 *   2b. 試聽一堂：家長選方便聯繫的時間（action = 'trial'）→ 寄「預約成功」給家長、寄通知給管理信箱
 *   候補的人在第 1 步就直接寄「候補登記」給家長、寄通知給管理信箱，不會請家長轉帳
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
// 有人完成報名／預約時，寄通知給這個信箱；信件都由部署這支程式的帳號寄出
const NOTIFY_EMAIL = 'randdlaboratary@gmail.com';
const SENDER_NAME = '微研 Weiyan 科學教育';
const LINE_URL = 'https://lin.ee/rZW93KN';
const CLASS_TIME = '每週三 13:30–15:00（一期 8 堂，每堂 90 分鐘）';
const CLASS_PLACE = '強納森美語教室（高科大旁）';

const HEADERS = [
  '收到時間', '報名編號', '報名類型', '課程', '學生姓名', '年級', '就讀學校',
  '家長姓名', '家長電話', '家長Email', '緊急聯絡人', '緊急聯絡電話',
  '付款方式', '應繳金額', '早鳥', '轉帳後5碼', '方便聯繫時間',
  '得知管道', '備註', '不同意照片公開', '瀏覽風格', '個資同意',
  '整期報名序', '處理狀態', '確認信', '管理備註'
];
const KEYS = [
  null, 'refCode', 'type', 'course', 'studentName', 'grade', 'school',
  'parentName', 'parentPhone', 'parentEmail', 'emergencyName', 'emergencyPhone',
  'payment', 'amount', 'earlyBird', null, null,
  'source', 'note', 'photoOptOut', 'pageStyle', 'consent',
  null, null, null, null
];
const COL = name => HEADERS.indexOf(name) + 1; // 欄號（從 1 開始）
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const STATUS_WAIT_PAY = '待回報轉帳';
const STATUS_PAID = '已回報轉帳（待對帳）';
const STATUS_WAIT_TIME = '待選聯繫時間';
const STATUS_CONTACT = '待聯繫';

function spreadsheet_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_() {
  const ss = spreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  // 欄位跟這版程式不同（舊版留下的分頁）→ 改名保留，另開新分頁，不刪任何資料
  if (sh && sh.getLastRow() > 0) {
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (head.join('|') !== HEADERS.join('|')) {
      sh.setName(SHEET_NAME + '_舊版_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'MMdd-HHmm'));
      sh = null;
    }
  }
  if (!sh) sh = ss.insertSheet(SHEET_NAME, 0);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#06182F').setFontColor('#2FE6FF');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, sh.getMaxRows() - 1, HEADERS.length).setNumberFormat('@');
  }
  return sh;
}

// 開頭加 ' 讓試算表一律當文字存：
// 1. 以 = + - @ 開頭的內容不會被當成公式執行
// 2. 0 開頭的數字（電話、末 5 碼）不會被吃掉開頭的 0
function safe_(v) {
  const s = String(v == null ? '' : v).slice(0, 200);
  return /^[=+\-@]/.test(s) || /^0\d/.test(s) ? "'" + s : s;
}

// 寄信用：把使用者輸入的內容轉成純文字，避免被當成 HTML
function esc_(v) {
  return String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// 統計整期報名：seats＝佔名額的人數，waiting＝還在排候補的人數
// 處理狀態含「取消」「測試」的列不算；處理狀態以「候補（」開頭的算候補，
// 你把候補的人改成「已錄取」等其他狀態後，他就改算佔名額
function countFull_(sh) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return { seats: 0, waiting: 0 };
  const typeCol = COL('報名類型') - 1, statusCol = COL('處理狀態') - 1;
  let seats = 0, waiting = 0;
  sh.getRange(2, 1, n, HEADERS.length).getValues().forEach(r => {
    const st = String(r[statusCol]);
    if (r[typeCol] !== '報名整期' || /取消|測試/.test(st)) return;
    if (st.indexOf('候補（') === 0) waiting++; else seats++;
  });
  return { seats, waiting };
}

// 依報名編號找到那一列，回傳 { row: 列號, rec: {欄名: 值} }
function findRow_(sh, refCode) {
  const n = sh.getLastRow() - 1;
  if (n <= 0 || !refCode) return null;
  const vals = sh.getRange(2, 1, n, HEADERS.length).getValues();
  const i = vals.findIndex(r => String(r[COL('報名編號') - 1]) === String(refCode));
  if (i < 0) return null;
  const rec = {};
  HEADERS.forEach((h, k) => rec[h] = String(vals[i][k]));
  return { row: i + 2, rec };
}

function setCell_(sh, row, header, value) {
  sh.getRange(row, COL(header)).setValue(safe_(value));
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'payment') return json_(handlePayment_(data));
    if (data.action === 'trial') return json_(handleTrial_(data));
    return json_(handleRegister_(data));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// 第 1 步：報名表送出
function handleRegister_(data) {
  if (!data.studentName || !data.parentPhone || !data.parentName) throw new Error('缺少必要欄位');
  if (!EMAIL_RE.test(String(data.parentEmail || ''))) throw new Error('Email 格式不正確');
  const sh = getSheet_();
  const full = data.type === '報名整期';
  let seq = '', status = STATUS_WAIT_TIME, waitPos = 0;
  if (full) {
    const early = new Date() <= EARLY_BIRD_DEADLINE;
    data.earlyBird = early ? '是' : '否';
    data.amount = String(early ? PRICE_EARLY : PRICE_ORIGINAL);
    data.payment = '銀行轉帳';
    const c = countFull_(sh);
    seq = c.seats + c.waiting + 1;
    // 有空位而且沒有人在排候補，才直接給名額；否則排到候補最後面
    if (c.seats >= CAPACITY || c.waiting > 0) { waitPos = c.waiting + 1; status = `候補（第 ${waitPos} 位）`; }
    else status = STATUS_WAIT_PAY;
  } else {
    data.type = '試聽一堂';
    data.amount = '試聽（專人聯繫）';
    data.payment = data.earlyBird = '';
  }
  const row = HEADERS.map((h, i) => {
    if (i === 0) return now_();
    if (h === '整期報名序') return String(seq);
    if (h === '處理狀態') return status;
    return KEYS[i] ? safe_(data[KEYS[i]]) : '';
  });
  sh.appendRow(row);

  // 候補：不請家長轉帳，直接寄候補登記通知
  if (waitPos) {
    const rec = {};
    HEADERS.forEach((h, i) => rec[h] = String(row[i]).replace(/^'/, ''));
    mailParent_(rec, 'waitlist', { waitPos });
    mailAdmin_(rec, `【候補登記】第 ${waitPos} 位｜${rec['學生姓名']}`);
    setCell_(sh, sh.getLastRow(), '確認信', '已寄候補通知 ' + now_());
  }
  return { ok: true, refCode: data.refCode, type: data.type, waitlist: !!waitPos, waitPos, amount: data.amount, earlyBird: data.earlyBird };
}

// 第 2 步（整期）：回填轉帳末 5 碼
function handlePayment_(data) {
  const last5 = String(data.last5 || '').trim();
  if (!/^\d{5}$/.test(last5)) throw new Error('末 5 碼需為 5 位數字');
  const sh = getSheet_();
  const f = findRow_(sh, data.refCode);
  if (!f || f.rec['報名類型'] !== '報名整期') throw new Error('找不到這筆報名');
  if (f.rec['處理狀態'].indexOf('候補（') === 0) throw new Error('目前為候補，請先不要轉帳');
  setCell_(sh, f.row, '轉帳後5碼', last5);
  if (f.rec['處理狀態'] === STATUS_WAIT_PAY) setCell_(sh, f.row, '處理狀態', STATUS_PAID);
  f.rec['轉帳後5碼'] = last5;
  const first = !f.rec['確認信'];
  if (first) {
    mailParent_(f.rec, 'paid');
    mailAdmin_(f.rec, `【新報名】整期｜${f.rec['學生姓名']}｜末 5 碼 ${last5}`);
    setCell_(sh, f.row, '確認信', '已寄 ' + now_());
  } else {
    mailAdmin_(f.rec, `【更新末 5 碼】${f.rec['學生姓名']}｜${last5}`);
  }
  return { ok: true, refCode: data.refCode, mailed: first, email: f.rec['家長Email'] };
}

// 第 2 步（試聽）：選方便聯繫的時間
function handleTrial_(data) {
  const allowed = ['平日上午', '平日下午', '假日上午', '假日下午'];
  const times = (Array.isArray(data.contactTimes) ? data.contactTimes : []).filter(t => allowed.indexOf(t) >= 0);
  if (!times.length) throw new Error('請至少選一個方便聯繫的時間');
  const sh = getSheet_();
  const f = findRow_(sh, data.refCode);
  if (!f || f.rec['報名類型'] !== '試聽一堂') throw new Error('找不到這筆預約');
  const t = times.join('、');
  setCell_(sh, f.row, '方便聯繫時間', t);
  if (f.rec['處理狀態'] === STATUS_WAIT_TIME) setCell_(sh, f.row, '處理狀態', STATUS_CONTACT);
  f.rec['方便聯繫時間'] = t;
  const first = !f.rec['確認信'];
  if (first) {
    mailParent_(f.rec, 'trial');
    mailAdmin_(f.rec, `【試聽預約】${f.rec['學生姓名']}｜${t}`);
    setCell_(sh, f.row, '確認信', '已寄 ' + now_());
  }
  return { ok: true, refCode: data.refCode, mailed: first, email: f.rec['家長Email'] };
}

/* ================== 信件 ================== */

function card_(title, bodyHtml) {
  return `<div style="background:#f3f6fa;padding:24px 12px;font-family:'Noto Sans TC','PingFang TC','Microsoft JhengHei',sans-serif;color:#1b2f5e;">
  <div style="max-width:560px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;border:1px solid #dbe6f3;">
    <div style="background:#06182F;color:#2FE6FF;padding:18px 24px;font-size:13px;letter-spacing:.2em;font-weight:700;">WEIYAN SCIENCE LAB</div>
    <div style="padding:24px;">
      <h2 style="margin:0 0 14px;font-size:22px;color:#1b2f5e;">${title}</h2>
      ${bodyHtml}
      <p style="margin:22px 0 0;font-size:14px;color:#6b7a99;">有任何問題，歡迎直接回覆這封信，或加入 <a href="${LINE_URL}" style="color:#2F7DE0;">官方 LINE</a>。</p>
      <p style="margin:6px 0 0;font-size:14px;color:#6b7a99;">${SENDER_NAME}</p>
    </div>
  </div>
</div>`;
}

function table_(pairs) {
  return `<table style="width:100%;border-collapse:collapse;font-size:15px;margin:8px 0;">${pairs
    .filter(p => p[1] !== '' && p[1] != null)
    .map(p => `<tr><td style="padding:8px 10px;background:#f3f6fa;border:1px solid #e3ebf5;width:34%;color:#4e6390;white-space:nowrap;">${esc_(p[0])}</td><td style="padding:8px 10px;border:1px solid #e3ebf5;font-weight:700;">${esc_(p[1])}</td></tr>`)
    .join('')}</table>`;
}

function mailParent_(rec, kind, extra) {
  const to = rec['家長Email'];
  if (!EMAIL_RE.test(to)) return;
  const kid = rec['學生姓名'], hi = `<p style="margin:0 0 12px;font-size:16px;">${esc_(rec['家長姓名'])} 您好：</p>`;
  const amt = rec['應繳金額'] ? 'NT$' + Number(rec['應繳金額']).toLocaleString('en-US') + (rec['早鳥'] === '是' ? '（早鳥價）' : '') : '';
  let subject, html;
  if (kind === 'paid') {
    subject = `【報名成功】8 大科學原理常態班｜${kid}`;
    html = card_('報名成功！🎉', hi +
      `<p style="margin:0 0 12px;font-size:16px;line-height:1.7;">我們已收到 <b>${esc_(kid)}</b> 的報名與轉帳資訊。核對款項後，開課前會再以電話或 Email 通知上課細節。</p>` +
      table_([['報名編號', rec['報名編號']], ['課程', '8 大科學原理常態班'], ['上課時間', CLASS_TIME], ['上課地點', CLASS_PLACE], ['應繳金額', amt], ['轉帳末 5 碼', rec['轉帳後5碼']]]) +
      `<p style="margin:16px 0 0;font-size:17px;font-weight:700;color:#F68F09;">期待在課堂上見到 ${esc_(kid)}～</p>`);
  } else if (kind === 'trial') {
    subject = `【預約成功】試聽一堂｜${kid}`;
    html = card_('試聽預約成功！✨', hi +
      `<p style="margin:0 0 12px;font-size:16px;line-height:1.7;">我們已收到 <b>${esc_(kid)}</b> 的試聽預約，會在您方便的時段以電話聯繫，安排試聽日期。</p>` +
      table_([['報名編號', rec['報名編號']], ['方便聯繫時間', rec['方便聯繫時間']], ['聯絡電話', rec['家長電話']], ['上課時間', CLASS_TIME], ['上課地點', CLASS_PLACE]]) +
      `<p style="margin:16px 0 0;font-size:17px;font-weight:700;color:#F68F09;">期待相見～</p>`);
  } else {
    subject = `【候補登記】8 大科學原理常態班｜${kid}`;
    html = card_(`已列入候補（第 ${extra.waitPos} 位）`, hi +
      `<p style="margin:0 0 12px;font-size:16px;line-height:1.7;">本期名額目前已滿，<b>${esc_(kid)}</b> 已列入候補。<b>請先不要轉帳</b>，名額釋出時我們會依序以電話通知您。</p>` +
      table_([['報名編號', rec['報名編號']], ['候補順位', `第 ${extra.waitPos} 位`], ['上課時間', CLASS_TIME]]) +
      `<p style="margin:16px 0 0;font-size:17px;font-weight:700;color:#F68F09;">謝謝您的耐心，期待相見～</p>`);
  }
  MailApp.sendEmail({ to, subject, htmlBody: html, name: SENDER_NAME, replyTo: NOTIFY_EMAIL });
}

function mailAdmin_(rec, subject) {
  if (!NOTIFY_EMAIL) return;
  const amt = rec['應繳金額'] && /^\d+$/.test(rec['應繳金額']) ? 'NT$' + Number(rec['應繳金額']).toLocaleString('en-US') + (rec['早鳥'] === '是' ? '（早鳥）' : '') : rec['應繳金額'];
  const html = card_(esc_(subject.replace(/^【[^】]+】/, '')), table_([
    ['報名類型', rec['報名類型']], ['處理狀態', rec['處理狀態'] === STATUS_WAIT_PAY ? STATUS_PAID : (rec['處理狀態'] === STATUS_WAIT_TIME ? STATUS_CONTACT : rec['處理狀態'])],
    ['學生', `${rec['學生姓名']}（${rec['年級']}・${rec['就讀學校']}）`],
    ['家長', rec['家長姓名']], ['家長電話', rec['家長電話']], ['家長 Email', rec['家長Email']],
    ['緊急聯絡', `${rec['緊急聯絡人']}　${rec['緊急聯絡電話']}`],
    ['應繳金額', amt], ['轉帳末 5 碼', rec['轉帳後5碼']], ['方便聯繫時間', rec['方便聯繫時間']],
    ['整期報名序', rec['整期報名序']], ['得知管道', rec['得知管道']], ['備註', rec['備註']],
    ['照片公開', rec['不同意照片公開'] ? '⚠ 不同意公開' : ''],
    ['報名編號', rec['報名編號']], ['收到時間', rec['收到時間']]
  ]) + `<p style="margin:14px 0 0;"><a href="${spreadsheet_().getUrl()}" style="display:inline-block;background:#2F7DE0;color:#fff;padding:10px 18px;border-radius:10px;text-decoration:none;font-weight:700;">打開報名試算表</a></p>`);
  MailApp.sendEmail({ to: NOTIFY_EMAIL, subject, htmlBody: html, name: '雷切常態班報名系統' });
}

// 用瀏覽器開啟部署網址時顯示，方便確認有部署成功
function doGet() {
  return json_({ ok: true, message: '8 大科學原理常態班報名接收程式運作中' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
