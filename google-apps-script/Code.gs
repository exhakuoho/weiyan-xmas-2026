/**
 * 2026 聖誕科技闖關日 報名 → Google 試算表 接收程式
 * 使用方式見同資料夾「設定說明.md」
 *
 * 網站會分兩次送資料進來：
 *   1. 報名表送出（action 空白）→ 每位孩子新增一列（同一個報名編號），寄「報名資料已收到＋轉帳資訊」給家長
 *   2. 家長轉帳後回填末 5 碼（action = 'payment'）→ 寄「已收到轉帳資訊」給家長、寄通知給管理信箱
 *   工作人員核對款項後，把「處理狀態」改成「已付款」，再從試算表選單
 *   「🎄 聖誕報名 → 寄送付款確認信」寄出正式確認信
 *
 * 表單不因額滿自動停止收件、不設候補（依 10/21 內部規劃）；每場人數請看「場次統計」分頁。
 */

const SHEET_NAME = '報名資料';
const STATS_SHEET = '場次統計';
// 試算表 ID（網址 /d/ 與 /edit 中間那串）。填了就固定寫進這份試算表；
// 留空字串則寫進「這個 Apps Script 所附屬的試算表」（從試算表的 擴充功能 → Apps Script 建立時）
const SHEET_ID = '';
const PRICE = 800;          // 每位孩子活動費
const MAX_KIDS = 4;         // 一次最多替幾位孩子報名
const CAPACITY = 60;        // 每場規劃上限（只用在統計，不會擋報名）
const PAY_DAYS_BEFORE = 7;  // 活動日前幾天要完成轉帳
// 有人完成報名時，寄通知給這個信箱；信件都由部署這支程式的帳號寄出
const NOTIFY_EMAIL = 'randdlaboratary@gmail.com';
const NOTIFY_ADMIN = true;  // 每天寄信量大時可改 false，只寄給家長
const SENDER_NAME = '微研 Weiyan 科學教育';
const EVENT_NAME = '2026 聖誕科技闖關日';
const LINE_URL = 'https://lin.ee/rZW93KN';
const BANK_TEXT = '中國信託（822）營業部　帳號 901890457506';

// 場次：id 要和網站 index.html 的 SESSIONS 一致
const SESSIONS = {
  '1219AM': { date: '2026-12-19', slot: '上午', venue: '飛米樂高' },
  '1219PM': { date: '2026-12-19', slot: '下午', venue: '飛米樂高' },
  '1220AM': { date: '2026-12-20', slot: '上午', venue: '飛米樂高' },
  '1220PM': { date: '2026-12-20', slot: '下午', venue: '飛米樂高' },
  '1225AM': { date: '2026-12-25', slot: '上午', venue: '高科大建工校區行政大樓' },
  '1225PM': { date: '2026-12-25', slot: '下午', venue: '高科大建工校區行政大樓' },
  '1226AM': { date: '2026-12-26', slot: '上午', venue: '高科大建工校區行政大樓' },
  '1226PM': { date: '2026-12-26', slot: '下午', venue: '高科大建工校區行政大樓' },
  '1227AM': { date: '2026-12-27', slot: '上午', venue: '高科大建工校區行政大樓' },
  '1227PM': { date: '2026-12-27', slot: '下午', venue: '高科大建工校區行政大樓' }
};
const SLOT_TIME = { '上午': '09:00–12:00（最晚 10:30 進場）', '下午': '14:00–17:00（最晚 15:30 進場）' };
const WEEK = ['日', '一', '二', '三', '四', '五', '六'];

const HEADERS = [
  '收到時間', '報名編號', '孩子序', '場次', '活動日期', '時段', '活動地點',
  '孩子姓名', '出生年月日', '身分證／居留證號', '性別', '國籍', '年級',
  '飲食過敏', '特殊協助需求',
  '監護人姓名', '監護人電話', '家長Email', '緊急聯絡人', '緊急聯絡人關係', '緊急聯絡電話',
  '付款方式', '本筆應繳金額', '付款期限', '轉帳後5碼',
  '攝影同意', '過敏原已確認', '退費規則已讀', '得知管道', '備註', '瀏覽風格', '個資同意',
  '處理狀態', '確認信', '管理備註'
];
const COL = name => HEADERS.indexOf(name) + 1; // 欄號（從 1 開始）
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const PHONE_RE = /^0\d{8,9}$/;
const STATUS_WAIT_PAY = '待回報轉帳';
const STATUS_PAID_REPORTED = '已回報轉帳（待對帳）';
const STATUS_PAID = '已付款';

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
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#B3261E').setFontColor('#FFF6E5');
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

// 信件與通知裡只顯示部分證號：A12****789
function maskId_(v) {
  const s = String(v || '');
  return s.length > 6 ? s.slice(0, 3) + '****' + s.slice(-3) : s;
}

// 身分證、新式居留證（第 2 碼 8/9）、舊式居留證（第 2 碼 A–D）檢查碼
function twIdOK_(id) {
  const L = 'ABCDEFGHJKLMNPQRSTUVXYWZIO';
  if (!/^[A-Z][A-D0-9]\d{8}$/.test(id)) return false;
  const n = L.indexOf(id[0]) + 10;
  const d1 = /[A-D]/.test(id[1]) ? (L.indexOf(id[1]) + 10) % 10 : +id[1];
  const ds = [d1].concat(id.slice(2).split('').map(Number));
  const w = [8, 7, 6, 5, 4, 3, 2, 1, 1];
  let sum = Math.floor(n / 10) + (n % 10) * 9;
  ds.forEach((d, i) => sum += d * w[i]);
  return sum % 10 === 0;
}

function sessionInfo_(id) {
  const s = SESSIONS[id];
  if (!s) return null;
  const d = new Date(s.date + 'T00:00:00+08:00');
  const md = (x) => `${Number(Utilities.formatDate(x, 'Asia/Taipei', 'M'))}/${Number(Utilities.formatDate(x, 'Asia/Taipei', 'd'))}（${WEEK[Number(Utilities.formatDate(x, 'Asia/Taipei', 'u')) % 7]}）`;
  const due = new Date(d.getTime() - PAY_DAYS_BEFORE * 864e5);
  const today = new Date(Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd') + 'T00:00:00+08:00');
  return {
    id, date: s.date, slot: s.slot, venue: s.venue, start: d,
    label: `${md(d)}${s.slot} ${SLOT_TIME[s.slot]}`,
    dueText: due < today ? '請於報名當天完成轉帳' : `${md(due)} 23:59 前`
  };
}

// 依報名編號找出所有列（多位孩子同一個編號）
function findRows_(sh, refCode) {
  const n = sh.getLastRow() - 1;
  if (n <= 0 || !refCode) return [];
  const vals = sh.getRange(2, 1, n, HEADERS.length).getValues();
  const out = [];
  vals.forEach((r, i) => {
    if (String(r[COL('報名編號') - 1]) !== String(refCode)) return;
    const rec = {};
    HEADERS.forEach((h, k) => rec[h] = String(r[k]));
    out.push({ row: i + 2, rec });
  });
  return out;
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
    return json_(handleRegister_(data));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// 第 1 步：報名表送出
function handleRegister_(data) {
  const s = sessionInfo_(data.sessionId);
  if (!s) throw new Error('請選擇場次');
  if (s.start.getTime() + 864e5 <= Date.now()) throw new Error('這個場次已經結束，請改選其他場次');
  const kids = Array.isArray(data.kids) ? data.kids.slice(0, MAX_KIDS) : [];
  if (!kids.length) throw new Error('請填寫孩子資料');
  kids.forEach((k, i) => {
    if (!k.name || !k.birth || !k.idNo || !k.gender || !k.grade) throw new Error(`第 ${i + 1} 位孩子資料不完整`);
    k.idNo = String(k.idNo).toUpperCase().replace(/\s/g, '');
    if (k.nationality === '中華民國' && !twIdOK_(k.idNo)) throw new Error(`第 ${i + 1} 位孩子的身分證字號格式不正確`);
    if (!/^[A-Z0-9]{6,20}$/.test(k.idNo)) throw new Error(`第 ${i + 1} 位孩子的證件號碼格式不正確`);
  });
  if (!data.guardianName || !PHONE_RE.test(String(data.guardianPhone || ''))) throw new Error('請填寫監護人姓名與電話');
  if (!data.emergencyName || !PHONE_RE.test(String(data.emergencyPhone || ''))) throw new Error('請填寫緊急聯絡人與電話');
  if (!EMAIL_RE.test(String(data.email || ''))) throw new Error('Email 格式不正確');
  if (data.consent !== '同意' || data.refundAck !== '已讀' || data.allergenAck !== '已確認') throw new Error('請勾選同意事項');

  const sh = getSheet_();
  const total = PRICE * kids.length;
  kids.forEach((k, i) => {
    const v = {
      '收到時間': now_(), '報名編號': data.refCode, '孩子序': `${i + 1}/${kids.length}`,
      '場次': s.id, '活動日期': s.date, '時段': s.slot, '活動地點': s.venue,
      '孩子姓名': k.name, '出生年月日': k.birth, '身分證／居留證號': k.idNo, '性別': k.gender, '國籍': k.nationality, '年級': k.grade,
      '飲食過敏': k.allergy || '無', '特殊協助需求': k.needs || '',
      '監護人姓名': data.guardianName, '監護人電話': data.guardianPhone, '家長Email': data.email,
      '緊急聯絡人': data.emergencyName, '緊急聯絡人關係': data.emergencyRelation, '緊急聯絡電話': data.emergencyPhone,
      '付款方式': '銀行轉帳', '本筆應繳金額': i === 0 ? String(total) : '（併第 1 列）', '付款期限': s.dueText,
      '攝影同意': data.photoConsent, '過敏原已確認': data.allergenAck, '退費規則已讀': data.refundAck,
      '得知管道': data.source, '備註': data.note, '瀏覽風格': data.pageStyle, '個資同意': data.consent,
      '處理狀態': STATUS_WAIT_PAY
    };
    sh.appendRow(HEADERS.map(h => h in v ? safe_(v[h]) : ''));
  });
  const first = sh.getLastRow() - kids.length + 1;
  try {
    mailParent_('received', data, kids, s, total);
    setCell_(sh, first, '確認信', '已寄收件通知 ' + now_());
  } catch (err) {
    setCell_(sh, first, '確認信', '收件通知寄送失敗：' + err.message);
  }
  return { ok: true, refCode: data.refCode, amount: total, kids: kids.length, session: s.label, venue: s.venue, due: s.dueText };
}

// 第 2 步：回填轉帳末 5 碼
function handlePayment_(data) {
  const last5 = String(data.last5 || '').trim();
  if (!/^\d{5}$/.test(last5)) throw new Error('末 5 碼需為 5 位數字');
  const sh = getSheet_();
  const rows = findRows_(sh, data.refCode);
  if (!rows.length) throw new Error('找不到這筆報名');
  const firstTime = !rows[0].rec['轉帳後5碼'];
  rows.forEach(f => {
    setCell_(sh, f.row, '轉帳後5碼', last5);
    if (f.rec['處理狀態'] === STATUS_WAIT_PAY) setCell_(sh, f.row, '處理狀態', STATUS_PAID_REPORTED);
  });
  const rec = rows[0].rec, s = sessionInfo_(rec['場次']);
  const kids = rows.map(f => ({ name: f.rec['孩子姓名'], idNo: f.rec['身分證／居留證號'], grade: f.rec['年級'], allergy: f.rec['飲食過敏'] }));
  const d = { refCode: data.refCode, guardianName: rec['監護人姓名'], guardianPhone: rec['監護人電話'], email: rec['家長Email'], last5 };
  const total = PRICE * rows.length;
  if (firstTime) mailParent_('reported', d, kids, s, total);
  if (NOTIFY_ADMIN && NOTIFY_EMAIL) mailAdmin_(rows, s, last5, firstTime);
  return { ok: true, refCode: data.refCode, email: rec['家長Email'] };
}

/* ================== 試算表選單 ================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎄 聖誕報名')
    .addItem('寄送付款確認信（處理狀態＝已付款）', 'sendPaidConfirmations')
    .addItem('更新場次統計', 'refreshStats')
    .addToUi();
}

// 處理狀態改成「已付款」、但還沒寄過付款確認信的報名，各寄一封正式確認信
function sendPaidConfirmations() {
  const sh = getSheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return;
  const vals = sh.getRange(2, 1, n, HEADERS.length).getValues();
  const groups = {};
  vals.forEach((r, i) => {
    const ref = String(r[COL('報名編號') - 1]);
    (groups[ref] = groups[ref] || []).push({ row: i + 2, r });
  });
  let sent = 0;
  Object.keys(groups).forEach(ref => {
    const g = groups[ref], r0 = g[0].r;
    if (String(r0[COL('處理狀態') - 1]) !== STATUS_PAID) return;
    if (String(r0[COL('確認信') - 1]).indexOf('付款確認') >= 0) return;
    const s = sessionInfo_(String(r0[COL('場次') - 1]));
    const kids = g.map(x => ({ name: String(x.r[COL('孩子姓名') - 1]), idNo: String(x.r[COL('身分證／居留證號') - 1]) }));
    const d = { refCode: ref, guardianName: String(r0[COL('監護人姓名') - 1]), email: String(r0[COL('家長Email') - 1]), last5: String(r0[COL('轉帳後5碼') - 1]) };
    mailParent_('confirmed', d, kids, s, PRICE * g.length);
    setCell_(sh, g[0].row, '確認信', '已寄付款確認 ' + now_());
    sent++;
  });
  refreshStats();
  SpreadsheetApp.getUi().alert(`已寄出 ${sent} 封付款確認信。`);
}

// 每場人數（處理狀態含「取消」「測試」的列不算）
function refreshStats() {
  const ss = spreadsheet_(), sh = getSheet_();
  const n = sh.getLastRow() - 1;
  const cnt = {};
  Object.keys(SESSIONS).forEach(k => cnt[k] = { all: 0, paid: 0 });
  if (n > 0) sh.getRange(2, 1, n, HEADERS.length).getValues().forEach(r => {
    const st = String(r[COL('處理狀態') - 1]), id = String(r[COL('場次') - 1]);
    if (!cnt[id] || /取消|測試/.test(st)) return;
    cnt[id].all++;
    if (st === STATUS_PAID) cnt[id].paid++;
  });
  let out = ss.getSheetByName(STATS_SHEET);
  if (!out) out = ss.insertSheet(STATS_SHEET);
  out.clear();
  const rows = [['場次', '日期時段', '地點', '報名人數', '已付款', '規劃上限', '剩餘', '更新時間']];
  Object.keys(SESSIONS).forEach(k => {
    const s = sessionInfo_(k);
    rows.push([k, s.label, s.venue, cnt[k].all, cnt[k].paid, CAPACITY, CAPACITY - cnt[k].all, now_()]);
  });
  out.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  out.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#1F7A4A').setFontColor('#FFFFFF');
  out.setFrozenRows(1);
}

/* ================== 信件 ================== */

function card_(title, bodyHtml) {
  return `<div style="background:#FFF6EA;padding:24px 12px;font-family:'Noto Sans TC','PingFang TC','Microsoft JhengHei',sans-serif;color:#3B2416;">
  <div style="max-width:560px;margin:0 auto;background:#fff;border-radius:18px;overflow:hidden;border:1px solid #F1D9BE;">
    <div style="background:#B3261E;color:#FFE7A8;padding:18px 24px;font-size:13px;letter-spacing:.2em;font-weight:700;">🎄 WEIYAN CHRISTMAS LAB</div>
    <div style="padding:24px;">
      <h2 style="margin:0 0 14px;font-size:22px;color:#B3261E;">${title}</h2>
      ${bodyHtml}
      <p style="margin:22px 0 0;font-size:14px;color:#8A6B55;">有任何問題，歡迎直接回覆這封信，或加入 <a href="${LINE_URL}" style="color:#1F7A4A;">官方 LINE</a>。</p>
      <p style="margin:6px 0 0;font-size:14px;color:#8A6B55;">${SENDER_NAME}</p>
    </div>
  </div>
</div>`;
}

function table_(pairs) {
  return `<table style="width:100%;border-collapse:collapse;font-size:15px;margin:8px 0;">${pairs
    .filter(p => p[1] !== '' && p[1] != null)
    .map(p => `<tr><td style="padding:8px 10px;background:#FFF6EA;border:1px solid #F1E2CF;width:34%;color:#8A6B55;white-space:nowrap;">${esc_(p[0])}</td><td style="padding:8px 10px;border:1px solid #F1E2CF;font-weight:700;">${esc_(p[1])}</td></tr>`)
    .join('')}</table>`;
}

function mailParent_(kind, d, kids, s, total) {
  const to = d.email;
  if (!EMAIL_RE.test(to)) return;
  const names = kids.map(k => k.name).join('、');
  const hi = `<p style="margin:0 0 12px;font-size:16px;">${esc_(d.guardianName)} 您好：</p>`;
  const base = [['報名編號', d.refCode], ['參加孩子', `${names}（共 ${kids.length} 位）`], ['場次', s ? s.label : ''], ['地點', s ? s.venue : '']];
  const amt = 'NT$' + Number(total).toLocaleString('en-US');
  let subject, html;
  if (kind === 'received') {
    subject = `【報名資料已收到】${EVENT_NAME}｜${names}`;
    html = card_('報名資料已收到，還差轉帳一步！🎁', hi +
      `<p style="margin:0 0 12px;font-size:16px;line-height:1.7;">我們已收到 <b>${esc_(names)}</b> 的報名資料。請於 <b style="color:#B3261E;">${esc_(s.dueText)}</b> 完成轉帳，期限內保留名額；逾期未付款會釋出名額。</p>` +
      table_(base.concat([['應繳金額', amt], ['轉帳帳戶', BANK_TEXT], ['付款期限', s.dueText]])) +
      `<p style="margin:12px 0 0;font-size:15px;line-height:1.7;">轉帳後若報名頁面已關閉，<b>直接回覆這封信告訴我們帳號末 5 碼</b>即可。</p>`);
  } else if (kind === 'reported') {
    subject = `【已收到轉帳資訊】${EVENT_NAME}｜${names}`;
    html = card_('已收到您的轉帳資訊 ✨', hi +
      `<p style="margin:0 0 12px;font-size:16px;line-height:1.7;">工作人員會在活動前核對款項，核對完成後再寄一封<b>付款確認信</b>給您。</p>` +
      table_(base.concat([['應繳金額', amt], ['轉帳末 5 碼', d.last5]])) +
      `<p style="margin:16px 0 0;font-size:17px;font-weight:700;color:#1F7A4A;">期待在聖誕闖關日見到孩子～🎄</p>`);
  } else {
    subject = `【報名完成・付款確認】${EVENT_NAME}｜${names}`;
    html = card_('付款已確認，報名完成！🎉', hi +
      `<p style="margin:0 0 12px;font-size:16px;line-height:1.7;">款項已核對完成。活動當天請直接到<b>服務台報到</b>，領取紙本集點卡就能開始闖關；每位孩子約需 90 分鐘，尖峰時段可能延長。</p>` +
      table_(base.concat([['時間', s ? SLOT_TIME[s.slot] : ''], ['已付金額', amt]])) +
      `<p style="margin:12px 0 0;font-size:15px;line-height:1.7;">取消規則：距活動日 4 日以上取消退還 80%（每位 NT$640）；活動日前 3 日內取消或未到場不退款。名額可轉讓，最晚請於活動前 1 日提供新參加者的保險資料。</p>` +
      `<p style="margin:16px 0 0;font-size:17px;font-weight:700;color:#1F7A4A;">聖誕快樂，現場見！🎅</p>`);
  }
  MailApp.sendEmail({ to, subject, htmlBody: html, name: SENDER_NAME, replyTo: NOTIFY_EMAIL });
}

function mailAdmin_(rows, s, last5, firstTime) {
  const rec = rows[0].rec;
  const subject = `${firstTime ? '【新報名】' : '【更新末 5 碼】'}${s ? s.id : ''}｜${rows.map(f => f.rec['孩子姓名']).join('、')}｜末 5 碼 ${last5}`;
  const kids = rows.map(f => [`孩子 ${f.rec['孩子序']}`, `${f.rec['孩子姓名']}（${f.rec['年級']}・${maskId_(f.rec['身分證／居留證號'])}）${f.rec['飲食過敏'] && f.rec['飲食過敏'] !== '無' ? '　⚠ 過敏：' + f.rec['飲食過敏'] : ''}`]);
  const html = card_(esc_(subject.replace(/^【[^】]+】/, '')), table_([
    ['場次', s ? s.label : rec['場次']], ['地點', rec['活動地點']]
  ].concat(kids).concat([
    ['監護人', `${rec['監護人姓名']}　${rec['監護人電話']}`], ['Email', rec['家長Email']],
    ['緊急聯絡', `${rec['緊急聯絡人']}（${rec['緊急聯絡人關係']}）${rec['緊急聯絡電話']}`],
    ['應繳金額', 'NT$' + (PRICE * rows.length).toLocaleString('en-US')], ['轉帳末 5 碼', last5],
    ['攝影同意', rec['攝影同意']], ['得知管道', rec['得知管道']], ['備註', rec['備註']],
    ['報名編號', rec['報名編號']], ['收到時間', rec['收到時間']]
  ])) + `<p style="margin:14px 0 0;"><a href="${spreadsheet_().getUrl()}" style="display:inline-block;background:#1F7A4A;color:#fff;padding:10px 18px;border-radius:10px;text-decoration:none;font-weight:700;">打開報名試算表</a></p>`);
  MailApp.sendEmail({ to: NOTIFY_EMAIL, subject, htmlBody: html, name: '聖誕闖關日報名系統' });
}

// 用瀏覽器開啟部署網址時顯示，方便確認有部署成功
function doGet() {
  return json_({ ok: true, message: '2026 聖誕科技闖關日報名接收程式運作中' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
