/**
 * 多語種 AI+ESG+DEI 培訓報名表 → Google Sheet（含新報名寄信通知）
 * 貼到「多語種AI+ESG+DEI」試算表的 擴充功能 > Apps Script
 */
const SHEET_ID   = '16gycTqhNldhBP_JFJBuR9Sypa-NZJBalRHDp8HUfyo4';
const SHEET_NAME = '工作表1';
// 新報名通知信收件人（可填多個，用逗號分隔）
const NOTIFY_TO  = '858377355tw@gmail.com';
const HEADERS = ['時間','姓名','公司／單位','職稱','Email','電話','單位類型','預計參訓人數','想了解的項目','學習語言','問題或需求','個資同意','來源頁面','處理狀態','備註'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const p = (e && e.parameter) || {};

    if (!clean_(p.name) || !clean_(p.org) || !clean_(p.email) || p.agree !== '是') {
      return json_({ ok: false, error: 'missing' });
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p.email)) {
      return json_({ ok: false, error: 'email' });
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#EEF2EC');
      sh.setFrozenRows(1);
    }

    sh.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),
      clean_(p.name), clean_(p.org), clean_(p.title), clean_(p.email), clean_(p.phone),
      clean_(p.role), clean_(p.size), clean_(p.interest), clean_(p.lang),
      clean_(p.message), '是', clean_(p.source), '待處理', ''
    ]);
    try { notify_(p); } catch (mailErr) { /* 寄信失敗不影響資料寫入 */ }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function notify_(p) {
  const oneLine = function (v) { return clean_(v).replace(/[\r\n]+/g, ' '); };
  const body = [
    '收到一筆新的培訓報名／試用洽詢：', '',
    '姓名：' + oneLine(p.name),
    '公司／單位：' + oneLine(p.org),
    '職稱：' + oneLine(p.title),
    'Email：' + oneLine(p.email),
    '電話：' + oneLine(p.phone),
    '單位類型：' + oneLine(p.role),
    '預計參訓人數：' + oneLine(p.size),
    '想了解的項目：' + oneLine(p.interest),
    '學習語言：' + oneLine(p.lang), '',
    '問題或需求：', clean_(p.message), '',
    '試算表：https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit'
  ].join('\n');
  MailApp.sendEmail({
    to: NOTIFY_TO,
    subject: '【多語種AI+ESG 新報名】' + oneLine(p.name) + '｜' + oneLine(p.org),
    body: body,
    replyTo: oneLine(p.email),
    name: '培訓報名表'
  });
}

// 開網址測試用：看到 {"ok":true,...} 代表部署成功
function doGet() { return json_({ ok: true, service: 'AI+ESG+DEI signup' }); }

// 防公式注入 + 限長
function clean_(v) {
  let s = String(v == null ? '' : v).trim().slice(0, 2000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 在編輯器手動執行一次：授權、建立表頭、寄一封測試通知信
function testWrite() {
  doPost({ parameter: { name: '測試', org: '測試公司', email: 'test@example.com', role: '其他', size: '尚未確定', interest: '試用／示範', lang: '華語', agree: '是', message: '這是一筆測試，可刪除' } });
}
