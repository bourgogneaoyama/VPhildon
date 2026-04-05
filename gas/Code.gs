// ===== Google Apps Script（水代記録アプリ） =====
// デプロイ：ウェブアプリとして公開（アクセス：全員）

const SPREADSHEET_ID = '1gSTkxTNJzHh1hj6xusaU0cDKqnS0gUiCdEgPV6OhhKY';

function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

// --- CORS対応 ---
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON);
}

function createJsonOutput(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- メインルーター ---
function doGet(e) {
  const action = e.parameter.action;
  switch (action) {
    case 'getMembers':
      return getMembers();
    case 'getSummary':
      return getSummary(e.parameter.month);
    default:
      return createJsonOutput({ error: '不明なアクション' });
  }
}

function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const action = params.action;
  switch (action) {
    case 'record':
      return addRecord(params.name, params.amount);
    case 'cancel':
      return cancelRecord(params.name);
    default:
      return createJsonOutput({ error: '不明なアクション' });
  }
}

// --- メンバー一覧取得 ---
function getMembers() {
  const sheet = getSheet('メンバー');
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return createJsonOutput({ members: [] });
  const values = sheet.getRange(1, 1, lastRow, 1).getValues();
  const members = values.map(r => r[0]).filter(v => v !== '');
  return createJsonOutput({ members: members });
}

// --- 記録追加 ---
function addRecord(name, amount) {
  if (!name || !amount) {
    return createJsonOutput({ error: '名前と金額は必須です' });
  }
  const sheet = getSheet('記録');
  const now = new Date();
  sheet.appendRow([now, name, Number(amount)]);
  return createJsonOutput({ success: true, timestamp: now.toISOString(), name: name, amount: Number(amount) });
}

// --- 直近1件取り消し（当日分のみ） ---
function cancelRecord(name) {
  if (!name) {
    return createJsonOutput({ error: '名前は必須です' });
  }
  const sheet = getSheet('記録');
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return createJsonOutput({ error: '記録がありません' });
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 最終行から遡って当日＆同名の直近1件を探す
  for (let i = lastRow; i >= 1; i--) {
    const row = sheet.getRange(i, 1, 1, 3).getValues()[0];
    const ts = new Date(row[0]);
    ts.setHours(0, 0, 0, 0);
    if (ts.getTime() === today.getTime() && row[1] === name) {
      sheet.deleteRow(i);
      return createJsonOutput({ success: true, cancelled: { timestamp: row[0], name: row[1], amount: row[2] } });
    }
  }
  return createJsonOutput({ error: '本日の記録が見つかりません' });
}

// --- 月次集計 ---
function getSummary(month) {
  // month: "2026-04" 形式。省略時は当月
  const now = new Date();
  if (!month) {
    month = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  const [year, mon] = month.split('-').map(Number);
  const startDate = new Date(year, mon - 1, 1);
  const endDate = new Date(year, mon, 1);

  const sheet = getSheet('記録');
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return createJsonOutput({ month: month, summary: [] });

  const data = sheet.getRange(1, 1, lastRow, 3).getValues();
  const totals = {};

  data.forEach(row => {
    const ts = new Date(row[0]);
    if (ts >= startDate && ts < endDate) {
      const name = row[1];
      const amount = Number(row[2]) || 0;
      totals[name] = (totals[name] || 0) + amount;
    }
  });

  const summary = Object.keys(totals).sort().map(name => ({
    name: name,
    total: totals[name]
  }));

  return createJsonOutput({ month: month, summary: summary });
}
