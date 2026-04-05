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
      return addRecord(params.name, params.amount, params.memberId);
    case 'cancel':
      return cancelRecord(params.name);
    default:
      return createJsonOutput({ error: '不明なアクション' });
  }
}

// --- メンバー一覧取得 ---
// メンバーシート: A列=社員番号, B列=名前（1行目ヘッダー）
function getMembers() {
  const sheet = getSheet('メンバー');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return createJsonOutput({ members: [] });
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const members = values
    .filter(r => r[1] !== '')
    .map(r => ({ id: String(r[0]), name: String(r[1]) }));
  return createJsonOutput({ members: members });
}

// --- 記録追加 ---
// 記録シート: A列=タイムスタンプ, B列=社員番号, C列=名前, D列=金額
function addRecord(name, amount, memberId) {
  if (!name || !amount) {
    return createJsonOutput({ error: '名前と金額は必須です' });
  }
  const sheet = getSheet('記録');
  const now = new Date();
  sheet.appendRow([now, memberId || '', name, Number(amount)]);
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
  // 記録シート: A=タイムスタンプ, B=社員番号, C=名前, D=金額
  for (let i = lastRow; i >= 1; i--) {
    const row = sheet.getRange(i, 1, 1, 4).getValues()[0];
    const ts = new Date(row[0]);
    ts.setHours(0, 0, 0, 0);
    if (ts.getTime() === today.getTime() && row[2] === name) {
      sheet.deleteRow(i);
      return createJsonOutput({ success: true, cancelled: { timestamp: row[0], name: row[2], amount: row[3] } });
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

  // 記録シート: A=タイムスタンプ, B=社員番号, C=名前, D=金額
  const data = sheet.getRange(1, 1, lastRow, 4).getValues();
  const totals = {};

  data.forEach(row => {
    const ts = new Date(row[0]);
    if (ts >= startDate && ts < endDate) {
      const name = row[2];
      const amount = Number(row[3]) || 0;
      totals[name] = (totals[name] || 0) + amount;
    }
  });

  const summary = Object.keys(totals).sort().map(name => ({
    name: name,
    total: totals[name]
  }));

  return createJsonOutput({ month: month, summary: summary });
}

// --- 年度別集計シート自動作成 ---
// 毎年4月1日にトリガーで実行、または手動実行可
// createYearlySummarySheet() で当年度分を作成
// createYearlySummarySheet(2027) で指定年度を作成
function createYearlySummarySheet(fiscalYear) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 年度の判定（4月始まり）: 省略時は現在の年度
  if (!fiscalYear) {
    const now = new Date();
    fiscalYear = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
  }

  const sheetName = fiscalYear + '年度集計';

  // 既に存在する場合はスキップ
  if (ss.getSheetByName(sheetName)) {
    Logger.log(sheetName + ' は既に存在します');
    return;
  }

  // メンバー取得
  const memSheet = ss.getSheetByName('メンバー');
  const lastRow = memSheet.getLastRow();
  if (lastRow < 2) return;
  const members = memSheet.getRange(2, 1, lastRow - 1, 2).getValues().filter(function(r) { return r[1] !== ''; });

  // 月リスト（4月〜翌3月）
  var months = [];
  for (var m = 4; m <= 12; m++) months.push({ year: fiscalYear, month: m });
  for (var m = 1; m <= 3; m++) months.push({ year: fiscalYear + 1, month: m });

  // シート作成
  var sheet = ss.insertSheet(sheetName);

  // ヘッダー
  var header = ['社員番号', '名前'];
  months.forEach(function(d) { header.push(d.year + '年' + d.month + '月'); });
  header.push('年間合計');
  sheet.getRange(1, 1, 1, header.length).setValues([header]);

  // メンバー行（数式）
  var dataRows = [];
  members.forEach(function(mem, idx) {
    var row = [mem[0], mem[1]];
    var r = idx + 2;
    months.forEach(function(d) {
      var startDate = 'DATE(' + d.year + ',' + d.month + ',1)';
      var lastDay = 'EOMONTH(DATE(' + d.year + ',' + d.month + ',1),0)+0.99999';
      row.push('=SUMPRODUCT((記録!C2:C10000=B' + r + ')*(記録!A2:A10000>=' + startDate + ')*(記録!A2:A10000<=' + lastDay + ')*記録!D2:D10000)');
    });
    row.push('=SUM(C' + r + ':N' + r + ')');
    dataRows.push(row);
  });

  // 合計行
  var lastDataRow = members.length + 1;
  var totalRow = ['', '合計'];
  for (var i = 0; i < months.length + 1; i++) {
    var col = String.fromCharCode(67 + i);
    totalRow.push('=SUM(' + col + '2:' + col + lastDataRow + ')');
  }
  dataRows.push(totalRow);

  sheet.getRange(2, 1, dataRows.length, header.length).setValues(dataRows);

  // 書式調整
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, header.length);

  Logger.log(sheetName + ' を作成しました');
}

// --- 毎年4月に自動実行するためのトリガー設定 ---
// 初回に1度だけ手動実行してください
function setupYearlyTrigger() {
  // 既存のトリガーを削除（重複防止）
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'createYearlySummarySheet') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 毎年4月1日 9:00 に実行
  ScriptApp.newTrigger('createYearlySummarySheet')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  Logger.log('月次トリガーを設定しました（毎月1日に実行、4月のみシート作成）');
}
