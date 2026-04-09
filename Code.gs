// Google Apps Script - API専用
const SPREADSHEET_ID = '108gdBRGXehCmvfL26ZpLZhzykf1MzlgDHaJW_H5lDCk';

function doGet(e) {
  const action = e.parameter.action;
  let result;

  if (action === 'getMembers') {
    result = getMembers();
  } else if (action === 'getHistory') {
    result = getRecentPurchases();
  } else {
    result = { error: 'Unknown action' };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const action = params.action;
  let result;

  if (action === 'record') {
    result = recordPurchase(params.name, params.quantity);
  } else if (action === 'delete') {
    result = deletePurchase(params.id);
  } else {
    result = { error: 'Unknown action' };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getMembers() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('メンバー一覧');
  const data = sheet.getRange('A2:A').getValues();
  return data.filter(row => row[0] !== '').map(row => row[0]);
}

function recordPurchase(name, quantity) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('購入履歴');
  const id = Utilities.getUuid().substring(0, 8);
  const now = new Date();
  const amount = quantity * 100;
  sheet.appendRow([now, name, quantity, amount, id]);
  updateMonthlySummary();
  return { success: true, id: id, datetime: now.toLocaleString('ja-JP'), amount: amount };
}

function getRecentPurchases() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('購入履歴');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const startRow = Math.max(2, lastRow - 49);
  const numRows = lastRow - startRow + 1;
  const data = sheet.getRange(startRow, 1, numRows, 5).getValues();

  return data.filter(row => row[0] !== '').map(row => ({
    datetime: new Date(row[0]).toLocaleString('ja-JP'),
    name: row[1],
    quantity: row[2],
    amount: row[3],
    id: row[4]
  })).reverse();
}

function deletePurchase(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('購入履歴');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === id) {
      sheet.deleteRow(i + 1);
      updateMonthlySummary();
      return { success: true };
    }
  }
  return { success: false, message: '該当する記録が見つかりません' };
}

function updateMonthlySummary() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const historySheet = ss.getSheetByName('購入履歴');
  const summarySheet = ss.getSheetByName('月別集計');
  const memberSheet = ss.getSheetByName('メンバー一覧');

  const lastRow = historySheet.getLastRow();
  if (lastRow <= 1) {
    summarySheet.clear();
    summarySheet.getRange('A1').setValue('※データがありません');
    return;
  }

  const historyData = historySheet.getRange(2, 1, lastRow - 1, 4).getValues();

  const members = memberSheet.getRange('A2:A').getValues()
    .filter(r => r[0] !== '').map(r => r[0]);

  const summary = {};
  const allNames = new Set();
  historyData.forEach(row => {
    if (!row[0]) return;
    const date = new Date(row[0]);
    const monthKey = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM');
    const name = row[1];
    const quantity = row[2];

    if (!summary[monthKey]) summary[monthKey] = {};
    if (!summary[monthKey][name]) summary[monthKey][name] = 0;
    summary[monthKey][name] += quantity;
    allNames.add(name);
  });

  const sortedNames = [...members];
  allNames.forEach(name => {
    if (!sortedNames.includes(name)) sortedNames.push(name);
  });

  const months = Object.keys(summary).sort();

  summarySheet.clear();

  const headerRow = ['名前', ...months, '合計(本)', '合計(円)'];
  const numCols = headerRow.length;

  const dataRows = [];
  const monthTotals = months.map(() => 0);
  let grandTotal = 0;

  sortedNames.forEach(name => {
    const row = [name];
    let memberTotal = 0;
    months.forEach((month, i) => {
      const qty = (summary[month] && summary[month][name]) ? summary[month][name] : 0;
      row.push(qty || '');
      memberTotal += qty;
      monthTotals[i] += qty;
    });
    row.push(memberTotal);
    row.push(memberTotal * 100 + '円');
    grandTotal += memberTotal;
    dataRows.push(row);
  });

  const totalRow = ['合計', ...monthTotals, grandTotal, grandTotal * 100 + '円'];

  const allData = [headerRow, ...dataRows, totalRow];
  summarySheet.getRange(1, 1, allData.length, numCols).setValues(allData);

  summarySheet.getRange(1, 1, 1, numCols).setFontWeight('bold')
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  summarySheet.getRange(allData.length, 1, 1, numCols).setFontWeight('bold')
    .setBackground('#e8f0fe');

  summarySheet.getRange(2, 1, dataRows.length, 1).setFontWeight('bold');

  if (numCols > 1) {
    summarySheet.getRange(1, 2, allData.length, numCols - 1).setHorizontalAlignment('center');
  }

  summarySheet.setFrozenRows(1);
  summarySheet.setFrozenColumns(1);
}
