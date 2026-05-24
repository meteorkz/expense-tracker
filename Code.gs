// ============================================================
//  รายรับรายจ่าย — Google Apps Script API
//  วิธี Deploy:
//  1. เปิด Google Sheets → Extensions → Apps Script
//  2. วางโค้ดนี้ทั้งหมดใน Code.gs
//  3. Deploy → New deployment → Web app
//     Execute as: Me | Who has access: Anyone
//  4. Copy URL ไปวางใน Setup หน้าแรกของ App
// ============================================================

const SHEET_NAME = 'Transactions';
const SPREADSHEET_ID = '1Jg5zC0P9VVG74xQ8P5EgYnOxA3YZCNw3W2Z4J08Rnno';

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  let result;

  try {
    switch (p.action) {
      case 'get':    result = getTransactions();      break;
      case 'add':    result = addTransaction(p);      break;
      case 'delete': result = deleteTransaction(p.id); break;
      default:       result = { error: 'Unknown action: ' + p.action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getTransactions() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return { transactions: [] };

  const transactions = data.slice(1)
    .filter(r => r[0])
    .map(r => ({
      id:          String(r[0]),
      date:        (r[1] instanceof Date)
                     ? Utilities.formatDate(r[1], 'Asia/Bangkok', 'yyyy-MM-dd')
                     : String(r[1]),
      type:        String(r[2]),
      category:    String(r[3]),
      description: String(r[4]),
      amount:      Number(r[5])
    }));

  return { transactions };
}

function addTransaction(p) {
  const sheet = getOrCreateSheet();
  const id    = Date.now().toString();

  sheet.appendRow([
    id,
    p.date        || Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'),
    p.type        || '',
    p.category    || '',
    p.description || '',
    parseFloat(p.amount) || 0
  ]);

  return { success: true, id };
}

function deleteTransaction(id) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { error: 'Not found' };
}

function getOrCreateSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ID', 'Date', 'Type', 'Category', 'Description', 'Amount']);
    sheet.setFrozenRows(1);

    const hr = sheet.getRange(1, 1, 1, 6);
    hr.setBackground('#1E293B');
    hr.setFontColor('#FFFFFF');
    hr.setFontWeight('bold');

    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 110);
    sheet.setColumnWidth(3, 90);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(6, 100);
  }

  return sheet;
}
