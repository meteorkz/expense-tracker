// ============================================================
//  รายรับรายจ่าย — Google Apps Script API
//  วิธี Deploy:
//  1. เปิด Google Sheets → Extensions → Apps Script
//  2. วางโค้ดนี้ทั้งหมดใน Code.gs
//  3. Deploy → New deployment → Web app
//     Execute as: Me | Who has access: Anyone
//  4. Copy URL ไปวางใน Setup หน้าแรกของ App
// ============================================================

const SHEET_NAME   = 'Transactions';
const BUDGET_SHEET = 'Budgets';
const SPREADSHEET_ID = '1Jg5zC0P9VVG74xQ8P5EgYnOxA3YZCNw3W2Z4J08Rnno';

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  let result;

  try {
    switch (p.action) {
      case 'get':          result = getTransactions();        break;
      case 'add':          result = addTransaction(p);        break;
      case 'delete':       result = deleteTransaction(p.id);  break;
      case 'getBudgets':      result = getBudgets(p.period);      break;
      case 'setBudget':       result = setBudget(p);              break;
      case 'deleteBudget':    result = deleteBudget(p.id);        break;
      case 'clearBudgets':    result = clearBudgets();            break;
      case 'clearTransactions': result = clearTransactions();     break;
      case 'getTasks':    result = getTasks();           break;
      case 'addTask':     result = addTask(p);           break;
      case 'deleteTask':  result = deleteTaskById(p.id); break;
      default:             result = { error: 'Unknown action: ' + p.action };
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

// ── Budget Functions ──────────────────────────────────────────────

function periodStr(val) {
  if (val instanceof Date) {
    const s = Utilities.formatDate(val, 'Asia/Bangkok', 'yyyy-MM');
    return s;
  }
  return String(val);
}

function getBudgets(period) {
  const sheet = getOrCreateBudgetSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { budgets: [] };

  const budgets = data.slice(1)
    .filter(r => r[0] && (!period || periodStr(r[1]) === period))
    .map(r => ({
      id:       String(r[0]),
      period:   periodStr(r[1]),
      type:     String(r[2]),
      category: String(r[3]),
      amount:   Number(r[4])
    }));

  return { budgets };
}

function setBudget(p) {
  const sheet = getOrCreateBudgetSheet();
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (periodStr(data[i][1]) === String(p.period) && String(data[i][3]) === String(p.category)) {
      sheet.getRange(i + 1, 5).setValue(parseFloat(p.amount));
      return { success: true, id: String(data[i][0]) };
    }
  }

  const id  = Date.now().toString();
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 2).setNumberFormat('@');
  sheet.getRange(row, 1, 1, 5).setValues([[id, p.period, p.type || '', p.category, parseFloat(p.amount)]]);
  return { success: true, id };
}

function clearBudgets() {
  const sheet = getOrCreateBudgetSheet();
  const last  = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  sheet.getRange(2, 2, 1000, 1).setNumberFormat('@');
  return { success: true, cleared: last - 1 };
}

function clearTransactions() {
  const sheet = getOrCreateSheet();
  const last  = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  return { success: true, cleared: last - 1 };
}

function deleteBudget(id) {
  const sheet = getOrCreateBudgetSheet();
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

function getOrCreateBudgetSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(BUDGET_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(BUDGET_SHEET);
    sheet.appendRow(['ID', 'Period', 'Type', 'Category', 'Amount']);
    sheet.setFrozenRows(1);
    const hr = sheet.getRange(1, 1, 1, 5);
    hr.setBackground('#4F46E5'); hr.setFontColor('#FFFFFF'); hr.setFontWeight('bold');
    sheet.setColumnWidth(1, 160); sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 90);  sheet.setColumnWidth(4, 120); sheet.setColumnWidth(5, 100);
  }
  return sheet;
}

// ── Tasks Sheet ───────────────────────────────────────────────────

const TASKS_SHEET = 'Tasks';

function getTasks() {
  const sheet = getOrCreateTaskSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { tasks: [] };
  const tasks = data.slice(1).filter(r => r[0]).map(r => ({
    id:          String(r[0]),
    month:       String(r[1]),
    category:    String(r[2]),
    description: String(r[3]),
    amount:      Number(r[4])
  }));
  return { tasks };
}

function addTask(p) {
  const sheet = getOrCreateTaskSheet();
  const id    = Date.now().toString();
  const row   = sheet.getLastRow() + 1;
  sheet.getRange(row, 2).setNumberFormat('@');
  sheet.getRange(row, 1, 1, 5).setValues([[
    id, p.month, p.category, p.description || '', parseFloat(p.amount) || 0
  ]]);
  return { success: true, id };
}

function deleteTaskById(id) {
  const sheet = getOrCreateTaskSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

function getOrCreateTaskSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(TASKS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TASKS_SHEET);
    sheet.appendRow(['ID', 'Month', 'Category', 'Description', 'Amount']);
    sheet.setFrozenRows(1);
    const hr = sheet.getRange(1, 1, 1, 5);
    hr.setBackground('#0F766E'); hr.setFontColor('#FFFFFF'); hr.setFontWeight('bold');
    sheet.getRange(2, 2, 1000, 1).setNumberFormat('@');
    sheet.setColumnWidth(1, 160); sheet.setColumnWidth(2, 90);
    sheet.setColumnWidth(3, 100); sheet.setColumnWidth(4, 200); sheet.setColumnWidth(5, 100);
  }
  return sheet;
}

// ── Transaction Sheet ─────────────────────────────────────────────

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
