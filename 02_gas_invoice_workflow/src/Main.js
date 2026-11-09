/**
 * スプレッドシートの読み書きとメニュー。計算は Logic.js に分けている。
 * メールの送信は行わず、督促文は下書きとしてシートに出すだけにしている。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('請求自動化')
    .addItem('請求と入金状況を作成', 'runInvoiceWorkflow')
    .addSeparator()
    .addItem('サンプルデータを入れる', 'setupSampleData')
    .addToUi();
}

function toDateText_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value === null || value === undefined ? '' : value).trim().replace(/\//g, '-');
}

function toMonthText_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return String(value).trim().replace(/\//g, '-');
}

function readRecords_(spreadsheet, name) {
  var sheet = spreadsheet.getSheetByName(name);
  if (!sheet) throw new Error('シート「' + name + '」がありません。');
  return rowsToObjects(sheet.getDataRange().getValues());
}

function writeTable_(spreadsheet, name, header, rows) {
  var sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  sheet.clearContents();
  var table = [header].concat(rows);
  sheet.getRange(1, 1, table.length, header.length).setValues(table);
}

function readSettings_(spreadsheet) {
  var settings = {};
  readRecords_(spreadsheet, '設定').forEach(function (row) { settings[row['項目']] = row['値']; });
  var month = toMonthText_(settings['対象月']);
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('設定の対象月は 2026-10 の形式で入力してください。');
  var today = settings['基準日'] ? toDateText_(settings['基準日']) : toDateText_(new Date());
  if (!isValidDateString(today)) throw new Error('設定の基準日が日付ではありません。');
  return { month: month, today: today };
}

function readCustomers_(spreadsheet) {
  var customers = {};
  readRecords_(spreadsheet, '顧客マスタ').forEach(function (row) {
    customers[String(row['顧客コード'])] = { name: row['顧客名'], paymentDays: Number(row['支払日数']) };
  });
  return customers;
}

function runInvoiceWorkflow() {
  var spreadsheet = SpreadsheetApp.getActive();
  var settings = readSettings_(spreadsheet);
  var customers = readCustomers_(spreadsheet);

  var orders = readRecords_(spreadsheet, '受注一覧').map(function (order) {
    order['受注日'] = toDateText_(order['受注日']);
    return order;
  });
  var checked = validateOrders(orders, customers, settings.month);
  var invoices = buildInvoices(checked.valid, customers, settings.month, checked.heldCustomers);
  var payments = reconcilePayments(invoices, readRecords_(spreadsheet, '入金一覧'), settings.today);
  var issues = checked.issues.concat(payments.issues);

  writeTable_(spreadsheet, '請求一覧',
    ['請求番号', '対象月', '顧客コード', '顧客名', '件数', '10%対象', '10%消費税', '8%対象', '8%消費税', '請求額', '支払期限'],
    invoices.map(function (i) {
      return [i.invoiceNo, i.month, i.customerCode, i.customerName, i.orderCount, i.subtotal10, i.tax10, i.subtotal8, i.tax8, i.total, i.dueDate];
    }));
  writeTable_(spreadsheet, '入金状況',
    ['請求番号', '顧客名', '請求額', '入金額', '残額', '状況', '支払期限', '遅延日数', '督促文の下書き'],
    payments.statuses.map(function (s) {
      return [s.invoiceNo, s.customerName, s.total, s.paid, s.balance, s.status, s.dueDate, s.overdueDays, s.draft];
    }));
  writeTable_(spreadsheet, '確認リスト',
    ['種別', '受注ID・請求番号', '顧客コード', '内容'],
    issues.map(function (issue) { return [issue.code, issue.orderId, issue.customer, issue.detail]; }));

  SpreadsheetApp.getUi().alert(
    '請求 ' + invoices.length + '件を作成しました。確認が必要な項目は ' + issues.length + '件です。');
}

function setupSampleData() {
  var spreadsheet = SpreadsheetApp.getActive();
  if (spreadsheet.getSheetByName('受注一覧')) {
    SpreadsheetApp.getUi().alert('「受注一覧」が既にあるため、上書きせずに中止しました。');
    return;
  }
  writeTable_(spreadsheet, '設定', ['項目', '値'], [['対象月', "'2026-10"], ['基準日', "'2026-12-05"]]);
  writeTable_(spreadsheet, '顧客マスタ', ['顧客コード', '顧客名', '支払日数'], [
    ['C001', '架空商事株式会社', 30],
    ['C002', 'サンプル工業合同会社', 30],
    ['C003', 'テスト物産株式会社', 30],
    ['C004', '架空デザイン事務所', 60],
    ['C005', 'サンプル食品株式会社', 30]
  ]);
  writeTable_(spreadsheet, '受注一覧', ['受注ID', '受注日', '顧客コード', '品目', '数量', '単価', '税率'], [
    ['A-1001', "'2026-10-02", 'C001', '制作費', 1, 120000, 10],
    ['A-1002', "'2026-10-10", 'C001', '軽減税率対象品', 15, 777, 8],
    ['A-1003', "'2026-10-12", 'C002', '保守費', 3, 33333, 10],
    ['A-1004', "'2026-10-15", 'C003', '制作費', '三', 50000, 10],
    ['A-1005', "'2026-10-18", 'C004', '撮影費', 2, 45000, 10],
    ['A-1006', "'2026-10-20", 'C005', '資料作成', 4, 9800, 10],
    ['A-1007', "'2026-10-22", 'C999', '制作費', 1, 10000, 10],
    ['A-1008', "'2026-09-30", 'C001', '前月分（対象外）', 1, 99999, 10]
  ]);
  writeTable_(spreadsheet, '入金一覧', ['請求番号', '入金日', '入金額'], [
    ['INV-202610-C002', "'2026-11-20", 50000],
    ['INV-202610-C777', "'2026-11-21", 12345]
  ]);
  SpreadsheetApp.getUi().alert('サンプルデータを入れました。メニューから「請求と入金状況を作成」を実行してください。');
}
