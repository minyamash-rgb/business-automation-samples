// Logic.js の計算と、Main.js の処理の流れを検証する。
// Main.js はGoogleの代わりに下の簡易モックで動かしている。実際のGoogle環境での動作確認ではない。
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

class FakeSheet {
  constructor() { this.values = []; }
  getDataRange() { return { getValues: () => this.values.map((r) => [...r]) }; }
  clearContents() { this.values = []; }
  getRange(row, column, rows, columns) {
    return {
      setValues: (table) => {
        assert.equal(table.length, rows, '行数が範囲と一致すること');
        table.forEach((r) => assert.equal(r.length, columns, '列数が範囲と一致すること'));
        // シートの入力と同じく、先頭のアポストロフィは文字列の印として取り除く。
        this.values = table.map((r) => r.map((c) => (typeof c === 'string' ? c.replace(/^'/, '') : c)));
      },
    };
  }
}

function loadApp() {
  const sheets = new Map();
  const alerts = [];
  const spreadsheet = {
    getSheetByName: (name) => sheets.get(name) || null,
    insertSheet: (name) => { const s = new FakeSheet(); sheets.set(name, s); return s; },
  };
  const context = vm.createContext({
    SpreadsheetApp: {
      getActive: () => spreadsheet,
      getUi: () => ({ alert: (m) => alerts.push(m) }),
    },
    Utilities: { formatDate: (d) => d.toISOString().slice(0, 10) },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
  });
  for (const file of ['Logic.js', 'Main.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', file), 'utf8'), context, { filename: file });
  }
  return { app: context, sheets, alerts };
}

const { app: logic } = loadApp();
const customers = { C1: { name: '架空商事', paymentDays: 30 } };
const order = (id, quantity, price, tax, date = '2026-10-05', customer = 'C1') => ({
  受注ID: id, 受注日: date, 顧客コード: customer, 数量: quantity, 単価: price, 税率: tax,
});
// vm内で作られた配列は deepEqual で別物扱いになるため、通常の配列へ変換する。
const records = (sheet) => JSON.parse(JSON.stringify(logic.rowsToObjects(sheet.values)));

test('消費税は税率ごとに合計してから1回だけ切り捨てる', () => {
  const orders = [order('1', 1, 5, 10), order('2', 1, 5, 10), order('3', 1, 1005, 10), order('4', 3, 99, 8)];
  const [invoice] = logic.buildInvoices(orders, customers, '2026-10', {});
  assert.equal(invoice.subtotal10, 1015);
  assert.equal(invoice.tax10, 101);
  assert.equal(invoice.subtotal8, 297);
  assert.equal(invoice.tax8, 23);
  assert.equal(invoice.total, 1436);
});

test('支払期限は月末からの日数で、月またぎとうるう年も正しい', () => {
  const [invoice] = logic.buildInvoices([order('1', 1, 100, 10, '2028-01-10')], customers, '2028-01', {});
  assert.equal(logic.lastDayOfMonth('2028-02'), '2028-02-29');
  assert.equal(invoice.dueDate, '2028-03-01');
});

test('問題のある受注がある顧客は請求全体を保留する', () => {
  const result = logic.validateOrders([order('1', 1, 100, 10), order('2', '三', 100, 10)], customers, '2026-10');
  assert.equal(result.issues[0].code, 'INVALID_QUANTITY');
  assert.equal(result.heldCustomers.C1, true);
  assert.equal(logic.buildInvoices(result.valid, customers, '2026-10', result.heldCustomers).length, 0);
});

for (const [name, bad, code] of [
  ['税率', order('1', 1, 100, 5), 'INVALID_TAX'],
  ['単価', order('1', 1, '1,000', 10), 'INVALID_PRICE'],
  ['顧客', order('1', 1, 100, 10, '2026-10-05', 'X'), 'UNKNOWN_CUSTOMER'],
  ['日付', order('1', 1, 100, 10, '2026-13-40'), 'INVALID_DATE'],
]) {
  test(`不正な${name}を検出する`, () => {
    const result = logic.validateOrders([bad], customers, '2026-10');
    assert.equal(result.issues[0].code, code);
    assert.equal(result.valid.length, 0);
  });
}

test('重複した受注IDは両方を確認対象にする', () => {
  const result = logic.validateOrders([order('1', 1, 100, 10), order('1', 1, 100, 10)], customers, '2026-10');
  assert.equal(result.issues.filter((i) => i.code === 'DUPLICATE_ID').length, 2);
});

test('対象月以外の受注は無視する', () => {
  const result = logic.validateOrders([order('1', 1, 100, 10, '2026-09-30')], customers, '2026-10');
  assert.equal(result.valid.length + result.issues.length, 0);
});

const invoice = { invoiceNo: 'I1', customerName: '架空商事', total: 1100, dueDate: '2026-11-30' };
for (const [paid, status] of [[0, '未入金'], [100, '一部入金'], [1100, '入金済み'], [1200, '過入金']]) {
  test(`${paid}円の入金は「${status}」になる`, () => {
    const payments = paid ? [{ 請求番号: 'I1', 入金額: paid }] : [];
    assert.equal(logic.reconcilePayments([invoice], payments, '2026-12-05').statuses[0].status, status);
  });
}

test('期限当日は遅延にせず、翌日から督促文の下書きを作る', () => {
  assert.equal(logic.reconcilePayments([invoice], [], '2026-11-30').statuses[0].overdueDays, 0);
  const late = logic.reconcilePayments([invoice], [], '2026-12-01').statuses[0];
  assert.equal(late.overdueDays, 1);
  assert.match(late.draft, /1,100円/);
});

test('対象にない請求番号の入金は残額に含めず確認リストに出す', () => {
  const result = logic.reconcilePayments([invoice], [{ 請求番号: 'NG', 入金額: 500 }], '2026-12-05');
  assert.equal(result.issues[0].code, 'UNKNOWN_INVOICE');
  assert.equal(result.statuses[0].paid, 0);
});

test('サンプルデータで一連の処理が動き、再実行しても同じ結果になる', () => {
  const { app, sheets, alerts } = loadApp();
  app.setupSampleData();
  app.runInvoiceWorkflow();
  const first = JSON.stringify([...sheets]);
  app.runInvoiceWorkflow();
  assert.equal(JSON.stringify([...sheets]), first);

  const invoices = records(sheets.get('請求一覧'));
  assert.deepEqual(invoices.map((i) => i['顧客コード']), ['C001', 'C002', 'C004', 'C005']);
  const c001 = invoices[0];
  assert.equal(c001['10%消費税'], 12000);
  assert.equal(c001['8%対象'], 11655);
  assert.equal(c001['8%消費税'], 932);
  assert.equal(c001['請求額'], 144587);

  const status = Object.fromEntries(records(sheets.get('入金状況')).map((s) => [s['請求番号'], s]));
  assert.equal(status['INV-202610-C002']['状況'], '一部入金');
  assert.equal(status['INV-202610-C004']['遅延日数'], 0);
  assert.equal(status['INV-202610-C005']['遅延日数'], 5);
  assert.match(status['INV-202610-C005']['督促文の下書き'], /サンプル食品株式会社/);

  const codes = records(sheets.get('確認リスト')).map((i) => i['種別']).sort();
  assert.deepEqual(codes, ['INVALID_QUANTITY', 'UNKNOWN_CUSTOMER', 'UNKNOWN_INVOICE']);
  assert.match(alerts.at(-1), /請求 4件.*確認が必要な項目は 3件/);
});

test('既存データがあるときサンプル投入は上書きしない', () => {
  const { app, sheets, alerts } = loadApp();
  app.setupSampleData();
  sheets.get('受注一覧').values[1][3] = '利用者が編集した内容';
  app.setupSampleData();
  assert.equal(sheets.get('受注一覧').values[1][3], '利用者が編集した内容');
  assert.match(alerts.at(-1), /上書きせず/);
});

test('対象月の形式が違うときは分かりやすく止める', () => {
  const { app, sheets } = loadApp();
  app.setupSampleData();
  sheets.get('設定').values[1][1] = '10月';
  assert.throws(() => app.runInvoiceWorkflow(), /2026-10 の形式/);
});
