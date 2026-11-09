/**
 * 請求作成と入金確認の計算処理（サンプル）。
 * Googleのサービスに依存しないため、手元のNode.jsでもテストできる。
 * 日付はすべて 'yyyy-MM-dd' の文字列で扱う。
 */

/** 見出し行をキーにして、シートの2次元配列をオブジェクトの配列にする。 */
function rowsToObjects(values) {
  if (!values || values.length === 0) return [];
  var header = values[0].map(function (cell) { return String(cell).trim(); });
  return values.slice(1)
    .filter(function (row) { return row.some(function (cell) { return cell !== '' && cell !== null; }); })
    .map(function (row) {
      var record = {};
      header.forEach(function (name, index) { record[name] = row[index]; });
      return record;
    });
}

function isValidDateString(text) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return false;
  var date = new Date(text + 'T00:00:00Z');
  return !isNaN(date.getTime()) && date.toISOString().slice(0, 10) === text;
}

function addDays(dateText, days) {
  var date = new Date(dateText + 'T00:00:00Z');
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function lastDayOfMonth(month) {
  var parts = month.split('-').map(Number);
  return new Date(Date.UTC(parts[0], parts[1], 0)).toISOString().slice(0, 10);
}

function daysBetween(fromText, toText) {
  return Math.round((Date.parse(toText + 'T00:00:00Z') - Date.parse(fromText + 'T00:00:00Z')) / 86400000);
}

function isWholeNumber(value, minimum) {
  return typeof value === 'number' && Number.isInteger(value) && value >= minimum;
}

/**
 * 対象月の受注を確認する。
 * 問題のある受注がある顧客の請求は、不完全な請求を出さないよう保留する。
 */
function validateOrders(orders, customers, month) {
  var issues = [];
  var valid = [];
  var heldCustomers = {};
  var idCounts = {};

  orders.forEach(function (order) {
    var id = String(order['受注ID']);
    idCounts[id] = (idCounts[id] || 0) + 1;
  });

  orders.forEach(function (order) {
    var id = String(order['受注ID']);
    var customer = String(order['顧客コード'] || '');
    var dateText = String(order['受注日'] || '');
    var dateOk = isValidDateString(dateText);
    if (dateOk && dateText.slice(0, 7) !== month) return;

    var problems = [];
    if (!dateOk) problems.push(['INVALID_DATE', '受注日が日付ではありません: ' + dateText]);
    if (idCounts[id] > 1) problems.push(['DUPLICATE_ID', '受注IDが重複しています。']);
    if (!customers[customer]) problems.push(['UNKNOWN_CUSTOMER', '顧客マスタにありません: ' + customer]);
    if (!isWholeNumber(order['数量'], 1)) problems.push(['INVALID_QUANTITY', '数量が正の整数ではありません: ' + order['数量']]);
    if (!isWholeNumber(order['単価'], 0)) problems.push(['INVALID_PRICE', '単価が0以上の整数ではありません: ' + order['単価']]);
    if (order['税率'] !== 10 && order['税率'] !== 8) problems.push(['INVALID_TAX', '税率は8か10で入力してください: ' + order['税率']]);

    if (problems.length > 0) {
      if (customer) heldCustomers[customer] = true;
      problems.forEach(function (problem) {
        issues.push({ code: problem[0], orderId: id, customer: customer, detail: problem[1] });
      });
      return;
    }
    valid.push(order);
  });

  return { valid: valid, issues: issues, heldCustomers: heldCustomers };
}

/**
 * 顧客ごとの月次請求を作る。
 * 消費税は適格請求書の考え方に合わせ、請求書ごと・税率ごとに1回だけ切り捨てる。
 */
function buildInvoices(orders, customers, month, heldCustomers) {
  var grouped = {};
  orders.forEach(function (order) {
    var code = order['顧客コード'];
    if (heldCustomers[code]) return;
    grouped[code] = grouped[code] || { 10: 0, 8: 0, count: 0 };
    grouped[code][order['税率']] += order['数量'] * order['単価'];
    grouped[code].count += 1;
  });

  return Object.keys(grouped).sort().map(function (code) {
    var sums = grouped[code];
    var customer = customers[code];
    var tax10 = Math.floor(sums[10] * 10 / 100);
    var tax8 = Math.floor(sums[8] * 8 / 100);
    return {
      invoiceNo: 'INV-' + month.replace('-', '') + '-' + code,
      month: month,
      customerCode: code,
      customerName: customer.name,
      orderCount: sums.count,
      subtotal10: sums[10],
      tax10: tax10,
      subtotal8: sums[8],
      tax8: tax8,
      total: sums[10] + tax10 + sums[8] + tax8,
      dueDate: addDays(lastDayOfMonth(month), customer.paymentDays)
    };
  });
}

/** 請求と入金を突き合わせ、状況と督促文の下書きを作る。メールは送らない。 */
function reconcilePayments(invoices, payments, today) {
  var known = {};
  var received = {};
  var issues = [];
  invoices.forEach(function (invoice) { known[invoice.invoiceNo] = true; });

  payments.forEach(function (payment) {
    var number = String(payment['請求番号'] || '');
    var amount = payment['入金額'];
    if (!isWholeNumber(amount, 1)) {
      issues.push({ code: 'INVALID_PAYMENT', orderId: number, customer: '', detail: '入金額が正の整数ではありません: ' + amount });
      return;
    }
    if (!known[number]) {
      issues.push({ code: 'UNKNOWN_INVOICE', orderId: number, customer: '', detail: '対象月の請求にない入金です: ' + amount + '円' });
      return;
    }
    received[number] = (received[number] || 0) + amount;
  });

  var statuses = invoices.map(function (invoice) {
    var paid = received[invoice.invoiceNo] || 0;
    var balance = invoice.total - paid;
    var status = balance === 0 ? '入金済み' : balance < 0 ? '過入金' : paid > 0 ? '一部入金' : '未入金';
    var overdueDays = balance > 0 && today > invoice.dueDate ? daysBetween(invoice.dueDate, today) : 0;
    return {
      invoiceNo: invoice.invoiceNo,
      customerName: invoice.customerName,
      total: invoice.total,
      paid: paid,
      balance: balance,
      status: status,
      dueDate: invoice.dueDate,
      overdueDays: overdueDays,
      draft: overdueDays > 0 ? buildReminderDraft(invoice, balance) : ''
    };
  });

  return { statuses: statuses, issues: issues };
}

function buildReminderDraft(invoice, balance) {
  return [
    invoice.customerName + ' ご担当者様',
    '',
    'いつもお世話になっております。',
    '下記のご請求について、本日時点でご入金を確認できておりません。',
    '',
    '請求番号: ' + invoice.invoiceNo,
    '未入金額: ' + balance.toLocaleString('ja-JP') + '円',
    'お支払期限: ' + invoice.dueDate,
    '',
    '行き違いでお支払い済みの場合は、ご容赦ください。',
    'ご確認のほど、よろしくお願いいたします。'
  ].join('\n');
}
