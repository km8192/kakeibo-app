/**
 * 家計簿アプリ バックエンド（Google Apps Script）
 *
 * このファイルは、家計簿を保存するGoogleスプレッドシートに「拡張機能 > Apps Script」で
 * スクリプトエディタを開き、その中に貼り付けて使います（コンテナバインド方式）。
 * デプロイ手順は README.md を参照してください。
 */

// ===== シート定義 =====
var SHEETS = {
  TRANSACTIONS: '取引',
  CATEGORIES: 'カテゴリ',
  MEMBERS: 'メンバー',
  PAYMENT_METHODS: '支払い方法',
  FIXED_COSTS: '固定費',
  MONTHLY_BUDGET: '月次予算',
  SETTINGS: '設定'
};

var HEADERS = {};
HEADERS[SHEETS.TRANSACTIONS] = ['ID', '日付', '種別', '誰が', '支払い方法', 'カテゴリ', '金額', 'メモ', '登録日時'];
HEADERS[SHEETS.CATEGORIES] = ['ID', 'カテゴリ名', '種別', 'アイコン', '表示順', '予算上限（月）'];
HEADERS[SHEETS.MEMBERS] = ['ID', 'メンバー名', '表示順'];
HEADERS[SHEETS.PAYMENT_METHODS] = ['ID', '支払い方法名', '表示順'];
HEADERS[SHEETS.FIXED_COSTS] = ['ID', '名称', '誰の', '支払い方法', '紐づけるカテゴリ', '金額（月）', '自動更新', '表示順'];
HEADERS[SHEETS.MONTHLY_BUDGET] = ['年月', '金額'];
HEADERS[SHEETS.SETTINGS] = ['締め日'];

var TIMEZONE = 'Asia/Tokyo';

// ===== エントリーポイント =====

function doGet(e) {
  return handleRequest(e.parameter.action, e.parameter);
}

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    body = {};
  }
  return handleRequest(body.action, body.payload || {});
}

function handleRequest(action, params) {
  _sheetDataCache = {}; // 1リクエストの中で同じシートを何度も読み直さないためのキャッシュ
  ensureSheets();
  preloadSheetsBatch();
  var result;
  try {
    switch (action) {
      case 'getInitialData': result = apiGetInitialData(params); break;
      case 'getTransactions': result = apiGetTransactions(params); break;
      case 'addTransaction': result = apiAddTransaction(params); break;
      case 'updateTransaction': result = apiUpdateTransaction(params); break;
      case 'deleteTransaction': result = apiDeleteTransaction(params); break;

      case 'getMasters': result = apiGetMasters(); break;

      case 'addCategory': result = apiAddCategory(params); break;
      case 'updateCategory': result = apiUpdateCategory(params); break;
      case 'deleteCategory': result = apiDeleteCategory(params); break;

      case 'addMember': result = apiAddMember(params); break;
      case 'updateMember': result = apiUpdateMember(params); break;
      case 'deleteMember': result = apiDeleteMember(params); break;
      case 'reorderMembers': result = apiReorderMembers(params); break;

      case 'addPaymentMethod': result = apiAddPaymentMethod(params); break;
      case 'updatePaymentMethod': result = apiUpdatePaymentMethod(params); break;
      case 'deletePaymentMethod': result = apiDeletePaymentMethod(params); break;
      case 'reorderPaymentMethods': result = apiReorderPaymentMethods(params); break;

      case 'getFixedCosts': result = apiGetFixedCosts(params); break;
      case 'addFixedCost': result = apiAddFixedCost(params); break;
      case 'updateFixedCost': result = apiUpdateFixedCost(params); break;
      case 'deleteFixedCost': result = apiDeleteFixedCost(params); break;
      case 'reorderFixedCosts': result = apiReorderFixedCosts(params); break;

      case 'getAverageAmount': result = apiGetAverageAmount(params); break;
      case 'getAverageAmountsForMember': result = apiGetAverageAmountsForMember(params); break;

      case 'getMonthlyBudget': result = apiGetMonthlyBudget(params); break;
      case 'setMonthlyBudget': result = apiSetMonthlyBudget(params); break;
      case 'getWeeklyBreakdown': result = apiGetWeeklyBreakdown(params); break;

      case 'getSettings': result = apiGetSettings(); break;
      case 'updateSettings': result = apiUpdateSettings(params); break;

      case 'getReport': result = apiGetReport(params); break;
      case 'getCategoryYearComparison': result = apiGetCategoryYearComparison(params); break;

      default:
        throw new Error('不明なaction: ' + action);
    }
    return jsonResponse({ ok: true, data: result });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ===== 初期化（シート作成・シード投入） =====

function ensureSheets() {
  // 初期化済みなら、毎回シートの有無を確認する重い処理は丸ごとスキップする
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('sheetsInitialized') === 'true') return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(HEADERS).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    var isNew = false;
    if (!sheet) {
      sheet = ss.insertSheet(name);
      isNew = true;
    }
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
      sheet.setFrozenRows(1);
      isNew = true;
    }
    if (isNew) {
      seedSheet(name, sheet);
    }
  });

  // 「年月」列がスプレッドシート側で日付型に自動変換されないよう、プレーンテキスト形式に固定する
  ss.getSheetByName(SHEETS.MONTHLY_BUDGET).getRange('A2:A1000').setNumberFormat('@');

  // デフォルトシート「シート1」が残っていれば削除（新規スプレッドシートの初期化時のみ）
  var defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  props.setProperty('sheetsInitialized', 'true');
}

function seedSheet(name, sheet) {
  if (name === SHEETS.CATEGORIES && sheet.getLastRow() === 1) {
    var expense = [
      ['家賃', '支出', 'ti-home', 1, 80000],
      ['食費', '支出', 'ti-tools-kitchen-2', 2, 30000],
      ['交通費', '支出', 'ti-bus', 3, 5000],
      ['日用品', '支出', 'ti-shopping-cart', 4, 5000],
      ['交際費', '支出', 'ti-users', 5, 5000],
      ['電気代', '支出', 'ti-bolt', 6, ''],
      ['ガス代', '支出', 'ti-flame', 7, ''],
      ['水道代', '支出', 'ti-droplet', 8, ''],
      ['その他', '支出', 'ti-dots', 9, 5000]
    ];
    var income = [
      ['給与', '収入', 'ti-cash', 1, ''],
      ['その他', '収入', 'ti-dots', 2, '']
    ];
    expense.concat(income).forEach(function (row) {
      sheet.appendRow([generateId('c'), row[0], row[1], row[2], row[3], row[4]]);
    });
  }
  if (name === SHEETS.MEMBERS && sheet.getLastRow() === 1) {
    [['パパ', 1], ['ママ', 2]].forEach(function (row) {
      sheet.appendRow([generateId('m'), row[0], row[1]]);
    });
  }
  if (name === SHEETS.PAYMENT_METHODS && sheet.getLastRow() === 1) {
    [['現金', 1], ['クレジット', 2]].forEach(function (row) {
      sheet.appendRow([generateId('p'), row[0], row[1]]);
    });
  }
  if (name === SHEETS.SETTINGS && sheet.getLastRow() === 1) {
    sheet.appendRow(['']);
  }
}

function generateId(prefix) {
  return prefix + '_' + Utilities.getUuid();
}

// ===== シート読み書き共通ヘルパー =====

// 1回のAPI呼び出しの中で同じシートを何度も読み直すと遅くなるため、リクエスト単位でキャッシュする
// （handleRequestの先頭で毎回空にしているので、リクエストをまたいで古いデータが残ることはない）
var _sheetDataCache = {};

function getSheetData(sheetName) {
  if (_sheetDataCache[sheetName]) return _sheetDataCache[sheetName];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = values[i][j];
    }
    obj.__row = i + 1;
    rows.push(obj);
  }
  var data = { sheet: sheet, headers: headers, rows: rows };
  _sheetDataCache[sheetName] = data;
  return data;
}

// 項目名（ヘッダー）はHEADERS定数で分かっているため、書き込み専用の処理では
// シートの全データを読み込まずにシート参照だけを取得する（履歴が増えても遅くならないように）
function getSheetRef(sheetName) {
  if (_sheetDataCache[sheetName] && _sheetDataCache[sheetName].sheet) return _sheetDataCache[sheetName].sheet;
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

// 日付を含まないシートは、Sheets API（拡張サービス）のbatchGetで1回の通信にまとめて読み込む。
// 「取引」シートだけは日付列の型の扱いが変わってしまうリスクがあるため、対象に含めず
// 従来通りSpreadsheetAppで個別に読み込む（getSheetDataのフォールバックに任せる）。
// 拡張サービス「Sheets API」が未有効化の環境でも壊れないよう、失敗時は何もせず
// 通常の個別読み込みにフォールバックする。
var BATCHABLE_SHEETS = [SHEETS.CATEGORIES, SHEETS.MEMBERS, SHEETS.PAYMENT_METHODS, SHEETS.FIXED_COSTS, SHEETS.MONTHLY_BUDGET, SHEETS.SETTINGS];

function preloadSheetsBatch() {
  try {
    var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var res = Sheets.Spreadsheets.Values.batchGet(ssId, {
      ranges: BATCHABLE_SHEETS,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    res.valueRanges.forEach(function (vr, idx) {
      var name = BATCHABLE_SHEETS[idx];
      var values = vr.values || [];
      var headers = HEADERS[name];
      var rows = [];
      for (var i = 1; i < values.length; i++) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
          obj[headers[j]] = values[i][j] !== undefined ? values[i][j] : '';
        }
        obj.__row = i + 1;
        rows.push(obj);
      }
      _sheetDataCache[name] = { sheet: null, headers: headers, rows: rows };
    });
  } catch (err) {
    // Sheets APIが未有効化などの場合はここで何もしない（各APIが必要な時に個別に読みに行く）
  }
}

function appendRowByHeaders(sheetName, obj) {
  var headers = HEADERS[sheetName];
  var row = headers.map(function (h) {
    return obj[h] !== undefined ? obj[h] : '';
  });
  getSheetRef(sheetName).appendRow(row);
  delete _sheetDataCache[sheetName];
  return row;
}

function updateRowByHeaders(sheetName, rowNumber, obj) {
  var headers = HEADERS[sheetName];
  var row = headers.map(function (h) {
    return obj[h] !== undefined ? obj[h] : '';
  });
  getSheetRef(sheetName).getRange(rowNumber, 1, 1, headers.length).setValues([row]);
  delete _sheetDataCache[sheetName];
}

function deleteRowByRowNumber(sheetName, rowNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName(sheetName).deleteRow(rowNumber);
  delete _sheetDataCache[sheetName];
}

function findRowById(rows, id) {
  for (var i = 0; i < rows.length; i++) {
    if (rows[i]['ID'] === id) return rows[i];
  }
  return null;
}

function formatDate(date) {
  return Utilities.formatDate(new Date(date), TIMEZONE, 'yyyy-MM-dd');
}

function formatDateTime(date) {
  return Utilities.formatDate(new Date(date), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

// ===== 会計月（締め日）ロジック =====

function getClosingDay() {
  var rows = getSheetData(SHEETS.SETTINGS).rows;
  if (rows.length === 0) return 31;
  var d = rows[0]['締め日'];
  d = Number(d);
  if (!d || d < 1 || d > 31) return 31;
  return d;
}

// その年月における「実際の締め日」を返す（例：31日締め設定・2月なら28日/29日になる）
function closingDateOf(year, monthIndex0, closingDay) {
  var lastDay = new Date(year, monthIndex0 + 1, 0).getDate();
  var day = Math.min(closingDay, lastDay);
  return new Date(year, monthIndex0, day, 23, 59, 59, 999);
}

function fmtYm(y, m) {
  return y + '-' + (m < 10 ? '0' + m : String(m));
}

// スプレッドシート側が「年月」列（例："2026-09"）を日付型に自動変換してしまうことがあるため、
// Dateで返ってきた場合でも文字列キーに正規化してから比較する
function normalizeYearMonthKey(value) {
  if (value instanceof Date) {
    return fmtYm(value.getFullYear(), value.getMonth() + 1);
  }
  return String(value);
}

// 指定した日付が属する会計月（年月キー・開始日・終了日）を返す
function getFiscalMonth(dateInput, closingDay) {
  var date = new Date(dateInput);
  date = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0, 0);

  if (!closingDay || closingDay >= 31) {
    var y0 = date.getFullYear(), m0 = date.getMonth();
    return { key: fmtYm(y0, m0 + 1), startDate: new Date(y0, m0, 1), endDate: new Date(y0, m0 + 1, 0, 23, 59, 59, 999) };
  }

  var y = date.getFullYear(), m = date.getMonth();
  var thisClose = closingDateOf(y, m, closingDay);
  var belongY = y, belongM = m;
  if (date.getTime() > thisClose.getTime()) {
    belongM = m + 1;
    if (belongM > 11) { belongM = 0; belongY = y + 1; }
  }
  return fiscalMonthRangeOf(belongY, belongM, closingDay);
}

function fiscalMonthRangeOf(belongY, belongM, closingDay) {
  var endDate = closingDateOf(belongY, belongM, closingDay);
  var prevM = belongM - 1, prevY = belongY;
  if (prevM < 0) { prevM = 11; prevY = belongY - 1; }
  var prevClose = closingDateOf(prevY, prevM, closingDay);
  var startDate = new Date(prevClose.getFullYear(), prevClose.getMonth(), prevClose.getDate() + 1, 0, 0, 0, 0);
  return { key: fmtYm(belongY, belongM + 1), startDate: startDate, endDate: endDate };
}

// 「2026-09」のような年月キーから、その会計月の開始日・終了日を求める
function fiscalMonthRangeForKey(key, closingDay) {
  var parts = key.split('-');
  var y = Number(parts[0]), m = Number(parts[1]) - 1;
  if (!closingDay || closingDay >= 31) {
    return { key: key, startDate: new Date(y, m, 1), endDate: new Date(y, m + 1, 0, 23, 59, 59, 999) };
  }
  return fiscalMonthRangeOf(y, m, closingDay);
}

function shiftFiscalMonthKey(key, delta) {
  var parts = key.split('-');
  var y = Number(parts[0]), m = Number(parts[1]) - 1;
  m += delta;
  y += Math.floor(m / 12);
  m = ((m % 12) + 12) % 12;
  return fmtYm(y, m + 1);
}

// ===== 初期データ一括取得 =====

function apiGetInitialData(params) {
  var closingDay = getClosingDay();
  var monthKey = params && params.monthKey ? params.monthKey : getFiscalMonth(new Date(), closingDay).key;

  return {
    masters: apiGetMasters(),
    settings: { closingDay: closingDay },
    fixedCosts: apiGetFixedCosts({}),
    monthKey: monthKey,
    transactions: apiGetTransactions({ monthKey: monthKey }),
    monthlyBudget: apiGetMonthlyBudget({ monthKey: monthKey })
  };
}

// ===== 取引 =====

function computeSummaryForMonth(monthKey) {
  var closingDay = getClosingDay();
  var range = fiscalMonthRangeForKey(monthKey, closingDay);
  var txRows = getSheetData(SHEETS.TRANSACTIONS).rows.filter(function (r) {
    var d = new Date(r['日付']);
    return d.getTime() >= range.startDate.getTime() && d.getTime() <= range.endDate.getTime();
  });

  var income = 0;
  txRows.forEach(function (r) {
    if (r['種別'] === '収入') income += Number(r['金額']) || 0;
  });

  var fixedCostTotal = 0;
  getSheetData(SHEETS.FIXED_COSTS).rows.forEach(function (r) {
    fixedCostTotal += Number(r['金額（月）']) || 0;
  });

  // 支出＝「固定費に一致しない実際の支出」＋「登録済みの固定費合計」。
  // 家賃などを実際に取引として記録済みでも、その取引は固定費に一致するため
  // variableExpenseの時点で除外され、代わりに固定費側の金額が使われるので二重計上にならない
  var variableExpense = computeVariableExpense(txRows);
  var expense = variableExpense + fixedCostTotal;

  var balance = income - expense;
  var availableBalance = income - fixedCostTotal;

  var budgetRow = getSheetData(SHEETS.MONTHLY_BUDGET).rows.filter(function (r) { return normalizeYearMonthKey(r['年月']) === monthKey; })[0];
  var budgetAmount = budgetRow ? Number(budgetRow['金額']) || 0 : 0;

  return {
    monthKey: monthKey,
    income: income,
    expense: expense,
    balance: balance,
    fixedCostTotal: fixedCostTotal,
    availableBalance: availableBalance,
    budget: {
      amount: budgetAmount,
      used: variableExpense,
      remaining: budgetAmount - variableExpense,
      usageRate: budgetAmount > 0 ? variableExpense / budgetAmount : 0
    }
  };
}

// 固定費に「紐づくカテゴリ×誰が」の組み合わせに一致する支出は、ざっくり予算の対象から除外する
function computeVariableExpense(txRows) {
  var fixedPairs = getFixedCostPairs();
  var total = 0;
  txRows.forEach(function (r) {
    if (r['種別'] !== '支出') return;
    var pairKey = r['カテゴリ'] + '|' + r['誰が'];
    if (fixedPairs[pairKey]) return;
    total += Number(r['金額']) || 0;
  });
  return total;
}

function getFixedCostPairs() {
  var pairs = {};
  getSheetData(SHEETS.FIXED_COSTS).rows.forEach(function (r) {
    pairs[r['紐づけるカテゴリ'] + '|' + r['誰の']] = true;
  });
  return pairs;
}

function apiGetTransactions(params) {
  params = params || {};
  var closingDay = getClosingDay();
  var monthKey = params.monthKey || getFiscalMonth(new Date(), closingDay).key;
  var range = fiscalMonthRangeForKey(monthKey, closingDay);

  var rows = getSheetData(SHEETS.TRANSACTIONS).rows.filter(function (r) {
    var d = new Date(r['日付']);
    return d.getTime() >= range.startDate.getTime() && d.getTime() <= range.endDate.getTime();
  });

  if (params.category) rows = rows.filter(function (r) { return r['カテゴリ'] === params.category; });
  if (params.member) rows = rows.filter(function (r) { return r['誰が'] === params.member; });
  if (params.paymentMethod) rows = rows.filter(function (r) { return r['支払い方法'] === params.paymentMethod; });
  if (params.keyword) {
    var kw = String(params.keyword).toLowerCase();
    rows = rows.filter(function (r) {
      return String(r['メモ']).toLowerCase().indexOf(kw) >= 0 || String(r['金額']).indexOf(kw) >= 0;
    });
  }

  rows.sort(function (a, b) {
    var da = new Date(a['日付']).getTime(), db = new Date(b['日付']).getTime();
    if (db !== da) return db - da;
    return new Date(b['登録日時']).getTime() - new Date(a['登録日時']).getTime();
  });

  var list = rows.map(function (r) {
    return {
      id: r['ID'],
      date: formatDate(r['日付']),
      type: r['種別'],
      member: r['誰が'],
      paymentMethod: r['支払い方法'],
      category: r['カテゴリ'],
      amount: Number(r['金額']) || 0,
      memo: r['メモ'] || ''
    };
  });

  return {
    monthKey: monthKey,
    list: list,
    summary: computeSummaryForMonth(monthKey)
  };
}

function apiAddTransaction(p) {
  var amount = Number(p.amount);
  if (!p.date || !p.type || !p.member || !p.paymentMethod || !p.category || !amount || amount <= 0) {
    throw new Error('入力内容が不足しています');
  }
  var id = generateId('t');
  appendRowByHeaders(SHEETS.TRANSACTIONS, {
    'ID': id,
    '日付': p.date,
    '種別': p.type,
    '誰が': p.member,
    '支払い方法': p.paymentMethod,
    'カテゴリ': p.category,
    '金額': amount,
    'メモ': p.memo || '',
    '登録日時': formatDateTime(new Date())
  });

  if (p.type === '支出') {
    applyFixedCostAutoUpdate(p.category, p.member, amount);
  }

  var closingDay = getClosingDay();
  var monthKey = getFiscalMonth(p.date, closingDay).key;
  // 画面の再描画に必要な情報をまとめて返し、保存直後にもう一度通信しなくて済むようにする
  return { id: id, transactions: apiGetTransactions({ monthKey: monthKey }) };
}

function applyFixedCostAutoUpdate(category, member, amount) {
  var data = getSheetData(SHEETS.FIXED_COSTS);
  data.rows.forEach(function (r) {
    if (r['紐づけるカテゴリ'] === category && r['誰の'] === member && (r['自動更新'] === true || r['自動更新'] === 'TRUE')) {
      updateRowByHeaders(SHEETS.FIXED_COSTS, r.__row, Object.assign({}, r, { '金額（月）': amount }));
    }
  });
}

function apiUpdateTransaction(p) {
  var data = getSheetData(SHEETS.TRANSACTIONS);
  var row = findRowById(data.rows, p.id);
  if (!row) throw new Error('取引が見つかりません');
  var amount = Number(p.amount);
  if (!p.date || !p.type || !p.member || !p.paymentMethod || !p.category || !amount || amount <= 0) {
    throw new Error('入力内容が不足しています');
  }
  updateRowByHeaders(SHEETS.TRANSACTIONS, row.__row, {
    'ID': row['ID'],
    '日付': p.date,
    '種別': p.type,
    '誰が': p.member,
    '支払い方法': p.paymentMethod,
    'カテゴリ': p.category,
    '金額': amount,
    'メモ': p.memo || '',
    '登録日時': row['登録日時']
  });
  var closingDay = getClosingDay();
  var monthKey = getFiscalMonth(p.date, closingDay).key;
  return { transactions: apiGetTransactions({ monthKey: monthKey }) };
}

function apiDeleteTransaction(p) {
  var data = getSheetData(SHEETS.TRANSACTIONS);
  var row = findRowById(data.rows, p.id);
  if (!row) throw new Error('取引が見つかりません');
  var closingDay = getClosingDay();
  var monthKey = getFiscalMonth(row['日付'], closingDay).key;
  deleteRowByRowNumber(SHEETS.TRANSACTIONS, row.__row);
  return { transactions: apiGetTransactions({ monthKey: monthKey }) };
}

// ===== マスタ一括取得 =====

function apiGetMasters() {
  var categories = getSheetData(SHEETS.CATEGORIES).rows
    .map(function (r) {
      return { id: r['ID'], name: r['カテゴリ名'], type: r['種別'], icon: r['アイコン'], order: Number(r['表示順']) || 0, budgetLimit: r['予算上限（月）'] === '' ? null : Number(r['予算上限（月）']) };
    })
    .sort(function (a, b) { return a.order - b.order; });

  var members = getSheetData(SHEETS.MEMBERS).rows
    .map(function (r) { return { id: r['ID'], name: r['メンバー名'], order: Number(r['表示順']) || 0 }; })
    .sort(function (a, b) { return a.order - b.order; });

  var paymentMethods = getSheetData(SHEETS.PAYMENT_METHODS).rows
    .map(function (r) { return { id: r['ID'], name: r['支払い方法名'], order: Number(r['表示順']) || 0 }; })
    .sort(function (a, b) { return a.order - b.order; });

  return { categories: categories, members: members, paymentMethods: paymentMethods };
}

// ===== カテゴリ =====

function nextOrder(sheetName) {
  var rows = getSheetData(sheetName).rows;
  var max = 0;
  rows.forEach(function (r) { max = Math.max(max, Number(r['表示順']) || 0); });
  return max + 1;
}

function apiAddCategory(p) {
  if (!p.name || !p.type) throw new Error('カテゴリ名・種別は必須です');
  var id = generateId('c');
  appendRowByHeaders(SHEETS.CATEGORIES, {
    'ID': id, 'カテゴリ名': p.name, '種別': p.type, 'アイコン': p.icon || 'ti-dots',
    '表示順': nextOrder(SHEETS.CATEGORIES), '予算上限（月）': p.budgetLimit || ''
  });
  return apiGetMasters();
}

function apiUpdateCategory(p) {
  var data = getSheetData(SHEETS.CATEGORIES);
  var row = findRowById(data.rows, p.id);
  if (!row) throw new Error('カテゴリが見つかりません');
  updateRowByHeaders(SHEETS.CATEGORIES, row.__row, {
    'ID': row['ID'],
    'カテゴリ名': p.name !== undefined ? p.name : row['カテゴリ名'],
    '種別': p.type !== undefined ? p.type : row['種別'],
    'アイコン': p.icon !== undefined ? p.icon : row['アイコン'],
    '表示順': p.order !== undefined ? p.order : row['表示順'],
    '予算上限（月）': p.budgetLimit !== undefined ? (p.budgetLimit || '') : row['予算上限（月）']
  });
  return apiGetMasters();
}

function apiDeleteCategory(p) {
  var data = getSheetData(SHEETS.CATEGORIES);
  var row = findRowById(data.rows, p.id);
  if (!row) throw new Error('カテゴリが見つかりません');
  deleteRowByRowNumber(SHEETS.CATEGORIES, row.__row);
  return apiGetMasters();
}

// ===== メンバー =====

function apiAddMember(p) {
  if (!p.name) throw new Error('メンバー名は必須です');
  appendRowByHeaders(SHEETS.MEMBERS, { 'ID': generateId('m'), 'メンバー名': p.name, '表示順': nextOrder(SHEETS.MEMBERS) });
  return apiGetMasters();
}

function apiUpdateMember(p) {
  var data = getSheetData(SHEETS.MEMBERS);
  var row = findRowById(data.rows, p.id);
  if (!row) throw new Error('メンバーが見つかりません');
  updateRowByHeaders(SHEETS.MEMBERS, row.__row, { 'ID': row['ID'], 'メンバー名': p.name || row['メンバー名'], '表示順': row['表示順'] });
  return apiGetMasters();
}

function apiDeleteMember(p) {
  var data = getSheetData(SHEETS.MEMBERS);
  var row = findRowById(data.rows, p.id);
  if (!row) throw new Error('メンバーが見つかりません');
  deleteRowByRowNumber(SHEETS.MEMBERS, row.__row);
  return apiGetMasters();
}

function apiReorderMembers(p) {
  applyOrder(SHEETS.MEMBERS, p.ids);
  return apiGetMasters();
}

// ===== 支払い方法 =====

function apiAddPaymentMethod(p) {
  if (!p.name) throw new Error('支払い方法名は必須です');
  appendRowByHeaders(SHEETS.PAYMENT_METHODS, { 'ID': generateId('p'), '支払い方法名': p.name, '表示順': nextOrder(SHEETS.PAYMENT_METHODS) });
  return apiGetMasters();
}

function apiUpdatePaymentMethod(p) {
  var data = getSheetData(SHEETS.PAYMENT_METHODS);
  var row = findRowById(data.rows, p.id);
  if (!row) throw new Error('支払い方法が見つかりません');
  updateRowByHeaders(SHEETS.PAYMENT_METHODS, row.__row, { 'ID': row['ID'], '支払い方法名': p.name || row['支払い方法名'], '表示順': row['表示順'] });
  return apiGetMasters();
}

function apiDeletePaymentMethod(p) {
  var data = getSheetData(SHEETS.PAYMENT_METHODS);
  var row = findRowById(data.rows, p.id);
  if (!row) throw new Error('支払い方法が見つかりません');
  deleteRowByRowNumber(SHEETS.PAYMENT_METHODS, row.__row);
  return apiGetMasters();
}

function apiReorderPaymentMethods(p) {
  applyOrder(SHEETS.PAYMENT_METHODS, p.ids);
  return apiGetMasters();
}

function applyOrder(sheetName, ids) {
  var data = getSheetData(sheetName);
  ids.forEach(function (id, index) {
    var row = findRowById(data.rows, id);
    if (row) {
      var obj = Object.assign({}, row);
      obj['表示順'] = index + 1;
      updateRowByHeaders(sheetName, row.__row, obj);
    }
  });
}

// ===== 固定費 =====

function apiGetFixedCosts(params) {
  params = params || {};
  var rows = getSheetData(SHEETS.FIXED_COSTS).rows;
  if (params.paymentMethod) {
    rows = rows.filter(function (r) { return r['支払い方法'] === params.paymentMethod; });
  }
  rows.sort(function (a, b) { return (Number(a['表示順']) || 0) - (Number(b['表示順']) || 0); });

  var list = rows.map(function (r) {
    return {
      id: r['ID'], name: r['名称'], member: r['誰の'], paymentMethod: r['支払い方法'],
      category: r['紐づけるカテゴリ'], amount: Number(r['金額（月）']) || 0,
      autoUpdate: (r['自動更新'] === true || r['自動更新'] === 'TRUE'), order: Number(r['表示順']) || 0
    };
  });

  var subtotalsByMember = {};
  var total = 0;
  list.forEach(function (item) {
    subtotalsByMember[item.member] = (subtotalsByMember[item.member] || 0) + item.amount;
    total += item.amount;
  });

  return { list: list, subtotalsByMember: subtotalsByMember, total: total };
}

function apiAddFixedCost(p) {
  if (!p.name || !p.member || !p.paymentMethod || !p.category) throw new Error('入力内容が不足しています');
  appendRowByHeaders(SHEETS.FIXED_COSTS, {
    'ID': generateId('f'), '名称': p.name, '誰の': p.member, '支払い方法': p.paymentMethod,
    '紐づけるカテゴリ': p.category, '金額（月）': Number(p.amount) || 0,
    '自動更新': !!p.autoUpdate, '表示順': nextOrder(SHEETS.FIXED_COSTS)
  });
  return apiGetFixedCosts({});
}

function apiUpdateFixedCost(p) {
  var data = getSheetData(SHEETS.FIXED_COSTS);
  var row = findRowById(data.rows, p.id);
  if (!row) throw new Error('固定費が見つかりません');
  updateRowByHeaders(SHEETS.FIXED_COSTS, row.__row, {
    'ID': row['ID'],
    '名称': p.name !== undefined ? p.name : row['名称'],
    '誰の': p.member !== undefined ? p.member : row['誰の'],
    '支払い方法': p.paymentMethod !== undefined ? p.paymentMethod : row['支払い方法'],
    '紐づけるカテゴリ': p.category !== undefined ? p.category : row['紐づけるカテゴリ'],
    '金額（月）': p.amount !== undefined ? Number(p.amount) || 0 : row['金額（月）'],
    '自動更新': p.autoUpdate !== undefined ? !!p.autoUpdate : row['自動更新'],
    '表示順': row['表示順']
  });
  return apiGetFixedCosts({});
}

function apiDeleteFixedCost(p) {
  var data = getSheetData(SHEETS.FIXED_COSTS);
  var row = findRowById(data.rows, p.id);
  if (!row) throw new Error('固定費が見つかりません');
  deleteRowByRowNumber(SHEETS.FIXED_COSTS, row.__row);
  return apiGetFixedCosts({});
}

function apiReorderFixedCosts(p) {
  applyOrder(SHEETS.FIXED_COSTS, p.ids);
  return apiGetFixedCosts({});
}

// ===== 過去の平均額 =====

function apiGetAverageAmount(p) {
  var rows = getSheetData(SHEETS.TRANSACTIONS).rows.filter(function (r) {
    return r['誰が'] === p.member && r['カテゴリ'] === p.category;
  });
  if (rows.length === 0) return { average: 0, count: 0 };
  var sum = 0;
  rows.forEach(function (r) { sum += Number(r['金額']) || 0; });
  return { average: Math.round(sum / rows.length), count: rows.length };
}

// カテゴリボタンの数だけ通信するとApps Scriptの応答が遅くなるため、
// 誰が×種別を指定して、対象カテゴリすべての平均額を1回の通信でまとめて返す
function apiGetAverageAmountsForMember(p) {
  var rows = getSheetData(SHEETS.TRANSACTIONS).rows.filter(function (r) {
    return r['誰が'] === p.member && r['種別'] === (p.type || '支出');
  });
  var sums = {}, counts = {};
  rows.forEach(function (r) {
    var cat = r['カテゴリ'];
    sums[cat] = (sums[cat] || 0) + (Number(r['金額']) || 0);
    counts[cat] = (counts[cat] || 0) + 1;
  });
  var result = {};
  Object.keys(counts).forEach(function (cat) {
    result[cat] = { average: Math.round(sums[cat] / counts[cat]), count: counts[cat] };
  });
  return result;
}

// ===== 今月の予算（ざっくり） =====

function apiGetMonthlyBudget(p) {
  var monthKey = p.monthKey;
  var row = getSheetData(SHEETS.MONTHLY_BUDGET).rows.filter(function (r) { return normalizeYearMonthKey(r['年月']) === monthKey; })[0];
  return { monthKey: monthKey, amount: row ? Number(row['金額']) || 0 : null, summary: computeSummaryForMonth(monthKey) };
}

function apiSetMonthlyBudget(p) {
  var data = getSheetData(SHEETS.MONTHLY_BUDGET);
  var row = data.rows.filter(function (r) { return normalizeYearMonthKey(r['年月']) === p.monthKey; })[0];
  var amount = Number(p.amount) || 0;
  if (row) {
    updateRowByHeaders(SHEETS.MONTHLY_BUDGET, row.__row, { '年月': p.monthKey, '金額': amount });
    row['金額'] = amount;
  } else {
    appendRowByHeaders(SHEETS.MONTHLY_BUDGET, { '年月': p.monthKey, '金額': amount });
    data.rows.push({ '年月': p.monthKey, '金額': amount, __row: data.rows.length + 2 });
  }
  // 書き込み直後にもう一度シートを読み直さずに済むよう、メモリ上のキャッシュを更新済みの内容で復元する
  // （updateRowByHeaders/appendRowByHeadersが書き込み後にキャッシュを削除するため）
  _sheetDataCache[SHEETS.MONTHLY_BUDGET] = data;
  return { monthKey: p.monthKey, amount: amount, summary: computeSummaryForMonth(p.monthKey) };
}

// ===== 週別内訳 =====

function apiGetWeeklyBreakdown(p) {
  var closingDay = getClosingDay();
  var monthKey = p.monthKey;
  var range = fiscalMonthRangeForKey(monthKey, closingDay);
  var budgetRow = getSheetData(SHEETS.MONTHLY_BUDGET).rows.filter(function (r) { return normalizeYearMonthKey(r['年月']) === monthKey; })[0];
  var budgetAmount = budgetRow ? Number(budgetRow['金額']) || 0 : 0;

  var totalDays = Math.round((range.endDate - range.startDate) / 86400000) + 1;

  var txRows = getSheetData(SHEETS.TRANSACTIONS).rows.filter(function (r) {
    var d = new Date(r['日付']);
    return d.getTime() >= range.startDate.getTime() && d.getTime() <= range.endDate.getTime();
  });
  var fixedPairs = getFixedCostPairs();

  // 日曜始まりの暦週に分割
  var weeks = [];
  var cursor = new Date(range.startDate);
  while (cursor.getTime() <= range.endDate.getTime()) {
    var weekStart = new Date(cursor);
    var weekEndCandidate = new Date(cursor);
    weekEndCandidate.setDate(weekEndCandidate.getDate() + (6 - weekEndCandidate.getDay()));
    var weekEnd = weekEndCandidate.getTime() < range.endDate.getTime() ? weekEndCandidate : new Date(range.endDate);
    weekEnd.setHours(23, 59, 59, 999);
    weeks.push({ start: new Date(weekStart), end: new Date(weekEnd) });
    cursor = new Date(weekEnd);
    cursor.setDate(cursor.getDate() + 1);
    cursor.setHours(0, 0, 0, 0);
  }

  var result = weeks.map(function (w) {
    var daysInWeek = Math.round((w.end - w.start) / 86400000) + 1;
    var weekBudget = totalDays > 0 ? Math.round(budgetAmount * (daysInWeek / totalDays)) : 0;
    var actual = 0;
    txRows.forEach(function (r) {
      if (r['種別'] !== '支出') return;
      var pairKey = r['カテゴリ'] + '|' + r['誰が'];
      if (fixedPairs[pairKey]) return;
      var d = new Date(r['日付']);
      if (d.getTime() >= w.start.getTime() && d.getTime() <= w.end.getTime()) {
        actual += Number(r['金額']) || 0;
      }
    });
    return {
      label: formatDate(w.start) + '〜' + formatDate(w.end),
      startDate: formatDate(w.start),
      endDate: formatDate(w.end),
      budgetAmount: weekBudget,
      actualAmount: actual,
      usageRate: weekBudget > 0 ? actual / weekBudget : 0
    };
  });

  return { monthKey: monthKey, totalBudget: budgetAmount, weeks: result };
}

// ===== 設定（締め日） =====

function apiGetSettings() {
  return { closingDay: getClosingDay() === 31 ? null : getClosingDay() };
}

function apiUpdateSettings(p) {
  var data = getSheetData(SHEETS.SETTINGS);
  var closingDay = p.closingDay === null || p.closingDay === undefined || p.closingDay === '' ? '' : Number(p.closingDay);
  if (data.rows.length > 0) {
    updateRowByHeaders(SHEETS.SETTINGS, data.rows[0].__row, { '締め日': closingDay });
  } else {
    appendRowByHeaders(SHEETS.SETTINGS, { '締め日': closingDay });
  }
  return apiGetSettings();
}

// ===== 集計・グラフ =====

function apiGetReport(p) {
  var closingDay = getClosingDay();
  var monthKey = p.monthKey || getFiscalMonth(new Date(), closingDay).key;
  var prevMonthKey = shiftFiscalMonthKey(monthKey, -1);

  var current = computeSummaryForMonth(monthKey);
  var prev = computeSummaryForMonth(prevMonthKey);

  function pctChange(cur, prevVal) {
    if (!prevVal) return null;
    return (cur - prevVal) / prevVal;
  }

  var range = fiscalMonthRangeForKey(monthKey, closingDay);
  var allExpenseRows = getSheetData(SHEETS.TRANSACTIONS).rows.filter(function (r) {
    var d = new Date(r['日付']);
    return d.getTime() >= range.startDate.getTime() && d.getTime() <= range.endDate.getTime() && r['種別'] === '支出';
  });

  // カテゴリ別／誰が別／支払い方法別の内訳は、固定費に一致する取引を除外したうえで
  // 固定費側の登録金額を合算する（支出の合計とグラフの内訳の合計が一致するように）
  var fixedPairs = getFixedCostPairs();
  var variableRows = allExpenseRows.filter(function (r) {
    return !fixedPairs[r['カテゴリ'] + '|' + r['誰が']];
  });
  var fixedCosts = getSheetData(SHEETS.FIXED_COSTS).rows;

  function groupBy(txKey, fcKey) {
    var map = {};
    variableRows.forEach(function (r) {
      var k = r[txKey];
      map[k] = (map[k] || 0) + (Number(r['金額']) || 0);
    });
    fixedCosts.forEach(function (r) {
      var k = r[fcKey];
      map[k] = (map[k] || 0) + (Number(r['金額（月）']) || 0);
    });
    return Object.keys(map).map(function (k) { return { label: k, amount: map[k] }; })
      .filter(function (x) { return x.amount !== 0; })
      .sort(function (a, b) { return b.amount - a.amount; });
  }

  var dailyMap = {};
  // 日別の推移は特定の日付を持たない固定費を含められないため、実際に記録した取引のみで作る
  allExpenseRows.forEach(function (r) {
    var d = formatDate(r['日付']);
    dailyMap[d] = (dailyMap[d] || 0) + (Number(r['金額']) || 0);
  });
  var dailyTrend = [];
  var cursor = new Date(range.startDate);
  while (cursor.getTime() <= range.endDate.getTime()) {
    var key = formatDate(cursor);
    dailyTrend.push({ date: key, amount: dailyMap[key] || 0 });
    cursor.setDate(cursor.getDate() + 1);
  }

  return {
    monthKey: monthKey,
    prevMonthKey: prevMonthKey,
    income: current.income,
    expense: current.expense,
    balance: current.balance,
    incomeChangeRate: pctChange(current.income, prev.income),
    expenseChangeRate: pctChange(current.expense, prev.expense),
    balanceChangeRate: pctChange(current.balance, prev.balance),
    byCategory: groupBy('カテゴリ', '紐づけるカテゴリ'),
    byMember: groupBy('誰が', '誰の'),
    byPaymentMethod: groupBy('支払い方法', '支払い方法'),
    dailyTrend: dailyTrend
  };
}

// ===== カテゴリ別の前年比較（光熱費など季節変動のあるカテゴリ向け） =====

function sumCategoryForMonth(category, monthKey, closingDay) {
  var range = fiscalMonthRangeForKey(monthKey, closingDay);
  var total = 0;
  getSheetData(SHEETS.TRANSACTIONS).rows.forEach(function (r) {
    if (r['カテゴリ'] !== category || r['種別'] !== '支出') return;
    var d = new Date(r['日付']);
    if (d.getTime() >= range.startDate.getTime() && d.getTime() <= range.endDate.getTime()) {
      total += Number(r['金額']) || 0;
    }
  });
  return total;
}

function apiGetCategoryYearComparison(p) {
  var closingDay = getClosingDay();
  var baseKey = p.monthKey || getFiscalMonth(new Date(), closingDay).key;
  var baseYear = Number(baseKey.split('-')[0]);
  var months = [];
  for (var m = 1; m <= 12; m++) {
    months.push({
      month: m,
      current: sumCategoryForMonth(p.category, fmtYm(baseYear, m), closingDay),
      previous: sumCategoryForMonth(p.category, fmtYm(baseYear - 1, m), closingDay)
    });
  }
  return { category: p.category, year: baseYear, prevYear: baseYear - 1, months: months };
}
