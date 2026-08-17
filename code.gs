// データ用のスプシIDを定義
const SPREADSHEET_ID = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';

// シート名を定数化（マジックストリング排除）
const SHEET = { PRODUCTS: 'products', CONFIG: 'config', CASES: 'cases' };

// config のキー（英字識別子に統一）
const CONFIG_KEY = {
  TITLE: 'title',
  INTERVENTION_AMOUNT: 'intervention_amount',
  INTERVENTION_CASE: 'intervention_case',
  INTERVENTION_LABEL: 'intervention_label',
};

// キャッシュ有効秒数（0でキャッシュ無効）
const CACHE_TTL_SEC = 600;
const CACHE_KEY = 'appData';

// タイトル未設定時のフォールバック
const DEFAULT_TITLE = '闇医者お会計ツール';

// 公開ページのホワイトリスト（キー=?page値、値=HTMLファイル名）
const PAGES = {
  index:  { file: 'index',  title: DEFAULT_TITLE },
  readme: { file: 'readme', title: 'README｜' + DEFAULT_TITLE },
};

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// キャッシュを手動クリア（GASエディタから実行。データ更新の即時反映用）
function clearCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'index';
  const target = PAGES[page] || PAGES.index;

  // indexページのみ config のタイトルをタブ名へ反映（readmeは固定）
  let title = target.title;
  if (page === 'index') title = getConfigTitle_();

  return HtmlService.createHtmlOutputFromFile(target.file)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// config の title だけを軽量に取得（doGet用。無ければデフォルト）
function getConfigTitle_() {
  try {
    const sheet = getSheet_(SpreadsheetApp.openById(SPREADSHEET_ID), SHEET.CONFIG);
    const map = readKeyValue_(sheet);
    const v = map[CONFIG_KEY.TITLE];
    if (v != null && String(v).trim() !== '') return String(v).trim();
  } catch (err) {}
  return DEFAULT_TITLE;
}

function getData() {
  const cache = CacheService.getScriptCache();

  // キャッシュヒット時はスプレッドシートアクセスを回避
  if (CACHE_TTL_SEC > 0) {
    const cached = cache.get(CACHE_KEY);
    if (cached) return JSON.parse(cached);
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cases = readCases(getSheet_(ss, SHEET.CASES));
  const config = readConfig_(getSheet_(ss, SHEET.CONFIG));
  const products = readProducts(getSheet_(ss, SHEET.PRODUCTS), cases);

  // 有効な区分idの集合（intervention_case の妥当性検証に使う）
  const caseIds = cases.map(function (c) { return c.id; });

  // 中型介入設定の組み立て。amount と case が有効なときのみ enabled=true
  let intervention = { enabled: false, amount: 0, caseId: '', label: '' };
  const amt = config.numbers[CONFIG_KEY.INTERVENTION_AMOUNT];
  const caseId = config.strings[CONFIG_KEY.INTERVENTION_CASE];
  const ivLabel = config.strings[CONFIG_KEY.INTERVENTION_LABEL] || '中型介入';
  if (amt != null && isFinite(amt) && caseId && caseIds.indexOf(caseId) !== -1) {
    intervention = { enabled: true, amount: amt, caseId: caseId, label: ivLabel };
  }

  const data = {
    title: config.title,
    cases: cases,
    products: products,
    intervention: intervention,
  };

  if (CACHE_TTL_SEC > 0) {
    cache.put(CACHE_KEY, JSON.stringify(data), CACHE_TTL_SEC);
  }
  return data;
}

/* ===== 共通ヘルパ ===== */

// シート取得（存在しなければ明示的にエラー）
function getSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('シートが見つかりません: ' + name);
  return sheet;
}

// ヘッダー行から列名→インデックスのマップを作る（列順への依存を排除）
function headerIndex_(values) {
  const header = values[0] || [];
  const map = {};
  header.forEach(function (h, i) { map[String(h).trim()] = i; });
  return map;
}

// セルが空（未設定）かどうか。Number('')が0になる問題を避けるため生値で判定
function isBlank_(v) {
  return v === '' || v === null || v === undefined;
}

// key/value 2列シートを素直なオブジェクトにする（生値のまま）
function readKeyValue_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};
  const col = headerIndex_(values);
  const iKey = col.key != null ? col.key : 0;
  const iVal = col.value != null ? col.value : 1;
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][iKey]).trim();
    if (key === '') continue;
    map[key] = values[i][iVal];
  }
  return map;
}

/* ===== 各シート読み取り ===== */

// cases: id, label, rate, order を読み取り order 昇順で返す
function readCases(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const col = headerIndex_(values);
  const iId = col.id != null ? col.id : 0;
  const iLabel = col.label != null ? col.label : 1;
  const iRate = col.rate != null ? col.rate : 2;
  const iOrder = col.order != null ? col.order : 3;

  const cases = [];
  for (let i = 1; i < values.length; i++) {
    const id = String(values[i][iId]).trim();
    const label = String(values[i][iLabel]).trim();
    if (id === '' || label === '') continue;   // id/label 必須

    const rate = Number(values[i][iRate]);
    const order = Number(values[i][iOrder]);
    cases.push({
      id: id,
      label: label,
      rate: isFinite(rate) ? rate : 0,
      order: isFinite(order) ? order : 9999,
    });
  }
  cases.sort(function (a, b) { return a.order - b.order; });
  return cases;
}

// config: title（文字列）、および数値設定・文字列設定を仕分けて返す
function readConfig_(sheet) {
  const raw = readKeyValue_(sheet);

  let title = DEFAULT_TITLE;
  const t = raw[CONFIG_KEY.TITLE];
  if (t != null && String(t).trim() !== '') title = String(t).trim();

  // 数値として扱うキーと文字列として扱うキーを分離
  const numbers = {};
  const strings = {};
  Object.keys(raw).forEach(function (key) {
    if (key === CONFIG_KEY.TITLE) return;
    const v = raw[key];
    if (isBlank_(v)) return;                 // 空欄は未設定として無視
    const n = Number(v);
    if (isFinite(n) && String(v).trim() !== '') {
      numbers[key] = n;                      // 数値化できるものは numbers へ
    }
    strings[key] = String(v).trim();         // 文字列版も保持（idなど参照用）
  });

  return { title: title, numbers: numbers, strings: strings };
}

// products: name, price, view（区分idのカンマ区切り）を読み取り
function readProducts(sheet, cases) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const col = headerIndex_(values);
  const iName = col.name != null ? col.name : 0;
  const iPrice = col.price != null ? col.price : 1;
  const iView = col.view != null ? col.view : 2;

  const validIds = {};
  cases.forEach(function (c) { validIds[c.id] = true; });

  const products = [];
  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][iName]).trim();
    const rawPrice = values[i][iPrice];
    if (name === '' || isBlank_(rawPrice)) continue;
    const price = Number(rawPrice);
    if (!isFinite(price)) continue;

    // view：カンマ区切りの区分id。存在する id のみ採用（不正idは除外）
    const rawView = String(values[i][iView] || '');
    const views = rawView.split(',')
      .map(function (s) { return s.trim(); })
      .filter(function (s) { return s !== '' && validIds[s]; });

    products.push({ name: name, price: price, views: views });
  }
  return products;
}