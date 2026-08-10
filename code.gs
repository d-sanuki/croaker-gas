// データ用のスプシIDを定義
const SPREADSHEET_ID = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';

// シート名を定数化（マジックストリング排除）
const SHEET = { PRODUCTS: 'products', CONFIG: 'config' };

// キャッシュ有効秒数（データ更新頻度に応じて調整。0でキャッシュ無効）
const CACHE_TTL_SEC = 300;

// 公開ページのホワイトリスト（キー=?page値、値=HTMLファイル名）
const PAGES = {
  index:  { file: 'index',  title: '闇医者お会計ツール' },
  readme: { file: 'readme', title: 'README｜闇医者お会計ツール' },
};

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'index';
  const target = PAGES[page] || PAGES.index;

  return HtmlService.createHtmlOutputFromFile(target.file)
    .setTitle(target.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getData() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'appData';

  // キャッシュヒット時はスプレッドシートアクセスを回避
  if (CACHE_TTL_SEC > 0) {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const data = {
    products: readProducts(getSheet_(ss, SHEET.PRODUCTS)),
    rates: readRates(getSheet_(ss, SHEET.CONFIG)),
  };

  if (CACHE_TTL_SEC > 0) {
    cache.put(cacheKey, JSON.stringify(data), CACHE_TTL_SEC);
  }
  return data;
}

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

// products: name, price 列を読み取り
function readProducts(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const col = headerIndex_(values);
  const iName = col.name != null ? col.name : 0;   // ヘッダーが無ければ従来の列順にフォールバック
  const iPrice = col.price != null ? col.price : 1;

  const products = [];
  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][iName]).trim();
    const price = Number(values[i][iPrice]);
    // 名前が空、または価格が数値でない行はスキップ
    if (name === '' || !isFinite(price)) continue;
    products.push({ name: name, price: price });
  }
  return products;
}

// config: key, value 列を読み取り rates マップ化
function readRates(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};

  const col = headerIndex_(values);
  const iKey = col.key != null ? col.key : 0;
  const iVal = col.value != null ? col.value : 1;

  const rates = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][iKey]).trim();
    const val = Number(values[i][iVal]);
    if (key === '' || !isFinite(val)) continue;
    rates[key] = val;
  }
  return rates;
}
