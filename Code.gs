/**
 * 買い物メモアプリ - バックエンド (Google Apps Script)
 *
 * 【セットアップ手順】
 * 1. Google スプレッドシートを新規作成する
 * 2. メニュー「拡張機能」→「Apps Script」を開く
 * 3. デフォルトの Code.gs の中身を全部消して、このファイルの内容を貼り付ける
 * 4. 上の「実行」ボタンを一度押して setup() を実行し、シートの初期化をする
 *    (権限の承認を求められたら許可する)
 * 5. 右上の「デプロイ」→「新しいデプロイ」
 *    種類: ウェブアプリ
 *    実行するユーザー: 自分
 *    アクセスできるユーザー: 全員
 *    →デプロイして発行される URL (.../exec) をコピー
 * 6. アプリ(index.html)の設定画面にその URL を貼り付ける
 *
 * シート構成:
 *  - Categories: id, name, color, order
 *  - Items:      id, name, categoryId, status, createdAt, updatedAt
 *  - Templates:  id, name, categoryId, order   ※「よく使うもの」の登録リスト
 */

const CATEGORY_SHEET = 'Categories';
const ITEM_SHEET = 'Items';
const TEMPLATE_SHEET = 'Templates';

const DEFAULT_CATEGORIES = [
  { name: 'スーパー', color: '#4C9A8B' },
  { name: '薬局', color: '#8B5FBF' },
];

// 初期の「よく使うもの」サンプル (カテゴリのインデックスで指定: 0=スーパー, 1=薬局)
const DEFAULT_TEMPLATES = [
  { name: 'たまご', catIndex: 0 },
  { name: '牛乳', catIndex: 0 },
  { name: 'トイレットペーパー', catIndex: 0 },
  { name: '常備薬', catIndex: 1 },
];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let catSheet = ss.getSheetByName(CATEGORY_SHEET);
  if (!catSheet) catSheet = ss.insertSheet(CATEGORY_SHEET);
  catSheet.clear();
  catSheet.appendRow(['id', 'name', 'color', 'order']);

  let itemSheet = ss.getSheetByName(ITEM_SHEET);
  if (!itemSheet) itemSheet = ss.insertSheet(ITEM_SHEET);
  itemSheet.clear();
  itemSheet.appendRow(['id', 'name', 'categoryId', 'status', 'createdAt', 'updatedAt']);

  const catIds = [];
  DEFAULT_CATEGORIES.forEach((c, i) => {
    const id = Utilities.getUuid();
    catIds.push(id);
    catSheet.appendRow([id, c.name, c.color, i]);
  });

  let tplSheet = ss.getSheetByName(TEMPLATE_SHEET);
  if (!tplSheet) tplSheet = ss.insertSheet(TEMPLATE_SHEET);
  tplSheet.clear();
  tplSheet.appendRow(['id', 'name', 'categoryId', 'order']);
  DEFAULT_TEMPLATES.forEach((t, i) => {
    tplSheet.appendRow([Utilities.getUuid(), t.name, catIds[t.catIndex], i]);
  });

  // デフォルトの「シート1」が残っていれば削除
  const sheet1 = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (sheet1) ss.deleteSheet(sheet1);

  Logger.log('セットアップ完了。デプロイして URL をアプリに設定してください。');
}

/**
 * 既に運用中のスプレッドシートに「よく使うもの」機能を追加するための関数。
 * setup() を再実行すると既存のカテゴリ・アイテムが消えてしまうため、
 * 追加済みの環境では setup() の代わりにこちらを1回実行してください。
 * (関数選択プルダウンで migrateAddTemplates を選んで ▶ 実行)
 */
function migrateAddTemplates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tplSheet = ss.getSheetByName(TEMPLATE_SHEET);
  if (tplSheet) {
    Logger.log('Templates シートは既に存在します。何もしませんでした。');
    return;
  }
  tplSheet = ss.insertSheet(TEMPLATE_SHEET);
  tplSheet.appendRow(['id', 'name', 'categoryId', 'order']);

  const categories = getCategories();
  DEFAULT_TEMPLATES.forEach((t, i) => {
    const cat = categories[t.catIndex];
    if (!cat) return; // 対応するカテゴリが無ければスキップ
    tplSheet.appendRow([Utilities.getUuid(), t.name, cat.id, i]);
  });
  Logger.log('Templates シートを追加しました。この後「デプロイを管理」から新バージョンとして再デプロイしてください。');
}

function doGet(e) {
  const data = {
    categories: getCategories(),
    items: getItems(),
    templates: getTemplates(),
  };
  return jsonResponse(data);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'invalid json' }, 400);
  }

  const action = body.action;
  let result;
  try {
    switch (action) {
      case 'addItem':
        result = addItem(body.name, body.categoryId);
        break;
      case 'checkItem':
        result = checkItem(body.id);
        break;
      case 'deleteItem':
        result = deleteItem(body.id);
        break;
      case 'moveItem':
        result = moveItem(body.id, body.categoryId);
        break;
      case 'addCategory':
        result = addCategory(body.name, body.color);
        break;
      case 'updateCategory':
        result = updateCategory(body.id, body.name, body.color);
        break;
      case 'deleteCategory':
        result = deleteCategory(body.id);
        break;
      case 'addTemplate':
        result = addTemplate(body.name, body.categoryId);
        break;
      case 'deleteTemplate':
        result = deleteTemplate(body.id);
        break;
      default:
        return jsonResponse({ error: 'unknown action' }, 400);
    }
  } catch (err) {
    return jsonResponse({ error: String(err) }, 500);
  }

  return jsonResponse({
    ok: true,
    result: result,
    categories: getCategories(),
    items: getItems(),
    templates: getTemplates(),
  });
}

/* ---------- ヘルパー ---------- */

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .filter(r => r[0] !== '')
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => (obj[h] = r[i]));
      return obj;
    });
}

function findRowIndexById(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) return i + 1; // 1-indexed row number
  }
  return -1;
}

/* ---------- カテゴリ ---------- */

function getCategories() {
  return sheetToObjects(getSheet(CATEGORY_SHEET)).sort((a, b) => a.order - b.order);
}

function addCategory(name, color) {
  const sheet = getSheet(CATEGORY_SHEET);
  const id = Utilities.getUuid();
  const order = sheet.getLastRow(); // 現在の行数をそのままorderに使う
  sheet.appendRow([id, name, color || '#4C9A8B', order]);
  return { id: id };
}

function updateCategory(id, name, color) {
  const sheet = getSheet(CATEGORY_SHEET);
  const row = findRowIndexById(sheet, id);
  if (row === -1) throw new Error('category not found');
  if (name) sheet.getRange(row, 2).setValue(name);
  if (color) sheet.getRange(row, 3).setValue(color);
  return { id: id };
}

function deleteCategory(id) {
  const itemSheet = getSheet(ITEM_SHEET);
  const items = sheetToObjects(itemSheet);
  const hasItems = items.some(it => it.categoryId === id && it.status === 'active');
  if (hasItems) throw new Error('このカテゴリにはアイテムが残っています');

  const sheet = getSheet(CATEGORY_SHEET);
  const row = findRowIndexById(sheet, id);
  if (row === -1) throw new Error('category not found');
  sheet.deleteRow(row);

  // このカテゴリに紐づく「よく使うもの」も連鎖削除
  const tplSheet = getSheet(TEMPLATE_SHEET);
  const templates = sheetToObjects(tplSheet);
  templates.filter(t => t.categoryId === id).forEach(t => {
    const r = findRowIndexById(tplSheet, t.id);
    if (r !== -1) tplSheet.deleteRow(r);
  });

  return { id: id };
}

/* ---------- よく使うもの (テンプレート) ---------- */

function getTemplates() {
  return sheetToObjects(getSheet(TEMPLATE_SHEET)).sort((a, b) => a.order - b.order);
}

function addTemplate(name, categoryId) {
  if (!name || !categoryId) throw new Error('name and categoryId required');
  const sheet = getSheet(TEMPLATE_SHEET);
  const id = Utilities.getUuid();
  const order = sheet.getLastRow();
  sheet.appendRow([id, name, categoryId, order]);
  return { id: id };
}

function deleteTemplate(id) {
  const sheet = getSheet(TEMPLATE_SHEET);
  const row = findRowIndexById(sheet, id);
  if (row === -1) throw new Error('template not found');
  sheet.deleteRow(row);
  return { id: id };
}

/* ---------- アイテム ---------- */

function getItems() {
  return sheetToObjects(getSheet(ITEM_SHEET)).filter(it => it.status === 'active');
}

function addItem(name, categoryId) {
  const sheet = getSheet(ITEM_SHEET);
  const id = Utilities.getUuid();
  const now = new Date().toISOString();
  sheet.appendRow([id, name, categoryId, 'active', now, now]);
  return { id: id };
}

function checkItem(id) {
  const sheet = getSheet(ITEM_SHEET);
  const row = findRowIndexById(sheet, id);
  if (row === -1) throw new Error('item not found');
  sheet.getRange(row, 4).setValue('done');
  sheet.getRange(row, 6).setValue(new Date().toISOString());
  return { id: id };
}

function deleteItem(id) {
  const sheet = getSheet(ITEM_SHEET);
  const row = findRowIndexById(sheet, id);
  if (row === -1) throw new Error('item not found');
  sheet.deleteRow(row);
  return { id: id };
}

function moveItem(id, categoryId) {
  const sheet = getSheet(ITEM_SHEET);
  const row = findRowIndexById(sheet, id);
  if (row === -1) throw new Error('item not found');
  sheet.getRange(row, 3).setValue(categoryId);
  sheet.getRange(row, 6).setValue(new Date().toISOString());
  return { id: id };
}
