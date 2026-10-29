/**
 * マニュアルポータル（イトー寿薬局・やしお薬局）データ保存用 Apps Script
 *
 * 使い方（くわしくは設定手順を参照）
 *   1. Googleスプレッドシートを作り、［拡張機能］→［Apps Script］を開く
 *   2. このファイルの中身をすべて貼り付けて保存
 *   3. 関数「setup」を選んで［実行］（シートと初期データができる）
 *   4. ［プロジェクトの設定］→［スクリプト プロパティ］に PASSCODE を追加
 *   5. ［デプロイ］→［新しいデプロイ］→ 種類「ウェブアプリ」
 *        次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 *
 * 注意：患者さんの個人情報・処方情報は保存しない運用です。
 */

// シート名と列（1行目の見出し）
const SHEETS = {
  categories: ['id', 'name', 'icon', 'order'],
  manuals: ['id', 'categoryId', 'title', 'icon', 'body', 'updatedAt', 'updatedBy', 'order', 'deleted'],
  forms: ['id', 'name', 'icon', 'url', 'order', 'deleted'],
};

// 暗証番号を続けて間違えたときのロック（10分間に20回で、10分間ロック）
const MAX_FAILS = 20;
const LOCK_SECONDS = 600;

// ───────── 初期設定 ─────────

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const seed = seedData_();
  Object.keys(SHEETS).forEach(function (name) {
    const header = SHEETS[name];
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    // 「=」「-」で始まる文字が数式にならないよう、すべて書式なしテキストにする
    sh.getRange(1, 1, sh.getMaxRows(), header.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (sh.getLastRow() <= 1 && seed[name].length) {
      const rows = seed[name].map(function (o) {
        return header.map(function (h) { return o[h] === undefined ? '' : String(o[h]); });
      });
      sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    }
  });
  // 最初からある空のシートを片づける
  ['シート1', 'Sheet1'].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  const hasPass = !!PropertiesService.getScriptProperties().getProperty('PASSCODE');
  Logger.log('setup 完了。' + (hasPass ? 'PASSCODE は設定済みです。' : '次に スクリプト プロパティ に PASSCODE を追加してください。'));
}

function seedData_() {
  const now = new Date().toISOString();
  const categories = [
    { id: 'training', name: '新人・研修用マニュアル', icon: '📚' },
    { id: 'sop', name: '業務マニュアル・SOP', icon: '💊' },
    { id: 'trouble', name: 'トラブル対応・メンテナンス業務', icon: '🔧' },
    { id: 'rules', name: 'その他・共通ルール', icon: '📌' },
    { id: 'handout', name: '患者配布物', icon: '🏥' },
    { id: 'storage', name: 'データ保管', icon: '🗄️' },
  ].map(function (c, i) { c.order = i + 1; return c; });

  const m = function (id, categoryId, title, icon, body) {
    return { id: id, categoryId: categoryId, title: title, icon: icon, body: body || '',
             updatedAt: body ? now : '', updatedBy: '', deleted: '' };
  };
  const manuals = [
    m('orientation', 'training', '入社時オリエンテーション', '🏢'),
    m('basic-flow', 'training', '基本業務の流れ', '📝'),
    m('training-check', 'training', '研修チェックリスト', '✅'),
    m('dispensing', 'sop', '調剤業務手順', '⚗️'),
    m('counseling', 'sop', '服薬指導の手順', '🗣️'),
    m('stock', 'sop', '在庫管理・発注', '📦'),
    m('receipt', 'sop', 'レセプト・会計業務', '🧾'),
    m('rececom', 'sop', 'レセコン操作マニュアル', '🖥️'),
    m('e-yakureki', 'sop', '電子薬歴の使い方', '📱'),
    m('machines', 'sop', '分包機など機器操作', '⚙️'),
    m('proofit', 'trouble', 'proofitラベル交換', '🏷️',
      '動画 https://drive.google.com/file/d/1_BzRRRkOLsyQGISYfpdOupdDl_29NjRo/view\nラベル交換の手順を動画で確認できます。'),
    m('machine-error', 'trouble', '機器エラー復旧手順', '🖥️'),
    m('gooco', 'trouble', 'GooCo薬歴飛ばない', '📡',
      '動画 https://drive.google.com/file/d/1N8SUtZRz3Hzyy17KClV6qGYpOz5zhL8x/view\n薬歴データが送信できない場合の復旧手順を動画で確認できます。'),
    m('new-patient-master', 'trouble', '新患マスタ紐づけトラブルの対処法', '🔗'),
    m('daily', 'trouble', '日次メンテナンス業務', '🧹'),
    m('monthly', 'trouble', '月次メンテナンス業務', '📅'),
    m('emergency', 'rules', '緊急時対応', '🚨'),
    m('company-rules', 'rules', '社内ルール・規定', '📖'),
    m('okusuri-techo', 'handout', 'お薬手帳の使い方', '📒'),
    m('how-to-take', 'handout', '服用方法の説明書', '💊'),
    m('generic', 'handout', 'ジェネリック医薬品のご案内', '📄'),
    m('templates', 'storage', '各種テンプレート', '📝'),
    m('notifications', 'storage', '届出書類', '📋'),
    m('training-records', 'storage', '研修記録', '✅'),
  ].map(function (x, i) { x.order = i + 1; return x; });

  const forms = [
    { id: 'overtime', name: '残業申請', icon: '⏰', url: 'https://docs.google.com/forms/d/e/1FAIpQLSeP89kxMD33sGVL7WAAz2wAvmM5e2jbKrmSIxqwoTjDpHJ-0w/viewform' },
    { id: 'purchase', name: '購入申請', icon: '🛒', url: 'https://docs.google.com/forms/d/e/1FAIpQLSe8nWOpjV93YNCDRhkovB52JveH6CbWrGR1c8J0u3NTpyPJhg/viewform' },
    { id: 'request', name: '会社への要望・相談（ハラスメント相談含む）', icon: '💬', url: 'https://docs.google.com/forms/d/e/1FAIpQLSdLEVYpBuGnNpZ2uiU6It6xyNGO0TV_h70YQScYGV5nADSWxg/viewform' },
  ].map(function (f, i) { f.order = i + 1; f.deleted = ''; return f; });

  return { categories: categories, manuals: manuals, forms: forms };
}

// ───────── ウェブアプリの入口 ─────────

function doGet() {
  return respond_(function () { return readAll_(); });
}

function doPost(e) {
  return respond_(function () {
    let req;
    try { req = JSON.parse(e.postData.contents); } catch (err) { throw apiError_('badrequest'); }
    checkPasscode_(req.passcode);
    const data = req.data || {};
    const name = clip_(req.name, 20);
    switch (req.action) {
      case 'verify': return {};
      case 'saveManual': return withLock_(function () { return { saved: saveManual_(data, name) }; });
      case 'deleteManual': return withLock_(function () { markDeleted_('manuals', data.id); return {}; });
      case 'saveForm': return withLock_(function () { return { saved: saveForm_(data) }; });
      case 'deleteForm': return withLock_(function () { markDeleted_('forms', data.id); return {}; });
      default: throw apiError_('badrequest');
    }
  });
}

function respond_(fn) {
  let out;
  try {
    out = fn();
    out.ok = true;
  } catch (err) {
    out = { ok: false, error: err.code || 'server', message: err.code ? '' : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function apiError_(code) {
  const err = new Error(code);
  err.code = code;
  return err;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw apiError_('busy');
  try {
    const result = fn();
    SpreadsheetApp.flush();
    return Object.assign(result, readAll_());
  } finally {
    lock.releaseLock();
  }
}

function checkPasscode_(input) {
  const pass = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  if (!pass) throw apiError_('nopass');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) throw apiError_('locked');
  if (String(input || '') !== pass) {
    cache.put('fails', String(fails + 1), LOCK_SECONDS);
    Utilities.sleep(1000);
    throw apiError_('badpass');
  }
}

// ───────── 読み書き ─────────

function readAll_() {
  const cats = readSheet_('categories').rows;
  const manuals = readSheet_('manuals').rows.filter(function (r) { return !r.deleted; });
  const forms = readSheet_('forms').rows.filter(function (r) { return !r.deleted; });
  [cats, manuals, forms].forEach(function (list) {
    list.forEach(function (r) { r.order = Number(r.order) || 0; delete r._row; delete r.deleted; });
    list.sort(function (a, b) { return a.order - b.order; });
  });
  return { categories: cats, manuals: manuals, forms: forms, serverTime: new Date().toISOString() };
}

function readSheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw apiError_('nosetup');
  const header = SHEETS[name];
  const last = sh.getLastRow();
  const values = last > 1 ? sh.getRange(2, 1, last - 1, header.length).getDisplayValues() : [];
  const rows = [];
  values.forEach(function (v, i) {
    if (!v[0]) return;
    const o = { _row: i + 2 };
    header.forEach(function (h, j) { o[h] = v[j]; });
    rows.push(o);
  });
  return { sh: sh, header: header, rows: rows };
}

function writeRow_(table, obj) {
  const rowIndex = obj._row || table.sh.getLastRow() + 1;
  if (rowIndex > table.sh.getMaxRows()) table.sh.insertRowsAfter(table.sh.getMaxRows(), 100);
  const range = table.sh.getRange(rowIndex, 1, 1, table.header.length);
  range.setNumberFormat('@');
  range.setValues([table.header.map(function (h) { return obj[h] === undefined ? '' : String(obj[h]); })]);
}

function nextOrder_(rows) {
  return rows.reduce(function (max, r) { return Math.max(max, Number(r.order) || 0); }, 0) + 1;
}

function newId_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 10);
}

function clip_(value, max) {
  return String(value === undefined || value === null ? '' : value).trim().slice(0, max);
}

function saveManual_(d, name) {
  const table = readSheet_('manuals');
  const title = clip_(d.title, 100);
  if (!title) throw apiError_('notitle');
  const body = String(d.body || '');
  if (body.length > 30000) throw apiError_('toolong');
  const categoryId = clip_(d.categoryId, 40);
  const catOk = readSheet_('categories').rows.some(function (c) { return c.id === categoryId; });
  if (!catOk) throw apiError_('nocategory');

  let row;
  if (d.id) {
    row = table.rows.filter(function (r) { return r.id === d.id && !r.deleted; })[0];
    if (!row) throw apiError_('notfound');
    // 編集を始めたあとに、ほかの人が先に保存していたら止める
    if ((row.updatedAt || '') !== (d.baseUpdatedAt || '')) throw apiError_('conflict');
  } else {
    row = { id: newId_(), order: nextOrder_(table.rows), deleted: '' };
  }
  row.title = title;
  row.icon = clip_(d.icon, 8) || '📄';
  row.categoryId = categoryId;
  row.body = body;
  row.updatedAt = new Date().toISOString();
  row.updatedBy = name;
  writeRow_(table, row);
  return row.id;
}

function saveForm_(d) {
  const table = readSheet_('forms');
  const name = clip_(d.name, 60);
  if (!name) throw apiError_('notitle');
  const url = clip_(d.url, 500);
  if (!/^https:\/\/\S+$/.test(url)) throw apiError_('badurl');
  let row;
  if (d.id) {
    row = table.rows.filter(function (r) { return r.id === d.id && !r.deleted; })[0];
    if (!row) throw apiError_('notfound');
  } else {
    row = { id: newId_(), order: nextOrder_(table.rows), deleted: '' };
  }
  row.name = name;
  row.url = url;
  row.icon = clip_(d.icon, 8) || '📝';
  writeRow_(table, row);
  return row.id;
}

function markDeleted_(name, id) {
  const table = readSheet_(name);
  const row = table.rows.filter(function (r) { return r.id === id; })[0];
  if (!row) throw apiError_('notfound');
  row.deleted = new Date().toISOString();
  writeRow_(table, row);
}
