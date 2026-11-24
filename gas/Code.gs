/**
 * わが家の献立 — 保存用 Google Apps Script
 *
 * 役割：GitHub Pages のアプリから届いた「操作」をスプレッドシートに保存し、最新の状態を返すだけ。
 *   - 自動実行（トリガー）はありません。アプリから呼ばれたときだけ動きます。
 *   - メール送信や外部への送信はしません。
 *   - 書き込みはすべて「log」シートに記録します（いつ・どの端末が・何をしたか）。
 *
 * 止め方：Apps Script の「デプロイを管理」で、このデプロイをアーカイブする（アプリからは読めも書けもしなくなる）。
 * 合言葉：スクリプトプロパティ KEY に、ご自身で決めた文字を入れてください（このファイルには書かない）。
 */

const DATA = 'data';
const LOG = 'log';
const DOCS = ['dishes', 'templates', 'plan', 'checks'];
const EMPTY = { dishes: null, templates: { weeks: [] }, plan: { days: {} }, checks: { lists: {} } };

/** 最初に1回だけ手で実行：シートを用意する（データがあれば触らない） */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DATA);
  if (!sh) {
    sh = ss.insertSheet(DATA);
    sh.getRange(1, 1, 1, 3).setValues([['名前', '中身(JSON)', '更新日時']]).setFontWeight('bold');
    sh.getRange(2, 1, DOCS.length, 3).setValues(DOCS.map(d => [d, JSON.stringify(EMPTY[d]), '']));
    sh.setFrozenRows(1);
  }
  let lg = ss.getSheetByName(LOG);
  if (!lg) {
    lg = ss.insertSheet(LOG);
    lg.getRange(1, 1, 1, 3).setValues([['日時', '端末', '操作']]).setFontWeight('bold');
    lg.setFrozenRows(1);
  }
  const hasKey = !!PropertiesService.getScriptProperties().getProperty('KEY');
  const msg = 'シートの準備ができました。' + (hasKey ? '' : '\n次に「プロジェクトの設定」→「スクリプト プロパティ」で KEY（合言葉）を追加してください。');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function doGet() {
  return out_({ ok: true, app: 'kondate', note: 'アプリからは POST で使います' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  const want = PropertiesService.getScriptProperties().getProperty('KEY');
  if (!want) return out_({ ok: false, error: 'not_configured' });
  if (body.key !== want) return out_({ ok: false, error: 'bad_key' });

  const ops = Array.isArray(body.ops) ? body.ops.slice(0, 200) : [];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const st = read_();
    if (ops.length) {
      ops.forEach(o => apply_(st, o));
      prune_(st);
      write_(st);
      log_(String(body.device || '不明').slice(0, 30), ops);
    }
    return out_({ ok: true, state: st, at: new Date().toISOString() });
  } finally {
    lock.releaseLock();
  }
}

/** 操作を1つ反映（アプリ側の applyOp と同じ決まり） */
function apply_(st, o) {
  const days = st.plan.days, lists = st.checks.lists;
  switch (o && o.t) {
    case 'day': if (o.v) days[o.k] = o.v; else delete days[o.k]; break;
    case 'days': Object.keys(o.m || {}).forEach(k => { if (o.m[k]) days[k] = o.m[k]; else delete days[k]; }); break;
    case 'chk': { const l = lists[o.l] || (lists[o.l] = {}); if (o.v) l[o.n] = true; else delete l[o.n]; break; }
    case 'chkclr': delete lists[o.l]; break;
    case 'dishes': if (o.v && Array.isArray(o.v.items)) st.dishes = o.v; break;
    case 'init':
      if (!st.dishes || !st.dishes.items || !st.dishes.items.length) st.dishes = o.dishes;
      if (!st.templates.weeks || !st.templates.weeks.length) st.templates = o.templates;
      break;
  }
}

/** 古いデータの整理：チェックは直近12リスト、献立は400日より前を削除 */
function prune_(st) {
  const keys = Object.keys(st.checks.lists).sort();
  while (keys.length > 12) delete st.checks.lists[keys.shift()];
  const limit = Utilities.formatDate(new Date(Date.now() - 400 * 864e5), 'Asia/Tokyo', 'yyyy-MM-dd');
  Object.keys(st.plan.days).forEach(k => { if (k < limit) delete st.plan.days[k]; });
}

function read_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(DATA);
  if (!sh) throw new Error('setup を先に実行してください');
  const st = JSON.parse(JSON.stringify(EMPTY));
  sh.getRange(2, 1, DOCS.length, 2).getValues().forEach(([name, json]) => {
    if (DOCS.indexOf(name) >= 0 && json) { try { st[name] = JSON.parse(json); } catch (e) {} }
  });
  if (!st.plan || !st.plan.days) st.plan = { days: {} };
  if (!st.checks || !st.checks.lists) st.checks = { lists: {} };
  if (!st.templates || !st.templates.weeks) st.templates = { weeks: [] };
  return st;
}

function write_(st) {
  const sh = SpreadsheetApp.getActive().getSheetByName(DATA);
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  sh.getRange(2, 1, DOCS.length, 3).setValues(DOCS.map(d => [d, JSON.stringify(st[d]), now]));
}

function log_(device, ops) {
  const lg = SpreadsheetApp.getActive().getSheetByName(LOG);
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const label = { day: '献立', days: '献立(週)', chk: 'チェック', chkclr: 'チェック全解除', dishes: '料理リスト', init: '初期データ' };
  const text = ops.map(o => {
    if (o.t === 'day') return '献立 ' + o.k;
    if (o.t === 'chk') return (o.v ? '✓ ' : '□ ') + o.n;
    return label[o.t] || String(o.t);
  }).join(' / ').slice(0, 500);
  lg.appendRow([now, device, text]);
  if (lg.getLastRow() > 3000) lg.deleteRows(2, 1000);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
