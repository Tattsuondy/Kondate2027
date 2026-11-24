/**
 * わが家の献立 — 保存用 Google Apps Script（招待QR方式）
 *
 * 役割：GitHub Pages のアプリから届いた「操作」をスプレッドシートに保存し、最新の状態を返すだけ。
 *   - 合言葉はありません。端末ごとに発行する「鍵」で本人確認します。
 *     鍵はスプレッドシートにはハッシュ（元に戻せない形）でだけ残ります。
 *   - 新しい端末は「招待」でだけ登録できます。招待は10分間・1回限り有効です。
 *   - 自動実行（トリガー）・メール送信・外部への送信はありません。アプリから呼ばれたときだけ動きます。
 *   - 書き込み・招待・端末登録は、すべて「log」シートに記録します。
 *
 * 止め方：
 *   - 1台だけ止める … 「devices」シートのその行の「状態」を「停止」にする（または行を削除）
 *   - 全部止める   … Apps Script の「デプロイを管理」でこのデプロイをアーカイブする
 *
 * 最初の1台の登録：エディタで関数 createInvite を選んで実行 → スプレッドシートにQRが出る → スマホで読む
 */

const DATA = 'data', LOG = 'log', DEV = 'devices';
const DOCS = ['dishes', 'templates', 'plan', 'checks'];
const EMPTY = { dishes: null, templates: { weeks: [] }, plan: { days: {} }, checks: { lists: {} } };
const INVITE_MIN = 10;
const CODE_CHARS = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'; // 読み違えやすい I, L, O, 0, 1 は使わない

/** 最初に1回だけ手で実行：シートを用意する（データがあれば触らない） */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, DATA, ['名前', '中身(JSON)', '更新日時'], sh =>
    sh.getRange(2, 1, DOCS.length, 3).setValues(DOCS.map(d => [d, JSON.stringify(EMPTY[d]), ''])));
  ensureSheet_(ss, LOG, ['日時', '端末', '操作']);
  ensureSheet_(ss, DEV, ['端末ID', '名前', '鍵(ハッシュ)', '状態', '登録日時', '招待した端末', '最終利用']);
  alert_('シートの準備ができました。\n次に、デプロイ（ウェブアプリ）をしてから createInvite を実行してください。');
}

/** 招待QRを作る（エディタから実行。スプレッドシートを開いておくと、そこにQRが出ます） */
function createInvite() {
  const props = PropertiesService.getScriptProperties();
  let appUrl = props.getProperty('APP_URL');
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  if (!appUrl && ui) {
    const r = ui.prompt('アプリのURL', 'GitHub Pages のURLを入れてください（例：https://ユーザー名.github.io/kondate/）', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    appUrl = r.getResponseText().trim();
    if (!/^https:\/\/[^\s#]+$/.test(appUrl)) { alert_('URLの形が違います'); return; }
    props.setProperty('APP_URL', appUrl);
  }
  let name = 'パパのスマホ';
  if (ui) {
    const r = ui.prompt('どの端末用の招待ですか？', '例：パパのスマホ、ママのiPhone', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    name = (r.getResponseText().trim() || name).slice(0, 30);
  }
  const code = newInvite_(name, 'スプレッドシート');
  log_('スプレッドシート', '招待を作成（' + name + '）');
  const link = (appUrl || '') + '#inv=' + code;
  Logger.log('招待リンク（' + INVITE_MIN + '分有効・1回限り）：' + link);
  if (!ui) return;
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;text-align:center">' +
    '<div id="q" style="display:inline-block;padding:10px;background:#fff"></div>' +
    '<p style="font-size:22px;letter-spacing:3px;margin:8px 0"><b>' + fmt_(code) + '</b></p>' +
    '<p style="font-size:13px;color:#555">' + esc_(name) + ' 用。スマホのカメラで読んでください。<br>' + INVITE_MIN + '分間・1回だけ有効です。</p></div>' +
    '<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>' +
    '<script>new QRCode(document.getElementById("q"),{text:' + JSON.stringify(link) + ',width:220,height:220});</script>'
  ).setWidth(320).setHeight(380);
  ui.showModalDialog(html, '招待QR');
}

function doGet() {
  return out_({ ok: true, app: 'kondate' });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (b.a === 'claim') return out_(claim_(b));
    const dev = auth_(b.key);
    if (!dev) return out_({ ok: false, error: 'bad_device' });
    if (b.a === 'invite') {
      const name = String(b.name || '家族の端末').slice(0, 30);
      const code = newInvite_(name, dev.name);
      log_(dev.name, '招待を作成（' + name + '）');
      return out_({ ok: true, code: code, minutes: INVITE_MIN });
    }
    // 同期（a: 'sync'）
    const ops = Array.isArray(b.ops) ? b.ops.slice(0, 200) : [];
    const st = read_();
    if (ops.length) {
      ops.forEach(o => apply_(st, o));
      prune_(st);
      write_(st);
      log_(dev.name, describe_(ops));
    }
    touch_(dev);
    return out_({ ok: true, state: st, device: dev.name, at: new Date().toISOString() });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 招待と端末 ---------- */

function claim_(b) {
  const code = String(b.code || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (code.length !== 8) return { ok: false, error: 'bad_invite' };
  const props = PropertiesService.getScriptProperties();
  const pk = 'inv_' + hash_(code);
  const raw = props.getProperty(pk);
  if (!raw) return { ok: false, error: 'bad_invite' };
  props.deleteProperty(pk);                       // 1回限り：読んだ時点で消す
  const inv = JSON.parse(raw);
  if (Date.now() > inv.exp) return { ok: false, error: 'expired' };
  const key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const sh = SpreadsheetApp.getActive().getSheetByName(DEV);
  const id = 'D' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyMMddHHmmss');
  sh.appendRow([id, inv.name, hash_(key), '有効', now_(), inv.by, now_()]);
  log_(inv.name, '端末を登録（招待：' + inv.by + '）');
  return { ok: true, key: key, name: inv.name };
}

function newInvite_(name, by) {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(k => {                 // 期限切れの招待を掃除
    if (k.indexOf('inv_') === 0) { try { if (Date.now() > JSON.parse(all[k]).exp) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); } }
  });
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid());
  let code = '';
  for (let i = 0; i < 8; i++) code += CODE_CHARS[(bytes[i] + 256) % CODE_CHARS.length];
  props.setProperty('inv_' + hash_(code), JSON.stringify({ exp: Date.now() + INVITE_MIN * 60000, name: name, by: by }));
  return code;
}

function auth_(key) {
  if (!key || typeof key !== 'string' || key.length < 32) return null;
  const sh = SpreadsheetApp.getActive().getSheetByName(DEV);
  if (!sh || sh.getLastRow() < 2) return null;
  const h = hash_(key);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][2] === h && rows[i][3] !== '停止') return { row: i + 2, name: rows[i][1], last: rows[i][6] };
  }
  return null;
}

function touch_(dev) {                            // 最終利用は6時間に1回だけ更新（書き込みを減らす）
  const last = dev.last instanceof Date ? dev.last.getTime() : Date.parse(String(dev.last).replace(/\//g, '-'));
  if (!last || Date.now() - last > 6 * 3600000) {
    SpreadsheetApp.getActive().getSheetByName(DEV).getRange(dev.row, 7).setValue(now_());
  }
}

/* ---------- データ ---------- */

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
  SpreadsheetApp.getActive().getSheetByName(DATA)
    .getRange(2, 1, DOCS.length, 3).setValues(DOCS.map(d => [d, JSON.stringify(st[d]), now_()]));
}

/* ---------- 記録・小物 ---------- */

function describe_(ops) {
  const label = { days: '献立(週)', chkclr: 'チェック全解除', dishes: '料理リスト', init: '初期データ' };
  return ops.map(o => o.t === 'day' ? '献立 ' + o.k : o.t === 'chk' ? (o.v ? '✓ ' : '□ ') + o.n : (label[o.t] || String(o.t))).join(' / ').slice(0, 500);
}

function log_(who, text) {
  const lg = SpreadsheetApp.getActive().getSheetByName(LOG);
  lg.appendRow([now_(), String(who).slice(0, 30), text]);
  if (lg.getLastRow() > 3000) lg.deleteRows(2, 1000);
}

function ensureSheet_(ss, name, headers, init) {
  if (ss.getSheetByName(name)) return;
  const sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (init) init(sh);
}

function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function now_() { return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'); }
function fmt_(c) { return c.slice(0, 4) + '-' + c.slice(4); }
function esc_(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function alert_(m) { try { SpreadsheetApp.getUi().alert(m); } catch (e) { Logger.log(m); } }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
