const APP = {
  SHEET_LISTENERS: 'リスナー一覧',
  SHEET_SETTINGS: '選択肢設定',
  SHEET_LOGS: '操作ログ',
  PROP_SS_ID: 'SPREADSHEET_ID',
  PROP_ADMIN_PASS: 'ADMIN_PASSWORD',
  PROP_APP_NAME: 'APP_NAME',
  PROP_THEME: 'THEME_COLOR',
  SESSION_TTL: 21600,
  MAX_TEXT: 500,
};

function doGet(e) {
  const page = String((e && e.parameter && e.parameter.page) || 'index').toLowerCase();
  const tplName = page === 'admin' ? 'Admin' : page === 'edit' ? 'Edit' : 'Index';
  const t = HtmlService.createTemplateFromFile(tplName);
  t.appName = getProp_(APP.PROP_APP_NAME, 'リスナーブック');
  t.themeColor = getProp_(APP.PROP_THEME, '#7c5cff');
  t.editToken = String((e && e.parameter && e.parameter.token) || '');
  return t.evaluate()
    .setTitle(t.appName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function setupListenerBook() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty(APP.PROP_SS_ID);
  let ss;
  if (ssId) {
    ss = SpreadsheetApp.openById(ssId);
  } else {
    ss = SpreadsheetApp.create('リスナーブック データ');
    props.setProperty(APP.PROP_SS_ID, ss.getId());
  }

  setupListenersSheet_(ss);
  setupSettingsSheet_(ss);
  setupLogsSheet_(ss);

  if (!props.getProperty(APP.PROP_ADMIN_PASS)) props.setProperty(APP.PROP_ADMIN_PASS, 'change-me');
  if (!props.getProperty(APP.PROP_APP_NAME)) props.setProperty(APP.PROP_APP_NAME, 'リスナーブック');
  if (!props.getProperty(APP.PROP_THEME)) props.setProperty(APP.PROP_THEME, '#7c5cff');

  return {
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    initialAdminPassword: props.getProperty(APP.PROP_ADMIN_PASS),
    message: '初期設定が完了しました。管理パスワードを必ず変更してください。'
  };
}

function getPublicConfig() {
  const settings = readSettings_();
  return {
    appName: getProp_(APP.PROP_APP_NAME, 'リスナーブック'),
    themeColor: getProp_(APP.PROP_THEME, '#7c5cff'),
    streams: settings.streams,
    sources: settings.sources,
  };
}

function submitListener(form) {
  rateLimit_('submit', 10);
  const data = sanitizeListener_(form);
  if (!data.name) throw new Error('お名前は必須です。');
  if (String(form.website || '').trim()) throw new Error('送信に失敗しました。');

  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(APP.SHEET_LISTENERS);
  const values = sh.getDataRange().getValues();
  const duplicate = values.slice(1).some(r => normalize_(r[3]) === normalize_(data.name));

  const now = new Date();
  const id = Utilities.getUuid();
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  sh.appendRow([
    id, now, now, data.name, data.reading, data.since, data.source,
    data.favoriteStreams.join(' / '), data.gamesAnime, data.hobby, data.birthday,
    data.callName, data.message, '', '', token, data.publicScope, '有効'
  ]);

  log_('CREATE', id, data.name);
  return {
    ok: true,
    duplicateWarning: duplicate,
    editUrl: ScriptApp.getService().getUrl() + '?page=edit&token=' + encodeURIComponent(token)
  };
}

function getListenerForEdit(token) {
  rateLimit_('editRead', 30);
  const row = findByToken_(token);
  if (!row) throw new Error('編集URLが無効です。');
  return rowToListener_(row.values);
}

function updateListenerByToken(token, form) {
  rateLimit_('editWrite', 10);
  const found = findByToken_(token);
  if (!found) throw new Error('編集URLが無効です。');
  const data = sanitizeListener_(form);
  if (!data.name) throw new Error('お名前は必須です。');

  const sh = getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS);
  const row = found.row;
  const old = found.values;
  const updated = [
    old[0], old[1], new Date(), data.name, data.reading, data.since, data.source,
    data.favoriteStreams.join(' / '), data.gamesAnime, data.hobby, data.birthday,
    data.callName, data.message, old[13], old[14], old[15], data.publicScope, old[17]
  ];
  sh.getRange(row, 1, 1, updated.length).setValues([updated]);
  log_('UPDATE_BY_USER', old[0], data.name);
  return { ok: true };
}

function adminLogin(password) {
  const expected = getProp_(APP.PROP_ADMIN_PASS, 'change-me');
  if (String(password || '') !== expected) {
    rateLimit_('adminLoginFail', 5);
    throw new Error('管理パスワードが違います。');
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('ADMIN_' + token, '1', APP.SESSION_TTL);
  return { token, appName: getProp_(APP.PROP_APP_NAME, 'リスナーブック') };
}

function adminGetListeners(sessionToken, filters) {
  assertAdmin_(sessionToken);
  const sh = getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS);
  const values = sh.getDataRange().getValues().slice(1)
    .filter(r => r[17] !== '削除')
    .map(rowToListener_);

  const f = filters || {};
  let list = values;
  const q = normalize_(f.query || '');
  if (q) list = list.filter(x => normalize_(x.name + ' ' + x.reading + ' ' + x.callName).includes(q));
  if (f.tag) list = list.filter(x => x.adminTag === f.tag);
  if (f.stream) list = list.filter(x => x.favoriteStreams.includes(f.stream));

  switch (f.sort) {
    case 'name': list.sort((a,b) => (a.reading || a.name).localeCompare(b.reading || b.name, 'ja')); break;
    case 'birthday': list.sort((a,b) => birthdayKey_(a.birthday) - birthdayKey_(b.birthday)); break;
    default: list.sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt));
  }
  return { listeners: list, settings: readSettings_() };
}

function adminUpdateListener(sessionToken, id, patch) {
  assertAdmin_(sessionToken);
  const found = findById_(id);
  if (!found) throw new Error('対象データが見つかりません。');
  const sh = getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS);
  const row = found.row;
  const old = found.values;
  const next = old.slice();
  next[2] = new Date();
  if (patch.name !== undefined) next[3] = clean_(patch.name, 100);
  if (patch.reading !== undefined) next[4] = clean_(patch.reading, 100);
  if (patch.callName !== undefined) next[11] = clean_(patch.callName, 100);
  if (patch.adminTag !== undefined) next[13] = clean_(patch.adminTag, 100);
  if (patch.adminMemo !== undefined) next[14] = clean_(patch.adminMemo, 2000);
  sh.getRange(row, 1, 1, next.length).setValues([next]);
  log_('UPDATE_BY_ADMIN', id, next[3]);
  return { ok: true };
}

function adminDeleteListener(sessionToken, id) {
  assertAdmin_(sessionToken);
  const found = findById_(id);
  if (!found) throw new Error('対象データが見つかりません。');
  const sh = getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS);
  sh.getRange(found.row, 18).setValue('削除');
  sh.getRange(found.row, 3).setValue(new Date());
  log_('DELETE', id, found.values[3]);
  return { ok: true };
}

function adminExportCsv(sessionToken) {
  assertAdmin_(sessionToken);
  const sh = getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS);
  const values = sh.getDataRange().getDisplayValues().filter((r, i) => i === 0 || r[17] !== '削除');
  const csv = values.map(row => row.map(csvCell_).join(',')).join('\r\n');
  return { fileName: 'listener-book-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm') + '.csv', csv };
}

function adminChangePassword(sessionToken, currentPassword, newPassword) {
  assertAdmin_(sessionToken);
  if (String(currentPassword || '') !== getProp_(APP.PROP_ADMIN_PASS, 'change-me')) throw new Error('現在のパスワードが違います。');
  if (String(newPassword || '').length < 8) throw new Error('新しいパスワードは8文字以上にしてください。');
  PropertiesService.getScriptProperties().setProperty(APP.PROP_ADMIN_PASS, String(newPassword));
  log_('CHANGE_PASSWORD', '', '');
  return { ok: true };
}

function setupListenersSheet_(ss) {
  const headers = ['ID','登録日時','更新日時','お名前','読み方','推し始めた時期','知ったきっかけ','好きな配信','好きなゲーム・アニメ','趣味','誕生日','呼び方','メッセージ','管理タグ','管理メモ','編集トークン','公開範囲','状態'];
  let sh = ss.getSheetByName(APP.SHEET_LISTENERS);
  if (!sh) sh = ss.insertSheet(APP.SHEET_LISTENERS);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#ede9fe');
  sh.setFrozenRows(1);
}

function setupSettingsSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEET_SETTINGS);
  if (!sh) sh = ss.insertSheet(APP.SHEET_SETTINGS);
  if (sh.getLastRow() > 0) return;
  sh.getRange(1,1,1,3).setValues([['種別','値','有効']]).setFontWeight('bold').setBackground('#ede9fe');
  const rows = [
    ['好きな配信','雑談','TRUE'],['好きな配信','ゲーム','TRUE'],['好きな配信','歌','TRUE'],['好きな配信','企画','TRUE'],
    ['知ったきっかけ','YouTubeのおすすめ','TRUE'],['知ったきっかけ','X（旧Twitter）','TRUE'],['知ったきっかけ','切り抜き動画','TRUE'],['知ったきっかけ','友人・知人','TRUE'],['知ったきっかけ','その他','TRUE'],
    ['管理タグ','初見さん','TRUE'],['管理タグ','覚えた','TRUE'],['管理タグ','常連さん','TRUE'],['管理タグ','要確認','TRUE']
  ];
  sh.getRange(2,1,rows.length,3).setValues(rows);
  sh.setFrozenRows(1);
}

function setupLogsSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEET_LOGS);
  if (!sh) sh = ss.insertSheet(APP.SHEET_LOGS);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,5).setValues([['日時','操作','対象ID','名前','実行者']]);
  sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#ede9fe');
  sh.setFrozenRows(1);
}

function readSettings_() {
  const sh = getSpreadsheet_().getSheetByName(APP.SHEET_SETTINGS);
  const rows = sh.getDataRange().getDisplayValues().slice(1).filter(r => String(r[2]).toUpperCase() !== 'FALSE');
  const pick = type => rows.filter(r => r[0] === type).map(r => r[1]).filter(Boolean);
  return { streams: pick('好きな配信'), sources: pick('知ったきっかけ'), tags: pick('管理タグ') };
}

function getSpreadsheet_() {
  const id = getProp_(APP.PROP_SS_ID, '');
  if (!id) throw new Error('初期設定が未完了です。setupListenerBook を一度実行してください。');
  return SpreadsheetApp.openById(id);
}

function findByToken_(token) {
  if (!token) return null;
  const sh = getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS);
  const values = sh.getDataRange().getValues();
  for (let i=1; i<values.length; i++) if (String(values[i][15]) === String(token) && values[i][17] !== '削除') return { row: i+1, values: values[i] };
  return null;
}

function findById_(id) {
  const sh = getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS);
  const values = sh.getDataRange().getValues();
  for (let i=1; i<values.length; i++) if (String(values[i][0]) === String(id)) return { row: i+1, values: values[i] };
  return null;
}

function rowToListener_(r) {
  return {
    id: String(r[0]), createdAt: toIso_(r[1]), updatedAt: toIso_(r[2]), name: String(r[3] || ''), reading: String(r[4] || ''),
    since: String(r[5] || ''), source: String(r[6] || ''), favoriteStreams: String(r[7] || '').split(' / ').filter(Boolean),
    gamesAnime: String(r[8] || ''), hobby: String(r[9] || ''), birthday: String(r[10] || ''), callName: String(r[11] || ''),
    message: String(r[12] || ''), adminTag: String(r[13] || ''), adminMemo: String(r[14] || ''), publicScope: String(r[16] || '')
  };
}

function sanitizeListener_(form) {
  form = form || {};
  return {
    name: clean_(form.name, 100), reading: clean_(form.reading, 100), since: clean_(form.since, 100), source: clean_(form.source, 100),
    favoriteStreams: Array.isArray(form.favoriteStreams) ? form.favoriteStreams.map(x => clean_(x, 50)).filter(Boolean).slice(0, 20) : [],
    gamesAnime: clean_(form.gamesAnime, APP.MAX_TEXT), hobby: clean_(form.hobby, APP.MAX_TEXT), birthday: clean_(form.birthday, 20),
    callName: clean_(form.callName, 100), message: clean_(form.message, 1000), publicScope: clean_(form.publicScope, 100) || '配信者のみ',
  };
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[<>]/g, '').replace(/\u0000/g, '').trim().slice(0, max || APP.MAX_TEXT);
}
function normalize_(v) { return String(v || '').normalize('NFKC').replace(/[\s　]+/g, '').toLowerCase(); }
function getProp_(key, fallback) { return PropertiesService.getScriptProperties().getProperty(key) || fallback; }
function toIso_(v) { try { return new Date(v).toISOString(); } catch(e) { return ''; } }
function birthdayKey_(v) { const m=String(v||'').match(/(\d{1,2})\D+(\d{1,2})/); return m ? Number(m[1])*100+Number(m[2]) : 9999; }
function csvCell_(v) { const s=String(v==null?'':v).replace(/"/g,'""'); return '"'+s+'"'; }
function assertAdmin_(token) { if (!token || CacheService.getScriptCache().get('ADMIN_' + token) !== '1') throw new Error('管理セッションの有効期限が切れました。もう一度ログインしてください。'); }
function rateLimit_(key, seconds) { const c=CacheService.getScriptCache(), k='RL_'+key+'_'+(Session.getTemporaryActiveUserKey()||'anon'); if(c.get(k)) throw new Error('連続操作を検知しました。少し時間を空けてください。'); c.put(k,'1',seconds); }
function log_(action,id,name) { try { getSpreadsheet_().getSheetByName(APP.SHEET_LOGS).appendRow([new Date(),action,id,name,Session.getActiveUser().getEmail()||'anonymous']); } catch(e) {} }
