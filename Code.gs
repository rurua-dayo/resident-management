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

const LISTENER_HEADERS = [
  'ID','登録日時','更新日時','Xのお名前','Xの読み方','XのURL','YouTubeのお名前','YouTubeの読み方',
  '推し始めた時期','知ったきっかけ','よく見る配信','配信をよく見る時間帯',
  '血液型','MBTI診断','好きなゲーム・アニメ','趣味','誕生日','呼び方','メッセージ',
  '管理タグ','管理メモ','編集トークン','公開範囲','状態'
];

function doGet(e) {
  const page = String((e && e.parameter && e.parameter.page) || 'index').toLowerCase();
  const tplName = page === 'admin' ? 'Admin' : page === 'edit' ? 'Edit' : 'Index';
  const t = HtmlService.createTemplateFromFile(tplName);
  t.appName = getProp_(APP.PROP_APP_NAME, '物置の住人登録届');
  t.themeColor = getProp_(APP.PROP_THEME, '#7c5cff');
  t.editToken = String((e && e.parameter && e.parameter.token) || '');
  return t.evaluate().setTitle(t.appName).addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

function setupListenerBook() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty(APP.PROP_SS_ID);
  let ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.create('物置の住人登録届 データ');
  if (!ssId) props.setProperty(APP.PROP_SS_ID, ss.getId());

  setupListenersSheet_(ss);
  setupSettingsSheet_(ss);
  setupLogsSheet_(ss);

  props.setProperty(APP.PROP_ADMIN_PASS, 'mimorin');
  if (!props.getProperty(APP.PROP_APP_NAME)) props.setProperty(APP.PROP_APP_NAME, '物置の住人登録届');
  if (!props.getProperty(APP.PROP_THEME)) props.setProperty(APP.PROP_THEME, '#7c5cff');

  return {
    spreadsheetId: ss.getId(), spreadsheetUrl: ss.getUrl(),
    initialAdminPassword: props.getProperty(APP.PROP_ADMIN_PASS),
    message: '初期設定と項目更新が完了しました。管理パスワードを必ず変更してください。'
  };
}

function applyMonookiDesign() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(APP.PROP_APP_NAME, '物置の住人登録届');
  props.setProperty(APP.PROP_THEME, '#f28fb3');
  return 'サイト名とテーマカラーを更新しました。';
}

function getPublicConfig() {
  const settings = readSettings_();
  return { appName:getProp_(APP.PROP_APP_NAME,'物置の住人登録届'), themeColor:getProp_(APP.PROP_THEME,'#7c5cff'), streams:settings.streams, sources:settings.sources };
}

function submitListener(form) {
  rateLimit_('submit', 10);
  const data = sanitizeListener_(form);
  if (!data.xName) throw new Error('Xのお名前は必須です。');
  if (String(form.website || '').trim()) throw new Error('送信に失敗しました。');

  const sh = getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS);
  const values = sh.getDataRange().getValues();
  const duplicate = values.slice(1).some(r =>
    (data.xName && normalize_(r[3]) === normalize_(data.xName)) ||
    (data.youtubeName && normalize_(r[6]) === normalize_(data.youtubeName))
  );

  const now = new Date(), id = Utilities.getUuid();
  const token = Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
  sh.appendRow([
    id, now, now, data.xName, data.xReading, data.xUrl, data.youtubeName, data.youtubeReading,
    data.since, data.source, data.favoriteStreams.join(' / '), data.watchTime,
    data.bloodType, data.mbti, data.gamesAnime, data.hobby, data.birthday, data.callName, data.message,
    '', '', token, data.publicScope, '有効'
  ]);
  log_('CREATE', id, data.xName);
  return { ok:true, duplicateWarning:duplicate, editUrl:ScriptApp.getService().getUrl()+'?page=edit&token='+encodeURIComponent(token) };
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
  if (!data.xName) throw new Error('Xのお名前は必須です。');
  const old = found.values;
  const updated = [
    old[0], old[1], new Date(), data.xName, data.xReading, data.xUrl, data.youtubeName, data.youtubeReading,
    data.since, data.source, data.favoriteStreams.join(' / '), data.watchTime,
    data.bloodType, data.mbti, data.gamesAnime, data.hobby, data.birthday, data.callName, data.message,
    old[19], old[20], old[21], data.publicScope, old[23]
  ];
  getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS).getRange(found.row,1,1,updated.length).setValues([updated]);
  log_('UPDATE_BY_USER', old[0], data.xName);
  return { ok:true };
}

function adminLogin(password) {
  const expected = getProp_(APP.PROP_ADMIN_PASS, 'mimorin');
  if (String(password || '') !== expected) { rateLimit_('adminLoginFail',5); throw new Error('管理パスワードが違います。'); }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('ADMIN_'+token,'1',APP.SESSION_TTL);
  return { token, appName:getProp_(APP.PROP_APP_NAME,'物置の住人登録届') };
}

function adminGetListeners(sessionToken, filters) {
  assertAdmin_(sessionToken);
  const values = getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS).getDataRange().getValues().slice(1)
    .filter(r => r[23] !== '削除').map(rowToListener_);
  const f = filters || {}; let list = values; const q = normalize_(f.query || '');
  if (q) list = list.filter(x => normalize_([x.xName,x.xReading,x.xUrl,x.youtubeName,x.youtubeReading,x.callName].join(' ')).includes(q));
  if (f.tag) list = list.filter(x => x.adminTag === f.tag);
  if (f.stream) list = list.filter(x => x.favoriteStreams.includes(f.stream));
  switch (f.sort) {
    case 'name': list.sort((a,b)=>(a.xReading||a.xName||a.youtubeReading||a.youtubeName).localeCompare(b.xReading||b.xName||b.youtubeReading||b.youtubeName,'ja')); break;
    case 'birthday': list.sort((a,b)=>birthdayKey_(a.birthday)-birthdayKey_(b.birthday)); break;
    default: list.sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));
  }
  return { listeners:list, settings:readSettings_() };
}

function adminGetAllListeners(sessionToken) {
  assertAdmin_(sessionToken);
  const list = getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS).getDataRange().getValues().slice(1)
    .filter(r => r[23] !== '削除')
    .map(rowToListener_)
    .sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt));
  return list;
}

function adminUpdateListener(sessionToken, id, patch) {
  assertAdmin_(sessionToken);
  const found = findById_(id); if (!found) throw new Error('対象データが見つかりません。');
  const next = found.values.slice(); next[2] = new Date();
  if (patch.xName !== undefined) next[3] = clean_(patch.xName,100);
  if (patch.xReading !== undefined) next[4] = clean_(patch.xReading,100);
  if (patch.xUrl !== undefined) next[5] = cleanUrl_(patch.xUrl);
  if (patch.youtubeName !== undefined) next[6] = clean_(patch.youtubeName,100);
  if (patch.youtubeReading !== undefined) next[7] = clean_(patch.youtubeReading,100);
  if (patch.bloodType !== undefined) next[12] = clean_(patch.bloodType,20);
  if (patch.mbti !== undefined) next[13] = clean_(patch.mbti,200);
  if (patch.callName !== undefined) next[17] = clean_(patch.callName,100);
  if (patch.adminTag !== undefined) next[19] = clean_(patch.adminTag,100);
  if (patch.adminMemo !== undefined) next[20] = clean_(patch.adminMemo,2000);
  getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS).getRange(found.row,1,1,next.length).setValues([next]);
  log_('UPDATE_BY_ADMIN',id,next[3]); return {ok:true};
}

function adminDeleteListener(sessionToken, id) {
  assertAdmin_(sessionToken); const found = findById_(id); if (!found) throw new Error('対象データが見つかりません。');
  const sh = getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS);
  sh.getRange(found.row,24).setValue('削除'); sh.getRange(found.row,3).setValue(new Date());
  log_('DELETE',id,found.values[3]); return {ok:true};
}

function adminChangePassword(sessionToken,currentPassword,newPassword) {
  assertAdmin_(sessionToken);
  if (String(currentPassword||'') !== getProp_(APP.PROP_ADMIN_PASS,'mimorin')) throw new Error('現在のパスワードが違います。');
  if (String(newPassword||'').length < 8) throw new Error('新しいパスワードは8文字以上にしてください。');
  PropertiesService.getScriptProperties().setProperty(APP.PROP_ADMIN_PASS,String(newPassword)); log_('CHANGE_PASSWORD','',''); return {ok:true};
}

function setupListenersSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEET_LISTENERS);
  if (!sh) sh = ss.insertSheet(APP.SHEET_LISTENERS);

  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,LISTENER_HEADERS.length).setValues([LISTENER_HEADERS]);
  } else {
    const oldHeaders = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
    const oldData = sh.getDataRange().getValues().slice(1);
    const aliases = {
      'Xのお名前':['Xのお名前','お名前'],
      'Xの読み方':['Xの読み方','読み方']
    };
    const migrated = oldData.map(row => LISTENER_HEADERS.map(header => {
      const names = aliases[header] || [header];
      for (const name of names) {
        const idx = oldHeaders.indexOf(name);
        if (idx >= 0) return row[idx];
      }
      return '';
    }));
    sh.clearContents();
    sh.getRange(1,1,1,LISTENER_HEADERS.length).setValues([LISTENER_HEADERS]);
    if (migrated.length) sh.getRange(2,1,migrated.length,LISTENER_HEADERS.length).setValues(migrated);
  }
  sh.getRange(1,1,1,LISTENER_HEADERS.length).setFontWeight('bold').setBackground('#ede9fe');
  sh.setFrozenRows(1);
}

function setupSettingsSheet_(ss) {
  let sh = ss.getSheetByName(APP.SHEET_SETTINGS); if (!sh) sh = ss.insertSheet(APP.SHEET_SETTINGS);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,3).setValues([['種別','値','有効']]);
  let rows = sh.getDataRange().getDisplayValues().slice(1).filter(r=>r.some(Boolean));
  rows = rows.map(r => [r[0] === '好きな配信' ? 'よく見る配信' : r[0], r[1], r[2] || 'TRUE']);
  // 管理タグは指定の4種類へ置き換えるため、既存の管理タグ設定をいったん除外する。
  rows = rows.filter(r => r[0] !== '管理タグ');
  const defaults = [
    ['よく見る配信','雑談','TRUE'],['よく見る配信','ゲーム','TRUE'],['よく見る配信','歌','TRUE'],['よく見る配信','企画','TRUE'],
    ['よく見る配信','案件','TRUE'],['よく見る配信','正拳突き','TRUE'],['よく見る配信','朗読','TRUE'],['よく見る配信','おはよう','TRUE'],
    ['知ったきっかけ','YouTubeのおすすめ','TRUE'],['知ったきっかけ','X（旧Twitter）','TRUE'],['知ったきっかけ','切り抜き動画','TRUE'],['知ったきっかけ','友人・知人','TRUE'],['知ったきっかけ','その他','TRUE'],
    ['管理タグ','覚えた','TRUE'],['管理タグ','常連さん','TRUE'],['管理タグ','モデレーター','TRUE'],['管理タグ','要注意人物','TRUE']
  ];
  const keys = new Set(rows.map(r=>r[0]+'\t'+r[1]));
  defaults.forEach(r=>{ if(!keys.has(r[0]+'\t'+r[1])) rows.push(r); });
  sh.clearContents(); sh.getRange(1,1,1,3).setValues([['種別','値','有効']]);
  if (rows.length) sh.getRange(2,1,rows.length,3).setValues(rows);
  sh.getRange(1,1,1,3).setFontWeight('bold').setBackground('#ede9fe'); sh.setFrozenRows(1);
}

function setupLogsSheet_(ss) {
  let sh=ss.getSheetByName(APP.SHEET_LOGS); if(!sh) sh=ss.insertSheet(APP.SHEET_LOGS);
  if(sh.getLastRow()===0) sh.getRange(1,1,1,5).setValues([['日時','操作','対象ID','名前','実行者']]);
  sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#ede9fe'); sh.setFrozenRows(1);
}

function readSettings_() {
  const sh=getSpreadsheet_().getSheetByName(APP.SHEET_SETTINGS);
  const rows=sh.getDataRange().getDisplayValues().slice(1).filter(r=>String(r[2]).toUpperCase()!=='FALSE');
  const pick=type=>rows.filter(r=>r[0]===type).map(r=>r[1]).filter(Boolean);
  return { streams:pick('よく見る配信').concat(pick('好きな配信')).filter((v,i,a)=>a.indexOf(v)===i), sources:pick('知ったきっかけ'), tags:pick('管理タグ') };
}

function getSpreadsheet_() { const id=getProp_(APP.PROP_SS_ID,''); if(!id) throw new Error('初期設定が未完了です。setupListenerBook を一度実行してください。'); return SpreadsheetApp.openById(id); }
function findByToken_(token) { if(!token)return null; const sh=getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS), values=sh.getDataRange().getValues(); for(let i=1;i<values.length;i++) if(String(values[i][21])===String(token)&&values[i][23]!=='削除') return {row:i+1,values:values[i]}; return null; }
function findById_(id) { const sh=getSpreadsheet_().getSheetByName(APP.SHEET_LISTENERS), values=sh.getDataRange().getValues(); for(let i=1;i<values.length;i++) if(String(values[i][0])===String(id)) return {row:i+1,values:values[i]}; return null; }
function rowToListener_(r) {
  const editToken = String(r[21] || '');
  return {
    id:String(r[0]),createdAt:toIso_(r[1]),updatedAt:toIso_(r[2]),xName:String(r[3]||''),xReading:String(r[4]||''),xUrl:String(r[5]||''),
    youtubeName:String(r[6]||''),youtubeReading:String(r[7]||''),since:monthInput_(r[8]),sinceDisplay:formatMonthJa_(r[8]),source:String(r[9]||''),
    favoriteStreams:String(r[10]||'').split(' / ').filter(Boolean),watchTime:String(r[11]||''),bloodType:String(r[12]||''),mbti:String(r[13]||''),gamesAnime:String(r[14]||''),
    hobby:String(r[15]||''),birthday:dateInput_(r[16]),birthdayDisplay:formatDateJa_(r[16]),callName:String(r[17]||''),message:String(r[18]||''),
    adminTag:String(r[19]||''),adminMemo:String(r[20]||''),publicScope:String(r[22]||''),
    editUrl: editToken ? ScriptApp.getService().getUrl() + '?page=edit&token=' + encodeURIComponent(editToken) : ''
  };
}

function sanitizeListener_(form) { form=form||{}; return { xName:clean_(form.xName,100),xReading:clean_(form.xReading,100),xUrl:cleanUrl_(form.xUrl),youtubeName:clean_(form.youtubeName,100),youtubeReading:clean_(form.youtubeReading,100),since:clean_(form.since,100),source:clean_(form.source,100),favoriteStreams:Array.isArray(form.favoriteStreams)?form.favoriteStreams.map(x=>clean_(x,50)).filter(Boolean).slice(0,20):[],watchTime:clean_(form.watchTime,100),bloodType:clean_(form.bloodType,20),mbti:clean_(form.mbti,200),gamesAnime:clean_(form.gamesAnime,APP.MAX_TEXT),hobby:clean_(form.hobby,APP.MAX_TEXT),birthday:clean_(form.birthday,20),callName:clean_(form.callName,100),message:clean_(form.message,1000),publicScope:clean_(form.publicScope,100)||'配信者のみ' }; }
function cleanUrl_(v){const s=clean_(v,300);if(!s)return '';if(!/^https:\/\/(x\.com|twitter\.com)\//i.test(s))throw new Error('XのURLは https://x.com/ から始まるURLを入力してください。');return s}
function clean_(v,max){return String(v==null?'':v).replace(/[<>]/g,'').replace(/\u0000/g,'').trim().slice(0,max||APP.MAX_TEXT)}
function normalize_(v){return String(v||'').normalize('NFKC').replace(/[\s　]+/g,'').toLowerCase()}
function monthInput_(v){
  if(!v)return '';
  if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v))return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM');
  const s=String(v).trim();
  const m=s.match(/^(\d{4})[-\/年](\d{1,2})/);
  if(m)return m[1]+'-'+String(Number(m[2])).padStart(2,'0');
  const d=new Date(s);return isNaN(d)?s:Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM');
}
function dateInput_(v){
  if(!v)return '';
  if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v))return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const s=String(v).trim();
  const m=s.match(/^(\d{4})[-\/年](\d{1,2})[-\/月](\d{1,2})/);
  if(m)return m[1]+'-'+String(Number(m[2])).padStart(2,'0')+'-'+String(Number(m[3])).padStart(2,'0');
  const d=new Date(s);return isNaN(d)?s:Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd');
}
function formatMonthJa_(v){const s=monthInput_(v),m=s.match(/^(\d{4})-(\d{2})$/);return m?m[1]+'年'+Number(m[2])+'月':s}
function formatDateJa_(v){const s=dateInput_(v),m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?m[1]+'年'+Number(m[2])+'月'+Number(m[3])+'日':s}
function getProp_(key,fallback){return PropertiesService.getScriptProperties().getProperty(key)||fallback}
function toIso_(v){try{return new Date(v).toISOString()}catch(e){return ''}}
function birthdayKey_(v){const s=String(v||'');let m=s.match(/^\d{4}-(\d{2})-(\d{2})$/);if(m)return Number(m[1])*100+Number(m[2]);m=s.match(/(\d{1,2})\D+(\d{1,2})/);return m?Number(m[1])*100+Number(m[2]):9999}
function csvCell_(v){const s=String(v==null?'':v).replace(/"/g,'""');return '"'+s+'"'}
function assertAdmin_(token){if(!token||CacheService.getScriptCache().get('ADMIN_'+token)!=='1')throw new Error('管理セッションの有効期限が切れました。もう一度ログインしてください。')}
function rateLimit_(key,seconds){const c=CacheService.getScriptCache(),k='RL_'+key+'_'+(Session.getTemporaryActiveUserKey()||'anon');if(c.get(k))throw new Error('連続操作を検知しました。少し時間を空けてください。');c.put(k,'1',seconds)}
function log_(action,id,name){try{getSpreadsheet_().getSheetByName(APP.SHEET_LOGS).appendRow([new Date(),action,id,name,Session.getActiveUser().getEmail()||'anonymous'])}catch(e){}}
