// 婚禮照片後端：賓客免登入上傳到 Google 雲端硬碟，資料記在試算表「婚禮照片後台」。
// 後台就是那份試算表：「照片」「愛心」看資料，「設定」改設定，「統計」看數字。
// 部署：執行一次 setup() 並授權，再部署成網頁應用程式（執行身分：我；存取權：所有人）。

const P = PropertiesService.getScriptProperties();
const book = () => SpreadsheetApp.openById(P.getProperty('SHEET_ID'));
const json = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
// 去掉開頭的 = + - @，避免賓客輸入的文字在試算表裡被當成公式執行。
const clean = (s, n) => String(s == null ? '' : s).replace(/^[=+\-@\s]+/, '').slice(0, n);
const okId = s => /^[\w-]{8,80}$/.test(s);
// 通關金鑰只存在指令碼屬性裡，不進 GitHub；沒設定時一律拒絕。
const okKey = k => !!k && k === P.getProperty('KEY');

function setup() {
  if (!P.getProperty('FOLDER_ID')) {
    const f = DriveApp.createFolder('婚禮賓客照片');
    // 相簿和投影牆用雲端硬碟縮圖網址顯示，所以資料夾要開「知道連結的人可檢視」。
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    P.setProperty('FOLDER_ID', f.getId());
  }
  if (P.getProperty('SHEET_ID')) return;
  const ss = SpreadsheetApp.create('婚禮照片後台');
  const photos = ss.getSheets()[0].setName('照片');
  photos.appendRow(['時間', '檔案ID', '類型', '裝置ID', '姓名', '識別碼', '祝福語', '隱藏（填任何字就不顯示）']);
  // 把改版前已上傳的檔案補進來。
  const it = DriveApp.getFolderById(P.getProperty('FOLDER_ID')).getFiles();
  while (it.hasNext()) {
    const f = it.next();
    photos.appendRow([f.getDateCreated(), f.getId(), f.getMimeType().split('/')[0], '', '', '', '', '']);
  }
  ss.insertSheet('愛心').appendRow(['檔案ID', '裝置ID', '時間']);
  ss.insertSheet('統計').getRange(1, 1, 6, 2).setValues([
    ['上傳總數', '=COUNTA(\'照片\'!B2:B)'],
    ['照片', '=COUNTIF(\'照片\'!C2:C,"image")'],
    ['影片', '=COUNTIF(\'照片\'!C2:C,"video")'],
    ['上傳人數（裝置）', '=COUNTUNIQUE(\'照片\'!D2:D)'],
    ['愛心總數', '=COUNTA(\'愛心\'!A2:A)'],
    ['每人上傳數', '=IFERROR(QUERY(\'照片\'!B:E,"select E, count(B) where B is not null group by E order by count(B) desc label count(B) \'\'",1),"")'],
  ]);
  P.setProperty('SHEET_ID', ss.getId());
  upgrade();
}

const DEFAULTS = [
  ['maxVideoSeconds', 30, '影片長度上限（秒）'],
  ['maxFileMB', 300, '單一檔案大小上限（MB）'],
  ['uploadOpen', true, 'TRUE 開放上傳與留言；FALSE 關閉'],
  ['wallSeconds', 6, '投影牆每張顯示秒數'],
  ['review', false, 'TRUE 新上傳與彈幕先不公開，把該列「隱藏」欄的「待審」清掉才顯示；FALSE 直接顯示'],
  ['maxPerDevice', 50, '每個裝置最多可上傳幾個檔案（0 不限）'],
  ['wishPosition', '上', '彈幕跑馬燈位置：上 或 下'],
  ['wishSpeed', 120, '彈幕速度：每秒移動幾像素，數字越大越快'],
  ['wishSize', 40, '彈幕字體大小（像素）'],
  ['maxWishesPerDevice', 20, '每個裝置最多可留幾則彈幕（0 不限）'],
  ['wallVideo', '', '投影牆輪播影片：貼上 Google 雲端硬碟影片的分享連結（共用設定需為「知道連結的任何人」）'],
];

// 可重複執行：補上後台缺少的分頁與設定列，已有的不會動。改版後在編輯器執行一次。
function upgrade() {
  const ss = book();
  const st = ss.getSheetByName('設定') || ss.insertSheet('設定');
  if (!st.getLastRow()) st.appendRow(['設定', '值', '說明']);
  const have = st.getDataRange().getValues().map(r => r[0]);
  DEFAULTS.forEach(r => { if (have.indexOf(r[0]) < 0) st.appendRow(r); });
  if (!ss.getSheetByName('彈幕')) {
    ss.insertSheet('彈幕').appendRow(['時間', '裝置ID', '姓名', '識別碼', '內容', '隱藏（填任何字就不顯示）']);
  }
}

function settings(ss) {
  const o = {};
  ss.getSheetByName('設定').getDataRange().getValues().slice(1).forEach(r => { o[r[0]] = r[1]; });
  return o;
}

// 回傳設定與未隱藏的照片（新到舊），帶 ?user= 時附上該裝置是否按過愛心。
function doGet(e) {
  if (!okKey(e.parameter.k)) return json({ error: 'key' });
  const me = e.parameter.user || '';
  const ss = book();
  const likes = {}, mine = {};
  ss.getSheetByName('愛心').getDataRange().getValues().slice(1).forEach(r => {
    likes[r[0]] = (likes[r[0]] || 0) + 1;
    if (r[1] === me) mine[r[0]] = true;
  });
  const photos = ss.getSheetByName('照片').getDataRange().getValues().slice(1)
    .filter(r => r[1] && (!r[7] || (r[7] === '待審' && me && r[3] === me)))   // 待審的只給上傳者自己看
    .map(r => ({ id: r[1], v: r[2] === 'video', name: String(r[4]), tag: String(r[5]), msg: String(r[6]), wait: !!r[7],
                 likes: likes[r[1]] || 0, liked: !!mine[r[1]] }))
    .reverse();
  // 彈幕：未隱藏的最新 200 則，依收到時間排序；t 讓投影牆記得播到哪裡，每則只播一次。
  const wishes = ss.getSheetByName('彈幕').getDataRange().getValues().slice(1)
    .filter(r => r[4] && !r[5]).slice(-200)
    .map(r => ({ t: new Date(r[0]).getTime(), name: String(r[2]), text: String(r[4]) }))
    .sort((a, b) => a.t - b.t);
  return json({ settings: settings(ss), photos: photos, wishes: wishes });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!okKey(d.k)) return json({ error: 'key' });
    if (d.action === 'seat') return json(seat(d));
    if (!okId(d.user)) return json({ error: 'user' });
    if (d.action === 'like') return json(like(d));
    if (d.action === 'put') return json(put(d));
    if (d.action === 'wish') return json(wish(d));
    return json({ error: 'action' });
  } catch (err) {
    return json({ error: String(err) });
  }
}

// 同一裝置對同一張照片按一次是加愛心，再按一次是收回。
function like(d) {
  if (!okId(d.id)) return { error: 'id' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = book().getSheetByName('愛心');
    const i = sh.getDataRange().getValues().findIndex(r => r[0] === d.id && r[1] === d.user);
    if (i >= 0) sh.deleteRow(i + 1); else sh.appendRow([d.id, d.user, new Date()]);
    return { ok: true, liked: i < 0 };
  } finally {
    lock.releaseLock();
  }
}

// 分段上傳：每段由這裡轉送到 Drive 的續傳上傳，所以檔案大小不受 Apps Script 單次請求 50 MB 的限制。
// shortcut: 影片秒數只在賓客的瀏覽器檢查，這裡只擋檔案大小；要嚴格擋秒數得等 Drive 轉檔完再讀影片長度。
function put(d) {
  const ss = book(), s = settings(ss);
  if (!s.uploadOpen) return { error: 'closed' };
  if (!/^(image|video)\/[\w.+-]+$/.test(d.type)) return { error: 'type' };
  const total = Number(d.total), offset = Number(d.offset);
  if (!(total > 0 && total <= s.maxFileMB * 1048576 && offset >= 0)) return { error: 'size' };
  let url = d.session;
  if (!url) {
    if (s.maxPerDevice > 0 && ss.getSheetByName('照片').getDataRange().getValues()
        .filter(r => r[3] === d.user).length >= s.maxPerDevice) return { error: 'quota' };
    const r = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
                 'X-Upload-Content-Type': d.type, 'X-Upload-Content-Length': String(total) },
      payload: JSON.stringify({ name: clean(d.name, 100) || 'upload', parents: [P.getProperty('FOLDER_ID')] }),
    });
    const h = r.getHeaders();
    url = h.Location || h.location;
  } else if (String(url).indexOf('https://www.googleapis.com/upload/drive/v3/files?') !== 0) {
    return { error: 'session' };
  }
  const bytes = Utilities.base64Decode(d.data);
  const r = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/octet-stream',
    headers: { 'Content-Range': 'bytes ' + offset + '-' + (offset + bytes.length - 1) + '/' + total },
    payload: bytes,
    muteHttpExceptions: true,
  });
  const code = r.getResponseCode();
  if (code === 308) return { ok: true, session: url };
  if (code !== 200 && code !== 201) return { error: 'drive ' + code };
  const id = JSON.parse(r.getContentText()).id;
  ss.getSheetByName('照片').appendRow([new Date(), id, d.type.split('/')[0], d.user,
    clean(d.userName, 20), d.user.slice(0, 4), clean(d.msg, 100), s.review ? '待審' : '']);
  return { ok: true, done: true };
}

// 彈幕留言：不用傳照片，只留一句話，會以跑馬燈經過投影牆。
function wish(d) {
  const ss = book(), s = settings(ss), text = clean(d.text, 50);
  if (!s.uploadOpen) return { error: 'closed' };
  if (!text) return { error: 'empty' };
  const sh = ss.getSheetByName('彈幕');
  if (s.maxWishesPerDevice > 0 && sh.getDataRange().getValues()
      .filter(r => r[1] === d.user).length >= s.maxWishesPerDevice) return { error: 'quota' };
  sh.appendRow([new Date(), d.user, clean(d.userName, 20), d.user.slice(0, 4), text, s.review ? '待審' : '']);
  return { ok: true, review: !!s.review };
}

// 座位查詢：名單在另一份試算表（ID 存在指令碼屬性 SEAT_SHEET_ID，需分享給執行帳號），
// A 欄桌名、B 欄姓名、C 欄可選填桌號。至少要輸入兩個字、最多回十筆，避免被人整份撈走。
function seat(d) {
  const norm = x => String(x).replace(/\s+/g, '').toLowerCase();
  const q = norm(d.q || '');
  if (q.length < 2) return { error: 'short' };
  const rows = SpreadsheetApp.openById(P.getProperty('SEAT_SHEET_ID')).getSheets()[0].getDataRange().getValues()
    .map(r => ({ table: String(r[0]).trim(), name: String(r[1]).trim(), no: String(r[2] == null ? '' : r[2]).trim() }))
    .filter(r => r.name && r.name !== '姓名');
  const hits = rows.filter(r => norm(r.name).indexOf(q) >= 0).slice(0, 10);
  return { ok: true, hits: hits.map(h => ({ name: h.name, table: h.table, no: h.no,
    mates: rows.filter(o => o !== h && o.table === h.table).map(o => o.name) })) };
}

// 在編輯器執行這個函式可叫出授權視窗（put 需要「連線到外部服務」權限）。
function authorize() {
  UrlFetchApp.fetch('https://www.google.com/generate_204');
}

// 在編輯器執行：第一次會產生金鑰，之後只顯示。要換金鑰就到「專案設定 > 指令碼屬性」刪掉 KEY 再執行一次。
function showKey() {
  if (!P.getProperty('KEY')) P.setProperty('KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 16));
  Logger.log(P.getProperty('KEY'));
}
