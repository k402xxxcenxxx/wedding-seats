// 婚禮照片後端：賓客免登入上傳到 Google 雲端硬碟，資料記在試算表「婚禮照片後台」。
// 後台就是那份試算表：「照片」「愛心」看資料，「設定」改設定，「統計」看數字。
// 部署：執行一次 setup() 並授權，再部署成網頁應用程式（執行身分：我；存取權：所有人）。

const P = PropertiesService.getScriptProperties();
const book = () => SpreadsheetApp.openById(P.getProperty('SHEET_ID'));
const json = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
// 去掉開頭的 = + - @，避免賓客輸入的文字在試算表裡被當成公式執行。
const clean = (s, n) => String(s == null ? '' : s).replace(/^[=+\-@\s]+/, '').slice(0, n);
const okId = s => /^[\w-]{8,80}$/.test(s);

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
  ss.insertSheet('設定').getRange(1, 1, 5, 3).setValues([
    ['設定', '值', '說明'],
    ['maxVideoSeconds', 30, '影片長度上限（秒）'],
    ['maxFileMB', 300, '單一檔案大小上限（MB）'],
    ['uploadOpen', true, 'TRUE 開放上傳；FALSE 關閉上傳'],
    ['wallSeconds', 6, '投影牆每張顯示秒數'],
  ]);
  ss.insertSheet('統計').getRange(1, 1, 6, 2).setValues([
    ['上傳總數', '=COUNTA(\'照片\'!B2:B)'],
    ['照片', '=COUNTIF(\'照片\'!C2:C,"image")'],
    ['影片', '=COUNTIF(\'照片\'!C2:C,"video")'],
    ['上傳人數（裝置）', '=COUNTUNIQUE(\'照片\'!D2:D)'],
    ['愛心總數', '=COUNTA(\'愛心\'!A2:A)'],
    ['每人上傳數', '=IFERROR(QUERY(\'照片\'!B:E,"select E, count(B) where B is not null group by E order by count(B) desc label count(B) \'\'",1),"")'],
  ]);
  P.setProperty('SHEET_ID', ss.getId());
}

function settings(ss) {
  const o = {};
  ss.getSheetByName('設定').getDataRange().getValues().slice(1).forEach(r => { o[r[0]] = r[1]; });
  return o;
}

// 回傳設定與未隱藏的照片（新到舊），帶 ?user= 時附上該裝置是否按過愛心。
function doGet(e) {
  const me = e.parameter.user || '';
  const ss = book();
  const likes = {}, mine = {};
  ss.getSheetByName('愛心').getDataRange().getValues().slice(1).forEach(r => {
    likes[r[0]] = (likes[r[0]] || 0) + 1;
    if (r[1] === me) mine[r[0]] = true;
  });
  const photos = ss.getSheetByName('照片').getDataRange().getValues().slice(1)
    .filter(r => r[1] && !r[7])
    .map(r => ({ id: r[1], v: r[2] === 'video', name: String(r[4]), tag: String(r[5]), msg: String(r[6]),
                 likes: likes[r[1]] || 0, liked: !!mine[r[1]] }))
    .reverse();
  return json({ settings: settings(ss), photos: photos });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!okId(d.user)) return json({ error: 'user' });
    if (d.action === 'like') return json(like(d));
    if (d.action === 'put') return json(put(d));
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
    clean(d.userName, 20), d.user.slice(0, 4), clean(d.msg, 100), '']);
  return { ok: true, done: true };
}

// 在編輯器執行這個函式可叫出授權視窗（put 需要「連線到外部服務」權限）。
function authorize() {
  UrlFetchApp.fetch('https://www.google.com/generate_204');
}
