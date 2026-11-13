// 婚禮照片上傳後端：賓客免登入上傳到你的 Google 雲端硬碟資料夾，投影牆讀取清單。
// 部署：先執行一次 setup() 並授權，再「部署 > 新增部署 > 網頁應用程式」，
// 執行身分選「我」、存取權選「所有人」，把網址貼到 photos.html 和 wall.html 的 API。

function setup() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('FOLDER_ID')) return;
  const f = DriveApp.createFolder('婚禮賓客照片');
  // 投影牆用雲端硬碟縮圖網址顯示，所以資料夾要開「知道連結的人可檢視」。
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  p.setProperty('FOLDER_ID', f.getId());
}

function folder() {
  return DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!/^(image|video)\/[\w.+-]+$/.test(d.type)) return json({ error: 'type' });
    const name = String(d.name || 'upload').replace(/[\\/]/g, '_').slice(0, 100);
    folder().createFile(Utilities.newBlob(Utilities.base64Decode(d.data), d.type, name));
    return json({ ok: true });
  } catch (err) {
    return json({ error: String(err) });
  }
}

// shortcut: 每次都走訪整個資料夾，幾千個檔案內沒問題；更多時改用 Drive API 的 orderBy 分頁。
function doGet() {
  const out = [];
  const it = folder().getFiles();
  while (it.hasNext()) {
    const f = it.next();
    out.push({ id: f.getId(), t: f.getDateCreated().getTime() });
  }
  out.sort((a, b) => b.t - a.t);
  return json(out);
}
