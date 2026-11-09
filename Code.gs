/**
 * LINE網頁圖卡分享建立平台 — Google Apps Script 後端
 * 獨立的 Apps Script 專案：第一次執行 setup() 會自動建立試算表「LINE網頁圖卡」（記在指令碼屬性 SHEET_ID）。
 * 更新程式後執行一次 setup() 補齊欄位，再部署新版本。
 */

/** 資料用的試算表：第一次自動建立 */
function book() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  if (!id) { id = SpreadsheetApp.create('LINE網頁圖卡').getId(); props.setProperty('SHEET_ID', id); }
  return SpreadsheetApp.openById(id);
}

const CONFIG = {
  CHANNEL_ID: '2012412693', // LINE Developers → LINE Login channel → Basic settings
  IMAGE_FOLDER: 'LINE 網頁圖卡圖片', // 上傳的圖片存在雲端硬碟的這個資料夾
};

const SH = { CARDS: '圖卡', PUBLISHERS: '發布者' };

const HEADERS = {
  '圖卡': ['圖卡代碼', '標題', '網址', '說明', '圖片網址', '按鈕文字', '安裝按鈕', '色調', '網站圖示', '建立者', '建立時間', '發送次數', '最後發送', '外部瀏覽器'],
  '發布者': ['LINE userId', '名稱', '備註'],
};

// 「圖卡」工作表的欄位位置（從 0 開始）
const C = { ID: 0, TITLE: 1, URL: 2, DESC: 3, IMAGE: 4, BUTTON: 5, INSTALL: 6, THEME: 7, ICON: 8, CREATOR: 9, CREATED: 10, SENT: 11, LAST: 12, EXTERNAL: 13 };

/** 建立工作表並補齊標題列；更新程式後再執行一次也安全（第一次會要求雲端硬碟與連外權限） */
function setup() {
  const ss = book();
  Object.keys(HEADERS).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const h = HEADERS[name];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const first = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  sessionSheet();
  imageFolder();
  Logger.log('試算表：' + ss.getUrl());
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); // 先取得連外權限（讀取網頁資訊用）
}

function doGet() {
  return json({ ok: true, msg: '網頁圖卡平台運作中' });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.action === 'card') {
      const c = getCard(req.cardId);
      out = { ok: true, card: c ? publicCard(c) : null };
    } else {
      const user = req.session && req.action !== 'session' ? userFromSession(req.session) : verifyIdToken(req.idToken);
      const own = () => requireOwner(user, req.cardId);
      switch (req.action) {
        case 'session':      out = { ok: true, ...createSession(user) }; break;
        case 'logout':       deleteSession(req.session); out = { ok: true }; break;
        case 'me':           out = { ok: true, userId: user.userId, displayName: user.displayName, isAdmin: isPublisher(user.userId) }; break;
        case 'fetchMeta':    out = { ok: true, meta: fetchMeta(req.url) }; break;
        case 'uploadImage':  out = { ok: true, url: uploadImage(user, req.data, req.mime) }; break;
        case 'createCard':   out = { ok: true, card: createCard(user, req.card || {}) }; break;
        case 'updateCard':   own(); out = { ok: true, card: updateCard(req.cardId, req.card || {}) }; break;
        case 'markSent':     own(); out = { ok: true, card: markSent(req.cardId) }; break;
        case 'listCards':    out = { ok: true, cards: listCards(user) }; break;
        case 'deleteCards':  out = { ok: true, removed: deleteCards(user, req.cardIds || []) }; break;
        case 'myData':       out = { ok: true, ...myData(user) }; break;
        case 'deleteMyData': out = { ok: true, ...deleteMyData(user) }; break;
        default: throw new Error('未知的動作');
      }
    }
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return json(out);
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 平台通行證：登入一次，60 天內不用再登入 ---------- */
const SESSION_SHEET = '登入';
const SESSION_DAYS = 60;

function sessionSheet() {
  const ss = book();
  let sh = ss.getSheetByName(SESSION_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SESSION_SHEET);
    sh.appendRow(['通行證', 'LINE userId', 'LINE名稱', '到期', '建立']);
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function createSession(user) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const now = new Date(), exp = new Date(now.getTime() + SESSION_DAYS * 864e5);
  withLock(() => {
    const sh = sessionSheet();
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) if (data[i][3] instanceof Date && data[i][3] < now) sh.deleteRow(i + 1);
    sh.appendRow([token, user.userId, user.displayName, exp, now]);
  });
  return { token, expires: exp.getTime() };
}

function userFromSession(token) {
  const data = sessionSheet().getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token && data[i][3] instanceof Date && data[i][3] > now) return { userId: String(data[i][1]), displayName: String(data[i][2]) };
  }
  throw new Error('SESSION_EXPIRED');
}

function deleteSession(token) {
  if (!token) return;
  withLock(() => {
    const sh = sessionSheet(), data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) if (data[i][0] === token) sh.deleteRow(i + 1);
  });
}

/** 向 LINE 驗證 ID token，防止有人偽造身分 */
function verifyIdToken(idToken) {
  if (!idToken) throw new Error('缺少登入資訊，請在 LINE 中重新開啟');
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: { id_token: idToken, client_id: CONFIG.CHANNEL_ID },
    muteHttpExceptions: true,
  });
  const d = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) {
    const m = d.error_description || '驗證失敗';
    throw new Error(/expired/i.test(m) ? 'TOKEN_EXPIRED' : 'LINE 身分驗證失敗：' + m);
  }
  return { userId: d.sub, displayName: d.name || '' };
}

/* ---------- 工具 ---------- */
function sheet(name) { return book().getSheetByName(name); }
function rows(name) { const v = sheet(name).getDataRange().getValues(); v.shift(); return v; }
function tz() { return Session.getScriptTimeZone(); }
function isTrue(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function fmt(d, pattern) { return Utilities.formatDate(d, tz(), pattern || 'yyyy/MM/dd HH:mm'); }
function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function isPublisher(uid) { return rows(SH.PUBLISHERS).some(r => String(r[0]).trim() === uid); }
function creatorOf(row) { const m = /\((U[0-9a-f]+)\)\s*$/.exec(String(row[C.CREATOR] || '')); return m ? m[1] : ''; }
function requireOwner(user, id) {
  if (isPublisher(user.userId)) return;
  if (creatorOf(findCardRow(id).row) !== user.userId) throw new Error('只有圖卡建立者可以管理這張圖卡');
}

/* ---------- 讀取網頁資訊（標題、說明、預覽圖、網站圖示） ---------- */
function cleanUrl(u) {
  u = String(u || '').trim();
  if (!/^https?:\/\//i.test(u)) u = 'https://' + u.replace(/^\/+/, '');
  if (!/^https?:\/\/[^\s/]+\.[^\s/]+/i.test(u) || /\s/.test(u)) throw new Error('網址格式不對，請貼完整的網址');
  return u.slice(0, 900);
}

function decodeEntities(s) {
  return String(s || '').replace(/&(#x?[0-9a-f]+|amp|lt|gt|quot|apos|nbsp);/gi, (m, k) => {
    const map = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };
    if (map[k.toLowerCase()]) return map[k.toLowerCase()];
    const n = k[1] === 'x' || k[1] === 'X' ? parseInt(k.slice(2), 16) : parseInt(k.slice(1), 10);
    return isNaN(n) ? m : String.fromCodePoint(n);
  }).replace(/\s+/g, ' ').trim();
}

/** 把相對網址換成完整網址 */
function absUrl(href, base) {
  href = decodeEntities(href);
  if (!href) return '';
  if (/^https?:\/\//i.test(href)) return href;
  if (/^\/\//.test(href)) return 'https:' + href;
  const m = /^(https?:\/\/[^/?#]+)([^?#]*)/i.exec(base);
  if (!m) return '';
  if (href.charAt(0) === '/') return m[1] + href;
  const dir = m[2].replace(/[^/]*$/, '') || '/';
  const parts = (dir + href).split('/'), out = [];
  parts.forEach(p => { if (p === '..') out.pop(); else if (p !== '.') out.push(p); });
  return m[1] + out.join('/');
}

function fetchMeta(url) {
  url = cleanUrl(url);
  let res;
  try {
    res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; LineLinkCard/1.0; +https://github.com/isaiahovercomer)', 'Accept-Language': 'zh-TW,zh;q=0.9' } });
  } catch (e) {
    throw new Error('讀不到這個網頁，請確認網址；也可以直接自己填標題和說明');
  }
  const code = res.getResponseCode();
  if (code >= 400) throw new Error('網頁回應錯誤（' + code + '），可以直接自己填標題和說明');
  const type = String(res.getHeaders()['Content-Type'] || res.getHeaders()['content-type'] || '');
  if (type && !/html/i.test(type)) return { url, title: '', desc: '', image: /^image\//i.test(type) && /^https:/i.test(url) ? url : '', icon: '', pwa: false };
  let html = res.getContentText('UTF-8');
  const cs = /<meta[^>]+charset=["']?([\w-]+)/i.exec(html.slice(0, 3000));
  if (cs && !/utf-?8/i.test(cs[1])) { try { html = res.getContentText(cs[1]); } catch (e) {} }
  const head = (/<head[\s\S]*?<\/head>/i.exec(html) || [html.slice(0, 200000)])[0];

  const tags = head.match(/<(meta|link)\b[^>]*>/gi) || [];
  const attr = (tag, name) => { const m = new RegExp('\\b' + name + '\\s*=\\s*("([^"]*)"|\'([^\']*)\'|([^\\s>]+))', 'i').exec(tag); if (!m) return ''; return m[2] !== undefined ? m[2] : m[3] !== undefined ? m[3] : (m[4] || ''); };
  const meta = key => { const t = tags.find(t => /^<meta/i.test(t) && [attr(t, 'property'), attr(t, 'name')].some(v => v.toLowerCase() === key)); return t ? decodeEntities(attr(t, 'content')) : ''; };
  const links = rel => tags.filter(t => /^<link/i.test(t) && attr(t, 'rel').toLowerCase().split(/\s+/).includes(rel));

  const titleTag = decodeEntities((/<title[^>]*>([\s\S]*?)<\/title>/i.exec(head) || [])[1]);
  let image = absUrl(meta('og:image') || meta('og:image:url') || meta('twitter:image') || meta('twitter:image:src'), url);
  if (!/^https:/i.test(image)) image = ''; // LINE 圖卡只能用 https 圖片
  const iconTag = links('apple-touch-icon')[0] || links('icon').sort((a, b) => (parseInt(attr(b, 'sizes')) || 0) - (parseInt(attr(a, 'sizes')) || 0))[0];
  let icon = iconTag ? absUrl(attr(iconTag, 'href'), url) : '';
  if (!/^https:/i.test(icon) || /\.ico(\?|$)/i.test(icon)) icon = '';
  return {
    url,
    title: (meta('og:title') || meta('twitter:title') || titleTag).slice(0, 80),
    desc: (meta('og:description') || meta('description') || meta('twitter:description')).slice(0, 200),
    image,
    icon,
    pwa: links('manifest').length > 0,
  };
}

/* ---------- 圖片上傳 ---------- */
function imageFolder() {
  const it = DriveApp.getFoldersByName(CONFIG.IMAGE_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.IMAGE_FOLDER);
}

function uploadImage(user, data, mime) {
  mime = /^image\/(jpeg|png)$/.test(mime) ? mime : 'image/jpeg';
  const b64 = String(data || '').replace(/^data:[^,]+,/, '');
  if (!b64) throw new Error('沒有收到圖片');
  if (b64.length > 7000000) throw new Error('圖片太大，請換一張小一點的');
  const name = 'card-' + fmt(new Date(), 'yyyyMMdd-HHmmss') + (mime === 'image/png' ? '.png' : '.jpg');
  const file = imageFolder().createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime, name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // LINE 圖卡要能公開讀取
  file.setDescription('上傳者：' + user.displayName);
  return 'https://lh3.googleusercontent.com/d/' + file.getId();
}

/* ---------- 圖卡 ---------- */
const THEMES = ['classic', 'feast', 'prayer', 'gospel', 'youth', 'kids', 'study', 'service'];

function rowToCard(r) {
  return {
    id: String(r[C.ID]), title: String(r[C.TITLE]), url: String(r[C.URL]), desc: String(r[C.DESC] || ''),
    image: String(r[C.IMAGE] || ''), button: String(r[C.BUTTON] || '') || '開啟網頁',
    install: r[C.INSTALL] === '' ? true : isTrue(r[C.INSTALL]),
    external: r[C.EXTERNAL] === '' ? true : isTrue(r[C.EXTERNAL]),
    theme: THEMES.includes(String(r[C.THEME])) ? String(r[C.THEME]) : 'classic',
    icon: String(r[C.ICON] || ''),
    creator: String(r[C.CREATOR] || '').replace(/\s*\(U[0-9a-f]+\)\s*$/, ''),
    created: r[C.CREATED] instanceof Date ? fmt(r[C.CREATED], 'yyyy/M/d HH:mm') : String(r[C.CREATED] || ''),
    sent: Number(r[C.SENT]) || 0,
    last: r[C.LAST] instanceof Date ? fmt(r[C.LAST], 'yyyy/M/d HH:mm') : String(r[C.LAST] || ''),
  };
}
/** 給轉傳頁用：不帶建立者 */
function publicCard(c) { const o = Object.assign({}, c); delete o.creator; return o; }

function cleanCard(c) {
  const httpsOrEmpty = v => { v = String(v || '').trim(); return /^https:\/\/\S+$/i.test(v) ? v.slice(0, 1500) : ''; };
  const title = String(c.title || '').trim().slice(0, 80);
  if (!title) throw new Error('請填標題');
  return {
    title, url: cleanUrl(c.url), desc: String(c.desc || '').trim().slice(0, 300),
    image: httpsOrEmpty(c.image), button: String(c.button || '').trim().slice(0, 20) || '開啟網頁',
    install: c.install !== false, external: c.external !== false,
    theme: THEMES.includes(c.theme) ? c.theme : 'classic', icon: httpsOrEmpty(c.icon),
  };
}

function getCard(id) {
  const r = rows(SH.CARDS).find(r => String(r[C.ID]) === String(id));
  return r ? rowToCard(r) : null;
}

function findCardRow(id) {
  const data = sheet(SH.CARDS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (String(data[i][C.ID]) === String(id)) return { index: i + 1, row: data[i] };
  throw new Error('找不到這張圖卡，可能已經被刪除');
}

function createCard(user, input) {
  const c = cleanCard(input);
  const id = 'L' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  const row = [];
  row[C.ID] = id; row[C.TITLE] = c.title; row[C.URL] = c.url; row[C.DESC] = c.desc; row[C.IMAGE] = c.image;
  row[C.BUTTON] = c.button; row[C.INSTALL] = c.install; row[C.THEME] = c.theme; row[C.ICON] = c.icon;
  row[C.CREATOR] = user.displayName + ' (' + user.userId + ')'; row[C.CREATED] = new Date(); row[C.SENT] = 0; row[C.LAST] = '';
  row[C.EXTERNAL] = c.external;
  withLock(() => sheet(SH.CARDS).appendRow(row));
  return getCard(id);
}

function updateCard(id, input) {
  const c = cleanCard(input);
  withLock(() => {
    const { index } = findCardRow(id);
    const sh = sheet(SH.CARDS);
    sh.getRange(index, C.TITLE + 1, 1, 8).setValues([[c.title, c.url, c.desc, c.image, c.button, c.install, c.theme, c.icon]]);
    sh.getRange(index, C.EXTERNAL + 1).setValue(c.external);
  });
  return getCard(id);
}

function markSent(id) {
  withLock(() => {
    const { index, row } = findCardRow(id);
    sheet(SH.CARDS).getRange(index, C.SENT + 1, 1, 2).setValues([[(Number(row[C.SENT]) || 0) + 1, new Date()]]);
  });
  return getCard(id);
}

/** 自己建立的圖卡；管理員看得到全部。新的在前面 */
function listCards(user) {
  const admin = isPublisher(user.userId);
  return rows(SH.CARDS)
    .filter(r => r[C.ID] && (admin || creatorOf(r) === user.userId))
    .map(r => Object.assign(rowToCard(r), { mine: creatorOf(r) === user.userId }))
    .reverse();
}

function deleteCards(user, ids) {
  const want = new Set(ids.map(String));
  const admin = isPublisher(user.userId);
  return withLock(() => {
    const sh = sheet(SH.CARDS), data = sh.getDataRange().getValues();
    let n = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (!want.has(String(data[i][C.ID]))) continue;
      if (!admin && creatorOf(data[i]) !== user.userId) throw new Error('只能刪除自己建立的圖卡');
      sh.deleteRow(i + 1); n++;
    }
    return n;
  });
}

/* ---------- 我的個人資料 ---------- */
function myData(user) {
  const mine = rows(SH.CARDS).filter(r => creatorOf(r) === user.userId).map(rowToCard);
  return { found: mine.length > 0, lineName: user.displayName, cards: mine.map(c => ({ title: c.title, time: c.created })) };
}

/** 刪除自己建立的圖卡紀錄與登入通行證 */
function deleteMyData(user) {
  return withLock(() => {
    const del = (name, test) => {
      const sh = sheet(name), data = sh.getDataRange().getValues();
      let n = 0;
      for (let i = data.length - 1; i >= 1; i--) if (test(data[i])) { sh.deleteRow(i + 1); n++; }
      return n;
    };
    const cards = del(SH.CARDS, r => creatorOf(r) === user.userId);
    sessionSheet(); del(SESSION_SHEET, r => String(r[1]) === user.userId);
    return { cards };
  });
}
