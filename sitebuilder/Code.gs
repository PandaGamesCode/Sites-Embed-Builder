/*******************************************************************
 * SINGLE SITE — Full single-page website builder (Google Apps Script)
 * Files: Code.gs (backend) + index.html (UI, served via IFRAME)
 *
 * DEPLOYMENT (one-time):
 *  1. Create a Google Sheet. Extensions ▸ Apps Script. Paste these 2 files.
 *  2. Deploy ▸ New deployment ▸ Web app
 *       Execute as: Me  |  Who has access: Anyone in <your department>
 *  3. Open the web-app URL. First run authorizes scopes; all sheets/tabs
 *     (Users, Projects, Pages, Versions, Media, Assets, SiteData) auto-create.
 *
 * AUTH:
 *  - "Sign in with Google": client-side Google Identity Services obtains the
 *    account email, backend verifies it and opens an 8-hour session.
 *    Restrict allowed domains by setting script property ALLOWED_DOMAINS
 *    (comma separated, e.g. "student.dodea.edu,yourcompany.com"). Empty = allow all.
 *  - Username + password fallback for every user (hashed w/ PBKDF2-style salt).
 *  - Everyone is equal — no roles. All data is stored under the user's email.
 *******************************************************************/

var SHEET_NAMES = ['Users', 'Projects', 'Pages', 'Versions', 'Media', 'Assets', 'SiteData'];
var HEADERS = {
  Users:    ['email', 'name', 'picture', 'passHash', 'salt', 'createdAt', 'lastLogin'],
  Projects: ['id', 'email', 'name', 'themeJson', 'mode', 'published', 'createdAt', 'updatedAt'],
  Pages:    ['id', 'projectId', 'email', 'name', 'sortOrder', 'dataJson', 'updatedAt'],
  Versions: ['id', 'pageId', 'projectId', 'email', 'label', 'dataJson', 'createdAt'],
  Media:    ['id', 'email', 'kind', 'mimeType', 'size', 'b64', 'createdAt'],   // small images (<35k chars b64)
  Assets:   ['id', 'email', 'name', 'kind', 'mimeType', 'size', 'driveUrl', 'fileId', 'createdAt'], // Drive files
  SiteData: ['projectId', 'key', 'valueJson', 'updatedAt']                    // mini-database for live exports
};
var SESSION_HOURS = 8;
var MAX_VERSIONS_PER_PAGE = 20;

/* ------------------------------ bootstrap ------------------------------ */

function doGet(e) {
  ensureSetup_();
  if (e && e.parameter && e.parameter.action) {
    var res = doGetData(e);
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Single Site')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

function ensureSetup_() {
  var ss = getSS_();
  SHEET_NAMES.forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (!sh) sh = ss.insertSheet(n);
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS[n]);
  });
  // tidy default sheet
  var gd = ss.getSheetByName('Sheet1'); if (gd && ss.getSheets().length > 1) ss.deleteSheet(gd);
  SpreadsheetApp.flush();
}
var _ss = null;
function getSS_() {
  if (_ss) return _ss;
  var p = PropertiesService.getScriptProperties();
  if (p.SS_ID) { _ss = SpreadsheetApp.openById(p.SS_ID); return _ss; }
  _ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!_ss) { // standalone: create + remember
    _ss = SpreadsheetApp.create('Single Site Data');
    PropertiesService.getScriptProperties().setProperty('SS_ID', _ss.getId());
  }
  return _ss;
}
function sheet_(name) { return getSS_().getSheetByName(name) || (ensureSetup_(), getSS_().getSheetByName(name)); }
function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12) + Date.now().toString(36); }
function now_() { return new Date().toISOString(); }

/* ------------------------------ auth ------------------------------ */

function hashPw_(pw, salt) {
  var raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + '::' + pw + '::' + salt,
    Utilities.Charset.UTF_8
  );
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function googleLogin(email, name, picture) {
  ensureSetup_();
  email = String(email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Invalid Google account.');
  var allow = PropertiesService.getScriptProperties().getProperty('ALLOWED_DOMAINS');
  if (allow) {
    var doms = allow.split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(String);
    var dom = email.split('@')[1];
    if (doms.indexOf(dom) === -1) throw new Error('Only accounts from: ' + doms.join(', '));
  }
  var row = findUserRow_(email);
  if (!row) createUserRow_(email, name || email.split('@')[0], picture || '', null, null);
  else updateLastLogin_(row.rowNum, picture);
  return makeSession_(email);
}

function registerUser(username, password, displayName) {
  ensureSetup_();
  username = String(username || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,32}$/.test(username)) throw new Error('Username: 3–32 chars, letters/numbers/._- only.');
  if (!password || String(password).length < 6) throw new Error('Password must be at least 6 characters.');
  var email = username + '@single.site';
  if (findUserRow_(email)) throw new Error('Username already taken.');
  var salt = uid_() + uid_();
  createUserRow_(email, displayName || username, '', hashPw_(password, salt), salt);
  return makeSession_(email);
}

function passwordLogin(username, password) {
  ensureSetup_();
  username = String(username || '').trim().toLowerCase();
  var email = username.indexOf('@') > -1 ? username : username + '@single.site';
  var row = findUserRow_(email);
  if (!row || !row.values[3]) throw new Error('No password account found. Use Google sign-in or register.');
  if (hashPw_(password, row.values[4]) !== row.values[3]) throw new Error('Incorrect password.');
  updateLastLogin_(row.rowNum);
  return makeSession_(email);
}

function findUserRow_(email) {
  var sh = sheet_('Users'), last = sh.getLastRow();
  if (last < 2) return null;
  var vals = sh.getRange(2, 1, last - 1, 5).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]).toLowerCase() === email.toLowerCase())
    return { rowNum: i + 2, values: vals[i] };
  return null;
}
function createUserRow_(email, name, picture, passHash, salt) {
  sheet_('Users').appendRow([email, name, picture, passHash || '', salt || '', now_(), now_()]);
}
function updateLastLogin_(rowNum, picture) {
  var sh = sheet_('Users');
  sh.getRange(rowNum, 7).setValue(now_());
  if (picture) sh.getRange(rowNum, 3).setValue(picture);
}
function makeSession_(email) {
  var token = uid_() + uid_();
  PropertiesService.getScriptProperties().setProperty('sess_' + token,
    JSON.stringify({ email: email, exp: Date.now() + SESSION_HOURS * 3600 * 1000 }));
  cleanupSessions_();
  var u = findUserRow_(email);
  return { token: token, user: { email: email, name: u ? u.values[1] : email, picture: u ? u.values[2] : '' } };
}
function logout(token) { if (token) PropertiesService.getScriptProperties().deleteProperty('sess_' + token); return true; }
function cleanupSessions_() {
  var p = PropertiesService.getScriptProperties(), all = p.getProperties(), n = 0;
  for (var k in all) if (k.indexOf('sess_') === 0) {
    try { if (JSON.parse(all[k]).exp < Date.now()) { p.deleteProperty(k); n++; } } catch (e) { p.deleteProperty(k); }
  }
  return n;
}
function auth_(token) {
  if (!token) throw new Error('SESSION_EXPIRED');
  var s = PropertiesService.getScriptProperties().getProperty('sess_' + token);
  if (!s) throw new Error('SESSION_EXPIRED');
  var o = JSON.parse(s);
  if (o.exp < Date.now()) { PropertiesService.getScriptProperties().deleteProperty('sess_' + token); throw new Error('SESSION_EXPIRED'); }
  return o.email;
}

function checkSession(token) {
  try { var email = auth_(token); var u = findUserRow_(email);
    return { ok: true, user: { email: email, name: u ? u.values[1] : email, picture: u ? u.values[2] : '' } };
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ------------------------------ projects ------------------------------ */

function listProjects(token) {
  var email = auth_(token);
  var v = sheet_('Projects').getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) if (String(v[i][1]).toLowerCase() === email.toLowerCase()) {
    out.push({ id: v[i][0], name: v[i][2], mode: v[i][3], published: v[i][4] === true || v[i][4] === 'TRUE',
      updatedAt: v[i][7] ? String(v[i][7]) : '' });
  }
  out.sort(function (a, b) { return (b.updatedAt || '').localeCompare(a.updatedAt || ''); });
  return out;
}

function createProject(token, name, mode) {
  var email = auth_(token);
  var id = uid_();
  sheet_('Projects').appendRow([id, email, name || 'Untitled site',
    JSON.stringify(defaultTheme_()), mode || 'simple', false, now_(), now_()]);
  var pid = uid_();
  sheet_('Pages').appendRow([pid, id, email, 'Home', 1, JSON.stringify({ blocks: [defaultHeroBlock_()] }), now_()]);
  return { projectId: id, pageId: pid };
}

function renameProject(token, projectId, name) {
  var email = auth_(token); var r = projectRow_('Projects', projectId, email);
  r.sh.getRange(r.row, 3).setValue(name); r.sh.getRange(r.row, 8).setValue(now_()); return true;
}
function setProjectMode(token, projectId, mode) {
  var email = auth_(token); var r = projectRow_('Projects', projectId, email);
  r.sh.getRange(r.row, 4).setValue(mode); return true;
}
function deleteProject(token, projectId) {
  var email = auth_(token);
  deleteRowsWhere_('Pages', function (v) { return v[1] === projectId && v[2].toLowerCase() === email.toLowerCase(); });
  deleteRowsWhere_('Versions', function (v) { return v[2] === projectId && v[3].toLowerCase() === email.toLowerCase(); });
  deleteRowsWhere_('SiteData', function (v) { return v[0] === projectId; });
  deleteRowsWhere_('Projects', function (v) { return v[0] === projectId && v[1].toLowerCase() === email.toLowerCase(); });
  return true;
}

function projectRow_(sheetName, id, email) {
  var sh = sheet_(sheetName), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    var idCol = sheetName === 'Projects' ? 0 : (sheetName === 'Pages' ? 0 : 1);
    var emCol = sheetName === 'Projects' ? 1 : (sheetName === 'Pages' ? 2 : 3);
    if (v[i][idCol] === id && (!email || String(v[i][emCol]).toLowerCase() === email.toLowerCase()))
      return { sh: sh, row: i + 1, values: v[i] };
  }
  throw new Error('Not found: ' + id);
}
function deleteRowsWhere_(name, pred) {
  var sh = sheet_(name), v = sh.getDataRange().getValues(), del = [];
  for (var i = 1; i < v.length; i++) if (pred(v[i])) del.push(i + 1);
  for (var j = del.length - 1; j >= 0; j--) sh.deleteRow(del[j]);
}

/* ------------------------------ pages ------------------------------ */

function listPages(token, projectId) {
  var email = auth_(token);
  var v = sheet_('Pages').getDataRange().getValues(), out = [];
  for (var i = 1; i < v.length; i++) if (v[i][1] === projectId && String(v[i][2]).toLowerCase() === email.toLowerCase())
    out.push({ id: v[i][0], name: v[i][3], sortOrder: v[i][4], updatedAt: v[i][6] });
  out.sort(function (a, b) { return a.sortOrder - b.sortOrder; });
  return out;
}
function getPage(token, pageId) {
  var email = auth_(token);
  var r = projectRow_('Pages', pageId, email);
  return { id: r.values[0], projectId: r.values[1], name: r.values[3], sortOrder: r.values[4],
    data: parseJson_(r.values[5], { blocks: [] }), theme: parseJson_(themeForProject_(r.values[1]), defaultTheme_()) };
}
function savePage(token, pageId, blocks, opts) {
  opts = opts || {};
  var email = auth_(token);
  var r = projectRow_('Pages', pageId, email);
  var json = JSON.stringify(blocks);
  r.sh.getRange(r.row, 6).setValue(json);
  r.sh.getRange(r.row, 7).setValue(now_());
  sheet_('Projects').getRange(projectRow_('Projects', r.values[1], email).row, 8).setValue(now_());
  if (opts.versionLabel) addVersion(token, pageId, json, opts.versionLabel);
  return true;
}
function createPage(token, projectId, name) {
  var email = auth_(token);
  var v = sheet_('Pages').getDataRange().getValues(); var max = 0;
  for (var i = 1; i < v.length; i++) if (v[i][1] === projectId && v[i][4] > max) max = v[i][4];
  var id = uid_();
  sheet_('Pages').appendRow([id, projectId, email, name || ('Page ' + (max + 1)), max + 1,
    JSON.stringify({ blocks: [] }), now_()]);
  return { pageId: id };
}
function renamePage(token, pageId, name) {
  var email = auth_(token); var r = projectRow_('Pages', pageId, email);
  r.sh.getRange(r.row, 4).setValue(name); return true;
}
function deletePage(token, pageId) {
  var email = auth_(token); var r = projectRow_('Pages', pageId, email);
  if (listPages(token, r.values[1]).length <= 1) throw new Error('A project needs at least one page.');
  deleteRowsWhere_('Versions', function (v) { return v[1] === pageId; });
  r.sh.deleteRow(r.row); return true;
}
function movePage(token, pageId, dir) { // dir: -1 up, +1 down
  var email = auth_(token); var r = projectRow_('Pages', pageId, email);
  var pages = listPages(token, r.values[1]);
  var idx = -1; pages.forEach(function (p, i) { if (p.id === pageId) idx = i; });
  var j = idx + dir; if (j < 0 || j >= pages.length) return true;
  var sh = sheet_('Pages');
  sh.getRange(projectRow_('Pages', pages[idx].id, email).row, 5).setValue(j + 1);
  sh.getRange(projectRow_('Pages', pages[j].id, email).row, 5).setValue(idx + 1);
  return true;
}
function saveTheme(token, projectId, themeJson) {
  var email = auth_(token); var r = projectRow_('Projects', projectId, email);
  r.sh.getRange(r.row, 4).setValue(JSON.stringify(themeJson));
  r.sh.getRange(r.row, 8).setValue(now_());
  return true;
}
function themeForProject_(projectId) {
  var v = sheet_('Projects').getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === projectId) return v[i][3];
  return null;
}

/* ------------------------------ versions ------------------------------ */

function addVersion(token, pageId, dataJson, label) {
  var email = auth_(token);
  var r = projectRow_('Pages', pageId, email);
  sheet_('Versions').appendRow([uid_(), pageId, r.values[1], email, label || 'Save', dataJson, now_()]);
  // trim old
  var v = sheet_('Versions').getDataRange().getValues(), mine = [];
  for (var i = 1; i < v.length; i++) if (v[i][1] === pageId) mine.push({ row: i + 1, created: v[i][6] });
  if (mine.length > MAX_VERSIONS_PER_PAGE) {
    mine.slice(0, mine.length - MAX_VERSIONS_PER_PAGE).forEach(function (m) { sheet_('Versions').deleteRow(m.row); });
  }
  return true;
}
function listVersions(token, pageId) {
  var email = auth_(token);
  projectRow_('Pages', pageId, email);
  var v = sheet_('Versions').getDataRange().getValues(), out = [];
  for (var j = 1; j < v.length; j++) if (v[j][1] === pageId) out.push({ id: v[j][0], label: v[j][4], createdAt: String(v[j][6]), dataJson: v[j][5] });
  out.reverse();
  return out;
}
function restoreVersion(token, versionId) {
  var email = auth_(token);
  var v = sheet_('Versions').getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === versionId && String(v[i][3]).toLowerCase() === email.toLowerCase()) {
    var pr = projectRow_('Pages', v[i][1], email);
    addVersion(token, v[i][1], pr.values[5], 'Before restore');
    pr.sh.getRange(pr.row, 6).setValue(v[i][5]);
    pr.sh.getRange(pr.row, 7).setValue(now_());
    return true;
  }
  throw new Error('Version not found.');
}

/* ------------------------------ media / assets ------------------------------ */

function uploadImage(token, name, mimeType, b64) {
  var email = auth_(token);
  var bytes = Utilities.base64Decode(b64);
  if (b64.length < 35000) { // small → inline in sheet for instant serving
    sheet_('Media').appendRow([uid_(), email, 'image', mimeType, bytes.length, b64, now_()]);
    return { kind: 'inline', ref: name };
  }
  var folder = driveFolder_(email);
  var file = DriveApp.createFile(Utilities.newBlob(bytes, mimeType, name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  sheet_('Assets').appendRow([uid_(), email, name, 'image', mimeType, bytes.length,
    'https://lh3.googleusercontent.com/d/' + file.getId(), file.getId(), now_()]);
  return { kind: 'drive', url: 'https://lh3.googleusercontent.com/d/' + file.getId() };
}
function getImageInline(token, name) {
  var email = auth_(token);
  var v = sheet_('Media').getDataRange().getValues();
  for (var i = v.length - 1; i > 0; i--) if (String(v[i][1]).toLowerCase() === email.toLowerCase() && v[i][2] === name)
    return { mimeType: v[i][3], b64: v[i][5] };
  throw new Error('Image not found');
}
function listImages(token) {
  var email = auth_(token), out = [];
  sheet_('Media').getDataRange().getValues().forEach(function (r, i) {
    if (i && String(r[1]).toLowerCase() === email.toLowerCase()) out.push({ name: r[2], mime: r[3], size: r[4], inline: true });
  });
  sheet_('Assets').getDataRange().getValues().forEach(function (r, i) {
    if (i && String(r[1]).toLowerCase() === email.toLowerCase()) out.push({ name: r[2], mime: r[4], size: r[5], inline: false, url: r[6] });
  });
  return out;
}
function deleteImage(token, name) {
  var email = auth_(token);
  deleteRowsWhere_('Media', function (v) { return v[1] === name && String(v[2]).toLowerCase() === email.toLowerCase(); });
  deleteRowsWhere_('Assets', function (v) { return v[2] === name && String(v[1]).toLowerCase() === email.toLowerCase(); });
  return true;
}
function driveFolder_(email) {
  var p = PropertiesService.getScriptProperties(); var id = p.getProperty('FOLDER_' + email);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('Single Site ▸ ' + email);
  p.setProperty('FOLDER_' + email, f.getId()); return f;
}

/* ------------------------------ site data (mini DB for live exports) ------------------------------ */

function saveSiteData(token, projectId, key, valueJson) {
  var email = auth_(token); projectRow_('Projects', projectId, email);
  var sh = sheet_('SiteData'), v = sh.getDataRange().getValues(), found = 0;
  for (var i = 1; i < v.length; i++) if (v[i][0] === projectId && v[i][1] === key) { found = i + 1; break; }
  var cell = typeof valueJson === 'string' ? valueJson : JSON.stringify(valueJson);
  if (found) { sh.getRange(found, 3).setValue(cell); sh.getRange(found, 4).setValue(now_()); }
  else sh.appendRow([projectId, key, cell, now_()]);
  return true;
}
function loadSiteData(token, projectId) {
  var email = auth_(token); projectRow_('Projects', projectId, email);
  var v = sheet_('SiteData').getDataRange().getValues(), out = {};
  for (var i = 1; i < v.length; i++) if (v[i][0] === projectId) {
    try { out[v[i][1]] = JSON.parse(v[i][2]); } catch (e) {}
  }
  return out;
}
// public read used by LIVE exported html (no session — data is public by design)
function fetchSiteDataPublic(projectId, key) {
  var v = sheet_('SiteData').getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === projectId && (v[i][1] === key || !key)) {
    try { return { ok: true, key: v[i][1], value: JSON.parse(v[i][2]), updated: String(v[i][3]) }; } catch (e) {}
  }
  return { ok: false };
}

/* ------------------------------ export ------------------------------ */

function setPublished(token, projectId, on) {
  var email = auth_(token); var r = projectRow_('Projects', projectId, email);
  r.sh.getRange(r.row, 6).setValue(!!on); return true;
}
function exportPage(token, pageId, opts) {
  opts = opts || {};
  var email = auth_(token);
  var page = getPage(token, pageId);
  if (opts.blocks) page.data = { blocks: opts.blocks }; // export live editor state
  var pj = projectRow_('Projects', page.projectId, email);
  var mode = pj.values[3] || 'simple';
  var siteUrl = ScriptApp.getService().getUrl();
  return buildExportHtml_(page.data.blocks, page.theme, {
    mode: opts.mode || mode,
    live: !!opts.live,
    projectId: page.projectId,
    siteUrl: siteUrl,
    pageName: page.name,
    staticDate: opts.staticDate || null
  });
}

/* ============================ HTML EXPORT ENGINE ============================ */

function defaultTheme_() {
  return {
    font: "'Segoe UI', system-ui, -apple-system, sans-serif",
    headingFont: "'Segoe UI', system-ui, sans-serif",
    primary: '#4f46e5', text: '#1f2937', bg: '#ffffff', accent: '#f59e0b',
    radius: 14, maxWidth: 1080, sectionPad: 56, dark: false,
    anim: 'fade', animDuration: 0.6, shadow: 'soft'
  };
}
function defaultHeroBlock_() {
  return { type: 'hero', id: uid_(), props: {
    headline: 'Welcome to ' + 'Single Site', sub: 'Edit this text, drop in sections, and export a single HTML file.',
    image: '', imagePos: 'right', overlay: 35, height: 'md', align: 'left',
    btnText: 'Learn More', btnLink: '#section-1', btn2Text: '', btn2Link: '' } };
}
function parseJson_(s, fb) { try { return typeof s === 'string' ? JSON.parse(s) : (s || fb); } catch (e) { return fb; } }
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function nl2p_(s) { return String(s || '').split(/\n+/).map(function (t) { return '<p>' + esc_(t) + '</p>'; }).join(''); }

function buildExportHtml_(blocks, theme, cfg) {
  var adv = cfg.mode === 'advanced';
  var T = theme || defaultTheme_();
  var cssVars = '--pri:' + T.primary + ';--txt:' + (T.dark ? '#e5e7eb' : T.text) + ';--bg:' + (T.dark ? '#0f172a' : T.bg) +
    ';--acc:' + T.accent + ';--rad:' + T.radius + 'px;--mw:' + T.maxWidth + 'px;--pad:' + T.sectionPad + 'px;' +
    '--ff:' + T.font + ';--fh:' + T.headingFont + ';';
  var body = blocksToHtml_(blocks, cfg, adv);
  var extraHead = adv ? (T.extraHead || '') : '';
  var extraFoot = adv ? (T.extraFoot || '') : '';
  var runtime = buildRuntime_(cfg, adv, blocks);
  return '<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">\n' +
    '<title>' + esc_(cfg.pageName || 'Single Site page') + '</title>\n' +
    '<style>' + EXPORT_CSS + '</style>\n' + extraHead + '\n</head>\n<body class="' +
    (T.dark ? 'dark' : '') + ' anim-' + (T.anim || 'fade') + '" style="' + cssVars +
    'background:var(--bg);color:var(--txt);font-family:var(--ff)">\n' +
    '<main class="ss">' + body + '</main>\n' + runtime + '\n' + extraFoot + '\n</body>\n</html>';
}

function blockWrap_(b, inner, cfg, adv) {
  var p = b.props || {};
  var cls = 'blk blk-' + b.type;
  var st = [];
  if (adv) {
    if (p.bg) st.push('background:' + p.bg);
    if (p.textColor) st.push('color:' + p.textColor);
    if (p.padY != null && p.padY !== '') st.push('padding-top:' + p.padY + 'px;padding-bottom:' + p.padY + 'px');
    if (p.padX != null && p.padX !== '') st.push('padding-left:' + p.padX + 'px;padding-right:' + p.padX + 'px');
    if (p.width) st.push('max-width:' + p.width);
    if (p.border) st.push('border:' + p.border);
    if (p.radius != null && p.radius !== '') st.push('border-radius:' + p.radius + 'px');
    if (p.customCss) st.push.apply(st, []);
    if (p.anchor) cls += ' anchor-' + p.anchor;
    if (p.hiddenOn === 'mobile') cls += ' hide-mobile';
  }
  var animD = adv && p.animDelay ? ' data-anim-delay="' + p.animDelay + '"' : '';
  var idAttr = p.anchor ? ' id="' + esc_(p.anchor) + '"' : '';
  return '<section' + idAttr + ' class="' + cls + '" style="' + st.join(';') + '"' + animD + '>' + inner + '</section>';
}

function blocksToHtml_(blocks, cfg, adv) {
  return (blocks || []).map(function (b) { return renderBlock_(b, cfg, adv); }).join('\n');
}

function renderBlock_(b, cfg, adv) {
  var p = b.props || {}, t = b.type, h = '';
  switch (t) {
    case 'heading':
      h = '<h2 class="ss-h">' + esc_(p.text) + '</h2>' + (p.sub ? '<p class="ss-sub">' + esc_(p.sub) + '</p>' : ''); break;
    case 'text':
      h = '<div class="ss-text">' + nl2p_(p.html ? sanitizeLite_(p.text) : p.text) + '</div>'; break;
    case 'button':
      h = '<a class="ss-btn" href="' + esc_(p.link || '#') + '" target="' + (p.newtab ? '_blank' : '_self') + '">' + esc_(p.label) + '</a>'; break;
    case 'buttons':
      h = '<div class="ss-btngroup">' + (p.items || []).map(function (i) {
        return '<a class="ss-btn ' + (i.style === 'ghost' ? 'ghost' : '') + '" href="' + esc_(i.link || '#') + '">' + esc_(i.label) + '</a>';
      }).join('') + '</div>'; break;
    case 'hero':
      h = heroHtml_(p, adv); break;
    case 'image':
      h = '<figure class="ss-imgwrap"><img src="' + esc_(p.src) + '" alt="' + esc_(p.alt || '') + '"' +
        (adv && p.lazy === false ? '' : ' loading="lazy"') + (adv && p.maxw ? ' style="max-width:' + p.maxw + '"' : '') + '>' +
        (p.caption ? '<figcaption>' + esc_(p.caption) + '</figcaption>' : '') + '</figure>'; break;
    case 'gallery':
      h = '<div class="ss-gallery" style="--cols:' + (adv && p.cols ? p.cols : 3) + '">' +
        (p.items || []).map(function (im) {
          return '<img src="' + esc_(im.src) + '" alt="' + esc_(im.alt || '') + '" loading="lazy"' +
            (adv && p.lightbox !== false ? ' data-lightbox="1"' : '') + '>';
        }).join('') + '</div>'; break;
    case 'video':
      h = '<div class="ss-video"><iframe src="' + esc_(embedUrl_(p.url)) + '" frameborder="0" allowfullscreen loading="lazy"></iframe></div>'; break;
    case 'columns':
      h = '<div class="ss-cols" style="--cols:' + (p.count || 2) + ';--gap:' + (adv && p.gap != null ? p.gap : 24) + 'px">' +
        (p.cols || []).map(function (col) { return '<div class="ss-col">' + blocksToHtml_(col.blocks || [], cfg, adv) + '</div>'; }).join('') + '</div>'; break;
    case 'cards':
      h = '<div class="ss-cards" style="--cols:' + (adv && p.cols ? p.cols : 3) + '">' +
        (p.items || []).map(function (c) {
          return '<div class="ss-card">' + (c.icon ? '<div class="ss-cardicon">' + esc_(c.icon) + '</div>' : '') +
            '<h3>' + esc_(c.title) + '</h3><p>' + esc_(c.body) + '</p>' +
            (c.link ? '<a href="' + esc_(c.link) + '">→</a>' : '') + '</div>';
        }).join('') + '</div>'; break;
    case 'list':
      h = '<' + (p.ordered ? 'ol' : 'ul') + ' class="ss-list">' +
        (p.items || []).map(function (x) { return '<li>' + esc_(typeof x === 'string' ? x : x.text) + '</li>'; }).join('') + '</' + (p.ordered ? 'ol' : 'ul') + '>'; break;
    case 'steps':
      h = '<div class="ss-steps">' + (p.items || []).map(function (s, i) {
        return '<div class="ss-step"><span class="ss-num">' + (i + 1) + '</span><div><h3>' + esc_(s.title) + '</h3><p>' + esc_(s.body) + '</p></div></div>';
      }).join('') + '</div>'; break;
    case 'accordion':
      h = '<div class="ss-acc">' + (p.items || []).map(function (a) {
        return '<details><summary>' + esc_(a.q) + '</summary><div class="ss-acc-body">' + esc_(a.a) + '</div></details>';
      }).join('') + '</div>'; break;
    case 'tabs':
      h = '<div class="ss-tabs" data-tabs><div class="ss-tabbar">' + (p.items || []).map(function (tb, i) {
        return '<button class="ss-tab' + (i === 0 ? ' on' : '') + '" data-i="' + i + '">' + esc_(tb.label) + '</button>';
      }).join('') + '</div>' + (p.items || []).map(function (tb, i) {
        return '<div class="ss-tabpane' + (i === 0 ? ' on' : '') + '" data-i="' + i + '">' + esc_(tb.body) + '</div>';
      }).join('') + '</div>'; break;
    case 'table':
      h = '<div class="ss-tablewrap"><table class="ss-table"><thead><tr>' +
        (p.headers || []).map(function (x) { return '<th>' + esc_(x) + '</th>'; }).join('') + '</tr></thead><tbody>' +
        (p.rows || []).map(function (r) { return '<tr>' + r.map(function (x) { return '<td>' + esc_(x) + '</td>'; }).join('') + '</tr>'; }).join('') +
        '</tbody></table></div>'; break;
    case 'pricing':
      h = '<div class="ss-cards">' + (p.plans || []).map(function (pl) {
        return '<div class="ss-card ss-price' + (pl.featured ? ' featured' : '') + '"><h3>' + esc_(pl.name) + '</h3>' +
          '<div class="ss-amt">' + esc_(pl.price) + '<span>' + esc_(pl.period || '') + '</span></div><ul>' +
          (pl.features || []).map(function (f) { return '<li>' + esc_(f) + '</li>'; }).join('') + '</ul>' +
          '<a class="ss-btn" href="' + esc_(pl.link || '#') + '">' + esc_(pl.cta || 'Choose') + '</a></div>';
      }).join('') + '</div>'; break;
    case 'testimonial':
      h = '<div class="ss-quote"><p>“' + esc_(p.quote) + '”</p><footer>' + esc_(p.author) + (p.role ? ' · ' + esc_(p.role) : '') + '</footer></div>'; break;
    case 'cta':
      h = '<div class="ss-cta"><h2>' + esc_(p.title) + '</h2><p>' + esc_(p.sub) + '</p>' +
        '<a class="ss-btn acc" href="' + esc_(p.link || '#') + '">' + esc_(p.btnText || 'Get started') + '</a></div>'; break;
    case 'form':
      h = formHtml_(p, cfg, adv); break;
    case 'divider':
      h = '<hr class="ss-hr">'; break;
    case 'spacer':
      h = '<div style="height:' + (adv && p.height ? p.height : 40) + 'px"></div>'; break;
    case 'countdown':
      h = '<div class="ss-count" data-end="' + esc_(p.until || '') + '" data-msg="' + esc_(p.done || 'Done!') + '"><div><b>--</b><span>d</span></div><div><b>--</b><span>h</span></div><div><b>--</b><span>m</span></div><div><b>--</b><span>s</span></div></div>'; break;
    case 'clock':
      h = '<div class="ss-clock" data-fmt="' + esc_(p.format || 'time-date') + '">--:--</div>'; break;
    case 'marquee':
      h = '<div class="ss-marquee"><div>' + [0, 1].map(function () { return esc_(p.text); }).join(' • ') + ' • </div></div>'; break;
    case 'badge':
      h = '<div class="ss-badges">' + (p.items || []).map(function (x) { return '<span class="ss-badge">' + esc_(x) + '</span>'; }).join('') + '</div>'; break;
    case 'nav':
      h = '<nav class="ss-nav">' + (p.items || []).map(function (x) { return '<a href="' + esc_(x.link) + '">' + esc_(x.label) + '</a>'; }).join('') + '</nav>'; break;
    case 'html':
      h = adv ? String(p.html || '') : '<p class="ss-sub">(Advanced mode required for custom HTML)</p>'; break;
    case 'dynamic':
      h = dynamicHtml_(p, cfg, adv); break;
    case 'embed':
      h = '<div class="ss-video"><iframe src="' + esc_(p.url) + '" style="height:' + (p.height || 400) + 'px" frameborder="0" loading="lazy"></iframe></div>'; break;
    default:
      h = '<p>[unknown block]</p>';
  }
  return blockWrap_(b, h, cfg, adv);
}

function heroHtml_(p, adv) {
  var st = [];
  if (p.image) { st.push('background-image:linear-gradient(rgba(0,0,0,' + ((p.overlay || 0) / 100) + '),rgba(0,0,0,' + ((p.overlay || 0) / 100) + ')),url(\"' + p.image + '\")'); st.push('background-size:cover'); st.push('background-position:center'); }
  var hcls = 'ss-hero h-' + (p.height || 'md') + ' al-' + (p.align || 'left') + (p.image && p.imagePos === 'split' ? ' split' : '');
  var inner = '';
  if (p.image && p.imagePos === 'split') inner += '<img class="ss-heroimg" src="' + esc_(p.image) + '" alt="">';
  inner += '<div class="ss-herotext">' +
    (p.kicker && adv ? '<div class="ss-kicker">' + esc_(p.kicker) + '</div>' : '') +
    '<h1>' + esc_(p.headline) + '</h1>' + (p.sub ? '<p>' + esc_(p.sub) + '</p>' : '') +
    (p.btnText ? '<div class="ss-btngroup"><a class="ss-btn ' + (adv && p.btnStyle === 'ghost' ? 'ghost' : '') + '" href="' + esc_(p.btnLink || '#') + '">' + esc_(p.btnText) + '</a>' +
      (p.btn2Text ? '<a class="ss-btn ghost" href="' + esc_(p.btn2Link || '#') + '">' + esc_(p.btn2Text) + '</a>' : '') + '</div>' : '') +
    '</div>';
  return '<div class="' + hcls + '" style="' + st.join(';') + '">' + inner + '</div>';
}

function formHtml_(p, cfg, adv) {
  var fields = (p.fields || []).map(function (f) {
    var req = f.required ? ' required' : '';
    if (f.type === 'select') return '<label>' + esc_(f.label) + '<select name="' + esc_(f.name) + '"' + req + '>' +
      (f.options || []).map(function (o) { return '<option>' + esc_(o) + '</option>'; }).join('') + '</select></label>';
    if (f.type === 'textarea') return '<label>' + esc_(f.label) + '<textarea name="' + esc_(f.name) + '"' + req + '></textarea></label>';
    return '<label>' + esc_(f.label) + '<input type="' + esc_(f.type || 'text') + '" name="' + esc_(f.name) + '"' + req + '></label>';
  }).join('');
  var mode = adv && p.submitMode ? p.submitMode : 'mailto';
  var tail = '';
  if (mode === 'live' && cfg.siteUrl) tail = '<script>window.__ssFormEndpoint=' + JSON.stringify(cfg.siteUrl) + '?action=form&project=' + JSON.stringify(cfg.projectId) + '&key=' + JSON.stringify((adv && p.formKey) || 'responses') + ';<' + '/script>';
  return '<form class="ss-form" data-mode="' + mode + '" data-to="' + esc_(p.email || '') + '">' + fields +
    '<button class="ss-btn" type="submit">' + esc_(p.btnText || 'Send') + '</button>' +
    '<div class="ss-formmsg" hidden>' + esc_(p.successMsg || 'Thanks! Message sent.') + '</div></form>' + tail;
}

function dynamicHtml_(p, cfg, adv) {
  // The "live tools" block: renders rows from the project mini-database.
  var key = p.dataKey || 'items';
  var bind = p.bind || 'current'; // current | week | day | all | key:<...>
  var tpl = p.template || '{title}';
  var attrs = ' data-key="' + esc_(key) + '" data-bind="' + esc_(bind) + '" data-tpl="' + esc_(tpl) + '"' +
    ' data-mode="' + (cfg.live ? 'live' : 'static') + '"' +
    ' data-static="' + esc_(p.staticValue || '') + '"' +
    (adv && p.sortBy ? ' data-sort="' + esc_(p.sortBy) + '"' : '') +
    (adv && p.filterField ? ' data-filter-field="' + esc_(p.filterField) + '"' : '') +
    (adv && p.showPicker ? ' data-picker="1"' : '') +
    (adv && p.pickerLabel ? ' data-picker-label="' + esc_(p.pickerLabel) + '"' : '') +
    ' data-project="' + esc_(cfg.projectId) + '" data-site-url="' + esc_(cfg.siteUrl || '') + '"';
  var storeAttr = '';
  if (!cfg.live) {
    var store = null;
    try {
      var dv = sheet_('SiteData').getDataRange().getValues();
      for (var di = 1; di < dv.length; di++) if (dv[di][0] === cfg.projectId && dv[di][1] === key) { store = JSON.parse(dv[di][2]); break; }
    } catch (err) {}
    if (store && cfg.staticDate) { store = snapshotUpTo_(store, cfg.staticDate); delete store.weeks; }
    storeAttr = ' data-store="' + esc_(JSON.stringify(store == null ? '' : store)) + '"';
  }
  var picker = adv && p.showPicker ? '<div class="ss-picker"></div>' : '';
  return '<div class="ss-dyn" ' + attrs + storeAttr + '>' + picker + '<div class="ss-dynlist">' +
    (p.placeholderRows || []).map(function (r) { return dynItem_(tpl, r, adv, p); }).join('') +
    (p.empty ? '<p class="ss-empty">' + esc_(p.empty) + '</p>' : '') + '</div></div>';
}
function dynItem_(tpl, row, adv, p) {
  var html = String(tpl).replace(/\{(\w[\w.]*)\}/g, function (m, k) { return esc_(deepGet_(row, k)); });
  if (adv && p.itemClass) return '<div class="' + p.itemClass + '">' + html + '</div>';
  return '<div class="ss-dynitem">' + html + '</div>';
}
function deepGet_(o, path) { var c = o; String(path).split('.').forEach(function (k) { c = c && c[k]; }); return c == null ? '' : c; }
function snapshotUpTo_(store, dateStr) { // freeze a store as it was on/before dateStr (day keys YYYY-MM-DD)
  var out = {}; Object.keys(store || {}).forEach(function (k) {
    if (/^\d{4}-\d{2}-\d{2}$/.test(k)) { if (k <= dateStr) out[k] = store[k]; } else out[k] = store[k];
  });
  out.current = (function () { var days = Object.keys(out).filter(function (k) { return /^\d{4}-\d{2}-\d{2}$/.test(k); }).sort(); return days.length ? days[days.length - 1] : store.current; })();
  return out;
}

function embedUrl_(u) {
  u = String(u || '');
  var m = u.match(/(?:youtu\.be\/|v=)([\w-]{11})/);
  if (m) return 'https://www.youtube.com/embed/' + m[1];
  return u;
}
function sanitizeLite_(s) { // keep simple tags only
  return String(s || '').replace(/<script[\s\S]*?<\/script>/gi, '').replace(/on\w+="[^"]*"/gi, '');
}

/* ---------- runtime JS injected into exports ---------- */

function buildRuntime_(cfg, adv, blocks) {
  var js = RUNTIME_JS;
  if (!adv) js = js.replace(/\/\*ADV-START\*\/[\s\S]*?\/\*ADV-END\*\//g, '');
  if (cfg.live) js = js.replace('__LIVE__', 'true').replace(/__SITEURL__/g, cfg.siteUrl || '');
  else js = js.replace('__LIVE__', 'false');
  return '<script>' + js + '</script>';
}

var EXPORT_CSS = [
  '*{box-sizing:border-box;margin:0}',
  '.ss{max-width:var(--mw);margin:0 auto;padding:0 clamp(14px,4vw,32px)}',
  '.ss section.blk{padding:calc(var(--pad)*.6) 0}',
  '.ss-h{font-family:var(--fh);font-size:clamp(1.5rem,3.5vw,2.2rem);margin-bottom:.4em}',
  '.ss-sub{opacity:.75;font-size:1.05rem}',
  '.ss-text p{margin:.55em 0;line-height:1.65}',
  '.ss-btn{display:inline-block;background:var(--pri);color:#fff;padding:.72em 1.5em;border-radius:var(--rad);text-decoration:none;font-weight:600;border:2px solid var(--pri);transition:transform .15s, box-shadow .15s}',
  '.ss-btn:hover{transform:translateY(-2px);box-shadow:0 6px 18px rgba(0,0,0,.18)}',
  '.ss-btn.ghost{background:transparent;color:var(--pri)}',
  '.ss-btn.acc{background:var(--acc);border-color:var(--acc);color:#111}',
  '.ss-btngroup{display:flex;gap:12px;flex-wrap:wrap;margin-top:1em}',
  '.ss-hero{color:#fff;display:flex;align-items:center;min-height:340px;border-radius:calc(var(--rad)*1.4);overflow:hidden;margin:14px 0}',
  '.ss-hero.h-sm{min-height:220px}.ss-hero.h-md{min-height:340px}.ss-hero.h-lg{min-height:460px}.ss-hero.h-full{min-height:78vh}',
  '.ss-hero.al-center{justify-content:center;text-align:center}.ss-hero.al-right{justify-content:flex-end;text-align:right}',
  '.ss-hero.split{gap:28px;padding:28px}.ss-hero.split .ss-heroimg{flex:1;min-width:0;max-width:46%;object-fit:cover;border-radius:var(--rad);height:100%}',
  '.ss-herotext{padding:clamp(20px,5vw,48px)}.ss-herotext h1{font-family:var(--fh);font-size:clamp(1.7rem,4.6vw,3rem);line-height:1.15;margin-bottom:.35em}',
  '.ss-herotext p{font-size:clamp(1rem,2vw,1.25rem);opacity:.92;max-width:46ch}',
  '.ss-kicker{text-transform:uppercase;letter-spacing:.14em;font-size:.8rem;font-weight:700;color:var(--acc);margin-bottom:.5em}',
  '.ss-imgwrap img{width:100%;border-radius:var(--rad);display:block}',
  '.ss-imgwrap figcaption{font-size:.85rem;opacity:.7;margin-top:.5em;text-align:center}',
  '.ss-gallery{display:grid;grid-template-columns:repeat(auto-fill,minmax(min(100%,220px),1fr));gap:12px}',
  '.ss-gallery img{width:100%;height:100%;object-fit:cover;border-radius:var(--rad);cursor:zoom-in;aspect-ratio:4/3}',
  '.ss-video{position:relative;padding-top:56.25%;border-radius:var(--rad);overflow:hidden;background:#000}',
  '.ss-video iframe{position:absolute;inset:0;width:100%;height:100%}',
  '.ss-cols{display:grid;grid-template-columns:repeat(var(--cols),1fr);gap:var(--gap,24px)}',
  '@media(max-width:760px){.ss-cols{grid-template-columns:1fr}}',
  '.ss-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,250px),1fr));gap:18px}',
  '.ss-card{background:var(--bg);border:1px solid rgba(128,128,128,.25);border-radius:var(--rad);padding:22px;box-shadow:0 3px 14px rgba(0,0,0,.06);transition:transform .18s}',
  '.ss-card:hover{transform:translateY(-4px)}',
  '.ss-card h3{font-family:var(--fh);margin:.3em 0}.ss-card p{opacity:.8;line-height:1.55}',
  '.ss-cardicon{font-size:1.9rem}',
  '.ss-price .ss-amt{font-size:2rem;font-weight:800;margin:.2em 0}.ss-price .ss-amt span{font-size:.9rem;font-weight:400;opacity:.7}',
  '.ss-price ul{list-style:none;padding:0;margin:.8em 0}.ss-price li{padding:.3em 0;border-bottom:1px dashed rgba(128,128,128,.25)}',
  '.ss-price.featured{outline:3px solid var(--acc)}',
  '.ss-list{padding-left:1.4em;line-height:1.7}',
  '.ss-steps{display:grid;gap:16px}.ss-step{display:flex;gap:16px;align-items:flex-start}.ss-num{background:var(--pri);color:#fff;min-width:36px;height:36px;border-radius:50%;display:grid;place-items:center;font-weight:700}',
  '.ss-acc details{border:1px solid rgba(128,128,128,.3);border-radius:var(--rad);margin-bottom:10px;padding:0 16px}',
  '.ss-acc summary{cursor:pointer;font-weight:600;padding:14px 0}',
  '.ss-acc-body{padding-bottom:14px;opacity:.85;line-height:1.6;white-space:pre-line}',
  '.ss-tabbar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}.ss-tab{padding:.5em 1.1em;border-radius:999px;border:1px solid rgba(128,128,128,.4);background:transparent;color:inherit;cursor:pointer;font-weight:600}',
  '.ss-tab.on{background:var(--pri);color:#fff;border-color:var(--pri)}',
  '.ss-tabpane{display:none;line-height:1.65;white-space:pre-line}.ss-tabpane.on{display:block}',
  '.ss-tablewrap{overflow-x:auto}.ss-table{width:100%;border-collapse:collapse}.ss-table th,.ss-table td{border:1px solid rgba(128,128,128,.3);padding:10px 12px;text-align:left}.ss-table th{background:var(--pri);color:#fff}',
  '.ss-quote{border-left:4px solid var(--acc);padding:14px 22px;font-size:1.15rem;line-height:1.6}.ss-quote footer{margin-top:.6em;opacity:.7;font-size:.95rem}',
  '.ss-cta{text-align:center;background:linear-gradient(135deg,var(--pri),var(--acc));color:#fff;border-radius:calc(var(--rad)*1.4);padding:clamp(28px,6vw,56px)}',
  '.ss-cta h2{font-family:var(--fh);margin-bottom:.3em}.ss-cta p{opacity:.92;margin-bottom:1.2em}',
  '.ss-cta .ss-btn{background:#fff;color:var(--pri);border-color:#fff}',
  '.ss-form{display:grid;gap:14px;max-width:560px}.ss-form label{display:grid;gap:6px;font-weight:600;font-size:.95rem}',
  '.ss-form input,.ss-form select,.ss-form textarea{padding:.7em .9em;border:1px solid rgba(128,128,128,.45);border-radius:calc(var(--rad)*.6);font:inherit;background:var(--bg);color:var(--txt);width:100%}',
  '.ss-form textarea{min-height:110px}.ss-formmsg{padding:12px;border-radius:var(--rad);background:#d1fae5;color:#065f46;font-weight:600}',
  '.dark .ss-formmsg{background:#064e3b;color:#a7f3d0}',
  '.ss-hr{border:none;border-top:1px solid rgba(128,128,128,.35)}',
  '.ss-count{display:flex;gap:14px;justify-content:center}.ss-count div{background:var(--pri);color:#fff;border-radius:var(--rad);padding:14px 18px;text-align:center;min-width:70px}.ss-count b{font-size:1.7rem;display:block}.ss-count span{opacity:.8;font-size:.8rem}',
  '.ss-clock{text-align:center;font-size:1.6rem;font-weight:700;font-variant-numeric:tabular-nums}',
  '.ss-marquee{overflow:hidden;white-space:nowrap;border-block:1px solid rgba(128,128,128,.3);padding:10px 0;font-weight:600}.ss-marquee div{display:inline-block;padding-left:100%;animation:ssmq 22s linear infinite}',
  '@keyframes ssmq{to{transform:translateX(-100%)}}',
  '.ss-badges{display:flex;gap:8px;flex-wrap:wrap}.ss-badge{background:var(--pri);color:#fff;border-radius:999px;padding:.35em .9em;font-size:.85rem;font-weight:600}',
  '.ss-nav{display:flex;gap:18px;flex-wrap:wrap}.ss-nav a{color:var(--pri);font-weight:600;text-decoration:none}',
  '.ss-dynitem{padding:12px 16px;border:1px solid rgba(128,128,128,.3);border-radius:var(--rad);margin-bottom:10px;line-height:1.55}',
  '.ss-empty{opacity:.6;font-style:italic}',
  '.ss-picker{margin-bottom:12px;display:flex;gap:8px;flex-wrap:wrap}',
  '.ss-lb{position:fixed;inset:0;background:rgba(0,0,0,.85);display:grid;place-items:center;z-index:99;cursor:zoom-out}.ss-lb img{max-width:92%;max-height:88vh;border-radius:8px}',
  /* animations */
  '.anim-fade [data-an],.anim-slide [data-an],.anim-pop [data-an],.anim-blur [data-an]{opacity:0;transform:translateY(24px);transition:opacity .6s ease, transform .6s ease, filter .6s}',
  '.anim-fade [data-an].in{opacity:1}',
  '.anim-slide [data-an].in{opacity:1;transform:none!important}',
  '.anim-pop [data-an]{transform:scale(.9)}.anim-pop [data-an].in{opacity:1;transform:scale(1)}',
  '.anim-blur [data-an].in{opacity:1;filter:none}',
  '.anim-blur [data-an]{filter:blur(6px)}',
  '[data-anim-delay]{transition-delay:.2s}',
  '@media(max-width:640px){.hide-mobile{display:none!important}}',
  'body.dark{background:#0f172a!important;color:#e5e7eb!important}'
].join('\n');

var RUNTIME_JS = [
  '(function(){',
  'var LIVE=__LIVE__, SITE_URL="__SITEURL__";',
  '/* ---- scroll reveal animations ---- */',
  'var sel=".ss section.blk, .ss-card, .ss-dynitem, .ss-step, .ss-quote";',
  'document.querySelectorAll(sel).forEach(function(el,i){el.setAttribute("data-an","1");if(i%3===2)el.setAttribute("data-anim-delay","1")});',
  'if("IntersectionObserver" in window){var io=new IntersectionObserver(function(es){es.forEach(function(en){if(en.isIntersecting){en.target.classList.add("in");io.unobserve(en.target)}})},{threshold:.12});document.querySelectorAll("[data-an]").forEach(function(el){io.observe(el)})}else{document.querySelectorAll("[data-an]").forEach(function(el){el.classList.add("in")})}',
  /* ---- tabs ---- */
  'document.querySelectorAll("[data-tabs]").forEach(function(t){var bs=t.querySelectorAll(".ss-tab"),ps=t.querySelectorAll(".ss-tabpane");bs.forEach(function(b){b.addEventListener("click",function(){bs.forEach(function(x){x.classList.remove("on")});ps.forEach(function(x){x.classList.remove("on")});b.classList.add("on");ps[+b.dataset.i].classList.add("on")})})});',
  /* ---- lightbox ---- */
  'document.querySelectorAll(".ss-gallery img[data-lightbox]").forEach(function(im){im.addEventListener("click",function(){var d=document.createElement("div");d.className="ss-lb";var i2=new Image();i2.src=im.src;d.appendChild(i2);d.onclick=function(){d.remove()};document.body.appendChild(d)})});',
  /* ---- countdown & clock ---- */
  'setInterval(function(){document.querySelectorAll(".ss-count[data-end]").forEach(function(c){var e=new Date(c.dataset.end).getTime()-Date.now();var b=c.querySelectorAll("b");if(e<=0){c.innerHTML="<b style=\\"font-size:1.3rem\\">"+(c.dataset.msg||"Done!")+"</b>";return}if(b.length===4){b[0].textContent=Math.floor(e/864e5);b[1].textContent=Math.floor(e%864e5/36e5);b[2].textContent=Math.floor(e%36e5/6e4);b[3].textContent=Math.floor(e%6e4/1e3)}});document.querySelectorAll(".ss-clock").forEach(function(c){var n=new Date();c.textContent=c.dataset.fmt==="time"?n.toLocaleTimeString():n.toLocaleString()})},1000);',
  /* ---- forms ---- */
  'document.querySelectorAll(".ss-form").forEach(function(f){f.addEventListener("submit",function(ev){ev.preventDefault();var mode=f.dataset.mode;var data={};new FormData(f).forEach(function(v,k){data[k]=v});',
  ' if(mode==="mailto"){var q=Object.keys(data).map(function(k){return encodeURIComponent(k)+": "+encodeURIComponent(data[k])}).join("%0D%0A");location.href="mailto:"+f.dataset.to+"?subject=(site form)&body="+q;}',
  ' else if(mode==="live"&&window.__ssFormEndpoint){fetch(window.__ssFormEndpoint+"&d="+encodeURIComponent(JSON.stringify(data))).then(function(){f.reset();f.querySelector(".ss-formmsg").hidden=false}).catch(function(){alert("Could not send.")})}',
  ' else{f.reset();f.querySelector(".ss-formmsg").hidden=false}',
  '})});',
  /* ---- dynamic data blocks (live tools) ---- */
  'function pad(n){return("0"+n).slice(-2)}',
  'function ymd(dt){return dt.getFullYear()+"-"+pad(dt.getMonth()+1)+"-"+pad(dt.getDate())}',
  'function isoWeek(dt){var d=new Date(Date.UTC(dt.getFullYear(),dt.getMonth(),dt.getDate()));var dn=d.getUTCDay()||7;d.setUTCDate(d.getUTCDate()+4-dn);var ys=new Date(Date.UTC(d.getUTCFullYear(),0,1));return Math.ceil((((d-ys)/864e5)+1)/7)}',
  'function renderTpl(tpl,row){return String(tpl).replace(/\\{(\\w[\\w.]*)\\}/g,function(m,k){var c=row;String(k).split(".").forEach(function(x){c=c&&c[x]});return c==null?"":String(c)})}',
  'function curPeriod(st){var wk="W"+isoWeek(new Date());if(st&&st.weeks&&st.weeks[wk])return wk;if(st&&st.current)return st.current;return wk}',
  'function bindValue(bind,st){',
  ' if(bind==="all")return"all";',
  ' if(bind==="current")return curPeriod(st);',
  ' if(bind==="week")return"W"+isoWeek(new Date())+((st&&st.yearMatch)?"-"+new Date().getFullYear():"");',
  ' if(bind==="day")return ymd(new Date());',
  ' if(bind.indexOf("key:")===0)return st?st[bind.slice(4)]:bind.slice(4);',
  ' return bind}',
  'function collectDays(store){if(!store)return[];var out=[];Object.keys(store).forEach(function(k){if(/^\\d{4}-\\d{2}-\\d{2}$/.test(k))out.push({key:k,label:k})});return out.sort()}',
  'function fillDyn(el,rows){var tpl=el.dataset.tpl,list=el.querySelector(".ss-dynlist");var pick=el.querySelector(".ss-picker");',
  ' var keep=pick?pick.outerHTML:"";',
  ' var html=keep+rows.map(function(r){return"<div class=\\"ss-dynitem\\">"+renderTpl(tpl,r)+"</div>"}).join("");',
  ' if(!rows.length&&el.dataset.empty)html+="<p class=\\"ss-empty\\">"+el.dataset.empty+"</p>";',
  ' list.innerHTML=html;wirePickers(el)}',
  'function wirePickers(el){var pick=el.querySelector(".ss-picker");if(!pick)return;',
  ' var store=el.__store||{};',
  ' var keys=Object.keys(store).filter(function(k){return k!=="weeks"&&k!=="current"&&k!=="yearMatch"});',
  ' var lab=el.dataset.pickerLabel||"Show:";',
  ' var btns=keys.map(function(k){return"<button class=\\"ss-tab\\" data-pk=\\""+k+"\\">"+k+"</button>"}).join("");',
  ' var days=collectDays(store).map(function(d){return"<button class=\\"ss-tab\\" data-day=\\""+d.key+"\\" data-mode=\\"static\\">"+d.label+"</button>"}).join("");',
  ' pick.innerHTML="<span style=\\"font-weight:600;padding-top:6px\\">"+lab+"</span>"+btns+days;',
  ' pick.querySelectorAll("button").forEach(function(b){b.onclick=function(){',
  '   el.dataset.static=b.dataset.pk||b.dataset.day;el.dataset.mode="static";runDyn(el,true)}})}',
  'function runDyn(el,skipStore){var key=el.dataset.key,mode=el.dataset.mode,sv=el.dataset.static;',
  ' function apply(store){if(!skipStore)el.__store=store;else if(store)el.__store=store;',
  '  var val,data=[];',
  '  if(mode==="static"&&sv){val=sv}else{val=bindValue(el.dataset.bind,store)}',
  '  if(store){if(store[val]!=null)data=Array.isArray(store[val])?store[val]:[store[val]];else if(store.weeks&&store.weeks[val])data=store.weeks[val];else if(Array.isArray(store))data=store}',
  '  if(el.dataset.filterField)data=data.filter(function(r){return String(r[el.dataset.filterField])===el.dataset.filterVal});',
  '  if(el.dataset.sort)data=data.slice().sort(function(a,b){return String(a[el.dataset.sort]).localeCompare(String(b[el.dataset.sort]),undefined,{numeric:true})});',
  '  fillDyn(el,data)}',
  ' if(mode==="live"&&LIVE){fetch(SITE_URL+"?action=data&project="+el.dataset.project+"&key="+encodeURIComponent(key))',
  '  .then(function(r){return r.json()}).then(function(j){apply(j&&j.ok?j.value:null)}).catch(function(){apply(null)})}',
  ' else{var st=null;try{st=JSON.parse(el.getAttribute("data-store")||"null")}catch(e){}apply(st)}}',
  'document.querySelectorAll(".ss-dyn").forEach(function(el){runDyn(el)});',
  '/*ADV-START*/',
  'document.querySelectorAll("a[href^=\\"#\\"]").forEach(function(a){a.addEventListener("click",function(e){var t=document.querySelector(a.getAttribute("href"));if(t){e.preventDefault();t.scrollIntoView({behavior:"smooth"})}})});',
  '/*ADV-END*/',
  '})();'
].join('\n');

/* ---------- serve-time hooks so live exports can call us without OAuth ---------- */

function doGetData(e) {
  var action = e.parameter.action;
  if (action === 'data') return fetchSiteDataPublic(e.parameter.project, e.parameter.key);
  if (action === 'form') {
    var key = e.parameter.key || 'responses';
    var projectId = e.parameter.project || 'forms';
    try {
      var data = JSON.parse(decodeURIComponent(e.parameter.d || '{}'));
      var sh = sheet_('SiteData'), v = sh.getDataRange().getValues(), found = 0;
      for (var i = 1; i < v.length; i++) if (v[i][0] === projectId && v[i][1] === key) { found = i + 1; break; }
      var arr = [];
      if (found) { try { arr = JSON.parse(v[found - 1][2]); } catch (err) { arr = []; } }
      data._at = now_(); arr.push(data);
      var cell = JSON.stringify(arr);
      if (cell.length > 45000) cell = JSON.stringify(arr.slice(-200));
      if (found) { sh.getRange(found, 3).setValue(cell); sh.getRange(found, 4).setValue(now_()); }
      else sh.appendRow([projectId, key, cell, now_()]);
    } catch (err) { return { ok: false, error: String(err) }; }
    return { ok: true };
  }
  return { ok: false };
}
