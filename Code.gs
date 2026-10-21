/** WM Design portfolio manager. Run setupPortfolio() once before deployment. */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'admin';
  if (action === 'gallery') {
    var items = getPublicPortfolio_();
    var callback = (e.parameter.callback || '').trim();
    var output = JSON.stringify(items).replace(/</g, '\\u003c');
    if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callback)) {
      return ContentService.createTextOutput(callback + '(' + output + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('Admin')
    .setTitle('WM Design — Portfolio Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupPortfolio() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('PORTFOLIO_FOLDER_ID');
  var sheetId = props.getProperty('PORTFOLIO_SHEET_ID');
  if (!folderId || !sheetId) {
    var folder = DriveApp.createFolder('WM Design — Public Portfolio Images');
    var ss = SpreadsheetApp.create('WM Design — Portfolio Index');
    var sheet = ss.getSheets()[0];
    sheet.setName('Portfolio');
    sheet.appendRow(['id', 'title', 'category', 'description', 'fileId', 'imageUrl', 'createdAt', 'active']);
    sheet.setFrozenRows(1);
    folderId = folder.getId();
    sheetId = ss.getId();
    props.setProperties({PORTFOLIO_FOLDER_ID: folderId, PORTFOLIO_SHEET_ID: sheetId});
  }
  Logger.log('Portfolio folder: https://drive.google.com/drive/folders/' + folderId);
  Logger.log('Portfolio index: https://docs.google.com/spreadsheets/d/' + sheetId + '/edit');
  Logger.log('Next: set Script Property ADMIN_PASSWORD to a strong, private password, then deploy as Web App.');
  return 'Setup complete. Open Executions/Logs for the folder and spreadsheet links.';
}

function adminLogin(password) {
  return isAuthorized_(password);
}

function listPortfolioForAdmin(password) {
  requireAdmin_(password);
  return getRows_(true);
}

function savePortfolioItem(password, payload) {
  requireAdmin_(password);
  if (!payload || !payload.title || !payload.category || !payload.dataUrl) throw new Error('Please provide a title, category and image.');
  var match = String(payload.dataUrl).match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);
  if (!match) throw new Error('Use a JPG, PNG or WebP image.');
  var bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 4 * 1024 * 1024) throw new Error('Image must be 4 MB or smaller.');
  var fileName = String(payload.fileName || 'wm-portfolio-image').replace(/[\\/:*?"<>|]/g, '-').slice(0, 100);
  var blob = Utilities.newBlob(bytes, match[1], fileName);
  var folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('PORTFOLIO_FOLDER_ID'));
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var fileId = file.getId();
  var imageUrl = 'https://drive.google.com/uc?export=view&id=' + encodeURIComponent(fileId);
  var sheet = getPortfolioSheet_();
  var id = Utilities.getUuid();
  sheet.appendRow([id, clean_(payload.title, 120), clean_(payload.category, 60), clean_(payload.description || '', 500), fileId, imageUrl, new Date(), true]);
  return {ok: true, id: id};
}

function updatePortfolioItem(password, id, payload) {
  requireAdmin_(password);
  var sheet = getPortfolioSheet_();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id) && values[i][7] !== false) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[clean_(payload.title, 120), clean_(payload.category, 60), clean_(payload.description || '', 500)]]);
      return {ok: true};
    }
  }
  throw new Error('Portfolio item not found.');
}

function deletePortfolioItem(password, id) {
  requireAdmin_(password);
  var sheet = getPortfolioSheet_();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id) && values[i][7] !== false) {
      sheet.getRange(i + 1, 8).setValue(false);
      return {ok: true};
    }
  }
  throw new Error('Portfolio item not found.');
}

function getPublicPortfolio_() {
  try { return getRows_(true); } catch (err) { return []; }
}

function getRows_(activeOnly) {
  var sheet = getPortfolioSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (activeOnly && r[7] !== true) continue;
    if (!r[0]) continue;
    out.push({id:String(r[0]), title:String(r[1] || ''), category:String(r[2] || ''), description:String(r[3] || ''), fileId:String(r[4] || ''), imageUrl:String(r[5] || ''), createdAt:r[6] instanceof Date ? r[6].toISOString() : String(r[6] || ''), active:r[7] === true});
  }
  return out.reverse();
}

function getPortfolioSheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('PORTFOLIO_SHEET_ID');
  if (!id) throw new Error('Run setupPortfolio() once from the Apps Script editor.');
  return SpreadsheetApp.openById(id).getSheetByName('Portfolio');
}
function isAuthorized_(password) {
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return !!stored && String(password || '') === stored;
}
function requireAdmin_(password) {
  if (!isAuthorized_(password)) throw new Error('Incorrect admin password.');
}
function clean_(value, max) { return String(value || '').trim().replace(/[\u0000-\u001f]/g, '').slice(0, max); }
