/**
 * Konfigurasi Spreadsheet & Script Google Apps Script
 * Script ID: 17vmBd8EUVLSCBKf-EBvCkB7U-FQDWHw8av-h9bFyIWPv1WjUm_k9G8pI
 */
var SPREADSHEET_ID = '1Zc_xxMda_sUrmr-MxZdmCmAN31liNeCNw7shUQTLmg8';
var SHEET_NAME = 'History';

/**
 * Menangani akses GET:
 * 1. Pengambilan link tujuan berdasarkan ID (?action=get&id=...)
 * 2. Redirect langsung jika ada yang membuka link Apps Script (?r=ID)
 * 3. API sinkronisasi (?action=list, save, update, delete)
 */
function doGet(e) {
  // 1. API cepat untuk mengambil URL tujuan dari Web Vercel
  if (e && e.parameter && e.parameter.action === 'get') {
    var target = getRedirectUrl(e.parameter.id);
    return ContentService.createTextOutput(JSON.stringify({ success: !!target, url: target }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 2. Redirect fallback langsung jika dibuka lewat link Apps Script
  if (e && e.parameter && e.parameter.r) {
    var targetUrl = getRedirectUrl(e.parameter.r);
    if (targetUrl) {
      if (!/^https?:\/\//i.test(targetUrl)) {
        targetUrl = 'https://' + targetUrl;
      }
      var html = '<!DOCTYPE html><html><head>' +
        '<meta charset="UTF-8">' +
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
        '<meta http-equiv="refresh" content="0;url=' + encodeURI(targetUrl) + '">' +
        '<title>Mengarahkan...</title>' +
        '</head><body>' +
        '<script>window.location.replace("' + encodeURI(targetUrl) + '");<\/script>' +
        '</body></html>';
      return HtmlService.createHtmlOutput(html)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } else {
      return HtmlService.createHtmlOutput('<div style="font-family:sans-serif;text-align:center;padding:40px;"><h3>Link barcode tidak ditemukan.</h3></div>');
    }
  }

  // 3. API sinkronisasi CORS-friendly dari Vercel
  if (e && e.parameter && e.parameter.action) {
    var res = {};
    var action = e.parameter.action;
    if (action === 'list') {
      res = getHistory();
    } else if (action === 'update') {
      res = updateBarcode({ id: e.parameter.id, label: e.parameter.label, url: e.parameter.url });
    } else if (action === 'save') {
      res = saveBarcode({
        id: e.parameter.id,
        label: e.parameter.label,
        url: e.parameter.url,
        type: e.parameter.type,
        mode: e.parameter.mode,
        createdAt: e.parameter.createdAt
      });
    } else if (action === 'delete') {
      res = deleteBarcode(e.parameter.id);
    }
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Barcode Link Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    var result = {};
    if (body.action === 'save') {
      result = saveBarcode(body.payload || body);
    } else if (body.action === 'update') {
      result = updateBarcode(body.payload || body);
    } else if (body.action === 'delete') {
      result = deleteBarcode(body.id);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function isSameId(val1, val2) {
  if (val1 == null || val2 == null) return false;
  var s1 = String(val1).trim().replace(/^'+/, '');
  var s2 = String(val2).trim().replace(/^'+/, '');
  if (s1 === s2) return true;

  var clean1 = s1.replace(/^qr_/, '');
  var clean2 = s2.replace(/^qr_/, '');
  if (clean1 === clean2) return true;

  var n1 = parseFloat(clean1);
  var n2 = parseFloat(clean2);
  if (!isNaN(n1) && !isNaN(n2) && Math.abs(n1 - n2) < 0.001) return true;

  return false;
}

function getRedirectUrl(id) {
  try {
    var targetId = String(id || '').trim();
    if (!targetId) return null;

    var sheet = getOrCreateSheet_();
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
        for (var i = data.length - 1; i >= 0; i--) {
          if (isSameId(data[i][0], targetId)) {
            var url = String(data[i][2] || '').trim();
            if (url) return url;
          }
        }
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

function getHistory() {
  try {
    var sheet = getOrCreateSheet_();
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    var list = [];

    for (var i = values.length - 1; i >= 0; i--) {
      list.push({
        id: String(values[i][0]).replace(/^'+/, ''),
        label: String(values[i][1] || ''),
        url: String(values[i][2] || ''),
        type: String(values[i][3] || 'qr'),
        createdAt: String(values[i][4] || ''),
        mode: String(values[i][5] || 'dynamic')
      });
    }

    return list;
  } catch (err) {
    return [];
  }
}

function saveBarcode(payload) {
  try {
    var rawId = payload.id ? String(payload.id).trim() : '';
    var itemId = rawId ? rawId : ('qr_' + new Date().getTime());

    var item = {
      id: itemId,
      label: (payload.label || 'Tanpa Nama').trim(),
      url: (payload.url || '').trim(),
      type: payload.type || 'qr',
      createdAt: payload.createdAt || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'dd/MM/yyyy HH:mm'),
      mode: payload.mode || 'dynamic'
    };

    var sheet = getOrCreateSheet_();
    if (sheet) {
      sheet.appendRow(["'" + item.id, item.label, item.url, item.type, item.createdAt, item.mode]);
      SpreadsheetApp.flush();
    }

    return item;
  } catch (err) {
    throw new Error('Gagal menyimpan: ' + err.message);
  }
}

function updateBarcode(payload) {
  try {
    var targetId = String(payload.id || '').trim();
    var newLabel = (payload.label || '').trim();
    var newUrl = (payload.url || '').trim();

    if (!targetId) throw new Error('ID barcode tidak valid');

    var sheet = getOrCreateSheet_();
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (isSameId(ids[i][0], targetId)) {
            var row = i + 2;
            if (newLabel) sheet.getRange(row, 2).setValue(newLabel);
            if (newUrl) sheet.getRange(row, 3).setValue(newUrl);
            SpreadsheetApp.flush();
            return { success: true, id: targetId, label: newLabel, url: newUrl };
          }
        }
      }
    }
    return { success: true, id: targetId, warning: 'Baris tidak ditemukan di Sheets' };
  } catch (err) {
    throw new Error('Gagal memperbarui: ' + err.message);
  }
}

function deleteBarcode(id) {
  try {
    var targetId = String(id || '').trim();
    var sheet = getOrCreateSheet_();

    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = ids.length - 1; i >= 0; i--) {
          if (isSameId(ids[i][0], targetId)) {
            sheet.deleteRow(i + 2);
          }
        }
        SpreadsheetApp.flush();
      }
    }

    return { success: true, id: targetId };
  } catch (err) {
    throw new Error('Gagal menghapus: ' + err.message);
  }
}

function getOrCreateSheet_() {
  try {
    var ss;
    if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== '') {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID.trim());
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }

    if (!ss) return null;

    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['ID', 'Nama', 'Link URL', 'Tipe', 'Tanggal', 'Mode']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }
    return sheet;
  } catch (e) {
    console.error('Error getOrCreateSheet_: ' + e.message);
    return null;
  }
}
