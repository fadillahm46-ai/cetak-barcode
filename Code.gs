/**
 * Konfigurasi Spreadsheet & Script Google Apps Script
 * Script ID: 17vmBd8EUVLSCBKf-EBvCkB7U-FQDWHw8av-h9bFyIWPv1WjUm_k9G8pI
 */
var SPREADSHEET_ID = '1Zc_xxMda_sUrmr-MxZdmCmAN31liNeCNw7shUQTLmg8';
var SHEET_NAME = 'History';

/**
 * Menampilkan halaman utama Web App atau mengembalikan data JSON jika dipanggil dari Vercel (?action=list)
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'list') {
    var list = getHistory();
    return ContentService.createTextOutput(JSON.stringify(list))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Barcode Link Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Endpoint API POST untuk menerima aksi simpan & hapus dari Vercel
 */
function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    var result = {};
    if (body.action === 'save') {
      result = saveBarcode(body.payload || {});
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

/**
 * Mengambil seluruh daftar riwayat barcode dari Google Sheets
 */
function getHistory() {
  try {
    var sheet = getOrCreateSheet_();
    if (!sheet) {
      var raw = PropertiesService.getScriptProperties().getProperty('BARCODE_HISTORY');
      return raw ? JSON.parse(raw) : [];
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var list = [];

    // Urutkan dari yang terbaru (bawah sheet ke atas)
    for (var i = values.length - 1; i >= 0; i--) {
      list.push({
        id: String(values[i][0]),
        label: String(values[i][1] || ''),
        url: String(values[i][2] || ''),
        type: String(values[i][3] || 'qr'),
        createdAt: String(values[i][4] || '')
      });
    }

    return list;
  } catch (err) {
    return [];
  }
}

/**
 * Menyimpan data barcode baru ke baris Google Sheets
 */
function saveBarcode(payload) {
  try {
    var item = {
      id: payload.id || String(new Date().getTime()),
      label: (payload.label || 'Tanpa Nama').trim(),
      url: (payload.url || '').trim(),
      type: payload.type || 'qr',
      createdAt: payload.createdAt || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'dd/MM/yyyy HH:mm')
    };

    var sheet = getOrCreateSheet_();
    if (sheet) {
      sheet.appendRow([item.id, item.label, item.url, item.type, item.createdAt]);
    } else {
      var list = getHistory();
      list.unshift(item);
      PropertiesService.getScriptProperties().setProperty('BARCODE_HISTORY', JSON.stringify(list));
    }

    return item;
  } catch (err) {
    throw new Error('Gagal menyimpan: ' + err.message);
  }
}

/**
 * Menghapus barcode berdasarkan ID dari Google Sheets
 */
function deleteBarcode(id) {
  try {
    var targetId = String(id);
    var sheet = getOrCreateSheet_();

    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]) === targetId) {
            sheet.deleteRow(i + 2);
          }
        }
      }
    } else {
      var list = getHistory().filter(function(item) {
        return String(item.id) !== targetId;
      });
      PropertiesService.getScriptProperties().setProperty('BARCODE_HISTORY', JSON.stringify(list));
    }

    return { success: true, id: targetId };
  } catch (err) {
    throw new Error('Gagal menghapus: ' + err.message);
  }
}

/**
 * Helper: Membuka Google Sheets menggunakan ID yang telah ditentukan
 */
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
      sheet.appendRow(['ID', 'Nama', 'Link URL', 'Tipe', 'Tanggal']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    return sheet;
  } catch (e) {
    console.error('Error getOrCreateSheet_: ' + e.message);
    return null;
  }
}
