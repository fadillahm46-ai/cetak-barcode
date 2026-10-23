/**
 * Menampilkan halaman utama Web App atau mengembalikan JSON jika dipanggil dari Vercel (?action=list)
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
 * Endpoint API POST untuk menerima simpan/hapus dari Vercel
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
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
 * Menyimpan data barcode baru ke riwayat
 */
function saveBarcode(payload) {
  try {
    var item = {
      id: payload.id || String(new Date().getTime()),
      label: (payload.label || 'Tanpa Nama').trim(),
      url: (payload.url || '').trim(),
      type: payload.type || 'qr',
      createdAt: payload.createdAt || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
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
 * Menghapus satu item barcode berdasarkan ID
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
 * Helper: Mengambil atau membuat sheet 'History' jika terhubung dengan Google Sheets
 */
function getOrCreateSheet_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName('History');
    if (!sheet) {
      sheet = ss.insertSheet('History');
      sheet.appendRow(['ID', 'Nama', 'Link URL', 'Tipe', 'Tanggal']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    return sheet;
  } catch (e) {
    return null;
  }
}
