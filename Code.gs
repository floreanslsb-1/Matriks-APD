/**
 * SISTEM MANAJEMEN APD OTOMATIS
 * Versi: Row-Based LIST APD dengan Named Ranges
 */

// ── NAMED RANGE COLUMN READERS ─────────────────────────────────────────────
// Reads the LIST APD sheet using named ranges.
// Returns an array of row objects, one per data row (skips header row 1).

function getListApdRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Each named range covers a full column (e.g. 'LIST APD'!D:D).
  // We read them all, then zip into row objects.
  var rangeMap = {
    number:        'PPE_NUMBER',
    category:      'PPE_CATEGORY',
    name:          'PPE_NAME',
    specification: 'PPE_SPECIFICATION',
    pic:           'PPE_PIC',
    link:          'PPE_LINK',
    mid:           'PPE_MID',
    matDesc:       'PPE_MAT_DESC',
    size:          'PPE_SIZE',
    lifetime:      'PPE_LIFETIME',
    type:          'PPE_TYPE'
  };

  // Pull values for every named range column
  var cols = {};
  var maxLen = 0;
  for (var key in rangeMap) {
    var r = ss.getRangeByName(rangeMap[key]);
    if (!r) { cols[key] = []; continue; }
    var vals = r.getValues(); // 2-D array [[v],[v],...]
    cols[key] = vals.map(function(row) { return String(row[0] || '').trim(); });
    if (cols[key].length > maxLen) maxLen = cols[key].length;
  }

  // Zip into row objects, skip header (index 0) and empty-name rows
  var rows = [];
  for (var i = 2; i < maxLen; i++) {
    var name = (cols.name[i] || '').trim();
    if (!name) continue; // skip blank rows
    rows.push({
      number:        cols.number[i]        || '',
      category:      cols.category[i]      || '',
      name:          name,
      specification: cols.specification[i] || '',
      link:          cols.link[i]          || '',
      mid:           cols.mid[i]           || '',
      matDesc:       cols.matDesc[i]       || '',
      size:          cols.size[i]          || '',
      lifetime:      cols.lifetime[i]      || '',
      type:          cols.type[i]          || ''
    });
  }
  return rows;
}

// ── BUILD APD LOOKUP ────────────────────────────────────────────────────────
// Groups LIST APD rows by PPE_NAME.
// Output structure per name:
// {
//   nama, kategori, lifetime, specification,
//   variantType: 'single' | 'sized' | 'typed' | 'typed-sized',
//   images: { '': 'https://...' }  or  { 'Pria': '...', 'Wanita': '...' },
//   link: '...',   // representative image (first available)
//   variants: [{ type, size, mid, link }]
// }

function buildApdLookup() {
  var rows = getListApdRows();
  var lookup = {};

  rows.forEach(function(row) {
    var name = row.name; // already trimmed
    if (!name) return;

    if (!lookup[name]) {
      lookup[name] = {
        nama:          name,
        kategori:      row.category,
        lifetime:      row.lifetime,
        specification: row.specification,
        link:          row.link || '',
        images:        {},
        variants:      []
      };
    }

    var entry = lookup[name];

    // Keep first available representative link
    if (!entry.link && row.link) entry.link = row.link;

    // Store first image per type (empty string key = no type)
    var typeKey = row.type || '';
    if (!entry.images[typeKey] && row.link) {
      entry.images[typeKey] = row.link;
    }

    // Always keep most-complete lifetime
    if (!entry.lifetime && row.lifetime) entry.lifetime = row.lifetime;

    entry.variants.push({
      type: row.type || '',
      size: (row.size === '-' ? '' : row.size) || '',
      mid:  row.mid  || '',
      link: row.link || ''
    });
  });

  // Determine variantType for each entry
  for (var name in lookup) {
    var entry = lookup[name];
    var hasType = entry.variants.some(function(v) { return v.type !== ''; });
    var hasSize = entry.variants.some(function(v) { return v.size !== ''; });

    if (hasType && hasSize)       entry.variantType = 'typed-sized';
    else if (hasType)             entry.variantType = 'typed';
    else if (hasSize)             entry.variantType = 'sized';
    else                          entry.variantType = 'single';
  }

  return lookup;
}

// ── GET DATA (for Matriks Risiko) ──────────────────────────────────────────
// Reads the 'All' sheet (consolidated view of all department sheets).
// Department sheet cells store PPE_NAME values, newline-separated for multi-PPE.

function getData() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAll    = ss.getSheetByName('All');
  var apdLookup   = buildApdLookup();

  var allData = sheetAll.getDataRange().getValues();
  var headers = allData[0];

  var COL = {
    id:         headers.indexOf('ID'),
    no:         headers.indexOf('No'),
    aspek:      findColIndex(headers, ['Asumsi', 'Kondisi', 'Bahaya', 'Aspek']),
    kepala:     findColIndex(headers, ['Pelindung Kepala']),
    mataMuka:   findColIndex(headers, ['Mata', 'Muka']),
    telinga:    findColIndex(headers, ['Telinga']),
    pernapasan: findColIndex(headers, ['Pernapasan']),
    tangan:     findColIndex(headers, ['Tangan']),
    kaki:       findColIndex(headers, ['Kaki']),
    pakaian:    findColIndex(headers, ['Pakaian']),
    jatuh:      findColIndex(headers, ['Jatuh']),
    ergonomi:   findColIndex(headers, ['Ergonomi']),
    section:    findColIndex(headers, ['Section'])
  };

  var categoryMap = {
    'Pelindung Kepala':      COL.kepala,
    'Pelindung Mata & Muka': COL.mataMuka,
    'Pelindung Telinga':     COL.telinga,
    'Pelindung Pernapasan':  COL.pernapasan,
    'Pelindung Tangan':      COL.tangan,
    'Pelindung Kaki':        COL.kaki,
    'Pakaian Pelindung':     COL.pakaian,
    'Pelindung Jatuh':       COL.jatuh,
    'Ergonomi':              COL.ergonomi
  };

  var result = [];

  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    if (!row[COL.id] || String(row[COL.id]).trim() === '') continue;

    var apdList = [];

    for (var kategori in categoryMap) {
      var colIdx = categoryMap[kategori];
      if (colIdx === -1) continue;

      var cellVal = String(row[colIdx] || '').trim();
      if (!cellVal) continue;

      // Support multiple PPE in one cell (newline-separated)
      var ppeNames = cellVal.split('\n').map(function(s) { return s.trim(); }).filter(Boolean);

      ppeNames.forEach(function(ppeName) {
        // Exact match by PPE_NAME (case-insensitive, trimmed)
        var apdDetail = findInLookup(apdLookup, ppeName);

        apdList.push({
          kategori:    kategori,
          nama:        ppeName,
          apdData:     apdDetail || null,  // full lookup entry (variants, images, etc.)
          imageLink:   apdDetail ? apdDetail.link : '',
          lifetime:    apdDetail ? apdDetail.lifetime : '',
          variantType: apdDetail ? apdDetail.variantType : 'single'
        });
      });
    }

    result.push({
      id:      String(row[COL.id] || '').trim(),
      no:      row[COL.no] || i,
      resiko:  String(row[COL.aspek] || '').trim(),
      section: String(row[COL.section >= 0 ? COL.section : headers.length - 1] || '').trim(),
      apd:     apdList
    });
  }

  return result;
}

// Case-insensitive, trimmed exact match on PPE_NAME
function findInLookup(lookup, name) {
  var key = name.trim();
  // Direct match first
  if (lookup[key]) return lookup[key];
  // Case-insensitive fallback
  var lower = key.toLowerCase();
  for (var k in lookup) {
    if (k.toLowerCase() === lower) return lookup[k];
  }
  return null;
}

// ── GET APD CATALOG ────────────────────────────────────────────────────────
// Returns one entry per unique PPE_NAME for the catalog table.
// Aggregates all sizes and types per name.

function getApdCatalog() {
  var lookup = buildApdLookup();
  var catalog = [];

  for (var name in lookup) {
    var entry = lookup[name];

    // Collect unique sizes and types for display
    var sizes = [];
    var types = [];
    entry.variants.forEach(function(v) {
      if (v.size && sizes.indexOf(v.size) === -1) sizes.push(v.size);
      if (v.type && types.indexOf(v.type) === -1) types.push(v.type);
    });

    catalog.push({
      nama:         entry.nama,
      kategori:     entry.kategori,
      specification:entry.specification,
      link:         entry.link,
      lifetime:     entry.lifetime,
      variantType:  entry.variantType,
      sizes:        sizes,   // ['6','7','8'] or []
      types:        types,   // ['Pria','Wanita'] or []
      variants:     entry.variants,
      images:       entry.images
    });
  }

  // Sort by category then name
  catalog.sort(function(a, b) {
    var catCmp = a.kategori.localeCompare(b.kategori);
    return catCmp !== 0 ? catCmp : a.nama.localeCompare(b.nama);
  });

  return catalog;
}

// ── ON EDIT ────────────────────────────────────────────────────────────────
// Simplified: only handles timestamp update and background highlight.
// Auto-fill of APD columns (H-P) has been removed.

function onEdit(e) {
  var range = e.range;
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

  var sheet     = range.getSheet();
  var sheetName = sheet.getName();

  // Skip protected sheets
  var proteksi = ['ASUMSI BASIC', 'LIST APD', 'MASTER COPY', 'All'];
  if (proteksi.indexOf(sheetName) !== -1) return;

  // Only act in data area: row 18+, columns E-P (5-16)
  if (range.getRow() < 18 || range.getColumn() < 5 || range.getColumn() > 16) return;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    // Highlight edited cell
    sheet.getRange('E18:P100').setBackground(null);
    range.setBackground('#d9ead3');

    // Update timestamp via named range
    var timestampCell = null;
    var allNamedRanges = sheet.getNamedRanges();

    for (var i = 0; i < allNamedRanges.length; i++) {
      if (allNamedRanges[i].getName().toLowerCase().endsWith('tanggalupdate')) {
        timestampCell = allNamedRanges[i].getRange().getCell(1, 1);
        break;
      }
    }

    if (!timestampCell) {
      var globalRange = e.source.getRangeByName('TanggalUpdate');
      if (globalRange && globalRange.getSheet().getName() === sheetName) {
        timestampCell = globalRange.getCell(1, 1);
      }
    }

    if (timestampCell) {
      timestampCell.setValue(new Date());
      timestampCell.setNumberFormat('dd/mm/yyyy HH:mm');
      timestampCell.setBackground('#d9ead3');
    }

  } catch (f) {
    console.log('Error pada onEdit: ' + f.toString());
  } finally {
    lock.releaseLock();
  }
}

// ── ON SELECTION CHANGE ────────────────────────────────────────────────────
function onSelectionChange(e) {
  var range = e.range;
  var sheet = range.getSheet();

  var displayCell;
  try {
    displayCell = sheet.getRange('LihatAPD');
  } catch(err) { return; }

  var defaultFormula = "='LIST APD'!BC5";
  var val = range.getValue().toString().trim();

  if (!val || val === '🔎 Lihat Gambar' || sheet.getName() === 'LIST APD') {
    if (displayCell.getFormula() !== defaultFormula) displayCell.setFormula(defaultFormula);
    return;
  }

  try {
    var dbRange = e.source.getRangeByName('DatabaseGambarAPD');
    if (!dbRange) return;

    var dataNama  = dbRange.getValues()[0].map(function(s) { return s.toString().trim(); });
    var kolomIndex = dataNama.indexOf(val);

    if (kolomIndex !== -1) {
      var kolomExcel = dbRange.getColumn() + kolomIndex;
      var barisGambar = dbRange.getRow() + 1;
      var formula = "='LIST APD'!" + columnToLetter(kolomExcel) + barisGambar;
      if (displayCell.getFormula() !== formula) displayCell.setFormula(formula);
    } else {
      if (displayCell.getFormula() !== defaultFormula) displayCell.setFormula(defaultFormula);
    }
  } catch (err) {
    displayCell.setFormula(defaultFormula);
  }
}

// ── DEPARTMENT MANAGEMENT ──────────────────────────────────────────────────
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 SISTEM APD')
      .addItem('➕ Tambah Departemen Baru',      'tambahDepartemenDialog')
      .addItem('🗑️ Hapus Departemen',            'hapusDepartemenDialog')
      .addItem('🔄 Sinkronisasi Named Range',     'setupNamedRangesMassal')
      .addItem('🔁 Sinkronisasi Nama APD',        'sinkronisasiNamaAPD')
      .addItem('🔁 Retry Gambar Gagal Saja',      'bakeGagalSaja')
      .addItem('🍞 Update Web App (Admin)',        'bakeData')
      .addToUi();
}

function tambahDepartemenDialog() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Tambah Departemen', 'Masukkan Nama Departemen Baru:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    var namaBaru = response.getResponseText().trim();
    if (namaBaru) buatSheetBaru(namaBaru);
  }
}

function hapusDepartemenDialog() {
  var ui        = SpreadsheetApp.getUi();
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAktif = ss.getActiveSheet();
  var namaSheet  = sheetAktif.getName();

  var proteksi = ['ASUMSI BASIC', 'LIST APD', 'MASTER COPY', 'All'];
  if (proteksi.indexOf(namaSheet) !== -1) {
    ui.alert('⚠️ Sheet ini tidak boleh dihapus!');
    return;
  }

  var response = ui.alert('Konfirmasi', 'Hapus departemen "' + namaSheet + '"?', ui.ButtonSet.YES_NO);
  if (response === ui.Button.YES) {
    ss.deleteSheet(sheetAktif);
    setupNamedRangesMassal();
  }
}

function buatSheetBaru(nama) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var template = ss.getSheetByName('MASTER COPY');
  if (!template) return;

  var newSheet = template.copyTo(ss).setName(nama);
  newSheet.showSheet();

  try {
    newSheet.getRange('NamaDepartemen').setValue(nama);
  } catch(e) { Logger.log('Named Range NamaDepartemen tidak ditemukan'); }

  setupNamedRangesMassal();
}

function setupNamedRangesMassal() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var namedRanges = ss.getNamedRanges();

  for (var i = 0; i < namedRanges.length; i++) {
    try { namedRanges[i].getRange(); }
    catch (e) { namedRanges[i].remove(); }
  }
  Logger.log('Pembersihan Named Range rusak selesai.');
}

// ── SINKRONISASI NAMA APD ──────────────────────────────────────────────────
// Updates department sheet cells to use PPE_NAME (from LIST APD) instead of old spec strings.
// Matches by trimmed case-insensitive comparison of the cell value against PPE_NAME.

function sinkronisasiNamaAPD() {
  var ui  = SpreadsheetApp.getUi();
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var lookup = buildApdLookup();

  var proteksi  = ['ASUMSI BASIC', 'LIST APD', 'MASTER COPY', 'All'];
  var deptSheets = ss.getSheets().filter(function(s) {
    return proteksi.indexOf(s.getName()) === -1;
  });

  if (!deptSheets.length) {
    ui.alert('⚠️ Tidak ada sheet departemen ditemukan.');
    return;
  }

  var totalReplaced = 0;
  var log = [];

  deptSheets.forEach(function(sheet) {
    var data = sheet.getDataRange().getValues();

    for (var r = 17; r < data.length; r++) {
      for (var c = 4; c <= 15; c++) {
        var cellVal = String(data[r][c] || '').trim();
        if (!cellVal) continue;

        // Handle newline-separated multi-PPE cells
        var lines = cellVal.split('\n').map(function(s) { return s.trim(); }).filter(Boolean);
        var newLines = lines.map(function(line) {
          var match = findInLookup(lookup, line);
          if (match && match.nama !== line) {
            log.push(sheet.getName() + ' R' + (r+1) + 'C' + (c+1) + ': "' + line + '" → "' + match.nama + '"');
            totalReplaced++;
            return match.nama;
          }
          return line;
        });

        var newVal = newLines.join('\n');
        if (newVal !== cellVal) {
          sheet.getRange(r + 1, c + 1).setValue(newVal);
        }
      }
    }
  });

  console.log('--- SYNC LOG ---\n' + log.join('\n'));

  if (totalReplaced === 0) {
    ui.alert('✅ Sinkronisasi Selesai', 'Semua data sudah sinkron.', ui.ButtonSet.OK);
  } else {
    ui.alert('✅ Sinkronisasi Selesai!',
      totalReplaced + ' sel diperbarui di ' + deptSheets.length + ' sheet.\nDetail di Execution Log.',
      ui.ButtonSet.OK);
  }
}

// ── WEB APP ────────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Matriks APD - PT. Sayap Mas Utama')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── IMAGE HANDLING (unchanged) ─────────────────────────────────────────────
function getImageBase64(url) {
  if (!url || url.trim() === '') return '';
  try {
    var driveIdMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
                       url.match(/[?&]id=([a-zA-Z0-9_-]+)/);

    if (driveIdMatch) {
      var fileId = driveIdMatch[1];
      var token  = ScriptApp.getOAuthToken();

      try {
        var thumbUrl  = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w150';
        var thumbResp = UrlFetchApp.fetch(thumbUrl, {
          muteHttpExceptions: true,
          headers: { 'Authorization': 'Bearer ' + token }
        });
        var thumbMime = (thumbResp.getHeaders()['Content-Type'] || thumbResp.getHeaders()['content-type'] || '');
        if (thumbResp.getResponseCode() === 200 && thumbMime.includes('image/')) {
          var b64 = Utilities.base64Encode(thumbResp.getBlob().getBytes());
          return 'data:' + thumbMime.split(';')[0].trim() + ';base64,' + b64;
        }
      } catch(thumbErr) {
        console.log('Thumbnail failed: ' + thumbErr);
      }

      Utilities.sleep(500);
      var exportUrl  = 'https://drive.google.com/uc?export=download&id=' + fileId;
      var exportResp = UrlFetchApp.fetch(exportUrl, {
        followRedirects: true, muteHttpExceptions: true,
        headers: { 'Authorization': 'Bearer ' + token }
      });
      var exportMime = (exportResp.getHeaders()['Content-Type'] || exportResp.getHeaders()['content-type'] || '');
      if (exportResp.getResponseCode() === 200 && exportMime.includes('image/')) {
        var b64 = Utilities.base64Encode(exportResp.getBlob().getBytes());
        return 'data:' + exportMime.split(';')[0].trim() + ';base64,' + b64;
      }

      Utilities.sleep(300);
      var apiUrl  = 'https://www.googleapis.com/drive/v3/files/' + fileId + '?alt=media&supportsAllDrives=true';
      var apiResp = UrlFetchApp.fetch(apiUrl, {
        muteHttpExceptions: true,
        headers: { 'Authorization': 'Bearer ' + token }
      });
      var apiMime = (apiResp.getHeaders()['Content-Type'] || apiResp.getHeaders()['content-type'] || '');
      if (apiResp.getResponseCode() === 200 && apiMime.includes('image/')) {
        var b64 = Utilities.base64Encode(apiResp.getBlob().getBytes());
        return 'data:' + apiMime.split(';')[0].trim() + ';base64,' + b64;
      }

      console.error('All methods failed for: ' + fileId);
      return '';
    }

    if (url.startsWith('http')) {
      Utilities.sleep(200);
      var resp = UrlFetchApp.fetch(url, { followRedirects: true, muteHttpExceptions: true });
      var mime = (resp.getHeaders()['Content-Type'] || resp.getHeaders()['content-type'] || '');
      if (resp.getResponseCode() === 200 && mime.includes('image/')) {
        var b64 = Utilities.base64Encode(resp.getBlob().getBytes());
        return 'data:' + mime.split(';')[0].trim() + ';base64,' + b64;
      }
    }
    return '';
  } catch(e) {
    console.error('getImageBase64 error: ' + url + ' | ' + e.toString());
    return '';
  }
}

function getImageBase64Cached(url) {
  if (!url || url.trim() === '') return '';
  var cache = CacheService.getScriptCache();
  var key   = 'img_' + Utilities.base64Encode(url).substring(0, 240);
  var cached = cache.get(key);
  if (cached) return cached;
  var result = getImageBase64(url);
  if (result) {
    var CHUNK = 90000;
    if (result.length <= CHUNK) {
      try { cache.put(key, result, 21600); } catch(e) {}
    } else {
      var chunks = Math.ceil(result.length / CHUNK);
      for (var i = 0; i < chunks; i++) {
        try { cache.put(key + '_' + i, result.substring(i * CHUNK, (i+1) * CHUNK), 21600); } catch(e) {}
      }
      try { cache.put(key + '_chunks', String(chunks), 21600); } catch(e) {}
    }
  }
  return result;
}

function getImageCached(url) {
  if (!url || url.trim() === '') return '';
  var cache = CacheService.getScriptCache();
  var key   = 'img_' + Utilities.base64Encode(url).substring(0, 240);
  var single = cache.get(key);
  if (single) return single;
  var chunksStr = cache.get(key + '_chunks');
  if (chunksStr) {
    var chunks = parseInt(chunksStr), result = '';
    for (var i = 0; i < chunks; i++) {
      var part = cache.get(key + '_' + i);
      if (!part) return '';
      result += part;
    }
    return result;
  }
  return getImageBase64Cached(url);
}

function clearImageCache() {
  return { cleared: true, time: new Date().toISOString() };
}

// ── OFFLINE IMAGE STORE ────────────────────────────────────────────────────
var STORE_FOLDER_NAME = 'APD_ImageStore';
var INDEX_FILENAME    = 'APD_Index.json';

function getStoreFolder() {
  var folders = DriveApp.getFoldersByName(STORE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(STORE_FOLDER_NAME);
}

function loadIndex() {
  var folder = getStoreFolder();
  var files  = folder.getFilesByName(INDEX_FILENAME);
  if (!files.hasNext()) return {};
  try { return JSON.parse(files.next().getBlob().getDataAsString()); }
  catch(e) { return {}; }
}

function saveIndex(index) {
  var folder = getStoreFolder();
  var json   = JSON.stringify(index);
  var files  = folder.getFilesByName(INDEX_FILENAME);
  if (files.hasNext()) { files.next().setContent(json); }
  else { folder.createFile(INDEX_FILENAME, json, MimeType.PLAIN_TEXT); }
}

function isDriveFileStale(url, indexEntry) {
  try {
    var match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (!match) return false;
    var lastModified = DriveApp.getFileById(match[1]).getLastUpdated().getTime();
    return lastModified > (indexEntry.lastModified || 0);
  } catch(e) { return false; }
}

function loadImageStore() {
  var result = {};
  try {
    var folder = getStoreFolder();
    var index  = loadIndex();
    Object.keys(index).forEach(function(url) {
      var entry    = index[url];
      var filename = typeof entry === 'string' ? entry : entry.filename;
      var files    = folder.getFilesByName(filename);
      if (!files.hasNext()) return;
      try {
        var content = files.next().getBlob().getDataAsString();
        if (content && content.startsWith('data:image/')) result[url] = content;
      } catch(e) {}
    });
  } catch(e) { console.error('loadImageStore error: ' + e.toString()); }
  return result;
}

function getImageStore() { return loadImageStore(); }

function hashCode(str) {
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash |= 0;
  }
  return hash;
}

// ── BAKE DATA ──────────────────────────────────────────────────────────────
// Collects all unique image URLs from LIST APD (via named range PPE_LINK),
// fetches/caches them, embeds into baked_data.json.

function bakeData() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('⏳ Memproses...', 'Sedang memproses data + gambar.\nKlik OK untuk mulai.', ui.ButtonSet.OK);

  try {
    // Collect all unique URLs from PPE_LINK named range
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var linkRange = ss.getRangeByName('PPE_LINK');
    var linkVals  = linkRange ? linkRange.getValues() : [];
    var allUrls   = [];
    linkVals.forEach(function(row) {
      var url = String(row[0] || '').trim();
      if (url && allUrls.indexOf(url) === -1) allUrls.push(url);
    });

    var folder = getStoreFolder();
    var index  = loadIndex();
    var store  = {};
    var ok = 0, skipped = 0, fail = 0;

    allUrls.forEach(function(url, idx) {
      var needsFetch = true;
      var reason     = 'new URL';

      if (index[url]) {
        var entry    = index[url];
        var filename = typeof entry === 'string' ? entry : entry.filename;
        var files    = folder.getFilesByName(filename);
        if (files.hasNext()) {
          try {
            var content = files.next().getBlob().getDataAsString();
            if (content && content.startsWith('data:image/')) {
              if (typeof entry === 'object' && isDriveFileStale(url, entry)) {
                reason = 'source file changed';
                delete index[url];
              } else {
                store[url] = content;
                skipped++;
                needsFetch = false;
              }
            } else { reason = 'corrupted'; }
          } catch(e) { reason = 'unreadable'; }
        } else { reason = 'missing file'; delete index[url]; }
      }

      if (!needsFetch) return;

      if ((ok + fail) > 0 && (ok + fail) % 3 === 0) {
        console.log('💤 Throttling...');
        Utilities.sleep(2000);
      }

      try {
        var base64 = getImageBase64(url);
        if (base64 && base64.startsWith('data:image/')) {
          store[url] = base64;
          var newFilename = 'img_' + (ok + skipped) + '_' + Math.abs(hashCode(url)) + '.txt';
          var existCheck  = folder.getFilesByName(newFilename);
          if (existCheck.hasNext()) { existCheck.next().setContent(base64); }
          else { folder.createFile(newFilename, base64, MimeType.PLAIN_TEXT); }
          var driveMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
          var lastMod = 0;
          if (driveMatch) {
            try { lastMod = DriveApp.getFileById(driveMatch[1]).getLastUpdated().getTime(); } catch(e) {}
          }
          index[url] = { filename: newFilename, lastModified: lastMod };
          ok++;
          console.log('✅ ' + ok + ': ' + url.substring(0, 60));
        } else { fail++; console.log('❌ Failed: ' + url.substring(0, 60)); }
      } catch(e) {
        fail++;
        console.log('❌ Error: ' + e.toString());
        Utilities.sleep(3000);
      }
    });

    saveIndex(index);

    // Embed images into data
    var data = getData();
    data.forEach(function(item) {
      item.apd.forEach(function(a) {
        if (a.imageLink && store[a.imageLink]) a._base64 = store[a.imageLink];
        // Also embed per-type images
        if (a.apdData && a.apdData.images) {
          a.apdData._base64Images = {};
          Object.keys(a.apdData.images).forEach(function(typeKey) {
            var imgUrl = a.apdData.images[typeKey];
            if (imgUrl && store[imgUrl]) a.apdData._base64Images[typeKey] = store[imgUrl];
          });
        }
      });
    });

    // Save baked_data.json
    var folderName  = 'Matriks APD Web';
    var webFolder;
    var webFolders  = DriveApp.getFoldersByName(folderName);
    if (webFolders.hasNext()) { webFolder = webFolders.next(); }
    else { webFolder = DriveApp.createFolder(folderName); }

    var json     = JSON.stringify(data);
    var fileName = 'baked_data.json';
    var existing = webFolder.getFilesByName(fileName);
    if (existing.hasNext()) { existing.next().setContent(json); }
    else { webFolder.createFile(fileName, json, MimeType.PLAIN_TEXT); }

    webFolder.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);

    ui.alert('✅ Selesai!',
      'Web app berhasil diupdate!\n\n' +
      '✅ Gambar baru: ' + ok + '\n' +
      '⏭️ Dilewati: ' + skipped + '\n' +
      '❌ Gagal: ' + fail,
      ui.ButtonSet.OK);

  } catch(e) {
    ui.alert('❌ Error', e.toString(), ui.ButtonSet.OK);
  }
}

function getBakedData() {
  try {
    var folders = DriveApp.getFoldersByName('Matriks APD Web');
    if (!folders.hasNext()) return null;
    var folder = folders.next();
    var files  = folder.getFilesByName('baked_data.json');
    if (!files.hasNext()) return null;
    return JSON.parse(files.next().getBlob().getDataAsString());
  } catch(e) { console.error('getBakedData error: ' + e.toString()); return null; }
}

function bakeGagalSaja() {
  var ui        = SpreadsheetApp.getUi();
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var linkRange = ss.getRangeByName('PPE_LINK');
  var linkVals  = linkRange ? linkRange.getValues() : [];

  var folder = getStoreFolder();
  var index  = loadIndex();
  var ok = 0, fail = 0;

  linkVals.forEach(function(row) {
    var url = String(row[0] || '').trim();
    if (!url) return;

    var needsFetch = true;
    if (index[url]) {
      var entry    = index[url];
      var filename = typeof entry === 'string' ? entry : entry.filename;
      var files    = folder.getFilesByName(filename);
      if (files.hasNext()) {
        try {
          var content = files.next().getBlob().getDataAsString();
          if (content && content.startsWith('data:image/')) needsFetch = false;
        } catch(e) {}
      }
    }
    if (!needsFetch) return;

    Utilities.sleep(2000);
    try {
      var base64 = getImageBase64(url);
      if (base64 && base64.startsWith('data:image/')) {
        var filename = 'img_retry_' + Math.abs(hashCode(url)) + '.txt';
        var existCheck = folder.getFilesByName(filename);
        if (existCheck.hasNext()) { existCheck.next().setContent(base64); }
        else { folder.createFile(filename, base64, MimeType.PLAIN_TEXT); }
        index[url] = filename;
        ok++;
      } else { fail++; }
    } catch(e) { fail++; Utilities.sleep(5000); }
  });

  saveIndex(index);
  ui.alert('✅ Retry Selesai', 'Berhasil: ' + ok + '\nMasih gagal: ' + fail, ui.ButtonSet.OK);
}

function prewarmImageCache() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var linkRange = ss.getRangeByName('PPE_LINK');
  if (!linkRange) return 0;
  var linkVals = linkRange.getValues();
  var warmed   = 0;
  linkVals.forEach(function(row) {
    var url = String(row[0] || '').trim();
    if (!url) return;
    var cache = CacheService.getScriptCache();
    var key   = 'img_' + Utilities.base64Encode(url).substring(0, 240);
    if (cache.get(key) || cache.get(key + '_chunks')) { warmed++; return; }
    getImageBase64Cached(url);
    warmed++;
  });
  return warmed;
}

// ── UTILITIES ──────────────────────────────────────────────────────────────
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp   = (column - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function findColIndex(headers, keywords) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase();
    if (keywords.every(function(k) { return h.includes(k.toLowerCase()); })) return i;
  }
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase();
    if (keywords.some(function(k)  { return h.includes(k.toLowerCase()); })) return i;
  }
  return -1;
}

function clearAllCache() {
  var ui = SpreadsheetApp.getUi();
  try {
    var folders = DriveApp.getFoldersByName('APD_ImageStore');
    if (folders.hasNext()) {
      var folder = folders.next(), files = folder.getFiles(), count = 0;
      while (files.hasNext()) { files.next().setTrashed(true); count++; }
      console.log('Deleted ' + count + ' files from APD_ImageStore');
    }
    var webFolders = DriveApp.getFoldersByName('Matriks APD Web');
    if (webFolders.hasNext()) {
      var webFolder = webFolders.next();
      var baked = webFolder.getFilesByName('baked_data.json');
      if (baked.hasNext()) baked.next().setTrashed(true);
    }
    ui.alert('✅ Cache Dibersihkan!', 'Jalankan "Update Web App" untuk rebuild.', ui.ButtonSet.OK);
  } catch(e) { ui.alert('❌ Error', e.toString(), ui.ButtonSet.OK); }
}
function submitFeedback(payload) {
  var result = { success: false, sheetSaved: false, emailSent: false, error: '' };

  // Declare at top scope so email block can access them
  var fotoUrls = [];
  var ssUrl    = '';

  // ── 1. Save to sheet ────────────────────────────────────────────────────
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    ssUrl     = ss.getUrl();
    var sheet = ss.getSheetByName('Feedback');

    if (!sheet) {
      sheet = ss.insertSheet('Feedback');
      sheet.appendRow(['Timestamp', 'Jenis', 'Nama', 'Deskripsi', 'Tab', 'Browser', 'Foto 1', 'Foto 2', 'Foto 3']);
      sheet.setFrozenRows(1);
      var header = sheet.getRange(1, 1, 1, 9);
      header.setBackground('#2d3748');
      header.setFontColor('#ffffff');
      header.setFontWeight('bold');
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(4, 400);
    }

    // ── Handle photos: save to Drive and get URLs ──
    var fotos = payload.foto || [];
    if (!Array.isArray(fotos)) fotos = fotos ? [fotos] : [];

    var folder;
    try {
      var folders = DriveApp.getFoldersByName('Feedback Photos - Matriks APD');
      folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Feedback Photos - Matriks APD');
    } catch(fe) {
      folder = DriveApp.getRootFolder();
    }

    fotos.forEach(function(base64, i) {
      try {
        if (!base64 || !base64.startsWith('data:image/')) return;
        var parts = base64.split(',');
        var mime  = parts[0].match(/:(.*?);/)[1];
        var ext   = mime.split('/')[1] || 'jpg';
        var blob  = Utilities.newBlob(
          Utilities.base64Decode(parts[1]), mime,
          'feedback_' + new Date().getTime() + '_' + i + '.' + ext
        );
        var file  = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fotoUrls.push(file.getUrl());
      } catch(fe) {
        console.error('Photo upload error:', fe);
      }
    });

    // Pad to 3 columns
    while (fotoUrls.length < 3) fotoUrls.push('');

    sheet.appendRow([
      new Date(),
      payload.jenis     || '—',
      payload.nama      || 'Anonim',
      payload.deskripsi,
      payload.tab       || '—',
      payload.browser   || '—',
      fotoUrls[0],
      fotoUrls[1],
      fotoUrls[2]
    ]);

    result.sheetSaved = true;

  } catch(e) {
    result.error = 'Sheet error: ' + e.toString();
    console.error(result.error);
    return result;
  }

  // ── 2. Send email ───────────────────────────────────────────────────────
  try {
    var ADMIN_EMAIL = 'floreansalsabila.irdana@wingscorp.com';
    var subject     = '📣 [Feedback Matriks APD] ' + (payload.jenis || '');

    // Build photo HTML for email
    var fotoHtml = '';
    if (fotoUrls.filter(Boolean).length) {
      fotoHtml =
        '<div style="margin-top:16px;">' +
          '<p style="margin:0 0 8px;font-size:12px;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:0.5px;">📷 Foto</p>' +
          '<div style="display:flex;gap:8px;flex-wrap:wrap;">' +
            fotoUrls.filter(Boolean).map(function(url, i) {
              return '<a href="' + url + '" target="_blank" style="display:inline-block;background:#f3f4f6;border:1px solid #e5e7eb;border-radius:8px;padding:8px 14px;font-size:12px;color:#2d3748;text-decoration:none;font-weight:600;">📎 Foto ' + (i + 1) + '</a>';
            }).join('') +
          '</div>' +
        '</div>';
    }

    var plainBody =
      'Feedback baru masuk di Matriks APD:\n\n' +
      'Jenis  : ' + (payload.jenis || '—') + '\n' +
      'Nama   : ' + (payload.nama  || 'Anonim') + '\n' +
      'Tab    : ' + (payload.tab   || '—') + '\n' +
      'Waktu  : ' + new Date().toLocaleString('id-ID') + '\n\n' +
      'Pesan:\n' + payload.deskripsi + '\n\n' +
      'Lihat feedback dan foto di spreadsheet: ' + ssUrl;

    var htmlBody =
      '<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden;">' +
        '<div style="background:linear-gradient(135deg,#2d3748,#3d4a5c);padding:20px 24px;">' +
          '<h2 style="color:white;margin:0;font-size:16px;">📣 Feedback Baru — Matriks APD</h2>' +
        '</div>' +
        '<div style="padding:24px;">' +
          '<table style="width:100%;border-collapse:collapse;font-size:14px;">' +
            '<tr style="border-bottom:1px solid #f3f4f6;">' +
              '<td style="padding:10px 0;color:#6b7280;width:80px;">🏷️ Jenis</td>' +
              '<td style="padding:10px 0;font-weight:600;color:#1f2937;">' + (payload.jenis || '—') + '</td>' +
            '</tr>' +
            '<tr style="border-bottom:1px solid #f3f4f6;">' +
              '<td style="padding:10px 0;color:#6b7280;">👤 Nama</td>' +
              '<td style="padding:10px 0;font-weight:600;color:#1f2937;">' + (payload.nama || 'Anonim') + '</td>' +
            '</tr>' +
            '<tr style="border-bottom:1px solid #f3f4f6;">' +
              '<td style="padding:10px 0;color:#6b7280;">📄 Tab</td>' +
              '<td style="padding:10px 0;font-weight:600;color:#1f2937;">' + (payload.tab || '—') + '</td>' +
            '</tr>' +
            '<tr>' +
              '<td style="padding:10px 0;color:#6b7280;">🕐 Waktu</td>' +
              '<td style="padding:10px 0;font-weight:600;color:#1f2937;">' + new Date().toLocaleString('id-ID') + '</td>' +
            '</tr>' +
          '</table>' +
          '<div style="margin-top:20px;background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;padding:16px;">' +
            '<p style="margin:0 0 8px;font-size:12px;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:0.5px;">📝 Pesan</p>' +
            '<p style="margin:0;font-size:14px;color:#1f2937;line-height:1.6;white-space:pre-wrap;">' + payload.deskripsi + '</p>' +
          '</div>' +
          fotoHtml +
        '</div>' +
        '<div style="background:#f9fafb;padding:12px 24px;border-top:1px solid #e5e7eb;font-size:11px;color:#9ca3af;">' +
          'Lihat semua feedback dan foto di <a href="' + ssUrl + '" style="color:#c0392b;font-weight:600;">spreadsheet</a>.' +
        '</div>' +
      '</div>';

    MailApp.sendEmail({
      to:       ADMIN_EMAIL,
      subject:  subject,
      body:     plainBody,
      htmlBody: htmlBody
    });

    result.emailSent = true;

  } catch(e) {
    result.error = 'Email error: ' + e.toString();
    console.error(result.error);
  }

  result.success = result.sheetSaved;
  return result;
}
