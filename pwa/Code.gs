/**
 * code.gs — Backend untuk Dashboard SK
 * Google Apps Script Web App
 *
 * Sumber data: Spreadsheet aktif (tanpa hardcode ID).
 * Deteksi sheet, header, dan kolom secara dinamis.
 */

// ═══════════════════════════════════════════════════════════════════
//  ENTRY POINT
// ═══════════════════════════════════════════════════════════════════

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard Surat Keputusan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ═══════════════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════════════

var BULAN = [
  'Januari','Februari','Maret','April','Mei','Juni',
  'Juli','Agustus','September','Oktober','November','Desember'
];

/**
 * parseDate_ — Ubah nilai sel jadi Date object.
 * Support: Date object (GAS native), dd/MM/yyyy, yyyy-MM-dd, dd-MM-yyyy.
 * Catatan: selalu gunakan dd/MM/yyyy (bukan MM/dd) untuk format slash.
 */
function parseDate_(val) {
  if (!val) return null;

  // Jika sudah Date object (format spreadsheet default)
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }

  var s = val.toString().trim();
  if (!s) return null;

  var p, d;
  if (s.includes('/')) {
    // Deteksi format: coba dd/MM/yyyy dulu (format Indonesia)
    p = s.split('/');
    if (p.length === 3) {
      var day   = parseInt(p[0], 10);
      var month = parseInt(p[1], 10);
      var year  = parseInt(p[2], 10);
      // Validasi range
      if (day >= 1 && day <= 31 && month >= 1 && month <= 12 && year > 1900) {
        d = new Date(year, month - 1, day);
      } else {
        // Fallback: coba MM/dd/yyyy
        d = new Date(s);
      }
    } else {
      d = new Date(s);
    }
  } else if (s.includes('-')) {
    p = s.split('-');
    d = p[0].length === 4 ? new Date(s) : new Date(+p[2], +p[1] - 1, +p[0]);
  } else {
    d = new Date(s);
  }
  return (d && !isNaN(d.getTime())) ? d : null;
}

/**
 * colIndex_ — Cari index kolom pertama yang cocok dengan salah satu keyword.
 * Case-insensitive, partial match. Return -1 jika tidak ada.
 */
function colIndex_(headers, keywords) {
  for (var i = 0; i < headers.length; i++) {
    var h = (headers[i] || '').toString().toUpperCase();
    for (var k = 0; k < keywords.length; k++) {
      if (h.includes(keywords[k].toUpperCase())) return i;
    }
  }
  return -1;
}

/**
 * colIndexAll_ — Cari SEMUA index kolom yang mengandung keyword.
 */
function colIndexAll_(headers, keyword) {
  var kw = keyword.toUpperCase();
  var out = [];
  for (var i = 0; i < headers.length; i++) {
    if ((headers[i] || '').toString().toUpperCase().includes(kw)) out.push(i);
  }
  return out;
}

/**
 * detectHeaderRow_ — Deteksi baris header di sheet.
 * Scan 5 baris pertama, pilih yang paling banyak sel non-kosong.
 * Return { row: 1-based, headers: string[] }
 */
function detectHeaderRow_(sheet) {
  var maxScan = Math.min(sheet.getLastRow(), 5);
  var lastCol = sheet.getLastColumn();
  var bestRow = 1, bestCount = 0, bestHeaders = [];

  for (var r = 1; r <= maxScan; r++) {
    var row = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    var count = 0;
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== '' && row[c] !== null && row[c] !== undefined) count++;
    }
    if (count > bestCount) {
      bestCount = count;
      bestRow = r;
      bestHeaders = row;
    }
  }
  return { row: bestRow, headers: bestHeaders };
}

/**
 * normStatus_ — Normalisasi status string.
 */
function normStatus_(s) {
  return (s || '').toString().toLowerCase().trim();
}

// ═══════════════════════════════════════════════════════════════════
//  FIND LATEST UPDATE DATE
// ═══════════════════════════════════════════════════════════════════

/**
 * findLatestUpdateDate_ — Cari tanggal last update terbaru dari 2D array data.
 *
 * Strategi deteksi kolom (urutan prioritas):
 *   1. Kolom 'LAST UPDATE' / variasi nama (PERBARUI, UPDATE, TERAKHIR, dsb.)
 *   2. Fallback ke kolom 'TANGGAL MASUK' / 'TANGGAL DITERIMA' / dsb.
 *
 * Setelah kolom ditemukan, scan semua baris untuk mencari Date terbaru.
 * Return string 'dd MMMM yyyy' yang aman dikirim via google.script.run,
 * atau string kosong jika tidak ada data valid.
 *
 * @param {string[]} headers  — Array header kolom (baris pertama data).
 * @param {Array[]}  rows     — 2D array data baris.
 * @param {string}   tz       — Timezone string dari Session.getScriptTimeZone().
 * @returns {string}  Tanggal terformat 'dd MMMM yyyy', atau '' jika tidak ditemukan.
 */
function findLatestUpdateDate_(headers, rows, tz) {
  // ── 1. Deteksi kolom secara dinamis ───────────────────────────────

  // Cari kolom Last Update (prioritas utama)
  var updateIdx = colIndex_(headers, [
    'LAST UPDATE', 'LASTUPDATE', 'UPDATE',
    'TERAKHIR DIPERBARUI', 'TANGGAL UPDATE', 'TANGGAL PERUBAHAN',
    'DIPERBARUI', 'PERBARUI'
  ]);

  // Fallback: cari kolom Tanggal Masuk / Tanggal Diterima
  var fallbackIdx = -1;
  if (updateIdx === -1) {
    fallbackIdx = colIndex_(headers, [
      'TANGGAL MASUK', 'TANGGAL TERIMA', 'TANGGAL DITERIMA',
      'TGL MASUK', 'DATE IN', 'TANGGAL'
    ]);
  }

  var colIdx = updateIdx !== -1 ? updateIdx : fallbackIdx;

  // Jika tidak ada kolom apapun yang cocok, return kosong
  if (colIdx === -1 || rows.length === 0) return '';

  // ── 2. Scan semua baris — cari Date terbaru ────────────────────────

  var latestDate = null;

  for (var r = 0; r < rows.length; r++) {
    var cellVal = rows[r][colIdx];

    // Skip sel kosong
    if (cellVal === '' || cellVal === null || cellVal === undefined) continue;

    var d;
    // Date object langsung dari spreadsheet (format sel Tanggal)
    if (cellVal instanceof Date) {
      d = isNaN(cellVal.getTime()) ? null : cellVal;
    } else {
      d = parseDate_(cellVal);
    }

    if (!d) continue;

    // Perbarui latestDate jika d lebih baru
    if (!latestDate || d.getTime() > latestDate.getTime()) {
      latestDate = d;
    }
  }

  if (!latestDate) return '';

  // ── 3. Format timezone-safe menggunakan Utilities.formatDate ────────
  // Utilities.formatDate menggunakan Java SimpleDateFormat.
  // Gunakan BULAN[] untuk nama bulan Indonesia (MMMM di Java = English locale).
  var day      = Utilities.formatDate(latestDate, tz, 'dd');      // '09'
  var monthIdx = parseInt(Utilities.formatDate(latestDate, tz, 'M'), 10) - 1; // 0-based
  var year     = Utilities.formatDate(latestDate, tz, 'yyyy');    // '2029'

  return day + ' ' + BULAN[monthIdx] + ' ' + year; // '09 September 2029'
}

// ═══════════════════════════════════════════════════════════════════
//  LAST UPDATE FROM INVENTARIS SHEET (KOLOM E)
// ═══════════════════════════════════════════════════════════════════

/**
 * getLastUpdateFromInventarisColE_ — Ambil tanggal terbaru dari kolom E
 * sheet 'Inventaris 2019 - Sekarang', difilter berdasarkan tahun berjalan.
 *
 * Logika:
 *   1. Baca semua nilai kolom E (index 4, 0-based) mulai baris 2.
 *   2. Filter hanya baris yang tahunnya = tahun berjalan.
 *   3. Dari baris yang lolos filter, cari tanggal terbesar (terbaru).
 *   4. Format hasil sebagai 'dd Bulan yyyy' (Indonesian locale).
 *
 * @param {Sheet}  sheet — Sheet 'Inventaris 2019 - Sekarang'.
 * @param {string} tz    — Timezone string dari Session.getScriptTimeZone().
 * @returns {string} Tanggal terformat 'dd Bulan yyyy', atau '' jika tidak ada.
 */
function getLastUpdateFromInventarisColE_(sheet, tz) {
  if (!sheet) return '';

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';

  var curYear = new Date().getFullYear();

  // Baca kolom E saja (kolom ke-5, 1-based) dari baris 2 hingga lastRow
  var colEValues = sheet.getRange(2, 5, lastRow - 1, 1).getValues();

  var latestDate = null;

  for (var i = 0; i < colEValues.length; i++) {
    var cellVal = colEValues[i][0];

    // Skip sel kosong
    if (cellVal === '' || cellVal === null || cellVal === undefined) continue;

    // Parse ke Date object
    var d;
    if (cellVal instanceof Date) {
      d = isNaN(cellVal.getTime()) ? null : cellVal;
    } else {
      d = parseDate_(cellVal);
    }
    if (!d) continue;

    // Filter: hanya data tahun berjalan
    var rowYear = parseInt(Utilities.formatDate(d, tz, 'yyyy'), 10);
    if (rowYear !== curYear) continue;

    // Perbarui tanggal terbaru
    if (!latestDate || d.getTime() > latestDate.getTime()) {
      latestDate = d;
    }
  }

  if (!latestDate) return '';

  // Format timezone-safe
  var day      = Utilities.formatDate(latestDate, tz, 'dd');             // '09'
  var monthIdx = parseInt(Utilities.formatDate(latestDate, tz, 'M'), 10) - 1; // 0-based
  var year     = Utilities.formatDate(latestDate, tz, 'yyyy');            // '2026'

  return day + ' ' + BULAN[monthIdx] + ' ' + year; // '09 September 2026'
}

/**
 * formatMonthYearFromCell_ — Validasi dan konversi nilai sel menjadi string 'Bulan Tahun'.
 * Contoh: '9/9/2029' → 'September 2029', Date(2029, 8, 9) → 'September 2029'.
 * Mengembalikan string 'Bulan Tahun' jika sel berisi tanggal valid, atau '' jika tidak.
 *
 * @param {*}      cellVal — Nilai dari sel spreadsheet (Date, string, number).
 * @param {string} tz      — Timezone string.
 * @returns {string}       — 'Bulan Tahun' (misal: 'September 2029') atau ''.
 */
function formatMonthYearFromCell_(cellVal, tz) {
  if (cellVal === '' || cellVal === null || cellVal === undefined) return '';

  // 1. Jika bertipe Date (format default Google Sheets untuk sel tanggal)
  if (cellVal instanceof Date) {
    if (isNaN(cellVal.getTime())) return '';
    var yearVal = parseInt(Utilities.formatDate(cellVal, tz, 'yyyy'), 10);
    if (isNaN(yearVal) || yearVal < 1900) return '';
    var mIdx = parseInt(Utilities.formatDate(cellVal, tz, 'M'), 10) - 1;
    if (mIdx >= 0 && mIdx < 12) {
      return BULAN[mIdx] + ' ' + yearVal;
    }
    return '';
  }

  // 2. Jika bertipe number (serial number tanggal Google Sheets / Excel, misal 47369)
  if (typeof cellVal === 'number' && !isNaN(cellVal) && cellVal > 30000 && cellVal < 60000) {
    var dSerial = new Date(Math.round((cellVal - 25569) * 86400 * 1000));
    if (!isNaN(dSerial.getTime())) {
      var y = parseInt(Utilities.formatDate(dSerial, tz, 'yyyy'), 10);
      var m = parseInt(Utilities.formatDate(dSerial, tz, 'M'), 10) - 1;
      if (y > 1900 && m >= 0 && m < 12) {
        return BULAN[m] + ' ' + y;
      }
    }
  }

  // 3. Jika bertipe string
  var s = cellVal.toString().trim();
  if (!s) return '';

  // Format slash: contoh 9/9/2029 atau 09/09/2029 atau 2029/09/09
  if (s.indexOf('/') !== -1) {
    var p = s.split('/');
    if (p.length === 3) {
      var p0 = parseInt(p[0].trim(), 10);
      var p1 = parseInt(p[1].trim(), 10);
      var p2 = parseInt(p[2].trim(), 10);
      if (!isNaN(p0) && !isNaN(p1) && !isNaN(p2)) {
        var day, month, year;
        if (p[0].trim().length === 4) {
          // yyyy/MM/dd
          year = p0; month = p1; day = p2;
        } else {
          // dd/MM/yyyy atau MM/dd/yyyy
          year = p2 < 100 ? (2000 + p2) : p2;
          if (p0 > 12 && p1 <= 12) {
            day = p0; month = p1;
          } else if (p1 > 12 && p0 <= 12) {
            month = p0; day = p1;
          } else {
            // Standar Indonesia: dd/MM/yyyy (misal 9/9/2029 -> day: 9, month: 9)
            day = p0; month = p1;
          }
        }
        if (year > 1900 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          return BULAN[month - 1] + ' ' + year;
        }
      }
    }
  }

  // Format dash: contoh 2029-09-09 atau 09-09-2029
  if (s.indexOf('-') !== -1) {
    var p = s.split('-');
    if (p.length === 3) {
      var p0 = parseInt(p[0].trim(), 10);
      var p1 = parseInt(p[1].trim(), 10);
      var p2 = parseInt(p[2].trim(), 10);
      if (!isNaN(p0) && !isNaN(p1) && !isNaN(p2)) {
        var day, month, year;
        if (p[0].trim().length === 4) {
          year = p0; month = p1; day = p2;
        } else {
          year = p2 < 100 ? (2000 + p2) : p2;
          if (p0 > 12 && p1 <= 12) {
            day = p0; month = p1;
          } else if (p1 > 12 && p0 <= 12) {
            month = p0; day = p1;
          } else {
            day = p0; month = p1;
          }
        }
        if (year > 1900 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          return BULAN[month - 1] + ' ' + year;
        }
      }
    }
  }

  // Format teks nama bulan: contoh 'September 2029' atau '9 September 2029'
  var sLower = s.toLowerCase();
  for (var b = 0; b < BULAN.length; b++) {
    if (sLower.indexOf(BULAN[b].toLowerCase()) !== -1) {
      var matchYear = s.match(/\b(19\d\d|20\d\d)\b/);
      if (matchYear) {
        return BULAN[b] + ' ' + matchYear[1];
      }
    }
  }

  // Fallback: parseDate_
  var dFallback = parseDate_(cellVal);
  if (dFallback && !isNaN(dFallback.getTime())) {
    var yF = parseInt(Utilities.formatDate(dFallback, tz, 'yyyy'), 10);
    if (yF > 1900) {
      var mF = parseInt(Utilities.formatDate(dFallback, tz, 'M'), 10) - 1;
      if (mF >= 0 && mF < 12) {
        return BULAN[mF] + ' ' + yF;
      }
    }
  }

  return '';
}

// ═══════════════════════════════════════════════════════════════════
//  LAST UPDATE FROM TREND CARD (ANALITIK TREN TAHUNAN)
// ═══════════════════════════════════════════════════════════════════

/**
 * getLastUpdateFromTrend_ — Mengambil nilai Pembaruan Terakhir berdasarkan
 * data source dan logic existing pada Card Analitik Tren Tahunan.
 *
 * Mengambil tahun terakhir dari data tren tahunan (sheet 'Raw Data' A17:B30,
 * atau fallback monthlyTrend dari data utama).
 * Jika informasi bulan tersedia untuk tahun tersebut, diformat 'Bulan Tahun'
 * (contoh: 'September 2024' atau 'September 2029').
 * Jika tidak ada informasi bulan, ditampilkan tahun saja (contoh: '2024').
 *
 * @param {Object|null} rawKPI    — Hasil readKPIFromRawData_(ss).
 * @param {Object}      analytics — Hasil computeAnalytics_(headers, rows).
 * @param {string[]}    headers   — Array header kolom.
 * @param {Array[]}     rows      — 2D array baris data.
 * @param {string}      tz        — Timezone string.
 * @returns {string}              — 'Bulan Tahun' atau 'Tahun'.
 */
function getLastUpdateFromTrend_(rawKPI, analytics, headers, rows, tz) {
  // 1. Cek apakah ada cell asli terakhir dari Card Analitik Tren Tahunan (A17:B30)
  var lastCell = (rawKPI && rawKPI.lastTrendCell != null) ? rawKPI.lastTrendCell : null;

  // Jika lastCell berisi tanggal valid lengkap (Date atau string dengan bulan)
  if (lastCell != null) {
    var formattedFromCell = formatMonthYearFromCell_(lastCell, tz);
    if (formattedFromCell) {
      return formattedFromCell;
    }
  }

  // 2. Tentukan labels tren dari Card Analitik Tren Tahunan
  var trendLabels = [];
  if (rawKPI && rawKPI.trendLabels && rawKPI.trendLabels.length > 0) {
    trendLabels = rawKPI.trendLabels;
  } else if (analytics && analytics.monthlyTrend && analytics.monthlyTrend.labels && analytics.monthlyTrend.labels.length > 0) {
    trendLabels = analytics.monthlyTrend.labels;
  }

  if (trendLabels.length === 0) return '';

  // 3. Ambil label terakhir dari tren tahunan
  var lastLabel = trendLabels[trendLabels.length - 1];
  if (!lastLabel) return '';

  var lastStr = lastLabel.toString().trim();

  // Cek apakah label terakhir sendiri sudah berupa format tanggal / Bulan Tahun
  var formattedDirect = formatMonthYearFromCell_(lastStr, tz);
  if (formattedDirect) {
    return formattedDirect;
  }

  // 4. Ekstrak 4-digit tahun dari label terakhir Card Analitik Tren Tahunan
  var yearMatch = lastStr.match(/\b(19\d\d|20\d\d)\b/);
  var targetYear = yearMatch ? parseInt(yearMatch[1], 10) : parseInt(lastStr, 10);
  if (isNaN(targetYear) || targetYear < 1900) {
    return lastStr;
  }

  var targetYearStr = String(targetYear);

  // 5. Cek apakah data bulan tersedia untuk tahun tersebut
  var foundMonthIdx = -1;

  // 5a. Dari analytics.lastUpdate (Date object terbaru dari sheet data)
  if (analytics && analytics.lastUpdate instanceof Date && !isNaN(analytics.lastUpdate.getTime())) {
    var luYear = Utilities.formatDate(analytics.lastUpdate, tz, 'yyyy');
    if (luYear === targetYearStr) {
      foundMonthIdx = parseInt(Utilities.formatDate(analytics.lastUpdate, tz, 'M'), 10) - 1;
    }
  }

  // 5b. Dari findLatestUpdateDate_ (format string 'dd Bulan yyyy')
  if (foundMonthIdx < 0 && headers && rows) {
    var latestStr = findLatestUpdateDate_(headers, rows, tz);
    if (latestStr) {
      var parts = latestStr.split(' ');
      if (parts.length === 3 && parts[2] === targetYearStr) {
        var bIdx = BULAN.indexOf(parts[1]);
        if (bIdx >= 0) foundMonthIdx = bIdx;
      }
    }
  }

  // 5c. Dari analytics.monthlyIncoming jika tahun berjalan cocok
  if (foundMonthIdx < 0 && analytics && analytics.monthlyIncoming) {
    var curYear = new Date().getFullYear();
    if (curYear === targetYear && analytics.monthlyIncoming.currentMonthName) {
      var bIdx = BULAN.indexOf(analytics.monthlyIncoming.currentMonthName);
      if (bIdx >= 0) foundMonthIdx = bIdx;
    }
  }

  // 6. Return format 'Bulan Tahun' jika bulan tersedia, atau 'Tahun' jika tidak
  if (foundMonthIdx >= 0 && foundMonthIdx < 12) {
    return BULAN[foundMonthIdx] + ' ' + targetYearStr;
  }

  return targetYearStr;
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET RESOLVER
// ═══════════════════════════════════════════════════════════════════

/**
 * resolveSheet_ — Tentukan sheet sumber data.
 * Prioritas: "Pelacakan SK" → sheet aktif.
 */
function resolveSheet_(ss) {
  var sheet = ss.getSheetByName('Pelacakan SK');
  if (sheet) return sheet;

  // Coba beberapa nama alternatif
  var altNames = ['Data SK', 'Tracking SK', 'Pelacakan', 'Data'];
  for (var i = 0; i < altNames.length; i++) {
    sheet = ss.getSheetByName(altNames[i]);
    if (sheet) return sheet;
  }

  // Fallback ke sheet aktif
  return ss.getActiveSheet();
}

/**
 * getInventarisSheet_ — Cari sheet 'Inventaris 2019 - Sekarang' sebagai master data.
 */
function getInventarisSheet_(ss) {
  var targetSheet = ss.getSheetByName('Inventaris 2019 - Sekarang');
  if (!targetSheet) {
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      var n = sheets[s].getName().trim().toLowerCase();
      if (n === 'inventaris 2019 - sekarang' || n === 'inventaris 2019-sekarang') {
        targetSheet = sheets[s];
        break;
      }
    }
    if (!targetSheet) {
      for (var s = 0; s < sheets.length; s++) {
        var n = sheets[s].getName().trim().toLowerCase();
        if (n.indexOf('inventaris') !== -1 && n.indexOf('2019') !== -1) {
          targetSheet = sheets[s];
          break;
        }
      }
    }
    if (!targetSheet) {
      for (var s = 0; s < sheets.length; s++) {
        var n = sheets[s].getName().trim().toLowerCase();
        if (n.indexOf('inventaris') !== -1) {
          targetSheet = sheets[s];
          break;
        }
      }
    }
  }
  return targetSheet;
}

/**
 * extractYearFromCell_ — Ekstrak 4-digit tahun dari nilai sel (Date, number, string).
 */
function extractYearFromCell_(cellVal, tz) {
  if (cellVal === '' || cellVal === null || cellVal === undefined) return null;
  if (cellVal instanceof Date) {
    if (isNaN(cellVal.getTime())) return null;
    var y = parseInt(Utilities.formatDate(cellVal, tz, 'yyyy'), 10);
    return (y >= 1900 && y <= 2100) ? y : null;
  }
  if (typeof cellVal === 'number' && !isNaN(cellVal) && cellVal > 30000 && cellVal < 60000) {
    var dSerial = new Date(Math.round((cellVal - 25569) * 86400 * 1000));
    if (!isNaN(dSerial.getTime())) {
      var y = parseInt(Utilities.formatDate(dSerial, tz, 'yyyy'), 10);
      return (y >= 1900 && y <= 2100) ? y : null;
    }
  }
  var s = cellVal.toString().trim();
  if (!s) return null;
  var match = s.match(/\b(19\d\d|20\d\d)\b/);
  if (match) {
    var y = parseInt(match[1], 10);
    return (y >= 1900 && y <= 2100) ? y : null;
  }
  var d = parseDate_(cellVal);
  if (d && !isNaN(d.getTime())) {
    var y = parseInt(Utilities.formatDate(d, tz, 'yyyy'), 10);
    return (y >= 1900 && y <= 2100) ? y : null;
  }
  return null;
}

/**
 * extractMonthFromCell_ — Ekstrak nomor bulan (1-12) dari nilai sel tanggal.
 */
function extractMonthFromCell_(cellVal, tz) {
  if (cellVal === '' || cellVal === null || cellVal === undefined) return null;
  if (cellVal instanceof Date) {
    if (isNaN(cellVal.getTime())) return null;
    return parseInt(Utilities.formatDate(cellVal, tz, 'M'), 10); // 1-12
  }
  if (typeof cellVal === 'number' && !isNaN(cellVal) && cellVal > 30000 && cellVal < 60000) {
    var dSerial = new Date(Math.round((cellVal - 25569) * 86400 * 1000));
    if (!isNaN(dSerial.getTime())) {
      return parseInt(Utilities.formatDate(dSerial, tz, 'M'), 10);
    }
  }
  var s = cellVal.toString().trim();
  if (!s) return null;
  if (s.indexOf('/') !== -1) {
    var p = s.split('/');
    if (p.length === 3) {
      var p0 = parseInt(p[0].trim(), 10);
      var p1 = parseInt(p[1].trim(), 10);
      if (p[0].trim().length === 4) return p1; // yyyy/MM/dd
      if (p0 > 12) return p1; // dd/MM/yyyy
      if (p1 > 12) return p0; // MM/dd/yyyy
      return p1; // default Indonesia dd/MM/yyyy
    }
  }
  if (s.indexOf('-') !== -1) {
    var p = s.split('-');
    if (p.length === 3) {
      return parseInt(p[1].trim(), 10);
    }
  }
  var sLower = s.toLowerCase();
  for (var b = 0; b < BULAN.length; b++) {
    if (sLower.indexOf(BULAN[b].toLowerCase()) !== -1) {
      return b + 1;
    }
  }
  var d = parseDate_(cellVal);
  if (d && !isNaN(d.getTime())) {
    return parseInt(Utilities.formatDate(d, tz, 'M'), 10);
  }
  return null;
}

/**
 * getAvailableYears_ — Ambil daftar tahun unik dari kolom E di sheet Inventaris.
 * Hanya mencakup tahun yang valid (>= 2019) sesuai nama sheet 'Inventaris 2019 - Sekarang'.
 */
function getAvailableYears_(sheet, tz) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // Kolom E (kolom ke-5, 1-based) dari baris 2 sampai lastRow
  var colEValues = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
  var yearSet = {};
  for (var i = 0; i < colEValues.length; i++) {
    var y = extractYearFromCell_(colEValues[i][0], tz);
    // Hanya sertakan tahun >= 2019 (sesuai cakupan sheet Inventaris 2019 - Sekarang)
    // Ini juga menghapus tahun tidak valid seperti 2001 yang berasal dari serial date spreadsheet
    if (y && y >= 2019) yearSet[y] = true;
  }
  var years = Object.keys(yearSet).map(function(y) { return parseInt(y, 10); });
  years.sort(function(a, b) { return b - a; });
  return years;
}

// ═══════════════════════════════════════════════════════════════════
//  YEARLY TREND FROM INVENTARIS KOLOM E
// ═══════════════════════════════════════════════════════════════════

/**
 * computeYearlyTrendFromInventaris_ — Hitung Analitik Tren Tahunan secara langsung
 * dari kolom E sheet 'Inventaris 2019 - Sekarang'.
 *
 * Logika:
 *   Untuk setiap baris, baca tanggal pada kolom E → ekstrak tahun.
 *   Kelompokkan dan hitung jumlah berkas per tahun:
 *     Jumlah Berkas Tahun X = COUNT seluruh baris kolom E yang tanggalnya memiliki tahun X
 *   Urutkan hasil berdasarkan tahun ascending (untuk chart).
 *   Tidak menggunakan referensi cell rekap, KPI, atau hasil perhitungan sebelumnya.
 *
 * @param {Sheet}  sheet — Sheet 'Inventaris 2019 - Sekarang'.
 * @param {string} tz    — Timezone string dari Session.getScriptTimeZone().
 * @returns {{ labels: string[], values: number[], totalYears: number }}
 *          labels: array tahun sebagai string (ascending), e.g. ['2019','2020',...]
 *          values: jumlah berkas per tahun sesuai urutan labels.
 *          totalYears: jumlah tahun berbeda yang ditemukan.
 */
function computeYearlyTrendFromInventaris_(sheet, tz) {
  if (!sheet) return { labels: [], values: [], totalYears: 0 };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { labels: [], values: [], totalYears: 0 };

  // Baca seluruh kolom E (kolom ke-5, 1-based) dari baris 2 hingga lastRow
  var colEValues = sheet.getRange(2, 5, lastRow - 1, 1).getValues();

  // yearCountMap: { tahun (number) → jumlah berkas }
  var yearCountMap = {};

  for (var i = 0; i < colEValues.length; i++) {
    var cellVal = colEValues[i][0];

    // Skip sel kosong
    if (cellVal === '' || cellVal === null || cellVal === undefined) continue;

    // Ekstrak tahun dari nilai sel
    var y = extractYearFromCell_(cellVal, tz);

    // Hanya sertakan tahun >= 2019 (sesuai cakupan sheet)
    if (!y || y < 2019) continue;

    // Hitung berkas untuk tahun ini
    yearCountMap[y] = (yearCountMap[y] || 0) + 1;
  }

  // Ambil daftar tahun dan urutkan ascending untuk tampilan chart
  var sortedYears = Object.keys(yearCountMap).map(function(y) { return parseInt(y, 10); });
  sortedYears.sort(function(a, b) { return a - b; });

  var labels = sortedYears.map(function(y) { return String(y); });
  var values = sortedYears.map(function(y) { return yearCountMap[y]; });

  return {
    labels: labels,
    values: values,
    totalYears: sortedYears.length
  };
}

// ═══════════════════════════════════════════════════════════════════
//  ANALYTICS ENGINE
// ═══════════════════════════════════════════════════════════════════

/**
 * computeAnalytics_ — Hitung semua metrik dashboard dari data mentah.
 *
 * @param {string[]} headers — Header kolom.
 * @param {Array[]}  rows    — Data baris (2D array).
 * @returns {Object} Analytics object lengkap.
 */
function computeAnalytics_(headers, rows) {
  // ── Deteksi kolom ──────────────────────────────────────────────
  var noIdx    = colIndex_(headers, ['NO.', 'NO ', 'NOMOR']);
  var judulIdx = colIndex_(headers, ['JUDUL', 'NAMA KEPUTUSAN', 'PERIHAL']);
  var statusIdx = colIndex_(headers, ['STATUS PENGERJAAN', 'STATUS']);
  var jenisIdx  = colIndex_(headers, ['JENIS SK', 'JENIS KEPUTUSAN', 'JENIS']);

  // Tanggal masuk & selesai — cari semua kolom TANGGAL lalu bedakan
  var dateCols = colIndexAll_(headers, 'TANGGAL');
  var tglMasukIdx = -1, tglSelesaiIdx = -1;

  for (var i = 0; i < dateCols.length; i++) {
    var h = (headers[dateCols[i]] || '').toString().toUpperCase();
    if (h.includes('MASUK') || h.includes('DITERIMA') || h.includes('TERIMA') || h.includes('TERIMA')) {
      tglMasukIdx = dateCols[i];
    } else if (h.includes('SELESAI') || h.includes('DITETAPKAN') || h.includes('TETAP') || h.includes('KELUAR')) {
      tglSelesaiIdx = dateCols[i];
    }
  }
  // Fallback: juga cek keyword bahasa Inggris
  if (tglMasukIdx === -1) tglMasukIdx = colIndex_(headers, ['TANGGAL MASUK', 'DATE IN', 'TANGGAL_INPUT', 'TANGGAL TERIMA']);
  if (tglSelesaiIdx === -1) tglSelesaiIdx = colIndex_(headers, ['TANGGAL SELESAI', 'DATE DONE', 'TANGGAL_TETAP', 'TANGGAL DITETAPKAN']);
  // Last resort: kolom tanggal pertama & terakhir
  if (tglMasukIdx === -1 && dateCols.length >= 1) tglMasukIdx = dateCols[0];
  if (tglSelesaiIdx === -1 && dateCols.length >= 2) tglSelesaiIdx = dateCols[dateCols.length - 1];

  // Kolom "last update" — cari spesifik, fallback ke tglMasuk
  var updateIdx = colIndex_(headers, ['LAST UPDATE', 'UPDATE', 'TERAKHIR DIPERBARUI', 'TANGGAL UPDATE', 'TANGGAL PERUBAHAN']);

  // Perangkat daerah
  var deptIdx = colIndex_(headers, [
    'PERANGKAT DAERAH', 'PERANGKAT', 'INSTANSI', 'SKPD', 'OPD',
    'ORGANISASI', 'DINAS', 'BADAN', 'BIRO', 'KANTOR',
    'UNIT KERJA', 'LEMBAGA', 'DEPARTEMEN'
  ]);

  // ── Inisialisasi accumulator ───────────────────────────────────
  var counts = { total: 0, done: 0, proc: 0, pend: 0, rej: 0 };
  var typeMap = {};       // jenis SK → count
  var monthMap = {};      // 'YYYY-MM' → count
  var procDays = [];      // array selisih hari
  var deptMap = {};       // dept → { gubernur, sekda, total }
  var latestDate = null;  // tanggal terbaru untuk "last update"

  var now = new Date();
  var curYear = now.getFullYear();
  var curMonth = now.getMonth(); // 0-based

  // ── Iterasi data ───────────────────────────────────────────────
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];

    // Skip baris kosong / header berulang
    var noVal  = noIdx >= 0 ? row[noIdx] : (r + 1);
    var judVal = judulIdx >= 0 ? row[judulIdx] : true;
    if (!noVal || !judVal) continue;
    var noStr = (noVal || '').toString().toUpperCase();
    if (noStr.includes('URAIAN') || noStr.includes('JUMLAH') || noStr.includes('TOTAL')) continue;

    counts.total++;

    // ── Status ───────────────────────────────────────────────────
    var status = statusIdx >= 0 ? normStatus_(row[statusIdx]) : '';
    if (status === 'done' || status === 'selesai' || status === 'completed') {
      counts.done++;
    } else if (status.includes('progres') || status.includes('process') || status.includes('proses')) {
      counts.proc++;
    } else if (status === 'pending' || status === 'tertunda') {
      counts.pend++;
    } else if (status === 'rejected' || status === 'ditolak') {
      counts.rej++;
    }

    // ── Jenis keputusan ──────────────────────────────────────────
    var jenis = jenisIdx >= 0 ? (row[jenisIdx] || '').toString().trim() : '';
    if (jenis) {
      var jl = jenis.toLowerCase();
      var key;
      if (jl.includes('gubernur')) key = 'SK Gubernur';
      else if (jl.includes('sekda') || jl.includes('sekretaris daerah')) key = 'SK Sekda';
      else key = jenis.length > 35 ? jenis.substring(0, 35) + '…' : jenis;
      typeMap[key] = (typeMap[key] || 0) + 1;
    }

    // ── Tanggal masuk → monthly trend ────────────────────────────
    var dMasuk = tglMasukIdx >= 0 ? parseDate_(row[tglMasukIdx]) : null;
    if (dMasuk) {
      var mk = dMasuk.getFullYear() + '-' + String(dMasuk.getMonth() + 1).padStart(2, '0');
      monthMap[mk] = (monthMap[mk] || 0) + 1;
    }

    // ── Processing time (completed only) ─────────────────────────
    var dSelesai = tglSelesaiIdx >= 0 ? parseDate_(row[tglSelesaiIdx]) : null;
    if (dMasuk && dSelesai && (status === 'done' || status === 'selesai' || status === 'completed')) {
      var days = Math.round((dSelesai.getTime() - dMasuk.getTime()) / 86400000);
      if (days >= 0) procDays.push(days);
    }

    // ── Department ───────────────────────────────────────────────
    var dept = deptIdx >= 0 ? (row[deptIdx] || '').toString().trim() : '';
    if (dept) {
      if (!deptMap[dept]) deptMap[dept] = { gubernur: 0, sekda: 0, total: 0 };
      deptMap[dept].total++;
      if (jenis && jenis.toLowerCase().includes('gubernur')) deptMap[dept].gubernur++;
      else if (jenis && (jenis.toLowerCase().includes('sekda') || jenis.toLowerCase().includes('sekretaris daerah'))) deptMap[dept].sekda++;
    }

    // ── Last update detection ────────────────────────────────────
    var updateVal = updateIdx >= 0 ? parseDate_(row[updateIdx]) : null;
    if (!updateVal && tglMasukIdx >= 0) updateVal = dMasuk; // fallback
    if (updateVal && (!latestDate || updateVal > latestDate)) latestDate = updateVal;
  }

  // ── Post: Monthly average ──────────────────────────────────────
  var curKey = curYear + '-' + String(curMonth + 1).padStart(2, '0');
  var curMonthCount = monthMap[curKey] || 0;
  var activeMo = 0, totalYear = 0;
  for (var k in monthMap) {
    if (k.indexOf(String(curYear)) === 0) { activeMo++; totalYear += monthMap[k]; }
  }
  var monthAvg = activeMo > 0 ? Math.round(totalYear / activeMo) : 0;

  // ── Post: Average processing time ──────────────────────────────
  var avgProc = 0;
  if (procDays.length > 0) {
    var sum = 0;
    for (var p = 0; p < procDays.length; p++) sum += procDays[p];
    avgProc = Math.round(sum / procDays.length);
  }

  // ── Post: Monthly trend (sorted) ──────────────────────────────
  var sortedKeys = Object.keys(monthMap).sort();
  var tLabels = [], tValues = [];
  for (var m = 0; m < sortedKeys.length; m++) {
    var pp = sortedKeys[m].split('-');
    tLabels.push(BULAN[parseInt(pp[1], 10) - 1]);
    tValues.push(monthMap[sortedKeys[m]]);
  }

  // ── Post: Department ranking (sorted desc) ────────────────────
  var dEntries = [];
  for (var dn in deptMap) {
    dEntries.push({
      name: dn.length > 45 ? dn.substring(0, 45) + '…' : dn,
      gubernur: deptMap[dn].gubernur,
      sekda: deptMap[dn].sekda,
      total: deptMap[dn].total
    });
  }
  dEntries.sort(function(a, b) { return b.total - a.total; });
  var dLabels = [], dGov = [], dSek = [];
  for (var d = 0; d < dEntries.length; d++) {
    dLabels.push(dEntries[d].name);
    dGov.push(dEntries[d].gubernur);
    dSek.push(dEntries[d].sekda);
  }

  // ── Return analytics ───────────────────────────────────────────
  return {
    statusCounts: { total: counts.total, done: counts.done, proc: counts.proc, pend: counts.pend, rej: counts.rej },
    monthlyIncoming: { currentMonthCount: curMonthCount, average: monthAvg, activeMonths: activeMo, currentMonthName: BULAN[curMonth] },
    processingTime: { average: avgProc, sampleSize: procDays.length },
    decisionTypes: { labels: Object.keys(typeMap), values: Object.keys(typeMap).map(function(x) { return typeMap[x]; }) },
    statusMonitoring: { pending: counts.pend, onProgress: counts.proc, completed: counts.done, rejected: counts.rej },
    monthlyTrend: { labels: tLabels, values: tValues },
    departmentRanking: { labels: dLabels, gubernur: dGov, sekda: dSek },
    lastUpdate: latestDate
  };
}

/**
 * computeYearlyAnalytics_ — Agregasi metrik untuk periode tahunan 12 bulan (Januari–Desember)
 * berdasarkan data baris sheet Inventaris 2019 - Sekarang yang telah difilter berdasarkan tahun pada kolom E.
 *
 * @param {string[]} headers    — Header kolom dari sheet Inventaris.
 * @param {Array[]}  rows       — Baris data yang sudah dicocokkan dengan targetYear pada kolom E.
 * @param {number}   targetYear — Tahun terpilih.
 * @param {string}   tz         — Timezone string.
 * @returns {Object} Analytics object lengkap untuk halaman Data Tahunan.
 */
function computeYearlyAnalytics_(headers, rows, targetYear, tz) {
  var noIdx     = colIndex_(headers, ['NO.', 'NO ', 'NOMOR']);
  var judulIdx  = colIndex_(headers, ['JUDUL', 'NAMA KEPUTUSAN', 'PERIHAL']);
  var jenisIdx  = colIndex_(headers, ['JENIS SK', 'JENIS KEPUTUSAN', 'JENIS']);

  // Kolom G (0-based index 6) adalah kolom STATUS di sheet Inventaris 2019 - Sekarang.
  // Hardcode ke kolom G sesuai struktur sheet master, tidak menggunakan deteksi header dinamis
  // agar status tidak salah dipetakan ke kolom lain.
  var statusIdx = 6; // Kolom G (0-based)

  // Kolom E adalah kolom tanggal masuk utama di sheet Inventaris
  var tglMasukIdx = 4; // Kolom E (0-based)
  var dateCols = colIndexAll_(headers, 'TANGGAL');
  var tglSelesaiIdx = -1;
  for (var i = 0; i < dateCols.length; i++) {
    var h = (headers[dateCols[i]] || '').toString().toUpperCase();
    if (h.includes('SELESAI') || h.includes('DITETAPKAN') || h.includes('TETAP') || h.includes('KELUAR')) {
      tglSelesaiIdx = dateCols[i];
    }
  }
  if (tglSelesaiIdx === -1) {
    tglSelesaiIdx = colIndex_(headers, ['TANGGAL SELESAI', 'DATE DONE', 'TANGGAL_TETAP', 'TANGGAL DITETAPKAN']);
  }

  var deptIdx = colIndex_(headers, [
    'PERANGKAT DAERAH', 'PERANGKAT', 'INSTANSI', 'SKPD', 'OPD',
    'ORGANISASI', 'DINAS', 'BADAN', 'BIRO', 'KANTOR',
    'UNIT KERJA', 'LEMBAGA', 'DEPARTEMEN'
  ]);

  var counts = { total: 0, done: 0, proc: 0, pend: 0, rej: 0 };
  var typeMap = { 'SK Gubernur': 0, 'SK Sekda': 0 };

  // Inisialisasi 12 bulan (Januari = 1 s.d. Desember = 12) dengan 0
  var monthCounts = {};
  for (var m = 1; m <= 12; m++) {
    monthCounts[m] = 0;
  }
  var procDays = [];
  var deptMap = {};
  var latestDateInYear = null;

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];

    var noVal  = noIdx >= 0 ? row[noIdx] : (r + 1);
    var judVal = judulIdx >= 0 ? row[judulIdx] : true;
    if (!noVal || !judVal) continue;
    var noStr = (noVal || '').toString().toUpperCase();
    if (noStr.includes('URAIAN') || noStr.includes('JUMLAH') || noStr.includes('TOTAL')) continue;

    counts.total++;

    // ── Status ──
    var status = statusIdx >= 0 ? normStatus_(row[statusIdx]) : '';
    if (status === 'done' || status === 'selesai' || status === 'completed') {
      counts.done++;
    } else if (status.includes('progres') || status.includes('process') || status.includes('proses')) {
      counts.proc++;
    } else if (status === 'pending' || status === 'tertunda') {
      counts.pend++;
    } else if (status === 'rejected' || status === 'ditolak') {
      counts.rej++;
    }

    // ── Jenis keputusan ──
    var jenis = jenisIdx >= 0 ? (row[jenisIdx] || '').toString().trim() : '';
    if (jenis) {
      var jl = jenis.toLowerCase();
      if (jl.includes('gubernur')) typeMap['SK Gubernur']++;
      else if (jl.includes('sekda') || jl.includes('sekretaris daerah')) typeMap['SK Sekda']++;
      else {
        var key = jenis.length > 35 ? jenis.substring(0, 35) + '…' : jenis;
        typeMap[key] = (typeMap[key] || 0) + 1;
      }
    }

    // ── Agregasi 12 bulan dari Kolom E ──
    var cellDateVal = row[tglMasukIdx];
    var mNum = extractMonthFromCell_(cellDateVal, tz);
    if (mNum >= 1 && mNum <= 12) {
      monthCounts[mNum]++;
    }

    // ── Processing time (completed only) ──
    var dMasuk = parseDate_(cellDateVal);
    var dSelesai = tglSelesaiIdx >= 0 ? parseDate_(row[tglSelesaiIdx]) : null;
    if (dMasuk && dSelesai && (status === 'done' || status === 'selesai' || status === 'completed')) {
      var days = Math.round((dSelesai.getTime() - dMasuk.getTime()) / 86400000);
      if (days >= 0) procDays.push(days);
    }

    // ── Department ──
    var dept = deptIdx >= 0 ? (row[deptIdx] || '').toString().trim() : '';
    if (dept) {
      if (!deptMap[dept]) deptMap[dept] = { gubernur: 0, sekda: 0, total: 0 };
      deptMap[dept].total++;
      if (jenis && jenis.toLowerCase().includes('gubernur')) deptMap[dept].gubernur++;
      else if (jenis && (jenis.toLowerCase().includes('sekda') || jenis.toLowerCase().includes('sekretaris daerah'))) deptMap[dept].sekda++;
    }

    // ── Latest date in year ──
    if (dMasuk && (!latestDateInYear || dMasuk > latestDateInYear)) {
      latestDateInYear = dMasuk;
    }
  }

  // Rata-rata per bulan dalam 12 bulan (Januari - Desember)
  var avgMonthly = counts.total > 0 ? Math.round(counts.total / 12) : 0;

  // Rata-rata waktu pengerjaan
  var avgProc = 0;
  if (procDays.length > 0) {
    var sum = 0;
    for (var p = 0; p < procDays.length; p++) sum += procDays[p];
    avgProc = Math.round(sum / procDays.length);
  }

  // Tren 12 Bulan (Januari - Desember)
  var monthlyTrendValues = [];
  for (var m = 1; m <= 12; m++) {
    monthlyTrendValues.push(monthCounts[m]);
  }

  // Ranking OPD
  var dEntries = [];
  for (var dn in deptMap) {
    dEntries.push({
      name: dn.length > 45 ? dn.substring(0, 45) + '…' : dn,
      gubernur: deptMap[dn].gubernur,
      sekda: deptMap[dn].sekda,
      total: deptMap[dn].total
    });
  }
  dEntries.sort(function(a, b) { return b.total - a.total; });
  var dLabels = [], dGov = [], dSek = [];
  for (var d = 0; d < dEntries.length; d++) {
    dLabels.push(dEntries[d].name);
    dGov.push(dEntries[d].gubernur);
    dSek.push(dEntries[d].sekda);
  }

  // Format last update untuk tahun terpilih
  var lastUpdateStr = '';
  if (latestDateInYear) {
    var mIdx = parseInt(Utilities.formatDate(latestDateInYear, tz, 'M'), 10) - 1;
    lastUpdateStr = BULAN[mIdx] + ' ' + targetYear;
  } else {
    lastUpdateStr = String(targetYear);
  }

  return {
    statusCounts: counts,
    monthlyIncoming: {
      average: avgMonthly,
      totalYear: counts.total,
      targetYear: targetYear,
      period: 'Januari - Desember'
    },
    processingTime: { average: avgProc, sampleSize: procDays.length },
    decisionTypes: {
      labels: Object.keys(typeMap),
      values: Object.keys(typeMap).map(function(k) { return typeMap[k]; }),
      total: counts.total
    },
    statusMonitoring: {
      completed: counts.done,
      onProgress: counts.proc,
      pending: counts.pend,
      rejected: counts.rej
    },
    monthlyTrend: {
      labels: BULAN.slice(), // 12 bulan penuh: Januari s.d. Desember
      values: monthlyTrendValues
    },
    departmentRanking: { labels: dLabels, gubernur: dGov, sekda: dSek },
    lastUpdateStr: lastUpdateStr
  };
}

// ═══════════════════════════════════════════════════════════════════
//  RAWDATA KPI READER
// ═══════════════════════════════════════════════════════════════════

/**
 * readKPIFromRawData_ — Baca nilai KPI dari sheet "Raw Data".
 *
 * Mapping cell:
 *   B8  → On Progress
 *   B9  → Pending
 *   B10 → Rejected SK
 *   B11 → Completed SK
 *   B12 → Total Received
 *   D8  → Rata-rata SK Masuk/Tahun
 *   D9  → Rata-rata Lama Pengerjaan SK (hari)
 *   D10 → Jumlah tahun tren data
 *   A17:B30 → Data tren tahunan (label tahun, jumlah berkas)
 *
 * @param {Spreadsheet} ss — Active spreadsheet.
 * @returns {Object|null} KPI values atau null jika sheet tidak ada.
 */
function readKPIFromRawData_(ss) {
  var sheet = ss.getSheetByName('Raw Data') || ss.getSheetByName('raw data') || ss.getSheetByName('RAW DATA');
  if (!sheet) {
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      if (sheets[s].getName().trim().toLowerCase() === 'raw data') {
        sheet = sheets[s];
        break;
      }
    }
  }
  if (!sheet) return null;

  // ── Jenis Keputusan dari cell B1:B3 ────────────────────────────
  // B1 → Total Jumlah
  // B2 → Gubernur
  // B3 → Sekda
  var totalDecision = parseInt(sheet.getRange('B1').getValue(), 10) || 0;
  var gubernurDecision = parseInt(sheet.getRange('B2').getValue(), 10) || 0;
  var sekdaDecision = parseInt(sheet.getRange('B3').getValue(), 10) || 0;
  if (totalDecision <= 0) {
    totalDecision = gubernurDecision + sekdaDecision;
  }

  // ── Status KPI dari cell B8:B12 ─────────────────────────────────
  var proc = parseInt(sheet.getRange('B8').getValue(), 10) || 0;
  var pend = parseInt(sheet.getRange('B9').getValue(), 10) || 0;
  var rej  = parseInt(sheet.getRange('B10').getValue(), 10) || 0;
  var done = parseInt(sheet.getRange('B11').getValue(), 10) || 0;
  var total = parseInt(sheet.getRange('B12').getValue(), 10) || 0;

  // Rata-rata KPI dari kolom D
  var avgMonthly  = parseFloat(sheet.getRange('D8').getValue()) || 0;
  var avgProcTime = parseFloat(sheet.getRange('D9').getValue()) || 0;
  var trendYears  = parseInt(sheet.getRange('D10').getValue(), 10) || 0;

  // Trend data tahunan dari A17:B30
  var trendLabels = [];
  var trendValues = [];
  var lastTrendCell = null;
  try {
    var trendRange = sheet.getRange('A17:B30').getValues();
    for (var i = 0; i < trendRange.length; i++) {
      var yearLabel = trendRange[i][0];
      var yearValue = trendRange[i][1];
      // Skip empty rows
      if (yearLabel === '' || yearLabel === null || yearLabel === undefined) continue;
      var labelStr = '';
      if (yearLabel instanceof Date) {
        labelStr = Utilities.formatDate(yearLabel, Session.getScriptTimeZone(), 'yyyy');
      } else {
        labelStr = yearLabel.toString().trim();
      }
      if (labelStr === '') continue;
      trendLabels.push(labelStr);
      trendValues.push(parseInt(yearValue, 10) || 0);
      lastTrendCell = yearLabel;
    }
  } catch (e) {
    console.error('Error reading trend data from A17:B30:', e);
  }

  // Data ranking perangkat daerah dari F3:H53
  var deptLabels = [];
  var deptGubernur = [];
  var deptSekda = [];
  try {
    var deptRange = sheet.getRange('F3:H53').getValues();
    for (var i = 0; i < deptRange.length; i++) {
      var opdName = deptRange[i][0];
      var gubernurVal = deptRange[i][1];
      var sekdaVal = deptRange[i][2];
      // Skip empty rows
      if (opdName === '' || opdName === null || opdName === undefined) continue;
      var opdStr = opdName.toString().trim();
      if (opdStr === '') continue;
      deptLabels.push(opdStr);
      deptGubernur.push(parseInt(gubernurVal, 10) || 0);
      deptSekda.push(parseInt(sekdaVal, 10) || 0);
    }
  } catch (e) {
    console.error('Error reading department data from F3:H53:', e);
  }

  // Fallback: jika Total Received kosong/tidak valid, hitung dari komponen
  if (total <= 0) {
    total = done + proc + pend + rej;
  }

  return {
    totalDecision: totalDecision,
    gubernurDecision: gubernurDecision,
    sekdaDecision: sekdaDecision,
    done: done, proc: proc, pend: pend, rej: rej,
    total: total,
    avgMonthly: avgMonthly,
    avgProcTime: avgProcTime,
    trendYears: trendYears,
    trendLabels: trendLabels,
    trendValues: trendValues,
    lastTrendCell: lastTrendCell,
    deptLabels: deptLabels,
    deptGubernur: deptGubernur,
    deptSekda: deptSekda
  };
}

// ═══════════════════════════════════════════════════════════════════
//  PUBLIC API
// ═══════════════════════════════════════════════════════════════════

/**
 * emptyResponse_ — Respons kosong yang valid (tidak pernah null).
 */
function emptyResponse_(msg) {
  return { analytics: null, lastUpdate: '', error: msg || null };
}

/**
 * getDashboardData — API utama yang dipanggil frontend via google.script.run.
 * Return: { analytics, lastUpdate, error }
 */
function getDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return emptyResponse_('Spreadsheet tidak ditemukan.');

    var tz = Session.getScriptTimeZone();
    var sheet = resolveSheet_(ss);
    if (!sheet || sheet.getLastRow() < 2) return emptyResponse_('Sheet data kosong atau tidak ditemukan.');

    // ── Deteksi header & baca data ──────────────────────────────
    var info = detectHeaderRow_(sheet);
    var headers = info.headers;
    var dataStartRow = info.row + 1;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < dataStartRow || lastCol < 1) {
      return emptyResponse_('Tidak ada data di bawah baris header.');
    }

    var rows = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol).getValues();

    // ── Hitung analytics ────────────────────────────────────────
    var analytics = computeAnalytics_(headers, rows);

    // ── Override KPI dari sheet Raw Data (cell B1-B3, B8-B12, D8-D10) ──
    var rawKPI = readKPIFromRawData_(ss);
    if (rawKPI) {
      // Override Jenis Keputusan dari B1:B3 (B1: Total, B2: Gubernur, B3: Sekda)
      if (rawKPI.totalDecision > 0 || rawKPI.gubernurDecision > 0 || rawKPI.sekdaDecision > 0) {
        analytics.decisionTypes = {
          total: rawKPI.totalDecision,
          gubernur: rawKPI.gubernurDecision,
          sekda: rawKPI.sekdaDecision,
          labels: ['SK Gubernur', 'SK Sekda'],
          values: [rawKPI.gubernurDecision, rawKPI.sekdaDecision]
        };
      }
      analytics.statusCounts = {
        total: rawKPI.total,
        done:  rawKPI.done,
        proc:  rawKPI.proc,
        pend:  rawKPI.pend,
        rej:   rawKPI.rej
      };
      analytics.statusMonitoring = {
        completed:  rawKPI.done,
        onProgress: rawKPI.proc,
        pending:    rawKPI.pend,
        rejected:   rawKPI.rej
      };
      // Override rata-rata dari D8 & D9
      if (analytics.monthlyIncoming) {
        analytics.monthlyIncoming.average = rawKPI.avgMonthly;
      }
      if (analytics.processingTime) {
        analytics.processingTime.average = rawKPI.avgProcTime;
      }
      // Override ranking perangkat daerah dari F3:H53 (jika ada data)
      if (rawKPI.deptLabels.length > 0) {
        analytics.departmentRanking = {
          labels: rawKPI.deptLabels,
          gubernur: rawKPI.deptGubernur,
          sekda: rawKPI.deptSekda
        };
      }
    }

    // ── Analitik Tren Tahunan — sumber UTAMA: kolom E Inventaris ──
    // Hitung langsung COUNT per tahun dari kolom E sheet Inventaris 2019 - Sekarang.
    // Tidak menggunakan referensi cell rekap, KPI, atau Raw Data A17:B30.
    var inventarisSheetForTrend = getInventarisSheet_(ss);
    var yearlyTrend = computeYearlyTrendFromInventaris_(inventarisSheetForTrend, tz);
    if (yearlyTrend.labels.length > 0) {
      analytics.monthlyTrend = {
        labels: yearlyTrend.labels,
        values: yearlyTrend.values
      };
      analytics.trendYears = yearlyTrend.totalYears;
    }

    // ── Format last update (Dashboard Utama) ──────────────────────
    // Sumber UTAMA: kolom E sheet 'Inventaris 2019 - Sekarang', difilter tahun berjalan.
    // Mengambil tanggal terbaru dari data tahun berjalan pada kolom E.
    var inventarisSheetForUpdate = getInventarisSheet_(ss);
    var lastUpdateStr = getLastUpdateFromInventarisColE_(inventarisSheetForUpdate, tz);

    // Fallback 1: tahun terakhir dari Card Analitik Tren Tahunan (Raw Data A17:B30)
    if (!lastUpdateStr) {
      lastUpdateStr = getLastUpdateFromTrend_(rawKPI, analytics, headers, rows, tz);
    }

    // Fallback 2: scan dinamis kolom 'LAST UPDATE' / 'TANGGAL MASUK' dari data utama
    if (!lastUpdateStr) {
      lastUpdateStr = findLatestUpdateDate_(headers, rows, tz);
      // findLatestUpdateDate_ mengembalikan 'dd Bulan yyyy'; ambil 'Bulan yyyy' saja
      if (lastUpdateStr) {
        var parts = lastUpdateStr.split(' ');
        if (parts.length === 3) lastUpdateStr = parts[1] + ' ' + parts[2];
      }
    }

    // Fallback 3: analytics.lastUpdate dari komputasi analytics
    if (!lastUpdateStr && analytics.lastUpdate) {
      var mIdx = parseInt(Utilities.formatDate(analytics.lastUpdate, tz, 'M'), 10) - 1;
      var yyyy = Utilities.formatDate(analytics.lastUpdate, tz, 'yyyy');
      lastUpdateStr = BULAN[mIdx] + ' ' + yyyy;
    }
    // Hapus field Date object sebelum dikirim ke client (tidak serializable)
    delete analytics.lastUpdate;

    // ── Dapatkan URL spreadsheet untuk tombol Rekap Data ────────
    var spreadsheetUrl = ss.getUrl();

    // ── Ambil daftar tahun dari sheet Inventaris 2019 - Sekarang ──
    var inventarisSheet = getInventarisSheet_(ss);
    var availableYears = [];
    if (inventarisSheet) {
      availableYears = getAvailableYears_(inventarisSheet, tz);
    }

    return { 
      analytics: analytics, 
      lastUpdate: lastUpdateStr, 
      availableYears: availableYears,
      spreadsheetUrl: spreadsheetUrl,
      error: null 
    };

  } catch (e) {
    console.error('getDashboardData error:', e);
    return emptyResponse_('Error server: ' + (e.message || e));
  }
}

/**
 * getYearlyData — API utama untuk halaman khusus Data Tahunan.
 * Mengambil data dari sheet 'Inventaris 2019 - Sekarang',
 * memfilter data berdasarkan tahun pada kolom E,
 * dan menghitung seluruh metrik untuk periode 12 bulan (Januari–Desember).
 *
 * @param {number|string} targetYear — Tahun yang dipilih.
 * @returns {Object} Data tahunan lengkap untuk frontend.
 */
function getYearlyData(targetYear) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { error: 'Spreadsheet tidak ditemukan.' };

    var tz = Session.getScriptTimeZone();
    var sheet = getInventarisSheet_(ss);
    if (!sheet || sheet.getLastRow() < 2) {
      return { error: 'Sheet Inventaris 2019 - Sekarang tidak ditemukan atau kosong.' };
    }

    var info = detectHeaderRow_(sheet);
    var headers = info.headers;
    var dataStartRow = info.row + 1;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < dataStartRow || lastCol < 1) {
      return { error: 'Tidak ada data di sheet Inventaris.' };
    }

    var allRows = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol).getValues();

    // Kolom E (kolom ke-5, 0-based index 4)
    var colEIdx = 4;
    var yearSet = {};
    for (var r = 0; r < allRows.length; r++) {
      var y = extractYearFromCell_(allRows[r][colEIdx], tz);
      // Hanya sertakan tahun >= 2019 (sesuai cakupan sheet Inventaris 2019 - Sekarang)
      // Menghapus tahun tidak valid seperti 2001 yang mungkin berasal dari serial date
      if (y && y >= 2019) yearSet[y] = true;
    }
    var availableYears = Object.keys(yearSet).map(function(y) { return parseInt(y, 10); });
    availableYears.sort(function(a, b) { return b - a; });

    // Tentukan tahun yang dipilih
    var parsedYear = parseInt(targetYear, 10);
    if (isNaN(parsedYear) || parsedYear < 1900) {
      parsedYear = availableYears.length > 0 ? availableYears[0] : new Date().getFullYear();
    }

    // Filter seluruh data dengan mencocokkan tahun pada kolom E dengan tahun terpilih
    var filteredRows = [];
    for (var r = 0; r < allRows.length; r++) {
      var row = allRows[r];
      var rowYear = extractYearFromCell_(row[colEIdx], tz);
      if (rowYear === parsedYear) {
        filteredRows.push(row);
      }
    }

    // Hitung agregasi untuk periode tahunan 12 bulan (Januari–Desember)
    var analytics = computeYearlyAnalytics_(headers, filteredRows, parsedYear, tz);

    return {
      targetYear: parsedYear,
      availableYears: availableYears,
      analytics: analytics,
      lastUpdate: analytics.lastUpdateStr,
      spreadsheetUrl: ss.getUrl(),
      error: null
    };

  } catch (e) {
    console.error('getYearlyData error:', e);
    return { error: 'Gagal memproses data tahunan: ' + (e.message || e) };
  }
}
