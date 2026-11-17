/**
 * backend/Code.gs — API Backend untuk Dashboard SK
 * Google Apps Script Web App (JSON API)
 *
 * Endpoint tersedia via doGet/doPost dengan parameter `action`:
 *   - getDashboardData            → data dashboard utama
 *   - getYearlyData&targetYear=YYYY → data tahunan per tahun
 *   - ping                        → health check
 *
 * Sumber data: Spreadsheet aktif (script terikat / bound).
 * Struktur data, header, dan kolom tetap dideteksi dinamis.
 *
 * CATATAN MIGRASI:
 *   - doGet tidak lagi mengembalikan HTML. Frontend statis berada di Vercel.
 *   - Akses lintas-origin ditangani oleh api/proxy.mjs di Vercel.
 */

// ═══════════════════════════════════════════════════════════════════
//  API ENTRY POINT
// ═══════════════════════════════════════════════════════════════════

function doGet(e) {
  return handleApiRequest_(e);
}

function doPost(e) {
  return handleApiRequest_(e);
}

/**
 * handleApiRequest_ — Router utama API JSON.
 * Menggabungkan parameter query string dan body JSON (jika ada),
 * lalu meneruskan ke handler yang sesuai.
 *
 * @param {Object} e — Event GAS (doGet/doPost).
 * @returns {ContentService.TextOutput} — JSON response.
 */
function handleApiRequest_(e) {
  var params = {};

  // 1. Merge query string parameters
  if (e && e.parameter) {
    for (var k in e.parameter) {
      if (Object.prototype.hasOwnProperty.call(e.parameter, k)) {
        params[k] = e.parameter[k];
      }
    }
  }

  // 2. Merge POST body (JSON)
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      for (var kb in body) {
        if (Object.prototype.hasOwnProperty.call(body, kb)) {
          params[kb] = body[kb];
        }
      }
    } catch (err) {
      console.warn('handleApiRequest_ — body bukan JSON valid:', err);
    }
  }

  var action = params.action || '';
  var result;

  try {
    switch (action) {
      case 'getDashboardData':
        result = getDashboardData();
        break;
      case 'getYearlyData':
        result = getYearlyData(params.targetYear || params.year);
        break;
      case 'ping':
        result = { ok: true, timestamp: new Date().toISOString() };
        break;
      default:
        result = {
          ok: true,
          message: 'Dashboard SK API',
          endpoints: ['getDashboardData', 'getYearlyData', 'ping'],
          usage: 'Tambahkan ?action=<namaEndpoint>[&targetYear=YYYY] pada URL.'
        };
    }
  } catch (err) {
    console.error('API error [' + action + ']:', err);
    result = { error: err.message || String(err) };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * include — Dipertahankan untuk kompatibilitas mundur.
 * Tidak dipakai oleh API mode, tetapi aman jika dipanggil manual.
 */
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
 */
function parseDate_(val) {
  if (!val) return null;

  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }

  var s = val.toString().trim();
  if (!s) return null;

  var p, d;
  if (s.includes('/')) {
    p = s.split('/');
    if (p.length === 3) {
      var day   = parseInt(p[0], 10);
      var month = parseInt(p[1], 10);
      var year  = parseInt(p[2], 10);
      if (day >= 1 && day <= 31 && month >= 1 && month <= 12 && year > 1900) {
        d = new Date(year, month - 1, day);
      } else {
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

function colIndex_(headers, keywords) {
  for (var i = 0; i < headers.length; i++) {
    var h = (headers[i] || '').toString().toUpperCase();
    for (var k = 0; k < keywords.length; k++) {
      if (h.includes(keywords[k].toUpperCase())) return i;
    }
  }
  return -1;
}

function colIndexAll_(headers, keyword) {
  var kw = keyword.toUpperCase();
  var out = [];
  for (var i = 0; i < headers.length; i++) {
    if ((headers[i] || '').toString().toUpperCase().includes(kw)) out.push(i);
  }
  return out;
}

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

function normStatus_(s) {
  return (s || '').toString().toLowerCase().trim();
}

// ═══════════════════════════════════════════════════════════════════
//  FIND LATEST UPDATE DATE
// ═══════════════════════════════════════════════════════════════════

function findLatestUpdateDate_(headers, rows, tz) {
  var updateIdx = colIndex_(headers, [
    'LAST UPDATE', 'LASTUPDATE', 'UPDATE',
    'TERAKHIR DIPERBARUI', 'TANGGAL UPDATE', 'TANGGAL PERUBAHAN',
    'DIPERBARUI', 'PERBARUI'
  ]);

  var fallbackIdx = -1;
  if (updateIdx === -1) {
    fallbackIdx = colIndex_(headers, [
      'TANGGAL MASUK', 'TANGGAL TERIMA', 'TANGGAL DITERIMA',
      'TGL MASUK', 'DATE IN', 'TANGGAL'
    ]);
  }

  var colIdx = updateIdx !== -1 ? updateIdx : fallbackIdx;
  if (colIdx === -1 || rows.length === 0) return '';

  var latestDate = null;
  for (var r = 0; r < rows.length; r++) {
    var cellVal = rows[r][colIdx];
    if (cellVal === '' || cellVal === null || cellVal === undefined) continue;
    var d;
    if (cellVal instanceof Date) {
      d = isNaN(cellVal.getTime()) ? null : cellVal;
    } else {
      d = parseDate_(cellVal);
    }
    if (!d) continue;
    if (!latestDate || d.getTime() > latestDate.getTime()) {
      latestDate = d;
    }
  }

  if (!latestDate) return '';

  var day      = Utilities.formatDate(latestDate, tz, 'dd');
  var monthIdx = parseInt(Utilities.formatDate(latestDate, tz, 'M'), 10) - 1;
  var year     = Utilities.formatDate(latestDate, tz, 'yyyy');
  return day + ' ' + BULAN[monthIdx] + ' ' + year;
}

// ═══════════════════════════════════════════════════════════════════
//  LAST UPDATE FROM INVENTARIS SHEET (KOLOM E)
// ═══════════════════════════════════════════════════════════════════

function getLastUpdateFromInventarisColE_(sheet, tz) {
  if (!sheet) return '';
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';

  var curYear = new Date().getFullYear();
  var colEValues = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
  var latestDate = null;

  for (var i = 0; i < colEValues.length; i++) {
    var cellVal = colEValues[i][0];
    if (cellVal === '' || cellVal === null || cellVal === undefined) continue;

    var d;
    if (cellVal instanceof Date) {
      d = isNaN(cellVal.getTime()) ? null : cellVal;
    } else {
      d = parseDate_(cellVal);
    }
    if (!d) continue;

    var rowYear = parseInt(Utilities.formatDate(d, tz, 'yyyy'), 10);
    if (rowYear !== curYear) continue;

    if (!latestDate || d.getTime() > latestDate.getTime()) {
      latestDate = d;
    }
  }

  if (!latestDate) return '';

  var day      = Utilities.formatDate(latestDate, tz, 'dd');
  var monthIdx = parseInt(Utilities.formatDate(latestDate, tz, 'M'), 10) - 1;
  var year     = Utilities.formatDate(latestDate, tz, 'yyyy');
  return day + ' ' + BULAN[monthIdx] + ' ' + year;
}

function formatMonthYearFromCell_(cellVal, tz) {
  if (cellVal === '' || cellVal === null || cellVal === undefined) return '';

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

  var s = cellVal.toString().trim();
  if (!s) return '';

  if (s.indexOf('/') !== -1) {
    var p = s.split('/');
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

  var sLower = s.toLowerCase();
  for (var b = 0; b < BULAN.length; b++) {
    if (sLower.indexOf(BULAN[b].toLowerCase()) !== -1) {
      var matchYear = s.match(/\b(19\d\d|20\d\d)\b/);
      if (matchYear) {
        return BULAN[b] + ' ' + matchYear[1];
      }
    }
  }

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

function getLastUpdateFromTrend_(rawKPI, analytics, headers, rows, tz) {
  var lastCell = (rawKPI && rawKPI.lastTrendCell != null) ? rawKPI.lastTrendCell : null;

  if (lastCell != null) {
    var formattedFromCell = formatMonthYearFromCell_(lastCell, tz);
    if (formattedFromCell) {
      return formattedFromCell;
    }
  }

  var trendLabels = [];
  if (rawKPI && rawKPI.trendLabels && rawKPI.trendLabels.length > 0) {
    trendLabels = rawKPI.trendLabels;
  } else if (analytics && analytics.monthlyTrend && analytics.monthlyTrend.labels && analytics.monthlyTrend.labels.length > 0) {
    trendLabels = analytics.monthlyTrend.labels;
  }

  if (trendLabels.length === 0) return '';

  var lastLabel = trendLabels[trendLabels.length - 1];
  if (!lastLabel) return '';

  var lastStr = lastLabel.toString().trim();

  var formattedDirect = formatMonthYearFromCell_(lastStr, tz);
  if (formattedDirect) {
    return formattedDirect;
  }

  var yearMatch = lastStr.match(/\b(19\d\d|20\d\d)\b/);
  var targetYear = yearMatch ? parseInt(yearMatch[1], 10) : parseInt(lastStr, 10);
  if (isNaN(targetYear) || targetYear < 1900) {
    return lastStr;
  }

  var targetYearStr = String(targetYear);
  var foundMonthIdx = -1;

  if (analytics && analytics.lastUpdate instanceof Date && !isNaN(analytics.lastUpdate.getTime())) {
    var luYear = Utilities.formatDate(analytics.lastUpdate, tz, 'yyyy');
    if (luYear === targetYearStr) {
      foundMonthIdx = parseInt(Utilities.formatDate(analytics.lastUpdate, tz, 'M'), 10) - 1;
    }
  }

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

  if (foundMonthIdx < 0 && analytics && analytics.monthlyIncoming) {
    var curYear = new Date().getFullYear();
    if (curYear === targetYear && analytics.monthlyIncoming.currentMonthName) {
      var bIdx2 = BULAN.indexOf(analytics.monthlyIncoming.currentMonthName);
      if (bIdx2 >= 0) foundMonthIdx = bIdx2;
    }
  }

  if (foundMonthIdx >= 0 && foundMonthIdx < 12) {
    return BULAN[foundMonthIdx] + ' ' + targetYearStr;
  }

  return targetYearStr;
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET RESOLVER
// ═══════════════════════════════════════════════════════════════════

function resolveSheet_(ss) {
  var sheet = ss.getSheetByName('Pelacakan SK');
  if (sheet) return sheet;

  var altNames = ['Data SK', 'Tracking SK', 'Pelacakan', 'Data'];
  for (var i = 0; i < altNames.length; i++) {
    sheet = ss.getSheetByName(altNames[i]);
    if (sheet) return sheet;
  }

  return ss.getActiveSheet();
}

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
      for (var s2 = 0; s2 < sheets.length; s2++) {
        var n2 = sheets[s2].getName().trim().toLowerCase();
        if (n2.indexOf('inventaris') !== -1 && n2.indexOf('2019') !== -1) {
          targetSheet = sheets[s2];
          break;
        }
      }
    }
    if (!targetSheet) {
      for (var s3 = 0; s3 < sheets.length; s3++) {
        var n3 = sheets[s3].getName().trim().toLowerCase();
        if (n3.indexOf('inventaris') !== -1) {
          targetSheet = sheets[s3];
          break;
        }
      }
    }
  }
  return targetSheet;
}

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
      var yS = parseInt(Utilities.formatDate(dSerial, tz, 'yyyy'), 10);
      return (yS >= 1900 && yS <= 2100) ? yS : null;
    }
  }
  var s = cellVal.toString().trim();
  if (!s) return null;
  var match = s.match(/\b(19\d\d|20\d\d)\b/);
  if (match) {
    var yM = parseInt(match[1], 10);
    return (yM >= 1900 && yM <= 2100) ? yM : null;
  }
  var d = parseDate_(cellVal);
  if (d && !isNaN(d.getTime())) {
    var yP = parseInt(Utilities.formatDate(d, tz, 'yyyy'), 10);
    return (yP >= 1900 && yP <= 2100) ? yP : null;
  }
  return null;
}

function extractMonthFromCell_(cellVal, tz) {
  if (cellVal === '' || cellVal === null || cellVal === undefined) return null;
  if (cellVal instanceof Date) {
    if (isNaN(cellVal.getTime())) return null;
    return parseInt(Utilities.formatDate(cellVal, tz, 'M'), 10);
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
      if (p[0].trim().length === 4) return p1;
      if (p0 > 12) return p1;
      if (p1 > 12) return p0;
      return p1;
    }
  }
  if (s.indexOf('-') !== -1) {
    var p2 = s.split('-');
    if (p2.length === 3) {
      return parseInt(p2[1].trim(), 10);
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

function getAvailableYears_(sheet, tz) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var colEValues = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
  var yearSet = {};
  for (var i = 0; i < colEValues.length; i++) {
    var y = extractYearFromCell_(colEValues[i][0], tz);
    if (y && y >= 2019) yearSet[y] = true;
  }
  var years = Object.keys(yearSet).map(function(y) { return parseInt(y, 10); });
  years.sort(function(a, b) { return b - a; });
  return years;
}

// ═══════════════════════════════════════════════════════════════════
//  YEARLY TREND FROM INVENTARIS KOLOM E
// ═══════════════════════════════════════════════════════════════════

function computeYearlyTrendFromInventaris_(sheet, tz) {
  if (!sheet) return { labels: [], values: [], totalYears: 0 };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { labels: [], values: [], totalYears: 0 };

  var colEValues = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
  var yearCountMap = {};

  for (var i = 0; i < colEValues.length; i++) {
    var cellVal = colEValues[i][0];
    if (cellVal === '' || cellVal === null || cellVal === undefined) continue;
    var y = extractYearFromCell_(cellVal, tz);
    if (!y || y < 2019) continue;
    yearCountMap[y] = (yearCountMap[y] || 0) + 1;
  }

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

function computeAnalytics_(headers, rows) {
  var noIdx     = colIndex_(headers, ['NO.', 'NO ', 'NOMOR']);
  var judulIdx  = colIndex_(headers, ['JUDUL', 'NAMA KEPUTUSAN', 'PERIHAL']);
  var statusIdx = colIndex_(headers, ['STATUS PENGERJAAN', 'STATUS']);
  var jenisIdx  = colIndex_(headers, ['JENIS SK', 'JENIS KEPUTUSAN', 'JENIS']);

  var dateCols = colIndexAll_(headers, 'TANGGAL');
  var tglMasukIdx = -1, tglSelesaiIdx = -1;

  for (var i = 0; i < dateCols.length; i++) {
    var h = (headers[dateCols[i]] || '').toString().toUpperCase();
    if (h.includes('MASUK') || h.includes('DITERIMA') || h.includes('TERIMA')) {
      tglMasukIdx = dateCols[i];
    } else if (h.includes('SELESAI') || h.includes('DITETAPKAN') || h.includes('TETAP') || h.includes('KELUAR')) {
      tglSelesaiIdx = dateCols[i];
    }
  }
  if (tglMasukIdx === -1) tglMasukIdx = colIndex_(headers, ['TANGGAL MASUK', 'DATE IN', 'TANGGAL_INPUT', 'TANGGAL TERIMA']);
  if (tglSelesaiIdx === -1) tglSelesaiIdx = colIndex_(headers, ['TANGGAL SELESAI', 'DATE DONE', 'TANGGAL_TETAP', 'TANGGAL DITETAPKAN']);
  if (tglMasukIdx === -1 && dateCols.length >= 1) tglMasukIdx = dateCols[0];
  if (tglSelesaiIdx === -1 && dateCols.length >= 2) tglSelesaiIdx = dateCols[dateCols.length - 1];

  var updateIdx = colIndex_(headers, ['LAST UPDATE', 'UPDATE', 'TERAKHIR DIPERBARUI', 'TANGGAL UPDATE', 'TANGGAL PERUBAHAN']);

  var deptIdx = colIndex_(headers, [
    'PERANGKAT DAERAH', 'PERANGKAT', 'INSTANSI', 'SKPD', 'OPD',
    'ORGANISASI', 'DINAS', 'BADAN', 'BIRO', 'KANTOR',
    'UNIT KERJA', 'LEMBAGA', 'DEPARTEMEN'
  ]);

  var counts = { total: 0, done: 0, proc: 0, pend: 0, rej: 0 };
  var typeMap = {};
  var monthMap = {};
  var procDays = [];
  var deptMap = {};
  var latestDate = null;

  var now = new Date();
  var curYear = now.getFullYear();
  var curMonth = now.getMonth();

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];

    var noVal  = noIdx >= 0 ? row[noIdx] : (r + 1);
    var judVal = judulIdx >= 0 ? row[judulIdx] : true;
    if (!noVal || !judVal) continue;
    var noStr = (noVal || '').toString().toUpperCase();
    if (noStr.includes('URAIAN') || noStr.includes('JUMLAH') || noStr.includes('TOTAL')) continue;

    counts.total++;

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

    var jenis = jenisIdx >= 0 ? (row[jenisIdx] || '').toString().trim() : '';
    if (jenis) {
      var jl = jenis.toLowerCase();
      var key;
      if (jl.includes('gubernur')) key = 'SK Gubernur';
      else if (jl.includes('sekda') || jl.includes('sekretaris daerah')) key = 'SK Sekda';
      else key = jenis.length > 35 ? jenis.substring(0, 35) + '…' : jenis;
      typeMap[key] = (typeMap[key] || 0) + 1;
    }

    var dMasuk = tglMasukIdx >= 0 ? parseDate_(row[tglMasukIdx]) : null;
    if (dMasuk) {
      var mk = dMasuk.getFullYear() + '-' + String(dMasuk.getMonth() + 1).padStart(2, '0');
      monthMap[mk] = (monthMap[mk] || 0) + 1;
    }

    var dSelesai = tglSelesaiIdx >= 0 ? parseDate_(row[tglSelesaiIdx]) : null;
    if (dMasuk && dSelesai && (status === 'done' || status === 'selesai' || status === 'completed')) {
      var days = Math.round((dSelesai.getTime() - dMasuk.getTime()) / 86400000);
      if (days >= 0) procDays.push(days);
    }

    var dept = deptIdx >= 0 ? (row[deptIdx] || '').toString().trim() : '';
    if (dept) {
      if (!deptMap[dept]) deptMap[dept] = { gubernur: 0, sekda: 0, total: 0 };
      deptMap[dept].total++;
      if (jenis && jenis.toLowerCase().includes('gubernur')) deptMap[dept].gubernur++;
      else if (jenis && (jenis.toLowerCase().includes('sekda') || jenis.toLowerCase().includes('sekretaris daerah'))) deptMap[dept].sekda++;
    }

    var updateVal = updateIdx >= 0 ? parseDate_(row[updateIdx]) : null;
    if (!updateVal && tglMasukIdx >= 0) updateVal = dMasuk;
    if (updateVal && (!latestDate || updateVal > latestDate)) latestDate = updateVal;
  }

  var curKey = curYear + '-' + String(curMonth + 1).padStart(2, '0');
  var curMonthCount = monthMap[curKey] || 0;
  var activeMo = 0, totalYear = 0;
  for (var k in monthMap) {
    if (k.indexOf(String(curYear)) === 0) { activeMo++; totalYear += monthMap[k]; }
  }
  var monthAvg = activeMo > 0 ? Math.round(totalYear / activeMo) : 0;

  var avgProc = 0;
  if (procDays.length > 0) {
    var sum = 0;
    for (var p = 0; p < procDays.length; p++) sum += procDays[p];
    avgProc = Math.round(sum / procDays.length);
  }

  var sortedKeys = Object.keys(monthMap).sort();
  var tLabels = [], tValues = [];
  for (var m = 0; m < sortedKeys.length; m++) {
    var pp = sortedKeys[m].split('-');
    tLabels.push(BULAN[parseInt(pp[1], 10) - 1]);
    tValues.push(monthMap[sortedKeys[m]]);
  }

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

function computeYearlyAnalytics_(headers, rows, targetYear, tz) {
  var noIdx     = colIndex_(headers, ['NO.', 'NO ', 'NOMOR']);
  var judulIdx  = colIndex_(headers, ['JUDUL', 'NAMA KEPUTUSAN', 'PERIHAL']);
  var jenisIdx  = colIndex_(headers, ['JENIS SK', 'JENIS KEPUTUSAN', 'JENIS']);

  var statusIdx = 6;
  var tglMasukIdx = 4;

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

    var cellDateVal = row[tglMasukIdx];
    var mNum = extractMonthFromCell_(cellDateVal, tz);
    if (mNum >= 1 && mNum <= 12) {
      monthCounts[mNum]++;
    }

    var dMasuk = parseDate_(cellDateVal);
    var dSelesai = tglSelesaiIdx >= 0 ? parseDate_(row[tglSelesaiIdx]) : null;
    if (dMasuk && dSelesai && (status === 'done' || status === 'selesai' || status === 'completed')) {
      var days = Math.round((dSelesai.getTime() - dMasuk.getTime()) / 86400000);
      if (days >= 0) procDays.push(days);
    }

    var dept = deptIdx >= 0 ? (row[deptIdx] || '').toString().trim() : '';
    if (dept) {
      if (!deptMap[dept]) deptMap[dept] = { gubernur: 0, sekda: 0, total: 0 };
      deptMap[dept].total++;
      if (jenis && jenis.toLowerCase().includes('gubernur')) deptMap[dept].gubernur++;
      else if (jenis && (jenis.toLowerCase().includes('sekda') || jenis.toLowerCase().includes('sekretaris daerah'))) deptMap[dept].sekda++;
    }

    if (dMasuk && (!latestDateInYear || dMasuk > latestDateInYear)) {
      latestDateInYear = dMasuk;
    }
  }

  var avgMonthly = counts.total > 0 ? Math.round(counts.total / 12) : 0;

  var avgProc = 0;
  if (procDays.length > 0) {
    var sum = 0;
    for (var p = 0; p < procDays.length; p++) sum += procDays[p];
    avgProc = Math.round(sum / procDays.length);
  }

  var monthlyTrendValues = [];
  for (var m2 = 1; m2 <= 12; m2++) {
    monthlyTrendValues.push(monthCounts[m2]);
  }

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
      labels: BULAN.slice(),
      values: monthlyTrendValues
    },
    departmentRanking: { labels: dLabels, gubernur: dGov, sekda: dSek },
    lastUpdateStr: lastUpdateStr
  };
}

// ═══════════════════════════════════════════════════════════════════
//  RAWDATA KPI READER
// ═══════════════════════════════════════════════════════════════════

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

  var totalDecision = parseInt(sheet.getRange('B1').getValue(), 10) || 0;
  var gubernurDecision = parseInt(sheet.getRange('B2').getValue(), 10) || 0;
  var sekdaDecision = parseInt(sheet.getRange('B3').getValue(), 10) || 0;
  if (totalDecision <= 0) {
    totalDecision = gubernurDecision + sekdaDecision;
  }

  var proc = parseInt(sheet.getRange('B8').getValue(), 10) || 0;
  var pend = parseInt(sheet.getRange('B9').getValue(), 10) || 0;
  var rej  = parseInt(sheet.getRange('B10').getValue(), 10) || 0;
  var done = parseInt(sheet.getRange('B11').getValue(), 10) || 0;
  var total = parseInt(sheet.getRange('B12').getValue(), 10) || 0;

  var avgMonthly  = parseFloat(sheet.getRange('D8').getValue()) || 0;
  var avgProcTime = parseFloat(sheet.getRange('D9').getValue()) || 0;
  var trendYears  = parseInt(sheet.getRange('D10').getValue(), 10) || 0;

  var trendLabels = [];
  var trendValues = [];
  var lastTrendCell = null;
  try {
    var trendRange = sheet.getRange('A17:B30').getValues();
    for (var i = 0; i < trendRange.length; i++) {
      var yearLabel = trendRange[i][0];
      var yearValue = trendRange[i][1];
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

  var deptLabels = [];
  var deptGubernur = [];
  var deptSekda = [];
  try {
    var deptRange = sheet.getRange('F3:H53').getValues();
    for (var j = 0; j < deptRange.length; j++) {
      var opdName = deptRange[j][0];
      var gubernurVal = deptRange[j][1];
      var sekdaVal = deptRange[j][2];
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
//  PUBLIC API (internal — dipanggil via handleApiRequest_)
// ═══════════════════════════════════════════════════════════════════

function emptyResponse_(msg) {
  return { analytics: null, lastUpdate: '', error: msg || null };
}

function getDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return emptyResponse_('Spreadsheet tidak ditemukan.');

    var tz = Session.getScriptTimeZone();
    var sheet = resolveSheet_(ss);
    if (!sheet || sheet.getLastRow() < 2) return emptyResponse_('Sheet data kosong atau tidak ditemukan.');

    var info = detectHeaderRow_(sheet);
    var headers = info.headers;
    var dataStartRow = info.row + 1;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < dataStartRow || lastCol < 1) {
      return emptyResponse_('Tidak ada data di bawah baris header.');
    }

    var rows = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol).getValues();

    var analytics = computeAnalytics_(headers, rows);

    var rawKPI = readKPIFromRawData_(ss);
    if (rawKPI) {
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
      if (analytics.monthlyIncoming) {
        analytics.monthlyIncoming.average = rawKPI.avgMonthly;
      }
      if (analytics.processingTime) {
        analytics.processingTime.average = rawKPI.avgProcTime;
      }
      if (rawKPI.deptLabels.length > 0) {
        analytics.departmentRanking = {
          labels: rawKPI.deptLabels,
          gubernur: rawKPI.deptGubernur,
          sekda: rawKPI.deptSekda
        };
      }
    }

    var inventarisSheetForTrend = getInventarisSheet_(ss);
    var yearlyTrend = computeYearlyTrendFromInventaris_(inventarisSheetForTrend, tz);
    if (yearlyTrend.labels.length > 0) {
      analytics.monthlyTrend = {
        labels: yearlyTrend.labels,
        values: yearlyTrend.values
      };
      analytics.trendYears = yearlyTrend.totalYears;
    }

    var inventarisSheetForUpdate = getInventarisSheet_(ss);
    var lastUpdateStr = getLastUpdateFromInventarisColE_(inventarisSheetForUpdate, tz);

    if (!lastUpdateStr) {
      lastUpdateStr = getLastUpdateFromTrend_(rawKPI, analytics, headers, rows, tz);
    }

    if (!lastUpdateStr) {
      lastUpdateStr = findLatestUpdateDate_(headers, rows, tz);
      if (lastUpdateStr) {
        var parts = lastUpdateStr.split(' ');
        if (parts.length === 3) lastUpdateStr = parts[1] + ' ' + parts[2];
      }
    }

    if (!lastUpdateStr && analytics.lastUpdate) {
      var mIdx = parseInt(Utilities.formatDate(analytics.lastUpdate, tz, 'M'), 10) - 1;
      var yyyy = Utilities.formatDate(analytics.lastUpdate, tz, 'yyyy');
      lastUpdateStr = BULAN[mIdx] + ' ' + yyyy;
    }
    delete analytics.lastUpdate;

    var spreadsheetUrl = ss.getUrl();

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

    var colEIdx = 4;
    var yearSet = {};
    for (var r = 0; r < allRows.length; r++) {
      var y = extractYearFromCell_(allRows[r][colEIdx], tz);
      if (y && y >= 2019) yearSet[y] = true;
    }
    var availableYears = Object.keys(yearSet).map(function(y) { return parseInt(y, 10); });
    availableYears.sort(function(a, b) { return b - a; });

    var parsedYear = parseInt(targetYear, 10);
    if (isNaN(parsedYear) || parsedYear < 1900) {
      parsedYear = availableYears.length > 0 ? availableYears[0] : new Date().getFullYear();
    }

    var filteredRows = [];
    for (var r2 = 0; r2 < allRows.length; r2++) {
      var row = allRows[r2];
      var rowYear = extractYearFromCell_(row[colEIdx], tz);
      if (rowYear === parsedYear) {
        filteredRows.push(row);
      }
    }

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