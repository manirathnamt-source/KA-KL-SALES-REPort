// ============================================================
//  KA & KL Store Dashboard — Apps Script FINAL (Fast Version)
//  Optimised: reads only needed rows/cols, no full sheet scans
// ============================================================

var PROPS    = PropertiesService.getScriptProperties();
var USER_KEY = 'active_users';
var USER_TTL = 5 * 60 * 1000;

var MONTH_DAYS = {
  APR:30,MAY:31,JUN:30,JUL:31,AUG:31,
  SEP:30,OCT:31,NOV:30,DEC:31,JAN:31,FEB:28,MAR:31
};

var MONTH_NAMES_MAP = {
  'APRIL':'APR','APR':'APR','MAY':'MAY','JUNE':'JUN','JUN':'JUN',
  'JULY':'JUL','JUL':'JUL','AUGUST':'AUG','AUG':'AUG',
  'SEPTEMBER':'SEP','SEP':'SEP','OCTOBER':'OCT','OCT':'OCT',
  'NOVEMBER':'NOV','NOV':'NOV','DECEMBER':'DEC','DEC':'DEC',
  'JANUARY':'JAN','JAN':'JAN','FEBRUARY':'FEB','FEB':'FEB',
  'MARCH':'MAR','MAR':'MAR'
};

function doGet(e) {
  var p        = (e && e.parameter) ? e.parameter : {};
  var sheet    = p.sheet    || 'APR';
  var callback = p.callback || '';
  var action   = p.action   || 'data';
  var userId   = p.userId   || '';
  var userName = p.userName || '';

  var result;
  try {
    if (action === 'heartbeat') {
      updatePresence(userId, userName);
      result = { users: getActiveUsers(), ts: Date.now() };

    } else if (action === 'leave') {
      removeUser(userId);
      result = { ok: true };

    } else if (action === 'debug_targets') {
      var ss0 = SpreadsheetApp.getActiveSpreadsheet();
      var kt  = loadKPITargets(ss0, sheet);
      result  = {
        totalStores: Object.keys(kt).length,
        kpiSample:   objSlice(kt, 5),
        message:     'Check kpiSample — conversion should be e.g. 41 not 0.41'
      };

    } else {
      if (userId) updatePresence(userId, userName);
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var ws = ss.getSheetByName(sheet);
      if (!ws) {
        var names = ss.getSheets().map(function(s){ return s.getName(); });
        result = { error: 'Sheet "' + sheet + '" not found. Available: ' + names.join(', ') };
      } else {
        var kpiTargets = loadKPITargets(ss, sheet);
        result = parseSheet(ws, sheet, kpiTargets);
        result.activeUsers = getActiveUsers();
      }
    }
  } catch(err) {
    result = { error: err.toString() + ' | Stack: ' + (err.stack||'').substring(0, 200) };
  }

  var json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function objSlice(obj, n) {
  var out = {}, i = 0;
  for (var k in obj) { if (i++ >= n) break; out[k] = obj[k]; }
  return out;
}

// ── Load KPI targets — FAST version ──
// Only reads first ~60 rows of STORE TARGETS (enough for header + 45 stores)
function loadKPITargets(ss, sheetName) {
  var targets = {};
  try {
    var ts = ss.getSheetByName('STORE TARGETS');
    if (!ts) return targets;

    var lastRow = Math.min(ts.getLastRow(), 65);   // max 65 rows needed
    var lastCol = Math.min(ts.getLastColumn(), 30); // max 30 cols needed
    if (lastRow < 3 || lastCol < 5) return targets;

    var rows = ts.getRange(1, 1, lastRow, lastCol).getValues();
    var sheetUp = sheetName.toUpperCase();

    // ── Find month column and KPI header row ──
    var monthStartCol = -1;
    var kpiHeaderRow  = -1;

    // Scan first 8 rows for month name
    for (var i = 0; i < Math.min(8, rows.length); i++) {
      var row = rows[i];
      for (var c = 4; c < row.length; c++) {
        var cv = String(row[c] || '').trim().toUpperCase().replace(/\s+/g,'');
        var mapped = MONTH_NAMES_MAP[cv] || MONTH_NAMES_MAP[String(row[c]||'').trim().toUpperCase()];
        if (mapped === sheetUp) {
          monthStartCol = c;
          // Look for KPI sub-headers in next 3 rows
          for (var j = i+1; j < Math.min(i+4, rows.length); j++) {
            var jRow = rows[j];
            var jText = jRow.slice(c, Math.min(c+8, jRow.length))
                            .map(function(v){ return String(v||'').toUpperCase(); }).join('|');
            if (jText.indexOf('SALES') !== -1 || jText.indexOf('BILL') !== -1) {
              kpiHeaderRow = j;
              break;
            }
          }
          break;
        }
      }
      if (monthStartCol >= 0) break;
    }

    // Fallback: look for a row with "Sales Target" + "Bill Target" keywords
    if (kpiHeaderRow === -1) {
      for (var i2 = 0; i2 < Math.min(8, rows.length); i2++) {
        var rText = rows[i2].slice(0, Math.min(20, rows[i2].length))
                            .map(function(v){ return String(v||'').toUpperCase(); }).join('|');
        if ((rText.indexOf('SALES') !== -1 || rText.indexOf('REVENUE') !== -1) &&
            (rText.indexOf('BILL') !== -1 || rText.indexOf('CONV') !== -1)) {
          kpiHeaderRow = i2;
          if (monthStartCol === -1) {
            // Find first SALES col in this row
            for (var c2 = 4; c2 < rows[i2].length; c2++) {
              if (String(rows[i2][c2]||'').toUpperCase().indexOf('SALES') !== -1) {
                monthStartCol = c2; break;
              }
            }
          }
          break;
        }
      }
    }

    // Ultimate fallback — use month position
    if (monthStartCol === -1) {
      var ORDER = ['APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC','JAN','FEB','MAR'];
      var mi = ORDER.indexOf(sheetUp);
      monthStartCol = (mi >= 0) ? 4 + mi * 6 : 4;
    }
    if (kpiHeaderRow === -1) kpiHeaderRow = 3; // row 4 (0-indexed = 3)

    // ── Map KPI columns from header row ──
    var colMap = { sales:-1, bills:-1, conversion:-1, abv:-1, walkIns:-1, upt:-1 };
    if (kpiHeaderRow >= 0 && kpiHeaderRow < rows.length) {
      var hRow = rows[kpiHeaderRow];
      var scan_from = Math.max(4, monthStartCol - 1);
      var scan_to   = Math.min(scan_from + 10, hRow.length);
      for (var c3 = scan_from; c3 < scan_to; c3++) {
        var h = String(hRow[c3] || '').replace(/[\s\n]/g,'').toUpperCase();
        if (h.indexOf('SALES') !== -1 || h.indexOf('REVENUE') !== -1) {
          if (colMap.sales === -1) colMap.sales = c3;
        } else if (h.indexOf('BILL') !== -1) {
          if (colMap.bills === -1) colMap.bills = c3;
        } else if (h.indexOf('CONV') !== -1) {
          if (colMap.conversion === -1) colMap.conversion = c3;
        } else if (h.indexOf('ABV') !== -1 || h.indexOf('ATV') !== -1) {
          if (colMap.abv === -1) colMap.abv = c3;
        } else if (h.indexOf('WALK') !== -1 || h.indexOf('FOOT') !== -1) {
          if (colMap.walkIns === -1) colMap.walkIns = c3;
        } else if (h.indexOf('UPT') !== -1 || h.indexOf('UNIT') !== -1) {
          if (colMap.upt === -1) colMap.upt = c3;
        }
      }
    }
    // If Sales col still not found, use monthStartCol
    if (colMap.sales === -1) colMap.sales = monthStartCol;

    // ── Parse store rows ──
    var SKIP = ['TOTAL','SUMMARY','GRAND','REGION','★','STORE NAME','S.NO'];
    var dataStart = kpiHeaderRow >= 0 ? kpiHeaderRow + 1 : 4;

    for (var r = dataStart; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row[1]) continue;
      var name = String(row[1]).trim();
      if (name.length < 3) continue;
      var skip = SKIP.some(function(s){ return name.toUpperCase().indexOf(s) !== -1; });
      if (skip) continue;
      if (!name.match(/\([A-Z]{3}[A-Z0-9]+\)/)) continue; // must have store code

      var codeM = name.match(/\(([A-Z0-9]+)\)/);
      var code  = codeM ? codeM[1] : name.substring(0,8).toUpperCase().replace(/\s/g,'');

      var salesVal  = colMap.sales      >= 0 ? parseNum(row[colMap.sales])      : 0;
      var billsVal  = colMap.bills      >= 0 ? parseNum(row[colMap.bills])      : 0;
      var convRaw   = colMap.conversion >= 0 ? parseNum(row[colMap.conversion]) : 0;
      var abvVal    = colMap.abv        >= 0 ? parseNum(row[colMap.abv])        : 0;
      var walkVal   = colMap.walkIns    >= 0 ? parseNum(row[colMap.walkIns])    : 0;
      var uptVal    = colMap.upt        >= 0 ? parseNum(row[colMap.upt])        : 0;

      // Fix: Google Sheets stores 41% as 0.41 — convert to percentage
      var convVal = (convRaw > 0 && convRaw < 1) ? Math.round(convRaw * 1000) / 10 : convRaw;

      targets[code] = {
        sales:      salesVal,
        bills:      billsVal,
        conversion: convVal,
        abv:        abvVal,
        walkIns:    walkVal,
        upt:        uptVal
      };
    }

    Logger.log('KPI targets loaded: ' + Object.keys(targets).length + ' stores | colMap: ' + JSON.stringify(colMap));
  } catch(err) {
    Logger.log('loadKPITargets error: ' + err.toString());
  }
  return targets;
}

// ── Parse sheet — FAST version ──
// Reads only data rows, limits column scan
function parseSheet(ws, sheetName, kpiTargets) {
  var lastRow = ws.getLastRow();
  var lastCol = ws.getLastColumn();

  if (lastRow < 4 || lastCol < 6) {
    return { stores:[], sheet:sheetName, dayLabels:[], daysCount:0, totalDaysInMonth:30 };
  }

  // Read only first 5 rows to find structure
  var headerRows = ws.getRange(1, 1, Math.min(5, lastRow), Math.min(lastCol, 140)).getValues();
  var row3 = headerRows[2] || [];
  var row4 = headerRows[3] || [];

  // ── Find SALES column positions from row 4 ──
  var salesCols = [];
  for (var c = 5; c < row4.length; c++) {
    var ft = String(row4[c] || '').replace(/\s/g,'').toUpperCase();
    var lt = String(row3[c] || '').replace(/\s/g,'').toUpperCase();
    if (ft.indexOf('SALES') !== -1 &&
        lt.indexOf('MTD')     === -1 &&
        lt.indexOf('SUMMARY') === -1 &&
        lt.indexOf('TOTAL')   === -1) {
      salesCols.push(c);
    }
  }
  // Fallback
  if (salesCols.length === 0) {
    for (var c2 = 5; c2 < row3.length; c2 += 4) {
      var v3 = String(row3[c2]||'').replace(/\s/g,'').toUpperCase();
      if (v3 && v3.indexOf('MTD')===-1 && v3.indexOf('SUMMARY')===-1) salesCols.push(c2);
    }
  }

  var sheetUp          = sheetName.toUpperCase();
  var knownDays        = MONTH_DAYS[sheetUp] || 30;
  var totalDaysInMonth = Math.min(salesCols.length, knownDays);
  var dayLabels        = salesCols.slice(0, totalDaysInMonth).map(function(sc, i) {
    return String(row3[sc] || '').trim() || (sheetName + '-' + (i+1));
  });

  // ── Trending cutoff = yesterday ──
  var todayDay  = new Date().getDate();
  var cutoff    = todayDay - 1; // yesterday
  var lastIdx   = -1;

  // Read store data rows to find last day with data
  // Only read up to row 60 for speed
  var dataRows = ws.getRange(5, 1, Math.min(lastRow - 4, 56), Math.min(lastCol, 134)).getValues();

  for (var di = 0; di < totalDaysInMonth; di++) {
    if ((di + 1) > cutoff) break;
    var sc = salesCols[di];
    var dayTotal = 0;
    for (var ri = 0; ri < dataRows.length; ri++) {
      if (sc < dataRows[ri].length) dayTotal += parseNum(dataRows[ri][sc]);
    }
    if (dayTotal > 0) lastIdx = di;
  }

  // Fallback: any data at all
  if (lastIdx === -1) {
    for (var di2 = 0; di2 < Math.min(totalDaysInMonth, cutoff+1); di2++) {
      var sc2 = salesCols[di2];
      var dt = 0;
      for (var ri2 = 0; ri2 < dataRows.length; ri2++) {
        if (sc2 < dataRows[ri2].length) dt += parseNum(dataRows[ri2][sc2]);
      }
      if (dt > 0) lastIdx = di2;
    }
  }
  if (lastIdx === -1) lastIdx = Math.max(0, Math.min(cutoff, totalDaysInMonth-1));

  var activeCols      = salesCols.slice(0, lastIdx + 1);
  var activeDayLabels = dayLabels.slice(0, lastIdx + 1);
  var lastSalesCol    = salesCols[lastIdx] || salesCols[0] || 5;
  var daysElapsed     = activeCols.length;
  var daysLeft        = Math.max(0, totalDaysInMonth - daysElapsed);

  // ── Parse stores ──
  var SKIP   = ['TOTAL','SUMMARY','GRAND','REGION','★'];
  var stores = [];

  for (var row = 0; row < dataRows.length; row++) {
    var r = dataRows[row];
    if (!r || !r[1]) continue;
    var name = String(r[1]).trim();
    if (name.length < 3 || name.toUpperCase().indexOf('STORE NAME') !== -1) continue;
    if (SKIP.some(function(s){ return name.toUpperCase().indexOf(s) !== -1; })) continue;

    var rm = String(r[2]||'').trim();
    var cm = String(r[3]||'').trim();
    var target = parseNum(r[4]);

    var ms=0, mb=0, mq=0, mw=0, daySales=[];
    for (var d = 0; d < activeCols.length; d++) {
      var col = activeCols[d];
      var sv = col     < r.length ? parseNum(r[col])     : 0;
      var bv = col + 1 < r.length ? parseNum(r[col + 1]) : 0;
      var qv = col + 2 < r.length ? parseNum(r[col + 2]) : 0;
      var wv = col + 3 < r.length ? parseNum(r[col + 3]) : 0;
      daySales.push(sv); ms += sv; mb += bv; mq += qv; mw += wv;
    }

    var todaySales = lastSalesCol     < r.length ? parseNum(r[lastSalesCol])     : 0;
    var todayBills = lastSalesCol + 1 < r.length ? parseNum(r[lastSalesCol + 1]) : 0;
    var todayWalk  = lastSalesCol + 3 < r.length ? parseNum(r[lastSalesCol + 3]) : 0;

    var pct  = target > 0 ? Math.round(ms / target * 100) : 0;
    var abv  = mb > 0 ? Math.round(ms / mb) : 0;
    var upt  = mb > 0 ? Math.round(mq / mb * 100) / 100 : 0;
    var conv = mw > 0 ? Math.round(mb / mw * 1000) / 10 : 0;

    var codeM = name.match(/\(([A-Z0-9]+)\)/);
    var code  = codeM ? codeM[1] : name.substring(0,8).toUpperCase().replace(/\s/g,'');
    var cityM = code.match(/^([A-Z]{3})/);
    var city  = cityM ? cityM[1] : '?';

    var kt       = kpiTargets[code] || { sales:0,bills:0,conversion:0,abv:0,walkIns:0,upt:0 };
    var salesTgt = kt.sales || target;

    function proj(a) { return daysElapsed > 0 ? Math.round(a + (a / daysElapsed) * daysLeft) : a; }
    function kpiT(a, t) {
      if (!t) return null;
      var p = proj(a), pc = Math.round(p / t * 100);
      return { actual: Math.round(a*100)/100, projected: Math.round(p*100)/100, target: t, pct: pc, status: pc >= 100 ? 'green' : 'red' };
    }
    function rateT(a, t) {
      if (!t) return null;
      var pc = Math.round(a / t * 100);
      return { actual: Math.round(a*100)/100, projected: Math.round(a*100)/100, target: t, pct: pc, status: pc >= 100 ? 'green' : 'red' };
    }

    var projTotal   = proj(ms);
    var trendPct    = salesTgt > 0 ? Math.round(projTotal / salesTgt * 100) : 0;
    var trendStatus = trendPct >= 100 ? 'on-track' : trendPct >= 85 ? 'at-risk' : 'behind';

    if (target > 0 || ms > 0) {
      stores.push({
        name: name, code: code, city: city, rm: rm, cm: cm,
        target:       Math.round(salesTgt),
        mtdSales:     Math.round(ms),
        mtdBills:     Math.round(mb),
        mtdSoldQty:   Math.round(mq),
        mtdWalkIns:   Math.round(mw),
        todaySales:   Math.round(todaySales),
        todayBills:   Math.round(todayBills),
        todayWalkIns: Math.round(todayWalk),
        pct: pct, abv: abv, upt: upt, convActual: conv,
        projectedTotal: projTotal, trendingPct: trendPct, trendStatus: trendStatus,
        daysCount: daysElapsed, daySales: daySales,
        kpiTargets: {
          sales:      Math.round(salesTgt),
          bills:      Math.round(kt.bills  || 0),
          conversion: kt.conversion || 0,
          abv:        Math.round(kt.abv    || 0),
          walkIns:    Math.round(kt.walkIns|| 0),
          upt:        kt.upt || 0
        },
        kpiTrending: {
          sales:      kpiT(ms,  salesTgt),
          bills:      kpiT(mb,  kt.bills),
          walkIns:    kpiT(mw,  kt.walkIns),
          conversion: rateT(conv, kt.conversion),
          abv:        rateT(abv,  kt.abv),
          upt:        rateT(upt,  kt.upt)
        }
      });
    }
  }

  var allSheets = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets().map(function(s){ return s.getName(); });

  return {
    stores:           stores,
    sheet:            sheetName,
    sheets:           allSheets,
    dayLabels:        activeDayLabels,
    daysCount:        daysElapsed,
    totalDaysInMonth: totalDaysInMonth,
    cutoffDay:        cutoff,
    todayDate:        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy')
  };
}

function parseNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  var str = String(val).trim();
  var s = str.replace(/[₹,%\s]/g, '').replace(/[^0-9.-]/g, '');
  return parseFloat(s) || 0;
}

// ── Presence tracking ──
function updatePresence(u, n) {
  if (!u) return;
  var d = {};
  try { d = JSON.parse(PROPS.getProperty(USER_KEY) || '{}'); } catch(e) {}
  d[u] = { name: n || u, ts: Date.now() };
  PROPS.setProperty(USER_KEY, JSON.stringify(d));
}
function removeUser(u) {
  if (!u) return;
  var d = {};
  try { d = JSON.parse(PROPS.getProperty(USER_KEY) || '{}'); } catch(e) {}
  delete d[u];
  PROPS.setProperty(USER_KEY, JSON.stringify(d));
}
function getActiveUsers() {
  var d = {};
  try { d = JSON.parse(PROPS.getProperty(USER_KEY) || '{}'); } catch(e) { return []; }
  var now = Date.now(), a = [], changed = false;
  Object.keys(d).forEach(function(id) {
    if (now - d[id].ts < USER_TTL) { a.push({ id: id, name: d[id].name, ts: d[id].ts }); }
    else { delete d[id]; changed = true; }
  });
  if (changed) PROPS.setProperty(USER_KEY, JSON.stringify(d));
  a.sort(function(a, b) { return b.ts - a.ts; });
  return a;
}
