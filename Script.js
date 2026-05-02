// ════════════════════════════════════════════════════════════
//  KA & KL MAIN KPI SCRIPT  v4
//  Sheet: 1Wp_0upauCVviyhw8ZM9XFVbIMRDgEYBZykh1_rjk0S8
//
//  ACTUAL SHEET LAYOUT (confirmed from data):
//  Each month is a separate sheet named APR, MAY, JUN...
//
//  APRIL sheet rows (0-based):
//  Row 0: Title "KA & KL REGION | APRIL 2025 | DAY-WISE..."
//  Row 1: blank/info row
//  Row 2: "S.No | STORE NAME | RM | CM | Target | ... | Apr-1 | | | | Apr-2..."
//  Row 3: "| | | | Sales | Conv | ABV | WalkIns | UPT | SALES(₹) | BILLS | SOLD QTY | WALK INS | ..."
//  Row 4+: Store data rows
//
//  TARGET SHEET: "STORE TARGETS" with cols A=Code, B-G=targets
//  OR targets are in the month sheet itself cols 5-9 (idx 4-8)
//
//  DAY DATA LAYOUT:
//  Col 5 (idx4)=Sales Target, 6=Conv%, 7=ABV, 8=WalkIns, 9=UPT
//  Col 10 (idx9)=Apr-1 Sales, 11=Bills, 12=SoldQty, 13=WalkIns
//  Col 14 (idx13)=Apr-2 Sales, 15=Bills, 16=SoldQty, 17=WalkIns
//  Pattern: Day N starts at col (10 + (N-1)*4), i.e. 0-based idx = 9+(N-1)*4
//
//  DEPLOY:
//  1. Paste into Apps Script
//  2. Save
//  3. Deploy → New deployment → Web app
//     Execute as: Me  |  Who has access: Anyone (even anonymous)
//  4. Copy the /exec URL → paste in dashboard top bar
// ════════════════════════════════════════════════════════════

var SHEET_ID    = '1Wp_0upauCVviyhw8ZM9XFVbIMRDgEYBZykh1_rjk0S8';
var MONTH_ORDER = ['APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC','JAN','FEB','MAR'];
var MONTH_DAYS  = {APR:30,MAY:31,JUN:30,JUL:31,AUG:31,SEP:30,OCT:31,NOV:30,DEC:31,JAN:31,FEB:28,MAR:31};
var MONTH_ABBR  = {APR:'Apr',MAY:'May',JUN:'Jun',JUL:'Jul',AUG:'Aug',SEP:'Sep',OCT:'Oct',NOV:'Nov',DEC:'Dec',JAN:'Jan',FEB:'Feb',MAR:'Mar'};
var QTR_MAP     = {APR:'Q1',MAY:'Q1',JUN:'Q1',JUL:'Q2',AUG:'Q2',SEP:'Q2',OCT:'Q3',NOV:'Q3',DEC:'Q3',JAN:'Q4',FEB:'Q4',MAR:'Q4'};

function doGet(e) {
  var p = e && e.parameter ? e.parameter : {};
  var cb = p.callback || '';
  var result;
  try {
    if (p.action === 'ytd') {
      result = getYTD(p.upToMonth);
    } else {
      result = getMonthly(p.sheet || 'APR');
    }
  } catch(err) {
    result = { error: 'Server error: ' + err.toString() };
  }
  var json = JSON.stringify(result);
  if (cb) return ContentService.createTextOutput(cb + '(' + json + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════
// MONTHLY DATA
// ════════════════════════════════════════
function getMonthly(monthName) {
  monthName = (monthName || 'APR').toUpperCase().trim();
  var totalDays = MONTH_DAYS[monthName] || 30;
  var mAbbr     = MONTH_ABBR[monthName] || monthName;

  // Open sheet
  var ss     = SpreadsheetApp.openById(SHEET_ID);
  var sheet  = findSheet(ss, monthName);
  if (!sheet) {
    var available = ss.getSheets().map(function(s){return s.getName();}).join(', ');
    return { error: 'Sheet "' + monthName + '" not found. Available: ' + available };
  }

  var allData = sheet.getDataRange().getValues();
  Logger.log(monthName + ': ' + allData.length + ' rows x ' + (allData[0]||[]).length + ' cols');

  // ── Find the header row (has "S.No" or "STORE NAME") ──
  var headerRowIdx = -1;
  for (var r = 0; r < Math.min(allData.length, 12); r++) {
    var r0 = String(allData[r][0] || '').trim().toLowerCase();
    var r1 = String(allData[r][1] || '').trim().toLowerCase();
    if (r0 === 's.no' || r1 === 'store name' || r1.indexOf('store name') !== -1) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx < 0) {
    // Last resort: find row with "S.No" anywhere
    for (var r2 = 0; r2 < Math.min(allData.length, 12); r2++) {
      for (var c2 = 0; c2 < Math.min(allData[r2].length, 6); c2++) {
        if (String(allData[r2][c2]).trim().toLowerCase() === 's.no') {
          headerRowIdx = r2; break;
        }
      }
      if (headerRowIdx >= 0) break;
    }
  }
  if (headerRowIdx < 0) headerRowIdx = 3; // fallback default

  var firstDataRow = headerRowIdx + 2; // skip header + sub-header
  var headerRow    = allData[headerRowIdx] || [];

  // ── Find day columns ──
  // Method 1: look for "Apr-1", "Apr-2" etc in header row
  var dayLabels    = [];
  var dayStartCols = []; // 0-based col index of SALES for each day

  for (var c = 0; c < headerRow.length; c++) {
    var h = String(headerRow[c] || '').trim();
    // Match "Apr-1", "May-12", "Jun-3" etc
    if (h.match(/^[A-Za-z]{3}-\d{1,2}$/)) {
      dayLabels.push(h);
      dayStartCols.push(c);
    }
  }

  // Method 2: if none found, use fixed layout
  // Known layout: Target cols at idx 4-8 (cols 5-9), Day1 SALES at idx 9 (col 10)
  if (dayLabels.length === 0) {
    Logger.log('No day labels in header row ' + headerRowIdx + ' — using fixed layout');
    for (var d = 1; d <= totalDays; d++) {
      dayLabels.push(mAbbr + '-' + d);
      dayStartCols.push(9 + (d - 1) * 4); // 0-based: col 10 = idx 9 for day 1
    }
  }

  Logger.log('Day cols found: ' + dayLabels.length + ' starting at col ' + (dayStartCols[0]+1));

  // ── Parse stores ──
  var stores = [];
  for (var r3 = firstDataRow; r3 < allData.length; r3++) {
    var row = allData[r3];
    var sno = String(row[0] || '').trim();
    if (!sno || isNaN(parseInt(sno))) continue; // skip non-data rows

    var storeFull = String(row[1] || '').trim();
    if (!storeFull) continue;
    if (storeFull.toUpperCase().indexOf('WAREHOUSE') !== -1) continue;
    if (storeFull.toUpperCase().indexOf('TOTAL') !== -1) continue;

    // Extract code from "BLR - JAYNAGAR (BLRJAY)"
    var code = storeFull;
    var cm2  = storeFull.match(/\(([A-Z0-9]{4,8})\)\s*$/);
    if (cm2) code = cm2[1];

    var rm = String(row[2] || '').trim();
    var cm = String(row[3] || '').trim();

    // Targets: idx 4=SalesTarget, 5=Conv%, 6=ABV, 7=WalkIns, 8=UPT
    var tgtSales   = numVal(row[4]);
    var tgtConv    = pctVal(row[5]);  // "53%" or 0.53
    var tgtABV     = numVal(row[6]);
    var tgtWalkIns = numVal(row[7]);
    var tgtUPT     = pctVal(row[8]);
    var tgtBills   = tgtWalkIns > 0 && tgtConv > 0 ? Math.round(tgtWalkIns * tgtConv / 100) : 0;

    // Day-wise data
    var daySales = [], dayBills = [], dayQty = [], dayWalkIns = [];
    for (var di = 0; di < dayStartCols.length; di++) {
      var c3 = dayStartCols[di];
      daySales.push(numVal(row[c3]));
      dayBills.push(numVal(row[c3 + 1]));
      dayQty.push(numVal(row[c3 + 2]));
      dayWalkIns.push(numVal(row[c3 + 3]));
    }

    // MTD calculations
    var mtdSales   = arrSum(daySales);
    var mtdBills   = arrSum(dayBills);
    var mtdQty     = arrSum(dayQty);
    var mtdWalkIns = arrSum(dayWalkIns);
    var daysWithData = daySales.filter(function(v){ return v > 0; }).length || 1;

    var projTotal   = Math.round(mtdSales / daysWithData * totalDays);
    var pct         = tgtSales > 0 ? Math.round(mtdSales / tgtSales * 100) : 0;
    var trendingPct = tgtSales > 0 ? Math.round(projTotal / tgtSales * 100) : 0;
    var convActual  = mtdWalkIns > 0 ? Math.round(mtdBills / mtdWalkIns * 1000) / 10 : 0;
    var abv         = mtdBills > 0 ? Math.round(mtdSales / mtdBills) : 0;
    var upt         = mtdBills > 0 ? Math.round(mtdQty / mtdBills * 100) / 100 : 0;

    stores.push({
      code: code,
      name: storeFull,
      rm: rm,
      cm: cm,
      target: tgtSales,
      mtdSales: mtdSales,
      projectedTotal: projTotal,
      pct: pct,
      trendingPct: trendingPct,
      mtdBills: mtdBills,
      mtdWalkIns: mtdWalkIns,
      mtdSoldQty: mtdQty,
      convActual: convActual,
      abv: abv,
      upt: upt,
      todaySales: daySales[daysWithData - 1] || 0,
      daysCount: daysWithData,
      daySales: daySales,
      dayBills: dayBills,
      kpiTargets: {
        sales: tgtSales,
        bills: tgtBills,
        walkIns: tgtWalkIns,
        conversion: tgtConv,
        abv: tgtABV,
        upt: tgtUPT
      },
      kpiTrending: {
        walkIns: { pct: tgtWalkIns > 0 ? Math.round(mtdWalkIns / tgtWalkIns * 100) : null }
      }
    });
  }

  Logger.log('Stores: ' + stores.length + ' | Days: ' + dayLabels.length + ' | MTD sample: ' + (stores[0] ? stores[0].mtdSales : 'n/a'));

  return {
    stores: stores,
    dayLabels: dayLabels,
    totalDaysInMonth: totalDays,
    elapsed: daysWithData || dayLabels.length,
    remaining: Math.max(0, totalDays - (daysWithData || dayLabels.length)),
    sheets: getAvailableSheets(ss),
    month: monthName,
    generatedAt: new Date().toISOString()
  };
}

// ════════════════════════════════════════
// YTD DATA
// ════════════════════════════════════════
function getYTD(upToMonth) {
  upToMonth = (upToMonth || 'APR').toUpperCase().trim();
  var upToIdx = MONTH_ORDER.indexOf(upToMonth);
  if (upToIdx < 0) upToIdx = 0;

  var monthsToRead = MONTH_ORDER.slice(0, upToIdx + 1);
  Logger.log('YTD: reading months ' + monthsToRead.join(', '));

  // Aggregate per store across all months
  var storeAgg = {}; // code → aggregated data

  for (var mi = 0; mi < monthsToRead.length; mi++) {
    var m = monthsToRead[mi];
    var mData;
    try {
      mData = getMonthly(m);
    } catch(e) {
      Logger.log('Error reading ' + m + ': ' + e.toString());
      continue;
    }

    // Skip months with errors or no data
    if (!mData || mData.error || !mData.stores || mData.stores.length === 0) {
      Logger.log('Skipping ' + m + ': ' + (mData && mData.error ? mData.error : 'no stores'));
      continue;
    }

    Logger.log(m + ': ' + mData.stores.length + ' stores');

    for (var si = 0; si < mData.stores.length; si++) {
      var s = mData.stores[si];
      var key = s.code || s.name;
      if (!key) continue;

      if (!storeAgg[key]) {
        storeAgg[key] = {
          code: s.code,
          name: s.name,
          rm:   s.rm,
          cm:   s.cm,
          ytdSales: 0, ytdTarget: 0,
          ytdBills: 0, ytdWalkIns: 0, ytdQty: 0,
          monthlyData: {}
        };
      }

      var a = storeAgg[key];
      a.ytdSales   += s.mtdSales   || 0;
      a.ytdTarget  += s.target     || 0;
      a.ytdBills   += s.mtdBills   || 0;
      a.ytdWalkIns += s.mtdWalkIns || 0;
      a.ytdQty     += s.mtdSoldQty || 0;
      a.monthlyData[m] = {
        sales:  s.mtdSales || 0,
        target: s.target   || 0,
        pct:    s.pct      || 0,
        bills:  s.mtdBills || 0
      };
    }
  }

  // Format store list
  var stores = Object.values(storeAgg).map(function(a) {
    var pct     = a.ytdTarget > 0 ? Math.round(a.ytdSales / a.ytdTarget * 100) : 0;
    var abv     = a.ytdBills  > 0 ? Math.round(a.ytdSales / a.ytdBills) : 0;
    var upt     = a.ytdBills  > 0 ? Math.round(a.ytdQty   / a.ytdBills * 100) / 100 : 0;
    var conv    = a.ytdWalkIns > 0 ? Math.round(a.ytdBills / a.ytdWalkIns * 1000) / 10 : 0;
    return {
      code:        a.code,
      name:        a.name,
      rm:          a.rm,
      cm:          a.cm,
      ytdSales:    Math.round(a.ytdSales),
      ytdTarget:   Math.round(a.ytdTarget),
      ytdBills:    Math.round(a.ytdBills),
      ytdWalkIns:  Math.round(a.ytdWalkIns),
      ytdQty:      Math.round(a.ytdQty),
      pct:         pct,
      abv:         abv,
      upt:         upt,
      convActual:  conv,
      monthlyData: a.monthlyData
    };
  }).sort(function(a, b) { return b.ytdSales - a.ytdSales; });

  Logger.log('YTD done: ' + stores.length + ' stores across ' + monthsToRead.length + ' months');

  return {
    stores:     stores,
    months:     monthsToRead,
    upToMonth:  upToMonth,
    quarter:    QTR_MAP[upToMonth] || 'Q1',
    generatedAt: new Date().toISOString()
  };
}

// ════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════
function findSheet(ss, name) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toUpperCase() === name.toUpperCase()) return sheets[i];
  }
  // Also try partial match e.g. "APRIL" → "APR"
  for (var j = 0; j < sheets.length; j++) {
    var sn = sheets[j].getName().toUpperCase();
    if (sn.indexOf(name) === 0 || name.indexOf(sn) === 0) return sheets[j];
  }
  return null;
}

function getAvailableSheets(ss) {
  return ss.getSheets()
    .map(function(s) { return s.getName().toUpperCase(); })
    .filter(function(n) { return MONTH_ORDER.indexOf(n) !== -1; });
}

function numVal(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  var s = String(v).replace(/[₹,\s]/g, '').trim();
  if (s === '' || s === '#REF!' || s === '#DIV/0!' || s === '#N/A') return 0;
  return parseFloat(s) || 0;
}

function pctVal(v) {
  // Handles both "53%" (string) and 0.53 (fraction) and 53 (number)
  if (v === null || v === undefined || v === '') return 0;
  var s = String(v).trim();
  if (s.indexOf('%') !== -1) return parseFloat(s) || 0; // "53%" → 53
  var n = parseFloat(s) || 0;
  if (n > 0 && n <= 1) return Math.round(n * 100); // 0.53 → 53
  return n; // already a percentage number
}

function arrSum(arr) {
  return (arr || []).reduce(function(a, b) { return a + (b || 0); }, 0);
}
