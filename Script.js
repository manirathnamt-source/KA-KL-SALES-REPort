// ============================================================// ============================================================
//  KA & KL Store Dashboard — Apps Script FINAL (Real-time)
//
//  HOW REAL-TIME SYNC WORKS:
//  1. doGet() reads FRESH data from sheet every single call
//  2. No time-driven trigger needed — delete any you have
//  3. No caching — we add ?t= timestamp to bust CDN cache
//  4. Dashboard polls every 60s — data appears within 1 min
//
//  NO onEdit trigger needed — polling is the correct approach
// ============================================================

var SPREADSHEET_ID = '1Wp_0upauCVviyhw8ZM9XFVbIMRDgEYBZykh1_rjk0S8';
var SECRET_KEY     = '4NJqtPi77ctO3Ec5acEenwwa17XHEFhT';

var PROPS    = PropertiesService.getScriptProperties();
var USER_KEY = 'active_users';
var USER_TTL = 5 * 60 * 1000; // 5 min

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

// ══════════════════════════════════════════════════════════
//  doGet — Entry point. Runs FRESH every call. No caching.
// ══════════════════════════════════════════════════════════
function doGet(e) {
  var p        = (e && e.parameter) ? e.parameter : {};
  var sheet    = p.sheet    || 'APR';
  var callback = p.callback || '';
  var action   = p.action   || 'data';
  var userId   = p.userId   || '';
  var userName = p.userName || '';
  // p.t is a timestamp added by dashboard to bust CDN cache — we ignore it

  // ── API key check ──
  if (p.key !== SECRET_KEY) {
    var denied = JSON.stringify({ error: 'Unauthorized' });
    if (callback) return ContentService.createTextOutput(callback+'('+denied+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(denied).setMimeType(ContentService.MimeType.JSON);
  }

  var result;
  try {
    // openById reads your live sheet directly — always fresh
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (action === 'heartbeat') {
      updatePresence(userId, userName);
      result = { users: getActiveUsers(), ts: Date.now() };

    } else if (action === 'leave') {
      removeUser(userId);
      result = { ok: true };

    } else if (action === 'ping') {
      // Lightweight check — no sheet read, just confirms script is alive
      result = { ok: true, ts: Date.now(), sheet: sheet };

    } else if (action === 'debug_targets') {
      var kt0 = loadKPITargets(ss, sheet);
      result  = { totalStores: Object.keys(kt0).length, kpiSample: objSlice(kt0, 5) };

    } else {
      // ── Main data fetch — always reads live sheet ──
      if (userId) updatePresence(userId, userName);
      var ws = ss.getSheetByName(sheet);
      if (!ws) {
        var names = ss.getSheets().map(function(s){ return s.getName(); });
        result = { error: 'Sheet "'+sheet+'" not found. Available: '+names.join(', ') };
      } else {
        var kpiTargets = loadKPITargets(ss, sheet);
        result = parseSheet(ws, sheet, kpiTargets);
        result.activeUsers = getActiveUsers();
        result.fetchedAt   = Utilities.formatDate(
          new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm:ss'
        );
      }
    }
  } catch(err) {
    result = { error: err.toString() };
  }

  var json = JSON.stringify(result);

  // Return with JAVASCRIPT mime type for JSONP
  // Google does NOT cache JAVASCRIPT responses the same way — this helps avoid stale data
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════
//  DELETE any time-driven triggers you have!
//  This function helps you clean them up — run it once manually
// ══════════════════════════════════════════════════════════
function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  Logger.log('Deleted ' + triggers.length + ' trigger(s). Real-time polling handles sync now.');
}

// ══════════════════════════════════════════════════════════
//  KPI TARGETS — reads STORE TARGETS sheet
// ══════════════════════════════════════════════════════════
function loadKPITargets(ss, sheetName) {
  var targets = {};
  try {
    var ts = ss.getSheetByName('STORE TARGETS');
    if (!ts) return targets;
    var lastRow = Math.min(ts.getLastRow(), 65);
    var lastCol = Math.min(ts.getLastColumn(), 30);
    if (lastRow < 3 || lastCol < 5) return targets;

    var rows = ts.getRange(1, 1, lastRow, lastCol).getValues();
    var sheetUp = sheetName.toUpperCase();

    var monthStartCol = -1, kpiHeaderRow = -1;
    for (var i = 0; i < Math.min(8, rows.length); i++) {
      for (var c = 4; c < rows[i].length; c++) {
        var cv = String(rows[i][c]||'').trim().toUpperCase().replace(/\s+/g,'');
        if (MONTH_NAMES_MAP[cv] === sheetUp) {
          monthStartCol = c;
          for (var j = i+1; j < Math.min(i+4, rows.length); j++) {
            var jt = rows[j].slice(c, Math.min(c+8, rows[j].length))
                            .map(function(v){return String(v||'').toUpperCase();}).join('|');
            if (jt.indexOf('SALES')!==-1||jt.indexOf('BILL')!==-1){kpiHeaderRow=j;break;}
          }
          break;
        }
      }
      if (monthStartCol >= 0) break;
    }
    if (kpiHeaderRow === -1) {
      for (var i2 = 0; i2 < Math.min(8, rows.length); i2++) {
        var rt = rows[i2].slice(0,Math.min(20,rows[i2].length))
                         .map(function(v){return String(v||'').toUpperCase();}).join('|');
        if (rt.indexOf('SALES')!==-1&&(rt.indexOf('BILL')!==-1||rt.indexOf('CONV')!==-1)){
          kpiHeaderRow=i2;
          if (monthStartCol===-1) for(var c2=4;c2<rows[i2].length;c2++){if(String(rows[i2][c2]||'').toUpperCase().indexOf('SALES')!==-1){monthStartCol=c2;break;}}
          break;
        }
      }
    }
    if (monthStartCol===-1){var ORD=['APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC','JAN','FEB','MAR'];var mi=ORD.indexOf(sheetUp);monthStartCol=(mi>=0)?4+mi*6:4;}
    if (kpiHeaderRow===-1) kpiHeaderRow=3;

    var colMap={sales:-1,bills:-1,conversion:-1,abv:-1,walkIns:-1,upt:-1};
    if (kpiHeaderRow<rows.length){
      var hRow=rows[kpiHeaderRow],sf=Math.max(4,monthStartCol-1),st=Math.min(sf+10,hRow.length);
      for(var c3=sf;c3<st;c3++){
        var h=String(hRow[c3]||'').replace(/[\s\n]/g,'').toUpperCase();
        if((h.indexOf('SALES')!==-1||h.indexOf('REVENUE')!==-1)&&colMap.sales===-1) colMap.sales=c3;
        else if(h.indexOf('BILL')!==-1&&colMap.bills===-1) colMap.bills=c3;
        else if(h.indexOf('CONV')!==-1&&colMap.conversion===-1) colMap.conversion=c3;
        else if((h.indexOf('ABV')!==-1||h.indexOf('ATV')!==-1)&&colMap.abv===-1) colMap.abv=c3;
        else if((h.indexOf('WALK')!==-1||h.indexOf('FOOT')!==-1)&&colMap.walkIns===-1) colMap.walkIns=c3;
        else if((h.indexOf('UPT')!==-1||h.indexOf('UNIT')!==-1)&&colMap.upt===-1) colMap.upt=c3;
      }
    }
    if (colMap.sales===-1) colMap.sales=monthStartCol;

    var SKIP=['TOTAL','SUMMARY','GRAND','REGION','★','STORE NAME','S.NO'];
    for(var r=kpiHeaderRow+1;r<rows.length;r++){
      var row2=rows[r];if(!row2||!row2[1])continue;
      var name=String(row2[1]).trim();if(name.length<3)continue;
      if(SKIP.some(function(s){return name.toUpperCase().indexOf(s)!==-1;}))continue;
      if(!name.match(/\([A-Z]{3}[A-Z0-9]+\)/))continue;
      var codeM=name.match(/\(([A-Z0-9]+)\)/);
      var code=codeM?codeM[1]:name.substring(0,8).toUpperCase().replace(/\s/g,'');
      var convRaw=colMap.conversion>=0?parseNum(row2[colMap.conversion]):0;
      var convVal=(convRaw>0&&convRaw<1)?Math.round(convRaw*1000)/10:convRaw;
      targets[code]={
        sales:   colMap.sales>=0   ?parseNum(row2[colMap.sales])  :0,
        bills:   colMap.bills>=0   ?parseNum(row2[colMap.bills])  :0,
        conversion: convVal,
        abv:     colMap.abv>=0     ?parseNum(row2[colMap.abv])    :0,
        walkIns: colMap.walkIns>=0 ?parseNum(row2[colMap.walkIns]):0,
        upt:     colMap.upt>=0     ?parseNum(row2[colMap.upt])    :0
      };
    }
  } catch(err){Logger.log('loadKPITargets: '+err);}
  return targets;
}

// ══════════════════════════════════════════════════════════
//  PARSE SHEET — reads live day-wise data
// ══════════════════════════════════════════════════════════
function parseSheet(ws, sheetName, kpiTargets) {
  var lastRow=ws.getLastRow(), lastCol=ws.getLastColumn();
  if(lastRow<4||lastCol<6) return{stores:[],sheet:sheetName,dayLabels:[],daysCount:0,totalDaysInMonth:30};

  var hdr=ws.getRange(1,1,Math.min(5,lastRow),Math.min(lastCol,140)).getValues();
  var row3=hdr[2]||[], row4=hdr[3]||[];

  // Find SALES columns (exclude MTD summary)
  var salesCols=[];
  for(var c=5;c<row4.length;c++){
    var ft=String(row4[c]||'').replace(/\s/g,'').toUpperCase();
    var lt=String(row3[c]||'').replace(/\s/g,'').toUpperCase();
    if(ft.indexOf('SALES')!==-1&&lt.indexOf('MTD')===-1&&lt.indexOf('SUMMARY')===-1&&lt.indexOf('TOTAL')===-1)
      salesCols.push(c);
  }
  if(salesCols.length===0){
    for(var c2=5;c2<row3.length;c2+=4){
      var v3=String(row3[c2]||'').replace(/\s/g,'').toUpperCase();
      if(v3&&v3.indexOf('MTD')===-1&&v3.indexOf('SUMMARY')===-1)salesCols.push(c2);
    }
  }

  var sheetUp=sheetName.toUpperCase();
  var knownDays=MONTH_DAYS[sheetUp]||30;
  var totalDaysInMonth=Math.min(salesCols.length,knownDays);
  var dayLabels=salesCols.slice(0,totalDaysInMonth).map(function(sc,i){
    return String(row3[sc]||'').trim()||(sheetName+'-'+(i+1));
  });

  // ── Trending cutoff = today - 1 ──
  // Data is entered by stores throughout the day and syncs live.
  // We use today's data IF any store has entered it (todayHasData check).
  // Otherwise fall back to yesterday.
  var todayDay=new Date().getDate();
  var todayIdx=todayDay-1; // 0-based index for today (Apr-15 = index 14)

  // Read store data rows
  var dataRows=ws.getRange(5,1,Math.min(lastRow-4,56),Math.min(lastCol,134)).getValues();

  // Check if ANY store has entered data for today
  var todayHasData=false;
  if(todayIdx<totalDaysInMonth && todayIdx<salesCols.length){
    var todayCol=salesCols[todayIdx];
    for(var ri=0;ri<dataRows.length;ri++){
      if(todayCol<dataRows[ri].length&&parseNum(dataRows[ri][todayCol])>0){
        todayHasData=true; break;
      }
    }
  }

  // cutoff = today if any store entered today's data, else yesterday
  var cutoff = todayHasData ? todayIdx : todayIdx - 1;
  cutoff = Math.max(0, Math.min(cutoff, totalDaysInMonth-1));

  // Find last active day up to cutoff
  var lastIdx=-1;
  for(var di=0;di<=cutoff;di++){
    var sc=salesCols[di],dayTot=0;
    for(var ri2=0;ri2<dataRows.length;ri2++){
      if(sc<dataRows[ri2].length)dayTot+=parseNum(dataRows[ri2][sc]);
    }
    if(dayTot>0)lastIdx=di;
  }
  if(lastIdx===-1)lastIdx=Math.max(0,cutoff);

  var activeCols=salesCols.slice(0,lastIdx+1);
  var activeDayLabels=dayLabels.slice(0,lastIdx+1);
  var lastSalesCol=salesCols[lastIdx]||salesCols[0]||5;
  var daysElapsed=activeCols.length;
  var daysLeft=Math.max(0,totalDaysInMonth-daysElapsed);

  var SKIP=['TOTAL','SUMMARY','GRAND','REGION','★'];
  var stores=[];

  for(var row=0;row<dataRows.length;row++){
    var r=dataRows[row];
    if(!r||!r[1])continue;
    var name=String(r[1]).trim();
    if(name.length<3||name.toUpperCase().indexOf('STORE NAME')!==-1)continue;
    if(SKIP.some(function(s){return name.toUpperCase().indexOf(s)!==-1;}))continue;

    var rm=String(r[2]||'').trim(),cm=String(r[3]||'').trim(),target=parseNum(r[4]);
    var ms=0,mb=0,mq=0,mw=0,daySales=[];

    for(var d=0;d<activeCols.length;d++){
      var col=activeCols[d];
      var sv=col<r.length?parseNum(r[col]):0;
      var bv=col+1<r.length?parseNum(r[col+1]):0;
      var qv=col+2<r.length?parseNum(r[col+2]):0;
      var wv=col+3<r.length?parseNum(r[col+3]):0;
      daySales.push(sv);ms+=sv;mb+=bv;mq+=qv;mw+=wv;
    }

    var todaySales=lastSalesCol<r.length?parseNum(r[lastSalesCol]):0;
    var todayBills=lastSalesCol+1<r.length?parseNum(r[lastSalesCol+1]):0;
    var todayWalk =lastSalesCol+3<r.length?parseNum(r[lastSalesCol+3]):0;
    var pct =target>0?Math.round(ms/target*100):0;
    var abv =mb>0?Math.round(ms/mb):0;
    var upt =mb>0?Math.round(mq/mb*100)/100:0;
    var conv=mw>0?Math.round(mb/mw*1000)/10:0;

    var codeM=name.match(/\(([A-Z0-9]+)\)/);
    var code=codeM?codeM[1]:name.substring(0,8).toUpperCase().replace(/\s/g,'');
    var cityM=code.match(/^([A-Z]{3})/);
    var city=cityM?cityM[1]:'?';
    var kt=kpiTargets[code]||{sales:0,bills:0,conversion:0,abv:0,walkIns:0,upt:0};
    var salesTgt=kt.sales||target;

    function proj(a){return daysElapsed>0?Math.round(a+(a/daysElapsed)*daysLeft):a;}
    function kpiT(a,t){if(!t)return null;var p=proj(a),pc=Math.round(p/t*100);return{actual:Math.round(a*100)/100,projected:p,target:t,pct:pc,status:pc>=100?'green':'red'};}
    function rateT(a,t){if(!t)return null;var pc=Math.round(a/t*100);return{actual:Math.round(a*100)/100,projected:Math.round(a*100)/100,target:t,pct:pc,status:pc>=100?'green':'red'};}

    var projTotal=proj(ms);
    var trendPct=salesTgt>0?Math.round(projTotal/salesTgt*100):0;

    if(target>0||ms>0){
      stores.push({
        name:name,code:code,city:city,rm:rm,cm:cm,
        target:Math.round(salesTgt),mtdSales:Math.round(ms),mtdBills:Math.round(mb),
        mtdSoldQty:Math.round(mq),mtdWalkIns:Math.round(mw),
        todaySales:Math.round(todaySales),todayBills:Math.round(todayBills),todayWalkIns:Math.round(todayWalk),
        pct:pct,abv:abv,upt:upt,convActual:conv,
        projectedTotal:projTotal,trendingPct:trendPct,
        trendStatus:trendPct>=100?'on-track':trendPct>=85?'at-risk':'behind',
        daysCount:daysElapsed,daySales:daySales,
        todayHasData:todayHasData,
        kpiTargets:{sales:Math.round(salesTgt),bills:Math.round(kt.bills||0),conversion:kt.conversion||0,abv:Math.round(kt.abv||0),walkIns:Math.round(kt.walkIns||0),upt:kt.upt||0},
        kpiTrending:{
          sales:kpiT(ms,salesTgt),bills:kpiT(mb,kt.bills),
          walkIns:kpiT(mw,kt.walkIns),conversion:rateT(conv,kt.conversion),
          abv:rateT(abv,kt.abv),upt:rateT(upt,kt.upt)
        }
      });
    }
  }

  var allSheets=SpreadsheetApp.openById(SPREADSHEET_ID).getSheets().map(function(s){return s.getName();});
  return{
    stores:stores,sheet:sheetName,sheets:allSheets,
    dayLabels:activeDayLabels,daysCount:daysElapsed,
    totalDaysInMonth:totalDaysInMonth,cutoffDay:cutoff,
    todayHasData:todayHasData,
    fetchedAt:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd-MMM-yyyy HH:mm:ss')
  };
}

// ══════════════════════════════════════════════════════════
//  Run this ONCE manually to remove old time-driven triggers
// ══════════════════════════════════════════════════════════
function deleteAllTriggers(){
  var t=ScriptApp.getProjectTriggers();
  t.forEach(function(x){ScriptApp.deleteTrigger(x);});
  Logger.log('Removed '+t.length+' trigger(s).');
}

// ── Helpers ──
function parseNum(val){if(val===null||val===undefined||val==='')return 0;if(typeof val==='number')return isNaN(val)?0:val;var s=String(val).replace(/[₹,%\s]/g,'').replace(/[^0-9.-]/g,'');return parseFloat(s)||0;}
function objSlice(obj,n){var o={},i=0;for(var k in obj){if(i++>=n)break;o[k]=obj[k];}return o;}
function updatePresence(u,n){if(!u)return;var d={};try{d=JSON.parse(PROPS.getProperty(USER_KEY)||'{}');}catch(e){}d[u]={name:n||u,ts:Date.now()};PROPS.setProperty(USER_KEY,JSON.stringify(d));}
function removeUser(u){if(!u)return;var d={};try{d=JSON.parse(PROPS.getProperty(USER_KEY)||'{}');}catch(e){}delete d[u];PROPS.setProperty(USER_KEY,JSON.stringify(d));}
function getActiveUsers(){var d={};try{d=JSON.parse(PROPS.getProperty(USER_KEY)||'{}');}catch(e){return[];}var now=Date.now(),a=[],ch=false;Object.keys(d).forEach(function(id){if(now-d[id].ts<USER_TTL)a.push({id:id,name:d[id].name,ts:d[id].ts});else{delete d[id];ch=true;}});if(ch)PROPS.setProperty(USER_KEY,JSON.stringify(d));a.sort(function(a,b){return b.ts-a.ts;});return a;}
//  KA & KL Store Dashboard — Apps Script FINAL (Real-time)
//
//  HOW REAL-TIME SYNC WORKS:
//  1. doGet() reads FRESH data from sheet every single call
//  2. No time-driven trigger needed — delete any you have
//  3. No caching — we add ?t= timestamp to bust CDN cache
//  4. Dashboard polls every 60s — data appears within 1 min
//
//  NO onEdit trigger needed — polling is the correct approach
// ============================================================

var SPREADSHEET_ID = '1Wp_0upauCVviyhw8ZM9XFVbIMRDgEYBZykh1_rjk0S8';
var SECRET_KEY     = '4NJqtPi77ctO3Ec5acEenwwa17XHEFhT';

var PROPS    = PropertiesService.getScriptProperties();
var USER_KEY = 'active_users';
var USER_TTL = 5 * 60 * 1000; // 5 min

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

// ══════════════════════════════════════════════════════════
//  doGet — Entry point. Runs FRESH every call. No caching.
// ══════════════════════════════════════════════════════════
function doGet(e) {
  var p        = (e && e.parameter) ? e.parameter : {};
  var sheet    = p.sheet    || 'APR';
  var callback = p.callback || '';
  var action   = p.action   || 'data';
  var userId   = p.userId   || '';
  var userName = p.userName || '';
  // p.t is a timestamp added by dashboard to bust CDN cache — we ignore it

  // ── API key check ──
  if (p.key !== SECRET_KEY) {
    var denied = JSON.stringify({ error: 'Unauthorized' });
    if (callback) return ContentService.createTextOutput(callback+'('+denied+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(denied).setMimeType(ContentService.MimeType.JSON);
  }

  var result;
  try {
    // openById reads your live sheet directly — always fresh
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (action === 'heartbeat') {
      updatePresence(userId, userName);
      result = { users: getActiveUsers(), ts: Date.now() };

    } else if (action === 'leave') {
      removeUser(userId);
      result = { ok: true };

    } else if (action === 'ping') {
      // Lightweight check — no sheet read, just confirms script is alive
      result = { ok: true, ts: Date.now(), sheet: sheet };

    } else if (action === 'debug_targets') {
      var kt0 = loadKPITargets(ss, sheet);
      result  = { totalStores: Object.keys(kt0).length, kpiSample: objSlice(kt0, 5) };

    } else {
      // ── Main data fetch — always reads live sheet ──
      if (userId) updatePresence(userId, userName);
      var ws = ss.getSheetByName(sheet);
      if (!ws) {
        var names = ss.getSheets().map(function(s){ return s.getName(); });
        result = { error: 'Sheet "'+sheet+'" not found. Available: '+names.join(', ') };
      } else {
        var kpiTargets = loadKPITargets(ss, sheet);
        result = parseSheet(ws, sheet, kpiTargets);
        result.activeUsers = getActiveUsers();
        result.fetchedAt   = Utilities.formatDate(
          new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm:ss'
        );
      }
    }
  } catch(err) {
    result = { error: err.toString() };
  }

  var json = JSON.stringify(result);

  // Return with JAVASCRIPT mime type for JSONP
  // Google does NOT cache JAVASCRIPT responses the same way — this helps avoid stale data
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════
//  DELETE any time-driven triggers you have!
//  This function helps you clean them up — run it once manually
// ══════════════════════════════════════════════════════════
function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  Logger.log('Deleted ' + triggers.length + ' trigger(s). Real-time polling handles sync now.');
}

// ══════════════════════════════════════════════════════════
//  KPI TARGETS — reads STORE TARGETS sheet
// ══════════════════════════════════════════════════════════
function loadKPITargets(ss, sheetName) {
  var targets = {};
  try {
    var ts = ss.getSheetByName('STORE TARGETS');
    if (!ts) return targets;
    var lastRow = Math.min(ts.getLastRow(), 65);
    var lastCol = Math.min(ts.getLastColumn(), 30);
    if (lastRow < 3 || lastCol < 5) return targets;

    var rows = ts.getRange(1, 1, lastRow, lastCol).getValues();
    var sheetUp = sheetName.toUpperCase();

    var monthStartCol = -1, kpiHeaderRow = -1;
    for (var i = 0; i < Math.min(8, rows.length); i++) {
      for (var c = 4; c < rows[i].length; c++) {
        var cv = String(rows[i][c]||'').trim().toUpperCase().replace(/\s+/g,'');
        if (MONTH_NAMES_MAP[cv] === sheetUp) {
          monthStartCol = c;
          for (var j = i+1; j < Math.min(i+4, rows.length); j++) {
            var jt = rows[j].slice(c, Math.min(c+8, rows[j].length))
                            .map(function(v){return String(v||'').toUpperCase();}).join('|');
            if (jt.indexOf('SALES')!==-1||jt.indexOf('BILL')!==-1){kpiHeaderRow=j;break;}
          }
          break;
        }
      }
      if (monthStartCol >= 0) break;
    }
    if (kpiHeaderRow === -1) {
      for (var i2 = 0; i2 < Math.min(8, rows.length); i2++) {
        var rt = rows[i2].slice(0,Math.min(20,rows[i2].length))
                         .map(function(v){return String(v||'').toUpperCase();}).join('|');
        if (rt.indexOf('SALES')!==-1&&(rt.indexOf('BILL')!==-1||rt.indexOf('CONV')!==-1)){
          kpiHeaderRow=i2;
          if (monthStartCol===-1) for(var c2=4;c2<rows[i2].length;c2++){if(String(rows[i2][c2]||'').toUpperCase().indexOf('SALES')!==-1){monthStartCol=c2;break;}}
          break;
        }
      }
    }
    if (monthStartCol===-1){var ORD=['APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC','JAN','FEB','MAR'];var mi=ORD.indexOf(sheetUp);monthStartCol=(mi>=0)?4+mi*6:4;}
    if (kpiHeaderRow===-1) kpiHeaderRow=3;

    var colMap={sales:-1,bills:-1,conversion:-1,abv:-1,walkIns:-1,upt:-1};
    if (kpiHeaderRow<rows.length){
      var hRow=rows[kpiHeaderRow],sf=Math.max(4,monthStartCol-1),st=Math.min(sf+10,hRow.length);
      for(var c3=sf;c3<st;c3++){
        var h=String(hRow[c3]||'').replace(/[\s\n]/g,'').toUpperCase();
        if((h.indexOf('SALES')!==-1||h.indexOf('REVENUE')!==-1)&&colMap.sales===-1) colMap.sales=c3;
        else if(h.indexOf('BILL')!==-1&&colMap.bills===-1) colMap.bills=c3;
        else if(h.indexOf('CONV')!==-1&&colMap.conversion===-1) colMap.conversion=c3;
        else if((h.indexOf('ABV')!==-1||h.indexOf('ATV')!==-1)&&colMap.abv===-1) colMap.abv=c3;
        else if((h.indexOf('WALK')!==-1||h.indexOf('FOOT')!==-1)&&colMap.walkIns===-1) colMap.walkIns=c3;
        else if((h.indexOf('UPT')!==-1||h.indexOf('UNIT')!==-1)&&colMap.upt===-1) colMap.upt=c3;
      }
    }
    if (colMap.sales===-1) colMap.sales=monthStartCol;

    var SKIP=['TOTAL','SUMMARY','GRAND','REGION','★','STORE NAME','S.NO'];
    for(var r=kpiHeaderRow+1;r<rows.length;r++){
      var row2=rows[r];if(!row2||!row2[1])continue;
      var name=String(row2[1]).trim();if(name.length<3)continue;
      if(SKIP.some(function(s){return name.toUpperCase().indexOf(s)!==-1;}))continue;
      if(!name.match(/\([A-Z]{3}[A-Z0-9]+\)/))continue;
      var codeM=name.match(/\(([A-Z0-9]+)\)/);
      var code=codeM?codeM[1]:name.substring(0,8).toUpperCase().replace(/\s/g,'');
      var convRaw=colMap.conversion>=0?parseNum(row2[colMap.conversion]):0;
      var convVal=(convRaw>0&&convRaw<1)?Math.round(convRaw*1000)/10:convRaw;
      targets[code]={
        sales:   colMap.sales>=0   ?parseNum(row2[colMap.sales])  :0,
        bills:   colMap.bills>=0   ?parseNum(row2[colMap.bills])  :0,
        conversion: convVal,
        abv:     colMap.abv>=0     ?parseNum(row2[colMap.abv])    :0,
        walkIns: colMap.walkIns>=0 ?parseNum(row2[colMap.walkIns]):0,
        upt:     colMap.upt>=0     ?parseNum(row2[colMap.upt])    :0
      };
    }
  } catch(err){Logger.log('loadKPITargets: '+err);}
  return targets;
}

// ══════════════════════════════════════════════════════════
//  PARSE SHEET — reads live day-wise data
// ══════════════════════════════════════════════════════════
function parseSheet(ws, sheetName, kpiTargets) {
  var lastRow=ws.getLastRow(), lastCol=ws.getLastColumn();
  if(lastRow<4||lastCol<6) return{stores:[],sheet:sheetName,dayLabels:[],daysCount:0,totalDaysInMonth:30};

  var hdr=ws.getRange(1,1,Math.min(5,lastRow),Math.min(lastCol,140)).getValues();
  var row3=hdr[2]||[], row4=hdr[3]||[];

  // Find SALES columns (exclude MTD summary)
  var salesCols=[];
  for(var c=5;c<row4.length;c++){
    var ft=String(row4[c]||'').replace(/\s/g,'').toUpperCase();
    var lt=String(row3[c]||'').replace(/\s/g,'').toUpperCase();
    if(ft.indexOf('SALES')!==-1&&lt.indexOf('MTD')===-1&&lt.indexOf('SUMMARY')===-1&&lt.indexOf('TOTAL')===-1)
      salesCols.push(c);
  }
  if(salesCols.length===0){
    for(var c2=5;c2<row3.length;c2+=4){
      var v3=String(row3[c2]||'').replace(/\s/g,'').toUpperCase();
      if(v3&&v3.indexOf('MTD')===-1&&v3.indexOf('SUMMARY')===-1)salesCols.push(c2);
    }
  }

  var sheetUp=sheetName.toUpperCase();
  var knownDays=MONTH_DAYS[sheetUp]||30;
  var totalDaysInMonth=Math.min(salesCols.length,knownDays);
  var dayLabels=salesCols.slice(0,totalDaysInMonth).map(function(sc,i){
    return String(row3[sc]||'').trim()||(sheetName+'-'+(i+1));
  });

  // ── Trending cutoff = today - 1 ──
  // Data is entered by stores throughout the day and syncs live.
  // We use today's data IF any store has entered it (todayHasData check).
  // Otherwise fall back to yesterday.
  var todayDay=new Date().getDate();
  var todayIdx=todayDay-1; // 0-based index for today (Apr-15 = index 14)

  // Read store data rows
  var dataRows=ws.getRange(5,1,Math.min(lastRow-4,56),Math.min(lastCol,134)).getValues();

  // Check if ANY store has entered data for today
  var todayHasData=false;
  if(todayIdx<totalDaysInMonth && todayIdx<salesCols.length){
    var todayCol=salesCols[todayIdx];
    for(var ri=0;ri<dataRows.length;ri++){
      if(todayCol<dataRows[ri].length&&parseNum(dataRows[ri][todayCol])>0){
        todayHasData=true; break;
      }
    }
  }

  // cutoff = today if any store entered today's data, else yesterday
  var cutoff = todayHasData ? todayIdx : todayIdx - 1;
  cutoff = Math.max(0, Math.min(cutoff, totalDaysInMonth-1));

  // Find last active day up to cutoff
  var lastIdx=-1;
  for(var di=0;di<=cutoff;di++){
    var sc=salesCols[di],dayTot=0;
    for(var ri2=0;ri2<dataRows.length;ri2++){
      if(sc<dataRows[ri2].length)dayTot+=parseNum(dataRows[ri2][sc]);
    }
    if(dayTot>0)lastIdx=di;
  }
  if(lastIdx===-1)lastIdx=Math.max(0,cutoff);

  var activeCols=salesCols.slice(0,lastIdx+1);
  var activeDayLabels=dayLabels.slice(0,lastIdx+1);
  var lastSalesCol=salesCols[lastIdx]||salesCols[0]||5;
  var daysElapsed=activeCols.length;
  var daysLeft=Math.max(0,totalDaysInMonth-daysElapsed);

  var SKIP=['TOTAL','SUMMARY','GRAND','REGION','★'];
  var stores=[];

  for(var row=0;row<dataRows.length;row++){
    var r=dataRows[row];
    if(!r||!r[1])continue;
    var name=String(r[1]).trim();
    if(name.length<3||name.toUpperCase().indexOf('STORE NAME')!==-1)continue;
    if(SKIP.some(function(s){return name.toUpperCase().indexOf(s)!==-1;}))continue;

    var rm=String(r[2]||'').trim(),cm=String(r[3]||'').trim(),target=parseNum(r[4]);
    var ms=0,mb=0,mq=0,mw=0,daySales=[];

    for(var d=0;d<activeCols.length;d++){
      var col=activeCols[d];
      var sv=col<r.length?parseNum(r[col]):0;
      var bv=col+1<r.length?parseNum(r[col+1]):0;
      var qv=col+2<r.length?parseNum(r[col+2]):0;
      var wv=col+3<r.length?parseNum(r[col+3]):0;
      daySales.push(sv);ms+=sv;mb+=bv;mq+=qv;mw+=wv;
    }

    var todaySales=lastSalesCol<r.length?parseNum(r[lastSalesCol]):0;
    var todayBills=lastSalesCol+1<r.length?parseNum(r[lastSalesCol+1]):0;
    var todayWalk =lastSalesCol+3<r.length?parseNum(r[lastSalesCol+3]):0;
    var pct =target>0?Math.round(ms/target*100):0;
    var abv =mb>0?Math.round(ms/mb):0;
    var upt =mb>0?Math.round(mq/mb*100)/100:0;
    var conv=mw>0?Math.round(mb/mw*1000)/10:0;

    var codeM=name.match(/\(([A-Z0-9]+)\)/);
    var code=codeM?codeM[1]:name.substring(0,8).toUpperCase().replace(/\s/g,'');
    var cityM=code.match(/^([A-Z]{3})/);
    var city=cityM?cityM[1]:'?';
    var kt=kpiTargets[code]||{sales:0,bills:0,conversion:0,abv:0,walkIns:0,upt:0};
    var salesTgt=kt.sales||target;

    function proj(a){return daysElapsed>0?Math.round(a+(a/daysElapsed)*daysLeft):a;}
    function kpiT(a,t){if(!t)return null;var p=proj(a),pc=Math.round(p/t*100);return{actual:Math.round(a*100)/100,projected:p,target:t,pct:pc,status:pc>=100?'green':'red'};}
    function rateT(a,t){if(!t)return null;var pc=Math.round(a/t*100);return{actual:Math.round(a*100)/100,projected:Math.round(a*100)/100,target:t,pct:pc,status:pc>=100?'green':'red'};}

    var projTotal=proj(ms);
    var trendPct=salesTgt>0?Math.round(projTotal/salesTgt*100):0;

    if(target>0||ms>0){
      stores.push({
        name:name,code:code,city:city,rm:rm,cm:cm,
        target:Math.round(salesTgt),mtdSales:Math.round(ms),mtdBills:Math.round(mb),
        mtdSoldQty:Math.round(mq),mtdWalkIns:Math.round(mw),
        todaySales:Math.round(todaySales),todayBills:Math.round(todayBills),todayWalkIns:Math.round(todayWalk),
        pct:pct,abv:abv,upt:upt,convActual:conv,
        projectedTotal:projTotal,trendingPct:trendPct,
        trendStatus:trendPct>=100?'on-track':trendPct>=85?'at-risk':'behind',
        daysCount:daysElapsed,daySales:daySales,
        todayHasData:todayHasData,
        kpiTargets:{sales:Math.round(salesTgt),bills:Math.round(kt.bills||0),conversion:kt.conversion||0,abv:Math.round(kt.abv||0),walkIns:Math.round(kt.walkIns||0),upt:kt.upt||0},
        kpiTrending:{
          sales:kpiT(ms,salesTgt),bills:kpiT(mb,kt.bills),
          walkIns:kpiT(mw,kt.walkIns),conversion:rateT(conv,kt.conversion),
          abv:rateT(abv,kt.abv),upt:rateT(upt,kt.upt)
        }
      });
    }
  }

  var allSheets=SpreadsheetApp.openById(SPREADSHEET_ID).getSheets().map(function(s){return s.getName();});
  return{
    stores:stores,sheet:sheetName,sheets:allSheets,
    dayLabels:activeDayLabels,daysCount:daysElapsed,
    totalDaysInMonth:totalDaysInMonth,cutoffDay:cutoff,
    todayHasData:todayHasData,
    fetchedAt:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd-MMM-yyyy HH:mm:ss')
  };
}

// ══════════════════════════════════════════════════════════
//  Run this ONCE manually to remove old time-driven triggers
// ══════════════════════════════════════════════════════════
function deleteAllTriggers(){
  var t=ScriptApp.getProjectTriggers();
  t.forEach(function(x){ScriptApp.deleteTrigger(x);});
  Logger.log('Removed '+t.length+' trigger(s).');
}

// ── Helpers ──
function parseNum(val){if(val===null||val===undefined||val==='')return 0;if(typeof val==='number')return isNaN(val)?0:val;var s=String(val).replace(/[₹,%\s]/g,'').replace(/[^0-9.-]/g,'');return parseFloat(s)||0;}
function objSlice(obj,n){var o={},i=0;for(var k in obj){if(i++>=n)break;o[k]=obj[k];}return o;}
function updatePresence(u,n){if(!u)return;var d={};try{d=JSON.parse(PROPS.getProperty(USER_KEY)||'{}');}catch(e){}d[u]={name:n||u,ts:Date.now()};PROPS.setProperty(USER_KEY,JSON.stringify(d));}
function removeUser(u){if(!u)return;var d={};try{d=JSON.parse(PROPS.getProperty(USER_KEY)||'{}');}catch(e){}delete d[u];PROPS.setProperty(USER_KEY,JSON.stringify(d));}
function getActiveUsers(){var d={};try{d=JSON.parse(PROPS.getProperty(USER_KEY)||'{}');}catch(e){return[];}var now=Date.now(),a=[],ch=false;Object.keys(d).forEach(function(id){if(now-d[id].ts<USER_TTL)a.push({id:id,name:d[id].name,ts:d[id].ts});else{delete d[id];ch=true;}});if(ch)PROPS.setProperty(USER_KEY,JSON.stringify(d));a.sort(function(a,b){return b.ts-a.ts;});return a;}
