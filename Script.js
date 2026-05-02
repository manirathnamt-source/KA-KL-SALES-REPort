// ════════════════════════════════════════════════════════════
//  KA & KL MAIN KPI SCRIPT
//  Sheet: 1Wp_0upauCVviyhw8ZM9XFVbIMRDgEYBZykh1_rjk0S8
//
//  SHEET LAYOUT (same for all months):
//  Row 1: Title row
//  Row 2: Blank / totals info
//  Row 3: Header row: S.No | STORE NAME | RM | CM | Target | ... | Apr-1 | | | | Apr-2 | ...
//  Row 4: Sub-header: Sales | Conv | ABV | Walk-ins | UPT | SALES(₹) | BILLS | SOLD QTY | WALK INS | ...
//  Row 5+: Store data rows
//
//  Col layout (1-indexed):
//  1=S.No, 2=Store Name, 3=RM, 4=CM, 5=Sales Target, 6=Conv Target
//  7=ABV Target, 8=Walk-ins Target, 9=UPT Target
//  Then groups of 4 per day: SALES | BILLS | SOLD QTY | WALK INS
//  Day 1 starts at col 10, Day 2 at col 14, Day N at col 10+(N-1)*4
//  MTD Summary at end (after last day)
//
//  DEPLOY: Apps Script → Deploy → Web App → Anyone
// ════════════════════════════════════════════════════════════

var SHEET_ID = '1Wp_0upauCVviyhw8ZM9XFVbIMRDgEYBZykh1_rjk0S8';
var MONTH_ORDER = ['APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC','JAN','FEB','MAR'];
var MONTH_DAYS  = {APR:30,MAY:31,JUN:30,JUL:31,AUG:31,SEP:30,OCT:31,NOV:30,DEC:31,JAN:31,FEB:28,MAR:31};
var MONTH_ABBR  = {APR:'Apr',MAY:'May',JUN:'Jun',JUL:'Jul',AUG:'Aug',SEP:'Sep',OCT:'Oct',NOV:'Nov',DEC:'Dec',JAN:'Jan',FEB:'Feb',MAR:'Mar'};

// FY starts April — quarter mapping
var QUARTER_MAP = {APR:'Q1',MAY:'Q1',JUN:'Q1',JUL:'Q2',AUG:'Q2',SEP:'Q2',OCT:'Q3',NOV:'Q3',DEC:'Q3',JAN:'Q4',FEB:'Q4',MAR:'Q4'};

function doGet(e){
  var p=e&&e.parameter?e.parameter:{}, cb=p.callback||'';
  var result;
  try{
    var action=p.action||'monthly';
    if(action==='ytd')         result=getYTD(p.upToMonth);
    else if(action==='monthly') result=getMonthly(p.sheet||'APR');
    else                        result=getMonthly(p.sheet||'APR');
  }catch(err){result={error:err.toString()};}
  var json=JSON.stringify(result);
  if(cb) return ContentService.createTextOutput(cb+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ── Monthly data (existing behaviour) ──
function getMonthly(monthName){
  var ss=SpreadsheetApp.openById(SHEET_ID);
  // Find sheet by name (case-insensitive)
  var sheet=null;
  var sheets=ss.getSheets();
  for(var i=0;i<sheets.length;i++){
    if(sheets[i].getName().toUpperCase()===monthName.toUpperCase()){
      sheet=sheets[i];break;
    }
  }
  if(!sheet) return{error:'Sheet "'+monthName+'" not found. Sheets: '+sheets.map(function(s){return s.getName();}).join(', ')};

  var data=sheet.getDataRange().getValues();
  if(data.length<5) return{error:'Sheet has too few rows'};

  // Find the header rows
  // Row index 2 (0-based) = S.No | STORE NAME | RM | CM | Target | ... | Apr-1 | ...
  // Row index 3 (0-based) = Sales | Conv | ABV | Walk-ins | UPT | SALES(₹) | BILLS | SOLD QTY | WALK INS | ...
  // Find which row has "S.No" or "STORE NAME"
  var headerRowIdx=-1, subHeaderRowIdx=-1, firstDataRowIdx=-1;
  for(var r=0;r<Math.min(data.length,10);r++){
    var cell0=String(data[r][0]||'').trim().toLowerCase();
    var cell1=String(data[r][1]||'').trim().toLowerCase();
    if(cell0==='s.no'||cell1.indexOf('store name')!==-1||cell1.indexOf('store')!==-1){
      headerRowIdx=r;
      subHeaderRowIdx=r+1;
      firstDataRowIdx=r+2;
      break;
    }
  }
  if(headerRowIdx<0){
    // Fallback: look for row with "Apr-" or "May-" day labels
    for(var r2=0;r2<Math.min(data.length,8);r2++){
      var rowStr=data[r2].join('|').toLowerCase();
      if(rowStr.indexOf('apr-')+rowStr.indexOf('may-')+rowStr.indexOf('jun-')>-3){
        headerRowIdx=r2;subHeaderRowIdx=r2+1;firstDataRowIdx=r2+2;break;
      }
    }
  }
  if(headerRowIdx<0) return{error:'Could not find header row. Check sheet structure.'};

  var headerRow=data[headerRowIdx];
  var today=new Date();
  var totalDays=MONTH_DAYS[monthName.toUpperCase()]||30;

  // Find day columns by looking for "Mon-N" patterns in header row
  // Each day group = [SALES, BILLS, SOLD QTY, WALK INS]
  var dayLabels=[];
  var dayStartCols=[]; // 0-based col index where each day's SALES column is

  for(var c=0;c<headerRow.length;c++){
    var h=String(headerRow[c]||'').trim();
    // Match "Apr-1", "Apr-2", "May-1" etc
    if(h.match(/^[A-Za-z]{3}-\d{1,2}$/)){
      dayLabels.push(h);
      dayStartCols.push(c);
    }
  }

  // If no day labels found, try to parse from target area
  // Col 5 = Sales Target (idx 4), then col 10 = Day 1 (idx 9)
  if(dayLabels.length===0){
    Logger.log('No day labels found in header. Using fixed layout: day 1 at col 10 (idx 9)');
    var startCol=9; // 0-based index of Apr-1 SALES column
    var mabbr=MONTH_ABBR[monthName.toUpperCase()]||monthName;
    for(var d=1;d<=totalDays;d++){
      dayLabels.push(mabbr+'-'+d);
      dayStartCols.push(startCol+(d-1)*4);
    }
  }

  Logger.log('Month: '+monthName+' | Days found: '+dayLabels.length+' | Day cols: '+dayStartCols.slice(0,3).join(','));

  // Find today's elapsed days
  var elapsed=dayLabels.length; // use actual data days available
  var remaining=Math.max(0,totalDays-elapsed);

  // Parse stores
  var stores=[];
  for(var r3=firstDataRowIdx;r3<data.length;r3++){
    var row=data[r3];
    var sno=String(row[0]||'').trim();
    var storeFull=String(row[1]||'').trim();
    if(!storeFull||storeFull.toUpperCase().indexOf('WAREHOUSE')!==-1||sno==='S.No'||sno==='') continue;
    if(!sno||isNaN(parseInt(sno))) continue; // skip non-data rows

    // Extract store code from name like "BLR - JAYNAGAR (BLRJAY)"
    var code=storeFull;
    var codeMatch=storeFull.match(/\(([A-Z0-9]{4,8})\)$/);
    if(codeMatch) code=codeMatch[1];

    var rm   =String(row[2]||'').trim();
    var cm   =String(row[3]||'').trim();

    // Targets at cols 5-9 (idx 4-8)
    var tgtSales  =num(row[4]);
    var tgtConv   =parseFloat(String(row[5]||'').replace(/[^0-9.]/g,''))||0; // "53%"
    var tgtABV    =num(row[6]);
    var tgtWalkIns=num(row[7]);
    var tgtUPT    =parseFloat(String(row[8]||'').replace(/[^0-9.]/g,''))||0;
    var tgtBills  =Math.round(tgtWalkIns*(tgtConv/100));

    // Day-wise data
    var daySales=[], dayBills=[], dayWalkIns=[], dayQty=[];
    for(var di=0;di<dayStartCols.length;di++){
      var c2=dayStartCols[di];
      daySales.push(num(row[c2]));
      dayBills.push(num(row[c2+1]));
      dayQty.push(num(row[c2+2]));
      dayWalkIns.push(num(row[c2+3]));
    }

    // MTD calculations
    var mtdSales=sum(daySales), mtdBills=sum(dayBills), mtdQty=sum(dayQty), mtdWalkIns=sum(dayWalkIns);
    var daysWithData=daySales.filter(function(v){return v>0;}).length;
    if(daysWithData===0) daysWithData=1;

    var projTotal=Math.round(mtdSales/daysWithData*totalDays);
    var pct=tgtSales?Math.round(mtdSales/tgtSales*100):0;
    var trendingPct=tgtSales?Math.round(projTotal/tgtSales*100):0;
    var convActual=mtdWalkIns>0?Math.round(mtdBills/mtdWalkIns*1000)/10:0;
    var abv=mtdBills>0?Math.round(mtdSales/mtdBills):0;
    var upt=mtdBills>0?Math.round(mtdQty/mtdBills*100)/100:0;

    stores.push({
      code:code, name:storeFull, rm:rm, cm:cm,
      target:tgtSales, mtdSales:mtdSales, projectedTotal:projTotal,
      pct:pct, trendingPct:trendingPct,
      mtdBills:mtdBills, mtdWalkIns:mtdWalkIns, mtdSoldQty:mtdQty,
      convActual:convActual, abv:abv, upt:upt,
      todaySales:daySales[daysWithData-1]||0,
      daysCount:daysWithData,
      daySales:daySales, dayBills:dayBills,
      kpiTargets:{
        sales:tgtSales, bills:tgtBills, walkIns:tgtWalkIns,
        conversion:tgtConv, abv:tgtABV, upt:tgtUPT
      },
      kpiTrending:{
        walkIns:{pct:tgtWalkIns?Math.round(mtdWalkIns/tgtWalkIns*100):null}
      }
    });
  }

  Logger.log('Stores parsed: '+stores.length+' | days: '+dayLabels.length);
  return{
    stores:stores,
    dayLabels:dayLabels,
    totalDaysInMonth:totalDays,
    elapsed:daysWithData||elapsed,
    remaining:remaining,
    sheets:getAvailableSheets(ss),
    month:monthName,
    generatedAt:new Date().toISOString()
  };
}

// ── YTD data — reads all months from APR to upToMonth ──
function getYTD(upToMonth){
  upToMonth=(upToMonth||'APR').toUpperCase();
  var ss=SpreadsheetApp.openById(SHEET_ID);

  // Determine which months to include (APR = start of FY)
  var upToIdx=MONTH_ORDER.indexOf(upToMonth);
  if(upToIdx<0) upToIdx=0;
  var monthsToRead=MONTH_ORDER.slice(0,upToIdx+1);

  Logger.log('YTD: reading months '+monthsToRead.join(', '));

  // Read each month and aggregate per store
  var storeAgg={}; // code → {name,rm,cm,sales:{},targets:{},bills:{},walkIns:{},qty:{}}

  monthsToRead.forEach(function(m){
    var mData=getMonthly(m);
    if(!mData||mData.error||!mData.stores) return;
    mData.stores.forEach(function(s){
      var key=s.code;
      if(!storeAgg[key])storeAgg[key]={
        code:s.code,name:s.name,rm:s.rm,cm:s.cm,
        ytdSales:0,ytdTarget:0,ytdBills:0,ytdWalkIns:0,ytdQty:0,
        monthlyData:{},targets:{}
      };
      var a=storeAgg[key];
      a.ytdSales  +=s.mtdSales;
      a.ytdTarget +=s.target||0;
      a.ytdBills  +=s.mtdBills;
      a.ytdWalkIns+=s.mtdWalkIns;
      a.ytdQty    +=s.mtdSoldQty;
      a.monthlyData[m]={sales:s.mtdSales,target:s.target||0,pct:s.pct,bills:s.mtdBills,walkIns:s.mtdWalkIns};
      a.targets[m]=s.kpiTargets||{};
    });
  });

  // Format output
  var stores=Object.values(storeAgg).map(function(a){
    var pct=a.ytdTarget?Math.round(a.ytdSales/a.ytdTarget*100):0;
    var abv=a.ytdBills?Math.round(a.ytdSales/a.ytdBills):0;
    var upt=a.ytdBills?Math.round(a.ytdQty/a.ytdBills*100)/100:0;
    var conv=a.ytdWalkIns?Math.round(a.ytdBills/a.ytdWalkIns*1000)/10:0;
    return{
      code:a.code,name:a.name,rm:a.rm,cm:a.cm,
      ytdSales:R(a.ytdSales),ytdTarget:R(a.ytdTarget),
      ytdBills:R(a.ytdBills),ytdWalkIns:R(a.ytdWalkIns),ytdQty:R(a.ytdQty),
      pct:pct,abv:abv,upt:upt,convActual:conv,
      monthlyData:a.monthlyData
    };
  }).sort(function(a,b){return b.ytdSales-a.ytdSales;});

  // RM/CM aggregates
  var rmAgg={},cmAgg={};
  stores.forEach(function(s){
    if(!rmAgg[s.rm])rmAgg[s.rm]={name:s.rm,ytdSales:0,ytdTarget:0,ytdBills:0,ytdWalkIns:0,stores:0};
    if(!cmAgg[s.cm])cmAgg[s.cm]={name:s.cm,ytdSales:0,ytdTarget:0,ytdBills:0,ytdWalkIns:0,stores:0};
    rmAgg[s.rm].ytdSales+=s.ytdSales;rmAgg[s.rm].ytdTarget+=s.ytdTarget;rmAgg[s.rm].stores++;
    cmAgg[s.cm].ytdSales+=s.ytdSales;cmAgg[s.cm].ytdTarget+=s.ytdTarget;cmAgg[s.cm].stores++;
    rmAgg[s.rm].ytdBills+=s.ytdBills;cmAgg[s.cm].ytdBills+=s.ytdBills;
    rmAgg[s.rm].ytdWalkIns+=s.ytdWalkIns;cmAgg[s.cm].ytdWalkIns+=s.ytdWalkIns;
  });
  ['rmAgg','cmAgg'].forEach(function(key){
    var obj=key==='rmAgg'?rmAgg:cmAgg;
    Object.values(obj).forEach(function(g){
      g.pct=g.ytdTarget?Math.round(g.ytdSales/g.ytdTarget*100):0;
      g.abv=g.ytdBills?Math.round(g.ytdSales/g.ytdBills):0;
      g.conv=g.ytdWalkIns?Math.round(g.ytdBills/g.ytdWalkIns*1000)/10:0;
    });
  });

  var quarter=QUARTER_MAP[upToMonth]||'Q1';
  return{
    stores:stores,
    rm:Object.values(rmAgg).sort(function(a,b){return b.ytdSales-a.ytdSales;}),
    cm:Object.values(cmAgg).sort(function(a,b){return b.ytdSales-a.ytdSales;}),
    months:monthsToRead,
    upToMonth:upToMonth,
    quarter:quarter,
    generatedAt:new Date().toISOString()
  };
}

function getAvailableSheets(ss){
  return ss.getSheets()
    .map(function(s){return s.getName().toUpperCase();})
    .filter(function(n){return MONTH_ORDER.indexOf(n)!==-1;});
}

function num(v){if(!v&&v!==0)return 0;if(typeof v==='number')return isNaN(v)?0:v;var s=String(v).replace(/[₹,\s]/g,'');return parseFloat(s)||0;}
function sum(arr){return(arr||[]).reduce(function(a,b){return a+(b||0);},0);}
function R(n){return Math.round(n);}
