// ============================================================
// Apps Script Web App โ€” เธฃเธฑเธเธเนเธญเธกเธนเธฅเธเธฅเธ—เธ”เธชเธญเธเธเธญเธเธเธฃเธตเธ•เธเธฒเธ Cowork
// เธงเธดเธเธต Deploy:
//   1. เน€เธเธดเธ” Sheet "เธเธฅเธ—เธ”เธชเธญเธเธเธญเธเธเธฃเธตเธ•" โ’ Extensions โ’ Apps Script
//   2. เธงเธฒเธเนเธเนเธ”เธเธตเนเธ—เธฑเนเธเธซเธกเธ” โ’ Save
//   3. Deploy โ’ New deployment โ’ Web app
//      - Execute as: Me
//      - Who has access: Anyone
//   4. Copy URL เธ—เธตเนเนเธ”เน โ’ เนเธเนเธเธเธฅเธฑเธเธกเธฒเน€เธเธทเนเธญเนเธชเนเนเธ Cowork skill
// ============================================================

const SHEET_NAME = 'เธเธฅเธ—เธ”เธชเธญเธเธเธญเธเธเธฃเธตเธ•';
const SPREADSHEET_ID = '1ZYGnV8AqyR3a0uTNRftkguEdBPrQOjLQbWU6ZhpiFyk';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);

    if (!sh) {
      return respond(false, 'เนเธกเนเธเธ Sheet: ' + SHEET_NAME);
    }

    // เธซเธฒเนเธ–เธงเธชเธธเธ”เธ—เนเธฒเธขเธเธฒเธ column A เธเธฃเธดเธเน (เนเธกเนเธเธฑเธเธชเธนเธ•เธฃเธ—เธตเนเธฅเธฒเธเธฅเธเนเธ)
    var colA = sh.getRange('A:A').getValues();
    var lastRow = 1;
    for (var i = colA.length - 1; i >= 1; i--) {
      if (colA[i][0] !== '') { lastRow = i + 1; break; }
    }
    var newRow = lastRow + 1;
    sh.getRange(newRow, 1, 1, 11).setValues([[
      data.sample_date,   // A: เธงเธฑเธเธ—เธตเนเน€เธเนเธเธ•เธฑเธงเธญเธขเนเธฒเธ
      data.test_date,     // B: เธงเธฑเธเธ—เธตเนเธ—เธ”เธชเธญเธ
      data.age_days,      // C: เธญเธฒเธขเธธ(เธงเธฑเธ)
      data.formula_name,  // D: เธเธทเนเธญเธชเธนเธ•เธฃ
      data.cube_size,     // E: เธเธเธฒเธ” cube
      data.result1_kn,    // F: เธฅเธนเธ1(kN)
      data.result2_kn,    // G: เธฅเธนเธ2(kN)
      data.result3_kn,    // H: เธฅเธนเธ3(kN)
      data.avg_kn,        // I: เน€เธเธฅเธตเนเธข kN
      data.avg_mpa,       // J: เน€เธเธฅเธตเนเธข MPa
      data.avg_ksc        // K: เน€เธเธฅเธตเนเธข KSC
    ]]);

    Logger.log('เธเธฑเธเธ—เธถเธเธชเธณเน€เธฃเนเธ: ' + JSON.stringify(data));
    return respond(true, 'เธเธฑเธเธ—เธถเธเธชเธณเน€เธฃเนเธ');

  } catch (err) {
    Logger.log('Error: ' + err.message);
    return respond(false, err.message);
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return respond(false, 'เนเธกเนเธเธ Sheet');
    const rows = sh.getDataRange().getValues();
    const headers = rows[0];
    const data = rows.slice(1).filter(r => r[0]).map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        const v = r[i];
        if (v instanceof Date) {
          obj[h] = Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd');
        } else {
          obj[h] = v;
        }
      });
      return obj;
    });

    // เธญเนเธฒเธ tab วัตถุดิบ (เธเธนเธ/เธซเธดเธ/เธ—เธฃเธฒเธข เธ—เธธเธ 10 เธงเธฑเธ)
    var materials = [];
    var matSh = ss.getSheetByName('วัตถุดิบ');
    if (matSh) {
      var matRows = matSh.getDataRange().getValues();
      var matH = matRows[0];
      materials = matRows.slice(1).filter(r => r[0]).map(r => {
        var obj = {};
        matH.forEach((h, i) => { obj[h] = r[i] === '' ? null : r[i]; });
        return obj;
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data: data, materials: materials }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return respond(false, err.message);
  }
}

// ============================================================
// เธเธณเน€เธเนเธฒเธเนเธญเธกเธนเธฅเธเธฒเธ Sheet "เธชเธณเน€เธเธฒเธเธญเธ เธเธณเธเธณเธฅเธฑเธเธญเธฑเธ”เธเนเธญเธเธเธนเธ"
// เธฃเธฑเธ function เธเธตเนเธเธฃเธฑเนเธเน€เธ”เธตเธขเธงเน€เธเธทเนเธญ import เธเนเธญเธกเธนเธฅเน€เธเนเธฒ
// ============================================================
function importFromOldSheet() {
  var SOURCE_ID = '1Ii0Ocr-PAIp1If2R-wMskzVteHkrGwtVuXXpvqKdKz4';
  var TARGET_NAME = 'เธเธฅเธ—เธ”เธชเธญเธเธเธญเธเธเธฃเธตเธ•';

  var src = SpreadsheetApp.openById(SOURCE_ID);
  var srcSheet = src.getSheetByName('เธเนเธฒเธเธณเธญเธฑเธ”เธญเธฑเธ”เธ—เธฑเนเธเธเธต69');
  if (!srcSheet) { Logger.log('เนเธกเนเธเธ sheet เธเนเธฒเธเธณเธญเธฑเธ”เธญเธฑเธ”เธ—เธฑเนเธเธเธต69'); return; }
  Logger.log('เธเธ sheet: ' + srcSheet.getName() + ' (' + (srcSheet.getLastRow()-1) + ' เนเธ–เธง)');
  var rows = srcSheet.getDataRange().getValues();
  var headers = rows[0].map(function(x){ return String(x).trim(); });
  var ci = {};
  headers.forEach(function(n,i){ ci[n] = i; });

  // เธซเธฒ index เธเธญเธฅเธฑเธกเธเน
  var iDate   = ci['เธงเธฑเธเธ—เธตเน'] !== undefined ? ci['เธงเธฑเธเธ—เธตเน'] : 0;
  var iForm   = ci['เธชเธนเธ•เธฃ']  !== undefined ? ci['เธชเธนเธ•เธฃ']  : 1;
  var iSet    = ci['เธเธธเธ”เธ—เธตเน'] !== undefined ? ci['เธเธธเธ”เธ—เธตเน'] : 2;
  var iAge    = ci['เธญเธฒเธขเธธ (เธงเธฑเธ)'] !== undefined ? ci['เธญเธฒเธขเธธ (เธงเธฑเธ)'] : (ci['เธญเธฒเธขเธธ(เธงเธฑเธ)'] !== undefined ? ci['เธญเธฒเธขเธธ(เธงเธฑเธ)'] : 3);
  var iKN     = ci['เนเธฃเธเธเธ” (kN)'] !== undefined ? ci['เนเธฃเธเธเธ” (kN)'] : (ci['เนเธฃเธเธเธ”(kN)'] !== undefined ? ci['เนเธฃเธเธเธ”(kN)'] : 4);
  var iKSC    = ci['เธเธณเธฅเธฑเธเธญเธฑเธ” (KSC)'] !== undefined ? ci['เธเธณเธฅเธฑเธเธญเธฑเธ” (KSC)'] : (ci['เธเธณเธฅเธฑเธเธญเธฑเธ”(KSC)'] !== undefined ? ci['เธเธณเธฅเธฑเธเธญเธฑเธ”(KSC)'] : 5);

  // เธเธฑเธ”เธเธฅเธธเนเธก 3 เธเนเธญเธ
  var groups = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[iDate]) continue;
    var rawDate = r[iDate];
    var age     = Number(r[iAge] || 0);
    var formula = String(r[iForm] || '').trim();
    var set     = String(r[iSet] || '1').trim();
    var kn      = Number(r[iKN]  || 0);
    var ksc     = Number(r[iKSC] || 0);

    // เนเธเธฅเธเธงเธฑเธเธ—เธตเน
    var testDate = '';
    if (rawDate instanceof Date) {
      testDate = Utilities.formatDate(rawDate, 'Asia/Bangkok', 'yyyy-MM-dd');
    } else {
      var parts = String(rawDate).split('/');
      if (parts.length === 3) {
        var y = Number(parts[2]) < 100 ? 2000 + Number(parts[2]) : Number(parts[2]);
        testDate = y + '-' + parts[0].padStart(2,'0') + '-' + parts[1].padStart(2,'0');
      }
    }
    if (!testDate) continue;
    // เน€เธเธเธฒเธฐเธเธต 2026 (เธ.เธจ. 2569) เน€เธ—เนเธฒเธเธฑเนเธ
    if (testDate.slice(0,4) !== '2026') continue;

    // เธงเธฑเธเธ—เธตเนเนเธเธเธตเธ• = เธงเธฑเธเธ—เธตเนเน€เธเนเธเธ•เธฑเธงเธญเธขเนเธฒเธ, เธเธณเธเธงเธ“เธงเธฑเธเธ—เธ”เธชเธญเธ = เธงเธฑเธเน€เธเนเธ + เธญเธฒเธขเธธ
    var sampleDate = testDate;
    var td2 = new Date(sampleDate);
    td2.setDate(td2.getDate() + age);
    testDate = Utilities.formatDate(td2, 'Asia/Bangkok', 'yyyy-MM-dd');

    var key = sampleDate + '|' + formula + '|' + set + '|' + age;
    if (!groups[key]) groups[key] = { sampleDate:sampleDate, testDate:testDate, age:age, formula:formula, kns:[], kscs:[] };
    groups[key].kns.push(kn);
    groups[key].kscs.push(ksc);
  }

  // เธฅเนเธฒเธเธเนเธญเธกเธนเธฅเน€เธ”เธดเธกเธ—เธฑเนเธเธซเธกเธ” (เน€เธเนเธเนเธ–เธงเธซเธฑเธงเนเธงเน)
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(TARGET_NAME);
  if (!sh) { Logger.log('เนเธกเนเธเธ target sheet: ' + TARGET_NAME); return; }
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).clearContent();
  }
  Logger.log('เธฅเนเธฒเธเธเนเธญเธกเธนเธฅเน€เธ”เธดเธกเนเธฅเนเธง');

  var added = 0;
  var keys = Object.keys(groups);
  for (var k = 0; k < keys.length; k++) {
    var g = groups[keys[k]];
    if (g.kns.length < 1) continue;
    var r1 = g.kns[0] || '';
    var r2 = g.kns[1] || '';
    var r3 = g.kns[2] || '';
    var avgKn  = g.kns.reduce(function(a,b){return a+b;},0) / g.kns.length;
    var avgKsc = g.kscs.reduce(function(a,b){return a+b;},0) / g.kscs.length;
    sh.appendRow([
      g.sampleDate,
      g.testDate,
      g.age,
      g.formula,
      '15x15',
      r1, r2, r3,
      Math.round(avgKn*100)/100,
      '',
      Math.round(avgKsc*10)/10
    ]);
    added++;
  }

  Logger.log('เธเธณเน€เธเนเธฒเธชเธณเน€เธฃเนเธ: ' + added + ' เนเธ–เธง');
  try { SpreadsheetApp.getUi().alert('โ… เธเธณเน€เธเนเธฒเธชเธณเน€เธฃเนเธ ' + added + ' เนเธ–เธง'); }
  catch(e) { Logger.log('done'); }
}

// ============================================================
// เน€เธกเธเธนเน€เธฅเธทเธญเธเน€เธ”เธทเธญเธ โ€” เธชเธฃเนเธฒเธเธ•เธญเธเน€เธเธดเธ” Sheet
// ============================================================
function onOpen() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('๐“… เน€เธฅเธทเธญเธเน€เธ”เธทเธญเธ');

  // เธซเธฒเน€เธ”เธทเธญเธเธ—เธฑเนเธเธซเธกเธ”เธ—เธตเนเธกเธตเนเธเธเธตเธ•
  if (sh && sh.getLastRow() > 1) {
    var dates = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    var months = {};
    dates.forEach(function(r) {
      var d = r[0];
      var ym = d instanceof Date
        ? Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM')
        : String(d).slice(0, 7);
      if (ym && ym.length === 7) months[ym] = true;
    });
    var sortedMonths = Object.keys(months).sort().reverse();
    sortedMonths.forEach(function(ym) {
      var parts = ym.split('-');
      var be = Number(parts[0]) + 543;
      var label = 'เน€เธ”เธทเธญเธ ' + parts[1] + '/' + String(be).slice(2) + '  (' + ym + ')';
      menu.addItem(label, 'showMonth_' + ym.replace('-', '_'));
    });
  }

  menu.addSeparator();
  menu.addItem('โ… เนเธชเธ”เธเธ—เธฑเนเธเธซเธกเธ”', 'showAllRows');
  menu.addItem('๐”’ เนเธเนเน€เธ”เธทเธญเธเธฅเนเธฒเธชเธธเธ”', 'showLatestMonthOnly');
  menu.addToUi();
}

// dynamic stubs โ€” Apps Script เธ•เนเธญเธเธเธฒเธฃ function เธเธฃเธดเธ
// เธชเธฃเนเธฒเธเธญเธฑเธ•เนเธเธกเธฑเธ•เธดเธ•เธญเธ onOpen เนเธ•เนเธ•เนเธญเธเธกเธต handler เธเธฅเธฒเธ
function handleMonthMenu(ym) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() <= 1) return;
  var dates = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < dates.length; i++) {
    var d = dates[i][0];
    var rowYM = d instanceof Date
      ? Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM')
      : String(d).slice(0, 7);
    if (rowYM === ym) sh.showRows(i + 2, 1);
    else sh.hideRows(i + 2, 1);
  }
  try { SpreadsheetApp.getUi().alert('โ… เนเธชเธ”เธเน€เธเธเธฒเธฐ ' + ym); } catch(e) {}
}

function showAllRows() {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (sh && sh.getLastRow() > 1) sh.showRows(2, sh.getLastRow() - 1);
  try { SpreadsheetApp.getUi().alert('โ… เนเธชเธ”เธเธ—เธฑเนเธเธซเธกเธ”เนเธฅเนเธง'); } catch(e) {}
}

// stubs เธชเธณเธซเธฃเธฑเธเนเธ•เนเธฅเธฐเน€เธ”เธทเธญเธ (เน€เธเธดเนเธกเนเธ”เนเธ–เนเธฒเธกเธตเน€เธ”เธทเธญเธเนเธซเธกเน)
function showMonth_2026_02(){handleMonthMenu('2026-02');}
function showMonth_2026_03(){handleMonthMenu('2026-03');}
function showMonth_2026_04(){handleMonthMenu('2026-04');}
function showMonth_2026_05(){handleMonthMenu('2026-05');}
function showMonth_2026_06(){handleMonthMenu('2026-06');}
function showMonth_2026_07(){handleMonthMenu('2026-07');}
function showMonth_2026_08(){handleMonthMenu('2026-08');}
function showMonth_2026_09(){handleMonthMenu('2026-09');}
function showMonth_2026_10(){handleMonthMenu('2026-10');}
function showMonth_2026_11(){handleMonthMenu('2026-11');}
function showMonth_2026_12(){handleMonthMenu('2026-12');}

// ============================================================
// เธเนเธญเธเนเธ–เธงเน€เธ”เธทเธญเธเธเนเธญเธ เนเธชเธ”เธเนเธเนเน€เธ”เธทเธญเธเธฅเนเธฒเธชเธธเธ”
// ============================================================
function showLatestMonthOnly() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { Logger.log('เนเธกเนเธเธ sheet'); return; }

  var lastRow = sh.getLastRow();
  if (lastRow <= 1) return; // เธกเธตเนเธเนเธซเธฑเธง

  // เธซเธฒเน€เธ”เธทเธญเธเธฅเนเธฒเธชเธธเธ”เธเธฒเธเธเธญเธฅเธฑเธกเธเน A (เธงเธฑเธเธ—เธตเนเน€เธเนเธเธ•เธฑเธงเธญเธขเนเธฒเธ)
  var dates = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  var latestYM = '';
  dates.forEach(function(r) {
    var d = r[0];
    var ym = d instanceof Date
      ? Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM')
      : String(d).slice(0, 7);
    if (ym > latestYM) latestYM = ym;
  });

  if (!latestYM) return;
  Logger.log('เน€เธ”เธทเธญเธเธฅเนเธฒเธชเธธเธ”: ' + latestYM);

  // เนเธชเธ”เธ/เธเนเธญเธเธ—เธตเธฅเธฐเนเธ–เธง
  for (var i = 0; i < dates.length; i++) {
    var d = dates[i][0];
    var ym = d instanceof Date
      ? Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM')
      : String(d).slice(0, 7);
    var row = i + 2; // +2 เน€เธเธฃเธฒเธฐเน€เธฃเธดเนเธกเนเธ–เธง 2
    if (ym === latestYM) {
      sh.showRows(row, 1);
    } else {
      sh.hideRows(row, 1);
    }
  }

  Logger.log('เน€เธชเธฃเนเธเธชเธดเนเธ โ€” เนเธชเธ”เธเน€เธเธเธฒเธฐ ' + latestYM);
  try { SpreadsheetApp.getUi().alert('โ… เนเธชเธ”เธเน€เธเธเธฒเธฐเน€เธ”เธทเธญเธ ' + latestYM); }
  catch(e) { Logger.log('done'); }
}

// ============================================================
// เนเธเนเธเนเธฒเน€เธเธฅเธตเนเธข kN เนเธฅเธฐ KSC เธ—เธธเธเนเธ–เธงเนเธซเนเธ–เธนเธเธ•เนเธญเธ
// เธเธญเธฅเธฑเธกเธเน: F=เธฅเธนเธ1, G=เธฅเธนเธ2, H=เธฅเธนเธ3, I=เน€เธเธฅเธตเนเธขkN, K=เน€เธเธฅเธตเนเธขKSC
// ============================================================
function fixAllKscValues() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { Logger.log('เนเธกเนเธเธ sheet'); return; }

  var lastRow = sh.getLastRow();
  if (lastRow <= 1) return;

  var range = sh.getRange(2, 1, lastRow - 1, 11); // A2:K
  var values = range.getValues();
  var fixed = 0;

  for (var i = 0; i < values.length; i++) {
    var kn1 = Number(values[i][5]) || 0; // F
    var kn2 = Number(values[i][6]) || 0; // G
    var kn3 = Number(values[i][7]) || 0; // H

    if (kn1 > 0 && kn2 > 0 && kn3 > 0) {
      var avgKn  = Math.round((kn1 + kn2 + kn3) / 3 * 100) / 100;
      var avgMpa = Math.round(avgKn / 22.5 * 100) / 100;
      var avgKsc = Math.round(avgKn / 22.5 * 10.197 * 10) / 10;

      values[i][8]  = avgKn;  // I: เน€เธเธฅเธตเนเธข kN
      values[i][9]  = avgMpa; // J: เน€เธเธฅเธตเนเธข MPa
      values[i][10] = avgKsc; // K: เน€เธเธฅเธตเนเธข KSC
      fixed++;
    }
  }

  range.setValues(values);
  Logger.log('เนเธเนเนเธเนเธฅเนเธง ' + fixed + ' เนเธ–เธง');
  try { SpreadsheetApp.getUi().alert('โ… เนเธเนเนเธเนเธฅเนเธง ' + fixed + ' เนเธ–เธง'); }
  catch(e) { Logger.log('done'); }
}

// ============================================================
// เธชเธฃเนเธฒเธ tab "วัตถุดิบ" เธเธฃเนเธญเธก header โ€” เธฃเธฑเธเธเธฃเธฑเนเธเน€เธ”เธตเธขเธง
// ============================================================
function setupMaterialsSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('วัตถุดิบ');
  if (!sh) {
    sh = ss.insertSheet('วัตถุดิบ');
    Logger.log('เธชเธฃเนเธฒเธ sheet วัตถุดิบ เนเธฅเนเธง');
  } else {
    Logger.log('sheet วัตถุดิบ เธกเธตเธญเธขเธนเนเนเธฅเนเธง');
  }
  // เนเธชเน header เธ–เนเธฒเธขเธฑเธเนเธกเนเธกเธต
  var headers = ['เธเนเธงเธ','เธเธนเธเธฃเธงเธก','เธเธนเธเน€เธชเธฒเนเธซเธเน','เธเธนเธI18','เธซเธดเธ3/4','เธซเธดเธ1','เธ—เธฃเธฒเธข'];
  var firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  var hasHeader = firstRow[0] !== '';
  if (!hasHeader) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    Logger.log('เนเธชเน header เนเธฅเนเธง');
  }
  try { SpreadsheetApp.getUi().alert('โ… เธเธฃเนเธญเธกเนเธฅเนเธง! เธเธฃเธญเธเธเนเธญเธกเธนเธฅเนเธ tab "วัตถุดิบ" เนเธ”เนเน€เธฅเธข\n\nHeader: เธเนเธงเธ | เธเธนเธเธฃเธงเธก | เธเธนเธเน€เธชเธฒเนเธซเธเน | เธเธนเธI18 | เธซเธดเธ3/4 | เธซเธดเธ1 | เธ—เธฃเธฒเธข'); }
  catch(e) { Logger.log('done'); }
}

// Auto-fill เน€เธ”เธทเธญเธเธเธต (เธเธญเธฅเธฑเธกเธเน L) เน€เธกเธทเนเธญเธเธฃเธญเธเธงเธฑเธเธ—เธตเนเน€เธเนเธเธ•เธฑเธงเธญเธขเนเธฒเธ (เธเธญเธฅเธฑเธกเธเน A)
function onEdit(e) {
  if (!e || !e.range) return;
  var range = e.range;
  var sh = range.getSheet();
  if (sh.getName() !== SHEET_NAME) return;
  if (range.getColumn() !== 1) return; // เธเธญเธฅเธฑเธกเธเน A เน€เธ—เนเธฒเธเธฑเนเธ
  var row = range.getRow();
  if (row <= 1) return; // เธเนเธฒเธกเธซเธฑเธงเธ•เธฒเธฃเธฒเธ
  var val = range.getValue();
  if (!val) return;
  var d = new Date(val);
  if (isNaN(d.getTime())) return;
  var ym = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM');
  sh.getRange(row, 12).setValue(ym); // เธเธญเธฅเธฑเธกเธเน L = 12
}

function respond(success, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: success, message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

