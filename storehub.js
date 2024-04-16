const ss = SpreadsheetApp.getActiveSpreadsheet();
const sTEMP = ss.getSheetByName('temp');
const CONSTANTS = ss.getSheetByName("Constants");
const sh = ss.getSheetByName(CONSTANTS.getRange('H2').getValue());
sTEMP.getRange(2, 2, sTEMP.getLastRow() - 1, sTEMP.getLastColumn() - 1).setNumberFormat('@STRING@');
const TIMEZONE = Session.getScriptTimeZone();

const STAFF_MONTHLY_SALARY = CONSTANTS.getRange('H3').getValue();
const STAFF_HOURLY_SALARY = STAFF_MONTHLY_SALARY === "" ?
                            CONSTANTS.getRange('H4').getValue() : STAFF_MONTHLY_SALARY/26/7.5;
const IS_HOURLY = STAFF_MONTHLY_SALARY === "";
const OT_RATE = STAFF_HOURLY_SALARY * 1.5;
const NS_RATE = CONSTANTS.getRange('B6').getValue(); 
const NS_DATES = expandStringToNumbers(CONSTANTS.getRange('H5').getValue());
const MC_DATES = IS_HOURLY ? [] : expandStringToNumbers(CONSTANTS.getRange('H6').getValue()); 
const AL_DATES = IS_HOURLY ? [] : expandStringToNumbers(CONSTANTS.getRange('H7').getValue());
const UL_DATES = IS_HOURLY ? [] : expandStringToNumbers(CONSTANTS.getRange('H8').getValue());
const WARNING_DATES = expandStringToNumbers(CONSTANTS.getRange('H9').getValue());

const SHEET_DATE = new Date(sTEMP.getRange('D3').getValue().split(" ")[0]);
const PH_DATES = expandStringToNumbers(CONSTANTS.getRange('B7').getValue());
const ONE_HR = 3600000;
const MONTH = SHEET_DATE.getMonth() + 1;
const YEAR = SHEET_DATE.getFullYear();
const LAST_DAY_OF_MONTH = new Date(YEAR,MONTH,0).getDate();
const NS_START = CONSTANTS.getRange('B4').getValue();
const NS_END = CONSTANTS.getRange('B5').getValue();
const NORMAL_WORK_HOURS = CONSTANTS.getRange('B1').getValue()-CONSTANTS.getRange('B2').getValue();
const WORKLESS_TH = CONSTANTS.getRange('B13').getDisplayValue();
const ROW_OFFSET = 1;
const COLUMN_OFFSET = 1;
const ARRAY_DAY_OFFSET = ROW_OFFSET + 1;
const ARRAY_TIME_OFFSET = COLUMN_OFFSET + 1;

// Utility -----------------------------------------------------------------------------------
function init() {

}

function fillOutRange(range, fillItem) {
  var fill = (fillItem === undefined)? "" : fillItem;

  //Get the max row length out of all rows in range.
  var initialValue = 0;
  var maxRowLen = range.reduce(function(acc, cur) {
    return Math.max(acc, cur.length);
  }, initialValue);

  //Fill shorter rows to match max with selecte value.
  var filled = range.map(function(row){
    var dif = maxRowLen - row.length;
    if(dif > 0){
      var arizzle = [];
      for(var i = 0; i <  dif; i++){arizzle[i] = fill};
      row = row.concat(arizzle);
    }
    return row;
  });
  return filled;
}

function msToTime(s) {
  var ms = s % 1000;
  s = (s - ms) / 1000;
  var secs = s % 60;
  s = (s - secs) / 60;
  var mins = s % 60;
  var hrs = (s - mins) / 60;

  return hrs + ':' + mins + ':' + secs; // milliSecs are not shown but you can use ms if needed
}

function strDrToInt(str) {
  if (str === "") return 0;
  let d = str.split(':');
  let hr = parseInt(d[0]);
  let min = parseInt(d[1]);
  return hr + (min / 60);
}

function expandStringToNumbers(str) {
  if (str === '') return [];
  return str.split(',').flatMap(s => {
    if(!s.includes('-')) return +s;
    const [min, max] = s.split('-');
    return Array.from({ length: max - min + 1 }, (_, n) => n + +min);
  });
} 

function numToAlphabet(num) {
  return (num + 9).toString(36).toUpperCase();
}

// Utility -----------------------------------------------------------------------------------

function highlightPH() {
  PH_DATES.forEach(p => sh.getRange(parseInt(p) + ROW_OFFSET, 1, 1,  sh.getLastColumn() - 4).setBackgroundColor("#F6B26B"));
}

function highlightMC() {
  MC_DATES.forEach(p => {
    sh.getRange(parseInt(p) + ROW_OFFSET, 1, 1,  sh.getLastColumn() - 4).setBackgroundColor("#CFE2F3");
    sh.getRange(parseInt(p) + ROW_OFFSET, 2).setValue('MC');
  });
}

function highlightAL() {
  AL_DATES.forEach(p => {
    sh.getRange(parseInt(p) + ROW_OFFSET, 1, 1,  sh.getLastColumn() - 4).setBackgroundColor("#FFF2CC");
    sh.getRange(parseInt(p) + ROW_OFFSET, 2).setValue('AL');
  });
}

function highlightUL() {
  UL_DATES.forEach(p => {
    sh.getRange(parseInt(p) + ROW_OFFSET, 1, 1,  sh.getLastColumn() - 4).setBackgroundColor("#E6B8AF");
    sh.getRange(parseInt(p) + ROW_OFFSET, 2).setValue('UL');
  });
}

function highlightWarning() {
  if (WARNING_DATES.length > 0) sh.getRangeList(WARNING_DATES.map(w => `B${w+1}`)).setBackground("red");
}

// If clock time is before OPENING, or between CLOSING_EARLY and CLOSING
function highlightToBeEditedTime() {
  const days = sh.getRange(ARRAY_DAY_OFFSET, ARRAY_TIME_OFFSET, LAST_DAY_OF_MONTH, sh.getLastColumn() - COLUMN_OFFSET).getValues();
  for (i = 0; i < days.length; i++) {
    if (days[i][0] === "") continue;
    const currDate = days[i][0].toLocaleDateString();
    const OPENING = new Date(`${currDate} ${CONSTANTS.getRange('B10').getValue()}`);
    const CLOSING_EARLY = new Date(`${currDate} ${CONSTANTS.getRange('B11').getValue()}`);
    const CLOSING = new Date(`${currDate} ${CONSTANTS.getRange('B12').getValue()}`);
    for (j = 0; j < days[i].length; j++) {
      const curr = days[i][j];
      if (curr && curr < OPENING || (curr < CLOSING && curr >= CLOSING_EARLY)) {
        sh.getRange(i + ARRAY_DAY_OFFSET, j + ARRAY_TIME_OFFSET, 1, 1).setBackgroundColor("red");
      }
    }
  }
}

function drawWeekBorder() {
  const columns = sh.getLastColumn() > 5 ? sh.getLastColumn() - 4 : sh.getLastColumn();
  Array.from({length: LAST_DAY_OF_MONTH}, (_, i) => i + 1)
    .forEach(d => {
      if (new Date(`${MONTH}/${d}/${YEAR}`).getDay() === 1)
        sh.getRange(d + ROW_OFFSET, 1, 1, columns).setBorder(true, null, null, null, null, null);
    });
}

function finalFormat() {
  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  sh.autoResizeColumns(1, lastCol);
  sh.setColumnWidth(lastCol - 3, 5);
  sh.setColumnWidths(2, lastCol - 5, 50);
  sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).setBackgroundColor(null);
  // Work less
  sh.getRange(1, lastCol - 4, lastRow, 1).setBorder(true, true, true, true, false, false);
  // Night shift
  sh.getRange(1, lastCol - 5, lastRow, 1).setBorder(true, true, true, true, false, false);
  // OT
  sh.getRange(1, lastCol - 6, lastRow, 1).setBorder(true, true, true, true, false, false);
  // Normal hrs
  sh.getRange(1, lastCol - 7, lastRow, 1).setBorder(true, true, true, true, false, false);
  // Total hrs
  sh.getRange(1, lastCol - 8, lastRow, 1).setBorder(true, true, true, true, false, false);
  // Date first column
  sh.getRange(1, 1, lastRow, 1).setBorder(true, true, true, true, false, false);
  // Times
  sh.getRange(1, 2, lastRow, lastCol - 10).setBorder(true, true, true, true, false, false);
  // First date row
  sh.getRange(1, 1, 1, lastCol - 4).setBorder(true, true, true, true, null, null);
  drawWeekBorder();
  highlightUL();
  highlightMC();
  highlightAL();
  highlightPH();
  highlightWarning();
}

function calcTotal() {
  const data = sh.getRange(ROW_OFFSET + 1, sh.getLastColumn() - 3, sh.getLastRow() - ROW_OFFSET, 4).getDisplayValues();
  let normalHrs = 0;
  let phHrs = 0;
  let normalOT = 0;
  let phOT = 0;
  let nightShift = 0;
  let workLess = 0;
  for (i = 0; i < data.length; i++) {
    if (PH_DATES.includes(i + 1)) {
      phHrs += strDrToInt(data[i][0]);
      phOT += strDrToInt(data[i][1]);
      continue;
    }
    normalHrs += strDrToInt(data[i][0]);
    normalOT += strDrToInt(data[i][1]);
    nightShift += strDrToInt(data[i][2]);
    if (strDrToInt(data[i][3]) > strDrToInt(WORKLESS_TH)) workLess += strDrToInt(data[i][3]);
  }
  const total = [
    ['', 'Pay rate', !IS_HOURLY ? STAFF_MONTHLY_SALARY : STAFF_HOURLY_SALARY],
    ['Normal hrs', normalHrs, normalHrs * STAFF_HOURLY_SALARY],
    ['PH hrs', phHrs, phHrs * STAFF_HOURLY_SALARY * 2],
    ['Normal OT', normalOT, normalOT * STAFF_HOURLY_SALARY * 1.5],
    ['PH OT', phOT, phOT * STAFF_HOURLY_SALARY * 3],
    ['Night shift hrs', nightShift, IS_HOURLY ? nightShift * NS_RATE : nightShift * (NS_RATE - STAFF_HOURLY_SALARY)],
    !IS_HOURLY ? ['Work less', workLess, workLess * STAFF_HOURLY_SALARY] : [],
    [, 'Overtime Total', (phHrs * STAFF_HOURLY_SALARY * 2) + (normalOT * STAFF_HOURLY_SALARY * 1.5) + (phOT * STAFF_HOURLY_SALARY * 3)]
  ];

  sh.getRange(1, sh.getLastColumn() + 2, total.length, total[0].length).setValues(fillOutRange(total))
  .setNumberFormat("0.00").setBorder(true, true, true, true, true, true);
}

function calculateHours() {
  const days = sh.getRange(ARRAY_DAY_OFFSET, 1, LAST_DAY_OF_MONTH, sh.getLastColumn()).getValues();
  const hours = [];
  hours[0] = ["Total hrs", "Normal hrs", "OT", "Night shift", "Work less"];
  
  for (i = 0; i < days.length; i++) {
    //const currDate = days[i][0].toLocaleDateString();
    //const nsStartDate = new Date(`${currDate} ${NS_START}`);
    const currDate = days[i][0];
    const nsStartDate = new Date(`${currDate.getFullYear()}-${currDate.getMonth() + 1}-${currDate.getDate()} ${NS_START}`);
    const hasNS = NS_DATES.includes(days[i][0].getDate());
    const isPH = PH_DATES.includes(days[i][0].getDate());
    const isAL = AL_DATES.includes(days[i][0].getDate());
    const isMC = MC_DATES.includes(days[i][0].getDate());
    let totalWorkHrs = 0;
    let normalHrs = 0;
    let dayShiftHrs = 0;
    let nightShiftHrs = 0;
    let otHrs = 0;
    let workLess = 0;
  
    for (j = 1; j < days[i].length; j+=2) {
      const clockIn = days[i][j];
      const clockOut = days[i][j+1];
      if (clockIn === "") break;
      const shiftDuration = clockOut - clockIn;

      // Calc total work hrs
      totalWorkHrs = totalWorkHrs + shiftDuration;

      if (clockOut <= nsStartDate) dayShiftHrs = dayShiftHrs + shiftDuration;
      else if (clockIn >= nsStartDate) nightShiftHrs = nightShiftHrs + shiftDuration;
      else {
        dayShiftHrs = dayShiftHrs + (nsStartDate - clockIn);
        nightShiftHrs = nightShiftHrs + (clockOut - nsStartDate);
      }
    }

    // Calc for normal case
    if (totalWorkHrs >= NORMAL_WORK_HOURS) {
      normalHrs = NORMAL_WORK_HOURS;
      otHrs = totalWorkHrs - NORMAL_WORK_HOURS;
    } else if (totalWorkHrs > 0) {
      normalHrs = totalWorkHrs;
      workLess = NORMAL_WORK_HOURS - totalWorkHrs;
    }

    // Calc for night shift case
    if (hasNS && !isPH) {
      if (normalHrs > dayShiftHrs) normalHrs = dayShiftHrs;
      if (OT_RATE >= NS_RATE) {
        nightShiftHrs = otHrs >= nightShiftHrs ? 0 : nightShiftHrs - otHrs;
      } else {
        otHrs = otHrs >= nightShiftHrs ? otHrs - nightShiftHrs : 0;
      }
    } else {
      nightShiftHrs = 0;
    }

    // Calc for MC or AL
    if ((isMC || isAL) && !isPH) {
      totalWorkHrs = NORMAL_WORK_HOURS;
      normalHrs = NORMAL_WORK_HOURS;
    }

    hours[i+1] = [
      totalWorkHrs > 0 ? msToTime(totalWorkHrs) : "", 
      normalHrs > 0 ? msToTime(normalHrs) : "",
      otHrs > 0 ? msToTime(otHrs) : "",
      nightShiftHrs > 0 ? msToTime(nightShiftHrs) : "",
      workLess > 0 ? msToTime(workLess) : "",
    ];
  }
  
  sh.getRange(1, sh.getLastColumn() + 1, sh.getLastRow(), hours[0].length).setValues(fillOutRange(hours));
  sh.getRange(2, 2, sh.getLastRow() - ROW_OFFSET, sh.getLastColumn() - COLUMN_OFFSET).setNumberFormat("HH:mm");
}

function createTimesheet(name, data) {
  const days = [];
  days[0] = ["Date"];
  for (day = 1; day <= LAST_DAY_OF_MONTH; day++) {
    days[day] = [new Date(YEAR, MONTH - 1, day)];
  }

  let toInsertDay = 1; 
  const FIRST = `${YEAR}-${MONTH}-${toInsertDay}`;
  let prevDate = new Date(FIRST);
  let nsEndDate = new Date(`${YEAR}-${MONTH}-${toInsertDay + 1} ${NS_END}`);
  let opening = new Date(`${FIRST} ${CONSTANTS.getRange('B10').getValue()}`);
  let closingEarly = new Date(`${FIRST} ${CONSTANTS.getRange('B11').getValue()}`);
  let closing = new Date(`${FIRST} ${CONSTANTS.getRange('B12').getValue()}`);
  const toChecks = [];
  data.forEach(e => {
    const d = e.split(" ");
    const date = d[0];
    const time = d[2];
    const currDate = new Date(`${date} ${time}`);

    // Move to next day 
    if (currDate > nsEndDate) {
      // Check if less clock record
      if (days[toInsertDay].length % 2 === 0) toChecks.push(`A${toInsertDay + ROW_OFFSET}`);
      
      toInsertDay = currDate.getDate();
      nsEndDate.setDate(toInsertDay + 1);
      opening.setDate(toInsertDay);
      closingEarly.setDate(toInsertDay);
      closing.setDate(toInsertDay);
    }    

    days[toInsertDay] = [...days[toInsertDay], currDate];
    // Check if clock out time < 1hr or early opening/closing
    if ((currDate - prevDate) < ONE_HR
      || currDate < opening 
      || (currDate < closing && currDate >= closingEarly)) 
    toChecks.push(`${numToAlphabet(days[toInsertDay].length)}${toInsertDay + ROW_OFFSET}`);
    
    prevDate = currDate;
    // Add night shift time into timesheet
    /*const nsStartDate = new Date(`${YEAR}-${MONTH}-${toInsertDay} ${NS_START}`);
    const addNS = NS_DATES.includes(toInsertDay)
                    && days[toInsertDay].findIndex(d => d.getTime() === nsStartDate.getTime()) === -1 
                    && currDate > nsStartDate;
    days[toInsertDay] = addNS ? [...days[toInsertDay], nsStartDate, nsStartDate, currDate] : [...days[toInsertDay], currDate];*/
  })

  let maxRowLen = days.reduce((acc, cur) => Math.max(acc, cur.length), 0);

  const outputSh = ss.insertSheet(name);
  outputSh.getRange(1, 1, LAST_DAY_OF_MONTH + ROW_OFFSET, maxRowLen).setValues(fillOutRange(days));
  outputSh.getRange(1, 1, LAST_DAY_OF_MONTH + ROW_OFFSET).setNumberFormat("dd/MM/yyyy");
  outputSh.getRange(2, 2, LAST_DAY_OF_MONTH + ROW_OFFSET, maxRowLen).setNumberFormat("HH:mm");
  outputSh.getRangeList(toChecks).setBackground("red");
}

function generateFull() {
  let curr = '';
  let data = [];
  sTEMP.getRange(2, 2, sTEMP.getLastRow() - 1, sTEMP.getLastColumn() - 1).getValues().forEach((e, i) => {
    if (e[0] !== '') {
      if (curr !== '') createTimesheet(curr, data);
      curr = e[0];
      data = [];
      return;
    }
    if (e[2] !== '') data.push(e[2]);
    if (e[3] !== '') data.push(e[3]);
  });

  // Create for last staff
  createTimesheet(curr, data);
}

function sanitize() {
  for(i = 1; i <= 73; i+=3) {
    sh.getRange('A'+(i+1)).setValue(sh.getRange('A'+i).getValue().substring(19,30));
  }
  const TO_DELETE = [1,3,4,6,7,9,10,12,13,15,16,18,19,21,22,24,25,27,28,30,31,33,34,36,37,39,40,42,43,45,46,48,49,51,52,54,55,57,58,60,61,63,64,66,67,69,70,72,73,75,76,78,79,81,82,84,85,87,88,90,91,93,94,96,97,99,100].reverse();
  for(i = 0; i < TO_DELETE.length; i++) {
    sh.deleteRow(TO_DELETE[i]);
  }
}

function test() {
//sh.getRange('A2').setValue();
  const hasNS = NS_DATES.includes(10);
  const isPH = PH_DATES.includes(10);
  sh.getRange('B1').setValue(hasNS);
  sh.getRange('C1').setValue(isPH);
  sh.getRange('D1').setValue('abc');
}

function main() {
  //sh.getRangeList(['A1:A2']).setBackground("red");
  createTimesheet('a', ['02/26/2024 Monday 10:07', '02/26/2024 Monday 15:35', '02/26/2024 Monday 16:33', '02/27/2024 Tuesday 13:10','02/27/2024 Tuesday 17:41','02/27/2024 Tuesday 18:47','02/27/2024 Tuesday 21:35','02/27/2024 Tuesday 21:51','02/28/2024 Wednesday 0:36'])
}

