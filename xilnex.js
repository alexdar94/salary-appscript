const ss = SpreadsheetApp.getActiveSpreadsheet();
const sTEMP = ss.getSheetByName('temp');
const CONSTANTS = ss.getSheetByName("Constants");
const sh = ss.getSheetByName(CONSTANTS.getRange('H2').getValue());
sTEMP.getRange(2, 2, sTEMP.getLastRow() - 1, sTEMP.getLastColumn() - 1).setNumberFormat('@STRING@');
const TIMEZONE = Session.getScriptTimeZone();

const G2DATE = sTEMP.getRange('G2').getValue().split('/')
const SHEET_DATE = new Date(`${G2DATE[1]}/${G2DATE[0]}/${G2DATE[2]}`);
const ONE_HR = 3600000;
const MONTH = SHEET_DATE.getMonth() + 1;
const YEAR = SHEET_DATE.getFullYear();
const LAST_DAY_OF_MONTH = new Date(YEAR,MONTH,0).getDate();
const FT_WORK_DAYS = 26;
const ROW_OFFSET = 1;
const COLUMN_OFFSET = 1;
const ARRAY_DAY_OFFSET = ROW_OFFSET + 1;
const ARRAY_TIME_OFFSET = COLUMN_OFFSET + 1;

// Constants sheet values
const SIX_WORKDAY_HRS = (CONSTANTS.getRange('B1').getValue()-CONSTANTS.getRange('B2').getValue()) / (1000 * 60 * 60);
const FIVE_WORKDAY_HRS = (CONSTANTS.getRange('C1').getValue()-CONSTANTS.getRange('B2').getValue()) / (1000 * 60 * 60);
const STAFF_MONTHLY_SALARY = CONSTANTS.getRange('H3').getValue();
const STAFF_HOURLY_SALARY = STAFF_MONTHLY_SALARY === "" ?
                            CONSTANTS.getRange('H4').getValue() : STAFF_MONTHLY_SALARY/FT_WORK_DAYS/SIX_WORKDAY_HRS;
const ALLOWANCE = CONSTANTS.getRange('H5').getValue();
const IS_HOURLY = STAFF_MONTHLY_SALARY === "";
const OT_RATE = STAFF_HOURLY_SALARY * 1.5;
const NS_RATE = CONSTANTS.getRange('B6').getValue(); 
const NS_DATES = expandStringToNumbers(CONSTANTS.getRange('J2').getValue());
const NINEHR_DATES = expandStringToNumbers(CONSTANTS.getRange('J7').getValue());
const MC_DATES = IS_HOURLY ? [] : expandStringToNumbers(CONSTANTS.getRange('J3').getValue()); 
const AL_DATES = IS_HOURLY ? [] : expandStringToNumbers(CONSTANTS.getRange('J4').getValue());
const UL_DATES = IS_HOURLY ? [] : expandStringToNumbers(CONSTANTS.getRange('J5').getValue());
const WARNING_DATES = expandStringToNumbers(CONSTANTS.getRange('J6').getValue());
const PH_DATES = expandStringToNumbers(CONSTANTS.getRange('B7').getValue());
const TRIPLE_DATES = expandStringToNumbers(CONSTANTS.getRange('J8').getValue());
const NS_START = CONSTANTS.getRange('B4').getValue();
const NS_END = CONSTANTS.getRange('B5').getValue();
const WORKLESS_TH = CONSTANTS.getRange('B13').getDisplayValue();

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

  hrs = hrs < 10 ? "0" + hrs : hrs;
  mins = mins < 10 ? "0" + mins : mins;
  return hrs + ':' + mins; 
}

function intToMs(s) {
  return s * 60 * 60 * 1000;
}

function durationStrToInt(s) {
  if (s === "") return 0;
  const [hrs, mins] = s.split(':').map(Number);
  return hrs + (mins / 60);
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

function isCellEmpty(cell) {
  const value = cell.getValue();
  return value === null || value === undefined || value === '';
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
  
function highlightDurationErrors() {
  const lastColumn = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  sh.getRange(1, lastColumn - 4, lastRow, 5).setBackgroundColor("white");

  // Iterate through rows where Column B is not empty
  for (let i = 2; i <= lastRow; i++) {
    if (!isCellEmpty(sh.getRange(i, 2))) {
      // Get the last 5 columns of the current row
      const totalHrsCell = sh.getRange(i, lastColumn - 4); // Fifth last (Total hrs)
      //const normalCell = sh.getRange(i, lastColumn - 3);  // Fourth last (Normal)
      const otCell = sh.getRange(i, lastColumn - 2);      // Third last (OT)
      const nightShCell = sh.getRange(i, lastColumn - 1); // Second last (Night sh)
      //const workLessCell = sh.getRange(i, lastColumn);    // Last (Work less)

      if (isCellEmpty(totalHrsCell)) {
        totalHrsCell.setBackground('red');
      } else {
        // Check if Total hrs has duration more than 12 hours
        if (durationStrToInt(totalHrsCell.getValue()) > 12) { // 12 hours in milliseconds
          totalHrsCell.setBackground('red');
        }
        // Check if OT has duration more than 4.5 hours
        if (!isCellEmpty(otCell) && durationStrToInt(otCell.getValue()) > 4.5) { // 12 hours in milliseconds
          otCell.setBackground('red');
        }
        // Check if Night sh has duration more than 4.5 hours
        if (!isCellEmpty(nightShCell) && durationStrToInt(nightShCell.getValue()) > 4.5) { // 12 hours in milliseconds
          nightShCell.setBackground('red');
        }
      }
    }
  }
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

function clearControl() {
  CONSTANTS.getRange('H2:H5').clearContent();
  CONSTANTS.getRange('J2:J8').clearContent();
}

function drawWeekBorder(mSheet) {
  const mSh = mSheet !== undefined ? mSheet : sh; 
  const columns = mSh.getLastColumn() > 5 ? mSh.getLastColumn() - 4 : mSh.getLastColumn();
  Array.from({length: LAST_DAY_OF_MONTH}, (_, i) => i + 1)
    .forEach(d => {
      if (new Date(`${MONTH}/${d}/${YEAR}`).getDay() === 1)
        mSh.getRange(d + ROW_OFFSET, 1, 1, columns).setBorder(true, null, null, null, null, null);
    });
}

function finalFormat() {
  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  sh.autoResizeColumns(1, lastCol);
  sh.setColumnWidth(lastCol - 3, 5);
  sh.setColumnWidths(2, lastCol - 5, 50);
  sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).setBackgroundColor(null).setFontFamily("Arial");
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
  sh.getRange(1, 2, lastRow, lastCol - 10).setBorder(true, true, true, true, false, false).setNumberFormat('HH:mm');
  // First date row
  sh.getRange(1, 1, 1, lastCol - 4).setBorder(true, true, true, true, null, null);
  drawWeekBorder();
  highlightUL();
  highlightMC();
  highlightAL();
  highlightPH();
  highlightWarning();
}

function calcAllowance(fTotalHrs, fWorkLess) {
  let a = 0;
  let hr = fTotalHrs;
  if (IS_HOURLY) {
    a = hr * ALLOWANCE;
  } else {
    hr = ((FT_WORK_DAYS - MC_DATES.length - UL_DATES.length) * SIX_WORKDAY_HRS) - fWorkLess;
    a = (hr / (FT_WORK_DAYS * SIX_WORKDAY_HRS)) * ALLOWANCE; 
    a = a > ALLOWANCE ? ALLOWANCE : a;
  }
  return ['Allowance', hr, a];
}

function calcTotal() {
  const data = sh.getRange(ROW_OFFSET + 1, sh.getLastColumn() - 4, sh.getLastRow() - ROW_OFFSET, 5).getValues();
  let fTotalHrs = 0;
  let fNormalHrs = 0;
  let fPhHrs = 0;
  let fNormalOT = 0;
  let fPhOT = 0;
  let fTripleHrs = 0;
  let fNightShift = 0;
  let fWorkLess = 0;
  for (i = 0; i < data.length; i++) {
    const [total, normal, ot, nightShift, workLess] = data[i];
    if (total === '') continue;
    fTotalHrs += durationStrToInt(total);
    if (TRIPLE_DATES.includes(i + 1)) {
      fTripleHrs += durationStrToInt(total);
    } else if (PH_DATES.includes(i + 1)) {
      fPhHrs += durationStrToInt(normal);
      fPhOT += durationStrToInt(ot);
    } else {
      fNormalHrs += durationStrToInt(normal);
      fNormalOT += durationStrToInt(ot);
      fNightShift += durationStrToInt(nightShift);
      if (durationStrToInt(workLess) > durationStrToInt(WORKLESS_TH)) fWorkLess += durationStrToInt(workLess);
    }
  }
  const normalOTSal = fNormalOT * STAFF_HOURLY_SALARY * 1.5;
  const phOTSal = fPhOT * STAFF_HOURLY_SALARY * 3;  
  const phSal = fPhHrs * STAFF_HOURLY_SALARY * 2;
  const tripleSal = fTripleHrs * STAFF_HOURLY_SALARY * 3;
  const total = [
    ['', 'Pay rate', !IS_HOURLY ? STAFF_MONTHLY_SALARY : STAFF_HOURLY_SALARY],
    ['Normal hrs', fNormalHrs, fNormalHrs * STAFF_HOURLY_SALARY],
    ['Normal OT', fNormalOT, normalOTSal],
    ['PH hrs', fPhHrs, phSal],
    ['PH OT', fPhOT, phOTSal],
    fTripleHrs !== 0 ? ['Triple hrs', fTripleHrs, tripleSal] : [],
    fNightShift !== 0 ? ['Night shift hrs', fNightShift, IS_HOURLY ? fNightShift * NS_RATE : fNightShift * (NS_RATE - STAFF_HOURLY_SALARY)] : [],
    !IS_HOURLY ? ['Work less', fWorkLess, fWorkLess * STAFF_HOURLY_SALARY] : [],
    calcAllowance(fTotalHrs, fWorkLess),
    [, 'Overtime Total', phSal + normalOTSal + phOTSal + tripleSal]
  ].filter(row => row.length > 0);

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
    const hasNS = NS_DATES.includes(currDate.getDate());
    const isPH = PH_DATES.includes(currDate.getDate());
    const NORMAL_WORK_HOURS = NINEHR_DATES.includes(currDate.getDate()) ? intToMs(FIVE_WORKDAY_HRS) : intToMs(SIX_WORKDAY_HRS);

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

    if (totalWorkHrs === 0) {
      hours[i+1] = ["", "", "", "", ""];
    } else {
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

      hours[i+1] = [
        totalWorkHrs > 0 ? msToTime(totalWorkHrs) : "", 
        normalHrs > 0 ? msToTime(normalHrs) : "",
        otHrs > 0 ? msToTime(otHrs) : "",
        nightShiftHrs > 0 ? msToTime(nightShiftHrs) : "",
        !IS_HOURLY && workLess > 0 ? msToTime(workLess) : "",
      ];
    }
  }
  
  sh.getRange(1, sh.getLastColumn() + 1, sh.getLastRow(), hours[0].length)
  .setNumberFormat("@STRING@").setValues(fillOutRange(hours));
  highlightDurationErrors();
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
  if (toChecks.length > 0) outputSh.getRangeList(toChecks).setBackground("red");
  drawWeekBorder(outputSh);
}

function generateFull() {
  let data = {};
  sTEMP.getRange(2, 3, sTEMP.getLastRow() - 1, 6).getValues().forEach((e, i) => {
    const d = e[4].split("/"); 
    const record = `${d[1]}/${d[0]}/${d[2]} X ${e[5]}`;
    if (data[e[0]] === undefined) {
      data[e[0]] = [record];
    } else {
      data[e[0]] = [...data[e[0]], record];
    }
  });
  
  for (let k in data) {
    if (data.hasOwnProperty(k)) {
      createTimesheet(k, data[k]);
    }
  }
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
  /*createTimesheet('a', ['02/26/2024 Monday 10:07', '02/26/2024 Monday 15:35', '02/26/2024 Monday 16:33', '02/27/2024 Tuesday 13:10','02/27/2024 Tuesday 17:41','02/27/2024 Tuesday 18:47','02/27/2024 Tuesday 21:35','02/27/2024 Tuesday 21:51','02/28/2024 Wednesday 0:36'])*/
  console.log(new Date('03/20/2024'))
}

