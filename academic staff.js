
//-----Set up sheet variable here----
const MAIN_SHEET_NAME = "Duty roster for Sem2 & Summer"
const METADATA_SHEET_NAME = "metadata";
const STAFF_LIST_SHEET_NAME = "Staff List";

//first row/column number(start from 1) of dataArea(contain Name info)
const DATA_AREA_START_ROW = 7;
const DATA_AREA_START_COL = 5;
const DATA_AREA_END_COL = 9;
const DEPT_NAME_ROW = 6;
const DATE_TIME_COL = 3;

const LOOP_START_ROW = 30; //for generateForAll
const LOOP_END_ROW = 40;

// const MORNING_NOON_BOUNDARY_COL = 10;
// const NOON_NIGHT_BOUNDARY_COL = 16;
//------Setup sheet variable end------



const MAIN_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
const METADATA_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(METADATA_SHEET_NAME);
const STAFF_LIST_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_LIST_SHEET_NAME);
let lastRow = MAIN_SHEET.getLastRow();

const url = { };



function installableOnEdit(e) {
    reloadUrl();

    if (
        e.source.getSheetName() === MAIN_SHEET_NAME &&
        e.range.getColumn() >= DATA_AREA_START_COL && e.range.getColumn() <= DATA_AREA_END_COL
    ){
        let row = e.range.getRow();
        let column = e.range.getColumn();
        generateEventForCell(row, column);
    }
    else if (
        e.source.getSheetName() === STAFF_LIST_SHEET_NAME &&
        e.range.getColumn() >= 1 && e.range.getColumn() <= 10
    ) {
        let row = e.range.getRow();
        let column = e.range.getColumn();
        updateCalendarStaffInfo(row, column);
    }
}



function generateEventForAll() {
    reloadUrl();

    for (let i = DATA_AREA_START_ROW; i <= lastRow; i++) {
        if (i < LOOP_START_ROW) continue;
        if (i === LOOP_END_ROW) break;
        for (let j = DATA_AREA_START_COL; j<= DATA_AREA_END_COL; j++) {
            console.log(`${i}//${j}`);
            generateEventForCell(i, j);
        }
    }
}

function  generateEventForCell(row, column) {
    let eventTitle = MAIN_SHEET.getRange(row, column).getValue();
    if (eventTitle.includes('>')) {
        let index = eventTitle.indexOf('>');
        eventTitle = eventTitle.slice(index + 1);
    }

    let deptName = MAIN_SHEET.getRange(DEPT_NAME_ROW, column).getValue();
    let calendarIdCell = "";
    switch (deptName) {
        case 'CIVIL':
            calendarIdCell = "B2";
            break;
        case 'CS':
            calendarIdCell = "B3";
            break;
        case 'EEE':
            calendarIdCell = "B4";
            break;
        case 'IMSE':
            calendarIdCell = "B5";
            break;
        case 'ME':
            calendarIdCell = "B6";
            break;
        default:
            return;
    }

    let calendarID = METADATA_SHEET.getRange(calendarIdCell).getValue();
    let calendar = CalendarApp.getCalendarById(calendarID);
    let eventId = METADATA_SHEET.getRange(row, column).getValue();

    if (eventId.length > 0) {
        let toBeDeleted = calendar.getEventById(eventId);
        try {
            toBeDeleted.deleteEvent();
        } catch (error) {
            console.log("cannot find " + eventId + " in the calendar " + calendar.getName());
        }
        METADATA_SHEET.getRange(row, column).setValue("");
        MAIN_SHEET.getRange(row, column).clearNote();
    }

    if (eventTitle === '' ||
        eventTitle === '/'
    ) return;

    // if staff has (nickname), show only the nickname
    try {
        if (eventTitle.toString().includes('(') && eventTitle.toString().includes(')'))
            eventTitle = eventTitle.match(/\((.+?)\)/g)[0].replace(/[()]/g,''); // extract (nickname) => get rid of "(", ")"
    } catch (e) {
        return;
    }


    let startDate = MAIN_SHEET.getRange(row, DATE_TIME_COL).getValue();


    let startTime_hour = 13;
    let endTime_hour = 17;
    // if (column < MORNING_NOON_BOUNDARY_COL) {
    //     startTime_hour = 9;
    //     endTime_hour = 13;
    // } else if (column < NOON_NIGHT_BOUNDARY_COL) {
    //     startTime_hour = 13;
    //     endTime_hour = 17;
    // } else {
    //     startTime_hour = 17;
    //     endTime_hour = 19;
    // }
    let startDateTime = new Date(new Date(startDate).getTime() + startTime_hour * 60 * 60 * 1000);
    let endDateTime = new Date(new Date(startDate).getTime() + endTime_hour * 60 * 60 * 1000);

    let descriptionUrl = `<a href = ${url[eventTitle]}>Profile<a>`;
    eventId = calendar.createEvent('[' + deptName + '] ' + eventTitle, startDateTime, endDateTime, {description: descriptionUrl}).getId();
    //eventId = calendar.createEvent('[' + deptName + ']' + eventTitle, startDateTime, endDateTime).getId()
    console.log(eventId, deptName, eventTitle, startDateTime, endDateTime);

    METADATA_SHEET.getRange(row, column).setValue(eventId);
    MAIN_SHEET.getRange(row, column).setNote("Synchronized to calendar");
}

function reloadUrl() {
    let _lastRow = STAFF_LIST_SHEET.getLastRow();
    //console.log(MAIN_SHEET.getRange('B' + lastRow).getValue());

    let dataArea = STAFF_LIST_SHEET.getRange("A2:J" + _lastRow).getValues();
    for(let i = 0; i< dataArea.length; i++) {
        let row = dataArea[i];
        let j = 0;
        for (k = 0; k < 5; k++)
        {
            if (row[j] == "") j = j+2;
            else url[row[j++]] = row[j++];
        }
    }
}


function updateCalendarStaffInfo(row, column) {
    let deptNameCol = Math.floor((column % 2) + column - 1);
    let _deptName = STAFF_LIST_SHEET.getRange(1, deptNameCol).getValues().toString();
    let mainSheetColForDept = null;
    let calendarIdCell = null;
    switch (_deptName) {
        case 'CIVIL':
            mainSheetColForDept = 5;
            calendarIdCell = "B2";
            break;
        case 'CS':
            mainSheetColForDept = 6;
            calendarIdCell = "B3";
            break;
        case 'EEE':
            mainSheetColForDept = 7;
            calendarIdCell = "B4";
            break;
        case 'IMSE':
            mainSheetColForDept = 9;
            calendarIdCell = "B5";
            break;
        case 'ME':
            mainSheetColForDept = 8;
            calendarIdCell = "B6";
            break;
        default:
            return;
    }

    for (let i = DATA_AREA_START_ROW; i <= lastRow; i++) {
        let title = MAIN_SHEET.getRange(i,mainSheetColForDept).getValues().toString();
        //Dept Column -> offset = 1; Profile link column -> offset = 0;
        let offset = column % 2;
        if (title === STAFF_LIST_SHEET.getRange(row,column - 1 + offset).getValues().toString())
        {
            generateEventForCell(i,mainSheetColForDept);
        }
    }
}

