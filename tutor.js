
//-----Set up sheet variable here----
const MAIN_SHEET_NAME = "ITTS duty roster (Sem2 & Summer)"
const METADATA_SHEET_NAME = "metadata";

//first row/column number(start from 1) of dataArea(contain Name info)
const DATA_AREA_START_ROW = 18;
const DATA_AREA_START_COL = 5;
const DATA_AREA_END_COL = 21;
//const DEPT_NAME_ROW = 17;
const DTAE_TIME_COL = 3;


const MORNING_NOON_BOUNDARY_COL = 10;
const NOON_NIGHT_BOUNDARY_COL = 16;

const LOOP_START_ROW = 97; //for generateForAll
const LOOP_END_ROW = 108;
//------Setup sheet variable end------



let mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
let subSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(METADATA_SHEET_NAME);
let lastRow = mainSheet.getLastRow();


function installableOnEdit(e) {
    if (
        e.source.getSheetName() === MAIN_SHEET_NAME &&
        e.range.getColumn() >= DATA_AREA_START_COL && e.range.getColumn() <= DATA_AREA_END_COL
    ){

        let row = e.range.getRow();
        let column = e.range.getColumn();
        generateEventForCell(row, column);
    }
}


function generateEventForAll() {


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

    let eventTitle = mainSheet.getRange(row, column).getValue();

    if (eventTitle === '' ||
        eventTitle === '/' ||
        eventTitle.toString().toUpperCase().includes('(ABSEN') ||
        eventTitle.toString().toUpperCase().includes('(CAN')
    ) return;

    // if (eventTitle.includes('>')) {
    //     let index = eventTitle.indexOf('>');
    //     eventTitle = eventTitle.slice(index + 1);
    // }

    //let deptName = mainSheet.getRange(DEPT_NAME_ROW, column).getValue();

    // match dept name with the prefix
    let deptName = "";
    try {
        deptName = eventTitle.match(/\[(.+?)\]/g)[0].replace(/\[|]/g,''); // extract [dept] => get rid of "[", "]"
    } catch (e) {
        return;
    }

    // if tutor has (nickname), show "[DEPT] nickname"
    try {
        if (eventTitle.toString().includes('(') && eventTitle.toString().includes(')'))
            eventTitle = '[' + deptName + '] ' + eventTitle.match(/\((.+?)\)/g)[0].replace(/[()]/g,''); // extract (nickname) => get rid of "(", ")"
    } catch (e) {
        return;
    }

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
            //console.log(`Cannot match the dept name.`);
            return;
    }

    let calendarID = subSheet.getRange(calendarIdCell).getValue();
    let calendar = CalendarApp.getCalendarById(calendarID);
    let eventId = subSheet.getRange(row, column).getValue();
    if (eventId.length > 0) {
        let toBeDeleted = calendar.getEventById(eventId);
        try {
            toBeDeleted.deleteEvent();
        } catch (error) {
            console.log("cannot find " + eventId + " in the calendar " + calendar.getName());
        }

        subSheet.getRange(row, column).setValue("");
        mainSheet.getRange(row, column).clearNote();
    }



    let startDate = mainSheet.getRange(row, DTAE_TIME_COL).getValue();
    let startTime_hour = null;
    let endTime_hour = null;
    if (column < MORNING_NOON_BOUNDARY_COL) {
        startTime_hour = 9;
        endTime_hour = 13;
    } else if (column < NOON_NIGHT_BOUNDARY_COL) {
        startTime_hour = 13;
        endTime_hour = 17;
    } else {
        startTime_hour = 17;
        endTime_hour = 19;
    }
    let startDateTime = new Date(new Date(startDate).getTime() + startTime_hour * 60 * 60 * 1000);
    let endDateTime = new Date(new Date(startDate).getTime() + endTime_hour * 60 * 60 * 1000);

    //eventId = calendar.createEvent('[' + deptName + ']' + eventTitle, startDateTime, endDateTime, {description: descriptionUrl}).getId();
    //eventId = calendar.createEvent('[' + deptName + '] ' + eventTitle, startDateTime, endDateTime).getId()
    eventId = calendar.createEvent(eventTitle, startDateTime, endDateTime).getId();
    console.log(eventId, deptName, eventTitle, startDateTime, endDateTime);

    subSheet.getRange(row, column).setValue(eventId);
    mainSheet.getRange(row, column).setNote("Synchronized to calendar");
}

