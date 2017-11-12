// Spreadsheet Data
var data, sheet, lastRow, lastCol;

// Title Columns
var dateHeader, startTime, endTime, duration, title, 
    description, address, eventIDCol, calendarCol, eventColor;

// Event Data
var startDate;
var endDate;

/* Updates our calendar according to the given spreadsheet */
function main() {
  data = getSpreadsheetData();
  findTitles();
  for (var row=1;row<lastRow;row++) {
    
    try {                           // Try to parse event data
      getDuration(row);
      var calendar = getCalendar(row);
      
      if ((data[row][eventIDCol].length == 0) || (eventExists(row, calendar) == false)) {   // Checks if the event exists
        createEvent(row, calendar);
      }
      
      
      // If it reaches here, then everything went well
      setCell(title, row, null, "#c6ffc3", ""); 
    } catch (errorMsg) {
      setCell(title, row, null, "#ffb7ab", "Error - "+errorMsg);
    }
    if (row%5 == 0)  // Google limits how fast we can edit events
      Utilities.sleep(4000);
  }
}

/* 
* If the event exists, update the event then return true
* If the event does not exist, return false
*/
function eventExists(row, calendar) {
  var checkExistingEvents = calendar.getEventById(data[row][eventIDCol]);
  if (checkExistingEvents != null) {
    updateEvent(row, checkExistingEvents, calendar); 
    return true;
  } else
    return false;
}

/* Sets the column headers */
function findTitles() {
  for (var col=0;col<lastCol;col++) {
    switch (data[0][col].toUpperCase()) {
      case "DATE" :
        dateHeader = col;
        break;
      case "START TIME" :
        startTime = col;
        break;
      case "END TIME" :
        endTime = col;
        break;
      case "DURATION" :
        duration = col;
        break;
      case "EVENT TITLE" :
        title = col;
        break;
      case "DESCRIPTION" :
        description = col;
        break;
      case "LOCATION" :
        address=col;
        break;
      case "CALENDAR" :
        calendarCol=col;
        break;
      case "EVENT ID" :
        eventIDCol=col;
        break;
      case "EVENT COLOR" :
        eventColor=col;
        break;
    }
  }
}

/**
* Gets the calendar we want to work with 
* If calendar is left blank, it defaults to the user's primary calendar
* Otherwise, it looks for a calendar specified by the user
**/
function getCalendar(row, calendar) {
  if (data[row][calendarCol].length == 0) 
    return CalendarApp.getDefaultCalendar();
  
  var calendars = CalendarApp.getCalendarsByName(data[row][calendarCol]);
  if (calendars.length == 1)
    return calendars[0];
  else if (calendars.length > 1) 
    throw ("More than one calendar with that name");
  else
    throw ("No calendar found by name: "+data[row][calendarCol]);
}

/**
* Updates existing event 
*/
function updateEvent(row, existingEvent, calendar) {
  existingEvent.setTime(startDate, endDate);
  existingEvent.setTitle(data[row][title]);
  existingEvent.setLocation(data[row][address]);
  existingEvent.setDescription(data[row][description]);
}

/**
* Creates the event on Google Calendar
* Sets the cell to the returned event ID so it can be found later
**/
function createEvent(row, calendar) {
  var eventID = calendar.createEvent(data[row][title], startDate, endDate, 
                                     {location: data[row][address], 
                                      description: data[row][description],
                                      colorId: data[row][eventColor]});
  setCell(eventIDCol, row, eventID.getId());  
}

/**
* Gets the startDate and endDate
**/
function getDuration(row) {
  startDate = new Date(Date.parse(data[row][dateHeader]
                                  +" "+data[row][startTime]));
  checkDate(startDate, "Date and/or Start Time");
  
  if (data[row][duration].length > 0) {
    endDate = new Date(startDate);
    addHours(endDate, Number(data[row][duration]));
  } else if (data[row][endTime].length > 0) {
    endDate = new Date(data[row][dateHeader]+" "+data[row][endTime]);  
  } else
    throw ("user does not have either a duration or an endTime");
  
  checkDate(endDate, "End Time or Duration");
  if (endDate <= startDate) 
    throw ("something wrong with end time or start time.");
}

/** Checks if our created date is valid **/
function checkDate(date, string) {
  if (isNaN(date.getTime())) 
    throw ("cell not valid: "+string); 
}

/** Modifies given date by adding # of given hours to it **/
function addHours(date, hours) {   
  date.setHours(date.getHours() + hours);  
}

/* Opens up spreadsheet and gets data */
function getSpreadsheetData() {      
  sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  lastRow = sheet.getLastRow();
  lastCol = sheet.getLastColumn();
  return sheet.getDataRange().getValues();
}

/**
* Sets the desired cell's new background color and note
* @column - numeric value, function then converts it to a character
* @row - numeric value
* @text - cell's new text (optional)
* @color - color for cell (optional)
* @note - note for celll  (optional)
**/
function setCell(column, row, text, color, note) {  
  var cell = sheet.getRange(String.fromCharCode(97 + column)+""+(row+1));
  if (text != null)
    cell.setValue(text);
  if (note != null)
    cell.setNote(note);
  if (color != null)
    cell.setBackground(color);
}

/* Adds a menu when the spreadsheet is opened */
function onOpen() {  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];  
  menuEntries.push({name: "Start", functionName: "main"}); 
  sheet.addMenu("* Sync Calendar *", menuEntries);  
}
