// Note that at various times in this document, [comment] denotes places where private, secure information needs to be entered.
// Remove the entirety of those entries (including the brackets) and place in the information (e.g. [calendarid] becomes 123456asdf)

function StartClockifyTimer() {
// Starts a timer based on the current sheet
// Step 1: Find PID
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var FileNo = sheet.getSheetName();
  var url = 'https://api.clockify.me/api/v1/workspaces/[workspaceid]/projects?name='+FileNo
  const header = {
    "headers": {
      "X-Api-Key" : "[X-Api-Key]",
      "content-type" : "application/json",
      "Accept" : "*/*"
      }
  };
  var response = UrlFetchApp.fetch(url, header)
  var json = response.getContentText();
  var data = JSON.parse(json);
  var PID = data[0]["id"];
  Logger.log(PID);

//Step 2: Use PID to start timer using current task and current time
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(11, 1, lastRow - 1, 4);
  range.sort(3);
  var task = sheet.getRange('A8').getValue();
  var date = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  var headers = {"X-Api-Key" : "[X-Api-Key]", "content-type" : "application/json"};
  var payload = JSON.stringify({'start' : date, 'projectId' : PID, 'description' : task, 'billable' : 'true'});
  var clockifyoptions = {
  'muteHttpExceptions' : true,
  'method' : 'post',
  'headers' : headers,
  'payload' : payload
  };
  UrlFetchApp.fetch('https://api.clockify.me/api/v1/workspaces/[workspaceid]/time-entries/', clockifyoptions);
}

function SearchFolder() {
// Searches Google Drive for the folder for the Active Sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var searchTerm = ss.getSheetName();
  var Folders = DriveApp.searchFolders("title contains '"+searchTerm.replace("'","\'")+"' and trashed = false and hidden = false");
  var Folder = Folders.next();
  sheet.getRange('K6').setFormula("=HYPERLINK(\""+Folder.getUrl()+"\",\""+searchTerm+"\")");
}

//Functions for Adding Timelines
function AddTimeline() {
  var shFile = ss.getActiveSheet();
  var rgeTimelines = shFile.getRange("F8:F").getValues();
  var shTimelines = ss.getSheetByName("Timelines");
  var optTimelines = ss.getSheetByName("Sheet Operators").getRange("A3:B").getValues();
  while(rgeTimelines[rgeTimelines.length - 1] === 0){
    rgeTimelines.pop();
    }
  var newTimeline = rgeTimelines.filter(String).length;
  var selection = shFile.getRange("F1").getValue();
  if (selection == "") {
    Browser.msgBox('Please Make a Selection')
    }
  else {
    for (i=0; i<optTimelines.length; i++) {
        if (optTimelines [i][0] == selection) {
          var rgeTimeline = optTimelines[i][1];
        }
    }
    shTimelines.getRange(rgeTimeline).copyTo(shFile.getRange('F10:M10').offset(newTimeline - 1,0));
    shFile.getRange("F1").clearContent();
  }
}

function gotoFile() {
  // Function to sort then return to file list
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cursheet = ss.getActiveSheet();
  var lastRow = cursheet.getLastRow();
  var range = cursheet.getRange(11, 1, lastRow - 1, 4);
  range.sort(3);
  exportDiarization();
  var nexttaskdate = new Date(cursheet.getRange('C8').getValue());
  if(nexttaskdate.toString() == "Invalid Date") {
    Browser.msgBox('Warning: Bad Date Detected for Next Task. Please Fix.');
    return;
    }
  ss.setActiveSheet(ss.getSheetByName("Files"));
  sortFile();
}

// Export Diarization
function exportDiarization() {
  var sheet = ss.getActiveSheet();
  var shid = sheet.getSheetId();
  var range = sheet.getRange(3, 2, 3, 4); // Range: Key Deadlines Description and Dates (B3:E6) 
  var fileno = sheet.getSheetName();
  var data = range.getValues();
  var calID = "[calendaraddress]@group.calendar.google.com";
  var cal = CalendarApp.getCalendarById(calID);
  var formulas = sheet.getRange(3, 2, 1, 3).getFormulas(); // Snagging one line of formulae which get broken in this process; tofix?
  //var Folders = DriveApp.searchFolders("title contains '"+fileno.replace("'","\'")+"' and trashed = false and hidden = false");
  //var Folder = Folders.next();
  for (i=0; i<data.length; i++) {
    var row = data[i];
    var description = row[0];     // First column (B - "Description")
    var date = new Date(row[2]);  // Third column (D - "Date")
      // Check if date is invalid; if so, trim that row if no date or alert if broken.
      if(date.toString() == "Invalid Date"){
        if(row[2].trim() == ""){
          if (row[3] != ""){
          var delevent = cal.getEventById(row[3]);
          delevent.deleteEvent();
          row[3] = "";
          }
          continue;
          }
      else {                 
        Browser.msgBox('Warning: Bad Date Detected; Entry Not Updated');
        continue; 
        }
    }
    var id = row[3];              // Fourth column (E - Dates, written in white)
    // Check if event already exists, update it if it does
    try {
      var event = cal.getEventById(id);
    }
    // Catches an exception if no event exists
    catch (e) {
    }
    if (!event) {
      var newEvent = cal.createAllDayEvent(fileno+' - '+description, date, {description: "https://docs.google.com/spreadsheets/d/[documentidofspreadsheet]/edit#gid="+shid}).getId();
      row[3] = newEvent;  // Update the data array with event ID
    }
    else {
        event.setTitle(fileno+' - '+description);
        event.setAllDayDate(date);
        event.setLocation("https://docs.google.com/spreadsheets/d/[documentidofspreadsheet]/edit#gid="+shid);
        }
    debugger;
  }
  // Record all event IDs to spreadsheet and restore formulas in first row
  range.setValues(data);
  sheet.getRange(3, 2, 1, 3).setFormulas(formulas);
}

// Calendar Deleter (for closing files)
function deleteEvents() {
  var sheet = ss.getActiveSheet();
  var range = sheet.getRange(3, 2, 3, 4);
  var fileno = sheet.getSheetName();
  var data = range.getValues();
  var calID = "[calendarid]@group.calendar.google.com";
  var cal = CalendarApp.getCalendarById(calID);
  var formulas = sheet.getRange(3, 2, 1, 3).getFormulas();
  for (i=0; i<data.length; i++) {
    var row = data[i];
    var description = row[0];           // First column
    var date = new Date(row[2]);  // Third column
    var id = row[3];              // Fourth column == eventId
    // Check if event already exists, update it if it does
    try {
      var event = cal.getEventById(id);
    }
    catch (e) {
      // do nothing - we just want to avoid the exception when event doesn't exist
    }
    if (!event) {
      // var newEvent = "None";
      row[3] = "";
    }
    else {
      event.deleteEvent();
      row[3] = "";
    }
    debugger;
  }
  // Record all event IDs to spreadsheet and restore formulas in first row
  range.setValues(data);
  sheet.getRange(3, 2, 1, 3).setFormulas(formulas);
}

// Calendar Event Creator
function createAppointment() {
  var sheet = ss.getActiveSheet();
  var row = ss.getActiveSheet().getActiveRange().getRow();
  var shid = sheet.getSheetId();
  var fileno = sheet.getSheetName();
  var calID = "tmeyers@macleanwest.com";
  var cal = CalendarApp.getCalendarById(calID);
  var meetingType = ss.getActiveSheet().getRange(row,1).getValue();
  var date = ss.getActiveSheet().getRange(row,3).getValue();
  var newEvent = cal.createEvent(meetingType+" ("+fileno+")", new Date(date.getTime()+8*3600000), new Date(date.getTime()+9*3600000),{description: "https://docs.google.com/spreadsheets/d/[documentidofspreadsheet]/edit#gid="+shid}).getId();
  var splitEventId = newEvent.split('@');
  var eventUrl = "https://www.google.com/calendar/event?eid="+Utilities.base64Encode(splitEventId[0] + " " + calID).toString().replace('=','');
  sheet.getRange(row,1).setFormula("=HYPERLINK(\""+eventUrl+"\",\""+meetingType+"\")");
  }
  
function addEmail() {
   var sheet = ss.getActiveSheet();
   var content = ss.getActiveSheet().getActiveRange().getValue();
   var email = Browser.inputBox("E-mail Address to Insert");
   sheet.getActiveRange().setFormula("=HYPERLINK(\"mailto:"+email+"\",\""+content+"\")");
 }
