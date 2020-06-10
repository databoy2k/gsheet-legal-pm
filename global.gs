// Global Variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var shFiles = ss.getSheetByName ("Files");

// Checking for Weekends in Dates
function onEdit(e) {
  var cell = e.range.getCell(1, 1);
  var val = cell.getValue();
  if ((val instanceof Date) && (val.getDay() == 0 || val.getDay() == 6)) {
    var fri = new Date();
    var mon = new Date();
    if (val.getDay() == 0) {
      fri.setDate(val.getDate()-2);
      mon.setDate(val.getDate()+1);
      }
    else {
      fri.setDate(val.getDate()-1);
      mon.setDate(val.getDate()+2);
      }
    Browser.msgBox('Warning: Date Occurs on Weekend\\nPrior Friday Date: '+fri.getDate()+'\\nNext Monday Date: '+mon.getDate());
//    cell.activate();
  }
}

// Create Menu
  SpreadsheetApp.getUi()
  .createMenu('Practice Management')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Add or Delete')
    .addItem('New File', 'duplicateSheet')
    .addItem('Delete Sheet', 'deleteSheet')
    .addItem('Delete Calendar Entries', 'deleteEvents'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Navigation')
    .addItem('Go to Selected Sheet', 'gotoSheet')
    .addItem('Go to File List', 'gotoFile')
    .addItem('Sort File List', 'sortFile')
    .addItem('Sort Sheet Tabs', 'sortSheets'))
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Functions')
    .addItem('Create Appointment from Task', 'createAppointment')
    .addItem('Add Email to Current Cell', 'addEmail')
    .addItem('Add Hyperlink for Folder', 'SearchFolder')
    .addItem('Update Diarization Calendar', 'exportDiarization')
    .addItem('Confirm Hyperlinks', 'ConfirmHyperlinks')
    .addItem('Clockify Start Time', 'StartClockifyTimer'))
  .addToUi();
  
function sortSheets () {
  var sheetNameArray = [];
  var sheets = ss.getSheets(); 
  for (var i = 0; i < sheets.length; i++) {
    sheetNameArray.push(sheets[i].getName());
  }
  sheetNameArray.sort();
  console.log(sheetNameArray);
  for( var j = 0; j < sheets.length; j++ ) {
    ss.setActiveSheet(ss.getSheetByName(sheetNameArray[j]));
    ss.moveActiveSheet(j + 1);
  }
  ["Files", "Court Information", "Timelines", "Template", "Sheet Operators"].forEach(function (name, index) { 
    ss.setActiveSheet(ss.getSheetByName(name)); 
    ss.moveActiveSheet(1 + index); 
    ss.getActiveSheet().hideSheet();
    });
}
