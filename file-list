// Function to create sheets for new files
function duplicateSheet() {
  var FileNumber = Browser.inputBox("New File Number");
  var newRow = shFiles.getLastRow() + 1;
  var cell = shFiles.getRange('A1').offset(newRow,0);
  var row = shFiles.getRange('B1:I1').offset(newRow,0);
  var formulas = shFiles.getRange("B3:I3").getFormulasR1C1();
  var shTemplate = ss.getSheetByName('Template');
  var shNew = shTemplate.copyTo(ss).setName(FileNumber);
  var headers = {"X-Api-Key" : "[X-Api-Key]", "content-type" : "application/json"};
  var payload = JSON.stringify({'name' : FileNumber, 'clientId' : "[clientid]", 'color' : "#0053ed", 'billable' : "true", 'isPublic' : "true", 'hourlyRate' : {'amount' : '30000', 'currency': "CAD"}});
  var clockifyoptions = {
  'muteHttpExceptions' : true,
  'method' : 'post',
  'headers' : headers,
  'payload' : payload
  };
  var protections = shTemplate.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var p = protections[i];
    var rangeNotation = p.getRange().getA1Notation();
    var p2 = shNew.getRange(rangeNotation).protect();
    p2.setDescription(p.getDescription());
    p2.setWarningOnly(p.isWarningOnly());
  }
  shFiles.insertRowAfter(shFiles.getLastRow());
  cell.setValue(FileNumber);
  row.setFormulasR1C1(formulas);
  ss.setActiveSheet(ss.getSheetByName(FileNumber));
  var shid = ss.getSheetByName(FileNumber).getSheetId();
  MailApp.sendEmail("[emailaddress]", FileNumber, "A new file has been opened. Remember to create a label. The spreadsheet link is: https://docs.google.com/spreadsheets/d/[spreadsheetid]/edit#gid="+shid);
  ss.getRange('K6').setValue(FileNumber);
  UrlFetchApp.fetch('https://api.clockify.me/api/v1/workspaces/[workspaceid]/projects/', clockifyoptions);
  SearchFolder();
}

function gotoSheet() {
  // Function to go to selected sheet and sort
   var row = ss.getActiveSheet().getActiveRange().getRow();
   var FileNumber = ss.getActiveSheet().getRange(row,1).getValue();
   var lastRow = ss.getSheetByName(FileNumber).getLastRow();
   var range = ss.getSheetByName(FileNumber).getRange(13, 1, lastRow - 1, 4);
   range.sort(3);
   ss.setActiveSheet(ss.getSheetByName(FileNumber));
   SearchFolder();
}

function sortFile() {
// (Manual) Sort Function
  var sheet = ss.getSheetByName("Files");
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(3, 1, lastRow - 1, lastCol);
  range.sort([{column: 4, ascending: true}, {column: 6, ascending: true}]);
}

function deleteSheet(){
// Function to delete a sheet when closing a file
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var row = ss.getActiveSheet().getActiveRange().getRow();
  var FileNumber = ss.getActiveSheet().getRange(row,1).getValue();
  var sheet = ss.getSheetByName(FileNumber);
  var destination = SpreadsheetApp.openById('[exporttargetid]');
  if (Browser.msgBox('Do you want to Archive File '+FileNumber+'?',Browser.Buttons.YES_NO) == 'yes') {
    sheet.copyTo(destination);
    ss.deleteRow(row);
    ss.setActiveSheet(ss.getSheetByName(FileNumber));
    deleteEvents();
    ss.deleteSheet(ss.getSheetByName(FileNumber));
    ss.setActiveSheet(ss.getSheetByName("Files"));
    MailApp.sendEmail("[emailaddress]", FileNumber, "This file has been closed. Delete Project from Clockify (manual) and export your e-mails for closing. https://takeout.google.com/settings/takeout");
  }
}
