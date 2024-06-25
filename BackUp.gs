function archiveCopy() {
var file= DriveApp.getFileById("");
var destination= DriveApp.getFolderById("");

var timeZone=Session.getScriptTimeZone();
var formattedDate= Utilities.formatDate(new Date(),timeZone,"yyyy-MM-dd' 'HH:mm:ss");
var name= SpreadsheetApp.getActiveSpreadsheet().getName()+"Copy"+formattedDate;


  file.makeCopy(name,destination);
  
}
function fillEmptyRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("T1:T2985");
  var values = range.getValues();
  var maxId = 0;
  var nextId = 1;
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == "") {
      while (maxId >= nextId) {
        nextId++;
      }
      values[i][0] = nextId;
      maxId = nextId;
    } else {
      maxId = Math.max(maxId, values[i][0]);
    }
  }
  range.setValues(values);
}

