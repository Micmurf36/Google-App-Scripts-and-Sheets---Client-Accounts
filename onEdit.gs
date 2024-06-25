function onEdit(e) {
  var sheet = e.range.getSheet();
  var range = sheet.getRange('B1');

  // Only run the script if the edit was made in cell B1
   if (sheet.getName() == "Data") {
    var columnToSort = 2; // Column B
    var range = sheet.getDataRange();
    var firstRow = range.getRow(); 
    if (e.range.getRow() != firstRow) {
      range.sort([
        {column: columnToSort, ascending: true},
        {column: columnToSort, ascending: false, nullsLast: true}
      ]);
    }
  } else if (sheet.getName() == "Account") {

    if (e.range.getColumn() != range.getColumn() || e.range.getRow() != range.getRow()) {
    return;
    }
    // Clear previous data
    sheet.getRange("C6:C").clearContent();
    sheet.getRange("E5:E").clearContent();

    // Delay for 1 second
    Utilities.sleep(500);

    var values = sheet.getRange('E5:E').getValues();
    var lastRow = values.filter(String).length + 4; // count non-empty rows, add 4 to account for header rows
    var emptyRow = 0;

    for (var i = 4; i < lastRow - 2; i++) { // start from row 5, end before the last 2 header rows
      if (values[i][0] == "") {
        emptyRow = i + 1;
        break;
      }
    }
  }



  if (emptyRow > 0) {
    var dataRange = sheet.getRange(5, 5, lastRow - 4);
    var dataValues = dataRange.getValues();
    var sum = 0;
    for (var i = 0; i < dataValues.length; i++) {
      sum += dataValues[i][0];
    }
    var sumCell = sheet.getRange(lastRow + 2, 5);
    sumCell.setValue(sum);
    var balanceCell = sumCell.offset(0, -2);
    balanceCell.setValue("Balance").setFontWeight("bold");
    sheet.getRange(lastRow + 2, 3).setValue("Balance").setFontWeight("bold");
    
    // Set borders
    sheet.getRange(lastRow, 3, 1, 3).setBorder(false, false, false, false, false, false);
    sheet.getRange(lastRow + 1, 3, 1, 3).setBorder(true, false, false, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(5, 3, lastRow - 5, 3).setBorder(false, false, false, false, false, false);
  }
  else {
    sheet.getRange(lastRow + 2, 5).clearContent();
    sheet.getRange(lastRow + 2, 3).clearContent();
    sheet.getRange(lastRow, 3, 1, 3).setBorder(false, false, false, false, false, false);
    sheet.getRange(lastRow + 1, 3, 1, 3).setBorder(false, false, false, false, false, false);
    sheet.getRange(5, 3, lastRow - 5, 3).setBorder(false, false, false, false, false, false);
  }
}
function onFormSubmit(e) {
  var sheet = e.source.getSheetByName("Data");
  var columnToSort = 2; // Column B
  var range = sheet.getDataRange();
  range.sort([
    {column: columnToSort, ascending: true},
    {column: columnToSort, ascending: false, nullsLast: true}
  ]);
}
