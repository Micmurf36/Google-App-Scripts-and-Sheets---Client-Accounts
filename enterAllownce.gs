function enterAllowance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var enterSheet = ss.getSheetByName("Enter Allowance");
  var dataSheet = ss.getSheetByName("Data");

  // Find last row with data in column D on Enter Allowance
  var lastRow = enterSheet.getLastRow();
  while (lastRow >= 3 && enterSheet.getRange("D" + lastRow).isBlank()) {
    lastRow--;
  }

  // Get date from H1 on Enter Allowance
  var date = enterSheet.getRange("G1").getValue();

  // Get the maximum transaction ID in column T and add 1
  var maxTransactionId = Math.max.apply(null, dataSheet.getRange("T:T").getValues().flat().filter(Boolean));
  var transactionId = isNaN(maxTransactionId) ? 1 : maxTransactionId + 1;

  // Show confirmation dialog
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "Are you sure you want to submit the allowance transactions?",
    ui.ButtonSet.YES_NO);

  // Proceed only if user clicked "Yes"
  if (response == ui.Button.YES) {
    // Loop through each row and copy data to Data sheet
    for (var i = 3; i <= lastRow; i++) {
      var dValue = enterSheet.getRange("D" + i).getValue();
      var eValue = enterSheet.getRange("E" + i).getValue();
      if (eValue > 0) {
        var newRow = [          dValue,          date,          "Allowance - Blue Petty Cash",          -Math.abs(eValue),          "",          "",          "",          "",          "",          "",          "",          "",          "",          "",          "",          "",          "",          "",          "",          transactionId++        ];
        dataSheet.appendRow(newRow);
      }
    }
  } else {
    // User clicked "No" or closed the dialog, so do nothing
    return;
  }
}
function enterGrocery() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var enterSheet = ss.getSheetByName("Enter Allowance");
  var dataSheet = ss.getSheetByName("Data");

  // Find last row with data in column D on Enter Allowance
  var lastRow = enterSheet.getLastRow();
  while (lastRow >= 3 && enterSheet.getRange("D" + lastRow).isBlank()) {
    lastRow--;
  }

  // Get date from G1 on Enter Allowance
  var date = enterSheet.getRange("G1").getValue();

  // Get the maximum transaction ID in column T and add 1
  var maxTransactionId = Math.max.apply(null, dataSheet.getRange("T:T").getValues().flat().filter(Boolean));
  var transactionId = isNaN(maxTransactionId) ? 1 : maxTransactionId + 1;

  // Show confirmation dialog
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Are you sure you want to submit the grocery transactions?", ui.ButtonSet.YES_NO);

  // Proceed only if user clicked "Yes"
  if (response == ui.Button.YES) {
    // Loop through each row and copy data to Data sheet if an "X" is found in column B
    for (var i = 3; i <= lastRow; i++) {
      var bValue = enterSheet.getRange("B" + i).getValue();
      if (bValue == "X") {
        var dValue = enterSheet.getRange("D" + i).getValue();
        var newRow = [dValue, date, "Fry's - Gift Card", -85, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", transactionId++];
        dataSheet.appendRow(newRow);
      }
    }
  } else {
    // User clicked "No" or closed the dialog, so do nothing
    return;
  }
}


