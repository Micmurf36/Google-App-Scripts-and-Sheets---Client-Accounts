function generateMonthlyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName("Data");
  var monthlyReportSheet = ss.getSheetByName("Monthly Report");
  var dateInB1 = monthlyReportSheet.getRange("B1").getValue();
  var dataRange = dataSheet.getDataRange();
  var dataValues = dataRange.getValues();
  var tuitionOwed = [];
  
  for (var i = 1; i < dataValues.length; i++) {
    var clientName = dataValues[i][0];
    var intakeDateString = dataValues[i][9]; // assuming column J is the 10th column (index 9)
    var intakeDate = new Date(intakeDateString);

    if (intakeDate && !isNaN(intakeDate.getTime())) {
      // Check if there is a valid intake date
      var monthsSinceIntake = monthsBetween(intakeDate, dateInB1);
      var tuitionAmount = 0;
      
      if (monthsSinceIntake <= 6) {
        for (var j = 10; j <= 15 && j <= (monthsSinceIntake + 9); j++) {
          tuitionAmount += dataValues[i][j];
        }
      } else {
        tuitionAmount = dataValues[i][15];
      }
      
      if (tuitionAmount > 0) {
        tuitionOwed.push([clientName, tuitionAmount]);
      }
    }
  }
  
  var tuitionRange = monthlyReportSheet.getRange(4, 4, tuitionOwed.length, 2);
  tuitionRange.setValues(tuitionOwed);
}


function monthsBetween(date1, date2) {
  var startOfMonth = new Date(date1.getFullYear(), date1.getMonth(), 1);
  var endOfMonth = new Date(date1.getFullYear(), date1.getMonth() + 1, 0);
  var daysInFirstMonth = Math.min(endOfMonth.getDate() - date1.getDate() + 1, endOfMonth.getDate());
  var daysBetween = Math.floor((date2 - date1) / (1000 * 60 * 60 * 24));
  var months = (date2.getFullYear() - date1.getFullYear()) * 12;
  months -= date1.getMonth();
  months += date2.getMonth();
  if (daysBetween >= endOfMonth.getDate() - date1.getDate() + 1) {
    months += 1;
  } else if (daysBetween >= daysInFirstMonth) {
    months += 1 - (daysBetween - daysInFirstMonth) / (endOfMonth.getDate() - date1.getDate() + 1);
  }
  return months <= 0 ? 0 : months;
}



function generateWhatWasPaidReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName("Data");
  var monthlyReportSheet = ss.getSheetByName("Monthly Report");
  var dataRange = dataSheet.getDataRange();
  var dataValues = dataRange.getValues();
  var monthlyReportRange = monthlyReportSheet.getDataRange();
  var monthlyReportValues = monthlyReportRange.getValues();
  var whatWasPaid = [];
  
  for (var i = 3; i < monthlyReportValues.length; i++) {
    var clientName = monthlyReportValues[i][3];
    if (clientName) {
      var paidAmount = 0;
      for (var j = 1; j < dataValues.length; j++) {
        if (dataValues[j][0] == clientName && dataValues[j][3] > 0) {
          paidAmount += Math.round(dataValues[j][3] / 500) * 500;
        }
      }
      whatWasPaid.push([paidAmount]);
    }
  }
  
  var whatWasPaidRange = monthlyReportSheet.getRange(4, 6, whatWasPaid.length, 1);
  whatWasPaidRange.setValues(whatWasPaid);
}

function searchForCheckNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthlyReportSheet = ss.getSheetByName("Monthly Report");
  const dataSheet = ss.getSheetByName("Data");
  const checkSearchSheet = ss.getSheetByName("Check Search");
  const checkNumberToSearch = checkSearchSheet.getRange("B3").getValue();
  
  const data = dataSheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const description = row[2];
    const checkNumberIndex = description.indexOf(checkNumberToSearch);
    if (checkNumberIndex !== -1) {
      const valuesToPrint = [row[0], row[1], row[2], row[3]];
      checkSearchSheet.getRange("A7:D7").setValues([valuesToPrint]);
      return;
    }
  }
  

}


function searchForCheckNumberRange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const searchCheckSheet = ss.getSheetByName("Check Search");
  const startingCheckNumber = searchCheckSheet.getRange("B13").getValue();
  const endingCheckNumber = searchCheckSheet.getRange("B14").getValue();
  const numChecksToPrint = endingCheckNumber - startingCheckNumber + 1;
  const valuesToPrint = [];
  const dataSheet = ss.getSheetByName("Data");

  const data = dataSheet.getDataRange().getValues();
  for (let i = 0; i < data.length && valuesToPrint.length < numChecksToPrint; i++) {
    const row = data[i];
    const description = row[2];
    const checkNumber = getCheckNumberFromDescription(description);
    if (checkNumber >= startingCheckNumber && checkNumber <= endingCheckNumber) {
      valuesToPrint.push([row[0], row[1], row[2], row[3]]);
    }
  }
  
  const outputRange = searchCheckSheet.getRange("A18:D" + (17 + valuesToPrint.length));
  outputRange.clearContent();
  if (valuesToPrint.length > 0) {
    outputRange.setValues(valuesToPrint);
  } else {
    outputRange.setValue("No checks were found");
  }
}

function getCheckNumberFromDescription(description) {
  const matches = description.match(/([0-9]+)$/);
  if (matches && matches.length > 1) {
    return parseInt(matches[1]);
  }
  return -1;
}
function clearCheckSearchResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const searchCheckSheet = ss.getSheetByName("Check Search");
  
  // Clear the search results
  const outputRange = searchCheckSheet.getRange("A18:D" + searchCheckSheet.getLastRow());
  outputRange.clearContent();

  // Clear the single row values
  searchCheckSheet.getRange("A7:D7").clearContent();
}
function searchTransactions() {
  // Get the start and end dates from the Date Search sheet
  var startDate = new Date(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Date Search').getRange('B2').getValue());
  var endDate = new Date(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Date Search').getRange('B3').getValue());

  // Get the transactions data from the Data sheet
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var data = dataSheet.getDataRange().getValues();

  // Initialize an array to hold the results
  var results = [];

  // Loop through the data to find transactions within the date range
  for (var i = 1; i < data.length; i++) {
    var date = new Date(data[i][1]); // Assumes dates are in column B
    if (date >= startDate && date <= endDate) {
      var name = data[i][0]; // Assumes names are in column A
      var description = data[i][2]; // Assumes descriptions are in column C
      var amount = data[i][3]; // Assumes amounts are in column D
      results.push([name, date, description, amount]);
    }
  }

  // Write the results to the Date Search sheet
  var searchSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Date Search');
  searchSheet.getRange('A7:D').clearContent(); // Clear any existing content in the range
  if (results.length > 0) {
    searchSheet.getRange(7, 1, results.length, 4).setValues(results);
  }
}

function clearSearchResults() {
  var searchSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Date Search');
  searchSheet.getRange('A7:D').clearContent();
}
function searchTransactionsChecks() {
  // Get the start and end dates from the Date Search sheet
  var startDate = new Date(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Date Search').getRange('B2').getValue());
  var endDate = new Date(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Date Search').getRange('B3').getValue());

  // Get the transactions data from the Data sheet
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var data = dataSheet.getDataRange().getValues();

  // Initialize an array to hold the results
  var results = [];

  // Loop through the data to find transactions within the date range
  for (var i = 1; i < data.length; i++) {
    var date = new Date(data[i][1]); // Assumes dates are in column B
    var description = data[i][2]; // Assumes descriptions are in column C
    if (date >= startDate && date <= endDate && description.includes("ck#")) {
      var name = data[i][0]; // Assumes names are in column A
      var amount = data[i][3]; // Assumes amounts are in column D
      results.push([name, date, description, amount]);
    }
  }

  // Sort the results array by date in ascending order
  results.sort(function(a, b) {
    return a[1] - b[1];
  });

  // Write the results to the Date Search sheet
  var searchSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Date Search');
  searchSheet.getRange('A7:D').clearContent(); // Clear any existing content in the range
  if (results.length > 0) {
    searchSheet.getRange(7, 1, results.length, 4).setValues(results);
  }
}
