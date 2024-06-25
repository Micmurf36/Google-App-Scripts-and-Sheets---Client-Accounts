function searchRecordInv() {
  var myGooglSheet = SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm = myGooglSheet.getSheetByName("Create Invoice"); //declare a variable and set with the User Form worksheet
  var datasheet = myGooglSheet.getSheetByName("Data"); //declare a variable and set with the Database worksheet
  
  var str = shUserForm.getRange("B2").getValue();
  var values = datasheet.getDataRange().getValues(); //getting the entire values from the used range and assigning it to values variable
  
  for (var i = 0; i < values.length; i++) {
    var rowValue = values[i];
    if (rowValue[0] == str && rowValue[4] != "") { // add check for non-empty value in column E
      // set values in cells B4 to B30
      shUserForm.getRange("B4").setValue(rowValue[0]);
      shUserForm.getRange("B6").setValue(rowValue[4]);
      shUserForm.getRange("B8").setValue(rowValue[5]);
      shUserForm.getRange("B10").setValue(rowValue[6]);
      shUserForm.getRange("B12").setValue(rowValue[7]);
      shUserForm.getRange("B14").setValue(rowValue[8]);
      shUserForm.getRange("B16").setValue(rowValue[9]);
      shUserForm.getRange("B18").setValue(rowValue[10]);
      shUserForm.getRange("B20").setValue(rowValue[11]);
      shUserForm.getRange("B22").setValue(rowValue[12]);
      shUserForm.getRange("B24").setValue(rowValue[13]);
      shUserForm.getRange("B26").setValue(rowValue[14]);
      shUserForm.getRange("B28").setValue(rowValue[15]);
      shUserForm.getRange("B43").setValue(rowValue[16]);
      return;
    }
  }
 
  if(valuesFound = false){
    //to create the instance of the user-interface environment to use the messagebox features
    var ui = SpreadsheetApp.getUi();
    ui.alert("No record found!");
  }
}

function searchReceiptTrans() {
  var myGooglSheet = SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm = myGooglSheet.getSheetByName("Create Invoice"); //declare a variable and set with the User Form worksheet
  var datasheet = myGooglSheet.getSheetByName("Data"); //declare a variable and set with the Database worksheet
  
  var selectedCell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getCurrentCell();
  var transactionId;
  if (shUserForm.getRange("H4").getValue() !== "") {
    transactionId = shUserForm.getRange("H4").getValue();
    Logger.log("Using transaction ID from H4: " + transactionId);
  } else {
    transactionId = selectedCell.getValue();
    Logger.log("Using transaction ID from selected cell: " + transactionId);
  }
  
  var values = datasheet.getDataRange().getValues(); //getting the entire values from the used range and assigning it to values variable
  
  for (var i = 0; i < values.length; i++) {
    var rowValue = values[i];
    if (rowValue[19] == transactionId) { //compare transaction ID in column T with selected cell value
      // set values in cells B4 to B30 the transaction ID is in column T (i.e., column 19)
      shUserForm.getRange("H8").setValue(rowValue[1]);
      shUserForm.getRange("I8").setValue(rowValue[2]);
      shUserForm.getRange("J8").setValue(rowValue[3]);      
      return; //come out from the search function
    }
  }
}

function moveValues1() {
  var myGooglSheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var shUserForm = myGooglSheet.getSheetByName("Create Invoice"); 
  var h8 = shUserForm.getRange("H8").getValue();
  var i8 = shUserForm.getRange("I8").getValue();
  var j8 = shUserForm.getRange("J8").getValue();
  shUserForm.getRange("H11").setValue(h8);
  shUserForm.getRange("I11").setValue(i8);
  shUserForm.getRange("J11").setValue(j8);
}

function moveValues2() {
  var myGooglSheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var shUserForm = myGooglSheet.getSheetByName("Create Invoice"); 
  var h8 = shUserForm.getRange("H8").getValue();
  var i8 = shUserForm.getRange("I8").getValue();
  var j8 = shUserForm.getRange("J8").getValue();
  shUserForm.getRange("H12").setValue(h8);
  shUserForm.getRange("I12").setValue(i8);
  shUserForm.getRange("J12").setValue(j8);
}

function moveValues3() {
  var myGooglSheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var shUserForm = myGooglSheet.getSheetByName("Create Invoice"); 
  var h8 = shUserForm.getRange("H8").getValue();
  var i8 = shUserForm.getRange("I8").getValue();
  var j8 = shUserForm.getRange("J8").getValue();
  shUserForm.getRange("H13").setValue(h8);
  shUserForm.getRange("I13").setValue(i8);
  shUserForm.getRange("J13").setValue(j8);
}

function clearBoxReceipt() {
  var myGooglSheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var shUserForm = myGooglSheet.getSheetByName("Create Invoice"); 
  shUserForm.getRange("H8:J13").clearContent();
}

function createInvoiceEmail() { 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Create Invoice');
  var email = sheet.getRange('B12').getValue(); 
  var name = sheet.getRange('B10').getValue().split(" ")[0];
  var date = new Date(sheet.getRange('E11').getValue());
  var clientName = sheet.getRange('B4').getValue().split(" ")[0];
  var dayOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][date.getDay()];
  var month = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][date.getMonth()];
  var dateString = dayOfWeek + ', ' + month + ' ' + date.getDate() + ', ' + date.getFullYear();
  var subject = clientName + " " + sheet.getRange('E7').getValue() + " Invoice";
  var message = "Hello " + name + ",\n\n" + clientName + " is beginning his next month here at the treatment facility this coming " + dateString + ". I have attached the invoice.\n\n" + "We offer several payment options including debit/credit card, check, wire transfer, and other options. If you have previously paid by credit or debit card, we may have your card on file and all we need is your authorization to run the payment. Please let us know if you need any additional information for the other payment options.\n\n" + "Thank you for your continued support of " + clientName + ".\n\n" + "Please don't hesitate to contact us if you have any questions or concerns. We wish you a great day!\n\n" + "Best regards,\n" + "Admin"; 
  GmailApp.createDraft(email, subject, message); 
}

function getInsTotal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var createInvoiceSheet = ss.getSheetByName("Create Invoice");
  var dataSheet = ss.getSheetByName("Data");
  
  var nameToSearch = createInvoiceSheet.getRange("B2").getValue().toString();
  var total = 0;
  
  var data = dataSheet.getRange("A:D").getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var nameInData = row[0].toString();
    var amount = row[3];
    
    if (nameInData.indexOf(nameToSearch + " Ins") !== -1 || nameInData.indexOf(nameToSearch + " INS") !== -1) {
      total += amount;
    }
  }
  
  createInvoiceSheet.getRange("D14").setValue(nameToSearch + " Ins");
  createInvoiceSheet.getRange("E14").setValue(total);
}

function createPDF(ssId, sheet, pdfName) {
  const fr = 0, fc = 0, lc = 9, lr = 27;
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');

  // Gets the folder in Drive where the PDFs are stored.
  const folder = getFolderByName_("OUTPUT_FOLDER_NAME");

  const pdfFile = folder.createFile(blob);
  return pdfFile;
}

function findPositiveTransactions() {
  var myGooglSheet = SpreadsheetApp.getActiveSpreadsheet();
  var shUserForm = myGooglSheet.getSheetByName("Create Invoice");

  var lastRow = shUserForm.getLastRow();
  var dates = shUserForm.getRange("L2:L" + lastRow).getValues();
  var descriptions = shUserForm.getRange("N2:N" + lastRow).getValues();
  var amounts = shUserForm.getRange("O2:O" + lastRow).getValues();

  var positiveTransactions = [];

  for (var i = 0; i < amounts.length; i++) {
    var transactionAmount = amounts[i][0];
    var transactionDescription = descriptions[i][0];
    var transactionDate = dates[i][0];

    if (transactionAmount > 0 && !transactionDescription.includes('(')) {
      positiveTransactions.push([transactionDate, transactionDescription, transactionAmount]);
    }
  }

  if (positiveTransactions.length > 0) {
    positiveTransactions.sort(function(a, b) {
      return new Date(a[0]) - new Date(b[0]); // Sort transactions by date in ascending order
    });

    var startRow = 11; // Starting row for printing positive transactions

    for (var j = 0; j < positiveTransactions.length; j++) {
      var row = startRow + j;
      shUserForm.getRange(row, 8).setValue(positiveTransactions[j][0]); // Date in column H
      shUserForm.getRange(row, 9).setValue(positiveTransactions[j][1]); // Description in column I
      shUserForm.getRange(row, 10).setValue(positiveTransactions[j][2]); // Amount in column J
    }
  }
}
