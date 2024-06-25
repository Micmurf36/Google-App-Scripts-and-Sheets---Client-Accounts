function validateTrans(){
 
  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm    = myGooglSheet.getSheetByName("Enter Transaction"); //delcare a variable and set with the User Form worksheet
 
  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
 
    //Assigning white as default background color
 
  shUserForm.getRange("B6").setBackground('#FFFFFF');
  shUserForm.getRange("B8").setBackground('#FFFFFF');
  shUserForm.getRange("B10").setBackground('#FFFFFF');
 
  
//Validating Employee ID
  if(shUserForm.getRange("B6").isBlank()==true){
    ui.alert("Please enter Date.");
    shUserForm.getRange("B6").activate();
    shUserForm.getRange("B6").setBackground('#FF0000');
    return false;
  }
  //Validating Gender
  else if(shUserForm.getRange("B8").isBlank()==true){
    ui.alert("Please enter description");
    shUserForm.getRange("B8").activate();
    shUserForm.getRange("B8").setBackground('#FF0000');
    return false;
  }
  //Validating Email ID
  else if(shUserForm.getRange("B10").isBlank()==true){
    ui.alert("Please enter the amount");
    shUserForm.getRange("B10").activate();
    shUserForm.getRange("B10").setBackground('#FF0000');
    return false;
  }
  
  return true;

}
  // Function to submit the data to Database sheet
function submitDataTrans() {
  var myGooglSheet = SpreadsheetApp.getActiveSpreadsheet();
  var shUserForm = myGooglSheet.getSheetByName("Enter Transaction");
  var datasheet = myGooglSheet.getSheetByName("Data");
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert("Submit", 'Do you want to submit the data?',ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) {return;} 
 
  if (validateTrans()==true) {
    var blankRow = datasheet.getLastRow() + 1;

    datasheet.getRange(blankRow, 1).setValue(shUserForm.getRange("B4").getValue());
    datasheet.getRange(blankRow, 2).setValue(shUserForm.getRange("B6").getValue());
    datasheet.getRange(blankRow, 3).setValue(shUserForm.getRange("B8").getValue());
    datasheet.getRange(blankRow, 4).setValue(shUserForm.getRange("B10").getValue());
    datasheet.getRange(blankRow, 18).setValue(new Date()).setNumberFormat('yyyy-mm-dd h:mm');
    datasheet.getRange(blankRow, 19).setValue(Session.getActiveUser().getEmail());

    // Generate the next transaction ID
    var lastTransactionId = datasheet.getRange("T1:T").getValues().filter(String).map(function(value) {
      return parseInt(value);
    }).sort(function(a, b) {
      return b - a;
    })[0];

    var nextTransactionId = lastTransactionId ? lastTransactionId + 1 : 1;
    datasheet.getRange(blankRow, 20).setValue(nextTransactionId);
    
    ui.alert('Transaction Saved ' + shUserForm.getRange("B4").getValue());    
    
    shUserForm.getRange("B8").clear();
    shUserForm.getRange("B10").clear();
  }
}


function searchRecordTrans() {
  var myGooglSheet = SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm = myGooglSheet.getSheetByName("Enter Transaction"); //delcare a variable and set with the User Form worksheet
  var datasheet = myGooglSheet.getSheetByName("Data"); ////delcare a variable and set with the Database worksheet
  
  var selectedCell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getCurrentCell();
  var transactionId;
  if (shUserForm.getRange("B12").getValue() !== "") {
    transactionId = shUserForm.getRange("B12").getValue();
    Logger.log("Using transaction ID from B12: " + transactionId);
  } else {
    transactionId = selectedCell.getValue();
    Logger.log("Using transaction ID from selected cell: " + transactionId);
  }
  
  var values = datasheet.getDataRange().getValues(); //getting the entire values from the used range and assigning it to values variable
  
  for (var i = 0; i < values.length; i++) {
    var rowValue = values[i];
    if (rowValue[19] == transactionId) { //compare transaction ID in column T with selected cell value
      // set values in cells B4 to B30 the transaction ID is in column T (i.e., column 19)
      shUserForm.getRange("B4").setValue(rowValue[0]);
      shUserForm.getRange("B6").setValue(rowValue[1]);
      shUserForm.getRange("B8").setValue(rowValue[2]);
      shUserForm.getRange("B10").setValue(rowValue[3]);
      shUserForm.getRange("B12").setValue(rowValue[19]);
      return; //come out from the search function
    }
  }
}




function clearFormTrans() 
{
  var myGoogleSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm    = myGoogleSheet.getSheetByName("Enter Transaction"); //declare a variable and set with the User Form worksheet
 
  //to create the instance of the user-interface environment to use the alert features
  var ui = SpreadsheetApp.getUi();
 
  // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
  // close the dialog by clicking the close button in its title bar.
  var response = ui.alert("Reset Confirmation", 'Do you want to reset this form?',ui.ButtonSet.YES_NO);
 
 // Checking the user response and proceed with clearing the form if user selects Yes
 if (response == ui.Button.YES) 
  {
     
  shUserForm.getRange("B4").clear(); //Search Field
  shUserForm.getRange("B6").clear();// Employeey ID
  shUserForm.getRange("B8").clear(); // Employee Name
  shUserForm.getRange("B10").clear(); // Gender
  shUserForm.getRange("B12").clear(); // Gender
 //Assigning white as default background color
 
 shUserForm.getRange("B4").setBackground('#FFFFFF');
 shUserForm.getRange("B6").setBackground('#FFFFFF');
 shUserForm.getRange("B8").setBackground('#FFFFFF');
 shUserForm.getRange("B10").setBackground('#FFFFFF');
 shUserForm.getRange("B12").setBackground('#FFFFFF');
 
  return true ;
  
  }
}

function modifyRecordTrans() {
  var myGooglSheet = SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm = myGooglSheet.getSheetByName("Enter Transaction"); //declare a variable and set with the Account Entry worksheet
  var datasheet = myGooglSheet.getSheetByName("Data"); //declare a variable and set with the Database worksheet

  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();

  // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
  // close the dialog by clicking the close button in its title bar.
  var response = ui.alert("Submit", 'Do you want to edit the data?', ui.ButtonSet.YES_NO);

  // Checking the user response and proceed with clearing the form if user selects Yes
  if (response == ui.Button.NO) {
    return; //exit from this function
  }

  var str = shUserForm.getRange("B12").getValue();
  var values = datasheet.getDataRange().getValues(); //getting the entire values from the used range and assigning it to values variable
  for (var i = 0; i < values.length; i++) {
    var rowValue = values[i];
    if (rowValue[19] == str) {
      var row = i + 1;

      datasheet.getRange(row, 1).setValue(shUserForm.getRange("B4").getValue()); 
      datasheet.getRange(row, 2).setValue(shUserForm.getRange("B6").getValue());
      datasheet.getRange(row, 3).setValue(shUserForm.getRange("B8").getValue()); 
      datasheet.getRange(row, 4).setValue(shUserForm.getRange("B10").getValue()); 
      datasheet.getRange(row, 20).setValue(shUserForm.getRange("B12").getValue()); 

      // date function to update the current date and time as submittted on
      datasheet.getRange(row, 21).setValue(new Date()).setNumberFormat('yyyy-mm-dd h:mm'); //Submitted On
        
      //get the email address of the person running the script and update as Submitted By
      datasheet.getRange(row, 22).setValue(Session.getActiveUser().getEmail()); //Submitted By
        
      ui.alert(' "Data updated for - Client ' + shUserForm.getRange("B4").getValue() +' "');

      //Clearing the data from the Data Entry Form
      shUserForm.getRange("B4").clear() ;     
      shUserForm.getRange("B6").clear() ;
      shUserForm.getRange("B8").clear() ;
      shUserForm.getRange("B10").clear() ;
      shUserForm.getRange("B12").clear() ;

      return; //come out from the search function
    }
  }
}
function clearEmptyG() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Enter Transaction");
  var dataSheet = SpreadsheetApp.getActive().getSheetByName("Data");
  var filterFormula = sheet.getRange("G4").getFormula(); // assuming the filter formula is in cell G4

  // Get the range of values displayed by the filter formula
  var range = dataSheet.getRange(filterFormula.substring(filterFormula.indexOf("(") + 1, filterFormula.lastIndexOf(")")));

  // Loop through each row in the range and clear the corresponding cell in column G if the corresponding cell in column F is empty
  for (var i = 1; i <= range.getNumRows(); i++) {
    var row = range.getCell(i, 1, 1, range.getNumColumns()).getValues()[0];
    var fValue = row[5];
    var gValue = row[6];
    if (fValue === "" && gValue !== "") {
      range.getCell(i, 7).clearContent();
    }
  }
}
function deleteRecordTrans() {
  var myGooglSheet = SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm = myGooglSheet.getSheetByName("Enter Transaction"); //declare a variable and set with the Account Entry worksheet
  var datasheet = myGooglSheet.getSheetByName("Data"); //declare a variable and set with the Database worksheet

  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();

  // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
  // close the dialog by clicking the close button in its title bar.
  var response = ui.alert("Submit", 'Do you want to delete the transaction?', ui.ButtonSet.YES_NO);

  // Checking the user response and proceed with deleting the row if user selects Yes
  if (response == ui.Button.NO) {
    return; //exit from this function
  }

  var str = shUserForm.getRange("B12").getValue();
  var values = datasheet.getDataRange().getValues(); //getting the entire values from the used range and assigning it to values variable
  for (var i = 0; i < values.length; i++) {
    var rowValue = values[i];
    if (rowValue[19] == str) {
      var row = i + 1;
      
      datasheet.deleteRow(row); // delete the row
      
      ui.alert(' "Data deleted for - Client ' + shUserForm.getRange("B4").getValue() +' "');

      //Clearing the data from the Data Entry Form
      shUserForm.getRange("B4").clear() ;     
      shUserForm.getRange("B6").clear() ;
      shUserForm.getRange("B8").clear() ;
      shUserForm.getRange("B10").clear() ;
      shUserForm.getRange("B12").clear() ;

      return; //come out from the search function
    }
  }
}

