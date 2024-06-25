function validateEntry(){
 
  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm    = myGooglSheet.getSheetByName("Account Entry"); //delcare a variable and set with the User Form worksheet
 
  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
 
    //Assigning white as default background color
 
  shUserForm.getRange("B4").setBackground('#FFFFFF');
  shUserForm.getRange("B6").setBackground('#FFFFFF');
  shUserForm.getRange("B10").setBackground('#FFFFFF');
  shUserForm.getRange("B12").setBackground('#FFFFFF');
  shUserForm.getRange("B14").setBackground('#FFFFFF');
  shUserForm.getRange("B16").setBackground('#FFFFFF');
  shUserForm.getRange("B18").setBackground('#FFFFFF');
  
//Validating Employee ID
  if(shUserForm.getRange("B4").isBlank()==true){
    ui.alert("Please enter Account Name.");
    shUserForm.getRange("B4").activate();
    shUserForm.getRange("B4").setBackground('#FF0000');
    return false;
  }
 
 //Validating Employee Name
  else if(shUserForm.getRange("B6").isBlank()==true){
    ui.alert("Please enter Name.");
    shUserForm.getRange("B6").activate();
    shUserForm.getRange("B6").setBackground('#FF0000');
    return false;
  }
  //Validating Gender
  else if(shUserForm.getRange("B10").isBlank()==true){
    ui.alert("Please enter payor name.");
    shUserForm.getRange("B10").activate();
    shUserForm.getRange("B10").setBackground('#FF0000');
    return false;
  }
  //Validating Email ID
  else if(shUserForm.getRange("B12").isBlank()==true){
    ui.alert("Please enter a valid Email.");
    shUserForm.getRange("B12").activate();
    shUserForm.getRange("B12").setBackground('#FF0000');
    return false;
  }
  //Validating Department
  else if(shUserForm.getRange("B14").isBlank()==true){
    ui.alert("Please enter phone number.");
    shUserForm.getRange("B14").activate();
    shUserForm.getRange("B14").setBackground('#FF0000');
    return false;
  }
  //Validating Address
  else if(shUserForm.getRange("B16").isBlank()==true){
    ui.alert("Please enter date of intake.");
    shUserForm.getRange("B16").activate();
    shUserForm.getRange("B16").setBackground('#FF0000');
    return false;
  }
  else if(shUserForm.getRange("B18").isBlank()==true){
    ui.alert("Please enter first month payment details.");
    shUserForm.getRange("B18").activate();
    shUserForm.getRange("B18").setBackground('#FF0000');
    return false;
  }

  return true;

}
  // Function to submit the data to Database sheet
function submitData() {
     
  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 

  var shUserForm= myGooglSheet.getSheetByName("Account Entry"); //delcare a variable and set with the User Form worksheet

  var datasheet = myGooglSheet.getSheetByName("Data"); ////delcare a variable and set with the Database worksheet

  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
  
  // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
  // close the dialog by clicking the close button in its title bar.
  var response = ui.alert("Submit", 'Do you want to submit the data?',ui.ButtonSet.YES_NO);

  // Checking the user response and proceed with clearing the form if user selects Yes
  if (response == ui.Button.NO) 
  {return;//exit from this function
  } 
  var response = ui.alert("Submit", 'Are we billing insurance for this client?',ui.ButtonSet.YES_NO);
   if (response == ui.Button.YES) {
    var blankRowIns = datasheet.getLastRow() + 1;
    datasheet.getRange(blankRowIns, 1).setValue(shUserForm.getRange("B4").getValue().concat(" INS"));
   }
var title = shUserForm.getRange("B4").getValue();
var intakeDate = shUserForm.getRange("B16").getValue();
var calendar = CalendarApp.getCalendarById("c_a7s1780ic2j2e6mcir1e4jfq6g@group.calendar.google.com"); // Replace [CALENDAR_ID] with your calendar ID
Logger.log(intakeDate)

// Create a new calendar event
var recurrence = CalendarApp.newRecurrence().addMonthlyRule();
calendar.createAllDayEventSeries(title, new Date(intakeDate), recurrence);
       
  //Validating the entry. If validation is true then proceed with transferring the data to Database sheet
 if (validateEntry()==true) {
  
    var blankRow=datasheet.getLastRow()+1; //identify the next blank row

    datasheet.getRange(blankRow, 1).setValue(shUserForm.getRange("B4").getValue()); //Account Name
    datasheet.getRange(blankRow, 5).setValue(shUserForm.getRange("B6").getValue()); //Full Name
    datasheet.getRange(blankRow, 6).setValue(shUserForm.getRange("B8").getValue()); //Address
    datasheet.getRange(blankRow, 7).setValue(shUserForm.getRange("B10").getValue()); // Payor Name
    datasheet.getRange(blankRow, 8).setValue(shUserForm.getRange("B12").getValue()); //payor email
    datasheet.getRange(blankRow, 9).setValue(shUserForm.getRange("B14").getValue());// payor phone number
    datasheet.getRange(blankRow, 10).setValue(shUserForm.getRange("B16").getValue());//Date of intake
    datasheet.getRange(blankRow, 11).setValue(shUserForm.getRange("B18").getValue());//Month One
    datasheet.getRange(blankRow, 12).setValue(shUserForm.getRange("B20").getValue());
    datasheet.getRange(blankRow, 13).setValue(shUserForm.getRange("B22").getValue());
    datasheet.getRange(blankRow, 14).setValue(shUserForm.getRange("B24").getValue());
    datasheet.getRange(blankRow, 15).setValue(shUserForm.getRange("B26").getValue());
    datasheet.getRange(blankRow, 16).setValue(shUserForm.getRange("B28").getValue());
    datasheet.getRange(blankRow, 17).setValue(shUserForm.getRange("B30").getValue());
    // date function to update the current date and time as submittted on
    datasheet.getRange(blankRow, 18).setValue(new Date()).setNumberFormat('yyyy-mm-dd h:mm'); //Submitted On
    
    //get the email address of the person running the script and update as Submitted By
    datasheet.getRange(blankRow, 19).setValue(Session.getActiveUser().getEmail()); //Submitted By
    
    ui.alert(' "New Client Data Saved: ' + shUserForm.getRange("B4").getValue() +' "');
  
  //Clearnign the data from the Data Entry Form

    shUserForm.getRange("B4").clear();
    shUserForm.getRange("B6").clear();
    shUserForm.getRange("B8").clear();
    shUserForm.getRange("B10").clear();
    shUserForm.getRange("B12").clear();
    shUserForm.getRange("B14").clear();
    shUserForm.getRange("B16").clear();
    shUserForm.getRange("B18").clear();
    shUserForm.getRange("B20").clear();
    shUserForm.getRange("B22").clear();
    shUserForm.getRange("B24").clear();
    shUserForm.getRange("B26").clear();
    shUserForm.getRange("B28").clear();
    shUserForm.getRange("B30").clear();
      
 }
  
}
function searchRecord() {
  
  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm= myGooglSheet.getSheetByName("Account Entry"); //delcare a variable and set with the User Form worksheet
  var datasheet = myGooglSheet.getSheetByName("Data"); ////delcare a variable and set with the Database worksheet
  
  var str       = shUserForm.getRange("G4").getValue();
  var values    = datasheet.getDataRange().getValues(); //getting the entire values from the used range and assigning it to values variable
  
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
    shUserForm.getRange("B30").setValue(rowValue[16]);
    return;
  }
}
 
if(valuesFound=false){
  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
  ui.alert("No record found!");
 }
  
}


//Function to delete the record
 
function deleteRow() {
  
  var myGooglSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm= myGooglSheet.getSheetByName("Account Entry"); //delcare a variable and set with the User Form worksheet
  var datasheet = myGooglSheet.getSheetByName("Data"); ////delcare a variable and set with the Database worksheet
 
  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
  
  // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
  // close the dialog by clicking the close button in its title bar.
  var response = ui.alert("Submit", 'Do you want to delete the record?',ui.ButtonSet.YES_NO);
 
 // Checking the user response and proceed with clearing the form if user selects Yes
 if (response == ui.Button.NO) 
    {return;//exit from this function
 } 
    
  var str       = shUserForm.getRange("G4").getValue();
  var values    = datasheet.getDataRange().getValues(); //getting the entire values from the used range and assigning it to values variable
  
  var valuesFound=false; //variable to store boolean value to validate whether values found or not
  
  for (var i = 0; i<values.length; i++) 
    {
    var rowValue = values[i]; //declaraing a variable and storing the value
   
    //checking the first value of the record is equal to search item
    if (rowValue[0] == str) {
      
      var  iRow = i+1; //identify the row number
      datasheet.deleteRow(iRow) ; //deleting the row
 
      //message to confirm the action
      ui.alert(' "Record deleted' + shUserForm.getRange("G4").getValue() +' "');
 
      //Clearing the user form
      shUserForm.getRange("B4").clear() ;     
      shUserForm.getRange("B6").clear() ;
      shUserForm.getRange("B8").clear() ;
      shUserForm.getRange("B10").clear() ;
      shUserForm.getRange("B12").clear() ;
      shUserForm.getRange("B14").clear() ;
      shUserForm.getRange("B16").clear() ;
      shUserForm.getRange("B16").clear() ;
      shUserForm.getRange("B18").clear() ;
      shUserForm.getRange("B20").clear() ;
      shUserForm.getRange("B22").clear() ;
      shUserForm.getRange("B24").clear() ;
      shUserForm.getRange("B26").clear() ;
      shUserForm.getRange("B28").clear() ;
      shUserForm.getRange("B30").clear() ;
 
      valuesFound=true;
      return; //come out from the search function
      }
  }
 
if(valuesFound==false){
  //to create the instance of the user-interface environment to use the messagebox features
  var ui = SpreadsheetApp.getUi();
  ui.alert("No record found!");
 }
} 
function modifyRecord() {
  var myGooglSheet = SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm = myGooglSheet.getSheetByName("Account Entry"); //declare a variable and set with the Account Entry worksheet
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

  var str = shUserForm.getRange("B4").getValue();
  var values = datasheet.getDataRange().getValues(); //getting the entire values from the used range and assigning it to values variable

  var matchingRows = values.filter(function(row) {
    return row[0] === str && row[4] !== ''; //only return rows with matching value in column A and non-empty value in column E
  });

  // Check if only one row matches the search criteria
  if (matchingRows.length !== 1) {
    ui.alert("Multiple rows match the search criteria. Please ensure there is only one matching row and try again.");
    return;
  }

  // Only proceed if one row matches the search criteria
  var rowToUpdate = matchingRows[0];
  var iRow = values.indexOf(rowToUpdate) + 1;

      datasheet.getRange(iRow, 1).setValue(shUserForm.getRange("B4").getValue()); 
      datasheet.getRange(iRow, 5).setValue(shUserForm.getRange("B6").getValue());
      datasheet.getRange(iRow, 6).setValue(shUserForm.getRange("B8").getValue()); 
      datasheet.getRange(iRow, 7).setValue(shUserForm.getRange("B10").getValue()); 
      datasheet.getRange(iRow, 8).setValue(shUserForm.getRange("B12").getValue()); 
      datasheet.getRange(iRow, 9).setValue(shUserForm.getRange("B14").getValue()); 
      datasheet.getRange(iRow, 10).setValue(shUserForm.getRange("B16").getValue());
      datasheet.getRange(iRow, 11).setValue(shUserForm.getRange("B18").getValue());
      datasheet.getRange(iRow, 12).setValue(shUserForm.getRange("B20").getValue());
      datasheet.getRange(iRow, 13).setValue(shUserForm.getRange("B22").getValue());
      datasheet.getRange(iRow, 14).setValue(shUserForm.getRange("B24").getValue());
      datasheet.getRange(iRow, 15).setValue(shUserForm.getRange("B26").getValue());
      datasheet.getRange(iRow, 16).setValue(shUserForm.getRange("B28").getValue());
      datasheet.getRange(iRow, 17).setValue(shUserForm.getRange("B30").getValue());
   
      // date function to update the current date and time as submittted on
      datasheet.getRange(iRow, 21).setValue(new Date()).setNumberFormat('yyyy-mm-dd h:mm'); //Submitted On
    
      //get the email address of the person running the script and update as Submitted By
      datasheet.getRange(iRow, 22).setValue(Session.getActiveUser().getEmail()); //Submitted By
    
      ui.alert(' "Data updated for - Client' + shUserForm.getRange("B4").getValue() +' "');
  
    //Clearnign the data from the Data Entry Form

      shUserForm.getRange("B4").clear() ;     
      shUserForm.getRange("B6").clear() ;
      shUserForm.getRange("B8").clear() ;
      shUserForm.getRange("B10").clear() ;
      shUserForm.getRange("B12").clear() ;
      shUserForm.getRange("B14").clear() ;
      shUserForm.getRange("B16").clear() ;      
      shUserForm.getRange("B18").clear() ;
      shUserForm.getRange("B20").clear() ;
      shUserForm.getRange("B22").clear() ;
      shUserForm.getRange("B24").clear() ;
      shUserForm.getRange("B26").clear() ;
      shUserForm.getRange("B28").clear() ;
      shUserForm.getRange("B30").clear() ;

      valuesFound=true;
      return; //come out from the search function
      }
