function clearForm() 
{
  var myGoogleSheet= SpreadsheetApp.getActiveSpreadsheet(); //declare a variable and set with active google sheet 
  var shUserForm    = myGoogleSheet.getSheetByName("Account Entry"); //declare a variable and set with the User Form worksheet
 
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
  shUserForm.getRange("B12").clear(); // Email ID
  shUserForm.getRange("B14").clear(); //Department
  shUserForm.getRange("B16").clear();//Address
  shUserForm.getRange("B18").clear();
  shUserForm.getRange("B20").clear();
  shUserForm.getRange("B22").clear();
  shUserForm.getRange("B24").clear();
  shUserForm.getRange("B26").clear();
  shUserForm.getRange("B28").clear();
  shUserForm.getRange("B30").clear();
 //Assigning white as default background color
 
 shUserForm.getRange("B4").setBackground('#FFFFFF');
 shUserForm.getRange("B6").setBackground('#FFFFFF');
 shUserForm.getRange("B8").setBackground('#FFFFFF');
 shUserForm.getRange("B10").setBackground('#FFFFFF');
 shUserForm.getRange("B12").setBackground('#FFFFFF');
 shUserForm.getRange("B14").setBackground('#FFFFFF');
 shUserForm.getRange("B16").setBackground('#FFFFFF');
 shUserForm.getRange("B18").setBackground('#FFFFFF');
 shUserForm.getRange("B20").setBackground('#FFFFFF');
 shUserForm.getRange("B22").setBackground('#FFFFFF');
 shUserForm.getRange("B24").setBackground('#FFFFFF');
 shUserForm.getRange("B26").setBackground('#FFFFFF');
 shUserForm.getRange("B28").setBackground('#FFFFFF');
 shUserForm.getRange("B30").setBackground('#FFFFFF');
 
  return true ;
  
  }
}
