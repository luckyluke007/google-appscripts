# Google Appscript Approval and Reject event - with email notification

- Yes/No column move row to Approval or Reject sheet and sent out email notifications

```
//Moved row data base on column B answer Yes or No
var APPROVED = '<p>Approved Message</p>';
var REJECTED = '<p>Rejected Message</p>';
var EMAIL_SENT = "Sent";

function processApproval(event) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var r = s.getActiveRange();
 
  if(s.getName() =="CPF" && r.getColumn() == 2) { //CPF Sheet name of the list to look for value column 2 for Yes or No
    var targetSheetName = null;
    var message = null;
    
    switch(r.getValue()) {
      case "yes":
      case "Yes":
      case "YES":
        targetSheetName = "Approved"; // If Yes move data to sheet Approved
        message = APPROVED;
        break;
        
      case "no":
      case "No":
      case "NO":
        targetSheetName = "Rejected"; // If No move data to sheet Approved
        message = REJECTED;
        break;
    }
    
    if(targetSheetName != null) {
      // Move row to appropriate sheet
      var row = r.getRow();
      var targetSheet = ss.getSheetByName(targetSheetName);
      var targetRange = targetSheet.getRange(targetSheet.getLastRow()+1,1);
      var numColumns = s.getLastColumn();
      
      s.getRange(row, 1, 1, numColumns).moveTo(targetRange);
      targetSheet.getRange(targetRange.getRow(), 24).setValue(message);

      // Send email notification
      var email = targetSheet.getRange(targetRange.getRow(), 22).getValue(); // Getting email value 
      var subject = targetSheet.getRange(targetRange.getRow(), 3).getValue(); // Getting column 3 for email subject
      var summary = targetSheet.getRange(targetRange.getRow(), 25).getValue(); // Getting summary column for email body
      MailApp.sendEmail(email, subject, "",{htmlBody:message + summary});

      // Mark row as sent
      targetSheet.getRange(targetRange.getRow(), 23).setValue(EMAIL_SENT); // Send mail to notify.

      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  } 
}
```
