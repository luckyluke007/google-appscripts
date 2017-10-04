# Google Appscript Approval and Reject event - with email notification

- Yes/No column move row to Approval or Reject sheet and sent out email notifications

```
//Moved row data base on column B answer Yes or No
var APPROVED = '<p>Approved Message</p>';
var REJECTED = '<p>Rejected Message</p>';
var EMAIL_SENT = "Sent";

function processApproval(event) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = event.source.getActiveSheet();
  var r = event.source.getActiveRange();
 
  if(s.getName() =="CPF" && r.getColumn() == 2) {
    var targetSheetName = null;
    var message = null;
    
    switch(r.getValue()) {
      case "yes":
      case "Yes":
      case "YES":
        targetSheetName = "Approved";
        message = APPROVED;
        break;
        
      case "no":
      case "No":
      case "NO":
        targetSheetName = "Rejected";
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
      var email = targetSheet.getRange(targetRange.getRow(), 22).getValue();
      var subject = targetSheet.getRange(targetRange.getRow(), 3).getValue();
      var summary = targetSheet.getRange(targetRange.getRow(), 25).getValue();
      MailApp.sendEmail(email, subject, "",{htmlBody:message + summary});

      // Mark row as sent
      targetSheet.getRange(targetRange.getRow(), 23).setValue(EMAIL_SENT);

      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  } 
}
```
