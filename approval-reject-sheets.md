# Google Appscript Approval and Reject event - with email notification

- Yes/No column move row to Approval or Reject sheet and sent out email notifications

```
//Approved
var APPROVED = '<p>Approved Message</p>';
var REJECTED = '<p>Rejected Message</p>';
var EMAIL_SENT = "Sent";

function processApproval(event) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = event.source.getActiveSheet();
  var r = event.source.getActiveRange();
  
  // Main Sheet Name - Set Yes or No on column too or assign column you want
  if(s.getName() =="Sheet1" && r.getColumn() == 2) {
    var targetSheetName = null;
    var message = null;
    
    switch(r.getValue()) {
      case "Yes":
      case "YES":
        targetSheetName = "Approved";
        message = APPROVED;
        break;
        
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
      
      // Apply message (Message set row 24)
      s.getRange(row, 1, 1, numColumns).moveTo(targetRange);
      targetSheet.getRange(targetRange.getRow(), 24).setValue(message);

      // Send email notification - (Email is set row 22)
      var email = targetSheet.getRange(targetRange.getRow(), 22).getValue();
      var subject = "Chancellor's Participation Request";
      MailApp.sendEmail(email, subject, "",{htmlBody:message});

      // Mark row as sent - status email sent (Set row 23)
      targetSheet.getRange(targetRange.getRow(), 23).setValue(EMAIL_SENT);

      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  } 
}
```
