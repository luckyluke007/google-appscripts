# Google Appscript Approval and Reject event - with email notification

- Yes/No column move row to Approval or Reject sheet and sent out email notifications

```
//Approved
var APPROVED = '<p>Your request for the Chancellor to participate in your upcoming event has been approved.</p><p>Please complete the <a href="http://chancellor.ucsc.edu/files/chancellor_briefing_template.zip">Briefing Form</a> with all details on the event. <span style="color:red">Note: The Briefing Form must be submitted to Margaret McGuire one week before the event. To submit Briefing Form <a href="https://ucsc.wufoo.com/forms/chancellors-event-briefing-packet/">click here</a>.</span></p>';
var REJECTED = '<p>Thank you for completing a Chancellor’s Participation Form for your upcoming event.  Unfortunately, schedules do not permit the Chancellor to participate.  You might consider asking other campus leadership to attend.</p><p>Thank you again for contacting us.</p>';
var EMAIL_SENT = "Sent";

function processApproval(event) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = event.source.getActiveSheet();
  var r = event.source.getActiveRange();
  
  // Main Sheet Name - Set Yes or No on column too or assign column you want
  if(s.getName() =="Main Sheet Name Goes Here" && r.getColumn() == 2) {
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
      
      // Apply message Approval or Rejected (Row 24 or change row #)
      s.getRange(row, 1, 1, numColumns).moveTo(targetRange);
      targetSheet.getRange(targetRange.getRow(), 24).setValue(message);

      // Send email notification - get email address (row 22 or change row #)
      var email = targetSheet.getRange(targetRange.getRow(), 22).getValue();
      var subject = "Chancellor's Participation Request";
      MailApp.sendEmail(email, subject, "",{htmlBody:message});

      // Mark row as sent - status email sent (row 23 or change row #)
      targetSheet.getRange(targetRange.getRow(), 23).setValue(EMAIL_SENT);

      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  } 
}
```
