# Google Appscript Approval and Reject event - with email notification

- Yes/No column move row to Approval or Reject sheet and sent out email notifications

```
var APPROVED = '<p>Your request for the Campus Provost/Executive Vice Chancellor participation at your upcoming event has been approved.</p><p>Please complete the <a href="https://cpevc.ucsc.edu/files/cpevc_briefing_template.zip">Briefing Form</a> with all details on the event. <span style="color:red">Note: The Briefing Form must be submitted to Alison Schwab one week before the event. To submit Briefing Form <a href="https://ucsc.wufoo.com/forms/cpevcs-event-briefing-packet/">click here</a>.</span></p><p>Please be in contact with Marc Deslardins at <a href="mailto:madesjar@ucsc.edu">madesjar@ucsc.edu</a> for speaking notes if you have requested the CP/EVC to make remarks.</p><p>If you have any changes to your request or there are changes to the event please contact Roxann Gonzales at <a href="mailto:rgonza49@ucsc.edu">rgonza49@ucsc.edu</a>, 459-3885.</p><p>Thank you.</p>';
var REJECTED = "<p>Unfortunately, due to competing priorities, your request for the Campus Provost/Executive Vice Chancellor's participation at your upcoming event is respectfully declined.</p><p>Thank you.</p>";;
var EMAIL_SENT = "Sent";

function processApproval(event) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = event.source.getActiveSheet();
  var r = event.source.getActiveRange();
 
  if(s.getName() =="REQUESTS" && r.getColumn() == 2) {
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
      targetSheet.getRange(targetRange.getRow(), 27).setValue(message);

      // Send email notification
      var email = targetSheet.getRange(targetRange.getRow(), 25).getValue();
      var subject = targetSheet.getRange(targetRange.getRow(), 3).getValue();
      var summary = targetSheet.getRange(targetRange.getRow(), 28).getValue();
      MailApp.sendEmail(email, subject, "",{htmlBody:message + summary});

      // Mark row as sent
      targetSheet.getRange(targetRange.getRow(), 26).setValue(EMAIL_SENT);

      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  } 
}
```
