# Google Appscript Sorting rows by date column

- Yes/No column move row to Approval or Reject sheet and sent out email notifications

```
//Archiving by date
function approveRequests() {

  // Initialising
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = ss.getSheetByName("APPROVED");
  var pastSheet = ss.getSheetByName("ARCHIVE");
  var lastColumn = scheduleSheet.getLastColumn();

  // Check all values from your "APPROVED" sheet
  for(var i = scheduleSheet.getLastRow(); i > 0; i--){

    // Check if the value is a valid date
    var dateCell = scheduleSheet.getRange(i, 1).getValue();
    if(isValidDate(dateCell)){
      var today = new Date();
      var test = new Date(dateCell);

      // If the value is a valid date and is a past date, we remove it from the sheet to paste on the other sheet
      if(test < today){

        var rangeToMove = scheduleSheet.getRange(i, 1, 1, scheduleSheet.getLastColumn()).getValues();
        pastSheet.getRange(pastSheet.getLastRow() + 1, 1, 1, scheduleSheet.getLastColumn()).setValues(rangeToMove);
        scheduleSheet.deleteRow(i);
      }
    }
  }
}

// Check is a valid date
function isValidDate(value) {
  var dateWrapper = new Date(value);
  return !isNaN(dateWrapper.getDate());
}
```
