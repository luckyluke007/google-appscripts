# Google Appscript Sorting rows by date column

- Yes/No column move row to Approval or Reject sheet and sent out email notifications

```
function onOpen(event) {
  var sheet = event.source.getSheetByName("APPROVED"); //Sheet name in qoute
  var editedCell = sheet.getActiveCell();
 
  var columnToSortBy = 1; //column one date 
  var tableRange = "A2:X955"; //sort the rows

    if(editedCell.getColumn() == columnToSortBy){
      var range = sheet.getRange(tableRange);
      range.sort( { column : columnToSortBy, ascending: false } );
  }
}
```
