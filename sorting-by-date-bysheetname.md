# Google Appscript Sorting rows by date column

- Yes/No column move row to Approval or Reject sheet and sent out email notifications

```
//Sorting by date in column 1
function onOpen(event) {
  var sheet = event.source.getSheetByName("APPROVED");
  var editedCell = sheet.getActiveCell();
 
  var columnToSortBy = 1;
  var tableRange = "A2:AB955";

    if(editedCell.getColumn() == columnToSortBy){
      var range = sheet.getRange(tableRange);
      range.sort( { column : columnToSortBy, ascending: false } );
  }
}
```
