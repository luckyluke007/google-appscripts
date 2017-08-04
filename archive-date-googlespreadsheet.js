function oldArchive() {
	 var ss = SpreadsheetApp.getActiveSpreadsheet();
	 var approvedSheet = ss.getSheetByName('APPROVED');
	 var archivedSheet = ss.getSheetByName('ARCHIVE');
	 var lastColumn = scheduleSheet.getLastColumn();

	 for (var i = scheduleSheet.getLastRow(); i > 0; i--) {

	 		var dataCell  = scheduleSheet.getRange(i, 1).getValue();
	 		if(isValidDate(datecell)) {
	 			var today = new Date();
	 			var test = new Date(dateCell);

	 			if(test < today){
	 				var rangeToMove = scheduleSheet.getRange(i, 1, 1, scheduleSheet.getLastColumn()).getValues();
	 				pastSheet.getRange(pastSheet.getLastRow() + 1, 1, 1, scheduleSheet.getLastColumn()).setValues(rangeToMove);
	 				scheduleSheet.deleteRow(i + 1);
	 			}
	 		}
	 }
}

function isValidDate(value) {
	var dateWrapper = new Date(value);
	return !isNaN(dateWrapper.getDate());
}