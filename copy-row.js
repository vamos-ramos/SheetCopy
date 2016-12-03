// Property of Lazara Ramos


/*
 * assumes source data in sheet named Needed
 * target sheet of move to named Acquired
 * test column with yes/no is col 4 or D
 */

function onEdit(event) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // the whole spreadsheet
  var s = event.source.getActiveSheet(); // the source sheet
  var r = event.source.getActiveRange(); // the area the event fired in

  if(s.getName() == "Ingredient-Request" && r.getColumn() == 6 && r.getValue() == "yes") {
    var row = r.getRow(); // the row where the edit event occurred
    var numColumns = s.getLastColumn(); // Returns the position of the last column that has content.
    var targetSheet = ss.getSheetByName("form-submitted");
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1); // write to the first blank row, starting from far left column
    s.getRange(row, 1, 1, numColumns).copyTo(target);
  }

  if(s.getName() == "form-submitted" && r.getColumn() == 7 && r.getValue() == "yes") {
    var row = r.getRow(); // the row where the edit event occurred
    var numColumns = s.getLastColumn(); // Returns the position of the last column that has content.
    var targetSheet = ss.getSheetByName("added-to-WT");
    var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1); // write to the first blank row, starting from far left column
    s.getRange(row, 1, 1, numColumns).copyTo(target);
  }
}
