// Property of Lazara Ramos

function spreadSheetEdit(event) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // the whole spreadsheet
  var s = event.source.getActiveSheet(); // the source sheet
  var r = event.source.getActiveRange(); // the area the event fired in

  // entry complete
  if(isEntryComplete(s, r)) {
    processEvent(ss, s, r, "Entry Complete");
  }

  // approvedbyCulinary
  if(approvedbyCulinary(s, r)) {
    processEvent(ss, s, r, "Approved by Culinary");
  }

  // rejectedbyCulinary - rejection email needed here
  if(rejectedbyCulinary(s, r)) {
    processEvent(ss, s, r, "Rejections");
    sendEmail(s, r, 4, "culinary");
  }

  // approvedbyNutrition
  if(approvedbyNutrition(s, r)) {
    processEvent(ss, s, r, "Approved by Nutrition");
  }

  // rejectedbyNutrition - rejection email needed here
  if(rejectedbyNutrition(s, r)) {
    processEvent(ss, s, r, "Rejections");
    sendEmail(s, r, 4, "nutrition");
  }
}

function processEvent(spreadsheet, activeSheet, activeRange, sheetName) {
  var row = activeRange.getRow();
  var numColumns = activeSheet.getLastColumn();
  var targetSheet = spreadsheet.getSheetByName(sheetName);
  var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
  activeSheet.getRange(row, 1, 1, numColumns).copyTo(target);
}

function isEntryComplete(activeSheet, activeRange) {
  return activeSheet.getName() == "BRGR", "Vegetable Sides" && activeRange.getColumn() == 10 && activeRange.getValue() == "Entry Complete";
}

function approvedbyCulinary(activeSheet, activeRange) {
  return activeSheet.getName() == "Entry Complete" && activeRange.getColumn() == 13 && activeRange.getValue() == "GB";
}

function rejectedbyCulinary(activeSheet, activeRange) {
  return activeSheet.getName() == "Entry Complete" && activeRange.getColumn() == 13 && activeRange.getValue() == "reject";
}

function approvedbyNutrition(activeSheet, activeRange) {
  return activeSheet.getName() == "Approved by Culinary" && activeRange.getColumn() == 15 && activeRange.getValue() == "KM" ;
}

function rejectedbyNutrition(activeSheet, activeRange) {
  return activeSheet.getName() == "Approved by Culinary" && activeRange.getColumn() == 15 && activeRange.getValue() == "reject";
}


function sendEmail(activeSheet, activeRange, emailColIndex, rejectionType) {
  var row = activeRange.getRow();
  var numColumns = activeSheet.getLastColumn();

  var emailCell = activeSheet.getRange(row, emailColIndex, 1, 1);
  var emailAddress = emailCell.getDisplayValues()[0][0];

  var mrfColIndex = 5;
  var mrfCell = activeSheet.getRange(row, mrfColIndex, 1, 1);
  var mrf = mrfCell.getDisplayValues()[0][0];

  var recipeColIndex = 7;
  var recipeCell = activeSheet.getRange(row, recipeColIndex, 1, 1);
  var recipe = recipeCell.getDisplayValues()[0][0];

  var emailBody = "MRF number: " + mrf + "\n" + "Recipe name: " + recipe + "\n";

  if(rejectionType == "culinary") {
    var culNotesColIndex = 12;
    var culinaryNotesCell = activeSheet.getRange(row, culNotesColIndex, 1, 1);
    var culinaryNotes = culinaryNotesCell.getDisplayValues()[0][0];
    emailBody += "Culinary notes: " + culinaryNotes;
  } else {
    var nutNotesColIndex = 14;
    var nutNotesCell = activeSheet.getRange(row, nutNotesColIndex, 1, 1);
    var nutritionNotes = nutNotesCell.getDisplayValues()[0][0];
    emailBody += "Nutrition notes: " + nutritionNotes;
  }

  MailApp.sendEmail(
    emailAddress,
    "Your recipe has been rejected",
    emailBody
  );
}
