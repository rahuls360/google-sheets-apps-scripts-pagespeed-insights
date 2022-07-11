// Menu 3
//Log the values and clear PageSpeed Results data:
function runLog() {
  var columnNumberToWatch = 13; // column A = 1, B = 2, etc.
  var valueToWatch = 'complete';
  var sheetNameToMoveTheRowTo = 'Log';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Results').activate();
  var cell = sheet.getRange('M6:M');
  var type = SpreadsheetApp.CopyPasteType.PASTE_VALUES;
  var lastRow = sheet.getLastRow();

  var Avals = ss.getRange('M6:M').getValues();
  var Alast = Avals.filter(String).length;

  if (
    sheet.getName() != sheetNameToMoveTheRowTo &&
    cell.getColumn() == columnNumberToWatch &&
    cell.getValue().toLowerCase() == valueToWatch
  ) {
    var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
    var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    sheet
      .getRange(cell.getRow(), 2, Alast, sheet.getLastColumn())
      .copyTo(targetRange, type, false);
    sheet.getRange('C6:M').clearContent();
  }
}
