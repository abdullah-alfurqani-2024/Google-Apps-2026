function diagnose() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Copy of Report 2"); // ← غيّر هذا
  
  var startRow = 14;
  var checkCol = 6;
  var lastRow = sheet.getLastRow();
  
  Logger.log("آخر صف: " + lastRow);
  
  var data = sheet.getRange(startRow, checkCol, lastRow - startRow + 1, 1).getValues();
  
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === "" || data[i][0] === null) {
      Logger.log("صف فارغ: " + (startRow + i) + " | القيمة: [" + data[i][0] + "]");
    }
  }
}
