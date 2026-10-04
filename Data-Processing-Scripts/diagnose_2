function diagnose_2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Copy of Report 2"); // ← غيّر هذا
  
  var startRow = 14;
  var checkCol = 6;
  var lastRow = sheet.getLastRow();
  
  Logger.log("آخر صف: " + lastRow);
  
  var data = sheet.getRange(startRow, checkCol, lastRow - startRow + 1, 1).getValues();
  
  // اطبع آخر 20 صف لنرى ما فيها
  for (var i = data.length - 20; i < data.length; i++) {
    Logger.log("صف " + (startRow + i) + " | القيمة: [" + data[i][0] + "] | النوع: " + typeof data[i][0]);
  }
}
