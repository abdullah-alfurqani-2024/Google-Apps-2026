function deleteEmptyRowsAfterQuery() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Copy of Report 2"); // ← غيّر هذا
  
  var startRow = 14;  // بداية الـ QUERY
  var checkCol = 6;   // العمود F (يدوي حقيقي)
  var lastRow = sheet.getLastRow();
  
  if (lastRow < startRow) return;
  
  var data = sheet.getRange(startRow, checkCol, lastRow - startRow + 1, 1).getValues();
  
  // نحذف من الأسفل لأعلى لتجنب اختلال الأرقام
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === "" || data[i][0] === null) {
      sheet.deleteRow(startRow + i);
    }
  }
}
