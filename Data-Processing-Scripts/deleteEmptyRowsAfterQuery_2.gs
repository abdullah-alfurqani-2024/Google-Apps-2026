function deleteEmptyRowsAfterQuery_2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Copy of Report 2"); // ← غيّر هذا
  
  var lastDataRow = sheet.getLastRow();
  var maxRows = sheet.getMaxRows();
  
  // إذا يوجد صفوف فارغة بعد آخر بيان
  if (maxRows > lastDataRow) {
    sheet.deleteRows(lastDataRow + 1, maxRows - lastDataRow);
  }
}
