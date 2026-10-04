function copySheetsForLooker() {
  // ==== CONFIG ====
  var SOURCE_FILE_ID = "1yaWxaA9_pc4HBlPD_6Vxu_LR5MKwUH2LtZEhBYio-3w";      // source file
  var DEST_FILE_ID   = "1ADAHt37H-au7Nt3ZhFsK7Wj4sRTC8-Lpe2798Xnt2r4";        // destination (formatted) file
  
  // Define which sheets and ranges to copy:
  // Each entry: {sheet: "SheetName", sourceRange: "A2:D50", destStart: "B3"}
  var SHEETS_TO_COPY = [
    {sheet: "Report", sourceRange: "B14:M", destStart: "B2"}
  ];
  // ==================

  var source = SpreadsheetApp.openById(SOURCE_FILE_ID);
  var dest   = SpreadsheetApp.openById(DEST_FILE_ID);

  SHEETS_TO_COPY.forEach(function(config) {
    var s = source.getSheetByName(config.sheet);
    var d = dest.getSheetByName(config.sheet);

    if (s && d) {
      // get data from source range
      var srcRange = s.getRange(config.sourceRange);
      var data = srcRange.getValues();

      // get destination starting cell
      var destCell = d.getRange(config.destStart);
      var row = destCell.getRow();
      var col = destCell.getColumn();

      // clear destination area of same size (content only, keep formatting)
      d.getRange(row, col, data.length, data[0].length).clearContent();

      // paste values
      d.getRange(row, col, data.length, data[0].length).setValues(data);
    }
  });

  Logger.log("Data copied successfully into formatted file (Tasks-Reports-Looker-Live): " + dest.getUrl());
}
