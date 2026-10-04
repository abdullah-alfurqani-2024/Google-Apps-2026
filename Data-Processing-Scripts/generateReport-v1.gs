// الكود في الملف الثاني (Data Script)
function generateReportFromLibrary(startDateInput, endDateInput) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("جميع المهام") || ss.getSheets()[0]; 
  
  if (!masterSheet) {
    return { success: false, error: "لم يتم العثور على الورقة المطلوبة." };
  }

  // ⚙️ إعدادات صف البداية والأعمدة
  var START_ROW = 37;
  var DATE_COL = 7;
  var STATUS_COL = 12;
  var TASK_TYPE_COL = 10;
  var TASK_CATEGORY_COL = 11;

  var startParts = startDateInput.split('-');
  var start = new Date(startParts[0], startParts[1] - 1, startParts[2], 0, 0, 0).getTime();
  
  var endParts = endDateInput.split('-');
  var end = new Date(endParts[0], endParts[1] - 1, endParts[2], 23, 59, 59).getTime();
  
  var arabicMonths = [
    "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
    "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
  ];
  
  var monthIndex = parseInt(startParts[1], 10) - 1;
  var monthNameArabic = arabicMonths[monthIndex];
  var yearString = startParts[0];

  // قراءة كامل البيانات الخام
  var lastRow = masterSheet.getLastRow();
  var lastCol = masterSheet.getLastColumn();
  
  if (lastRow < START_ROW) {
    return { success: false, error: "لا توجد بيانات كافية في الشيت." };
  }
  
  var data = masterSheet.getRange(1, 1, lastRow, lastCol).getValues();

  var totalTasks = 0;
  var completedTasks = 0;
  var pendingTasks = 0;

  var TotalTaskRequest = 0;
  var TotalTechReq = 0;
  var TotalEvaluation = 0;
  var TotalOthers = 0;
  
  function parseSheetDate(cellValue) {
    if (!cellValue) return null;
    if (cellValue instanceof Date) return cellValue.getTime();
    if (typeof cellValue === 'string') {
      var parts = cellValue.trim().split(/[\/\-\.]/);
      if (parts.length === 3) {
        if (parts[0].length === 4) return new Date(parts[0], parts[1] - 1, parts[2]).getTime();
        else if (parts[2].length === 4) return new Date(parts[2], parts[1] - 1, parts[0]).getTime();
      }
      var parsed = Date.parse(cellValue);
      if (!isNaN(parsed)) return parsed;
    }
    return null;
  }

  for (var i = START_ROW-1; i < data.length; i++) {
    var rawDate = data[i][DATE_COL-1]; // العمود B
    var rowTimestamp = parseSheetDate(rawDate);
    
    if (rowTimestamp && rowTimestamp >= start && rowTimestamp <= end) {
      totalTasks++;
      
      var status = String(data[i][STATUS_COL-1]).trim(); // العمود F (الحالة)
      if (status === "مكتملة" || status === "مكتمل" || status.toLowerCase() === "completed") {
        completedTasks++;
        var taskType = String(data[i][TASK_TYPE_COL-1]).trim();
        if (taskType === "طلب خدمة") {TotalTaskRequest ++;}
        if (taskType === "مراسلة") {
          var taskCategory = String(data[i][TASK_CATEGORY_COL-1]).trim();
          if (taskCategory === "المواصفات الفنية") {TotalTechReq ++;}
          if (taskCategory === "التحليل والتقييم") {TotalEvaluation ++;}
          if (taskCategory === "أخرى") {TotalOthers ++;}
        }
      } else {
        pendingTasks++;
      }
    }
  }
  
  // IDs نموذج التقرير ومجلد الحفظ
  const TEMPLATE_ID   = '1EFRw1teEpAorZUNE6OF89ydSBoIgkX4iqbcSnPEmTlY';
  const FOLDER_ID = '1ib6QCtktE8L_Y0Va46ZBJ8-4dj5n4Hn1';
  
  try {
    var templateFile = DriveApp.getFileById(TEMPLATE_ID);
    var targetFolder = DriveApp.getFolderById(FOLDER_ID);
    
    var newDocName = "تقرير الدعم الفني - " + monthNameArabic + " " + yearString;
    var newDocFile = templateFile.makeCopy(newDocName, targetFolder);
    var doc = DocumentApp.openById(newDocFile.getId());
    var body = doc.getBody();
    
    body.replaceText("<<الشهر>>", monthNameArabic);
    body.replaceText("<<السنة>>", yearString);
    body.replaceText("<<تاريخ_البدء>>", startDateInput);
    body.replaceText("<<تاريخ_النهاية>>", endDateInput);
    body.replaceText("<<إجمالي_المهام>>", totalTasks);
    body.replaceText("<<المهام_المكتملة>>", completedTasks);
    body.replaceText("<<المهام_المعلقة>>", pendingTasks);
    
    doc.saveAndClose();
    
    return {
      success: true,
      url: newDocFile.getUrl(),
      total: totalTasks
    };
    
  } catch (e) {
    return {
      success: false,
      error: e.toString()
    };
  }
}
