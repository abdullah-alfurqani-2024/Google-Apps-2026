/**
 * ============================================================================
 * مكتبة توليد تقارير الدعم الفني الآلية (ReportLibrary)
 * ============================================================================
 */

function generateReportFromLibrary(startDateInput, endDateInput) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Report") || ss.getSheets()[0]; 
  
  if (!masterSheet) {
    return { success: false, error: "لم يتم العثور على ورقة البيانات المطلوبة." };
  }

  // ⚙️ [تعديل الإعدادات هنا] ------------------------------------------------
  var START_ROW = 37;
  var DATE_COL = 7;
  var STATUS_COL = 12;
  var TASK_TYPE_COL = 10;
  var TASK_CATEGORY_COL = 11;

  // IDs نموذج التقرير ومجلد الحفظ
  const TEMPLATE_ID   = '1EFRw1teEpAorZUNE6OF89ydSBoIgkX4iqbcSnPEmTlY'; // v-5
  // const TEMPLATE_ID   = '1AvSlcKgz-GOpSruRcBCUtZZAD_UV5_qd3hNUnKBKYtY'; // v-4
  const FOLDER_ID = '1ib6QCtktE8L_Y0Va46ZBJ8-4dj5n4Hn1';
  // ------------------------------------------------------------------------

  // 1. معالجة وتجهيز نطاق التواريخ
  var startParts = startDateInput.split('-');
  var endParts = endDateInput.split('-');

  var startDateObj = new Date(startParts[0], parseInt(startParts[1], 10) - 1, startParts[2]);
  startDateObj.setHours(0, 0, 0, 0);
  var startTime = startDateObj.getTime();

  var endDateObj = new Date(endParts[0], parseInt(endParts[1], 10) - 1, endParts[2]);
  endDateObj.setHours(23, 59, 59, 999);
  var endTime = endDateObj.getTime();

  var arabicMonths = [
    "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
    "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
  ];
  var monthIndex = parseInt(startParts[1], 10) - 1;
  var monthNameArabic = arabicMonths[monthIndex] || "";
  var yearString = String(startParts[0]);

  // 2. قراءة البيانات الخام مباشرة لتجاوز الفلاتر
  var lastRow = masterSheet.getLastRow();
  var lastCol = masterSheet.getLastColumn();
  
  if (lastRow < START_ROW) {
    return { success: false, error: "لا توجد بيانات كافية في ورقة العمل." };
  }
  
  var data = masterSheet.getRange(1, 1, lastRow, lastCol).getValues();

  var totalTasks = 0;
  var completedTasks = 0;
  var pendingTasks = 0;

  var TotalTaskRequest = 0;
  var TotalTechReq = 0;
  var TotalEvaluation = 0;
  var TotalOthers = 0;
  var TotalCalls = 0;
  var TotalDirectAssignments = 0;
  var TotalDirectRequests = 0;

  // 3. فلترة البيانات وحساب الإحصائيات
  for (var i = START_ROW - 1; i < data.length; i++) {
    var rawDate = data[i][DATE_COL - 1];
    if (!rawDate) continue;

    var rowDateObj = parseSheetDate(rawDate);
    
    if (rowDateObj) {
      var rowTime = rowDateObj.getTime();

      if (rowTime >= startTime && rowTime <= endTime) {
        totalTasks++;
        var status = String(data[i][STATUS_COL - 1] || "").trim().toLowerCase();
        
        if (status === "مكتملة" || status === "مكتمل" || status.toLowerCase() === "completed") {
        completedTasks++;
        var taskType = String(data[i][TASK_TYPE_COL-1]).trim();
        if (taskType === "طلب خدمة") {TotalTaskRequest ++;}
        if (taskType === "اتصال") {TotalCalls ++;}
        if (taskType === "طلب مباشر") {TotalDirectRequests ++;}
        if (taskType === "تكليف مباشر") {TotalDirectAssignments ++;}
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
  }

  // 4. إنشاء المستند واستبدال الوسوم
  try {
    var templateFile = DriveApp.getFileById(TEMPLATE_ID);
    var targetFolder = DriveApp.getFolderById(FOLDER_ID);
    
    var newDocName = "تقرير الدعم الفني - " + monthNameArabic + " " + yearString;
    var newDocFile = templateFile.makeCopy(newDocName, targetFolder);
    var doc = DocumentApp.openById(newDocFile.getId());

    var replacements = {
      "<<طلبات_الخدمة>>": String(TotalTaskRequest),
      "<<مراسلات_مواصفات>>": String(TotalTechReq),
      "<<مراسلات_تحليل>>": String(TotalEvaluation),
      "<<مراسلات_أخرى>>": String(TotalOthers),
      "<<اتصال>>": String(TotalCalls),
      "<<طلب_مباشر>>": String(TotalDirectRequests),
      "<<تكليف_مباشر>>": String(TotalDirectAssignments),
      "<<الشهر>>": String(monthNameArabic || ""),
      "<<السنة>>": String(yearString || "")
    };

    replaceAllTags(doc, replacements);
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

/**
 * دالة استبدال الوسوم الشاملة (تغطي النصوص والجداول والهيدر والفوتر)
 */
function replaceAllTags(doc, replacements) {
  var body = doc.getBody();
  
  var numChildren = body.getNumChildren();
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    var type = child.getType();
    
    if (type == DocumentApp.ElementType.TABLE) {
      var table = child.asTable();
      for (var r = 0; r < table.getNumRows(); r++) {
        var row = table.getRow(r);
        for (var c = 0; c < row.getNumCells(); c++) {
          var cell = row.getCell(c);
          for (var key in replacements) {
            cell.replaceText(key, replacements[key]);
          }
        }
      }
    } else {
      for (var key in replacements) {
        body.replaceText(key, replacements[key]);
      }
    }
  }

  var header = doc.getHeader();
  if (header) {
    for (var key in replacements) { header.replaceText(key, replacements[key]); }
  }

  var footer = doc.getFooter();
  if (footer) {
    for (var key in replacements) { footer.replaceText(key, replacements[key]); }
  }
}

/**
 * دالة تحليل التواريخ المرنة
 */
function parseSheetDate(cellValue) {
  if (!cellValue) return null;

  if (cellValue instanceof Date) {
    var d = new Date(cellValue);
    d.setHours(12, 0, 0, 0);
    return d;
  }

  if (typeof cellValue === 'string') {
    var str = cellValue.trim();
    if (str === "") return null;

    var datePart = str.split(" ")[0];
    var parts = datePart.split(/[\/\-\.]/);

    if (parts.length === 3) {
      var p1 = parseInt(parts[0], 10);
      var p2 = parseInt(parts[1], 10);
      var p3 = parseInt(parts[2], 10);

      if (parts[0].length === 4) {
        return new Date(p1, p2 - 1, p3, 12, 0, 0);
      } else if (parts[2].length === 4) {
        return new Date(p3, p2 - 1, p1, 12, 0, 0);
      }
    }

    var parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      parsed.setHours(12, 0, 0, 0);
      return parsed;
    }
  }

  return null;
}
