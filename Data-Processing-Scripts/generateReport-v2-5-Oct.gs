/**
 * ============================================================================
 * مكتبة توليد تقارير الدعم الفني الآلية (ReportLibrary) - مع الرسم البياني الشامل
 * ============================================================================
 */

function generateReportFromLibrary(startDateInput, endDateInput) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Report") || ss.getSheets()[0]; 
  
  if (!masterSheet) {
    return { success: false, error: "لم يتم العثور على ورقة البيانات المطلوبة." };
  }

  // ⚙️ [تعديل الإعدادات والأعمدة هنا] ----------------------------------------
  var START_ROW = 37;
  var DATE_COL = 7;
  var STATUS_COL = 12;
  var TASK_TYPE_COL = 10;
  var TASK_CATEGORY_COL = 11;

  // IDs نموذج التقرير ومجلد الحفظ
  const TEMPLATE_ID   = '1ZD6-gNFgNjBZrxSWQxVDa5bvh8rylYDBVhuDaFSzjKk'; // v-7
  // const TEMPLATE_ID   = '1EFRw1teEpAorZUNE6OF89ydSBoIgkX4iqbcSnPEmTlY'; // v-5
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

  // عدادات المتغيرات الـ 7 الجديدة
  var counts = {
    TotalTaskRequest: 0,   // طلبات الخدمة
    TotalTechReq: 0,       // مراسلات مواصفات
    TotalEvaluation: 0,   // مراسلات تحليل
    TotalOthers: 0,      // مراسلات أخرى
    TotalCalls: 0,             // اتصال
    TotalDirectRequests: 0,    // طلب مباشر
    TotalDirectAssignments: 0,  // تكليف مباشر
    TotalOtherRequests: 0      // طلبات أخرى
  };

  // 3. فلترة البيانات وحساب الإحصائيات
  for (var i = START_ROW - 1; i < data.length; i++) {
    var rawDate = data[i][DATE_COL - 1];
    if (!rawDate) continue;

    var rowDateObj = parseSheetDate(rawDate);
    
    if (rowDateObj) {
      var rowTime = rowDateObj.getTime();

      if (rowTime >= startTime && rowTime <= endTime) {
        totalTasks++;
        
        // حساب حالة المهمة
        var status = String(data[i][STATUS_COL - 1] || "").trim().toLowerCase();
        if (status === "مكتملة" || status === "مكتمل" || status === "completed") {
          completedTasks++;

          // تصنيف ونسبة المتغيرات الـ 7 بناءً على عمود نوع الطلب
          var typeVal = String(data[i][TASK_TYPE_COL - 1] || "").trim().toLowerCase();

          if (typeVal.indexOf("طلب خدمة") !== -1) {
            counts.TotalTaskRequest++;
          }
          else if (typeVal.indexOf("اتصال") !== -1) {
            counts.TotalCalls++;
          }
          else if (typeVal.indexOf("طلب مباشر") !== -1) {
            counts.TotalDirectRequests++;
          }
          else if (typeVal.indexOf("تكليف مباشر") !== -1) {
            counts.TotalDirectAssignments++;
          }
          // else if (typeVal.indexOf("أخرى") !== -1) {
          //   counts.TotalOtherRequests++;
          // }
          else if (typeVal.indexOf("مراسلة") !== -1) {
            // counts.TotalCalls++;
            var typeValC = String(data[i][TASK_CATEGORY_COL - 1] || "").trim().toLowerCase();

            if (typeValC.indexOf("المواصفات الفنية") !== -1) {
              counts.TotalTechReq++;
            }
            if (typeValC.indexOf("التحليل والتقييم") !== -1) {
              counts.TotalEvaluation++;
            }
            if (typeValC.indexOf("أخرى") !== -1) {
              counts.TotalOthers++;
            }
          }
          else if (typeVal.indexOf("أخرى") !== -1) {
            counts.TotalOtherRequests++;
          }


          // ← أضف هذا مؤقتاً لترى القيم غير المعروفة
          // else {
          //   if (typeVal !== "") {
          //     Logger.log("قيمة غير مصنفة: '" + typeVal + "'");
          //   }
          // }
        } else {
          pendingTasks++;
        }
      }
    }
  }

  // Logger.log("=== النتائج ===");
  // Logger.log("TotalTaskRequest: "       + counts.TotalTaskRequest);
  // Logger.log("TotalTechReq: "           + counts.TotalTechReq);
  // Logger.log("TotalEvaluation: "        + counts.TotalEvaluation);
  // Logger.log("TotalOthers: "            + counts.TotalOthers);
  // Logger.log("TotalCalls: "             + counts.TotalCalls);
  // Logger.log("TotalDirectRequests: "    + counts.TotalDirectRequests);
  // Logger.log("TotalDirectAssignments: " + counts.TotalDirectAssignments);
  // Logger.log("TotalOtherRequests: "     + counts.TotalOtherRequests);

  // 4. إنشاء المستند واستبدال الوسوم والرسوم البيانية
  try {
    var templateFile = DriveApp.getFileById(TEMPLATE_ID);
    var targetFolder = DriveApp.getFolderById(FOLDER_ID);
    
    var newDocName = "تقرير الدعم الفني - " + monthNameArabic + " " + yearString;
    var newDocFile = templateFile.makeCopy(newDocName, targetFolder);
    var doc = DocumentApp.openById(newDocFile.getId());

    // تجهيز قائمة الوسوم الاستبدالية بما فيها المتغيرات الـ 7
    var replacements = {
      "<<الشهر>>": String(monthNameArabic || ""),
      "<<السنة>>": String(yearString || ""),
      // "<<من_تاريخ>>":         startDateInput,   // ← جديد
      // "<<إلى_تاريخ>>":        endDateInput,      // ← جديد
      "<<من_تاريخ>>":  formatArabicDate(startDateInput),
      "<<إلى_تاريخ>>": formatArabicDate(endDateInput),
      "<<طلبات_الخدمة>>": String(counts.TotalTaskRequest),
      "<<مراسلات_مواصفات>>": String(counts.TotalTechReq),
      "<<مراسلات_تحليل>>": String(counts.TotalEvaluation),
      "<<مراسلات_أخرى>>": String(counts.TotalOthers),
      "<<اتصال>>": String(counts.TotalCalls),
      "<<طلب_مباشر>>": String(counts.TotalDirectRequests),
      "<<تكليف_مباشر>>": String(counts.TotalDirectAssignments),
      "<<طلبات_أخرى>>": String(counts.TotalOtherRequests),
      "<<مجموع_عدد_المهام>>": String(completedTasks)
    };

    // أ) استبدال الوسوم النصية والجداول
    replaceAllTags(doc, replacements);

    // ب) إنشاء الرسم البياني الأفقي (Bar Chart) الشامل وإدراجه
    if (totalTasks > 0) {
      var chartImage = createMultiVariableChart_dnt_all(masterSheet, counts, formatArabicDate(startDateInput), formatArabicDate(endDateInput));
      replaceTagWithImage_dnt_all(doc, "<<الرسم_البياني>>", chartImage);
    } else {
      doc.getBody().replaceText("<<الرسم_البياني>>", "لا توجد بيانات للعرض البياني خلال هذه الفترة");
    }

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
 * إنشاء رسم بياني تنفيذي فاخر (Doughnut أو Radar) بناءً على تفضيلات التصميم
 */

function createMultiVariableChart(sheet, counts, monthName, year) {

  var categories = [
    ["طلب خدمة",    counts.TotalTaskRequest       || 0],
    ["مواصفات",     counts.TotalTechReq           || 0],
    ["تحليل",       counts.TotalEvaluation        || 0],
    ["مراسلات أخرى", counts.TotalOthers           || 0],
    ["طلب مباشر",   counts.TotalDirectRequests    || 0],
    ["طلبات أخرى",   counts.TotalOtherRequests    || 0],
    ["تكليف مباشر", counts.TotalDirectAssignments || 0],
    ["اتصال",        counts.TotalCalls             || 0],
  ];

  categories.sort(function(a, b) { return b[1] - a[1]; });

  var CW = 900;
  var CH = 560;

  _chartW = CW;
  _chartH = CH;

  var title = "توزيع المهام المكتملة" +
    ((monthName && year) ? " · " + monthName + " " + year : "");

  // كتابة البيانات في الشيت مؤقتاً
  var startRow = sheet.getLastRow() + 5;
  var tableData = [["النوع", "العدد"]];
  categories.forEach(function(c) { tableData.push([c[0], c[1]]); });
  sheet.getRange(startRow, 1, tableData.length, 2).setValues(tableData);

  // بناء الرسم وإدراجه في الشيت
  var builder = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(startRow, 1, tableData.length, 2))
    .setPosition(startRow, 4, 0, 0)
    .setOption("title", title)
    .setOption("titleTextStyle", { color: "#0C2340", fontSize: 14, bold: true })
    .setOption("width",  900)
    .setOption("height", 560)
    .setOption("backgroundColor", { fill: "#FFFFFF" })
    .setOption("colors", ["#1A56DB"])
    .setOption("legend", { position: "none" })
    .setOption("chartArea", { left: "25%", top: "12%", width: "70%", height: "76%" })
    .setOption("hAxis", {
      minValue:  0,
      gridlines: { color: "#E2E8F0", count: 5 },
      textStyle: { color: "#64748B", fontSize: 11 }
    })
    .setOption("vAxis", {
      textStyle: { color: "#0F172A", bold: true, fontSize: 11 }
    });

  // ← الفرق الجوهري: أدرج أولاً ثم استخرج من المُدرَج
  sheet.insertChart(builder.build());

  var charts    = sheet.getCharts();
  var inserted  = charts[charts.length - 1];
  var imageBlob = inserted.getAs("image/png");

  // تنظيف الشيت
  sheet.removeChart(inserted);
  sheet.getRange(startRow, 1, tableData.length, 2).clearContent();

  Logger.log("✅ Chart OK");
  return imageBlob;
}

function replaceTagWithImage(doc, tag, imageBlob) {
  if (!imageBlob) {
    doc.getBody().replaceText(tag, "[تعذّر إنشاء الرسم البياني]");
    return;
  }

  var body  = doc.getBody();
  var found = body.findText(tag);
  if (!found) return;

  // الفقرة الحاوية للوسم دائماً
  var para = found.getElement().getParent().asParagraph();

  // استبدل الوسم بمسافة واحدة — لا تفريغ لا حذف
  body.replaceText(tag, " ");
  para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // هل الفقرة داخل خلية جدول؟
  var isInCell = para.getParent().getType() === DocumentApp.ElementType.TABLE_CELL;

  var img = para.appendInlineImage(imageBlob);
  img.setWidth(isInCell ? 320 : 451);
  img.setHeight(isInCell ? 199 : 280);
}

function replaceTagWithImage_dnt(doc, tag, imageBlob) {
  if (!imageBlob) {
    doc.getBody().replaceText(tag, "[تعذّر إنشاء الرسم البياني]");
    return;
  }

  var body  = doc.getBody();
  var found = body.findText(tag);
  if (!found) return;

  // الفقرة الحاوية للوسم دائماً
  var para = found.getElement().getParent().asParagraph();

  // استبدل الوسم بمسافة واحدة — لا تفريغ لا حذف
  body.replaceText(tag, " ");
  para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // هل الفقرة داخل خلية جدول؟
  var isInCell = para.getParent().getType() === DocumentApp.ElementType.TABLE_CELL;

  var img = para.appendInlineImage(imageBlob);
  // الدونات أقرب للمربع — A4 عرض مناسب مع ارتفاع أكبر
  img.setWidth(420);
  img.setHeight(Math.round(420 * _chartH / _chartW)); // 420 × 520/640 = 341
}

function replaceTagWithImage_dnt_all(doc, tag, imageBlob) {
  if (!imageBlob) {
    doc.getBody().replaceText(tag, "[تعذّر إنشاء الرسم البياني]");
    return;
  }

  var body  = doc.getBody();
  var found = body.findText(tag);
  if (!found) return;

  // الفقرة الحاوية للوسم دائماً
  var para = found.getElement().getParent().asParagraph();

  // استبدل الوسم بمسافة واحدة — لا تفريغ لا حذف
  body.replaceText(tag, " ");
  para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // هل الفقرة داخل خلية جدول؟
  var isInCell = para.getParent().getType() === DocumentApp.ElementType.TABLE_CELL;

  var img = para.appendInlineImage(imageBlob);
  // الدونات أقرب للمربع — A4 عرض مناسب مع ارتفاع أكبر
  img.setWidth(451);
  img.setHeight(Math.round(451 * _chartH / _chartW)); // 451 × 700/900 ≈ 351
}

/**
 * دالة استبدال الوسوم الشاملة
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

function createMultiVariableChart_dnt(sheet, counts, monthName, year) {

  // دمج الفئات الصغيرة في "أخرى" للرسم فقط — الأرقام الأصلية في التقرير
  var othersVal = (counts.TotalDirectRequests    || 0)
                + (counts.TotalDirectAssignments || 0)
                + (counts.TotalCalls             || 0);

  var categories = [
    ["طلب خدمة",     counts.TotalTaskRequest || 0],
    ["مواصفات فنية", counts.TotalTechReq     || 0],
    ["تحليل وتقييم", counts.TotalEvaluation  || 0],
    ["مراسلات أخرى", counts.TotalOthers      || 0],
    ["طلب/تكليف/اتصال", othersVal],
  ];

  // احذف الفئات الصفر لتجنب شرائح فارغة
  categories = categories.filter(function(c) { return c[1] > 0; });
  categories.sort(function(a, b) { return b[1] - a[1]; });

  var CW = 640;
  var CH = 520;

  var title = "توزيع المهام المكتملة" +
    ((monthName && year) ? " · " + monthName + " " + year : "");

  var startRow = sheet.getLastRow() + 5;
  sheet.getRange(startRow, 1, categories.length, 2).setValues(categories);

  var colors = ["#1A56DB", "#0C2340", "#059669", "#7C3AED", "#D97706",
                "#60A5FA", "#EC4899", "#0D9488"];

  var builder = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(startRow, 1, categories.length, 2))
    .setPosition(startRow, 4, 0, 0)
    .setOption("title",         title)
    .setOption("titleTextStyle", { color: "#0C2340", fontSize: 14, bold: true })
    .setOption("width",          CW)
    .setOption("height",         CH)
    .setOption("backgroundColor", { fill: "#FFFFFF" })
    .setOption("colors",          colors)
    .setOption("pieHole",         0.45)   // ← هذا يحوّله لدونات
    .setOption("legend", {
      position:  "right",
      textStyle: { color: "#334155", fontSize: 12, bold: true }
    })
    .setOption("pieSliceText", "percentage") // نسبة % على كل شريحة
    .setOption("pieSliceTextStyle", {
      color:    "#FFFFFF",
      fontSize: 12,
      bold:     true
    })
    .setOption("chartArea", { left: "5%", top: "15%", width: "60%", height: "75%" });

  sheet.insertChart(builder.build());
  var charts   = sheet.getCharts();
  var inserted = charts[charts.length - 1];
  var blob     = inserted.getAs("image/png");

  sheet.removeChart(inserted);
  sheet.getRange(startRow, 1, categories.length, 2).clearContent();

  // حدّث الأبعاد للإدراج في الدوكس
  _chartW = CW;
  _chartH = CH;

  Logger.log("✅ Donut chart OK");
  return blob;
}

function createMultiVariableChart_dnt_all(sheet, counts, monthName, year) {

  var raw = [
    ["طلبات خدمة",     counts.TotalTaskRequest || 0],
    ["مراسلات - مواصفات فنية", counts.TotalTechReq || 0],
    ["مراسلات - تحليل وتقييم", counts.TotalEvaluation || 0],
    ["مراسلات أخرى", counts.TotalOthers || 0],
    ["طلبات مباشر",    counts.TotalDirectRequests || 0],
    ["تكليف مباشر",  counts.TotalDirectAssignments || 0],
    ["اتصالات",         counts.TotalCalls || 0],
    ["طلبات أخرى",         counts.TotalOtherRequests || 0],
  ];

  // احذف الفئات الصفر
  var filtered = raw.filter(function(c) { return c[1] > 0; });
  filtered.sort(function(a, b) { return b[1] - a[1]; });

  // احسب المجموع ثم ادمج النسبة في الاسم ← تظهر في legend
  var total = filtered.reduce(function(sum, c) { return sum + c[1]; }, 0);
  var categories = filtered.map(function(c) {
    var pct = total > 0 ? Math.round(c[1] * 100 / total) : 0;
    return [c[0] + "   " + pct + "%", c[1]];
  });

  // var CW = 700;
  // var CH = 540;
  var CW = 900;
  var CH = 700;
  _chartW = CW;
  _chartH = CH;

  var title = "توزيع المهام المكتملة خلال الفترة من " + monthName + " إلى " + year;

  var startRow = sheet.getLastRow() + 5;
  sheet.getRange(startRow, 1, categories.length, 2).setValues(categories);

  var colors = [
    "#1A56DB", // أزرق
    "#0C2340", // كحلي
    "#059669", // أخضر
    "#7C3AED", // بنفسجي
    "#D97706", // عنبري
    "#60A5FA", // أزرق فاتح
    "#EC4899", // وردي
    "#0D9488", // تيل
  ];

  var builder = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(startRow, 1, categories.length, 2))
    .setPosition(startRow, 4, 0, 0)
    .setOption("title", title)
    .setOption("titleTextStyle", { color: "#0C2340", fontSize: 14, bold: true })
    .setOption("width", CW)
    .setOption("height", CH)
    .setOption("backgroundColor", { fill: "#FFFFFF" })
    .setOption("colors", colors)
    .setOption("pieHole", 0.45)
    // .setOption("pieSliceText", "none") // لا نص على الشرائح — النسبة في legend فقط
    .setOption("pieSliceText", "percentage")
    .setOption("pieSliceTextStyle", {
      color:    "#FFFFFF",
      fontSize: 20,
      bold:     true
    })
    .setOption("legend", {
      position:  "right",
      textStyle: { color: "#334155", fontSize: 25 }
    })
    .setOption("chartArea", { left: "5%", top: "15%", width: "58%", height: "76%" });

  sheet.insertChart(builder.build());
  var charts   = sheet.getCharts();
  var inserted = charts[charts.length - 1];
  var blob     = inserted.getAs("image/png");

  sheet.removeChart(inserted);
  sheet.getRange(startRow, 1, categories.length, 2).clearContent();

  Logger.log("✅ Donut OK — " + categories.length + " فئات");
  return blob;
}

function formatArabicDate(dateStr) {
  var arabicMonths = [
    "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
    "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
  ];
  var parts = dateStr.split("-");
  var day   = parseInt(parts[2], 10);
  var month = arabicMonths[parseInt(parts[1], 10) - 1];
  var year  = parts[0];
  return day + " " + month + " " + year;
}
