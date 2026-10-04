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
  const TEMPLATE_ID   = '1nITCSSbp9j6Ktfsctg8x-aeRJEr4UBdQ15vM5ZHTfVI'; // v-6
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
    TotalDirectAssignments: 0  // تكليف مباشر
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
        } else {
          pendingTasks++;
        }

        // تصنيف ونسبة المتغيرات الـ 7 بناءً على عمود نوع الطلب
        var typeVal = String(data[i][TASK_TYPE_COL - 1] || "").trim().toLowerCase();

        if (typeVal.indexOf("خدمة") !== -1 || typeVal.indexOf("طلبات الخدمة") !== -1) {
          counts.TotalTaskRequest++;
        } else if (typeVal.indexOf("مواصفات") !== -1) {
          counts.TotalTechReq++;
        } else if (typeVal.indexOf("تحليل") !== -1) {
          counts.TotalEvaluation++;
        } else if (typeVal.indexOf("مراسلات") !== -1 || typeVal.indexOf("أخرى") !== -1) {
          counts.TotalOthers++;
        } else if (typeVal.indexOf("اتصال") !== -1 || typeVal.indexOf("هاتف") !== -1) {
          counts.TotalCalls++;
        } else if (typeVal.indexOf("طلب مباشر") !== -1) {
          counts.TotalDirectRequests++;
        } else if (typeVal.indexOf("تكليف") !== -1 || typeVal.indexOf("تكليف مباشر") !== -1) {
          counts.TotalDirectAssignments++;
        }
      }
    }
  }

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
      "<<طلبات_الخدمة>>": String(counts.TotalTaskRequest),
      "<<مراسلات_مواصفات>>": String(counts.TotalTechReq),
      "<<مراسلات_تحليل>>": String(counts.TotalEvaluation),
      "<<مراسلات_أخرى>>": String(counts.TotalOthers),
      "<<اتصال>>": String(counts.TotalCalls),
      "<<طلب_مباشر>>": String(counts.TotalDirectRequests),
      "<<تكليف_مباشر>>": String(counts.TotalDirectAssignments)
    };

    // أ) استبدال الوسوم النصية والجداول
    replaceAllTags(doc, replacements);

    // ب) إنشاء الرسم البياني الأفقي (Bar Chart) الشامل وإدراجه
    // if (totalTasks > 0) {
    //   // var chartImage = createMultiVariableChart(masterSheet, counts);
    //   var chartImage = createMultiVariableChart(masterSheet, counts, monthNameArabic, yearString);
    //   replaceTagWithImage(doc, "<<الرسم_البياني>>", chartImage);
    // } else {
    //   doc.getBody().replaceText("<<الرسم_البياني>>", "لا توجد بيانات للعرض البياني خلال هذه الفترة");
    // }

    // ب) إنشاء الرسم البياني الأفقي (Bar Chart) الشامل وإدراجه
    if (totalTasks > 0) {
      var chartResult = createMultiVariableChart(masterSheet, counts, monthNameArabic, yearString);
      replaceTagWithImage(doc, "<<الرسم_البياني>>", chartResult);
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
 * دالة إنشاء رسم بياني أفقِي/أعمدة (Bar Chart) للمتغيرات الـ 7
 */
// function createMultiVariableChart(sheet, counts) {
//   var tempRow = sheet.getLastRow() + 5;
  
//   // تجهيز جدول البيانات المؤقت
//   var chartData = [
//     ["نوع الطلب", "العدد"],
//     ["طلبات الخدمة", counts.TotalTaskRequest],
//     ["مراسلات مواصفات", counts.TotalTechReq],
//     ["مراسلات تحليل", counts.TotalEvaluation],
//     ["مراسلات أخرى", counts.TotalOthers],
//     ["اتصال", counts.TotalCalls],
//     ["طلب مباشر", counts.TotalDirectRequests],
//     ["تكليف مباشر", counts.TotalDirectAssignments]
//   ];

//   sheet.getRange(tempRow, 1, 8, 2).setValues(chartData);
//   var dataRange = sheet.getRange(tempRow, 1, 8, 2);
  
//   // بناء رسم بياني أفقِي منظم ومناسب لعدد الفئات (7 فئات)
//   var chartBuilder = sheet.newChart()
//     .setChartType(Charts.ChartType.BAR) // رسم أفقِي لتظهر أسماء الفئات بوضوح
//     .addRange(dataRange)
//     .setPosition(tempRow, 4, 0, 0)
//     .setOption('title', 'توزيع المهام حسب نوع الطلب والخدمة')
//     .setOption('width', 550)
//     .setOption('height', 350)
//     .setOption('colors', ['#1A365D']) // لون أزرق رسمي كحلي
//     .setOption('legend', {position: 'none'}); // إخفاء دليل الألوان للتبسيط

//   var chart = chartBuilder.build();
//   sheet.insertChart(chart);

//   // استخراج الصورة وتنظيف الشيت
//   var imageBlob = chart.getAs('image/png');
//   sheet.removeChart(chart);
//   sheet.getRange(tempRow, 1, 8, 2).clear();

//   return imageBlob;
// }

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

  // عدادات المتغيرات الـ 7 الجديدة
  var counts = {
    TotalTaskRequest: 0,   // طلبات الخدمة
    TotalTechReq: 0,       // مراسلات مواصفات
    TotalEvaluation: 0,   // مراسلات تحليل
    TotalOthers: 0,      // مراسلات أخرى
    TotalCalls: 0,             // اتصال
    TotalDirectRequests: 0,    // طلب مباشر
    TotalDirectAssignments: 0  // تكليف مباشر
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
          // else if (typeVal.indexOf("مراسلات مواصفات") !== -1) {
          //   counts.TotalTechReq++;
          // }
          // else if (typeVal.indexOf("مراسلات تحليل") !== -1) {
          //   counts.TotalEvaluation++;
          // }
          // else if (typeVal.indexOf("مراسلات مراسلات") !== -1 || typeVal.indexOf("أخرى") !== -1) {
          //   counts.TotalOthers++;
          // }
          else if (typeVal.indexOf("اتصال") !== -1) {
            counts.TotalCalls++;
          }
          else if (typeVal.indexOf("طلب مباشر") !== -1) {
            counts.TotalDirectRequests++;
          }
          else if (typeVal.indexOf("تكليف مباشر") !== -1) {
            counts.TotalDirectAssignments++;
          }
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
        } else {
          pendingTasks++;
        }
      }
    }
  }

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
      "<<طلبات_الخدمة>>": String(counts.TotalTaskRequest),
      "<<مراسلات_مواصفات>>": String(counts.TotalTechReq),
      "<<مراسلات_تحليل>>": String(counts.TotalEvaluation),
      "<<مراسلات_أخرى>>": String(counts.TotalOthers),
      "<<اتصال>>": String(counts.TotalCalls),
      "<<طلب_مباشر>>": String(counts.TotalDirectRequests),
      "<<تكليف_مباشر>>": String(counts.TotalDirectAssignments),
      "<<مجموع_عدد_المهام>>": String(completedTasks)
    };

    // أ) استبدال الوسوم النصية والجداول
    replaceAllTags(doc, replacements);

    // ب) إنشاء الرسم البياني الأفقي (Bar Chart) الشامل وإدراجه
    if (totalTasks > 0) {
      // var chartImage = createMultiVariableChart(masterSheet, counts);
      var chartImage = createMultiVariableChart(masterSheet, counts, monthNameArabic, yearString);
      replaceTagWithImage(doc, "<<الرسم_البياني>>", chartImage);
    } else {
      doc.getBody().replaceText("<<الرسم_البياني>>", "لا توجد بيانات للعرض البياني خلال هذه الفترة");
    }

    // 
    if (totalTasks > 0) {
      var chartImage = createMultiVariableChart(masterSheet, counts, monthNameArabic, yearString);
      replaceTagWithImage(doc, "<<الرسم_البياني>>", chartImage);
    }
    // 

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
 * دالة إنشاء رسم بياني احترافي، أنيق، ومريح للعين (3D Column Chart)
 */
// function createMultiVariableChart(sheet, counts) {
//   var tempRow = sheet.getLastRow() + 5;
  
//   // 1. إعداد البيانات المؤقتة
//   var chartData = [
//     ["نوع الطلب", "العدد"],
//     ["طلبات الخدمة", counts.TotalTaskRequest],
//     ["مراسلات مواصفات", counts.TotalTechReq],
//     ["مراسلات تحليل", counts.TotalEvaluation],
//     ["مراسلات أخرى", counts.TotalOthers],
//     ["اتصال", counts.TotalCalls],
//     ["طلب مباشر", counts.TotalDirectRequests],
//     ["تكليف مباشر", counts.TotalDirectAssignments]
//   ];

//   sheet.getRange(tempRow, 1, 8, 2).setValues(chartData);
//   var dataRange = sheet.getRange(tempRow, 1, 8, 2);
  
//   // 2. بناء الرسم البياني بتنسيقات بصرية راقية
//   var chartBuilder = sheet.newChart()
//     .setChartType(Charts.ChartType.COLUMN) // أعمدة رأسية حديثة
//     .addRange(dataRange)
//     .setPosition(tempRow, 4, 0, 0)
//     // العناوين والنصوص
//     .setOption('title', 'توزيع المهام والطلبات حسب الفئة')
//     .setOption('titleTextStyle', {
//       color: '#1A365D',
//       fontSize: 16,
//       bold: true
//     })
//     // الأبعاد وخلفية الرسم
//     .setOption('width', 600)
//     .setOption('height', 360)
//     .setOption('backgroundColor', '#F7FAFC') // خلفية رمادية هادئة جداً
//     .setOption('colors', ['#2B6CB0'])        // لون أزرق إلكتروني أنيق
//     .setOption('is3D', true)                 // لمسة تجسيم 3D هادئة
//     .setOption('legend', { position: 'none' }) // إخفاء دليل الألوان غير الضروي
//     // تحسين المحاور والأرقام
//     .setOption('vAxis', {
//       title: 'عدد الطلبات',
//       titleTextStyle: { color: '#4A5568', bold: true, fontSize: 12 },
//       gridlines: { color: '#E2E8F0', count: 5 }, // خطوط شبكة خفيفة جداً
//       baselineColor: '#CBD5E0'
//     })
//     .setOption('hAxis', {
//       slantedText: true,           // إمالة النص إذا كانت الأسماء طويلة لمنع التداخل
//       slantedTextAngle: 30,
//       textStyle: { color: '#2D3748', fontSize: 11, bold: true }
//     });

//   var chart = chartBuilder.build();
//   sheet.insertChart(chart);

//   // 3. استخراج الصورة عالية الدقة وتنظيف المكان
//   var imageBlob = chart.getAs('image/png');
//   sheet.removeChart(chart);
//   sheet.getRange(tempRow, 1, 8, 2).clear();

//   return imageBlob;
// }

/**
 * إنشاء رسم بياني احترافي للغاية ومريح للعين باستخدام QuickChart API
 */
// function createMultiVariableChart(sheet, counts) {
//   // 1. تجهيز التسميات والبيانات
//   var labels = [
//     "طلبات الخدمة",
//     "مراسلات مواصفات",
//     "مراسلات تحليل",
//     "مراسلات أخرى",
//     "اتصال",
//     "طلب مباشر",
//     "تكليف مباشر"
//   ];
  
//   var dataValues = [
//     counts.TotalTaskRequest || 0,
//     counts.TotalTechReq || 0,
//     counts.TotalEvaluation || 0,
//     counts.TotalOthers || 0,
//     counts.TotalCalls || 0,
//     counts.TotalDirectRequests || 0,
//     counts.TotalDirectAssignments || 0
//   ];

//   // 2. إعداد إعدادات Chart.js الاحترافية
//   var chartConfig = {
//     type: 'bar',
//     data: {
//       labels: labels,
//       datasets: [{
//         label: 'عدد الطلبات',
//         data: dataValues,
//         backgroundColor: 'rgba(54, 99, 230, 0.85)', // أزرق عصري ومريح
//         borderColor: '#2563EB',
//         borderWidth: 1.5,
//         borderRadius: 8, // حواف دائرية أنيقة للأشرطة
//         borderSkipped: false
//       }]
//     },
//     options: {
//       responsive: true,
//       plugins: {
//         title: {
//           display: true,
//           text: 'توزيع المهام والطلبات حسب الفئة',
//           color: '#1E293B',
//           font: { size: 18, weight: 'bold', family: 'sans-serif' },
//           padding: { bottom: 20 }
//         },
//         legend: { display: false }, // إخفاء الدليل لعدم الحاجة له
//         datalabels: { // عرض الأرقام فوق كل عمود
//           anchor: 'end',
//           align: 'top',
//           color: '#334155',
//           font: { weight: 'bold', size: 12 }
//         }
//       },
//       scales: {
//         x: {
//           grid: { display: false }, // إخفاء خطوط الشبكة العمودية
//           ticks: {
//             color: '#475569',
//             font: { size: 11, weight: '600' }
//           }
//         },
//         y: {
//           beginAtZero: true,
//           grid: { color: '#F1F5F9' }, // خطوط أفقية خفيفة جداً
//           ticks: {
//             color: '#64748B',
//             precision: 0
//           }
//         }
//       }
//     }
//   };

//   // 3. جلب الصورة عالية الدقة عبر طلب HTTP
//   var url = 'https://quickchart.io/chart?w=650&h=380&devicePixelRatio=2&bkg=white&c=' + encodeURIComponent(JSON.stringify(chartConfig));
  
//   try {
//     var response = UrlFetchApp.fetch(url);
//     return response.getBlob().setName("chart.png");
//   } catch (e) {
//     Logger.log("خطأ في جلب الرسم البياني: " + e.toString());
//     return null;
//   }
// }

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
    ["تكليف مباشر", counts.TotalDirectAssignments || 0],
    ["اتصال",        counts.TotalCalls             || 0],
  ];

  categories.sort(function(a, b) { return b[1] - a[1]; });

  _chartW = 600;
  _chartH = 400;

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
    .setOption("width",  600)
    .setOption("height", 400)
    .setOption("backgroundColor", { fill: "#FFFFFF" })
    .setOption("colors", ["#1A56DB"])
    .setOption("legend", { position: "none" })
    .setOption("chartArea", { left: "28%", top: "16%", width: "65%", height: "72%" })
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

// function createMultiVariableChart(sheet, counts, monthName, year) {

//   var categories = [
//     ["طلب خدمة",     counts.TotalTaskRequest       || 0],
//     ["مواصفات فنية", counts.TotalTechReq           || 0],
//     ["تحليل وتقييم", counts.TotalEvaluation        || 0],
//     ["مراسلات أخرى", counts.TotalOthers            || 0],
//     ["طلب مباشر",    counts.TotalDirectRequests    || 0],
//     ["تكليف مباشر",  counts.TotalDirectAssignments || 0],
//     ["اتصال",         counts.TotalCalls             || 0],
//   ];

//   // ترتيب تنازلي
//   categories.sort(function(a, b) { return b[1] - a[1]; });

//   var CW = 600;
//   var CH = 400;
//   _chartW = CW;
//   _chartH = CH;

//   var title2 = (monthName && year) ? (monthName + " " + year) : "";

//   // كتابة البيانات في صفوف مؤقتة أسفل الشيت
//   var tempRow = sheet.getLastRow() + 5;
//   var chartData = [["نوع المهمة", "العدد"]].concat(categories);
//   sheet.getRange(tempRow, 1, chartData.length, 2).setValues(chartData);
//   var dataRange = sheet.getRange(tempRow, 1, chartData.length, 2);

//   // بناء الرسم البياني داخل Sheets — يرسم العربي بشكل صحيح
//   var chart = sheet.newChart()
//     .setChartType(Charts.ChartType.BAR)
//     .addRange(dataRange)
//     .setPosition(tempRow, 4, 0, 0)
//     .setOption("title", "توزيع المهام المكتملة" + (title2 ? " · " + title2 : ""))
//     .setOption("titleTextStyle", { color: "#0C2340", fontSize: 15, bold: true })
//     .setOption("width",  CW)
//     .setOption("height", CH)
//     .setOption("backgroundColor", { fill: "#FFFFFF" })
//     .setOption("colors", ["#1A56DB"])
//     .setOption("legend", { position: "none" })
//     .setOption("chartArea", { left: "150", top: 60, width: "400", height: "320" })
//     .setOption("hAxis", {
//       minValue:  0,
//       gridlines: { color: "#E2E8F0", count: 5 },
//       textStyle: { color: "#64748B", fontSize: 11 }
//     })
//     .setOption("vAxis", {
//       textStyle: { color: "#0F172A", bold: true, fontSize: 12 }
//     })
//     .build();

//   sheet.insertChart(chart);
//   var imageBlob = chart.getAs("image/png");

//   // تنظيف: احذف الرسم والبيانات المؤقتة
//   sheet.removeChart(chart);
//   sheet.getRange(tempRow, 1, chartData.length, 2).clear();

//   Logger.log("✅ Sheets chart OK — " + CW + "×" + CH);
//   return imageBlob;
// }

// function createMultiVariableChart(sheet, counts, monthName, year) {

//   // var categories = [
//   //   { label: "طلب خدمة",         value: counts.TotalTaskRequest       || 0, color: "#2563EB" },
//   //   { label: "مراسلة · مواصفات", value: counts.TotalTechReq           || 0, color: "#0C2340" },
//   //   { label: "مراسلة · تحليل",   value: counts.TotalEvaluation        || 0, color: "#1A56DB" },
//   //   { label: "مراسلة · أخرى",    value: counts.TotalOthers            || 0, color: "#60A5FA" },
//   //   { label: "طلب مباشر",        value: counts.TotalDirectRequests    || 0, color: "#059669" },
//   //   { label: "تكليف مباشر",      value: counts.TotalDirectAssignments || 0, color: "#7C3AED" },
//   //   { label: "اتصال",             value: counts.TotalCalls             || 0, color: "#D97706" },
//   // ];

//   // ── التسميات المختصرة (تحل مشكلة القطع والانكسار) ──
//   var categories = [
//     { label: "طلب خدمة",     value: counts.TotalTaskRequest       || 0, color: "#2563EB" },
//     { label: "مواصفات فنية", value: counts.TotalTechReq           || 0, color: "#0C2340" },
//     { label: "تحليل وتقييم", value: counts.TotalEvaluation        || 0, color: "#1A56DB" },
//     { label: "مراسلات أخرى", value: counts.TotalOthers            || 0, color: "#60A5FA" },
//     { label: "طلب مباشر",    value: counts.TotalDirectRequests    || 0, color: "#059669" },
//     { label: "تكليف مباشر",  value: counts.TotalDirectAssignments || 0, color: "#7C3AED" },
//     { label: "اتصال",         value: counts.TotalCalls             || 0, color: "#D97706" },
//   ];

//   categories.sort(function(a, b) { return b.value - a.value; });

//   // var CW = 445;
//   // var CH = Math.max(320, categories.length * 42 + 120);

//   // احفظ الأبعاد للدالة الثانية
//   // _chartW = CW;
//   // _chartH = CH;

//   // ── أبعاد الرسم: أوسع قليلاً لإعطاء المحور مساحة ──
//   var CW = 520;
//   var CH = Math.max(340, categories.length * 44 + 130);
//   _chartW = CW;
//   _chartH = CH;

//   var title2 = (monthName && year) ? (monthName + " " + year) : "";

//   var cfg = {
//     type: "bar",
//     data: {
//       labels: categories.map(function(c) { return c.label; }),
//       datasets: [{
//         data:            categories.map(function(c) { return c.value; }),
//         backgroundColor: categories.map(function(c) { return c.color; }),
//         borderRadius:    5,
//         borderSkipped:   false,
//         barPercentage:   0.68
//       }]
//     },
//     options: {
//       indexAxis: "y",
//       plugins: {
//         legend: { display: false },
//         // title: {
//         //   display: true,
//         //   text: title2
//         //     ? ["توزيع المهام المكتملة حسب النوع", title2]
//         //     : ["توزيع المهام المكتملة حسب النوع"],
//         //   color: "#0C2340",
//         //   font:  { size: 14, weight: "bold", family: "Arial" },
//         //   padding: { bottom: 14 }
//         // },
//         // ── عنوان يوضّح أن مواصفات/تحليل/أخرى هي فئة المراسلات ──
//         title: {
//           display: true,
//           text: title2
//             ? ["توزيع المهام المكتملة حسب النوع  |  مواصفات · تحليل · أخرى = مراسلات", title2]
//             : ["توزيع المهام المكتملة حسب النوع  |  مواصفات · تحليل · أخرى = مراسلات"],
//           color: "#0C2340",
//           font:  { size: 13, weight: "bold", family: "Arial" },
//           padding: { bottom: 14 }
//         },
//         datalabels: {
//           display:   true,
//           anchor:    "end",
//           align:     "end",
//           clamp:     true,
//           color:     "#0C2340",
//           font:      { weight: "bold", size: 11, family: "Arial" },
//           formatter: function(val) { return val > 0 ? val : ""; }
//         }
//       },
//       scales: {
//         x: {
//           beginAtZero: true,
//           grid:  { color: "#E2E8F0" },
//           ticks: { color: "#64748B", font: { size: 10 } }
//         },
//         // y: {
//         //   grid:  { display: false },
//         //   ticks: { color: "#0F172A", font: { weight: "bold", size: 11 } }
//         // }
//         // ── المحور Y: حجم خط أصغر + منع الكسر ──
//         y: {
//           grid: { display: false },
//           ticks: {
//             color: "#0F172A",
//             font:  { weight: "bold", size: 10, family: "Arial" },
//             maxRotation: 0,
//             minRotation: 0,
//             autoSkip: false,
//           }
//         },
//       },
//       layout: { padding: { right: 45, top: 4, bottom: 6, left: 2 } }
//     }
//   };

//   // var url = "https://quickchart.io/chart"
//   //   + "?v=3&w=" + CW + "&h=" + CH
//   //   + "&devicePixelRatio=2&bkg=%23FFFFFF"
//   //   + "&c=" + encodeURIComponent(JSON.stringify(cfg));

//   // ── رابط QuickChart بنفس العرض الجديد ──
//   var url = "https://quickchart.io/chart"
//     + "?v=3&w=" + CW + "&h=" + CH
//     + "&devicePixelRatio=2&bkg=%23FFFFFF"
//     + "&c=" + encodeURIComponent(JSON.stringify(cfg));

//   try {
//     var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
//     if (res.getResponseCode() === 200) {
//       Logger.log("✅ QuickChart OK");
//       return res.getBlob();
//     }
//     Logger.log("QuickChart HTTP: " + res.getResponseCode());
//   } catch (e) {
//     Logger.log("QuickChart error: " + e.message);
//   }

//   // Fallback
//   try {
//     var dt = Charts.newDataTable()
//       .addColumn(Charts.ColumnType.STRING, "النوع")
//       .addColumn(Charts.ColumnType.NUMBER, "العدد");
//     categories.forEach(function(c) { dt.addRow([c.label, c.value]); });

//     return Charts.newBarChart()
//       .setDataTable(dt)
//       .setTitle("توزيع المهام · " + (title2 || ""))
//       .setOption("titleTextStyle",  { fontSize: 13, bold: true, color: "#0C2340" })
//       .setOption("backgroundColor", { fill: "#FFFFFF" })
//       .setOption("colors",          ["#1A56DB"])
//       .setOption("legend",          { position: "none" })
//       .setOption("chartArea",       { left: 30, top: 55, width: "68%", height: "82%" })
//       .setOption("hAxis", { minValue: 0, gridlines: { color: "#E2E8F0" }, textStyle: { color: "#64748B" } })
//       .setOption("vAxis", { textStyle: { color: "#0F172A", bold: true } })
//       .setDimensions(CW, CH)
//       .build()
//       .getAs("image/png");
//   } catch (e2) {
//     Logger.log("Fallback error: " + e2.message);
//     return null;
//   }
// }

// function createMultiVariableChart(sheet, counts, monthName, year) {

//   var categories = [
//     { label: "طلب خدمة",         value: counts.TotalTaskRequest       || 0, color: "#2563EB" },
//     { label: "مراسلة · مواصفات", value: counts.TotalTechReq           || 0, color: "#0C2340" },
//     { label: "مراسلة · تحليل",   value: counts.TotalEvaluation        || 0, color: "#1A56DB" },
//     { label: "مراسلة · أخرى",    value: counts.TotalOthers            || 0, color: "#60A5FA" },
//     { label: "طلب مباشر",        value: counts.TotalDirectRequests    || 0, color: "#059669" },
//     { label: "تكليف مباشر",      value: counts.TotalDirectAssignments || 0, color: "#7C3AED" },
//     { label: "اتصال",             value: counts.TotalCalls             || 0, color: "#D97706" },
//   ];

//   categories.sort(function(a, b) { return b.value - a.value; });

//   var labels = categories.map(function(c) { return c.label; });
//   var values = categories.map(function(c) { return c.value; });
//   var colors = categories.map(function(c) { return c.color; });
//   var title2 = (monthName && year) ? (monthName + " " + year) : "";

//   // ── المفتاح: عرض الرسم = عرض العرض في الدوكس (445px ≈ 440pt)
//   // → النصوص بحجمها الطبيعي، لا ضغط ولا تكبير
//   var CHART_W = 445;
//   var CHART_H = Math.max(320, categories.length * 42 + 120);

//   var chartConfig = {
//     type: "bar",
//     data: {
//       labels: labels,
//       datasets: [{
//         data:            values,
//         backgroundColor: colors,
//         borderRadius:    5,
//         borderSkipped:   false,
//         barPercentage:   0.68
//       }]
//     },
//     options: {
//       indexAxis: "y",
//       plugins: {
//         legend: { display: false },
//         title: {
//           display: true,
//           text: title2
//             ? ["توزيع المهام المكتملة حسب النوع", title2]
//             : ["توزيع المهام المكتملة حسب النوع"],
//           color: "#0C2340",
//           font:  { size: 14, weight: "bold", family: "Arial" },
//           padding: { bottom: 14 }
//         },
//         datalabels: {
//           display:   true,
//           anchor:    "end",
//           align:     "end",
//           clamp:     true,
//           color:     "#0C2340",
//           font:      { weight: "bold", size: 11, family: "Arial" },
//           formatter: function(val) { return val > 0 ? val : ""; }
//         }
//       },
//       scales: {
//         x: {
//           beginAtZero: true,
//           grid:  { color: "#E2E8F0", lineWidth: 1 },
//           ticks: { color: "#64748B", font: { size: 10 } }
//         },
//         y: {
//           grid:  { display: false },
//           ticks: { color: "#0F172A", font: { weight: "bold", size: 11 } }
//         }
//       },
//       layout: { padding: { right: 45, top: 4, bottom: 6, left: 2 } }
//     }
//   };

//   // devicePixelRatio=2 → صورة 890×2H فعلية (حادة)
//   // عرض العرض في الدوكس (440pt) ≈ عرض الرسم المنطقي (445px) → 1:1
//   var url = "https://quickchart.io/chart"
//     + "?v=3"
//     + "&w=" + CHART_W
//     + "&h=" + CHART_H
//     + "&devicePixelRatio=2"
//     + "&bkg=%23FFFFFF"
//     + "&c=" + encodeURIComponent(JSON.stringify(chartConfig));

//   try {
//     var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
//     if (res.getResponseCode() === 200) {
//       Logger.log("✅ QuickChart OK — " + CHART_W + "×" + CHART_H);
//       // أعد الصورة مع الأبعاد المنطقية
//       return {
//         blob:   res.getBlob().setName("chart.png"),
//         width:  CHART_W,
//         height: CHART_H
//       };
//     }
//     Logger.log("QuickChart HTTP " + res.getResponseCode());
//   } catch (e) {
//     Logger.log("QuickChart: " + e.message);
//   }

//   // Fallback: Charts API
//   try {
//     var dt = Charts.newDataTable()
//       .addColumn(Charts.ColumnType.STRING, "النوع")
//       .addColumn(Charts.ColumnType.NUMBER, "العدد");
//     categories.forEach(function(c) { dt.addRow([c.label, c.value]); });

//     var blob = Charts.newBarChart()
//       .setDataTable(dt)
//       .setTitle("توزيع المهام · " + (title2 || ""))
//       .setOption("titleTextStyle",  { fontSize: 13, bold: true, color: "#0C2340" })
//       .setOption("backgroundColor", { fill: "#FFFFFF" })
//       .setOption("colors",          ["#1A56DB"])
//       .setOption("legend",          { position: "none" })
//       .setOption("chartArea",       { left: 30, top: 55, width: "68%", height: "82%" })
//       .setOption("hAxis", { minValue: 0, gridlines: { color: "#E2E8F0" }, textStyle: { color: "#64748B" } })
//       .setOption("vAxis", { textStyle: { color: "#0F172A", bold: true } })
//       .setDimensions(CHART_W * 2, CHART_H * 2)
//       .build()
//       .getAs("image/png");

//     return { blob: blob, width: CHART_W, height: CHART_H };
//   } catch (e2) {
//     Logger.log("Charts API: " + e2.message);
//     return null;
//   }
// }

// function createMultiVariableChart(sheet, counts, monthName, year) {

//   var categories = [
//     { label: "طلب خدمة",         value: counts.TotalTaskRequest       || 0, color: "#2563EB" },
//     { label: "مراسلة · مواصفات", value: counts.TotalTechReq           || 0, color: "#0C2340" },
//     { label: "مراسلة · تحليل",   value: counts.TotalEvaluation        || 0, color: "#1A56DB" },
//     { label: "مراسلة · أخرى",    value: counts.TotalOthers            || 0, color: "#60A5FA" },
//     { label: "طلب مباشر",        value: counts.TotalDirectRequests    || 0, color: "#059669" },
//     { label: "تكليف مباشر",      value: counts.TotalDirectAssignments || 0, color: "#7C3AED" },
//     { label: "اتصال",             value: counts.TotalCalls             || 0, color: "#D97706" },
//   ];

//   categories.sort(function(a, b) { return b.value - a.value; });

//   var labels = categories.map(function(c) { return c.label; });
//   var values = categories.map(function(c) { return c.value; });
//   var colors = categories.map(function(c) { return c.color; });
//   var title2 = (monthName && year) ? (monthName + " " + year) : "";

//   // الارتفاع مضبوط ليناسب صفحة A4
//   var chartHeight = (categories.length * 40) + 130;

//   var chartConfig = {
//     type: "bar",
//     data: {
//       labels: labels,
//       datasets: [{
//         data:            values,
//         backgroundColor: colors,
//         borderRadius:    6,
//         borderSkipped:   false,
//         barPercentage:   0.70
//       }]
//     },
//     options: {
//       indexAxis: "y",
//       plugins: {
//         legend: { display: false },
//         title: {
//           display: true,
//           text: title2
//             ? ["توزيع المهام المكتملة حسب النوع", title2]
//             : ["توزيع المهام المكتملة حسب النوع"],
//           color: "#0C2340",
//           font: { size: 15, weight: "bold", family: "Arial" },
//           padding: { bottom: 18 }
//         },
//         datalabels: {
//           display:   true,
//           anchor:    "end",
//           align:     "end",
//           clamp:     true,
//           color:     "#0C2340",
//           font:      { weight: "bold", size: 12, family: "Arial" },
//           formatter: function(val) { return val > 0 ? val : ""; }
//         }
//       },
//       scales: {
//         x: {
//           beginAtZero: true,
//           grid:  { color: "#E2E8F0", lineWidth: 1 },
//           ticks: { color: "#64748B", font: { size: 11 } }
//         },
//         y: {
//           grid:  { display: false },
//           ticks: { color: "#0F172A", font: { weight: "bold", size: 11 } }
//         }
//       },
//       layout: { padding: { right: 55, top: 6, bottom: 8, left: 4 } }
//     }
//   };

//   // w=800 * devicePixelRatio=2 → صورة 1600px فعلية (حادة في الطباعة)
//   var url = "https://quickchart.io/chart"
//     + "?v=3"
//     + "&w=800"
//     + "&h=" + chartHeight
//     + "&devicePixelRatio=2"
//     + "&bkg=%23FFFFFF"
//     + "&c=" + encodeURIComponent(JSON.stringify(chartConfig));

//   try {
//     var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
//     if (res.getResponseCode() === 200) {
//       Logger.log("✅ QuickChart — h=" + chartHeight);
//       return res.getBlob().setName("chart.png");
//     }
//     Logger.log("QuickChart HTTP " + res.getResponseCode());
//   } catch (e) {
//     Logger.log("QuickChart: " + e.message);
//   }

//   // Fallback: Charts API
//   try {
//     var dt = Charts.newDataTable()
//       .addColumn(Charts.ColumnType.STRING, "النوع")
//       .addColumn(Charts.ColumnType.NUMBER, "العدد");
//     categories.forEach(function(c) { dt.addRow([c.label, c.value]); });

//     return Charts.newBarChart()
//       .setDataTable(dt)
//       .setTitle("توزيع المهام · " + (title2 || ""))
//       .setOption("titleTextStyle",  { fontSize: 14, bold: true, color: "#0C2340" })
//       .setOption("backgroundColor", { fill: "#F8FAFF" })
//       .setOption("colors",          ["#1A56DB"])
//       .setOption("legend",          { position: "none" })
//       .setOption("chartArea",       { left: 30, top: 55, width: "68%", height: "83%" })
//       .setOption("hAxis", { minValue: 0, gridlines: { color: "#E2E8F0" }, textStyle: { color: "#64748B" } })
//       .setOption("vAxis", { textStyle: { color: "#0F172A", bold: true } })
//       .setDimensions(800, chartHeight)
//       .build()
//       .getAs("image/png");
//   } catch (e2) {
//     Logger.log("Charts API: " + e2.message);
//     return null;
//   }
// }

// function createMultiVariableChart(sheet, counts, monthName, year) {

//   var categories = [
//     { label: "طلب خدمة",         value: counts.TotalTaskRequest       || 0, color: "#2563EB" },
//     { label: "مراسلة · مواصفات", value: counts.TotalTechReq           || 0, color: "#0C2340" },
//     { label: "مراسلة · تحليل",   value: counts.TotalEvaluation        || 0, color: "#1A56DB" },
//     { label: "مراسلة · أخرى",    value: counts.TotalOthers            || 0, color: "#60A5FA" },
//     { label: "طلب مباشر",        value: counts.TotalDirectRequests    || 0, color: "#059669" },
//     { label: "تكليف مباشر",      value: counts.TotalDirectAssignments || 0, color: "#7C3AED" },
//     { label: "اتصال",             value: counts.TotalCalls             || 0, color: "#D97706" },
//   ];

//   categories.sort(function(a, b) { return b.value - a.value; });

//   var labels     = categories.map(function(c) { return c.label; });
//   var values     = categories.map(function(c) { return c.value; });
//   var colors     = categories.map(function(c) { return c.color; });
//   var titleLine2 = (monthName && year) ? (monthName + " " + year) : "";
//   var chartHeight = categories.length * 52 + 160;

//   var chartConfig = {
//     type: "bar",
//     data: {
//       labels: labels,
//       datasets: [{
//         data:            values,
//         backgroundColor: colors,
//         borderRadius:    8,
//         borderSkipped:   false,
//         barPercentage:   0.72
//       }]
//     },
//     options: {
//       indexAxis: "y",
//       plugins: {
//         legend: { display: false },
//         title: {
//           display: true,
//           text: titleLine2
//             ? ["توزيع المهام المكتملة حسب النوع", titleLine2]
//             : ["توزيع المهام المكتملة حسب النوع"],
//           color: "#0C2340",
//           font:  { size: 16, weight: "bold", family: "Arial" },
//           padding: { bottom: 20 }
//         },
//         datalabels: {
//           display:   true,
//           anchor:    "end",
//           align:     "end",
//           clamp:     true,
//           color:     "#0C2340",
//           font:      { weight: "bold", size: 13, family: "Arial" },
//           formatter: function(val) { return val > 0 ? val : ""; }
//         }
//       },
//       scales: {
//         x: {
//           beginAtZero: true,
//           grid:  { color: "#E2E8F0", lineWidth: 1 },
//           ticks: { color: "#64748B", font: { size: 12 } }
//         },
//         y: {
//           grid:  { display: false },
//           ticks: { color: "#0F172A", font: { weight: "bold", size: 12 } }
//         }
//       },
//       layout: { padding: { right: 65, top: 8, bottom: 10, left: 4 } }
//     }
//   };

//   // ── QuickChart (الطريقة الأولى) ──
//   // var url = "https://quickchart.io/chart"
//   //   + "?v=3&w=900&h=" + chartHeight
//   //   + "&bkg=%23FFFFFF"
//   //   + "&c=" + encodeURIComponent(JSON.stringify(chartConfig));

//   var url = "https://quickchart.io/chart"
//     + "?v=3&w=900&h=" + chartHeight
//     + "&devicePixelRatio=2"
//     + "&bkg=%23FFFFFF"
//     + "&c=" + encodeURIComponent(JSON.stringify(chartConfig));

//   try {
//     var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
//     if (res.getResponseCode() === 200) {
//       Logger.log("✅ QuickChart: تم توليد الرسم بنجاح");
//       return res.getBlob().setName("chart.png");
//     }
//     Logger.log("QuickChart HTTP " + res.getResponseCode());
//   } catch (e) {
//     Logger.log("QuickChart error: " + e.message);
//   }

//   // ── Charts API الاحتياطي ──
//   Logger.log("التحويل إلى Charts API الاحتياطي");
//   try {
//     var dt = Charts.newDataTable()
//       .addColumn(Charts.ColumnType.STRING, "النوع")
//       .addColumn(Charts.ColumnType.NUMBER, "العدد");
//     categories.forEach(function(c) { dt.addRow([c.label, c.value]); });

//     return Charts.newBarChart()
//       .setDataTable(dt)
//       .setTitle("توزيع المهام المكتملة · " + (titleLine2 || ""))
//       .setOption("titleTextStyle",  { fontSize: 15, bold: true, color: "#0C2340" })
//       .setOption("backgroundColor", { fill: "#F8FAFF" })
//       .setOption("colors",          ["#1A56DB"])
//       .setOption("legend",          { position: "none" })
//       .setOption("chartArea",       { left: 30, top: 60, width: "70%", height: "85%" })
//       .setOption("hAxis", { minValue: 0, gridlines: { color: "#E2E8F0" }, textStyle: { color: "#64748B" } })
//       .setOption("vAxis", { textStyle: { color: "#0F172A", bold: true } })
//       .setDimensions(900, chartHeight)
//       .build()
//       .getAs("image/png");
//   } catch (e2) {
//     Logger.log("Charts API error: " + e2.message);
//     return null;
//   }
// }

// function createMultiVariableChart(sheet, counts) {
//   var labels = [
//     "طلبات الخدمة",
//     "مراسلات مواصفات",
//     "مراسلات تحليل",
//     "مراسلات أخرى",
//     "اتصال",
//     "طلب مباشر",
//     "تكليف مباشر"
//   ];
  
//   var dataValues = [
//     counts.TotalTaskRequest || 0,
//     counts.TotalTechReq || 0,
//     counts.TotalEvaluation || 0,
//     counts.TotalOthers || 0,
//     counts.TotalCalls || 0,
//     counts.TotalDirectRequests || 0,
//     counts.TotalDirectAssignments || 0
//   ];

//   // ألوان هادئة متعددة ومتناسقة (Pastel Multi-color)
//   var softColors = [
//     '#3B82F6', // أزرق هادئ
//     '#10B981', // أخضر زمردي
//     '#F59E0B', // برتقالي هادئ
//     '#8B5CF6', // بنفسجي دافئ
//     '#EC4899', // وردي خافت
//     '#06B6D4', // أزرق سماوي
//     '#64748B'  // رمادي داكن هادئ
//   ];

//   // ----------------------------------------------------
//   // الخيار الأول: الدائري التفاعلي / دونات (Doughnut Chart)
//   // ----------------------------------------------------
//   var doughnutConfig = {
//     type: 'doughnut',
//     data: {
//       labels: labels,
//       datasets: [{
//         data: dataValues,
//         backgroundColor: softColors,
//         borderWidth: 2,
//         borderColor: '#FFFFFF'
//       }]
//     },
//     options: {
//       plugins: {
//         title: {
//           display: true,
//           text: 'توزيع المهام والطلبات (نسبة وحجم)',
//           color: '#1E293B',
//           font: { size: 18, weight: 'bold', family: 'sans-serif' },
//           padding: { bottom: 15 }
//         },
//         legend: {
//           display: true,
//           position: 'right', // دليل الألوان على اليمين
//           labels: {
//             color: '#334155',
//             font: { size: 12, weight: '600' },
//             padding: 15,
//             usePointStyle: true // نقاط دائرية بدلاً من المربعات
//           }
//         },
//         // إظهار الرقم المباشر والنسبة المئوية % معاً
//         datalabels: {
//           color: '#FFFFFF',
//           font: { weight: 'bold', size: 12 },
//           formatter: (value, ctx) => {
//             let sum = ctx.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
//             if (sum === 0 || value === 0) return '';
//             let percentage = (value * 100 / sum).toFixed(1) + "%";
//             return `${value}\n(${percentage})`; // إظهار الرقم ثم النسبة
//           }
//         }
//       },
//       cutout: '55%' // تفريغ عصري في الوسط
//     }
//   };

//   // توليد رابط الصورة للرسم الدائري عبر QuickChart API بدقة عالية
//   var chartUrl = 'https://quickchart.io/chart?w=700&h=400&devicePixelRatio=2&bkg=white&c=' + encodeURIComponent(JSON.stringify(doughnutConfig));
  
//   try {
//     var response = UrlFetchApp.fetch(chartUrl);
//     return response.getBlob().setName("doughnut_chart.png");
//   } catch (e) {
//     Logger.log("خطأ في جلب الرسم البياني: " + e.toString());
//     return null;
//   }
// }

/**
 * دالة آمنة لإدراج الصورة في موقع الوسم بدون أخطاء النصوص الفارغة
 */
// function replaceTagWithImage(doc, tag, imageBlob) {
//   var body = doc.getBody();
//   var found = body.findText(tag);

//   if (found) {
//     var element = found.getElement();
//     var parent = element.getParent();

//     // 1. استبدال نص الوسم بمسافة واحدة لتجنب خطأ النص الفارغ Empty Text Element
//     body.replaceText(tag, " ");

//     // 2. إدراج الصورة داخل الفقرة الحاوية للوسم
//     if (parent.getType() == DocumentApp.ElementType.PARAGRAPH) {
//       var paragraph = parent.asParagraph();
//       var img = paragraph.appendInlineImage(imageBlob);
      
//       // ضبط أبعاد الصورة
//       img.setWidth(450);
//       // img.setHeight(300);
//     } 
//     // إذا كان الوسم داخل خلية جدول
//     else {
//       var container = parent.getParent();
//       if (container && container.getType() == DocumentApp.ElementType.TABLE_CELL) {
//         var cell = container.asTableCell();
//         var img = cell.appendParagraph("").appendInlineImage(imageBlob);
        
//         img.setWidth(320);
//         // img.setHeight(390);
//       }
//     }
//   }
// }

// function replaceTagWithImage(doc, tag, imageBlob) {
//   if (!imageBlob) {
//     doc.getBody().replaceText(tag, "[تعذّر إنشاء الرسم البياني]");
//     return;
//   }

//   var body  = doc.getBody();
//   var found = body.findText(tag);
//   if (!found) return;

//   var parent = found.getElement().getParent();

//   if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
//     var para = parent.asParagraph();

//     // امسح الوسم تماماً وهيّئ الفقرة
//     para.setText("");
//     para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
//     para.setSpacingBefore(6);
//     para.setSpacingAfter(6);

//     // أدرج الصورة — العرض فقط، لا تضع setHeight أبداً
//     var img = para.appendInlineImage(imageBlob);
//     img.setWidth(440);

//   } else if (parent.getType() === DocumentApp.ElementType.TABLE_CELL) {
//     var cell = parent.asTableCell();
//     cell.clear();
//     var newPara = cell.appendParagraph("");
//     newPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
//     var img2 = newPara.appendInlineImage(imageBlob);
//     img2.setWidth(300);
//   }
// }

// function replaceTagWithImage(doc, tag, chartResult) {
//   // chartResult يمكن أن يكون { blob, width, height } أو blob مباشرة (للتوافق)
//   // var imageBlob, logW, logH;
//   // if (chartResult && chartResult.blob) {
//   //   imageBlob = chartResult.blob;
//   //   logW      = chartResult.width  || 445;
//   //   logH      = chartResult.height || 320;
//   // } else {
//   //   imageBlob = chartResult;
//   //   logW = 445; logH = 320;
//   // }

//   // ── استخراج الـ Blob والأبعاد من النتيجة ──
//   var imageBlob = (chartResult && chartResult.blob) ? chartResult.blob : chartResult;
//   var logW      = (chartResult && chartResult.width)  ? chartResult.width  : 445;
//   var logH      = (chartResult && chartResult.height) ? chartResult.height : 320;
//   var displayW  = 440;
//   var displayH  = Math.round(displayW * logH / logW);
//   // ─────────────────────────────────────────

//   if (!imageBlob) {
//     doc.getBody().replaceText(tag, "[تعذّر إنشاء الرسم البياني]");
//     return;
//   }

//   var body  = doc.getBody();
//   var found = body.findText(tag);
//   if (!found) return;

//   var parent = found.getElement().getParent();

//   // ── عرض العرض في الدوكس وحساب الارتفاع بنفس النسبة ──
//   var displayW = 440;
//   var displayH = Math.round(displayW * logH / logW);  // نسبة صحيحة دائماً

//   if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
//     var para = parent.asParagraph();
//     para.setText("");
//     para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
//     para.setSpacingBefore(4);
//     para.setSpacingAfter(4);

//     var img = para.appendInlineImage(imageBlob);
//     img.setWidth(displayW);
//     img.setHeight(displayH);

//   } else if (parent.getType() === DocumentApp.ElementType.TABLE_CELL) {
//     var cell = parent.asTableCell();
//     cell.clear();
//     var p2  = cell.appendParagraph("");
//     p2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
//     var img2 = p2.appendInlineImage(imageBlob);
//     // داخل الجدول: عرض أصغر قليلاً
//     var cellW = 320;
//     img2.setWidth(cellW);
//     img2.setHeight(Math.round(cellW * logH / logW));
//   }
// }

// function replaceTagWithImage(doc, tag, imageBlob) {
//   if (!imageBlob) {
//     doc.getBody().replaceText(tag, "[تعذّر إنشاء الرسم البياني]");
//     return;
//   }

//   Logger.log("imageBlob type: " + typeof imageBlob);

//   var body  = doc.getBody();
//   var found = body.findText(tag);
//   if (!found) return;

//   var parent = found.getElement().getParent();
//   // var dW = 440;
//   // var dH = Math.round(dW * _chartH / _chartW);
//   // ── في replaceTagWithImage: العرض يبقى 440، الارتفاع يُحسب من النسبة ──
//   var dW = 440;
//   var dH = Math.round(dW * _chartH / _chartW); // = 440 × CH/520

//   if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
//     var para = parent.asParagraph();
//     para.setText("");
//     para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
//     var img = para.appendInlineImage(imageBlob);
//     img.setWidth(dW);
//     img.setHeight(dH);

//   } else if (parent.getType() === DocumentApp.ElementType.TABLE_CELL) {
//     var cell = parent.asTableCell();
//     cell.clear();
//     var p2   = cell.appendParagraph("");
//     p2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
//     var cW   = 320;
//     var img2 = p2.appendInlineImage(imageBlob);
//     img2.setWidth(cW);
//     img2.setHeight(Math.round(cW * _chartH / _chartW));
//   }
// }

// function replaceTagWithImage(doc, tag, imageBlob) {
//   if (!imageBlob) {
//     doc.getBody().replaceText(tag, "[تعذّر إنشاء الرسم البياني]");
//     return;
//   }

//   var body  = doc.getBody();
//   var found = body.findText(tag);
//   if (!found) return;

//   var parent = found.getElement().getParent();

//   // حساب الارتفاع من نسبة الرسم (600×400 = نسبة 2:3)
//   var dW = 440;
//   var dH = Math.round(dW * _chartH / _chartW); // 440 × 400/600 ≈ 293

//   if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
//     var para = parent.asParagraph();
//     para.setText("");
//     para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
//     var img = para.appendInlineImage(imageBlob);
//     img.setWidth(dW);
//     img.setHeight(dH);

//   } else if (parent.getType() === DocumentApp.ElementType.TABLE_CELL) {
//     var cell = parent.asTableCell();
//     cell.clear();
//     var p2   = cell.appendParagraph("");
//     p2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
//     var img2 = p2.appendInlineImage(imageBlob);
//     img2.setWidth(320);
//     img2.setHeight(Math.round(320 * _chartH / _chartW));
//   }
// }

function replaceTagWithImage(doc, tag, imageBlob) {
  if (!imageBlob) {
    doc.getBody().replaceText(tag, "[تعذّر إنشاء الرسم البياني]");
    return;
  }

  var body  = doc.getBody();
  var found = body.findText(tag);
  if (!found) return;

  var parent = found.getElement().getParent();
  var dW = 440;
  var dH = Math.round(dW * _chartH / _chartW); // 440 × 420/600 = 308

  if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
    var para = parent.asParagraph();
    para.setText("");
    para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    var img = para.appendInlineImage(imageBlob);
    img.setWidth(dW);
    img.setHeight(dH);

  } else if (parent.getType() === DocumentApp.ElementType.TABLE_CELL) {
    var cell = parent.asTableCell();
    cell.clear();
    var p2   = cell.appendParagraph("");
    p2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    var img2 = p2.appendInlineImage(imageBlob);
    img2.setWidth(320);
    img2.setHeight(Math.round(320 * _chartH / _chartW));
  }
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
 * دالة آمنة لإدراج الصورة في موقع الوسم بدون أخطاء النصوص الفارغة
 */
function replaceTagWithImage(doc, tag, imageBlob) {
  var body = doc.getBody();
  var found = body.findText(tag);

  if (found) {
    var element = found.getElement();
    var parent = element.getParent();

    // 1. استبدال نص الوسم بمسافة واحدة لتجنب خطأ النص الفارغ Empty Text Element
    body.replaceText(tag, " ");

    // 2. إدراج الصورة داخل الفقرة الحاوية للوسم
    if (parent.getType() == DocumentApp.ElementType.PARAGRAPH) {
      var paragraph = parent.asParagraph();
      var img = paragraph.appendInlineImage(imageBlob);
      
      // ضبط أبعاد الصورة
      img.setWidth(480);
      img.setHeight(300);
    } 
    // إذا كان الوسم داخل خلية جدول
    else {
      var container = parent.getParent();
      if (container && container.getType() == DocumentApp.ElementType.TABLE_CELL) {
        var cell = container.asTableCell();
        var img = cell.appendParagraph("").appendInlineImage(imageBlob);
        
        img.setWidth(480);
        img.setHeight(390);
      }
    }
  }
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
