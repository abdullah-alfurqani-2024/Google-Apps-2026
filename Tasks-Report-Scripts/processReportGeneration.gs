function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('التقارير الشهرية 📊')
    .addItem('إنشاء تقرير شهري', 'openDatePickerDialog')
    .addToUi();
}

// فتح نافذة التقويم في الملف الرئيسي
function openDatePickerDialog() {
  var html = HtmlService.createHtmlOutputFromFile('DateDialog')
      .setWidth(380)
      .setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'اختيار فترة التقرير');
}

// الدالة التي تستقبل التواريخ وتوجهها للمكتبة في الملف الثاني
function processReportGeneration(startDateInput, endDateInput) {
  // استدعاء الدالة الموجودة في مكتبة الملف الثاني مباشرة
  // اختر اسم اسم المكتبة كما أضفتها (مثلاً ReportLibrary)
  // return ReportLibrary.generateReportFromLibrary(startDateInput, endDateInput);
  return TasksReportLibrary.generateReportFromLibrary(startDateInput, endDateInput);
}
