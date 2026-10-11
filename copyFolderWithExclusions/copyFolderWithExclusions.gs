/**
 * إعدادات السكربت - عدّل القيم هنا حسب حاجتك
 */
const CONFIG2 = {
  // SOURCE_FOLDER_ID: '',
  // DEST_PARENT_FOLDER_ID: '',
  // NEW_FOLDER_NAME: '',
  // EXCLUDE_NAMES: [''],

  SOURCE_FOLDER_ID: '1POn66-jw_Qw64q02Uskz-YeB7uJ8CKxR',
  DEST_PARENT_FOLDER_ID: '1vpxScxCWDzs9PkB7LHNx-s6WN8MSAvxe',
  // NEW_FOLDER_NAME: 'Technical Support - Omaninfo - Backup',
  // EXCLUDE_NAMES: ['Reports','Tasks-Report-Backups']
  EXCLUDE_NAMES: ['Backup - Jun-2026','Backup - Jul-2026','Backup - Aug-2026','Backup - Sep-2026','Generated-Reports']

  // Backup
  // SOURCE_FOLDER_ID: '1E1AmwLlAlbrr0iwAovqniweKnbZfLhLc',
  // DEST_PARENT_FOLDER_ID: '1vpxScxCWDzs9PkB7LHNx-s6WN8MSAvxe',
  // NEW_FOLDER_NAME: 'TestScript-01072026',
  // EXCLUDE_NAMES: ['Reports','Tasks-Report-Backups']

  // استثناء ملفات حسب الامتداد (بدون نقطة)
  // EXCLUDE_EXTENSIONS: ['tmp', 'log'],

  // استثناء أي اسم يحتوي على كلمة معينة (اختياري)
  // EXCLUDE_CONTAINS: ['backup', 'draft']
};

function copyFolderWithExclusions() {
  const sourceFolder = DriveApp.getFolderById(CONFIG2.SOURCE_FOLDER_ID);
  const destParent = DriveApp.getFolderById(CONFIG2.DEST_PARENT_FOLDER_ID);

  // توليد الاسم تلقائياً: اسم المجلد المصدر + التاريخ
  const newFolderName = generateFolderName(sourceFolder.getName());

  const newFolder = destParent.createFolder(newFolderName);
  Logger.log('تم إنشاء المجلد الجديد: ' + newFolder.getName());

  copyContents(sourceFolder, newFolder);

  Logger.log('اكتملت عملية النسخ بنجاح.');
}

function generateFolderName(sourceName) {
  const timeZone = Session.getScriptTimeZone();
  // const dateStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  const dateStr = Utilities.formatDate(new Date(), timeZone, 'yyyyMMdd-HHmmss');
  return sourceName + ' - ' + dateStr;
}

function copyContents(sourceFolder, destFolder) {
  const files = sourceFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (shouldExclude(file.getName())) {
      Logger.log('تم تجاهل الملف: ' + file.getName());
      continue;
    }
    file.makeCopy(file.getName(), destFolder);
  }

  const subFolders = sourceFolder.getFolders();
  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    if (shouldExclude(subFolder.getName())) {
      Logger.log('تم تجاهل المجلد: ' + subFolder.getName());
      continue;
    }
    const newSubFolder = destFolder.createFolder(subFolder.getName());
    copyContents(subFolder, newSubFolder);
  }
}

function shouldExclude(name) {
  if (CONFIG2.EXCLUDE_NAMES.includes(name)) return true;

  const ext = name.includes('.') ? name.split('.').pop().toLowerCase() : '';
  if (ext && CONFIG2.EXCLUDE_EXTENSIONS.includes(ext)) return true;

  // for (const keyword of CONFIG.EXCLUDE_CONTAINS) {
  //   if (name.toLowerCase().includes(keyword.toLowerCase())) return true;
  // }

  return false;
}
