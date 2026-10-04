function backupWholeFileWithFormat() {
  // ==== CONFIG ====
  var SOURCE_FILE_ID = "1INoQeNyBsgfIz-4UBkN3wzf8DR7dI_9SmmofbNJK0Qc"; 
  var TARGET_FOLDER_ID = "1OJ7wzigwTvNiAK4QIXH1mDfgxTzF6Ggl"; // optional: put a folder ID here
  // ==================

  var sourceFile = DriveApp.getFileById(SOURCE_FILE_ID);

  // create a timestamped name
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  //var backupName = sourceFile.getName() + " - Backup " + ts;
  var backupName = "Tasks-Reports-Backup - " + ts;

  // make a full copy (preserves formatting, charts, everything)
  var backupFile = sourceFile.makeCopy(backupName);

  // move backup into target folder
  if (TARGET_FOLDER_ID) {
    var folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    folder.addFile(backupFile);
    DriveApp.getRootFolder().removeFile(backupFile); // remove from root
  } else {
    // put in same folder as source
    var parents = sourceFile.getParents();
    if (parents.hasNext()) {
      var srcFolder = parents.next();
      srcFolder.addFile(backupFile);
      DriveApp.getRootFolder().removeFile(backupFile);
    }
  }

  Logger.log("Backup created: " + backupName + " — " + backupFile.getUrl());
}
