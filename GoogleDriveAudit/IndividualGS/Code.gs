function onOpen() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('List Files and Folders')
    .addItem('List All Files and Folders', 'listFilesAndFolders')
    .addItem('List Main Folder Files (no subfolder files)', 'listFilesNoSubs')
    .addItem('List Folders ONLY', 'listFoldersOnly')
    .addItem('List Files ONLY', 'listFilesOnly')
    .addToUi();
};
