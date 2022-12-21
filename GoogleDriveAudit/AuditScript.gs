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
 
 
//---------------------------------------------------------------------
// LISTING FILES AND FOLDERS
//---------------------------------------------------------------------
 
function listFilesAndFolders(){
 var folderId = Browser.inputBox('Enter folder ID', Browser.Buttons.OK_CANCEL);
 if (folderId === "") {
   Browser.msgBox('Folder ID is invalid');
   return;
 }
 getFolderTree(folderId, true);
};
// Get Folder Tree
function getFolderTree(folderId, listAll) {
 try {
   // Get folder by id
   var parentFolder = DriveApp.getFolderById(folderId);
   var parentName = parentFolder.getName();
 
   // Initialise the sheet
   sheet = setupSheet('Full audit (including subfiles/subfolders) of ', parentName, parentFolder);
 
   //Get files that are not in a subfolder
   var files = parentFolder.getFiles();
   while (listAll & files.hasNext()) {
     var childFile = files.next();
     sheet.appendRow(writeData(childFile, parentName));
   }
   // Get subfolders
   getChildFolders(parentName, parentFolder, sheet, listAll);
 } catch (e) {
   Logger.log(e.toString());
 }
 sheet.appendRow(['---FINISHED---']);
};
// Get the list of files and folders and their metadata in recursive mode
function getChildFolders(parentName, parent, sheet, listAll) {
 var childFolders = parent.getFolders();
 // List folders inside the folder
 while (childFolders.hasNext()) {
   var childFolder = childFolders.next();
   //  var folderId = childFolder.getId();
   sheet.appendRow(writeData(childFolder, parentName));
 
   // List files inside the folder
   var files = childFolder.getFiles();
   while (listAll & files.hasNext()) {
     var childFile = files.next();
     sheet.appendRow(writeData(childFile, parentName+'//'+childFolder.getName()));
   }
 // Recursive call of the subfolder
 getChildFolders(parentName + "//" + childFolder.getName(), childFolder, sheet, listAll);
 }
};
 
 
//---------------------------------------------------------------------
// LISTING FILES AND FOLDERS ONLY IN THE MAIN FOLDER
//---------------------------------------------------------------------
 
function listFilesNoSubs(){
 var folderId = Browser.inputBox('Enter folder ID', Browser.Buttons.OK_CANCEL);
 if (folderId === "") {
   Browser.msgBox('Folder ID is invalid');
   return;
 }
 getFolderNoSubs(folderId, true);
};
 
// Get Folder Files (but not searching subfolders)
function getFolderNoSubs(folderId, listAll) {
 try {
   // Get folder by id
   var parentFolder = DriveApp.getFolderById(folderId);
   var parentName = parentFolder.getName();
 
   // Initialise the sheet
   sheet = setupSheet('Audit (EXCLUDING subfiles/subfolders) of ', parentName, parentFolder);
 
   //Get files that are not in a subfolder
   var files = parentFolder.getFiles();
   while (listAll & files.hasNext()) {
     var childFile = files.next();
     // Logger.log(childFile.getName());
     sheet.appendRow(writeData(childFile, parentName));
   }
 
   // Get subfolders
 
   var childFolders = parentFolder.getFolders();
   // List folders inside the folder
   while (childFolders.hasNext()) {
   var childFolder = childFolders.next();
   sheet.appendRow(writeData(childFolder, parentName));
 }
 } catch (e) {
   Logger.log(e.toString());
 }
 sheet.appendRow(['---FINISHED WITHOUT GOING INTO SUBFOLDERS---']);
};
 
 
//---------------------------------------------------------------------
// LISTING FOLDERS ONLY
//---------------------------------------------------------------------
 
function listFoldersOnly(){
 var folderId = Browser.inputBox('Enter folder ID', Browser.Buttons.OK_CANCEL);
 if (folderId === "") {
   Browser.msgBox('Folder ID is invalid');
   return;
 }
 getFolderOnlyTree(folderId, true);
};
// Get Folder Tree
function getFolderOnlyTree(folderId, listAll) {
 try {
   // Get folder by id
   var parentFolder = DriveApp.getFolderById(folderId);
   var parentName = parentFolder.getName();
  
   // Initialise the sheet
   sheet = setupSheet('Folder-only audit (including subfolders) of ', parentName, parentFolder);
 
   // Get subfolders
   getChildFoldersOnly(parentName, parentFolder, sheet, listAll);
 } catch (e) {
   Logger.log(e.toString());
 }
 sheet.appendRow(['---FINISHED---']);
 };
 
 // Get the list of files and folders and their metadata in recursive mode
 function getChildFoldersOnly(parentName, parent, sheet, listAll) {
 var childFolders = parent.getFolders();
 // List folders inside the folder
 while (childFolders.hasNext()) {
   var childFolder = childFolders.next();
   sheet.appendRow(writeData(childFolder, parentName));
   // Recursive call of the subfolder
   getChildFoldersOnly(parentName + "//" + childFolder.getName(), childFolder, sheet, listAll);
 }
};
 
 
//---------------------------------------------------------------------
// LISTING FILES ONLY
//---------------------------------------------------------------------
 
function listFilesOnly(){
 var folderId = Browser.inputBox('Enter folder ID', Browser.Buttons.OK_CANCEL);
 if (folderId === "") {
   Browser.msgBox('Folder ID is invalid');
   return;
 }
 getFileOnlyTree(folderId, true);
 };
 
 // Get Folder Tree
 function getFileOnlyTree(folderId, listAll) {
 try {
   // Get folder by id
   var parentFolder = DriveApp.getFolderById(folderId);
   var parentName = parentFolder.getName();
 
   // Initialise the sheet
   sheet = setupSheet('File-only audit (including subfiles) of ', parentName, parentFolder);
  
   // Get subfolders
   getChildFilesOnly(parentName, parentFolder, sheet, listAll);
 } catch (e) {
   Logger.log(e.toString());
 }
 sheet.appendRow(['---FINISHED---']);
};
 
// Get the list of files and folders and their metadata in recursive mode
function getChildFilesOnly(parentName, parent, sheet, listAll) {
 try {
   //Get files that are not in a subfolder
   var files = parent.getFiles();
   while (listAll & files.hasNext()) {
     var childFile = files.next();
     sheet.appendRow(writeData(childFile, parentName));
   }
   //Go into the subfolders for more files
   var childFolders = parent.getFolders();
   while (childFolders.hasNext()) {
     var childFolder = childFolders.next();
     // Recursive call of the subfolder
     getChildFilesOnly(parentName + "//" + childFolder.getName(), childFolder, sheet, listAll);
   }
 } catch (e) {
   Logger.log(e.toString());
 }
};
 
 
//---------------------------------------------------------------------
// HELPER FUNCTIONS
//---------------------------------------------------------------------
 
// Initialise the sheet
function setupSheet(funcStr, parentName, parentFolder) { 
   var ssBook = SpreadsheetApp.getActiveSpreadsheet();
   var datetime = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "MMddyyyy HH:mm:ss");
   var sheet = ssBook.insertSheet('Audit '+datetime);
   sheet.appendRow([funcStr+parentName+': '+parentFolder.getUrl()]);
   sheet.appendRow(["Full Path", "Name","Type" ,"Date Created", "Last Updated", "URL", "File/Folder ID", "Parent Folder ID", "Size","Owner Email", "Editor Emails", "Viewer/Commenter Emails"]);
   sheet.setFrozenRows(2);
 
   return sheet
}
 
// Write data about a file, folder
function writeData(childFile, parentName) {
 
 var filetype;
 try {
   filetype = getFileType(childFile.getMimeType());
 } catch (e) {
   filetype = 'Folder';
 }
 
 var data = [
   parentName + "//" + childFile.getName(),
   childFile.getName(),
   filetype,
   childFile.getDateCreated(),
   childFile.getLastUpdated(),
   childFile.getUrl(),
   childFile.getId(),
   childFile.getParents().next().getId(),
   childFile.getSize()/1024
 ];
 try {
   data.push(childFile.getOwner().getEmail());
 } catch (e) {
   data.push('')
 }
 try {
   data.push(childFile.getEditors().map(e => e.getEmail()).join('; '));
 } catch (e) {
   data.push('')
 }
 try {
   data.push(childFile.getViewers().map(e => e.getEmail()).join('; '));
 } catch (e) {
   data.push('')
 }
 return data;
}
 
// Parse the file mime type to return a more "readable" type
function getFileType(mimeType) {
 switch (mimeType) {
   case MimeType.GOOGLE_APPS_SCRIPT:
     return 'Google Apps Script';
   case MimeType.GOOGLE_DRAWINGS:
     return 'Google Drawings';
   case MimeType.GOOGLE_DOCS:
     return 'Google Docs';
   case MimeType.GOOGLE_FORMS:
     return 'Google Forms';
   case MimeType.GOOGLE_SHEETS:
     return 'Google Sheets';
   case MimeType.GOOGLE_SLIDES:
     return'Google Slides';
   case MimeType.FOLDER:
     return 'Folder';
   case MimeType.BMP:
     return 'BMP';
   case MimeType.GIF:
     return 'GIF';
   case MimeType.JPEG:
     return 'JPEG';
   case MimeType.PNG:
     return 'PNG';
   case MimeType.SVG:
     return 'SVG';
   case MimeType.PDF:
     return 'PDF';
   case MimeType.CSS:
     return 'CSS';
   case MimeType.CSV:
     return 'CSV';
   case MimeType.HTML:
     return 'HTML';
   case MimeType.JAVASCRIPT:
     return 'JavaScript';
   case MimeType.PLAIN_TEXT:
     return 'Plain Text';
   case MimeType.RTF:
     return 'Rich Text';
   case MimeType.SHORTCUT:
     return 'Shortcut';
   case "application/x-iwork-keynote-sffkey":
     return 'Apple Keynote';
   case "application/x-iwork-numbers-sffnumbers":
     return 'Apple Numbers';
   case "application/x-iwork-pages-sffpages":
     return 'Apple Pages';
   case MimeType.OPENDOCUMENT_GRAPHICS:
     return 'OpenDocument Graphics';
   case MimeType.OPENDOCUMENT_PRESENTATION:
     return 'OpenDocument Presentation';
   case MimeType.OPENDOCUMENT_SPREADSHEET:
     return 'OpenDocument Spreadsheet';
   case MimeType.OPENDOCUMENT_TEXT:
     return 'OpenDocument Word';
   case MimeType.MICROSOFT_EXCEL:
    return 'Microsoft Excel';
   case MimeType.MICROSOFT_EXCEL_LEGACY:
     return 'Microsoft Excel';
   case MimeType.MICROSOFT_POWERPOINT:
     return 'Microsoft PowerPoint';
   case MimeType.MICROSOFT_POWERPOINT_LEGACY:
     return 'Microsoft PowerPoint';
   case MimeType.MICROSOFT_WORD:
     return 'Microsoft Word';
   case MimeType.MICROSOFT_WORD_LEGACY:
     return 'Microsoft Word';
   case MimeType.ZIP:
     return 'ZIP';
   default:
     return mimeType;
 }
}
