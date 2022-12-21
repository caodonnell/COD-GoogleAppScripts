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
