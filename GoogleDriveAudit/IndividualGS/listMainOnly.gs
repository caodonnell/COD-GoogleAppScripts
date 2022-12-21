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
