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
