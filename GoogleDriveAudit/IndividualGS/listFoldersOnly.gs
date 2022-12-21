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
