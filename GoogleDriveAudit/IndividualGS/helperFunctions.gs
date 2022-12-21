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
