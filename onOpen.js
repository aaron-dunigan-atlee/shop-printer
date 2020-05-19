function onOpen(e) {
  var ui = SpreadsheetApp.getUi()
  ui.createMenu('⬆️Import')
  .addItem('Process Imported files', 'processDroppedFilesSlack')
  .addToUi()
  ui.createMenu('🌐App')
  .addItem('Open Web App', 'openWebApp')
  .addToUi()
  ui.createMenu('📁Folders')
  .addItem('🟡Open Drop Folder', 'openDropFolder')  
  .addItem('🟢Open Already Processed Folder', 'openDestinationFolder')
  .addSeparator()
  .addItem('🔴Open Error Folder','openErrorFolder')
  .addToUi()
  
}


function openDropFolder(){
  openFolder(DROP_FOLDER_ID);
}

function openErrorFolder(){
  openFolder(ERROR_FOLDER_ID);
}
    
function openDestinationFolder(){
  openFolder(DEST_FOLDER_ID);
}

function openFolder(folderId){
  var folder = DriveApp.getFolderById(folderId);
  var titleMessage = folder.getName();
  var clickMessage = 'Click here to open';
  var url = 'https://drive.google.com/drive/u/1/folders/'+folderId;
  showAnchor(titleMessage, clickMessage, url);
}

function showAnchor(titleMessage, clickMessage, url) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var html = '<html><body><a href="' + url + '" target="blank" onclick="google.script.host.close()">' + clickMessage + '</a></body></html>';
  var ui = HtmlService.createHtmlOutput(html)
  SpreadsheetApp.getUi().showModelessDialog(ui, titleMessage);
}

function openWebApp(){
  var titleMessage = 'Go to Live App'
  var clickMessage = 'Click here to open';
  var url = 'https://script.google.com/a/ultimatecabinetsok.com/macros/s/AKfycbyi8slYD2AF6NSBDOOBxYILZ9-WVZEsTM8KBVKdgRXNB-NTD_I/exec';
  showAnchor(titleMessage, clickMessage, url);
}