var DROP_FOLDER_ID = '1E0CCIQ0NjoBUxPLxMaiqm2q0keaI4M8P'
var ARCHIVE_FOLDER_ID = '1xO_2xeSgNTYdLQlnJ8U4eHalMIoucHB5'
var ERROR_FOLDER_ID = '1M2mjtnIprw2Gp8Ps5zX0P_toBiO25TvT'

function processDroppedFilesSlack() {
  return runWithSlackReporting('processDroppedFiles')
}

/**
 * Scan the drop folder for files and process each one.
 * If successful, move to Archive folder.  
 * Otherwise move to Error folder and attempt to process next file.
 */
function processDroppedFiles() {
  var dropFolder = DriveApp.getFolderById(DROP_FOLDER_ID);
  var errorFolder = DriveApp.getFolderById(ERROR_FOLDER_ID);
  var archiveFolder = DriveApp.getFolderById(ARCHIVE_FOLDER_ID);
  
  var files = dropFolder.getFiles();
  if (!files.hasNext()) {
    console.log('No files found in drop folder.')
    return; 
  }
  var ss = SpreadsheetApp.getActive();
  var jobsSheet = ss.getSheetByName(JOB_SHEET_NAME)
  var cabinetsSheet = ss.getSheetByName(CABINET_SHEET_NAME)
  while (files.hasNext()) { 
    let file = files.next()
    try {
      processFile(file)
      archiveFolder.addFile(file)
      dropFolder.removeFile(file)
    } catch(err) {
      slackError(err, true, "Error processing file in Drop folder. Moved to Error folder.")
      errorFolder.addFile(file)
      dropFolder.removeFile(file)
    }
  }

  // Private functions
  // -----------------

  /**
   * Get rows data from csv file.
   * @param {DriveApp.File} file 
   */
  function processFile(file) {
    file = file || DriveApp.getFileById('1u_Q2V4iGWRsTmQJiYkNlEQNeNLt4FNf8')
    var csvRows = Utilities.parseCsv(file.getBlob().getDataAsString('UTF-16'))
    csvRows.shift() // One useless row.
    // Get the job name
    var jobNameString = csvRows.shift()[0] // e.g. "Job: Default (Ridervale-Baker)"
    var match = jobNameString.match(/Job: Default \((.+)\)/)
    if (!match) match = jobNameString.match(/Job:\s*(.+)/)
    var jobName = match ? match[1] : jobNameString
    console.log('Found file ' + file.getName() + ' for job "' + jobName + '"')

    var rows = getObjects_(csvRows, normalizeHeaders(CSV_HEADERS))
    console.log('Found ' + rows.length + ' cabinets.')
    // Add some data to cabinets objects
    var pieceCount = 0;
    rows.forEach(function(x, index){
      x.job = jobName;
      x.pieceCount = pieceCount + 1;
      pieceCount += x.qty;
    })
    rows.forEach(function(x){
      x.totalCount = pieceCount
    })

    setRowsData(cabinetsSheet, rows, {writeMethod: 'append'})

    // Write job object to job sheet
    var job = {
      name: jobName,
      cabinetCount: pieceCount,
      cabinetsCompleted: 0
    }
    setRowsData(jobsSheet, [job], {writeMethod: 'append'})
  }

}

/**
 * Extract text from a pdf stored in Google Drive.
 * @param {string} fileId 
 */
function extractPdfText(fileId) {
  var pdfFile = DriveApp.getFileById(fileId);
  var blob = pdfFile.getBlob();
  var resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  };
 
  // Needs Advanced Drive service. 
  try {
    var file = Drive.Files.insert(resource, blob, {ocr: true, ocrLanguage: "en"});
  } catch(err) {
    // We were experiencing "Internal Error" so if this happens we'll try one more time.
    slackWarn("extractPdfText: failed to call Drive.Files.insert\n" + err.message + "\nTrying again.")
    var file = Drive.Files.insert(resource, blob, {ocr: true, ocrLanguage: "en"});
  }
  
  var doc = DocumentApp.openById(file.id);
  var text = doc.getBody().getText();
  var file = DriveApp.getFileById(file.id);
  file.setTrashed(true);
  return text;
}


function convertCsvToSheets(fileId) {
  var sheetsFileJson = Drive.Files.copy({}, fileId, {convert: true}); 
  var sheet = SpreadsheetApp.openById(sheetsFileJson.id).getSheets()[0];
  rows = getRowsData(sheet,null,{getMetadata: true})
  var sheetsFile = DriveApp.getFileById(sheetsFileJson.id)
  sheetsFile.setTrashed(true)
}