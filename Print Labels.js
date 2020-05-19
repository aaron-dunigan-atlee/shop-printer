function printCabinetLabelSlack(id, job) {
  return runWithSlackReporting('printCabinetLabel', [id, job])
}

/**
 * Print a cabinet label.  Called from web app.
 * @param {string} id 
 * @param {string} jobName Name of job
 */
function printCabinetLabel(id, jobName) {
  console.log("Printing label for cabinet with id " + id + " on job " + jobName)
  var cabinet = getCabinets(jobName).find(function(x){return x.id == id})
  if (!cabinet) slackError("Failed to print label: Couldn't find cabinet with id " + id + " on job " + jobName)
  for (var i=0; i<cabinet.qty; i++) {
    var templateId = fillTemplate(CABINET_LABEL_TEMPLATE_ID, cabinet)
    var pdfFile = toPdf(templateId, null, true)
    if (DEBUG) {
      console.log('Debug "print": file is at ' + pdfFile.getUrl())
    } else {
      sendPdfToPrinter(pdfFile, 'Cabinet Labels Outfit')
      DriveApp.removeFile(pdfFile)
    }
    cabinet.pieceCount++
  }

  // Mark as printed
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CABINET_SHEET_NAME)
  var cabinetRow = sheet.getRange(cabinet.sheetRow, 1, 1, sheet.getMaxColumns());
  cabinet.labelPrinted = true;
  setRowsData(sheet, [cabinet], {firstRowIndex: cabinet.sheetRow, startHeader: 'Label Printed', endHeader: 'Label Printed'})

  // Update 'Cabinets Completed' on Jobs sheet
  var jobsSheet = ss.getSheetByName(JOB_SHEET_NAME);
  var jobsData = getRowsData(jobsSheet, null, {getMetadata: true})
  var job = jobsData.find(function(row){return row.name == cabinet.job})
  if (job) {
    job.cabinetsCompleted += cabinet.qty
    if (job.cabinetsCompleted == job.cabinetCount) job.jobCompleted = true;
    setRowsData(jobsSheet, [job], {firstRowIndex: job.sheetRow})
    if (job.jobCompleted) printEndRunLabel(job)
  } else {
    slackWarn("printingCabinetLabel: Couldn't update cabinets completed because this job was not found on Jobs sheet: " + cabinet.job)
  }
  return
}

/**
 * Print a label when a  job is finished. 
 * @param {Object} job Job row object
 */
function printEndRunLabel(job) {
  console.log("Printing end run label for job " + job.name)
  var templateId = fillTemplate(END_RUN_TEMPLATE_ID, job)
  var pdfFile = toPdf(templateId, null, true)
  if (DEBUG) {
    console.log('Debug "print": file is at ' + pdfFile.getUrl())
  } else {
    sendPdfToPrinter(pdfFile, 'End Run Outfit', cabinet.qty)
    DriveApp.removeFile(pdfFile)
  }

}

/**
 * Specify printer by name.  Defaults to 'Cabinet Labels Outfit'
 * @param {File} pdfFile 
 * @param {string} printerName 
 * @param {integer} quantity
 */
function sendPdfToPrinter(pdfFile, printerName, quantity) {
  var printerId
  if (printerName) printerId = PRINTER_IDS[printerName]
  if (!printerName) printerId = PRINTER_IDS['Cabinet Labels Outfit']
  printPdf(pdfFile, printerId, quantity)
}

function printSampleLabel() {
  var pdfFile = toPdf(CABINET_LABEL_TEMPLATE_ID, PRINTED_FOLDER_ID)
  var printJobId = printPdf(pdfFile, 69452243)
  console.log('Sample label printed with id ' + printJobId)
}

function setLabelPageFormat(document) {
  document = document || DocumentApp.openById('1UQpf_herS6ASRA4QvAmOtRcCDDhGXb2KEIi2u9G2UQE')
  var body = document.getBody()
  body.setPageHeight(216).setPageWidth(288) // 3 x 4 inches.
  var style = {};
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FFFFFF' //'#DDDDDD';// For testing, gray background
  style[DocumentApp.Attribute.MARGIN_BOTTOM] = 0.25;
  style[DocumentApp.Attribute.MARGIN_LEFT] = 0.25;
  style[DocumentApp.Attribute.MARGIN_RIGHT] = 0.25;
  style[DocumentApp.Attribute.MARGIN_TOP] = 0.25;
  body.setAttributes(style)
}


/**
 * Move Drive file to a destination folder and remove it from all other folders.
 * @param {file} file 
 * @param {folder} destinationFolder 
 */
function moveFile(file, destinationFolder) {
  // Get previous parent folders.
  var oldParents = file.getParents();
  // Add file to destination folder.
  destinationFolder.addFile(file);
  // Remove previous parents.
  while (oldParents.hasNext()) {
    var oldParent = oldParents.next();
    // In case the destination folder was already a parent, don't remove it.
    if (oldParent.getId() != destinationFolder.getId()) {
      oldParent.removeFile(file);
    }
  }
}

/**
 * Convert Drive file to pdf and move to desired folder.
 * @param {string} fileId 
 * @param {string} folderId 
 * @param {boolean} removeOriginal  Whether to delete the original after converting.
 */
function toPdf(fileId, folderId, removeOriginal) {
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var newPdfFile = DriveApp.createFile(blob);
  var acmeFolder = folderId? DriveApp.getFolderById(folderId) : file.getParents().next();
  moveFile(newPdfFile, acmeFolder);
  if (removeOriginal) file.setTrashed(true)
  return newPdfFile;
}