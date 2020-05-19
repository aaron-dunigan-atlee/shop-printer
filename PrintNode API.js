var PRINT_NODE_API_KEY = PropertiesService.getScriptProperties().getProperty('print_node_key')
var PRINT_NODE_BASE_URL = 'https://api.printnode.com'

function logPrinters() {
  var endpoint = '/printers'
  var options = {
    method: 'GET',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(PRINT_NODE_API_KEY)
    }
  }
  var result = UrlFetchApp.fetch(PRINT_NODE_BASE_URL + endpoint, options)
  Logger.log(result)
}

function testPrintNode() {
  var endpoint = '/whoami'
  var options = {
    method: 'GET',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(PRINT_NODE_API_KEY)
    }
  }
  var result = UrlFetchApp.fetch(PRINT_NODE_BASE_URL + endpoint, options)
  Logger.log(result)
}

/**
 * See https://www.printnode.com/en/docs/api/curl#printjobs
 * @param {DriveApp.File} pdfFile 
 * @param {integer} printerId 
 * @param {integer} quantity 
 */
function printPdf(pdfFile, printerId, quantity) {
  var endpoint = '/printjobs'
  var payload = {
    printerId: printerId,
    title: pdfFile.getName(),
    contentType: 'pdf_base64',
    content: Utilities.base64Encode(pdfFile.getBlob().getBytes()),
    source: 'Google Apps Script',
    qty: quantity || 1,
    options: {
      fit_to_page: true
    }
  }

  var options = {
    method: 'POST',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(PRINT_NODE_API_KEY)
    },
    payload: payload
  }
  var result = UrlFetchApp.fetch(PRINT_NODE_BASE_URL + endpoint, options)
  return result;
}