/**
 * Make a copy of a template and fill its fields.
 * @param {string} templateId 
 * @param {Object} replacementObject 
 * @param {DriveApp.Folder} destinationFolder 
 * @param {string} filename 
 * @param {boolean} replaceEmptyFields  Defaults to true.  If false, leave placeholder if there is no value for the field.
 */
function fillTemplate(templateId, replacementObject, destinationFolder, filename, replaceEmptyFields) {
  if (typeof replaceEmptyFields === 'undefined') replaceEmptyFields = true;
  var file = DriveApp.getFileById(templateId);
  filename = filename || file.getName()
  destinationFolder = destinationFolder || file.getParents().next();
  var templateAsFile = file.makeCopy(filename, destinationFolder);
  var templateId = templateAsFile.getId();
  var templateAsDoc = DocumentApp.openById(templateId);
  var requests = [];
  // Add requests for fields.  
  // Start with a template of empty strings in case any fields are missing.
  var templateObject = getEmptyTemplateObject(templateAsDoc);

  for (var prop in templateObject) {
    if(replacementObject[prop] != undefined){
      templateObject[prop] = replacementObject[prop].toString();
    } 
  }
  requests = buildRequests(templateObject, replaceEmptyFields);
  // Batch update all requests

  // Requires advanced Docs service
  var response = Docs.Documents.batchUpdate({'requests': requests}, templateId);
  console.log(JSON.stringify(response))
  
  templateAsDoc.saveAndClose();
  return templateAsDoc.getId();
}


/**
 * Find names of all fields, written as {fieldName}, in the document,
 * and returns an object with each field assigned an empty string. 
 * @param {Document} document 
 */
function getEmptyTemplateObject(document) {
  var body = document.getBody();
  var searchPattern = /{.*?}/g;
  var bodyText = body.getText();
  var patternMatch = bodyText.match(searchPattern)
  if (!patternMatch) return {}
  var matches = patternMatch.map(function(text){
    return text.slice(1, -1);
  });
  var templateObject = {}
  matches.forEach(function(fieldName){
    templateObject[fieldName] = '';
  });
  return templateObject;
}

/**
 * Create a Docs API request to replace text for each field in replacementObject,
 * and append to the requests array.
 * @param {Array} requests 
 * @param {Object} replacementObject 
 */
function buildRequests(replacementObject, replaceEmptyFields) {
  var requests = []
  for (var prop in replacementObject){
    // i.e. if replaceEmptyFields, replace all, but if not, only replace if there is a value for prop.
    if (replaceEmptyFields || replacementObject[prop]) {
      var request = {
        'replaceAllText': {
          'containsText': {'text': "{"+prop+"}"},
          'replaceText': replacementObject[prop]
        }
      };
      requests.push(request);
    }
  }
  return requests;
}
