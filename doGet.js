function doGet(e) {
  var page = e.parameter.page || 'index';
  console.log('Loading page %s with parameters %s', page, JSON.stringify(e.parameter));
  // Context object to be passed to page template.
  var context = {
    page: page,
    urlParams: e.parameter
  };


  try{
    var html = renderPage(context);
  } catch(err) {
    slackError(err,true);
    var html = HtmlService.createHtmlOutputFromFile('error').setTitle('Ultimate Cabinets');
  }
  return html
}

function include(filename, context){
  var template = HtmlService.createTemplateFromFile(filename);
  if (context) template.context = context;
  return template.evaluate().getContent();
}

function renderPage(context) {
  var template = HtmlService.createTemplateFromFile('base');
  template.context = context;
  return template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Ultimate Cabinets')
    .setFaviconUrl(FAVICON_URL);
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}