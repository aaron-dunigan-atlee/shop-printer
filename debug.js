function activateDebug(){
  var myDate = new Date();
  myDate.setHours( myDate.getHours() + 1 );
  var cache = CacheService.getUserCache()
  console.warn('Entering debug mode until %s EST',Utilities.formatDate(myDate, 'EST', 'hh:mm a'))
  cache.put('debug','true',3600) // one hour
}

function deactivateDebug(){
  var myDate = new Date();
  var cache = CacheService.getUserCache()
  console.warn('Left debug mode at %s EST',Utilities.formatDate(myDate, 'EST', 'hh:mm a'))
  cache.remove('debug')
}

function checkDebug(){
  var cache = CacheService.getUserCache()
  var debug = cache.get('debug') ? 'on' : 'off'
  var ui = SpreadsheetApp.getUi()
  if (ui) ui.alert('Debug mode is: '+debug);
  console.log('Debug mode is: '+debug);
}

function getDebug(){
 var debugStatus = (CacheService.getUserCache().get('debug') == 'true')
 if (debugStatus) console.log('Debug mode is on.')
 return debugStatus
}

var DEBUG = getDebug();

function addDebugMenus() {
    SpreadsheetApp.getUi().createMenu('🕷Debug')
    .addItem('Enter Debug Mode', 'activateDebug')
    .addItem('Exit Debug Mode', 'deactivateDebug')
    .addItem('Check Debug Mode Status', 'checkDebug')
    .addToUi()
}
