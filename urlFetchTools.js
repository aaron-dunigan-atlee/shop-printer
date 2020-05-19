var FetchTools = (function (ns) {
  var BATCH_SIZE = 10;
  var DEFAULT_RETRIES = 3;
  
  ns.backoffOne = function(url, options ,retries){
    retries = retries || DEFAULT_RETRIES;
    var success = false;
    for (var attempt = 0; attempt <= retries; attempt++) {
      if (attempt > 0) Utilities.sleep(5000 * Math.pow(2, attempt)); //don't wait the first time
      try{
        var responses = UrlFetchApp.fetch(url, options); 
        success = true;
      } catch (err) {
        if(attempt == retries){
          throw err
          break;
        }
        console.error('Fetch error, '+(retries-attempt)+' retries remain.\n'+err)
      }
      if (success) break;
    }  
    return responses
  } // fetchTools.backoffOne()
  
  ns.backoffBatches = function(requests,retries){
    retries = retries || DEFAULT_RETRIES;
    var responses = [];
    for (var i = 0; i < Math.ceil(requests.length/BATCH_SIZE); i++){
      var miniBatch = [];
      for (var j = 0; j < BATCH_SIZE; j++){
        if (i*BATCH_SIZE+j>=requests.length) break;
        miniBatch.push(requests[i*BATCH_SIZE+j]);
      }
  
      try{
        var miniBatchResponses = ns.backoffAll(miniBatch,retries);
      } catch(err) {
        console.error('Error sending requests '+(BATCH_SIZE*i)+' to '+(Math.min(requests.length,BATCH_SIZE*(1+i)))+': '+err);
      }
    
      if (miniBatchResponses.length){
        responses = responses.concat(miniBatchResponses)
      }
    }
    return responses;
  } // fetchTools.backoffBatches()

  ns.backoffAll = function(requests,retries){
    retries = retries || DEFAULT_RETRIES;
    var success = false;
    for (var attempt = 0; attempt <= retries; attempt++) {
      if (attempt > 0) Utilities.sleep(5000 * Math.pow(2, attempt)); //don't wait the first time
      try{
        var responses = UrlFetchApp.fetchAll(requests); 
        success = true;
      } catch (err) {
        if(/Invalid argument/.test(err.message)){
          throw err
          break;
        }
        console.error('FetchAll error, '+(retries-attempt)+' retries remain.\n'+err)
      }
      if (success) break;
    }  
    return responses
  } // fetchTools.backoffAll()
  
  return ns
})(FetchTools || {})
