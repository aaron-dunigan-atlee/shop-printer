/**
 * Return array of cabinet objects for the job.  Used by cabinet-list.html 
 * @param {string} jobName 
 */
function getCabinets(jobName) {
  var cabinets = getRowsData(SpreadsheetApp.getActive().getSheetByName(CABINET_SHEET_NAME), null, {getMetadata: true})
    .filter(function(cabinet){
      return cabinet.job === jobName
    })
  console.log('getCabinets: Found ' + cabinets.length + ' cabinets for job ' + jobName)
  return cabinets
}


/**
 * Return array of job objects that are active.
 * @param {string} jobName 
 */
function getActiveJobs() {
  var jobs = getRowsData(SpreadsheetApp.getActive().getSheetByName(JOB_SHEET_NAME))
    .filter(function(job){
      return !job.jobCompleted
    })
  console.log('getJobs: Found ' + jobs.length + ' active jobs.')
  return jobs
}