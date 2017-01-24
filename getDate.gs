function getDate(str){
  var patternDate = /^(\d{4})\-(\d{1,2})\-(\d{1,2})/
  var arrayDate = str.match(patternDate);
  var date = new Date(arrayDate[1], arrayDate[2] - 1, arrayDate[3]);
  return(date)
}