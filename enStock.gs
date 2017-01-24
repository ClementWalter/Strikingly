function enStock(){
  // import other data
  var plan_ss = SpreadsheetApp.getActiveSpreadsheet()
  var attribution_s = plan_ss.getSheetByName('Attribution')
  var attribution_data = attribution_s.getDataRange().getValues()
  // S : 10 ; M : 11 ; L : 12 ; X : 13
  
  var ventes_s = plan_ss.getSheetByName('Ventes')
  var ventes_data = ventes_s.getDataRange().getValues()
  // design : 17 ; taille : 16
  var designCol = 17
  var tailleCol = 16
  
  var n_lines = ventes_data.length
  var enStock_col = 21
  for(var i = 1;i<n_lines;i++){
    // look for right line
    var look = 2
    var design = attribution_data[look][1].toLowerCase();
    ventes_data[i][designCol] = ventes_data[i][designCol].toLowerCase()
 //   Logger.log(ventes_data[i][designCol])
    if(ventes_data[i][designCol]!=""){
      while(design!=ventes_data[i][designCol]){
        look = look + 1
        design = attribution_data[look][1].toLowerCase();
      }
      
      // look for right size
      var size = 10
      while(attribution_data[1][size]!=ventes_data[i][tailleCol]){
        size = size + 1
      }
      Logger.log(attribution_data[look][size])
      ventes_s.getRange(i+1,enStock_col,1,1).setValue(attribution_data[look][size])
    }
  }
}