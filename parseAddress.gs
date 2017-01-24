function parseAddress(){
  // mettre l'id
 // Maps.setAuthentication('xxx@gmail.com', 'key');
  // initialisation du geocoder
  var geocoder = Maps.newGeocoder()
  // chercher les adresses non formattees dans 'Ventes'
  var plan_ss = SpreadsheetApp.getActiveSpreadsheet()
  var ventes_s = plan_ss.getActiveSheet()
  var ventes_data = ventes_s.getDataRange().getValues()
  var adresse;
  var rue = 7;
  var cp = 8;
  var pays = 9;
  // col 8,9,10 = ville, CP, Pays
  var n_ventes = ventes_data.length
  Logger.log(n_ventes)
//  var i = 2
  for(var i=1;i<n_ventes;i++){
    if(ventes_data[i][cp]==" " & ventes_data[i][rue]!=""){
      Logger.log(ventes_data[i][rue])
     try{
        var response = geocoder.geocode(ventes_data[i][rue]);
        Utilities.sleep(1000)
        Logger.log(response.results[0].formatted_address)
        adresse = response.results[0].formatted_address
        adresse = adresse.split(",")
        Logger.log(adresse)
        ventes_s.getRange(i+1,rue+1,1,3).setValues(new Array(adresse))
      }
      catch(e) {
        Logger.log(e)
      }
//*/
    }
 }
}
