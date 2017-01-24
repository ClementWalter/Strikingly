function getMap(){
  // Add markers to the map.
  var center = 'Paris france'
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Entrer code postal :');
  
  var arrondissements = response.getResponseText().split(",")
  var nArr = arrondissements.length
  for(var arrInd=0; arrInd<nArr; arrInd++){
    var arrondissement = parseInt(arrondissements[arrInd])
    var mapInfo = ['Id', 'Nom', 'Adresse', 'CP', 'Commande']
    var map = Maps.newStaticMap()       
    .setSize(10000, 10000)
    .setCenter(arrondissements[arrInd] + "arrondissement Paris")
    .setZoom(16)
    
    // Add markers for the nearbye train stations.
    map.setMarkerStyle(Maps.StaticMap.MarkerSize.MID, Maps.StaticMap.Color.BLUE, '0');  
    map.addMarker(center)
    map.setMarkerStyle(Maps.StaticMap.MarkerSize.MID, Maps.StaticMap.Color.RED, 'T');
    // get address
    var plan_ss = SpreadsheetApp.getActiveSpreadsheet()
    var livraison_s = plan_ss.getSheetByName('Etiquettes agrégées')
    var livraison_data = livraison_s.getDataRange().getValues()
    var mapInfo_ss = SpreadsheetApp.create('Info ' + arrondissement)
    var mapInfo_s = mapInfo_ss.getSheets()[0]
    mapInfo_s.getRange(1,1,1,5).setValues(new Array(mapInfo))
    var ref,name,address, cp, commandeId;
    var inc = 1;
    var n_data = livraison_data.length/5
    for(var i=0;i<n_data;i++){
      ref = livraison_data[5*i][0].toUpperCase()
      name = livraison_data[5*i + 1][0]
      address = livraison_data[5*i + 2][0]
      cp = livraison_data[5*i + 3][0].slice(0,6)
      try{
        cp = parseInt(cp)
      }
      catch(e) {
        cp = 0
      }
      if (cp === arrondissement) {
        commandeId = columnToLetter(inc);inc = inc + 1
        mapInfo = [commandeId,name,address,cp,ref]
        address = address + cp
        map.setMarkerStyle(Maps.StaticMap.MarkerSize.MID, Maps.StaticMap.Color.RED, commandeId);
        map.addMarker(address)
        Logger.log(new Array(mapInfo))
        mapInfo_s.getRange(inc,1,1,5).setValues(new Array(mapInfo))
      }
    }
    
    //  Logger.log(map.getMapUrl())
    DriveApp.createFile(Utilities.newBlob(map.getMapImage(), 'image/png', arrondissement + '.png'))
    var files = DriveApp.getFilesByName(arrondissement + '.png')
    var file = files.next()
    var calFolder = DriveApp.getFoldersByName('Cartes livraison')
    var mapsFolder = calFolder.next()
    file.makeCopy(mapsFolder)
    DriveApp.removeFile(file)
    
    files = DriveApp.getFilesByName('Info ' + arrondissement)
    file = files.next()
    Logger.log(file)
    mapsFolder.addFile(file)
 //   file.makeCopy('Info ' + arrondissement, mapsFolder)
    DriveApp.removeFile(file)
  }
  
  /*  // Send the map in an email.
  var toAddress = Session.getActiveUser().getEmail();
  MailApp.sendEmail(toAddress, arrondissement, 'Please open: ' + map.getMapUrl(), {
  htmlBody: 'See below.<br/><img src="cid:mapImage">',
  inlineImages: {
  mapImage: Utilities.newBlob(map.getMapImage(), 'image/png')
  }
  });
  //*/
}

function columnToLetter(column)
{
  var temp, letter = '';
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  return letter;
}