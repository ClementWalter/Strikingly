function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Ventes').
  addItem('Télech. orders', 'openDialog').
  addItem('Update ventes', 'importData').
  addItem('Update stocks', 'enStock').
  addItem('Parse adresses', 'parseAddress').
  addItem('Faire étiquettes', 'faireEtiquette').
  addItem('Créer carte', 'getMap').
  addItem('Créer ordre Wing', 'faireWing').
  addItem('Créer ordre Cubyn', 'faireCubyn').
  addItem('Générer une facture', 'getFacture').
  addToUi();
}

function onEdit(event){
  enStock()
}