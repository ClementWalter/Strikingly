function faireCubyn(){
  var plan_ss = SpreadsheetApp.getActiveSpreadsheet()
  var ventes_s = plan_ss.getSheetByName('Ventes')
  var ventes_data = ventes_s.getDataRange().getValues()
  var nAdresses = ventes_data.length
  var designCol = "T".charCodeAt(0) - 65
  var tailleCol = "S".charCodeAt(0) - 65
  var nameCol = "E".charCodeAt(0) - 65
  var adresseCol = "H".charCodeAt(0) - 65
  var cpCol = "I".charCodeAt(0) - 65
  var paysCol = "K".charCodeAt(0) - 65
  //  var livreCol = "U".charCodeAt(0) - 65
  var suiviCol = "U".charCodeAt(0) - 65
  var statusCol = "B".charCodeAt(0) - 65
  var prixCol = "L".charCodeAt(0) - 65
  var refCol = "C".charCodeAt(0) - 65
  var emailCol = "F".charCodeAt(0) - 65
  var complementCol = "Q".charCodeAt(0) - 65
  var telCol = "G".charCodeAt(0) - 65
  var dateCol = "D".charCodeAt(0) - 65
  var paiementCol = "P".charCodeAt(0) - 65
  var shipCol = "N".charCodeAt(0) - 65
  var stateCol = "J".charCodeAt(0) - 65;
  var curRef, nextRef, prix, article, prenom, name, email, tel, adresse, complement, cp, ville, state, pays, iso, transporteur, colis, point, paiement, date, ship;
  var writeIn = 'Cubyn'
  var type = "Cubyn Colissimo"
  var assurance;
  var complementB = ""
  var raisonSoc = ""
  var mobile = "";
  
  var cubyn_s = plan_ss.getSheetByName(writeIn)
  var nRow = cubyn_s.getMaxRows()
  var nCol = cubyn_s.getMaxColumns()
  if(nRow>2){
    cubyn_s.getRange(3, 1, nRow-2, nCol).clearContent()
  }
  
  // get date info
  var d = new Date()
  var month = d.getMonth() + 1
  var day = d.getDate()
  if(day<10) day = "0" + day
  if(month < 10) month = "0" + month
  
  // Ask for clearing pending delivery
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Marquer les commandes comme livrées et générer le fichier .xlsx ?', ui.ButtonSet.YES_NO);
  var factureBouton = ui.alert('Générer les factures ?', ui.ButtonSet.YES_NO);
  var livreFill = [""];
  var status = "pending"
  
  // Process the user's response.
  if (response == ui.Button.YES) {
    livreFill = ["Cubyn" + "_" + d.getFullYear() + month + day]
    status = "completed"
  }
  
  // get directories info
  var livraisonFolder = DriveApp.getFoldersByName("Livraison").next()
  var invoiceFolder = DriveApp.getFolderById("0B_CI-LC3jYGIZzdRMlg3elY4SVE")
  var commandesFolder = DriveApp.getFolderById("0B_CI-LC3jYGIbmNuaTB5SXRGNUU")
  
  // Write invoices
  var facture
  var facture_s
  var factureBlob
  var facturePdf
  
  var newData;
  var i = 1
  var writeLine = 3
  var nextLivre;
  var internalCount;
  var deliveryDate = d.getFullYear() + "/" + month + "/" + day
  
  while(i<nAdresses){
    if(ventes_data[i][suiviCol]=="" & ventes_data[i][statusCol]=='pending'){
      try{
        // info colis
        curRef = ventes_data[i][refCol]
        prix = getPrix(ventes_data[i][designCol])
        article = ventes_data[i][tailleCol] + ventes_data[i][designCol].slice(0,1) + ventes_data[i][designCol].split(" ")[1]
        article = article.toUpperCase()
        // info client
        name = ventes_data[i][nameCol].split(" ")
        prenom = name[0].replace(/,/g,' ')
        name = name.slice(1, name.length).join(" ")
        adresse = ventes_data[i][adresseCol].replace(/,/g,' ')
        complement = adresse.slice(30, adresse.length)
        adresse = adresse.slice(0, 30)
        complementB = ventes_data[i][complementCol].replace(/[^a-zA-Z0-9]/g, ' ')
        cp = ventes_data[i][cpCol].split(" ")
        ville = cp.slice(1, cp.length).join(" ").replace(/,/g,' ')
        cp = cp[0]
        pays = ventes_data[i][paysCol]
        iso = getISOCountry(pays);        
        email = ventes_data[i][emailCol]
        tel = ventes_data[i][telCol]
        paiement = ventes_data[i][paiementCol]
        date = ventes_data[i][dateCol]        
        ship = ventes_data[i][shipCol]
        
        if(pays!="France"){
          if (factureBouton == ui.Button.YES) {
            // get invoice model
            facture = DriveApp.getFileById("1WaMYmGyhE4gt2sjUVcfcN8ukAbBm92_bdSuy69j-GMY").makeCopy("Invoice_"+curRef)
            facture_s = SpreadsheetApp.openById(facture.getId()).getSheetByName('Pro forma invoice')
            
            // write general data
            facture_s.getRange("G2").setValue(d.getFullYear() + "/" + month + "/" + day)
            facture_s.getRange("G3").setValue(curRef)
            facture_s.getRange("E8").setValue("Nom / Name : " + name.toUpperCase())
            facture_s.getRange("E9").setValue("Prénom / First name : " + prenom)
            facture_s.getRange("E10").setValue("Adresse / Address : " + adresse)
            facture_s.getRange("E11").setValue("Ville / City : " + ville + " ; CP / ZIP : " + cp)
            facture_s.getRange("E12").setValue("Pays / Country : " + pays)
            facture_s.getRange("E13").setValue("Téléphone / Phone : " + tel)
            facture_s.getRange("E14").setValue("Email : " + email)
            facture_s.getRange("A18").setValue("Paiement reçu le / payment received on : " + date)
            facture_s.getRange("A19").setValue("Moyen de paiement / payment method : " + paiement)
            facture_s.getRange("G39").setValue(ship)
          }
        }
        type = "Cubyn Colissimo"
        internalCount = 1
        newData = [curRef.toString(), type, prix, ""];
        newData.push(name, prenom, raisonSoc, adresse, complement, complementB, ville, cp, iso, email, tel)
        
        if((i+1)<nAdresses) {
          nextRef = ventes_data[i+1][refCol]
          nextLivre = ventes_data[i+1][suiviCol]
          while((i+1)<nAdresses & curRef==nextRef & nextLivre==""){
            ventes_s.getRange(i+1, suiviCol + 1, 1, 1).setValues(new Array(livreFill))
            ventes_s.getRange(i+1, statusCol + 1, 1, 1).setValue(status)
            if(pays!="France") {
              // add order to the invoice
              if (factureBouton == ui.Button.YES) {
                facture_s.getRange(28 + internalCount, 2, 1, 1).setValue("Calchemise")
                facture_s.getRange(28 + internalCount, 3, 1, 1).setValue(article)
                facture_s.getRange(28 + internalCount, 6, 1, 1).setValue(prix)
                facture_s.getRange(28 + internalCount, 5, 1, 1).setValue(1)
              }
            }
            internalCount = internalCount + 1
            
            // go to next line and prepare next newData
            i = i + 1
            prix = getPrix(ventes_data[i][designCol])
            newData[2] = newData[2] + prix
            article = ventes_data[i][tailleCol] + ventes_data[i][designCol].slice(0,1) + ventes_data[i][designCol].split(" ")[1]
            article = article.toUpperCase()
            
            // update nextRef and nextLivre if not at the end of the table
            if((i+1)<nAdresses) {
              nextRef = ventes_data[i+1][refCol]
              nextLivre = ventes_data[i+1][suiviCol]
            }
            
          }
        }
        cubyn_s.getRange(writeLine, 1, 1, newData.length).setValues(new Array(newData))
        cubyn_s.getRange(writeLine, 1, 1, newData.length).clearFormat()
        
        ventes_s.getRange(i+1, suiviCol + 1, 1, 1).setValues(new Array(livreFill))
        ventes_s.getRange(i+1, statusCol + 1, 1, 1).setValue(status)
        
        if(pays!="France" & (factureBouton == ui.Button.YES)) {
          // add order to the invoice
          facture_s.getRange(28 + internalCount, 2, 1, 1).setValue("Calchemise")
          facture_s.getRange(28 + internalCount, 3, 1, 1).setValue(article)
          facture_s.getRange(28 + internalCount, 6, 1, 1).setValue(newData[2])
          facture_s.getRange(28 + internalCount, 5, 1, 1).setValue(1)
          SpreadsheetApp.flush()
          factureBlob = facture.getBlob().getAs('application/pdf');
          facturePdf = invoiceFolder.createFile(factureBlob);
          commandesFolder.removeFile(facture)
        }
        
        writeLine = writeLine + 1
      }
      catch(e){
        Logger.log(e)
      }
    }
    i = i + 1
  }
  
  // Write new file
  if (response == ui.Button.YES) {
    SpreadsheetApp.flush();
    var fileName = writeIn + "_" + d.getFullYear() + month + day
    var wingFileXlsx = DriveApp.createFile(getExcel(cubyn_s, fileName))
    livraisonFolder.addFile(wingFileXlsx)
    DriveApp.removeFile(wingFileXlsx)
  }
}