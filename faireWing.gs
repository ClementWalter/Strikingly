function faireWing(){
  var plan_ss = SpreadsheetApp.getActiveSpreadsheet()
  var ventes_s = plan_ss.getSheetByName('Ventes')
  var ventes_data = ventes_s.getDataRange().getValues()
  var nAdresses = ventes_data.length
  var designCol = "S".charCodeAt(0) - 65
  var tailleCol = "R".charCodeAt(0) - 65
  var nameCol = "E".charCodeAt(0) - 65
  var adresseCol = "H".charCodeAt(0) - 65
  var cpCol = "I".charCodeAt(0) - 65
  var paysCol = "K".charCodeAt(0) - 65
  var livreCol = "U".charCodeAt(0) - 65
  var suiviCol = "T".charCodeAt(0) - 65
  var statusCol = "B".charCodeAt(0) - 65
  var prixCol = "L".charCodeAt(0) - 65
  var refCol = "C".charCodeAt(0) - 65
  var emailCol = "F".charCodeAt(0) - 65
  var complementCol = "P".charCodeAt(0) - 65
  var telCol = "G".charCodeAt(0) - 65
  var dateCol = "D".charCodeAt(0) - 65
  var paiementCol = "O".charCodeAt(0) - 65
  var shipCol = "M".charCodeAt(0) - 65
  var stateCol = "J".charCodeAt(0) - 65;
  var curRef, nextRef, prix, article, prenom, name, email, tel, adresse, complement, cp, ville, state, pays, iso, transporteur, type, colis, point, paiement, date, ship;
  var writeIn = 'Wing'
  var assurance = ""
  var complementB = ""
  var raisonSoc = ""
  var quantite = 1;
  var mobile = "";

  transporteur = "La Poste"
  colis = "Colissimo "
  point = ""
  
  var wing_s = plan_ss.getSheetByName(writeIn)
  wing_s.clearContents()
  
  // Ask for clearing pending delivery
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Marquer les commandes comme livrées et générer le fichier .csv ?', ui.ButtonSet.YES_NO);
  var factureBouton = ui.alert('Générer les factures ?', ui.ButtonSet.YES_NO);
  var livreFill = ["",""];
  // Process the user's response.
  var d = new Date()
  var month = d.getMonth() + 1
  var day = d.getDate()
  if(day<10) day = "0" + day
  if(month < 10) month = "0" + month
  if (response == ui.Button.YES) {
    livreFill = ["Wing" + "_" + d.getFullYear() + month + day,"1"]
  }
  
  var livraisonFolder = DriveApp.getFoldersByName("Livraison").next()
  var invoiceFolder = DriveApp.getFolderById("0B_CI-LC3jYGIZzdRMlg3elY4SVE")
  var commandesFolder = DriveApp.getFolderById("0B_CI-LC3jYGIbmNuaTB5SXRGNUU")
  
  // Write in a new file
  var csvStringFR, csvString;
  
  // Write invoices
  var facture
  var facture_s
  var factureBlob
  var facturePdf
  
  // Make header
//  var header = ["id", "Ref Commande", "Code Logistique", "Assurance", "Prix", "Articles", "Prenom", "Nom", "Raison Sociale", "Email", "Telephone", "Adresse", "Complement", "Complement Bis", "Code Postal", "Ville", "Pays", "Transporteur", "Type de colis", "Point Relais"]
 /* var header = ["store",
                "orderReference",
                "logisticId", "insurance", "itemPrice", "itemDesignation", "itemSku", "shCode", "itemQty", "itemWeight", "itemLength", "itemWidth", "itemHeight",
                "firstName", "lastName", "company", "email", "phone", "mobile", "addressLine1", "addressLine2", "additional", "zip", "city", "state", "country",
                "transporter", "service", "relayPoint", "return", "ftd", "machinable", "cod", "codValue"];//*/
    var header = ["Order reference*", "Item Name*", "Item quantity", "Item price", "Item insurance",
                  "First Name", "Last Name*", "Company", "Email", "Phone", "Address Line 1*", "Address Line 2", "Additional address details", "Postcode*", "City*", "State / Region", "Country* (ISO code)",
                  "Carrier and Service*", "Relay Point"];
  
  wing_s.getRange(1, 1, 1, header.length).setValues( new Array(header) );
  csvStringFR = header.join(";")
  csvString = header.join(",")
  
  var newData;
  var id = 0;
  var codeLog = 0;
  var i = 1
  var writeLine = 2
  var nextLivre;
  var internalCount;
  var deliveryDate = d.getFullYear() + "/" + month + "/" + day
  
  while(i<nAdresses){
    if(ventes_data[i][suiviCol]=="" & ventes_data[i][statusCol]=='pending'){
      try{
        curRef = ventes_data[i][refCol]
        codeLog = codeLog + 1
        prix = getPrix(ventes_data[i][designCol])
        article = ventes_data[i][tailleCol] + ventes_data[i][designCol].slice(0,1) + ventes_data[i][designCol].split(" ")[1]
        article = article.toUpperCase()
        name = ventes_data[i][nameCol].split(" ")
        prenom = name[0].replace(/,/g,' ')
        name = name.slice(1, name.length).join(" ")
        email = ventes_data[i][emailCol]
        tel = ventes_data[i][telCol]
        adresse = ventes_data[i][adresseCol].replace(/,/g,' ')
        complement = adresse.slice(30, adresse.length)
        adresse = adresse.slice(0, 30)
        complementB = ventes_data[i][complementCol].replace(/[^a-zA-Z0-9]/g, ' ')
        cp = ventes_data[i][cpCol].split(" ")
        ville = cp.slice(1, cp.length).join(" ").replace(/,/g,' ')
        cp = cp[0]
        state = ventes_data[i][stateCol]
        if(state=="undefined") state="";
        pays = ventes_data[i][paysCol]
        iso = getISOCountry(pays);
        paiement = ventes_data[i][paiementCol]
        date = ventes_data[i][dateCol]
        ship = ventes_data[i][shipCol]

        if(pays!="France"){
          type = colis + "International"
        //  point = "Non"
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
        } else {
          type = colis + "Domicile Sans Signature"
         // point = "Oui"
        }
        internalCount = 1
        /*
        newData = [writeLine - 1, curRef, codeLog, assurance, prix, article]
        newData.push(prenom, name, raisonSoc, email, tel, adresse, complement, complementB, cp, ville, pays)
        newData.push(transporteur, type, point)
        
        newData = ["", curRef, "", assurance, prix, article, "", "", quantite, "", "", "", ""]
        newData.push(prenom, name, raisonSoc, email, tel, mobile, adresse, complement, complementB, cp, ville, state, pays)
        newData.push(transporteur, type, point)
        //*/
        newData = [curRef, article, quantite, prix, assurance];
        newData.push(prenom, name, raisonSoc, email, tel, adresse, complement, complementB, cp, ville, state, iso)
        newData.push(transporteur + " - " + type, point)
        
        if((i+1)<nAdresses) {
          nextRef = ventes_data[i+1][refCol]
          nextLivre = ventes_data[i+1][suiviCol]
          while((i+1)<nAdresses & curRef==nextRef & nextLivre==""){
            // modify newData entry
           // newData[2] = codeLog + "_" + internalCount
            // write newData
            wing_s.getRange(writeLine, 1, 1, newData.length).setValues(new Array(newData))
            ventes_s.getRange(i+1, livreCol, 1, 2).setValues(new Array(livreFill))
            csvStringFR = csvStringFR + "\n" + newData.join(";")
            csvString = csvString + "\n" + newData.join(",")
            if(pays!="France") {
              // add order to the invoice
              if (factureBouton == ui.Button.YES) {
                facture_s.getRange(28 + internalCount, 2, 1, 1).setValue("Calchemise")
                facture_s.getRange(28 + internalCount, 3, 1, 1).setValue(newData[1])
                facture_s.getRange(28 + internalCount, 6, 1, 1).setValue(newData[3])
                facture_s.getRange(28 + internalCount, 5, 1, 1).setValue(1)
              }
            }
            internalCount = internalCount + 1
            
            // go to next line and prepare next newData
            i = i + 1
            writeLine = writeLine + 1
          //  newData[0] = writeLine - 1
            article = ventes_data[i][tailleCol] + ventes_data[i][designCol].slice(0,1) + ventes_data[i][designCol].split(" ")[1]
            article = article.toUpperCase()
            newData[1] = article
            newData[3] = getPrix(ventes_data[i][designCol])
            // update nextRef and nextLivre if not at the end of the table
            if((i+1)<nAdresses) {
              nextRef = ventes_data[i+1][refCol]
              nextLivre = ventes_data[i+1][suiviCol]
            }
            // if end of while loop
    /*        if(curRef!=nextRef){
              newData[2] = codeLog + "_" + internalCount
            }//*/
          }
        }
        wing_s.getRange(writeLine, 1, 1, newData.length).setValues(new Array(newData))
        
        
        ventes_s.getRange(i+1, livreCol, 1, 2).setValues(new Array(livreFill))
        csvStringFR = csvStringFR + "\n" + newData.join(";")
        csvString = csvString + "\n" + newData.join(",")
        if(pays!="France" & (factureBouton == ui.Button.YES)) {
          // add order to the invoice
          facture_s.getRange(28 + internalCount, 2, 1, 1).setValue("Calchemise")
          facture_s.getRange(28 + internalCount, 3, 1, 1).setValue(newData[1])
          facture_s.getRange(28 + internalCount, 6, 1, 1).setValue(newData[3])
          facture_s.getRange(28 + internalCount, 5, 1, 1).setValue(1)
          SpreadsheetApp.flush()
          factureBlob = facture.getBlob().getAs('application/pdf');
          facturePdf = invoiceFolder.createFile(factureBlob);
          commandesFolder.removeFile(facture)
        }
        
        writeLine = writeLine + 1
  //      newData[0] = writeLine - 1
      }
      catch(e){
        Logger.log(e)
      }
    }
    i = i + 1
  }
  
  // Write csvStringFR into a new file
  if (response == ui.Button.YES) {
    SpreadsheetApp.flush();
    var fileName = writeIn + "_" + d.getFullYear() + month + day
    var wingFileCSV = DriveApp.createFile(fileName +".csv", csvStringFR)
    var wingFileXlsx = DriveApp.createFile(getExcel(wing_s, fileName))
    livraisonFolder.addFile(wingFileCSV)
    livraisonFolder.addFile(wingFileXlsx)
    DriveApp.removeFile(wingFileCSV)
    DriveApp.removeFile(wingFileXlsx)
  }
}