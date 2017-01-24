function importData() {
  
  // import strikingly data
  var fi = DriveApp.getFilesByName('orders.csv')
  if ( fi.hasNext() ) { // proceed if ".csv" file exists in the reports folder
    var file = fi.next();
    var csv = file.getBlob().getDataAsString();
    var csvData = CSVToArray(csv);

    // write data indices
    var status = 0, order = 1, created = 2, name = 3, email = 4, phone = 5, address = 6, city = 7, state = 8, country = 9, zip = 10,
        products = 11, subtotal = 12, shipping = 13, coupon = 14, discount = 15, total = 17, payment = 18, customer = 19;
    
    // define some global variables
    var i = 1;
    var i_insert = 0;
    var newdata, line;
    var mod_livraison = "";

    // this variable will store the data
    var lines = [[]]
    
    while(i<(csvData.length-1)) {
      newdata = csvData[i]
      newdata[subtotal] = parseInt(newdata[subtotal].slice(1))
      newdata[shipping] = parseInt(newdata[shipping].slice(1))
      newdata[total] = parseInt(newdata[total].slice(1))
      newdata[created] = getDate(newdata[created])
      
      CA = CA + newdata[total]
      var modeles = newdata[products].split(",")
      var n_modeles = modeles.length
      var n_calch = 0;
      if(newdata[coupon]=="LIVRAISONOKLMTAVU" | newdata[coupon]=="FREESHIPPING"){
        newdata[shipping] = 0
        mod_livraison = "main propre"
      }
      for(var modele_loop = 0; modele_loop < n_modeles; modele_loop++){
        var com = getCommande(modeles[modele_loop])
        n_calch += parseInt(com[3])
      }
      for(var modele_loop = 0; modele_loop < n_modeles; modele_loop++){
        var com = getCommande(modeles[modele_loop])
        var n_com = parseInt(com[3])        
        for(var com_loop = 0; com_loop < n_com; com_loop++){
          // Le header est status, order, created, name, email, phone, address,
          // zip + " " + city, , state, country, subtotal, shipping, total, payment, customer, commande info
          line = [newdata[status], newdata[order], newdata[created], newdata[name], newdata[email], newdata[phone], newdata[address],
                  newdata[zip] + " " + newdata[city], newdata[state], newdata[country], eval(newdata[subtotal]/(n_calch)),
                    eval((newdata[total] - newdata[shipping])/n_calch),
                      eval(newdata[shipping]/(n_calch)), eval(newdata[total]/(n_calch)),
                        newdata[payment], newdata[customer]];
          line = line.concat(com.slice(0,3)).concat([mod_livraison,"1"])
          lines[i_insert + com_loop] = line
        }
        i_insert = i_insert + n_com
      }
      i = i + 1
    }
    
    // import other data
    var hors_strik_ss = SpreadsheetApp.openById('1Md7Lp7LwfR4lY95LoTThTmCBAVWgadr1up3Bi2L-I5o')
    var hors_strik_s = hors_strik_ss.getSheetByName('Orders')
    var hors_strik = hors_strik_s.getDataRange().getValues()
    var nHS = hors_strik.length
    /*    // update CA calcul
    var cur = hors_strik[1][6]
    CA = CA + hors_strik[1][10]
    for(var i = 2;i<nHS; i++){
    var newName = hors_strik[i][6]
    if(newName!=cur){
    CA = CA + hors_strik[i][10]
    cur = newName
    }
    }
    //*/
    
    // merge all data
    var all_lines = lines.concat(hors_strik.slice(1))
    all_lines.sort(function(a, b) {
      return b[2] - a[2]
    })
    
    // get number of old data
    var writeIn = 'Ventes'
    var plan_ss = SpreadsheetApp.getActiveSpreadsheet()
    var ventes_s = plan_ss.getSheetByName(writeIn)
    var ventes = ventes_s.getDataRange().getValues()
    var nVentes = ventes.length -1 //header
    
    // put new data into the spreahsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var newsheet = ss.getSheetByName(writeIn)
    newdata = csvData[0]
    var header
    var n_lines = all_lines.length
    var new_line;
    header = ['Com #', newdata[status], newdata[order], newdata[created], newdata[name], newdata[email], newdata[phone], newdata[address],
              newdata[zip] + " " + newdata[city], newdata[state], newdata[country],
                //                newdata[subtotal],
                "Prix catalogue", "Prix réel",
                  newdata[shipping], newdata[total],
                    newdata[payment], newdata[customer], "Modèle", "Taille", "Désignation", "N° Suivi/Retour", "Volume"]
                    newsheet.getRange(1, 1, 1, header.length).setValues( new Array(header) );
    var start = n_lines
    if(writeIn=='Ventes'){
      start = n_lines - nVentes
    }
    for(var lines_loop = n_lines; lines_loop>0; lines_loop--){
      var tmp = [n_lines - lines_loop + 1]
      all_lines[lines_loop-1] = tmp.concat(all_lines[lines_loop-1])
      new_line = all_lines[lines_loop-1]
      if(lines_loop<=start) {
        newsheet.getRange(n_lines - lines_loop + 2, 1, 1, header.length).setValues( new Array(new_line) );//*/
        newsheet.getRange(n_lines - lines_loop + 2, header.length + 1).setFormula(
          "=CONCATENATE(year(D" +
          eval(new_line[0]+1) +
          "), \"/\", if(MONTH(D" +
            eval(new_line[0]+1) +
              ")<10, CONCATENATE(\"0\", MONTH(D" +
                eval(new_line[0]+1) +
                ")), MONTH(D" +
                  eval(new_line[0]+1) +
                    ")))"
        );
        //  }
        /*   else{
        // newsheet.getRange(n_lines - lines_loop + 2, 1, 1, 2).setValues( new Array(new_line.slice(0,2)) );
        newsheet.getRange(n_lines - lines_loop + 2, 12, 1, 4).setValues( new Array(new_line.slice(11,15)) );
        }//*/
        //Logger.log(new_line[13].toString().replace(/\\./g,','))
        newsheet.getRange(n_lines - lines_loop + 2, 13).setFormula("=O" + eval(new_line[0]+1) + "-N" + eval(new_line[0]+1));
        newsheet.getRange(n_lines - lines_loop + 2, 14).setFormula("=IF(U" + eval(new_line[0]+1) + " = \"main propre\"; 0; " + new_line[13].toString().replace(/\\./g,',') + ")") 
      }
    }
  }
}