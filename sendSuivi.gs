function sendSuivi(){
  var url = "www.csuivi.courrier.laposte.fr/suivi/index?id="
  
  // chercher les adresses non formattees dans 'Ventes'
  var plan_ss = SpreadsheetApp.getActiveSpreadsheet()
  var ventes_s = plan_ss.getSheetByName('Ventes')
  var ventes_data = ventes_s.getDataRange().getValues()
  var mailCol = 5
  var suiviCol = 23
  var modeleCol = 17
  var tailleCol = 16
  var nVentes = ventes_data.length
  var curMail, curSuivi, isPoste;
  for(var i=1;i<3;i++){
    curSuivi = ventes_data[i][suiviCol]
    isPoste = curSuivi.slice(0,1)
    if(isPoste==1){
      //      curMail = ventes_data[i][mailCol]
      curMail = 'clement0walter@gmail.com'
      // Send the map in an email.
      var curUrl = url+curSuivi;
      MailApp.sendEmail({
        to: curMail,
        subject: "Ton calchemise est parti !",
        htmlBody: "Eh oui, ce magnifique " +
        ventes_data[i][modeleCol] + ":" + ventes_data[i][tailleCol] +
        "court déjà, hyper à l'aise, à travers la France ! Pour le suivre : <br>" +
        curUrl + "<br>" +
        "s'il n'est pas déjà sur ton corps ! <br>" +
        "à bientôt sur calchemise.com, et Joyeux Noël bien sûr! <br> <br>" +
        "PS : en cas de commandes multiples, il y a un colis, et donc un numéro, par calchemise !"
      });
    }
  }
}