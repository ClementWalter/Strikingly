function getPrix(str){
  var prix = 69
  switch(str) {
      // Rayure
    case'Rayure 7':prix = 45;break;
    case'Rayure 10':prix = 45;break;
    case'Rayure 11':prix = 50;break;
    case'Rayure 11b':prix = 55;break;
    case'Rayure 16':prix = 45;break;// rayure polyester
    case'Rayure 17':prix = 69;break;
    case'Rayure 18':prix = 60;break;
      
      // Carreaux
    case'Carreaux 15':prix = 50;break;
    case'Carreaux 16':prix = 69;break;
    case'Carreaux 17':prix = 69;break;
      
      // Aleatoire
    case'Pistil 1':prix = 69;break;
    case'Pois 2':prix = 50;break;
      
      // Uni
    case'Uni 1b':prix = 50;break;
    case'Uni 9':prix = 69;break;
      
      // Wax
    case'Wax 11':prix = 60;break;
    case'Wax 12':prix = 60;break;
    case'Wax 13':prix = 69;break;
    case'Wax 14':prix = 69;break;
    case'Wax 15':prix = 69;break;
    case'Wax 4d':prix = 69;break;
      
      // Liberty
    case'Liberty 16':prix = 50;break;
    case'Liberty 19':prix = 50;break;
    case'Liberty 20':prix = 69;break;
    case'Liberty 21':prix = 69;break;
  }
  
  // str = str.toUpperCase()
  //  Logger.log(str)
  //  var prix = 60
  //  if(str=="U") {prix = 55}
  //  if(str=="W") {prix = 65}
  // retirer TVA
  prix = prix/1.2
  return(prix)
}