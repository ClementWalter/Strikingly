function getCommande(str){
  str = str.slice(11); // retirer 'Calchemise '
  str = str.split(":"); // séparer modèle et taille
  if(str.length==1){
    str = str.concat("no size*1")
  }
  var modele = str[0];
  modele = modele.split(" ")
  var nMod = modele.length
  modele[nMod - 1] = modele[nMod-1].split("*")[0]
  var designation = modele.join(" ")
  if(modele[1]=="carreaux"){
    designation = modele[1] + " " + modele[2]
  }
  if(modele[1]=="pois"){
    designation = modele[1] + " " + modele[2]
  }
  if(modele[0].slice(0,3)=="ray"){
    designation = "rayure " + modele[1]
  }
  if(modele[1]=="africain"){
    designation = "wax " + modele[2]
  }
  if(modele[0]=="wax"){
    designation = "wax " + modele[1]
  }
  if(modele[1]=="Noël"){
    designation = "noel " + modele[2]
  }
  if(modele[1]=="Noel"){
    designation = "noel " + modele[2]
  }  
  modele.pop()
  modele = modele.join(" ")
  var taille = str[1];
  var n_com = taille[taille.length-1];
  taille = taille.split("*");
  taille = taille[0]
  var com = [modele, taille, designation, n_com]
  return(com)
}