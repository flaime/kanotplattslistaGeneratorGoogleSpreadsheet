//Av: Linus Ahlin-Hamberg 20016

/*


b�sta f�rb�ttringarna jag kan se att man b�r/kan g�ra �r att �ndra s� att den inte anropar getValue s� mycket p� raderna kring 106 d� det blir ett anrop per styck till exelarket skulle vara b�ttre att fr�ga en eller n�gon f� g�ng f�r nu �r det tillochmed flera lopar s� blir m�nga g�nger...
Sedan raderar den bara ett "fast" antal rader ca 500 stycken n�r man "rensar" �r ju inte heller s� bra... (men om man inte bygger ut det mycket s� kommer det inte beh�vas �ndras eller vara ett problem) D� det nu "bara" finns va 300 plattser/rader som fylls...


kanothuset
K1
B4:E17

K2
G4:P17

K4
AC4:AF17
K5
AH6:AQ17
K6
AS6:BC17
K7
BD6:BG17
K8
BI6:BR17
*/

var spredshetID = "1x0-VV1nx_2GryMArovs2S49CJxIZdP-N7C1Up9mW7HM";
var tabbelNumer = 1088317021;
var namnPaListfliken = "Lista";
var namnPaKanothusFliken = "kanothuset kanotplatser";
var namnPaLadaOchUtbygnad = "Ladan + utbygnad kanotplatser";

function minListFunction(){
  
  var ss = getSpreadsheetListanKorekt(); //getActiveSpreadsheet();
  

  
  
  var ss = SpreadsheetApp.openById(spredshetID);//getActiveSpreadsheet();
  var ssKHus = ss.getSheets()[0];
  
  
  
  
  var dorrar = ['B4:E17','G4:P17','R4:AA17','AC4:AF17', 'AH6:AQ17', 'AS6:BC17', 'BD6:BG17', 'BI6:BR17'];
  var dorrarLadanUtbygnad = ['BA7:BF20','AR7:AW20', 'AK7:AP20', 'AD7:AI20','W7:AB20','P7:U20','I7:N20', 'B7:G20'];
  var rader = [];
  var nastarad = 0;
  liveLogga("startar listskapandet");
  //g� igenom alla d�rrar en f�r en f�r kanothuset
  for(var i = 0; i < dorrar.length; i++){
    var plattserTillDorren = plockaUtPlattser(ssKHus.getRange(dorrar[i]));
    rader = rader.concat(plattserTillDorren);
  }
  liveLogga("Har l�st in alla kanotplattser i kanothuset");
  //g�r igenom ladan samt utbygnaden
  var ssLadanUtutbygnad = getSpreadsheetLadaOchUtbygnadKorekt();
  for(var i = 0; i < dorrarLadanUtbygnad.length; i++){
    var plattserTillDorren = plockaUtPlattser(ssLadanUtutbygnad.getRange(dorrarLadanUtbygnad[i]));
    rader = rader.concat(plattserTillDorren);
  }
  liveLogga("Har l�st in alla kanotplattser i ladan samt utbygnaden");
  
  _skrivUtLista(rader);
  
  
  rensaLiveloggenMenVantaAntalSecInnan(5);
}

function _skrivUtLista(rader){
  liveLogga("B�rjar skriva ut alla kanotplattser");  
  var tal = rader.length;
  liveLogga("Det finns just nu " + tal + " antal utskrivna kanotplattser");
  getSpreadsheetListanKorekt().getRange(5,2,tal,3).setValues(rader);
  /*var vilkenRad = 5;
  var ss =getSpreadsheetListanKorekt();
  Logger.log("fj�rde klar");
  liveLogga("fj�rde klar");
  for(var i = 0; i < rader.length; i++){
     //B5 �r f�rsta rutan
    skrivTillCell("B" + vilkenRad, rader[i][0]);
    skrivTillCell("C" + vilkenRad, rader[i][1]);
    skrivTillCell("D" + vilkenRad, rader[i][2]);

    vilkenRad = vilkenRad +1;
    
  }*/
  liveLogga("Allt klart");
  liveLogga("Vars�god");
}



var platts = 1;
function plockaUtPlattser(enDorr){
  var plattser = [];
  var x =2;
  var y = 2;
  platts = 0;
  for(i = 0; i < antalPlattserPaOmrader(enDorr); i++){
    var rad = [];
    rad[0] = enDorr.getCell(x-1,y-1).getValue()+""; //plattsnamnet
    rad[1] = enDorr.getCell(x,y-1).getValue()+""; //namnet p� personen
    rad[2] = enDorr.getCell(x,y).getValue() +""; //komentaren
    //Logger.log(rad);
    if(inehallerSaker(rad)==true){
      plattser[platts] = rad;
      platts = platts + 1;
    }
    
    if(enDorr.getHeight() >= x+2){
         x = x+2;
      
    }else if(enDorr.getWidth() >= y+2){
      x=2;
      y = y+2;
    }
    //Logger.log("Rad = " + enDorr.getCell(1,1).getValue());
  }
  return plattser;
}

function antalPlattserPaOmrader(enDorr){
  return (enDorr.getHeight() * enDorr.getWidth())/4
  
}

function inehallerSaker(platts){
  if(platts[0].length > 0){
   return true;
  }else if(platts[1].length > 0){
   return true;
  }if(platts[2].length > 0){
   return true;
  }
  return false;
}


function skrivTillCell(cellnamn, vadAttSkriva){
  getSpreadsheetListanKorekt().getRange(cellnamn).setValue(vadAttSkriva);
}

function DOUBLE(input) {
  return input * 2;
}



function rensa(){
  liveLogga("B�rjar att rensa listan");
  getSpreadsheetListanKorekt().getRange('B5:D507').clearContent();
  liveLogga("Listan �r nu rensad");
  liveLogga("Vars�god");
  rensaLiveloggenMenVantaAntalSecInnan(5);
}



function getSpreadsheetListanKorekt(){
  var ss = SpreadsheetApp.openById(spredshetID);
  return ss.getSheetByName(namnPaListfliken);
}

function getSpreadsheetKanothusetKorekt(){
  var ss = SpreadsheetApp.openById(spredshetID);
  return ss.getSheetByName(namnPaKanothusFliken);
}

function getSpreadsheetLadaOchUtbygnadKorekt(){
  var ss = SpreadsheetApp.openById(spredshetID);
  return ss.getSheetByName(namnPaLadaOchUtbygnad);
  
}



function skrivaUt(){
  skrivTillCell("B6","testrar");
  
  
}
var nummer = 8;
function liveLogga(attLogga){
  
  skrivTillCell("F"+nummer,attLogga);
  nummer = nummer +1;
}

function rensaLiveloggenMenVantaAntalSecInnan(sec){
 Utilities.sleep(sec * 1000);
  rensaLiveloggen();
}
function rensaLiveloggen(){
  liveLogga("");
  getSpreadsheetListanKorekt().getRange("F8:F" + nummer).clearContent();
  nummer = 8;
}
