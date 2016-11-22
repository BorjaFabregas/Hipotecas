function calculoTAE(){
  
  //Declaracion de variables
  var email, ingresos, gastos, titulares, edad, finalidad,
      tipoVivienda, valorHipotecado, valorActual,
      importeRestante, necesidad, importeRequest, tiempo;
  var tae;
  var interesNominal, banco;
  var valorFinalHipoteca;
  
  banco = ["BBVA","ING","Cajamar","Unicaja","Santander","La Caixa","Caja Rural","EVO"];
  
  
  //Obtencion hoja de calculo
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respuestas");
  
  var row=sheet.getLastRow();
  
  //Obtencion de valores para cada variable
  email = sheet.getRange("B"+row).getValue();
  titulares = sheet.getRange("C"+row).getValue();
  ingresos = sheet.getRange("D"+row).getValue();
  gastos = sheet.getRange("E"+row).getValue();
  edad = sheet.getRange("F"+row).getValue();
  finalidad = sheet.getRange("G"+row).getValue();
  importeRequest = valorActual = sheet.getRange("O"+row).getValue();
  tiempo = valorActual = sheet.getRange("P"+row).getValue();
  switch(finalidad){
    case "Comprar una vivienda":
      
      var sheetTAE=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Comprar una vivienda");
      var interesNominal = sheetTAE.getRange(1,2,9,1).getValues();
      
      tipoVivienda = sheet.getRange("H"+row).getValue();
      if(tipoVivienda == "Primera vivienda"){
        valorHipotecado = sheet.getRange("I"+row).getValue();
      }else{
        //Si es segunda vivienda comision de 5000€
        valorHipotecado = sheet.getRange("I"+row).getValue();
        valorHipotecado+=5000;
      }
      
      
      //Evaluacion TAE
      var i = 0;
      do{
        tae = Math.pow((1+(interesNominal[i]/tiempo)), tiempo)-1;
        var bancoTAE = banco[i];
        i++;
      }while(interesNominal[i]<interesNominal[i-1]&&i<interesNominal.length);
      
      valorFinalHipoteca = (valorHipotecado-importeRequest)*tae;
      
      Logger.log("Banco: " + bancoTAE + "\tValor Hipoteca: " + valorFinalHipoteca + "\tInteres Nominal: " + interesNominal[i]); 
      break;
      
    case "Cambiar mi hipoteca":
      
      var sheetTAE=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cambiar hipoteca");
      var interesNominal = sheetTAE.getRange(1,2,9,1).getValues();
      
      tipoVivienda = sheet.getRange("J"+row).getValue();
      if(tipoVivienda == "Primera vivienda"){
        valorActual = sheet.getRange("K"+row).getValue();  
      }else{
        //Si es segunda vivienda comision de 5000€
        valorActual = sheet.getRange("K"+row).getValue();
        valorActual+=5000;
      }     
      importeRestante = sheet.getRange("L"+row).getValue();
      
      
      //Evaluacion TAE
      var i = 0;
      do{
        tae = Math.pow((1+(interesNominal[i]/tiempo)), tiempo)-1;
        var bancoTAE = banco[i];
        i++;
      }while(interesNominal[i]<interesNominal[i-1]&&i<interesNominal.length);
      
      valorFinalHipoteca = (valorActual-importeRequest)*tae;
      
      Logger.log("Banco: " + bancoTAE + "\tValor Hipoteca: " + valorFinalHipoteca + "\tInteres Nominal: " + interesNominal[i]); 
      break;
      
    case "Hipotecar mi casa":
      
      var sheetTAE=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hipotecar mi casa");
      var interesNominal = sheetTAE.getRange(1,2,9,1).getValues();
      
      necesidad = sheet.getRange("M"+row).getValue();
      valorActual = sheet.getRange("N"+row).getValue();
      
      
      //Evaluacion TAE
      var i = 0;
      do{
        tae = Math.pow((1+(interesNominal[i]/tiempo)), tiempo)-1;
        var bancoTAE = banco[i];
        i++;
      }while(interesNominal[i]<interesNominal[i-1]&&i<interesNominal.length);
      
      valorFinalHipoteca = (valorActual-importeRequest)*tae;
      Logger.log("Banco: " + bancoTAE + "\tValor Hipoteca: " + valorFinalHipoteca + "\tInteres Nominal: " + interesNominal[i]); 
  }
  
  //Obtener fecha
  var hoy = new Date();
  var mes = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "MM");
  var dia = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd");
  var anio = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "YYYY");
  var fecha = dia + "/" + mes + "/" + anio + "\n";
  
  //Elaboracion E-mail
  
  var subject = "Rastreador de Hipotecas";
  var cabecera = "Estimado Cliente\n";
  var body = fecha + cabecera;
  body = body + "\nDe acuerdo con las condiciones que Ud. nos indicó en el formulario, " +
    "le informamos de que la mejor opcion para su hipoteca es la de " + bancoTAE + ", " +
      "con un Interes Nominal de " + interesNominal[i] + "%, " +
        " y un Valor Final de Hipoteca de " + valorFinalHipoteca + " €";
  body = body + "\n\n";
  body = body + "Tambien, le ofrecemos los siguientes planes segun nuestra Base de Datos: \n\n";
  for(i = 1;i <= 7;i++){
    body = body + "Banco: " + banco[i] +"\t\tInteres Nominal: " + interesNominal[i+1] + "%\n";
  }
  body = body + "\nUn Saludo";
  
  Logger.log(body);
  
  GmailApp.sendEmail(email, subject, body);
  
}