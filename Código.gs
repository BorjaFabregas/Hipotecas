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
      tipoVivienda = sheet.getRange("H"+row).getValue();
      if(tipoVivienda == "Primera vivienda"){
        valorHipotecado = sheet.getRange("I"+row).getValue();
      }else{
        //Si es segunda vivienda comision de 5000€
        valorHipotecado = sheet.getRange("I"+row).getValue();
        valorHipotecado+=5000;
      }
      
      interesNominal = [7.3,1.5,1.51,2.22,2.67,2.79,3.2,4.5];
      //Evaluacion TAE
      var i = 0;
      do{
        tae = Math.pow((1+(interesNominal[i]/tiempo)), tiempo)-1;
        var bancoTAE = banco[i];
        i++;
      }while(interesNominal[i]<interesNominal[i-1]&&i<interesNominal.length);
      
      valorFinalHipoteca = (valorHipotecado-importeRequest)*tae;
        
      break;
    case "Cambiar mi hipoteca":
      tipoVivienda = sheet.getRange("J"+row).getValue();
      if(tipoVivienda == "Primera vivienda"){
        valorActual = sheet.getRange("K"+row).getValue();  
      }else{
        //Si es segunda vivienda comision de 5000€
        valorActual = sheet.getRange("K"+row).getValue();
        valorActual+=5000;
      }     
      importeRestante = sheet.getRange("L"+row).getValue();
      
      interesNominal = [8.2,5.5,3.4,5.3,6.7,6.5,4.9,4.5];
      //Evaluacion TAE
      var i = 0;
      do{
        tae = Math.pow((1+(interesNominal[i]/tiempo)), tiempo)-1;
        var bancoTAE = banco[i];
        i++;
      }while(interesNominal[i]<interesNominal[i-1]&&i<interesNominal.length);
      
      valorFinalHipoteca = (valorActual-importeRequest)*tae;
      
      break;
    case "Hipotecar mi casa":
      necesidad = sheet.getRange("M"+row).getValue();
      valorActual = sheet.getRange("N"+row).getValue();
      
      interesNominal = [8.9,6.7,1.51,2.22,2.99,3.4,3.2,4.5];
      //Evaluacion TAE
      var i = 0;
      do{
        tae = Math.pow((1+(interesNominal[i]/tiempo)), tiempo)-1;
        var bancoTAE = banco[i];
        i++;
      }while(interesNominal[i]<interesNominal[i-1]&&i<interesNominal.length);
      
      valorFinalHipoteca = (valorActual-importeRequest)*tae;
  }
  
  
  
  
  
  /*
  var hoy = new Date();
  var mes = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "MM");
  var dia = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd");
  var anio = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "YYYY");
  var fecha = dia + "/" + mes + "/" + anio;*/
  
  /*Elaboracion E-mail
  var subject = "Resolucion de Solicitud de Linea de Credito";
  var cabecera = "\n\n" + nombre + " " + apellidos  + "\n" + dir + "\n" + prov + "\t" + cp + "\n\n\n";
  var body = fecha + cabecera;
  body = body + "Estimado " + nombre + " con esta carta se le hace de su conocimiento de que su línea de credito ha sido " + credito;
  body = body + "\n\n";
  if(credito == "aceptada"){
    body = body + "En breve nos pondremos en contacto con usted.\n\nUn cordial saludo."
  }else{
    body = body + "Sin más por el momento nos despedimos."
  }  
  Logger.log(body);
  GmailApp.sendEmail(email, subject, body);*/
  
}