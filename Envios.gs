// MENU propio
function onOpen() {
    console.log("===== [onOpen] Run ======");
    
    var Ui = SpreadsheetApp.getUi();
    Ui.createMenu('<< Hoakeen: Bienvenido ENVIOS >>')
    .addSubMenu((Ui.createMenu('Estas en la cuenta correcta?')
                 .addItem('1 - ENVIAR EMAILS', 'emailEnvios')
                 .addSeparator()                
                ))
    .addSeparator()
    // .addSubMenu((Ui.createMenu('[LISTAS - SISTEMA]')
    //              .addItem('9 - LISTA CLASES (TODAS)', 'CLASSES_LIST')    
    //              .addSeparator()
    //              .addItem('9 - LISTA CLASES (Implementacion TODAS)', 'CLASSES_LIST_IMPLEMENTACION_all')    
    //              .addSeparator()
    //           ))
    
    .addToUi();
  }

function emailEnvios() {
  var fila = 0;
  var sede = 1;
  var nombreAlumno = 2;
  var userLCB = 7;
  var password = 8;
  var classroomName = 9;
  var linkMeet = 10;
  var classroomLink = 11;
  var emailReply = 13;
  var emailsEnvio = 14;
  // var emailModelo = 18;
  var escribirResultado = 23 // empieza por 1


// Ventana preguntando si quieres correr el script
var ui = SpreadsheetApp.getUi();
var response = ui.alert('Estas en la cuenta correcta? (bienvenidos.noreply@)???? ',
ui.ButtonSet.YES_NO);
console.log("Esperando confirmación en planilla...");
var usuarioCorrecto = SpreadsheetApp.getUi().prompt('Ingresa el usuario correcto');
 // Process the user's response.
 if (response == ui.Button.YES) {
      console.log("Todo Correcto, corremos script");
      var logArray =[];
      //logArray.push([ "Sede", "Nombre Alumno", "Usuario", "Contraseña", "Meet Link", "gClassroom","Class Link", "ENVIADO", "ENVIADO A:"]);
      console.log(logArray);

      var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ENVIOS");
      var data = ws.getRange("A3:W" + ws.getLastRow()).getValues(); 
      
      // DATA Envios NEW
      dataNEW = data.filter(function(r) { return r[19] == "NEW" && r[20] == "SI" && r[21] == "PERMITIDO" && r[22] != "ENVIADO" });
      console.log("NEW");
      console.log(dataNEW);

      if (dataNEW && dataNEW.length) {
        
          // Decido que plantilla/Templates de emails HTML uso
          var emailTemp = HtmlService.createTemplateFromFile("emailsNuevos");
          var asuntoEmailNuevo = "¡Bienvenido! Ya podés acceder a tu aula interactiva del Liceo 2021"

              dataNEW.forEach(function(row){
              console.log("Siguiente Registro NEW: ");
              emailTemp.fila = row[fila];
              emailTemp.sede = row[sede];
              emailTemp.nombreAlumno = row[nombreAlumno];
              emailTemp.userLCB = row[userLCB];
              emailTemp.password = row[password];
              emailTemp.linkMeet = row[linkMeet];
              emailTemp.classroomName = row[classroomName];    
              emailTemp.classroomLink = row[classroomLink];
              emailTemp.emailsEnvio = row[emailsEnvio];
              emailTemp.emailReply = row[emailReply];
              emailTemp.linkFaq = 'https://docs.google.com/document/d/1-xOYcEuuFb8mMw1P72xc774V2YFPuNISPbR_LwJHRpg/';
              emailTemp.linkTutorial = 'https://drive.google.com/file/d/1EHUk4bv60ilsvvm8LhiU10sXkTrQidXR/view?usp=sharing';
              emailTemp.linkCalendar = 'https://calendar.google.com/';
              
              // console.log(emailTemp.nombreAlumno)
              // console.log("Fila: "+emailTemp.fila);

            try {   
              var htmlMessage = emailTemp.evaluate().getContent();

              MailApp.sendEmail({
                to: emailTemp.emailsEnvio,
                replyTo: emailTemp.emailReply, 
                cc: emailTemp.userLCB,
                bcc: emailTemp.emailReply,
                subject: asuntoEmailNuevo,
                htmlBody: htmlMessage,
                body: "Si no puede ver el contenido de este correo, por favor comunícate con la secretaría de tu sede del Liceo Británico.",
              })

              console.log("Email enviado")   
              var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
              ss.getRange(emailTemp.fila, escribirResultado).setValue('ENVIADO');

              var newArray = [emailTemp.sede, emailTemp.userLCB, emailTemp.password, emailTemp.classroomName ,emailTemp.classroomLink, emailTemp.linkMeet, new Date(),emailTemp.emailsEnvio];
              logArray.push(newArray);

              } catch (errorEmail) {
                console.log(errorEmail);
                }
            });

              var ssLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LOG-REGISTRO");
              var ultimaFila = ssLog.getLastRow();
              console.log("UltimaFila");
              console.log(ultimaFila);
              ssLog.getRange(ultimaFila+1, 1, logArray.length, logArray[0].length).setValues(logArray);
        }  else {
            console.log("No encuentra data NEW que cumpla con el filtro")
               } 
        

        // DATA Envios OLD
          dataOLD = data.filter(function(r) { return r[19] == "OLD" && r[20] == "SI" && r[21] == "PERMITIDO" && r[22] != "ENVIADO" });
          console.log("OLD");
          console.log(dataOLD);

        var logArray =[]; // reseteamos el Array de Resultados

        if (dataOLD && dataOLD.length) {
          
          // Decido que plantilla/Templates de emails HTML uso
          var emailTemp = HtmlService.createTemplateFromFile("emailsViejos");
          var asuntoEmailViejo = "¡Bienvenido nuevamente! Ya podés acceder a tu aula interactiva del Liceo 2021"

              dataOLD.forEach(function(row){
              console.log("Siguiente Registro OLD: ");
              emailTemp.fila = row[fila];
              emailTemp.sede = row[sede];
              emailTemp.nombreAlumno = row[nombreAlumno];
              emailTemp.userLCB = row[userLCB];
              emailTemp.password = row[password];
              emailTemp.linkMeet = row[linkMeet];
              emailTemp.classroomName = row[classroomName];    
              emailTemp.classroomLink = row[classroomLink];
              emailTemp.emailsEnvio = row[emailsEnvio];
              emailTemp.emailReply = row[emailReply];
              emailTemp.linkFaq = 'https://docs.google.com/document/d/1-xOYcEuuFb8mMw1P72xc774V2YFPuNISPbR_LwJHRpg/';
              emailTemp.linkTutorial = 'https://drive.google.com/file/d/1EHUk4bv60ilsvvm8LhiU10sXkTrQidXR/view?usp=sharing';
              emailTemp.linkCalendar = 'https://calendar.google.com/';

              
              console.log(emailTemp.nombreAlumno)
              console.log("Fila: "+emailTemp.fila);

            try {   
              var htmlMessage = emailTemp.evaluate().getContent();


              MailApp.sendEmail({
                to: emailTemp.emailsEnvio,
                replyTo: emailTemp.emailReply,
                cc: emailTemp.userLCB,
                bcc: emailTemp.emailReply, 
                subject: asuntoEmailViejo,
                htmlBody: htmlMessage,
                body: "Si no puede ver el contenido de este correo, por favor comunícate con la secretaría de tu sede del Liceo Británico.",
              })

              console.log("Email enviado")   
              var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
              ss.getRange(emailTemp.fila, escribirResultado).setValue('ENVIADO');

              var newArray = [emailTemp.sede, emailTemp.userLCB, emailTemp.password, emailTemp.classroomName ,emailTemp.classroomLink, emailTemp.linkMeet, new Date(),emailTemp.emailsEnvio];
              logArray.push(newArray);

              } catch (errorEmail) {
                console.log(errorEmail);
                }
            });
              var ssLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LOG-REGISTRO");
              var ultimaFila = ssLog.getLastRow();
              console.log("UltimaFila");
              console.log(ultimaFila);
              ssLog.getRange(ultimaFila+1, 1, logArray.length, logArray[0].length).setValues(logArray);
          
          } else {
            console.log("No encuentra data OLD que cumpla con el filtro")
            } 
} else {
    console.log("Decidiste NO CORRER el script de envios O escribiste el usuario incorrecto")
}

}
