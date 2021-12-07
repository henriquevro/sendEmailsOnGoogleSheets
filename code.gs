// This constant is written on the Flag Column where the respective e-mail was sucessfully sended
var EMAIL_SENT = 'EMAIL_SENT';
var NO_EMAIL = 'NO_EMAIL'

/**
 * Send non-duplicate e-mails from choosed spreadsheet
 */

function sendEmails() {
  
  //Chosing e activating the correct spreadsheet tab of the WorkSheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Pessoas + Form Links')


  var startRow = 2; // Primeira linha de dados para ser processado
  var numRows = 350; // Número de linhas a serem processadas
  var dataRange = sheet.getRange(startRow, 1, numRows, 25); // Fetch o range de células A2:T3
  var data = dataRange.getValues(); // Fetch valores para cada linha no Range.

  var picture1 = DriveApp.getFileById('<picID>'); //public with link https://drive.google.com/file/d/<picID>/view?usp=sharing

  var inlineImages = {};
  inlineImages[picture1.getId()] = picture1.getBlob();

  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[1]; // Primeira coluna
    var emailSent = row[16]; // Coluna Q
    var nomeCompleto = row[6]
    var arraydoNomeCompleto = nomeCompleto.split(" "); //Split by space
    var primeiroNome = arraydoNomeCompleto[0]
    //var message = "Olá, " + primeiroNome + "! \n \n" + "Aqui está o formulário para registro do seu crachá:\n \n" + row[15]; // Coluna P 

    if (emailSent !== EMAIL_SENT) { // Impede o envio de e-mails duplicados
      if (emailAddress == "") {
        sheet.getRange(startRow + i, 17).setValue(NO_EMAIL); // Atualização imediata, caso haja interrupção
      } else {
      MailApp.sendEmail({
        to: emailAddress, 
        subject: "Olá, " + primeiroNome + "!   Aqui está o formulário para seu novo crachá",
        htmlBody: 
        "Olá, " + primeiroNome + '! <br> <br> Aqui está o formulário para registro do seu crachá: <br> <br> <a href=" ' + row[15] + '"> Clique aqui para acessar seu formulário </a> <br> <br>' +
        '<br><img src="cid:' + picture1.getId() + '" /><br>',
        inlineImages: inlineImages   
        });
      sheet.getRange(startRow + i, 17).setValue(EMAIL_SENT); // Atualização imediata, caso haja interrupção
      SpreadsheetApp.flush();
      }
    }
  }
}
