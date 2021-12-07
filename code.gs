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


  var startRow = 2; // First Row of data
  var numRows = 350; // Number of rows to be processed
  var dataRange = sheet.getRange(startRow, 1, numRows, 25); // Fetch the range of cells
  var data = dataRange.getValues(); // Fetch values of range cells

  var picture1 = DriveApp.getFileById('<picID>'); //Insertion of inline image public with link https://drive.google.com/file/d/<picID>/view?usp=sharing

  var inlineImages = {};
  inlineImages[picture1.getId()] = picture1.getBlob();

  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[1]; // Primeira coluna
    var emailSent = row[16]; // Coluna Q
    var nomeCompleto = row[6]
    var arraydoNomeCompleto = nomeCompleto.split(" "); //Split by space
    var primeiroNome = arraydoNomeCompleto[0]

    if (emailSent !== EMAIL_SENT) { // Impede o envio de e-mails duplicados
      if (emailAddress == "") {
        sheet.getRange(startRow + i, 17).setValue(NO_EMAIL); // Seta flag para NO_EMAIL, caso o campo de e-mail esteja vazio
      } else {
      MailApp.sendEmail({
        to: emailAddress, 
        subject: "Olá, " + primeiroNome + "!   Aqui está o email com seus dados",
        htmlBody: 
        "Olá, " + primeiroNome + '! <br> <br> Aqui está o email com seu formulário personalizado: <br> <br> <a href=" ' + row[15] + '"> Clique aqui para acessar seu formulário </a> <br> <br>' +
        '<br><img src="cid:' + picture1.getId() + '" /><br>',
        inlineImages: inlineImages   
        });
      sheet.getRange(startRow + i, 17).setValue(EMAIL_SENT); // Atualização imediata, caso haja interrupção
      SpreadsheetApp.flush();
      }
    }
  }
}
