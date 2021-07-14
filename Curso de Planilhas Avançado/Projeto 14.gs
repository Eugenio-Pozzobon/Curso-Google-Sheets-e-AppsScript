function recieveForms() {

  var Ssheet = SpreadsheetApp.openById("1-VYKWmumigG3LKTiKAdHFhwhgT6Ahw6RnCfmx-LQuwo");
  
  var inscritosSheet = Ssheet.getSheetByName("14. Formulário de Inscrição");
  
  var inscritos = inscritosSheet.getRange("A2:F");
  
  // lê os dados de cada linha da planilha de inscritos no formulario
  for(var i=1;i<inscritos.getLastRow(); i++){
    var email = inscritos.getCell(i, 2);
    var name = inscritos.getCell(i, 3);
    var curso = inscritos.getCell(i, 5);

    // se o email não tiver sido enviado ainda, envia o email.
    if(!inscritos.getCell(i, 6).getValue()){
      if(name.getValue()!=""){

        var message = "Prezado(a) " + name.getValue() + ",\n\nVocê se inscreveu na Jornada de Minicursos da EPEM para o curso: " +  curso.getValue() + "\nEm caso de dúvidas entre em contato conosco! Fique no aguardo de novas informações. \n\nAtt.\n\nEPEM \nEmail Enviado Automaticamente"

        var subject = "Inscrição na Jornada Confirmada";
        MailApp.sendEmail(email.getValue(), subject, message);
        inscritos.getCell(i, 6).setValue(true)
      }
    }
  }
}
