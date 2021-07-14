function validarCodigo() {
  var Ssheet = SpreadsheetApp.openById("1-VYKWmumigG3LKTiKAdHFhwhgT6Ahw6RnCfmx-LQuwo");
  
  var inscritosSheet = Ssheet.getSheetByName("15. Inscritos");
  var requisicoesSheet = Ssheet.getSheetByName("16. Validador");
  
  var inscritos = inscritosSheet.getRange("A2:F");
  
  var requisicoes = requisicoesSheet.getRange("A2:D");


  //checa todas as requisições
  for(var i=1;i<requisicoes.getLastRow(); i++){

  var confirmation = false

    //se a linha não tiver sido conferida anteriormente, indicando novo valor teste
    if(!requisicoes.getCell(i, 4).getValue()){
      var codeRequested = requisicoes.getCell(i, 3)
      var email = requisicoes.getCell(i, 2)

      //se a linha tiver um email
      if(email.getValue() != ""){

        //lê todos os dados dos incritos para procurar correspondência dos valores
        for(var j=1;j<inscritos.getLastRow(); j++){

          //se o inscrito estava apto a receber certificado 
          if(inscritos.getCell(j, 6).getValue()){
            var code = inscritos.getCell(j, 5);

            //checa o valor do código e compara com o valor recebido pelo usuário
            if(code.getValue() == codeRequested.getValue()){
                
              //envia email confirmando
              var message = "Prezado(a),\n\n O código [" +  codeRequested.getValue() + "] informado é válido\n\nAtt.\n\nEPEM"
              var subject = "Validação de Certificado";
              MailApp.sendEmail(email.getValue(), subject, message);
              confirmation = true
            }
          }
        }

        if (!confirmation){
          //envia e-mail informando que não encontrou correspondência
          var message = "Prezado(a),\n\n O código [" +  codeRequested.getValue() + "] informado NÃO é válido\n\nAtt.\n\nEPEM"
          var subject = "Validação de Certificado";
          MailApp.sendEmail(email.getValue(), subject, message);
        }
        
        requisicoes.getCell(i, 4).setValue(true)
      }
    }
  }
}
