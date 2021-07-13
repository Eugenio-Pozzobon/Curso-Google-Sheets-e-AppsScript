function gerarCertificados() {

  var slides = SlidesApp.openById("1_S52B_qq7qESdBh_fMzu754WsyDTW8H0OUd-uSp4HtU");
  
  var Ssheet = SpreadsheetApp.openById("1-VYKWmumigG3LKTiKAdHFhwhgT6Ahw6RnCfmx-LQuwo");
  
  var inscritosSheet = Ssheet.getSheetByName("15. Inscritos");
  
  var inscritos = inscritosSheet.getRange("A2:F");
  
  // lê os dados de cada linha da planilha de inscritos
  for(var i=1;i<inscritos.getLastRow(); i++){
      var name = inscritos.getCell(i, 1);
      var curso = inscritos.getCell(i, 3);
      var horas = inscritos.getCell(i, 4);
      var code = inscritos.getCell(i, 5);
      
      var pCm = (72)/2.55

      // se o inscrito estiver com a presença mínima exigida, gera o texto e insere o texto numa folha do google slides
      if(inscritos.getCell(i, 6).getValue()){

        if(name.getValue()!=""){

        var slide=slides.appendSlide();
          
        var certificadoMainText = "Certificado"
        var mainTextBox = slide.insertTextBox(certificadoMainText , 8*pCm, 1*pCm , 9*pCm, 1*pCm);

        var certificadoAuxText1 = "Certificamos que, "
        var certificadoAuxText2 = " participou como ouvinte do minicurso de "
        var certificadoAuxText3 = ", promovido pela Escola Piloto de Engenharia Mecânica, realizado entre os dias 10 e 14 de julho, no Centro de Tecnologia da Universidade Federal de Santa Maria, totalizando uma carga horária total de "

        var text = certificadoAuxText1 + name.getValue() + certificadoAuxText2 + curso.getValue() + certificadoAuxText3 + horas.getValue() 
        var auxTextBox =  slide.insertTextBox(text , 4*pCm, 5*pCm , 16*pCm, 4*pCm);

        var certificadoValidationText = "Codigo de Autenticação: " + code.getValue()
        var validationTextBox = slide.insertTextBox(certificadoValidationText , 17*pCm, 13*pCm, 8*pCm, 2*pCm);

        slide.refreshSlide();
        
        }
      }
  }  
}

