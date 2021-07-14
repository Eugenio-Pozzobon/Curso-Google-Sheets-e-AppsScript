function gerarCertificados() {

  var slides = SlidesApp.openById("1_S52B_qq7qESdBh_fMzu754WsyDTW8H0OUd-uSp4HtU");
  
  var Ssheet = SpreadsheetApp.openById("1-VYKWmumigG3LKTiKAdHFhwhgT6Ahw6RnCfmx-LQuwo");
  
  var inscritosSheet = Ssheet.getSheetByName("15. Inscritos");
  
  var inscritos = inscritosSheet.getRange("A2:F");
  
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Gerador de Certificados',
     'Você deseja apagar os certificados que já foram gerados?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    var pages = slides.getSlides();

    for(var i = pages.length-1; i>=0; i--){
      pages[i].remove();
    }

  } else {
    // User clicked "No" or X in the alert.
  }

  // lê os dados de cada linha da planilha de inscritos
  for(var i=1;i<inscritos.getLastRow(); i++){
      var name = inscritos.getCell(i, 1);
      var curso = inscritos.getCell(i, 3);
      var horas = inscritos.getCell(i, 5);
      var code = inscritos.getCell(i, 6);
      
      var pCm = (72)/2.55

      // se o inscrito estiver com a presença mínima exigida, gera o texto e insere o texto numa folha do google slides
      if(inscritos.getCell(i, 4).getValue()){

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
  
  ui.alert('Certificados Gerados. Acesse o Google Slides');  
}

function gerarNovoCertificado() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var nameResult = ui.prompt(
      'Adicionar Certificado',
      'Qual o NOME que constará no Certificado?',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var namebutton = nameResult.getSelectedButton();
  var name = nameResult.getResponseText();
  var curso = "";
  var ch = "0h";
  var validationCode = 0;

  if (namebutton == ui.Button.OK) {
    // User clicked "OK".

    var cursoResult = ui.prompt(
    'Adicionar Certificado',
    'Qual o CURSO que constará no Certificado?',
    ui.ButtonSet.OK_CANCEL);

    // Process the user's response.
    var buttonCurso = cursoResult.getSelectedButton();
    curso = cursoResult.getResponseText();

    if (buttonCurso == ui.Button.OK) {
      switch(curso){ //checa se o curso informado consta e atribui a carga horária relativa ao curso
        case "MATLAB":
          ch = "8h";
          break;
        case "Microcontroladores":
          ch = "6h";
          break;
        case "LaTeX":
          ch = "4h";
          break;
        case "HP 50g":
          ch = "5h";
          break;
        default:
          ch = "0h";
          break;
      }
      
      console.log(curso)
      console.log(ch)

      if(ch != "0h"){
        //open your Spread sheet by passing id
        var Ssheet= SpreadsheetApp.openById("1-VYKWmumigG3LKTiKAdHFhwhgT6Ahw6RnCfmx-LQuwo");
        var inscritosSheet = Ssheet.getSheetByName("15. Inscritos");
        
        //add new row with recieved parameter from client
        validationCode = 101
        inscritosSheet.appendRow([name,"",curso,true,ch,validationCode]);  


        var pCm = (72)/2.55

        //adiciona o slide
        var slides = SlidesApp.openById("1_S52B_qq7qESdBh_fMzu754WsyDTW8H0OUd-uSp4HtU");
        var slide=slides.appendSlide();
          
        var certificadoMainText = "Certificado"
        var mainTextBox = slide.insertTextBox(certificadoMainText , 8*pCm, 1*pCm , 9*pCm, 1*pCm);

        var certificadoAuxText1 = "Certificamos que, "
        var certificadoAuxText2 = " participou como ouvinte do minicurso de "
        var certificadoAuxText3 = ", promovido pela Escola Piloto de Engenharia Mecânica, realizado entre os dias 10 e 14 de julho, no Centro de Tecnologia da Universidade Federal de Santa Maria, totalizando uma carga horária total de "

        var text = certificadoAuxText1 + name + certificadoAuxText2 + curso + certificadoAuxText3 + ch
        var auxTextBox =  slide.insertTextBox(text , 4*pCm, 5*pCm , 16*pCm, 4*pCm);

        var certificadoValidationText = "Codigo de Autenticação: " + validationCode
        var validationTextBox = slide.insertTextBox(certificadoValidationText , 17*pCm, 13*pCm, 8*pCm, 2*pCm);

        slide.refreshSlide();

      }else{
        ui.alert("O curso informado não consta na Jornada")
      }
    }

  }
}

