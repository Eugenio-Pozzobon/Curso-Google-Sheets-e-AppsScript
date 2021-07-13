function Atualizar() {

  //Remove os filtros
  var spreadsheet = SpreadsheetApp.getActive();
    try {
    spreadsheet.getActiveSheet().getFilter().remove();
  } catch (e) {
    // Logs an ERROR message.
    // console.error(e);
  }

  //Abre a página de tarefas e pega a seleção das células desejadas
  var tarefas = spreadsheet.getSheetByName("13. Tarefas");
  var tarefas_celulas = tarefas.getRange('A4:E');
  var concluidas = spreadsheet.getSheetByName("13.5 Concluídas");
  var concluidas_celulas = concluidas.getRange('A2:C');

  // lê a planilha linha por linha
  for(var i=1;i<tarefas_celulas.getLastRow()-2; i++){

    var delayDate = (tarefas_celulas.getCell(i,4).getValue());

    //Se tiver algo escrito na célula de prorrogar, adiciona isso na data final na mesma linha
    if (delayDate != ""){
      var oldDate = new Date(tarefas_celulas.getCell(i,2).getValue());

      var newDate = new Date(oldDate);
      newDate.setDate(oldDate.getDate() + delayDate);

      tarefas_celulas.getCell(i,2).setValue(newDate);
      tarefas_celulas.getCell(i,4).setValue("");
    }

    //verifica se existem tarefas concluídas, se sim transfere os dados para a outra planilha
    try {

      if(tarefas_celulas.getCell(i,5).getValue()==true){

        concluidas.appendRow([
          tarefas_celulas.getCell(i,1).getValue(),
          tarefas_celulas.getCell(i,2).getValue(),
          tarefas_celulas.getCell(i,5).getValue()]
        );

        tarefas.deleteRow(i+3);
        tarefas_celulas = tarefas.getRange('A4:E');
        i=0;  
      }
    } catch (e) {
      // Logs an ERROR message.
      // console.error(e);
    }
  }

  //verifica se existem tarefas deixaram de ser concluídas, se sim transfere os dados de volta para as tarefas
  for(var i=1;i<concluidas_celulas.getLastRow()-1; i++){
    if(concluidas_celulas.getCell(i,3).getValue()==false){

      tarefas.appendRow([
        concluidas_celulas.getCell(i,1).getValue(),
        concluidas_celulas.getCell(i,2).getValue(),
        "",
        "",
        false]
      );

      concluidas.deleteRow(i+1)
      concluidas_celulas = concluidas.getRange('A2:C');
      i=0;  
    }
  }


  // seleciona a planilha novamente e reordena os dados em ordem de prioridade de data
  tarefas.getRange('A4:E').activate();
  tarefas.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  tarefas.getRange('C4').activate();
  tarefas.getActiveRange().autoFill(spreadsheet.getRange('C4:C'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  
  tarefas.getRange('A3:E').activate();
  tarefas.getRange('A3:E').createFilter();
  tarefas.getRange('B3').activate();
  tarefas.getFilter().sort(2, true);

};
