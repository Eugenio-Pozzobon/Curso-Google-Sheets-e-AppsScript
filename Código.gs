/** @OnlyCurrentDoc */

function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu('Tasks');
  menu.addItem('Filtrar Urgentes', 'Urgentes').addToUi();
  menu.addItem('Filtrar Não Urgentes', 'NaoUrgentes').addToUi();
  menu.addItem('Atualizar Dados', 'Atualizar').addToUi();
}

function Urgentes() {
  var spreadsheet = SpreadsheetApp.getActive();
  try {
    spreadsheet.getActiveSheet().getFilter().remove();
  } catch (e) {
    // Logs an ERROR message.
    // console.error(e);
  }
  spreadsheet.getRange('A3:E').activate();
  spreadsheet.getRange('A3:E').createFilter();
  spreadsheet.getRange('B3').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .whenFormulaSatisfied('=IF(ISBLANK(B4);TRUE;DATEDIF(TODAY();B4;"D")<7)')
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(2, criteria);
};

function NaoUrgentes() {
  var spreadsheet = SpreadsheetApp.getActive();
    try {
    spreadsheet.getActiveSheet().getFilter().remove();
  } catch (e) {
    // Logs an ERROR message.
    // console.error(e);
  }
  spreadsheet.getRange('A3:E').activate();
  spreadsheet.getRange('A3:E').createFilter();
  spreadsheet.getRange('B3').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .whenFormulaSatisfied('=IF(ISBLANK(B4);TRUE;DATEDIF(TODAY();B4;"D")>=7)')
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(2, criteria);
};

function Atualizar() {
  var spreadsheet = SpreadsheetApp.getActive();
    try {
    spreadsheet.getActiveSheet().getFilter().remove();
  } catch (e) {
    // Logs an ERROR message.
    // console.error(e);
  }

  var tarefas = spreadsheet.getSheetByName("Tarefas");
  var tarefas_celulas = tarefas.getRange('A4:E');
  var concluidas = spreadsheet.getSheetByName("Concluídas");
  var concluidas_celulas = concluidas.getRange('A2:C');

  for(var i=1;i<tarefas_celulas.getLastRow()-2; i++){


    var delayDate = (tarefas_celulas.getCell(i,4).getValue());

    if (delayDate != ""){
      var oldDate = new Date(tarefas_celulas.getCell(i,2).getValue());

      var newDate = new Date(oldDate);
      newDate.setDate(oldDate.getDate() + delayDate);

      tarefas_celulas.getCell(i,2).setValue(newDate);
      tarefas_celulas.getCell(i,4).setValue("");
    }

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

  tarefas.getRange('A4:E').activate();
  tarefas.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  tarefas.getRange('C4').activate();
  tarefas.getActiveRange().autoFill(spreadsheet.getRange('C4:C'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  
  tarefas.getRange('A3:E').activate();
  tarefas.getRange('A3:E').createFilter();
  tarefas.getRange('B3').activate();
  tarefas.getFilter().sort(2, true);

};
