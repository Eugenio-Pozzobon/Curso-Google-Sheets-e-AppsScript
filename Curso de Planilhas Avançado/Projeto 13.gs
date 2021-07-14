function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu('Tasks');
  menu.addItem('Filtrar Urgentes', 'Urgentes').addToUi();
  menu.addItem('Filtrar Não Urgentes', 'NaoUrgentes').addToUi();
  menu.addItem('Atualizar Dados', 'Atualizar').addToUi();

  
  var menuJornada = SpreadsheetApp.getUi().createMenu('Jornada');
  menuJornada.addItem('Gerar Certificados', 'gerarCertificados').addToUi();
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