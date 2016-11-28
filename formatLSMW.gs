function formatoLSMW() {
  var sheet 	= SpreadsheetApp.getActiveSheet();
  var data	= sheet.getDataRange().getValues();
  var newSheet 	= SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();
  var newSheet 	= SpreadsheetApp.getActiveSheet().clear();
  var pedido; 
  var item;
  for (var i = 0; i < data.length; i++) {
    if (pedido != data[i][0]) {
      newSheet.appendRow(['H', data[i][0]]);
    }
    if ((item != data[i][1]) || (pedido != data[i][0])) {         
      newSheet.appendRow(['I', data[i][1], 'ZFE', data[i][2], 'U']);
      newSheet.appendRow(['C', data[i][1], 'ZFE', 'X', 'X', 'X', 'X', 'X']);
    }
    pedido = data[i][0];
    item = data[i][1];
  }
}