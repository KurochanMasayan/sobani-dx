function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PDF\')
    .addItem('PDF�������', 'exportToPDF')
    .addToUi();
}

function exportToPDF() {
  // PDF\_�o�g��W~Y
  SpreadsheetApp.getUi().alert('PDF\_���ň�gY');
}