function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gerar Contrato') // Nome do menu ao lado de "Ajuda"
    .addItem('Gerar Contrato', 'gerarContratos') // Nome da opção e função vinculada
    .addToUi();
}
