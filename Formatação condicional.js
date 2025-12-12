function aplicarFormatacaoCondicional() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName("Planilha2");

  // Limpa regras antigas (opcional)
  aba.clearConditionalFormatRules();

  const ultimaLinha = aba.getLastRow();
  const intervalo = aba.getRange(`A3:I${ultimaLinha}`);

  const regra = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$P3=TRUE')  // Coluna F é a das caixas de seleção
    .setBackground('#A8E6A3')           // Verde claro (pode trocar a cor)
    .setRanges([intervalo])
    .build();

  aba.setConditionalFormatRules([regra]);
}