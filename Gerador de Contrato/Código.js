function gerarContratos() {
  const ui = SpreadsheetApp.getUi();
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dados = planilha.getDataRange().getValues();
  const cabecalho = dados[0];
  const linhas = dados.slice(1);

  // 📌 MAPA DE MODELOS – (Opção A)
  const modelos = {
    "Contrato Pós 1": "ID",
    "Contrato Pós 2": "ID",
    "Contrato Pós 3": "ID",
    "Contrato Pós 4": "ID",
    "Contrato Pós 5": "ID",
    "Contrato Pós 6": "ID"
  };

  const idPastaPrincipal = 'ID';
  const pastaPrincipal = DriveApp.getFolderById(idPastaPrincipal);

  let indexStatus = cabecalho.indexOf('Status');
  let indexLinkPasta = cabecalho.indexOf('Link da Pasta');
  let indexLinkDocumento = cabecalho.indexOf('Link do Documento');
  let indexLinkPDF = cabecalho.indexOf('Link do PDF');
  const indexSelecao = cabecalho.indexOf('Selecionar');
  const indexModelo = cabecalho.indexOf('Modelo');
  const indexNomeBase = cabecalho.indexOf('Nome Base') !== -1 ? cabecalho.indexOf('Nome Base') : 0;
  const indexNomeCliente = cabecalho.indexOf('Nome Cliente') !== -1 ? cabecalho.indexOf('Nome Cliente') : 1;

  // Criar colunas automaticamente se não existirem
  if (indexLinkDocumento === -1) {
    planilha.getRange(1, cabecalho.length + 1).setValue('Link do Documento');
    planilha.getRange(1, cabecalho.length + 2).setValue('Link do PDF');
    planilha.getRange(1, cabecalho.length + 3).setValue('Link da Pasta');

    const newHead = planilha.getDataRange().getValues()[0];
    indexLinkDocumento = newHead.indexOf('Link do Documento');
    indexLinkPDF = newHead.indexOf('Link do PDF');
    indexLinkPasta = newHead.indexOf('Link da Pasta');
    indexStatus = newHead.indexOf('Status');
  }

  const numColunas = planilha.getLastColumn();

  const linksDocumento = [];
  const linksPdf = [];
  const linksPastas = [];
  const statusValores = [];
  const backgrounds = [];

  linhas.forEach((linha, i) => {
    const selecionar = linha[indexSelecao];
    const modeloEscolhido = linha[indexModelo];
    const emptyRich = SpreadsheetApp.newRichTextValue().setText('').build();

    const linhaIndex = i + 2;

    // Linha não selecionada → mantém valores antigos
    if (selecionar !== true) {
      const linkDocAtual = planilha.getRange(linhaIndex, indexLinkDocumento + 1).getRichTextValue();
      const linkPdfAtual = planilha.getRange(linhaIndex, indexLinkPDF + 1).getRichTextValue();
      const linkPastaAtual = planilha.getRange(linhaIndex, indexLinkPasta + 1).getRichTextValue();
      const statusAtual = planilha.getRange(linhaIndex, indexStatus + 1).getValue();
      const fundoAtual = planilha.getRange(linhaIndex, 1, 1, numColunas).getBackgrounds()[0];

      linksDocumento.push([linkDocAtual]);
      linksPdf.push([linkPdfAtual]);
      linksPastas.push([linkPastaAtual]);
      statusValores.push([statusAtual]);
      backgrounds.push([...fundoAtual]);
      return;
    }

    // Verifica modelo informado
    if (!modeloEscolhido || !modelos[modeloEscolhido]) {
      statusValores.push([`Modelo inválido: ${modeloEscolhido}`]);
      linksDocumento.push([emptyRich]);
      linksPdf.push([emptyRich]);
      linksPastas.push([emptyRich]);
      backgrounds.push(new Array(numColunas).fill('#ffccc7'));
      return;
    }

    const idModelo = modelos[modeloEscolhido];

    const nomeBase = String(linha[indexNomeBase]).trim();
    const nomeCliente = String(linha[indexNomeCliente]).trim();

    if (!nomeBase || !nomeCliente) {
      statusValores.push(['Dados incompletos']);
      linksDocumento.push([emptyRich]);
      linksPdf.push([emptyRich]);
      linksPastas.push([emptyRich]);
      backgrounds.push(new Array(numColunas).fill('#fff1f0'));
      return;
    }

    const nomePasta = `${nomeBase} - ${nomeCliente}`;
    const dadosSubstituir = {};

    cabecalho.forEach((col, j) => {
      if (col.startsWith('#')) {
        let valor = linha[j];

        if (valor instanceof Date) {
          const dia = String(valor.getDate()).padStart(2, '0');
          const mes = String(valor.getMonth() + 1).padStart(2, '0');
          const ano = valor.getFullYear();
          valor = `${dia}/${mes}/${ano}`;
        } else if (typeof valor === 'number' && col.toLowerCase().includes('valor')) {
          valor = `R$ ${valor.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.')}`;
        }

        dadosSubstituir[col] = valor;
      }
    });

    try {
      const pastas = pastaPrincipal.getFoldersByName(nomePasta);
      const pastaDestino = pastas.hasNext()
        ? pastas.next()
        : pastaPrincipal.createFolder(nomePasta);

      const modelo = DriveApp.getFileById(idModelo);
      const copia = modelo.makeCopy(nomePasta, pastaDestino);
      const doc = DocumentApp.openById(copia.getId());
      const corpo = doc.getBody();

      Object.keys(dadosSubstituir).forEach(tag => {
        corpo.replaceText(tag, dadosSubstituir[tag]);
      });

      doc.saveAndClose();

      const pdfBlob = copia.getAs(MimeType.PDF);
      const pdf = pastaDestino.createFile(pdfBlob.setName(nomePasta + '.pdf'));

      const richDoc = SpreadsheetApp.newRichTextValue().setText(nomePasta).setLinkUrl(copia.getUrl()).build();
      const richPdf = SpreadsheetApp.newRichTextValue().setText(nomePasta).setLinkUrl(pdf.getUrl()).build();
      const richPasta = SpreadsheetApp.newRichTextValue().setText(nomePasta).setLinkUrl(pastaDestino.getUrl()).build();

      linksDocumento.push([richDoc]);
      linksPdf.push([richPdf]);
      linksPastas.push([richPasta]);
      statusValores.push(['Gerado']);
      backgrounds.push(new Array(numColunas).fill('#d9f7be'));

    } catch (e) {
      statusValores.push([`Erro: ${e.message}`]);
      linksDocumento.push([emptyRich]);
      linksPdf.push([emptyRich]);
      linksPastas.push([emptyRich]);
      backgrounds.push(new Array(numColunas).fill('#fff1f0'));
    }
  });

  const startRow = 2;
  planilha.getRange(startRow, indexLinkDocumento + 1, linksDocumento.length, 1).setRichTextValues(linksDocumento);
  planilha.getRange(startRow, indexLinkPDF + 1, linksPdf.length, 1).setRichTextValues(linksPdf);
  planilha.getRange(startRow, indexLinkPasta + 1, linksPastas.length, 1).setRichTextValues(linksPastas);
  planilha.getRange(startRow, indexStatus + 1, statusValores.length, 1).setValues(statusValores);
  planilha.getRange(startRow, 1, backgrounds.length, numColunas).setBackgrounds(backgrounds);

  SpreadsheetApp.flush();
  ui.alert("Contratos gerados com sucesso!");
}