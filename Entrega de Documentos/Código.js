function onFormSubmit(e) {
  var responses = e.namedValues;

  var nomeAlunoOriginal = responses["Nome Completo"][0];
  var nomeAluno = capitalizarNome(nomeAlunoOriginal);

  var cpfAluno = responses["CPF"][0];
  var cpfLimpo = cpfAluno.replace(/\D/g, "");

  var pastaPrincipalId = "1dCkIHrt60xqIPmS_vUKrjWxw4OMc2u9y";
  var pastaPrincipal = DriveApp.getFolderById(pastaPrincipalId);

  var nomeLimpo = nomeAluno.replace(/[\\/:*?"<>|]/g, "_");
  var nomePasta = nomeLimpo + " - " + cpfLimpo;

  var pastas = pastaPrincipal.getFolders();
  var pastaAluno = null;

  while (pastas.hasNext()) {
    var pasta = pastas.next();
    if (pasta.getName().indexOf(cpfLimpo) !== -1) {
      pastaAluno = pasta;
      break;
    }
  }

  if (!pastaAluno) {
    pastaAluno = pastaPrincipal.createFolder(nomePasta);
  }

  var camposUpload = {
    "CPF e RG (ou CNH – Carteira Nacional de Habilitação)": "Documento de Identidade",
    "Comprovante de residência": "Comprovante de Residencia",
    "Cópia da frente do diploma de graduação": "Diploma Frente",
    "Cópia do verso do diploma de graduação": "Diploma Verso"
  };

  for (var campo in camposUpload) {
    var arquivos = responses[campo];
    if (arquivos && arquivos.length > 0) {
      arquivos.forEach(function(links) {
        var arquivosLinks = links.split(", ");
        arquivosLinks.forEach(function(link) {
          var fileId = getFileIdFromUrl(link);
          if (fileId) {
            try {
              var file = DriveApp.getFileById(fileId);
              file.setName(camposUpload[campo]);
              pastaAluno.addFile(file);

              var parents = file.getParents();
              if (parents.hasNext()) {
                parents.next().removeFile(file);
              }

              Logger.log("Arquivo movido e renomeado: " + file.getName());
            } catch (err) {
              Logger.log("Erro ao processar arquivo do campo " + campo + ": " + err);
            }
          }
        });
      });
    }
  }

  Logger.log("Pasta organizada para " + pastaAluno.getName());
}

function capitalizarNome(nome) {
  return nome.toLowerCase().split(" ").map(function(palavra) {
    return palavra.charAt(0).toUpperCase() + palavra.slice(1);
  }).join(" ");
}

function getFileIdFromUrl(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}