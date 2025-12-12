function testDrive() {
  const folderId = "1hxb0hPZoTqwo4cQwXVDv-B5jeYrK4vVK"; // substitua pelo ID que você está usando
  try {
    const pasta = DriveApp.getFolderById(folderId);
    Logger.log("Acesso OK. Nome da pasta: " + pasta.getName());
    SpreadsheetApp.getUi().alert("Acesso OK. Nome da pasta: " + pasta.getName());
  } catch (e) {
    Logger.log("Erro ao acessar pasta: " + e.message);
    SpreadsheetApp.getUi().alert("Erro ao acessar pasta:\n" + e.message);
  }
}

