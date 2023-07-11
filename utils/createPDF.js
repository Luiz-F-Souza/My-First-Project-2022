export function createPDF(planilhaID, aba, nomeDoPDF, idPastaNoDrive,isPortrait) {
  console.log('começou a gerar o pdf')

    const url = "https://docs.google.com/spreadsheets/d/" + planilhaID + "/export" +
      "?format=pdf&" +
      "size=A4&" +
      "fzr=true&" +
      `portrait=${isPortrait}&` +
      "fitw=true&" +
      "gridlines=false&" +
      "printtitle=false&" +
      "top_margin=0.5&" +
      "bottom_margin=0.25&" +
      "left_margin=0.5&" +
      "right_margin=0.5&" +
      "sheetnames=false&" +
      "pagenum=UNDEFINED&" +
      "attachment=true&" +
      "gid=" + aba.getSheetId()
      
    const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
    const blob = UrlFetchApp.fetch(url, params).getBlob().setName(nomeDoPDF + '.pdf');

    // Gets the folder in Drive where the PDFs are stored.
    const folder = DriveApp.getFolderById(idPastaNoDrive);

    const pdfFile = folder.createFile(blob);

    console.log('terminou de gerar o pdf')
    return pdfFile;
}