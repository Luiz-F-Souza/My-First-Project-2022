export function criarLinha({sheet, dadosObj,startColumn,endColumn}) {
   
  const arrayComDados = []

  for(const key in dadosObj){

    const value = dadosObj[key]
    arrayComDados.push(value)
  }
  
  // arrayComDados.push(emailDoUsuario)
  
  try{

    sheet.appendRow(arrayComDados)
  
    const numeroDaLinhaCriada = sheet.getLastRow()

    const linhaCriada = sheet.getRange(`${startColumn}${numeroDaLinhaCriada}:${endColumn}${numeroDaLinhaCriada}`)

    linhaCriada.setBackground('white')
    linhaCriada.setFontSize(9)
    linhaCriada.setFontWeight('normal')
    linhaCriada.setFontColor('#202124')
    
    sheet.setRowHeight(numeroDaLinhaCriada,40)

    return { foiCriada: true, numeroDaLinhaCriada }
  }
  catch(err){
    console.log(err)
    SpreadsheetApp.getUi().alert(err, SpreadsheetApp.getUi().ButtonSet.OK)

    return { foiCriada: false, numeroDaLinhaCriada: null}
  }
  
}