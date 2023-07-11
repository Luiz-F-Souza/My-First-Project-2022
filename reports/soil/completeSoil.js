import { createPDF } from "../../utils/createPDF.js"

export function modeloLaudoSolo({interessado,municipio,convenio,numSolicitacao,dataEntrada,obsAoLab,col1,col2, comRecomendacao, qtdAnalises}){

  SpreadsheetApp.flush()

  const planilhaModelo = SpreadsheetApp.openById('ID')
  const abaModelo = planilhaModelo.getSheetByName('MODELO')

  const planilhaDestino = SpreadsheetApp.openById('DESTINY ID')

  const celulas = {

      interessado: abaModelo.getRange('B6:L6'),
      municipio: abaModelo.getRange('B7:L7'),
      convenio: abaModelo.getRange('B8:L8'),
      numSolicitacao: abaModelo.getRange('T6:W6'),
      dataEntrada: abaModelo.getRange('T7:W7'),
  
      comRecomendacao: abaModelo.getRange('O27:S27'),
      col1: abaModelo.getRange('D30:J36'),
      col2: abaModelo.getRange('K30:R36'),
      obsAoLab: abaModelo.getRange('H37:N37'),
      qtdAnalises: abaModelo.getRange('T30:W31')

  }


  celulas.interessado.setValue(`INTERESSADO: ${interessado}`)
  celulas.municipio.setValue(municipio)
  celulas.convenio.setValue(convenio)
  celulas.numSolicitacao.setValue(`SOLICITAÇÃO: ${numSolicitacao}`)
  celulas.dataEntrada.setValue(dataEntrada)
  

  celulas.comRecomendacao.setValue(comRecomendacao === "SIM" ? "Com recomendação" : " ")

  console.log(comRecomendacao,'solo completa.')
  celulas.col1.setValue(col1)
  celulas.col2.setValue(col2)
  celulas.obsAoLab.setValue(obsAoLab)
  celulas.qtdAnalises.setValue(qtdAnalises)


  let abaCopia
  try{
    abaCopia = abaModelo.copyTo(planilhaDestino)
    abaCopia.setTabColor('red')
    abaCopia.setName(numSolicitacao)
    abaCopia.activate()
    planilhaDestino.moveActiveSheet(1)
    
  }
  catch(err){
    console.log(err)
    return { erroMsgCompleto: `${err},ERRO ENVIANDO O LAUDO. VERIFIQUE COM O LAB SE FOI OU NÃO ENVIADO`}
  }

  try{
    createPDF("ID",abaCopia,`${numSolicitacao} - ${interessado}`,'ID',false)

    abaCopia.hideColumn(abaCopia.getRange('A:A'))
  }
  catch(err){
    console.log(err)
    return { erroMsgCompleto: `${err},ERRO AO GERAR O PDF`}
  }
  

 
  return { erroMsgCompleto: null}

}