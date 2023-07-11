import { createPDF } from "../../utils/createPDF.js"

export function modeloLaudoSoloSimples({interessado,municipio,convenio,numSolicitacao,dataEntrada,obsAoLab,col1,col2,comEnxofre, comRecomendacao, qtdAnalises}){
  
  SpreadsheetApp.flush()

 
  const planilhaModelo = SpreadsheetApp.openById("LAYOYT ID")
  const abaModelo = planilhaModelo.getSheetByName('MODELO')
  

  const planilhaDestino = SpreadsheetApp.openById('DESTINY ID')
  
  const celulasModelo = {
    interessado: abaModelo.getRange('B6:J6'),
    municipio:  abaModelo.getRange('B7:J7'),
    convenio: abaModelo.getRange('B8:J8'),
    numSolicitacao: abaModelo.getRange('P6:S6'),
    dataEntrada: abaModelo.getRange('P7:S7'),
    comRecomendacao: abaModelo.getRange('M27:O27'),
    obsAoLab: abaModelo.getRange('H36:M38'),
    col1: abaModelo.getRange('D29:I35'),
    col2: abaModelo.getRange('J29:N35'),
    qtdAnalises: abaModelo.getRange('O28:R29')

    
  }

  
  celulasModelo.interessado.setValue(`INTERESSADO: ${interessado}`)
  celulasModelo.municipio.setValue(municipio)
  celulasModelo.convenio.setValue(convenio)
  celulasModelo.numSolicitacao.setValue(`SOLICITAÇÃO: ${numSolicitacao}`)
  celulasModelo.dataEntrada.setValue(dataEntrada)

  celulasModelo.comRecomendacao.setValue(comRecomendacao === "SIM" ? "Com recomendação" : "")
  console.log(comRecomendacao,'solo simples.')
  celulasModelo.obsAoLab.setValue(obsAoLab)
  celulasModelo.col1.setValue(col1)
  celulasModelo.col2.setValue(col2)
  celulasModelo.qtdAnalises.setValue(qtdAnalises)



  let abaCopia 

  try{
    abaCopia = abaModelo.copyTo(planilhaDestino)

    const backgroundColor = comEnxofre ? analisarElementoBackground : "#fff"

    
    abaCopia.getRange('E9:E9').setBackground(backgroundColor)
    
     
    abaCopia.setTabColor('red')
    abaCopia.setName(numSolicitacao)
    abaCopia.activate()
    planilhaDestino.moveActiveSheet(1)
  }
  catch(err){
    console.log(err,`ERRO AO COLOCAR O NOME NA ABA DO SOLO SIMPLES ${numSolicitacao} - ${interessado}`)
    return { erroMsgSimples: `${err},ERRO ENVIANDO O LAUDO. VERIFIQUE COM O LAB SE FOI OU NÃO ENVIADO`}
  }
  

  try{
    createPDF('ID',abaCopia,`${numSolicitacao} - ${interessado}`,"ID",false)
    abaCopia.hideColumn(abaCopia.getRange('A:A'))
  }
  catch(err){
    console.log(err,`ERRO AO CRIAR O PDF SOLO SIMPLES ${numSolicitacao} - ${interessado}`)
    return { erroMsgSimples: `${err},ERRO AO GERAR O PDF`}
  }


  return { erroMsgSimples: null}

}