import { createPDF } from "../../utils/createPDF.js"
import { setandoBackgroundElemento } from "../../utils/SettingBackground.js"

export function modeloLaudoSoloFisica({interessado,municipio,convenio,comRecomendacao,numSolicitacao,dataEntrada,obsAoLab,col1,col2,comUmidade,comGranulometria,areiaGrossa,areiaFina,comArgilaNatural,comCondutividade,comDensidade, qtdAnalises}){
  
 
  SpreadsheetApp.flush()

  const planilhaModelo = SpreadsheetApp.openById("LAYOUT ID")
  const abaModelo = planilhaModelo.getSheetByName('MODELO')
  

  const planilhaDestino = SpreadsheetApp.openById('DESTINY ID')
  
  const celulasModelo = {
    interessado: abaModelo.getRange('B6:H6'),
    municipio:  abaModelo.getRange('B7:H7'),
    convenio: abaModelo.getRange('B8:H8'),
    numSolicitacao: abaModelo.getRange('L6:N6'),
    dataEntrada: abaModelo.getRange('L7:N7'),
    obsAoLab: abaModelo.getRange('H42:K43'),
    col1: abaModelo.getRange('D33:G41'),
    col2: abaModelo.getRange('H33:J41'),
    comRecomendacao: abaModelo.getRange('H44:K44'),
    qtdAnalises: abaModelo.getRange('M36:N36')
  
  }

  
  celulasModelo.interessado.setValue(`INTERESSADO: ${interessado}`)
  celulasModelo.municipio.setValue(municipio)
  celulasModelo.convenio.setValue(convenio)
  celulasModelo.numSolicitacao.setValue(`SOLICITAÇÃO: ${numSolicitacao}`)
  celulasModelo.dataEntrada.setValue(dataEntrada)

  comRecomendacao = comRecomendacao === "SIM" ? comRecomendacao = "Com recomendação" : ""
  celulasModelo.obsAoLab.setValue(obsAoLab)
  celulasModelo.comRecomendacao.setValue(comRecomendacao)
  celulasModelo.col1.setValue(col1)
  celulasModelo.col2.setValue(col2)
  celulasModelo.qtdAnalises.setValue(qtdAnalises)



  let abaCopia 
  

  try{
    console.log('começou a copiar as infos pra planilha laudos')
    abaCopia = abaModelo.copyTo(planilhaDestino)

    // Colocando a cor nos elementos a analisar

    setandoBackgroundElemento(abaCopia,'D9:E10',comUmidade)
    setandoBackgroundElemento(abaCopia,'F9:J9',comGranulometria)
    setandoBackgroundElemento(abaCopia,'F11:F11',areiaGrossa)
    setandoBackgroundElemento(abaCopia,'G11:G11',areiaFina)
    setandoBackgroundElemento(abaCopia,'K9:K10',comArgilaNatural)
    setandoBackgroundElemento(abaCopia,'L9:L10',comCondutividade)
    setandoBackgroundElemento(abaCopia,'M9:N10',comDensidade)

    abaCopia.setTabColor('red')
    abaCopia.setName(numSolicitacao)
    abaCopia.activate()
    planilhaDestino.moveActiveSheet(1)

    
    
    
    console.log('terminou de copiar as infos para planilha laudos')
  }
  catch(err){
    console.log(err,`ERRO AO COLOCAR O NOME NA ABA DO SOLO FÍSICA ${numSolicitacao} - ${interessado}`)
    return { erroMsgFisica: `${err},ERRO ENVIANDO O LAUDO. VERIFIQUE COM O LAB SE FOI OU NÃO ENVIADO`}
  }
  

  try{
    console.log('entrou no try do pdf')
    createPDF('ID',abaCopia,`${numSolicitacao} - ${interessado}`,"ID",false)
    
    abaCopia.hideColumn(abaCopia.getRange('A:A'))

  }
  catch(err){
    return { erroMsgFisica: `${err},ERRO AO GERAR O PDF`}
  }



  return { erroMsgFisica: null}
}