import { createPDF } from "../../utils/createPDF.js"

export function enviarBacteriologica(subTipoDeAnalise,objParaCadastro,numSolicitacao){

  SpreadsheetApp.flush()

  const planilhaModelo = SpreadsheetApp.openById('ID LAYOUT')
  const planilhaLaudo = SpreadsheetApp.openById('ID REPORT')

  const abaSelecionadaModelo = planilhaModelo.getSheetByName(subTipoDeAnalise)
  const abaRascunhoModelo = planilhaModelo.getSheetByName("RASCUNHO ÁGUA")


  const celulasAReceberDados = {
    numSolicitacao: 'G11:J11',
    nome: 'B13:J13',
    enderecoDaColeta: "B14:J14",
    localDaColeta: "B15:J15",
    dataDaColeta: "B16:E16",
    horaDaColeta: "F16:J16",
    dataEHoraDeEntradaNoLab: "B17:J17",
    origemDaAmostra: 'B18:F18',
    convenio: 'H18:J18',
    textoColetadoPor: "B38:J39",
    rascunhoColetadoPor:'F21:J21',
    rascunhoObsAoLab: 'B61:F62'
    
  }

  abaSelecionadaModelo.getRange(celulasAReceberDados.numSolicitacao).setValue(`SOLICITAÇÃO: ${numSolicitacao}`)

  abaSelecionadaModelo.getRange(celulasAReceberDados.nome).setValue(`INTERESSADO: ${objParaCadastro.nome}`)

  abaSelecionadaModelo.getRange(celulasAReceberDados.enderecoDaColeta).setValue(`ENDEREÇO: ${objParaCadastro.enderecoDaColeta}`)
  abaSelecionadaModelo.getRange(celulasAReceberDados.localDaColeta).setValue(`LOCAL DA COLETA: ${objParaCadastro.localDaColeta}`)

  abaSelecionadaModelo.getRange(celulasAReceberDados.dataDaColeta).setValue(`DATA DA COLETA: ${objParaCadastro.dataDaColeta}`)
  abaSelecionadaModelo.getRange(celulasAReceberDados.horaDaColeta).setValue(`HORA DA COLETA: ${objParaCadastro.horaDaColeta}`)

  abaSelecionadaModelo.getRange(celulasAReceberDados.dataEHoraDeEntradaNoLab).setValue(`DATA / HORA DE ENTRADA NO LABORATÓRIO: ${objParaCadastro.dataEHoraDeEntradaNoLab}`)

  abaSelecionadaModelo.getRange(celulasAReceberDados.origemDaAmostra).setValue(`ORIGEM DA AMOSTRA: ${objParaCadastro.origemDaAmostra}`)
  abaSelecionadaModelo.getRange(celulasAReceberDados.convenio).setValue(`CONVÊNIO: ${objParaCadastro.convenio}`)


  if(objParaCadastro.responsavelPelaColeta === "O Interessado") objParaCadastro.responsavelPelaColeta = "Proprietário"
  abaSelecionadaModelo.getRange(celulasAReceberDados.textoColetadoPor).setValue(`Obs.: Amostra coletada por: ${objParaCadastro.responsavelPelaColeta}. Sendo de nossa responsabilidade somente o exame realizado no referido material.`)


  const erroBacteriologica = {
    abaPrincipal: null,
    abaRascunho: null
  }

  try{
    const abaLaudo = abaSelecionadaModelo.copyTo(planilhaLaudo)
    abaLaudo.activate()
    abaLaudo.setTabColor('red')
    abaLaudo.setName(numSolicitacao)

    planilhaLaudo.moveActiveSheet(1)

  }
  catch(err){

    erroBacteriologica.abaPrincipal = err
  }

  

  
  abaRascunhoModelo.getRange(celulasAReceberDados.numSolicitacao).setValue(`SOLICITAÇÃO: ${numSolicitacao}`)

  abaRascunhoModelo.getRange(celulasAReceberDados.nome).setValue(`INTERESSADO: ${objParaCadastro.nome}`)

  abaRascunhoModelo.getRange(celulasAReceberDados.enderecoDaColeta).setValue(`ENDEREÇO: ${objParaCadastro.enderecoDaColeta}`)
  abaRascunhoModelo.getRange(celulasAReceberDados.localDaColeta).setValue(`LOCAL DA COLETA: ${objParaCadastro.localDaColeta}`)

  abaRascunhoModelo.getRange(celulasAReceberDados.dataDaColeta).setValue(`DATA DA COLETA: ${objParaCadastro.dataDaColeta}`)
  abaRascunhoModelo.getRange(celulasAReceberDados.horaDaColeta).setValue(`HORA DA COLETA: ${objParaCadastro.horaDaColeta}`)

  abaRascunhoModelo.getRange(celulasAReceberDados.dataEHoraDeEntradaNoLab).setValue(`DATA / HORA DE ENTRADA NO LABORATÓRIO: ${objParaCadastro.dataEHoraDeEntradaNoLab}`)

  abaRascunhoModelo.getRange(celulasAReceberDados.origemDaAmostra).setValue(`ORIGEM DA AMOSTRA: ${objParaCadastro.origemDaAmostra}`)

  abaRascunhoModelo.getRange(celulasAReceberDados.convenio).setValue(`CONVÊNIO: ${objParaCadastro.convenio}`)

  abaRascunhoModelo.getRange(celulasAReceberDados.rascunhoColetadoPor).setValue(`COLETADO POR: ${objParaCadastro.responsavelPelaColeta}`)

  abaRascunhoModelo.getRange(celulasAReceberDados.rascunhoObsAoLab).setValue(objParaCadastro.obsAoLab)

  try{
    const abaLaudoRascunho = abaRascunhoModelo.copyTo(planilhaLaudo)
    abaLaudoRascunho.activate()
    abaLaudoRascunho.setTabColor('red')
    abaLaudoRascunho.setName(`RASCUNHO - ${numSolicitacao}`)

    planilhaLaudo.moveActiveSheet(2)

    createPDF('ID',abaLaudoRascunho,`${numSolicitacao} - ${objParaCadastro.nome}`,'ID',true)
 

  }
  catch(err){

    erroBacteriologica.abaRascunho = err
  }
  



  

  return { erroBacteriologica }

}