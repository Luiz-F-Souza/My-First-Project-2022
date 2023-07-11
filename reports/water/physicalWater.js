import { createPDF } from "../../utils/createPDF.js"

export function enviarFisicoQuimica(subTipoDeAnalise,objParaCadastro,numSolicitacao){
  SpreadsheetApp.flush()


 const planilhaModelo = SpreadsheetApp.openById('LAYOUT ID')
 const planilhaLaudo = SpreadsheetApp.openById('REPORT ID')


 const abaRascunho = planilhaModelo.getSheetByName(`RASCUNHO - MODELO (${subTipoDeAnalise})`)


 const celulasAbaRascunho = {

   nome: abaRascunho.getRange('B6:L6'),
   numSolicitacao: abaRascunho.getRange('S6:W6'),
   municipio: abaRascunho.getRange('B7:L7'),
   dataDeEntrada: abaRascunho.getRange("S7:W7"),
   convenio: abaRascunho.getRange('B8:L8'),
   localDaColeta: abaRascunho.getRange("D27:J31"),
   enderecoDaColeta: abaRascunho.getRange("D32:J34"),
   obsAoLab: abaRascunho.getRange('K27:Q33'),
   chumboECadmio: abaRascunho.getRange('N9:O9'),
   comDureza: abaRascunho.getRange('W9:W9')
   
 }

 if(subTipoDeAnalise === "Irrigação") {
   celulasAbaRascunho.numSolicitacao = abaRascunho.getRange('U6:W6')
   celulasAbaRascunho.dataDeEntrada = abaRascunho.getRange("U7:W7")
 }


 const erroFisicoQuimica = {
   abaPrincipal: null,
   abaRascunho: null
 }


 try{

   celulasAbaRascunho.nome.setValue(`INTERESSADO: ${objParaCadastro.nome}`)
   celulasAbaRascunho.numSolicitacao.setValue(`SOLICITAÇÃO: ${numSolicitacao}`)
   celulasAbaRascunho.municipio.setValue(`MUNICÍPIO: ${objParaCadastro.enderecoDaColeta.municipio}`)
   celulasAbaRascunho.convenio.setValue(`CONVÊNIO: ${objParaCadastro.convenio}`)
   celulasAbaRascunho.dataDeEntrada.setValue(`DATA DE ENTRADA: ${objParaCadastro.dataDaColeta}`)

   celulasAbaRascunho.localDaColeta.setValue(`LOCAL DA COLETA: ${objParaCadastro.localDaColeta}`)
   
   celulasAbaRascunho.enderecoDaColeta.setValue(`ENDEREÇO DA COLETA: ${objParaCadastro.enderecoDaColeta.ruaOuAvenida} ${objParaCadastro.enderecoDaColeta.nomeDaRuaOuAvenida}, ${objParaCadastro.enderecoDaColeta.numero} - ${objParaCadastro.enderecoDaColeta.bairro} - ${objParaCadastro.enderecoDaColeta.municipio}`)
  
   celulasAbaRascunho.chumboECadmio.setBackground(objParaCadastro.comChumboECadimio ? analisarElementoBackground : "#fff")
 
   celulasAbaRascunho.comDureza.setBackground(objParaCadastro.comDureza ? analisarElementoBackground : '#fff')

   celulasAbaRascunho.obsAoLab.setValue(objParaCadastro.obsAoLab)

   const abaCopia = abaRascunho.copyTo(planilhaLaudo)

   abaCopia.setName(`RASCUNHO ${numSolicitacao}`)
   abaCopia.setTabColor('red')
   abaCopia.activate()

   planilhaLaudo.moveActiveSheet(1)

   createPDF("ID",abaCopia,`${numSolicitacao} - ${objParaCadastro.nome}`,"ID",false)

 }
 catch(err){
   console.log(err,"erro criando a aba laudo de agua fisico quimica")
   erroFisicoQuimica.abaPrincipal = err
 }

 if(subTipoDeAnalise != "Consumo") return { erroFisicoQuimica }

 const abaModeloConsumo = planilhaModelo.getSheetByName('MODELO CONSUMO')

 const celulasAbaConsumo = {
   Interessado: abaModeloConsumo.getRange("B14:J14"),
   municipio: abaModeloConsumo.getRange("B15:J15"),
   localDaColeta: abaModeloConsumo.getRange("B17:J17"),
   enderecoDaColeta: abaModeloConsumo.getRange("B16:J16"),
   responsavelPelaColeta: abaModeloConsumo.getRange("B19:G19"),
   numSolicitacao: abaModeloConsumo.getRange("H11:J11"),
   dataDeEntrada: abaModeloConsumo.getRange("H12:J12"),
   convenio: abaModeloConsumo.getRange("H19:J19"),
   obsAoLab: abaModeloConsumo.getRange("B47:F50")
 }

 try{
   celulasAbaConsumo.Interessado.setValue(`INTERESSADO: ${objParaCadastro.nome}`)
   celulasAbaConsumo.municipio.setValue(`MUNICÍPIO: ${objParaCadastro.enderecoDaColeta.municipio}`)
   celulasAbaConsumo.localDaColeta.setValue(`LOCAL DA COLETA: ${objParaCadastro.localDaColeta}`)
   celulasAbaConsumo.enderecoDaColeta.setValue(`ENDEREÇO DA COLETA: ${objParaCadastro.enderecoDaColeta.ruaOuAvenida} ${objParaCadastro.enderecoDaColeta.nomeDaRuaOuAvenida}, ${objParaCadastro.enderecoDaColeta.numero} - ${objParaCadastro.enderecoDaColeta.bairro} - ${objParaCadastro.enderecoDaColeta.municipio}`)
   celulasAbaConsumo.responsavelPelaColeta.setValue(`RESPONSÁVEL PELA COLETA: ${objParaCadastro.responsavelPelaColeta}`)
   celulasAbaConsumo.numSolicitacao.setValue(`SOLICITAÇÃO: ${numSolicitacao}`)
   celulasAbaConsumo.dataDeEntrada.setValue(`DATA DE ENTRADA: ${objParaCadastro.dataDaColeta}`)
   celulasAbaConsumo.convenio.setValue(`CONVÊNIO: ${objParaCadastro.convenio}`)
   celulasAbaConsumo.obsAoLab.setValue(objParaCadastro.obsAoLab)

   const abaCopia = abaModeloConsumo.copyTo(planilhaLaudo)
   abaCopia.setTabColor('red')
   abaCopia.setName(numSolicitacao)
   abaCopia.activate()

   planilhaLaudo.moveActiveSheet(1)
 }
 catch(err){
   console.log(err)
   erroFisicoQuimica.abaPrincipal = err
 }
 

 return { erroFisicoQuimica }

}