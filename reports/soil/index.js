import { criarLinha } from "../../utils/createNewLine.js"
import { prevenirDuploClick } from "../../utils/preventDoubleClick.js"
import { modeloLaudoSolo } from "./completeSoil.js"
import { modeloLaudoSoloSimples } from "./simpleSoil.js"
import { modeloLaudoSoloFisica } from "./physicalSoil.js"

export function enviarSolo(){
  console.log('entrou na func enviarSolo')
  const ui = SpreadsheetApp.getUi()

  ui.showModelessDialog(
    HtmlService.createHtmlOutput(`
      <div>
        <h3>A operação solicitada foi iniciada...</h3>

        <p>Por favor, aguarde cerca de 10 segundos antes de clicar em outro botão</p>
      </div>
    `).setWidth(1600).setHeight(400),
    "Envio de SOLO"
  )

  const planilhaRecepcao = SpreadsheetApp.openById('SPREADSHEET ID')
  const dataObj = new Date()
  const data = dataObj.toLocaleDateString('pt-BR',{month:'2-digit',year:'numeric',day:'2-digit'})
  const dataEHora = dataObj.toLocaleDateString('pt-BR', {month:'2-digit',day:'2-digit',year:'2-digit',hour:'2-digit',minute:'2-digit'})
  const ano = dataObj.toLocaleDateString('pt-BR',{year:'numeric'})

  const abaEnviarSolo = planilhaRecepcao.getSheetByName("ENVIAR SOLOS")

  const botaoEnviarSolo = abaEnviarSolo.getDrawings()[0]
  if(botaoEnviarSolo.getOnAction() === 'prevenirDuploClick') return prevenirDuploClick()

  SpreadsheetApp.flush()
  botaoEnviarSolo.setOnAction('prevenirDuploClick')
 

  const celulasComDados = {

    informacaoPessoal:{
      nome: abaEnviarSolo.getRange("L3:P3"),
      email:abaEnviarSolo.getRange("L5:P5"),
      celular: abaEnviarSolo.getRange("L7:P7"),
      empresaVinculo: abaEnviarSolo.getRange("L9:P9"),
      linkDaPastaNoDrive: abaEnviarSolo.getRange("L11:P11"),
      cpf: abaEnviarSolo.getRange("B3:D3"),
      cnpj: abaEnviarSolo.getRange("F3:H3")
    },

    informacaoAnalise:{
      qtdAnalises: abaEnviarSolo.getRange("B6:D6"),
      tipoDeConvenio: abaEnviarSolo.getRange("F6:H6"),
      tipoDeAnalise: abaEnviarSolo.getRange("B9:D9"),
      comRecomendacao: abaEnviarSolo.getRange("F9:H9"),
      municipio: abaEnviarSolo.getRange("B12:H12"),

      obs:{
        coluna1: abaEnviarSolo.getRange("B15:H17"),
        coluna2: abaEnviarSolo.getRange("J15:N17"),
        infoAoLab: abaEnviarSolo.getRange("B20:H22"),
      },
      
      elementos: {
        comEnxofre: abaEnviarSolo.getRange("J20:J20"),
        comArgila: abaEnviarSolo.getRange("K20:K20"),
        comCondutividade: abaEnviarSolo.getRange("L20:L20"),
        comDensidade: abaEnviarSolo.getRange("M20:M20"),
        comUmidade: abaEnviarSolo.getRange("N20:N20"),
        comGranulometria: abaEnviarSolo.getRange("J22:J22"),
        comAreiaGrossa: abaEnviarSolo.getRange("K22:K22"),
        comAreiaFina: abaEnviarSolo.getRange("L22:L22")
      }
      
    }
  }

  const valores = {

    informacaoPessoal:{
      nome: celulasComDados.informacaoPessoal.nome.getValue(),
      email:celulasComDados.informacaoPessoal.email.getValue(),
      celular: celulasComDados.informacaoPessoal.celular.getValue(),
      empresaVinculo: celulasComDados.informacaoPessoal.empresaVinculo.getValue(),
      linkDaPastaNoDrive: celulasComDados.informacaoPessoal.linkDaPastaNoDrive.getValue(),
      cpf: celulasComDados.informacaoPessoal.cpf.getValue(),
      cnpj: celulasComDados.informacaoPessoal.cnpj.getValue()
    },

    informacaoAnalise:{
      qtdAnalises: celulasComDados.informacaoAnalise.qtdAnalises.getValue(),
      tipoDeConvenio: celulasComDados.informacaoAnalise.tipoDeConvenio.getValue(),
      tipoDeAnalise: celulasComDados.informacaoAnalise.tipoDeAnalise.getValue(),
      comRecomendacao: celulasComDados.informacaoAnalise.comRecomendacao.getValue(),
      municipio: celulasComDados.informacaoAnalise.municipio.getValue(),

      obs:{
        coluna1: celulasComDados.informacaoAnalise.obs.coluna1.getValue(),
        coluna2: celulasComDados.informacaoAnalise.obs.coluna2.getValue(),
        infoAoLab: celulasComDados.informacaoAnalise.obs.infoAoLab.getValue(),
      },
      
      elementos: {
        comEnxofre: celulasComDados.informacaoAnalise.elementos.comEnxofre.getValue(),
        comArgila: celulasComDados.informacaoAnalise.elementos.comArgila.getValue(),
        comCondutividade: celulasComDados.informacaoAnalise.elementos.comCondutividade.getValue(),
        comDensidade: celulasComDados.informacaoAnalise.elementos.comDensidade.getValue(),
        comUmidade: celulasComDados.informacaoAnalise.elementos.comUmidade.getValue(),
        comGranulometria: celulasComDados.informacaoAnalise.elementos.comGranulometria.getValue(),
        comAreiaGrossa: celulasComDados.informacaoAnalise.elementos.comAreiaGrossa.getValue(),
        comAreiaFina: celulasComDados.informacaoAnalise.elementos.comAreiaFina.getValue()
      }
      
    }

  }

  if(valores.informacaoPessoal.cpf && valores.informacaoPessoal.cnpj){

     ui.showModelessDialog(
      HtmlService.createHtmlOutput(`
        <div>
          <h3>A operação solicitada foi finalizada com erro.</h3>

          <p>Os campos de cpf e cnpj estão marcados ao mesmo tempo.</p>

        </div>
      `).setWidth(500).setHeight(250),
    "Envio de SOLO"
    )

    botaoEnviarSolo.setOnAction('enviarSolo')

    return Error('cpf e cnpj selecionados ao mesmo tempo')
  }


  const historyObj = {
    checkBox: null,
    data: dataEHora,
    solicitacao: null,
    qtdAnalise: valores.informacaoAnalise.qtdAnalises,
    cpf: valores.informacaoPessoal.cpf ? valores.informacaoPessoal.cpf : valores.informacaoPessoal.cnpj,
    nome: valores.informacaoPessoal.nome,
    tel: valores.informacaoPessoal.celular,
    email: valores.informacaoPessoal.email,
    linkDaPasta: valores.informacaoPessoal.linkDaPastaNoDrive,
    empresaVinculo: valores.informacaoPessoal.empresaVinculo,
    municipio: valores.informacaoAnalise.municipio,
    obsCol1: valores.informacaoAnalise.obs.coluna1,
    obsCol2: valores.informacaoAnalise.obs.coluna2,
    infoAoLab: valores.informacaoAnalise.obs.infoAoLab,
    comRecomendacao: valores.informacaoAnalise.comRecomendacao
  }

  // utilizado para gerar o numero da solicitação
  let prefixoDaNumeracao = "SC/"
  let endColumn = "P"

  switch(valores.informacaoAnalise.tipoDeAnalise){

    case 'SOLO COMPLETO':
    historyObj.emailDoUsuario = emailDoUsuario
    break;

    case "SOLO SIMPLES":
      historyObj.comEnxofre = valores.informacaoAnalise.elementos.comEnxofre
      historyObj.emailDoUsuario = emailDoUsuario

      prefixoDaNumeracao = "SS/"
      endColumn = "Q"
      break;
    
    case "SOLO FÍSICA":
      historyObj.comArgila = valores.informacaoAnalise.elementos.comArgila
      historyObj.comCondutividade = valores.informacaoAnalise.elementos.comCondutividade
      historyObj.comDensidade = valores.informacaoAnalise.elementos.comDensidade
      historyObj.comUmidade = valores.informacaoAnalise.elementos.comUmidade
      historyObj.comGranulometria = valores.informacaoAnalise.elementos.comGranulometria
      historyObj.comAreiaGrossa = valores.informacaoAnalise.elementos.comAreiaGrossa
      historyObj.comAreiaFina = valores.informacaoAnalise.elementos.comAreiaFina
      historyObj.emailDoUsuario = emailDoUsuario

      prefixoDaNumeracao = "SF/"
      endColumn = "W"
    break;
    
  }


  const abaSelecionada = planilhaRecepcao.getSheetByName(valores.informacaoAnalise.tipoDeAnalise)

 
  const { foiCriada, numeroDaLinhaCriada } = criarLinha({sheet: abaSelecionada, dadosObj: historyObj, startColumn:'A',endColumn})


  if(!foiCriada) return Error('A linha não foi criada')

  // COLOCANDO O CHECK BOX NA COLUNA A (O CHECKBOX É USADO PARA ENVIAR EMAILS AOS CLIENTES QUE TIVEREM ESTE CAMPO MARCADO)
  abaSelecionada.getRange(`A${numeroDaLinhaCriada}:A${numeroDaLinhaCriada}`).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox())
  
  
  
  
  // COLOCANDO O Nº DA SOLICITAÇÃO NA COLUNA C. O - 3 É PQ A LINHA INICIAL DA CONTAGEM É A 4, DAI VOLTA O VALOR PARA 1
  const numSolicitacao = `${prefixoDaNumeracao}${(numeroDaLinhaCriada - 3).toLocaleString(undefined,{minimumIntegerDigits: 6, useGrouping: false})}.${ano}`

  abaSelecionada.getRange(`C${numeroDaLinhaCriada}:C${numeroDaLinhaCriada}`)
  .setValue(numSolicitacao)

  const erro = {
    messagem: null
  }
  

  switch(valores.informacaoAnalise.tipoDeAnalise){

    case 'SOLO COMPLETO':
      const { erroMsgCompleto } = modeloLaudoSolo({
        interessado: historyObj.nome,
        municipio: `MUNICÍPIO: ${historyObj.municipio}`,
        convenio: `CONVÊNIO: ${valores.informacaoAnalise.tipoDeConvenio}`,
        numSolicitacao: numSolicitacao,
        dataEntrada: `DATA DE ENTRADA: ${data}`,
        obsAoLab: historyObj.infoAoLab,
        col1: historyObj.obsCol1,
        col2: historyObj.obsCol2,
        comRecomendacao: historyObj.comRecomendacao,
        qtdAnalises: `QUANTIDADE DE ANÁLISES: ${valores.informacaoAnalise.qtdAnalises}`

      })

      erro.messagem = erroMsgCompleto
    break;

    
    case 'SOLO SIMPLES':

    const { erroMsgSimples } = modeloLaudoSoloSimples({
        interessado: historyObj.nome,
        municipio: `MUNICÍPIO: ${historyObj.municipio}`,
        convenio: `CONVÊNIO: ${valores.informacaoAnalise.tipoDeConvenio}`,
        numSolicitacao: numSolicitacao,
        dataEntrada: `DATA DE ENTRADA: ${data}`,
        obsAoLab: historyObj.infoAoLab,
        col1: historyObj.obsCol1,
        col2: historyObj.obsCol2,
        comEnxofre: historyObj.comEnxofre,
        comRecomendacao: historyObj.comRecomendacao,
        qtdAnalises: `QUANTIDADE DE ANÁLISES: ${valores.informacaoAnalise.qtdAnalises}`

      })

      erro.messagem = erroMsgSimples
    break;

    case 'SOLO FÍSICA':

     const { erroMsgFisica } = modeloLaudoSoloFisica({
        interessado: historyObj.nome,
        municipio: `MUNICÍPIO: ${historyObj.municipio}`,
        convenio: `CONVÊNIO: ${valores.informacaoAnalise.tipoDeConvenio}`,
        numSolicitacao: numSolicitacao,
        dataEntrada: `                DATA DE ENTRADA: ${data}`,
        obsAoLab: historyObj.infoAoLab,
        col1: historyObj.obsCol1,
        col2: historyObj.obsCol2,
        comRecomendacao: historyObj.comRecomendacao,
        comUmidade: historyObj.comUmidade,
        comGranulometria: historyObj.comGranulometria,
        areiaGrossa:  historyObj.comAreiaGrossa,
        areiaFina: historyObj.comAreiaFina,
        comArgilaNatural: historyObj.comArgila,
        comCondutividade: historyObj.comCondutividade ,
        comDensidade:historyObj.comDensidade,
        qtdAnalises: `QUANTIDADE DE ANÁLISES: ${valores.informacaoAnalise.qtdAnalises}`
      })


      erro.messagem = erroMsgFisica

    break;
  }


  console.log('parte final com as ui')
  if (erro.messagem){
    ui.showModelessDialog(
          HtmlService.createHtmlOutput(`
            <div>
              <h3>Ocorreu um erro inesperado ao enviar a solicitação.</h3>

              <p>ERRO:</p>
              <p>${erro.messagem}</p>
              
            </div>
          `).setWidth(500).setHeight(250),
          "Envio de SOLO"
        )

        botaoEnviarSolo.setOnAction('enviarSolo')

        return
  }


   ui.showModelessDialog(
        HtmlService.createHtmlOutput(`
          <div>
            <h3>A operação solicitada foi FINALIZADA.</h3>

            <p>Parabéns por aguardar até aqui kkk</p>
          </div>
        `).setWidth(500).setHeight(250),
        "Envio de SOLO"
  )

  console.log('terminou com as ui')
  botaoEnviarSolo.setOnAction('enviarSolo')
 
  
  
}