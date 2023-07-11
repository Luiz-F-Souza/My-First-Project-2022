import { criarLinha } from "../../utils/createNewLine.js"
import { prevenirDuploClick } from "../../utils/preventDoubleClick.js"
import { enviarBacteriologica } from './bacteriologicalWater.js'
import { enviarFisicoQuimica } from "./physicalWater.js"

export function enviarAgua() {

  SpreadsheetApp.flush()

  const ui = SpreadsheetApp.getUi()

  ui.showModelessDialog(
    HtmlService.createHtmlOutput(`
      <div>
        <h3>A operação solicitada foi iniciada...</h3>

        <p>Por favor, aguarde cerca de 10 segundos antes de clicar em outro botão</p>
      </div>
    `).setWidth(1600).setHeight(400),
    "Envio de ÁGUA"
  )

  const planilha = SpreadsheetApp.openById('ID')

  const abaDados = planilha.getSheetByName('ENVIAR ÁGUAS')

  const abaHistoricoBacteriologica = planilha.getSheetByName('ÁGUA BACTERIOLÓGICA')
  const abaHistoricoFisicoQuimica = planilha.getSheetByName('ÁGUA FÍSICO-QUÍMICA')

  const botaoEnviarAgua = abaDados.getDrawings()[0]
  if(botaoEnviarAgua.getOnAction() === 'prevenirDuploClick') return prevenirDuploClick()
  botaoEnviarAgua.setOnAction('prevenirDuploClick')
  


  const dataObj = new Date()

  const data = dataObj.toLocaleDateString('pt-BR',{month:'2-digit',day:'2-digit',year:'numeric'})
  const hora = dataObj.getHours().toLocaleString('pt-BR',{minimumIntegerDigits:2})
  const minuto = dataObj.getMinutes().toLocaleString('pt-BR',{minimumIntegerDigits:2})
  const ano = dataObj.getFullYear()

  const formatoDataHoraLaudo = `${data} - ${hora}h${minuto}min`
  
  const celulasComDados = {
    cpf: abaDados.getRange('B3:D3'),
    cnpj: abaDados.getRange('F3:H3'),
    nome: abaDados.getRange('L3:P3'),
    email: abaDados.getRange('L5:P5'),
    celular: abaDados.getRange('L7:P7'),
    empresaVinculo: abaDados.getRange('L9:P9'),
    linkDaPastaNoDrive: abaDados.getRange('L11:P11'),
    qtdAnalises: abaDados.getRange('B6:D6'),
    convenio: abaDados.getRange('F6:H6'),
    tipoDeAnalise: abaDados.getRange('B9:D9'),
    subTipoDeAnalise: abaDados.getRange('F9:H9'),
    comDureza: abaDados.getRange('B12:C12'),
    comChumboECadimio: abaDados.getRange('D12:F12'),
    origemDaAmostra: abaDados.getRange('G12:H12'),
    localDaColeta: abaDados.getRange('B15:H15'),
    dataDaColeta: abaDados.getRange('B18:D18'),
    horaDaColeta: abaDados.getRange('F18:H18'),
    responsavelPelaColeta: abaDados.getRange('B21:H21'),
    enderecoDaColeta: {
      ruaOuAvenida: abaDados.getRange('B24:B24'),
      nomeDaRuaOuAvenida: abaDados.getRange('C24:H24'),
      numero: abaDados.getRange('C25:H25'),
      bairro: abaDados.getRange('C26:H26'),
      municipio: abaDados.getRange('C27:H27')
    },
    obsAoLab: abaDados.getRange('B30:H35')
  }

  
  const dados = {
    cpf: celulasComDados.cpf.getValue(),
    cnpj: celulasComDados.cnpj.getValue(),
    nome: celulasComDados.nome.getValue(),
    email: celulasComDados.email.getValue(),
    celular: celulasComDados.celular.getValue(),
    empresaVinculo: celulasComDados.empresaVinculo.getValue(),
    linkDaPastaNoDrive: celulasComDados.linkDaPastaNoDrive.getValue(),
    qtdAnalises: celulasComDados.qtdAnalises.getValue(),
    convenio: celulasComDados.convenio.getValue().trim(),
    tipoDeAnalise: celulasComDados.tipoDeAnalise.getValue(),
    subTipoDeAnalise: celulasComDados.subTipoDeAnalise.getValue(),
    comDureza: celulasComDados.comDureza.getValue(),
    comChumboECadimio: celulasComDados.comChumboECadimio.getValue(),
    origemDaAmostra: celulasComDados.origemDaAmostra.getValue(),
    localDaColeta: celulasComDados.localDaColeta.getValue().trim(),
    dataDaColeta: celulasComDados.dataDaColeta.getValue().toLocaleDateString('pt-BR',{month:'2-digit',day:'2-digit',year:'numeric'}),
    horaDaColeta: celulasComDados.horaDaColeta.getValue().toString().trim(),
    responsavelPelaColeta: celulasComDados.responsavelPelaColeta.getValue().trim(),
    enderecoDaColeta: {
      ruaOuAvenida: celulasComDados.enderecoDaColeta.ruaOuAvenida.getValue(),
      nomeDaRuaOuAvenida: celulasComDados.enderecoDaColeta.nomeDaRuaOuAvenida.getValue().trim(),
      numero: celulasComDados.enderecoDaColeta.numero.getValue().toString().trim(),
      bairro: celulasComDados.enderecoDaColeta.bairro.getValue().trim(),
      municipio: celulasComDados.enderecoDaColeta.municipio.getValue().trim()
    },
    obsAoLab: celulasComDados.obsAoLab.getValue()

  }

  const enderecoDaColeta = `${dados.enderecoDaColeta.ruaOuAvenida} ${dados.enderecoDaColeta.nomeDaRuaOuAvenida}, ${dados.enderecoDaColeta.numero} - ${dados.enderecoDaColeta.bairro} - ${dados.enderecoDaColeta.municipio} `


  if(dados.cpf && dados.cnpj){

     ui.showModelessDialog(
      HtmlService.createHtmlOutput(`
        <div>
          <h3>A operação solicitada foi finalizada com erro.</h3>

          <p>Os campos de cpf e cnpj estão marcados ao mesmo tempo.</p>

        </div>
      `).setWidth(500).setHeight(250),
    "Envio de ÁGUA"
    )


    botaoEnviarAgua.setOnAction('enviarAgua')
    return Error('cpf e cnpj selecionados')
  }
  if(!dados.cpf && !dados.cnpj){

     ui.showModelessDialog(
      HtmlService.createHtmlOutput(`
        <div>
          <h3>A operação solicitada foi finalizada com erro.</h3>

          <p>Os campos de cpf e cnpj estão vazios. Selecione um dos dois antes de enviar.</p>

        </div>
      `).setWidth(500).setHeight(250),
    "Envio de ÁGUA"
    )


    botaoEnviarAgua.setOnAction('enviarAgua')
    return Error('cpf e cnpj selecionados')
  }

  const objParaCadastro = {
    enviar: null,
    dataEHoraDeEntradaNoLab: formatoDataHoraLaudo,
    numSolicitacao: null,
    cpfOuCnpj: dados.cpf ? dados.cpf : dados.cnpj,
    nome: dados.nome,
    telefone: dados.celular,
    email: dados.email,
    linkDaPasta: dados.linkDaPastaNoDrive,
    empresaVinculo: dados.empresaVinculo,
    tipoDeAnalise: dados.subTipoDeAnalise,

  }

  let prefixoDaNumeracao 
  let abaSelecionada
  let endColumn

  switch(dados.tipoDeAnalise){
    case 'BACTERIOLÓGICA':
      objParaCadastro.origemDaAmostra = dados.origemDaAmostra
      objParaCadastro.localDaColeta = dados.localDaColeta
      objParaCadastro.dataDaColeta = dados.dataDaColeta
      objParaCadastro.horaDaColeta = dados.horaDaColeta
      objParaCadastro.responsavelPelaColeta = dados.responsavelPelaColeta
      objParaCadastro.enderecoDaColeta = enderecoDaColeta
      objParaCadastro.convenio = dados.convenio
      objParaCadastro.obsAoLab = dados.obsAoLab
      objParaCadastro.emailDoUsuario = emailDoUsuario

      prefixoDaNumeracao = "BA/"
      abaSelecionada = abaHistoricoBacteriologica
      endColumn = "S"
    break;

    case 'ÁGUA FÍSICO - QUÍMICA':
      objParaCadastro.qtdAnalise = dados.qtdAnalises
      objParaCadastro.comDureza = dados.comDureza
      objParaCadastro.comChumboEcadmio = dados.comChumboECadimio
      objParaCadastro.localDaColeta = dados.localDaColeta
      objParaCadastro.responsavelPelaColeta = dados.responsavelPelaColeta
      objParaCadastro.enderecoDaColeta = enderecoDaColeta
      objParaCadastro.convenio = dados.convenio
      objParaCadastro.obsAoLab = dados.obsAoLab
      objParaCadastro.emailDoUsuario = emailDoUsuario

      prefixoDaNumeracao = "AC/"
      abaSelecionada = abaHistoricoFisicoQuimica
      endColumn = "T"
    break;
  }


  const { foiCriada, numeroDaLinhaCriada } = criarLinha({sheet: abaSelecionada, dadosObj: objParaCadastro, startColumn: "A", endColumn})
  
  if(!foiCriada) {
    console.log(Error('Não foi possível criar linha no envio da água'))

    ui.showModelessDialog(
    HtmlService.createHtmlOutput(`
      <div>
        <h3>A operação solicitada foi FINALIZADA.</h3>

        <p>Deu ruim, não conseguimos criar a linha da análise</p>

      </div>
    `).setWidth(500).setHeight(250),
    "Envio de ÁGUA"
    )

    return
  }

  // COLOCANDO O CHECK BOX NA COLUNA A (O CHECKBOX É USADO PARA ENVIAR EMAILS AOS CLIENTES QUE TIVEREM ESTE CAMPO MARCADO)
  abaSelecionada.getRange(`A${numeroDaLinhaCriada}:A${numeroDaLinhaCriada}`).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox())

  // COLOCANDO O Nº DA SOLICITAÇÃO NA COLUNA C. O - 3 É PQ A LINHA INICIAL DA CONTAGEM É A 4, DAI VOLTA O VALOR PARA 1
  const numSolicitacao = `${prefixoDaNumeracao}${(numeroDaLinhaCriada - 3).toLocaleString(undefined,{minimumIntegerDigits: 4, useGrouping: false})}.${ano}`

  abaSelecionada.getRange(`C${numeroDaLinhaCriada}:C${numeroDaLinhaCriada}`)
  .setValue(numSolicitacao)

  const erros = {
    abaPrincipal: null,
    abaRascunho: null
  }

  switch(dados.tipoDeAnalise){

    case 'BACTERIOLÓGICA':
      const { erroBacteriologica } = enviarBacteriologica(dados.subTipoDeAnalise,objParaCadastro,numSolicitacao)

      erros.abaPrincipal = erroBacteriologica.abaPrincipal
      erros.abaRascunho = erroBacteriologica.abaRascunho
    break;

    case 'ÁGUA FÍSICO - QUÍMICA':

      const { erroFisicoQuimica } = enviarFisicoQuimica(dados.subTipoDeAnalise,dados,numSolicitacao)

      erros.abaPrincipal = erroFisicoQuimica.abaPrincipal
      erros.abaRascunho = erroFisicoQuimica.abaRascunho

    break;
  }
  


  if(erros.abaPrincipal || erros.abaRascunho){
       ui.showModelessDialog(
          HtmlService.createHtmlOutput(`
            <div>
              <h3>Ocorreu um erro inesperado ao enviar a solicitação.</h3>

              <p>ERRO:</p>
              <p>Criar laudo: ${erros.abaPrincipal}</p>
              <p>Criar rascunho: ${erros.abaRascunho}</p>
            </div>
          `).setWidth(500).setHeight(250),
          "Envio de ÁGUA"
        )

      botaoEnviarAgua.setOnAction('enviarAgua')
  
      return
  }

  ui.showModelessDialog(
        HtmlService.createHtmlOutput(`
          <div>
            <h3>A operação solicitada foi FINALIZADA.</h3>

            <p>Parabéns por aguardar até aqui kkk</p>
          </div>
        `).setWidth(500).setHeight(250),
        "Envio de ÁGUA"
  )

  botaoEnviarAgua.setOnAction('enviarAgua')
  

}