import { prevenirDuploClick } from './utils/preventDoubleClick.js'
import { criarLinha } from './utils/createNewLine.js'
import { criarPastaNoDrive } from "./drive/creatingClientFolder.js"


// FIRED BY ACTIONS
export function cadastrarCliente(){

  SpreadsheetApp.flush()


  const planilhas = {
    cadastros: planilhaRecepcao.getSheetByName('CADASTROS'),
    clientes: planilhaRecepcao.getSheetByName('CLIENTES'),
    empresas: planilhaRecepcao.getSheetByName('EMPRESAS DE COLETA')
  }

  const botaoCadastrar = planilhas.cadastros.getDrawings()[0]

  if(botaoCadastrar.getOnAction() === "prevenirDuploClick") return prevenirDuploClick()

   ui.showModelessDialog(
    HtmlService.createHtmlOutput(`
      <div>
        <h3>A operação solicitada foi iniciada...</h3>

        <p>Por favor, aguarde cerca de 10 segundos antes de clicar em outro botão</p>
      </div>
    `).setWidth(1600).setHeight(400),
    "Cadastro de cliente / empresa"
  )

  
  botaoCadastrar.setOnAction('prevenirDuploClick')
  
  

  const celulasComValores = {
    nome: planilhas.cadastros.getRange("B5:F5"),
    cpf: planilhas.cadastros.getRange("H5:L5"),
    email: planilhas.cadastros.getRange("B8:F8"),
    celular: planilhas.cadastros.getRange("H8:L8"),
    tipoDeCadastro: planilhas.cadastros.getRange("D10:F10"),
    empresaVinculo: planilhas.cadastros.getRange("J10:L10")
  }
  
  const data = new Date().toLocaleDateString('pt-BR',{month:'2-digit',day:'2-digit',year:'numeric',hour:'2-digit', minute:'2-digit'})
  const tipoDeCadastro = celulasComValores.tipoDeCadastro.getValue()

  const valoresCadastrados = {
    id: planilhas.clientes.getLastRow(),
    cpf: celulasComValores.cpf.getValue().toString().trim(),
    nome: celulasComValores.nome.getValue().trim(),
    email: celulasComValores.email.getValue().trim(),
    celular: celulasComValores.celular.getValue().toString().trim(),
    empresaDeColeta: celulasComValores.empresaVinculo.getValue(),
    linkDaPastaNoDrive: null,
    idDaPastaNoDrive: null,
    responsavelPeloCadastro: emailDoUsuario,
    data
  }

  const valoresCadastradosParaEmpresa = {
      id: planilhas.empresas.getLastRow() ,
      cnpj: valoresCadastrados.cpf,
      nome: valoresCadastrados.nome,
      email: valoresCadastrados.email,
      celular: valoresCadastrados.celular,
      colunaOculta: null, // criei apenas para manter o mesmo nº de colunas que a pag de cliente e assim facilitar a inserção de dados
      linkDaPastaNoDrive: null,
      idDaPastaNoDrive: null,
      responsavelPeloCadastro: emailDoUsuario,
      data,
      nomeNaListagem: valoresCadastrados.nome, // é uma coluna extra que duplica o nome da empresa, fiz isso pq pra puxar o nome na hr de cadastrar os cliente precisamos do "Nenhuma" aparecendo como opção em 'Empresa vínculo'
  }

  // VERIFICANDO SE TODOS OS DADOS ESTÃO CORRETOS

    const validandoDados = function(estaCorreto,celula,erroMsg){

      if(estaCorreto) return celula.setBorder(true,true,true,true,false,false,'#b7b7b7',SpreadsheetApp.BorderStyle.SOLID)

      celula.setBorder(true,true,true,true,false,false,'red',SpreadsheetApp.BorderStyle.SOLID)
      botaoCadastrar.setOnAction('cadastrarCliente')

       ui.showModelessDialog(
        HtmlService.createHtmlOutput(`
          <div>
            <h3>Não foi possível concluir o cadastro</h3>

            <p>${erroMsg}</p>
          </div>
        `).setWidth(800).setHeight(400),
        "Cadastro de cliente / empresa"
      )
    }
  
    // NOME
    if(valoresCadastrados.nome.length < 7) return validandoDados(false,celulasComValores.nome,"DIGITE O NOME COMPLETO")
    validandoDados(true,celulasComValores.nome)

    // CPF / CNPJ
    if(!valoresCadastrados.cpf.match(/([0-9]){11,14}/gi)) return validandoDados(false,celulasComValores.cpf,"CNPJ OU CPF INCORRETO!")
    validandoDados(true,celulasComValores.cpf)
      //FORMATANDO O CPF / CNPJ 
        if(valoresCadastrados.cpf.length === 14){
          const novoValor = valoresCadastrados.cpf.split('')
          novoValor.splice(2,0,".")
          novoValor.splice(6,0,".")
          novoValor.splice(10,0,"/")
          novoValor.splice(15,0,"-")
          valoresCadastrados.cpf = novoValor.join('')
          valoresCadastradosParaEmpresa.cnpj = valoresCadastrados.cpf
          } 

          if(valoresCadastrados.cpf.length === 11){
            const novoValor = valoresCadastrados.cpf.split('')
            novoValor.splice(3,0,".")
            novoValor.splice(7,0,".")
            novoValor.splice(11,0,"-")
           
            valoresCadastrados.cpf = novoValor.join('')
          }
      

    // EMAIL
    if(!valoresCadastrados.email.match(/([a-z0-9._-]){3,}@([a-z0-9]){3,}\.([a-z]){2,}/gi)) return validandoDados(false,celulasComValores.email,"DIGITE UM EMAIL VÁLIDO.")
    validandoDados(true,celulasComValores.email) 

    // CELULAR
    if(!valoresCadastrados.celular.match(/([0-9]){9,14}/gi)) return validandoDados(false,celulasComValores.celular,"DIGITE UM CELULAR VÁLIDO")
    validandoDados(true,celulasComValores.celular) 

      // FORMATANDO CELULAR PARA (xx) xxxxx-xxx
      valoresCadastrados.celular = valoresCadastrados.celular.split('')
      valoresCadastrados.celular.splice(0,0,"(")
      valoresCadastrados.celular.splice(3,0,")")
      valoresCadastrados.celular.splice(4,0," ")
      valoresCadastrados.celular.splice(9,0,"-")
      valoresCadastrados.celular = valoresCadastrados.celular.join('')
    

    // TIPO DE CADASTRO
    if(tipoDeCadastro === "") return validandoDados(false,celulasComValores.tipoDeCadastro,"SELECIONE O TIPO DE CADASTRO")
    validandoDados(true,celulasComValores.tipoDeCadastro) 

    // EMPRESA VINCULO
    if(tipoDeCadastro === "CLIENTE") 
      if(valoresCadastrados.empresaDeColeta === "") return validandoDados(false,celulasComValores.empresaVinculo,"SE NÃO DESEJA VINCULAR A NENHUMA EMPRESA, MARQUE 'NENHUMA'.")
    validandoDados(true,celulasComValores.empresaVinculo) 


  
  // FIM DA VERIFICAÇÃO


  const planilhaAlvo = tipoDeCadastro === "CLIENTE" ? planilhas.clientes : planilhas.empresas

  const objParaCadastro = tipoDeCadastro === "CLIENTE" ? valoresCadastrados : valoresCadastradosParaEmpresa

  // CRIANDO LINHA (APENAS SE TODOS OS DADOS ACIMA ESTIVEREM CORRETOS)

  const cpfsCadastrados = planilhaAlvo.getRange(`B1:B`).getValues().flat()
  console.log(cpfsCadastrados, 'CPFS CADASTRADOS')
  console.log(cpfsCadastrados.includes(valoresCadastrados.cpf),"INCLUDES ??")

  if(cpfsCadastrados.includes(valoresCadastrados.cpf)){
      botaoCadastrar.setOnAction('cadastrarCliente')

       ui.showModelessDialog(
        HtmlService.createHtmlOutput(`
          <div>
            <h3>Não foi possível concluir o cadastro</h3>

            <p>CPF / CNPJ JÁ CADASTRADO !! </p>
            <P>LINHA CADASTRADA: ${cpfsCadastrados.indexOf(valoresCadastrados.cpf)}</p>
          </div>
        `).setWidth(800).setHeight(400),
        "Cadastro de cliente / empresa"
      )

    return Error(`Cliente já cadastrado, LINHA: ${cpfsCadastrados.indexOf(valoresCadastrados.cpf)}`)
  } 
  
  const { foiCriada,numeroDaLinhaCriada } = criarLinha({sheet: planilhaAlvo ,dadosObj: objParaCadastro, startColumn: 'A', endColumn: 'J'})
  if(!foiCriada) return botaoCadastrar.setOnAction('cadastrarCliente')



  // CRIANDO PASTA NO DRIVE 
  const nomeDaPastaNoDrive = `${valoresCadastrados.nome} ( ${valoresCadastrados.cpf} )`

  const linhaDaEmpresaVinculo = tipoDeCadastro === "CLIENTE" ? 
    planilhas.empresas.getRange(`C2:C${planilhas.empresas.getLastRow()}`).getValues().flat().indexOf(valoresCadastrados.empresaDeColeta)  + 2
    : 
    null
 
  // questiono aqui se é maior que 1 pq se for 1 significa que não encontrou a empresa e retornou o cabeçalho
  const idEmpresaVinculo = linhaDaEmpresaVinculo > 1  ? 
    planilhas.empresas.getRange(`H${linhaDaEmpresaVinculo}:H${linhaDaEmpresaVinculo}`).getValue()
    :
    null

  const { linkDaPasta, idDaPasta } = criarPastaNoDrive(tipoDeCadastro,nomeDaPastaNoDrive, idEmpresaVinculo)

  // COLOCANDO LINK DO DRIVE E ID DA PASTA NOS CAMPOS CORRETOS
  planilhaAlvo.getRange(`G${numeroDaLinhaCriada}:G${numeroDaLinhaCriada}`).setValue(linkDaPasta)
  planilhaAlvo.getRange(`H${numeroDaLinhaCriada}:H${numeroDaLinhaCriada}`).setValue(idDaPasta)

  // LIMPANDO DADOS APÓS CADASTRO COM SUCESSO
  for(const key in celulasComValores){
    celulasComValores[key].setValue('')
  }
  
  
       ui.showModelessDialog(
        HtmlService.createHtmlOutput(`
          <div>
              <h3>OPERAÇÃO FINALIZADA COM SUCESSO!!</h3>
          </div>
        `).setWidth(800).setHeight(400),
        "Cadastro de cliente / empresa"
      )
  planilhaAlvo.getRange(`A2:J${numeroDaLinhaCriada}`).sort({column: 3, ascending: true})
  botaoCadastrar.setOnAction('cadastrarCliente')
}