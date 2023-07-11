
// FIRED BY ACTIONS
export function enviarEmailComLaudoSolos() {

  
  const planilha = SpreadsheetApp.getActive()
  SpreadsheetApp.flush()

  const aba = planilha.getActiveSheet()

   ui.showModelessDialog(
    HtmlService.createHtmlOutput(`
      <div>
        <h3>A operação solicitada foi iniciada...</h3>

        <p>Por favor, aguarde antes de clicar em outro botão</p>
      </div>
    `).setWidth(1600).setHeight(400),
    "Envio de emails"
  )

  const checkBoxArray = aba.getRange(`A4:A${aba.getLastRow()}`).getValues().flat()
 

  checkBoxArray.forEach((checkBox,index) => {

   const linha = index + 4
   

    if(checkBox){

      const nome = aba.getRange(`F${linha}:F${linha}`).getValue()
      const email = aba.getRange(`H${linha}:H${linha}`).getValue()
      console.log(email,'email')
      const linkDaPasta = aba.getRange(`I${linha}:I${linha}`).getValue()
      const empresaVinculo = aba.getRange(`J${linha}:J${linha}`).getValue()

      const subject = empresaVinculo ? 
        `O laudo do seu cliente, ${nome} está pronto.` 
          : 
        `${nome}, seu laudo está pronto.`


      const htmlBody = empresaVinculo ?
        `<h3>${empresaVinculo}, <br/>
          Atualizamos a pasta que contém os laudos de seu cliente ${nome}:</h3>
          <a href="${linkDaPasta}" target="_blank"><h3>Clique aqui para acessar sua pasta</h3></a>
          <p><em>Este link é direcionado apenas a você e concede acesso a seus laudos,recomendamos que <strong>Não</strong> o compartilhe com mais      ninguém.</p>
          <br/>
          <p>Caso possua alguma dúvida, favor entrar em contato com:</p>
          <ul><li><a href="tel:00000000000">00 00000-0000</a></li>
          <li><a href:"mailto:email@email.com">email@email.com</a></li></ul>

          <p>Agradecemos a preferência e somos gratos por te ter como parceiro (a)</p>
          <br/>
          <h5>Aproveitamos para te convidar a nos seguir em nosso <a href="https://www.instagram.com/fundenor.org132/" target="_blank">instagram</a> e ficar por dentro das novidades: <a href="https://www.instagram.com/fundenor.org132/" target="_blank"><em>@fundenor.org132</em></a></h5>`

          :

          `<h3>${nome},<br/>
          Atualizamos a pasta que contém seus laudos:</h3>
          <a href="${linkDaPasta}" target="_blank"><h3>Clique aqui para acessar sua pasta</h3></a>
          <p><em>Este link é direcionado apenas a você e concede acesso a seus laudos,recomendamos que <strong>Não</strong> o compartilhe com mais      ninguém.</p>
          <br/>
          <p>Caso possua alguma dúvida, favor entrar em contato com:</p>
          <ul><li><a href="tel:00000000000">00 00000-0000</a></li>
          <li><a href:"mailto:email@email.com">email@email.com</a></li></ul>

          <p>Agradecemos a preferência e somos gratos por te ter como parceiro (a)</p>
          <br/>
          <h5>Aproveitamos para te convidar a nos seguir em nosso <a href="https://www.instagram.com/fundenor.org132/" target="_blank">instagram</a> e ficar por dentro das novidades: <a href="https://www.instagram.com/fundenor.org132/" target="_blank"><em>@fundenor.org132</em></a></h5>
        `

    
      try{
        MailApp.sendEmail({
          to: email,
          bcc: "email@email.com",
          subject,
          htmlBody,
          replyTo: 'email@email.com'
        })

        aba.getRange(`A${linha}`).setValue(false)
      }
      catch(err){
        console.log(err)
      }
      
    }
  })

  ui.showModelessDialog(
      HtmlService.createHtmlOutput(`
        <div>
          <h3>A operação finalizada...</h3>

          <p>Pode prosseguir com sua vida kkkk</p>
        </div>
      `).setWidth(1600).setHeight(400),
      "Envio de emails"
    )


}