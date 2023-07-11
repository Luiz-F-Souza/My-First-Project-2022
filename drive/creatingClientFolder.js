export function criarPastaNoDrive(tipoDoDrive,nomeDaPasta,idEmpresaVinculo){


  const drive = tipoDoDrive === "CLIENTE" ? 
      DriveApp.getFolderById('NORMAL CLIENT FOLDER ID')
      : 
      DriveApp.getFolderById('COMPANY FOLDER ID') 
      
  
  const pastaEmpresaVinculo = idEmpresaVinculo ? DriveApp.getFolderById(idEmpresaVinculo) : null

  const createdFolder = pastaEmpresaVinculo ? 
  pastaEmpresaVinculo.createFolder(nomeDaPasta)
  :
  drive.createFolder(nomeDaPasta)
  

  // COMPARTILHANDO A PASTA COM QUALQUER PESSOA COM O LINK (AVISAR AO CLIENTE QUE NÃO PODE COMPARTILHAR O LINK COM OS OUTROS)
  if(!pastaEmpresaVinculo) createdFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW)

  const linkDaPasta = createdFolder.getUrl()
  const idDaPasta = createdFolder.getId()

  


  return { idDaPasta, linkDaPasta }
  
}