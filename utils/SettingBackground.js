export async function setandoBackgroundElemento(aba,celula,analisarElemento){
  const cor = analisarElemento ? analisarElementoBackground : 'fff'
  await aba.getRange(celula).setBackground(cor)
}