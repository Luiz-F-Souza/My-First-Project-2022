export function prevenirDuploClick(){

  const ui = SpreadsheetApp.getUi()
  ui.alert("Multiplos cliques simultâneos.","Se acalme mulher, você acabou de apertar. Espera terminar o processo, toma um café e curte o clima!!!",ui.ButtonSet.OK)

  return
}