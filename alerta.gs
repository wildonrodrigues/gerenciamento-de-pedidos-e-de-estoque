function alertMessage() {
  var result = SpreadsheetApp.getUi().alert("Nova funcionalidade habilitada:  balanço por período ou data. Nela, você pode consultar as quantidades e a posição de estoque atual. ⚠️⚠️⚠️ Caso for colar algum dado, utilize a função 'CRTL + SHIFT+ V' para preservar a formatação.  ");
  if(result === SpreadsheetApp.getUi().Button.OK)  {
   //Take some action
   SpreadsheetApp.getActive().toast("About to take some action ...");
   }
}
