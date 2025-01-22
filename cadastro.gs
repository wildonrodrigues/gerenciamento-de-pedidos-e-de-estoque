var ss =  SpreadsheetApp.getActiveSpreadsheet();

function lanc() {

  var abaCadastrar =  ss.getSheetByName('Cadastrar');
  var abaa = ss.getSheetByName('Lançamentos');

  var dadosCadastrar = abaCadastrar.getRange('d1:d20').getValues();

  var data= dadosCadastrar[0][0]
  var  lac = dadosCadastrar[1][0]
  var  tip = dadosCadastrar[3][0]
  var  pos = dadosCadastrar[7][0]
  var  siad = dadosCadastrar[4][0]
  var  pro = dadosCadastrar[5][0]
  var  lote = dadosCadastrar[6][0]
  var  und = dadosCadastrar[10][0]
  var  qpc = dadosCadastrar[11][0]
  var  qdc = dadosCadastrar[12][0]
  var  qtu = dadosCadastrar[13][0]
  var  fab = dadosCadastrar[8][0]
  var  val = dadosCadastrar[9][0]
  var  vun = dadosCadastrar[14][0]
  var  vtp = dadosCadastrar[15][0]
  var  nf = dadosCadastrar[2][0]
  var  mot = dadosCadastrar[16][0]
  

  var linhaVazia = abaa.getLastRow()+1;
  var linhaVazia = abaa.getRange('c:c').getValues().filter(function (item) { return item != ""}).length + 1;

  var linhaDados = [[data, lac, tip, pos, siad, pro, lote, und, qpc, qdc, qtu, fab, val, vun, vtp, nf, mot]]
  
  abaa.getRange(linhaVazia,3,linhaDados.length,linhaDados[0].length).setValues(linhaDados)
}




