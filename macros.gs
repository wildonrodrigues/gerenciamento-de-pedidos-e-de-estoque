function a2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('F7').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por SIAD'), true);
  spreadsheet.getRange('F7').activate();
  bs();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Cadastrar'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('impressao'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DISTRIBUIÇÃO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Balanço por data'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Lançamentos'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por SIAD'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Buscas combinadas'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Movimentação da posição'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LISTA DE PRODUTOS'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Ocupação'), true);
  spreadsheet.getActiveSheet().hideSheet();
};

function a1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('F7').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Cadastrar'), true);
  spreadsheet.getRange('F7').activate();
  ca();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('impressao'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DISTRIBUIÇÃO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Balanço por data'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Lançamentos'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por SIAD'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Buscas combinadas'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Movimentação da posição'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LISTA DE PRODUTOS'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Ocupação'), true);
  spreadsheet.getActiveSheet().hideSheet();
};



function a3() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('I10').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
  spreadsheet.getRange('I10').activate();
  bl();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Cadastrar'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('impressao'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DISTRIBUIÇÃO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Balanço por data'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por SIAD'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Lançamentos'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Buscas combinadas'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Movimentação da posição'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LISTA DE PRODUTOS'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Ocupação'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.getSheetByName('Lançamentos').showSheet()
  .activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Lançamentos'), true);
  spreadsheet.getSheetByName('Ocupação').showSheet()
  .activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Ocupação'), true);
  spreadsheet.getSheetByName('Busca por lote').showSheet()
  .activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
};

function a4() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('L12').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Movimentação da posição'), true);
  spreadsheet.getRange('L12').activate();
  mp();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Balanço por data'), true);
  spreadsheet.getActiveSheet().hideSheet()
  .hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por SIAD'), true);
  spreadsheet.getActiveSheet().hideSheet()
  .hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
  spreadsheet.getActiveSheet().hideSheet()
  .hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Lançamentos'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Buscas combinadas'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LISTA DE PRODUTOS'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Ocupação'), true);
  spreadsheet.getActiveSheet().hideSheet();
};

function a5() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('O15').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Buscas combinadas'), true);
  spreadsheet.getRange('O15').activate();
  bcc();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Lançamentos'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Movimentação da posição'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LISTA DE PRODUTOS'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Ocupação'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
  spreadsheet.getActiveSheet().hideSheet()
  .hideSheet()
  .hideSheet()
  .hideSheet()
  .hideSheet()
  .hideSheet();
};

function a6() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E21').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DISTRIBUIÇÃO'), true);
  spreadsheet.getRange('E21').activate();
  disl();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Cadastrar'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('impressao'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Balanço por data'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por SIAD'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Lançamentos'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Buscas combinadas'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Movimentação da posição'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LISTA DE PRODUTOS'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Ocupação'), true);
  spreadsheet.getActiveSheet().hideSheet();
};

function a7() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E21').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Balanço por data'), true);
  spreadsheet.getRange('E21').activate();
  bd();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DISTRIBUIÇÃO'), true);
  spreadsheet.getActiveSheet().hideSheet()
  .hideSheet()
  .hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por SIAD'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Lançamentos'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Buscas combinadas'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Movimentação da posição'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LISTA DE PRODUTOS'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Ocupação'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Cadastrar'), true);
  spreadsheet.getActiveSheet().hideSheet();
};

function A8() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('N16').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Lançamentos'), true);
  spreadsheet.getRange('N16').activate();
  la();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Cadastrar'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('impressao'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DISTRIBUIÇÃO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Balanço por data'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por SIAD'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Buscas combinadas'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Movimentação da posição'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LISTA DE PRODUTOS'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Ocupação'), true);
  spreadsheet.getActiveSheet().hideSheet();
};

function a9() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('M24').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LISTA DE PRODUTOS'), true);
  spreadsheet.getRange('M24').activate();
  lp();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Movimentação da posição'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Ocupação'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Buscas combinadas'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Lançamentos'), true);
  spreadsheet.getActiveSheet().hideSheet()
  .hideSheet()
  .hideSheet()
  .hideSheet()
  .hideSheet()
  .hideSheet();
};

function a10() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('M24').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Ocupação'), true);
  spreadsheet.getRange('M24').activate();
  oc();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
  spreadsheet.getActiveSheet().hideSheet()
  .hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Buscas combinadas'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Movimentação da posição'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Lançamentos'), true);
  spreadsheet.getActiveSheet().hideSheet()
  .hideSheet()
  .hideSheet()
  .hideSheet()
  .hideSheet()
  .hideSheet();
};

function a11() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('M24').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
  spreadsheet.getRange('M24').activate();
  re();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por lote'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Balanço por data'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Busca por SIAD'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('impressao'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DISTRIBUIÇÃO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Lançamentos'), true);
  spreadsheet.getActiveSheet().hideSheet()
  .hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Buscas combinadas'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Movimentação da posição'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LISTA DE PRODUTOS'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Ocupação'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.getSheetByName('RECEBIMENTO').showSheet()
  .activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RECEBIMENTO'), true);
};
