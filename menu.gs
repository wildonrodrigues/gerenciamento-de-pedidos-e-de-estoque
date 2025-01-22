function abrirOutraPagina(nomePagina) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(nomePagina);

  if (sheet) {
    spreadsheet.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('A aba "' + nomePagina + '" não foi encontrada.');
  }
}

function ca() {
abrirOutraPagina('Cadastrar');
}

function bd() {
  abrirOutraPagina('Balanço por data');
}

function la() {
  abrirOutraPagina('Lançamentos');
}

function bs() {
  abrirOutraPagina('Busca por SIAD');
}

function bl() {
  abrirOutraPagina('Busca por lote');
}

function mp() {
  abrirOutraPagina('Movimentação da posição');
}

function lp() {
  abrirOutraPagina('LISTA DE PRODUTOS');
}

function re() {
  abrirOutraPagina('RECEBIMENTO');
}

function disl() {
  abrirOutraPagina('DISTRIBUIÇÃO');
}

function oc() {
  abrirOutraPagina('Ocupação');
}

function menu() {
  abrirOutraPagina('MENU');
}

function bcc() {
  abrirOutraPagina('Buscas combinadas');
}
