// armazena a planilha corrente em uma variavel
var ss 		= SpreadsheetApp.getActiveSpreadsheet();
var sheet 	= ss.getActiveSheet();
var ui 		= SpreadsheetApp.getUi();

//--------------------------------------------------------
//Funcao Principal, chamada de um botao criado na planilha
//--------------------------------------------------------
function Main() {
  // Busca os dados da tabela ativa
  var aDados = getValores(sheet);
  // Pergunta de teste para confirmar a execucao
  var response = ui.alert('Executar', 'Deseja calcular os valores', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    incluiVal(aDados);
  }
  
  Logger.log(aDados); // Loga os valores 
}

//-----------------------------------------------
// Busca todos os dados da coluna
//-----------------------------------------------
function getValores(){
  // funcao getRange row, column, numRows, numColumns
  return sheet.getRange(2,1,sheet.getLastRow(), ss.getLastColumn()).getValues();
}

//-----------------------------------------------
// Calcula e Inclui os valores nas outras celulas
//-----------------------------------------------
function incluiVal(aDados){
 
  var nSoma 	= 0;
  var nCout 	= 0;
  var nMaior	= 0;
  var nMenor	= 99;
  var sLastWord = "";
  var sFirsWord = "";
  var nTamStr 	= 0;
  var nQtdStr 	= 0;
  
  // Varre todas as celulas para calcular os valores
  for(i=0; i< aDados.length ; i++ ){
    if (aDados[i][1] != "" ) {

      // Regra de negocio
      nSoma += aDados[i][1]; // Soma
      nCout++;
      if (aDados[i][1] > nMaior){
        nMaior = aDados[i][1]; 
      }
      if (aDados[i][1] < nMenor){
        nMenor = aDados[i][1]; 
      }
      nTamStr = aDados[i][0].length;
      sFirsWord += aDados[i][0].substring(0,1);	
      sLastWord += aDados[i][0].substring(nTamStr-1,nTamStr);
      // busca se contem uma parte da string
      if (aDados[i][0].toUpperCase().includes("EL")){
          nQtdStr++;
      }
      
    }
  }
  
  // Alimenta as celulas com o resultado, todos na coluna 7 da planilha
  sheet.getRange(2,7,1,1).setValue(nSoma);
  sheet.getRange(3,7,1,1).setValue(nSoma/nCout);
  sheet.getRange(4,7,1,1).setValue(nMaior);
  sheet.getRange(5,7,1,1).setValue(nMenor);
  sheet.getRange(6,7,1,1).setValue(sFirsWord);
  sheet.getRange(7,7,1,1).setValue(sLastWord);
  sheet.getRange(8,7,1,1).setValue(nQtdStr);

  return;
}

//-----------------------------------
// Limpa os valores
//-----------------------------------
function ClearCells(){
  
  // funcao getRange row, column, numRows, numColumns
  // limpa os valores células
  sheet.getRange(2,7,7,1).clearContent();
  
  return;
}