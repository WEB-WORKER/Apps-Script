function switchers(x,y,sheetname){
  var regex = new RegExp(x,'g');
  var letters = ["A", "F", "G", "H"];
  var i, j;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetname);
  var lastrow = sheet.getLastRow();
  for (j in letters){
     for (i = 2; i < lastrow; i++){
     var src = sheet.getRange(letters[j]+i).getFormula().toString();
     var trg = sheet.getRange(letters[j]+i); 
     var ch = src.replace(regex, y);
     trg.setValue(ch);
     }
   } 
 }

function contfresh(){
    switchers("Настройки Контекст 1", "Настройки Контекст 2", "Индикаторы Контекст");  
}
function contold(){
    switchers("Настройки Контекст 2", "Настройки Контекст 1", "Индикаторы Контекст");  
}
function seofresh(){
    switchers("Настройки SEO 1", "Настройки SEO 2", "Индикаторы SEO");
}
function seoold(){
    switchers("Настройки SEO 2", "Настройки SEO 1", "Индикаторы SEO");
}

  
