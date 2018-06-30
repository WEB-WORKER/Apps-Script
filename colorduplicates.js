function colorduplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data1 = [];
  var data2 = [];
  var data3 = [];
  var data4 = []; 
  var i, j, key, x, y, rng, letter, value, lastrow;  
  var selection = ss.getSelection();
  var activeRange = selection.getActiveRange();
  
  letter = activeRange.getA1Notation().slice(0,1).toString();

  lastrow = sheet.getLastRow();

  
  for(j = 0, i = 1; i <= lastrow; i++, j++){
    rng = sheet.getRange(letter+i);
    value = rng.getValue();
    data1[j] = value;
  } 
  
  for(i in data1){ 
    data2[i] = data1[i];
  }
  data2.sort(); 
  
  for(i in data2){ 
    if (data2[i] != 0){
        data3.push(data2[i]);
    }
  }  
  
  for(i = 0; i<(data3.length-1); i++){
    x = data3[i];
    y = data3[i+1];
    if (x == y){
        data4.push(x);
      } 
  }
  
  for(i = 0, key = 1; i < data1.length; i++,key++){
     for(j in data4){
      if (data1[i] == data4[j]){
        sheet.getRange(letter+key).setBackground('#ea9999');
      }
     }    
  }
};

function clearcolor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var letter;
  var selection = ss.getSelection();
  var activeRange = selection.getActiveRange();  
  letter = activeRange.getA1Notation().slice(0,1).toString();  
  ss.getActiveRangeList().setBackground('#ffffff');  
  ss.getRange(letter+"2").setBackground('#3c78d8');
};
