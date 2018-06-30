function datetables() { 
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ascont = ss.getSheetByName("Контекст");
  var asseo = ss.getSheetByName("SEO");  
  var contlastrow = ascont.getLastRow();
  var seolastrow = asseo.getLastRow();
  
  var sites = {
     cont:{       
       idlink:{},       
       dtstartlink:{},       
     },
    seo:{      
      idlink:{},      
      dtstartlink:{},      
  },
sheetdateslink: "Текущие даты",
unitablelink: "сводная таблица"
} 
  
  for (j = 2, key = 1; key < contlastrow; j++, key++){
	sites.cont.idlink[key] = ("B"+j);
    sites.cont.dtstartlink[key] = ("C"+j);   
  }  
  
  for (j = 2, key = 1; key < seolastrow; j++, key++){
	sites.seo.idlink[key] = ("B"+j);
    sites.seo.dtstartlink[key] = ("C"+j);  
  }  
  
  var asdates = ss.getSheetByName(sites.sheetdateslink);
  
  var key, j, id, spreadsheet, sheetunitable, dtstart;
  
  var day = ["01","02","03","04","05","06","07","08","09","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31"];  
  var txtmonth = asdates.getRange("A1").getValue();
  var month, monthlengh;
  var year = asdates.getRange("B1").getValue();
  
  var forclear = {contentsOnly: true, skipFilteredRows: true}
    
  //Выбрать длину и номер месяца в зависимости от указанных параметров в таблице "Текущие даты":
    switch(txtmonth)
    {
      case "Январь":
      month = ".01.";
      monthlengh = new Date(year,1,0);
      break;
      
      case "Февраль":
      month = ".02.";      
      monthlengh = new Date(year,2,0);
      break;
              
      case "Март":
      month = ".03.";
      monthlengh = new Date(year,3,0);      
      break;
        
      case "Апрель":
      month = ".04.";
      monthlengh = new Date(year,4,0);
      break;
        
      case "Май":
      month = ".05.";
      monthlengh = new Date(year,5,0);
      break;
        
      case "Июнь":
      month = ".06.";
      monthlengh = new Date(year,6,0);
      break;
      
      case "Июль":
      month = ".07.";
      monthlengh = new Date(year,7,0);
      break;
      
      case "Август":
      month = ".08.";
      monthlengh = new Date(year,8,0);
      break;
        
      case "Сентябрь":
      month = ".09.";
      monthlengh = new Date(year,9,0);
      break;
        
      case "Октябрь":
      month = ".10.";
      monthlengh = new Date(year,10,0);
      break;
      
      case "Ноябрь":
      month = ".11.";
      monthlengh = new Date(year,11,0);
      break;
        
      case "Декабрь":
      month = ".12.";
      monthlengh = new Date(year,12,0);
      break;
        
      default:
        throw "stop";
      break;
    }
  
  //Преобразовать monthlengh к строке, выбрать только то, что нужно (slice) и преобразовать monthlengh к числу:
  monthlengh = Number(monthlengh.toString().slice(8,10));


  function dates(sheetunitable, dtstart, monthlengh, day, month, year){
  //Прописать числа в сводной таблице на определенный месяц:
  for (j = 0; j<monthlengh; j++) {           
    sheetunitable.getRange("A"+(dtstart+j)).setValue(day[j]+month+year);
  }
  //Очистить лишние числа если месяц короче 31:
  if (monthlengh == 28) {
    sheetunitable.getRange("A"+(dtstart+monthlengh)).setValue("");
    sheetunitable.getRange("A"+(dtstart+monthlengh+1)).setValue("");
    sheetunitable.getRange("A"+(dtstart+monthlengh+2)).setValue("");
  }
  if (monthlengh == 29) {
    sheetunitable.getRange("A"+(dtstart+monthlengh)).setValue("");
    sheetunitable.getRange("A"+(dtstart+monthlengh+1)).setValue("");      
  }
  if (monthlengh == 30) {
    sheetunitable.getRange("A"+(dtstart+monthlengh)).setValue("");      
  } 
 } 
  
 for (key in sites.cont.idlink){
   id = ascont.getRange(sites.cont.idlink[key]).getValue();
   spreadsheet = SpreadsheetApp.openById(id);
   sheetunitable = spreadsheet.getSheetByName(sites.unitablelink);
   dtstart = ascont.getRange(sites.cont.dtstartlink[key]).getValue();  
   dates(sheetunitable, dtstart, monthlengh, day, month, year);         
         
 }
 for (key in sites.seo.idlink){
   id = asseo.getRange(sites.seo.idlink[key]).getValue();
   spreadsheet = SpreadsheetApp.openById(id);
   sheetunitable = spreadsheet.getSheetByName(sites.unitablelink);
   dtstart = asseo.getRange(sites.seo.dtstartlink[key]).getValue();
   dates(sheetunitable, dtstart, monthlengh, day, month, year);
       
 }  
};

