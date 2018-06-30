function cpranges(x) {
  var cont = {
    //source:
      sitenumber:{},
      id:{},      
      dtstart:{},      
      firstltr:{},
      lastltr:{},
      unitablelink: "сводная таблица",
      srcrng:{},
    //target:
      idmaintbl: "dfhjkshfjsee8227cscsjcksjckq203j",      
      sheet:{},
      dtstarget:{},
      trgrng:{}    
  }      
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ascont = ss.getSheetByName("Контекст");
  var source, target, key, chk1, chk2, chk3,intrv;
  
  cont.sitenumber = x;
  cont.sheet = ascont.getRange("A"+(cont.sitenumber+1)).getValue();
  cont.id = ascont.getRange("B"+(cont.sitenumber+1)).getValue();  
  cont.dtstart = ascont.getRange("C"+(cont.sitenumber+1)).getValue();
  cont.dtstarget = ascont.getRange("F"+(cont.sitenumber+1)).getValue();
  cont.firstltr = ascont.getRange("D"+(cont.sitenumber+1)).getValue();
  cont.lastltr = ascont.getRange("E"+(cont.sitenumber+1)).getValue();
  
  chk1 = SpreadsheetApp.openById(cont.id).getSheetByName(cont.unitablelink).getRange("A"+(cont.dtstart+28)).getValue();      
  chk2 = SpreadsheetApp.openById(cont.id).getSheetByName(cont.unitablelink).getRange("A"+(cont.dtstart+29)).getValue();  
  chk3 = SpreadsheetApp.openById(cont.id).getSheetByName(cont.unitablelink).getRange("A"+(cont.dtstart+30)).getValue();
  
  if (chk1 == 0 ){      
      intrv = 27;
  }
  else if (chk2 == 0 ){     
     intrv = 28;
  }
  else if (chk3 == 0 ){     
     intrv = 29;
  } else {    
    intrv = 30;
  }  
  
  cont.srcrng = (cont.firstltr+cont.dtstart+":"+cont.lastltr+(cont.dtstart+intrv)).toString();
  cont.trgrng = (cont.firstltr+cont.dtstarget+":"+cont.lastltr+(cont.dtstarget+intrv)).toString();
    
  source = SpreadsheetApp.openById(cont.id).getSheetByName(cont.unitablelink).getRange(cont.srcrng);
  target = SpreadsheetApp.openById(cont.idmaintbl).getSheetByName(cont.sheet).getRange(cont.trgrng);
  target.setValues(source.getValues());  
   
 }

function cpallranges(){  
  var lastrow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Контекст").getLastRow();
  for (i = 1; i < lastrow; i++){
    cpcont(i);   
  }
}
  
