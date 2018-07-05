function onOpen(){
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Scholarship Options');
  var item = menu.addItem('Generate Award Letter(s)', 'Test');
  var item2 = menu.addItem('Generate Bill-ID Email', 'Bill_ID');
  var item3 = menu.addItem('Generate Nomination Email', 'newScol');
  item.addToUi();
  item2.addToUi();
  item3.addToUi();  
}

//function onEdit(e) {
//  Logger.log(e);
//  var spreadsheet = SpreadsheetApp.getActive();
//  for(var i = 2; i<40; i++){
//    var Type = spreadsheet.getRange('G'+i).getValue();
//    //Logger.log(Type);
//    var Tuition = spreadsheet.getRange('J'+i).getValue();
//    //Logger.log(Type+" : "+Tuition);
//    if(Type == "TYPE 2"){
//      //Logger.log(Type +" is equal to TYPE 2");
//      if(Tuition > 9000){
//        //Logger.Log(Tuition +"is Greater than 9,000");
//        spreadsheet.getRange('G'+i).activate();
//        spreadsheet.getActiveRangeList().setFontColor('red');
//        spreadsheet.getRange('J'+i).activate();
//        spreadsheet.getActiveRangeList().setFontColor('red');
//      }else {
//        spreadsheet.getRange('G'+i).activate();
//        spreadsheet.getActiveRangeList().setFontColor('black');
//        spreadsheet.getRange('J'+i).activate();
//        spreadsheet.getActiveRangeList().setFontColor('black');
//      }
//    }    
//  }
//}
function FindCols(ColNames) {
  var Col_Index = [];
  var TuitionCols = [];
  var CustomColNames = ColNames;
  var Internal_Data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SMR");
  var data = Internal_Data.getDataRange().getValues();  
  for(var i = 0; i < CustomColNames.length; i++){
    Col_Index[i] = data[0].indexOf(CustomColNames[i]);   
  }
  for(var i = 0; i < Col_Index.length; i++){
    var Range = Internal_Data.getRange(2,Col_Index[i]+1,data.length-1);
    TuitionCols[i] = Range.getA1Notation();
  }
  return TuitionCols;  
}

function CreateMemo(Date,College,Term,Lname,Fname,SSAN,Type,Max) {
  var template = DriveApp.getFileById('1ihx8O0zdSkZIpNogUQWAOw2jlllnq9Rr_EoGHUuAoQg');
  var NewMemo = template.makeCopy();
  
  
  NewMemo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  NewMemo.setName(Lname+Fname+Term);
  
  var doc = DocumentApp.openById(NewMemo.getId());
  var body = doc.getBody();
  body.replaceText('\\$date', Date);
  body.replaceText('\\$college', College);
  body.replaceText('\\$term', Term);
  body.replaceText('\\$l_name', Lname);
  body.replaceText('\\$f_name', Fname);
  body.replaceText('\\$ssan', SSAN);
  body.replaceText('\\$type', Type);
  body.replaceText('\\$max', Max);
  doc.saveAndClose();
  DriveApp.createFile(doc.getAs('application/pdf'))
 
}

function Test(){
CreateMemo("1 July 2018","MANHATTAN COLLEGE", "Fall 2018", "HAWTHORNE","ULAN","0-1562","TYPE-2","$9,000");
}
function newScol(){
  pass
}

function Bill_ID(){
  pass
}

function DataMine(){
  var ColNames = ['LAST NAME','FIRST NAME','TERM','SSAN','COLLEGE','TYPE','ESTIMATE','DATE'];
  var Internal_Data = SpreadsheetApp.getActiveSpreadsheet();
}
