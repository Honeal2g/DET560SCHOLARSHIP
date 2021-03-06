function onOpen(){
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Scholarship Options');
  var item = menu.addItem('Generate Award Letter(s)', 'DataMine');
  var item2 = menu.addItem('Generate Bill-ID Email', 'Bill_ID');
  var item3 = menu.addItem('Generate Nomination Email', 'ScholarshipEmail');
  item.addToUi();
  item2.addToUi();
  item3.addToUi();  
}
function CreateMemo(Date,College,Term,Lname,Fname,SSAN,Type,Max){
  var template = DriveApp.getFileById('1ihx8O0zdSkZIpNogUQWAOw2jlllnq9Rr_EoGHUuAoQg');
  var NewMemo = template.makeCopy();  
  NewMemo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  NewMemo.setName(Lname+Fname+Term);//NEED TO IMPROVE NOMENCLATURE
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
  var Award_ID = NewMemo.getId();
  return Award_ID;
}
function ScholarshipEmail(id){
  var email = "Honeal2g@gmail.com";
  var subject = "Air Force ROTC Scholarship Award Letter";  
  var body = "This is a test";
  
  var Memo = DriveApp.getFileById(id);
  var PDF_Memo = Memo.getAs(MimeType.PDF);
  
  MailApp.sendEmail(email, subject, body, {attachments: [PDF_Memo]}); 
  //Write function that finds all ID's(Array) for the same school
  //Write function that makes PDF's from (ID_Array)
  //Send Email to Individual POC's with Bulk PDF
}

function Bill_ID(){//Can this email be sent from EDU?
   pass
}

function DataMine(){
  var ColRange = [];
  var ColNames = ['Generated Award','Last Name','First Name','Term','SSAN','College','TYPE','Estimate','Date','AwardLTR-IDS'];
  var Cols = FindCols(ColNames);
  var Internal_Data = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  for(var i = 0; i < Cols.length; i++){
    ColRange[i] = Internal_Data.getRange(Cols[i]);
  }  
  var AwardLtr = ColRange[0].getDisplayValues();
  var Lnames = ColRange[1].getDisplayValues(); 
  var Fnames = ColRange[2].getDisplayValues(); 
  var Terms = ColRange[3].getDisplayValues();
  var SSANs = ColRange[4].getDisplayValues(); 
  var Colleges = ColRange[5].getDisplayValues(); 
  var TYPEs = ColRange[6].getDisplayValues();
  var Estimates = ColRange[7].getDisplayValues(); 
  var Dates = ColRange[8].getDisplayValues();  
  
  for(var i = 0; i < AwardLtr.length; i++){
    if(AwardLtr[i]=="NO"){
      var id = CreateMemo(Dates[i],Colleges[i],Terms[i],Lnames[i],Fnames[i],SSANs[i],TYPEs[i],Estimates[i]);
      ColRange[9].getCell(i+1,1).setValue(id);
      ColRange[0].getCell(i+1,1).setValue("YES");
      ScholarshipEmail(id) //Here for test purposes ONLY 
      break; //Here for test purposes ONLY       
    }
  }    
}

function FindCols(ColNames){
  var Col_Index = [];
  var TuitionCols = [];
  var CustomColNames = ColNames;
  var Internal_Data = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = Internal_Data.getDataRange().getValues();  
  for(var i = 0; i < CustomColNames.length; i++){
    Col_Index[i] = data[0].indexOf(CustomColNames[i]);   
  }
  for(var i = 0; i < Col_Index.length; i++){
    var Range = Internal_Data.getRange(2,Col_Index[i]+1,data.length-2);
    TuitionCols[i] = Range.getA1Notation();
  }
  return TuitionCols;  
}