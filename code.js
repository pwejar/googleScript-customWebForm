function doGet(e){
 var rdv= HtmlService.createTemplateFromFile("home");
     var ss = SpreadsheetApp.openById("1gW<----------->n1-w");
  var sheet=ss.getSheetByName("webOptions");
  var theUser =getUser();
  rdv.user =theUser;
 rdv.bay_list=sheet.getRange(3,2,sheet.getRange("B2").getValue()).getValues();
 rdv.sorting_batch_list=sheet.getRange(3,6,sheet.getRange("F2").getValue()).getValues();
 rdv.cutting_type_list=sheet.getRange(3,4,sheet.getRange("D2").getValue()).getValues();
 rdv.tunnel_list=sheet.getRange(3,3,sheet.getRange("C2").getValue()).getValues();
 rdv.GC_list=sheet.getRange(3,5,sheet.getRange("E2").getValue()).getValues();

 rdv.supervisor_list =sheet.getRange(3,7,sheet.getRange("G2").getValue()).getValues();
  Logger.log("tt");
   return rdv.evaluate();
  
}
//************************************************************************************************* defining include in html tag eg. <?!=include("home-js");?>

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
//*************************************************************************************************paste clipping

function pasteClipping(paste,supervisor,jsDate){
  var timeStamp=new Date();
  var newArr=[];//to add extra data in the array
  var user=getUser();
  Logger.log(user);
 //_------------------------------------------------generating newArr 
  for (var i = 0; i < paste.length; i++) {
            if (paste[i][0]!="") {
              paste[i].unshift(supervisor);
              paste[i].unshift(jsDate);
              paste[i].unshift(timeStamp);
              paste[i].unshift(user);
              var total=(+paste[i][5]*20)+(+paste[i][6]);
              paste[i].push(total);
              var id=""+jsDate+paste[i][4]+total;
              paste[i].push(id);
              newArr.push(paste[i]);   
        }
    }
   
  ///-----------------------------------------------------------------------------------pasting newArr
  var ss = SpreadsheetApp.openById("1gW<----------->n1-w");
  var sheet=ss.getSheetByName("pottingData");
  //------------------getting last row using data region
  var dataR=sheet.getRange("A1:I").getValues();
  var lrIndex;
  for(var i=dataR.length-1;i>=0;i--){
  lrIndex=i
  var row=dataR[i];
    var isBlank=row.every(function(c){return c=="";});
    if(!isBlank){break;}
   }
  
  var lr=lrIndex+1;
  //----------------------------------------------------final pasting
  
  var keys=sheet.getRange("I:I").getValues();
  
  Logger.log(keys);
 
  var theIDs = [];


for(var i = 0; i < keys.length; i++)
{
    theIDs = theIDs.concat(keys[i]);
}

   var check=theIDs.indexOf(newArr[0][8])
  
  if(check<0){
  sheet.getRange(lr+1,1,newArr.length,newArr[0].length).setValues(newArr);
    Logger.log(newArr);
    return "success";
  }else{
    return "there was duplicates";
  }
  
}
//*************************************************************************************************getUser();
function getUser(){
  var userEmail=Session.getActiveUser().getEmail()
return userEmail;
}
//*************************************************************************************************paste placing

function pastePlacing(placing_data,supervisor,jsDate,user,batch,placers){
  var timeStamp=new Date();
  var newArr=[];//to add extra data in the array
  user =getUser();
 //_------------------------------------------------generating newArr 
  for (var i = 0; i < placing_data.length; i++) {
            if (placing_data[i][0]!="") {
              placing_data[i].unshift(batch);
              placing_data[i].unshift(supervisor);
              placing_data[i].unshift(jsDate);
              placing_data[i].unshift(placers);
              var total=3200-((+placing_data[i][8])*200+(+placing_data[i][9]*10)+(+placing_data[i][10]));
              var id=""+jsDate+placing_data[i][3]+placing_data[i][4]+placing_data[i][5]+total;
              placing_data[i].unshift(timeStamp);
              placing_data[i].unshift(user);
              placing_data[i].push(total);
              placing_data[i].push(id);
              newArr.push(placing_data[i]);   
        }
    }
   
  ///-----------------------------------------------------------------------------------pasting newArr
  var ss = SpreadsheetApp.openById("1gW<----------->n1-w");
  var sheet=ss.getSheetByName("placingData");
  //------------------getting last row using data region
  var dataR=sheet.getRange("A1:O").getValues();
  var lrIndex;
  for(var i=dataR.length-1;i>=0;i--){
  lrIndex=i
  var row=dataR[i];
    var isBlank=row.every(function(c){return c=="";});
    if(!isBlank){break;}
   }
  
  var lr=lrIndex+1;
  //----------------------------------------------------final pasting
  
  var keys=sheet.getRange(2,15,lr+1,1).getValues();
  
  
 
  var theIDs = [];


for(var i = 0; i < keys.length; i++)
{
    theIDs = theIDs.concat(keys[i]);
}

   var check=theIDs.indexOf(newArr[0][14])
  
  if(check<0){
  sheet.getRange(lr+1,1,newArr.length,newArr[0].length).setValues(newArr);
    Logger.log(keys);
    return "success";
  }else{
    return "there was duplicates";
    //indexSupT();
  }
  
}

//*************************************************************************************************paste sorting/cleaning/stock

function pasteSorting(sheetname,paste,supervisor,jsDate,user,batch,placers){
  var timeStamp=new Date();
  var newArr=[];//to add extra data in the array
  user =getUser();
 //_------------------------------------------------generating newArr 
  for (var i = 0; i < paste.length; i++) {
            if (paste[i][0]!="") {
              
              var total=((+paste[i][4])*200+(+paste[i][5]*10)+(+paste[i][6]));
              var id=""+jsDate+paste[i][0]+paste[i][1]+paste[i][1]+total;
              
              if(batch!=null){
                paste[i].unshift(batch);}
              paste[i].unshift(supervisor);
              paste[i].unshift(jsDate);
              paste[i].unshift(placers);
              
              paste[i].unshift(timeStamp);
              paste[i].unshift(user);
              paste[i].push(total);
              paste[i].push(id);
              newArr.push(paste[i]);   
        }
    }
   
  ///-----------------------------------------------------------------------------------pasting newArr
  var ss = SpreadsheetApp.openById("1gW<----------->n1-w");
  var sheet=ss.getSheetByName(sheetname);
  //------------------getting last row using data region
  var dataR=sheet.getRange("A1:O").getValues();
  var lrIndex;
  for(var i=dataR.length-1;i>=0;i--){
  lrIndex=i
  var row=dataR[i];
    var isBlank=row.every(function(c){return c=="";});
    if(!isBlank){break;}
   }
  
  var lr=lrIndex+1;
  //----------------------------------------------------final pasting
  
  var keys=sheet.getRange(2,15,lr+1,1).getValues();
  
  
 
  var theIDs = [];


for(var i = 0; i < keys.length; i++)
{
    theIDs = theIDs.concat(keys[i]);
}

   var check=theIDs.indexOf(newArr[0][14])
  
  if(check<0){
  sheet.getRange(lr+1,1,newArr.length,newArr[0].length).setValues(newArr);
    Logger.log(keys);
    return "success";
  }else{
    return "there was duplicates";
  }
  
}

//*************************************************************************************************indexSup();
function indexSup(value){
var ss = SpreadsheetApp.openById("1gW<----------->n1-w");
  var sheet=ss.getSheetByName("webOptions");
  var lastrow=sheet.getRange("G2").getValue();
  var data=sheet.getRange(4,7, lastrow, 2).getDisplayValues();

 
  for (var i=0; i<data.length; i++){
  
    if(data[i][0]==value){
      return data[i][1];
       break;
    }
  }}