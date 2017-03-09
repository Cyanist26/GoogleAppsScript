function onOpen() {
  ui = SpreadsheetApp.getUi();
  ui 
  .createMenu('custom menu') 
  .addItem('Visibility Test','visibilityTest')
  .addToUi();
}

function visibilityTest(){
  var html = HtmlService.createHtmlOutputFromFile('visibilityTest')
      .setWidth(500)
      .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Visibility Test');
}

function getInfo(){
  try{
    var info = [];
    
    var userEmail = Session.getEffectiveUser().getEmail();
    var result = getData(userEmail);
    info.push(userEmail);
    info.push(result[0]);
    info.push(result[1]);
    info.push(result[2]);
    info.push(result[3]);
    info.push(result[4]);
    info.push(result[5]);
    info.push(result[6]);
    return info;
  }
  catch(e){
    throw e;
  }
}

function getData(email) {
  try{
  var dataFile = SpreadsheetApp.openById("1LbYVRqbvG9M-pm9mqU6jhtOZ__kF8-Kaag_-a5v4ZYg");
  var dataSheet = dataFile.getSheetByName("工作表1");
  var dataRange = dataSheet.getDataRange();
  var data = dataRange.getValues();
  
  var emailIndex = FtofsStandardLibrary.getIndexByContent(true, email, data);
  for( var i = 0; i < emailIndex.length; i++ )
  {
    if( data[(emailIndex[i][0])-1][4] != "離職" )
    {
      var result = [];
      result.push(data[(emailIndex[i][0])-1][0]);
      result.push(data[(emailIndex[i][0])-1][1]);
      result.push(data[(emailIndex[i][0])-1][2]);
      result.push(data[(emailIndex[i][0])-1][3]);
      result.push(data[(emailIndex[i][0])-1][4]);
      result.push(data[(emailIndex[i][0])-1][5]);
      result.push(data[(emailIndex[i][0])-1][6]);
      return result;
    }
  }
  }
  catch(e)
  {
    throw e;
  }
}