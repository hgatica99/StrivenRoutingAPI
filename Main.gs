
function myFunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get's all exisiting sheet names.
  var sheetNamesFinal = ALLSHEETSNAMES();
  console.log(sheetNamesFinal);

  var URL_STRING = "STRIVEN API KEY";
  var RouteOptURL = "https://routexl.com/";
  var RouteString = 'These are your addressES for today:'
  
  var res = UrlFetchApp.fetch(URL_STRING);
  var json = res.getContentText();
  var data = JSON.parse(json);
  
  const AddressList = [data.data];
  const ListOfTechs = [];
  const TechsNJobs = {};

  // This will check for all Technicians, and then group all the address they will be visiting
  for(var x = 0, addylength = AddressList[0].length; x < addylength; x++){
    var tech = AddressList[0][x].Technician;
    var jobAddy = AddressList[0][x].ShipToAddress;
    var TechIndex = 0;

  // Checks if the Tech name already exists, if not, add it on.
    if(ListOfTechs.includes(tech)){
      TechIndex = ListOfTechs.indexOf(tech);
    } else {
      ListOfTechs.push(tech);
      TechIndex = ListOfTechs.indexOf(tech);
      TechsNJobs[ListOfTechs[TechIndex]] = [];
    }
    TechsNJobs[ListOfTechs[TechIndex]].push(jobAddy);
  }

  // Checks to see if the Tech already has a sheet available to him, if not, create one. Since this uses the ListsOfTechs, the indexes should match
  for(var x = 0, listLength = ListOfTechs.length; x < listLength; x++){
    if (sheetNamesFinal.includes(ListOfTechs[x])){
      console.log("This Sheet Exsists");
    } else {
      ss.insertSheet(ListOfTechs[x], x+1);
    }
  }

  sheetNamesFinal = ALLSHEETSNAMES();
  sheetNamesFinal.shift();

  keys = Object.keys(TechsNJobs);;

  //Going through each Tech/Sheet and populating the rows with the proper addresses
  for(var z = 0, techListLength = keys.length; z < techListLength; z++){
    var addy = "";
    var addyListLength = TechsNJobs[keys[z]][0].length;
    var TechSheet = ss.getSheetByName((sheetNamesFinal[z]));

    TechSheet.getRange(1,1).setValue(RouteOptURL);
    TechSheet.getRange(3,1).setValue(RouteString);

    const techAddys = TechsNJobs[ListOfTechs[z]];
    console.log(techAddys);

    for (var a = 0, listLength = addyListLength; a < listLength; a++){
      addy = TechsNJobs[keys[z]][a];
      TechSheet.getRange(a+5,1).setValue(addy);
      if( a == listLength-1){
        TechSheet.autoResizeColumn(1);
      }
    }

  }
}

//Returns all current sheet names to avoid doubling
function ALLSHEETSNAMES(){
  let ssa = SpreadsheetApp.getActive();
  let sheets = ssa.getSheets();
  let sheetNames = [];
  sheets.forEach(function (sheet){
    sheetNames.push(sheet.getName());
  });
  return sheetNames;
}
