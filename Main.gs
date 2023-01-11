const ss = SpreadsheetApp.getActiveSpreadsheet();
const apiKey = 'STRIVEN_REPORT_API'
const routeXUrl = "https://routexl.com/"
const myFunction = () => {
  let allSheets = getAllSheets()
  let repAddress = getUniqueSalesReps()
  setFinalSheets(allSheets, repAddress)
  loadFinalSheets(repAddress, allSheets)

}
const getAllSheets = () => {
  let sheets =  SpreadsheetApp.getActive().getSheets()
  let sheetNames  = [];
  sheets.forEach((sheet) => {
    if (!sheetNames.includes(sheet)) sheetNames.push(sheet)
  })
  return sheetNames
}
const getUniqueSalesReps = () => {
  const estimateApiData = JSON.parse(UrlFetchApp.fetch(apiKey).getContentText()).data
  let uniqueRepDic = {}
  for(let x = 0; x < estimateApiData.length; x++){
    const { ShipToAddress: address, SalesRep: rep } = estimateApiData[x]
    if(uniqueRepDic[rep] === undefined) uniqueRepDic[`${rep}`] = {
      'Addresses' : []
    }
    if(uniqueRepDic[rep] != undefined && !uniqueRepDic[`${rep}`]['Addresses'].includes(address)) uniqueRepDic[`${rep}`]['Addresses'].push(address)
  }
  return uniqueRepDic
}
const getSheetNames = (sheets) => {
  let finalList = []
  sheets.forEach((sheet) => {
    finalList.push(sheet.getName())
  })
  return finalList
}
const getMissingSheetIndexes = (repNames, sheets) => {
  let missingNameIndex = []
  for(let x = 0; x < sheets.length; x++){
    if(!repNames.includes(sheets[x].getName())){
      console.log(missingNameIndex)
      missingNameIndex.push(x)
    }
  }
  return missingNameIndex
}
const getMissingRepNames = (repNames, sheetNames) =>{
  let missingRepNames = []
  for(let x = 0; x < repNames.length; x++){
    if(!sheetNames.includes(repNames[x])) missingRepNames.push(repNames[x])
  }
  return missingRepNames
}
const setFinalSheets = (sheets, reps) => {
  let repNames = Object.keys(reps)
  let sheetKeys = Object.keys(sheets)
  let sheetNames = getSheetNames(sheets)
  let repNamesLength = repNames.length
  let sheetsLength = sheets.length
  let sheetDifference, x
    
  switch(true){
    case (repNamesLength > sheetsLength): 
      sheetDifference = (repNamesLength - sheetsLength)
      x = 0
      while(x < sheetDifference){
        ss.insertSheet()
        x++
      }
      break;
    case (repNamesLength < sheetsLength):
      sheetDifference = (sheetsLength - repNamesLength)
      console.log("Too Many Sheets")
      for(let x = sheetKeys.length-1; x >= repNamesLength; x--){
        ss.deleteSheet(sheets[x])
        sheets.splice(x, 1)
        sheetKeys = Object.keys(sheets)
      }
      break;
  }

  // We need to update this "Sheets" instance because the sheet objects are getting adjust above. Without this, the code below throws errors
  sheets = getAllSheets()
  sheetNames = getSheetNames(sheets)

  let missingRepNames = getMissingRepNames(repNames, sheetNames)
  let missingIndexs = getMissingSheetIndexes(repNames, sheets)
  console.log('',repNames, '\n', sheetNames,'\n', missingRepNames,'\n', missingIndexs)

  for(let x = 0 ; x < missingIndexs.length; x++){
    sheets[missingIndexs[x]].setName(missingRepNames[x])
  }
}
const clearSheets = (finalSheets) => {
  for(let x = 0; x < finalSheets.length; x++){
    finalSheets[x].getRange('A1:O500').clearContent()
  }
}
const loadFinalSheets = (repAddress, sheets) => {
  let addresses;
  clearSheets(sheets)
  sheets = getAllSheets()
  for(let x = 0; x < sheets.length; x++){
    let currentSheet = sheets[x]
    currentSheet.getRange(1,1).setValue(routeXUrl)
    if(repAddress[currentSheet.getName()] != undefined){
      addresses = ((repAddress[currentSheet.getName()]['Addresses']))
      for(let a = 0; a < addresses.length; a++){
        addy = addresses[a]
        currentSheet.getRange(a+3, 1).setValue(addy)
      }
    }
    currentSheet.autoResizeColumn(1);
  }
}
