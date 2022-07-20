function CreateKeywords(){
  const app = SpreadsheetApp;
  const activeSpreadsheet = app.getActiveSpreadsheet();  
  const commonKeyWordSheet = activeSpreadsheet.getSheetByName('Common Keyword Research');
  
  // Ask user
  var ui = SpreadsheetApp.getUi();
  var userInput = ui.prompt("Write language abbreviation (ENG, SE, DK, NO, FI, DE, NL, FR, IT, PL)");
  var userInput = userInput.getResponseText();

  // Get all keywords
  const allDevices = commonKeyWordSheet.getRange("A2:A").getValues();
  // Remove empty cells from all values
  const devices = allDevices.filter(e=>e.join().replace(/,/g, "").length);    
  // Get keywords
  const keywordColumn = getKeywordColumn(userInput, commonKeyWordSheet);
  // Create keywords
  const keywords = createKeywords(devices, keywordColumn);
  // Clear and create sheet
  const keywordResearchSheet = createAndClearSheet(activeSpreadsheet, 'Keyword Research');    
  // write sheet
  writeKeywordResearchSheet(keywords, keywordResearchSheet);
}

// Delete existing values from sheet (Except headers)
function createAndClearSheet(activeSpreadsheet, sheetName) {
  const importSheet = activeSpreadsheet.getSheetByName(sheetName);
  if(!importSheet) {
      importSheet = activeSpreadsheet.insertSheet(sheetName);
  } else {
      importSheet.getRange("A2:AD").clear();
  }
  return importSheet;
}

function getKeywordColumn(userInput, commonKeyWordSheet) {
// Get user input  
const lang = userInput.toLowerCase();
// Determing which column (commonKeyWordSheet) to use, according to the user
switch(lang){
  case 'eng':
    var range = 'B2:B';
    break;
  
  case 'se':
    var range = 'C2:C';
    break;
  
  case 'dk':
    var range = 'D2:D';
    break;
  
  case 'no':
    var range = 'E2:E';
    break;
  
  case 'fi':
    var range = 'F2:F';
    break;
  
  case 'de':
    var range = 'G2:G';
    break;   

  case 'nl':
    var range = 'H2:H';
    break; 

  case 'fr':
    var range = 'I2:I';
    break; 

  case 'it':
    var range = 'J2:J';
    break;

  case 'pl':
    var range = 'K2:K';
    break;        
  default:
    var range = 'B2:B';      
}
// Get column and remove empty
let allKeywords = commonKeyWordSheet.getRange(range).getValues();   
const keywordColumn = allKeywords.filter(e=>e.join().replace(/,/g, "").length);  
return keywordColumn;
}

function createKeywords(devices, keywordColumn) {
let keywords = [];
devices.forEach(function(device){
  keywordColumn.forEach(function(kw) {
    var keyword = kw[0];

    if(keyword.includes("[DEVICE]")) {
      var keyword = keyword.replace("[DEVICE]", device)        
      keywords.push(keyword);
    }

  });
});
//
return keywords;  
}

function writeKeywordResearchSheet(keywords, keywordResearchSheet) {
keywords.forEach(function(keyword) {
  let lastRow = keywordResearchSheet.getLastRow();
  let nextRow = lastRow + 1;    
  keywordResearchSheet.getRange('A' + nextRow).setValue(keyword);
});
}
