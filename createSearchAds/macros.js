/** @OnlyCurrentDoc */

function CreateSearchAds() {
  const app = SpreadsheetApp;
  const activeSpreadsheet = app.getActiveSpreadsheet();  
  const initSheet = activeSpreadsheet.getSheetByName('Init');
  const txtSheet = activeSpreadsheet.getSheetByName('Txt');

  // Get models column. Get all values.
  let allModels = initSheet.getRange("A2:M").getValues();          
  // Remove all empty from models
  const models = allModels.filter(e=>e.join().replace(/,/g, "").length);          
  
  // Get Device values
  const phoneKeywords2d = txtSheet.getRange("A2:A").getValues().filter(String);  
  const phoneKeywords = [].concat(...phoneKeywords2d);
  const phoneHeadlines2d = txtSheet.getRange("B2:B").getValues().filter(String);
  const phoneHeadlines = [].concat(...phoneHeadlines2d);
  const phoneDescriptions2d = txtSheet.getRange("C2:C").getValues().filter(String);
  const phoneDescriptions = [].concat(...phoneDescriptions2d);
  const watchKeywords2d = txtSheet.getRange("D2:D").getValues().filter(String);  
  const watchKeywords = [].concat(...watchKeywords2d);    
  const watchHeadlines2d = txtSheet.getRange("E2:E").getValues().filter(String);
  const watchHeadlines = [].concat(...watchHeadlines2d);
  const watchDescriptions2d = txtSheet.getRange("F2:F").getValues().filter(String);
  const watchDescriptions = [].concat(...watchDescriptions2d);
  const calloutExtensionTxt = txtSheet.getRange("G2:G").getValues().filter(String);  
  // Gather device values
  const keywords = [phoneKeywords, watchKeywords];
  const headlines = [phoneHeadlines, watchHeadlines];
  const descriptions = [phoneDescriptions, watchDescriptions];

  // Create adGroups objects  
  const adGroups = createAdGroup(models, keywords, calloutExtensionTxt);
  const textAds = createTextAds(models, headlines, descriptions);  

  // Clearing sheets, so we ensure we only have new data
  const adGroupSheet = createAndClearSheet(activeSpreadsheet, 'Ad Group');
  const keywordSheet = createAndClearSheet(activeSpreadsheet, 'Keywords');
  const adsSheet = createAndClearSheet(activeSpreadsheet, 'Ads');
  const sitelinkExtSheet = createAndClearSheet(activeSpreadsheet, 'Sitelink Extensions');
  const calloutExtSheet = createAndClearSheet(activeSpreadsheet, 'Callout Extensions');  
  const priceExtSheet = createAndClearSheet(activeSpreadsheet, 'Price Extensions');  

  // Write ad group sheet
  writeAdGroupSheet(adGroupSheet, adGroups);
  // Write Keywords Sheet
  writeKeywordSheet(keywordSheet, adGroups);
  // Write ad sheet
  writeAdSheet(adsSheet, textAds);
  // Write sitelink extenstion sheet
  writeSitelinkExtSheet(sitelinkExtSheet, adGroups);
  // Write callout extenstion sheet
  writeCalloutExtSheet(calloutExtSheet, adGroups);
  // Write price extension sheet
  writePriceExtSheet(priceExtSheet, adGroups);
}

// Delete existing values
function createAndClearSheet(activeSpreadsheet, sheetName) {
  const importSheet = activeSpreadsheet.getSheetByName(sheetName);
  if(!importSheet) {
    importSheet = activeSpreadsheet.insertSheet(sheetName);
  } else {
    importSheet.getRange("A2:AD").clear();
  }
  return importSheet;
}

function createAdGroup(models, keywords, calloutExtensionTxt) {
  const adGroups = [];
  models.forEach(function(model) {    
    let adGroup = {
      'campaign': [model[0]],
      'name': [model[1]],
      // 'finalUrl': [model[3]],      
      'finalUrl': createFinalUrl(model[3]),      
      'maxCPC' : [model[4]],
      'maxCPM' : [model[5]],
      'targetCPM' : [model[6]],
      'language' : [model[11]],
      'currency' : [model[12]],
      'keywords': createKeywords(model, keywords),
      'sitelinkExtension' : createSitelinkExtension(model),
      'calloutExtension' : createCalloutExtension(model, calloutExtensionTxt),
      'priceExtension' : createPriceExtension(model),
    };
    // Push created adGroup into array
    adGroups.push(adGroup);
  });
  return adGroups;
}

function createFinalUrl(url){
  // Remove any parameters
  if(url.includes('?')) {
    let urlSplit = url.split('?');    
    var url = urlSplit[0]   
  }
  return url;
}

function createKeywords(model, keywordList) {  
  const name = model[1];
  const type = model[2];  
  // Chosing which keyword list to use
  switch(type.toLowerCase()) {
    case 'smartphone':
      var keywords = keywordList[0];
    break;

    case 'smartwatch':
      var keywords = keywordList[1];
    break;

    default:
      var keywords = keywordList[0];
    break;
  }
  // The generated keywords we'll return
  const gKeywords = []
  // Loop through keyword list
  keywords.forEach(function(keyword) {            
    // If keyword is a broad modifier
    if(keyword.includes('+')) {          
      // Generate Broad modifier (Every word in model with a plus; +fitbit +charge +4)
      modifiedName = name.split(" ").map(s => '+' + s).join(' ');
      var gKeyword = keyword.replaceAll('{DEVICE}', modifiedName);      
      // Generate Broad modifier (Every word in model, execpt brand, with a plus; +charge +4)
      // modifiedName = name.split(" ").slice(1).map(s => '+' + s).join(' ');
      // var gKeyword = keyword.replaceAll('+{DEVICE}', modifiedName);      
    } else {
      var gKeyword = keyword.replaceAll('{DEVICE}', name);
    };
    // Push generated keyword to array    
    gKeywords.push([gKeyword]);            
  });          
  return gKeywords;
}

function createPriceExtension(model){  
  const url = model[3];
  const prices = getParameterByName('price', url);
  var lowestPrice = '';
  if(prices != null) {
    const splitPrices = prices.split('-');
    var lowestPrice = splitPrices[0];    
  }   
  return lowestPrice;
}

function writePriceExtSheet(priceExtSheet, adGroups) {
  adGroups.forEach(function(adGroup) {
    let lastRow = priceExtSheet.getLastRow();
    let nextRow = lastRow + 1;    
    priceExtSheet.getRange('A' + nextRow).setValue(adGroup.campaign);
    priceExtSheet.getRange('B' + nextRow).setValue(adGroup.name);
    priceExtSheet.getRange('C' + nextRow).setValue(adGroup.priceExtension);    
    priceExtSheet.getRange('H' + nextRow).setValue(adGroup.language);    
    priceExtSheet.getRange('I' + nextRow).setValue(adGroup.currency);    
  });
}

function createTextAds(models, headlineList, descriptionList) {
  // A list we'll push objects to.  
  let textAds = [];  
  // How many ads we want to create
  models.forEach(function(model) {
    for(i = 0; i <= 2; i++){    
      let ads = {
        'campaign' : model[0],
        'adGroup' : model[1],
        'headlines' : generateAdDeviceContent(headlineList, model, 30, 15),
        'descriptions' : generateAdDeviceContent(descriptionList, model, 90, 4),
        'path1': createRecursePath(model[1]),      
        'finalUrl': model[3],        
      };      
      textAds.push(ads);
    }  
  });
  
  return textAds;
}
