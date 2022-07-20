function writeAdGroupSheet(adGroupSheet, adGroups) {
  adGroups.forEach(function(adGroup) {
    let lastRow = adGroupSheet.getLastRow();
    let nextRow = lastRow + 1;    
    adGroupSheet.getRange('A' + nextRow).setValue(adGroup.campaign);
    adGroupSheet.getRange('B' + nextRow).setValue(adGroup.name);
    adGroupSheet.getRange('C' + nextRow).setValue(adGroup.maxCPC);
    adGroupSheet.getRange('D' + nextRow).setValue(adGroup.maxCPM);
    adGroupSheet.getRange('E' + nextRow).setValue(adGroup.targetCPM);    
    adGroupSheet.getRange('F' + nextRow).setValue('Google Sheets Macro');    
  });
}

function writeKeywordSheet(keywordSheet, adGroups) {
  adGroups.forEach(function(adGroup) {    
    adGroup.keywords.forEach(function(keyword) {
      let lastRow = keywordSheet.getLastRow();
      let nextRow = lastRow + 1;    
      keywordSheet.getRange('A' + nextRow).setValue(adGroup.campaign);
      keywordSheet.getRange('B' + nextRow).setValue(adGroup.name);
      keywordSheet.getRange('C' + nextRow).setValue(keyword);
      keywordSheet.getRange('D' + nextRow).setValue(adGroup.finalUrl);
    });    
  });
}

function writeAdSheet(adsSheet, textAds) {  
  textAds.forEach(function(textAd) {    
    let lastRow = adsSheet.getLastRow();
    let nextRow = lastRow + 1;    
    adsSheet.getRange('A' + nextRow).setValue(textAd.campaign);
    adsSheet.getRange('B' + nextRow).setValue(textAd.adGroup);
    // // Write Headlines, Starting from column 4 (D)    
    adsSheet.getRange(nextRow, 3, 1, textAd.headlines.length).setValues([textAd.headlines]);
    // // Write Descriptions
    adsSheet.getRange(nextRow, 18, 1, textAd.descriptions.length).setValues([textAd.descriptions]);
    // // Write Path 1
    adsSheet.getRange('V' + nextRow).setValue(textAd.path1);      
    // // Write Final Url
    adsSheet.getRange(nextRow, 24, 1).setValue(textAd.finalUrl);
  }); 
}

function writeSitelinkExtSheet(sitelinkExtSheet, adGroups) {
  adGroups.forEach(function(adGroup) {
    adGroup['sitelinkExtension'].forEach(function(ext) {
      // Split sitelink from comma (text and url)
      const extSplit = ext.split(';');
      const linkText = extSplit[0];
      const finalUrl = extSplit[1];
      
      let lastRow = sitelinkExtSheet.getLastRow();
      let nextRow = lastRow + 1;    
      sitelinkExtSheet.getRange('A' + nextRow).setValue(adGroup.campaign);
      sitelinkExtSheet.getRange('B' + nextRow).setValue(adGroup.name);
      sitelinkExtSheet.getRange('C' + nextRow).setValue(linkText);
      sitelinkExtSheet.getRange('D' + nextRow).setValue(finalUrl);      
    });
  });
}

function writeCalloutExtSheet(calloutExtSheet, adGroups) {
  adGroups.forEach(function(adGroup) {
    adGroup['calloutExtension'].forEach(function(ext) {
      let lastRow = calloutExtSheet.getLastRow();
      let nextRow = lastRow + 1;
      calloutExtSheet.getRange('A' + nextRow).setValue(adGroup.campaign);
      calloutExtSheet.getRange('B' + nextRow).setValue(adGroup.name);
      calloutExtSheet.getRange('C' + nextRow).setValue(ext);
    });    
  });
}




