/** @OnlyCurrentDoc */
function ConvertMagentoToMatrixifySheet() {
    const app = SpreadsheetApp;
    const activeSpreadsheet = app.getActiveSpreadsheet();  
    const initSheet = activeSpreadsheet.getSheetByName('Magento');
    
    // Get Get all values.
    let allValues = initSheet.getRange("A2:J").getValues();      
    // Remove empty cells from all values
    const values = allValues.filter(e=>e.join().replace(/,/g, "").length);
    const finalValues = createFinalValues(values);
    // Clear Sheets
    const matrixify = createAndClearSheet(activeSpreadsheet, 'Matrixify');
    // Write sheets
    writeMatrixifySheet(matrixify, finalValues);
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
  
  /**
   * A object that contains all values, we'll write to the sheet later on
   */
  function createFinalValues(values){
    const finalValues = [];
  
    for (const [index, value] of values.entries()) {
      const prevValue = values[index-1];
      const sku = value[0];
      const ean = value[1];
      const title = value[2];
      const manufacturer = value[3];
      const model = value[4];
      const additional_images = value[5];
      const description = value[6];
      const color = value[7];
      const qty = value[8];
      const price = value[9]; 
  
      let finalValue = {
        'handle' : createHandle(sku),
        'variantSku' : sku,
        'title': createParentContent(value, prevValue, title),
        'command': 'MERGE',
        'variantCommand' : 'MERGE',
        'body': createParentContent(value, prevValue, description),
        'vendor' : 'urrem.dk',	
        'tags' : manufacturer + ',' + model,	
        'option1Name' : 'Color',	
        'option1Value' : color,	
        'option2Name' : 'Size',	
        'option2Value' : '',	
        'variantGrams' : '50',
        'variantInventoryTracker' : 'shopify',	
        'variantInventoryQty' : qty,	
        'variantInventoryPolicy' : 'deny',	
        'variantFulfillmentService': 'manual',
        'variantPrice' : price,		
        'variantRequiresShipping' : 'true',	
        'variantTaxable' : 'true',	
        'variantWeightUnit' : 'kg',
        'seoTitle' : '',	
        'seoDescription' : '',	      
        'status' : 'active',	
        'imageSrc' : createAdditionalImages(additional_images),
        'imageAltText' : createImageAltText(color),	
        'variantCountryofOrigin' : 'SE',	
        'standardizedProductType' : '342 - Apparel & Accessories > Jewelry > Watch Accessories > Watch Bands',      
      };
  
      finalValues.push(finalValue);
  
    }
    return finalValues;
  }
  
  function createHandle(sku){
      let skuSplit = sku.split('-');    
      // remove last item of skuSplit
      skuSplit.pop();
      let handle = skuSplit.join('-');    
      return handle;
  }
  
  /**
   * A function that checks if current SKU belongs to the previous SKU
   * If product belongs to the same serie, only the parent (The first product) should have content (Title or desc)
   */
  function createParentContent(currentValue, previousValue, item) {  
    var item = item;
    // If first item, there's no previousValue
    if(previousValue){
      const currentSku = currentValue[0];
      const previousSku = previousValue[0];  
      // Check if currentSku and previousSku belongs to same serie
      let currentSkuSplit = currentSku.split('-');
      currentSkuSplit.pop();
      let previousSkuSplit = previousSku.split('-');
      previousSkuSplit.pop();
      if (currentSkuSplit.join('-') === previousSkuSplit.join('-')){      
        var item = '';
      } 
    }
    return item;
  }
  
  /**
   * If colors contain more than one word, use the first
   */
  function createImageAltText(color){  
    if(color.includes('/')){
      const colorSplit = color.split('/');
      var color = colorSplit[0];     
    }
    return '#color_' + color.toLowerCase();  
  }
  
  function createAdditionalImages(addImages){
    let additionalImages = [];
    if(addImages){
      let imagesSplit = addImages.split(';');        
      let imagesUrls = imagesSplit.map(i => 'https://lux-case.com/media/catalog/product/' + i);
      additionalImages.push(imagesUrls);   
    }
    return additionalImages;
  }
  
  function writeMatrixifySheet(sheet, finalValues){
    finalValues.forEach(function(finalValue) {
      let lastRow = sheet.getLastRow();
      let nextRow = lastRow + 1;    
      sheet.getRange('A' + nextRow).setValue(finalValue.handle);
      sheet.getRange('B' + nextRow).setValue(finalValue.variantSku);
      sheet.getRange('C' + nextRow).setValue(finalValue.title);
      sheet.getRange('D' + nextRow).setValue(finalValue.command);
      sheet.getRange('E' + nextRow).setValue(finalValue.variantCommand);
      sheet.getRange('F' + nextRow).setValue(finalValue.body);
      sheet.getRange('G' + nextRow).setValue(finalValue.vendor);
      sheet.getRange('H' + nextRow).setValue(finalValue.tags);
      sheet.getRange('I' + nextRow).setValue(finalValue.option1Name);
      sheet.getRange('J' + nextRow).setValue(finalValue.option1Value);
      sheet.getRange('K' + nextRow).setValue(finalValue.option2Name);
      sheet.getRange('L' + nextRow).setValue(finalValue.option2Value);
      sheet.getRange('M' + nextRow).setValue(finalValue.variantGrams);
      sheet.getRange('N' + nextRow).setValue(finalValue.variantInventoryTracker);
      sheet.getRange('O' + nextRow).setValue(finalValue.variantInventoryQty);
      sheet.getRange('P' + nextRow).setValue(finalValue.variantInventoryPolicy);
      sheet.getRange('Q' + nextRow).setValue(finalValue.variantFulfillmentService);
      sheet.getRange('R' + nextRow).setValue(finalValue.variantPrice);
      sheet.getRange('S' + nextRow).setValue(finalValue.variantRequiresShipping);
      sheet.getRange('T' + nextRow).setValue(finalValue.variantTaxable);
      sheet.getRange('U' + nextRow).setValue(finalValue.variantWeightUnit);
      sheet.getRange('V' + nextRow).setValue(finalValue.seoTitle);
      sheet.getRange('W' + nextRow).setValue(finalValue.seoDescription);
      sheet.getRange('X' + nextRow).setValue(finalValue.variantWeightUnit);
      sheet.getRange('Y' + nextRow).setValue(finalValue.status);
      if(finalValue.imageSrc[0]){
        sheet.getRange('Z' + nextRow).setValue(finalValue.imageSrc[0].join(','));
      }
      sheet.getRange('AA' + nextRow).setValue(finalValue.imageAltText);
      sheet.getRange('AB' + nextRow).setValue(finalValue.variantCountryofOrigin);
      sheet.getRange('AC' + nextRow).setValue(finalValue.standardizedProductType);
  
    });
  }