/** @OnlyCurrentDoc */

/**
 * Create the Matrixify sheet from Magento Sheet
 */

 function ConvertMagentoToMatrixifySheet() {
  const app = SpreadsheetApp;
  const activeSpreadsheet = app.getActiveSpreadsheet();  
  const initSheet = activeSpreadsheet.getSheetByName('Magento');
  // Get Get all values.
  let allValues = initSheet.getRange("A2:J").getValues();
  // Remove empty cells from all values
  const values = allValues.filter(e=>e.join().replace(/,/g, "").length);
  var finalValues = createFinalValues(values);
  const intColorSizeObj = GetColorsAndSizeFromName();
  var finalValues = replaceDuplicateColorsInSerie(finalValues, intColorSizeObj);
  
  // Clear Sheets
  const matrixify = createAndClearSheet(activeSpreadsheet, 'Matrixify');
  const colorsFromNameSheet = createAndClearSheet(activeSpreadsheet, 'ColorsFromName');        
  // Write Sheets
  writeMatrixifySheet(matrixify, finalValues);  
  writeColorsFromNameSheet(colorsFromNameSheet, intColorSizeObj);
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
      'title': createParentContent(value, prevValue, title, 'name'),
      'command': 'MERGE',
      'variantCommand' : 'MERGE',
      'tagsCommand' : 'MERGE',
      'body': createParentContent(value, prevValue, description, 'body'),
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
function createParentContent(currentValue, previousValue, item, type) {  
  var item = item;
  if(type === 'name' && item.includes(' - ')) {
    // Remove everything after last dash
    let currentNameSplit = item.split(' - ');
    currentNameSplit.pop()
    // cast name back to string
    var item = currentNameSplit.join(' ');
  }

  if(type === 'body'){
    // Remove html
    // if something between < and >
    const htmlRegex = new RegExp("\<.*?\>");    
    if(htmlRegex.test(item)){
      var item = item.split(htmlRegex).join('');            
    }    
  }

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
  //var str = "White / White Edge";
  // Replace all special chars (Non letters).
  var color = color.replace(/[^a-zA-Z]/g, "-");
  // Split the string
  var colorSplit = color.split('-');
  // Remove all duplicates (When there's mulitple - in a row)
  var colorSplit = colorSplit.filter(n => n);
  // Cast back to string
  var color = colorSplit.join('-');
  
  return '#color_' + color.toLowerCase();  
}

function createAdditionalImages(addImages){
  let additionalImages = [];
  if(addImages){
    let imagesSplit = addImages.split(',');        
    let imagesUrls = imagesSplit.map(i => 'https://lux-case.com/media/catalog/product' + i);
    additionalImages.push(imagesUrls);   
  }
  return additionalImages;
}

/**
 * If there is duplicates value of colors, 
 * replace it with colors from the INT name (MagentoNameColors)
 */
function replaceDuplicateColorsInSerie(finalValues, intColorSizeObj){    
  finalValues.map((el, i) => {
    finalValues.find((element, index) => {
      // If handle and color is the same (Duplicate values)
      if (i !== index && element.handle === el.handle && element.color === el.color) {     
        // Go into colors object (From INT sheet), find matching sku and replace the color in finalValues          
        intColorSizeObj.find(colorSizeObj => {
          if (colorSizeObj['sku'] === el.variantSku) {
            // If there is a color from the name
            if(colorSizeObj['color']) {
              // Replace color attribute
              finalValues[i].option1Value = colorSizeObj['color'];              
              // Replace alt tag from the new color
              finalValues[i].imageAltText = colorSizeObj['imageAltText'];
            }            
          };
        });
      }
    });
  });
  return finalValues;
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
    sheet.getRange('F' + nextRow).setValue(finalValue.tagsCommand);
    sheet.getRange('G' + nextRow).setValue(finalValue.body);
    sheet.getRange('H' + nextRow).setValue(finalValue.vendor);
    sheet.getRange('I' + nextRow).setValue(finalValue.tags);
    sheet.getRange('J' + nextRow).setValue(finalValue.option1Name);
    sheet.getRange('K' + nextRow).setValue(finalValue.option1Value);
    sheet.getRange('L' + nextRow).setValue(finalValue.option2Name);
    sheet.getRange('M' + nextRow).setValue(finalValue.option2Value);
    sheet.getRange('N' + nextRow).setValue(finalValue.variantGrams);
    sheet.getRange('O' + nextRow).setValue(finalValue.variantInventoryTracker);
    sheet.getRange('P' + nextRow).setValue(finalValue.variantInventoryQty);
    sheet.getRange('Q' + nextRow).setValue(finalValue.variantInventoryPolicy);
    sheet.getRange('R' + nextRow).setValue(finalValue.variantFulfillmentService);
    sheet.getRange('S' + nextRow).setValue(finalValue.variantPrice);
    sheet.getRange('T' + nextRow).setValue(finalValue.variantRequiresShipping);
    sheet.getRange('U' + nextRow).setValue(finalValue.variantTaxable);
    sheet.getRange('V' + nextRow).setValue(finalValue.variantWeightUnit);
    sheet.getRange('W' + nextRow).setValue(finalValue.seoTitle);
    sheet.getRange('X' + nextRow).setValue(finalValue.seoDescription);    
    sheet.getRange('Y' + nextRow).setValue(finalValue.status);
    if(finalValue.imageSrc[0]){
      sheet.getRange('Z' + nextRow).setValue(finalValue.imageSrc[0].join(';'));
    }
    sheet.getRange('AA' + nextRow).setValue(finalValue.imageAltText);
    sheet.getRange('AB' + nextRow).setValue(finalValue.variantCountryofOrigin);
    sheet.getRange('AC' + nextRow).setValue(finalValue.standardizedProductType);

  });
}

/**
 * Create ColorsFromName sheet from MagentoName (After last dash)
 */
function GetColorsAndSizeFromName() {
  const app = SpreadsheetApp;
  const activeSpreadsheet = app.getActiveSpreadsheet();  
  const initSheet = activeSpreadsheet.getSheetByName('MagentoNameColors');
  // Get values
  let allValues = initSheet.getRange("A2:B").getValues();
  const values = allValues.filter(e=>e.join().replace(/,/g, "").length);  
  const colorSizeObj = getColorsAndSizeFromName(values);
  return colorSizeObj;
}

function getColorsAndSizeFromName(values){
  let colorsObjList = [];
  values.forEach(function(value) {    
    const name = value[1];
    var size = 'One-size';
    // Get everything after last dash    
    let nameSplit = name.split(' - ');        
    var afterLastDash = nameSplit.pop();

    if(afterLastDash.includes(' Size: ')) {
      let afterLastDashSplit = afterLastDash.split(':');
      var size = afterLastDashSplit.pop();          
      // remove spaces
      var size = size.split(' ').join('');
      var afterLastDash = afterLastDash.replace(' Size: ' + size, '');
      // Format the color name
      if(afterLastDash.endsWith(' /')) {
        var afterLastDash = afterLastDash.split(' /');
        afterLastDash.pop();
        // cast back to string
        var afterLastDash = afterLastDash.join(' ');
      }
    }    
    if(afterLastDash === name){
      var afterLastDash = '';
    }   
    
    let colorsObj = {
      'sku': value[0],
      'name': name,
      'color': afterLastDash,
      'imageAltText' : createImageAltText(afterLastDash),
      'size': size,
    };
    colorsObjList.push(colorsObj);
  });
  return colorsObjList;
  #
}

function writeColorsFromNameSheet(sheet, colorSizes){
  colorSizes.forEach(function(value){
    let lastRow = sheet.getLastRow();
    let nextRow = lastRow + 1;        
    sheet.getRange('A' + nextRow).setValue(value.sku);
    sheet.getRange('B' + nextRow).setValue(value.name);
    sheet.getRange('C' + nextRow).setValue(value.color);
    sheet.getRange('D' + nextRow).setValue(value.imageAltText);
    sheet.getRange('E' + nextRow).setValue(value.size);
  });
}