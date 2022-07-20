// Get random number of elements from array
function getMultipleRandom(arr, num) {
  const shuffled = [...arr].sort(() => 0.5 - Math.random());
  return shuffled.slice(0, num);
}

// Remove duplicates from two-dimensional array
function multiDimensionalUnique(arr) {
  var uniques = [];
  var itemsFound = {};
  for(var i = 0, l = arr.length; i < l; i++) {
    var stringified = JSON.stringify(arr[i]);
    if(itemsFound[stringified]) { continue; }
    uniques.push(arr[i]);
    itemsFound[stringified] = true;
  }
  return uniques;
}

// Create paths for ad group / Recursive
function createRecursePath(name) {  
  var path = slugify(name);  
  // If path is above 15
  if (path.length > 15) {     
    // split the slug, remove first element, join, and try again
    var path = path.split('-').splice(1).join('-');              
    // call yourself again
    return createRecursePath(path);         
  } else {
    // When the path is below 15 it's all good, stop calling yourself and return the path
    return path;
  }
}

function slugify(text) {
  const from = "ãàáäâẽèéëêìíïîõòóöôùúüûñç·/_,:;"
  const to = "aaaaaeeeeeiiiiooooouuuunc------"
  const newText = text.split('').map(
    (letter, i) => letter.replace(new RegExp(from.charAt(i), 'g'), to.charAt(i)))
  return newText
    .toString()                     // Cast to string
    .toLowerCase()                  // Convert the string to lowercase letters
    .trim()                         // Remove whitespace from both sides of a string
    .replace(/\s+/g, '-')           // Replace spaces with -
    .replace(/&/g, '-y-')           // Replace & with 'and'
    .replace(/[^\w\-]+/g, '')       // Remove all non-word chars
    .replace(/\-\-+/g, '-');        // Replace multiple - with single -
}

function getParameterByName(name, url) {
    name = name.replace(/[\[\]]/g, '\\$&');
    var regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)'),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, ' '));
}

// Generate ad text, where variables is removed (Texts generated: Headlines, Description)
function generateAdDeviceContent(contentsList, model, maxLength, maxItems) {  
  const name = model[1];
  const type = model[2];    
  // Chosing which keyword list to use
  switch(type.toLowerCase()) {
    case 'smartphone':
      var contents = contentsList[0];
    break;

    case 'smartwatch':
      var contents = contentsList[1];
    break;

    default:
      var contents = contentsList[0];
    break;
  }

  // Generated Content
  var gContents = [];

  contents.forEach(function(content) {
    // With brand in name / {DEVICE}
    let wBrand = content.replaceAll('{DEVICE}', name);  
    // If there's keyword in text, we need a different wordcount
    if(wBrand.includes('{KeyWord:') === true){
      if(wBrand.length < maxLength + 10) {
        gContents.push([wBrand]);
      }
    }    
    if(wBrand.length < maxLength && wBrand.includes('{KeyWord:') === false) {
      gContents.push([wBrand]);
    }

    // Without Brand in name / {DEVICE}, if the name contains more than two words (Eg. iPhone 11 would just be 11)
    if(name.split(' ').length > 2){
      let woBrand = content.replaceAll('{DEVICE}', name.split(" ").slice(1).join(' '));
      if(woBrand.includes('{KeyWord:') === true){
        if(woBrand.length < maxLength + 10) {
          gContents.push([woBrand]);
        }
      }
      if(woBrand.length < maxLength && woBrand.includes('{KeyWord:') === false) {
        gContents.push([woBrand]);      
      }
    }
  });

  // If there's more than maxItems generated headlines, pick maxItems random
  if(gContents.length > maxItems) {
    var gContents = getMultipleRandom(gContents, maxItems);
  }  

  // Clean up the array, so we'll only return unique values
  var gContents = multiDimensionalUnique(gContents);  
  return gContents;
}
