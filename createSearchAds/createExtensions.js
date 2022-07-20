function createSitelinkExtension(model) {  
    // Gather all extensions
    let extensions = [model[7], model[8], model[9], model[10]];
    // Remove empty values
    extensions = extensions.filter(n => n);
    return extensions;
  }
  
  function createCalloutExtension(model, calloutExtensionTxt){
    // We are getting to random extensions
    const randomExtensions = getMultipleRandom(calloutExtensionTxt, 10);
    return randomExtensions; 
  }