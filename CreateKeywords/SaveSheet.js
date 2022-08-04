// Save active sheet
function saveSheetToDriveFolder(){  
  // Ask user
  let ui = SpreadsheetApp.getUi();
  const userInput = ui.prompt("Please, save file as Type Model Country (EG. Smartphone - iPhone Y - SE) ");
  let newFileName = userInput.getResponseText();
  // The active spreadsheet
  const currentSheet = SpreadsheetApp.getActiveSheet();  
  const sheetName = currentSheet.getName();
  const spreadsheetID = currentSheet.getParent().getId();

  // Get sheetname so we can find the correct folder
  if(sheetName == 'Smartphone Template'){
    var googleDriveFolderId = '1DBwFu0MblZjCI5zP1deDadfABuIArYPr'
  }
  if(sheetName == 'Smartwatch Template'){
    var googleDriveFolderId = '1_WB_JdQ5dlKc2UJ0oJd5OF9SNsEWweKK'
  }
  const googleDriveFolder = DriveApp.getFolderById(googleDriveFolderId);
  // copy the spreadsheet
  // Append date to the newFileName sheet name, so we don't accidentally overwrite others
  newFileName = newFileName + ' - ' + Math.floor(Date.now() / 1000);

  const spreadSheetCopy = SpreadsheetApp.open(DriveApp.getFileById(spreadsheetID).makeCopy(newFileName, googleDriveFolder));
  // Delete the sheet we don't need in the copy (We only want the saved text)
  const sheetsCopy = spreadSheetCopy.getSheets();
  for(i = 0; i < sheetsCopy.length; i++){
    if(sheetsCopy[i].getSheetName() != sheetName){
      spreadSheetCopy.deleteSheet(sheetsCopy[i]);
    }
  }

  MailApp.sendEmail({
    to: "marketing5@wepack.se",
    subject: "New Lux-case text has been written!",
    htmlBody: "Filename: " + newFileName + '<a href="https://drive.google.com/drive/folders/1JqaSgrXvHII6T1_Zzsu9T7ErkTfLJ_Dt"> Link to Google Drive</a>',
  });
}