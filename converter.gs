function main() {   

  //These IDs can be found at the end of the URL for the Google Drive folder.
  
  var excelFolderId = '0B5hrtNbfIu_mVndyeWNYMHVCR00';
  var sheetsFolderId = '0B5hrtNbfIu_maW8tcXRDX1ZXQms';
  
  convertExcelFiles2Sheets_(excelFolderId,sheetsFolderId);
}


/*
 * Converts excel files in the excel folder to Google Sheets with the same name and places Google Sheets in the sheets folder.
 * @param {String} excelFolderId The ID of the Google Drive folder where the excel files are located.
 * @param {String} sheetsFolderId The ID of the Google Drive folder where the Google Sheet files will be placed.
 */
function convertExcelFiles2Sheets_(excelFolderId,sheetsFolderId) {
  
  // Loop through code for every file in excel folder
  var xlsFiles = DriveApp.getFolderById(excelFolderId).getFiles();
  while (xlsFiles.hasNext()) {
    
    var xlsFile = xlsFiles.next();
    var xlsFileName = xlsFile.getName();
    var googleFile = DriveApp.getFolderById(sheetsFolderId).getFilesByName(xlsFileName);
    var googleFileExists = googleFile.hasNext();
    
    if (googleFileExists){
      
      var googleFileID = googleFile.next().getId();
      var uploadResponse = APIGoogleDriveUpload_(googleFileID, xlsFile, 'put');
           
    } else {
      
      var uploadResponse = APIGoogleDriveUpload_(null, xlsFile, 'post');

      // Update file name and drive folder
      var updateID = JSON.parse(uploadResponse.getContentText()).id;
      var updateResponse = APIGoogleDriveUpdate_(updateID, xlsFileName, sheetsFolderId);
      
    }
  }
  
  
  // Utility API Functions
  
  function APIGoogleDriveUpload_(uploadID, xlsFile, method) {
    
    var convertParm = '';
    if (uploadID == null){
      uploadID = '';
      convertParm = '&convert=true';
    }
        
    var url = 'https://www.googleapis.com/upload/drive/v2/files/' + uploadID + '?uploadType=media' + convertParm;
    var uploadParams = {
      method: method,
      contentType: 'application/vnd.ms-excel',
      contentLength: xlsFile.getBlob().getBytes().length,
      headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
      payload: xlsFile.getBlob().getBytes()
    };
       
    return UrlFetchApp.fetch(url, uploadParams);
  }
  
  function APIGoogleDriveUpdate_(updateID, xlsFileName, sheetsFolderId) {
    
    var url = 'https://www.googleapis.com/drive/v2/files/' + updateID;
    var payloadData = {
      title: xlsFileName, 
      parents: [{id: sheetsFolderId}]
    };
    var updateParams = {
      method:'put',
      headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
      contentType: 'application/json',
      payload: JSON.stringify(payloadData)
    };
    
    return UrlFetchApp.fetch(url, updateParams);
  }
  
}