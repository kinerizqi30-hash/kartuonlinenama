var SHEET_NAME = "Sheet1"; // Sesuaikan dengan nama sheet Anda

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  var id = e.parameter.id;
  if (!id) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "ID is required"}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0]; 
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = rows[i][j];
      }
      return ContentService.createTextOutput(JSON.stringify(obj))
                           .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({status: "not found"}))
                       .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    var dataParam = e.parameter.data;
    if (!dataParam) {
      return ContentService.createTextOutput(JSON.stringify({status: "error", message: "No data provided"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    var data = JSON.parse(dataParam);
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0];
    
    var rowIndex = -1;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.id.toString()) {
        rowIndex = i + 1;
        break;
      }
    }
    
    var rowData = [
      data.id,
      data.name || "",
      data.title || "",
      data.company || "",
      data.phone || "",
      data.email || "",
      data.bio || "",
      data.avatar || ""
    ];
    
    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch (err) {
    // PERBAIKAN DI SINI (menambahkan kurung tutup yang kurang)
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
