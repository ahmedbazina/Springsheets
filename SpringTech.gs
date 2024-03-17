//https://www.youtube.com/watch?v=WSU1Mnuh4Pw
//https://chromewebstore.google.com/detail/google-apps-script-github/lfjcgcmkmjjlieihflfhjopckgpelofo

////////////////////////////////////////////////////////////////////////////// Triggers /////////////////////////////////////////////////////////////////////////////////////////////////////////////

// This function is triggered when a cell is edited
function onEdit(e) {
  // Get the active sheet and the edited range
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var sheetName = sheet.getName();

  if (sheetName === 'Repairs - Worksheet') {
    // Call functions to handle various columns in the 'Repairs - Worksheet' sheet
    autoDateColumnA(range);
    handleProductIDColumn(range);
  } 
}


// This function is triggered when the "Scanned Devices" is changed
function onChange(e) {
  var sourceSheet = e.source.getActiveSheet();
  var sheetName = sourceSheet.getName();

  if (sheetName === 'Scanned Devices') { 
    var sourceSpreadsheetId = "1IyeTZu7CS8jfCnKmtXWjJyhRPOv8JE8eaK8bK9BTlFo"; // Replace with the actual source spreadsheet ID
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    var sourceSheetName = "Scanned Devices";
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
    
    if (sourceSheet) {
      syncStockSheet(sourceSheet); // Call the syncStockSheet function to update the stock sheet
    }
  }
}

// This function is triggered when a row is deleted
function onDelete(e) {
  var sheetName = e.source.getSheet().getName();

  if (sheetName === 'Repairs - Worksheet') {
    updatePartUsageCounts();
    orderQTY();
  }
}