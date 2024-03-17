////////////////////////////////////////////////////////////////////////////// Repair - Worksheet ///////////////////////////////////////////////////////////////////////////////////////////////////

// Handle changes in date columnA
function autoDateColumnA(range) {
  if (range.getColumn() === 2) {
    var supplierValue = range.getValue();
    var row = range.getRow();
    var dateCell = range.offset(0, -1);

    if (supplierValue === "") {
      dateCell.setValue("");
    } else if (row > 1) {
      var currentDate = new Date();
      var formattedDate = Utilities.formatDate(currentDate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yy");
      dateCell.setValue(formattedDate);
    }
  }
}

// Handle changes in Product ID column
function handleProductIDColumn(range) {
  if (range.getColumn() === 3) {
    var productID = range.getValue();
    var repairsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Repairs - Worksheet');
    var stockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Parts - Prices');
    var data = stockSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == productID) {
        var currentRow = range.getRow();
        populateRelevantDetails(currentRow, data[i]);
        break;
      } else {
        var currentRow = range.getRow();
        clearAutoPopulatedCells(currentRow);
      }
    }
  }
}

// Populate relevant details into Repairs sheet
function populateRelevantDetails(row, rowData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowValues = [
    rowData[3], // Brand
    rowData[4], // Model
    rowData[2], // Supplier
    rowData[5], // Part
    rowData[7]  // Price
  ];
  sheet.getRange(row, 4, 1, rowValues.length).setValues([rowValues]);
}

// Clear auto-populated cells
function clearAutoPopulatedCells(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(row, 4, 1, 5).clearContent();
}