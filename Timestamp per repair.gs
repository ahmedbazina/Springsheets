////////////////////////////////////////////////////////////////////////////////////TimeStamp per repair///////////////////////////////////////////////////////////////////////////////////////////

// Function to handle changes in Barcode column in 'Repairs - Worksheet' and update 'Timeframe per repair' sheet
function handleBarcodeColumn(e) {
  // Get the edited range
  var editedRange = e.range;
  var sheet = e.source.getActiveSheet();

  // Check if the edited column is the barcode column in 'Repairs - Worksheet' (column B)
  if (sheet.getName() === 'Repairs - Worksheet' && editedRange.getColumn() === 2) {
    // Get the 'Timeframe per repair' sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timeframeSheet = ss.getSheetByName('Timeframe per repair');

    // Get the values from the edited range
    var barcodeValues = editedRange.getValues();

    // Loop through each cell in the edited range
    for (var i = 0; i < barcodeValues.length; i++) {
      var barcodeValue = barcodeValues[i][0]; // Get the barcode value from the 2D array

      // Calculate the last row with content in 'Timeframe per repair' sheet in column A
      var lastRow = timeframeSheet.getLastRow();

      if (barcodeValue === "") {
        // Log that the barcode is empty
        Logger.log("Barcode is empty. No data added to Timeframe per repair.");
      } else {
        // Get the current date and time
        var currentDate = new Date();

        // Format the date as 'dd/mm/yy'
        var formattedDate = Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "dd/MM/yy");

        // Format the time as 'HH:mm'
        var formattedTime = Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "HH:mm");

        // Set the barcode in 'Timeframe per repair' column C
        timeframeSheet.getRange(lastRow + 1, 3).setValue(barcodeValue); // Column C
        Logger.log("Set Barcode in Timeframe per repair Column C: " + barcodeValue);

        // Set the formatted date in 'Timeframe per repair' column A
        timeframeSheet.getRange(lastRow + 1, 1).setValue(formattedDate); // Column A
        Logger.log("Set Date in Timeframe per repair Column A: " + formattedDate);

        // Set the formatted time in 'Timeframe per repair' column B
        timeframeSheet.getRange(lastRow + 1, 2).setValue(formattedTime); // Column B
        Logger.log("Set Time in Timeframe per repair Column B: " + formattedTime);

        // Generate the formula for column D based on the row number
        var statusFormula = '=IFERROR(IF(INDEX(\'Repairs - Worksheet\'!I:I, MATCH(C' + (lastRow + 1) + ', \'Repairs - Worksheet\'!B:B, 0)), "Fixed", "Not Fixed"), "")';
        timeframeSheet.getRange(lastRow + 1, 4).setFormula(statusFormula); // Column D
        Logger.log("Set Formula in Timeframe per repair Column D: " + statusFormula);
      }
    }
  }
}

// Function to update timestamps based on status changes in 'Timeframe per repair' sheet
function updateTimestampBasedOnStatus() {
  // Get the 'Timeframe per repair' sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeframeSheet = ss.getSheetByName('Timeframe per repair');

  // Get the data range in columns D and E of 'Timeframe per repair' sheet
  var statusRange = timeframeSheet.getRange("D:D");
  var timestampRange = timeframeSheet.getRange("E:E");

  // Get the values in columns D and E
  var statusValues = statusRange.getValues();
  var timestampValues = timestampRange.getValues();

  // Get the last row in the columns
  var lastRow = statusRange.getLastRow();

  // Get the current date and time
  var currentDate = new Date();
  var formattedTime = Utilities.formatDate(currentDate, ss.getSpreadsheetTimeZone(), "HH:mm");

  // Loop through each value in column D
  for (var i = 1; i <= lastRow; i++) {
    var statusValue = statusValues[i - 1][0].toString().toLowerCase(); // Convert to lowercase for case insensitivity

    // Check if the status is "fixed" and the corresponding cell in column E is empty
    if (statusValue === "fixed" && timestampValues[i - 1][0] === "") {
      // Add a timestamp in column E in the same row
      timeframeSheet.getRange(i, 5).setValue(formattedTime);
    } else if (statusValue === "not fixed") {
      // Clear the cell in column E in the same row
      timeframeSheet.getRange(i, 5).clearContent();
    }
  }
}

