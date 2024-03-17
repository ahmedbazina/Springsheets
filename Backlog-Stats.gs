// Define a function called backlogStatstest
function statisticsData() {
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ms = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1IyeTZu7CS8jfCnKmtXWjJyhRPOv8JE8eaK8bK9BTlFo/edit?pli=1");

  // Get references to different sheets in the spreadsheet
  var statsSheet = ss.getSheetByName('Statistics');
  var failedSheet = ms.getSheetByName('Failed repairs');
  var scannedDevicesSheet = ms.getSheetByName('Scanned Devices')

  // Get data from various sheets
  var failedData = failedSheet.getRange("B2:D" + failedSheet.getLastRow()).getValues();
  var scannedData = scannedDevicesSheet.getRange("B2:F" + scannedDevicesSheet.getLastRow()).getValues();

  // Get data from the 'Helper' sheet
  var helperSheet = ss.getSheetByName('Helper');
  var modelsRange = helperSheet.getRange("E2:E30");
  var modelsData = modelsRange.getValues();

  // Create an array of models
  var models = modelsData.map(function(row) {
    return row[0];
  });

  // Initialize empty counters for various counts
  // Buffer Form scanned devices data
  var stockCounts = {};
  var repairsCounts = {};
  var polishCounts = {};
  var refurbCounts = {};
  var nandCounts = {};

  // Failed Form data
  var failedCounts = {};
  var sidelinedCounts = {};
  var customerReturnsCounts = {};


  // Loop through stock data and count models based on comment
  for (var i = 0; i < scannedData.length; i++) {
    var model = scannedData[i][2]; // Model in column D
    var comment = scannedData[i][4]; // Comment in column F

    // Check if the model is in the list
    if (models.indexOf(model) !== -1) {
      // Update stockCounts for all comments
      stockCounts[model] = (stockCounts[model] || 0) + 1;

      // Update counts based on specific comments
      if (comment === "Repair") {
        repairsCounts[model] = (repairsCounts[model] || 0) + 1;
      } else if (comment === "Polish") {
        polishCounts[model] = (polishCounts[model] || 0) + 1;
      } else if (comment === "Refurb") {
        refurbCounts[model] = (refurbCounts[model] || 0) + 1;
      } else if (comment === "NAND") {
        nandCounts[model] = (nandCounts[model] || 0) + 1;
      }
    }
  }

  // Loop through the failed data and count models based on location
  for (var i = 0; i < failedData.length; i++) {
    var model = failedData[i][1]; // Model in column C
    var location = failedData[i][2]; // Location in column D

    // Check if the model is in the list and update counts based on location
    if (models.indexOf(model) !== -1) {
      if (location === "QC Fails") {
        failedCounts[model] = (failedCounts[model] || 0) + 1;
      } else if (location === "Mid Repair/Sidelined") {
        sidelinedCounts[model] = (sidelinedCounts[model] || 0) + 1;
      } else if (location === "Customer Return") {
        customerReturnsCounts[model] = (customerReturnsCounts[model] || 0) + 1;
      }
    }
  }

  // Prepare the final output array
  var output = models.map(function(model) {
    return [model, stockCounts[model] || 0, repairsCounts[model] || 0, polishCounts[model] || 0, refurbCounts[model] || 0, nandCounts[model] || 0, failedCounts[model] || 0, customerReturnsCounts[model] || 0, sidelinedCounts[model] || 0];
  });

  // Update the 'statsSheet' with the output data
  statsSheet.getRange(1, 1, 1, 9).setValues([["Model", "Total Stock", "Repairs", "Polish", "Glass Refurb", "Nand", "QC Fails", "Customer Returns", "Sidelined"]]);
  statsSheet.getRange("A2:I30").clearContent();
  statsSheet.getRange(2, 1, output.length, 9).setValues(output);

  // Log the results for debugging purposes
  // Logger.log("List of iPhone Models: " + models.join(', '));
  // Logger.log("Stock Counts: " + JSON.stringify(stockCounts) + "\n\nRepairs Counts: " + JSON.stringify(repairsCounts) + "\n\nPolish Counts: " + JSON.stringify(polishCounts) + "\n\nRefurb Counts: " + JSON.stringify(refurbCounts) + "\n\nNAND Counts: " + JSON.stringify(nandCounts));
  // Logger.log("Failed Counts: " + JSON.stringify(failedCounts) + "\n\nSidelined Counts: " + JSON.stringify(sidelinedCounts) + "\n\nCustomer Returns Counts: " + JSON.stringify(customerReturnsCounts));
}


// Function to delete rows in 'Scanned Devices' sheet based on matching barcodes in 'Repairs - Worksheet' and 'Barcode - Sales' sheets
function deleteRepairedDevices() {
  // Replace these URLs with your actual Google Sheets URLs
  var masterSheetUrl = 'https://docs.google.com/spreadsheets/d/1IyeTZu7CS8jfCnKmtXWjJyhRPOv8JE8eaK8bK9BTlFo/edit?pli=1#gid=403108654';
  var springTechSheetUrl = 'https://docs.google.com/spreadsheets/d/1mSIayi1JQBA5EesKbSiUQV-EmY4PhqNtwyLbt1N91do/edit#gid=446545246';

  var masterSheet = SpreadsheetApp.openByUrl(masterSheetUrl);
  var repairworkSheet = SpreadsheetApp.openByUrl(springTechSheetUrl);

  // Scanned Devices and Repair Worksheet data
  var scannedDevicesSheet = masterSheet.getSheetByName('Scanned Devices');
  var repairsSheet = repairworkSheet.getSheetByName('Repairs - Worksheet');

  // Failed devices and Repairs QC data
  var failedDevicesSheet = masterSheet.getSheetByName('Failed Repairs');
  var repairsQCSheet = repairworkSheet.getSheetByName('Repairs - QC');

  var scannedDevicesBarcodes = scannedDevicesSheet.getRange('B2:B').getValues().flat();
  var failedDevicesBarcodes = failedDevicesSheet.getRange('B2:B').getValues().flat();

  deleteRepairedDevicesInScannedDevices(scannedDevicesSheet, repairsSheet, scannedDevicesBarcodes);
  deleteRepairedDevicesInFailedRepairs(failedDevicesSheet, repairsQCSheet, failedDevicesBarcodes);
  deleteScannedDevicesInFailedRepairs(scannedDevicesSheet, failedDevicesSheet, scannedDevicesBarcodes);  
}

function deleteRepairedDevicesInScannedDevices(scannedDevicesSheet, repairsSheet, scannedDevicesBarcodes) {
  var repairsBarcodes = repairsSheet.getRange('B:B').getValues().flat();

  // Assuming the first row is the header row, start the loop from the second row (index 1)
  for (var i = scannedDevicesBarcodes.length - 1; i > 0; i--) {
    var barcode = scannedDevicesBarcodes[i];
    var inRepairs = repairsBarcodes.includes(barcode);

    if (inRepairs) {
      // Log the deletion
      Logger.log('Deleted row in Scanned Devices where Barcode = %s', barcode);

      // Match found in Repairs, delete the row in Scanned Devices
      scannedDevicesSheet.deleteRow(i + 2); // Adjust the index to account for header row
    }
  }
}

function deleteRepairedDevicesInFailedRepairs(failedDevicesSheet, repairsQCSheet, failedDevicesBarcodes) {
  var repairsQCBarcodes = repairsQCSheet.getRange('B:B').getValues().flat();

  // Assuming the first row is the header row, start the loop from the second row (index 1)
  for (var i = failedDevicesBarcodes.length - 1; i > 0; i--) {
    var qcBarcode = failedDevicesBarcodes[i];
    var inQCRepairs = repairsQCBarcodes.includes(qcBarcode);

    if (inQCRepairs) {
      // Log the deletion
      Logger.log('Deleted row in Failed Repairs where Barcode = %s', qcBarcode);

      // Match found in Repairs, delete the row in Scanned Devices
      failedDevicesSheet.deleteRow(i + 2); // Adjust the index to account for header row
    }
  }
}



function deleteScannedDevicesInFailedRepairs(scannedDevicesSheet, failedDevicesSheet, scannedDevicesBarcodes) {
  var failedDevicesBarcodes = failedDevicesSheet.getRange('B:B').getValues().flat();
  var failedDevicesLocations = failedDevicesSheet.getRange('D:D').getValues().flat();

  // Assuming the first row is the header row, start the loop from the second row (index 1)
  for (var i = scannedDevicesBarcodes.length - 1; i > 0; i--) {
    var barcode = scannedDevicesBarcodes[i];
    var inFailedRepairs = failedDevicesBarcodes.includes(barcode);
    var locationIndex = failedDevicesBarcodes.indexOf(barcode);
    var location = inFailedRepairs ? failedDevicesLocations[locationIndex] : "";

    if (inFailedRepairs && location === "Mid Repair/Sidelined") {
      // Log the deletion
      Logger.log('Deleted row in Scanned Devices where Barcode = %s', barcode);

      // Match found in Failed Repairs with the specified location, delete the row in Scanned Devices
      scannedDevicesSheet.deleteRow(i + 2); // Adjust the index to account for header row
    }
  }
}
