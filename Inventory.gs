////////////////////////////////////////////////////////////////////////////////Inventory//////////////////////////////////////////////////////////////////////////////////////////////////////

// This function is the main entry point for synchronizing stock inventory
function syncStockInventory() {
  partsToInventory();
  updatePartUsageCounts();
  stockLvl();
  updateToOrder();
}

// This function adds parts to the inventory sheet
function partsToInventory() {
  var partsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parts - Prices");
  var inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory");
  var helperSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Helper");

  // Get the list of models from Helper sheet
  var modelsRange = helperSheet.getRange("E2:E30");
  var modelsData = modelsRange.getValues();

  // Create an efficient lookup data structure for existing inventory data
  var existingData = {};
  var inventoryData = inventorySheet.getDataRange().getValues();
  for (var i = 0; i < inventoryData.length; i++) {
    var key = inventoryData[i].slice(0, 5).join(";");
    existingData[key] = true;
  }

  // Extract the list of models
  var modelsOrder = modelsData.map(function (row) {
    return row[0];
  });

  // Get all parts data from Parts sheet
  var allPartsData = partsSheet.getRange("B2:F" + partsSheet.getLastRow()).getValues();
  var newPartsData = [];

  for (var i = 0; i < modelsOrder.length; i++) {
    var model = modelsOrder[i];

    for (var j = 0; j < allPartsData.length; j++) {
      var part = allPartsData[j];
      var key = part.slice(0, 5).join(";");

      // Check if the part doesn't exist in inventory, matches the current model, and is not "Glass + OCA + Frame"
      if (!existingData[key] && part[3] === model && part[4] !== "Glass + OCA + Frame") {
        // Modify the part number to preserve leading zeros
        part[0] = "'" + part[0];
        newPartsData.push(part);
        existingData[key] = true;
      }
    }
  }

  // Add new parts data to the inventory sheet
  if (newPartsData.length > 0) {
    inventorySheet.getRange(inventorySheet.getLastRow() + 1, 1, newPartsData.length, newPartsData[0].length).setValues(newPartsData);
  }
}


// This function updates part usage counts in the Inventory sheet
function updatePartUsageCounts() {
  var stockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var repairsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Repairs - Worksheet');
  var startRow = 3596; // Specify the starting row (row 2788)

  // Get the last row with data in the 'Repairs - Worksheet' sheet
  var lastRowRepairs = Math.max(startRow, repairsSheet.getLastRow());

  // Check if there's any data in the repairs sheet
  if (lastRowRepairs >= startRow) {
    // Get repairs data starting from the specified row
    var repairsData = repairsSheet.getRange(startRow, 3, lastRowRepairs - startRow + 1, 1).getValues();

    // Calculate usage counts for parts in repairs data
    var usageCounts = calculateUsageCounts(repairsData);

    // Get product IDs from the stock data
    var stockData = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 1).getValues();

    var newUsageCounts = [];

    for (var i = 0; i < stockData.length; i++) {
      var productID = stockData[i][0];
      newUsageCounts.push([usageCounts[productID] || 0]);
    }

    // Update part usage counts in the Inventory sheet
    stockSheet.getRange(2, 8, newUsageCounts.length, 1).setValues(newUsageCounts);
  } else {
    Logger.log("No data in 'Repairs - Worksheet'");
  }
}

// This function calculates usage counts for each product ID in the repairs data
function calculateUsageCounts(repairsData) {
  var usageCounts = {};

  for (var i = 0; i < repairsData.length; i++) {
    var productID = repairsData[i][0];

    // Check if the product ID is not empty and update the usage count
    if (productID && productID !== "") {
      usageCounts[productID] = (usageCounts[productID] || 0) + 1;
    }
  }

  return usageCounts;
}

// This function updates stock levels based on the difference between columns F, G, and H
function stockLvl() {
  var stockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var lastRow = stockSheet.getLastRow();
  var valuesF = stockSheet.getRange(2, 6, lastRow - 1, 1).getValues();
  var valuesG = stockSheet.getRange(2, 7, lastRow - 1, 1).getValues();
  var valuesH = stockSheet.getRange(2, 8, lastRow - 1, 1).getValues();
  var resultValues = [];

  for (var i = 0; i < valuesF.length; i++) {
    var diff = valuesF[i][0] - valuesG[i][0] - valuesH[i][0];
    resultValues.push([diff]);
  }

  // Update stock levels in column I
  stockSheet.getRange(2, 9, resultValues.length, 1).setValues(resultValues);
}




function updateToOrder() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var helperSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Helper');
  var range = sheet.getDataRange();
  var values = range.getValues();

  // Fetch the minimum levels from the Helper sheet
  var helperData = helperSheet.getRange("S2:T" + helperSheet.getLastRow()).getValues();
  var minLevels = {};
  for (var j = 0; j < helperData.length; j++) {
    var label = helperData[j][0];
    var value = helperData[j][1];
    minLevels[label] = value;
  }

  for (var i = 1; i < values.length; i++) {
    var part = values[i][4];
    var supplier = values[i][1];
    var stockLevel = values[i][8];
    var toOrder = 0; // Initialize to 0 to prevent doubling

    // Check conditions for different parts and "Rewa" supplier
    if (supplier === 'Rewa') {
      var minLevel = minLevels[part] || 0; // Use the value from Helper sheet or default to 0

      if (part.includes('Screen') && stockLevel <= minLevel) {
        toOrder = Math.max(minLevel - stockLevel, 0); // Calculate based on the difference
      } else if (part === 'Battery' && stockLevel <= minLevel) {
        toOrder = Math.max(minLevel - stockLevel, 0); // Calculate based on the difference
      } else if ((part === 'Glass + OCA' || part === 'Frame') && stockLevel <= minLevel) {
        toOrder = Math.max(minLevel - stockLevel, 0); // Calculate based on the difference
      } else if (part.includes('Housing') && stockLevel <= 5) {
        toOrder = Math.max(5 - stockLevel, 0); // Calculate based on the difference
      }
    }

    // Update the "To Order" value in the sheet directly
    sheet.getRange(i + 1, 10).setValue(toOrder);
  }
}
