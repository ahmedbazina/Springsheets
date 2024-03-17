//////////////////////////////////////////////////////////////// this section is for Sidebar menu with the different forms available ///////////////////////////////////////////////////////////////

// Custom menu to display the sidebar on specific sheets
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Data entry")
    .addItem("Add QC repairs", "showQCForm") // Add a menu item "Add New QC" that triggers the showQCForm function
    .addItem("Add Buffer Devices", "showBufferForm") // Add a menu item "Add New Stock" that triggers the showBufferForm function
    .addItem("Add New Stock & Faulty Parts", "showStockForm") // Add a menu item "Add New Stock" that triggers the showBufferForm function
    .addItem("QC/Returns/Sideline", "showFailedForm") // Add a menu item "Add New Stock" that triggers the showBufferForm function
    .addToUi(); // Add the menu to the Google Sheets UI
}

// Function to show the QC Form sidebar
function showQCForm() {
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile("QC.html").setTitle("Please enter your QC Repairs"));
}

// Function to show the Buffer Form sidebar
function showBufferForm() {
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile("Buffer.html").setTitle("Please enter new buffer phones"));
}

// Function to show the Stock sidebar
function showStockForm() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile("Stock").setTitle("Please enter new stock and faulty parts"));
}

// Function to show the QC Form sidebar
function showFailedForm() {
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile("Fails.html").setTitle("Please enter your failed repairs"));
}

// Function to get unique supplier data from the Inventory sheet
function getSupplierData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var data = sheet.getRange('B2:B' + sheet.getLastRow()).getValues().flat();
  var uniqueSuppliers = [...new Set(data)];

  Logger.log("Unique Suppliers: " + uniqueSuppliers); // Log the unique suppliers

  return uniqueSuppliers;
}

// Function to get unique model data from the Inventory sheet based on the selected supplier
function getModelData(selectedSupplier) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var data = sheet.getDataRange().getValues();
  var uniqueModels = [];

  for (var i = 1; i < data.length; i++) {
    var supplier = data[i][1]; // Assuming Supplier is in column B (index 1)
    var model = data[i][3]; // Assuming Model is in column D (index 3)

    if (supplier === selectedSupplier && model) {
      uniqueModels.push(model);
    }
  }

  // Remove duplicates from uniqueModels array
  uniqueModels = [...new Set(uniqueModels)];

  Logger.log("Unique Models: " + uniqueModels);

  return uniqueModels;
}

// Function to get unique part data from the Inventory sheet based on the selected model and supplier
function getPartData(selectedModel, selectedSupplier) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var data = sheet.getDataRange().getValues();
  var uniqueParts = [];

  for (var i = 1; i < data.length; i++) {
    var supplier = data[i][1]; // Assuming Supplier is in column B (index 1)
    var model = data[i][3]; // Assuming Model is in column D (index 3)
    var part = data[i][4]; // Assuming Part is in column E (index 4)

    if (supplier === selectedSupplier && model === selectedModel && part) {
      uniqueParts.push(part);
    }
  }

  // Remove duplicates from uniqueParts array
  uniqueParts = [...new Set(uniqueParts)];

  Logger.log("Unique Parts: " + uniqueParts);

  return uniqueParts;
}

function submitRecord(supplier, model, part, quantity) {
  // Log the values received
  Logger.log("Received values - Supplier: " + supplier + ", Model: " + model + ", Part: " + part + ", Quantity: " + quantity);

  // Get the active spreadsheet and the "Inventory" sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Inventory');

  // Find the row that matches the selected supplier, model, and part
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    // Log the values in the current row for debugging
    Logger.log("Checking row " + i + " - Supplier: " + data[i][1] + ", Model: " + data[i][3] + ", Part: " + data[i][4]);

    if (data[i][1] === supplier && data[i][3] === model && data[i][4] === part) {
      // Get the existing quantity from column F (index 5)
      var existingQuantity = parseFloat(data[i][5]);
      if (isNaN(existingQuantity)) {
        existingQuantity = 0; // Set to 0 if no existing quantity
      }

      // Update the quantity in column F with the combined quantity
      sheet.getRange(i + 1, 6).setValue(existingQuantity + quantity);
      return existingQuantity + quantity;
    }
  }

  // If the combination of supplier, model, and part doesn't exist, you can handle it here.
  // For example, you can log an error or show a message to the user.
  Logger.log("Record not found for the selected combination - Supplier: " + supplier + ", Model: " + model + ", Part: " + part);
  return null; // Return null if the record is not found
}

// Function to get unique faulty supplier data from the Inventory sheet
function getFaultySupplierData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var data = sheet.getRange('B2:B' + sheet.getLastRow()).getValues().flat();
  var uniqueSuppliers = [...new Set(data)];

  Logger.log("Unique Faulty Suppliers: " + uniqueSuppliers); // Log the unique faulty suppliers

  return uniqueSuppliers;
}

// Function to get unique faulty model data from the Inventory sheet based on the selected faulty supplier
function getFaultyModelData(selectedSupplier) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var data = sheet.getDataRange().getValues();
  var uniqueModels = [];

  for (var i = 1; i < data.length; i++) {
    var supplier = data[i][1]; // Assuming Supplier is in column B (index 1)
    var model = data[i][3]; // Assuming Model is in column D (index 3)

    if (supplier === selectedSupplier && model) {
      uniqueModels.push(model);
    }
  }

  // Remove duplicates from uniqueModels array
  uniqueModels = [...new Set(uniqueModels)];

  Logger.log("Unique Faulty Models: " + uniqueModels);

  return uniqueModels;
}

// Function to get unique faulty part data from the Inventory sheet based on the selected faulty model and supplier
function getFaultyPartData(selectedModel, selectedSupplier) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var data = sheet.getDataRange().getValues();
  var uniqueParts = [];

  for (var i = 1; i < data.length; i++) {
    var supplier = data[i][1]; // Assuming Supplier is in column B (index 1)
    var model = data[i][3]; // Assuming Model is in column D (index 3)
    var part = data[i][4]; // Assuming Part is in column E (index 4)

    if (supplier === selectedSupplier && model === selectedModel && part) {
      uniqueParts.push(part);
    }
  }

  // Remove duplicates from uniqueParts array
  uniqueParts = [...new Set(uniqueParts)];

  Logger.log("Unique Faulty Parts: " + uniqueParts);

  return uniqueParts;
}


function submitFaultyRecord(fsupplier, fmodel, fpart, fquantity) {
  // Log the values received
  Logger.log("Received values - Supplier: " + fsupplier + ", Model: " + fmodel + ", Part: " + fpart + ", Quantity: " + fquantity);

  // Get the active spreadsheet and the "Inventory" sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Inventory');

  // Find the row that matches the selected fmodel and fpart
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    // Log the values in the current row for debugging
    Logger.log("Checking row " + i + " - fSupplier: " + data[i][1] + ", fModel: " + data[i][3] + ", fPart: " + data[i][4]);

    if (data[i][1] === fsupplier && data[i][3] === fmodel && data[i][4] === fpart) {
      // Get the existing quantity from column G (index 6)
      var existingfQuantity = parseFloat(data[i][6]);
      if (isNaN(existingfQuantity)) {
        existingfQuantity = 0; // Set to 0 if no existing fquantity
      }

      // Update the fquantity in column G with the combined fquantity
      sheet.getRange(i + 1, 7).setValue(existingfQuantity + fquantity);
      return existingfQuantity + fquantity;
    }
  }

  // If the combination of model and fpart doesn't exist, you can handle it here.
  // For example, you can log an error or show a message to the user.
  Logger.log("Record not found for the selected combination - Model: " + fmodel + ", Part: " + fpart);
  return null; // Return null if the record is not found
}



