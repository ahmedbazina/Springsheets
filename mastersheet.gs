// Global variable for Spring Tech Sheet ID
var SPRING_TECH_SHEET_ID = "1mSIayi1JQBA5EesKbSiUQV-EmY4PhqNtwyLbt1N91do";
var RESPONSES_SHEET_NAME = "QC & Returns"; // Name of the responses sheet
var REPAIRS_SHEET_NAME = "Repairs - QC"; // Name of the repairs sheet

function onFormSubmit(e) {
  // Check if the form submission is from the "QC & Returns" form
  var returnReason = e.namedValues['Return'][0];
  if (!(returnReason === 'QC Return' || returnReason === 'Customer Return')) {
    // If not, exit the function
    return;
  }

  // Get the submitted data from the QC & Returns
  var timestamp = e.values[0];
  var barcode = e.values[1];
  var repairs = e.values[3].split(', ');
  var time = e.values[4];
  var technician = e.values[5];
  var comments = e.values[6];

  // Format the timestamp to "dd/mm/yy"
  var formattedTimestamp = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), "dd/MM/yy");

  // Open the "Spring Tech" document by ID
  var springTechSheet = SpreadsheetApp.openById(SPRING_TECH_SHEET_ID);

  // Find the target sheet by name ("Repairs - QC")
  var springTechDataSheet = null;
  var sheets = springTechSheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === REPAIRS_SHEET_NAME) {
      springTechDataSheet = sheets[i];
      break;
    }
  }

  // Check if the target sheet was found
  if (springTechDataSheet !== null) {
    // Iterate through each repair and append it as a separate row
    for (var i = 0; i < repairs.length; i++) {
      // Determine the last row in the target sheet to append data below existing entries
      var lastRow = springTechDataSheet.getLastRow() + 1;

      // Create an array for the row data
      var rowData = [];

      // If it's the first repair, include the barcode in the row data
      if (i === 0) {
        rowData.push(formattedTimestamp, barcode, returnReason, repairs[i], time, technician, comments);
      } else {
        rowData.push('', '', returnReason, repairs[i], time, technician, comments);
      }

      // Add the row data to the target sheet
      springTechDataSheet.getRange(lastRow, 1, 1, rowData.length).setValues([rowData]);

      // If it's the last repair, apply a bottom border to the entire row from column A to G
      if (i === repairs.length - 1) {
        var lastRowRange = springTechDataSheet.getRange(lastRow, 1, 1, rowData.length);
        lastRowRange.setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
      }
    }
  } else {
    // Log an error message if the target sheet was not found
    Logger.log("Target sheet not found: " + REPAIRS_SHEET_NAME);
  }
}


//This script essentially does the following:

//1.Opens the specified Google Sheets document using its ID.
//2.Retrieves the "Failed Repairs" sheet from the opened spreadsheet.
//3.Opens the specified Google Form using its ID.
//4.Extracts unique values from the "Barcode" column in the sheet and combines them into a pattern using '|' as a delimiter.
//5.Filters the form items to find the one with the title "Barcode" and converts it to a TextItem.
//6.Creates a text validation rule that requires the text to not match the pattern derived from the spreadsheet data.
//7.Applies the created text validation rule to the "Barcode" form item in the Google Form.
//This will help ensure that the input for the "Barcode" field in the Google Form does not match any of the values already present in the specified sheet.

// The ID of the Google Sheets document containing the data
var sheetId = "1IyeTZu7CS8jfCnKmtXWjJyhRPOv8JE8eaK8bK9BTlFo";

// Form and sheet mappings
var formSheetMapping = {
    "1VGNi_zQOvPu6x8hCMRE2BOCfseYW7itoQXxcv9gSvfk": "Failed Repairs",
    "1xOaKPQ5acK6Xuf96521bsivgANe6dbPqOFeUrFmhnHQ": "Scanned Devices"
};

function checkDuplicates() {
    // Open the Google Sheets document using the provided ID
    var ss = SpreadsheetApp.openById(sheetId);

    // Loop through each form and corresponding sheet
    for (var formId in formSheetMapping) {
        // Get the specific sheet by name
        var sheet = ss.getSheetByName(formSheetMapping[formId]);

        // Open the Google Form using the corresponding ID
        var form = FormApp.openById(formId);

        // Barcode
        // Extract unique values from the "Barcode" column in the spreadsheet and join them using a '|' delimiter
        var data = [...new Set(sheet.getDataRange().getDisplayValues().map(row => row[1]))].join('|');

        // Filter form items to find the one with the title "Barcode" and treat it as a TextItem
        var item = form.getItems(FormApp.ItemType.TEXT).filter(item => item.getTitle() == 'Barcode')[0].asTextItem();

        // Create a regular expression pattern using the unique values from the spreadsheet
        var pattern = `(${data})`;

        // Create a text validation rule for the form item
        var textValidation = FormApp.createTextValidation()
            .setHelpText("Already in sheet !!")
            .requireTextDoesNotMatchPattern(pattern)
            .build();

        // Apply the text validation rule to the "Barcode" form item
        item.setValidation(textValidation);
    }
}

