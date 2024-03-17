///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Google Forms read data and auto populate the form with data from helper sheet to make life easier change data in 1 place and it updates everywhere its meant to be

// Combined function to run both getDataFromHelper and populateQcForm
function runDataAndFormPopulation() {
  getDataFromHelper(); // Get data from the spreadsheet
  populateQcForm();   // Populate QC Form
  populateStockForm(); // Populate Stock Form
  populateFailedForm(); //Populate Failed Form
}

// this function reads data from a specific sheet in a Google Spreadsheet "Helper" and organizes it into an object where each column's title becomes a key, and the associated values are stored as an array of choices. This organized data can then be used for various purposes, such as populating questions in a Google Form, as demonstrated in your previous script. 
function getDataFromHelper(){
  // Get the active Google Spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  // Get a specific sheet within the spreadsheet by its name
  const sheet = ss.getSheetByName("Helper");
  // Get all the data (including header) from the specified sheet and convert it to an array
  const [header, ...data] = sheet.getDataRange().getDisplayValues();
  // Initialize an empty object to store the choices
  const choices = {}
  // Iterate through each column (title) in the header
  header.forEach(function(title, index) {
    // For each column, create an array of choices by extracting data from the corresponding column
    // Filter out any empty values from the array
    choices[title] = data.map(row => row[index]).filter(e => e !== "");
  });
  // Return the choices object containing the data from the sheet
  return choices;
}

// This function is to auto-populate a QC form with all the data needed, where this script opens a specific Google Form by its ID, retrieves the form items (questions), and then populates the choices for the "Repairs" and "Technicians" questions based on the data obtained from your spreadsheet using the getDataFromHelper function.
function populateQcForm() {
  // Replace with your Google Form ID
  const Google_Form_ID = "19x6HTJ7RPTl9xFZttSmlhOWjXxQ-KwlaocT3AfkQtu8";
  // Open the Google Form by its ID
  const googleForm = FormApp.openById(Google_Form_ID);
  // Get all the items (questions) in the Google Form
  const items = googleForm.getItems();
  // Get the choices (data) from the spreadsheet using the getDataFromHelper function
  const choices = getDataFromHelper();

  // Iterate through each item (question) in the Google Form
  items.forEach(function (item) {
    // Get the title (text) of the current item (question)
    const itemTitle = item.getTitle();
    
    // Check if the item is a "Repairs" question and is of type CHECKBOX
    if (itemTitle === "Repairs" && item.getType() === FormApp.ItemType.CHECKBOX) {
      // Populate choices for "Repairs" as checkboxes in question 3
      item.asCheckboxItem().setChoiceValues(choices["Repairs"]);
    }
    // Check if the item is a "Technicians" question and is of type LIST (dropdown)
    else if (itemTitle === "Technicians" && item.getType() === FormApp.ItemType.LIST) {
      // Populate choices for "Technicians" as a dropdown in question 5
      item.asListItem().setChoiceValues(choices["Technicians"]);
    }
  });
}

function populateStockForm() {
  // Replace with your Google Form ID
  const Google_Form_ID = "1xOaKPQ5acK6Xuf96521bsivgANe6dbPqOFeUrFmhnHQ";
  const googleForm = FormApp.openById(Google_Form_ID);
  const items = googleForm.getItems();
  const choices = getDataFromHelper(); // Retrieve choices from your spreadsheet

  items.forEach(function (item) {
    const itemTitle = item.getTitle(); // Get the title of the current item
    
    // Check if the item is the "Brands" section and is a LIST (dropdown) type
    if (itemTitle === "Brands" && item.getType() === FormApp.ItemType.LIST) {
      // Filter "Brands" to include only "Apple"
      const uniqueBrands = ["Apple"];
      item.asListItem().setChoiceValues(uniqueBrands); // Set choices for "Brands"
    }
    // Check if the item is the "Models" section and is a LIST (dropdown) type
    else if (itemTitle === "Models" && item.getType() === FormApp.ItemType.LIST) {
      // Filter "Models" to include anything that has "iPhone" but exclude specific models
      const filteredModels = choices["Models"].filter(model => {
        return model.includes("iPhone") && !model.includes("iPhone 12/12 Pro") && !model.includes("iPhone 8/SE20/SE22");
      });
      item.asListItem().setChoiceValues(filteredModels); // Set choices for "Models"
    }
    // Check if the item is the "Colours" section and is a LIST (dropdown) type
    else if (itemTitle === "Colours" && item.getType() === FormApp.ItemType.LIST) {
      // Filter "Models" to include anything that has "Housing" in the name
      const filteredModels = choices["Parts"].filter(model => model.includes("Housing"));
      item.asListItem().setChoiceValues(filteredModels); // Set choices for "Colours"
    }
    // Check if the item is the "Comment" section and is a LIST (dropdown) type
    else if (itemTitle === "Comment" && item.getType() === FormApp.ItemType.LIST) {
      item.asListItem().setChoiceValues(choices["Status"]); // Set choices for "Comment"
    }
  });
}

function populateFailedForm() {
  // Replace with your Google Form ID
  const Google_Form_ID = "1VGNi_zQOvPu6x8hCMRE2BOCfseYW7itoQXxcv9gSvfk";
  const googleForm = FormApp.openById(Google_Form_ID);
  const items = googleForm.getItems();
  const choices = getDataFromHelper(); // Retrieve choices from your spreadsheet

  items.forEach(function (item) {
    const itemTitle = item.getTitle(); // Get the title of the current item

    // Check if the item is the "Models" section and is a LIST (dropdown) type
     if (itemTitle === "Models" && item.getType() === FormApp.ItemType.LIST) {
      // Filter "Models" to include anything that has "iPhone" but exclude specific models
      const filteredModels = choices["Models"].filter(model => {
        return model.includes("iPhone") && !model.includes("iPhone 12/12 Pro") && !model.includes("iPhone 8/SE20/SE22");
      });
      item.asListItem().setChoiceValues(filteredModels); // Set choices for "Models"
    }
  });
}