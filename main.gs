// Function 1: Adds the custom menu to the top of Google Sheets when it opens
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SPH Roster Menu')
    .addItem('Generate New Duty List', 'generateDutyList')
    .addToUi();
}

// Function 2: The primary automation interface engine using standard alert buttons
function generateDutyList() {
  var ui = SpreadsheetApp.getUi();
  
  // Step A: Open a standard message box with built-in YES and NO selection choices
  var response = ui.alert(
    'Duty List Selection', 
    'Which duty list do you want to create?\n\n👉 Click YES for: Department\n👉 Click NO for: Residents', 
    ui.ButtonSet.YES_NO
  );
  
  // Store Mom's selected option choice in a text string variable
  var listType = "";
  if (response == ui.Button.YES) {
    listType = "Department";
  } else if (response == ui.Button.NO) {
    listType = "Residents";
  } else {
    // If Mom closes the alert window out with the X button, stop the script safely right here
    return; 
  }

  // Step B: Prompt Mom to select the targeted Month and Year (using MM-YY format context)
  var monthYearResponse = ui.prompt(
    'Date Selection', 
    'Enter the Month and Year (Format: MM-YY, e.g., 02-26):', 
    ui.ButtonSet.OK_CANCEL
  );
  
  // Add a Cancel protection safeguard block
  if (monthYearResponse.getSelectedButton() !== ui.Button.OK) {
    return; // Stops the script instantly if she clicks the Cancel button
  }

  // Clean up extra blank spaces around her typed text entry
  var monthYear = monthYearResponse.getResponseText().trim();
  var parts = monthYear.split('-');
  
  var inputMonth = parseInt(parts[0], 10); // Extracts the month number integer (e.g., 2)
  var shortYear = parseInt(parts[1], 10);  // Extracts the 2-digit year integer (e.g., 26)
  
  // Convert the 2-digit year into a full 4-digit calendar year value for the engine loops
  var inputYear = 2000 + shortYear;
  
  var monthNames = [
    "January", "February", "March", "April", "May", "June", 
    "July", "August", "September", "October", "November", "December"
  ];
  var fullMonthName = monthNames[inputMonth - 1];
  
  // Step C: Pass the collected input arguments to our layout builder engine file
  buildSpreadsheetLayout(listType, fullMonthName, inputMonth, inputYear);
}