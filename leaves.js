// Function 12: Gathers active names and launches the custom leave configuration window modal
function launchLeaveFormWindow() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var userProps = PropertiesService.getUserProperties();
  
  var listType = spreadsheet.getName().includes("Residents") ? "Residents" : "Department";
  
  // Step A: Attempt to load the roster history from the cloud locker
  var savedData = userProps.getProperty(listType + "_Saved_Roster");
  var colorMapping = savedData ? JSON.parse(savedData) : null;
  
  // Safeguard: Reconstruct defaults using pure colors if cloud properties return empty on first run
  if (!colorMapping) {
    if (listType === "Residents") {
      colorMapping = { 
        "ABJ": "#ff00ff", // Pure Purple
        "ANJ": "#ffff00", // Pure Yellow
        "MRA": "#00ff00", // Pure Green
        "ZUV": "#ff0000"  // Pure Red
      };
    } else {
      colorMapping = { 
        "ASC": "#0000ff", // Pure Blue
        "MSK": "#00ffff"  // Pure Cyan
      };
    }
  }
  
  var htmlTemplate = HtmlService.createTemplateFromFile('leaveForm');
  htmlTemplate.listType = listType;
  htmlTemplate.allowedNames = Object.keys(colorMapping); // Passes active team names array to front-end loop
  
  var window = htmlTemplate.evaluate().setWidth(450).setHeight(450);
  ui.showModalDialog(window, '🏥 Resident Leave Registrar');
}

// Function 13: Processes submitted leave items and saves them safely to cloud properties memory
function processLeaveForm(leavePayload) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var userProps = PropertiesService.getUserProperties();
  var listType = spreadsheet.getName().includes("Residents") ? "Residents" : "Department";
  
  var leaveMapping = {};
  
  for (var name in leavePayload) {
    var rawText = leavePayload[name].trim();
    if (rawText === "") {
      leaveMapping[name] = []; // Map a clean empty array list if a resident has no leaves listed
    } else {
      // Split entries by comma, parse out the characters, and filter away empty formatting space errors
      leaveMapping[name] = rawText.split(',').map(function(num) {
        return parseInt(num.trim(), 10);
      }).filter(function(num) { return !isNaN(num); });
    }
  }
  
  // Save the calculated calendar index mappings object as a flat string data block
  userProps.setProperty(listType + "_Current_Leaves", JSON.stringify(leaveMapping));
  
  return "SUCCESS";
}