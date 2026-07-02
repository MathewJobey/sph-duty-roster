// Function 5: Handoff Helper - Lets the leave window launch the weekend coordinator cleanly
function triggerWeekendFormLaunch() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var userProps = PropertiesService.getUserProperties();
  var listType = spreadsheet.getName().includes("Residents") ? "Residents" : "Department";
  
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

  assignWeekendShifts(colorMapping);
}

// Function 6: Scans the sheet chronologically and displays the custom HTML form window modal
function assignWeekendShifts(colorMapping) {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var lastRow = sheet.getLastRow();
  
  var weekendDays = [];
  
  for (var r = 5; r <= lastRow; r++) {
    var dayValue = sheet.getRange(r, 2).getValue().toString().toLowerCase();
    var fullDateText = sheet.getRange(r, 1).getValue().toString();
    var dayNumber = fullDateText.split('-')[0]; 
    
    if (dayValue === "sat" || dayValue === "sun") {
      weekendDays.push({
        row: r,
        dayName: dayValue.toUpperCase(),
        dateNum: dayNumber
      });
    }
  }
  
  var htmlTemplate = HtmlService.createTemplateFromFile('weekendForm');
  var listType = spreadsheet.getName().includes("Residents") ? "Residents" : "Department";
  var allowedNamesString = Object.keys(colorMapping).join(', ');
  
  // Load recorded leaves for visual display and client-side JavaScript execution
  var userProps = PropertiesService.getUserProperties();
  var rawLeaves = userProps.getProperty(listType + "_Current_Leaves");
  var leaveMapping = rawLeaves ? JSON.parse(rawLeaves) : {};
  
  // Create a clean readable line for the UI header block: "ABJ: 4, 12 | ANJ: None"
  var leaveEntries = [];
  for (var resName in colorMapping) {
    var datesArray = leaveMapping[resName] || [];
    var datesText = (datesArray.length > 0) ? datesArray.join(', ') : "None";
    leaveEntries.push("<b>" + resName + "</b>: " + datesText);
  }
  var visualLeavesText = leaveEntries.join(' &nbsp;|&nbsp; ');
  
  // Inject variables into our HTML template container workspace context
  htmlTemplate.weekends = weekendDays;
  htmlTemplate.listType = listType;
  htmlTemplate.allowedNames = allowedNamesString;
  htmlTemplate.visualLeaves = visualLeavesText; 
  htmlTemplate.leavesJSON = JSON.stringify(leaveMapping); 
  
  var htmlWindow = htmlTemplate.evaluate()
      .setWidth(450)
      .setHeight(560); 
      
  ui.showModalDialog(htmlWindow, '🏥 Weekend Shift Coordinator');
}

// Function 7: The receiving terminal that processes submitted form inputs from the HTML client side
function processFormSubmission(formEntries) {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var userProps = PropertiesService.getUserProperties();
  
  var listType = spreadsheet.getName().includes("Residents") ? "Residents" : "Department";
  var savedData = userProps.getProperty(listType + "_Saved_Roster");
  var colorMapping = savedData ? JSON.parse(savedData) : null;
  
  // Safeguard: Reconstruct defaults if properties memory returns empty
  if (!colorMapping) {
    if (listType === "Residents") {
      colorMapping = { "ABJ": "#ff00ff", "ANJ": "#ffff00", "MRA": "#00ff00", "ZUV": "#ff0000" };
    } else {
      colorMapping = { "ASC": "#0000ff", "MSK": "#00ffff" };
    }
  }
  
  // Loop through the data Mom just confirmed on the HTML screen layout dashboard preview
  for (var i = 0; i < formEntries.length; i++) {
    var entry = formEntries[i];
    var row = entry.row;
    var rawText = entry.textInput.trim();
    
    if (rawText === "") continue;
    
    var parts = rawText.split(',');
    var doc1 = parts[0] ? parts[0].trim().toUpperCase() : "";
    var doc2 = parts[1] ? parts[1].trim().toUpperCase() : "";
    
    // Backup check to ensure row safety data pairs are complete
    if (doc1 === "" || doc2 === "") {
      ui.alert('⚠️ Missing Entry', 'Row for ' + entry.dayName + ' ' + entry.dateNum + ' must have two names separated by a comma. Row skipped.', ui.ButtonSet.OK);
      continue;
    }
    
    // Backup check to ensure codes exist on the master registry roster checklist
    if (colorMapping[doc1] === undefined || colorMapping[doc2] === undefined) {
      ui.alert('⚠️ Unrecognized Name', 'One of the names entered for ' + entry.dayName + ' ' + entry.dateNum + ' does not exist in your roster list. Row skipped.', ui.ButtonSet.OK);
      continue;
    }
    
    // Write 1st Call values safely to Column 3 (Column C)
    var cell1st = sheet.getRange(row, 3);
    cell1st.setValue(doc1).setBackground(colorMapping[doc1]).setFontWeight("bold");
    
    // Write 2nd Call values safely to Column 4 (Column D)
    var cell2nd = sheet.getRange(row, 4);
    cell2nd.setValue(doc2).setBackground(colorMapping[doc2]).setFontWeight("bold");
  }
  
  // Force Google Sheets layout framework to update cells immediately
  SpreadsheetApp.flush();
  
  // Return a clean success receipt message back to the HTML interface handler
  return "SUCCESS"; 
}