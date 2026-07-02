// Function 8: Determines whether to execute automatic balancing or launch the manual weekday form
function startWeekdayPhase() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Roster Completion', 
    'Would you like to automatically randomize and balance the remaining weekdays?', 
    ui.ButtonSet.YES_NO
  );  
  
  if (response === ui.Button.YES) {
    autoFillRemainingWeekdays();
  } else if (response === ui.Button.NO) {
    launchWeekdayFormWindow();
  }
}

// Function 9: The intelligent constraint allocator engine with an interactive random retry loop
function autoFillRemainingWeekdays() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var lastRow = sheet.getLastRow();
  var userProps = PropertiesService.getUserProperties();
  
  var listType = spreadsheet.getName().includes("Residents") ? "Residents" : "Department";
  var colorMapping = JSON.parse(userProps.getProperty(listType + "_Saved_Roster"));
  
  if (!colorMapping) {
    if (listType === "Residents") {
      colorMapping = { "ABJ": "#ff00ff", "ANJ": "#ffff00", "MRA": "#00ff00", "ZUV": "#ff0000" };
    } else {
      colorMapping = { "ASC": "#0000ff", "MSK": "#00ffff" };
    }
  }
  
  var leaveMapping = JSON.parse(userProps.getProperty(listType + "_Current_Leaves") || "{}");
  var residentList = Object.keys(colorMapping);
  
  // --- STEP 1: CAPTURE AND ISOLATE OPEN WEEKDAY ROWS ---
  var openWeekdayRows = [];
  for (var r = 5; r <= lastRow; r++) {
    var check1st = sheet.getRange(r, 3).getValue().toString().trim();
    var check2nd = sheet.getRange(r, 4).getValue().toString().trim();
    var dayName = sheet.getRange(r, 2).getValue().toString().toLowerCase();
    
    // If the row is empty and it is not a weekend, record it as an available open slot
    if (check1st === "" && check2nd === "" && dayName !== "sat" && dayName !== "sun") {
      openWeekdayRows.push(r);
    }
  }

  // --- STEP 2: LAUNCH THE INTERACTIVE RETRY LOOP ---
  while (true) {
    
    // Clear out any assignments from a previous optimization run to start fresh
    openWeekdayRows.forEach(function(rowNum) {
      sheet.getRange(rowNum, 3).setValue("").setBackground("#ffffff");
      sheet.getRange(rowNum, 4).setValue("").setBackground("#ffffff");
    });
    
    // Reset scorecard tracking tallies back to baseline metrics
    var stats = {};
    residentList.forEach(function(name) {
      var personLeaves = leaveMapping[name] || [];
      stats[name] = { firstCalls: 0, secondCalls: 0, totalLeavesCount: personLeaves.length };
    });
    
    // Re-tally baseline weekend shifts placed manually by Mom
    for (var r = 5; r <= lastRow; r++) {
      var weekend1 = sheet.getRange(r, 3).getValue().toString().trim().toUpperCase();
      var weekend2 = sheet.getRange(r, 4).getValue().toString().trim().toUpperCase();
      if (stats[weekend1] !== undefined) stats[weekend1].firstCalls++;
      if (stats[weekend2] !== undefined) stats[weekend2].secondCalls++;
    }
    
    // PASS 1 - Initial Distribution (With Integrated Randomization)
    for (var row = 5; row <= lastRow; row++) {
      var cell1st = sheet.getRange(row, 3);
      var cell2nd = sheet.getRange(row, 4);
      
      // Only process the rows we isolated as open weekdays
      if (openWeekdayRows.indexOf(row) !== -1) {
        var fullDateText = sheet.getRange(row, 1).getValue().toString();
        var dayNumInt = parseInt(fullDateText.split('-')[0], 10);
        
        var available1st = residentList.filter(function(name) {
          var leaves = leaveMapping[name] || [];
          return leaves.indexOf(dayNumInt) === -1;
        });
        
        var chosen1st = ""; 
        
        if (available1st.length > 0) {
          // Shuffle the available pool randomly to break sequence order bias
          available1st.sort(function() { return Math.random() - 0.5; });
          
          chosen1st = available1st.reduce(function(lowest, current) {
            return (stats[current].firstCalls < stats[lowest].firstCalls) ? current : lowest;
          }, available1st[0]);
          
          stats[chosen1st].firstCalls++;
          cell1st.setValue(chosen1st).setBackground(colorMapping[chosen1st]).setFontWeight("bold");
        } else {
          cell1st.setValue("").setBackground("#ffffff");
        }
        
        var available2nd = available1st.filter(function(name) { return name !== chosen1st; });
        
        if (available2nd.length > 0 && chosen1st !== "") {
          // Shuffle the secondary pool randomly as well
          available2nd.sort(function() { return Math.random() - 0.5; });
          
          var chosen2nd = available2nd.reduce(function(lowest, current) {
            return (stats[current].secondCalls < stats[lowest].secondCalls) ? current : lowest;
          }, available2nd[0]);
          
          stats[chosen2nd].secondCalls++;
          cell2nd.setValue(chosen2nd).setBackground(colorMapping[chosen2nd]).setFontWeight("bold");
        } else {
          cell2nd.setValue("").setBackground("#ffffff");
        }
      }
    }
    
    // PASS 2 - Global Readjustment Proximity Optimizer Loops (Fatigue Control Barrier)
    for (var optimizationCycle = 0; optimizationCycle < 10; optimizationCycle++) {
      var wasAnySwapMade = false;
      
      for (var currentRow = 5; currentRow <= lastRow; currentRow++) {
        var currentDayValue = sheet.getRange(currentRow, 2).getValue().toString().toLowerCase();
        if (currentDayValue === "sat" || currentDayValue === "sun") continue;
        
        var dateTextCurrent = sheet.getRange(currentRow, 1).getValue().toString();
        var dayNumCurrent = parseInt(dateTextCurrent.split('-')[0], 10);
        
        var name1st = sheet.getRange(currentRow, 3).getValue().toString().trim().toUpperCase();
        var name2nd = sheet.getRange(currentRow, 4).getValue().toString().trim().toUpperCase();
        
        var yesterday1st = sheet.getRange(currentRow - 1, 3).getValue().toString().trim().toUpperCase();
        var yesterday2nd = sheet.getRange(currentRow - 1, 4).getValue().toString().trim().toUpperCase();
        
        var tomorrow1st = (currentRow < lastRow) ? sheet.getRange(currentRow + 1, 3).getValue().toString().trim().toUpperCase() : "";
        var tomorrow2nd = (currentRow < lastRow) ? sheet.getRange(currentRow + 1, 4).getValue().toString().trim().toUpperCase() : "";
        
        if (name1st !== "" && (name1st === yesterday1st || name1st === tomorrow1st)) {
          var didSwap = attemptWeekdaySwap(sheet, currentRow, 3, name1st, dayNumCurrent, leaveMapping, lastRow, colorMapping);
          if (didSwap === true) wasAnySwapMade = true;
        }
        if (name2nd !== "" && (name2nd === yesterday2nd || name2nd === tomorrow2nd)) {
          var didSwap = attemptWeekdaySwap(sheet, currentRow, 4, name2nd, dayNumCurrent, leaveMapping, lastRow, colorMapping);
          if (didSwap === true) wasAnySwapMade = true;
        }
      }
      
      if (wasAnySwapMade === false) {
        break; 
      }
    }
    
    // --- STEP 3: TALLY THE ABSOLUTE FINAL SCORECARD GRID ---
    var finalReportStats = {};
    residentList.forEach(function(name) {
      finalReportStats[name] = { firsts: 0, seconds: 0 };
    });
    
    for (var r = 5; r <= lastRow; r++) {
      var final1st = sheet.getRange(r, 3).getValue().toString().trim().toUpperCase();
      var final2nd = sheet.getRange(r, 4).getValue().toString().trim().toUpperCase();
      if (finalReportStats[final1st] !== undefined) finalReportStats[final1st].firsts++;
      if (finalReportStats[final2nd] !== undefined) finalReportStats[final2nd].seconds++;
    }
    
    var scorecardSummaryText = "📊 FINAL WORKLOAD DISTRIBUTION:\n";
    residentList.forEach(function(name) {
      scorecardSummaryText += "\n👤 " + name + "  ➔  1st Call: " + finalReportStats[name].firsts + " | 2nd Call: " + finalReportStats[name].seconds;
    });
    
    // Render layout updates onto the sheet in real-time so Mom can see the background
    SpreadsheetApp.flush();
    
    // Display the interactive optimization review panel prompt
    var response = SpreadsheetApp.getUi().alert(
      '🚀 Optimization Complete', 
      'Review the stats below. Does this distribution look good?\n\n' + 
      scorecardSummaryText + '\n\n' +
      '• Click [ YES ] to Confirm & Finish Roster\n' +
      '• Click [ NO ] to Retry Optimization (Shuffles Names)', 
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (response === SpreadsheetApp.getUi().Button.YES) {
      break; // Mom is happy with the draft! Break out of the infinite while loop and exit cleanly
    }
  }
}

// --- MASTER OPTIMIZATION SWAPPER HELPER ---
function attemptWeekdaySwap(sheet, targetRow, targetColumn, conflictedPerson, targetDayNum, leaveMapping, lastRow, colorMapping) {
  for (var checkRow = 5; checkRow <= lastRow; checkRow++) {
    if (checkRow === targetRow) continue; 
    
    var checkDayValue = sheet.getRange(checkRow, 2).getValue().toString().toLowerCase();
    if (checkDayValue === "sat" || checkDayValue === "sun") continue; 
    
    var dateTextCheck = sheet.getRange(checkRow, 1).getValue().toString();
    var dayNumCheck = parseInt(dateTextCheck.split('-')[0], 10);
    
    var alternativePerson = sheet.getRange(checkRow, targetColumn).getValue().toString().trim().toUpperCase();
    
    if (alternativePerson === "") continue;
    
    var alternativeLeaves = leaveMapping[alternativePerson] || [];
    var conflictedLeaves = leaveMapping[conflictedPerson] || [];
    
    if (alternativeLeaves.indexOf(targetDayNum) !== -1) continue; 
    if (conflictedLeaves.indexOf(dayNumCheck) !== -1) continue;   
    
    var targetPartnerColumn = (targetColumn === 3) ? 4 : 3;
    var targetRowPartner = sheet.getRange(targetRow, targetPartnerColumn).getValue().toString().trim().toUpperCase();
    var checkRowPartner = sheet.getRange(checkRow, targetPartnerColumn).getValue().toString().trim().toUpperCase();
    
    if (alternativePerson === targetRowPartner || conflictedPerson === checkRowPartner) continue;
    
    var targetPrev = sheet.getRange(targetRow - 1, targetColumn).getValue().toString().trim().toUpperCase();
    var targetNext = sheet.getRange(targetRow + 1, targetColumn).getValue().toString().trim().toUpperCase();
    var checkPrev = sheet.getRange(checkRow - 1, targetColumn).getValue().toString().trim().toUpperCase();
    var checkNext = (checkRow < lastRow) ? sheet.getRange(checkRow + 1, targetColumn).getValue().toString().trim().toUpperCase() : "";
    
    if (alternativePerson !== targetPrev && alternativePerson !== targetNext && 
        conflictedPerson !== checkPrev && conflictedPerson !== checkNext) {
      
      sheet.getRange(targetRow, targetColumn).setValue(alternativePerson).setBackground(colorMapping[alternativePerson]);
      sheet.getRange(checkRow, targetColumn).setValue(conflictedPerson).setBackground(colorMapping[conflictedPerson]);
      
      return true; 
    }
  }
  return false; 
}

// Function 10: Gathers blank rows and fires up the custom HTML window panel with leave filters
function launchWeekdayFormWindow() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var lastRow = sheet.getLastRow();
  var userProps = PropertiesService.getUserProperties();
  
  var listType = spreadsheet.getName().includes("Residents") ? "Residents" : "Department";
  var savedData = userProps.getProperty(listType + "_Saved_Roster");
  var colorMapping = savedData ? JSON.parse(savedData) : null;
  
  if (!colorMapping) {
    if (listType === "Residents") {
      colorMapping = { "ABJ": "#ff00ff", "ANJ": "#ffff00", "MRA": "#00ff00", "ZUV": "#ff0000" };
    } else {
      colorMapping = { "ASC": "#0000ff", "MSK": "#00ffff" };
    }
  }
  
  var allowedNamesString = Object.keys(colorMapping).join(', ');
  var rawLeaves = userProps.getProperty(listType + "_Current_Leaves");
  var leaveMapping = rawLeaves ? JSON.parse(rawLeaves) : {};
  
  var leaveEntries = [];
  for (var resName in colorMapping) {
    var datesArray = leaveMapping[resName] || [];
    var datesText = (datesArray.length > 0) ? datesArray.join(', ') : "None";
    leaveEntries.push("<b>" + resName + "</b>: " + datesText);
  }
  var visualLeavesText = leaveEntries.join(' &nbsp;|&nbsp; ');
  
  var openDays = [];
  for (var r = 5; r <= lastRow; r++) {
    if (sheet.getRange(r, 3).getValue() === "") {
      var dayName = sheet.getRange(r, 2).getValue().toString().toUpperCase();
      var dateText = sheet.getRange(r, 1).getValue().toString().split('-')[0];
      openDays.push({ row: r, dayName: dayName, dateNum: dateText });
    }
  }
  
  var weekendStats = {};
  Object.keys(colorMapping).forEach(function(name) {
    weekendStats[name] = { firstCalls: 0, secondCalls: 0 };
  });
  
  for (var r = 5; r <= lastRow; r++) {
    var weekend1 = sheet.getRange(r, 3).getValue().toString().trim().toUpperCase();
    var weekend2 = sheet.getRange(r, 4).getValue().toString().trim().toUpperCase();
    if (weekendStats[weekend1] !== undefined) weekendStats[weekend1].firstCalls++;
    if (weekendStats[weekend2] !== undefined) weekendStats[weekend2].secondCalls++;
  }
  
  var htmlTemplate = HtmlService.createTemplateFromFile('weekdayForm');
  htmlTemplate.weekendStatsJSON = JSON.stringify(weekendStats);
  htmlTemplate.weekdays = openDays;
  htmlTemplate.listType = listType;
  htmlTemplate.allowedNames = allowedNamesString;
  htmlTemplate.visualLeaves = visualLeavesText;
  htmlTemplate.leavesJSON = JSON.stringify(leaveMapping); 
  
  var window = htmlTemplate.evaluate()
      .setWidth(450)
      .setHeight(560); 
      
  ui.showModalDialog(window, '🏥 Weekday Shift Coordinator');
}

// Function 11: Processes inputs sent back from the manual weekday HTML user form layout interface
function processWeekdayForm(payload) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var userProps = PropertiesService.getUserProperties();
  
  var listType = spreadsheet.getName().includes("Residents") ? "Residents" : "Department";
  var colorMapping = JSON.parse(userProps.getProperty(listType + "_Saved_Roster"));
  
  if (!colorMapping) {
    if (listType === "Residents") {
      colorMapping = { "ABJ": "#ff00ff", "ANJ": "#ffff00", "MRA": "#00ff00", "ZUV": "#ff0000" };
    } else {
      colorMapping = { "ASC": "#0000ff", "MSK": "#00ffff" };
    }
  }
  
  payload.forEach(function(entry) {
    var rawText = entry.textInput.trim();
    if (rawText === "") return; 
    
    var parts = rawText.split(',');
    var doc1 = parts[0] ? parts[0].trim().toUpperCase() : "";
    var doc2 = parts[1] ? parts[1].trim().toUpperCase() : "";
    
    if (colorMapping[doc1] !== undefined && colorMapping[doc2] !== undefined) {
      sheet.getRange(entry.row, 3).setValue(doc1).setBackground(colorMapping[doc1]).setFontWeight("bold");
      sheet.getRange(entry.row, 4).setValue(doc2).setBackground(colorMapping[doc2]).setFontWeight("bold");
    }
  });
  
  SpreadsheetApp.flush();
}