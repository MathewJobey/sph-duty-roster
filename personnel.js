// Function 4: The custom roster builder engine with integrated cloud storage properties memory
function setupResidentsAndColors(listType) {
  var ui = SpreadsheetApp.getUi();
  
  // Step 1: Connect to the spreadsheet account's cloud properties memory storage
  var userProps = PropertiesService.getUserProperties();
  var saveKey = listType + "_Saved_Roster"; // Generates a unique key string like "Residents_Saved_Roster"
  
  // Step 2: Create a translation dictionary that turns hex color codes into colored square emojis
  var emojiTranslator = {
    "#ff00ff": "🟪", // Pure Purple
    "#ffff00": "🟨", // Pure Yellow
    "#00ff00": "🟩", // Pure Green
    "#ff0000": "🟥", // Pure Red
    "#0000ff": "🟦", // Pure Blue
    "#00ffff": "🟦"  // Pure Cyan (Maps to blue emoji for uniform clarity)
  };

  // Step 3: Check for a saved list from last month, otherwise load system defaults
  var savedRosterString = userProps.getProperty(saveKey);
  var currentMapping = {};
  
  if (savedRosterString !== null) {
    // A saved roster exists! Parse the JSON string back into a usable data object
    currentMapping = JSON.parse(savedRosterString);
  } else {
    // No cloud memory found. Load default base configurations depending on selection
    if (listType === "Residents") {
      currentMapping = {
        "ABJ": "#ff00ff",
        "ANJ": "#ffff00",
        "MRA": "#00ff00",
        "ZUV": "#ff0000"
      };
    } else {
      currentMapping = {
        "ASC": "#0000ff",
        "MSK": "#00ffff"
      };
    }
  }

  // Step 4: Build a visual layout list using our emoji translator icons
  var emojiPromptText = "";
  for (var res in currentMapping) {
    var color = currentMapping[res];
    var icon = emojiTranslator[color] || "⬜"; // Fallback to white box icon if a custom hex code breaks
    emojiPromptText += icon + " " + res + "\n";
  }

  // Step 5: Present the verification message box layout to Mom
  var choiceResponse = ui.alert(
    'Roster Memory Loaded', 
    'Would you like to use this ' + listType + ' list?\n\n' + emojiPromptText, 
    ui.ButtonSet.YES_NO
  );
  
  var finalizedColorMapping = {};
  
  if (choiceResponse === ui.Button.YES) {
    // Mom accepted the loaded list! Lock it in as our active operational mapping configuration
    finalizedColorMapping = currentMapping;
    
  } else if (choiceResponse === ui.Button.NO) {
    // Step 6: Mom wants a fresh roster. Open a clean text prompt window framework box
    var customResponse = ui.prompt(
      'New Roster Setup', 
      'Enter the new 3-letter codes separated by commas\n(e.g., KBC, LMR, PQT):', 
      ui.ButtonSet.OK_CANCEL
    );
    
    // Safely stop execution if she hits Cancel or closes the window prompt box out
    if (customResponse.getSelectedButton() !== ui.Button.OK) return null;
    
    var rawText = customResponse.getResponseText();
    var customNames = rawText.split(',').map(function(name) {
      return name.trim().toUpperCase(); // Removes extra accidental spaces and makes letters capital
    });
    
    // Fallback cycle array containing our vivid theme color palettes
    var fallbackPastels = ['#ff00ff', '#ffff00', '#00ff00', '#ff0000', '#0000ff', '#00ffff'];
    var newSummaryText = "New " + listType + " Roster Saved!\n\n";
    
    // Loop through her new inputs and assign colors step-by-step
    for (var i = 0; i < customNames.length; i++) {
      var name = customNames[i];
      var colorIndex = i % fallbackPastels.length; // Keeps loop safe even if she enters 20 names
      var assignedColor = fallbackPastels[colorIndex];
      
      finalizedColorMapping[name] = assignedColor;
      newSummaryText += emojiTranslator[assignedColor] + " " + name + "\n";
    }
    
    // Step 7: Write this newly generated list structure into cloud memory for next month
    userProps.setProperty(saveKey, JSON.stringify(finalizedColorMapping));
    
    // Display a confirmation window showing the new colored roster array setup
    ui.alert('✅ Saved for Next Month', newSummaryText, ui.ButtonSet.OK);
    
  } else {
    return null; // Stop runtime execution cleanly if window was bypassed
  }
  
  // Step 8: Route execution smoothly into your custom layout leave wizard
  launchLeaveFormWindow();
}