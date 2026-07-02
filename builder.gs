// Function 3: Builds the grid visual layout using the calendar variables passed to it
function buildSpreadsheetLayout(listType, fullMonthName, inputMonth, inputYear) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Step 1: Clear and Reset the Canvas
  sheet.clear();
  sheet.clearFormats();
  
  // Step 2: Create a short 2-digit year text string with an apostrophe (e.g., '26)
  var yearText = "'" + (inputYear - 2000);
  
  // Step 3: Execute File Renaming (Updated to use SPH privacy indicator)
  var fileName = fullMonthName + "_" + yearText + "-" + "Plastic_Surgery" + "-" + listType + "_Duty_List-SPH";
  spreadsheet.setName(fileName);
  
  // Step 4: Build the Main Large Heading (Print-Optimized Visual Hierarchy)
  var mainTitle = fullMonthName + " " + yearText + " " + "Plastic Surgery" + " " + listType + " Duty List";
  var locationTitle = "SPH";
  
  // Row 1 Setup: The large dominant title banner
  sheet.getRange("A1").setValue(mainTitle);
  sheet.getRange("A1:D1").merge();
  sheet.getRange("A1").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center"); 
  
  // Row 2 Setup: The smaller, elegant secondary location text
  var locationCell = sheet.getRange("A2");
  locationCell.setValue(locationTitle);
  sheet.getRange("A2:D2").merge();
  locationCell.setFontSize(11);
  locationCell.setFontWeight("normal");
  locationCell.setFontStyle("italic");
  locationCell.setFontColor("#5f6368"); 
  locationCell.setHorizontalAlignment("center");
  
  // Step 5: Set Up Column Headers on Row 4 (Row 3 is left intentionally blank as a spacer)
  var headers = ["Date", "Day", "1st Call", "2nd Call"];
  sheet.getRange("A4:D4").setValues([headers]).setFontWeight("bold").setBackground("#DADADA").setFontSize(12);
  
  // Step 6: Draw the Dynamic Calendar Rows and Highlight Weekends
  var daysInMonth = new Date(inputYear, inputMonth, 0).getDate();
  var dayNames = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];
  var weekendColor = "#8FC7FF"; 
  
  for (var i = 1; i <= daysInMonth; i++) {
    var row = i + 4; // Data rows start at Row 5 because of headers
    
    var dayStr = i < 10 ? "0" + i : i;
    var monthStr = inputMonth < 10 ? "0" + inputMonth : inputMonth;
    var dateValue = dayStr + "-" + monthStr + "-" + inputYear;
    
    var tempDate = new Date(inputYear, inputMonth - 1, i);
    var dayValue = dayNames[tempDate.getDay()];
    
    sheet.getRange(row, 1).setValue(dateValue).setNumberFormat("@").setFontSize(11);
    sheet.getRange(row, 2).setValue(dayValue).setFontSize(11);
    
    if (dayValue === "sat" || dayValue === "sun") {
      sheet.getRange(row, 1, 1, 2).setBackground(weekendColor);
    }
    
    // Set data row heights to be tall and readable for printing
    sheet.setRowHeight(row, 28); 
  }
  
  // Step 7: Set custom row heights for the top titles and column headers
  sheet.setRowHeight(1, 35); // Main Title Row
  sheet.getRange("A1").setVerticalAlignment("middle");
  
  sheet.setRowHeight(2, 22); // Subtitle Row
  sheet.getRange("A2").setVerticalAlignment("middle");
  
  sheet.setRowHeight(3, 15); // Blank Spacer Row
  
  sheet.setRowHeight(4, 30); // Column Headers Row
  sheet.getRange("A4:D4").setVerticalAlignment("middle").setFontSize(11);
  
  // Step 8: Set explicit column widths to perfectly fit standard paper dimensions
  sheet.setColumnWidth(1, 120); // Date Column width in pixels
  sheet.setColumnWidth(2, 120); // Day Column width in pixels
  sheet.setColumnWidth(3, 120); // 1st Call Column width in pixels
  sheet.setColumnWidth(4, 120); // 2nd Call Column width in pixels
  
  // Step 9: Center align all text columns nicely across the dataset grid
  sheet.getRange(1, 1, daysInMonth + 4, 4).setHorizontalAlignment("center");
  SpreadsheetApp.flush();

  // Trigger the next automated routing function to collect resident roster metrics
  setupResidentsAndColors(listType);
}