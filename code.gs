const API_KEY = "Your API_Key";
const API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=" + API_KEY;

// This creates the menu at the top of your sheet
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 AI Agent')
    .addItem('Formula Assistant', 'showPrompt')
    .addItem('Suggest Dashboard', 'suggestDashboard')
    .addItem('Create Pivot Dashboard', 'createPivotDashboard') // Add this line!
    .addToUi();
}

// This opens a popup box to ask your question
function showPrompt() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Get the request
  const response = ui.prompt('AI Formula Assistant', 'What formula do you need?', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const userRequest = response.getResponseText();

  // 2. Get the target cell
  const cellResponse = ui.prompt('Target Cell', 'Which cell (e.g., A1)?', ui.ButtonSet.OK_CANCEL);
  if (cellResponse.getSelectedButton() !== ui.Button.OK) return;
  const targetCell = cellResponse.getResponseText();
  // Add this line inside showPrompt
     const headers = getHeaders();

  

  // 3. Create the "Strict" prompt
  const systemPrompt = `You are a Google Sheets expert. 
  Context: The spreadsheet has these headers: ${headers}.
- Employee id is in column A. 
- Names are in Column B.
- Departments are in Column C.
- Salaries are in Column D.
- Location is in column E
Rule: Return ONLY the Google Sheets formula starting with =. No explanations.
Request: ${userRequest}`;

  
  const aiResponse = callGemini(systemPrompt);

  // 4. Write to the sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  try {
    sheet.getRange(targetCell).setFormula(aiResponse);
    ui.alert('Success!', 'Formula written: ' + aiResponse, ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('Error', 'Check if the cell address is correct: ' + e.message, ui.ButtonSet.OK);
  }
}

// Our core AI logic (now called by the popup)
function callGemini(prompt) {
  const payload = { "contents": [{"parts": [{"text": prompt}]}] };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true 
  };

  const response = UrlFetchApp.fetch(API_URL, options);
  const data = JSON.parse(response.getContentText());
  return data.candidates[0].content.parts[0].text;
}

function getHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // This looks at Row 1, Column 1, and grabs everything to the right
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headerValues = headerRange.getValues()[0]; // [0] gets the first (and only) row
  
  return headerValues.join(", "); // Returns them as a simple list: "Date, Sales, Region"
}

function suggestDashboard() {
  const ui = SpreadsheetApp.getUi();
  const headers = getHeaders();
  
  const prompt = "I have a spreadsheet with these headers: " + headers + 
                 ". Suggest 3 specific charts or summary tables for a dashboard. " +
                 "Be brief and technical.";
  
  const suggestion = callGemini(prompt);
  ui.alert('📊 Dashboard Ideas', suggestion, ui.ButtonSet.OK);
}

// Don't forget to add it to your menu in onOpen()!
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 AI Agent')
    .addItem('Formula Assistant', 'showPrompt')
    .addItem('Suggest Dashboard', 'suggestDashboard') // New item!
    .addItem('Create Dept Summary', 'createDepartmentSummary') // Add this!
    .addItem('Create Pivot Dashboard', 'createPivotDashboard') // Add this line!
    .addToUi();
}

/**
 * Grabs the headers from Row 1 of the active sheet.
 * @return {string} A comma-separated list of column names.
 */
function getHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get all data in the first row from column A to the last used column
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return "No headers found";
  
  const headerValues = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  
  // Filter out any empty cells and join them with commas
  return headerValues.filter(String).join(", ");
}

function createDepartmentSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  // 1. Get all departments from Column C (starting at row 2)
  const deptRange = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  
  // 2. Find unique departments using a "Set"
  const uniqueDepts = [...new Set(deptRange.flat())].filter(String);
  
  // 3. Set headers for our summary table in F1 and G1
  sheet.getRange("F1").setValue("Department");
  sheet.getRange("G1").setValue("Total Salary");
  
  // 4. Loop through unique departments and write the SUMIF formula
  for (let i = 0; i < uniqueDepts.length; i++) {
    const deptName = uniqueDepts[i];
    const row = i + 2;
    
    sheet.getRange(row, 6).setValue(deptName); // Write Dept Name in Col F
    
    // Write the SUMIF formula in Col G
    // It says: Look in C:C for the name in F, then sum D:D
    sheet.getRange(row, 7).setFormula(`=SUMIF(C:C, F${row}, D:D)`);
  }
  
  SpreadsheetApp.getUi().alert('Summary Table Created in Columns F & G!');
}

function createPivotDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  const sourceData = sourceSheet.getRange("B1:D" + sourceSheet.getLastRow());
  
  // 1. Create a new sheet for the dashboard
  let dashSheet = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard");
  dashSheet.clear(); // Clean start
  
  // 2. Add the Pivot Table
  const pivotTable = dashSheet.getRange("A1").createPivotTable(sourceData);
  
  // Add Departments as Rows (Column C is index 2 relative to B:D)
  pivotTable.addRowGroup(3); 
  
  // Add Salaries as Values (Column D is index 3 relative to B:D)
  pivotTable.addPivotValue(4, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  
  // 3. Create a Bar Chart from the Pivot Table
  const chart = dashSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(dashSheet.getRange("A1:B10")) // Targets the pivot results
    .setPosition(2, 4, 0, 0)
    .setOption('title', 'Total Salary Spend by Department')
    .build();
    
  dashSheet.insertChart(chart);
  SpreadsheetApp.getUi().alert("Dashboard created in the 'Dashboard' tab! 🚀");
}
