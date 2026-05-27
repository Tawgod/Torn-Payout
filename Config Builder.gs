/**
 * Builds the Config tab if it doesn't exist.
 * Ensures the structure matches the repository requirements.
 */
function buildConfigTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheetName = (typeof SETTINGS !== 'undefined' && SETTINGS.configSheet) ? SETTINGS.configSheet : "Config";
  
  let configSheet = ss.getSheetByName(configSheetName);
  
  // If the sheet already exists, we do nothing to avoid overwriting your current API Keys/IDs.
  if (configSheet) {
    ss.toast("Config tab already exists. Skipping build.", "Setup", 3);
    return;
  }
  
  // Create the sheet
  configSheet = ss.insertSheet(configSheetName);
  
  // Define the exact layout found in your Config.csv
  const configData = [
    ["API Key", ""],             // Cell B1
    ["Attack Data", ""],         // Cell B2
    ["Archive Sheet ID", ""],    // Cell B3
    ["Target Faction ID", ""],   // Cell B4
    ["Official War Start", ""],  // Cell B5
    ["Internal War ID", ""],     // Cell B6
    ["Faction ID", ""],          // Cell B7
    ["Public sheet ID", ""]      // Cell B8
  ];
  
  // Inject labels and values
  configSheet.getRange(1, 1, configData.length, 2).setValues(configData);
  
  // Styling to match your repository aesthetics
  configSheet.getRange(1, 1, configData.length, 1).setFontWeight("bold").setBackground("#f3f3f3");
  configSheet.setColumnWidth(1, 150);
  configSheet.setColumnWidth(2, 350);
  
  ss.toast("Config tab successfully created!", "Setup", 5);
}