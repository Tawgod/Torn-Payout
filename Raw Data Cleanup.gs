// ==========================================
// AUTOMATED RAW DATA CLEANUP (CAPACITY-BASED)
// ==========================================
function autoPruneData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Data"); 
  
  if (!dataSheet) {
    console.error("Error: 'Data' sheet not found.");
    return;
  }

  // --- GOOGLE SHEETS LIMITS & MARGINS ---
  const MAX_CELL_LIMIT = 10000000; // Hard Google Sheets Limit
  const TRIGGER_CAPACITY = 0.90;   // Trigger cleanup at 90% (Leaves 10% margin of error)
  const TARGET_CAPACITY = 0.80;    // Clean down to 80% capacity

  const numCols = dataSheet.getMaxColumns(); // Dynamically checks columns (18 if A-R)
  const currentTotalRows = dataSheet.getMaxRows(); // Total physical rows counting towards the limit
  const currentTotalCells = currentTotalRows * numCols; 
  
  // Calculate specific thresholds based on current column count
  const triggerCellCount = MAX_CELL_LIMIT * TRIGGER_CAPACITY; // 9,000,000 cells
  const targetCellCount = MAX_CELL_LIMIT * TARGET_CAPACITY;   // 8,000,000 cells

  if (currentTotalCells >= triggerCellCount) {
    // Calculate how many rows equal 8,000,000 cells
    const targetRowCount = Math.floor(targetCellCount / numCols); 
    const rowsToDelete = currentTotalRows - targetRowCount;
    
    // Safety check to ensure we don't accidentally wipe the header (Row 1)
    if (rowsToDelete > 0 && currentTotalRows > rowsToDelete + 1) {
      // Delete oldest entries starting at row 2
      dataSheet.deleteRows(2, rowsToDelete);
      
      console.log(`⚠️ Cleanup Triggered: Sheet reached ${currentTotalCells} cells.`);
      console.log(`✅ Success: Deleted the oldest ${rowsToDelete} rows. Sheet is now back to ${targetRowCount} rows (~80% capacity).`);
    } else {
      console.error("Error: Deletion math attempted to delete too many rows. Check column count.");
    }
  } else {
    // Sheet is still healthy
    const percentUsed = ((currentTotalCells / MAX_CELL_LIMIT) * 100).toFixed(2);
    console.log(`Sheet healthy: Currently at ${percentUsed}% capacity (${currentTotalCells} / 10,000,000 cells). Trigger is at 90%.`);
  }
}