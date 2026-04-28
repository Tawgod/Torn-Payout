function onEdit(e) {
  // If run manually from the editor, exit immediately
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  
  // Only monitor the Dashboard tab
  if (sheet.getName() !== SETTINGS.dashboardSheet) return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  let value = e.value;
  
  // If the cell was cleared, treat the value as 0
  if (value === undefined || value === "") {
    value = "";
  }
  
  // 1. If you type into 'Actual Cache Value' (Row 12, Col 3)
  if (row === 12 && col === 3) {
    // Copy the exact number over to 'Total Revenue' (Row 3, Col 12)
    sheet.getRange("L3").setValue(value);
  }
  
  // 2. If you type into 'Total Revenue' (Row 3, Col 12)
  if (row === 3 && col === 12) {
    // Copy the exact number over to 'Actual Cache Value' (Row 12, Col 3)
    sheet.getRange("C12").setValue(value);
  }
}