function handleDashboardSyncEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== SETTINGS.dashboardSheet) return;

  const editedA1 = e.range.getA1Notation();
  const actualCache = findLabelCell_(sheet, "Actual Cache Value");
  const totalRevenue = findLabelCell_(sheet, "Total Revenue");

  if (actualCache && editedA1 === actualCache.valueRange.getA1Notation()) {
    if (totalRevenue) totalRevenue.valueRange.setValue(e.value === undefined ? "" : e.value);
    return;
  }

  if (totalRevenue && editedA1 === totalRevenue.valueRange.getA1Notation()) {
    if (actualCache) actualCache.valueRange.setValue(e.value === undefined ? "" : e.value);
  }
}
