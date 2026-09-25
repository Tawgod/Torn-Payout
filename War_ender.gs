function checkWarEnd() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  
  if (!dashSheet) return;
  
  const warId = dashSheet.getRange("C5").getValue(); // NOW READS FROM DASHBOARD

  if (!warId || warId === "No Data" || warId === "") return;

  const currentResult = dashSheet.getRange("C9").getValue();
  if (currentResult === "Victory" || currentResult === "Defeat") return;

  let warJson;
  try {
    warJson = tornApiRequest_("torn", warId, "rankedwars");
  } catch (e) { return; }

  if (warJson.error || !warJson.rankedwars || !warJson.rankedwars[warId]) return;

  const warData = warJson.rankedwars[warId];
  if (warData.war.end === 0) return; 
  
  let myFactionId;
  try {
    myFactionId = tornApiRequest_("user", "", "profile").faction.faction_id.toString();
  } catch(e) { return; }

  const winnerId = warData.war.winner.toString();
  const resultText = (winnerId === myFactionId) ? "Victory" : "Defeat";

  let cacheString = "None";
  let totalEstValue = 0;
  const myRewards = warData.factions[myFactionId]?.rewards?.items;

  if (myRewards && Object.keys(myRewards).length > 0) {
    const itemRes = tornApiRequest_("torn", "", "items");
    const allItems = itemRes.items || {};

    let cacheArr = [];
    for (let key in myRewards) {
      let rwItem = myRewards[key]; 
      cacheArr.push(`${rwItem.quantity}x ${rwItem.name}`);

      for (let itemId in allItems) {
        if (allItems[itemId].name === rwItem.name) {
          totalEstValue += (allItems[itemId].market_value * rwItem.quantity);
          break;
        }
      }
    }
    cacheString = cacheArr.join(", ");
  }

  // Write Results shifted down one row
  dashSheet.getRange("C9").setValue(resultText);
  dashSheet.getRange("C10").setValue(cacheString);
  dashSheet.getRange("C11").setValue(totalEstValue);
  
  ss.toast(`War Ended! Result: ${resultText}.`, "System", 8);
}