function checkWarEnd() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(SETTINGS.dashboardSheet);
  
  if (!dashSheet) return;
  
  const warId = labelValue_(dashSheet, "War ID", "") || labelValue_(dashSheet, "War Report ID", "");

  if (!warId || warId === "No Data" || warId === "") return;

  const currentResult = labelValue_(dashSheet, "Outcome (Result)", "");
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

  // Write named Dashboard fields so layout changes cannot redirect the results.
  setLabelValue_(dashSheet, "Outcome (Result)", resultText);
  setLabelValue_(dashSheet, "Caches / Items Won", cacheString);
  setLabelValue_(dashSheet, "Est. Cache Value", totalEstValue);
  setLabelValue_(dashSheet, "Actual Cache Value", totalEstValue);
  
  ss.toast(`War Ended! Result: ${resultText}.`, "System", 8);
}