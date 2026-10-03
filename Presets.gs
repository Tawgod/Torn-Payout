const PAYOUT_PRESETS = {
  "40664": {
    defaultCut: 0.06,
    defaultMaxCut: "",
    presets: {
      "Termed Win":  { cut: 0.06, maxCut: "", weights: [1,0.5,0,0,0.5,"",0.2,0,0.2,0,0.5,0,0] },
      "Termed Loss": { cut: 0.06, maxCut: "", weights: [1,0.5,0,0,0.5,"",0.2,0,0.2,0,0.5,0,0] },
      "Real War":    { cut: 0.06, maxCut: "", weights: [1,0.5,0,0,0.5,"",0.2,0,0.2,0,0.5,0,0] }
    }
  },
  "46442": {
    defaultCut: 0.20,
    defaultMaxCut: 1000000000,
    presets: {
      "Termed Win":  { cut: 0.20, maxCut: 1000000000, weights: [1,0.5,0,0,0.5,"",0.2,0,0.2,0,0.5,0,0] },
      "Termed Loss": { cut: 0.20, maxCut: 1000000000, weights: [1,0.5,0,0,0.5,"",0.2,0,0.2,0,0.5,0,0] },
      "Real War":    { cut: 0.20, maxCut: 1000000000, weights: [1,0.5,0,0,0.5,"",0.2,0,0.2,0,0.5,0,0] }
    }
  }
};

function getFactionIdForPreset_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName((typeof SETTINGS !== "undefined" && SETTINGS.configSheet) ? SETTINGS.configSheet : "Config");
  if (!config) return "";
  return String(config.getRange((typeof SETTINGS !== "undefined" && SETTINGS.factionIdCell) ? SETTINGS.factionIdCell : "B7").getValue() || "").trim();
}

function getPresetConfig_() {
  const factionId = getFactionIdForPreset_();
  return PAYOUT_PRESETS[factionId] || PAYOUT_PRESETS["40664"];
}

function applySelectedPayoutPreset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName((typeof SETTINGS !== "undefined" && SETTINGS.dashboardSheet) ? SETTINGS.dashboardSheet : "Dashboard");
  const payouts = ss.getSheetByName((typeof SETTINGS !== "undefined" && SETTINGS.payoutSheet) ? SETTINGS.payoutSheet : "Payouts");
  if (!dash || !payouts) throw new Error("Dashboard and Payouts tabs are required.");

  const data = dash.getDataRange().getValues();
  let presetName = "";
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length - 1; c++) {
      if (String(data[r][c] || "").trim().toLowerCase() === "payout preset") {
        presetName = String(data[r][c + 1] || "").trim();
        break;
      }
    }
    if (presetName) break;
  }

  if (!presetName || presetName === "Custom") {
    ss.toast("Custom preset selected. Existing values were left unchanged.", "Payout Preset", 4);
    return;
  }

  const cfg = getPresetConfig_();
  const preset = cfg.presets[presetName];
  if (!preset) throw new Error("Unknown payout preset: " + presetName);

  if (!setDashboardValueByLabel_(dash, "Faction Cut %", preset.cut)) {
    throw new Error("Faction Cut % field was not found on the Dashboard.");
  }
  if (!setDashboardValueByLabel_(dash, "Max Faction Cut ($)", preset.maxCut)) {
    throw new Error("Max Faction Cut ($) field was not found on the Dashboard.");
  }

  const weights = preset.weights.slice(0, 13);
  while (weights.length < 13) weights.push(0);
  payouts.getRange(1, 5, 1, 13).setValues([weights]);

  SpreadsheetApp.flush();
  if (typeof runPayoutMath === "function") runPayoutMath();
  if (typeof refreshDashboard === "function") refreshDashboard();

  ss.toast(
    presetName + " applied: " + Math.round(Number(preset.cut || 0) * 100) + "% faction cut" +
      (preset.maxCut ? ", $" + Number(preset.maxCut).toLocaleString() + " max cut" : ""),
    "Payout Preset",
    5
  );
}

function onEditPayoutPreset_(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== ((typeof SETTINGS !== "undefined" && SETTINGS.dashboardSheet) ? SETTINGS.dashboardSheet : "Dashboard")) return;

  const labelCell = sheet.getRange(e.range.getRow(), Math.max(1, e.range.getColumn() - 1)).getDisplayValue();
  if (String(labelCell || "").trim().toLowerCase() !== "payout preset") return;

  try {
    applySelectedPayoutPreset();
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Preset error: " + err.message, "Payout Preset", 6);
  }
}

function onEdit(e) {
  onEditPayoutPreset_(e);
}
