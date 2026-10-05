function normalizeHeader_(value) {
  return String(value === null || value === undefined ? "" : value).trim().toLowerCase();
}

function headerMapFromRow_(row) {
  const map = {};
  (row || []).forEach((value, index) => {
    const key = normalizeHeader_(value);
    if (key && map[key] === undefined) map[key] = index;
  });
  return map;
}

function headerIndex_(map, names, required) {
  const list = Array.isArray(names) ? names : [names];
  for (const name of list) {
    const key = normalizeHeader_(name);
    if (map[key] !== undefined) return map[key];
  }
  if (required) throw new Error("Required header not found: " + list.join(" / "));
  return -1;
}

function findLabelCell_(sheet, label) {
  if (!sheet) return null;
  const wanted = normalizeHeader_(label);
  const data = sheet.getDataRange().getValues();
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length - 1; c++) {
      if (normalizeHeader_(data[r][c]) === wanted) {
        return {
          labelRange: sheet.getRange(r + 1, c + 1),
          valueRange: sheet.getRange(r + 1, c + 2),
          value: data[r][c + 1]
        };
      }
    }
  }
  return null;
}

function labelValue_(sheet, label, fallback) {
  const hit = findLabelCell_(sheet, label);
  return hit ? hit.value : fallback;
}

function setLabelValue_(sheet, label, value) {
  const hit = findLabelCell_(sheet, label);
  if (!hit) return false;
  hit.valueRange.setValue(value === null || value === undefined ? "" : value);
  return true;
}


function findLabelCellInRange_(sheet, rangeA1, label) {
  if (!sheet) return null;
  const wanted = normalizeHeader_(label);
  const range = sheet.getRange(rangeA1);
  const data = range.getValues();
  const startRow = range.getRow();
  const startCol = range.getColumn();
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length - 1; c++) {
      if (normalizeHeader_(data[r][c]) === wanted) {
        return {
          labelRange: sheet.getRange(startRow + r, startCol + c),
          valueRange: sheet.getRange(startRow + r, startCol + c + 1),
          value: data[r][c + 1]
        };
      }
    }
  }
  return null;
}
