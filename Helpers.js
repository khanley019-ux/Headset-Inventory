/**
 * Returns the first empty data row.
 * Data starts at Row 3.
 */
function getFirstEmptyRow(sheet, column) {

  const startRow = 3;

  const values = sheet
    .getRange(startRow, column, sheet.getMaxRows() - startRow + 1, 1)
    .getValues();

  for (let i = 0; i < values.length; i++) {

    if (values[i][0] === "") {
      return startRow + i;
    }

  }

  return sheet.getLastRow() + 1;

}

/**
 * Returns the next Asset ID.
 * Data starts at Row 3.
 */
function getNextAssetId() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.INVENTORY);

  const lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    return "HS-0001";
  }

  const assetIds = sheet
    .getRange(3, 2, lastRow - 2, 1)
    .getValues()
    .flat();

  let highest = 0;

  assetIds.forEach(function(id) {

    if (!id) return;

    const match = id.toString().match(/^HS-(\d+)$/);

    if (match) {
      highest = Math.max(highest, parseInt(match[1], 10));
    }

  });

  return "HS-" + String(highest + 1).padStart(4, "0");

}


/**
 * Returns all available headsets.
 * Data starts at Row 3.
 */
function getAvailableHeadsets() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.INVENTORY);

  const lastRow = sheet.getLastRow();

  if (lastRow < 3) return [];

  const data = sheet.getRange(3, 2, lastRow - 2, 5).getValues();

  return data
    .filter(row => row[4] === "Available")
    .map(row => ({
      assetId: row[0],
      brand: row[1]
    }));

}