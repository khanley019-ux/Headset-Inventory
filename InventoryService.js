/**
 * Returns a headset by Asset ID.
 */
function getHeadsetByAssetId(assetId) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.INVENTORY);

  const lastRow = sheet.getLastRow();

  if (lastRow < 3) return null;

  const data = sheet.getRange(3, 2, lastRow - 2, 14).getValues();

  assetId = String(assetId).trim();

  for (let i = 0; i < data.length; i++) {

    const currentAssetId = String(data[i][0]).trim();

    Logger.log("Searching for: " + assetId);
Logger.log("Current Row Asset ID: " + currentAssetId);

if (currentAssetId == assetId) {

      return {
        assetId: data[i][0],
        brand: data[i][1],
        model: data[i][2],
        serialNumber: data[i][3],
        status: data[i][4],
        condition: data[i][5],
        lastUser: data[i][6],
        assignedTo: data[i][7],
        dateIssued: data[i][8],
        dateReturned: data[i][9],
        purchaseDate: data[i][10],
        purchasePrice: data[i][11],
        ageMonths: data[i][12],
        depreciatedValue: data[i][13]
      };

    }

  }

  Logger.log("Asset not found: " + assetId);

  return null;

}