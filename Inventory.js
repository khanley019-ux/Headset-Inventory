function getInventory() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.INVENTORY);

  const data = sheet.getDataRange().getValues();

  data.shift(); // Remove header row

  return data.map(row => ({
    assetId: row[1],       // Column B
    brand: row[2],         // Column C
    model: row[3],         // Column D
    serialNumber: row[4],  // Column E
    lastUser: row[7],      // Column H
    status: row[5]         // Column F
  })).filter(item => item.assetId !== "");

}