function saveHeadset(headset) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.INVENTORY);


  // =========================================
  // FIND NEXT EMPTY INVENTORY ROW
  // =========================================

  const row = getFirstEmptyRow(sheet, 2);


  // =========================================
  // SAVE HEADSET
  // =========================================

  sheet.getRange(row, 2).setValue(headset.assetId);          // B - Asset ID
  sheet.getRange(row, 3).setValue(headset.brand);            // C - Brand
  sheet.getRange(row, 4).setValue(headset.model);            // D - Model
  sheet.getRange(row, 5).setValue(headset.serialNumber);     // E - Serial Number
  sheet.getRange(row, 7).setValue(headset.condition);         // G - Condition
  sheet.getRange(row, 12).setValue(headset.purchaseDate);    // L - Purchase Date
  sheet.getRange(row, 13).setValue(headset.purchasePrice);   // M - Purchase Price


  SpreadsheetApp.flush();


  // =========================================
  // ACTIVITY LOG
  // =========================================
  //
  // IMPORTANT:
  // The frontend can send the logged-in IT
  // user inside headset.currentITUser.
  //

  const user = headset.currentITUser || {};


  saveActivityLog({

    user: {
      username: user.username || "",
      fullName: user.fullName || "",
      position: user.position || ""
    },

    action: "ADDED HEADSET",

    assetId: headset.assetId || "",

    employee: "",

    description:
      "Added " +
      (headset.brand || "") +
      " " +
      (headset.model || "") +
      " to inventory."

  });


  // =========================================
  // SUCCESS
  // =========================================

  return true;

}

function getInventory() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.INVENTORY);

  const lastRow = sheet.getLastRow();

  if (lastRow < 3) return [];

  // Get all data from Row 3 onward
  const data = sheet.getRange(3, 2, lastRow - 2, 14).getDisplayValues();

  const inventory = [];

  data.forEach(function(row) {

    const assetId = row[0].toString().trim();

    // Skip completely empty rows
    if (assetId === "") return;

    inventory.push({
      assetId: assetId,
      brand: row[1],
      model: row[2],
      serialNumber: row[3],
      status: row[4],
      condition: row[5],
      lastUser: row[6],
      assignedTo: row[7],
      dateIssued: row[8],
      dateReturned: row[9],
      purchaseDate: row[10],
      purchasePrice: row[11],
      ageMonths: row[12],
      depreciatedValue: row[13]
    });

  });

  return inventory;

}