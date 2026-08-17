/**
 * Returns a headset by Asset ID.
 */
function getHeadsetByAssetId(assetId) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.SHEETS.INVENTORY
    );


  // =========================================
  // SHEET CHECK
  // =========================================

  if (!sheet) {

    throw new Error(
      "Inventory sheet was not found."
    );

  }


  // =========================================
  // NORMALIZE ASSET ID
  // =========================================

  const searchAssetId =
    String(assetId || "")
      .trim()
      .toUpperCase();


  if (!searchAssetId) {

    return null;

  }


  // =========================================
  // GET LAST ROW
  // =========================================

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 3) {

    return null;

  }


  // =========================================
  // READ INVENTORY
  //
  // B:O
  // B = Asset ID
  // C = Brand
  // D = Model
  // E = Serial Number
  // F = Status
  // G = Condition
  // H = Last User
  // I = Assigned To
  // J = Date Issued
  // K = Date Returned
  // L = Purchase Date
  // M = Purchase Price
  // N = Age Months
  // O = Depreciated Value
  // =========================================

  const data =
    sheet
      .getRange(
        3,
        2,
        lastRow - 2,
        14
      )
      .getDisplayValues();


  // =========================================
  // SEARCH
  // =========================================

  for (
    let i = 0;
    i < data.length;
    i++
  ) {

    const currentAssetId =
      String(data[i][0] || "")
        .trim()
        .toUpperCase();


    console.log(
      "Searching: " +
      searchAssetId +
      " | Current: " +
      currentAssetId
    );


    if (
      currentAssetId ===
      searchAssetId
    ) {

      return {

        assetId:
          data[i][0],

        brand:
          data[i][1],

        model:
          data[i][2],

        serialNumber:
          data[i][3],

        status:
          data[i][4],

        condition:
          data[i][5],

        lastUser:
          data[i][6],

        assignedTo:
          data[i][7],

        dateIssued:
          data[i][8],

        dateReturned:
          data[i][9],

        purchaseDate:
          data[i][10],

        purchasePrice:
          data[i][11],

        ageMonths:
          data[i][12],

        depreciatedValue:
          data[i][13]

      };

    }

  }


  // =========================================
  // NOT FOUND
  // =========================================

  console.log(
    "Asset not found: " +
    searchAssetId
  );


  return null;

}