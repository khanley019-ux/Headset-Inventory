/**
 * =========================================
 * BORROW HEADSET
 * =========================================
 *
 * Borrow Logs structure:
 *
 * B = Date
 * C = Agent Name
 * D = Asset ID
 * E = Headset Brand
 * F = Issued By
 * G = Status
 *
 * Data starts at Row 2
 */


/**
 * Get headsets that can currently be borrowed.
 *
 * We use the Headset Master Inventory
 * as the source of available headsets.
 */
function getBorrowableHeadsets() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.INVENTORY);

  if (!sheet) {
    throw new Error("Headset Master Inventory sheet not found.");
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    return [];
  }

  // Inventory starts at Row 3, Columns B:O
  const data = sheet
    .getRange(3, 2, lastRow - 2, 14)
    .getValues();

  return data
    .filter(function(row) {

      const assetId = String(row[0] || "").trim();
      const status = String(row[4] || "").trim().toLowerCase();

      return (
        assetId !== "" &&
        status === "available"
      );

    })
    .map(function(row) {

      return {

        assetId: row[0],
        brand: row[1],
        model: row[2],
        serialNumber: row[3],
        status: row[4],
        condition: row[5]

      };

    });

}


/**
 * Save a new Borrow transaction.
 *
 * Writes to:
 *
 * B = Date
 * C = Agent Name
 * D = Asset ID
 * E = Headset Brand
 * F = Issued By
 * G = Status
 */
function borrowHeadset(data) {

  if (!data) {
    throw new Error("Borrow data is missing.");
  }

  if (!data.agentName) {
    throw new Error("Agent Name is required.");
  }

  if (!data.assetId) {
    throw new Error("Asset ID is required.");
  }

  if (!data.brand) {
    throw new Error("Headset Brand is required.");
  }


  // =========================================
  // GET BORROW LOGS SHEET
  // =========================================

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.BORROW);


  if (!sheet) {
    throw new Error("Borrow Logs sheet not found.");
  }


  // =========================================
  // GET LOGGED-IN IT USER
  // =========================================

  const currentUser =
    data.currentITUser || {};

  const itUsername =
    currentUser.username || "";

  const itFullName =
    currentUser.fullName ||
    data.issuedBy ||
    "Unknown IT Staff";

  const itPosition =
    currentUser.position ||
    "IT Staff";


  // =========================================
  // FIND NEXT ROW
  // =========================================

  const row =
    getFirstEmptyRow(sheet, 2);


  // =========================================
  // SAVE BORROW TRANSACTION
  // =========================================

  // B - Date
  sheet
    .getRange(row, 2)
    .setValue(new Date());


  // C - Agent Name
  sheet
    .getRange(row, 3)
    .setValue(data.agentName);


  // D - Asset ID
  sheet
    .getRange(row, 4)
    .setValue(data.assetId);


  // E - Headset Brand
  sheet
    .getRange(row, 5)
    .setValue(data.brand);


  // F - Issued By
  sheet
    .getRange(row, 6)
    .setValue(itFullName);


  // G - Status
  sheet
    .getRange(row, 7)
    .setValue("Borrowed");


  SpreadsheetApp.flush();


  // =========================================
  // ACTIVITY LOG
  // =========================================

  logActivity({

    itUsername:
      itUsername,

    itFullName:
      itFullName,

    itPosition:
      itPosition,

    action:
      "BORROWED HEADSET",

    assetId:
      data.assetId,

    employee:
      data.agentName,

    description:
      "Borrowed headset " +
      data.assetId +
      " by " +
      data.agentName

  });


  // =========================================
  // RETURN RESULT
  // =========================================

  return {

    success: true,

    row: row

  };

}