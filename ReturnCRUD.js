/**
 * =========================================
 * RETURN HEADSET
 * =========================================
 *
 * Return Logs:
 *
 * B = Date Returned
 * C = Agent Name
 * D = Asset ID
 * E = Headset Brand
 * F = Status Upon Return
 * G = Checked By
 * H = Remarks
 *
 * Data starts at Row 3.
 */


/**
 * Find a headset by Asset ID for returning.
 *
 * A headset can be returned if its current
 * Master Inventory status is:
 *
 * - Borrowed
 * - Assigned
 *
 * For Borrowed:
 * Agent Name comes from Borrow Logs.
 *
 * For Assigned:
 * Agent Name comes from Assigned To
 * in Master Inventory.
 */
function searchReturnableHeadset(assetId) {

  if (!assetId) {
    return null;
  }


  const inventorySheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.INVENTORY
      );


  if (!inventorySheet) {
    throw new Error(
      "Headset Master Inventory sheet not found."
    );
  }


  const lastRow =
    inventorySheet.getLastRow();


  if (lastRow < 3) {
    return null;
  }


  /*
   * Master Inventory:
   *
   * B:O
   *
   * B = Asset ID
   * C = Brand
   * D = Model
   * E = Serial Number
   * F = Status
   * G = Condition
   * H = Last User
   * I = Assigned To
   * J = Date Issued
   * K = Date Returned
   * L = Purchase Date
   * M = Purchase Price
   * N = Age Months
   * O = Depreciated Value
   */

  const data =
    inventorySheet
      .getRange(
        3,
        2,
        lastRow - 2,
        14
      )
      .getValues();


  const searchId =
    String(assetId)
      .trim()
      .toLowerCase();


  for (let i = 0; i < data.length; i++) {

    const row = data[i];


    const rowAssetId =
      String(row[0] || "")
        .trim();


    if (
      rowAssetId.toLowerCase() !== searchId
    ) {
      continue;
    }


    const status =
      String(row[4] || "")
        .trim();


    const normalizedStatus =
      status.toLowerCase();


    /*
     * Only Borrowed and Assigned
     * headsets can be returned.
     */

    if (
      normalizedStatus !== "borrowed" &&
      normalizedStatus !== "assigned"
    ) {

      return {

        found: false,

        message:
          "This headset cannot be returned because its current status is " +
          (status || "Unknown") +
          "."

      };

    }


    let agentName = "";


    /*
     * =========================================
     * BORROWED HEADSET
     * =========================================
     */

    if (normalizedStatus === "borrowed") {

      const borrowRecord =
        getActiveBorrowRecord(
          rowAssetId
        );


      if (borrowRecord) {

        agentName =
          borrowRecord.agentName || "";

      }

    }


    /*
     * =========================================
     * ASSIGNED HEADSET
     * =========================================
     */

    if (normalizedStatus === "assigned") {

      agentName =
        row[7] || ""; // I - Assigned To

    }


    /*
     * If Assigned To is empty, use Last User
     * as a fallback.
     */

    if (!agentName) {

      agentName =
        row[6] || ""; // H - Last User

    }


    return {

      found: true,

      assetId: row[0],

      brand: row[1],

      model: row[2],

      serialNumber: row[3],

      status: row[4],

      condition: row[5],

      lastUser: row[6],

      assignedTo: row[7],

      agentName: agentName

    };

  }


  return {

    found: false,

    message:
      "Headset with Asset ID " +
      assetId +
      " was not found."

  };

}


/**
 * Get the active Borrow Logs record.
 *
 * Borrow Logs:
 *
 * B = Date
 * C = Agent Name
 * D = Asset ID
 * E = Headset Brand
 * F = Issued By
 * G = Status
 *
 * Data starts at Row 2.
 */
function getActiveBorrowRecord(assetId) {

  if (!assetId) {
    return null;
  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.BORROW
      );


  if (!sheet) {
    throw new Error(
      "Borrow Logs sheet not found."
    );
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return null;
  }


  const data =
    sheet
      .getRange(
        2,
        2,
        lastRow - 1,
        6
      )
      .getValues();


  const searchId =
    String(assetId)
      .trim()
      .toLowerCase();


  for (let i = 0; i < data.length; i++) {

    const row = data[i];


    const rowAssetId =
      String(row[2] || "")
        .trim()
        .toLowerCase();


    const status =
      String(row[5] || "")
        .trim()
        .toLowerCase();


    if (
      rowAssetId === searchId &&
      status === "borrowed"
    ) {

      return {

        row: i + 2,

        date: row[0],

        agentName: row[1],

        assetId: row[2],

        brand: row[3],

        issuedBy: row[4],

        status: row[5]

      };

    }

  }


  return null;

}


/**
 * Save a Return transaction.
 *
 * Return Logs:
 *
 * B = Date Returned
 * C = Agent Name
 * D = Asset ID
 * E = Headset Brand
 * F = Status Upon Return
 * G = Checked By
 * H = Remarks
 *
 * Data starts at Row 3.
 */
function returnHeadset(data) {

  if (!data) {
    throw new Error(
      "Return data is missing."
    );
  }


  if (!data.assetId) {
    throw new Error(
      "Asset ID is required."
    );
  }


  if (!data.statusUponReturn) {
    throw new Error(
      "Status Upon Return is required."
    );
  }


  if (!data.checkedBy) {
    throw new Error(
      "Checked By is required."
    );
  }


  const allowedStatuses = [
    "Good",
    "New",
    "Fair",
    "Damaged",
    "Lost"
  ];


  if (
    allowedStatuses.indexOf(
      data.statusUponReturn
    ) === -1
  ) {

    throw new Error(
      "Invalid Status Upon Return."
    );

  }


  /*
   * Find the headset in Master Inventory.
   */

  const headset =
    searchReturnableHeadset(
      data.assetId
    );


  if (
    !headset ||
    !headset.found
  ) {

    throw new Error(
      headset &&
      headset.message
        ? headset.message
        : "Headset could not be found."
    );

  }


  /*
   * =========================================
   * RETURN LOGS
   * =========================================
   */

  const returnSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.RETURN
      );


  if (!returnSheet) {

    throw new Error(
      "Return Logs sheet not found."
    );

  }


  /*
   * Return Logs starts at Row 3.
   */

  const returnRow =
    getFirstEmptyRow(
      returnSheet,
      2,
      3
    );


  /*
   * B - Date Returned
   */

  returnSheet
    .getRange(
      returnRow,
      2
    )
    .setValue(
      new Date()
    );


  /*
   * C - Agent Name
   */

  returnSheet
    .getRange(
      returnRow,
      3
    )
    .setValue(
      headset.agentName || ""
    );


  /*
   * D - Asset ID
   */

  returnSheet
    .getRange(
      returnRow,
      4
    )
    .setValue(
      headset.assetId
    );


  /*
   * E - Headset Brand
   */

  returnSheet
    .getRange(
      returnRow,
      5
    )
    .setValue(
      headset.brand
    );


  /*
   * F - Status Upon Return
   */

  returnSheet
    .getRange(
      returnRow,
      6
    )
    .setValue(
      data.statusUponReturn
    );


  /*
   * G - Checked By
   */

  returnSheet
    .getRange(
      returnRow,
      7
    )
    .setValue(
      data.checkedBy
    );


  /*
   * H - Remarks
   */

  returnSheet
    .getRange(
      returnRow,
      8
    )
    .setValue(
      data.remarks || ""
    );


  /*
   * =========================================
   * BORROWED HEADSET
   * =========================================
   *
   * If this headset came from Borrow Logs,
   * change:
   *
   * Borrowed → Returned
   */

  if (
    String(headset.status)
      .trim()
      .toLowerCase() === "borrowed"
  ) {

    const borrowRecord =
      getActiveBorrowRecord(
        headset.assetId
      );


    if (borrowRecord) {

      const borrowSheet =
        SpreadsheetApp
          .getActiveSpreadsheet()
          .getSheetByName(
            CONFIG.SHEETS.BORROW
          );


      if (borrowSheet) {

        borrowSheet
          .getRange(
            borrowRecord.row,
            7
          )
          .setValue(
            "Returned"
          );

      }

    }

  }


  SpreadsheetApp.flush();


  return {

    success: true,

    returnRow: returnRow,

    assetId: headset.assetId,

    previousStatus: headset.status,

    agentName: headset.agentName

  };

}