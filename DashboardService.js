/**
 * ============================================
 * DASHBOARD SERVICE
 * ============================================
 *
 * Reads the Headset Master Inventory and logs
 * to provide dashboard statistics.
 */


/**
 * Get dashboard statistics.
 */
function getDashboardStats() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const inventorySheet =
    ss.getSheetByName(
      CONFIG.SHEETS.INVENTORY
    );


  if (!inventorySheet) {
    throw new Error(
      "Headset Master Inventory sheet not found."
    );
  }


  const stats = {

    total: 0,

    available: 0,

    issued: 0,

    borrowed: 0,

    damaged: 0,

    repair: 0,

    lost: 0,

    issuedToday: 0,

    returnedToday: 0

  };


  /*
   * ==========================================
   * MASTER INVENTORY
   * ==========================================
   *
   * Data starts at Row 3.
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

  const lastRow =
    inventorySheet.getLastRow();


  if (lastRow >= 3) {

    const data =
      inventorySheet
        .getRange(
          3,
          2,
          lastRow - 2,
          14
        )
        .getValues();


    data.forEach(function(row) {

      const assetId =
        String(row[0] || "").trim();


      /*
       * Ignore empty inventory rows.
       */

      if (!assetId) {
        return;
      }


      stats.total++;


      const status =
        String(row[4] || "")
          .trim()
          .toLowerCase();


      switch (status) {

        case "available":

          stats.available++;

          break;


        case "assigned":

        case "issued":

          stats.issued++;

          break;


        case "borrowed":

          stats.borrowed++;

          break;


        case "damaged":

          stats.damaged++;

          break;


        case "under repair":

        case "repair":

          stats.repair++;

          break;


        case "lost":

          stats.lost++;

          break;

      }

    });

  }


  /*
   * ==========================================
   * TODAY
   * ==========================================
   */

  const today =
    new Date();


  today.setHours(
    0,
    0,
    0,
    0
  );


  /*
   * ==========================================
   * ISSUANCE LOGS
   * ==========================================
   *
   * B = Date
   * C = Agent Name
   * D = Asset ID
   * E = Headset Brand
   * F = Issued By
   * G = Status
   *
   * Data starts Row 2.
   */

  const issueSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.ISSUE
    );


  if (issueSheet) {

    const issueLastRow =
      issueSheet.getLastRow();


    if (issueLastRow >= 2) {

      const issueData =
        issueSheet
          .getRange(
            2,
            2,
            issueLastRow - 1,
            6
          )
          .getValues();


      issueData.forEach(function(row) {

        const date =
          row[0];


        if (
          date instanceof Date
        ) {

          const logDate =
            new Date(date);


          logDate.setHours(
            0,
            0,
            0,
            0
          );


          if (
            logDate.getTime() ===
            today.getTime()
          ) {

            stats.issuedToday++;

          }

        }

      });

    }

  }


  /*
   * ==========================================
   * RETURN LOGS
   * ==========================================
   *
   * B = Date Returned
   * C = Agent Name
   * D = Asset ID
   * E = Headset Brand
   * F = Status Upon Return
   * G = Checked By
   * H = Remarks
   *
   * Data starts Row 3.
   */

  const returnSheet =
    ss.getSheetByName(
      CONFIG.SHEETS.RETURN
    );


  if (returnSheet) {

    const returnLastRow =
      returnSheet.getLastRow();


    if (returnLastRow >= 3) {

      const returnData =
        returnSheet
          .getRange(
            3,
            2,
            returnLastRow - 2,
            7
          )
          .getValues();


      returnData.forEach(function(row) {

        const date =
          row[0];


        if (
          date instanceof Date
        ) {

          const logDate =
            new Date(date);


          logDate.setHours(
            0,
            0,
            0,
            0
          );


          if (
            logDate.getTime() ===
            today.getTime()
          ) {

            stats.returnedToday++;

          }

        }

      });

    }

  }


  return stats;

}