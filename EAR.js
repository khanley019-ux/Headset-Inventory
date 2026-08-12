/**
 * ============================================
 * EAR - EQUIPMENT ACKNOWLEDGEMENT RECEIPT
 * ============================================
 *
 * Retrieves the information needed to generate
 * an EAR for a specific headset.
 *
 * The EAR is NOT generated automatically.
 * Data is retrieved only when the user clicks
 * the Generate EAR button.
 */

function getEARData(assetId) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  assetId = String(assetId || "").trim();

  if (!assetId) {
    throw new Error("Asset ID is required.");
  }

  // ==========================================
  // FIND HEADSET IN MASTER INVENTORY
  // ==========================================

  const inventorySheet =
    ss.getSheetByName("Headset Master Inventory");

  if (!inventorySheet) {
    throw new Error(
      'Sheet "Headset Master Inventory" was not found.'
    );
  }

  const inventoryLastRow =
    inventorySheet.getLastRow();

  if (inventoryLastRow < 3) {
    throw new Error(
      "No headset inventory data found."
    );
  }

  /*
   * Inventory:
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
   * N = Age
   * O = Depreciated Value
   */

  const inventoryData =
    inventorySheet
      .getRange(
        3,
        2,
        inventoryLastRow - 2,
        14
      )
      .getValues();

  let headset = null;

  inventoryData.forEach(function(row) {

    const rowAssetId =
      String(row[0] || "").trim();

    if (
      rowAssetId === assetId &&
      !headset
    ) {

      headset = {

        assetId: row[0],

        brand: row[1],

        model: row[2],

        serialNumber: row[3],

        status: row[4],

        condition: row[5],

        lastUser: row[6],

        assignedTo: row[7],

        dateIssued: row[8],

        dateReturned: row[9]

      };

    }

  });

  if (!headset) {

    throw new Error(
      "Headset " +
      assetId +
      " was not found in inventory."
    );

  }

  // ==========================================
  // FIND MOST RECENT ISSUANCE
  // ==========================================

  const issuanceSheet =
    ss.getSheetByName("Issuance Logs");

  if (!issuanceSheet) {

    throw new Error(
      'Sheet "Issuance Logs" was not found.'
    );

  }

  const issuanceLastRow =
    issuanceSheet.getLastRow();

  let latestIssuance = null;

  if (issuanceLastRow >= 2) {

    /*
     * Issuance Logs:
     *
     * B = Date
     * C = Agent Name
     * D = Asset ID
     * E = Headset Brand
     * F = Issued By
     * G = Status
     */

    const issuanceData =
      issuanceSheet
        .getRange(
          2,
          2,
          issuanceLastRow - 1,
          6
        )
        .getValues();

    issuanceData.forEach(function(row) {

      const date =
        row[0];

      const employee =
        row[1];

      const rowAssetId =
        String(row[2] || "").trim();

      const brand =
        row[3];

      const issuedBy =
        row[4];

      const status =
        String(row[5] || "")
          .trim()
          .toLowerCase();

      if (
        rowAssetId !== assetId
      ) {
        return;
      }

      const currentRecord = {

        date: date,

        employee: employee,

        assetId: rowAssetId,

        brand: brand,

        issuedBy: issuedBy,

        status: status

      };

      if (!latestIssuance) {

        latestIssuance =
          currentRecord;

        return;

      }

      const currentDate =
        date instanceof Date
          ? date.getTime()
          : new Date(date).getTime();

      const latestDate =
        latestIssuance.date instanceof Date
          ? latestIssuance.date.getTime()
          : new Date(
              latestIssuance.date
            ).getTime();

      /*
       * Prefer the currently assigned issuance.
       */

      if (
        status === "assigned" &&
        latestIssuance.status !== "assigned"
      ) {

        latestIssuance =
          currentRecord;

        return;

      }

      /*
       * Otherwise use the newest issuance.
       */

      if (
        !isNaN(currentDate) &&
        (
          isNaN(latestDate) ||
          currentDate > latestDate
        )
      ) {

        latestIssuance =
          currentRecord;

      }

    });

  }

  if (!latestIssuance) {

    throw new Error(
      "No issuance record was found for " +
      assetId +
      "."
    );

  }

  // ==========================================
  // DATE
  // ==========================================

  let issuedDate =
    latestIssuance.date ||
    headset.dateIssued ||
    new Date();

  if (issuedDate instanceof Date) {

    issuedDate =
      issuedDate.toISOString();

  }

  // ==========================================
  // BUILD EAR DATA
  // ==========================================

  const earData = {

    equipmentName: "Headset",

    assetId:
      String(
        headset.assetId ||
        assetId
      ),

    brand:
      String(
        headset.brand ||
        latestIssuance.brand ||
        ""
      ),

    model:
      String(
        headset.model ||
        ""
      ),

    serialNumber:
      String(
        headset.serialNumber ||
        ""
      ),

    condition:
      String(
        headset.condition ||
        ""
      ),

    // ========================================
    // RECEIVED BY
    // ========================================

    receivedName:
      String(
        latestIssuance.employee ||
        headset.assignedTo ||
        ""
      ),

    // Position intentionally blank
    receivedPosition: "",

    receivedDate:
      issuedDate,

    // ========================================
    // ISSUED BY
    // ========================================

    issuedBy:
      String(
        latestIssuance.issuedBy ||
        ""
      ),

    // Position intentionally blank
    issuedByPosition: "",

    // Same date as Received By
    issuedDate:
      issuedDate

  };

  console.log(
    "EAR DATA:",
    JSON.stringify(
      earData,
      null,
      2
    )
  );

  return earData;

}


/**
 * ============================================
 * TEST EAR DATA
 * ============================================
 */

function testGetEARData() {

  const assetId = "HS-0001";

  const result =
    getEARData(assetId);

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

}

/**
 * ============================================
 * DIPH LOGO
 * ============================================
 *
 * Gets the DIPH logo from Google Drive
 * and converts it to a Base64 image so
 * the EAR can print reliably.
 */

function getDIPHLogo() {

  const fileId =
    "1I-QV-4GbNlgJKUhS71ds-odh9tiZVO8c";

  const file =
    DriveApp.getFileById(fileId);

  const blob =
    file.getBlob();

  const contentType =
    blob.getContentType();

  const base64 =
    Utilities.base64Encode(
      blob.getBytes()
    );

  return (
    "data:" +
    contentType +
    ";base64," +
    base64
  );

}