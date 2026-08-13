function saveIssueTransaction(issue) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.SHEETS.ISSUE
    );


  // =========================================
  // VALIDATE ISSUE DATA
  // =========================================

  if (!issue) {
    throw new Error(
      "Issue information is required."
    );
  }

  if (!issue.agentName) {
    throw new Error(
      "Agent Name is required."
    );
  }

  if (!issue.assetId) {
    throw new Error(
      "Asset ID is required."
    );
  }


  // =========================================
  // FIND HEADSET
  // =========================================

  const headset =
    getHeadsetByAssetId(
      issue.assetId
    );


  if (!headset) {

    throw new Error(
      "Headset not found."
    );

  }


  // =========================================
  // GET LOGGED-IN IT USER
  // =========================================

  const currentUser =
    issue.currentITUser || {};


  const itUsername =
    currentUser.username ||
    "";

  const itFullName =
    currentUser.fullName ||
    issue.issuedBy ||
    "Unknown IT Staff";

  const itPosition =
    currentUser.position ||
    "IT Staff";


  // =========================================
  // SAVE TO ISSUANCE LOGS
  // =========================================

  const row =
    getFirstEmptyRow(
      sheet,
      2
    );


  sheet
    .getRange(
      row,
      2,
      1,
      6
    )
    .setValues([[
      new Date(),

      issue.agentName,

      headset.assetId,

      headset.brand,

      itFullName,

      "Issued"
    ]]);


  SpreadsheetApp.flush();


  // =========================================
  // ACTIVITY LOG
  // =========================================

  saveActivityLog({

    user: {

      username:
        itUsername,

      fullName:
        itFullName,

      position:
        itPosition

    },

    action:
      "ISSUED HEADSET",

    assetId:
      headset.assetId,

    employee:
      issue.agentName,

    description:
      "Issued headset " +
      headset.assetId +
      " to " +
      issue.agentName +
      "."

  });


  // =========================================
  // FLUSH CHANGES
  // =========================================

  SpreadsheetApp.flush();


  // =========================================
  // SUCCESS
  // =========================================

  return {

    success: true,

    assetId:
      headset.assetId,

    employee:
      issue.agentName,

    issuedBy:
      itFullName

  };

}