function saveIssueTransaction(issue) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss
    .getSheetByName(CONFIG.SHEETS.ISSUE);


  // =========================================
  // FIND HEADSET
  // =========================================

  const headset = getHeadsetByAssetId(issue.assetId);


  if (!headset) {
    throw new Error("Headset not found.");
  }


  // =========================================
  // SAVE TO ISSUANCE LOGS
  // =========================================

  const row = getFirstEmptyRow(sheet, 2);


  sheet.getRange(row, 2, 1, 6).setValues([[
    new Date(),
    issue.agentName,
    headset.assetId,
    headset.brand,
    issue.issuedBy,
    "Issued"
  ]]);


  SpreadsheetApp.flush();


  // =========================================
  // ACTIVITY LOG
  // =========================================

  const currentUser =
    issue.currentITUser || {};

  logActivity({

    itUsername:
      currentUser.username || "",

    itFullName:
      currentUser.fullName ||
      issue.issuedBy ||
      "Unknown IT Staff",

    itPosition:
      currentUser.position ||
      "IT Staff",

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
      issue.agentName

  });


  // =========================================
  // DONE
  // =========================================

  SpreadsheetApp.flush();


  return true;

}