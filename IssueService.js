function saveIssueTransaction(issue) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.ISSUE);

  const headset = getHeadsetByAssetId(issue.assetId);

  if (!headset) {
    throw new Error("Headset not found.");
  }

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

  return true;

}