function getDashboardStats() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.INVENTORY);

  const data = sheet.getDataRange().getValues();

  let total = 0;
  let available = 0;
  let assigned = 0;
  let repair = 0;

  for (let i = 1; i < data.length; i++) {

    const assetId = data[i][1]; // Column B
    const status = String(data[i][5]).trim(); // Column F

    if (assetId === "") continue;

    total++;

    if (status === "Available") {
      available++;
    }

    if (status === "Assigned") {
      assigned++;
    }

    if (status === "Under Repair") {
      repair++;
    }

  }

  return {
    total: total,
    available: available,
    assigned: assigned,
    repair: repair
  };

}

function getRecentTransactions() {

  // Get the same transaction data used
  // by the Transaction History page

  const transactions =
    getTransactionHistory();


  // Return only the latest 5 transactions

  return transactions.slice(0, 5);

}