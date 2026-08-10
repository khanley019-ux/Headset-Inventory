function getRecentTransactions() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const transactions = [];


  // ==================================================
  // HELPER - FORMAT DATE
  // ==================================================

  function formatDate(date) {

    if (!date) {
      return "";
    }

    if (date instanceof Date) {

      return Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        "MM/dd/yyyy"
      );

    }

    return String(date);

  }


  // ==================================================
  // ISSUANCE LOGS
  // ==================================================

  const issueSheet =
    ss.getSheetByName("Issuance Logs");


  if (issueSheet) {

    const lastRow =
      issueSheet.getLastRow();


    if (lastRow >= 2) {

      const data =
        issueSheet
          .getRange(
            2,
            2,
            lastRow - 1,
            6
          )
          .getValues();


      data.forEach(function(row) {

        const date =
          row[0];

        const employee =
          row[1];

        const assetId =
          row[2];

        const brand =
          row[3];


        if (!assetId) {
          return;
        }


        transactions.push({

          date: formatDate(date),

          action: "Issued",

          assetId: String(assetId),

          brand: String(brand || ""),

          employee: String(employee || "")

        });

      });

    }

  }


  // ==================================================
  // BORROW LOGS
  // ==================================================

  const borrowSheet =
    ss.getSheetByName("Borrow Logs");


  if (borrowSheet) {

    const lastRow =
      borrowSheet.getLastRow();


    if (lastRow >= 2) {

      const data =
        borrowSheet
          .getRange(
            2,
            2,
            lastRow - 1,
            6
          )
          .getValues();


      data.forEach(function(row) {

        const date =
          row[0];

        const employee =
          row[1];

        const assetId =
          row[2];

        const brand =
          row[3];


        if (!assetId) {
          return;
        }


        transactions.push({

          date: formatDate(date),

          action: "Borrowed",

          assetId: String(assetId),

          brand: String(brand || ""),

          employee: String(employee || "")

        });

      });

    }

  }


  // ==================================================
  // RETURN LOGS
  // ==================================================

  const returnSheet =
    ss.getSheetByName("Return Logs");


  if (returnSheet) {

    const lastRow =
      returnSheet.getLastRow();


    if (lastRow >= 3) {

      const data =
        returnSheet
          .getRange(
            3,
            2,
            lastRow - 2,
            7
          )
          .getValues();


      data.forEach(function(row) {

        const date =
          row[0];

        const employee =
          row[1];

        const assetId =
          row[2];

        const brand =
          row[3];


        if (!assetId) {
          return;
        }


        transactions.push({

          date: formatDate(date),

          action: "Returned",

          assetId: String(assetId),

          brand: String(brand || ""),

          employee: String(employee || "")

        });

      });

    }

  }


  // ==================================================
  // SORT NEWEST FIRST
  // ==================================================

  transactions.sort(function(a, b) {

    const dateA =
      new Date(a.date).getTime();

    const dateB =
      new Date(b.date).getTime();


    return dateB - dateA;

  });


  // ==================================================
  // RETURN LATEST 10
  // ==================================================

  return transactions.slice(0, 10);

}