function getTransactionHistory() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const transactions = [];


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
            2,      // Row 2
            2,      // Column B
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

        const issuedBy =
          row[4];

        const status =
          row[5];


        if (!assetId) {
          return;
        }


        transactions.push({

          date: date,

          action: "Issued",

          assetId:
            String(assetId),

          brand:
            String(brand || ""),

          employee:
            String(employee || ""),

          details:
            issuedBy
              ? "Issued by " + issuedBy
              : "",

          source:
            "Issuance Logs"

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

        const issuedBy =
          row[4];

        const status =
          String(row[5] || "")
            .toLowerCase()
            .trim();


        if (!assetId) {
          return;
        }


        let action =
          "Borrowed";


        if (status === "returned") {

          action =
            "Returned";

        }


        transactions.push({

          date: date,

          action: action,

          assetId:
            String(assetId),

          brand:
            String(brand || ""),

          employee:
            String(employee || ""),

          details:
            issuedBy
              ? "Processed by " + issuedBy
              : "",

          source:
            "Borrow Logs"

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
            3,      // Row 3
            2,      // Column B
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

        const statusUponReturn =
          row[4];

        const checkedBy =
          row[5];

        const remarks =
          row[6];


        if (!assetId) {
          return;
        }


        let details = "";


        if (statusUponReturn) {

          details =
            String(statusUponReturn);

        }


        if (checkedBy) {

          details +=
            details
              ? " • Checked by " + checkedBy
              : "Checked by " + checkedBy;

        }


        if (remarks) {

          details +=
            details
              ? " • " + remarks
              : String(remarks);

        }


        transactions.push({

          date: date,

          action: "Returned",

          assetId:
            String(assetId),

          brand:
            String(brand || ""),

          employee:
            String(employee || ""),

          details:
            details,

          source:
            "Return Logs"

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
  // CONVERT DATES FOR GOOGLE.SCRIPT.RUN
  // ==================================================

  transactions.forEach(function(transaction) {

    if (transaction.date instanceof Date) {

      transaction.date =
        transaction.date.toISOString();

    }

  });


  console.log(
    "Total transactions:",
    transactions.length
  );


  return transactions;

}

function testGetTransactionHistory() {

  const result =
    getTransactionHistory();

  console.log(
    "Returned transactions:",
    result.length
  );

  console.log(
    JSON.stringify(
      result.slice(0, 5),
      null,
      2
    )
  );

}