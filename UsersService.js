function updateActivitySummary() {

  // =========================================
  // TOTAL ACTIVITIES
  // =========================================

  const total =
    activityLogs.length;


  // =========================================
  // UNIQUE HEADSETS IN ACTIVITY LOG
  // =========================================

  const assets = new Set();

  activityLogs.forEach(function(log) {

    if (log.assetId) {

      assets.add(
        String(log.assetId).trim()
      );

    }

  });


  // =========================================
  // TODAY'S ACTIVITIES
  // =========================================

  const today =
    new Date();

  today.setHours(
    0,
    0,
    0,
    0
  );


  const todayCount =
    activityLogs.filter(function(log) {

      if (!log.timestamp) {
        return false;
      }

      const date =
        new Date(log.timestamp);

      date.setHours(
        0,
        0,
        0,
        0
      );

      return (
        date.getTime() ===
        today.getTime()
      );

    }).length;


  // =========================================
  // GET REPORT ELEMENTS
  // =========================================

  const totalElement =
    document.getElementById(
      "activityTotal"
    );

  const staffElement =
    document.getElementById(
      "activityITStaff"
    );

  const headsetElement =
    document.getElementById(
      "activityHeadsets"
    );

  const todayElement =
    document.getElementById(
      "activityToday"
    );


  // =========================================
  // UPDATE TOTAL ACTIVITIES
  // =========================================

  if (totalElement) {

    totalElement.textContent =
      total;

  }


  // =========================================
  // UPDATE HEADSETS
  // =========================================

  if (headsetElement) {

    headsetElement.textContent =
      assets.size;

  }


  // =========================================
  // UPDATE TODAY
  // =========================================

  if (todayElement) {

    todayElement.textContent =
      todayCount;

  }


  // =========================================
  // GET REAL IT STAFF COUNT
  // =========================================

  if (staffElement) {

    staffElement.textContent = "...";


    google.script.run

      .withSuccessHandler(function(count) {

        staffElement.textContent =
          Number(count) || 0;

      })

      .withFailureHandler(function(error) {

        console.error(
          "IT Staff count error:",
          error
        );

        staffElement.textContent =
          "0";

      })

      .getActiveITStaffCount();

  }

}

function getActiveITStaffCount() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName("Users");

  if (!sheet) {

    throw new Error(
      'Users sheet was not found.'
    );

  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    return 0;

  }

  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        5
      )
      .getValues();

  let count = 0;

  data.forEach(function(row) {

    const username =
      String(row[0] || "")
        .trim();

    const status =
      String(row[4] || "")
        .trim()
        .toLowerCase();

    if (
      username &&
      status === "active"
    ) {

      count++;

    }

  });

  return count;

}