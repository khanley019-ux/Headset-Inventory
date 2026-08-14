/**
 * =========================================
 * ACTIVITY LOGS
 * =========================================
 *
 * Records actions performed by logged-in IT
 * staff.
 *
 * Sheet required:
 * Activity Logs
 *
 * Columns:
 * A = Timestamp
 * B = IT Username
 * C = IT Full Name
 * D = IT Position
 * E = Action
 * F = Asset ID
 * G = Employee
 * H = Description
 */

/**
 * Create Activity Logs sheet if it doesn't exist.
 */
function setupActivityLogs() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("Activity Logs");

  if (!sheet) {

    sheet = ss.insertSheet("Activity Logs");

  }

  // Add headers only if the sheet is empty
  if (sheet.getLastRow() === 0) {

    sheet.getRange(1, 1, 1, 8).setValues([[
      "Timestamp",
      "IT Username",
      "IT Full Name",
      "IT Position",
      "Action",
      "Asset ID",
      "Employee",
      "Description"
    ]]);

    sheet.getRange(1, 1, 1, 8)
      .setFontWeight("bold");

    sheet.setFrozenRows(1);

  }

  return "Activity Logs ready.";

}


/**
 * =========================================
 * LOG AN IT ACTIVITY
 * =========================================
 *
 * Example:
 *
 * logActivity(
 *   "ISSUED",
 *   "HS-0124",
 *   "Juan Dela Cruz",
 *   "Issued headset to employee"
 * );
 */
function logActivity(activity) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName("Activity Logs");

  if (!sheet) {
    throw new Error(
      'Activity Logs sheet was not found.'
    );
  }


  // =========================================
  // GET IT USER INFORMATION
  // =========================================

  const itUsername =
    activity.itUsername ||
    activity.username ||
    "";

  const itFullName =
    activity.itFullName ||
    activity.fullName ||
    activity.performedBy ||
    "Unknown IT Staff";

  const itPosition =
    activity.itPosition ||
    activity.position ||
    "IT Staff";


  // =========================================
  // GET ACTIVITY INFORMATION
  // =========================================

  const action =
    activity.action ||
    "UNKNOWN";

  const assetId =
    activity.assetId ||
    "";

  const employee =
    activity.employee ||
    activity.agentName ||
    "";

  const description =
    activity.description ||
    "";


  // =========================================
  // SAVE ACTIVITY
  // =========================================

  sheet.appendRow([

    new Date(),

    itUsername,

    itFullName,

    itPosition,

    action,

    assetId,

    employee,

    description

  ]);


  SpreadsheetApp.flush();


  return true;
}


/**
 * =========================================
 * GET CURRENT IT USER
 * =========================================
 *
 * The frontend stores the logged-in user
 * in window.currentITUser.
 *
 * Because Apps Script server-side functions
 * cannot directly access window.currentITUser,
 * we temporarily receive the user information
 * from the frontend when an activity is logged.
 *
 * This helper is kept as a fallback.
 */
function getCurrentITUserForLog_() {

  return {
    username: "",
    fullName: "",
    position: ""
  };

}


/**
 * =========================================
 * LOG ACTIVITY WITH USER INFORMATION
 * =========================================
 *
 * This is the function the frontend should
 * call when recording an activity.
 */
function saveActivityLog(activity) {

  if (!activity) {
    throw new Error("Activity information is required.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("Activity Logs");

  if (!sheet) {

    setupActivityLogs();

    sheet = ss.getSheetByName("Activity Logs");

  }

  const user = activity.user || {};

  sheet.appendRow([

    new Date(),

    user.username || "",
    user.fullName || "",
    user.position || "",

    activity.action || "",
    activity.assetId || "",
    activity.employee || "",
    activity.description || ""

  ]);

  return {
    success: true,
    message: "Activity logged successfully."
  };

}


/**
 * =========================================
 * GET ACTIVITY LOGS
 * =========================================
 */
/**
 * =========================================
 * GET ACTIVITY LOGS
 * =========================================
 *
 * Returns all Activity Logs for the
 * IT Activity Report.
 */
function getActivityLogs() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName("Activity Logs");


  // =========================================
  // SHEET CHECK
  // =========================================

  if (!sheet) {

    throw new Error(
      'Sheet "Activity Logs" was not found.'
    );

  }


  // =========================================
  // CHECK FOR DATA
  // =========================================

  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return [];

  }


  // =========================================
  // READ ACTIVITY LOGS
  // =========================================

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        8
      )
      .getValues();


  // =========================================
  // CONVERT TO SAFE JSON OBJECTS
  // =========================================

  const logs =
    values

      .filter(function(row) {

        return row.some(function(value) {

          return String(value || "").trim() !== "";

        });

      })

      .map(function(row) {

        let timestamp = "";

        if (row[0]) {

          if (row[0] instanceof Date) {

            timestamp =
              row[0].toISOString();

          } else {

            timestamp =
              String(row[0]);

          }

        }


        return {

          timestamp: timestamp,

          username:
            String(row[1] || "").trim(),

          fullName:
            String(row[2] || "").trim(),

          position:
            String(row[3] || "").trim(),

          action:
            String(row[4] || "").trim(),

          assetId:
            String(row[5] || "").trim(),

          employee:
            String(row[6] || "").trim(),

          description:
            String(row[7] || "").trim()

        };

      });


  // =========================================
  // NEWEST ACTIVITY FIRST
  // =========================================

  logs.reverse();


  console.log(
    "Activity Logs retrieved: " +
    logs.length
  );


  console.log(
    JSON.stringify(logs)
  );


  return logs;

}