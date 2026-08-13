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
function getActivityLogs() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName("Activity Logs");

  if (!sheet || sheet.getLastRow() < 2) {

    return [];

  }

  const lastRow = sheet.getLastRow();

  const values =
    sheet.getRange(
      2,
      1,
      lastRow - 1,
      8
    ).getValues();

  return values.map(function(row) {

    return {

      timestamp: row[0],

      username: row[1],

      fullName: row[2],

      position: row[3],

      action: row[4],

      assetId: row[5],

      employee: row[6],

      description: row[7]

    };

  }).reverse();

}

function getActivityLogs() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName("Activity Logs");

  if (!sheet) {

    throw new Error(
      'Sheet "Activity Logs" was not found.'
    );

  }

  const lastRow =
    sheet.getLastRow();

  // Header only / no records
  if (lastRow < 2) {
    return [];
  }

  /*
   * Activity Logs:
   *
   * A = Timestamp
   * B = IT Username
   * C = IT Full Name
   * D = IT Position
   * E = Action
   * F = Asset ID
   * G = Employee
   * H = Description
   */

  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        8
      )
      .getDisplayValues();


  const logs = [];


  data.forEach(function(row) {

    // Ignore completely empty rows
    if (
      row.every(function(value) {
        return String(value).trim() === "";
      })
    ) {
      return;
    }


    logs.push({

      timestamp:
        row[0] || "",

      username:
        row[1] || "",

      fullName:
        row[2] || "",

      position:
        row[3] || "",

      action:
        row[4] || "",

      assetId:
        row[5] || "",

      employee:
        row[6] || "",

      description:
        row[7] || ""

    });

  });


  console.log(
    "Activity Logs retrieved: " +
    logs.length
  );


  return logs;

}