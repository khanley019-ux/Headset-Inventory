/**
 * ============================================
 * IT USER AUTHENTICATION
 * ============================================
 *
 * Checks the username and password against
 * the "Users" sheet.
 *
 * Users sheet columns:
 *
 * A = Username
 * B = Password
 * C = Full Name
 * D = Position
 * E = Status
 */

function authenticateUser(username, password) {

  username = String(username || "").trim();
  password = String(password || "");

  if (!username || !password) {
    return {
      success: false,
      message: "Please enter your username and password."
    };
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName("Users");

  if (!sheet) {
    return {
      success: false,
      message: 'Users sheet was not found.'
    };
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return {
      success: false,
      message: "No IT user accounts have been configured."
    };
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

  for (let i = 0; i < data.length; i++) {

    const row =
      data[i];

    const sheetUsername =
      String(row[0] || "").trim();

    const sheetPassword =
      String(row[1] || "");

    const fullName =
      String(row[2] || "").trim();

    const position =
      String(row[3] || "").trim();

    const status =
      String(row[4] || "")
        .trim()
        .toLowerCase();

    // Check username
    if (sheetUsername !== username) {
      continue;
    }

    // Check account status
    if (status !== "active") {
      return {
        success: false,
        message: "This IT account is inactive."
      };
    }

    // Check password
    if (sheetPassword !== password) {
      return {
        success: false,
        message: "Invalid username or password."
      };
    }

    // =========================================
    // LOGIN SUCCESS
    // =========================================

    return {
      success: true,
      username: sheetUsername,
      fullName: fullName || sheetUsername,
      position: position || "IT Staff"
    };
  }

  return {
    success: false,
    message: "Invalid username or password."
  };
}

/**
 * ============================================
 * ACTIVITY LOGGING
 * ============================================
 *
 * Records actions performed by IT staff.
 *
 * Activity Logs columns:
 * A = Timestamp
 * B = IT User
 * C = Action
 * D = Description
 * E = Asset ID
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
  // GET ACTIVITY INFORMATION
  // =========================================

  const timestamp = new Date();

  const itUser =
    activity.itUser ||
    activity.performedBy ||
    "Unknown IT Staff";

  const action =
    activity.action ||
    "UNKNOWN";

  const description =
    activity.description ||
    "";

  const assetId =
    activity.assetId ||
    "";


  // =========================================
  // SAVE ACTIVITY
  // =========================================

  sheet.appendRow([
    timestamp,
    itUser,
    action,
    description,
    assetId
  ]);


  SpreadsheetApp.flush();


  return true;
}