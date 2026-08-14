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

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName("Users");

  if (!sheet) {
    return {
      success: false,
      message: "Users sheet was not found."
    };
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      success: false,
      message: "No IT user accounts have been configured."
    };
  }


  /*
   * Users sheet:
   *
   * A = Username
   * B = Password Hash
   * C = Full Name
   * D = Position
   * E = Status
   * F = Salt
   * G = Must Change Password
   */

  const data = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      7
    )
    .getValues();


  for (let i = 0; i < data.length; i++) {

    const row = data[i];


    const sheetUsername =
      String(row[0] || "").trim();

    const storedHash =
      String(row[1] || "").trim();

    const fullName =
      String(row[2] || "").trim();

    const position =
      String(row[3] || "").trim();

    const status =
      String(row[4] || "")
        .trim()
        .toLowerCase();

    const salt =
      String(row[5] || "").trim();

    const mustChangePassword =
      String(row[6] || "")
        .trim()
        .toLowerCase();


    // =========================================
    // USERNAME
    // =========================================

    if (sheetUsername !== username) {
      continue;
    }


    // =========================================
    // ACCOUNT STATUS
    // =========================================

    if (status !== "active") {

      return {
        success: false,
        message: "This IT account is inactive."
      };

    }


    // =========================================
    // PASSWORD DATA
    // =========================================

    if (!storedHash || !salt) {

      return {
        success: false,
        message:
          "This account has not been configured for secure password authentication."
      };

    }


    // =========================================
    // VERIFY PASSWORD
    // =========================================

    const validPassword =
      verifyPassword(
        password,
        storedHash,
        salt
      );


    if (!validPassword) {

      return {
        success: false,
        message: "Invalid username or password."
      };

    }


    // =========================================
    // CHECK MUST CHANGE PASSWORD
    // =========================================

    const requiresPasswordChange =
      mustChangePassword === "true" ||
      mustChangePassword === "yes" ||
      mustChangePassword === "1";


    // =========================================
    // LOGIN SUCCESS
    // =========================================

    return {

      success: true,

      username:
        sheetUsername,

      fullName:
        fullName ||
        sheetUsername,

      position:
        position ||
        "IT Staff",

      mustChangePassword:
        requiresPasswordChange

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

function initializeMustChangePassword() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName("Users");

  if (!sheet) {
    throw new Error("Users sheet was not found.");
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return {
      success: false,
      message: "No users found."
    };
  }

  // G = Must Change Password
  const range =
    sheet.getRange(
      2,
      7,
      lastRow - 1,
      1
    );

  const values =
    range.getValues();

  for (let i = 0; i < values.length; i++) {

    // Only initialize blank values.
    // Do NOT overwrite TRUE values.
    if (
      values[i][0] === "" ||
      values[i][0] === null
    ) {

      values[i][0] = false;

    }

  }

  range.setValues(values);

  SpreadsheetApp.flush();

  return {
    success: true,
    message:
      "Must Change Password values initialized."
  };

}