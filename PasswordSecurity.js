/**
 * =========================================
 * PASSWORD SECURITY
 * =========================================
 *
 * Stores passwords as:
 *
 * SHA-256(password + salt)
 *
 * The actual password is never stored.
 *
 * Users sheet:
 *
 * A = Username
 * B = Password Hash
 * C = Full Name
 * D = Position
 * E = Status
 * F = Salt
 */


/**
 * Generate a random salt.
 */
function generatePasswordSalt() {

  const randomBytes =
    Utilities
      .getUuid()
      .replace(/-/g, "");

  return randomBytes;

}


/**
 * Create a SHA-256 password hash.
 */
function hashPassword(password, salt) {

  password =
    String(password || "");

  salt =
    String(salt || "");


  const value =
    password +
    salt;


  const digest =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      value,
      Utilities.Charset.UTF_8
    );


  return digest
    .map(function(byte) {

      const value =
        byte < 0
          ? byte + 256
          : byte;

      return (
        value
          .toString(16)
          .padStart(2, "0")
      );

    })
    .join("");

}


/**
 * Create a complete password record.
 */
function createPasswordRecord(password) {

  if (!password) {

    throw new Error(
      "Password is required."
    );

  }


  const salt =
    generatePasswordSalt();


  const hash =
    hashPassword(
      password,
      salt
    );


  return {

    hash: hash,

    salt: salt

  };

}


/**
 * Verify a password against
 * a stored hash and salt.
 */
function verifyPassword(
  password,
  storedHash,
  salt
) {

  if (
    !password ||
    !storedHash ||
    !salt
  ) {

    return false;

  }


  const calculatedHash =
    hashPassword(
      password,
      salt
    );


  return calculatedHash ===
    String(storedHash).trim();

}

/**
 * =========================================
 * MIGRATE EXISTING PASSWORDS
 * =========================================
 *
 * ONE-TIME MIGRATION
 *
 * Converts existing plaintext passwords
 * in Users!B into:
 *
 * Password Hash + Salt
 *
 * Users sheet:
 *
 * A = Username
 * B = Password Hash
 * C = Full Name
 * D = Position
 * E = Status
 * F = Salt
 *
 * IMPORTANT:
 * Run this ONLY ONCE.
 */
function migrateExistingPasswords() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName("Users");


  if (!sheet) {

    throw new Error(
      "Users sheet was not found."
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    throw new Error(
      "No users found in the Users sheet."
    );

  }


  /*
   * Get A:F
   */

  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();


  let migratedCount = 0;


  for (
    let i = 0;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    const username =
      String(row[0] || "")
        .trim();


    const existingPassword =
      String(row[1] || "");


    const existingSalt =
      String(row[5] || "")
        .trim();


    /*
     * Skip empty users.
     */

    if (!username) {

      continue;

    }


    /*
     * If Salt already exists,
     * this account has already
     * been migrated.
     */

    if (existingSalt) {

      continue;

    }


    /*
     * Skip accounts with no password.
     */

    if (!existingPassword) {

      continue;

    }


    /*
     * Create secure password record.
     */

    const passwordRecord =
      createPasswordRecord(
        existingPassword
      );


    /*
     * B = Password Hash
     */

    sheet
      .getRange(
        i + 2,
        2
      )
      .setValue(
        passwordRecord.hash
      );


    /*
     * F = Salt
     */

    sheet
      .getRange(
        i + 2,
        6
      )
      .setValue(
        passwordRecord.salt
      );


    migratedCount++;

  }


  SpreadsheetApp.flush();


  return {

    success: true,

    migrated:
      migratedCount,

    message:
      migratedCount +
      " user account(s) migrated successfully."

  };

}