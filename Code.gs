// ============================================================
//  EVERGREEN ENTERPRISES — USER PIN SET / RESET
//  AppScript Backend
//  Employee Register  : 1HWKZPhKRhcuvxBgyyN8zRt8p-SzYmKjJWiOdCgykBHs
//  UserSecrets Sheet  : 1hN4VEDNpVLD3lKuBPYCTOaViv7UpveRfud2d2gy15D0
// ============================================================

// ---------- CONFIG ----------
const EMP_SHEET_ID   = '1HWKZPhKRhcuvxBgyyN8zRt8p-SzYmKjJWiOdCgykBHs';
const EMP_TAB        = '0_EmployeeRegister_Live';

const PIN_SHEET_ID   = '1hN4VEDNpVLD3lKuBPYCTOaViv7UpveRfud2d2gy15D0';
const PIN_TAB        = 'UserSecrets';

// Employee Register columns (0-based index)
const COL_EMP_EMAIL  = 15;   // Column P
const COL_EMP_NAME   = 8;    // Column I
const COL_EMP_REF    = 4;    // Column E
const COL_EMP_STATUS = 23;   // Column X
const ACTIVE_STATUS  = 'CURRENT';

// UserSecrets columns (1-based for sheet range; 0-based for array reads)
// A=MailID | B=UserName | C=EmpRef | D=CurrentPIN | E=ModifiedPIN | F=ModificationDate
const DEFAULT_PIN    = '1234';

// ---------- ENTRY POINT ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Evergreen Enterprises — User PIN Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- STEP 1 : VALIDATE EMPLOYEE ----------
// Called from frontend when user submits their email.
// Returns { success, name, empRef, message }
function validateEmployee(email) {
  try {
    email = email.trim().toLowerCase();

    const sheet = SpreadsheetApp.openById(EMP_SHEET_ID)
                    .getSheetByName(EMP_TAB);
    const data  = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row       = data[i];
      const rowEmail  = (row[COL_EMP_EMAIL] || '').toString().trim().toLowerCase();

      if (rowEmail !== email) continue;

      const status = (row[COL_EMP_STATUS] || '').toString().trim();
      if (status !== ACTIVE_STATUS) {
        return {
          success : false,
          message : 'Your account is not active. Please contact HR.'
        };
      }

      return {
        success : true,
        name    : (row[COL_EMP_NAME] || '').toString().trim(),
        empRef  : (row[COL_EMP_REF]  || '').toString().trim()
      };
    }

    return { success: false, message: 'Email ID not found in Employee Register.' };

  } catch (e) {
    return { success: false, message: 'Server error: ' + e.message };
  }
}

// ---------- STEP 2 : SET / RESET PIN ----------
// Called from frontend after employee is validated.
// Returns { success, message }
function setPin(email, userName, empRef, currentPin, newPin) {
  try {
    email      = email.trim().toLowerCase();
    currentPin = currentPin.toString().trim();
    newPin     = newPin.toString().trim();

    // --- Validate PIN format (4–12 digits, numbers only) ---
    if (!/^\d{4,12}$/.test(newPin)) {
      return { success: false, message: 'New PIN must be 4–12 digits (numbers only).' };
    }
    if (newPin === currentPin) {
      return { success: false, message: 'New PIN must be different from Current PIN.' };
    }

    const sheet  = SpreadsheetApp.openById(PIN_SHEET_ID).getSheetByName(PIN_TAB);
    const data   = sheet.getDataRange().getValues();

    let userRow     = -1;   // 1-based sheet row
    let activePin   = DEFAULT_PIN;

    for (let i = 1; i < data.length; i++) {
      const rowEmail = (data[i][0] || '').toString().trim().toLowerCase();
      if (rowEmail !== email) continue;

      userRow   = i + 1;  // convert to 1-based
      // Active PIN = Modified PIN (col E, index 4); fall back to default if blank
      const modifiedPin = (data[i][4] || '').toString().trim();
      activePin = modifiedPin !== '' ? modifiedPin : DEFAULT_PIN;
      break;
    }

    // --- Verify current PIN ---
    if (currentPin !== activePin) {
      const hint = (userRow === -1)
        ? 'First-time setup: use 1234 as your Current PIN.'
        : 'Current PIN is incorrect.';
      return { success: false, message: hint };
    }

    const now = new Date();

    if (userRow === -1) {
      // New user — append a fresh row
      // D (CurrentPIN) = default PIN used to authenticate; E (ModifiedPIN) = new PIN chosen
      sheet.appendRow([email, userName, empRef, DEFAULT_PIN, newPin, now]);
    } else {
      // Existing user — update in place
      sheet.getRange(userRow, 4).setValue(activePin);  // D: record PIN that was current
      sheet.getRange(userRow, 5).setValue(newPin);     // E: new active PIN
      sheet.getRange(userRow, 6).setValue(now);        // F: modification date
    }

    return { success: true, message: 'PIN updated successfully!' };

  } catch (e) {
    return { success: false, message: 'Server error: ' + e.message };
  }
}
