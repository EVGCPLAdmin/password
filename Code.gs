// ============================================================
//  EVERGREEN ENTERPRISES — USER PIN SET / RESET
//  AppScript Backend — API Mode (called from GitHub Pages)
//  Employee Register  : 1HWKZPhKRhcuvxBgyyN8zRt8p-SzYmKjJWiOdCgykBHs
//  UserSecrets Sheet  : 1hN4VEDNpVLD3lKuBPYCTOaViv7UpveRfud2d2gy15D0
// ============================================================

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
const DEFAULT_PIN    = '1234';

// ── CORS helper ──────────────────────────────────────────────
function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ENTRY POINT ──────────────────────────────────────────────
// Handles both:
//   ?action=validateEmployee&email=xxx
//   ?action=setPin&email=xxx&userName=xxx&empRef=xxx&currentPin=xxx&newPin=xxx
function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === 'validateEmployee') {
      return corsResponse(validateEmployee(e.parameter.email));
    }

    if (action === 'setPin') {
      return corsResponse(setPin(
        e.parameter.email,
        e.parameter.userName,
        e.parameter.empRef,
        e.parameter.currentPin,
        e.parameter.newPin
      ));
    }

    // No action param — return API info
    return corsResponse({ status: 'Evergreen PIN API running' });

  } catch(err) {
    return corsResponse({ success: false, message: 'Server error: ' + err.message });
  }
}

// ── VALIDATE EMPLOYEE ─────────────────────────────────────────
function validateEmployee(email) {
  try {
    email = (email || '').trim().toLowerCase();
    const sheet = SpreadsheetApp.openById(EMP_SHEET_ID).getSheetByName(EMP_TAB);
    const data  = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row      = data[i];
      const rowEmail = (row[COL_EMP_EMAIL] || '').toString().trim().toLowerCase();
      if (rowEmail !== email) continue;

      const status = (row[COL_EMP_STATUS] || '').toString().trim().toUpperCase();
      if (status !== ACTIVE_STATUS) {
        return { success: false, message: 'Your account is not active. Please contact HR.' };
      }
      return {
        success : true,
        name    : (row[COL_EMP_NAME] || '').toString().trim(),
        empRef  : (row[COL_EMP_REF]  || '').toString().trim()
      };
    }
    return { success: false, message: 'Email ID not found in Employee Register.' };

  } catch(e) {
    return { success: false, message: 'Server error: ' + e.message };
  }
}

// ── SET / RESET PIN ──────────────────────────────────────────
function setPin(email, userName, empRef, currentPin, newPin) {
  try {
    email      = (email || '').trim().toLowerCase();
    currentPin = (currentPin || '').toString().trim();
    newPin     = (newPin || '').toString().trim();

    if (!/^\d{4,12}$/.test(newPin)) {
      return { success: false, message: 'New PIN must be 4–12 digits (numbers only).' };
    }
    if (newPin === currentPin) {
      return { success: false, message: 'New PIN must be different from Current PIN.' };
    }

    const sheet  = SpreadsheetApp.openById(PIN_SHEET_ID).getSheetByName(PIN_TAB);
    const data   = sheet.getDataRange().getValues();

    let userRow   = -1;
    let activePin = DEFAULT_PIN;

    for (let i = 1; i < data.length; i++) {
      const rowEmail = (data[i][0] || '').toString().trim().toLowerCase();
      if (rowEmail !== email) continue;
      userRow   = i + 1;
      const mod = (data[i][4] || '').toString().trim();
      activePin = mod !== '' ? mod : DEFAULT_PIN;
      break;
    }

    if (currentPin !== activePin) {
      return {
        success: false,
        message: userRow === -1
          ? 'First-time setup: use 1234 as your Current PIN.'
          : 'Current PIN is incorrect.'
      };
    }

    const now = new Date();
    if (userRow === -1) {
      sheet.appendRow([email, userName, empRef, DEFAULT_PIN, newPin, now]);
    } else {
      sheet.getRange(userRow, 4).setValue(activePin);
      sheet.getRange(userRow, 5).setValue(newPin);
      sheet.getRange(userRow, 6).setValue(now);
    }

    return { success: true, message: 'PIN updated successfully!' };

  } catch(e) {
    return { success: false, message: 'Server error: ' + e.message };
  }
}
