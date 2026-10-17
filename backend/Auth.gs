// ============================================================
// Auth.gs - checks that a request comes from a signed-in admin
// ============================================================
// The app sends the user's Firebase login token (idToken) with each
// protected request. We ask Google's Identity Toolkit whether that token
// is valid and which email it belongs to. Only ADMIN_AUTH_EMAIL passes.
//
// ROLLOUT SWITCH:
//   false = log unauthenticated requests in the Debug Log but still allow them
//   true  = block them
const ENFORCE_ADMIN_AUTH = true;

const ADMIN_AUTH_EMAIL = 'booking@thebarnstorm.com';
// Firebase web API key (same public value that is in index.html)
const FIREBASE_WEB_API_KEY = 'AIzaSyDhYBUzbpBBv40-X6SLASGikxoMUWcCU5Q';
const AUTH_CACHE_SECONDS = 300;

function verifiedEmailFromToken_(idToken) {
  if (!idToken || typeof idToken !== 'string' || idToken.length < 100) return '';

  // Cache by a fingerprint of the token so we do not call Google every time
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken);
  const cacheKey = 'auth_' + digest.map(function (b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('');
  const cache = CacheService.getScriptCache();
  const hit = cache.get(cacheKey);
  if (hit !== null) return hit; // may be '' for a cached failure

  let email = '';
  try {
    const resp = UrlFetchApp.fetch(
      'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_WEB_API_KEY,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ idToken: idToken }),
        muteHttpExceptions: true
      }
    );
    if (resp.getResponseCode() === 200) {
      const data = JSON.parse(resp.getContentText());
      const user = data.users && data.users[0];
      if (user && user.email) email = String(user.email).toLowerCase();
    }
  } catch (err) {
    debugLog_('Auth check error: ' + err);
    return ''; // do not cache errors
  }
  cache.put(cacheKey, email, AUTH_CACHE_SECONDS);
  return email;
}

// Returns null if the request may proceed, or a response to send back if not.
function requireAdmin_(idToken, actionName) {
  const email = verifiedEmailFromToken_(idToken);
  if (email === ADMIN_AUTH_EMAIL.toLowerCase()) return null;

  debugLog_('AUTH: ' + (ENFORCE_ADMIN_AUTH ? 'BLOCKED' : 'would block (not enforcing)') +
    ' action=' + actionName + ' tokenProvided=' + (idToken ? 'yes' : 'no') + ' email=' + (email || 'none'));
  if (!ENFORCE_ADMIN_AUTH) return null;

  return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'unauthorized' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// N/A requests: only the member being marked N/A may send one
// ============================================================
// ROLLOUT SWITCH (same idea as ENFORCE_ADMIN_AUTH above):
//   false = log requests that WOULD be blocked in the Debug Log, but allow them
//   true  = block them
// Check the Debug Log for lines starting "AUTH: N/A ... would block" after real
// members have used the N/A button. If real members show up there, fix their
// names first, then switch this to true.
const ENFORCE_NA_SENDER_AUTH = false;

// Like verifiedEmailFromToken_, but also returns the login's display name.
// Returns { email: '', displayName: '' } when the token is missing or invalid.
function verifiedUserFromToken_(idToken) {
  const none = { email: '', displayName: '' };
  if (!idToken || typeof idToken !== 'string' || idToken.length < 100) return none;

  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken);
  const cacheKey = 'authu_' + digest.map(function (b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('');
  const cache = CacheService.getScriptCache();
  const hit = cache.get(cacheKey);
  if (hit !== null) {
    try { return JSON.parse(hit); } catch (e) { /* fall through and look it up again */ }
  }

  let result = none;
  try {
    const resp = UrlFetchApp.fetch(
      'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_WEB_API_KEY,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ idToken: idToken }),
        muteHttpExceptions: true
      }
    );
    if (resp.getResponseCode() === 200) {
      const data = JSON.parse(resp.getContentText());
      const user = data.users && data.users[0];
      if (user && user.email) {
        result = { email: String(user.email).toLowerCase(), displayName: String(user.displayName || '') };
      }
    }
  } catch (err) {
    debugLog_('Auth user lookup error: ' + err);
    return none; // do not cache errors
  }
  cache.put(cacheKey, JSON.stringify(result), AUTH_CACHE_SECONDS);
  return result;
}

// Same strict rule the app uses: full name matches exactly, or BOTH first and last name match.
function namesMatch_(a, b) {
  const x = String(a || '').toLowerCase().trim();
  const y = String(b || '').toLowerCase().trim();
  if (!x || !y) return false;
  if (x === y) return true;
  const xp = x.split(/[\s@._-]+/).filter(function (p) { return p.length > 0; });
  const yp = y.split(/[\s@._-]+/).filter(function (p) { return p.length > 0; });
  if (xp.length < 2 || yp.length < 2) return false;
  return xp[0] === yp[0] && xp[xp.length - 1] === yp[yp.length - 1];
}

// The name the app saved for this email on the Push Tokens tab ('' if none).
function pushTokenNameForEmail_(email) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PUSH_TOKENS_SHEET_NAME);
    if (!sheet) return '';
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if ((values[i][1] || '').toString().toLowerCase() === email) return (values[i][0] || '').toString();
    }
  } catch (err) {
    debugLog_('Push Tokens name lookup error: ' + err);
  }
  return '';
}

// Returns null if the N/A request may proceed, or a response to send back if not.
function checkNaSender_(params) {
  const user = verifiedUserFromToken_(params.idToken);
  const memberName = params.memberName;
  let reason = '';

  if (!user.email) {
    reason = 'no valid login token';
  } else if (user.email === ADMIN_AUTH_EMAIL.toLowerCase()) {
    return null; // the booking account may mark anyone
  } else {
    const emailName = user.email.split('@')[0];
    const savedName = pushTokenNameForEmail_(user.email);
    if (namesMatch_(user.displayName, memberName) ||
        namesMatch_(savedName, memberName) ||
        namesMatch_(emailName, memberName)) {
      return null;
    }
    reason = 'login ' + user.email + ' (name "' + user.displayName + '", saved name "' + savedName + '") does not match "' + memberName + '"';
  }

  debugLog_('AUTH: N/A ' + (ENFORCE_NA_SENDER_AUTH ? 'BLOCKED' : 'would block (not enforcing)') + ' - ' + reason);
  if (!ENFORCE_NA_SENDER_AUTH) return null;

  return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'unauthorized' }))
    .setMimeType(ContentService.MimeType.JSON);
}
