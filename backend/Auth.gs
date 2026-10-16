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
