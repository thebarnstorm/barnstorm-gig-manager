/**
 * ====================================================================
 * PUSH NOTIFICATIONS - Firebase Cloud Messaging (FCM) integration
 * ====================================================================
 * This is a NEW file to add to your existing Apps Script project.
 * In the Apps Script editor: click the "+" next to Files > Script,
 * name it "PushNotifications", and paste this in.
 *
 * SETUP REQUIRED BEFORE THIS WORKS:
 *
 * 1. In the Firebase Console: Project Settings (gear icon) >
 *    Service Accounts tab > click "Generate new private key."
 *    This downloads a .json file. Keep it safe, do not share it publicly.
 *
 * 2. Open that downloaded .json file in a text editor and copy the
 *    ENTIRE contents.
 *
 * 3. In THIS Apps Script project: click Project Settings (gear icon
 *    on the left sidebar) > scroll to "Script Properties" > click
 *    "Add script property" and add:
 *       Property: FCM_SERVICE_ACCOUNT_JSON
 *       Value:    (paste the entire JSON file contents here)
 *
 *    Add a second property:
 *       Property: FCM_PROJECT_ID
 *       Value:    barnstorm-gig-manager
 *
 * 4. In your existing doPost(e) function, find where you parse the
 *    incoming request data, and add this near the top, before your
 *    existing N/A-handling logic:
 *
 *       if (data.action === 'savePushToken') {
 *         return handleSavePushToken(data);
 *       }
 *
 *    (The exact variable name for your parsed data may differ -
 *    match whatever you already call it, e.g. "data" or "requestData".)
 *
 * 5. Redeploy the Apps Script web app (same process you always use),
 *    and send Claude the new URL to update in index.html.
 *
 * A note on this code: the token-signing logic below (getFcmAccessToken_)
 * uses a standard, widely-documented Apps Script pattern for authenticating
 * as a Google service account. It's worth testing carefully with
 * testSendPush() below before relying on it, and checking the Apps Script
 * docs for Utilities.computeRsaSha256Signature if you hit any errors.
 * ====================================================================
 */

const PUSH_TOKENS_SHEET_NAME = 'Push Tokens';

/**
 * Saves or updates a device's push notification token.
 * Called from doPost() when action === 'savePushToken'.
 */
function handleSavePushToken(data) {
  const name = data.name || '';
  const email = data.email || '';
  const token = data.token || '';

  if (!token || !email) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Missing token or email' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PUSH_TOKENS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(PUSH_TOKENS_SHEET_NAME);
    sheet.appendRow(['Name', 'Email', 'Token', 'Updated At']);
  }

  const values = sheet.getDataRange().getValues();
  let existingRow = -1;
  for (let i = 1; i < values.length; i++) {
    if ((values[i][1] || '').toString().toLowerCase() === email.toLowerCase()) {
      existingRow = i + 1; // sheet rows are 1-indexed
      break;
    }
  }

  const now = new Date();
  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, 4).setValues([[name, email, token, now]]);
  } else {
    sheet.appendRow([name, email, token, now]);
  }

  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Gets a short-lived access token for calling the FCM API,
 * authenticated as your Firebase service account.
 */
function getFcmAccessToken_() {
  const props = PropertiesService.getScriptProperties();
  const serviceAccountJson = props.getProperty('FCM_SERVICE_ACCOUNT_JSON');
  if (!serviceAccountJson) {
    throw new Error('Missing FCM_SERVICE_ACCOUNT_JSON script property. See setup instructions at the top of this file.');
  }
  const serviceAccount = JSON.parse(serviceAccountJson);

  const header = { alg: 'RS256', typ: 'JWT' };
  const now = Math.floor(Date.now() / 1000);
  const claimSet = {
    iss: serviceAccount.client_email,
    scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  };

  const base64Encode = (obj) => Utilities.base64EncodeWebSafe(JSON.stringify(obj)).replace(/=+$/, '');
  const toSign = base64Encode(header) + '.' + base64Encode(claimSet);
  const signatureBytes = Utilities.computeRsaSha256Signature(toSign, serviceAccount.private_key);
  const signature = Utilities.base64EncodeWebSafe(signatureBytes).replace(/=+$/, '');
  const jwt = toSign + '.' + signature;

  const response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  if (!result.access_token) {
    throw new Error('Failed to get FCM access token: ' + response.getContentText());
  }
  return result.access_token;
}

/**
 * Sends a push notification to a single device token.
 * title/body show in the notification. data is optional extra info
 * (e.g. { gigId: '123' }) your app could use later to jump to a gig.
 */
function sendPushToToken_(token, title, body, data) {
  const props = PropertiesService.getScriptProperties();
  const projectId = props.getProperty('FCM_PROJECT_ID');
  if (!projectId) {
    throw new Error('Missing FCM_PROJECT_ID script property.');
  }

  const accessToken = getFcmAccessToken_();
  const url = 'https://fcm.googleapis.com/v1/projects/' + projectId + '/messages:send';

  // NOTE: we intentionally use only "data" here, not "notification".
  // If both are present, browsers show the notification automatically AND
  // our service worker shows it again manually, causing duplicates.
  const message = {
    message: {
      token: token,
      data: Object.assign({ title: title, body: body }, data || {})
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + accessToken },
    payload: JSON.stringify(message),
    muteHttpExceptions: true
  });

  Logger.log('FCM response: ' + response.getContentText());
  return response.getContentText();
}

/**
 * TEST FUNCTION.
 * Run this manually from the Apps Script editor (pick "testSendPush"
 * from the function dropdown at the top, then click Run) to confirm
 * the whole pipeline works before we wire up real triggers.
 *
 * Before running: change nameToTest to match someone who has already
 * tapped "Enable Notifications" in the app (so a token exists for them
 * in the "Push Tokens" sheet tab).
 */
function testSendPush() {
  const nameToTest = 'booking'; // <-- change this to a real name

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PUSH_TOKENS_SHEET_NAME);
  if (!sheet) {
    Logger.log('No "Push Tokens" sheet found yet - has anyone enabled notifications in the app?');
    return;
  }

  const values = sheet.getDataRange().getValues();
  let token = null;
  for (let i = 1; i < values.length; i++) {
    if ((values[i][0] || '').toString().toLowerCase() === nameToTest.toLowerCase()) {
      token = values[i][2];
      break;
    }
  }

  if (!token) {
    Logger.log('No saved token found for "' + nameToTest + '". Check the Push Tokens tab for the exact name on file.');
    return;
  }

  const result = sendPushToToken_(token, 'Test Notification', 'If you see this, push notifications are working!', {});
  Logger.log(result);
}