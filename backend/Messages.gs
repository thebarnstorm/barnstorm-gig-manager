/**
 * ====================================================================
 * MESSAGES - Twilio (SMS/WhatsApp) → Firestore
 * ====================================================================
 * This is a NEW file to add to your existing Apps Script project.
 * In the Apps Script editor: click the "+" next to Files > Script,
 * name it "Messages", and paste this in.
 *
 * SETUP REQUIRED BEFORE THIS WORKS:
 *
 * 1. This reuses the SAME service account you already set up for push
 *    notifications (FCM_SERVICE_ACCOUNT_JSON in Script Properties).
 *    No new credentials needed - IF that account has permission to
 *    write to Firestore. This is the one thing I'm not fully certain
 *    about: Firebase's default service account usually has broad
 *    project access, but if you get a "permission denied" error when
 *    testing, that's the likely cause, and we'll need to grant it the
 *    "Cloud Datastore User" role in Google Cloud Console > IAM.
 *
 * 2. In your existing doPost(e) function in Code.gs, add this as the
 *    VERY FIRST LINE inside the function (before the JSON.parse line):
 *
 *       if (e.parameter && e.parameter.From && e.parameter.Body !== undefined) {
 *         return handleIncomingTwilioMessage(e.parameter);
 *       }
 *
 *    This is necessary because Twilio sends a different data format
 *    than the rest of your app, so it needs to be caught before your
 *    existing JSON.parse(e.postData.contents) line runs.
 *
 * 3. Redeploy the Apps Script web app (Deploy > Manage deployments >
 *    pencil icon > New version > Deploy), same as always.
 *
 * 4. In Twilio Console: go to Phone Numbers > Manage > Active Numbers,
 *    click your number, scroll to "Messaging Configuration", and set
 *    "A message comes in" to your Apps Script exec URL, method POST.
 *    For WhatsApp: go to Messaging > Try it out > Send a WhatsApp
 *    message (or your WhatsApp Sender settings) and set the same
 *    webhook URL there.
 *
 * 5. Test by texting your Twilio number and checking the Messages
 *    tab in the app (only visible to booking@thebarnstorm.com).
 * ====================================================================
 */

const FIRESTORE_PROJECT_ID = 'barnstorm-gig-manager';
const DEBUG_LOG_SHEET_NAME = 'Debug Log';

/**
 * Writes a line to a "Debug Log" tab in your spreadsheet. This is a
 * temporary troubleshooting aid - safe to ignore or delete this
 * sheet tab later once things are working reliably.
 */
function debugLog_(message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DEBUG_LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(DEBUG_LOG_SHEET_NAME);
      sheet.appendRow(['Timestamp', 'Message']);
    }
    sheet.appendRow([new Date(), message]);
  } catch (err) {
    // If even the debug log fails, there's nothing more we can do here.
  }
}

/**
 * Handles an incoming SMS or WhatsApp message from Twilio's webhook.
 * Twilio POSTs form fields (not JSON) - e.parameter.From, .Body, etc.
 */
function handleIncomingTwilioMessage(params) {
  debugLog_('Received Twilio webhook. From=' + params.From + ' Body=' + params.Body);

  try {
    let from = params.From || '';
    const isWhatsapp = from.indexOf('whatsapp:') === 0;
    if (isWhatsapp) {
      from = from.replace('whatsapp:', '');
    }

    const messageData = {
      from: from,
      to: (params.To || '').replace('whatsapp:', ''),
      body: params.Body || '',
      channel: isWhatsapp ? 'whatsapp' : 'sms',
      direction: 'inbound',
      read: false,
      timestamp: new Date().toISOString(),
      twilioSid: params.MessageSid || params.SmsMessageSid || ''
    };

    saveMessageToFirestore_(messageData);
    debugLog_('Firestore write succeeded.');

    if (isWhatsapp) {
      if (params.MessageType === 'button' && params.OriginalRepliedMessageSid) {
        checkAndUpdateReferralResponseFromButton_(params.OriginalRepliedMessageSid, params.ButtonPayload, params.ButtonText);
      } else {
        checkAndUpdateReferralResponse_(from, messageData.body);
      }
    }
  } catch (err) {
    debugLog_('ERROR: ' + err.toString());
    // Still respond to Twilio below so it doesn't retry/error on their end.
  }

  // Empty TwiML response = "received, no auto-reply."
  return ContentService.createTextOutput('<?xml version="1.0" encoding="UTF-8"?><Response></Response>')
    .setMimeType(ContentService.MimeType.XML);
}

/**
 * Gets a short-lived access token scoped for Firestore, authenticated
 * as your Firebase service account. Same JWT pattern as the FCM code,
 * just a different scope.
 */
function getFirestoreAccessToken_() {
  const props = PropertiesService.getScriptProperties();
  const serviceAccountJson = props.getProperty('FCM_SERVICE_ACCOUNT_JSON');
  if (!serviceAccountJson) {
    throw new Error('Missing FCM_SERVICE_ACCOUNT_JSON script property.');
  }
  const serviceAccount = JSON.parse(serviceAccountJson);

  const header = { alg: 'RS256', typ: 'JWT' };
  const now = Math.floor(Date.now() / 1000);
  const claimSet = {
    iss: serviceAccount.client_email,
    scope: 'https://www.googleapis.com/auth/datastore',
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
    throw new Error('Failed to get Firestore access token: ' + response.getContentText());
  }
  return result.access_token;
}

/**
 * Handles a "mark these messages as read" request from the app.
 * params: { messageIds: [docId, docId, ...] }
 */
function handleMarkMessagesRead(params) {
  try {
    const accessToken = getFirestoreAccessToken_();
    const ids = params.messageIds || [];

    ids.forEach(id => {
      const url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID +
        '/databases/(default)/documents/messages/' + id + '?updateMask.fieldPaths=read';

      const response = UrlFetchApp.fetch(url, {
        method: 'patch',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + accessToken },
        payload: JSON.stringify({ fields: { read: { booleanValue: true } } }),
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      if (code < 200 || code >= 300) {
        debugLog_('ERROR marking message ' + id + ' as read (' + code + '): ' + response.getContentText());
      }
    });

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    debugLog_('ERROR marking messages read: ' + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Writes one message document to the "messages" collection in Firestore
 * via the Firestore REST API.
 */
function saveMessageToFirestore_(messageData) {
  const accessToken = getFirestoreAccessToken_();
  const url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID +
    '/databases/(default)/documents/messages';

  // Firestore's REST API needs each field wrapped with its type, e.g.
  // {"stringValue": "..."} instead of just "...". This is just how
  // their REST format works, nothing to change here.
  const body = {
    fields: {
      from: { stringValue: messageData.from },
      to: { stringValue: messageData.to },
      body: { stringValue: messageData.body },
      channel: { stringValue: messageData.channel },
      direction: { stringValue: messageData.direction },
      read: { booleanValue: messageData.read },
      timestamp: { timestampValue: messageData.timestamp },
      twilioSid: { stringValue: messageData.twilioSid || '' }
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + accessToken },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  Logger.log('Firestore write response: ' + response.getContentText());

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Firestore write failed (' + code + '): ' + response.getContentText());
  }

  return response.getContentText();
}

/**
 * TEST FUNCTION. Run this manually from the Apps Script editor
 * (pick "testSaveMessage" from the function dropdown, click Run)
 * to confirm the Firestore connection works before wiring up Twilio.
 * Check the "Messages" tab in the app afterward - a test message
 * should appear.
 */
function testSaveMessage() {
  const result = saveMessageToFirestore_({
    from: '+15555550123',
    to: '+15555550199',
    body: 'This is a test message from Apps Script.',
    channel: 'sms',
    direction: 'inbound',
    read: false,
    timestamp: new Date().toISOString()
  });
  Logger.log(result);
}

/**
 * ====================================================================
 * PHASE 2 - Sending replies
 * ====================================================================
 * ADDITIONAL SETUP REQUIRED:
 *
 * 1. In Twilio Console, go to Account > API keys & tokens (or the
 *    Console dashboard) and find your Account SID and Auth Token.
 *
 * 2. In this Apps Script project's Script Properties (Project Settings
 *    > Script Properties), add:
 *       TWILIO_ACCOUNT_SID      = (your Account SID, starts with AC...)
 *       TWILIO_AUTH_TOKEN       = (your Auth Token)
 *       TWILIO_PHONE_NUMBER     = (your SMS number in E.164 format,
 *                                  e.g. +15517332381 - no spaces/dashes)
 *       TWILIO_WHATSAPP_NUMBER  = (your WhatsApp sender number in
 *                                  E.164 format, e.g. +16469833275 -
 *                                  this can be a different number than
 *                                  TWILIO_PHONE_NUMBER)
 *
 * 3. In doPost(e) in Code.gs, add this alongside your other action
 *    checks (after the savePushToken check, before the gigId logic):
 *
 *       if (params.action === 'sendMessage') {
 *         return handleSendMessage(params);
 *       }
 *
 * 4. Redeploy, same as always.
 * ====================================================================
 */

/**
 * Handles a "send a reply" request from the app's Messages tab.
 * params: { to, body, channel }
 */
function handleSendMessage(params) {
  try {
    const sendResult = sendTwilioMessage_(params.to, params.body, params.channel);

    const props = PropertiesService.getScriptProperties();
    const isWhatsapp = params.channel === 'whatsapp';
    const fromProperty = isWhatsapp ? 'TWILIO_WHATSAPP_NUMBER' : 'TWILIO_PHONE_NUMBER';

    // Save our own outbound message to Firestore too, so it shows up
    // in the thread the same way an incoming message does.
    saveMessageToFirestore_({
      from: props.getProperty(fromProperty) || '',
      to: params.to,
      body: params.body,
      channel: params.channel || 'sms',
      direction: 'outbound',
      read: true,
      timestamp: new Date().toISOString(),
      twilioSid: sendResult.sid || ''
    });

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    debugLog_('ERROR sending message: ' + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Sends an SMS or WhatsApp message via Twilio's REST API.
 * SMS sends from TWILIO_PHONE_NUMBER; WhatsApp sends from the
 * separately-registered TWILIO_WHATSAPP_NUMBER. Returns the parsed
 * Twilio response (includes .sid, the Twilio message ID).
 */
function sendTwilioMessage_(to, body, channel) {
  const props = PropertiesService.getScriptProperties();
  const accountSid = props.getProperty('TWILIO_ACCOUNT_SID');
  const authToken = props.getProperty('TWILIO_AUTH_TOKEN');
  const isWhatsapp = channel === 'whatsapp';
  const fromNumber = props.getProperty(isWhatsapp ? 'TWILIO_WHATSAPP_NUMBER' : 'TWILIO_PHONE_NUMBER');

  if (!accountSid || !authToken || !fromNumber) {
    throw new Error('Missing TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, or ' +
      (isWhatsapp ? 'TWILIO_WHATSAPP_NUMBER' : 'TWILIO_PHONE_NUMBER') + ' script property.');
  }

  const url = 'https://api.twilio.com/2010-04-01/Accounts/' + accountSid + '/Messages.json';

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(accountSid + ':' + authToken)
    },
    payload: {
      To: isWhatsapp ? 'whatsapp:' + to : to,
      From: isWhatsapp ? 'whatsapp:' + fromNumber : fromNumber,
      Body: body
    },
    muteHttpExceptions: true
  });

  Logger.log('Twilio send response: ' + response.getContentText());

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Twilio send failed (' + code + '): ' + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}

/**
 * TEST FUNCTION. Run manually from the Apps Script editor to confirm
 * sending works before testing from the app itself. Change toNumber
 * to your own phone number first.
 */
function testSendMessage() {
  const toNumber = '+15555550123'; // <-- change this to your real number
  handleSendMessage({ to: toNumber, body: 'Test reply from Barnstorm Gig Manager.', channel: 'sms' });
}

/**
 * ====================================================================
 * ADD LEAD - add a lead to the Messages tab WITHOUT sending a text
 * ====================================================================
 * Saves an empty placeholder entry so the lead shows up in the SMS list.
 * params: { to }
 *
 * ADDITIONAL SETUP REQUIRED:
 * In doPost(e) in Code.gs, add this right below the sendMessage check:
 *
 *       if (params.action === 'addLead') {
 *         return handleAddLead(params);
 *       }
 *
 * Then redeploy (Deploy > Manage deployments > pencil > New version).
 * ====================================================================
 */
function handleAddLead(params) {
  try {
    const to = String(params.to || '').trim();
    if (!/^\+\d{8,15}$/.test(to)) {
      throw new Error('Invalid phone number: ' + to);
    }

    saveMessageToFirestore_({
      from: '',
      to: to,
      body: '',
      channel: 'sms',
      direction: 'outbound',
      read: true,
      timestamp: new Date().toISOString(),
      twilioSid: ''
    });

    debugLog_('Manual lead added: ' + to);
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    debugLog_('ERROR adding lead: ' + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}