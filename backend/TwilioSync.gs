/**
 * ====================================================================
 * TWILIO SYNC - Backfill outbound SMS messages sent through Twilio
 * by anything other than this app (e.g. your Zapier workflow).
 * ====================================================================
 * This is a NEW file to add to your existing Apps Script project.
 * In the Apps Script editor: click the "+" next to Files > Script,
 * name it "TwilioSync", and paste this in.
 *
 * WHY THIS EXISTS: Zapier's "Send SMS" action talks directly to
 * Twilio - it never passes through our Apps Script, so nothing tells
 * Firestore that message happened. This periodically asks Twilio
 * directly, "show me your outbound SMS," and records anything we
 * don't already have.
 *
 * SETUP REQUIRED:
 *
 * 1. Run syncOutboundSmsMessages() once manually (pick it from the
 *    function dropdown, click Run) to confirm it works and check the
 *    Debug Log tab for what it found.
 *
 * 2. Set up a recurring trigger so this runs automatically:
 *    In the Apps Script editor, click the alarm clock icon (Triggers)
 *    in the left sidebar > Add Trigger. Set:
 *      - Function to run: syncOutboundSmsMessages
 *      - Event source: Time-driven
 *      - Type: Minutes timer
 *      - Interval: Every 5 minutes (or whatever feels right)
 *    Click Save.
 *
 * NOTE: This only backfills SMS (not WhatsApp), since that's what
 * your Zapier workflow sends. It only looks at messages sent FROM
 * your Twilio SMS number, so it won't touch WhatsApp threads.
 * ====================================================================
 */

/**
 * Fetches the set of Twilio message SIDs already recorded in
 * Firestore, so we know what NOT to re-add. Paginated, same pattern
 * as other collection-wide fetches in this project.
 */
function fetchExistingTwilioSids_() {
  const accessToken = getFirestoreAccessToken_();
  const sids = {};
  let pageToken = null;
  const maxPages = 20; // safety cap

  for (let page = 0; page < maxPages; page++) {
    let url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID +
      '/databases/(default)/documents/messages?pageSize=300';
    if (pageToken) {
      url += '&pageToken=' + encodeURIComponent(pageToken);
    }

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Firestore messages fetch failed (' + response.getResponseCode() + '): ' + response.getContentText());
    }

    const result = JSON.parse(response.getContentText());
    const docs = result.documents || [];
    docs.forEach(doc => {
      const sid = doc.fields && doc.fields.twilioSid ? doc.fields.twilioSid.stringValue : '';
      if (sid) sids[sid] = true;
    });

    pageToken = result.nextPageToken;
    if (!pageToken) break;
  }

  return sids;
}

/**
 * Fetches recent messages sent FROM your Twilio SMS number, using
 * Twilio's own Messages list API. Returns up to `maxMessages`,
 * newest first.
 */
function fetchRecentOutboundTwilioSms_(maxMessages) {
  const props = PropertiesService.getScriptProperties();
  const accountSid = props.getProperty('TWILIO_ACCOUNT_SID');
  const authToken = props.getProperty('TWILIO_AUTH_TOKEN');
  const smsNumber = props.getProperty('TWILIO_PHONE_NUMBER');

  if (!accountSid || !authToken || !smsNumber) {
    throw new Error('Missing TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, or TWILIO_PHONE_NUMBER script property.');
  }

  const authHeader = 'Basic ' + Utilities.base64Encode(accountSid + ':' + authToken);
  let messages = [];
  let url = 'https://api.twilio.com/2010-04-01/Accounts/' + accountSid + '/Messages.json?From=' +
    encodeURIComponent(smsNumber) + '&PageSize=100';

  while (url && messages.length < maxMessages) {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: authHeader },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Twilio messages fetch failed (' + response.getResponseCode() + '): ' + response.getContentText());
    }

    const result = JSON.parse(response.getContentText());
    messages = messages.concat(result.messages || []);

    // Twilio gives a relative next-page path, not a full URL
    url = result.next_page_uri ? 'https://api.twilio.com' + result.next_page_uri : null;
  }

  return messages.slice(0, maxMessages);
}

/**
 * Main sync function. Finds outbound SMS from Twilio that aren't yet
 * in Firestore, and adds them.
 */
function syncOutboundSmsMessages() {
  try {
    const existingSids = fetchExistingTwilioSids_();
    const recentMessages = fetchRecentOutboundTwilioSms_(200);

    let addedCount = 0;
    recentMessages.forEach(msg => {
      // Only outbound messages that aren't already recorded
      if (msg.direction.indexOf('outbound') !== 0) return;
      if (existingSids[msg.sid]) return;

      saveMessageToFirestore_({
        from: msg.from,
        to: msg.to,
        body: msg.body || '',
        channel: 'sms',
        direction: 'outbound',
        read: true,
        timestamp: new Date(msg.date_created).toISOString(),
        twilioSid: msg.sid
      });
      addedCount++;
    });

    debugLog_('Twilio SMS sync: checked ' + recentMessages.length + ', added ' + addedCount + ' new outbound message(s).');
  } catch (err) {
    debugLog_('ERROR syncing outbound SMS: ' + err.toString());
  }
}