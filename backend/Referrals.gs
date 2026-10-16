/**
 * ====================================================================
 * REFERRALS - Broadcast a lead referral to your bandleader roster,
 * and track who responds.
 * ====================================================================
 * This is a NEW file to add to your existing Apps Script project.
 * In the Apps Script editor: click the "+" next to Files > Script,
 * name it "Referrals", and paste this in.
 *
 * SETUP REQUIRED:
 *
 * 1. Your Google Sheet needs a tab named exactly "Bandleader Roster"
 *    with columns: Name (A), WhatsApp Number (B) - one bandleader
 *    per row, numbers in +1XXXXXXXXXX format.
 *
 * 2. In doGet(e) in Code.gs, add this alongside the lookupContact
 *    check:
 *
 *       if (e.parameter.action === 'getRoster') {
 *         return handleGetRoster();
 *       }
 *
 * 3. In doPost(e) in Code.gs, add this alongside your other action
 *    checks:
 *
 *       if (params.action === 'createReferral') {
 *         return handleCreateReferral(params);
 *       }
 *
 * 4. In Messages.gs, inside handleIncomingTwilioMessage(params), after
 *    the line that calls saveMessageToFirestore_(messageData), add:
 *
 *       if (isWhatsapp) {
 *         checkAndUpdateReferralResponse_(from, messageData.body);
 *       }
 *
 * 5. Redeploy, same as always.
 *
 * 6. Add this Script Property (Project Settings > Script Properties):
 *       TWILIO_REFERRAL_CONTENT_SID = HXc70ae07d5ba5c40c9300b7330b3c1c1e
 *
 * NOTE ON RESPONSES: sending now uses your approved Quick Reply
 * template, so bandleaders see actual tappable buttons. We still
 * need to verify the exact webhook data Twilio sends when a button
 * is tapped before we can read it precisely - see testSendReferralTemplate
 * below for a safe way to test that with just yourself.
 * ====================================================================
 */

const BANDLEADER_ROSTER_SHEET_NAME = 'Bandleader Roster';

/**
 * Reads your bandleader roster from the Google Sheet. Reads by header
 * name (Whatsapp Group, Band, Bandleader, Cell) rather than fixed
 * column positions, so reordering columns in the sheet won't break
 * this.
 */
function getBandleaderRoster_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BANDLEADER_ROSTER_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet tab "' + BANDLEADER_ROSTER_SHEET_NAME + '" not found.');
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => h.toString().trim().toLowerCase());
  const groupCol = headers.indexOf('whatsapp group');
  const bandCol = headers.indexOf('band');
  const bandleaderCol = headers.indexOf('bandleader');
  const cellCol = headers.indexOf('cell');

  if (cellCol === -1) {
    throw new Error('Could not find a "Cell" column in the Bandleader Roster sheet.');
  }
  if (groupCol === -1) {
    throw new Error('Could not find a "Whatsapp Group" column in the Bandleader Roster sheet.');
  }

  const roster = [];
  for (let i = 1; i < values.length; i++) {
    const group = (values[i][groupCol] || '').toString().trim();
    const band = bandCol !== -1 ? (values[i][bandCol] || '').toString().trim() : '';
    const bandleaderName = bandleaderCol !== -1 ? (values[i][bandleaderCol] || '').toString().trim() : '';
    const phone = (values[i][cellCol] || '').toString().trim();

    if (!phone || !group) continue;

    const displayName = band && bandleaderName ? (band + ' (' + bandleaderName + ')') : (band || bandleaderName || phone);
    roster.push({ group, name: displayName, band, phone });
  }
  return roster;
}

/**
 * Returns the distinct list of WhatsApp group names in the roster,
 * in the order they first appear.
 */
function getRosterGroupNames_() {
  const roster = getBandleaderRoster_();
  const seen = [];
  roster.forEach(r => {
    if (seen.indexOf(r.group) === -1) seen.push(r.group);
  });
  return seen;
}

/**
 * Handles a request from the app for the list of roster groups.
 * Called from doGet(e) when e.parameter.action === 'getRosterGroups'.
 */
function handleGetRosterGroups() {
  try {
    const groups = getRosterGroupNames_();
    return ContentService.createTextOutput(JSON.stringify({ success: true, groups: groups }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    debugLog_('ERROR getting roster groups: ' + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles a request from the app for the current roster.
 * Called from doGet(e) when e.parameter.action === 'getRoster'.
 */
function handleGetRoster() {
  try {
    const roster = getBandleaderRoster_();
    return ContentService.createTextOutput(JSON.stringify({ success: true, roster: roster }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    debugLog_('ERROR getting roster: ' + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Builds the referral message text sent to each bandleader.
 */
function buildReferralMessageText_(details) {
  const lines = [
    '\uD83D\uDC4B New inquiry - any takers?',
    '',
    '\uD83D\uDCC5 Date: ' + (details.date || ''),
    '\uD83D\uDC70 Event: ' + (details.eventType || ''),
    '\uD83D\uDCCD Venue: ' + (details.venue || ''),
    '\uD83D\uDC65 Headcount: ' + (details.headcount || ''),
    '\uD83D\uDCB0 Budget: ' + (details.budget || ''),
    '\uD83D\uDCDD More Info: ' + (details.moreInfo || ''),
    '',
    'Let me know asap if interested and I\'ll get the note out.'
  ];
  return lines.join('\n');
}

/**
 * Sends the approved "barnstorm_referral" WhatsApp Quick Reply
 * template to one bandleader, filling in the {{1}}-{{6}} placeholders.
 * Requires TWILIO_REFERRAL_CONTENT_SID in Script Properties (the
 * Content SID, starts with HX..., from the approved template).
 */
function sendReferralTemplateMessage_(to, details) {
  const props = PropertiesService.getScriptProperties();
  const accountSid = props.getProperty('TWILIO_ACCOUNT_SID');
  const authToken = props.getProperty('TWILIO_AUTH_TOKEN');
  const fromNumber = props.getProperty('TWILIO_WHATSAPP_NUMBER');
  const contentSid = props.getProperty('TWILIO_REFERRAL_CONTENT_SID');

  if (!accountSid || !authToken || !fromNumber || !contentSid) {
    throw new Error('Missing TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, TWILIO_WHATSAPP_NUMBER, or TWILIO_REFERRAL_CONTENT_SID script property.');
  }

  // WhatsApp template variables can't contain newlines, tabs, or more
  // than 4 consecutive spaces - flatten anything like that to a single
  // space so multi-line "More Info" text doesn't break the send.
  const clean = (value) => (value || '').toString().replace(/[\n\t\r]+/g, ' ').replace(/ {2,}/g, ' ').trim();

  const contentVariables = JSON.stringify({
    '1': clean(details.date),
    '2': clean(details.eventType),
    '3': clean(details.venue),
    '4': clean(details.headcount),
    '5': clean(details.budget),
    '6': clean(details.moreInfo)
  });

  const url = 'https://api.twilio.com/2010-04-01/Accounts/' + accountSid + '/Messages.json';
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(accountSid + ':' + authToken)
    },
    payload: {
      To: 'whatsapp:' + to,
      From: 'whatsapp:' + fromNumber,
      ContentSid: contentSid,
      ContentVariables: contentVariables
    },
    muteHttpExceptions: true
  });

  Logger.log('Twilio template send response: ' + response.getContentText());

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Twilio template send failed (' + code + '): ' + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}

/**
 * Handles a "broadcast this referral to selected WhatsApp groups"
 * request. Sends to every bandleader in the roster whose "Whatsapp
 * Group" matches one of the selected groups.
 * params: { details: {...}, groups: ["Golden Era", "New England", ...] }
 */
function handleCreateReferral(params) {
  try {
    const messageText = buildReferralMessageText_(params.details || {}); // used for the thread/Firestore copy only
    const now = new Date().toISOString();
    const insightlyLeadId = params.insightlyLeadId || '';

    // Create the parent referral record
    const referralId = saveReferralToFirestore_({
      details: params.details || {},
      createdAt: now,
      insightlyLeadId: insightlyLeadId
    });

    // Sending this out means The Barnstorm isn't taking the gig -
    // remove it from the Bands__c list, and clear the separate
    // Band__c field too in case anything else relies on it.
    if (insightlyLeadId) {
      try {
        removeBandFromInsightlyLead_(insightlyLeadId, 'The Barnstorm');
      } catch (err) {
        debugLog_('ERROR removing The Barnstorm from Bands__c for lead ' + insightlyLeadId + ': ' + err.toString());
      }
      try {
        clearBandFieldOnInsightlyLead_(insightlyLeadId);
      } catch (err) {
        debugLog_('ERROR clearing Band__c for lead ' + insightlyLeadId + ': ' + err.toString());
      }
    }

    const selectedPhones = params.individualPhones || [];
    const recipients = getBandleaderRoster_().filter(r => selectedPhones.indexOf(r.phone) !== -1);
    const results = [];
    recipients.forEach(r => {
      let sendError = null;
      let sendResult = null;
      try {
        sendResult = sendReferralTemplateMessage_(r.phone, params.details || {});
      } catch (err) {
        sendError = err.toString();
        debugLog_('ERROR sending referral to ' + r.phone + ': ' + sendError);
      }

      saveReferralRecipientToFirestore_({
        referralId: referralId,
        name: r.name,
        band: r.band,
        phone: r.phone,
        status: sendError ? 'failed' : 'pending',
        sentAt: now,
        respondedAt: null,
        responseText: '',
        sentMessageSid: sendResult && sendResult.sid ? sendResult.sid : '',
        insightlyLeadId: insightlyLeadId
      });

      // Also log it as an outbound WhatsApp message so it shows in
      // that bandleader's normal conversation thread too.
      if (!sendError) {
        saveMessageToFirestore_({
          from: PropertiesService.getScriptProperties().getProperty('TWILIO_WHATSAPP_NUMBER') || '',
          to: r.phone,
          body: messageText,
          channel: 'whatsapp',
          direction: 'outbound',
          read: true,
          timestamp: now,
          twilioSid: sendResult && sendResult.sid ? sendResult.sid : ''
        });
      }

      results.push({ name: r.name, phone: r.phone, success: !sendError });
    });

    return ContentService.createTextOutput(JSON.stringify({ success: true, referralId: referralId, results: results }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    debugLog_('ERROR creating referral: ' + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Writes the parent referral record to Firestore. Returns its ID.
 */
function saveReferralToFirestore_(referral) {
  const accessToken = getFirestoreAccessToken_();
  const url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID + '/databases/(default)/documents/referrals';

  const d = referral.details;
  const body = {
    fields: {
      date: { stringValue: d.date || '' },
      eventType: { stringValue: d.eventType || '' },
      venue: { stringValue: d.venue || '' },
      headcount: { stringValue: d.headcount || '' },
      budget: { stringValue: d.budget || '' },
      moreInfo: { stringValue: d.moreInfo || '' },
      createdAt: { timestampValue: referral.createdAt },
      insightlyLeadId: { stringValue: String(referral.insightlyLeadId || '') }
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + accessToken },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Firestore referral write failed (' + code + '): ' + response.getContentText());
  }

  // Extract the auto-generated document ID from the response name field,
  // e.g. ".../documents/referrals/AbCdEf123" -> "AbCdEf123"
  const result = JSON.parse(response.getContentText());
  const parts = result.name.split('/');
  return parts[parts.length - 1];
}

/**
 * Writes one recipient tracking record to Firestore.
 */
function saveReferralRecipientToFirestore_(recipient) {
  const accessToken = getFirestoreAccessToken_();
  const url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID + '/databases/(default)/documents/referralRecipients';

  const body = {
    fields: {
      referralId: { stringValue: recipient.referralId },
      name: { stringValue: recipient.name },
      band: { stringValue: recipient.band || '' },
      phone: { stringValue: recipient.phone },
      status: { stringValue: recipient.status },
      sentAt: { timestampValue: recipient.sentAt },
      respondedAt: recipient.respondedAt ? { timestampValue: recipient.respondedAt } : { nullValue: null },
      responseText: { stringValue: recipient.responseText || '' },
      sentMessageSid: { stringValue: recipient.sentMessageSid || '' },
      insightlyLeadId: { stringValue: String(recipient.insightlyLeadId || '') }
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + accessToken },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Firestore referral recipient write failed (' + code + '): ' + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());
  const parts = result.name.split('/');
  return parts[parts.length - 1];
}

/**
 * Updates a recipient's Firestore document by its document ID.
 */
function updateReferralRecipientInFirestore_(docId, updates) {
  const accessToken = getFirestoreAccessToken_();
  const fieldPaths = Object.keys(updates);
  const updateMaskParams = fieldPaths.map(f => 'updateMask.fieldPaths=' + encodeURIComponent(f)).join('&');
  const url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID +
    '/databases/(default)/documents/referralRecipients/' + docId + '?' + updateMaskParams;

  const fields = {};
  fieldPaths.forEach(f => {
    const v = updates[f];
    if (v === null) {
      fields[f] = { nullValue: null };
    } else if (f === 'respondedAt') {
      fields[f] = { timestampValue: v };
    } else {
      fields[f] = { stringValue: v };
    }
  });

  const response = UrlFetchApp.fetch(url, {
    method: 'patch',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + accessToken },
    payload: JSON.stringify({ fields: fields }),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Firestore referral recipient update failed (' + code + '): ' + response.getContentText());
  }
}

/**
 * Fetches all referralRecipients documents (paginated). Used to find
 * a pending referral matching an incoming reply.
 */
function fetchAllReferralRecipients_() {
  const accessToken = getFirestoreAccessToken_();
  let all = [];
  let pageToken = null;
  const maxPages = 10; // safety cap

  for (let page = 0; page < maxPages; page++) {
    let url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID +
      '/databases/(default)/documents/referralRecipients?pageSize=300';
    if (pageToken) {
      url += '&pageToken=' + encodeURIComponent(pageToken);
    }

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Firestore referralRecipients fetch failed (' + response.getResponseCode() + '): ' + response.getContentText());
    }

    const result = JSON.parse(response.getContentText());
    const docs = result.documents || [];
    docs.forEach(doc => {
      const parts = doc.name.split('/');
      const id = parts[parts.length - 1];
      const f = doc.fields || {};
      all.push({
        id: id,
        referralId: f.referralId ? f.referralId.stringValue : '',
        name: f.name ? f.name.stringValue : '',
        phone: f.phone ? f.phone.stringValue : '',
        status: f.status ? f.status.stringValue : '',
        sentAt: f.sentAt ? f.sentAt.timestampValue : '',
        sentMessageSid: f.sentMessageSid ? f.sentMessageSid.stringValue : '',
        band: f.band ? f.band.stringValue : '',
        insightlyLeadId: f.insightlyLeadId ? f.insightlyLeadId.stringValue : ''
      });
    });

    pageToken = result.nextPageToken;
    if (!pageToken) break;
  }

  return all;
}

/**
 * When an inbound WhatsApp message arrives and is a genuine button
 * tap on a referral template (MessageType 'button'), this matches it
 * precisely to the exact referral message it replied to, using
 * Twilio's OriginalRepliedMessageSid - no guessing needed.
 */
function checkAndUpdateReferralResponseFromButton_(originalMessageSid, buttonPayload, buttonText) {
  try {
    const all = fetchAllReferralRecipients_();
    const match = all.find(r => r.sentMessageSid === originalMessageSid);
    if (!match) {
      debugLog_('Button tap did not match any referral recipient (sid ' + originalMessageSid + ').');
      return;
    }

    const status = buttonPayload === 'accept' ? 'accepted' : (buttonPayload === 'decline' ? 'declined' : 'responded');

    updateReferralRecipientInFirestore_(match.id, {
      status: status,
      respondedAt: new Date().toISOString(),
      responseText: buttonText || buttonPayload || ''
    });

    debugLog_('Referral button response recorded: ' + match.name + ' -> ' + status);

    // If accepted and this referral was tied to a specific Insightly
    // lead, add this band to that lead's Bands__c field automatically.
    if (status === 'accepted' && match.insightlyLeadId && match.band) {
      try {
        appendBandToInsightlyLead_(match.insightlyLeadId, match.band);
      } catch (err) {
        debugLog_('ERROR updating Insightly Bands__c for lead ' + match.insightlyLeadId + ': ' + err.toString());
      }
    }
  } catch (err) {
    debugLog_('ERROR checking button referral response: ' + err.toString());
  }
}

/**
 * Fallback for when a bandleader types a reply instead of tapping a
 * button. Matches by phone number + most recently sent pending offer,
 * since there's no message-level reference to go on here.
 */
function checkAndUpdateReferralResponse_(fromPhone, bodyText) {
  try {
    const target = normalizePhoneDigits_(fromPhone);
    const all = fetchAllReferralRecipients_();

    const pendingMatches = all.filter(r => r.status === 'pending' && normalizePhoneDigits_(r.phone) === target);
    if (pendingMatches.length === 0) return; // not a referral reply

    // Most recently sent pending offer for this bandleader
    pendingMatches.sort((a, b) => new Date(b.sentAt).getTime() - new Date(a.sentAt).getTime());
    const match = pendingMatches[0];

    updateReferralRecipientInFirestore_(match.id, {
      status: 'responded',
      respondedAt: new Date().toISOString(),
      responseText: bodyText || ''
    });

    debugLog_('Referral response recorded for ' + fromPhone + ': "' + bodyText + '"');
  } catch (err) {
    debugLog_('ERROR checking referral response: ' + err.toString());
  }
}

/**
 * Deletes a single Firestore document by full collection + ID.
 */
function deleteFirestoreDoc_(collectionName, docId) {
  const accessToken = getFirestoreAccessToken_();
  const url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID +
    '/databases/(default)/documents/' + collectionName + '/' + docId;

  const response = UrlFetchApp.fetch(url, {
    method: 'delete',
    headers: { Authorization: 'Bearer ' + accessToken },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Firestore delete failed (' + code + '): ' + response.getContentText());
  }
}

/**
 * Handles a "clear this referral" request - deletes the referral
 * record and all of its recipient tracking records.
 * params: { referralId }
 */
function handleDeleteReferral(params) {
  try {
    const referralId = params.referralId;
    if (!referralId) {
      throw new Error('Missing referralId.');
    }

    const allRecipients = fetchAllReferralRecipients_();
    const toDelete = allRecipients.filter(r => r.referralId === referralId);
    toDelete.forEach(r => deleteFirestoreDoc_('referralRecipients', r.id));

    deleteFirestoreDoc_('referrals', referralId);

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    debugLog_('ERROR deleting referral: ' + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * TEST FUNCTION. Run manually to send yourself a real templated
 * referral (with actual tappable buttons), so you can verify sending
 * works and then tap a button to see what Twilio sends back.
 * Change toNumber to your own phone number first.
 */
function testSendReferralTemplate() {
  const toNumber = '+15555550123'; // <-- change this to your real number

  const result = sendReferralTemplateMessage_(toNumber, {
    date: 'Sat. 7/24/27',
    eventType: 'Wedding',
    venue: 'Test Venue',
    headcount: '150',
    budget: '$15,000-$20,000',
    moreInfo: 'This is a test referral.'
  });
  Logger.log(JSON.stringify(result));
}


// ====================================================================
// REFERRAL ARCHIVING
// Archives a referral (instead of deleting it) and auto-archives
// referrals once "Send Referral" is ticked on the Insightly lead.
// ====================================================================

/**
 * Reads all documents in the "referrals" collection (paginated).
 * Returns simple objects: { id, insightlyLeadId, archived, restoredManually }
 */
function fetchAllReferrals_() {
  const accessToken = getFirestoreAccessToken_();
  const all = [];
  let pageToken = null;
  const maxPages = 10; // safety cap

  for (let page = 0; page < maxPages; page++) {
    let url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID +
      '/databases/(default)/documents/referrals?pageSize=300';
    if (pageToken) {
      url += '&pageToken=' + encodeURIComponent(pageToken);
    }

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Firestore referrals fetch failed (' + response.getResponseCode() + '): ' + response.getContentText());
    }

    const result = JSON.parse(response.getContentText());
    const docs = result.documents || [];
    docs.forEach(doc => {
      const parts = doc.name.split('/');
      const id = parts[parts.length - 1];
      const f = doc.fields || {};
      all.push({
        id: id,
        insightlyLeadId: f.insightlyLeadId ? (f.insightlyLeadId.stringValue || String(f.insightlyLeadId.integerValue || '')) : '',
        archived: f.archived ? f.archived.booleanValue === true : false,
        restoredManually: f.restoredManually ? f.restoredManually.booleanValue === true : false
      });
    });

    pageToken = result.nextPageToken || null;
    if (!pageToken) break;
  }

  return all;
}

/**
 * Marks a referral as archived (or restored) in Firestore.
 * Archiving records a timestamp. Restoring flags it so the timer leaves it alone.
 */
function setReferralArchivedInFirestore_(referralId, archived) {
  const accessToken = getFirestoreAccessToken_();
  const fields = { archived: { booleanValue: archived } };
  const maskFields = ['archived'];

  if (archived) {
    fields.archivedAt = { timestampValue: new Date().toISOString() };
    maskFields.push('archivedAt');
  } else {
    fields.restoredManually = { booleanValue: true };
    maskFields.push('restoredManually');
  }

  // currentDocument.exists=true makes this fail (instead of creating a stray
  // record) if the referral ID doesn't exist.
  const url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID +
    '/databases/(default)/documents/referrals/' + referralId +
    '?currentDocument.exists=true&' +
    maskFields.map(m => 'updateMask.fieldPaths=' + encodeURIComponent(m)).join('&');

  const response = UrlFetchApp.fetch(url, {
    method: 'patch',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + accessToken },
    payload: JSON.stringify({ fields: fields }),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Firestore referral update failed (' + code + '): ' + response.getContentText());
  }
}

/**
 * Handles "archiveReferral" and "restoreReferral" requests from the app.
 * params: { referralId }
 */
function handleSetReferralArchived(params, archived) {
  try {
    const referralId = params.referralId;
    if (!referralId) {
      throw new Error('Missing referralId.');
    }
    setReferralArchivedInFirestore_(referralId, archived);
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    debugLog_('ERROR ' + (archived ? 'archiving' : 'restoring') + ' referral: ' + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Looks at one Insightly lead and returns true if "Send Referral" is ticked.
 */
function isSendReferralChecked_(leadId) {
  const url = getInsightlyBaseUrl_() + '/Leads/' + leadId;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: getInsightlyAuthHeader_() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Insightly lead fetch failed (' + response.getResponseCode() + ') for lead ' + leadId);
  }
  const lead = JSON.parse(response.getContentText());
  const value = getCustomFieldValue_(lead.CUSTOMFIELDS, 'Send_Referral__c');
  return value === true || String(value).toLowerCase() === 'true';
}

/**
 * RUNS ON A TIMER. Archives any active referral whose lead has
 * "Send Referral" ticked in Insightly. Never deletes anything.
 */
function autoArchiveSentReferrals() {
  try {
    const candidates = fetchAllReferrals_()
      .filter(r => !r.archived && !r.restoredManually && r.insightlyLeadId);

    candidates.forEach(r => {
      try {
        if (isSendReferralChecked_(r.insightlyLeadId)) {
          setReferralArchivedInFirestore_(r.id, true);
          debugLog_('Auto-archived referral ' + r.id + ' (Send Referral ticked on lead ' + r.insightlyLeadId + ')');
        }
      } catch (err) {
        debugLog_('ERROR auto-archive check for referral ' + r.id + ': ' + err.toString());
      }
    });
  } catch (err) {
    debugLog_('ERROR in autoArchiveSentReferrals: ' + err.toString());
  }
}

/**
 * RUN ONCE to set up the 10-minute timer. Safe to run again (replaces the old timer).
 */
function installAutoArchiveTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoArchiveSentReferrals') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('autoArchiveSentReferrals').timeBased().everyMinutes(10).create();
}

/**
 * TEST FUNCTION. Run manually. Logs what Insightly returns for
 * "Send Referral" on each referral's lead, so we can confirm the value format.
 */
function testSendReferralField() {
  fetchAllReferrals_().forEach(r => {
    if (!r.insightlyLeadId) return;
    const url = getInsightlyBaseUrl_() + '/Leads/' + r.insightlyLeadId;
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: getInsightlyAuthHeader_() },
      muteHttpExceptions: true
    });
    const lead = JSON.parse(response.getContentText());
    const raw = getCustomFieldValue_(lead.CUSTOMFIELDS, 'Send_Referral__c');
    const msg = 'Referral ' + r.id + ' lead ' + r.insightlyLeadId +
      ' Send_Referral__c raw value: ' + JSON.stringify(raw) + ' (type ' + typeof raw + ')' +
      ' -> counted as ticked: ' + isSendReferralChecked_(r.insightlyLeadId);
    Logger.log(msg);
    debugLog_(msg);
  });
}