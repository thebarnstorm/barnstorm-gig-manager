/**
 * ====================================================================
 * MANUAL LEAD SYNC - Automatically add manually created Insightly leads
 * to the Messages tab (no text needed, no button tap needed).
 * ====================================================================
 * This is a NEW file. In the Apps Script editor: click the "+" next to
 * Files > Script, name it "ManualLeadSync", and paste this in.
 *
 * HOW IT WORKS: Every few minutes this looks at your Insightly leads,
 * finds recent ones with "Added Manually" checked, and for any whose
 * phone number has no messages yet, saves an empty placeholder entry
 * (the same thing the "+ Add Lead" button does).
 *
 * SETUP:
 * 1. Run previewManualLeadSync() once (pick it from the function
 *    dropdown, click Run). It only LOGS what it would add; it changes
 *    nothing. Check the "Debug Log" tab in your Master Doc.
 * 2. If the preview looks right, run syncManualLeads() once for real.
 * 3. Add a trigger: alarm clock icon > Add Trigger > function
 *    syncManualLeads > Time-driven > Minutes timer > Every 10 minutes.
 * ====================================================================
 */

// Only leads created within this many days are considered. This keeps
// old manual leads from all flooding into the Messages tab at once.
const MANUAL_LEAD_LOOKBACK_DAYS = 14;

function previewManualLeadSync() {
  runManualLeadSync_(true);
}

function syncManualLeads() {
  runManualLeadSync_(false);
}

function isChecked_(value) {
  return value === true || String(value).toLowerCase() === 'true';
}

/**
 * Returns a lookup of phone numbers (last 10 digits) that already have
 * at least one SMS entry in Firestore. Throws if there are too many
 * messages to check safely, so we never add duplicates by mistake.
 */
function fetchExistingSmsContactNumbers_() {
  const accessToken = getFirestoreAccessToken_();
  const numbers = {};
  let pageToken = null;
  const maxPages = 40; // safety cap (12,000 messages)

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
    (result.documents || []).forEach(doc => {
      const f = doc.fields || {};
      const channel = f.channel && f.channel.stringValue ? f.channel.stringValue : 'sms';
      if (channel !== 'sms') return;
      const direction = f.direction && f.direction.stringValue ? f.direction.stringValue : '';
      const raw = direction === 'outbound'
        ? (f.to && f.to.stringValue)
        : (f.from && f.from.stringValue);
      const key = normalizePhoneDigits_(raw);
      if (key) numbers[key] = true;
    });

    pageToken = result.nextPageToken;
    if (!pageToken) return numbers;
  }

  throw new Error('Too many messages to check safely; stopping so no duplicates are added.');
}

/**
 * Fetches Insightly leads (fresh, not cached) that have "Added Manually"
 * checked and were created within the lookback window.
 */
function fetchRecentManualInsightlyLeads_() {
  const baseUrl = getInsightlyBaseUrl_();
  const authHeader = getInsightlyAuthHeader_();
  const cutoff = Date.now() - MANUAL_LEAD_LOOKBACK_DAYS * 24 * 60 * 60 * 1000;

  let allLeads = [];
  let skip = 0;
  const pageSize = 500;
  const maxPages = 20; // same cap as the rest of the app (10,000 leads)

  for (let page = 0; page < maxPages; page++) {
    const response = UrlFetchApp.fetch(baseUrl + '/Leads?top=' + pageSize + '&skip=' + skip, {
      method: 'get',
      headers: { Authorization: authHeader },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Insightly leads fetch failed (' + response.getResponseCode() + '): ' + response.getContentText());
    }
    const batch = JSON.parse(response.getContentText());
    allLeads = allLeads.concat(batch);
    if (batch.length < pageSize) break;
    skip += pageSize;
  }

  let manualCount = 0;
  let missingDateCount = 0;
  const recent = [];

  allLeads.forEach(l => {
    if (!isChecked_(getCustomFieldValue_(l.CUSTOMFIELDS, 'Added_Manually__c'))) return;
    manualCount++;

    const created = l.DATE_CREATED_UTC ? new Date(l.DATE_CREATED_UTC.toString().replace(' ', 'T') + 'Z') : null;
    if (!created || isNaN(created.getTime())) {
      missingDateCount++;
      return;
    }
    if (created.getTime() >= cutoff) recent.push(l);
  });

  debugLog_('Manual lead sync: scanned ' + allLeads.length + ' leads, ' + manualCount +
    ' marked manual, ' + recent.length + ' within last ' + MANUAL_LEAD_LOOKBACK_DAYS + ' days' +
    (missingDateCount ? ' (' + missingDateCount + ' skipped for missing created date)' : '') + '.');

  return recent;
}

function runManualLeadSync_(dryRun) {
  try {
    const leads = fetchRecentManualInsightlyLeads_();
    if (leads.length === 0) {
      debugLog_('Manual lead sync: nothing to do.');
      return;
    }

    const existing = fetchExistingSmsContactNumbers_();
    const seenThisRun = {};
    let added = 0;

    leads.forEach(l => {
      const digits = normalizePhoneDigits_(l.MOBILE || l.PHONE);
      const name = [l.FIRST_NAME, l.LAST_NAME].filter(Boolean).join(' ') || '(no name)';

      if (digits.length !== 10) {
        debugLog_('Manual lead sync: skipped ' + name + ' (phone missing or not a 10-digit US number).');
        return;
      }
      if (existing[digits] || seenThisRun[digits]) return; // already in the app

      seenThisRun[digits] = true;
      const e164 = '+1' + digits;

      if (dryRun) {
        debugLog_('PREVIEW: would add ' + name + ' (' + e164 + ')');
      } else {
        saveMessageToFirestore_({
          from: '',
          to: e164,
          body: '',
          channel: 'sms',
          direction: 'outbound',
          read: true,
          timestamp: new Date().toISOString(),
          twilioSid: ''
        });
        debugLog_('Manual lead sync: added ' + name + ' (' + e164 + ')');
      }
      added++;
    });

    debugLog_('Manual lead sync ' + (dryRun ? '(PREVIEW) ' : '') + 'finished: ' + added + ' lead(s) ' + (dryRun ? 'would be ' : '') + 'added.');
  } catch (err) {
    debugLog_('ERROR in manual lead sync: ' + err.toString());
  }
}