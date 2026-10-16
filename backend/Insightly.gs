/**
 * ====================================================================
 * INSIGHTLY - Contact lookup by phone number
 * ====================================================================
 * This is a NEW file to add to your existing Apps Script project.
 * In the Apps Script editor: click the "+" next to Files > Script,
 * name it "Insightly", and paste this in.
 *
 * ====================================================================
 */

function getInsightlyAuthHeader_() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('INSIGHTLY_API_KEY');
  if (!apiKey) {
    throw new Error('Missing INSIGHTLY_API_KEY script property.');
  }
  // Insightly uses Basic auth: API key as username, blank password.
  return 'Basic ' + Utilities.base64Encode(apiKey + ':');
}

function getInsightlyBaseUrl_() {
  const pod = PropertiesService.getScriptProperties().getProperty('INSIGHTLY_POD');
  if (!pod) {
    throw new Error('Missing INSIGHTLY_POD script property.');
  }
  return 'https://api.' + pod + '.insightly.com/v3.1';
}

/**
 * DIAGNOSTIC TEST FUNCTION. Run this manually first. It fetches your
 * first 3 contacts and logs their full raw JSON to the Debug Log tab,
 * so we can see the real field names for phone numbers before writing
 * the actual search-by-phone logic.
 */
/**
 * DIAGNOSTIC TEST FUNCTION #2. Run this to see the FULL picture for
 * one specific contact - including custom fields and any linked
 * Leads/Opportunities, which is likely where fields like Event Date,
 * Venue, and Budget actually live rather than on the Contact itself.
 * Change phoneToTest below if needed.
 */
function testInsightlyFullLookup() {
  const phoneToTest = '2525590231'; // <-- your test contact's number

  const url = getInsightlyBaseUrl_() + '/Contacts/Search?field_name=PHONE&field_value=' + encodeURIComponent(phoneToTest);
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: getInsightlyAuthHeader_() },
    muteHttpExceptions: true
  });

  debugLog_('Full lookup - Contact search HTTP ' + response.getResponseCode());
  const contactText = response.getContentText();
  logInChunks_('Full lookup - Contact', contactText);

  const contacts = JSON.parse(contactText);
  if (!contacts || contacts.length === 0) {
    debugLog_('Full lookup - no contact found for that phone number.');
    return;
  }

  const contact = contacts[0];
  const links = contact.LINKS || [];
  debugLog_('Full lookup - found ' + links.length + ' linked record(s): ' +
    links.map(l => l.LINK_OBJECT_NAME + '#' + l.LINK_OBJECT_ID).join(', '));

  // Fetch full detail on each linked Opportunity or Lead, since that's
  // likely where event-specific custom fields are stored.
  links.forEach(link => {
    if (link.LINK_OBJECT_NAME === 'Opportunity' || link.LINK_OBJECT_NAME === 'Lead') {
      const detailUrl = getInsightlyBaseUrl_() + '/' + link.LINK_OBJECT_NAME + 's/' + link.LINK_OBJECT_ID;
      const detailResponse = UrlFetchApp.fetch(detailUrl, {
        method: 'get',
        headers: { Authorization: getInsightlyAuthHeader_() },
        muteHttpExceptions: true
      });
      logInChunks_('Full lookup - ' + link.LINK_OBJECT_NAME + ' #' + link.LINK_OBJECT_ID, detailResponse.getContentText());
    }
  });
}

function logInChunks_(label, text) {
  const chunkSize = 45000;
  for (let i = 0; i < text.length; i += chunkSize) {
    debugLog_(label + ' [' + i + ']: ' + text.substring(i, i + chunkSize));
  }
}

/**
 * DIAGNOSTIC TEST FUNCTION #3. Fetches a few real Leads directly
 * (not Contacts) to see their full structure, including custom
 * fields like Event Date, Venue, and Budget.
 */
/**
 * DIAGNOSTIC TEST FUNCTION #5. Fetches one specific lead by phone
 * number and logs its raw Bands__c value, so we know the exact
 * format Insightly uses (plain text, newline-separated, etc.) before
 * writing code that updates it.
 */
function testInsightlyLeadBandsField() {
  const phoneToTest = '5164593497'; // <-- Madison Roth, from the screenshot

  const url = getInsightlyBaseUrl_() + '/Leads/Search?field_name=MOBILE&field_value=' + encodeURIComponent(phoneToTest);
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: getInsightlyAuthHeader_() },
    muteHttpExceptions: true
  });

  debugLog_('Bands field test - HTTP ' + response.getResponseCode());
  logInChunks_('Bands field test', response.getContentText());
}

function testInsightlyLeads() {
  const url = getInsightlyBaseUrl_() + '/Leads?top=5';
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: getInsightlyAuthHeader_() },
    muteHttpExceptions: true
  });

  debugLog_('Leads test - HTTP ' + response.getResponseCode());
  logInChunks_('Leads test', response.getContentText());
}

/**
 * DIAGNOSTIC TEST FUNCTION #4. Lists all custom field definitions for
 * Leads, so we can find the exact internal FIELD_NAME behind a label
 * you see in the Insightly UI (like "Guest Headcount Range").
 */
function testInsightlyCustomFields() {
  const url = getInsightlyBaseUrl_() + '/CustomFields';
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: getInsightlyAuthHeader_() },
    muteHttpExceptions: true
  });

  debugLog_('Custom fields - HTTP ' + response.getResponseCode());

  if (response.getResponseCode() === 200) {
    const fields = JSON.parse(response.getContentText());
    const leadFields = fields.filter(f => f.FIELD_FOR === 'LEAD');
    debugLog_('Custom fields - found ' + leadFields.length + ' Lead custom fields: ' +
      leadFields.map(f => f.FIELD_LABEL + ' => ' + f.FIELD_NAME).join(' | '));
  } else {
    logInChunks_('Custom fields - error body', response.getContentText());
  }
}

function testInsightlyContacts() {
  const url = getInsightlyBaseUrl_() + '/Contacts?top=3';

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: getInsightlyAuthHeader_() },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  debugLog_('Insightly test - HTTP ' + code);
  // Sheet cells have a size limit, so split long output across rows if needed.
  const chunkSize = 45000;
  for (let i = 0; i < text.length; i += chunkSize) {
    debugLog_('Insightly test body [' + i + ']: ' + text.substring(i, i + chunkSize));
  }

  Logger.log('HTTP ' + code);
  Logger.log(text);
}

/**
 * ====================================================================
 * Lead lookup by phone number (real implementation)
 * ====================================================================
 * Real prospective leads live as Insightly LEADS (not Contacts), with
 * their phone number in the MOBILE field (sometimes PHONE), and event
 * details as custom fields: Event_Date__c, Event_Type__c, Venue__c,
 * CityState__c, Budget__c.
 *
 * Insightly stores phone numbers as bare digit strings with no
 * formatting or country code, while Twilio sends full E.164 numbers
 * (e.g. +12525590231). This normalizes both to their last 10 digits
 * before comparing, so formatting differences don't cause a miss.
 *
 * This fetches your lead list (paginated, up to 2500 leads) and
 * matches client-side rather than using Insightly's Search endpoint,
 * since that endpoint requires an exact-format match and phone
 * formatting is inconsistent across records.
 *
 * Results are cached for 10 minutes to avoid re-fetching your whole
 * lead list every time you open a conversation.
 * ====================================================================
 */

const INSIGHTLY_LEADS_CACHE_KEY = 'insightly_leads_v1';
const INSIGHTLY_SOURCES_CACHE_KEY = 'insightly_lead_sources_v1';
const INSIGHTLY_CACHE_TTL = 600; // 10 minutes

function normalizePhoneDigits_(phone) {
  if (!phone) return '';
  const digits = phone.toString().replace(/\D/g, '');
  return digits.length > 10 ? digits.slice(-10) : digits;
}

function getCustomFieldValue_(customFields, fieldId) {
  const field = (customFields || []).find(f => f.FIELD_NAME === fieldId || f.CUSTOM_FIELD_ID === fieldId);
  return field ? field.FIELD_VALUE : '';
}

function formatInsightlyDate_(rawDate) {
  if (!rawDate) return '';
  try {
    // Insightly dates look like "2018-06-16 00:00:00" (UTC, date-only meaning)
    const isoish = rawDate.toString().replace(' ', 'T') + 'Z';
    const d = new Date(isoish);
    if (isNaN(d.getTime())) return rawDate.toString();
    return d.toLocaleDateString('en-US', { timeZone: 'UTC', year: 'numeric', month: 'long', day: 'numeric' });
  } catch (err) {
    return rawDate.toString();
  }
}

/**
 * Maps LEAD_SOURCE_ID to a human-readable name (e.g. "Website Contact Form").
 */
function fetchInsightlyLeadSourceMap_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(INSIGHTLY_SOURCES_CACHE_KEY);
  if (cached) return JSON.parse(cached);

  const url = getInsightlyBaseUrl_() + '/LeadSources';
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: getInsightlyAuthHeader_() },
    muteHttpExceptions: true
  });

  let map = {};
  if (response.getResponseCode() === 200) {
    const sources = JSON.parse(response.getContentText());
    sources.forEach(s => {
      map[s.LEAD_SOURCE_ID] = s.LEAD_SOURCE;
    });
    try {
      cache.put(INSIGHTLY_SOURCES_CACHE_KEY, JSON.stringify(map), 3600); // 1 hour, rarely changes
    } catch (err) {
      Logger.log('Lead source cache put error: ' + err);
    }
  } else {
    Logger.log('Could not fetch LeadSources (' + response.getResponseCode() + '): ' + response.getContentText());
  }
  return map;
}

function fetchAllInsightlyLeads_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(INSIGHTLY_LEADS_CACHE_KEY);
  if (cached) {
    return JSON.parse(cached);
  }

  const baseUrl = getInsightlyBaseUrl_();
  const authHeader = getInsightlyAuthHeader_();
  let allLeads = [];
  let skip = 0;
  const pageSize = 500;
  const maxPages = 20; // safety cap - 10,000 leads

  for (let page = 0; page < maxPages; page++) {
    const url = baseUrl + '/Leads?top=' + pageSize + '&skip=' + skip;
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: authHeader },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Insightly leads fetch failed (' + response.getResponseCode() + '): ' + response.getContentText());
    }
    const batch = JSON.parse(response.getContentText());
    allLeads = allLeads.concat(batch);
    if (batch.length < pageSize) break; // reached the last page
    skip += pageSize;
  }

  // Keep only what we need, to stay well under the cache's size limit
  const slim = allLeads.map(l => ({
    id: l.LEAD_ID,
    firstName: l.FIRST_NAME,
    lastName: l.LAST_NAME,
    email: l.EMAIL,
    brideName: getCustomFieldValue_(l.CUSTOMFIELDS, 'Brides_Name__c'),
    groomName: getCustomFieldValue_(l.CUSTOMFIELDS, 'Grooms_Name__c'),
    phones: [l.MOBILE, l.PHONE].filter(Boolean),
    eventDate: getCustomFieldValue_(l.CUSTOMFIELDS, 'Event_Date__c'),
    eventType: getCustomFieldValue_(l.CUSTOMFIELDS, 'Event_Type__c'),
    venue: getCustomFieldValue_(l.CUSTOMFIELDS, 'Venue__c'),
    cityState: getCustomFieldValue_(l.CUSTOMFIELDS, 'CityState__c'),
    budget: getCustomFieldValue_(l.CUSTOMFIELDS, 'Budget__c'),
    headcountRange: getCustomFieldValue_(l.CUSTOMFIELDS, 'Guest_Headcount_Range__c'),
    leadOrigin: getCustomFieldValue_(l.CUSTOMFIELDS, 'Lead_Origin__c'),
    description: l.LEAD_DESCRIPTION || '',
    otherInfo: getCustomFieldValue_(l.CUSTOMFIELDS, 'Other_Info__c')
  }));

  try {
    cache.put(INSIGHTLY_LEADS_CACHE_KEY, JSON.stringify(slim), INSIGHTLY_CACHE_TTL);
  } catch (err) {
    // If the lead list is too large to cache, we just skip caching -
    // lookups will be a bit slower but still work correctly.
    Logger.log('Insightly cache put error (continuing without cache): ' + err);
  }

  return slim;
}

function findInsightlyLeadByPhone_(phone) {
  const target = normalizePhoneDigits_(phone);
  if (!target) return null;

  const leads = fetchAllInsightlyLeads_();
  for (const l of leads) {
    for (const p of l.phones) {
      if (normalizePhoneDigits_(p) === target) {
        return l;
      }
    }
  }
  return null;
}

/**
 * Handles a "look up this phone number in Insightly" request.
 * Called from doGet(e) when e.parameter.action === 'lookupContact'.
 */
function handleLookupContact(phone) {
  try {
    const match = findInsightlyLeadByPhone_(phone);
    const pod = PropertiesService.getScriptProperties().getProperty('INSIGHTLY_POD');

    if (!match) {
      return ContentService.createTextOutput(JSON.stringify({ found: false }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const result = {
      found: true,
      leadId: match.id,
      name: [match.firstName, match.lastName].filter(Boolean).join(' ') || '(no name)',
      email: match.email || '',
      brideName: match.brideName || '',
      groomName: match.groomName || '',
      mobile: match.phones[0] || '',
      leadOrigin: match.leadOrigin || '',
      eventDate: formatInsightlyDate_(match.eventDate),
      eventType: match.eventType || '',
      venue: match.venue || '',
      cityState: match.cityState || '',
      budget: match.budget || '',
      headcountRange: match.headcountRange || '',
      description: match.description || '',
      otherInfo: match.otherInfo || '',
      insightlyUrl: 'https://crm.' + pod + '.insightly.com/details/Lead/' + match.id
    };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    debugLog_('ERROR looking up Insightly lead: ' + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ found: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles a "look up this lead directly by its Insightly ID" request.
 * Used by the Referrals list, which knows the lead's ID (saved when
 * the referral was created) but has no phone number of its own to
 * search by. Fetches the lead directly rather than searching the
 * whole cached list.
 * Called from doGet(e) when e.parameter.action === 'lookupContactById'.
 */
function handleLookupContactById(leadId) {
  try {
    if (!leadId) {
      return ContentService.createTextOutput(JSON.stringify({ found: false }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const url = getInsightlyBaseUrl_() + '/Leads/' + leadId;
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: getInsightlyAuthHeader_() },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      return ContentService.createTextOutput(JSON.stringify({ found: false }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const l = JSON.parse(response.getContentText());
    const pod = PropertiesService.getScriptProperties().getProperty('INSIGHTLY_POD');

    const result = {
      found: true,
      leadId: l.LEAD_ID,
      name: [l.FIRST_NAME, l.LAST_NAME].filter(Boolean).join(' ') || '(no name)',
      email: l.EMAIL || '',
      brideName: getCustomFieldValue_(l.CUSTOMFIELDS, 'Brides_Name__c'),
      groomName: getCustomFieldValue_(l.CUSTOMFIELDS, 'Grooms_Name__c'),
      mobile: l.MOBILE || l.PHONE || '',
      leadOrigin: getCustomFieldValue_(l.CUSTOMFIELDS, 'Lead_Origin__c'),
      eventDate: formatInsightlyDate_(getCustomFieldValue_(l.CUSTOMFIELDS, 'Event_Date__c')),
      eventType: getCustomFieldValue_(l.CUSTOMFIELDS, 'Event_Type__c'),
      venue: getCustomFieldValue_(l.CUSTOMFIELDS, 'Venue__c'),
      cityState: getCustomFieldValue_(l.CUSTOMFIELDS, 'CityState__c'),
      budget: getCustomFieldValue_(l.CUSTOMFIELDS, 'Budget__c'),
      headcountRange: getCustomFieldValue_(l.CUSTOMFIELDS, 'Guest_Headcount_Range__c'),
      description: l.LEAD_DESCRIPTION || '',
      otherInfo: getCustomFieldValue_(l.CUSTOMFIELDS, 'Other_Info__c'),
      insightlyUrl: 'https://crm.' + pod + '.insightly.com/details/Lead/' + l.LEAD_ID
    };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    debugLog_('ERROR looking up Insightly lead by id: ' + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ found: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles a "look up these phone numbers, all at once" request - used
 * for showing names next to conversations in a list. Fetches the
 * lead list only once (using the same cache as single lookups) no
 * matter how many phone numbers are asked for, rather than each
 * conversation triggering its own separate Insightly fetch.
 * Called from doGet(e) when e.parameter.action === 'lookupContactsBatch'.
 */
function handleLookupContactsBatch(phonesParam) {
  try {
    const phones = (phonesParam || '').split(',').map(p => p.trim()).filter(Boolean);
    const leads = fetchAllInsightlyLeads_(); // single fetch (or cache hit) for the whole batch

    const results = {};
    phones.forEach(phone => {
      const target = normalizePhoneDigits_(phone);
      const match = leads.find(l => l.phones.some(p => normalizePhoneDigits_(p) === target));
      results[phone] = match
        ? {
            found: true,
            name: [match.firstName, match.lastName].filter(Boolean).join(' ') || '(no name)',
            eventDate: formatInsightlyDate_(match.eventDate),
            venue: match.venue || '',
            cityState: match.cityState || ''
          }
        : { found: false };
    });

    return ContentService.createTextOutput(JSON.stringify({ success: true, results: results }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    debugLog_('ERROR batch looking up Insightly leads: ' + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Appends a band name to a Lead's "Bands__c" field, which Insightly
 * stores as a semicolon-separated list (e.g. "The Hook Club;The
 * Brooklyn Beats"). Skips if the band is already listed. Fetches the
 * full lead first and PUTs it back with just this field changed, to
 * avoid clobbering anything else on the record.
 */
function appendBandToInsightlyLead_(leadId, bandName) {
  const authHeader = getInsightlyAuthHeader_();
  const url = getInsightlyBaseUrl_() + '/Leads/' + leadId;

  const getResponse = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: authHeader },
    muteHttpExceptions: true
  });
  if (getResponse.getResponseCode() !== 200) {
    throw new Error('Could not fetch lead ' + leadId + ' (' + getResponse.getResponseCode() + '): ' + getResponse.getContentText());
  }

  const lead = JSON.parse(getResponse.getContentText());
  const customFields = lead.CUSTOMFIELDS || [];
  const bandsField = customFields.find(f => f.FIELD_NAME === 'Bands__c');

  const existingBands = bandsField && bandsField.FIELD_VALUE
    ? bandsField.FIELD_VALUE.split(';').map(b => b.trim()).filter(Boolean)
    : [];

  if (existingBands.indexOf(bandName) !== -1) {
    debugLog_('Band "' + bandName + '" already listed on lead ' + leadId + ' - skipping.');
    return; // already there, nothing to do
  }

  existingBands.push(bandName);
  const newValue = existingBands.join(';');

  if (bandsField) {
    bandsField.FIELD_VALUE = newValue;
  } else {
    customFields.push({ FIELD_NAME: 'Bands__c', FIELD_VALUE: newValue, CUSTOM_FIELD_ID: 'Bands__c' });
  }
  lead.CUSTOMFIELDS = customFields;

  const putResponse = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: authHeader },
    payload: JSON.stringify(lead),
    muteHttpExceptions: true
  });

  const code = putResponse.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Could not update lead ' + leadId + ' (' + code + '): ' + putResponse.getContentText());
  }

  debugLog_('Added "' + bandName + '" to Bands__c on lead ' + leadId + '.');
}

/**
 * Clears the "Band__c" field on a Lead - used when a referral is sent
 * out, since The Barnstorm is no longer the band handling that lead.
 * Fetches the full lead first and PUTs it back with just this field
 * changed, same pattern as appendBandToInsightlyLead_.
 */
function clearBandFieldOnInsightlyLead_(leadId) {
  const authHeader = getInsightlyAuthHeader_();
  const url = getInsightlyBaseUrl_() + '/Leads/' + leadId;

  const getResponse = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: authHeader },
    muteHttpExceptions: true
  });
  if (getResponse.getResponseCode() !== 200) {
    throw new Error('Could not fetch lead ' + leadId + ' (' + getResponse.getResponseCode() + '): ' + getResponse.getContentText());
  }

  const lead = JSON.parse(getResponse.getContentText());
  const customFields = lead.CUSTOMFIELDS || [];
  const bandField = customFields.find(f => f.FIELD_NAME === 'Band__c');

  if (!bandField || !bandField.FIELD_VALUE) {
    debugLog_('Band__c already empty on lead ' + leadId + ' - skipping.');
    return;
  }

  bandField.FIELD_VALUE = '';
  lead.CUSTOMFIELDS = customFields;

  const putResponse = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: authHeader },
    payload: JSON.stringify(lead),
    muteHttpExceptions: true
  });

  const code = putResponse.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Could not clear Band__c on lead ' + leadId + ' (' + code + '): ' + putResponse.getContentText());
  }

  debugLog_('Cleared Band__c on lead ' + leadId + '.');
}

/**
 * Removes a specific band from a Lead's "Bands__c" list (the same
 * semicolon-separated field appendBandToInsightlyLead_ adds to).
 * Used to remove "The Barnstorm" from that list once a referral is
 * sent out, since it's the entry Insightly's own automation seeds
 * onto every new lead by default.
 */
function removeBandFromInsightlyLead_(leadId, bandName) {
  const authHeader = getInsightlyAuthHeader_();
  const url = getInsightlyBaseUrl_() + '/Leads/' + leadId;

  const getResponse = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: authHeader },
    muteHttpExceptions: true
  });
  if (getResponse.getResponseCode() !== 200) {
    throw new Error('Could not fetch lead ' + leadId + ' (' + getResponse.getResponseCode() + '): ' + getResponse.getContentText());
  }

  const lead = JSON.parse(getResponse.getContentText());
  const customFields = lead.CUSTOMFIELDS || [];
  const bandsField = customFields.find(f => f.FIELD_NAME === 'Bands__c');

  if (!bandsField || !bandsField.FIELD_VALUE) {
    debugLog_('Bands__c already empty on lead ' + leadId + ' - nothing to remove.');
    return;
  }

  const existingBands = bandsField.FIELD_VALUE.split(';').map(b => b.trim()).filter(Boolean);
  const filteredBands = existingBands.filter(b => b !== bandName);

  if (filteredBands.length === existingBands.length) {
    debugLog_('"' + bandName + '" was not in Bands__c on lead ' + leadId + ' - nothing to remove.');
    return;
  }

  bandsField.FIELD_VALUE = filteredBands.join(';');
  lead.CUSTOMFIELDS = customFields;

  const putResponse = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: authHeader },
    payload: JSON.stringify(lead),
    muteHttpExceptions: true
  });

  const code = putResponse.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Could not update lead ' + leadId + ' (' + code + '): ' + putResponse.getContentText());
  }

  debugLog_('Removed "' + bandName + '" from Bands__c on lead ' + leadId + '.');
}
function testInsightlyLeadCount() {
  const baseUrl = getInsightlyBaseUrl_();
  const authHeader = getInsightlyAuthHeader_();
  const pageSize = 500;
  let total = 0, skip = 0, bridgetPos = -1;
  for (let page = 0; page < 40; page++) {
    const r = UrlFetchApp.fetch(baseUrl + '/Leads?top=' + pageSize + '&skip=' + skip, {
      method: 'get', headers: { Authorization: authHeader }, muteHttpExceptions: true
    });
    if (r.getResponseCode() !== 200) {
      debugLog_('Lead count test - HTTP ' + r.getResponseCode());
      return;
    }
    const batch = JSON.parse(r.getContentText());
    batch.forEach((l, i) => {
      if (bridgetPos === -1 && [l.MOBILE, l.PHONE].some(p => normalizePhoneDigits_(p) === '5167801744')) {
        bridgetPos = total + i + 1;
      }
    });
    total += batch.length;
    if (batch.length < pageSize) break;
    skip += pageSize;
  }
  debugLog_('Lead count test - Insightly returned ' + total + ' leads. Bridget is lead number ' + bridgetPos + ' (app only reads the first 2500).');
}