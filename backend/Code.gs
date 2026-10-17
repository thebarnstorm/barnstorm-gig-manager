// ─── Cache Settings ───────────────────────────────────────────
const CACHE_KEY_META = 'gigs_meta_v1';   // column/header structure
const CACHE_KEY_GIGS = 'gigs_data_v2';   // actual gig records
const GIGS_TTL  = 300;   // 5 minutes
const META_TTL  = 3600;  // 1 hour

function doGet(e) {
  // ── Admin-only actions need a valid login token (see Auth.gs) ──
  const ADMIN_GET_ACTIONS = ['lookupContact', 'lookupContactById', 'lookupContactsBatch', 'getRoster', 'getRosterGroups'];
  if (ADMIN_GET_ACTIONS.indexOf(e.parameter.action) !== -1) {
    const denied = requireAdmin_(e.parameter.idToken, e.parameter.action);
    if (denied) return denied;
  }

  // ── Handle Insightly contact lookups from the Messages tab ─────
  if (e.parameter.action === 'lookupContact') {
    return handleLookupContact(e.parameter.phone);
  }

  // ── Handle direct lead-by-ID lookups (used by the Referrals list) ──
  if (e.parameter.action === 'lookupContactById') {
    return handleLookupContactById(e.parameter.leadId);
  }

  // ── Handle batched Insightly contact lookups (for name lists) ──
  if (e.parameter.action === 'lookupContactsBatch') {
    return handleLookupContactsBatch(e.parameter.phones);
  }

  // ── Handle bandleader roster requests from the Messages tab ────
  if (e.parameter.action === 'getRoster') {
    return handleGetRoster();
  }

  // ── Handle roster group list requests from the Messages tab ────
  if (e.parameter.action === 'getRosterGroups') {
    return handleGetRosterGroups();
  }

  const bypassCache = e.parameter.bypass === 'true';
  const cache = CacheService.getScriptCache();

  // ── Try gig cache first ──────────────────────────────────────
  if (!bypassCache) {
    const cached = cache.get(CACHE_KEY_GIGS);
    if (cached) {
      Logger.log('Returning cached gig data');
      return ContentService.createTextOutput(cached)
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  Logger.log('Cache miss - fetching fresh data');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Barnstorm BI');

  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({
      error: 'Sheet named "Barnstorm BI" not found. Available sheets: ' +
        ss.getSheets().map(s => s.getName()).join(', ')
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: 'Not enough data (need at least 2 rows)' })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  // ── Try meta cache (headers/column structure) ────────────────
  let meta = null;
  if (!bypassCache) {
    const cachedMeta = cache.get(CACHE_KEY_META);
    if (cachedMeta) {
      meta = JSON.parse(cachedMeta);
      Logger.log('Using cached column metadata');
    }
  }

  // ── Build meta if not cached ─────────────────────────────────
  if (!meta) {
    Logger.log('Building column metadata');

    const row1Lower = data[0].map(h => h.toString().toLowerCase().trim());
    const row2Lower = data[1].map(h => h.toString().toLowerCase().trim());
    const headerIndicators = ['date', 'who', 'venue', 'event', 'location'];
    const row1HasHeaders = headerIndicators.some(i => row1Lower.includes(i));
    const row2HasHeaders = headerIndicators.some(i => row2Lower.includes(i));

    const headerRowIndex = (row2HasHeaders && !row1HasHeaders) ? 1 : 0;
    const dataStartIndex = headerRowIndex + 1;
    const headers = data[headerRowIndex].map(h => h.toString().toLowerCase().trim());

    function findColumn(terms) {
      for (let t of terms) {
        const idx = headers.indexOf(t.toLowerCase());
        if (idx !== -1) return idx;
      }
      return -1;
    }

    const musicianHeaders = ['gtr / vox', 'vox / emcee', 'drums', 'bass', 'keys', 'power vox', 'saxophone', 'trumpet', 'trombone', 'violin'];
    const productionHeaders = ['a1', 'a2', 'l1', 'l2', 'photog'];

    const musicianColIndices = [], musicianColHeaders = [];
    const productionColIndices = [], productionColHeaders = [];
    for (let i = 0; i < headers.length; i++) {
      if (musicianHeaders.includes(headers[i])) { musicianColIndices.push(i); musicianColHeaders.push(headers[i]); }
      if (productionHeaders.includes(headers[i])) { productionColIndices.push(i); productionColHeaders.push(headers[i]); }
    }

    meta = {
      headerRowIndex,
      dataStartIndex,
      headers,
      dateColIndex:      findColumn(['date', 'event date', 'gig date']),
      eventColIndex:     findColumn(['who', 'event', 'event name', 'name', 'client']),
      venueColIndex:     findColumn(['venue', 'location name']),
      typeColIndex:      findColumn(['type', 'gig type', 'event type']),
      addressColIndex:   findColumn(['location', 'address', 'venue address']),
      showInfoColIndex:  headers.indexOf('show info'),
      ceremonyColIndex:  findColumn(['ceremony?', 'ceremony']),
      cocktailColIndex:  findColumn(['cocktail hour?', 'cocktail hour']),
      afterPartyColIndex: findColumn(['after party dj', 'after party']),
      photoColIndex:     headers.indexOf('photo'),
      videoColIndex:     headers.indexOf('video'),
      audioColIndex:     headers.indexOf('audio'),
      photographerColIndex: headers.indexOf('photographer'),
      videoEditsColIndex: headers.indexOf('video edits'),
      vendorsColIndex:   headers.indexOf('vendors'),
      blogColIndex:      headers.indexOf('blog entry'),
      musicianColIndices,
      musicianColHeaders,
      productionColIndices,
      productionColHeaders
    };

    // Cache meta for 1 hour
    try {
      cache.put(CACHE_KEY_META, JSON.stringify(meta), META_TTL);
      Logger.log('Column metadata cached for 1 hour');
    } catch (err) {
      Logger.log('Meta cache error: ' + err);
    }
  }

  // ── Destructure meta ─────────────────────────────────────────
  const {
    dataStartIndex, headers,
    dateColIndex, eventColIndex, venueColIndex, typeColIndex, addressColIndex,
    showInfoColIndex, ceremonyColIndex, cocktailColIndex, afterPartyColIndex,
    photoColIndex, videoColIndex, audioColIndex, photographerColIndex,
    videoEditsColIndex, vendorsColIndex, blogColIndex,
    musicianColIndices, musicianColHeaders,
    productionColIndices, productionColHeaders
  } = meta;

  // ── Pre-fetch rich text for media/URL columns ─────────────────
  const numRows = data.length;
  const mediaColumns = [photoColIndex, videoColIndex, audioColIndex, videoEditsColIndex, blogColIndex, showInfoColIndex].filter(i => i !== -1);
  const richTextCache = {};

  if (mediaColumns.length > 0 && numRows > dataStartIndex) {
    try {
      for (const colIndex of mediaColumns) {
        const startRow = dataStartIndex + 1;
        const numDataRows = numRows - dataStartIndex;
        richTextCache[colIndex] = sheet.getRange(startRow, colIndex + 1, numDataRows, 1).getRichTextValues();
      }
    } catch (err) {
      Logger.log('Rich text fetch error: ' + err);
    }
  }

  function extractUrlFromCache(rowIndex, colIndex) {
    if (colIndex === -1) return null;
    try {
      const cacheIndex = rowIndex - dataStartIndex;
      if (richTextCache[colIndex] && richTextCache[colIndex][cacheIndex]) {
        const richText = richTextCache[colIndex][cacheIndex][0];
        if (richText) {
          const runs = richText.getRuns();
          if (runs && runs.length > 0) {
            const url = runs[0].getLinkUrl();
            if (url) return url;
          }
        }
      }
      const cellValue = data[rowIndex][colIndex];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith('http')) return cellValue;
    } catch (err) {
      Logger.log('URL extract error: ' + err);
    }
    return null;
  }

  // ── Flags a cell as explicitly marked "N/A" (as opposed to just empty) ──
  function isNAValue(value) {
    return value !== null && value !== undefined && value.toString().trim().toLowerCase() === 'n/a';
  }

  // ── Build gigs array ─────────────────────────────────────────
  const gigs = [];

  for (let i = dataStartIndex; i < data.length; i++) {
    const row = data[i];

    const date      = dateColIndex !== -1 ? row[dateColIndex] : null;
    const eventName = eventColIndex !== -1 ? row[eventColIndex] : '';
    const venue     = venueColIndex !== -1 ? row[venueColIndex] : '';
    const type      = typeColIndex !== -1 ? row[typeColIndex] : '';
    const address   = addressColIndex !== -1 ? row[addressColIndex] : '';
    const ceremony  = ceremonyColIndex !== -1 ? row[ceremonyColIndex] : '';
    const cocktailHour  = cocktailColIndex !== -1 ? row[cocktailColIndex] : '';
    const afterPartyDJ  = afterPartyColIndex !== -1 ? row[afterPartyColIndex] : '';
    const photographer  = photographerColIndex !== -1 ? row[photographerColIndex] : '';
    const vendors   = vendorsColIndex !== -1 ? row[vendorsColIndex] : '';

    const hasMusicians = musicianColIndices.some(idx => row[idx] && row[idx].toString().trim());
    if (!eventName && !venue && !hasMusicians) continue;

    let formattedDate = '';
    if (date) {
      try {
        const d = new Date(date);
        if (!isNaN(d.getTime())) formattedDate = d.toISOString().split('T')[0];
      } catch (err) {}
    }
    if (!formattedDate) continue;

    const musicians = musicianColIndices.map((colIdx, idx) => {
      const value = row[colIdx] ? row[colIdx].toString().trim() : '';
      return { role: musicianColHeaders[idx], name: (value.toUpperCase() === 'N/A' || value.toUpperCase() === 'NEED') ? '' : value, isNA: value.toUpperCase() === 'N/A', isNeed: value.toUpperCase() === 'NEED', originalName: value };
    });

    const production = productionColIndices.map((colIdx, idx) => {
      const value = row[colIdx] ? row[colIdx].toString().trim() : '';
      return { role: productionColHeaders[idx], name: (value.toUpperCase() === 'N/A' || value.toUpperCase() === 'NEED') ? '' : value, isNA: value.toUpperCase() === 'N/A', isNeed: value.toUpperCase() === 'NEED', originalName: value };
    });

    gigs.push({
      id: `gig-${i}`,
      date: formattedDate,
      eventName: eventName.toString().trim(),
      venue: venue.toString().trim(),
      type: type.toString().trim(),
      address: address.toString().trim(),
      showInfoUrl: extractUrlFromCache(i, showInfoColIndex),
      musicians,
      production,
      services: {
        ceremony: ceremony.toString().trim(),
        cocktailHour: cocktailHour.toString().trim(),
        afterPartyDJ: afterPartyDJ.toString().trim()
      },
      media: {
        photo: extractUrlFromCache(i, photoColIndex) || '',
        photoNA: isNAValue(row[photoColIndex]),
        video: extractUrlFromCache(i, videoColIndex) || '',
        videoNA: isNAValue(row[videoColIndex]),
        audio: extractUrlFromCache(i, audioColIndex) || '',
        audioNA: isNAValue(row[audioColIndex]),
        photographer: photographer.toString().trim(),
        videoEdits: extractUrlFromCache(i, videoEditsColIndex) || '',
        videoEditsNA: isNAValue(row[videoEditsColIndex]),
        vendors: vendors.toString().trim(),
        blog: extractUrlFromCache(i, blogColIndex) || '',
        blogNA: isNAValue(row[blogColIndex])
      }
    });
  }

  // ── Cache and return ─────────────────────────────────────────
  const jsonResult = JSON.stringify({
    gigs,
    cachedAt: new Date().toISOString()   // timestamp for "last synced" display later
  });

  try {
    cache.put(CACHE_KEY_GIGS, jsonResult, GIGS_TTL);
    Logger.log('Gig data cached for 5 minutes');
  } catch (err) {
    Logger.log('Gig cache error: ' + err);
  }

  return ContentService.createTextOutput(jsonResult)
    .setMimeType(ContentService.MimeType.JSON);
}


// ─── POST: Twilio messages + N/A functionality + Push Token saving ──────────────
function doPost(e) {
  try {
    Logger.log('doPost called');
    debugLog_('doPost fired. parameter=' + JSON.stringify(e.parameter) + ' | postData.type=' + (e.postData ? e.postData.type : 'none') + ' | contents=' + (e.postData ? String(e.postData.contents).replace(/"idToken":"[^"]*"/g, '"idToken":"[hidden]"') : 'none'));

    // ── NEW: handle incoming Twilio SMS/WhatsApp webhooks ──────
    // Twilio sends form fields (not JSON), so this must be checked
    // BEFORE the JSON.parse line below, or JSON.parse will fail on it.
    if (e.parameter && e.parameter.From && e.parameter.Body !== undefined) {
      const twilioResponse = handleIncomingTwilioMessage(e.parameter);
      // Send the admin a push alert. Wrapped so a push problem can never
      // stop the text from being saved or break Twilio's response.
      try {
        notifyAdminOfInboundText_(e.parameter);
      } catch (pushErr) {
        debugLog_('Inbound text push failed: ' + pushErr);
      }
      return twilioResponse;
    }

    const params = JSON.parse(e.postData.contents);

    // ── Admin-only actions need a valid login token (see Auth.gs) ──
    const ADMIN_POST_ACTIONS = ['sendMessage', 'addLead', 'markMessagesRead', 'createReferral', 'deleteReferral', 'archiveReferral', 'restoreReferral'];
    if (ADMIN_POST_ACTIONS.indexOf(params.action) !== -1) {
      const denied = requireAdmin_(params.idToken, params.action);
      if (denied) return denied;
    }

    // ── Handle push notification token saves ──────────────
    if (params.action === 'savePushToken') {
      return handleSavePushToken(params);
    }

    // ── Handle sending a reply from the Messages tab ───────
    if (params.action === 'sendMessage') {
      return handleSendMessage(params);
    }

    // ── Handle adding a lead (no text sent) from the Messages tab ───────
    if (params.action === 'addLead') {
      return handleAddLead(params);
    }

    // ── Handle marking messages as read ───────
    if (params.action === 'markMessagesRead') {
      return handleMarkMessagesRead(params);
    }

    // ── Handle broadcasting a referral to the bandleader roster ────
    if (params.action === 'createReferral') {
      return handleCreateReferral(params);
    }

    // ── Handle clearing a referral ────
    if (params.action === 'deleteReferral') {
      return handleDeleteReferral(params);
    }

    // ── Handle archiving / restoring a referral ────
    if (params.action === 'archiveReferral') {
      return handleSetReferralArchived(params, true);
    }
    if (params.action === 'restoreReferral') {
      return handleSetReferralArchived(params, false);
    }

    const gigId = params.gigId;
    const memberName = params.memberName;
    const type = params.type;

    // ── Only the member being marked N/A (or the booking account) may send this (see Auth.gs) ──
    const naDenied = checkNaSender_(params);
    if (naDenied) return naDenied;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Barnstorm BI');

    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Barnstorm BI sheet not found' })).setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getDataRange().getValues();
    const row1Lower = data[0].map(h => h.toString().toLowerCase().trim());
    const row2Lower = data[1].map(h => h.toString().toLowerCase().trim());
    const headerIndicators = ['date', 'who', 'venue', 'event', 'location'];
    const row2HasHeaders = headerIndicators.some(i => row2Lower.includes(i));
    const row1HasHeaders = headerIndicators.some(i => row1Lower.includes(i));
    const headerRowIndex = (row2HasHeaders && !row1HasHeaders) ? 1 : 0;
    const dataStartIndex = headerRowIndex + 1;
    const headers = data[headerRowIndex].map(h => h.toString().toLowerCase().trim());

    const dateColIndex = headers.findIndex(h => h.includes('date'));
    let rowIndex = -1;
    for (let i = dataStartIndex; i < data.length; i++) {
      if (`gig-${i}` === gigId) { rowIndex = i; break; }
    }

    if (rowIndex === -1) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Gig not found' })).setMimeType(ContentService.MimeType.JSON);
    }

    const musicianHeaders = ['gtr / vox', 'vox / emcee', 'drums', 'bass', 'keys', 'power vox', 'saxophone', 'trumpet', 'trombone', 'violin'];
    const productionHeaders = ['a1', 'a2', 'l1', 'l2', 'photog'];
    const columnsToSearch = type === 'musician' ? musicianHeaders : productionHeaders;

    let foundColIndex = -1;
    for (let j = 0; j < headers.length; j++) {
      if (columnsToSearch.includes(headers[j])) {
        const cellValue = data[rowIndex][j];
        if (cellValue && cellValue.toString().trim().toLowerCase() === memberName.toLowerCase()) {
          foundColIndex = j; break;
        }
      }
    }

    if (foundColIndex === -1) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Member not found in lineup' })).setMimeType(ContentService.MimeType.JSON);
    }

    sheet.getRange(rowIndex + 1, foundColIndex + 1).setValue('NEED');

    // Also invalidate the gig cache so next load reflects the change
    try {
      CacheService.getScriptCache().remove(CACHE_KEY_GIGS);
      Logger.log('Gig cache invalidated after N/A update');
    } catch (err) {
      Logger.log('Cache invalidation error: ' + err);
    }

    try {
      const gigDate = data[rowIndex][dateColIndex];
      const eventColIndex = headers.findIndex(h => h.includes('who') || h.includes('event') || h.includes('name'));
      const eventName = eventColIndex !== -1 ? data[rowIndex][eventColIndex] : 'Unknown Event';
      MailApp.sendEmail({
        to: 'booking@thebarnstorm.com',
        subject: `N/A Request: ${memberName} - ${eventName}`,
        body: `${memberName} has marked themselves as N/A for:\n\nEvent: ${eventName}\nDate: ${gigDate}\nRole: ${headers[foundColIndex]}\n\nPlease update the lineup accordingly.`
      });
    } catch (emailError) {
      Logger.log('Email error: ' + emailError);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('ERROR in doPost:', error);
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}