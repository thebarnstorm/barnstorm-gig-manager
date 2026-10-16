// ============================================================
// Reminders.gs - day-before gig reminder push notifications
// ============================================================
// Every morning at 9 AM Eastern, anyone listed on tomorrow's gig gets a
// push notification. Uses the same push setup as the rest of the app
// (PushNotifications.gs) and the "Push Tokens" tab in the Master Doc.
//
// People are matched by NAME: the name in the gig's musician/production
// cell must match the name saved on the Push Tokens tab (same rule the
// app uses for "My Gigs"). Blank, NEED and N/A cells are skipped.

const REMINDER_TIMEZONE = 'America/New_York';
const REMINDER_HOUR = 9;
const REMINDER_LAST_SENT_KEY = 'reminder_last_sent_for_date';
const REMINDER_ADMIN_EMAIL = 'booking@thebarnstorm.com';

/** RUN ONCE to set up the daily 9 AM timer. Safe to run again. */
function installDayBeforeReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDayBeforeReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDayBeforeReminders')
    .timeBased()
    .everyDays(1)
    .atHour(REMINDER_HOUR)
    .nearMinute(0)
    .inTimezone(REMINDER_TIMEZONE)
    .create();
  debugLog_('Day-before reminder timer installed (daily around ' + REMINDER_HOUR + ':00 ' + REMINDER_TIMEZONE + ').');
}

/** Runs on the daily timer. Sends reminders for tomorrow's gigs. */
function sendDayBeforeReminders() {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const tomorrow = Utilities.formatDate(new Date(Date.now() + 24 * 60 * 60 * 1000), tz, 'yyyy-MM-dd');

  // Never send twice for the same date, even if the timer fires twice
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(REMINDER_LAST_SENT_KEY) === tomorrow) {
    debugLog_('Reminders: already sent for ' + tomorrow + ', skipping.');
    return;
  }
  if (runReminders_(tomorrow, 'send')) {
    props.setProperty(REMINDER_LAST_SENT_KEY, tomorrow);
  }
}

/** Dry run for tomorrow: only logs who WOULD be notified. Sends nothing. */
function previewDayBeforeReminders() {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const tomorrow = Utilities.formatDate(new Date(Date.now() + 24 * 60 * 60 * 1000), tz, 'yyyy-MM-dd');
  runReminders_(tomorrow, 'preview');
}

/** Dry run for the NEXT upcoming gig (any date). Sends nothing. */
function previewNextGigReminder() {
  const next = findNextGigDate_();
  if (!next) { debugLog_('Reminders: no upcoming gig found.'); return; }
  runReminders_(next, 'preview');
}

/**
 * Real test with no risk to the band: builds the reminder for the NEXT
 * upcoming gig but sends it ONLY to booking@thebarnstorm.com's device.
 */
function testSendNextGigReminderToMe() {
  const next = findNextGigDate_();
  if (!next) { debugLog_('Reminders: no upcoming gig found.'); return; }
  runReminders_(next, 'test');
}

// ── internals ───────────────────────────────────────────────

function findNextGigDate_() {
  const gigs = readGigRowsForReminders_();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const dates = gigs.map(function (g) { return g.date; }).filter(function (d) { return d >= today; }).sort();
  return dates.length ? dates[0] : '';
}

/** Reads the Barnstorm BI tab into a simple list of gigs with their people. */
function readGigRowsForReminders_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Barnstorm BI');
  if (!sheet) throw new Error('Sheet "Barnstorm BI" not found.');
  const tz = ss.getSpreadsheetTimeZone();
  const data = sheet.getDataRange().getValues();
  if (data.length < 3) return [];

  // Same header-row rule as the app: headers are in row 1 or row 2
  const lower = function (row) { return row.map(function (h) { return h.toString().toLowerCase().trim(); }); };
  const indicators = ['date', 'who', 'venue', 'event', 'location'];
  const r1 = lower(data[0]), r2 = lower(data[1]);
  const r1Has = indicators.some(function (i) { return r1.indexOf(i) !== -1; });
  const r2Has = indicators.some(function (i) { return r2.indexOf(i) !== -1; });
  const headerRow = (r2Has && !r1Has) ? 1 : 0;
  const headers = lower(data[headerRow]);

  const roleHeaders = ['gtr / vox', 'vox / emcee', 'drums', 'bass', 'keys', 'power vox', 'saxophone', 'trumpet', 'trombone', 'violin', 'a1', 'a2', 'l1', 'l2', 'photog'];
  const roleCols = [];
  headers.forEach(function (h, i) { if (roleHeaders.indexOf(h) !== -1) roleCols.push({ index: i, role: h }); });
  const dateCol = headers.indexOf('date');
  const whoCol = headers.indexOf('who');
  const venueCol = headers.indexOf('venue');
  if (dateCol === -1) throw new Error('Could not find the Date column.');

  const gigs = [];
  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    const rawDate = row[dateCol];
    if (!rawDate) continue;
    const d = new Date(rawDate);
    if (isNaN(d.getTime())) continue;
    const who = whoCol !== -1 ? row[whoCol].toString().trim() : '';
    const venue = venueCol !== -1 ? row[venueCol].toString().trim() : '';
    if (!who && !venue) continue;

    const people = [];
    roleCols.forEach(function (rc) {
      const v = row[rc.index] ? row[rc.index].toString().trim() : '';
      const up = v.toUpperCase();
      if (!v || up === 'NEED' || up === 'N/A') return;
      people.push({ name: v, role: rc.role.toUpperCase() });
    });
    gigs.push({ date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'), who: who, venue: venue, people: people });
  }
  return gigs;
}

/** Name (lowercase) -> push token, from the Push Tokens tab. */
function readPushTokensByName_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PUSH_TOKENS_SHEET_NAME);
  const out = { byName: {}, byEmail: {} };
  if (!sheet) return out;
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const name = (values[i][0] || '').toString().toLowerCase().trim();
    const email = (values[i][1] || '').toString().toLowerCase().trim();
    const token = (values[i][2] || '').toString().trim();
    if (!token) continue;
    if (name) out.byName[name] = token;
    if (email) out.byEmail[email] = token;
  }
  return out;
}

function runReminders_(dateStr, mode) {
  const label = mode === 'preview' ? 'PREVIEW' : (mode === 'test' ? 'TEST' : 'SEND');
  try {
    const gigs = readGigRowsForReminders_().filter(function (g) { return g.date === dateStr; });
    if (!gigs.length) { debugLog_('Reminders (' + label + '): no gigs on ' + dateStr + '.'); return true; }

    const tokens = readPushTokensByName_();
    let sent = 0, noToken = 0;

    gigs.forEach(function (g) {
      const title = (mode === 'test' ? '[TEST] ' : '') + 'Gig tomorrow: ' + (g.who || 'The Barnstorm');
      g.people.forEach(function (p) {
        const body = (g.venue ? g.venue + ' - ' : '') + 'you are on ' + p.role;
        const token = mode === 'test' ? tokens.byEmail[REMINDER_ADMIN_EMAIL] : tokens.byName[p.name.toLowerCase().trim()];

        if (!token) {
          if (mode !== 'test') noToken++;
          if (mode === 'preview') debugLog_('Reminders (PREVIEW): ' + p.name + ' (' + p.role + ') has no push token yet - ' + (g.who || g.venue));
          return;
        }
        if (mode === 'preview') {
          debugLog_('Reminders (PREVIEW): would notify ' + p.name + ' (' + p.role + ') - ' + title + ' | ' + body);
          sent++;
          return;
        }
        if (mode === 'test' && sent > 0) return; // one test push is enough

        try {
          const resp = sendPushToToken_(token, title, body, { gigDate: dateStr });
          if (resp.indexOf('"error"') !== -1) {
            debugLog_('Reminders: push to ' + p.name + ' failed: ' + resp);
          } else {
            sent++;
          }
        } catch (err) {
          debugLog_('Reminders: push to ' + p.name + ' error: ' + err);
        }
      });
    });

    debugLog_('Reminders (' + label + ') for ' + dateStr + ': ' + gigs.length + ' gig(s), ' + sent + ' notification(s) ' +
      (mode === 'preview' ? 'would be sent' : 'sent') + ', ' + noToken + ' person(s) without a push token.');
    return true;
  } catch (err) {
    debugLog_('ERROR in reminders: ' + err.toString());
    return false;
  }
}
