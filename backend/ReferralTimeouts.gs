/**
 * ====================================================================
 * REFERRAL TIMEOUTS - auto-record "No Thx" for any bandleader who
 * hasn't responded to a referral within 3 "active" hours.
 * ====================================================================
 * Active hours are 9:00am to 10:00pm New York time. The clock pauses
 * overnight (10:00pm to 9:00am) and resumes at 9:00am.
 *
 * Example: a referral sent at 10:00pm expires at 12:00pm the next day.
 *
 * SETUP: none. The existing time-driven trigger on
 * processReferralTimeouts (every 10 minutes) keeps working as is.
 * ====================================================================
 */

const REFERRAL_AUTODECLINE_AFTER_MINUTES = 180; // active minutes of waiting before auto-decline
const REFERRAL_ACTIVE_START_HOUR = 9;           // 9:00am
const REFERRAL_ACTIVE_END_HOUR = 22;            // 10:00pm
const REFERRAL_TIMEZONE = 'America/New_York';   // handles daylight saving automatically

/**
 * Returns true if the given moment falls inside active hours
 * (9:00am up to, but not including, 10:00pm New York time).
 */
function isReferralActiveTime_(date) {
  const hour = Number(Utilities.formatDate(date, REFERRAL_TIMEZONE, 'H'));
  return hour >= REFERRAL_ACTIVE_START_HOUR && hour < REFERRAL_ACTIVE_END_HOUR;
}

/**
 * Counts how many active minutes passed between start and end.
 * Stops early once it reaches stopAt, since we only need to know
 * whether the limit has been hit.
 */
function activeMinutesElapsed_(start, end, stopAt) {
  let count = 0;
  let t = start.getTime();
  const endMs = end.getTime();
  while (t < endMs && count < stopAt) {
    if (isReferralActiveTime_(new Date(t))) count++;
    t += 60000; // step forward one minute
  }
  return count;
}

/**
 * Time-driven entry point (runs every 10 minutes via the trigger).
 * Finds every still-pending referral recipient that has waited
 * 3 active hours and auto-records them as declined, the same way a
 * real "No Thx!" tap would.
 */
function processReferralTimeouts() {
  const now = new Date();
  let all;
  try {
    all = fetchAllReferralRecipients_();
  } catch (err) {
    debugLog_('ERROR fetching referral recipients for timeout check: ' + err.toString());
    return;
  }

  const overdue = all.filter(r => {
    if (r.status !== 'pending') return false;
    const sentAt = new Date(r.sentAt);
    if (isNaN(sentAt.getTime())) return false; // skip records with a bad date
    const waited = activeMinutesElapsed_(sentAt, now, REFERRAL_AUTODECLINE_AFTER_MINUTES);
    return waited >= REFERRAL_AUTODECLINE_AFTER_MINUTES;
  });

  overdue.forEach(r => {
    try {
      updateReferralRecipientInFirestore_(r.id, {
        status: 'declined',
        respondedAt: now.toISOString(),
        responseText: 'No response (auto-declined after 3 hours, overnight hours excluded)'
      });
      debugLog_('Referral auto-declined for ' + r.name + ' (' + r.phone + ') after 3 active hours with no response');
    } catch (err) {
      debugLog_('ERROR auto-declining referral for ' + r.phone + ': ' + err.toString());
    }
  });
}
