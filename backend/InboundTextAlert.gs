/**
 * INBOUND TEXT ALERT
 * Add as a NEW file in your Apps Script project (click "+" next to Files,
 * choose Script, name it "InboundTextAlert", paste this in).
 *
 * Code.gs calls notifyAdminOfInboundText_() every time Twilio delivers an
 * incoming text. It sends a push to the admin's saved device token using
 * sendPushToToken_ from PushNotifications.gs.
 *
 * No secret needed: the call happens inside the script, not from outside.
 * Names come from Insightly via findInsightlyLeadByPhone_ (in Insightly.gs).
 */

const ADMIN_PUSH_EMAIL = 'booking@thebarnstorm.com';

// WhatsApp replies that should NOT trigger an alert (quick replies from bandleaders).
// Matching ignores capitalization, punctuation and extra spaces,
// so "Yes Pls", "yes pls!" and "YES  PLS." are all treated the same.
const SILENT_WHATSAPP_REPLIES = ['yes pls', 'no thx'];

function isSilentWhatsAppReply_(twilioParams) {
  const isWhatsApp = (twilioParams.From || '').toString().indexOf('whatsapp:') === 0;
  if (!isWhatsApp) return false;
  const cleaned = (twilioParams.Body || '').toString()
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  return SILENT_WHATSAPP_REPLIES.indexOf(cleaned) !== -1;
}

function notifyAdminOfInboundText_(twilioParams) {
  if (isSilentWhatsAppReply_(twilioParams)) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PUSH_TOKENS_SHEET_NAME);
  if (!sheet) return;

  // Find the admin's token (column B = Email, column C = Token)
  const values = sheet.getDataRange().getValues();
  let token = null;
  for (let i = 1; i < values.length; i++) {
    if ((values[i][1] || '').toString().toLowerCase() === ADMIN_PUSH_EMAIL.toLowerCase()) {
      token = values[i][2];
      break;
    }
  }
  if (!token) return;

  // Who it's from: CRM lead name (Insightly), else WhatsApp profile name,
  // else the phone number. A CRM problem never blocks the alert.
  const from = (twilioParams.From || '').toString().replace('whatsapp:', '');
  let who = '';
  try {
    const lead = findInsightlyLeadByPhone_(from);
    if (lead) {
      who = [lead.firstName, lead.lastName].filter(Boolean).join(' ');
    }
  } catch (lookupErr) {
    debugLog_('Inbound text CRM lookup failed: ' + lookupErr);
  }
  if (!who) who = (twilioParams.ProfileName || from || 'Unknown number').toString();

  let body = (twilioParams.Body || '').toString().replace(/\s+/g, ' ').trim();
  if (body.length > 120) body = body.substring(0, 117) + '...';
  if (!body) body = '(no message text)';

  // phone + channel let the app open this exact conversation when tapped
  const channel = (twilioParams.From || '').toString().indexOf('whatsapp:') === 0 ? 'whatsapp' : 'sms';
  sendPushToToken_(token, 'New text from ' + who, body, { view: 'messages', phone: from, channel: channel });
}

/**
 * TEST: pick "testInboundTextAlert" in the function dropdown and click Run.
 * You should get a push on your phone within a few seconds.
 */
function testInboundTextAlert() {
  // Use a real lead's phone number here to see their name in the alert
  notifyAdminOfInboundText_({ From: '+15555550123', Body: 'Hi! Yes, we are still interested in booking.' });
}

/**
 * TEST: checks the silent-reply rule without sending anything.
 * Run it, then open View > Logs. Every line should say "true" for the
 * first four and "false" for the last three.
 */
function testSilentReplies() {
  const cases = [
    ['whatsapp:+15555550123', 'Yes Pls'],
    ['whatsapp:+15555550123', 'yes pls!'],
    ['whatsapp:+15555550123', 'NO THX'],
    ['whatsapp:+15555550123', ' no   thx. '],
    ['whatsapp:+15555550123', 'Yes please, send details'],
    ['whatsapp:+15555550123', 'Can we talk?'],
    ['+15555550123', 'Yes Pls']   // regular SMS: still alerts
  ];
  cases.forEach(c => Logger.log(c[1] + ' -> silent: ' + isSilentWhatsAppReply_({ From: c[0], Body: c[1] })));
}
