// ─── Slack DM Send — QuestFlow ────────────────────────────────────────────────
// Requires ScriptProperties key: SLACK_BOT_TOKEN
// Required Slack bot scopes: chat:write, users:read, users:read.email, im:write
// ──────────────────────────────────────────────────────────────────────────────

var SLACK_TOKEN_KEY = 'SLACK_BOT_TOKEN';

/**
 * Main entry point called from the frontend via google.script.run.
 * payload: { personName: string, personEmail: string, message: string }
 * Returns: { ok: true, name: string } | throws Error
 */
function sendSlackMessageToPerson(payload) {
  if (!payload || !payload.message || !String(payload.message).trim()) {
    throw new Error('Please enter a message before sending.');
  }
  if (!payload.personEmail || !String(payload.personEmail).trim()) {
    throw new Error('No recipient email provided.');
  }

  var token = PropertiesService.getScriptProperties().getProperty(SLACK_TOKEN_KEY);
  if (!token || !token.trim()) {
    throw new Error('Slack is not connected. Add SLACK_BOT_TOKEN in Apps Script → Project Settings → Script Properties.');
  }
  token = token.trim();

  var slackUserId = lookupSlackUserByEmail(payload.personEmail.trim(), token);
  var channelId   = openSlackDm(slackUserId, token);
  postSlackMessage(channelId, String(payload.message).trim(), token);

  return { ok: true, name: payload.personName || payload.personEmail };
}

/**
 * Looks up a Slack user ID by email address.
 * Returns the Slack user ID string, or throws.
 */
function lookupSlackUserByEmail(email, token) {
  var res = UrlFetchApp.fetch(
    'https://slack.com/api/users.lookupByEmail?email=' + encodeURIComponent(email),
    {
      method: 'get',
      muteHttpExceptions: true,
      headers: { 'Authorization': 'Bearer ' + token },
    }
  );

  var data = JSON.parse(res.getContentText() || '{}');

  if (!data.ok) {
    if (data.error === 'users_not_found') {
      throw new Error('Could not find Slack user for this email: ' + email);
    }
    throw new Error('Slack lookup failed: ' + (data.error || 'unknown error'));
  }

  return data.user.id;
}

/**
 * Opens (or retrieves) a DM channel with a Slack user.
 * Returns the DM channel ID string, or throws.
 */
function openSlackDm(slackUserId, token) {
  var res = UrlFetchApp.fetch('https://slack.com/api/conversations.open', {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json; charset=utf-8',
    },
    payload: JSON.stringify({ users: slackUserId }),
  });

  var data = JSON.parse(res.getContentText() || '{}');

  if (!data.ok) {
    throw new Error('Could not open Slack DM: ' + (data.error || 'unknown error'));
  }

  return data.channel.id;
}

/**
 * Posts a message to a Slack channel/DM.
 * Throws if the send fails.
 */
function postSlackMessage(channelId, message, token) {
  var res = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json; charset=utf-8',
    },
    payload: JSON.stringify({ channel: channelId, text: message }),
  });

  var data = JSON.parse(res.getContentText() || '{}');

  if (!data.ok) {
    throw new Error('Slack message send failed: ' + (data.error || 'unknown error'));
  }
}

// ─── Test function ─────────────────────────────────────────────────────────────
// Edit RECIPIENT_EMAIL and MESSAGE, then run from Apps Script editor to test.
function testSendSlackMessageToPerson() {
  var RECIPIENT_EMAIL = 'teammate@yourcompany.com'; // ← change this
  var MESSAGE = 'Hi! This is a test check-in from QuestFlow.';

  try {
    var result = sendSlackMessageToPerson({
      personName:  'Test User',
      personEmail: RECIPIENT_EMAIL,
      message:     MESSAGE,
    });
    Logger.log('✓ Success: ' + JSON.stringify(result));
  } catch (e) {
    Logger.log('✗ Error: ' + e.message);
  }
}
