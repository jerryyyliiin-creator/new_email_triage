// ─── DocuSign Integration — QuestFlow ────────────────────────────────────────
// Supports two auth modes:
//   'apikey' — paste an access token from DocuSign developer portal
//   'jwt'    — enterprise JWT grant using RSA private key (no browser redirect)
//
// Script Properties used:
//   DS_AUTH_MODE        'apikey' | 'jwt'
//   DS_ACCOUNT_ID       Your DocuSign account ID (GUID)
//   DS_BASE_URL         e.g. https://na4.docusign.net/restapi
//   DS_ACCESS_TOKEN     (apikey mode) personal access token
//   DS_INTEGRATION_KEY  (jwt mode) OAuth integration key
//   DS_USER_ID          (jwt mode) DocuSign user ID (GUID) to impersonate
//   DS_PRIVATE_KEY      (jwt mode) RSA private key PEM string
// ─────────────────────────────────────────────────────────────────────────────

// ── Settings ─────────────────────────────────────────────────────────────────

function getDocuSignSettings() {
  var props = PropertiesService.getScriptProperties();
  return {
    authMode:       props.getProperty('DS_AUTH_MODE')       || 'apikey',
    accountId:      props.getProperty('DS_ACCOUNT_ID')      || '',
    baseUrl:        props.getProperty('DS_BASE_URL')        || 'https://demo.docusign.net/restapi',
    accessToken:    props.getProperty('DS_ACCESS_TOKEN')    || '',
    integrationKey: props.getProperty('DS_INTEGRATION_KEY') || '',
    userId:         props.getProperty('DS_USER_ID')         || '',
    hasPrivateKey:  !!(props.getProperty('DS_PRIVATE_KEY')),
  };
}

function saveDocuSignSettings(settings) {
  var props = PropertiesService.getScriptProperties();
  if (settings.authMode       !== undefined) props.setProperty('DS_AUTH_MODE',       settings.authMode);
  if (settings.accountId      !== undefined) props.setProperty('DS_ACCOUNT_ID',      settings.accountId);
  if (settings.baseUrl        !== undefined) props.setProperty('DS_BASE_URL',        settings.baseUrl || 'https://na4.docusign.net/restapi');
  if (settings.accessToken    !== undefined) props.setProperty('DS_ACCESS_TOKEN',    settings.accessToken);
  if (settings.integrationKey !== undefined) props.setProperty('DS_INTEGRATION_KEY', settings.integrationKey);
  if (settings.userId         !== undefined) props.setProperty('DS_USER_ID',         settings.userId);
  if (settings.privateKey     !== undefined && settings.privateKey.trim()) {
    props.setProperty('DS_PRIVATE_KEY', settings.privateKey.trim());
  }
  return { ok: true };
}

// ── Auth ──────────────────────────────────────────────────────────────────────

function getDocuSignToken_() {
  var s = getDocuSignSettings();
  if (s.authMode === 'jwt') return getJwtAccessToken_(s);
  if (!s.accessToken) throw new Error('DocuSign access token not configured. Add it in Settings → DocuSign.');
  return s.accessToken;
}

// ── PKCS#1 → PKCS#8 conversion ───────────────────────────────────────────────
// Apps Script's computeRsaSha256Signature requires PKCS#8 (BEGIN PRIVATE KEY).
// DocuSign generates PKCS#1 (BEGIN RSA PRIVATE KEY). This wraps it correctly.
function derEncode_(tag, bytes) {
  var len = bytes.length;
  var lenBytes = len < 128 ? [len] : len < 256 ? [0x81, len] : [0x82, (len >> 8) & 0xff, len & 0xff];
  return [tag].concat(lenBytes).concat(bytes);
}

function pkcs1ToPkcs8Pem_(rawInput) {
  // Strip headers, ALL whitespace, any non-base64 chars — accept any paste format
  var b64 = String(rawInput)
    .replace(/-----[^-]+-----/g, '')
    .replace(/[^A-Za-z0-9+/=]/g, '');
  if (!b64) throw new Error('RSA private key is empty after cleaning.');
  // Pad to multiple of 4
  while (b64.length % 4 !== 0) b64 += '=';
  var der = Utilities.base64Decode(b64, Utilities.Charset.UTF_8);
  var arr = [];
  for (var i = 0; i < der.length; i++) arr.push(der[i] & 0xff);
  var algSeq = [0x30,0x0d,0x06,0x09,0x2a,0x86,0x48,0x86,0xf7,0x0d,0x01,0x01,0x01,0x05,0x00];
  var octet  = derEncode_(0x04, arr);
  var inner  = [0x02,0x01,0x00].concat(algSeq).concat(octet);
  var pkcs8  = derEncode_(0x30, inner);
  var encoded = Utilities.base64Encode(pkcs8).match(/.{1,64}/g).join('\n');
  return '-----BEGIN PRIVATE KEY-----\n' + encoded + '\n-----END PRIVATE KEY-----';
}

// Returns the OAuth base host — sandbox when DS_BASE_URL contains "demo", production otherwise.
// Sandbox: account-d.docusign.com  |  Production: account.docusign.com
function dsAuthHost_(s) {
  var base = (s && s.baseUrl) || '';
  return base.indexOf('demo') !== -1 ? 'account-d.docusign.com' : 'account.docusign.com';
}

function getJwtAccessToken_(s) {
  if (!s.integrationKey) throw new Error('DocuSign Integration Key not configured.');
  if (!s.userId)         throw new Error('DocuSign User ID not configured.');
  var props = PropertiesService.getScriptProperties();
  var privateKey = props.getProperty('DS_PRIVATE_KEY');
  if (!privateKey) throw new Error('DocuSign RSA private key not configured.');
  // Convert to PKCS#8 (Apps Script requires it). pkcs1ToPkcs8Pem_ strips all
  // whitespace/headers so it works regardless of paste format.
  if (privateKey.indexOf('BEGIN PRIVATE KEY') !== -1 && privateKey.indexOf('BEGIN RSA') === -1) {
    // Already PKCS#8 — just normalise whitespace in the base64
    var b64only = privateKey.replace(/-----[^-]+-----/g,'').replace(/[^A-Za-z0-9+/=]/g,'');
    while (b64only.length % 4 !== 0) b64only += '=';
    var lines8 = b64only.match(/.{1,64}/g).join('\n');
    privateKey = '-----BEGIN PRIVATE KEY-----\n' + lines8 + '\n-----END PRIVATE KEY-----';
  } else {
    // PKCS#1 or raw base64 — convert to PKCS#8
    privateKey = pkcs1ToPkcs8Pem_(privateKey);
  }

  // Check cached token (valid for 55 min, DS tokens last 60 min)
  var cached    = props.getProperty('DS_JWT_TOKEN');
  var cachedExp = parseInt(props.getProperty('DS_JWT_TOKEN_EXP') || '0', 10);
  if (cached && cachedExp > Date.now() + 60000) return cached;

  // Build JWT
  var now = Math.floor(Date.now() / 1000);
  var header  = Utilities.base64EncodeWebSafe(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).replace(/=+$/, '');
  var payload = Utilities.base64EncodeWebSafe(JSON.stringify({
    iss: s.integrationKey,
    sub: s.userId,
    aud: dsAuthHost_(s),
    iat: now,
    exp: now + 3600,
    scope: 'signature impersonation',
  })).replace(/=+$/, '');
  var sigInput = header + '.' + payload;
  var sig = Utilities.base64EncodeWebSafe(
    Utilities.computeRsaSha256Signature(sigInput, privateKey)
  ).replace(/=+$/, '');
  var jwt = sigInput + '.' + sig;

  // Exchange JWT for access token
  var resp = UrlFetchApp.fetch('https://' + dsAuthHost_(s) + '/oauth/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + encodeURIComponent(jwt),
    muteHttpExceptions: true,
  });
  var body = JSON.parse(resp.getContentText());
  if (!body.access_token) {
    throw new Error('DocuSign JWT auth failed: ' + (body.error_description || body.error || 'Unknown error'));
  }
  props.setProperty('DS_JWT_TOKEN', body.access_token);
  props.setProperty('DS_JWT_TOKEN_EXP', String(Date.now() + 55 * 60 * 1000));
  return body.access_token;
}

// ── API helpers ───────────────────────────────────────────────────────────────

function dsRequest_(path, method, body) {
  var s     = getDocuSignSettings();
  var token = getDocuSignToken_();
  var url   = s.baseUrl + '/v2.1/accounts/' + s.accountId + path;
  var opts  = {
    method: method || 'get',
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
    muteHttpExceptions: true,
  };
  if (body) {
    opts.contentType = 'application/json';
    opts.payload = JSON.stringify(body);
  }
  var resp = UrlFetchApp.fetch(url, opts);
  var code = resp.getResponseCode();
  var text = resp.getContentText();
  if (code >= 400) {
    var err = {};
    try { err = JSON.parse(text); } catch(e) {}
    throw new Error(err.message || ('DocuSign API error ' + code + ': ' + text.slice(0, 200)));
  }
  return text ? JSON.parse(text) : {};
}

// ── Core functions (called from frontend) ────────────────────────────────────

/**
 * Returns envelopes in 'created' (draft) status — not yet sent.
 */
function getDraftEnvelopes() {
  var fromDate = new Date(Date.now() - 90*86400000).toISOString().slice(0, 10);
  var data = dsRequest_('/envelopes?status=created&from_date=' + fromDate + '&order_by=last_modified&order=desc&count=20');
  var envelopes = data.envelopes || [];
  return envelopes.map(function(env) {
    // Fetch recipients inline so queue cards can show signer avatars
    var signers = [];
    try {
      var recs = dsRequest_('/envelopes/' + env.envelopeId + '/recipients');
      signers = (recs.signers || []).map(function(r) {
        return { name: r.name, email: r.email, status: r.status || 'created' };
      });
    } catch(e) {}
    return {
      envelopeId:     env.envelopeId,
      subject:        env.emailSubject || '(no subject)',
      status:         env.status,
      created:        env.createdDateTime,
      lastModified:   env.lastModifiedDateTime,
      docCount:       parseInt(env.documentsCount  || '0', 10),
      recipientCount: parseInt(env.recipientsCount || '0', 10),
      signers:        signers,
    };
  });
}

/**
 * Uses Claude to extract key dates from an envelope subject + related email snippet.
 * Returns array of { label, date, daysUntil } objects.
 */
function extractKeyDates(subject, emailSnippet) {
  var apiKey = getApiKey_();
  if (!apiKey) return [];
  var today = new Date().toISOString().slice(0, 10);
  var context = 'Document: ' + (subject || '') + '\n' + (emailSnippet ? 'Email context: ' + emailSnippet : '');
  var prompt = [
    'Today is ' + today + '.',
    'Extract any important dates mentioned in the following contract/document context.',
    'Look for: expiry dates, renewal dates, termination dates, signature deadlines, effective dates, review dates.',
    'Return ONLY a JSON array of objects with keys: label (string, ≤30 chars), date (YYYY-MM-DD or null if unclear), note (string, ≤40 chars).',
    'Max 4 items. If no dates found, return [].',
    '',
    context,
  ].join('\n');

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 300,
      messages: [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });
  try {
    var result = JSON.parse(resp.getContentText());
    var text = result.content && result.content[0] && result.content[0].text || '[]';
    var match = text.match(/\[[\s\S]*\]/);
    var dates = JSON.parse(match ? match[0] : '[]');
    // Compute daysUntil for each
    var now = Date.now();
    return dates.map(function(d) {
      var daysUntil = null;
      if (d.date) {
        daysUntil = Math.round((new Date(d.date).getTime() - now) / 86400000);
      }
      return { label: d.label, date: d.date, note: d.note, daysUntil: daysUntil };
    });
  } catch(e) { return []; }
}

/**
 * Returns full details for one envelope: documents, recipients, existing message.
 */
function getEnvelopeDetails(envelopeId) {
  var env  = dsRequest_('/envelopes/' + envelopeId);
  var docs = dsRequest_('/envelopes/' + envelopeId + '/documents');
  var recs = dsRequest_('/envelopes/' + envelopeId + '/recipients');

  var documents = (docs.envelopeDocuments || [])
    .filter(function(d) { return d.type !== 'summary'; })
    .map(function(d) { return { id: d.documentId, name: d.name }; });

  var signers = (recs.signers || []).map(function(r) {
    return { name: r.name, email: r.email, routingOrder: r.routingOrder || 1 };
  });
  var ccList = (recs.carbonCopies || []).map(function(r) {
    return { name: r.name, email: r.email };
  });

  return {
    envelopeId:  envelopeId,
    subject:     env.emailSubject || '',
    existingMsg: env.emailBlurb   || '',
    status:      env.status,
    documents:   documents,
    signers:     signers,
    cc:          ccList,
    senderName:  (env.sender && env.sender.userName) || '',
    senderEmail: (env.sender && env.sender.email)    || '',
  };
}

/**
 * Uses Claude to draft the cover message for an envelope.
 * envelope: { subject, documents, signers, cc, senderName, senderEmail, existingMsg }
 */
function generateDocuSignMessage(envelope) {
  var apiKey = getApiKey_();
  if (!apiKey) throw new Error('Anthropic API key not configured.');

  var docNames  = (envelope.documents || []).map(function(d) { return d.name; }).join(', ') || envelope.subject;
  var signerStr = (envelope.signers || []).map(function(r) { return r.name + ' <' + r.email + '>'; }).join('; ');
  var ccStr     = (envelope.cc || []).map(function(r) { return r.name; }).join(', ');
  var senderCtx = envelope.senderName ? ('This is being sent by ' + envelope.senderName + '.') : '';
  var ccCtx     = ccStr ? ('The following people are CC\'d: ' + ccStr + '.') : '';

  var prompt = [
    'You write professional but warm DocuSign cover messages.',
    'Keep it concise — 4 to 6 sentences max. No bullet points. No subject line.',
    'Structure: greeting by first name → one-line description of what is being signed → any relevant approval/sender context → clear request to review and sign → polite sign-off.',
    'Do not use em-dashes. Do not use exclamation marks. Sound like a senior professional.',
    '',
    'Envelope details:',
    'Document(s): ' + docNames,
    'Recipient(s): ' + (signerStr || 'unknown'),
    senderCtx,
    ccCtx,
    '',
    'Write the message body only. No subject line. No "Subject:" prefix.',
  ].filter(Boolean).join('\n');

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: 'claude-sonnet-4-6',
      max_tokens: 400,
      messages: [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });
  var result = JSON.parse(resp.getContentText());
  if (!result.content || !result.content[0]) throw new Error('Claude did not return a message.');
  return { draft: result.content[0].text.trim() };
}

/**
 * Sets the cover message on the envelope and sends it.
 * message: the final text the user approved
 */
function sendDocuSignEnvelope(envelopeId, message) {
  // Update the email blurb (cover message)
  dsRequest_('/envelopes/' + envelopeId, 'put', { emailBlurb: message, status: 'sent' });
  return { ok: true, envelopeId: envelopeId };
}

/**
 * Update cover message only (without sending) — for save-as-draft.
 */
function updateDocuSignMessage(envelopeId, message) {
  dsRequest_('/envelopes/' + envelopeId, 'put', { emailBlurb: message });
  return { ok: true };
}

/**
 * Returns envelopes across all active statuses for the pipeline view.
 * Statuses: created, sent, delivered, completed (excludes voided/declined unless recent).
 */
function getEnvelopePipelineAll() {
  var fromDate = new Date(Date.now() - 30*86400000).toISOString().slice(0, 10);
  var statuses = ['created', 'sent', 'delivered', 'completed'];
  var all = [];
  statuses.forEach(function(status) {
    try {
      var data = dsRequest_('/envelopes?status=' + status + '&from_date=' + fromDate + '&order_by=last_modified&order=desc&count=10');
      (data.envelopes || []).forEach(function(env) { all.push(env); });
    } catch(e) {}
  });
  return all.map(function(env) {
    return {
      envelopeId:        env.envelopeId,
      subject:           env.emailSubject || '(no subject)',
      status:            env.status,
      lastModified:      env.lastModifiedDateTime,
      sentDateTime:      env.sentDateTime      || null,
      deliveredDateTime: env.deliveredDateTime || null,
      completedDateTime: env.completedDateTime || null,
      recipientCount:    parseInt(env.recipientsCount  || '0', 10),
      docCount:          parseInt(env.documentsCount   || '0', 10),
    };
  });
}

function _daysAgoIso_(days) {
  return new Date(Date.now() - days * 86400000).toISOString().slice(0, 10);
}

/**
 * Returns new signing events since the last call.
 * Compares current envelope statuses against a UserProperties cache.
 * Each returned event: { envelopeId, subject, status, signerName, signerEmail, timestamp, isNew }
 */
function getDocuSignNotifications() {
  var userProps = PropertiesService.getUserProperties();
  var cacheKey  = 'DS_STATUS_CACHE';
  var seenKey   = 'DS_SEEN_NOTIF';

  // Load cached statuses from last call
  var cached = {};
  try { cached = JSON.parse(userProps.getProperty(cacheKey) || '{}'); } catch(e) {}

  // Statuses that generate notifications
  var interestingStatuses = ['sent', 'delivered', 'completed', 'declined', 'voided'];
  var fromDate = _daysAgoIso_(14);
  var all = [];
  interestingStatuses.forEach(function(status) {
    try {
      var data = dsRequest_('/envelopes?status=' + status + '&from_date=' + fromDate + '&order_by=last_modified&order=desc&count=20');
      (data.envelopes || []).forEach(function(env) { all.push(env); });
    } catch(e) {}
  });

  // Also fetch recipients for completed envelopes to get signer names
  var events = [];
  var newCache = {};

  all.forEach(function(env) {
    var id     = env.envelopeId;
    var status = env.status;
    newCache[id] = status;

    // Build event if status changed or is new
    var prevStatus = cached[id];
    var isNew = prevStatus !== status;

    var timestamp = env.completedDateTime || env.deliveredDateTime || env.sentDateTime || env.lastModifiedDateTime || null;

    // Try to get signer details for completed envelopes
    var signerName = '';
    var signerEmail = '';
    if (status === 'completed' || status === 'delivered') {
      try {
        var recs = dsRequest_('/envelopes/' + id + '/recipients');
        var signers = recs.signers || [];
        // Find the first signer who has signed
        var signed = signers.filter(function(s) { return s.status === 'completed'; });
        if (signed.length) { signerName = signed[0].name; signerEmail = signed[0].email; }
        else if (signers.length) { signerName = signers[0].name; signerEmail = signers[0].email; }
      } catch(e) {}
    }

    events.push({
      envelopeId:  id,
      subject:     env.emailSubject || '(no subject)',
      status:      status,
      signerName:  signerName,
      signerEmail: signerEmail,
      timestamp:   timestamp,
      isNew:       isNew,
    });
  });

  // Persist updated cache
  try { userProps.setProperty(cacheKey, JSON.stringify(newCache)); } catch(e) {}

  // Sort: newest first
  events.sort(function(a, b) {
    return (b.timestamp || '') > (a.timestamp || '') ? 1 : -1;
  });

  return events;
}

/**
 * Marks all current notifications as seen (clears isNew flags by syncing cache).
 */
function markDocuSignNotificationsSeen() {
  var userProps = PropertiesService.getUserProperties();
  userProps.setProperty('DS_SEEN_NOTIF', String(Date.now()));
  return { ok: true };
}

/**
 * Searches Gmail for a thread whose subject matches the envelope subject.
 * Used to surface related email context in the detail panel.
 */
function searchRelatedEmailForEnvelope(subject) {
  var cleaned = String(subject || '').replace(/^re:\s*/i, '').replace(/"/g, '').trim().substring(0, 80);
  if (!cleaned) return null;
  try {
    var threads = GmailApp.search('subject:("' + cleaned + '")', 0, 3);
    if (!threads || !threads.length) return null;
    var thread = threads[0];
    var msgs = thread.getMessages();
    if (!msgs || !msgs.length) return null;
    var lastMsg = msgs[msgs.length - 1];
    var sender = parseSender_(lastMsg.getFrom());
    return {
      subject:      thread.getFirstMessageSubject() || subject,
      from:         sender.name || sender.email,
      fromEmail:    sender.email,
      date:         lastMsg.getDate().toISOString(),
      snippet:      lastMsg.getPlainBody().substring(0, 240).replace(/\s+/g, ' ').trim(),
      messageCount: msgs.length,
    };
  } catch(e) {
    return null;
  }
}

/**
 * Connection test — returns account info.
 */
function testDocuSignConnection() {
  var token = getDocuSignToken_();
  var s     = getDocuSignSettings();
  // Verify account info endpoint
  var resp  = UrlFetchApp.fetch('https://' + dsAuthHost_(s) + '/oauth/userinfo', {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  var info = JSON.parse(resp.getContentText());
  if (!info.sub) throw new Error('DocuSign auth failed — check your credentials.');
  var acct = (info.accounts || []).find(function(a) { return a.account_id === s.accountId; })
          || (info.accounts || [])[0]
          || {};
  return { ok: true, name: info.name, email: info.email, accountName: acct.account_name || '' };
}
