// ---------------------------------------------------------------------------
// Inbox Triage - Gmail Add-on + Web App
// ---------------------------------------------------------------------------

const API_KEY_PROP = 'ANTHROPIC_API_KEY';
const MODEL = 'claude-haiku-4-5-20251001';
const CLASSIFICATION_CACHE_PROP = 'EMAIL_CLASS_CACHE';
const CLASSIFICATION_CACHE_TTL_MS = 24 * 3600 * 1000;
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbyFMxu_SUl9OENyx88_dev2NAgDOspfLxHc2OVV1b96mT03iFfIbronX4SsSijpcUOEZQ/exec';

const CAT_META = {
  escalation:     { label: 'Escalation',     emoji: '🔥', icon: CardService.Icon.STAR },
  action_required:{ label: 'Action Required',emoji: '🔴', icon: CardService.Icon.EMAIL },
  calendar:       { label: 'Calendar',       emoji: '📅', icon: CardService.Icon.INVITE },
  awaiting:       { label: 'Awaiting',       emoji: '⏳', icon: CardService.Icon.DESCRIPTION },
  digest:         { label: 'Digest',         emoji: '📥', icon: CardService.Icon.EMAIL },
};

const CAT_ORDER = {
  escalation: 0,
  action_required: 1,
  calendar: 2,
  awaiting: 3,
  digest: 4,
};

const PRIORITY_LABELS = {
  5: '🔴 Critical',
  4: '🟠 High',
  3: '🟡 Medium',
  2: '🔵 Low',
  1: '⚪ Minimal',
};

const CAT_COLORS = {
  escalation:      '#D45C5C',
  action_required: '#C49A40',
  calendar:        '#00A651',
  awaiting:        '#004D3A',
  digest:          '#7C8C88',
};

const TRIAGE_LABEL_NAMES = {
  escalation:      'Triage/Escalation',
  action_required: 'Triage/Action Required',
  calendar:        'Triage/Calendar',
  awaiting:        'Triage/Awaiting',
  digest:          'Triage/Digest',
};

const RESOLVED_LABEL_NAME = 'Triage/Resolved';
const VALID_CATEGORIES = Object.keys(TRIAGE_LABEL_NAMES);
const VIP_SENDERS_PROP = 'VIP_SENDERS';

// ClickUp / Slack integration property keys.
// Declared at top-level (before any function definitions) to avoid temporal
// dead zone errors — functions such as getTaskHistory() reference these.
const CLICKUP_API_KEY_PROP   = 'CLICKUP_API_KEY';
const CLICKUP_LIST_ID_PROP   = 'CLICKUP_LIST_ID';  // legacy single-list (kept for migration)
const CLICKUP_LISTS_PROP     = 'CLICKUP_LISTS';
const CLICKUP_STUCK_DAYS_PROP = 'CLICKUP_STUCK_DAYS';
const SLACK_BOT_TOKEN_PROP   = 'SLACK_BOT_TOKEN';
const SLACK_OVERRIDE_CHANNEL_PROP = 'SLACK_OVERRIDE_CHANNEL';
const TASK_HISTORY_PROP      = 'TASK_HISTORY';

function getUserProps_() {
  return PropertiesService.getUserProperties();
}

function parseJson_(value, fallback) {
  try {
    return JSON.parse(value || '');
  } catch (e) {
    return fallback;
  }
}

function getApiKey_() {
  // User-level key overrides the shared script-level key
  var userKey = PropertiesService.getUserProperties().getProperty(API_KEY_PROP);
  if (userKey && userKey.trim()) return userKey.trim();
  return PropertiesService.getScriptProperties().getProperty(API_KEY_PROP);
}

function requireApiKey_() {
  const key = getApiKey_();
  if (!key) throw new Error('No API key set. Configure it in Settings or set ANTHROPIC_API_KEY in Script Properties.');
  return key;
}

function saveAnthropicKey(key) {
  var k = String(key || '').trim();
  if (!k) throw new Error('API key cannot be empty.');
  PropertiesService.getUserProperties().setProperty(API_KEY_PROP, k);
  return { ok: true };
}

function parseSender_(from) {
  const raw = from || '';
  const match = raw.match(/<([^>]+)>/);
  const email = (match ? match[1] : raw).trim();
  const name = raw.replace(/<[^>]+>/g, '').trim().replace(/^["']+|["']+$/g, '').trim() || email;
  return { name, email };
}

function normalizeCategory_(category) {
  return VALID_CATEGORIES.indexOf(category) === -1 ? 'digest' : category;
}

function normalizePriority_(priority) {
  const n = parseInt(priority, 10);
  return isNaN(n) ? 1 : Math.max(1, Math.min(5, n));
}

function normalizeEmailOrDomain_(value) {
  return String(value || '').toLowerCase().trim().replace(/^@/, '');
}

function extractJsonArray_(text) {
  const cleaned = String(text || '').replace(/```(?:json)?/gi, '').trim();
  const match = cleaned.match(/\[[\s\S]*\]/);
  if (!match) throw new Error('Could not parse AI response: ' + cleaned.substring(0, 200));
  return JSON.parse(match[0]);
}

function callClaude_(apiKey, system, user, maxTokens, model) {
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify({
      model: model || MODEL,
      max_tokens: maxTokens || 1000,
      temperature: 0,
      system,
      messages: [{ role: 'user', content: user }],
    }),
  });

  const body = res.getContentText();
  const data = parseJson_(body, null);
  if (!data) throw new Error('Invalid API response. HTTP ' + res.getResponseCode());
  if (data.error) throw new Error('API error: ' + data.error.message);
  const text = (data.content || []).map(function(block) { return block.text || ''; }).join('').trim();
  if (!text) throw new Error('Empty API response. HTTP ' + res.getResponseCode());
  return text;
}

function getLatestInboxEmails_(limit) {
  const threads = GmailApp.getInboxThreads(0, limit || 15);
  const allMessages = GmailApp.getMessagesForThreads(threads);

  return threads.map(function(thread, i) {
    const msgs = allMessages[i] || [];
    const msg = msgs[msgs.length - 1];
    if (!msg) return null;
    const sender = parseSender_(msg.getFrom());
    return {
      id: msg.getId(),
      subject: thread.getFirstMessageSubject() || '(no subject)',
      sender_name: sender.name,
      sender_email: sender.email,
      date: msg.getDate().toISOString(),
    };
  }).filter(Boolean);
}

// ---------------------------------------------------------------------------
// Smart labels
// ---------------------------------------------------------------------------

function getOrCreateLabel(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function applyTriageLabels(emails) {
  const labelCache = {};
  const getLabel = function(name) {
    if (!labelCache[name]) labelCache[name] = getOrCreateLabel(name);
    return labelCache[name];
  };

  emails.forEach(function(email) {
    const targetName = TRIAGE_LABEL_NAMES[email.category];
    if (!targetName || email.isFollowup) return;

    try {
      const thread = GmailApp.getMessageById(email.id).getThread();
      Object.keys(TRIAGE_LABEL_NAMES).forEach(function(category) {
        const name = TRIAGE_LABEL_NAMES[category];
        if (name === targetName) return;
        const existing = GmailApp.getUserLabelByName(name);
        if (existing) thread.removeLabel(existing);
      });
      thread.addLabel(getLabel(targetName));
    } catch (e) {
      console.warn('applyTriageLabels: could not label ' + email.id + ' - ' + e.message);
    }
  });
}

function removeTriageLabels(thread) {
  Object.keys(TRIAGE_LABEL_NAMES).forEach(function(category) {
    const label = GmailApp.getUserLabelByName(TRIAGE_LABEL_NAMES[category]);
    if (label) thread.removeLabel(label);
  });
}

function getSmartLabelsEnabled() {
  return getUserProps_().getProperty('SMART_LABELS_ENABLED') !== 'false';
}

function setSmartLabelsEnabled(val) {
  const enabled = val === true || val === 'true';
  getUserProps_().setProperty('SMART_LABELS_ENABLED', enabled ? 'true' : 'false');
  return enabled;
}

// ---------------------------------------------------------------------------
// VIP senders
// ---------------------------------------------------------------------------

function getVipSenders() {
  return parseJson_(getUserProps_().getProperty(VIP_SENDERS_PROP), []);
}

function setVipSenders_(senders) {
  const cleaned = Array.from(new Set((senders || [])
    .map(normalizeEmailOrDomain_)
    .filter(Boolean)))
    .slice(0, 200);
  getUserProps_().setProperty(VIP_SENDERS_PROP, JSON.stringify(cleaned));
  return cleaned;
}

function addVipSender(sender) {
  const value = normalizeEmailOrDomain_(sender);
  if (!value) return getVipSenders();
  const senders = getVipSenders();
  if (senders.indexOf(value) === -1) senders.push(value);
  return setVipSenders_(senders);
}

function removeVipSender(sender) {
  const value = normalizeEmailOrDomain_(sender);
  return setVipSenders_(getVipSenders().filter(function(item) { return item !== value; }));
}

function isVipSender_(email, vipSenders) {
  const sender = normalizeEmailOrDomain_(email);
  if (!sender) return false;
  const domain = sender.indexOf('@') === -1 ? sender : sender.split('@').pop();
  return (vipSenders || []).some(function(vip) {
    const value = normalizeEmailOrDomain_(vip);
    return sender === value || domain === value;
  });
}

// ---------------------------------------------------------------------------
// Task extraction
// ---------------------------------------------------------------------------

function getTasks() {
  return parseJson_(getUserProps_().getProperty('TASK_DATA'), []);
}

function getTaskHistory() {
  return parseJson_(PropertiesService.getUserProperties().getProperty(TASK_HISTORY_PROP), []);
}


function addTaskToClickUp(messageId, listId, tldr) {
  var apiKey = requireApiKey_();
  var msg = GmailApp.getMessageById(messageId);
  var sender = parseSender_(msg.getFrom());
  var subject = msg.getThread().getFirstMessageSubject();
  var body = cleanEmailBody_(msg.getPlainBody()).substring(0, 800);
  var threadId = msg.getThread().getId();
  var todayIso = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  var text = callClaude_(
    apiKey,
    'You are a JSON-only task extraction API. Return only a valid JSON array. No markdown, no prose.',
    "Today's date is " + todayIso + ". Identify the single most important action item the recipient must do from this email.\n" +
    "Return ONLY a JSON array with exactly ONE object: [{description (verb-first, max 12 words), dueDate (ISO yyyy-MM-dd date resolved from any relative phrases like 'by Friday' or 'next week' using today's date, or null if none mentioned), priority ('high'|'medium'|'low'), assigneeHint (first name of person who should handle this, or null)}]\n" +
    "If no clear action item exists, return [].\n\n" +
    "Subject: " + subject + "\nFrom: " + sender.name + "\nBody: " + body,
    400,
    MODEL
  );

  var allExtracted = parseJson_(text.match(/\[[\s\S]*\]/) ? text.match(/\[[\s\S]*\]/)[0] : '[]', []);
  if (!allExtracted.length) return { tasks: [], source: 'none' };
  // Hard limit: exactly 1 task per email
  var extracted = [allExtracted[0]];

  var cu = getClickUpSettings();
  var priorityMap = { high: 2, medium: 3, low: 4 };

  if (!cu.apiKey) {
    return { tasks: [], source: 'no_key', message: 'Add your ClickUp API key in Settings to sync tasks.' };
  }

  // Use explicitly passed listId, fall back to first configured list
  var targetListId = String(listId || '').replace(/\D/g, '').trim();
  if (!targetListId) {
    var configured = (cu.lists || []).filter(function(l) { return l.id; });
    if (!configured.length) {
      return { tasks: [], source: 'no_list', message: 'No ClickUp list configured. Add List IDs in Settings.' };
    }
    targetListId = configured[0].id;
  }

  var headers = { 'Authorization': cu.apiKey, 'Content-Type': 'application/json' };
  var threadUrl = 'https://mail.google.com/mail/#all/' + threadId;
  var created = [];
  var lastError = '';

  for (var ti = 0; ti < extracted.length; ti++) {
    var t = extracted[ti];
    var priority = ['high','medium','low'].indexOf(t.priority) === -1 ? 'medium' : t.priority;
    var descParts = [];
    if (tldr) descParts.push(tldr);
    descParts.push('From: ' + sender.name + ' <' + sender.email + '>');
    descParts.push('Subject: ' + subject);
    if (t.assigneeHint) descParts.push('Suggested assignee: ' + t.assigneeHint);
    descParts.push('Gmail thread: ' + threadUrl);
    var payload = {
      name: String(t.description || '').substring(0, 120),
      description: descParts.join('\n'),
      priority: priorityMap[priority],
    };
    if (t.dueDate) {
      try { payload.due_date = new Date(t.dueDate).getTime(); } catch(e) {}
    }
    try {
      var res = UrlFetchApp.fetch('https://api.clickup.com/api/v2/list/' + targetListId + '/task', {
        method: 'post',
        muteHttpExceptions: true,
        headers: headers,
        payload: JSON.stringify(payload),
      });
      var data = parseJson_(res.getContentText(), null);
      if (data && data.id) {
        created.push({ id: data.id, name: data.name, url: data.url || '', priority: priority });
      } else {
        lastError = (data && data.err) ? data.err : ('HTTP ' + res.getResponseCode());
        if (lastError && lastError.toLowerCase().indexOf('listid') !== -1) {
          PropertiesService.getUserProperties().deleteProperty(CLICKUP_LIST_ID_PROP);
          return { tasks: [], source: 'error', message: 'Invalid List ID for this list. Open Settings → ClickUp Lists and verify the numeric ID.' };
        }
        console.warn('ClickUp task creation failed: ' + lastError);
      }
    } catch(e) {
      lastError = e.message;
      console.warn('ClickUp task creation failed: ' + e.message);
    }
  }

  if (created.length) {
    var listEntry = null;
    (cu.lists || []).forEach(function(l) {
      if (l.id === targetListId) { listEntry = l; return; }
      (l.sublists || []).forEach(function(sl) { if (sl.id === targetListId) listEntry = sl; });
    });
    var histEntry = {
      id: created[0].id,
      name: created[0].name,
      url: created[0].url || '',
      listId: targetListId,
      listName: listEntry ? listEntry.name : '',
      createdAt: new Date().toISOString(),
      subject: subject,
    };
    try {
      var hist = parseJson_(PropertiesService.getUserProperties().getProperty(TASK_HISTORY_PROP), []);
      hist.unshift(histEntry);
      if (hist.length > 200) hist = hist.slice(0, 200);
      PropertiesService.getUserProperties().setProperty(TASK_HISTORY_PROP, JSON.stringify(hist));
    } catch(he) {}
    return { tasks: created, source: 'clickup' };
  }
  if (lastError) return { tasks: [], source: 'error', message: 'ClickUp error: ' + lastError };
  return { tasks: [], source: 'none' };
}

function markTaskDone(taskId, done) {
  const props = getUserProps_();
  const tasks = getTasks().map(function(task) {
    return task.id === taskId ? Object.assign({}, task, { done: done === true || done === 'true' }) : task;
  });
  props.setProperty('TASK_DATA', JSON.stringify(tasks));
  return tasks;
}

function clearDoneTasks() {
  const props = getUserProps_();
  const tasks = getTasks().filter(function(t) { return !t.done; });
  props.setProperty('TASK_DATA', JSON.stringify(tasks));
  return tasks;
}

// ---------------------------------------------------------------------------
// Follow-up tracker
// ---------------------------------------------------------------------------

function addFollowup(messageId, subject, senderName, senderEmail, emailDate, hours) {
  const props = getUserProps_();
  const data = parseJson_(props.getProperty('FOLLOWUP_DATA'), {});
  const now = new Date();
  data[messageId] = {
    subject,
    sender_name: senderName,
    sender_email: senderEmail,
    date: emailDate,
    addedAt: now.toISOString(),
    followupAt: new Date(now.getTime() + Number(hours || 24) * 3600000).toISOString(),
  };
  props.setProperty('FOLLOWUP_DATA', JSON.stringify(data));
  return 'added';
}

function removeFollowup(messageId) {
  const props = getUserProps_();
  const data = parseJson_(props.getProperty('FOLLOWUP_DATA'), {});
  delete data[messageId];
  props.setProperty('FOLLOWUP_DATA', JSON.stringify(data));
  return 'removed';
}

function processFollowups() {
  const props = getUserProps_();
  const data = parseJson_(props.getProperty('FOLLOWUP_DATA'), {});
  const keys = Object.keys(data);
  if (!keys.length) return { overdue: [], activeIds: [] };

  const now = new Date();
  const myEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const overdue = [];
  const activeIds = [];
  const toRemove = [];

  keys.forEach(function(id) {
    const followup = data[id];
    try {
      const messages = GmailApp.getMessageById(id).getThread().getMessages();
      const addedDate = new Date(followup.addedAt);
      const hasReply = messages.some(function(message) {
        const from = message.getFrom().toLowerCase();
        return message.getDate() > addedDate && (!myEmail || from.indexOf(myEmail) === -1);
      });

      if (hasReply) {
        toRemove.push(id);
        return;
      }

      const dueDate = new Date(followup.followupAt);
      if (now > dueDate) {
        overdue.push(Object.assign({}, followup, {
          id,
          hoursOverdue: Math.max(1, Math.round((now - dueDate) / 3600000)),
        }));
      } else {
        activeIds.push(id);
      }
    } catch (e) {
      toRemove.push(id);
    }
  });

  if (toRemove.length) {
    toRemove.forEach(function(id) { delete data[id]; });
    props.setProperty('FOLLOWUP_DATA', JSON.stringify(data));
  }

  return { overdue, activeIds };
}


// ---------------------------------------------------------------------------
// Custom rules
// ---------------------------------------------------------------------------

function getRules() {
  return parseJson_(getUserProps_().getProperty('TRIAGE_RULES'), []);
}

function addRule(rule) {
  const rules = getRules();
  rules.push({
    id: 'r_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
    condition: rule.condition,
    value: String(rule.value || '').trim(),
    category: normalizeCategory_(rule.category),
    priority: normalizePriority_(rule.priority || 3),
  });
  getUserProps_().setProperty('TRIAGE_RULES', JSON.stringify(rules));
  return rules;
}

function deleteRule(id) {
  const rules = getRules().filter(function(rule) { return rule.id !== id; });
  getUserProps_().setProperty('TRIAGE_RULES', JSON.stringify(rules));
  return rules;
}

function getLearnedRules() {
  return parseJson_(getUserProps_().getProperty('LEARNED_RULES'), []);
}

function learnFromCorrection(senderEmail, senderName, subject, newCategory) {
  var cat = normalizeCategory_(newCategory);
  var email = String(senderEmail || '').toLowerCase().trim();
  if (!email || !cat) return [];

  var props = getUserProps_();
  var learned = parseJson_(props.getProperty('LEARNED_RULES'), []);

  var existing = null;
  for (var i = 0; i < learned.length; i++) {
    if (learned[i].condition === 'sender_email' && learned[i].value === email) {
      existing = learned[i];
      break;
    }
  }

  if (existing) {
    existing.category = cat;
    existing.count = (existing.count || 1) + 1;
  } else {
    learned.push({
      id: 'lr_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
      condition: 'sender_email',
      value: email,
      category: cat,
      priority: 3,
      learned: true,
      learnedFrom: String(subject || '').slice(0, 60),
      count: 1,
    });
  }

  props.setProperty('LEARNED_RULES', JSON.stringify(learned));
  return learned;
}

function deleteLearnedRule(id) {
  var props = getUserProps_();
  var learned = parseJson_(props.getProperty('LEARNED_RULES'), []).filter(function(r) { return r.id !== id; });
  props.setProperty('LEARNED_RULES', JSON.stringify(learned));
  return learned;
}

function applyRules(emails) {
  const customRules = getRules();
  const learnedRules = getLearnedRules();
  const allRules = customRules.concat(learnedRules);
  if (!allRules.length) return emails;

  return emails.map(function(email) {
    for (let i = 0; i < allRules.length; i++) {
      const rule = allRules[i];
      if (matchesRule(email, rule)) {
        return Object.assign({}, email, {
          category: normalizeCategory_(rule.category),
          priority: normalizePriority_(rule.priority),
          summary: email.summary || 'Matched custom rule.',
        });
      }
    }
    return email;
  });
}

function enrichAttentionSignals_(emails) {
  const vipSenders = getVipSenders();
  return (emails || []).map(function(email) {
    const existing = email.attentionSignals || {};
    const category = normalizeCategory_(email.category);
    const priority = normalizePriority_(email.priority);
    const text = [
      email.subject || '',
      email.summary || '',
      email.sender_name || '',
      email.sender_email || '',
    ].join(' ').toLowerCase();

    const directAsk = existing.directAsk === true ||
      category === 'action_required' ||
      /\b(can you|could you|please|pls|need you to|action required|approval needed|please approve|review|sign off|respond|reply)\b/.test(text);
    const deadlineDriven = existing.deadlineDriven === true ||
      /\b(today|tomorrow|eod|end of day|asap|urgent|deadline|due|by \w+day|before \w+day|this week|next week|overdue)\b/.test(text);
    const criticalEscalation = existing.criticalEscalation === true ||
      category === 'escalation' ||
      priority >= 5 ||
      /\b(escalat|critical|blocker|sev[ -]?[0-2]|production down|outage|urgent)\b/.test(text);
    const vip = isVipSender_(email.sender_email, vipSenders);
    const overdueFollowup = existing.overdueFollowup === true || email.isFollowup === true;

    const reasons = [];
    if (criticalEscalation) reasons.push('Critical escalation');
    if (directAsk) reasons.push('Direct ask');
    if (deadlineDriven) reasons.push('Deadline-driven');
    if (vip) reasons.push('VIP sender');
    if (overdueFollowup) reasons.push('Overdue follow-up');

    return Object.assign({}, email, {
      attentionSignals: {
        criticalEscalation,
        directAsk,
        deadlineDriven,
        vip,
        overdueFollowup,
      },
      attentionReasons: reasons,
      isAttention: reasons.length > 0,
    });
  });
}

function matchesRule(email, rule) {
  const val = String(rule.value || '').toLowerCase().trim();
  if (!val) return false;

  const senderEmail = String(email.sender_email || '').toLowerCase();
  const senderName = String(email.sender_name || '').toLowerCase();
  const subject = String(email.subject || '').toLowerCase();

  switch (rule.condition) {
    case 'sender_domain':
      return senderEmail.indexOf('@' + val.replace(/^@/, '')) !== -1;
    case 'sender_email':
      return senderEmail === val;
    case 'sender_contains':
      return senderEmail.indexOf(val) !== -1 || senderName.indexOf(val) !== -1;
    case 'subject_contains':
      return subject.indexOf(val) !== -1;
    default:
      return false;
  }
}

// ---------------------------------------------------------------------------
// Email body / contact photos
// ---------------------------------------------------------------------------

function cleanEmailBody_(text) {
  if (!text) return '';

  // Truncate at common footer separators
  ['~~//~~', '~-~-~', '\n-- \n', '\n--\n'].forEach(function(marker) {
    var idx = text.indexOf(marker);
    if (idx > 80) text = text.substring(0, idx);
  });

  // Normalise line endings so all regexes work consistently
  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  // Strip Organizer/Guests block — cut everything from that line onward
  var orgIdx = text.search(/\nOrganizer\n/);
  if (orgIdx > 0) text = text.substring(0, orgIdx);

  // Strip other Google Calendar / notification boilerplate
  text = text.replace(/You are receiving this email[\s\S]*$/im, '');
  text = text.replace(/Invitation from Google Calendar[\s\S]*$/im, '');
  text = text.replace(/Reply for [^\n]+/gi, '');
  text = text.replace(/View all guest info[^\n]*/gi, '');
  text = text.replace(/Your attendance is (?:optional|required)[^\n]*/gi, '');

  // Remove bare long URLs (calendar event links, tracking pixels, etc.)
  text = text.replace(/https?:\/\/\S{60,}/g, '');

  // Collapse excess whitespace
  text = text.replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
  return text;
}

// Converts a PDF or Office doc blob to plain text via Drive's import conversion.
// Requires Drive API Advanced Service enabled (Services → Drive API in GAS editor).
function extractDocText_(blob, mimeType) {
  try {
    if (!blob || blob.getBytes().length > 6 * 1024 * 1024) return null; // skip >6 MB
    var file = Drive.Files.insert(
      { title: 'qf_tmp_' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
      blob,
      { convert: true }
    );
    var doc = DocumentApp.openById(file.id);
    var text = doc.getBody().getText();
    try { Drive.Files.remove(file.id); } catch(e) {}
    return text.substring(0, 3000).trim() || null;
  } catch(e) {
    return null;
  }
}

function getEmailBody(messageId) {
  const msg = GmailApp.getMessageById(messageId);
  const thread = msg.getThread();
  const messages = thread.getMessages().slice(-6);
  let myEmail = '';
  try { myEmail = (Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) {}

  const TEXT_MIME_TYPES = ['text/plain','text/csv','text/html','text/xml','application/json','application/xml'];
  const DOC_MIME_TYPES = [
    'application/pdf',
    'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ];

  const structured = messages.map(function(m, i) {
    const sender = parseSender_(m.getFrom());

    // Collect attachments — extract text from PDFs/Office docs on recent messages only
    var rawAttachments = [];
    try { rawAttachments = m.getAttachments({includeInlineImages: false, includeAttachments: true}); } catch(e) {}
    var isRecent = i >= messages.length - 3; // only extract from last 3 messages to avoid timeout
    var attachments = rawAttachments.map(function(a) {
      var name = a.getName() || 'attachment';
      var size = a.getSize();
      var mimeType = a.getContentType() || 'application/octet-stream';
      var text = null;
      if (TEXT_MIME_TYPES.indexOf(mimeType) !== -1) {
        try { text = a.getDataAsString().substring(0, 500); } catch(e) {}
      } else if (isRecent && DOC_MIME_TYPES.indexOf(mimeType) !== -1) {
        try { text = extractDocText_(a.copyBlob(), mimeType); } catch(e) {}
      }
      return { name: name, size: size, mimeType: mimeType, text: text };
    });

    return {
      index: i,
      from: sender.name,
      fromEmail: sender.email,
      date: m.getDate().toISOString(),
      body: cleanEmailBody_(m.getPlainBody()) || '(No content)',
      htmlBody: (function() {
        var h = m.getBody() || '';
        // Strip base64 inline images (they can be MB each and blow the serialization limit)
        h = h.replace(/src="data:[^"]{0,8}[^"]*"/g, 'src=""');
        h = h.replace(/src='data:[^']{0,8}[^']*'/g, "src=''");
        return h;
      })(),
      attachments: attachments,
      isMe: !!(myEmail && sender.email.toLowerCase() === myEmail),
      bullets: null,
    };
  });

  const apiKey = getApiKey_();
  if (apiKey) {
    try {
      const payload = structured.map(function(m) {
        var attachContext = m.attachments.length > 0
          ? ' [Attachments: ' + m.attachments.map(function(a) {
              if (a.text) return a.name + ' (extracted content: ' + a.text.substring(0, 600) + ')';
              return a.name + ' (binary, not extracted)';
            }).join('; ') + ']'
          : '';
        return { index: m.index, from: m.from, body: m.body.substring(0, 800) + attachContext };
      });
      const text = callClaude_(
        apiKey,
        'You are a JSON-only API. Return only a valid JSON array. No markdown, no prose.',
        'Summarize each email message into 3-5 bullet points. Each bullet must be under 15 words, factual, and start with a verb or key noun. Return a JSON array — each element: index (number), bullets (string array).\n\n' + JSON.stringify(payload),
        800
      );
      const summaries = extractJsonArray_(text);
      summaries.forEach(function(s) {
        if (structured[s.index] && Array.isArray(s.bullets)) {
          structured[s.index].bullets = s.bullets.slice(0, 5);
        }
      });
    } catch (e) {
      // Fall back to raw body if summarization fails
    }
  }

  // Fetch profile photos for all senders
  var senderEmails = structured.map(function(m) { return m.fromEmail; }).filter(Boolean);
  var uniqueSenderEmails = senderEmails.filter(function(e, i) { return senderEmails.indexOf(e) === i; });
  var photoMap = {};
  try { photoMap = uniqueSenderEmails.length ? fetchContactPhotos(uniqueSenderEmails) : {}; } catch(e) {}

  return structured.map(function(m) {
    return {
      from: m.from,
      fromEmail: m.fromEmail,
      date: m.date,
      body: m.body,
      htmlBody: m.htmlBody,
      attachments: m.attachments,
      bullets: m.bullets,
      isMe: m.isMe,
      photo: photoMap[m.fromEmail] || null,
    };
  });
}

function fetchContactPhotos(senderEmails) {
  const emails = (senderEmails || []).filter(Boolean).slice(0, 30);
  if (!emails.length) return {};

  const token = ScriptApp.getOAuthToken();

  // Primary: searchContacts (saved contacts + "other contacts" = everyone you've emailed)
  const contactRequests = emails.map(function(email) {
    return {
      url: 'https://people.googleapis.com/v1/people:searchContacts?query=' +
        encodeURIComponent(email) +
        '&readMask=photos,emailAddresses&pageSize=3' +
        '&sources=READ_SOURCE_TYPE_CONTACT' +
        '&sources=READ_SOURCE_TYPE_OTHER_CONTACT',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    };
  });

  const photoMap = {};

  UrlFetchApp.fetchAll(contactRequests).forEach(function(res, i) {
    try {
      const data = JSON.parse(res.getContentText());
      const results = data.results || [];
      // Find the result whose email address matches exactly
      for (var r = 0; r < results.length; r++) {
        var person = results[r].person || {};
        var emailAddrs = (person.emailAddresses || []).map(function(ea) { return (ea.value || '').toLowerCase(); });
        if (emailAddrs.indexOf(emails[i].toLowerCase()) === -1 && results.length > 1) continue;
        var photos = person.photos || [];
        var photo = photos.find(function(p) { return !p.default; }) || photos[0];
        if (photo && photo.url) { photoMap[emails[i]] = photo.url; break; }
      }
    } catch (e) { /* optional */ }
  });

  // Supplement: searchDirectoryPeople for Workspace org members not in contacts
  var missing = emails.filter(function(e) { return !photoMap[e]; });
  if (missing.length) {
    var dirRequests = missing.map(function(email) {
      return {
        url: 'https://people.googleapis.com/v1/people:searchDirectoryPeople?query=' +
          encodeURIComponent(email) +
          '&readMask=photos,emailAddresses&pageSize=1' +
          '&sources=DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE' +
          '&sources=DIRECTORY_SOURCE_TYPE_DOMAIN_CONTACT',
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true,
      };
    });
    UrlFetchApp.fetchAll(dirRequests).forEach(function(res, i) {
      try {
        const data = JSON.parse(res.getContentText());
        const people = data.people || [];
        if (!people.length) return;
        const photos = people[0].photos || [];
        const photo = photos.find(function(p) { return !p.default; }) || photos[0];
        if (photo && photo.url) photoMap[missing[i]] = photo.url;
      } catch (e) { /* optional */ }
    });
  }

  return photoMap;
}

function fetchContactPhotosForEmails(emails) {
  const uniqueEmails = Array.from(new Set((emails || [])
    .map(function(email) { return email.sender_email; })
    .filter(Boolean)));
  return fetchContactPhotos(uniqueEmails);
}

// ---------------------------------------------------------------------------
// Dismissed email persistence
// ---------------------------------------------------------------------------

function getDismissedIds() {
  return parseJson_(getUserProps_().getProperty('DISMISSED_IDS'), []);
}

function addDismissedId(messageId) {
  const props = getUserProps_();
  const ids = getDismissedIds();
  if (ids.indexOf(messageId) === -1) {
    ids.push(messageId);
    if (ids.length > 500) ids.splice(0, ids.length - 500);
    props.setProperty('DISMISSED_IDS', JSON.stringify(ids));
  }
}

function storeDismissed(messageId, action, msg, thread) {
  const props = getUserProps_();
  const data = parseJson_(props.getProperty('DISMISSED_DATA'), {});
  const sender = parseSender_(msg.getFrom());
  data[messageId] = {
    action,
    subject: thread.getFirstMessageSubject() || '(no subject)',
    sender_name: sender.name,
    sender_email: sender.email,
    date: msg.getDate().toISOString(),
    dismissedAt: new Date().toISOString(),
  };

  const keys = Object.keys(data);
  if (keys.length > 200) {
    keys
      .sort(function(a, b) { return new Date(data[a].dismissedAt) - new Date(data[b].dismissedAt); })
      .slice(0, keys.length - 200)
      .forEach(function(key) { delete data[key]; });
  }

  props.setProperty('DISMISSED_DATA', JSON.stringify(data));
  addDismissedId(messageId);
}


function clearDismissedEmails() {
  const props = getUserProps_();
  props.deleteProperty('DISMISSED_IDS');
  props.deleteProperty('DISMISSED_DATA');
  return 'cleared';
}


function resolveEmail(messageId) {
  const msg = GmailApp.getMessageById(messageId);
  const thread = msg.getThread();
  removeTriageLabels(thread);
  thread.addLabel(getOrCreateLabel(RESOLVED_LABEL_NAME));
  thread.moveToArchive();
  storeDismissed(messageId, 'resolve', msg, thread);
  removeFollowup(messageId);
  removeFromClassificationCache_(messageId);
  return 'resolved';
}

function trashEmail(messageId) {
  const msg = GmailApp.getMessageById(messageId);
  const thread = msg.getThread();
  removeTriageLabels(thread);
  thread.moveToTrash();
  storeDismissed(messageId, 'trash', msg, thread);
  removeFollowup(messageId);
  removeFromClassificationCache_(messageId);
  return 'trashed';
}

// ---------------------------------------------------------------------------
// Web app
// ---------------------------------------------------------------------------

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('QuestFlow')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getCurrentUser() {
  try {
    const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
    let name = '';
    let photo = null;
    try {
      const token = ScriptApp.getOAuthToken();
      const uiRes = UrlFetchApp.fetch(
        'https://www.googleapis.com/oauth2/v3/userinfo',
        { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
      );
      if (uiRes.getResponseCode() === 200) {
        const ui = JSON.parse(uiRes.getContentText());
        // Prefer full name over given_name (given_name may just be email prefix)
        const fullName = ui.name || '';
        const givenName = ui.given_name || '';
        // Use full name first word if it looks like a real name (not email-like)
        if (fullName && !/[@0-9_]/.test(fullName)) {
          name = fullName.split(/\s+/)[0];
        } else if (givenName && !/[@0-9_]/.test(givenName)) {
          name = givenName;
        }
        if (ui.picture) photo = ui.picture;
      }
    } catch (e) {}
    // Try People API for display name
    if (!name || /[@0-9_]/.test(name)) {
      try {
        const token = ScriptApp.getOAuthToken();
        const pRes = UrlFetchApp.fetch(
          'https://people.googleapis.com/v1/people/me?personFields=names',
          { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
        );
        if (pRes.getResponseCode() === 200) {
          const p = JSON.parse(pRes.getContentText());
          const names = p.names || [];
          const primary = names.find(function(n){ return n.metadata && n.metadata.primary; }) || names[0];
          if (primary && primary.givenName && !/[@0-9_]/.test(primary.givenName)) {
            name = primary.givenName;
          } else if (primary && primary.displayName && !/[@0-9_]/.test(primary.displayName)) {
            name = primary.displayName.split(/\s+/)[0];
          }
        }
      } catch (e) {}
    }
    return { name, email, photo };
  } catch (e) {
    return { name: '', email: '', photo: null };
  }
}

function queryTeamBrief(payload) {
  var apiKey = requireApiKey_();
  var query = String((payload && payload.query) || '').trim();
  var teamJson = String((payload && payload.team) || '[]');
  if (!query) return { answer: '', actions: [] };

  var systemPrompt = [
    'You are a concise team analytics assistant for a manager. Answer questions about team workload based on data provided.',
    'Always respond with valid JSON only — no markdown fences, no extra text.',
    'Format: { "answer": "prose answer (max 120 words, use **bold** for names/numbers)", "actions": [] }',
    'If the answer recommends messaging or reassigning specific people, include action objects in the actions array.',
    'Each action: { "type": "slack_draft", "label": "Draft message to [Name]", "toName": "Full Name", "toEmail": "email@example.com", "draft": "the pre-written Slack message as manager to that person" }',
    'Only include actions when you have a specific person to message with a clear purpose. Max 3 actions.',
    'Draft messages should be professional, brief (2-3 sentences), and written in first person from the manager.',
    'If no email is available in the team data, omit that action.',
  ].join('\n');

  var text = callClaude_(
    apiKey,
    systemPrompt,
    'Team data:\n' + teamJson + '\n\nManager question: ' + query,
    500,
    MODEL
  );

  try {
    // Strip markdown fences if Claude wrapped the JSON
    var clean = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();
    var parsed = JSON.parse(clean);
    return {
      answer: parsed.answer || '',
      actions: Array.isArray(parsed.actions) ? parsed.actions : [],
    };
  } catch(e) {
    return { answer: text.trim(), actions: [] };
  }
}

function queryInboxEmails(payload) {
  var apiKey = requireApiKey_();
  var query = String((payload && payload.query) || '').trim();
  var emailsMeta = parseJson_(String((payload && payload.emails) || '[]'), []);
  if (!query) return { answer: '' };

  // Fetch full body for up to 15 emails (quota-safe)
  var enriched = emailsMeta.slice(0, 15).map(function(e) {
    var body = '';
    try {
      var msg = GmailApp.getMessageById(e.id);
      body = cleanEmailBody_(msg.getPlainBody() || '').substring(0, 800);
    } catch(err) {}
    return {
      from: e.from || '',
      senderEmail: e.senderEmail || '',
      subject: e.subject || '',
      category: e.category || '',
      date: e.date || '',
      body: body || e.preview || '',
    };
  });

  var systemPrompt = [
    'You are a concise inbox assistant with access to the full content of the user\'s emails.',
    'Answer the user\'s question using the email bodies provided — quote or summarise specific content when asked.',
    'Always respond with valid JSON only — no markdown fences, no extra text.',
    'Format: { "answer": "prose answer (max 200 words, use **bold** for names/subjects/key facts)" }',
    'Be specific and accurate. If the exact information is in the email body, share it directly.',
    'If no email matches, say so clearly.',
  ].join('\n');

  var text = callClaude_(
    apiKey,
    systemPrompt,
    'Inbox emails with full content:\n' + JSON.stringify(enriched) + '\n\nUser question: ' + query,
    1000,
    MODEL
  );

  try {
    var clean = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();
    var parsed = JSON.parse(clean);
    return { answer: parsed.answer || '' };
  } catch(e) {
    return { answer: text.trim() };
  }
}

function summarizeThread(messageId) {
  var apiKey = requireApiKey_();
  var msg = GmailApp.getMessageById(messageId);
  var messages = msg.getThread().getMessages().slice(-8);
  var threadText = messages.map(function(m, i) {
    return '[' + (i + 1) + '] From: ' + parseSender_(m.getFrom()).name + '\n' +
           cleanEmailBody_(m.getPlainBody()).substring(0, 500);
  }).join('\n\n---\n\n');
  var text = callClaude_(
    apiKey,
    'You are a concise email summarizer. Return only valid JSON, no markdown.',
    'Summarize this email thread.\nReturn JSON: {"tldr":"one sentence ≤20 words","bullets":["...","...","..."]}\n\nThread:\n' + threadText,
    400
  );
  var m = text.match(/\{[\s\S]*\}/);
  return m ? parseJson_(m[0], { tldr: '', bullets: [] }) : { tldr: 'Could not summarize.', bullets: [] };
}


function draftTeamMemberMessage(payload) {
  var apiKey = requireApiKey_();
  var name = (payload.name || '').split(' ')[0] || 'hey';
  var context = payload.context || '';
  var text = callClaude_(
    apiKey,
    'You write casual, direct manager check-in Slack messages. No emojis. No exclamation marks. No filler phrases like "hope you\'re doing well". Output only the message text.',
    'Draft a casual Slack check-in from a manager to ' + name + '. The tone should be relaxed and collegial — like a quick message between teammates. Ask how things are going on their current work, check if there are any blockers or anything you can help unblock. Keep it to 2-3 short sentences.\n\nWorkload context:\n' + context,
    160,
    MODEL
  );
  return { draft: text.trim() };
}

function draftEscalationSlack(messageId) {
  var apiKey = requireApiKey_();
  var msg = GmailApp.getMessageById(messageId);
  var sender = parseSender_(msg.getFrom());
  var subject = msg.getThread().getFirstMessageSubject();
  var body = cleanEmailBody_(msg.getPlainBody()).substring(0, 600);
  var text = callClaude_(
    apiKey,
    'You write concise Slack escalation messages. Output only the message, no quotes or preamble.',
    'Draft a Slack message escalating this email to the team. State what the issue is, who it\'s from, and what action is needed. Max 3 sentences.\n\nFrom: ' + sender.name + '\nSubject: ' + subject + '\nBody: ' + body,
    200,
    MODEL
  );
  return { draft: text.trim() };
}


function fetchInboxSnapshot(limit) {
  requireApiKey_();

  const count = (limit && limit > 0) ? Math.min(limit, 100) : 15;
  const dismissedIds = new Set(getDismissedIds());
  const emails = getLatestInboxEmails_(count)
    .filter(function(email) { return !dismissedIds.has(email.id); })
    .slice(0, count)
    .map(function(email) {
      return Object.assign({}, email, {
        category: 'digest',
        priority: 1,
        summary: 'Waiting for AI classification...',
        photo: null,
      });
    });

  return {
    emails,
    smartLabelsEnabled: getSmartLabelsEnabled(),
  };
}

// ---------------------------------------------------------------------------
// Auto-bundle: entity detection (company / project / client)
// ---------------------------------------------------------------------------

function detectEmailEntities(emails) {
  var apiKey = requireApiKey_();

  var input = (emails || []).map(function(e) {
    return {
      id: e.id,
      sender_name: e.sender_name || '',
      sender_email: e.sender_email || '',
      subject: e.subject || '',
      summary: e.summary || '',
    };
  });

  if (!input.length) return [];

  var prompt =
    'Emails:\n' + JSON.stringify(input) +
    '\n\nFor each email identify the primary entity: company name, client, or project keyword.' +
    '\nRules:\n' +
    '- Use the sender company from their email domain (e.g. john@questrade.com → "Questrade").\n' +
    '- If domain is personal (gmail, hotmail, etc.) look for a project/client keyword in the subject.\n' +
    '- Return null when no clear entity can be determined.\n' +
    '- Be consistent: same company always gets the same name. Capitalize properly.\n' +
    'Return ONLY JSON: [{"id":"...","entity":"Name or null"},...]';

  var text = callClaude_(
    apiKey,
    'You are an email entity extractor. Identify the company, client, or project for each email. Return ONLY valid JSON.',
    prompt,
    500,
    MODEL
  );

  var match = text.match(/\[[\s\S]*\]/);
  return match ? parseJson_(match[0], []) : [];
}

// ---------------------------------------------------------------------------
// Natural language email search
// ---------------------------------------------------------------------------

function saveEmailSnapshotForSearch(emails) {
  var compact = (emails || []).map(function(e) {
    return {
      id: e.id,
      from: e.sender_name || e.sender_email || '',
      email: e.sender_email || '',
      subject: e.subject || '',
      summary: e.summary || '',
      category: e.category || '',
      date: e.date || '',
    };
  });
  CacheService.getUserCache().put(
    'EMAIL_SEARCH_SNAPSHOT',
    JSON.stringify(compact.slice(0, 50)),
    21600
  );
  return compact.length;
}

function searchEmailsNL(query) {
  requireApiKey_();
  if (!query || !String(query).trim()) return { results: [] };

  var raw = CacheService.getUserCache().get('EMAIL_SEARCH_SNAPSHOT');
  if (!raw) return { results: [], stale: true };

  var emails = parseJson_(raw, []);
  if (!emails.length) return { results: [] };

  var apiKey = getApiKey_();
  var userPrompt =
    'Emails:\n' + JSON.stringify(emails) +
    '\n\nQuery: "' + String(query).replace(/"/g, "'") + '"' +
    '\n\nReturn ONLY a JSON array: [{"id":"...","reason":"one sentence explaining why it matches"},...]. Return [] if no matches.';

  var text = callClaude_(
    apiKey,
    'You are a semantic email search assistant. Find emails matching the user query by meaning and intent, not just keywords. Return ONLY valid JSON.',
    userPrompt,
    600,
    MODEL
  );

  var match = text.match(/\[[\s\S]*\]/);
  var matches = match ? parseJson_(match[0], []) : [];

  var emailMap = {};
  emails.forEach(function(e) { emailMap[e.id] = e; });

  return {
    results: matches
      .filter(function(m) { return m && m.id && emailMap[m.id]; })
      .map(function(m) { return Object.assign({}, emailMap[m.id], { reason: m.reason || '' }); })
  };
}

function classifyInboxEmails(emails) {
  requireApiKey_();

  const input = (emails || []).filter(function(email) { return !email.isFollowup; }).map(function(email) {
    return {
      id: email.id,
      subject: email.subject,
      sender_name: email.sender_name,
      sender_email: email.sender_email,
      date: email.date,
    };
  });

  const classified = applyRules(classifyWithClaude(input));
  const enriched = enrichAttentionSignals_(attachFollowupCards_(classified));

  // Fetch contact photos for all senders
  const uniqueEmails = Array.from(new Set(enriched.map(function(e) { return e.sender_email; }).filter(Boolean)));
  const photoMap = uniqueEmails.length ? fetchContactPhotos(uniqueEmails) : {};
  const withPhotos = enriched.map(function(e) {
    return Object.assign({}, e, { photo: photoMap[e.sender_email] || null });
  });

  return {
    emails: withPhotos,
    smartLabelsEnabled: getSmartLabelsEnabled(),
  };
}


function attachFollowupCards_(emails) {
  let followups;
  try {
    followups = processFollowups();
  } catch (e) {
    console.warn('Follow-up processing failed: ' + e.message);
    followups = { overdue: [], activeIds: [] };
  }

  const activeSet = new Set(followups.activeIds);
  const overdueSet = new Set(followups.overdue.map(function(f) { return f.id; }));
  const overdueCards = followups.overdue.map(function(f) {
    return {
      id: f.id,
      subject: f.subject,
      sender_name: f.sender_name,
      sender_email: f.sender_email,
      date: f.date,
      category: 'escalation',
      priority: 5,
      summary: 'No reply after ' + f.hoursOverdue + 'h - follow-up overdue',
      isFollowup: true,
      attentionSignals: {
        criticalEscalation: true,
        directAsk: false,
        deadlineDriven: true,
        vip: false,
        overdueFollowup: true,
      },
      attentionReasons: ['Overdue follow-up', 'Deadline-driven', 'Critical escalation'],
      isAttention: true,
      photo: null,
    };
  });

  const marked = (emails || [])
    .filter(function(email) { return !overdueSet.has(email.id); })
    .map(function(email) {
      return activeSet.has(email.id) ? Object.assign({}, email, { hasFollowup: true }) : email;
    });

  return overdueCards.concat(marked);
}

// ---------------------------------------------------------------------------
// Reply drafting
// ---------------------------------------------------------------------------

function fetchAndDraftReply(messageId) {
  const apiKey = requireApiKey_();
  const myEmail = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const msg = GmailApp.getMessageById(messageId);
  const thread = msg.getThread();
  const subject = thread.getFirstMessageSubject();
  const messages = thread.getMessages().slice(-4);

  // Build thread context, labelling the logged-in user's messages as "(you)"
  const context = messages.map(function(message) {
    const parsed = parseSender_(message.getFrom());
    const isMe = myEmail && parsed.email.toLowerCase() === myEmail;
    return {
      label: isMe ? parsed.name + ' (you)' : parsed.name,
      isMe: isMe,
      body: message.getPlainBody().substring(0, 400).replace(/\s+/g, ' ').trim(),
    };
  });

  // The reply goes back to whoever sent the most recent message that is NOT the logged-in user
  const lastOther = messages.slice().reverse().find(function(m) {
    return parseSender_(m.getFrom()).email.toLowerCase() !== myEmail;
  });
  const replyToEmail = lastOther ? parseSender_(lastOther.getFrom()).email : parseSender_(messages[messages.length - 1].getFrom()).email;

  const draft = callClaude_(
    apiKey,
    'You are an email writing assistant for ' + (myEmail || 'the user') + '. ' +
    'Write replies strictly in first person AS that user — never from any other sender\'s perspective. ' +
    'Return ONLY the reply body text — no subject line, no greeting prefix like "Dear...", no sign-off.',
    'Draft a reply to this email thread. You are writing AS the person labelled "(you)" in the thread.\n\n' +
    'Subject: ' + subject + '\n\nThread (oldest → newest):\n' +
    context.map(function(m, i) { return '[' + (i + 1) + '] From ' + m.label + ':\n' + m.body; }).join('\n\n') +
    '\n\nWrite a professional reply from ' + (myEmail || 'the user') + '\'s perspective, under 100 words.',
    400
  ).trim();

  if (!draft) throw new Error('Empty draft response.');
  return { draft: draft, senderEmail: replyToEmail, subject: /^re:/i.test(subject) ? subject : 'Re: ' + subject };
}

function refineDraft(messageId, currentDraft, instruction, tone) {
  var apiKey = requireApiKey_();
  var msg = GmailApp.getMessageById(messageId);
  var subject = msg.getThread().getFirstMessageSubject();
  var emailContext = msg.getPlainBody().substring(0, 600).replace(/\s+/g, ' ').trim();

  var system = 'You are an email writing assistant. Rewrite the draft according to the instruction. Return ONLY the revised email text — no explanation, no subject line, no markdown.';
  var user =
    'Original email subject: ' + subject + '\n' +
    'Email context: ' + emailContext + '\n' +
    'Current draft:\n' + String(currentDraft || '').substring(0, 1500) + '\n\n' +
    'Tone: ' + (tone || 'Warm') + '\n' +
    'Instruction: ' + String(instruction || '').substring(0, 300) + '\n\n' +
    'Rewrite the draft applying the instruction. Keep the same general structure unless told otherwise.';

  var revised = callClaude_(apiKey, system, user, 800);
  return { draft: revised };
}

function sendDraftReply(messageId, body) {
  const replyBody = String(body || '').trim();
  if (!replyBody) throw new Error('Reply body is empty.');
  if (replyBody.length > 20000) throw new Error('Reply is too long to send.');

  const msg = GmailApp.getMessageById(messageId);
  const thread = msg.getThread();
  thread.reply(replyBody);
  removeTriageLabels(thread);
  thread.addLabel(getOrCreateLabel(RESOLVED_LABEL_NAME));
  thread.moveToArchive();
  storeDismissed(messageId, 'resolve', msg, thread);
  removeFollowup(messageId);
  return {
    status: 'sent',
    resolution: 'resolved',
    sentAt: new Date().toISOString(),
  };
}

function createCalendarEventFromEmail(messageId) {
  var apiKey = requireApiKey_();
  var msg = GmailApp.getMessageById(messageId);
  if (!msg) throw new Error('Message not found.');

  var subject = msg.getThread().getFirstMessageSubject();
  var body = msg.getPlainBody().substring(0, 1200).replace(/\s+/g, ' ').trim();
  var emailDate = msg.getDate().toISOString();

  var userPrompt =
    'Email received on: ' + emailDate + '\n' +
    'Subject: ' + subject + '\n' +
    'Body: ' + body + '\n\n' +
    'Extract the meeting or deadline details and return ONLY this JSON:\n' +
    '{"title":"...","date":"YYYY-MM-DD","startTime":"HH:MM","endTime":"HH:MM","location":"","description":"..."}\n' +
    'Rules:\n' +
    '- Resolve relative dates (e.g. "Thursday", "next week") from the email received date above.\n' +
    '- endTime: add 1 hour to startTime if not specified.\n' +
    '- If no time is mentioned, use "09:00".\n' +
    '- location and description may be empty strings.\n' +
    '- Return ONLY the JSON object, nothing else.';

  var text = callClaude_(apiKey, 'You are a calendar event extractor. Return ONLY valid JSON.', userPrompt, 300);
  var match = text.match(/\{[\s\S]*?\}/);
  if (!match) throw new Error('Could not extract event details from this email.');

  var details = parseJson_(match[0], null);
  if (!details || !details.title || !details.date) throw new Error('Could not determine event date or title.');

  var startStr = details.date + 'T' + (details.startTime || '09:00') + ':00';
  var endStr   = details.date + 'T' + (details.endTime   || '10:00') + ':00';
  var start = new Date(startStr);
  var end   = new Date(endStr);
  if (isNaN(start.getTime())) throw new Error('Invalid date parsed: ' + details.date);
  if (end <= start) end = new Date(start.getTime() + 3600000);

  var opts = { description: (details.description || '').trim() };
  if (details.location) opts.location = details.location;

  var event = CalendarApp.getDefaultCalendar().createEvent(details.title, start, end, opts);

  return {
    title: details.title,
    date: details.date,
    startTime: details.startTime || '09:00',
    endTime: details.endTime || '10:00',
    eventId: event.getId(),
  };
}

function getCalendarEventDetails(messageId) {
  var msg = GmailApp.getMessageById(messageId);
  if (!msg) throw new Error('Message not found.');

  var attachments = msg.getAttachments();
  var icsContent = null;
  for (var i = 0; i < attachments.length; i++) {
    var att = attachments[i];
    if (att.getName().slice(-4).toLowerCase() === '.ics' || att.getContentType() === 'text/calendar') {
      icsContent = att.getDataAsString();
      break;
    }
  }

  var result = { title: msg.getSubject(), date: null, time: null, location: null, meetLink: null, organizer: null, organizerEmail: null };

  if (icsContent) {
    // Extract only the VEVENT block to avoid matching DTSTART inside VTIMEZONE definitions
    var veventMatch = icsContent.match(/BEGIN:VEVENT([\s\S]*?)END:VEVENT/);
    var ev = veventMatch ? veventMatch[1] : icsContent;

    var dtstartM = ev.match(/DTSTART(?:;TZID=([^;:\r\n]+))?(?:;[^:\r\n]*)?:([^\r\n]+)/);
    var dtendM   = ev.match(/DTEND(?:;TZID=([^;:\r\n]+))?(?:;[^:\r\n]*)?:([^\r\n]+)/);
    var summary   = ev.match(/SUMMARY:([^\r\n]+)/);
    var locLine   = ev.match(/LOCATION:([^\r\n]+)/);
    var organizer = ev.match(/ORGANIZER(?:;CN=([^;:\r\n]+))?(?:;[^:\r\n]*)?:mailto:([^\r\n\s]+)/i);
    var descLine  = ev.match(/DESCRIPTION:([^\r\n]+)/);

    if (summary) result.title = summary[1].replace(/\\n/g, ' ').replace(/\\,/g, ',').trim();

    var userTz = Session.getScriptTimeZone();

    function parseDt(tzid, rawVal) {
      var v = rawVal.trim();
      var isUtc = /Z$/i.test(v);
      var clean = v.replace(/Z$/i, '');
      if (clean.length === 8) {
        // All-day date
        return { date: Utilities.parseDate(clean, tzid || userTz, 'yyyyMMdd'), allDay: true };
      }
      var dt;
      if (isUtc) {
        var iso = clean.slice(0,4)+'-'+clean.slice(4,6)+'-'+clean.slice(6,8)+'T'+clean.slice(9,11)+':'+clean.slice(11,13)+':'+clean.slice(13,15)+'Z';
        dt = new Date(iso);
      } else {
        dt = Utilities.parseDate(clean, tzid || userTz, 'yyyyMMdd\'T\'HHmmss');
      }
      return { date: dt, allDay: false };
    }

    if (dtstartM) {
      var ps = parseDt(dtstartM[1], dtstartM[2]);
      result.date = Utilities.formatDate(ps.date, userTz, 'EEEE, MMMM d, yyyy');
      if (!ps.allDay) {
        result.time = Utilities.formatDate(ps.date, userTz, 'h:mm a');
        if (dtendM) {
          var pe = parseDt(dtendM[1], dtendM[2]);
          if (!pe.allDay) result.time += ' – ' + Utilities.formatDate(pe.date, userTz, 'h:mm a');
        }
      }
    }

    // Extract Meet link from location or description first
    var searchText = (locLine ? locLine[1] : '') + ' ' + (descLine ? descLine[1] : '');
    var meetMatch = searchText.match(/https:\/\/meet\.google\.com\/[a-zA-Z0-9-]+/);
    if (meetMatch) result.meetLink = meetMatch[0];

    // Only set location if it's not just the Meet URL
    if (locLine) {
      var loc = locLine[1].replace(/\\n/g, ' ').replace(/\\,/g, ',').trim();
      if (loc && loc.indexOf('meet.google.com') === -1) result.location = loc;
    }

    if (organizer) {
      result.organizer = organizer[1] ? organizer[1].replace(/^"|"$/g,'').trim() : null;
      result.organizerEmail = organizer[2] ? organizer[2].trim() : null;
      if (!result.organizer && result.organizerEmail) result.organizer = result.organizerEmail;
    }
  }

  return result;
}

function rsvpCalendarEvent(messageId, response) {
  if (response !== 'yes' && response !== 'no' && response !== 'maybe') throw new Error('Invalid response.');

  var msg = GmailApp.getMessageById(messageId);
  if (!msg) throw new Error('Message not found.');

  // Parse UID, title and start time from ICS data.
  // Google Calendar invites often embed text/calendar inline in the MIME body
  // rather than as a named attachment, so we check both.
  var uid = null, startDate = null, eventTitle = null;

  function extractFromIcs(ics) {
    // Unfold RFC 5545 continuation lines
    var unfolded = ics.replace(/\r?\n[ \t]/g, '');
    var veventBlock = unfolded.match(/BEGIN:VEVENT([\s\S]*?)END:VEVENT/);
    var evBlock = veventBlock ? veventBlock[1] : unfolded;

    var uidMatch = evBlock.match(/^UID:([^\r\n]+)/m);
    if (uidMatch && !uid) uid = uidMatch[1].trim();

    var summaryMatch = evBlock.match(/^SUMMARY:([^\r\n]+)/m);
    if (summaryMatch && !eventTitle) eventTitle = summaryMatch[1].replace(/\\n/g, ' ').replace(/\\,/g, ',').trim();

    var dtstartMatch = evBlock.match(/DTSTART(?:;TZID=([^;:\r\n]+))?(?:;[^:\r\n]*)?:([^\r\n]+)/);
    if (dtstartMatch && !startDate) {
      var tzid = dtstartMatch[1] ? dtstartMatch[1].trim() : Session.getScriptTimeZone();
      var rawDt = dtstartMatch[2].trim();
      var isUtc = /Z$/i.test(rawDt);
      var clean = rawDt.replace(/Z$/i, '');
      if (clean.length > 8) {
        try {
          startDate = isUtc
            ? new Date(clean.slice(0,4)+'-'+clean.slice(4,6)+'-'+clean.slice(6,8)+'T'+clean.slice(9,11)+':'+clean.slice(11,13)+':'+clean.slice(13,15)+'Z')
            : Utilities.parseDate(clean, tzid, 'yyyyMMdd\'T\'HHmmss');
        } catch(e) {}
      }
    }
  }

  // 1. Named attachments (.ics or text/calendar content type)
  var attachments = msg.getAttachments();
  for (var i = 0; i < attachments.length; i++) {
    var att = attachments[i];
    if (att.getName().slice(-4).toLowerCase() !== '.ics' && att.getContentType().indexOf('calendar') === -1) continue;
    try { extractFromIcs(att.getDataAsString()); } catch(e) {}
    if (uid) break;
  }

  // 2. Inline text/calendar part embedded in the raw MIME body
  if (!uid) {
    try {
      var raw = msg.getRawContent();
      var calBlocks = raw.match(/BEGIN:VCALENDAR[\s\S]*?END:VCALENDAR/g) || [];
      for (var b = 0; b < calBlocks.length; b++) {
        extractFromIcs(calBlocks[b]);
        if (uid) break;
      }
    } catch(e) {}
  }

  // 3. Last resort: search calendar by subject if we have no UID but have a title
  if (!uid && !eventTitle) eventTitle = msg.getSubject();

  if (!uid && !startDate && !eventTitle) throw new Error('No calendar event data found in this message.');

  var userEmail = Session.getActiveUser().getEmail();
  var responseStatus = response === 'yes' ? 'accepted' : response === 'maybe' ? 'tentative' : 'declined';
  var guestStatus = response === 'yes' ? CalendarApp.GuestStatus.YES
                  : response === 'maybe' ? CalendarApp.GuestStatus.MAYBE
                  : CalendarApp.GuestStatus.NO;
  var token   = ScriptApp.getOAuthToken();
  var headers = { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' };
  var rsvped  = false;

  // Strategy 1: Calendar REST API — exact UID match, most reliable
  if (uid) {
    var calListRes = UrlFetchApp.fetch(
      'https://www.googleapis.com/calendar/v3/users/me/calendarList?maxResults=50',
      { headers: headers, muteHttpExceptions: true }
    );
    var calListCode = calListRes.getResponseCode();
    if (calListCode === 401 || calListCode === 403) throw new Error('NEEDS_AUTH');

    var calIds = (JSON.parse(calListRes.getContentText()).items || []).map(function(c) { return c.id; });
    var foundCalId = null, foundEvent = null;
    for (var c = 0; c < calIds.length; c++) {
      var searchRes = UrlFetchApp.fetch(
        'https://www.googleapis.com/calendar/v3/calendars/' + encodeURIComponent(calIds[c]) +
        '/events?iCalUID=' + encodeURIComponent(uid) + '&maxResults=1',
        { headers: headers, muteHttpExceptions: true }
      );
      if (searchRes.getResponseCode() !== 200) continue;
      var items = JSON.parse(searchRes.getContentText()).items || [];
      if (items.length) { foundCalId = calIds[c]; foundEvent = items[0]; break; }
    }

    if (foundEvent) {
      if (foundEvent.organizer && foundEvent.organizer.self) {
        rsvped = true;
      } else {
        var attendees = (foundEvent.attendees || []).map(function(a) {
          if (a.self || a.email === userEmail) return Object.assign({}, a, { responseStatus: responseStatus });
          return a;
        });
        if (!attendees.some(function(a) { return a.self || a.email === userEmail; })) {
          attendees.push({ email: userEmail, responseStatus: responseStatus });
        }
        var patchRes = UrlFetchApp.fetch(
          'https://www.googleapis.com/calendar/v3/calendars/' + encodeURIComponent(foundCalId) +
          '/events/' + foundEvent.id + '?sendUpdates=all',
          { method: 'patch', headers: headers, payload: JSON.stringify({ attendees: attendees }), muteHttpExceptions: true }
        );
        rsvped = patchRes.getResponseCode() < 300;
      }
    }
  }

  // Strategy 2: CalendarApp fallback — time-window search (works when UID lookup misses)
  if (!rsvped && startDate) {
    var windowStart = new Date(startDate.getTime() - 60000);
    var windowEnd   = new Date(startDate.getTime() + 3600000);
    var allCals = CalendarApp.getAllCalendars();
    outer: for (var j = 0; j < allCals.length; j++) {
      var evts = allCals[j].getEvents(windowStart, windowEnd);
      for (var k = 0; k < evts.length; k++) {
        if (eventTitle && evts[k].getTitle() !== eventTitle) continue;
        var myStatus = evts[k].getMyStatus();
        if (myStatus === CalendarApp.GuestStatus.OWNER) { rsvped = true; break outer; }
        try { evts[k].setMyStatus(guestStatus); rsvped = true; break outer; } catch(e) {
          console.log('CalendarApp setMyStatus failed: ' + e.message);
        }
      }
    }
  }

  // Strategy 3: Title-only search (no time info available) — last resort
  if (!rsvped && eventTitle && !startDate) {
    var now = new Date();
    var searchStart = new Date(now.getTime() - 7 * 24 * 3600000);
    var searchEnd   = new Date(now.getTime() + 90 * 24 * 3600000);
    var allCals2 = CalendarApp.getAllCalendars();
    outer2: for (var j2 = 0; j2 < allCals2.length; j2++) {
      var evts2 = allCals2[j2].getEvents(searchStart, searchEnd);
      for (var k2 = 0; k2 < evts2.length; k2++) {
        if (evts2[k2].getTitle() !== eventTitle) continue;
        var myStatus2 = evts2[k2].getMyStatus();
        if (myStatus2 === CalendarApp.GuestStatus.OWNER) { rsvped = true; break outer2; }
        try { evts2[k2].setMyStatus(guestStatus); rsvped = true; break outer2; } catch(e2) {}
      }
    }
  }

  if (!rsvped) throw new Error('EVENT_NOT_FOUND');

  var thread = msg.getThread();
  removeTriageLabels(thread);
  thread.addLabel(getOrCreateLabel(RESOLVED_LABEL_NAME));
  thread.moveToArchive();
  storeDismissed(messageId, 'resolve', msg, thread);
  removeFollowup(messageId);

  return { status: 'sent', response: response };
}

// ---------------------------------------------------------------------------
// Daily digest
// ---------------------------------------------------------------------------

function sendDailyDigest() {
  const key = getApiKey_();
  if (!key) return;

  try {
    const classified = classifyWithClaude(getLatestInboxEmails_(20));
    const userEmail = Session.getActiveUser().getEmail();
    const dateStr = new Date().toLocaleDateString('en-CA', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    });
    MailApp.sendEmail({
      to: userEmail,
      subject: 'Inbox Digest - ' + dateStr,
      htmlBody: buildDigestHtml(classified, dateStr),
    });
  } catch (e) {
    console.error('Digest failed: ' + e.message);
  }
}

function buildDigestHtml(emails, dateStr) {
  const counts = emails.reduce(function(acc, email) {
    acc[email.category] = (acc[email.category] || 0) + 1;
    return acc;
  }, {});
  const urgent = emails
    .filter(function(e) { return e.category === 'escalation' || e.category === 'action_required'; })
    .sort(function(a, b) { return b.priority - a.priority; });
  const calendar = emails.filter(function(e) { return e.category === 'calendar'; });
  const fyi = emails.filter(function(e) { return e.category === 'awaiting'; }).slice(0, 5);

  const row = function(e) {
    return '<tr><td style="padding:10px 14px;border-bottom:1px solid #2a2a2a">' +
      '<div style="display:flex;align-items:baseline;gap:8px;flex-wrap:wrap">' +
      '<span style="font-size:12px;color:#888;min-width:100px;flex-shrink:0">' + escapeHtml_(e.sender_name) + '</span>' +
      '<span style="font-size:13px;font-weight:600;color:#fff">' + escapeHtml_(e.subject) + '</span>' +
      '<span style="font-size:10px;padding:2px 7px;border-radius:4px;background:' + CAT_COLORS[e.category] + '22;color:' + CAT_COLORS[e.category] + ';font-weight:700;text-transform:uppercase;letter-spacing:.4px;white-space:nowrap">' +
      e.category.replace(/_/g, ' ') + '</span></div>' +
      '<div style="font-size:12px;color:#666;margin-top:3px">' + escapeHtml_(e.summary || '') + '</div></td></tr>';
  };

  const section = function(title, items) {
    if (!items.length) return '';
    return '<div style="margin-bottom:24px"><div style="font-size:11px;font-weight:700;color:#555;text-transform:uppercase;letter-spacing:.8px;margin-bottom:10px">' +
      title + ' (' + items.length + ')</div>' +
      '<table style="width:100%;border-collapse:collapse;background:#1a1a1a;border:1px solid #2a2a2a;border-radius:8px;overflow:hidden">' +
      items.map(row).join('') + '</table></div>';
  };

  return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#0f0f0f;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif">' +
    '<div style="max-width:620px;margin:0 auto;padding:32px 20px">' +
    '<div style="font-size:18px;font-weight:700;color:#fff">Morning Inbox Digest</div>' +
    '<div style="font-size:13px;color:#888;margin-top:4px">' + escapeHtml_(dateStr) + '</div>' +
    '<div style="display:flex;gap:8px;margin:24px 0">' +
    Object.keys(CAT_COLORS).filter(function(k) { return k !== 'digest'; }).map(function(k) {
      return '<div style="flex:1;background:#1a1a1a;border:1px solid #2a2a2a;border-top:2px solid ' + CAT_COLORS[k] + ';border-radius:8px;padding:12px;text-align:center">' +
        '<div style="font-size:22px;font-weight:700;color:' + ((counts[k] || 0) > 0 ? CAT_COLORS[k] : '#444') + '">' + (counts[k] || 0) + '</div>' +
        '<div style="font-size:10px;color:#666;text-transform:uppercase;letter-spacing:.5px;margin-top:3px">' + k.replace(/_/g, ' ') + '</div></div>';
    }).join('') +
    '</div>' + section('Needs Attention', urgent) + section('Calendar', calendar) + section('Awaiting', fyi) +
    '<div style="text-align:center;margin:28px 0"><a href="' + WEB_APP_URL + '" style="display:inline-block;background:#00c04b;color:#111;font-weight:700;font-size:14px;padding:12px 28px;border-radius:8px;text-decoration:none">Open Full Inbox Triage</a></div>' +
    '</div></body></html>';
}

function setupDailyDigest() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'sendDailyDigest'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendDailyDigest')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone(Session.getScriptTimeZone())
    .create();
  return 'active';
}

function removeDailyDigest() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'sendDailyDigest'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  return 'inactive';
}

function getDailyDigestStatus() {
  return ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'sendDailyDigest';
  }) ? 'active' : 'inactive';
}

// ---------------------------------------------------------------------------
// Gmail Add-on
// ---------------------------------------------------------------------------

function onHomepage() {
  return getApiKey_() ? buildHomeCard() : buildSetupCard();
}

function getFormInput_(e, fieldName) {
  if (e && e.formInput && e.formInput[fieldName] !== undefined) return e.formInput[fieldName];
  const inputs = e && e.commonEventObject && e.commonEventObject.formInputs;
  const field = inputs && inputs[fieldName];
  const values = field && (field.stringInputs && field.stringInputs.value);
  return values && values.length ? values[0] : '';
}

function buildSetupCard() {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Inbox Triage')
      .setSubtitle('One-time setup')
      .setImageUrl('https://www.gstatic.com/images/branding/product/2x/gmail_2020q4_48dp.png')
      .setImageStyle(CardService.ImageStyle.CIRCLE))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText('Enter your Anthropic API key to enable AI triage.'))
      .addWidget(CardService.newTextInput().setFieldName('api_key').setTitle('Anthropic API Key').setHint('sk-ant-...'))
      .addWidget(CardService.newTextButton()
        .setText('Save & Continue')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('saveApiKey'))))
    .build();
}

function saveApiKey(e) {
  const key = String(getFormInput_(e, 'api_key') || '').trim();
  if (!key.startsWith('sk-ant-')) return notify('Invalid key - must start with sk-ant-');

  PropertiesService.getScriptProperties().setProperty(API_KEY_PROP, key);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildHomeCard()))
    .setNotification(CardService.newNotification().setText('API key saved'))
    .build();
}

function buildHomeCard() {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Inbox Triage')
      .setSubtitle('AI-powered email prioritization')
      .setImageUrl('https://forex-brokers.ca/wp-content/uploads/2020/04/questrade-broker-forex-canada.png')
      .setImageStyle(CardService.ImageStyle.CIRCLE))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newDecoratedText()
        .setText('QuestFlow surfaces what needs your attention — escalations, action items, and follow-ups — so you can move fast and stay in control.')
        .setWrapText(true))
      .addWidget(CardService.newDivider())
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('Quick Triage')
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction().setFunctionName('runTriage')))
        .addButton(CardService.newTextButton()
          .setText('Open Full View')
          .setOpenLink(CardService.newOpenLink().setUrl(WEB_APP_URL).setOpenAs(CardService.OpenAs.FULL_SIZE))))
      .addWidget(CardService.newTextButton()
        .setText('Change API Key')
        .setOnClickAction(CardService.newAction().setFunctionName('showSetup'))))
    .build();
}

function showSetup() {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildSetupCard()))
    .build();
}

function runTriage() {
  try {
    const classified = classifyWithClaude(getLatestInboxEmails_(20));
    CacheService.getUserCache().put('last_triage', JSON.stringify(classified), 600);
    return buildTriageCard(classified, null);
  } catch (e) {
    return buildErrorCard(e.message);
  }
}

// ---------------------------------------------------------------------------
// Classification cache
// ---------------------------------------------------------------------------

function getClassificationCache_() {
  return parseJson_(getUserProps_().getProperty(CLASSIFICATION_CACHE_PROP), {});
}

function saveClassificationCache_(cache) {
  const cutoff = Date.now() - CLASSIFICATION_CACHE_TTL_MS;
  const pruned = {};
  Object.keys(cache).forEach(function(id) {
    if (cache[id].cachedAt && new Date(cache[id].cachedAt).getTime() > cutoff) {
      pruned[id] = cache[id];
    }
  });
  const keys = Object.keys(pruned);
  if (keys.length > 500) {
    keys.sort(function(a, b) { return new Date(pruned[a].cachedAt) - new Date(pruned[b].cachedAt); })
      .slice(0, keys.length - 500)
      .forEach(function(k) { delete pruned[k]; });
  }
  getUserProps_().setProperty(CLASSIFICATION_CACHE_PROP, JSON.stringify(pruned));
}

function removeFromClassificationCache_(messageId) {
  const cache = getClassificationCache_();
  if (cache[messageId]) {
    delete cache[messageId];
    getUserProps_().setProperty(CLASSIFICATION_CACHE_PROP, JSON.stringify(cache));
  }
}

function classifyWithClaude(emails) {
  if (!emails || !emails.length) return [];

  const apiKey = requireApiKey_();
  const cache = getClassificationCache_();
  const cutoffMs = Date.now() - CLASSIFICATION_CACHE_TTL_MS;
  const now = new Date().toISOString();

  const needsClassification = [];
  emails.forEach(function(email) {
    const hit = cache[email.id];
    if (!hit || new Date(hit.cachedAt).getTime() <= cutoffMs) {
      needsClassification.push(email);
    }
  });

  if (needsClassification.length > 0) {
    const payload = needsClassification.map(function(email) {
      var snippet = '';
      try {
        var msg = GmailApp.getMessageById(email.id);
        snippet = cleanEmailBody_(msg.getPlainBody() || msg.getBody() || '').substring(0, 150);
      } catch(e) {}
      return {
        id: email.id,
        subject: email.subject,
        sender_name: email.sender_name,
        date: email.date,
        body_snippet: snippet,
      };
    });

    const SYSTEM_PROMPT =
      'You are a JSON-only API. Respond with a valid JSON array only. No markdown, no prose, no code fences. Start with [ and end with ].';

    // Chunk into batches of 20 so output never truncates
    var CLASSIFICATION_INSTRUCTIONS =
      'Classify each email into exactly one category using this strict decision hierarchy (evaluate top-to-bottom; use the first that fits):\n\n' +
      '1. escalation — C-level/VP/director is escalating an issue; production outage or system-down incident; SLA breach; named-account customer escalation; or any explicit use of words like "escalate", "critical", "sev0", "sev1", "outage", "blocker". Priority 4-5.\n' +
      '2. action_required — Sender explicitly asks the recipient to: reply with a decision, approve or reject something, complete a specific task, submit or send something, fix a bug, or take a named action. Keywords: "can you", "please", "need you to", "action required", "approval needed", "please approve", "sign off", "review and respond". Priority 2-4.\n' +
      '3. calendar — Email is a meeting invite, scheduling request, calendar event notification, or a direct request to book or reschedule time. Priority 2-3.\n' +
      '4. awaiting — A reply or response is expected but no explicit action is demanded: ongoing conversation threads, status updates you are part of, emails where someone said they will get back to you, or informational emails from known contacts where a reply is natural but not urgent. Priority 1-2.\n' +
      '5. digest — Everything else: newsletters, marketing emails, automated system notifications, receipts, social alerts, bulk mail, or any email clearly not requiring a reply. Priority 1.\n\n' +
      'Return ONLY a JSON array. Each object: id (exact match), category, priority (1-5 integer), summary (max 12 words, factual), attentionSignals ({directAsk:boolean, deadlineDriven:boolean, criticalEscalation:boolean}).\n' +
      'directAsk = category is action_required or escalation.\n' +
      'deadlineDriven = email mentions today, tomorrow, EOD, ASAP, deadline, due date, overdue, or time-sensitive urgency.\n' +
      'criticalEscalation = category is escalation.\n\n';

    var classified = [];
    var BATCH = 20;
    for (var b = 0; b < payload.length; b += BATCH) {
      var chunk = payload.slice(b, b + BATCH);
      var chunkText = callClaude_(apiKey, SYSTEM_PROMPT, CLASSIFICATION_INSTRUCTIONS + JSON.stringify(chunk), 4000, MODEL);
      var chunkResult = extractJsonArray_(chunkText);
      classified = classified.concat(chunkResult);
    }

    classified.forEach(function(found) {
      if (!found.id) return;
      cache[found.id] = {
        category: normalizeCategory_(found.category || 'digest'),
        priority: normalizePriority_(found.priority || 1),
        summary: String(found.summary || 'No summary.').substring(0, 160),
        attentionSignals: {
          directAsk: found.attentionSignals && found.attentionSignals.directAsk === true,
          deadlineDriven: found.attentionSignals && found.attentionSignals.deadlineDriven === true,
          criticalEscalation: found.attentionSignals && found.attentionSignals.criticalEscalation === true,
        },
        cachedAt: now,
      };
    });

    saveClassificationCache_(cache);
  }

  return emails.map(function(email) {
    const found = cache[email.id] || {};
    return Object.assign({}, email, {
      category: normalizeCategory_(found.category || 'digest'),
      priority: normalizePriority_(found.priority || 1),
      summary: String(found.summary || 'No summary.').substring(0, 160),
      attentionSignals: {
        directAsk: found.attentionSignals && found.attentionSignals.directAsk === true,
        deadlineDriven: found.attentionSignals && found.attentionSignals.deadlineDriven === true,
        criticalEscalation: found.attentionSignals && found.attentionSignals.criticalEscalation === true,
      },
    });
  });
}

function buildTriageCard(emails, filterCat) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildTriageCard_(emails, filterCat)))
    .build();
}

function buildTriageCard_(emails, filterCat) {
  const counts = emails.reduce(function(acc, email) {
    acc[email.category] = (acc[email.category] || 0) + 1;
    return acc;
  }, {});
  const needAttn = (counts.action_required || 0) + (counts.escalation || 0);
  const active = filterCat || 'all';
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Inbox Triage')
      .setSubtitle(emails.length + ' emails - ' + needAttn + ' need attention')
      .setImageUrl('https://www.gstatic.com/images/branding/product/2x/gmail_2020q4_48dp.png')
      .setImageStyle(CardService.ImageStyle.CIRCLE));

  const grid = CardService.newGrid()
    .setNumColumns(3)
    .setOnClickAction(CardService.newAction().setFunctionName('handleGridClick'));

  Object.keys(CAT_META).forEach(function(key) {
    const meta = CAT_META[key];
    grid.addItem(CardService.newGridItem()
      .setIdentifier(key)
      .setTitle(String(counts[key] || 0))
      .setSubtitle(meta.label)
      .setTextAlignment(CardService.HorizontalAlignment.CENTER)
      .setLayout(CardService.GridItemLayout.TEXT_BELOW));
  });

  card.addSection(CardService.newCardSection()
    .setHeader('Tap a category to filter')
    .addWidget(grid)
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Refresh')
        .setOnClickAction(CardService.newAction().setFunctionName('runTriage')))
      .addButton(CardService.newTextButton()
        .setText('Show All')
        .setOnClickAction(CardService.newAction().setFunctionName('showAllCached')))
      .addButton(CardService.newTextButton()
        .setText('Full View')
        .setOpenLink(CardService.newOpenLink().setUrl(WEB_APP_URL).setOpenAs(CardService.OpenAs.FULL_SIZE)))));

  const sorted = emails.slice()
    .filter(function(email) { return active === 'all' || email.category === active; })
    .sort(function(a, b) {
      const ao = Object.prototype.hasOwnProperty.call(CAT_ORDER, a.category) ? CAT_ORDER[a.category] : 5;
      const bo = Object.prototype.hasOwnProperty.call(CAT_ORDER, b.category) ? CAT_ORDER[b.category] : 5;
      const cd = ao - bo;
      return cd !== 0 ? cd : b.priority - a.priority;
    });

  if (!sorted.length) {
    card.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText('No emails in this category.')));
    return card.build();
  }

  const groups = {};
  sorted.forEach(function(email) {
    if (!groups[email.category]) groups[email.category] = [];
    groups[email.category].push(email);
  });

  Object.keys(CAT_ORDER).forEach(function(cat) {
    const group = groups[cat];
    if (!group || !group.length) return;
    const meta = CAT_META[cat];
    const isUrgent = cat === 'escalation' || cat === 'action_required';
    const section = CardService.newCardSection()
      .setHeader(meta.emoji + ' ' + meta.label + ' (' + group.length + ')')
      .setCollapsible(!isUrgent)
      .setNumUncollapsibleWidgets(isUrgent ? group.length : 0);

    group.forEach(function(email) {
      section.addWidget(CardService.newDecoratedText()
        .setTopLabel(email.sender_name + ' - ' + (PRIORITY_LABELS[email.priority] || ''))
        .setText(email.subject)
        .setBottomLabel(email.summary)
        .setIcon(meta.icon)
        .setWrapText(true)
        .setOnClickAction(CardService.newAction().setFunctionName('openEmail').setParameters({ messageId: email.id })));
      section.addWidget(CardService.newDivider());
    });
    card.addSection(section);
  });

  return card.build();
}

function handleGridClick(e) {
  const cache = CacheService.getUserCache().get('last_triage');
  if (!cache) return notify('Session expired - please re-triage.');
  const cat = getGridIdentifier_(e);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildTriageCard_(JSON.parse(cache), cat)))
    .build();
}

function getGridIdentifier_(e) {
  if (e && e.formInput && e.formInput.grid_click_identifier) return e.formInput.grid_click_identifier;
  if (e && e.commonEventObject && e.commonEventObject.parameters) {
    return e.commonEventObject.parameters.grid_click_identifier || e.commonEventObject.parameters.identifier;
  }
  return null;
}

function showAllCached() {
  const cache = CacheService.getUserCache().get('last_triage');
  if (!cache) return notify('Session expired - please re-triage.');
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildTriageCard_(JSON.parse(cache), null)))
    .build();
}

function openEmail(e) {
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink().setUrl('https://mail.google.com/mail/u/0/#inbox/' + e.parameters.messageId))
    .build();
}

function buildErrorCard(msg) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(
      CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle('Something went wrong'))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newDecoratedText().setText('⚠️ ' + msg).setWrapText(true))
          .addWidget(CardService.newTextButton()
            .setText('Back')
            .setOnClickAction(CardService.newAction().setFunctionName('showHome'))))
        .build()))
    .build();
}

function showHome() {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(onHomepage()))
    .build();
}

function notify(msg) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(msg))
    .build();
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ---------------------------------------------------------------------------
// ClickUp integration
// ---------------------------------------------------------------------------

var CLICKUP_DEFAULT_LISTS = [
  { name: 'Legal Operations', id: '' },
  { name: 'Litigation',       id: '' },
  { name: 'Intake',           id: '' },
  { name: 'Project Roadmap',  id: '' },
  { name: 'Projects',         id: '' },
  { name: 'Automation',       id: '' },
];

function getClickUpSettings() {
  var props = PropertiesService.getScriptProperties();
  var listsJson = props.getProperty(CLICKUP_LISTS_PROP);
  var lists;
  if (listsJson) {
    lists = parseJson_(listsJson, CLICKUP_DEFAULT_LISTS);
  } else {
    // First run — seed defaults, migrate legacy single listId if present
    lists = CLICKUP_DEFAULT_LISTS.map(function(l) { return { name: l.name, id: l.id }; });
    var legacyId = props.getProperty(CLICKUP_LIST_ID_PROP) || '';
    if (legacyId) lists[0] = { name: lists[0].name, id: legacyId };
  }
  return {
    apiKey:               props.getProperty(CLICKUP_API_KEY_PROP) || '',
    lists:                lists,
    stuckDays:            parseInt(props.getProperty(CLICKUP_STUCK_DAYS_PROP) || '3', 10),
    slackToken:           props.getProperty(SLACK_BOT_TOKEN_PROP) || '',
    slackOverrideChannel: props.getProperty(SLACK_OVERRIDE_CHANNEL_PROP) || '',
  };
}

function saveSlackToken(token, overrideChannel) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(SLACK_BOT_TOKEN_PROP, String(token || '').trim());
  props.setProperty(SLACK_OVERRIDE_CHANNEL_PROP, String(overrideChannel || '').trim());
  return true;
}

function sendSlackCheckin(toEmail, message) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty(SLACK_BOT_TOKEN_PROP);
  if (!token) throw new Error('Slack bot token not configured. Add it in Team Brief settings.');

  var channel = props.getProperty(SLACK_OVERRIDE_CHANNEL_PROP) || '';

  if (!channel) {
    if (!toEmail) throw new Error('No recipient email provided.');
    var lookupRes = UrlFetchApp.fetch(
      'https://slack.com/api/users.lookupByEmail?email=' + encodeURIComponent(toEmail),
      { method: 'get', muteHttpExceptions: true, headers: { 'Authorization': 'Bearer ' + token } }
    );
    var lookupData = parseJson_(lookupRes.getContentText(), null);
    if (!lookupData || !lookupData.ok) {
      throw new Error('Could not find Slack user for ' + toEmail + '. ' + (lookupData && lookupData.error ? lookupData.error : ''));
    }
    channel = lookupData.user.id;
  }

  var sendRes = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    muteHttpExceptions: true,
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json; charset=utf-8' },
    payload: JSON.stringify({ channel: channel, text: message }),
  });
  var sendData = parseJson_(sendRes.getContentText(), null);
  if (!sendData || !sendData.ok) {
    throw new Error('Slack send failed: ' + (sendData && sendData.error ? sendData.error : 'unknown error'));
  }
  return { ok: true };
}



function testAnthropicConnection() {
  var apiKey = getApiKey_();
  if (!apiKey) throw new Error('No Anthropic API key configured.');
  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', muteHttpExceptions: true,
    headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: MODEL, max_tokens: 1, messages: [{ role: 'user', content: '.' }] }),
  });
  var data = parseJson_(res.getContentText(), null);
  if (res.getResponseCode() === 200 && data && data.content) return { ok: true };
  var msg = data && data.error ? data.error.message : ('HTTP ' + res.getResponseCode());
  throw new Error(msg);
}

function reassignClickUpTask(payload) {
  var taskId  = payload.taskId;
  var toEmail = payload.toEmail;
  var toName  = payload.toName;
  if (!taskId) throw new Error('No task ID provided.');
  var settings = getClickUpSettings();
  if (!settings.apiKey) throw new Error('ClickUp not configured.');
  var headers = { 'Authorization': settings.apiKey, 'Content-Type': 'application/json' };

  // Resolve email → ClickUp user ID via workspace members
  var teamRes = UrlFetchApp.fetch('https://api.clickup.com/api/v2/team', { method:'get', muteHttpExceptions:true, headers:headers });
  var teamData = parseJson_(teamRes.getContentText(), null);
  var userId = null;
  if (teamData && teamData.teams) {
    teamData.teams.forEach(function(team) {
      (team.members || []).forEach(function(member) {
        if (member.user && member.user.email === toEmail) userId = member.user.id;
      });
    });
  }
  if (!userId) throw new Error('Could not find ClickUp user for ' + toEmail + '. Make sure they are a workspace member.');

  // Remove existing assignees and set new one
  var body = JSON.stringify({ assignees: { add: [userId], rem: [] } });
  var res = UrlFetchApp.fetch('https://api.clickup.com/api/v2/task/' + taskId, {
    method: 'put', muteHttpExceptions: true, headers: headers, payload: body
  });
  var result = parseJson_(res.getContentText(), null);
  if (!result || result.err) throw new Error(result ? result.err : 'ClickUp API error');
  return { success: true, taskId: taskId, assignedTo: toName };
}

function testClickUpConnection() {
  var settings = getClickUpSettings();
  if (!settings.apiKey) throw new Error('No ClickUp API key configured.');
  return testClickUpKey(settings.apiKey);
}

function testClickUpKey(apiKey) {
  if (!apiKey || !String(apiKey).trim()) throw new Error('No ClickUp API key provided.');
  var res = UrlFetchApp.fetch('https://api.clickup.com/api/v2/team', {
    method: 'get', muteHttpExceptions: true,
    headers: { 'Authorization': String(apiKey).trim(), 'Content-Type': 'application/json' },
  });
  var data = parseJson_(res.getContentText(), null);
  if (data && data.teams) return { ok: true, workspaces: data.teams.length };
  throw new Error(data && data.err ? data.err : ('HTTP ' + res.getResponseCode()));
}

function testSlackConnection() {
  var token = PropertiesService.getScriptProperties().getProperty(SLACK_BOT_TOKEN_PROP);
  if (!token) throw new Error('No Slack bot token configured.');
  return testSlackToken(token);
}

function testSlackToken(token) {
  if (!token || !String(token).trim()) throw new Error('No Slack bot token provided.');
  var res = UrlFetchApp.fetch('https://slack.com/api/auth.test', {
    method: 'post', muteHttpExceptions: true,
    headers: { 'Authorization': 'Bearer ' + String(token).trim(), 'Content-Type': 'application/json' },
    payload: '{}',
  });
  var data = parseJson_(res.getContentText(), null);
  if (data && data.ok) return { ok: true, team: data.team };
  throw new Error(data && data.error ? data.error : ('HTTP ' + res.getResponseCode()));
}

function saveClickUpSettings(apiKey, lists) {
  var props = PropertiesService.getScriptProperties();
  var cleanKey = String(apiKey || '').trim();
  if (!cleanKey) throw new Error('API key is required.');
  if (!cleanKey.startsWith('pk_')) throw new Error('API key must start with pk_');
  if (cleanKey.length < 30) throw new Error('API key looks too short — copy it again from ClickUp.');
  props.setProperty(CLICKUP_API_KEY_PROP, cleanKey);
  if (lists && Array.isArray(lists)) {
    var cleaned = lists.map(function(l) {
      var id = String(l.id || '').replace(/\D/g, '').trim();
      if (id && !/^\d+$/.test(id)) throw new Error('List ID for "' + l.name + '" must be numeric only.');
      var sublists = Array.isArray(l.sublists) ? l.sublists.map(function(sl) {
        var slId = String(sl.id || '').replace(/\D/g, '').trim();
        if (slId && !/^\d+$/.test(slId)) throw new Error('Sublist ID for "' + sl.name + '" must be numeric only.');
        return { name: String(sl.name || '').trim(), id: slId };
      }) : [];
      return { name: String(l.name || '').trim(), id: id, sublists: sublists };
    });
    props.setProperty(CLICKUP_LISTS_PROP, JSON.stringify(cleaned));
    // Keep legacy prop in sync for backward compat
    var first = cleaned.find(function(l) { return l.id; });
    if (first) props.setProperty(CLICKUP_LIST_ID_PROP, first.id);
  }
  return true;
}

function clickUpFetch_(path, apiKey) {
  var res = UrlFetchApp.fetch('https://api.clickup.com/api/v2' + path, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'Authorization': apiKey,
      'Content-Type': 'application/json',
    },
  });
  var data = parseJson_(res.getContentText(), null);
  if (!data) throw new Error('ClickUp API error: HTTP ' + res.getResponseCode());
  if (data.err) throw new Error('ClickUp: ' + data.err);
  return data;
}

// Returns the most recent Gmail thread whose subject or participants overlap with
// a task name or assignee email. Used to surface email context alongside stuck tasks.
function findRelatedThread_(taskName, assigneeEmail) {
  try {
    var cleanName = (taskName || '').replace(/[^a-zA-Z0-9 ]/g, '').trim();
    var threads = [];
    if (cleanName.length > 4) {
      threads = GmailApp.search('"' + cleanName.substring(0, 40) + '"', 0, 1);
    }
    if (!threads.length && assigneeEmail) {
      threads = GmailApp.search('from:' + assigneeEmail + ' newer_than:14d', 0, 1);
    }
    if (!threads.length) return null;
    var thread = threads[0];
    var msg = thread.getMessages()[thread.getMessageCount() - 1];
    return {
      subject: thread.getFirstMessageSubject(),
      messageId: msg.getId(),
      date: msg.getDate().toISOString(),
    };
  } catch (e) {
    return null;
  }
}

function getTeamBrief(listId) {
  var settings = getClickUpSettings();
  if (!settings.apiKey) {
    throw new Error('ClickUp not configured. Add your API key in settings.');
  }

  // Try cache first (5-minute TTL)
  var cacheKey = 'TEAM_BRIEF_' + (listId || 'default');
  try {
    var cached = CacheService.getUserCache().get(cacheKey);
    if (cached) return parseJson_(cached, null);
  } catch(e) {}

  var now = Date.now();
  var stuckDays = settings.stuckDays || 3;
  var staleMs = stuckDays * 24 * 3600 * 1000;
  var IN_PROGRESS_KEYWORDS = ['in progress', 'in_progress', 'doing', 'active', 'started'];

  var params = '?include_closed=false&subtasks=true&page=0';
  var headers = { 'Authorization': settings.apiKey, 'Content-Type': 'application/json' };

  var targetId = String(listId || '').replace(/\D/g, '').trim();
  var configuredLists = (settings.lists || []).filter(function(l) { return l.id; });

  // Build endpoint list — workspace-wide only as last resort if no lists configured
  var endpoints = [];
  if (targetId) endpoints.push('/list/' + targetId + '/task' + params);
  configuredLists.forEach(function(l) {
    if (l.id !== targetId) endpoints.push('/list/' + l.id + '/task' + params);
  });

  // Only fall back to workspace-wide if no configured lists at all
  if (endpoints.length === 0) {
    var teamLookup = UrlFetchApp.fetch('https://api.clickup.com/api/v2/team', {
      method: 'get', muteHttpExceptions: true, headers: headers,
    });
    var teamInfo = parseJson_(teamLookup.getContentText(), null);
    if (!teamInfo || !teamInfo.teams || !teamInfo.teams.length) {
      throw new Error('No authorized workspaces found. Verify your API key.');
    }
    // Limit to first team only to avoid timeout
    endpoints.push('/team/' + teamInfo.teams[0].id + '/task' + params);
  }

  // Fetch from ALL endpoints and merge by task ID so tasks across multiple lists are included
  var seenTaskIds = {};
  var allTasks = [];
  var lastErr = '';
  var anySuccess = false;
  for (var ep = 0; ep < endpoints.length; ep++) {
    var res = UrlFetchApp.fetch('https://api.clickup.com/api/v2' + endpoints[ep], {
      method: 'get', muteHttpExceptions: true, headers: headers,
    });
    var parsed = parseJson_(res.getContentText(), null);
    if (parsed && parsed.tasks) {
      anySuccess = true;
      parsed.tasks.forEach(function(t) {
        if (!seenTaskIds[t.id]) { seenTaskIds[t.id] = true; allTasks.push(t); }
      });
    }
    if (parsed && parsed.err) lastErr = parsed.err;
  }
  if (!anySuccess) {
    throw new Error(lastErr || 'Could not fetch tasks. Verify your API key and List IDs in Settings.');
  }

  var memberMap = {};
  allTasks.forEach(function(task) {
    var dueDate = task.due_date ? parseInt(task.due_date) : null;
    var dateUpdated = task.date_updated ? parseInt(task.date_updated) : (task.date_created ? parseInt(task.date_created) : null);
    var isOverdue = !!(dueDate && dueDate > 0 && dueDate < now);
    var daysSinceUpdate = dateUpdated ? Math.floor((now - dateUpdated) / 86400000) : null;
    var isStale = !!(dateUpdated && (now - dateUpdated) > staleMs);
    var rawStatus = task.status ? (task.status.status || task.status.type || 'unknown') : 'unknown';
    var statusType = (task.status && task.status.type) ? task.status.type.toLowerCase() : '';
    var status = rawStatus.toLowerCase();
    var isClosed = statusType === 'closed' || status === 'complete' || status === 'completed' || status === 'done' || status === 'closed';
    var isInProgress = IN_PROGRESS_KEYWORDS.some(function(kw) { return status.indexOf(kw) !== -1; });
    var isStuck = !isClosed && (isOverdue || isStale || (isInProgress && daysSinceUpdate !== null && daysSinceUpdate >= stuckDays));

    var p = task.priority ? String(task.priority.priority || '').toLowerCase() : '';
    var priorityBucket = (p === 'urgent' || p === 'high') ? 'high' : (p === 'low') ? 'low' : 'medium';

    (task.assignees || []).forEach(function(assignee) {
      var key = String(assignee.id);
      if (!memberMap[key]) {
        memberMap[key] = {
          id: assignee.id,
          name: assignee.username || assignee.email || 'Unknown',
          email: assignee.email || '',
          allTasks: [],
          stuckTasks: [],
          priorityCounts: { high: 0, medium: 0, low: 0 },
        };
      }
      if (!isClosed) memberMap[key].allTasks.push(task);
      memberMap[key].priorityCounts[priorityBucket]++;
      if (isStuck) {
        memberMap[key].stuckTasks.push({
          id: task.id,
          name: task.name,
          status: rawStatus,
          dueDate: dueDate ? new Date(dueDate).toISOString().slice(0, 10) : null,
          daysSinceUpdate: daysSinceUpdate,
          isOverdue: isOverdue,
          isInProgress: isInProgress,
          isStale: isStale && !isOverdue,
        });
      }
    });
  });

  var members = Object.values(memberMap).filter(function(m) {
    return m.allTasks.length > 0;
  });

  if (!members.length) return { members: [], allClear: true };

  // Find linked Gmail threads for the top stuck task per member
  members.forEach(function(m) {
    m.stuckTasks = m.stuckTasks.slice(0, 5).map(function(t) {
      t.linkedThread = findRelatedThread_(t.name, m.email);
      return t;
    });
  });

  var apiKey = requireApiKey_();
  var memberSummaries = members.map(function(m) {
    return {
      name: m.name,
      email: m.email,
      totalOpen: m.allTasks.length,
      stuckCount: m.stuckTasks.length,
      stuckTasks: m.stuckTasks.map(function(t) {
        return {
          name: t.name,
          status: t.status,
          dueDate: t.dueDate,
          daysSinceUpdate: t.daysSinceUpdate,
          overdue: t.isOverdue,
          inProgress: t.isInProgress,
        };
      }),
    };
  });

  var text = callClaude_(
    apiKey,
    'You are a team manager assistant. Return only valid JSON. No markdown, no prose.',
    'For each team member, analyze their stuck tasks and generate:\n' +
    '- workloadScore: integer 1-5 (1=light, 5=overwhelmed based on open task count and stuck tasks)\n' +
    '- overloaded: boolean (true if score >= 4)\n' +
    '- suggestion: 1 sentence — the single most important unblocking action for the manager to take\n' +
    '- checkInMessage: a SHORT summary blurb (max 15 words). Do NOT include the person\'s name. Example: "4 stuck tasks — API docs 6 days no update, status: in progress"\n' +
    '- checkInDraft: a short, casual Slack DM (2-3 sentences). Friendly tone, no greeting/sign-off, no formal language. Mention the specific task name and ask if there are blockers or if they need help.\n\n' +
    'Return ONLY a JSON array in the same order as input: [{workloadScore,overloaded,suggestion,checkInMessage,checkInDraft},...]\n\n' +
    JSON.stringify(memberSummaries),
    1500
  );

  var match = text.match(/\[[\s\S]*\]/);
  var analysisArr = parseJson_(match ? match[0] : '[]', []);

  var result = members.map(function(m, i) {
    var ai = analysisArr[i] || {};
    return {
      name: m.name,
      email: m.email,
      totalOpen: m.allTasks.length,
      stuckCount: m.stuckTasks.length,
      workloadScore: Math.min(5, Math.max(1, parseInt(ai.workloadScore) || 3)),
      overloaded: !!ai.overloaded,
      suggestion: String(ai.suggestion || '').slice(0, 200),
      checkInMessage: String(ai.checkInMessage || '').slice(0, 120),
      checkInDraft: String(ai.checkInDraft || '').slice(0, 600),
      stuckTasks: m.stuckTasks,
      priorityCounts: m.priorityCounts,
    };
  });

  var finalResult = { members: result };

  // Cache for 5 minutes
  try {
    var json = JSON.stringify(finalResult);
    if (json.length < 100000) CacheService.getUserCache().put(cacheKey, json, 300);
  } catch(e) {}

  return finalResult;
}

// ---------------------------------------------------------------------------
// People Management Alerts
// ---------------------------------------------------------------------------

// ── QFG-inspired multi-factor workload scoring ───────────────────────────────
// Adapted from QFG Vendor Risk Framework v1.4
//   Event factors  (Impact × Likelihood, 1-16) → normalized 0-100: (raw-1)×100/15
//   Exposure factors (Impact only, 1-4)         → normalized 0-100: (raw-1)×100/3
//   Weights: deadline 35%, complexity 25%, stakeholder 25%, regulatory 15%
//   Tier thresholds: ≥75 Critical · 50-74 High · 25-49 Medium · <25 Low
//   Floor override: compliance or client-facing tasks at max → tier ≥ Medium
// ─────────────────────────────────────────────────────────────────────────────
function computeQFWorkloadScore_(tasks) {
  if (!tasks || !tasks.length) return { score: 0, tier: 'Low', factors: {} };

  var WEIGHTS = { deadline: 0.35, complexity: 0.25, stakeholder: 0.25, regulatory: 0.15 };
  var PRIORITY_IMPACT = { urgent: 4, high: 3, medium: 2, low: 1 };
  var now = Date.now();

  var totDeadline = 0, totComplexity = 0, totStakeholder = 0, totRegulatory = 0;
  var floorOverride = false;

  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    var name = String(t.name || t.description || '').toLowerCase();
    var priority = String(t.priority || 'low').toLowerCase();
    var impact = PRIORITY_IMPACT[priority] || 1;

    // 1. Deadline pressure — Event (Impact × Likelihood, 1-16)
    var isOverdue = t.isOverdue || false;
    var daysSince = t.daysSinceUpdate || (t.date_updated ? (now - Number(t.date_updated)) / 86400000 : 0);
    var isStuck = !t.isDone && daysSince >= 5;
    var likelihood = isOverdue ? 4 : isStuck ? 3 : daysSince > 7 ? 2 : 1;
    var deadlineRaw = Math.min(impact * likelihood, 16);
    totDeadline += (deadlineRaw - 1) * 100 / 15;

    // 2. Task complexity — Exposure (1-4): priority + age penalty
    var complexityRaw = Math.min(4, impact + (daysSince > 14 ? 1 : 0));
    totComplexity += (complexityRaw - 1) * 100 / 3;

    // 3. Stakeholder/customer impact — Exposure (1-4)
    var stakeRaw = 1;
    if (/\b(client|customer|external|partner|exec|ceo|board|vip|delivery|launch|demo|presentation)\b/.test(name)) stakeRaw = 3;
    if (/\b(board|ceo|investor|regulator|auditor|critical client)\b/.test(name)) stakeRaw = 4;
    totStakeholder += (stakeRaw - 1) * 100 / 3;
    if (stakeRaw >= 4) floorOverride = true;

    // 4. Regulatory/compliance flag — Exposure (1-4)
    var regRaw = 1;
    if (/\b(compliance|audit|legal|review|approval|sign.?off|contract)\b/.test(name)) regRaw = 2;
    if (/\b(regulatory|soc|gdpr|pii|sox|sec|kyc|aml|privacy|risk|policy)\b/.test(name)) regRaw = 3;
    if (/\b(breach|violation|penalty|enforcement|mandatory|critical audit)\b/.test(name)) regRaw = 4;
    totRegulatory += (regRaw - 1) * 100 / 3;
    if (regRaw >= 4) floorOverride = true;
  }

  var n = tasks.length;
  var avgDeadline    = totDeadline    / n;
  var avgComplexity  = totComplexity  / n;
  var avgStakeholder = totStakeholder / n;
  var avgRegulatory  = totRegulatory  / n;

  // Volume multiplier: each additional task adds 8% load, capped at 2×
  var volumeMult = Math.min(1 + (n - 1) * 0.08, 2.0);

  var raw = (
    avgDeadline    * WEIGHTS.deadline +
    avgComplexity  * WEIGHTS.complexity +
    avgStakeholder * WEIGHTS.stakeholder +
    avgRegulatory  * WEIGHTS.regulatory
  ) * volumeMult;

  var score = Math.round(Math.min(raw, 100) * 10) / 10;
  if (floorOverride && score < 25) score = 25;

  var tier = score >= 75 ? 'Critical' : score >= 50 ? 'High' : score >= 25 ? 'Medium' : 'Low';

  return {
    score: score,
    tier: tier,
    factors: {
      deadline:    Math.round(avgDeadline),
      complexity:  Math.round(avgComplexity),
      stakeholder: Math.round(avgStakeholder),
      regulatory:  Math.round(avgRegulatory),
    },
  };
}

// Fallback when only priorityCounts are available (no task detail array)
function computeFallbackScore_(member) {
  var pc = member.priorityCounts || {};
  var stuckCount = (member.stuckTasks || []).length;
  var n = Math.max(member.totalOpen || 1, 1);
  // Approximate normalized scores from priority buckets
  var urgentDeadline = (pc.urgent || 0) * 100;   // urgent overdue → max event
  var highDeadline   = (pc.high   || 0) * 46.7;  // high active → mid event
  var medDeadline    = (pc.medium || 0) * 13.3;
  var stuckDeadline  = stuckCount      * 73.3;   // stuck → likelihood=3, impact=4 → (12-1)×100/15
  var avgDeadline    = Math.min((urgentDeadline + highDeadline + medDeadline + stuckDeadline) / n, 100);
  var avgComplexity  = Math.min(((pc.urgent||0)*100 + (pc.high||0)*66.7 + (pc.medium||0)*33.3) / n, 100);
  var volumeMult     = Math.min(1 + (n - 1) * 0.08, 2.0);
  var raw = (avgDeadline * 0.35 + avgComplexity * 0.25) * volumeMult;
  var score = Math.round(Math.min(raw, 100) * 10) / 10;
  if ((pc.urgent || 0) > 0 && score < 50) score = 50;
  var tier = score >= 75 ? 'Critical' : score >= 50 ? 'High' : score >= 25 ? 'Medium' : 'Low';
  return { score: score, tier: tier, factors: {} };
}

function runPeopleManagementAlerts() {
  var brief = getTeamBrief();
  var members = brief.members || [];
  if (!members.length) return { sent: false, reason: 'No team members found' };

  var scoreObjects = members.map(function(m) {
    return m.tasks && m.tasks.length ? computeQFWorkloadScore_(m.tasks) : computeFallbackScore_(m);
  });
  var scores = scoreObjects.map(function(s) { return s.score; });

  var teamAvg = scores.length ? scores.reduce(function(a, b) { return a + b; }, 0) / scores.length : 0;

  // Section A: stuck tasks / escalations
  var sectionALines = [];
  members.forEach(function(m, i) {
    (m.stuckTasks || []).forEach(function(t) {
      var days = t.daysSinceUpdate || 0;
      var label = days > 5 ? '*ESCALATION*' : 'Stuck';
      sectionALines.push(label + ' — *' + m.name + '*: ' + t.name + ' (' + days + 'd idle)');
    });
  });

  // Section B: check-in recommendations
  var sectionBLines = [];
  members.forEach(function(m, i) {
    var checkIn = scores[i] > teamAvg * 1.3 || (m.stuckTasks || []).length >= 2;
    if (checkIn) {
      var reason = scores[i] > teamAvg * 1.3 ? 'high complexity score (' + scores[i].toFixed(1) + ' vs avg ' + teamAvg.toFixed(1) + ')' : 'multiple stuck tasks (' + m.stuckTasks.length + ')';
      sectionBLines.push('• *' + m.name + '* — ' + reason);
    }
  });

  // Section C: workload imbalance
  var sectionCLines = [];
  var redistributionSuggestions = '';
  if (members.length >= 3) {
    var maxScore = Math.max.apply(null, scores);
    var minScore = Math.min.apply(null, scores);
    var allClose = maxScore <= teamAvg * 1.2 && minScore >= teamAvg * 0.8;
    if (!allClose && maxScore > teamAvg * 1.5 && (maxScore - minScore) > teamAvg * 0.8) {
      var overloaded = members.filter(function(m, i) { return scores[i] > teamAvg * 1.2; });
      var underloaded = members.filter(function(m, i) { return scores[i] < teamAvg * 0.8; });
      sectionCLines.push('Overloaded: ' + overloaded.map(function(m) { return m.name; }).join(', '));
      sectionCLines.push('Underloaded: ' + underloaded.map(function(m) { return m.name; }).join(', '));
      try {
        var apiKey = requireApiKey_();
        var summary = members.map(function(m, i) { return m.name + ': score=' + scores[i].toFixed(1) + ' (' + scoreObjects[i].tier + '), open=' + m.totalOpen + ', stuck=' + (m.stuckTasks||[]).length; }).join('\n');
        redistributionSuggestions = callClaude_(
          apiKey,
          'You are a team management assistant. Be concise.',
          'The following team has an unbalanced workload. Suggest 2-3 specific task redistribution actions (one sentence each):\n' + summary,
          300
        );
        sectionCLines.push('Suggestions:\n' + redistributionSuggestions);
      } catch(e) {
        sectionCLines.push('(Could not generate AI suggestions: ' + e.message + ')');
      }
    }
  }

  var hasAlerts = sectionALines.length > 0 || sectionBLines.length > 0 || sectionCLines.length > 0;
  if (!hasAlerts) return { sent: false, reason: 'No alerts — team looks healthy' };

  var msg = '*QuestFlow People Management Alerts*\n\n';
  if (sectionALines.length) {
    msg += '*A. Stuck / Escalation*\n' + sectionALines.join('\n') + '\n\n';
  }
  if (sectionBLines.length) {
    msg += '*B. Check-in Recommendations*\n' + sectionBLines.join('\n') + '\n\n';
  }
  if (sectionCLines.length) {
    msg += '*C. Workload Imbalance*\n' + sectionCLines.join('\n') + '\n\n';
  }
  msg += '_Generated: ' + new Date().toISOString() + '_';

  var token = PropertiesService.getScriptProperties().getProperty(SLACK_BOT_TOKEN_PROP);
  if (!token || !token.trim()) throw new Error('Slack not configured. Set SLACK_BOT_TOKEN in Script Properties.');
  token = token.trim();

  var managerEmail = Session.getActiveUser().getEmail();
  var slackUserId = lookupSlackUserByEmail(managerEmail, token);
  var channelId = openSlackDm(slackUserId, token);
  postSlackMessage(channelId, msg, token);

  var alertCount = sectionALines.length + sectionBLines.length + (sectionCLines.length ? 1 : 0);
  return { sent: true, alertCount: alertCount };
}

function getPeopleAlertsSummary() {
  var brief = getTeamBrief();
  var members = brief.members || [];

  var scoreObjects = members.map(function(m) {
    return m.tasks && m.tasks.length ? computeQFWorkloadScore_(m.tasks) : computeFallbackScore_(m);
  });
  var scores = scoreObjects.map(function(s) { return s.score; });

  var teamAvg = scores.length ? scores.reduce(function(a, b) { return a + b; }, 0) / scores.length : 0;

  var memberResults = members.map(function(m, i) {
    var so = scoreObjects[i];
    var checkIn = scores[i] >= 50 || (m.stuckTasks || []).length >= 2;
    var checkInReason = '';
    if (checkIn) {
      checkInReason = so.tier === 'Critical' || so.tier === 'High'
        ? so.tier + ' workload tier (' + scores[i].toFixed(1) + ' · avg ' + teamAvg.toFixed(1) + ')'
        : 'Multiple stuck tasks (' + m.stuckTasks.length + ')';
    }
    return {
      name: m.name,
      email: m.email,
      complexityScore: scores[i],
      tier: so.tier,
      factors: so.factors,
      checkInFlag: checkIn,
      checkInReason: checkInReason,
    };
  });

  var stuckEscalations = [];
  members.forEach(function(m) {
    (m.stuckTasks || []).forEach(function(t) {
      stuckEscalations.push({ memberName: m.name, taskName: t.name, daysStuck: t.daysSinceUpdate || 0 });
    });
  });

  var imbalanced = false;
  var overloaded = [];
  var underloaded = [];
  var redistributionSuggestions = [];

  if (members.length >= 3) {
    var maxScore = Math.max.apply(null, scores);
    var minScore = Math.min.apply(null, scores);
    var allClose = maxScore <= teamAvg * 1.2 && minScore >= teamAvg * 0.8;
    var spreadSignificant = (maxScore - minScore) > 20; // 20pt spread on 0-100 scale is meaningful
    if (!allClose && maxScore > teamAvg * 1.4 && spreadSignificant) {
      imbalanced = true;
      members.forEach(function(m, i) {
        if (scores[i] >= 50) overloaded.push({ name: m.name, score: scores[i], tier: scoreObjects[i].tier });
        if (scores[i] < 25 && teamAvg >= 25) underloaded.push({ name: m.name, score: scores[i], tier: scoreObjects[i].tier });
      });
      try {
        var apiKey = requireApiKey_();
        var summary = members.map(function(m, i) { return m.name + ': score=' + scores[i].toFixed(1) + ' (' + scoreObjects[i].tier + '), open=' + m.totalOpen + ', stuck=' + (m.stuckTasks||[]).length; }).join('\n');
        var raw = callClaude_(
          apiKey,
          'You are a team management assistant. Be concise.',
          'The following team has an unbalanced workload. Suggest 2-3 specific task redistribution actions (one sentence each, as a JSON array of strings):\n' + summary,
          300
        );
        var match = raw.match(/\[[\s\S]*\]/);
        if (match) {
          var parsed = parseJson_(match[0], []);
          redistributionSuggestions = parsed;
        } else {
          redistributionSuggestions = [raw];
        }
      } catch(e) {
        redistributionSuggestions = [];
      }
    }
  }

  return {
    members: memberResults,
    teamAvg: Math.round(teamAvg * 10) / 10,
    imbalanced: imbalanced,
    overloaded: overloaded,
    underloaded: underloaded,
    redistributionSuggestions: redistributionSuggestions,
    stuckEscalations: stuckEscalations,
    generatedAt: new Date().toISOString(),
  };
}

function setupPeopleAlertsSchedule(frequencyHours) {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'runPeopleManagementAlerts'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runPeopleManagementAlerts')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  PropertiesService.getScriptProperties().setProperty('PEOPLE_ALERTS_ENABLED', 'true');
  PropertiesService.getScriptProperties().setProperty('PEOPLE_ALERTS_FREQ', String(frequencyHours || 24));
  return { ok: true };
}

function removePeopleAlertsSchedule() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'runPeopleManagementAlerts'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  PropertiesService.getScriptProperties().setProperty('PEOPLE_ALERTS_ENABLED', 'false');
  return { ok: true };
}

function getPeopleAlertsConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    enabled: props.getProperty('PEOPLE_ALERTS_ENABLED') === 'true',
    frequencyHours: parseInt(props.getProperty('PEOPLE_ALERTS_FREQ') || '24', 10),
  };
}

// ---------------------------------------------------------------------------
// DocuSign Integration
// ---------------------------------------------------------------------------

var DS_SETTINGS_PROP = 'DOCUSIGN_SETTINGS';
var DS_NOTIFICATIONS_PROP = 'DOCUSIGN_NOTIFICATIONS';

function getDocuSignSettings() {
  return parseJson_(PropertiesService.getUserProperties().getProperty(DS_SETTINGS_PROP), null);
}

function saveDocuSignSettings(payload) {
  PropertiesService.getUserProperties().setProperty(DS_SETTINGS_PROP, JSON.stringify(payload));
  return { ok: true };
}

function getDocuSignHeaders_() {
  var s = getDocuSignSettings();
  if (!s || !s.accessToken) throw new Error('DocuSign not configured. Add credentials in Settings → DocuSign.');
  return { 'Authorization': 'Bearer ' + s.accessToken, 'Content-Type': 'application/json' };
}

function getDocuSignBase_() {
  var s = getDocuSignSettings();
  if (!s || !s.accountId) throw new Error('DocuSign not configured.');
  var base = (s.baseUrl || 'https://na4.docusign.net/restapi').replace(/\/+$/, '');
  return base + '/v2.1/accounts/' + s.accountId;
}

function testDocuSignConnection() {
  var headers = getDocuSignHeaders_();
  var base = getDocuSignBase_();
  var res = UrlFetchApp.fetch(base, { method: 'get', muteHttpExceptions: true, headers: headers });
  var data = parseJson_(res.getContentText(), null);
  if (res.getResponseCode() >= 400) throw new Error(data && data.message ? data.message : 'HTTP ' + res.getResponseCode());
  return { name: (data && data.currentUserInfo && data.currentUserInfo.userName) || 'Connected', accountName: data && data.accountName || '' };
}

function getDraftEnvelopes() {
  var headers = getDocuSignHeaders_();
  var base = getDocuSignBase_();
  // Fetch drafts (created status)
  var url = base + '/envelopes?from_date=2020-01-01&status=created&order_by=last_modified&order=desc&count=50';
  var res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: headers });
  if (res.getResponseCode() >= 400) throw new Error('DocuSign API error: ' + res.getResponseCode());
  var data = parseJson_(res.getContentText(), null);
  var envelopes = (data && data.envelopes) || [];
  return envelopes.map(function(e) {
    return {
      envelopeId: e.envelopeId,
      subject: e.emailSubject || '(no subject)',
      status: e.status || 'created',
      lastModified: e.lastModifiedDateTime || e.createdDateTime || null,
      docCount: parseInt(e.documentsCount || '0', 10),
      signers: [],
    };
  });
}

function getEnvelopePipelineAll() {
  var headers = getDocuSignHeaders_();
  var base = getDocuSignBase_();
  // Fetch sent, delivered, and completed envelopes from the last 90 days
  var from = new Date(Date.now() - 90 * 24 * 3600 * 1000).toISOString().slice(0, 10);
  var statuses = ['sent', 'delivered', 'completed', 'declined', 'voided'];
  var all = [];
  var seen = {};
  for (var i = 0; i < statuses.length; i++) {
    var url = base + '/envelopes?from_date=' + from + '&status=' + statuses[i] + '&order_by=last_modified&order=desc&count=20';
    var res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: headers });
    if (res.getResponseCode() >= 400) continue;
    var data = parseJson_(res.getContentText(), null);
    ((data && data.envelopes) || []).forEach(function(e) {
      if (!seen[e.envelopeId]) {
        seen[e.envelopeId] = true;
        all.push({ envelopeId: e.envelopeId, subject: e.emailSubject || '(no subject)', status: e.status, lastModified: e.lastModifiedDateTime || null });
      }
    });
  }
  return all;
}

function getDocuSignNotifications() {
  var headers = getDocuSignHeaders_();
  var base = getDocuSignBase_();
  var from = new Date(Date.now() - 30 * 24 * 3600 * 1000).toISOString().slice(0, 10);
  var url = base + '/envelopes?from_date=' + from + '&status=completed&order_by=last_modified&order=desc&count=20';
  var res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: headers });
  if (res.getResponseCode() >= 400) return [];
  var data = parseJson_(res.getContentText(), null);
  var envelopes = (data && data.envelopes) || [];
  var seenProp = parseJson_(PropertiesService.getUserProperties().getProperty(DS_NOTIFICATIONS_PROP), {});
  return envelopes.map(function(e) {
    return {
      envelopeId: e.envelopeId,
      subject: e.emailSubject || '(no subject)',
      status: e.status,
      signerName: '',
      signerEmail: '',
      timestamp: e.lastModifiedDateTime || null,
      isNew: !seenProp[e.envelopeId],
    };
  });
}

function markDocuSignNotificationsSeen() {
  var headers = getDocuSignHeaders_();
  var base = getDocuSignBase_();
  var from = new Date(Date.now() - 30 * 24 * 3600 * 1000).toISOString().slice(0, 10);
  var url = base + '/envelopes?from_date=' + from + '&status=completed&count=20';
  var res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: headers });
  var data = parseJson_(res.getContentText(), null);
  var seen = {};
  ((data && data.envelopes) || []).forEach(function(e) { seen[e.envelopeId] = true; });
  PropertiesService.getUserProperties().setProperty(DS_NOTIFICATIONS_PROP, JSON.stringify(seen));
  return { ok: true };
}

function getEnvelopeDetails(envelopeId) {
  var headers = getDocuSignHeaders_();
  var base = getDocuSignBase_();
  var res = UrlFetchApp.fetch(base + '/envelopes/' + envelopeId, { method: 'get', muteHttpExceptions: true, headers: headers });
  if (res.getResponseCode() >= 400) throw new Error('Could not load envelope details.');
  var env = parseJson_(res.getContentText(), null);
  // Fetch recipients
  var rRes = UrlFetchApp.fetch(base + '/envelopes/' + envelopeId + '/recipients', { method: 'get', muteHttpExceptions: true, headers: headers });
  var rData = parseJson_(rRes.getContentText(), null);
  var signers = ((rData && rData.signers) || []).map(function(s) {
    return { name: s.name || '', email: s.email || '', status: s.status || 'created' };
  });
  // Fetch documents
  var dRes = UrlFetchApp.fetch(base + '/envelopes/' + envelopeId + '/documents', { method: 'get', muteHttpExceptions: true, headers: headers });
  var dData = parseJson_(dRes.getContentText(), null);
  var docs = ((dData && dData.envelopeDocuments) || []).filter(function(d){ return d.documentId !== 'certificate'; }).map(function(d) {
    return { name: d.name || 'Document', documentId: d.documentId };
  });
  return {
    envelopeId: envelopeId,
    subject: env.emailSubject || '(no subject)',
    status: env.status || 'created',
    emailBlurb: env.emailBlurb || '',
    signers: signers,
    documents: docs,
    lastModified: env.lastModifiedDateTime || null,
  };
}

function generateDocuSignMessage(detail) {
  var apiKey = PropertiesService.getUserProperties().getProperty(API_KEY_PROP) || PropertiesService.getScriptProperties().getProperty(API_KEY_PROP);
  if (!apiKey) return { draft: detail.emailBlurb || '' };
  var signerNames = (detail.signers || []).map(function(s) { return s.name; }).join(', ') || 'the recipient';
  var prompt = 'Write a professional, concise cover message for a DocuSign envelope titled "' + detail.subject + '" sent to ' + signerNames + '. The message should be 2-3 sentences, friendly, and explain they need to review and sign the attached document. Return only the message body, no greeting or signature.';
  var text = callClaude_(apiKey, 'You are a professional business writing assistant. Write clear, concise DocuSign cover messages.', prompt, 200, MODEL);
  return { draft: text.trim() };
}

function prepareEnvelopeDraft(emailId) {
  var apiKey = PropertiesService.getUserProperties().getProperty('ANTHROPIC_API_KEY') || PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  var msg = GmailApp.getMessageById(emailId);
  var thread = msg.getThread();
  var messages = thread.getMessages().slice(-6);
  var subject = thread.getFirstMessageSubject() || '';
  var sender = parseSender_(msg.getFrom());
  var todayIso = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Build thread context
  var threadText = messages.map(function(m, i) {
    var s = parseSender_(m.getFrom());
    return '[Message ' + (i+1) + '] From: ' + s.name + ' <' + s.email + '>\n' + cleanEmailBody_(m.getPlainBody()).substring(0, 400);
  }).join('\n\n');

  // Get CC addresses from latest message
  var ccStr = '';
  try { ccStr = msg.getCc() || ''; } catch(e) {}

  var systemPrompt = 'You are a professional business writing assistant preparing DocuSign envelopes. Extract structured info and write a cover message. Return ONLY valid JSON, no markdown.';
  var userPrompt = 'Today: ' + todayIso + '\nSubject: ' + subject + '\nSender: ' + sender.name + ' <' + sender.email + '>\nCC: ' + (ccStr || 'none') + '\n\nThread:\n' + threadText + '\n\nReturn a JSON object with these exact keys:\n- recipientName: full name of who should sign (string)\n- recipientEmail: their email address (string)\n- docType: short document label like "NDA", "Service Agreement", "Contract", "Approval Form" (max 4 words)\n- approvalContext: if CCs or thread mention approvers/senders, a short string like "Approved by Sarah Chen (CFO)", else null\n- coverMsg: full professional cover message with: greeting by first name, 1-line description of what needs signing, approval/sender context if relevant, request to review and sign, thanks and warm sign-off. Should feel personal not robotic.\n\nReturn ONLY the JSON object.';

  if (!apiKey) {
    return {
      recipientName: sender.name,
      recipientEmail: sender.email,
      docType: 'Document',
      approvalContext: null,
      coverMsg: 'Hi ' + sender.name.split(' ')[0] + ',\n\nPlease find attached the document for your review and signature.\n\nKindly sign at your earliest convenience.\n\nThank you,\nZach'
    };
  }

  var text = callClaude_(apiKey, systemPrompt, userPrompt, 600, MODEL);
  var match = text.match(/\{[\s\S]*\}/);
  var parsed = match ? parseJson_(match[0], null) : null;

  if (!parsed || !parsed.coverMsg) {
    return {
      recipientName: sender.name,
      recipientEmail: sender.email,
      docType: 'Document',
      approvalContext: null,
      coverMsg: text.trim() || ('Hi,\n\nPlease find attached the document for your review and signature.\n\nThank you,\nZach')
    };
  }
  return {
    recipientName: parsed.recipientName || sender.name,
    recipientEmail: parsed.recipientEmail || sender.email,
    docType: parsed.docType || 'Document',
    approvalContext: parsed.approvalContext || null,
    coverMsg: parsed.coverMsg || ''
  };
}

function sendDocuSignEnvelope(envelopeId, coverMessage) {
  var headers = getDocuSignHeaders_();
  var base = getDocuSignBase_();
  // Update email blurb then send (change status to 'sent')
  var updateRes = UrlFetchApp.fetch(base + '/envelopes/' + envelopeId, {
    method: 'put',
    muteHttpExceptions: true,
    headers: headers,
    payload: JSON.stringify({ emailBlurb: coverMessage, status: 'sent' }),
  });
  if (updateRes.getResponseCode() >= 400) {
    var err = parseJson_(updateRes.getContentText(), null);
    throw new Error((err && err.message) || 'Failed to send envelope.');
  }
  return { ok: true };
}

function searchRelatedEmailForEnvelope(subject) {
  if (!subject) return null;
  var words = subject.replace(/complete with docusign[:.]?\s*/i, '').replace(/[^\w\s]/g, ' ').trim().split(/\s+/).filter(function(w){ return w.length > 3; }).slice(0, 4);
  if (!words.length) return null;
  var query = words.join(' ');
  var threads = GmailApp.search(query, 0, 3);
  if (!threads.length) return null;
  var msg = threads[0].getMessages()[0];
  return { snippet: msg.getPlainBody().slice(0, 400), subject: msg.getSubject(), from: msg.getFrom() };
}


function resendDocuSignEnvelope(envelopeId) {
  var cfg = getDocuSignSettings();
  if (!cfg || !cfg.accessToken || !cfg.accountId) throw new Error('DocuSign not configured');
  var url = cfg.baseUrl + '/v2.1/accounts/' + cfg.accountId + '/envelopes/' + envelopeId + '/recipients?resend_envelope=true';
  var res = UrlFetchApp.fetch(url, {
    method: 'PUT',
    headers: { 'Authorization': 'Bearer ' + cfg.accessToken, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ signers: [] }),
    muteHttpExceptions: true,
  });
  return { status: res.getResponseCode() };
}

function sendEnvelopeWithFiles(payload) {
  var recipientName  = String((payload && payload.recipientName)  || '').trim();
  var recipientEmail = String((payload && payload.recipientEmail) || '').trim();
  var subject        = String((payload && payload.subject)        || 'Document for your signature').trim();
  var coverMsg       = String((payload && payload.coverMsg)       || '').trim();
  var fileIds        = (payload && payload.fileIds) || [];

  if (!recipientEmail) throw new Error('Recipient email is required.');

  var headers = getDocuSignHeaders_();
  var base    = getDocuSignBase_();

  var documents = fileIds.map(function(fileId, i) {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var b64  = Utilities.base64Encode(blob.getBytes());
    var mime = file.getMimeType() || '';
    var ext  = mime === 'application/pdf' ? 'pdf' : mime.indexOf('word') !== -1 ? 'docx' : 'pdf';
    return {
      documentId: String(i + 1),
      name: file.getName(),
      fileExtension: ext,
      documentBase64: b64,
    };
  });

  var envelopeBody = {
    emailSubject: subject,
    emailBlurb: coverMsg,
    documents: documents,
    recipients: {
      signers: [{
        recipientId: '1',
        name: recipientName || recipientEmail,
        email: recipientEmail,
        tabs: { signHereTabs: [{ documentId: '1', pageNumber: '1', xPosition: '100', yPosition: '700' }] },
      }],
    },
    status: 'sent',
  };

  var res = UrlFetchApp.fetch(base + '/envelopes', {
    method: 'post', muteHttpExceptions: true, headers: headers,
    payload: JSON.stringify(envelopeBody),
  });
  if (res.getResponseCode() >= 400) {
    var err = parseJson_(res.getContentText(), null);
    throw new Error((err && err.message) || 'Failed to create envelope. HTTP ' + res.getResponseCode());
  }
  var result = parseJson_(res.getContentText(), null);
  return { ok: true, envelopeId: result && result.envelopeId };
}

// ── MEETINGS / CORPORATE SECRETARY MODULE ────────────────────────────────────

var MEETING_CONFIG_PROP = 'MEETING_CONFIG';
var MEETING_RECORDS_PROP = 'MEETING_RECORDS';
var MEETING_ACTION_ITEMS_PROP = 'MEETING_ACTION_ITEMS';
var MEETING_COMMITTEES_PROP = 'MEETING_COMMITTEES';
var MEETING_ACTIVE_PROP = 'MEETING_ACTIVE_COMMITTEE';
var MEETING_STATUS_PROP = 'MEETING_STATUS';
var MEETING_AGENDA_PROP = 'MEETING_AGENDA';
var COI_REGISTRY_PROP = 'COI_REGISTRY';

function getActiveCommitteeCode_() {
  return PropertiesService.getUserProperties().getProperty(MEETING_ACTIVE_PROP) || 'QUESTBANK';
}

function getCommitteesList_() {
  var raw = PropertiesService.getUserProperties().getProperty(MEETING_COMMITTEES_PROP);
  return raw ? JSON.parse(raw) : ['QUESTBANK'];
}

function getDrivePickerToken() {
  return JSON.stringify({ token: ScriptApp.getOAuthToken() });
}

function getCalendarRooms() {
  var rooms = [];
  // Try enterprise Google Calendar room resources first
  try {
    var token = ScriptApp.getOAuthToken();
    var res = UrlFetchApp.fetch(
      'https://www.googleapis.com/calendar/v3/users/me/calendarList?maxResults=250',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (res.getResponseCode() === 200) {
      var data = parseJson_(res.getContentText(), null);
      var items = (data && data.items) || [];
      items.forEach(function(cal) {
        var id = cal.id || '';
        if (id.indexOf('@resource.calendar.google.com') !== -1) {
          var name = (cal.summaryOverride || cal.summary || id).trim();
          if (name) rooms.push(name);
        }
      });
    }
  } catch(e) {}
  // Fallback: scan recent/upcoming event locations
  if (rooms.length === 0) {
    try {
      var seen = {};
      var now = new Date();
      var events = CalendarApp.getDefaultCalendar().getEvents(
        new Date(now.getTime() - 90 * 86400000),
        new Date(now.getTime() + 180 * 86400000)
      );
      events.forEach(function(ev) {
        var loc = (ev.getLocation() || '').trim();
        if (loc && !seen[loc]) { seen[loc] = true; rooms.push(loc); }
      });
    } catch(e) {}
    rooms = rooms.slice(0, 30);
  }
  return JSON.stringify(rooms);
}

function getMeetingConfig() {
  var code = getActiveCommitteeCode_();
  var raw = PropertiesService.getUserProperties().getProperty(MEETING_CONFIG_PROP + '_' + code)
         || PropertiesService.getUserProperties().getProperty(MEETING_CONFIG_PROP);
  if (raw) return raw;
  return JSON.stringify({
    committeeCode: code, title: code + ' Committee Meeting',
    dayOfWeek: 3, hour: 10, minute: 0, durationMinutes: 60,
    room: '', attendees: [], workingFolderId: '', publishedFolderId: '', numQuarters: 2,
  });
}

function saveMeetingConfig(configJson) {
  var cfg = JSON.parse(configJson);
  var code = cfg.committeeCode || getActiveCommitteeCode_();
  PropertiesService.getUserProperties().setProperty(MEETING_CONFIG_PROP + '_' + code, configJson);
  PropertiesService.getUserProperties().setProperty(MEETING_ACTIVE_PROP, code);
  var list = getCommitteesList_();
  if (list.indexOf(code) === -1) { list.push(code); PropertiesService.getUserProperties().setProperty(MEETING_COMMITTEES_PROP, JSON.stringify(list)); }
  return 'ok';
}

function getCommittees() {
  var list = getCommitteesList_();
  var active = getActiveCommitteeCode_();
  var committees = list.map(function(code) {
    var raw = PropertiesService.getUserProperties().getProperty(MEETING_CONFIG_PROP + '_' + code);
    var cfg = raw ? JSON.parse(raw) : { committeeCode: code, title: code + ' Committee Meeting' };
    return { code: code, title: cfg.title || code };
  });
  return JSON.stringify({ committees: committees, active: active });
}

function setActiveCommittee(code) {
  PropertiesService.getUserProperties().setProperty(MEETING_ACTIVE_PROP, code);
  return getMeetingConfig();
}

function removeCommittee(code) {
  var list = getCommitteesList_();
  var idx = list.indexOf(code);
  if (idx !== -1) list.splice(idx, 1);
  if (list.length === 0) list = ['QUESTBANK'];
  PropertiesService.getUserProperties().setProperty(MEETING_COMMITTEES_PROP, JSON.stringify(list));
  PropertiesService.getUserProperties().deleteProperty(MEETING_CONFIG_PROP + '_' + code);
  if (getActiveCommitteeCode_() === code) PropertiesService.getUserProperties().setProperty(MEETING_ACTIVE_PROP, list[0]);
  return JSON.stringify({ remaining: list.length, active: getActiveCommitteeCode_() });
}

function getFirstFullBusinessWeekMonday_(year, month) {
  var d = new Date(year, month, 1);
  var dow = d.getDay();
  var offset = dow === 1 ? 0 : (8 - dow) % 7;
  return new Date(year, month, 1 + offset);
}

function calculateMeetingSchedule(numQuartersStr) {
  var numQuarters = parseInt(numQuartersStr, 10) || 2;
  var config = JSON.parse(getMeetingConfig());
  var dayOffset = (config.dayOfWeek || 3) - 1;
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth();
  var results = [];

  for (var q = 0; q < numQuarters; q++) {
    for (var m = 0; m < 3; m++) {
      var totalMonth = month + q * 3 + m;
      var y = year + Math.floor(totalMonth / 12);
      var mo = totalMonth % 12;
      var firstMonday = getFirstFullBusinessWeekMonday_(y, mo);
      var thirdMonday = new Date(firstMonday.getTime() + 14 * 86400000);
      var firstMeeting = new Date(firstMonday.getTime() + dayOffset * 86400000);
      var thirdMeeting = new Date(thirdMonday.getTime() + dayOffset * 86400000);

      function pushDate(dt) {
        if (dt < now) return;
        var lbl = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        var display = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
        results.push({ iso: dt.toISOString(), label: lbl, display: display });
      }
      pushDate(firstMeeting);
      pushDate(thirdMeeting);
    }
  }
  return JSON.stringify(results);
}

function scheduleMeetingsOnCalendar(datesJson) {
  var dates = JSON.parse(datesJson);
  var config = JSON.parse(getMeetingConfig());
  var calendar = CalendarApp.getDefaultCalendar();
  var created = [];
  var skipped = [];

  dates.forEach(function(d) {
    var base = new Date(d.iso);
    base.setHours(config.hour || 10, config.minute || 0, 0, 0);
    var end = new Date(base.getTime() + (config.durationMinutes || 60) * 60000);
    var title = config.title || 'Committee Meeting';
    var existing = calendar.getEvents(base, end);
    var duplicate = false;
    for (var i = 0; i < existing.length; i++) {
      if (existing[i].getTitle() === title) { duplicate = true; break; }
    }
    // Always create Drive subfolders regardless of whether the calendar event is new or duplicate
    var code2 = config.committeeCode || 'CMT';
    if (config.workingFolderId) { try { getMeetingSubfolder_(config.workingFolderId, d.label, code2); } catch(e2) {} }
    if (config.publishedFolderId) { try { getMeetingSubfolder_(config.publishedFolderId, d.label, code2); } catch(e2) {} }
    if (duplicate) {
      skipped.push({ date: d.label, display: d.display });
      return;
    }
    var coOrgs = config.coOrganizerEmails || [];
    var allGuests = (config.attendees || []).concat(coOrgs).filter(function(e,i,a){return a.indexOf(e)===i;});
    var descLines = ['Meeting Suite — ' + (config.committeeCode || 'Committee')];
    if (coOrgs.length) descLines.push('Organizers: ' + coOrgs.join(', '));
    var opts = {
      description: descLines.join('\n'),
      guests: allGuests.join(','),
      sendInvites: true,
    };
    if (config.room) opts.location = config.room;
    var ev = calendar.createEvent(title, base, end, opts);
    ev.setGuestsCanInviteOthers(false);
    ev.setGuestsCanModify(false);
    created.push({ id: ev.getId(), date: d.label, display: d.display });
  });
  return JSON.stringify({ created: created.length, skipped: skipped.length, events: created, alreadyExists: skipped });
}

function createAllSubfolders(datesJson) {
  var dates = JSON.parse(datesJson);
  var config = JSON.parse(getMeetingConfig());
  var code = config.committeeCode || 'CMT';
  var created = [];
  dates.forEach(function(d) {
    var entry = { date: d.label, working: null, published: null };
    if (config.workingFolderId) {
      try { var wf = getMeetingSubfolder_(config.workingFolderId, d.label, code); entry.working = wf.getUrl(); } catch(e) {}
    }
    if (config.publishedFolderId) {
      try { var pf = getMeetingSubfolder_(config.publishedFolderId, d.label, code); entry.published = pf.getUrl(); } catch(e) {}
    }
    created.push(entry);
  });
  return JSON.stringify({ created: created.length, folders: created });
}

function requestMeetingMaterials(dateLabel, attendeesJson) {
  var config = JSON.parse(getMeetingConfig());
  var attendees = JSON.parse(attendeesJson || '[]');
  if (!attendees.length) attendees = config.attendees || [];
  if (!attendees.length) throw new Error('No attendees configured. Add attendee emails in Meeting Settings.');
  var senderName = getSenderName_();
  var tmpl = getEmailTemplate_(config, 'materials');
  var vars = { committee: config.committeeCode||'CMT', title: config.title||'Committee Meeting', date: dateLabel, senderName: senderName };
  var subject = renderTemplate_(tmpl.subject, vars);
  var body = renderTemplate_(tmpl.body, vars);
  attendees.forEach(function(email) { GmailApp.sendEmail(email, subject, body, {name: senderName}); });
  return JSON.stringify({ sent: attendees.length });
}

function announceMeeting(dateLabel, attendeesJson) {
  var config = JSON.parse(getMeetingConfig());
  var attendees = JSON.parse(attendeesJson || '[]');
  if (!attendees.length) attendees = config.attendees || [];
  if (!attendees.length) throw new Error('No attendees configured.');
  var senderName = getSenderName_();
  var timeStr = (config.hour || 10) + ':' + ('0' + (config.minute || 0)).slice(-2);
  var tmpl = getEmailTemplate_(config, 'notice');
  var vars = { committee: config.committeeCode||'CMT', title: config.title||'Committee Meeting', date: dateLabel, time: timeStr, room: config.room||'TBD', senderName: senderName };
  var subject = renderTemplate_(tmpl.subject, vars);
  var body = renderTemplate_(tmpl.body, vars);
  attendees.forEach(function(email) { GmailApp.sendEmail(email, subject, body, {name: senderName}); });
  return JSON.stringify({ sent: attendees.length });
}

function transformNotesToMinutes(roughNotes, meetingContextJson) {
  var apiKey = requireApiKey_();
  var ctx = JSON.parse(meetingContextJson || '{}');
  var config = JSON.parse(getMeetingConfig());
  var system = 'You are the Meeting Suite for Questrade Financial Group, a Canadian financial services company regulated by CIRO. ' +
    'Your audience is senior executives and board members who expect precision, brevity, and formal language. ' +
    'Transform rough notes into polished committee meeting minutes following Questrade\'s standard template. ' +
    'Rules: (1) Use formal third-person past tense throughout. (2) Every discussion item must conclude with a clear RESOLVED statement. ' +
    '(3) Do not invent, embellish, or add details not present in the notes. (4) Action items must be numbered and attributed to a named individual with a due date. ' +
    '(5) Omit all informal language, filler, and small talk.';
  var conflictsSection = '';
  if (ctx.conflicts && ctx.conflicts.length > 0) {
    conflictsSection = 'CONFLICT OF INTEREST DECLARATIONS:\n' +
      ctx.conflicts.map(function(c) { return '• ' + c.member + ' declared a conflict regarding: ' + c.item; }).join('\n') +
      '\n(Include these in section 4 "CONFLICTS OF INTEREST" before business items.)\n\n';
  }
  var defaultTemplate =
    '═══════════════════════════════════════════════════\n' +
    '[COMMITTEE NAME] — MEETING MINUTES\nCONFIDENTIAL\n' +
    '═══════════════════════════════════════════════════\n\n' +
    'Date:      [date]\nTime:      [time or "As recorded"]\nLocation:  [room or "As configured"]\nSecretary: Meeting Suite Office\n\n' +
    'ATTENDEES PRESENT\n[bullet list of names]\n\nREGRETS\n[names or "None recorded"]\n\n' +
    '───────────────────────────────────────────────────\n' +
    '1. CALL TO ORDER\nThe Chair called the meeting to order.\n\n' +
    '2. APPROVAL OF AGENDA\nRESOLVED that the agenda be approved as presented.\n\n' +
    '3. APPROVAL OF PREVIOUS MINUTES\n[Include only if mentioned in notes. Otherwise omit.]\n\n' +
    '4. BUSINESS\n[For each item:]\n  4.X [Item Title]\n  [1-3 sentence summary, third-person past tense.]\n  RESOLVED: [Clear decision.]\n\n' +
    '5. ACTION ITEMS\n| # | Action | Owner | Due Date | Status |\n|---|--------|-------|----------|--------|\n[one row per item]\n\n' +
    '6. NEXT MEETING\n[Date if mentioned.]\n\n' +
    '7. ADJOURNMENT\nThere being no further business, the meeting was adjourned.\n\n' +
    '───────────────────────────────────────────────────\n' +
    'Certified correct:\n\nMeeting Suite: _______________________  Date: ___________\nChair: ____________________________________  Date: ___________';
  var minutesTemplate = (config.minutesTemplate && config.minutesTemplate.trim()) ? config.minutesTemplate : defaultTemplate;
  var user = 'MEETING: ' + (config.title || 'Committee Meeting') + ' (' + (config.committeeCode || '') + ')\n' +
    'DATE: ' + (ctx.date || '[Date]') + '\n' +
    'ATTENDEES: ' + (ctx.attendees || (config.attendees || []).join(', ') || '[List]') + '\n\n' +
    conflictsSection +
    'ROUGH NOTES:\n' + roughNotes + '\n\n' +
    'Produce minutes using EXACTLY this structure:\n\n' + minutesTemplate;
  return callClaude_(apiKey, system, user, 3000, MODEL);
}

function extractActionItemsFromText(minutesText, meetingDate) {
  var apiKey = requireApiKey_();
  var user = 'From these meeting minutes, extract every action item. ' +
    'Return ONLY a JSON array, no other text. Each object must have: ' +
    '{"action":"...","owner":"...","ownerEmail":"","dueDate":"YYYY-MM-DD or description","status":"open"}\n\n' + minutesText;
  var raw = callClaude_(apiKey, 'Return only a valid JSON array.', user, 1000, MODEL);
  try {
    var clean = raw.replace(/```json|```/g, '').trim();
    var items = JSON.parse(clean);
    var now = new Date().toISOString();
    return JSON.stringify(items.map(function(item, i) {
      item.id = 'ai_' + Date.now() + '_' + i;
      item.meetingDate = meetingDate || '';
      item.createdAt = now;
      item.status = 'open';
      return item;
    }));
  } catch(e) {
    return '[]';
  }
}

function getMeetingActionItems() {
  var code = getActiveCommitteeCode_();
  return PropertiesService.getUserProperties().getProperty(MEETING_ACTION_ITEMS_PROP + '_' + code)
      || PropertiesService.getUserProperties().getProperty(MEETING_ACTION_ITEMS_PROP)
      || '[]';
}

function saveMeetingActionItems(itemsJson) {
  var code = getActiveCommitteeCode_();
  PropertiesService.getUserProperties().setProperty(MEETING_ACTION_ITEMS_PROP + '_' + code, itemsJson);
  return 'ok';
}

function appendMeetingActionItems(newItemsJson) {
  var existing = JSON.parse(getMeetingActionItems());
  var newItems = JSON.parse(newItemsJson);
  var merged = newItems.concat(existing);
  var code = getActiveCommitteeCode_();
  PropertiesService.getUserProperties().setProperty(MEETING_ACTION_ITEMS_PROP + '_' + code, JSON.stringify(merged));
  return JSON.stringify({ total: merged.length });
}

function getMeetingSubfolder_(parentFolderId, meetingLabel, committeeCode) {
  var folderName = (committeeCode || 'CMT') + '_' + meetingLabel;
  var parent = DriveApp.getFolderById(parentFolderId);
  var existing = parent.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(folderName);
}

function enforceNamingConvention(meetingLabel) {
  var config = JSON.parse(getMeetingConfig());
  if (!config.workingFolderId) throw new Error('Working folder not configured.');
  var code = config.committeeCode || 'CMT';
  var subfolder = getMeetingSubfolder_(config.workingFolderId, meetingLabel, code);
  var files = subfolder.getFiles();
  var renamed = [];
  var prefix = code + '_' + meetingLabel + '_';
  while (files.hasNext()) {
    var f = files.next();
    var name = f.getName();
    if (name.indexOf(prefix) !== 0) {
      var newName = prefix + name;
      f.setName(newName);
      renamed.push({ old: name, new: newName });
    }
  }
  return JSON.stringify({ renamed: renamed.length, files: renamed, subfolder: subfolder.getName() });
}

function createMeetingSubfolders(meetingLabel) {
  var config = JSON.parse(getMeetingConfig());
  var code = config.committeeCode || 'CMT';
  var result = {};
  if (config.workingFolderId) {
    var wf = getMeetingSubfolder_(config.workingFolderId, meetingLabel, code);
    result.workingSubfolder = wf.getName();
    result.workingUrl = wf.getUrl();
  }
  if (config.publishedFolderId) {
    var pf = getMeetingSubfolder_(config.publishedFolderId, meetingLabel, code);
    result.publishedSubfolder = pf.getName();
    result.publishedUrl = pf.getUrl();
  }
  return JSON.stringify(result);
}

function publishMeetingMinutes(minutesText, meetingLabel, workingFolderIdStr) {
  var config = JSON.parse(getMeetingConfig());
  if (!config.publishedFolderId) throw new Error('Published folder ID not configured. Set it in Meeting Settings.');
  var code = config.committeeCode || 'CMT';
  var publishedSubfolder = getMeetingSubfolder_(config.publishedFolderId, meetingLabel, code);
  var minutesFileName = code + '_' + meetingLabel + '_MeetingMinutes';
  var doc = DocumentApp.create(minutesFileName);
  doc.getBody().setText(minutesText);
  doc.saveAndClose();
  var savedDoc = DocumentApp.openById(doc.getId());
  var pdf = savedDoc.getAs('application/pdf');
  pdf.setName(minutesFileName + '.pdf');
  var pdfFile = publishedSubfolder.createFile(pdf);
  DriveApp.getFileById(savedDoc.getId()).setTrashed(true);
  if (config.workingFolderId) {
    try {
      var workingSubfolder = getMeetingSubfolder_(config.workingFolderId, meetingLabel, code);
      var prefix = code + '_' + meetingLabel + '_';
      var wfiles = workingSubfolder.getFiles();
      while (wfiles.hasNext()) {
        var f = wfiles.next();
        var name = f.getName();
        var newName = name.indexOf(prefix) === 0 ? name : prefix + name;
        var mimeType = f.getMimeType();
        if (mimeType === 'application/pdf') {
          f.makeCopy(newName, publishedSubfolder);
          f.setTrashed(true);
        } else {
          try {
            var exportUrl = 'https://www.googleapis.com/drive/v3/files/' + f.getId() + '/export?mimeType=application/pdf';
            var token = ScriptApp.getOAuthToken();
            var response = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
            if (response.getResponseCode() === 200) {
              var blob = response.getBlob();
              blob.setName(newName.replace(/\.[^.]+$/, '') + '.pdf');
              publishedSubfolder.createFile(blob);
              f.setTrashed(true);
            } else {
              f.makeCopy(newName, publishedSubfolder);
              f.setTrashed(true);
            }
          } catch(convErr) {
            f.makeCopy(newName, publishedSubfolder);
            f.setTrashed(true);
          }
        }
      }
    } catch(e) {}
  }
  var attendees = config.attendees || [];
  if (attendees.length) {
    var senderName = getSenderName_();
    var tmpl = getEmailTemplate_(config, 'minutes');
    var vars = { committee: code, title: config.title||'Committee Meeting', date: meetingLabel, url: pdfFile.getUrl(), senderName: senderName };
    var subj = renderTemplate_(tmpl.subject, vars);
    var emailBody = renderTemplate_(tmpl.body, vars);
    attendees.forEach(function(email) { GmailApp.sendEmail(email, subj, emailBody, {name: senderName}); });
  }
  return JSON.stringify({ pdfUrl: pdfFile.getUrl(), pdfName: pdfFile.getName(), subfolder: publishedSubfolder.getName() });
}

function chaseActionItems() {
  var items = JSON.parse(getMeetingActionItems());
  var config = JSON.parse(getMeetingConfig());
  var open = items.filter(function(item) { return item.status === 'open'; });
  var chased = 0;
  var senderName = getSenderName_();
  var tmpl = getEmailTemplate_(config, 'chase');
  open.forEach(function(item) {
    if (!item.ownerEmail) return;
    var itemLine = 'Action: ' + (item.action||'') + '\nDue: ' + (item.dueDate||'TBD') + '\nMeeting: ' + (item.meetingDate||'');
    var vars = { committee: config.committeeCode||'CMT', title: config.title||'Committee Meeting', date: item.meetingDate||'', ownerName: item.owner||'', items: itemLine, senderName: senderName };
    var subj = renderTemplate_(tmpl.subject, vars);
    var body = renderTemplate_(tmpl.body, vars);
    try { GmailApp.sendEmail(item.ownerEmail, subj, body, {name: senderName}); chased++; } catch(e) {}
  });
  return JSON.stringify({ chased: chased, total: open.length });
}

function chaseSelectedItems(idsJson) {
  var ids = JSON.parse(idsJson);
  var items = JSON.parse(getMeetingActionItems());
  var config = JSON.parse(getMeetingConfig());
  var selected = items.filter(function(item) { return ids.indexOf(item.id) !== -1; });
  var chased = 0;
  var senderName = getSenderName_();
  var tmpl = getEmailTemplate_(config, 'chase');
  selected.forEach(function(item) {
    if (!item.ownerEmail) return;
    var itemLine = 'Action: ' + (item.action||'') + '\nDue: ' + (item.dueDate||'TBD') + '\nMeeting: ' + (item.meetingDate||'');
    var vars = { committee: config.committeeCode||'CMT', title: config.title||'Committee Meeting', date: item.meetingDate||'', ownerName: item.owner||'', items: itemLine, senderName: senderName };
    var subj = renderTemplate_(tmpl.subject, vars);
    var body = renderTemplate_(tmpl.body, vars);
    try { GmailApp.sendEmail(item.ownerEmail, subj, body, {name: senderName}); chased++; } catch(e) {}
  });
  return JSON.stringify({ chased: chased, total: selected.length });
}

function transcribeNotesImage(base64Image, mimeType) {
  var apiKey = requireApiKey_();
  var payload = JSON.stringify({
    model: MODEL,
    max_tokens: 2000,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mimeType || 'image/jpeg', data: base64Image } },
        { type: 'text', text: 'Transcribe ALL text visible in this image of meeting notes. Preserve structure, bullet points, names, numbers, and action items exactly as written. Output plain text only — no commentary, no preamble.' }
      ]
    }]
  });
  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
    payload: payload,
    muteHttpExceptions: true,
  });
  var result = JSON.parse(response.getContentText());
  if (result.error) throw new Error(result.error.message);
  return result.content[0].text;
}

function getMeetingChangeRequests() {
  var config = JSON.parse(getMeetingConfig());
  var code = config.committeeCode || 'Committee';
  var threads = GmailApp.search('subject:"Re: ' + code + '" newer_than:30d', 0, 20);
  var requests = [];
  threads.forEach(function(thread) {
    var msgs = thread.getMessages();
    msgs.forEach(function(msg) {
      if (msg.isUnread()) {
        var body = msg.getPlainBody().slice(0, 500);
        var from = msg.getFrom();
        var subj = msg.getSubject();
        var lower = body.toLowerCase();
        var type = 'general';
        if (lower.indexOf('reschedule') !== -1 || lower.indexOf('postpone') !== -1) type = 'reschedule';
        else if (lower.indexOf('invite') !== -1 || lower.indexOf('add ') !== -1) type = 'add_attendee';
        else if (lower.indexOf('cancel') !== -1 || lower.indexOf('decline') !== -1) type = 'cancel';
        else if (lower.indexOf('regret') !== -1 || lower.indexOf('cannot attend') !== -1) type = 'regrets';
        requests.push({ from: from, subject: subj, snippet: body.slice(0, 200), type: type, date: msg.getDate().toISOString() });
      }
    });
  });
  return JSON.stringify(requests);
}

function createActionItemsSheet() {
  var config = JSON.parse(getMeetingConfig());
  var code = config.committeeCode || 'CMT';
  var ss = SpreadsheetApp.create(code + ' — Action Items');
  var sheet = ss.getSheets()[0];
  sheet.setName('Action Items');
  sheet.getRange(1,1,1,6).setValues([['Action','Owner','Email','Due Date','Status','Meeting Date']]);
  sheet.getRange(1,1,1,6).setFontWeight('bold');
  sheet.setFrozenRows(1);
  // Save the ID back to config
  var raw = PropertiesService.getUserProperties().getProperty('MEETING_CONFIG_' + code)
         || PropertiesService.getUserProperties().getProperty('MEETING_CONFIG');
  if (raw) {
    var cfg = JSON.parse(raw);
    cfg.actionItemsSheetId = ss.getId();
    PropertiesService.getUserProperties().setProperty('MEETING_CONFIG_' + code, JSON.stringify(cfg));
  }
  return JSON.stringify({ id: ss.getId(), url: ss.getUrl(), name: ss.getName() });
}

function syncActionItemsFromSheet() {
  var config = JSON.parse(getMeetingConfig());
  var sheetId = config.actionItemsSheetId;
  if (!sheetId) throw new Error('No spreadsheet ID configured. Add it in Meeting Settings.');
  // Accept full URL or bare ID
  var ss;
  var urlMatch = sheetId.match(/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (urlMatch) {
    ss = SpreadsheetApp.openById(urlMatch[1]);
  } else {
    ss = SpreadsheetApp.openById(sheetId);
  }
  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return JSON.stringify({ imported: 0, total: 0 });
  var hdr = data[0].map(function(h) { return String(h).toLowerCase().trim(); });
  function col(names) { for (var i=0;i<names.length;i++){var x=hdr.indexOf(names[i]);if(x>=0)return x;} return -1; }
  var aIdx=col(['action','action item','description','task']);
  var oIdx=col(['owner','assigned to','name']);
  var eIdx=col(['email','owner email','email address']);
  var dIdx=col(['due date','due','deadline']);
  var sIdx=col(['status','state']);
  var mIdx=col(['meeting date','meeting','date']);
  var items = [];
  var now = new Date().toISOString();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var action = aIdx >= 0 ? String(row[aIdx]||'').trim() : '';
    if (!action) continue;
    items.push({
      id: 'sheet_' + i + '_' + Date.now(),
      action: action,
      owner:      oIdx>=0 ? String(row[oIdx]||'').trim() : '',
      ownerEmail: eIdx>=0 ? String(row[eIdx]||'').trim() : '',
      dueDate:    dIdx>=0 ? String(row[dIdx]||'').trim() : '',
      status:     sIdx>=0 ? String(row[sIdx]||'open').trim().toLowerCase() : 'open',
      meetingDate:mIdx>=0 ? String(row[mIdx]||'').trim() : '',
      createdAt: now, source: 'sheet',
    });
  }
  var existing = JSON.parse(getMeetingActionItems());
  var nonSheet = existing.filter(function(item) { return item.source !== 'sheet'; });
  var merged = items.concat(nonSheet);
  var code = getActiveCommitteeCode_();
  PropertiesService.getUserProperties().setProperty(MEETING_ACTION_ITEMS_PROP + '_' + code, JSON.stringify(merged));
  return JSON.stringify({ imported: items.length, total: merged.length });
}

function setupWeeklyChase() {
  removeWeeklyChase();
  ScriptApp.newTrigger('chaseActionItems').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  return 'ok';
}

function removeWeeklyChase() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'chaseActionItems') ScriptApp.deleteTrigger(t);
  });
  return 'ok';
}

function getWeeklyChaseStatus() {
  var active = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'chaseActionItems'; });
  return JSON.stringify({ active: active });
}


function replyToChangeRequest(msgId, replyType) {
  var msg = GmailApp.getMessageById(msgId);
  if (!msg) throw new Error('Message not found.');
  var replies = { acknowledge: 'Thank you for your message. The Meeting Suite has been notified and will follow up shortly.\n\nThis is an automated acknowledgement from QuestFlow.', reschedule: 'Thank you for your reschedule request. The Meeting Suite will review available dates and send a revised meeting notice.\n\nThis is an automated acknowledgement from QuestFlow.', regrets: 'Thank you for your regrets. Your absence has been noted by the Meeting Suite.\n\nThis is an automated acknowledgement from QuestFlow.' };
  var body = replies[replyType] || replies['acknowledge'];
  msg.getThread().reply(body);
  msg.markRead();
  return 'ok';
}

// ── EMAIL TEMPLATE ENGINE ─────────────────────────────────────────────────────
var EMAIL_DEFAULTS_ = {
  notice:      { subject: '{{committee}} Meeting — Notice of Meeting — {{date}}',
                 body:    'Dear Committee Members,\n\nThis is to confirm the upcoming {{title}}.\n\nDate: {{date}}\nTime: {{time}}\nLocation: {{room}}\n\nThe agenda and materials will be circulated in advance of the meeting.\n\nRegards,\n{{senderName}}\nMeeting Suite' },
  materials:   { subject: '{{committee}} Meeting — Materials Request — {{date}}',
                 body:    'Dear Committee Members,\n\nThe next {{title}} is scheduled for {{date}}.\n\nPlease submit your materials to the working folder two business days prior to the meeting. All materials must follow the naming convention: {{committee}}_{{date}}_[Description]\n\nPlease confirm your attendance or send regrets to the Meeting Suite.\n\nThank you,\n{{senderName}}\nMeeting Suite' },
  reminder:    { subject: 'Meeting Reminder: {{title}} — {{date}}',
                 body:    'Dear Committee Members,\n\nThis is a reminder that the {{title}} is scheduled for tomorrow.\n\nDate: {{date}}\nTime: {{time}}\nLocation: {{room}}\n\nPlease ensure you have reviewed the agenda and any circulated materials in advance.\n\nRegards,\n{{senderName}}\nMeeting Suite' },
  chase:       { subject: '{{committee}} — Action Item Follow-Up',
                 body:    'Dear {{ownerName}},\n\nThe following action item(s) from the {{title}} are outstanding and assigned to you:\n\n{{items}}\n\nPlease provide a status update to the Meeting Suite at your earliest convenience.\n\nRegards,\n{{senderName}}\nMeeting Suite' },
  package:     { subject: '{{committee}} Meeting — Pre-Meeting Package — {{date}}',
                 body:    'Dear Committee Members,\n\nPlease find below the pre-meeting package for the {{title}} scheduled for {{date}}.\n\nDate & Time: {{date}} at {{time}}\nLocation: {{room}}\n\n{{files}}\n\nPlease review all materials in advance of the meeting. If you have any questions, please contact the Meeting Suite.\n\nRegards,\n{{senderName}}\nMeeting Suite' },
  minutes:     { subject: '{{committee}} Meeting Minutes — Published — {{date}}',
                 body:    'Dear Committee Members,\n\nThe minutes for the {{date}} {{title}} have been published.\n\nMinutes: {{url}}\n\nPlease review and submit any corrections to the Meeting Suite within 5 business days.\n\nRegards,\n{{senderName}}\nMeeting Suite' },
  followup:    { subject: '{{committee}} — Action Item Follow-Up — {{date}}',
                 body:    'Dear {{ownerName}},\n\nFollowing the {{title}} on {{date}}, the following action item(s) are assigned to you:\n\n{{items}}\n\nPlease provide a status update to the Meeting Suite at your earliest convenience.\n\nRegards,\n{{senderName}}\nMeeting Suite' },
  circulation: { subject: '[DRAFT — FOR REVIEW] {{committee}} Meeting Minutes — {{date}}',
                 body:    'Dear Committee Members,\n\nPlease find below the DRAFT minutes for the {{title}} on {{date}} for your review.\n\nPlease reply to this email with any corrections or amendments within 5 business days. If you have no corrections, a simple reply of \'Approved as circulated\' is sufficient.\n\n─────────────────────────────────────────\n{{minutesText}}\n─────────────────────────────────────────\n\nThis draft is CONFIDENTIAL and for committee review only.\n\nRegards,\n{{senderName}}\nMeeting Suite' },
};

function getEmailTemplate_(config, key) {
  var tmpl = (config.emailTemplates || {})[key] || {};
  var def = EMAIL_DEFAULTS_[key] || { subject: '', body: '' };
  return { subject: tmpl.subject != null ? tmpl.subject : def.subject,
           body:    tmpl.body    != null ? tmpl.body    : def.body };
}

function renderTemplate_(tmpl, vars) {
  var out = tmpl;
  Object.keys(vars).forEach(function(k) {
    out = out.split('{{' + k + '}}').join(vars[k] != null ? String(vars[k]) : '');
  });
  return out;
}

function sendTestEmail(templateKey, configJson) {
  var config = configJson ? JSON.parse(configJson) : JSON.parse(getMeetingConfig());
  var senderName = config.senderName || getSenderName_();
  var myEmail = '';
  try { myEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail(); } catch(e) {}
  if (!myEmail) throw new Error('Could not determine your email address.');
  var tmpl = getEmailTemplate_(config, templateKey);
  var timeStr = (config.hour || 10) + ':' + ('0' + (config.minute || 0)).slice(-2);
  var vars = {
    committee: config.committeeCode || 'CMT',
    title: config.title || 'Committee Meeting',
    date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    time: timeStr,
    room: config.room || 'TBD',
    senderName: senderName,
    ownerName: senderName,
    url: 'https://drive.google.com (example)',
    files: 'MEETING MATERIALS:\n• Example_Agenda.pdf\n  https://drive.google.com/example',
    items: '1. [Sample action item]\n   Due: 2026-08-01',
    minutesText: '[Draft minutes content would appear here]',
  };
  var subject = '[TEST] ' + renderTemplate_(tmpl.subject, vars);
  var body = renderTemplate_(tmpl.body, vars);
  GmailApp.sendEmail(myEmail, subject, body, { name: senderName });
  return JSON.stringify({ to: myEmail });
}

// ── AUTO-AGENDA BUILDER ──────────────────────────────────────────────────────
function getSenderName_() {
  var config = {};
  try { config = JSON.parse(getMeetingConfig()); } catch(e) {}
  if (config.senderName) return config.senderName;
  var email = '';
  try { email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || ''; } catch(e) {}
  if (!email) return 'Meeting Suite';
  var local = email.split('@')[0];
  return local.replace(/[._\-]/g, ' ').replace(/\b\w/g, function(c) { return c.toUpperCase(); });
}
function buildAgenda(meetingLabel) {
  var config = JSON.parse(getMeetingConfig());
  if (!config.workingFolderId) throw new Error('Working folder not configured in Meeting Settings.');
  var code = config.committeeCode || 'CMT';
  var label = meetingLabel;

  // Pull open action items from previous meetings
  var allItems = JSON.parse(getMeetingActionItems());
  var openItems = allItems.filter(function(item) { return item.status !== 'done'; });

  // Get API key for Claude
  var apiKey = PropertiesService.getUserProperties().getProperty('ANTHROPIC_API_KEY') || '';
  var agendaBody;

  if (apiKey && openItems.length > 0) {
    var itemsSummary = openItems.slice(0, 20).map(function(item, i) {
      return (i+1) + '. [' + (item.owner||'TBD') + '] ' + item.action + (item.dueDate ? ' (due ' + item.dueDate + ')' : '');
    }).join('\n');
    var prompt = 'You are preparing a formal agenda for the ' + (config.title || 'Committee Meeting') + ' on ' + label + '.\n\n' +
      'Open action items from previous meetings:\n' + itemsSummary + '\n\n' +
      'Write a formal meeting agenda. Include: call to order, approval of previous agenda, approval of previous minutes, ' +
      'then a "Standing Business" section reviewing each open action item by owner, then "New Business" (leave as TBD), ' +
      'then next meeting date, then adjournment. Use numbered sections. Keep it concise and professional.';
    try {
      agendaBody = callClaude_(apiKey, 'You are a formal meeting agenda writer. Write a professional meeting agenda.', prompt, 1024, MODEL);
    } catch(e) {}
  }

  if (!agendaBody) {
    // Fallback: structured template with action items
    var actionSection = openItems.length > 0
      ? openItems.slice(0, 20).map(function(item, i) {
          return '   5.' + (i+1) + ' [' + (item.owner||'TBD') + '] ' + item.action + (item.dueDate ? ' — due ' + item.dueDate : '');
        }).join('\n')
      : '   (No outstanding action items)';
    agendaBody = [
      code + ' — ' + (config.title || 'Committee Meeting'),
      'MEETING AGENDA — CONFIDENTIAL',
      '',
      'Date: ' + label,
      'Time: ' + (config.hour||10) + ':' + ('0'+(config.minute||0)).slice(-2),
      'Location: ' + (config.room || 'TBD'),
      'Attendees: ' + (config.attendees || []).join(', '),
      '',
      '1. CALL TO ORDER',
      '',
      '2. APPROVAL OF AGENDA',
      '   RESOLVED that the agenda be approved as presented.',
      '',
      '3. APPROVAL OF PREVIOUS MINUTES',
      '   RESOLVED that the minutes of the previous meeting be approved as circulated.',
      '',
      '4. CONFLICTS OF INTEREST',
      '   Members to declare any conflicts with items on the agenda.',
      '',
      '5. STANDING BUSINESS — ACTION ITEM REVIEW',
      actionSection,
      '',
      '6. NEW BUSINESS',
      '   6.1 [To be determined]',
      '',
      '7. DATE OF NEXT MEETING',
      '   [To be confirmed]',
      '',
      '8. ADJOURNMENT',
    ].join('\n');
  }

  var title = code + '_' + label + '_Agenda';
  var subfolder = getMeetingSubfolder_(config.workingFolderId, label, code);
  var doc = DocumentApp.create(title);
  doc.getBody().setText(agendaBody);
  DriveApp.getFileById(doc.getId()).moveTo(subfolder);
  return JSON.stringify({ url: doc.getUrl(), name: title, subfolder: subfolder.getName(), actionItemCount: openItems.length });
}

// ── PRE-MEETING PACKAGE EMAILER ──────────────────────────────────────────────
function sendMeetingPackage(meetingLabel) {
  var config = JSON.parse(getMeetingConfig());
  var attendees = config.attendees || [];
  if (!attendees.length) throw new Error('No attendees configured. Add emails in Meeting Settings.');
  var code = config.committeeCode || 'CMT';

  // Find files in working subfolder for this meeting
  var links = [];
  if (config.workingFolderId) {
    try {
      var subfolder = getMeetingSubfolder_(config.workingFolderId, meetingLabel, code);
      var files = subfolder.getFiles();
      while (files.hasNext()) {
        var f = files.next();
        links.push({ name: f.getName(), url: f.getUrl() });
      }
    } catch(e) {}
  }

  var senderName = getSenderName_();
  var timeStr = (config.hour||10) + ':' + ('0'+(config.minute||0)).slice(-2);
  var filesText = links.length > 0
    ? 'MEETING MATERIALS:\n' + links.map(function(l){ return '• ' + l.name + '\n  ' + l.url; }).join('\n')
    : 'Materials will be circulated separately.';
  var tmpl = getEmailTemplate_(config, 'package');
  var vars = { committee: code, title: config.title||'Committee Meeting', date: meetingLabel, time: timeStr, room: config.room||'TBD', files: filesText, senderName: senderName };
  var subject = renderTemplate_(tmpl.subject, vars);
  var body = renderTemplate_(tmpl.body, vars);
  attendees.forEach(function(email) { GmailApp.sendEmail(email, subject, body, {name: senderName}); });
  return JSON.stringify({ sent: attendees.length, filesAttached: links.length, recipients: attendees });
}

// ── POST-MEETING ACTION FOLLOW-UPS ───────────────────────────────────────────
function sendActionFollowups(meetingLabel) {
  var config = JSON.parse(getMeetingConfig());
  var code = config.committeeCode || 'CMT';
  var allItems = JSON.parse(getMeetingActionItems());

  // Get items from this meeting OR all open items if no meeting filter
  var items = (meetingLabel && meetingLabel !== 'all')
    ? allItems.filter(function(i) { return i.meetingDate === meetingLabel && i.status !== 'done'; })
    : allItems.filter(function(i) { return i.status !== 'done'; });

  if (!items.length) return JSON.stringify({ sent: 0, message: 'No open action items for ' + (meetingLabel || 'any meeting') });

  // Group by owner email
  var byOwner = {};
  items.forEach(function(item) {
    var email = item.ownerEmail || '';
    if (!email) return;
    if (!byOwner[email]) byOwner[email] = [];
    byOwner[email].push(item);
  });

  var sent = 0;
  var skipped = 0;
  var senderName = getSenderName_();
  var tmpl = getEmailTemplate_(config, 'followup');
  Object.keys(byOwner).forEach(function(email) {
    var ownerItems = byOwner[email];
    var ownerName = ownerItems[0].owner || email;
    var itemsText = ownerItems.map(function(item, i) {
      var line = (i+1) + '. ' + item.action;
      if (item.dueDate) line += '\n   Due: ' + item.dueDate;
      if (item.meetingDate) line += '\n   Meeting: ' + item.meetingDate;
      return line;
    }).join('\n\n');
    var vars = { committee: code, title: config.title||'Committee Meeting', date: meetingLabel||'', ownerName: ownerName, items: itemsText, senderName: senderName };
    var subject = renderTemplate_(tmpl.subject, vars);
    var body = renderTemplate_(tmpl.body, vars);
    try { GmailApp.sendEmail(email, subject, body, {name: senderName}); sent++; } catch(e) { skipped++; }
  });

  var noEmail = items.filter(function(i) { return !i.ownerEmail; }).length;
  return JSON.stringify({ sent: sent, skipped: skipped, noEmail: noEmail, total: items.length });
}

// ── RESOLUTION REGISTER ───────────────────────────────────────────────────────
var MEETING_RESOLUTIONS_PROP = 'MEETING_RESOLUTIONS';

function getMeetingResolutions() {
  var code = getActiveCommitteeCode_();
  return PropertiesService.getUserProperties().getProperty(MEETING_RESOLUTIONS_PROP + '_' + code) || '[]';
}

function extractAndSaveResolutions(minutesText, meetingLabel) {
  var config = JSON.parse(getMeetingConfig());
  var code = config.committeeCode || 'CMT';
  var apiKey = PropertiesService.getUserProperties().getProperty('ANTHROPIC_API_KEY') || '';

  var resolved = [];
  if (apiKey) {
    try {
      var prompt = 'Extract every formal resolution from these meeting minutes. ' +
        'A resolution is a sentence starting with "RESOLVED" or "BE IT RESOLVED". ' +
        'Return ONLY a JSON array of strings — each string is the full resolution text (starting with RESOLVED). ' +
        'If none found, return []. No markdown, no explanation.\n\nMINUTES:\n' + minutesText;
      var raw = callClaude_(apiKey, 'Extract formal resolutions from meeting minutes. Return ONLY a JSON array of strings, each starting with RESOLVED. If none found, return [].', prompt, 1024, MODEL);
      resolved = parseJson_(raw.replace(/```json|```/g, '').trim(), []);
    } catch(e) {}
  }

  // Fallback: regex extraction
  if (!resolved.length) {
    var matches = minutesText.match(/RESOLVED[^\n]+/g) || [];
    resolved = matches.map(function(m) { return m.trim(); });
  }

  if (!resolved.length) return JSON.stringify({ count: 0 });

  var existing = JSON.parse(getMeetingResolutions());
  var year = meetingLabel ? meetingLabel.slice(0, 4) : new Date().getFullYear().toString();

  // Count existing resolutions for this year to assign sequential numbers
  var yearCount = existing.filter(function(r) { return r.id && r.id.indexOf('R-' + year) === 0; }).length;

  var newEntries = resolved.map(function(text, i) {
    yearCount++;
    var num = yearCount < 10 ? '00' + yearCount : yearCount < 100 ? '0' + yearCount : '' + yearCount;
    return {
      id: 'R-' + year + '-' + num,
      meetingDate: meetingLabel || '',
      committeeCode: code,
      text: text,
      extractedAt: new Date().toISOString()
    };
  });

  var merged = existing.concat(newEntries);
  PropertiesService.getUserProperties().setProperty(MEETING_RESOLUTIONS_PROP + '_' + code, JSON.stringify(merged));
  return JSON.stringify({ count: newEntries.length, resolutions: newEntries });
}

// ── MEETING STATUS TRACKER ────────────────────────────────────────────────────
function getMeetingStatuses() {
  var code = getActiveCommitteeCode_();
  return PropertiesService.getUserProperties().getProperty(MEETING_STATUS_PROP + '_' + code) || '{}';
}

function setMeetingStatus(meetingLabel, status) {
  var code = getActiveCommitteeCode_();
  var statuses = JSON.parse(getMeetingStatuses());
  statuses[meetingLabel] = status;
  PropertiesService.getUserProperties().setProperty(MEETING_STATUS_PROP + '_' + code, JSON.stringify(statuses));
  return 'ok';
}

// ── AGENDA BUILDER ────────────────────────────────────────────────────────────
function getAgendaItems(meetingLabel) {
  var code = getActiveCommitteeCode_();
  return PropertiesService.getUserProperties().getProperty(MEETING_AGENDA_PROP + '_' + code + '_' + meetingLabel) || '[]';
}

function saveAgendaItems(meetingLabel, itemsJson) {
  var code = getActiveCommitteeCode_();
  PropertiesService.getUserProperties().setProperty(MEETING_AGENDA_PROP + '_' + code + '_' + meetingLabel, itemsJson);
  return 'ok';
}

function exportAgendaDoc(meetingLabel, itemsJson) {
  var config = JSON.parse(getMeetingConfig());
  var items = JSON.parse(itemsJson || '[]');
  var code = config.committeeCode || 'CMT';
  var title = code + '_' + meetingLabel + '_Agenda';
  var doc = DocumentApp.create(title);
  var body = doc.getBody();
  body.appendParagraph(config.title || 'Committee Meeting').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('AGENDA').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Date: ' + meetingLabel);
  body.appendParagraph('Time: ' + (config.hour || '') + ':' + (String(config.minute || 0).padStart(2, '0')));
  body.appendParagraph('Location: ' + (config.room || 'TBD'));
  body.appendParagraph('');
  var totalMins = items.reduce(function(s, i) { return s + (parseInt(i.duration, 10) || 0); }, 0);
  items.forEach(function(item, idx) {
    var line = (idx + 1) + '. ' + (item.title || '') +
      (item.presenter ? ' — ' + item.presenter : '') +
      (item.duration ? ' (' + item.duration + ' min)' : '') +
      (item.type ? ' [' + item.type + ']' : '');
    body.appendParagraph(line);
  });
  body.appendParagraph('');
  body.appendParagraph('Total time: ' + totalMins + ' minutes');
  doc.saveAndClose();
  if (config.workingFolderId) {
    try {
      var file = DriveApp.getFileById(doc.getId());
      DriveApp.getFolderById(config.workingFolderId).addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch(e) {}
  }
  return JSON.stringify({ url: doc.getUrl(), id: doc.getId() });
}

// ── COI REGISTRY ─────────────────────────────────────────────────────────────
function getCOIRegistry() {
  var code = getActiveCommitteeCode_();
  return PropertiesService.getUserProperties().getProperty(COI_REGISTRY_PROP + '_' + code) || '[]';
}

function saveCOIEntries(meetingLabel, meetingDate, entriesJson) {
  var code = getActiveCommitteeCode_();
  var existing = JSON.parse(getCOIRegistry());
  var entries = JSON.parse(entriesJson || '[]');
  var now = new Date().toISOString();
  var newEntries = entries.map(function(e) {
    return { id: 'coi_' + Date.now() + '_' + Math.random().toString(36).slice(2,6),
      committeeCode: code, meetingLabel: meetingLabel, meetingDate: meetingDate || meetingLabel,
      member: e.member, item: e.item, savedAt: now };
  });
  var merged = existing.concat(newEntries);
  PropertiesService.getUserProperties().setProperty(COI_REGISTRY_PROP + '_' + code, JSON.stringify(merged));
  return JSON.stringify({ count: newEntries.length });
}

function deleteCOIEntry(entryId) {
  var code = getActiveCommitteeCode_();
  var existing = JSON.parse(getCOIRegistry());
  var filtered = existing.filter(function(e) { return e.id !== entryId; });
  PropertiesService.getUserProperties().setProperty(COI_REGISTRY_PROP + '_' + code, JSON.stringify(filtered));
  return 'ok';
}

// ── CROSS-PORTFOLIO RESOLUTION SEARCH ─────────────────────────────────────────
function getAllResolutions() {
  var list = getCommitteesList_();
  var all = [];
  list.forEach(function(code) {
    var raw = PropertiesService.getUserProperties().getProperty(MEETING_RESOLUTIONS_PROP + '_' + code);
    if (raw) {
      try {
        var entries = JSON.parse(raw);
        entries.forEach(function(e) { if (!e.committeeCode) e.committeeCode = code; });
        all = all.concat(entries);
      } catch(e) {}
    }
  });
  all.sort(function(a, b) { return (b.meetingDate || '').localeCompare(a.meetingDate || ''); });
  return JSON.stringify(all);
}

// ── MINUTES CIRCULATION ───────────────────────────────────────────────────────
function circulateDraftMinutes(minutesText, meetingLabel) {
  var config = JSON.parse(getMeetingConfig());
  var attendees = config.attendees || [];
  if (!attendees.length) throw new Error('No attendees configured. Add emails in Meeting Settings.');
  var code = config.committeeCode || 'CMT';
  var senderName = getSenderName_();

  var tmpl = getEmailTemplate_(config, 'circulation');
  var vars = {
    committee: code, title: config.title || 'Committee Meeting',
    date: meetingLabel || 'Pending', senderName: senderName,
    minutesText: minutesText
  };
  var subject = renderTemplate_(tmpl.subject, vars);
  var body = renderTemplate_(tmpl.body, vars);

  var sent = 0;
  attendees.forEach(function(email) {
    try {
      GmailApp.sendEmail(email, subject, body, { name: senderName });
      sent++;
    } catch(e) {}
  });
  return JSON.stringify({ sent: sent, meetingLabel: meetingLabel });
}

// ── MEETING REMINDERS ─────────────────────────────────────────────────────────
function setupMeetingReminders() {
  removeMeetingReminders();
  ScriptApp.newTrigger('sendMeetingReminders_').timeBased().everyDays(1).atHour(8).create();
  PropertiesService.getUserProperties().setProperty('MEETING_REMINDER_ACTIVE', 'true');
  return 'ok';
}

function removeMeetingReminders() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendMeetingReminders_') ScriptApp.deleteTrigger(t);
  });
  PropertiesService.getUserProperties().deleteProperty('MEETING_REMINDER_ACTIVE');
  return 'ok';
}

function getMeetingReminderStatus() {
  var active = PropertiesService.getUserProperties().getProperty('MEETING_REMINDER_ACTIVE') === 'true';
  return JSON.stringify({ active: active });
}

function sendMeetingReminders_() {
  var codes = getCommitteesList_();
  codes.forEach(function(code) {
    try {
      PropertiesService.getUserProperties().setProperty('MEETING_ACTIVE_PROP_TEMP', code);
      var raw = PropertiesService.getUserProperties().getProperty(MEETING_CONFIG_PROP + '_' + code);
      if (!raw) return;
      var config = JSON.parse(raw);
      var attendees = config.attendees || [];
      if (!attendees.length) return;
      var senderName = config.senderName || getSenderName_();

      // Calculate upcoming meeting dates for next 30 days
      var now = new Date();
      var horizon = new Date(now.getTime() + 30 * 86400000);
      var targetDay = config.dayOfWeek || 3;
      var reminderHours = 48;

      // Scan next 30 days for meeting dates matching committee schedule
      var d = new Date(now);
      d.setHours(0,0,0,0);
      while (d <= horizon) {
        if (d.getDay() === targetDay) {
          // Check if this is a 1st or 3rd occurrence of that weekday in the month
          var dayOfMonth = d.getDate();
          var firstOccurrence = dayOfMonth <= 7;
          var thirdOccurrence = dayOfMonth >= 15 && dayOfMonth <= 21;
          if (firstOccurrence || thirdOccurrence) {
            var meetingTime = new Date(d);
            meetingTime.setHours(config.hour || 10, config.minute || 0, 0, 0);
            var hoursUntil = (meetingTime - now) / 3600000;
            if (hoursUntil > 0 && hoursUntil <= reminderHours + 2) {
              var label = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
              var timeStr = (config.hour || 10) + ':' + ('0' + (config.minute || 0)).slice(-2);
              var reminderTmpl = getEmailTemplate_(config, 'reminder');
              var reminderVars = {
                committee: code, title: config.title || 'Committee Meeting',
                date: label, time: timeStr, room: config.room || 'TBD', senderName: senderName
              };
              var subject = renderTemplate_(reminderTmpl.subject, reminderVars);
              var body = renderTemplate_(reminderTmpl.body, reminderVars);
              attendees.forEach(function(email) {
                try { GmailApp.sendEmail(email, subject, body, { name: senderName }); } catch(e) {}
              });
            }
          }
        }
        d.setDate(d.getDate() + 1);
      }
    } catch(e) {}
  });
}

// ── COMMITTEE ONBOARDING ──────────────────────────────────────────────────────
function createCommitteeFolders(code, title) {
  var rootFolderName = 'QuestFlow — ' + (title || code);
  var root = DriveApp.createFolder(rootFolderName);
  var working = root.createFolder(code + ' — Working Documents');
  var published = root.createFolder(code + ' — Published Minutes');
  return JSON.stringify({
    rootId: root.getId(),
    rootUrl: root.getUrl(),
    workingId: working.getId(),
    workingUrl: working.getUrl(),
    publishedId: published.getId(),
    publishedUrl: published.getUrl()
  });
}
