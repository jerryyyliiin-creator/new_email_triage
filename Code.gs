// ---------------------------------------------------------------------------
// Inbox Triage - Gmail Add-on + Web App
// ---------------------------------------------------------------------------

const API_KEY_PROP = 'ANTHROPIC_API_KEY';
const MODEL = 'claude-haiku-4-5-20251001';
const CLASSIFICATION_CACHE_PROP = 'EMAIL_CLASS_CACHE';
const CLASSIFICATION_CACHE_TTL_MS = 24 * 3600 * 1000;
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxDNqLP0-NT77sc2n_l4zTugk3KfouqEcnZxCrSEfX-NLXJmS2HxFq2Cxn0At1rRC7c/exec';

const CAT_META = {
  escalation: { label: 'Escalation', emoji: '🔥', icon: CardService.Icon.STAR },
  action_required: { label: 'Action Required', emoji: '🔴', icon: CardService.Icon.EMAIL },
  calendar: { label: 'Calendar', emoji: '📅', icon: CardService.Icon.INVITE },
  fyi: { label: 'FYI', emoji: 'ℹ️', icon: CardService.Icon.DESCRIPTION },
  low_priority: { label: 'Low Priority', emoji: '📥', icon: CardService.Icon.EMAIL },
};

const CAT_ORDER = {
  escalation: 0,
  action_required: 1,
  calendar: 2,
  fyi: 3,
  low_priority: 4,
};

const PRIORITY_LABELS = {
  5: '🔴 Critical',
  4: '🟠 High',
  3: '🟡 Medium',
  2: '🔵 Low',
  1: '⚪ Minimal',
};

const CAT_COLORS = {
  escalation: '#ff4d4d',
  action_required: '#ff8c42',
  calendar: '#00c04b',
  fyi: '#5b9cf6',
  low_priority: '#707070',
};

const TRIAGE_LABEL_NAMES = {
  escalation: 'Triage/Escalation',
  action_required: 'Triage/Action Required',
  calendar: 'Triage/Calendar',
  fyi: 'Triage/FYI',
  low_priority: 'Triage/Low Priority',
};

const RESOLVED_LABEL_NAME = 'Triage/Resolved';
const VALID_CATEGORIES = Object.keys(TRIAGE_LABEL_NAMES);
const VIP_SENDERS_PROP = 'VIP_SENDERS';

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
  return PropertiesService.getScriptProperties().getProperty(API_KEY_PROP);
}

function requireApiKey_() {
  const key = getApiKey_();
  if (!key) throw new Error('No API key set. Open the Gmail sidebar add-on to configure it.');
  return key;
}

function parseSender_(from) {
  const raw = from || '';
  const match = raw.match(/<([^>]+)>/);
  const email = (match ? match[1] : raw).trim();
  const name = raw.replace(/<[^>]+>/g, '').replace(/^"|"$/g, '').trim() || email;
  return { name, email };
}

function normalizeCategory_(category) {
  return VALID_CATEGORIES.indexOf(category) === -1 ? 'low_priority' : category;
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

function callClaude_(apiKey, system, user, maxTokens) {
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify({
      model: MODEL,
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
  const threads = GmailApp.getInboxThreads(0, limit || 20);
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

function extractTasksFromEmails(emailData) {
  const apiKey = getApiKey_();
  if (!apiKey) return getTasks();

  const actionable = (emailData || []).filter(function(e) {
    return e.category === 'action_required' || e.category === 'escalation';
  });
  if (!actionable.length) return getTasks();

  const withBodies = actionable.map(function(email) {
    try {
      const body = GmailApp.getMessageById(email.id)
        .getPlainBody()
        .substring(0, 600)
        .replace(/\s+/g, ' ')
        .trim();
      return Object.assign({}, email, { body });
    } catch (e) {
      return Object.assign({}, email, { body: email.summary || '' });
    }
  });

  const payload = withBodies.map(function(email) {
    return {
      messageId: email.id,
      subject: email.subject,
      sender: email.sender_name,
      body: email.body,
    };
  });

  const text = callClaude_(
    apiKey,
    'You are a JSON-only task extraction API. Return only a valid JSON array. No markdown, no prose.',
    'Extract concrete action items the recipient must do from these emails. Return ONLY a JSON array.\n\n' +
      'Each object: messageId (exact match), tasks (array of {description (verb-first, max 12 words), dueDate (ISO date string if mentioned, else null), priority ("high"|"medium"|"low")}).\n' +
      'high = has a deadline or is urgent. medium = standard reply or action needed. low = optional or nice-to-have.\n' +
      'If no tasks for an email, return empty tasks array.\n\n' + JSON.stringify(payload),
    1500
  );

  const extracted = extractJsonArray_(text);
  const now = new Date().toISOString();
  const tasks = [];

  extracted.forEach(function(item) {
    const email = withBodies.find(function(e) { return e.id === item.messageId; });
    if (!email || !item.tasks || !item.tasks.length) return;

    item.tasks.forEach(function(task) {
      if (!task.description) return;
      const priority = ['high', 'medium', 'low'].indexOf(task.priority) === -1 ? 'medium' : task.priority;
      tasks.push({
        id: 't_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7),
        description: String(task.description).substring(0, 120),
        dueDate: task.dueDate || null,
        priority,
        messageId: email.id,
        subject: email.subject,
        sender_name: email.sender_name,
        sender_email: email.sender_email,
        addedAt: now,
        done: false,
      });
    });
  });

  const props = getUserProps_();
  const existing = parseJson_(props.getProperty('TASK_DATA'), []);
  const refreshedIds = new Set(actionable.map(function(e) { return e.id; }));
  const kept = existing.filter(function(t) { return !refreshedIds.has(t.messageId); });
  const merged = tasks.concat(kept).slice(0, 300);
  props.setProperty('TASK_DATA', JSON.stringify(merged));
  return merged;
}

function getTasks() {
  return parseJson_(getUserProps_().getProperty('TASK_DATA'), []);
}

function markTaskDone(taskId, done) {
  const props = getUserProps_();
  const tasks = getTasks().map(function(task) {
    return task.id === taskId ? Object.assign({}, task, { done: done === true || done === 'true' }) : task;
  });
  props.setProperty('TASK_DATA', JSON.stringify(tasks));
  return tasks;
}

function deleteTask(taskId) {
  const props = getUserProps_();
  const tasks = getTasks().filter(function(task) { return task.id !== taskId; });
  props.setProperty('TASK_DATA', JSON.stringify(tasks));
  return tasks;
}

function clearCompletedTasks() {
  const props = getUserProps_();
  const tasks = getTasks().filter(function(task) { return !task.done; });
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

function getFollowups() {
  const data = parseJson_(getUserProps_().getProperty('FOLLOWUP_DATA'), {});
  const now = new Date();
  return Object.keys(data).map(function(id) {
    const f = data[id];
    const dueDate = new Date(f.followupAt);
    const addedDate = new Date(f.addedAt);
    const isOverdue = now > dueDate;
    const sentDays = Math.floor((now - addedDate) / 86400000);
    let dueStr;
    if (isOverdue) {
      const d = Math.floor((now - dueDate) / 86400000);
      dueStr = 'Overdue ' + d + ' day' + (d === 1 ? '' : 's');
    } else {
      const d = Math.ceil((dueDate - now) / 86400000);
      dueStr = d === 0 ? 'Due today' : d === 1 ? 'Due tomorrow' : 'Due in ' + d + ' days';
    }
    return {
      id: id,
      to: f.sender_name || f.sender_email || '(unknown)',
      subject: f.subject || '(no subject)',
      sent: sentDays === 0 ? 'today' : sentDays + ' day' + (sentDays === 1 ? '' : 's') + ' ago',
      due: dueStr,
      overdue: isOverdue,
    };
  });
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

function applyRules(emails) {
  const rules = getRules();
  if (!rules.length) return emails;

  return emails.map(function(email) {
    for (let i = 0; i < rules.length; i++) {
      const rule = rules[i];
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

function getEmailBody(messageId) {
  const msg = GmailApp.getMessageById(messageId);
  const thread = msg.getThread();
  const messages = thread.getMessages().slice(-6);
  let myEmail = '';
  try { myEmail = (Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) {}

  const structured = messages.map(function(m, i) {
    const sender = parseSender_(m.getFrom());
    return {
      index: i,
      from: sender.name,
      fromEmail: sender.email,
      date: m.getDate().toISOString(),
      body: m.getPlainBody()
        .substring(0, 1200)
        .replace(/[ \t]+/g, ' ')
        .replace(/\n{3,}/g, '\n\n')
        .trim() || '(No content)',
      isMe: !!(myEmail && sender.email.toLowerCase() === myEmail),
      bullets: null,
    };
  });

  const apiKey = getApiKey_();
  if (apiKey) {
    try {
      const payload = structured.map(function(m) {
        return { index: m.index, from: m.from, body: m.body.substring(0, 800) };
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

  return structured.map(function(m) {
    return {
      from: m.from,
      fromEmail: m.fromEmail,
      date: m.date,
      body: m.body,
      bullets: m.bullets,
      isMe: m.isMe,
    };
  });
}

function fetchContactPhotos(senderEmails) {
  const emails = (senderEmails || []).filter(Boolean).slice(0, 20);
  if (!emails.length) return {};

  const token = ScriptApp.getOAuthToken();
  const requests = emails.map(function(email) {
    return {
      url: 'https://people.googleapis.com/v1/people:searchContacts?query=' +
        encodeURIComponent(email) +
        '&readMask=photos,emailAddresses&pageSize=1' +
        '&sources=READ_SOURCE_TYPE_CONTACT' +
        '&sources=READ_SOURCE_TYPE_OTHER_CONTACT',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    };
  });

  const photoMap = {};
  UrlFetchApp.fetchAll(requests).forEach(function(res, i) {
    try {
      const data = JSON.parse(res.getContentText());
      const results = data.results || [];
      if (!results.length) return;
      const photos = results[0].person.photos || [];
      const photo = photos.find(function(p) { return !p.default; }) || photos[0];
      if (photo && photo.url) photoMap[emails[i]] = photo.url;
    } catch (e) {
      // Contact photos are optional.
    }
  });

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

function fetchDismissedEmails() {
  const data = parseJson_(getUserProps_().getProperty('DISMISSED_DATA'), {});
  const resolved = [];
  const trashed = [];

  Object.keys(data).forEach(function(id) {
    const item = Object.assign({}, data[id], { id });
    if (item.action === 'resolve') resolved.push(item);
    else trashed.push(item);
  });

  resolved.sort(function(a, b) { return new Date(b.dismissedAt) - new Date(a.dismissedAt); });
  trashed.sort(function(a, b) { return new Date(b.dismissedAt) - new Date(a.dismissedAt); });

  const all = resolved.concat(trashed);
  const uniqueEmails = Array.from(new Set(all.map(function(e) { return e.sender_email; }).filter(Boolean)));
  const photoMap = fetchContactPhotos(uniqueEmails);

  return {
    resolved: resolved.map(function(e) { return Object.assign({}, e, { photo: photoMap[e.sender_email] || null }); }),
    trashed: trashed.map(function(e) { return Object.assign({}, e, { photo: photoMap[e.sender_email] || null }); }),
  };
}

function clearDismissedEmails() {
  const props = getUserProps_();
  props.deleteProperty('DISMISSED_IDS');
  props.deleteProperty('DISMISSED_DATA');
  return 'cleared';
}

function restoreEmail(messageId) {
  const props = getUserProps_();
  const ids = getDismissedIds().filter(function(id) { return id !== messageId; });
  props.setProperty('DISMISSED_IDS', JSON.stringify(ids));

  const data = parseJson_(props.getProperty('DISMISSED_DATA'), {});
  delete data[messageId];
  props.setProperty('DISMISSED_DATA', JSON.stringify(data));

  try {
    const thread = GmailApp.getMessageById(messageId).getThread();
    thread.moveToInbox();
    const label = GmailApp.getUserLabelByName(RESOLVED_LABEL_NAME);
    if (label) thread.removeLabel(label);
  } catch (e) {
    // If the message no longer exists, cleaning local state is still useful.
  }

  return 'restored';
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

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Inbox Triage')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function fetchAndTriageEmails() {
  const snapshot = fetchInboxSnapshot();
  return classifyInboxEmails(snapshot.emails);
}

function fetchInboxSnapshot() {
  requireApiKey_();

  const dismissedIds = new Set(getDismissedIds());
  const emails = getLatestInboxEmails_(20)
    .filter(function(email) { return !dismissedIds.has(email.id); })
    .slice(0, 20)
    .map(function(email) {
      return Object.assign({}, email, {
        category: 'low_priority',
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
  return {
    emails: enrichAttentionSignals_(attachFollowupCards_(classified)),
    smartLabelsEnabled: getSmartLabelsEnabled(),
  };
}

function applySmartLabelsForEmails(emails) {
  if (!getSmartLabelsEnabled()) return 'disabled';
  applyTriageLabels((emails || []).filter(function(email) { return !email.isFollowup; }));
  return 'applied';
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
  const msg = GmailApp.getMessageById(messageId);
  const thread = msg.getThread();
  const subject = thread.getFirstMessageSubject();
  const messages = thread.getMessages().slice(-4);

  const context = messages.map(function(message) {
    return {
      from: parseSender_(message.getFrom()).name,
      body: message.getPlainBody().substring(0, 400).replace(/\s+/g, ' ').trim(),
    };
  });

  const lastMsg = messages[messages.length - 1];
  const senderEmail = parseSender_(lastMsg.getFrom()).email;

  const draft = callClaude_(
    apiKey,
    'You are a professional email assistant. Write clear, concise, helpful email replies. Return ONLY the reply body - no subject, no signature.',
    'Draft a reply to this email thread.\n\nSubject: ' + subject + '\n\nThread:\n' +
      context.map(function(m, i) { return '[' + (i + 1) + '] From ' + m.from + ':\n' + m.body; }).join('\n\n') +
      '\n\nWrite a professional reply under 100 words.',
    400
  ).trim();

  if (!draft) throw new Error('Empty draft response.');
  return { draft, senderEmail, subject: /^re:/i.test(subject) ? subject : 'Re: ' + subject };
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

function checkCalendarAuth() {
  var token = ScriptApp.getOAuthToken();
  var res = UrlFetchApp.fetch(
    'https://www.googleapis.com/calendar/v3/users/me/calendarList?maxResults=1',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  var code = res.getResponseCode();
  if (code === 401 || code === 403) return { authorized: false };
  return { authorized: true };
}

function rsvpCalendarEvent(messageId, response) {
  if (response !== 'yes' && response !== 'no') throw new Error('Invalid response.');

  var msg = GmailApp.getMessageById(messageId);
  if (!msg) throw new Error('Message not found.');

  // Parse UID, title and start time from .ics — needed for both lookup strategies
  var uid = null, startDate = null, eventTitle = null;
  var attachments = msg.getAttachments();
  for (var i = 0; i < attachments.length; i++) {
    var att = attachments[i];
    if (att.getName().slice(-4).toLowerCase() !== '.ics' && att.getContentType() !== 'text/calendar') continue;
    // Unfold ICS lines (continuation lines start with a space/tab per RFC 5545)
    var ics = att.getDataAsString().replace(/\r?\n[ \t]/g, '');
    var veventBlock = ics.match(/BEGIN:VEVENT([\s\S]*?)END:VEVENT/);
    var evBlock = veventBlock ? veventBlock[1] : ics;

    var uidMatch = evBlock.match(/^UID:([^\r\n]+)/m);
    if (uidMatch) uid = uidMatch[1].trim();

    var summaryMatch = evBlock.match(/^SUMMARY:([^\r\n]+)/m);
    if (summaryMatch) eventTitle = summaryMatch[1].replace(/\\n/g, ' ').replace(/\\,/g, ',').trim();

    var dtstartMatch = evBlock.match(/DTSTART(?:;TZID=([^;:\r\n]+))?(?:;[^:\r\n]*)?:([^\r\n]+)/);
    if (dtstartMatch) {
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
    if (uid) break;
  }

  if (!uid) throw new Error('No UID found in calendar invite attachment.');

  var userEmail = Session.getActiveUser().getEmail();
  var responseStatus = response === 'yes' ? 'accepted' : 'declined';
  var guestStatus   = response === 'yes' ? CalendarApp.GuestStatus.YES : CalendarApp.GuestStatus.NO;
  var token   = ScriptApp.getOAuthToken();
  var headers = { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' };
  var rsvped  = false;

  // Strategy 1: Calendar REST API — exact UID match, most reliable
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
    // Organizer is automatically attending — no RSVP needed
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

  // Strategy 2: CalendarApp fallback — time-window search
  if (!rsvped && startDate) {
    var windowStart = new Date(startDate.getTime() - 60000);
    var windowEnd   = new Date(startDate.getTime() + 3600000);
    var allCals = CalendarApp.getAllCalendars();
    outer: for (var j = 0; j < allCals.length; j++) {
      var evts = allCals[j].getEvents(windowStart, windowEnd);
      for (var k = 0; k < evts.length; k++) {
        if (eventTitle && evts[k].getTitle() !== eventTitle) continue;
        var myStatus = evts[k].getMyStatus();
        // Organizer is automatically attending
        if (myStatus === CalendarApp.GuestStatus.OWNER) {
          rsvped = true;
          break outer;
        }
        try { evts[k].setMyStatus(guestStatus); rsvped = true; break outer; } catch(e) {
          console.log('CalendarApp setMyStatus failed: ' + e.message);
        }
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
      subject: '📬 Inbox Digest - ' + dateStr,
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
  const fyi = emails.filter(function(e) { return e.category === 'fyi'; }).slice(0, 5);

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
    Object.keys(CAT_COLORS).filter(function(k) { return k !== 'low_priority'; }).map(function(k) {
      return '<div style="flex:1;background:#1a1a1a;border:1px solid #2a2a2a;border-top:2px solid ' + CAT_COLORS[k] + ';border-radius:8px;padding:12px;text-align:center">' +
        '<div style="font-size:22px;font-weight:700;color:' + ((counts[k] || 0) > 0 ? CAT_COLORS[k] : '#444') + '">' + (counts[k] || 0) + '</div>' +
        '<div style="font-size:10px;color:#666;text-transform:uppercase;letter-spacing:.5px;margin-top:3px">' + k.replace(/_/g, ' ') + '</div></div>';
    }).join('') +
    '</div>' + section('Needs Attention', urgent) + section('Calendar', calendar) + section('FYI', fyi) +
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
    .inTimezone('America/Toronto')
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
      .setImageUrl('https://www.gstatic.com/images/branding/product/2x/gmail_2020q4_48dp.png')
      .setImageStyle(CardService.ImageStyle.CIRCLE))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newDecoratedText()
        .setText('Classifies your 20 most recent emails by urgency, category, and required action.')
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
      return {
        id: email.id,
        subject: email.subject,
        sender_name: email.sender_name,
        date: email.date,
      };
    });

    const SYSTEM_PROMPT =
      'You are a JSON-only API. Respond with a valid JSON array only. No markdown, no prose, no code fences. Start with [ and end with ].';

    const USER_PROMPT =
      'Classify each email into exactly one category using this strict decision hierarchy (evaluate top-to-bottom; use the first that fits):\n\n' +
      '1. escalation — C-level/VP/director is escalating an issue; production outage or system-down incident; SLA breach; named-account customer escalation; or any explicit use of words like "escalate", "critical", "sev0", "sev1", "outage", "blocker". Priority 4-5.\n' +
      '2. action_required — Sender explicitly asks the recipient to: reply with a decision, approve or reject something, complete a specific task, submit or send something, fix a bug, or take a named action. Keywords: "can you", "please", "need you to", "action required", "approval needed", "please approve", "sign off", "review and respond". Priority 2-4.\n' +
      '3. calendar — Email is a meeting invite, scheduling request, calendar event notification, or a direct request to book or reschedule time. Priority 2-3.\n' +
      '4. fyi — Informational only: status update, summary, report, announcement, or notification where no reply or action is expected. Priority 1-2.\n' +
      '5. low_priority — Everything else: newsletters, marketing, automated system notifications, receipts, social alerts, or anything that does not fit categories 1-4. Priority 1.\n\n' +
      'Return ONLY a JSON array. Each object: id (exact match), category, priority (1-5 integer), summary (max 12 words, factual), attentionSignals ({directAsk:boolean, deadlineDriven:boolean, criticalEscalation:boolean}).\n' +
      'directAsk = category is action_required or escalation.\n' +
      'deadlineDriven = email mentions today, tomorrow, EOD, ASAP, deadline, due date, overdue, or time-sensitive urgency.\n' +
      'criticalEscalation = category is escalation.\n\n' +
      JSON.stringify(payload);

    const text = callClaude_(apiKey, SYSTEM_PROMPT, USER_PROMPT, 2000);
    const classified = extractJsonArray_(text);

    classified.forEach(function(found) {
      if (!found.id) return;
      cache[found.id] = {
        category: normalizeCategory_(found.category || 'low_priority'),
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
      category: normalizeCategory_(found.category || 'low_priority'),
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
