/**
 * updateWorkOrderSubjects.gs
 *
 * Searches Gmail for matching Work Order messages and creates a draft
 * (to the original sender) with an updated subject, or optionally forwards
 * automatically. It marks processed messages using script properties and a label
 * so it won't reprocess messages.
 *
 * IMPORTANT:
 * - You cannot change the subject of an existing received message in Gmail.
 *   This script creates a draft or forwards the message with the new subject.
 * - Test with a conservative SEARCH_QUERY to avoid mass actions.
 *
 * Configuration (edit as needed):
 * - SEARCH_QUERY: Gmail search string to find relevant messages.
 * - REGEX_PATTERN: JS regex (string) to match in subject.
 * - REPLACEMENT: subject replacement (JS replacement string).
 * - LABEL_NAME: label added to processed threads.
 * - SEND_FORWARD: false = create a draft (to original sender); true = forward automatically to original recipients.
 *
 * Usage:
 * - Paste into Apps Script editor (https://script.google.com/) as a new project/file.
 * - Run 'processWorkOrderEmails' manually to test.
 * - If OK, add an installable trigger (time-driven) to run on schedule.
 */

const SEARCH_QUERY = 'subject:"Work Order" newer_than:30d'; // adjust as needed
const REGEX_PATTERN = '\\bWork Order\\b';                  // string for RegExp
const REPLACEMENT = 'Work Order [UPDATED]';               // replacement for subject
const LABEL_NAME = 'workorder-subject-updated';
const SEND_FORWARD = false; // false => create draft to original sender; true => forward automatically

function processWorkOrderEmails() {
  const label = getOrCreateLabel(LABEL_NAME);
  const query = `${SEARCH_QUERY} -label:${LABEL_NAME}`;
  const threads = GmailApp.search(query, 0, 100); // process up to 100 matching threads per run
  if (!threads || threads.length === 0) {
    Logger.log('No threads found for query: %s', query);
    return;
  }

  Logger.log('Found %s threads', threads.length);
  const regex = new RegExp(REGEX_PATTERN, 'i');
  for (const thread of threads) {
    const messages = thread.getMessages();
    let threadTouched = false;
    for (const message of messages) {
      try {
        const messageId = message.getId();
        if (isProcessed(messageId)) {
          continue;
        }
        const subject = message.getSubject() || '';
        if (!regex.test(subject)) {
          markProcessed(messageId); // mark as processed so we don't revisit it later
          continue;
        }

        const newSubject = subject.replace(regex, REPLACEMENT);
        const bodyHtml = buildForwardHtmlBody(message);
        const bodyPlain = message.getPlainBody() || '';

        if (SEND_FORWARD) {
          // Forward automatically to original "To" recipients
          const originalTo = message.getTo();
          if (!originalTo) {
            Logger.log('Skipping messageId=%s: no "To" recipients found for forwarding', messageId);
            markProcessed(messageId);
            continue;
          }
          Logger.log('Forwarding messageId=%s to=%s newSubject=%s', messageId, originalTo, newSubject);
          // forward(recipient, options) sends immediately
          message.forward(originalTo, { htmlBody: bodyHtml, subject: newSubject });
        } else {
          // Create a draft addressed to the original sender for review
          const recipient = extractSingleAddress(message.getFrom()) || Session.getActiveUser().getEmail();
          Logger.log('Creating draft to=%s newSubject=%s', recipient, newSubject);
          GmailApp.createDraft(recipient, newSubject, bodyPlain, { htmlBody: bodyHtml });
        }

        markProcessed(messageId);
        threadTouched = true;
      } catch (err) {
        Logger.log('Error processing message: %s', err && err.toString());
      }
    }

    if (threadTouched) {
      label.addToThread(thread);
    }
  }

  Logger.log('Processing complete.');
}

/* Helper: create a simple HTML body that preserves the original message with a header.
 * For security, we escape the original email body content using the escapeHtml function
 * to prevent potential XSS or HTML injection in the draft/forwarded message.
 */
function buildForwardHtmlBody(message) {
  const from = message.getFrom() || '';
  const date = message.getDate() ? message.getDate().toString() : '';
  const subject = message.getSubject() || '';
  // Use plain body and escape it for safer HTML output
  const originalBody = message.getPlainBody() || '';
  // Include original metadata and message in the draft/forward html
  return '<p><strong>Original message</strong></p>'
    + `<p>From: ${escapeHtml(from)}<br/>Date: ${escapeHtml(date)}<br/>Subject: ${escapeHtml(subject)}</p>`
    + '<hr/><pre style="white-space: pre-wrap; font-family: inherit;">' + escapeHtml(originalBody) + '</pre>';
}

/* Helper: basic HTML escape */
function escapeHtml(text) {
  if (!text) return '';
  return text.replace(/&/g, '&amp;')
             .replace(/</g, '&lt;')
             .replace(/>/g, '&gt;')
             .replace(/"/g, '&quot;')
             .replace(/'/g, '&#39;');
}

/* Helper: ensure label exists */
function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

/* Persist processed message IDs to avoid duplicate work */
function isProcessed(messageId) {
  if (!messageId) return false;
  const key = 'processed_' + messageId;
  const val = PropertiesService.getScriptProperties().getProperty(key);
  return val === '1';
}
function markProcessed(messageId) {
  if (!messageId) return;
  const key = 'processed_' + messageId;
  PropertiesService.getScriptProperties().setProperty(key, '1');
}

/* Try to extract a single email address from a "From:" value (simple) */
function extractSingleAddress(fromHeader) {
  if (!fromHeader) return null;
  // examples: "Foo Bar <foo@example.com>" or "foo@example.com"
  const m = fromHeader.match(/<([^>]+)>/);
  if (m && m[1]) return m[1];
  // fallback to first token that looks like an email
  const m2 = fromHeader.match(/([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/i);
  return m2 ? m2[1] : null;
}

/* Admin helper to clear processed markers (for testing) */
function clearProcessedMarkers() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const keys = Object.keys(all).filter(k => k.startsWith('processed_'));
  for (const k of keys) props.deleteProperty(k);
  Logger.log('Cleared %s processed markers', keys.length);
}
