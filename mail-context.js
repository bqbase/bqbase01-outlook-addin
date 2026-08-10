/*
============================================================================
 mail-context.js -- mail context extraction for the task pane
============================================================================

 Mirrors V4's mail_context.py: without this, the add-in is a generic chatbot
 sitting next to an email it cannot see. Adapted for Office.js's async,
 promise-free callback API and the read/compose API-shape differences
 discovered live while building the feasibility prototype (see STATE.md):
 item.subject is a plain string in READ mode but a Subject object requiring
 .getAsync() in COMPOSE mode; item.body always requires .getAsync() in both
 modes.

 Attachments (V4's attach_from_mail_item/find_original_message feature) are
 NOT ported in this first pass -- Office.js's attachment content APIs are
 more restricted than COM's (getAttachmentContentAsync exists but reading an
 arbitrary original message's attachments the way V4 does via COM has no
 direct Office.js equivalent; would need its own design pass). Tracked as a
 known gap in STATE.md, not silently dropped.
============================================================================
*/

const BASE_SYSTEM_PROMPT =
  "You are an email writing assistant helping the user draft or refine an " +
  "email in Microsoft Outlook. Be concise, natural, and match a " +
  "professional but warm tone unless told otherwise. When the user asks " +
  "for a draft, write the full email body text they can use directly. " +
  "Hold the draft to the length the email actually needs -- do not pad it " +
  "with unnecessary sections or restate things the recipient already " +
  "knows. Write what the user asked for; do not widen the task into a " +
  "longer or more elaborate email than requested.\n\n" +
  "Whenever your reply IS a complete, ready-to-insert email draft (the " +
  "user can take it as-is and put it straight into the email body), " +
  "structure your ENTIRE response as exactly this, with nothing else " +
  "before, between, or after it:\n" +
  "SUBJECT: <the proposed subject line, or literally \"(no change)\" if " +
  "you are not proposing a subject change -- e.g. because this is a " +
  "reply/forward and the existing subject is already fine>\n\n" +
  "<the email body text -- nothing else: no preamble like \"Here's a " +
  "draft:\", no closing remarks, no explanations>\n" +
  "<<<END EMAIL>>>\n" +
  "Only write remarks, questions, or explanations for the user AFTER the " +
  "\"<<<END EMAIL>>>\" line -- never before it or inside the body. " +
  "For anything that is NOT a complete draft -- answering a question, " +
  "brainstorming options, asking for clarification -- just respond in " +
  "plain prose as normal; do not use this structure at all.";

// item.subject's shape differs by mode -- see this file's header. Returns a
// Promise<string> either way so callers never need to branch themselves.
function getSubject(item) {
  return new Promise((resolve) => {
    if (typeof item.subject === "string") {
      resolve(item.subject);
      return;
    }
    item.subject.getAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : "");
    });
  });
}

function getBodyText(item) {
  return new Promise((resolve) => {
    if (!item.body) {
      resolve("");
      return;
    }
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : "");
    });
  });
}

// Office.js's compose item has no direct equivalent of COM's
// ConversationIndex-length check V4 uses (mail_context.is_reply_or_forward).
// The closest documented signal is Office.context.mailbox.item.conversationId
// being non-null: a reply/forward inherits the original thread's ID
// immediately, per Microsoft's own docs. Known, accepted imprecision (same
// "degrades to a slightly-off prompt, not a crash" category as V4's own
// COM-check trade-off): the SAME docs also say a brand-new (non-reply)
// compose item gets a NON-null conversationId too, once the user has typed a
// subject AND the item has been saved as a draft -- so a saved blank
// "Untitled" compose that later gets a subject could false-positive as a
// reply. Narrow in practice (most brand-new compose windows are never
// explicitly saved before the user starts chatting), not chased further.
function isReplyOrForward(item) {
  return Boolean(item.conversationId);
}

async function buildMailContextSummary(item) {
  const subject = await getSubject(item);
  if (isReplyOrForward(item)) {
    const body = await getBodyText(item);
    return (
      "The user is replying to or forwarding an existing email thread.\n" +
      "Subject: " + subject + "\n" +
      "Current draft body (includes the quoted original message below " +
      "the reply point):\n" + body
    );
  }
  return "The user is composing a brand new email (not a reply).\nSubject so far: " + subject;
}

async function buildSystemPrompt(item) {
  if (!item) return BASE_SYSTEM_PROMPT;
  try {
    const context = await buildMailContextSummary(item);
    return BASE_SYSTEM_PROMPT + "\n\n" + context;
  } catch (err) {
    // A read failure must not crash the add-in -- fall back to a
    // context-free prompt and keep going, same as mail_context.py's
    // build_system_prompt.
    return BASE_SYSTEM_PROMPT;
  }
}

async function describeContext(item) {
  if (!item) return "No mail item attached -- generic assistant, no email context.";
  try {
    const subject = (await getSubject(item)) || "(no subject yet)";
    const kind = isReplyOrForward(item) ? "reply/forward" : "new message";
    return 'Context: ' + kind + ' -- "' + subject + '"';
  } catch (err) {
    return "Mail context unreadable -- continuing without it.";
  }
}

// Matches word_insert.py's SUBJECT:/body/<<<END EMAIL>>> contract -- kept in
// sync deliberately (same marker strings) so a draft this add-in produces
// parses identically to how V4's own INSERT logic parses it, in case the
// two are ever compared or a user works across both.
const DRAFT_PATTERN = /^SUBJECT:\s*(.*?)\s*\r?\n\r?\n([\s\S]*?)(?:\r?\n<<<END EMAIL>>>|$)/i;
const NO_CHANGE_SUBJECTS = new Set(["", "(no change)", "no change"]);

function parseDraft(rawText) {
  const match = DRAFT_PATTERN.exec(rawText.trim());
  if (!match) return { proposedSubject: null, body: rawText };
  const subject = match[1].trim();
  const proposedSubject = NO_CHANGE_SUBJECTS.has(subject.toLowerCase()) ? null : subject;
  return { proposedSubject, body: match[2].trim() };
}
