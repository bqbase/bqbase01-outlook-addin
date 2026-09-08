/*
============================================================================
 taskpane.js -- main orchestrator
============================================================================

 Wires together mail-context.js (read the email), providers.js (the streaming
 OpenRouter call), and insert.js (write the draft back) into the actual chat
 UI: a history list, streaming deltas rendered live, INSERT enabled only
 after a successful draft, and auto-suggest when a reply is opened.
============================================================================
*/

let currentItem = null;
let history = []; // [{role, content}]
let lastDraft = null; // last successful draft, enables INSERT
let busy = false;
let pendingAttachments = []; // Attachment[] from attachments.js

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;
  currentItem = Office.context.mailbox.item;

  document.getElementById("screen-loading").classList.add("hidden");

  document.getElementById("insertBtn").addEventListener("click", onInsert);
  document.getElementById("input-box").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      onSend();
    }
  });
  document.getElementById("browseBtn").addEventListener("click", () => {
    document.getElementById("fileInput").click();
  });
  document.getElementById("fileInput").addEventListener("change", async (e) => {
    await addFiles(e.target.files);
    e.target.value = ""; // allow re-selecting the same file later
  });
  _wireDropTarget(document.getElementById("input-box"));
  // The transcript is ALSO a drop target, not just the input box -- a user
  // dragging a file toward the input box has a decent chance of releasing
  // it slightly high, over the transcript instead. Both route to the same
  // addFiles handler, matching how forgiving a real drop zone should be.
  _wireDropTarget(document.getElementById("transcript"));

  showChatScreen();
  maybeAutoSuggestReplyOptions();
});

// -- screen management --------------------------------------------------

function showChatScreen() {
  document.getElementById("screen-chat").classList.remove("hidden");
  updateStatusLine();
}

// -- chat ------------------------------------------------------------------

function updateStatusLine() {
  describeContext(currentItem).then((text) => {
    document.getElementById("status-line").textContent = text;
  });
}

function appendTurn(text, cssClass) {
  const transcript = document.getElementById("transcript");
  const div = document.createElement("div");
  div.className = "turn " + cssClass;
  div.textContent = text;
  transcript.appendChild(div);
  transcript.scrollTop = transcript.scrollHeight;
  return div;
}

// Disabling the input box is the whole busy guard now -- there is no submit
// button to disable, since Enter is the only way to submit.
function setBusy(isBusy) {
  busy = isBusy;
  document.getElementById("input-box").disabled = isBusy;
}

// -- attachments -------------------------------------------------------

function _wireDropTarget(el) {
  // dragover must call preventDefault(), or the browser's default
  // "not-a-drop-target" behavior wins and drop never fires -- standard
  // HTML5 Drag and Drop API requirement, not optional.
  el.addEventListener("dragover", (e) => {
    e.preventDefault();
    el.classList.add("drop-target-active");
  });
  el.addEventListener("dragleave", () => {
    el.classList.remove("drop-target-active");
  });
  el.addEventListener("drop", async (e) => {
    e.preventDefault();
    el.classList.remove("drop-target-active");
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
      await addFiles(e.dataTransfer.files);
    }
  });
}

async function addFiles(fileList) {
  if (busy) return; // busy guard on attach -- deliberately a silent no-op
  const room = MAX_ATTACHMENTS_PER_MESSAGE - pendingAttachments.length;
  if (room <= 0) return;
  const processed = await processFiles(Array.from(fileList).slice(0, room));
  pendingAttachments = pendingAttachments.concat(processed);
  renderAttachmentChips();
}

function removeAttachment(index) {
  if (busy) return;
  pendingAttachments.splice(index, 1);
  renderAttachmentChips();
}

function renderAttachmentChips() {
  const row = document.getElementById("chips-row");
  row.innerHTML = "";
  pendingAttachments.forEach((att, index) => {
    const chip = document.createElement("span");
    chip.className = "chip" + (att.error ? " chip-error" : "");
    chip.title = att.error || att.filename;
    const label = document.createElement("span");
    label.textContent = att.filename;
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.textContent = "×"; // ×
    removeBtn.addEventListener("click", () => removeAttachment(index));
    chip.appendChild(label);
    chip.appendChild(removeBtn);
    row.appendChild(chip);
  });
}

async function onSend() {
  if (busy) return;
  const inputBox = document.getElementById("input-box");
  const text = inputBox.value.trim();
  // A message with no typed text but pending attachments is valid --
  // e.g. drop a PDF and press Enter with nothing typed.
  if (!text && pendingAttachments.length === 0) return;

  let display = "You: " + (text || "(no message)");
  const okAttachments = pendingAttachments.filter((a) => !a.error);
  if (okAttachments.length) {
    display += "\n[Attached: " + okAttachments.map((a) => a.filename).join(", ") + "]";
  }
  appendTurn(display, "user");

  inputBox.value = "";
  const sentAttachments = pendingAttachments;
  pendingAttachments = [];
  renderAttachmentChips();
  await dispatchTurn(text, true, sentAttachments);
}

// Fires once on a reply/forward: a synthetic first turn, not shown as
// "You: ...". Auto-reading the ORIGINAL message's own attachments is not
// implemented -- see mail-context.js's header. This suggests reply angles
// from plain text context only. Manually drag-and-dropped files (this
// file's own addFiles/attachments.js) are a separate, supported feature.
async function maybeAutoSuggestReplyOptions() {
  if (!isReplyOrForward(currentItem)) return;
  appendTurn("[system] Reviewing the email -- suggesting reply options...", "system");
  await dispatchTurn(
    "Before drafting anything, suggest 2-3 different brief angles or " +
    "approaches I could take in replying to this email (one sentence " +
    "each). Don't write a full draft yet -- just the options.",
    false
  );
}

// A neutral label for the model's turns. It read "Claude:" until 2026-09-08
// while every call went to openai/gpt-5.6-luna, then briefly showed that
// model id -- accurate, but long and noisy sitting next to "You:".
// "Assistant:" stays true whichever model providers.js is pointed at, so it
// cannot go stale the way the old hardcoded name did.
const ASSISTANT_LABEL = "Assistant:\n";

// Renders the small subset of Markdown the model actually emits --
// **bold**, *italic*, `code` -- as real elements.
//
// NOTHING HERE USES innerHTML, deliberately. Model output is derived in part
// from an email the user RECEIVED, so treating it as markup would let a
// hostile or malformed message inject into the pane. Every piece of that
// text is placed with textContent or createTextNode and only strong/em/code
// elements are ever created, so the formatting is gained without giving the
// text any power to become markup.
//
// This affects the TRANSCRIPT ONLY. The draft that goes into the email is
// still inserted verbatim by insert.js, so Markdown in a draft body would
// reach the recipient as literal asterisks -- a separate, still-open issue.
//
// Kept identical to chrome_extension/sidepanel.js. Change both together.
// The fourth alternative is a line-leading ordered-list marker ("1.", "2."),
// which is not Markdown emphasis at all but is treated as bold here so a
// numbered list of reply options reads as a list. The "m" flag is what makes
// "^" match at every line start rather than only the first.
const INLINE_MARKDOWN = /\*\*([^*]+)\*\*|\*([^*\n]+)\*|`([^`\n]+)`|^(\d+\.)(?=\s)/gm;

function _markdownNodes(text) {
  const nodes = [];
  let last = 0;
  let match;
  INLINE_MARKDOWN.lastIndex = 0; // the regex is global and reused across turns
  while ((match = INLINE_MARKDOWN.exec(text)) !== null) {
    if (match.index > last) {
      nodes.push(document.createTextNode(text.slice(last, match.index)));
    }
    let tag = "strong";
    let inner = match[1];
    if (match[2] !== undefined) {
      tag = "em";
      inner = match[2];
    } else if (match[3] !== undefined) {
      tag = "code";
      inner = match[3];
    } else if (match[4] !== undefined) {
      inner = match[4]; // list marker -- stays "strong", styled with the label
    }
    const el = document.createElement(tag);
    el.textContent = inner; // textContent, never innerHTML -- see above
    nodes.push(el);
    last = match.index + match[0].length;
  }
  if (last < text.length) nodes.push(document.createTextNode(text.slice(last)));
  return nodes;
}

function renderAssistantTurn(el, text) {
  el.textContent = "";
  el.appendChild(document.createTextNode(ASSISTANT_LABEL));
  for (const node of _markdownNodes(text)) el.appendChild(node);
}

async function dispatchTurn(text, isDraftCandidate, attachments) {
  const content = attachments && attachments.length ? buildMessageContent(text, attachments) : text;
  history.push({ role: "user", content: content });
  setBusy(true);

  const systemPrompt = await buildSystemPrompt(currentItem);

  const assistantDiv = appendTurn(ASSISTANT_LABEL, "assistant");
  let streamedText = "";
  const onDelta = (chunk) => {
    streamedText += chunk;
    renderAssistantTurn(assistantDiv, streamedText);
    const transcript = document.getElementById("transcript");
    transcript.scrollTop = transcript.scrollHeight;
  };

  const result = await sendTurnStreaming(systemPrompt, history, onDelta);

  setBusy(false);

  if (result.ok) {
    history.push({ role: "assistant", content: result.text });
    // Default: re-render the final text from the authoritative result,
    // not the accumulated deltas -- providers.js's contract is that the
    // terminal result is authoritative, not the streamed pieces.
    let displayText = result.text;
    if (isDraftCandidate) {
      lastDraft = result.text;
      document.getElementById("insertBtn").disabled = false;
      // Swap the raw SUBJECT:/body/<<<END EMAIL>>> scaffolding for a
      // human-readable rendering. Confirmed live during testing that
      // without this, a real draft showed the literal "SUBJECT: ..." and
      // "<<<END EMAIL>>>" markers in the chat transcript even though the
      // actual INSERT (insert.js, which already calls parseDraft itself)
      // correctly stripped them from the compose body. A no-op if the
      // model's reply didn't use the contract -- parseDraft then returns
      // the text unchanged.
      const { proposedSubject, body } = parseDraft(result.text);
      if (proposedSubject !== null || body !== result.text) {
        displayText = proposedSubject ? "Subject: " + proposedSubject + "\n\n" + body : body;
      }
    }
    renderAssistantTurn(assistantDiv, displayText);
  } else {
    // Roll back the just-pushed user turn on failure -- a retry
    // shouldn't desync history with a dangling unanswered user turn.
    history.pop();
    assistantDiv.remove();
    const message = result.refusalCategory
      ? "The model declined to respond (" + result.refusalCategory + ")."
      : result.error;
    appendTurn(message, "error");
  }
}

async function onInsert() {
  if (!lastDraft) return;
  const result = await insertDraft(currentItem, lastDraft);
  appendTurn("[system] " + result.message, result.ok ? "system" : "error");
}
