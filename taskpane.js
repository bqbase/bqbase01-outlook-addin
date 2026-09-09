/*
============================================================================
 taskpane.js -- main orchestrator
============================================================================

 Wires together mail-context.js (read the email), providers.js (the streaming
 model call), and insert.js (write the draft back) into the actual chat
 UI: a history list, streaming deltas rendered live, INSERT enabled only
 after a successful draft, and auto-suggest when a reply is opened.
============================================================================
*/

let currentItem = null;
let history = []; // [{role, content}]
let lastDraft = null; // last successful draft, enables INSERT
let busy = false;
let pendingAttachments = []; // Attachment[] from attachments.js
let boundEmail = ""; // mailbox the service says this token belongs to

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) return;
  currentItem = Office.context.mailbox.item;

  // Must happen before ANY call to the service, including the token check on
  // the settings screen: the subscription is tied to this mailbox, and a call
  // that omits the address cannot bind a fresh token to it.
  setMailbox(_mailboxAddress());

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

  document.getElementById("settingsBtn").addEventListener("click", () => showSettings(true));
  document.getElementById("saveTokenBtn").addEventListener("click", onSaveToken);
  document.getElementById("cancelTokenBtn").addEventListener("click", () => {
    showChatScreen();
  });
  document.getElementById("reviewAttachBtn").addEventListener("click", onReviewAttachments);

  // A customer with no token has nothing to look at, so send them straight to
  // the one field that matters rather than a chat screen that cannot work.
  if (!getToken()) {
    showSettings(false);
    return;
  }
  showChatScreen();
  await refreshCoinBar();
  maybeAutoSuggestReplyOptions();
});

// The SIGNED-IN USER's address, which is deliberately not "the mailbox
// currently open": on a shared or delegated mailbox userProfile reports the
// person, and the licence should follow the person who bought it.
//
// Wrapped because it is absent on some hosts. Empty means the Worker cannot
// check the binding and lets the call through on the token alone.
function _mailboxAddress() {
  try {
    const profile = Office.context.mailbox.userProfile;
    return (profile && profile.emailAddress) || "";
  } catch (err) {
    return "";
  }
}

// -- settings ------------------------------------------------------------

// canCancel is false on first run: there is no chat screen to go back to yet.
function showSettings(canCancel) {
  document.getElementById("screen-chat").classList.add("hidden");
  document.getElementById("screen-settings").classList.remove("hidden");
  document.getElementById("token-input").value = getToken();
  document.getElementById("settings-error").classList.add("hidden");
  document.getElementById("cancelTokenBtn").classList.toggle("hidden", !canCancel);
  renderMailboxLine();
}

// Which mailbox this pane is signed in as, and -- once the service has
// answered -- which mailbox the token is registered to. The two differ only
// when a token has been pasted into the wrong account, which is worth showing
// here rather than leaving the customer to discover it as a refusal later.
function renderMailboxLine() {
  const el = document.getElementById("settings-mailbox");
  const here = _mailboxAddress();
  const parts = [];
  if (here) parts.push("This mailbox: " + here);
  if (boundEmail && boundEmail.toLowerCase() !== here.toLowerCase()) {
    parts.push("Token registered to: " + boundEmail);
  }
  el.textContent = parts.join(" · ");
  el.classList.toggle("hidden", parts.length === 0);
}

// The token is VERIFIED against the service before it is accepted, so a typo
// is caught here rather than surfacing as a confusing failure on the first
// real request.
async function onSaveToken() {
  const field = document.getElementById("token-input");
  const errorBox = document.getElementById("settings-error");
  const button = document.getElementById("saveTokenBtn");
  const value = field.value.trim();
  if (!value) {
    errorBox.textContent = "Paste your access token first.";
    errorBox.classList.remove("hidden");
    return;
  }

  button.disabled = true;
  const previous = getToken();
  setToken(value);
  const result = await fetchBalance();
  button.disabled = false;

  if (!result.ok) {
    setToken(previous); // do not leave a bad token stored
    errorBox.textContent = result.error;
    errorBox.classList.remove("hidden");
    return;
  }

  showChatScreen();
  renderCoinBar(result.balance);
  maybeAutoSuggestReplyOptions();
}

// -- the coin counter ----------------------------------------------------

function renderCoinBar(balance) {
  const el = document.getElementById("coin-count");
  if (!balance) {
    el.textContent = "";
    return;
  }
  if (balance.email) boundEmail = balance.email; // for the settings screen
  const left = balance.coins_left;
  el.textContent = left + (left === 1 ? " coin" : " coins") +
    " · resets " + _shortDate(balance.resets_at);
  el.classList.toggle("negative", left < 0);
}

// Updated from the headers that ride back on every turn, which avoids a
// second round trip to /balance after each reply.
function updateCoinBarFromTurn(coins) {
  if (!coins || coins.left === null) return;
  const el = document.getElementById("coin-count");
  const resets = coins.resets ? " · resets " + _shortDate(coins.resets) : "";
  el.textContent = coins.left + (coins.left === 1 ? " coin" : " coins") + resets;
  el.classList.toggle("negative", coins.left < 0);
}

async function refreshCoinBar() {
  const result = await fetchBalance();
  if (result.ok) renderCoinBar(result.balance);
}

// "2026-10-08" -> "8 Oct". Parsed as parts rather than through Date, which
// would read a bare YYYY-MM-DD as UTC and can show the previous day.
function _shortDate(iso) {
  if (!iso) return "";
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const [, m, d] = iso.split("-").map(Number);
  return d + " " + (months[m - 1] || "");
}

// -- screen management --------------------------------------------------

function showChatScreen() {
  document.getElementById("screen-settings").classList.add("hidden");
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
  document.getElementById("reviewAttachBtn").disabled = isBusy;
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
  renderAttachRow();
}

// Attaching a file costs NOTHING. Reviewing it is a separate, priced action,
// and the price is shown before the button is pressed -- the customer is
// charged exactly this number, never a different one worked out afterwards.
function renderAttachRow() {
  const row = document.getElementById("attach-row");
  const usable = pendingAttachments.filter((a) => !a.error);
  if (usable.length === 0) {
    row.classList.add("hidden");
    return;
  }
  const coins = estimateCoins(usable);
  document.getElementById("attach-estimate").textContent =
    coins + (coins === 1 ? " coin" : " coins");
  document.getElementById("reviewAttachBtn").textContent =
    usable.length === 1 ? "Review attachment" : "Review " + usable.length + " attachments";
  row.classList.remove("hidden");
}

async function onSend() {
  if (busy) return;
  const inputBox = document.getElementById("input-box");
  const text = inputBox.value.trim();
  if (!text) return; // attachments have their own button now

  appendTurn("You: " + text, "user");
  inputBox.value = "";
  // Attachments deliberately stay put. A reply is one coin; reviewing a
  // document is a separate action the customer chooses and is quoted for.
  await dispatchTurn(text, true, null, "reply");
}

// The priced action. Sends the attachments with whatever instruction is in
// the box, then clears them so the same document cannot be silently charged
// for twice.
async function onReviewAttachments() {
  if (busy) return;
  const usable = pendingAttachments.filter((a) => !a.error);
  if (usable.length === 0) return;

  const inputBox = document.getElementById("input-box");
  const text = inputBox.value.trim();
  const coins = estimateCoins(usable);

  let line = "You: [Reviewing " + usable.map((a) => a.filename).join(", ") +
    " -- " + coins + (coins === 1 ? " coin" : " coins") + "]";
  if (text) line += "\n" + text;
  appendTurn(line, "user");

  inputBox.value = "";
  pendingAttachments = [];
  renderAttachmentChips();
  await dispatchTurn(
    text || "Review the attached document and tell me what matters for my reply.",
    true, usable, "attachment"
  );
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
    false, null, "review"
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

// purpose decides the PRICE, not the behaviour: "review" is free, "reply" is
// one coin, "attachment" is the quoted estimate. The Worker charges it; this
// only declares which kind of call it is.
async function dispatchTurn(text, isDraftCandidate, attachments, purpose) {
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

  const result = await sendTurnStreaming(systemPrompt, history, onDelta, purpose);

  setBusy(false);
  // Updated even when the turn FAILED: a refusal may itself be the reason
  // (a suspended account), and a stale counter is worse than none.
  updateCoinBarFromTurn(result.coins);

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
