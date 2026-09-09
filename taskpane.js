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
// The credit limit from the last /balance. The per-turn coin headers do not
// carry it, so it is remembered here to keep the bar consistent between a
// full refresh and a turn update.
let lastFloor = null;
// The balance as the SERVICE last reported it. Only /balance writes this,
// never the per-turn header -- the header carries the charge the Worker
// intended before its meter decided whether to refund it.
let lastBalance = null;

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) return;
  currentItem = Office.context.mailbox.item;

  // Must happen before ANY call to the service, including the token check on
  // the settings screen: the subscription is registered to one mailbox, and a
  // call that omits the address cannot be checked against it.
  setMailbox(_mailboxAddress());

  // Before the first getToken(), or startup reads only this machine's copy
  // and sends a returning customer back to the settings screen.
  setTokenStore(_roamingTokenStore());

  document.getElementById("screen-loading").classList.add("hidden");

  document.getElementById("insertBtn").addEventListener("click", onInsert);
  document.getElementById("replyBtn").addEventListener("click", () => onReplyWithDraft(false));
  document.getElementById("replyAllBtn").addEventListener("click", () => onReplyWithDraft(true));
  configureDraftButtons();
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
  document.getElementById("cancelTokenBtn").addEventListener("click", async () => {
    showChatScreen();
    // Startup skips this when /balance fails, so backing out of the settings
    // screen is the customer's only route to it. Without this a transient
    // network blip removed the priced attachment row for the whole session.
    await loadItemAttachments();
  });
  document.getElementById("reviewAttachBtn").addEventListener("click", onReviewAttachments);

  // A customer with no token has nothing to look at, so send them straight to
  // the one field that matters rather than a chat screen that cannot work.
  if (!getToken()) {
    showSettings(false);
    return;
  }
  showChatScreen();
  const balance = await refreshCoinBar();
  // A token the service REJECTS used to leave a chat screen with an empty
  // coin bar and no explanation -- which is what a wrong, unregistered or
  // suspended account looks like from the customer's side. Send them to the
  // one screen that can say why, with the service's own wording.
  if (!balance.ok) {
    showSettings(Boolean(getToken()));
    document.getElementById("settings-error").textContent = balance.error;
    document.getElementById("settings-error").classList.remove("hidden");
    return;
  }
  // Only NOW is it safe to copy a machine-local token into the mailbox: the
  // service has confirmed the token is valid AND registered to this very
  // mailbox. Migrating at startup copied whatever token sat on this machine
  // into whichever mailbox happened to be open.
  _migrateTokenToMailbox();
  await loadItemAttachments();
  maybeAutoSuggestReplyOptions();
});

// Puts the open message's own attachments in front of the customer as a
// priced button. NOT reviewed automatically -- that would spend coins nobody
// asked to spend, and the whole point of the button is that the price is
// agreed before anything is charged.
//
// Present when READING a message and when FORWARDING one. Absent on Reply and
// Reply All, because Outlook itself drops the attachments from those drafts;
// see the note in attachments.js for why Office.js cannot go back for them.
// Guarded so it runs ONCE per pane. It is called from startup AND from
// onSaveToken, and re-opening settings to re-save a token used to append the
// message's attachments a second time -- the chips doubled, and so did the
// quote and the charge.
let itemAttachmentsLoaded = false;

async function loadItemAttachments() {
  if (itemAttachmentsLoaded) return;
  if (!attachmentApiAvailable(currentItem)) return;
  let entries;
  try {
    entries = describeItemAttachments(await listItemAttachments(currentItem));
  } catch (err) {
    return; // an unreadable attachment list must not stop the pane loading
  }
  // Set only once the list is actually in hand. Setting it before the try
  // meant one throw disabled the priced attachment path for the whole
  // session, with no way back -- including from the Cancel path that exists
  // precisely to recover from a failed start.
  itemAttachmentsLoaded = true;
  if (entries.length === 0) return;
  const room = MAX_ATTACHMENTS_PER_MESSAGE - pendingAttachments.length;
  const taken = entries.slice(0, Math.max(0, room));
  pendingAttachments = pendingAttachments.concat(taken);
  renderAttachmentChips();
  // An email with more attachments than the limit used to drop the surplus
  // without a word, so the customer would review 10 of 14 contracts believing
  // they had reviewed all of them. Said out loud instead.
  if (entries.length > taken.length) {
    appendTurn("[system] This email has " + entries.length + " attachments; only the first " +
      taken.length + " are listed. Remove some to reach the rest.", "system");
  }

  // Word, Excel, PowerPoint and text files are billed on the text INSIDE
  // them, which a compressed file's size does not predict. They are fetched
  // and extracted now so the quoted price is the real one. The row is already
  // on screen with the PDFs and images priced, and its total updates when
  // this finishes -- better than showing nothing while it runs.
  if (taken.some((a) => TEXT_PRODUCING.has(a.category))) {
    setAttachEstimatePending(true);
    try {
      await resolveTextAttachments(currentItem, taken);
    } finally {
      setAttachEstimatePending(false);
      renderAttachmentChips();
    }
  }
}

// A token store backed by the MAILBOX rather than by this machine, so one
// paste covers the desktop, the laptop and Outlook on the web.
//
// RoamingSettings is Mailbox 1.1, well under the 1.5 this add-in declares,
// but it is still feature-detected: a host that lacks it must fall back to
// localStorage rather than throw on startup.
function _roamingTokenStore() {
  try {
    const settings = Office.context.roamingSettings;
    if (!settings || typeof settings.get !== "function" ||
        typeof settings.set !== "function" || typeof settings.saveAsync !== "function") {
      return null;
    }
    return {
      load: () => settings.get(TOKEN_KEY) || "",
      save: (value) => {
        settings.set(TOKEN_KEY, value);
        // Fire and forget: the local copy has already been written, so a
        // failed save costs the customer one extra paste on their NEXT
        // device, not the use of the add-in here and now.
        settings.saveAsync(() => {});
      },
    };
  } catch (err) {
    return null;
  }
}

// Carries a token already pasted on THIS machine up to the mailbox, once.
// Without it, an existing customer would keep the per-machine behaviour
// forever and only new pastes would roam.
function _migrateTokenToMailbox() {
  try {
    // Only after the service has confirmed this token belongs to THIS
    // mailbox. boundEmail comes from /balance, so a token for someone else's
    // account is never written into this mailbox's roaming settings.
    const here = _mailboxAddress().toLowerCase();
    if (!here || !boundEmail || boundEmail.toLowerCase() !== here) return;
    const settings = Office.context.roamingSettings;
    if (!settings || typeof settings.get !== "function") return;
    if (settings.get(TOKEN_KEY)) return;                   // the mailbox already has one
    const local = localStorage.getItem(TOKEN_KEY);
    if (!local) return;
    settings.set(TOKEN_KEY, local);
    settings.saveAsync(() => {});
  } catch (err) {
    // Nothing to do -- the local copy still works on this machine.
  }
}

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
  // The same startup work Office.onReady does. Without this a customer who
  // has just pasted their token gets a chat screen with NO attachment row --
  // the priced review path is simply absent on first run, and only appears
  // if they happen to close and reopen the pane. Every new customer meets
  // that path first.
  await loadItemAttachments();
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
  if (typeof balance.coins_floor === "number") lastFloor = balance.coins_floor;
  lastBalance = balance.coins_left;
  const left = balance.coins_left;
  // The credit limit is shown ONLY once the balance is negative. Quoting it
  // to somebody comfortably in credit would advertise a wall they are
  // nowhere near; quoting it to somebody overdrawn is the one moment it
  // matters, because they can still act before they are stopped.
  const limit = typeof balance.coins_floor === "number" && left < 0
    ? " · limit " + balance.coins_floor
    : "";
  el.textContent = left + (left === 1 ? " coin" : " coins") +
    limit + " · resets " + _shortDate(balance.resets_at);
  el.classList.toggle("negative", left < 0);
  if (balance.blocked) {
    appendTurn("[system] This account has reached its credit limit and is paused. " +
      "Contact BQBase to settle up and continue.", "error");
  }
}

// Updated from the headers that ride back on every turn, which avoids a
// second round trip to /balance after each reply.
function updateCoinBarFromTurn(coins) {
  if (!coins || coins.left === null) return;
  const el = document.getElementById("coin-count");
  const resets = coins.resets ? " · resets " + _shortDate(coins.resets) : "";
  // The limit is shown once the balance goes negative, exactly as
  // renderCoinBar does. Without it the limit vanished the moment a turn
  // completed -- which is precisely when a customer heading for the floor
  // needs to see it. lastFloor is remembered from the last /balance, since
  // the per-turn headers do not carry it.
  const limit = typeof lastFloor === "number" && coins.left < 0
    ? " · limit " + lastFloor
    : "";
  el.textContent = coins.left + (coins.left === 1 ? " coin" : " coins") + limit + resets;
  el.classList.toggle("negative", coins.left < 0);
}

async function refreshCoinBar() {
  const result = await fetchBalance();
  if (result.ok) renderCoinBar(result.balance);
  return result;                                           // callers need the failure reason
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
  // Routed through renderAttachRow rather than assigned here, so every
  // reason the button should stay disabled is applied in ONE place: busy,
  // an extraction still running, and nothing ticked. Assigning it directly
  // re-enabled a button that had been disabled because no file was selected.
  renderAttachRow();
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
  const offered = Array.from(fileList);
  if (room <= 0) {
    appendTurn("[system] Already holding " + MAX_ATTACHMENTS_PER_MESSAGE +
      " files -- remove one before adding another.", "system");
    return;
  }
  if (offered.length > room) {
    appendTurn("[system] Only " + room + " of those " + offered.length +
      " files were added; " + MAX_ATTACHMENTS_PER_MESSAGE + " is the limit.", "system");
  }
  const processed = await processFiles(offered.slice(0, room));
  pendingAttachments = pendingAttachments.concat(processed);
  renderAttachmentChips();
}

function removeAttachment(index) {
  if (busy) return;
  pendingAttachments.splice(index, 1);
  renderAttachmentChips();
}

// The files that will actually be sent: usable, and ticked. Everything that
// prices or reviews goes through here, so the number quoted and the files
// reviewed can never drift apart from what the chips show.
function selectedAttachments() {
  return pendingAttachments.filter((a) => !a.error && a.selected !== false);
}

function renderAttachmentChips() {
  const row = document.getElementById("chips-row");
  row.innerHTML = "";
  pendingAttachments.forEach((att, index) => {
    const chip = document.createElement("label"); // label, so the text ticks it too
    chip.className = "chip" + (att.error ? " chip-error" : "") +
      (!att.error && att.selected === false ? " chip-off" : "");
    chip.title = att.error || att.filename;

    // A file that cannot be read gets no checkbox: there is nothing to
    // include, and an unticked box would imply it could be.
    if (!att.error) {
      const tick = document.createElement("input");
      tick.type = "checkbox";
      tick.checked = att.selected !== false;
      tick.addEventListener("change", () => {
        att.selected = tick.checked;
        renderAttachmentChips();
      });
      chip.appendChild(tick);
    }

    const label = document.createElement("span");
    label.textContent = att.filename;
    chip.appendChild(label);

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.textContent = "×"; // ×
    removeBtn.title = "Remove this file";
    // Stops the click reaching the surrounding label, which would otherwise
    // toggle the checkbox on its way out.
    removeBtn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      removeAttachment(index);
    });
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
  // With files present but all unticked the row STAYS, with the button dead.
  // Making it vanish would look like the attachments had been lost.
  const chosen = selectedAttachments();
  if (chosen.length === 0) {
    document.getElementById("attach-estimate").textContent = "tick a file to review";
    document.getElementById("reviewAttachBtn").textContent = "Review attachments";
    document.getElementById("reviewAttachBtn").disabled = true;
    row.classList.remove("hidden");
    return;
  }
  const coins = estimateCoins(chosen);
  // While an Office document is still being read, the total is not yet the
  // real one. Saying so beats showing a number that is about to change on its
  // own, which reads as the price moving after it was quoted.
  document.getElementById("attach-estimate").textContent = attachEstimatePending
    ? "reading documents..."
    : coins + (coins === 1 ? " coin" : " coins");
  // Says "3 of 5" only when some are unticked, so the common case where
  // everything is included stays uncluttered.
  const label = chosen.length === 1 ? "Review attachment"
    : chosen.length === usable.length ? "Review " + chosen.length + " attachments"
    : "Review " + chosen.length + " of " + usable.length + " attachments";
  document.getElementById("reviewAttachBtn").textContent = label;
  document.getElementById("reviewAttachBtn").disabled = attachEstimatePending || busy;
  row.classList.remove("hidden");
}

// True while Office documents are being fetched and unzipped to price them.
let attachEstimatePending = false;

function setAttachEstimatePending(pending) {
  attachEstimatePending = pending;
  renderAttachRow();
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
  const usable = selectedAttachments();
  if (usable.length === 0) return;

  const inputBox = document.getElementById("input-box");
  const text = inputBox.value.trim();
  const coins = estimateCoins(usable);

  let line = "You: [Reviewing " + usable.map((a) => a.filename).join(", ") +
    " -- " + coins + (coins === 1 ? " coin" : " coins") + "]";
  if (text) line += "\n" + text;
  appendTurn(line, "user");

  inputBox.value = "";
  // Only the files being reviewed leave the row. Anything left unticked stays
  // put, so a customer who reviews two of five contracts now can review the
  // rest afterwards without hunting the email down again.
  pendingAttachments = pendingAttachments.filter((a) => !usable.includes(a));
  renderAttachmentChips();

  // Files already on the message arrive as metadata only -- their bytes are
  // fetched here, AFTER the price has been quoted and agreed. Dropped files
  // were read when they were dropped and pass through untouched.
  setBusy(true);
  for (const entry of usable) {
    await fetchItemAttachment(currentItem, entry);
  }
  setBusy(false);

  // Office refuses to release some attachments, and it says so only now. A
  // file that failed is NOT sent, so the Worker never sees it and never
  // charges for it -- the customer pays less than the quote, never more.
  const ready = usable.filter((a) => !a.error);
  const failed = usable.filter((a) => a.error);
  if (failed.length) {
    appendTurn("Could not read: " +
      failed.map((a) => a.filename + " -- " + a.error).join("; "), "error");
  }
  if (ready.length === 0) {
    appendTurn("Nothing was reviewed, so no coins were used.", "error");
    return;
  }

  // With nothing typed, a review ends in three fresh reply options rather
  // than a summary: the point of reading the attachments is to reply better,
  // and the options the pane suggested when the email opened were written
  // without any of this. Mirrors the wording of maybeAutoSuggestReplyOptions
  // so the two read as the same feature at different moments.
  //
  // Anything TYPED wins. Someone who asks "does this contract mention a
  // penalty?" wants that answered, not three ways to reply.
  // NOT a draft candidate when it is going to answer with three options:
  // isDraftCandidate stores the reply as lastDraft and lights up Insert /
  // Reply, so leaving it true would let the customer paste a numbered list
  // of options into the email as though it were the reply.
  const wantsOptions = !text;
  const ok = await dispatchTurn(
    text || REVIEW_THEN_OPTIONS_PROMPT,
    !wantsOptions, ready, "attachment"
  );
  if (!ok) {
    // The files go back in the row rather than making the customer find the
    // email again.
    pendingAttachments = ready.concat(pendingAttachments);
    renderAttachmentChips();
    // NO NUMBER IS CLAIMED for what the lost turn cost, because the pane
    // cannot know it. Two attempts got this wrong: reading
    // X-BQBase-Coins-Charged reported a charge the Worker's meter then
    // refunded, and diffing the balance across a refresh blamed the review
    // for every coin spent since the last /balance -- the per-turn header
    // updates the bar but deliberately does not update that baseline. The
    // charge is also applied asynchronously in waitUntil, so an immediate
    // refresh can miss it entirely.
    //
    // So: refresh the counter, state it, and say plainly what a retry may
    // cost. A number the customer can check beats a claim they cannot.
    // The refresh's OWN result decides whether a number may be quoted. It
    // fails on exactly the network drop that killed the turn, and
    // lastBalance is only ever written by a SUCCESSFUL /balance -- so
    // quoting it regardless announced a figure from several charged turns
    // ago as "now", in the same sentence that told the customer to check it.
    const refreshed = await refreshCoinBar();
    const now = refreshed.ok && typeof lastBalance === "number"
      ? " Your balance is now " + lastBalance + (lastBalance === 1 ? " coin." : " coins.")
      : " The coin count could not be refreshed, so reopen the pane to see it.";
    appendTurn("[system] The answer did not arrive and the attachments are still" +
      " listed." + now + " A review that reached the model is charged even when" +
      " the answer is lost, so check the count before pressing Review again.",
      "error");
  }
}

const REVIEW_THEN_OPTIONS_PROMPT =
  "Read the attached file(s). Do NOT write a reply yet. First give me a " +
  "short account of what is in them that actually bears on how I answer " +
  "-- figures, dates, obligations, anything that changes the picture. " +
  "Then, combining that with the email itself, suggest 3 different " +
  "approaches I could take in replying (one sentence each), numbered 1, 2 " +
  "and 3. If the attachments change what a sensible reply looks like " +
  "compared with the email alone, say so.";

// Fires once on a reply/forward: a synthetic first turn, not shown as
// "You: ...". Auto-reading the ORIGINAL message's own attachments is not
// implemented -- see mail-context.js's header. This suggests reply angles
// from plain text context only. Manually drag-and-dropped files (this
// file's own addFiles/attachments.js) are a separate, supported feature.
let autoSuggested = false;

async function maybeAutoSuggestReplyOptions() {
  if (!isReplyOrForward(currentItem)) return;
  // Once per pane. Startup and onSaveToken both call this and are not
  // concurrent, so the busy guard below does not stop a second, duplicate
  // set of suggestions being generated after a token is saved.
  if (autoSuggested) return;
  autoSuggested = true;
  // Guarded like every other entry point. Saving a token calls this while
  // the pane's own startup may already have a turn in flight, and two
  // concurrent dispatchTurn calls interleave their pushes into `history`
  // and fight over setBusy.
  if (busy) return;
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
      enableDraftButtons();
      // Swap the raw SUBJECT:/body/<<<END EMAIL>>> scaffolding for a
      // human-readable rendering. Confirmed live during testing that
      // without this, a real draft showed the literal "SUBJECT: ..." and
      // "<<<END EMAIL>>>" markers in the chat transcript even though the
      // actual INSERT (insert.js, which already calls parseDraft itself)
      // correctly stripped them from the compose body. A no-op if the
      // model's reply didn't use the contract -- parseDraft then returns
      // the text unchanged.
      const { proposedSubject, body, remarks } = parseDraft(result.text);
      if (proposedSubject !== null || body !== result.text) {
        displayText = proposedSubject ? "Subject: " + proposedSubject + "\n\n" + body : body;
        // The model's own remarks are KEPT, separated from the draft. The
        // system prompt reserves the space after <<<END EMAIL>>> for exactly
        // this, and dropping it made that channel write-only: a model saying
        // "you wrote 'next Tuesday'; I assumed the 12th, confirm before
        // sending" was talking to nobody, and the customer could send a
        // reply agreeing to the wrong date having never seen the question.
        // Display only -- Insert and Reply re-parse and take the body alone.
        if (remarks) displayText += "\n\n--\n" + remarks;
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
  return result.ok;
}

async function onInsert() {
  if (!lastDraft) return;
  const result = await insertDraft(currentItem, lastDraft);
  appendTurn("[system] " + result.message, result.ok ? "system" : "error");
}

// Read mode cannot insert -- there is no draft to insert INTO -- so it offers
// to open one already containing the draft. Compose mode keeps the
// cursor-insert, which is the only thing that preserves a quoted thread.
function configureDraftButtons() {
  const readMode = isReadMode(currentItem);
  document.getElementById("insertBtn").classList.toggle("hidden", readMode);
  document.getElementById("replyBtn").classList.toggle("hidden", !readMode);
  document.getElementById("replyAllBtn").classList.toggle("hidden", !readMode);
}

// Enables whichever pair is on screen. Kept in one place so a new draft can
// never light up a button that is not the one being shown.
function enableDraftButtons() {
  for (const id of ["insertBtn", "replyBtn", "replyAllBtn"]) {
    document.getElementById(id).disabled = false;
  }
}

function onReplyWithDraft(replyAll) {
  if (!lastDraft) return;
  const result = replyWithDraft(currentItem, lastDraft, replyAll);
  appendTurn("[system] " + result.message, result.ok ? "system" : "error");
}
