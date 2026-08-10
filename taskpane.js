/*
============================================================================
 taskpane.js -- main orchestrator
============================================================================

 Wires together mail-context.js (read the email), settings.js (per-user
 provider/key), providers.js (the streaming API call), and insert.js (write
 the draft back) into the actual chat UI. Mirrors main_app.py's AssistantWindow
 at the level of what it does (history list, streaming deltas rendered live,
 INSERT enabled only after a successful draft, auto-suggest on reply open) --
 not a line-for-line port, since Tkinter's widget model and a browser DOM
 have nothing in common mechanically.
============================================================================
*/

let currentItem = null;
let history = []; // [{role, content}], mirrors main_app.py's self.history
let lastDraft = null; // mirrors main_app.py's self.last_draft
let busy = false;
let settings = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;
  currentItem = Office.context.mailbox.item;
  settings = loadSettings();

  document.getElementById("screen-loading").classList.add("hidden");

  document.getElementById("settingsBtn").addEventListener("click", showSettingsScreen);
  document.getElementById("cancelSettingsBtn").addEventListener("click", () => {
    if (hasValidSettings(settings)) showChatScreen();
    // If settings were never configured, Cancel has nowhere useful to go
    // back to -- stay on the Settings screen rather than show a broken chat
    // screen with no key.
  });
  document.getElementById("saveSettingsBtn").addEventListener("click", onSaveSettings);
  document.getElementById("providerSelect").addEventListener("change", onProviderChange);
  document.getElementById("sendBtn").addEventListener("click", onSend);
  document.getElementById("insertBtn").addEventListener("click", onInsert);
  document.getElementById("input-box").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      onSend();
    }
  });

  if (hasValidSettings(settings)) {
    showChatScreen();
    maybeAutoSuggestReplyOptions();
  } else {
    // First-ever launch: no key configured yet -- go straight to Settings
    // rather than show a chat screen that can only ever error. Novice-
    // friendly per the user's explicit ask, not an afterthought.
    showSettingsScreen();
  }
});

// -- screen management --------------------------------------------------

function showChatScreen() {
  document.getElementById("screen-settings").classList.add("hidden");
  document.getElementById("screen-chat").classList.remove("hidden");
  updateStatusLine();
}

function showSettingsScreen() {
  document.getElementById("providerSelect").value = settings.provider;
  document.getElementById("apiKeyInput").value = settings.apiKey;
  document.getElementById("modelInput").value = settings.model;
  document.getElementById("effortSelect").value = settings.effort;
  document.getElementById("settings-error").textContent = "";
  onProviderChange();
  document.getElementById("screen-chat").classList.add("hidden");
  document.getElementById("screen-settings").classList.remove("hidden");
}

function onProviderChange() {
  const provider = document.getElementById("providerSelect").value;
  const modelInput = document.getElementById("modelInput");
  const modelNote = document.getElementById("modelNote");
  const effortField = document.getElementById("effortField");
  if (!modelInput.value) {
    modelInput.value = PROVIDERS[provider].defaultModel;
  }
  modelNote.textContent = "Suggested default: " + PROVIDERS[provider].defaultModel;
  effortField.style.display = PROVIDERS[provider].supportsEffort ? "" : "none";
}

async function onSaveSettings() {
  const newSettings = {
    provider: document.getElementById("providerSelect").value,
    apiKey: document.getElementById("apiKeyInput").value.trim(),
    model: document.getElementById("modelInput").value.trim(),
    effort: document.getElementById("effortSelect").value,
  };
  if (!newSettings.apiKey) {
    document.getElementById("settings-error").textContent = "API key is required.";
    return;
  }
  if (!newSettings.model) {
    document.getElementById("settings-error").textContent = "Model is required.";
    return;
  }
  try {
    await saveSettings(newSettings);
    settings = newSettings;
    showChatScreen();
  } catch (err) {
    document.getElementById("settings-error").textContent = "Could not save settings: " + err.message;
  }
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

function setBusy(isBusy) {
  busy = isBusy;
  document.getElementById("sendBtn").disabled = isBusy;
  document.getElementById("input-box").disabled = isBusy;
}

async function onSend() {
  if (busy) return;
  const inputBox = document.getElementById("input-box");
  const text = inputBox.value.trim();
  if (!text) return;
  inputBox.value = "";
  appendTurn("You: " + text, "user");
  await dispatchTurn(text, true);
}

// Fires once on a reply/forward, mirroring main_app.py's
// _suggest_reply_options -- a synthetic first turn, not shown as "You: ...".
// Attachment auto-review (V4's find_original_message/attach_from_mail_item)
// is NOT ported here yet -- see mail-context.js's header for why; this only
// carries forward the "suggest reply angles automatically" behavior, over
// plain text context.
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

async function dispatchTurn(text, isDraftCandidate) {
  history.push({ role: "user", content: text });
  setBusy(true);

  const systemPrompt = await buildSystemPrompt(currentItem);

  const assistantDiv = appendTurn("Claude:\n", "assistant");
  let streamedText = "";
  const onDelta = (chunk) => {
    streamedText += chunk;
    assistantDiv.textContent = "Claude:\n" + streamedText;
    const transcript = document.getElementById("transcript");
    transcript.scrollTop = transcript.scrollHeight;
  };

  const result = await sendTurnStreaming(settings, systemPrompt, history, onDelta);

  setBusy(false);

  if (result.ok) {
    history.push({ role: "assistant", content: result.text });
    if (isDraftCandidate) {
      lastDraft = result.text;
      document.getElementById("insertBtn").disabled = false;
    }
    // Re-render the final text from the authoritative result, not the
    // accumulated deltas -- same "trust the terminal result, not the
    // streamed pieces" contract as V4's own claude_client/openrouter_client.
    assistantDiv.textContent = "Claude:\n" + result.text;
  } else {
    // Roll back the just-pushed user turn on failure, same as
    // main_app.py's _handle_claude_result -- a retry shouldn't desync
    // history with a dangling unanswered user turn.
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
