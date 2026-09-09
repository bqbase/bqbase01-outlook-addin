/*
============================================================================
 providers.js -- streaming chat call to the BQBase proxy
============================================================================

 Streams a chat turn: zero or more text deltas via the onDelta callback,
 then exactly one terminal result. One model, not user-configurable.

 This file does NOT know which provider is upstream, and should not: it
 talks only to the Worker, which decides. The model id sent here is the id
 the WORKER expects from clients; the Worker translates it for whichever
 provider it is pointed at. That indirection is what let the upstream move
 from OpenRouter to OpenAI on 2026-09-08 with no change to this file.

 An earlier revision supported OpenAI and Anthropic as well, with each user
 supplying their own key via a Settings screen. That was removed on request
 (2026-08-25): one provider, one model, no settings.

 NO API KEY IS IN THIS FILE, deliberately. A task pane is just a web page,
 so any key it held would be readable by anyone viewing the source. Requests
 instead go to a Cloudflare Worker (source in ../worker) which holds the key
 as an encrypted secret and adds the Authorization header server-side. The
 Worker also restricts calling origins and the model, so the endpoint being
 public does not mean it can be used for arbitrary spend.

 sendTurnStreaming(systemPrompt, history, onDelta) -> Promise<{ok, text,
 error, refusalCategory}>
   history: [{role: "user"|"assistant", content: string}], NOT including the
            system prompt (that's passed separately)
   onDelta: called with each incremental text chunk as it streams in
            (display-only -- the final resolved text is authoritative, not
            the concatenation of deltas, though in practice they should
            match)

 Never throws -- every failure path resolves {ok: false, error: "..."} so
 taskpane.js's caller doesn't need a try/catch around this by contract.
============================================================================
*/

// The proxy endpoint. NO API KEY LIVES HERE -- see this file's header.
// The word "openrouter" in this hostname is a LEFTOVER from the
// original provider; the Worker has called OpenAI directly since
// 2026-09-08. Renaming the Worker would break all three clients at
// once, so the stale name was kept deliberately.
const PROXY_URL = "https://bqbase-openrouter-proxy.bqbase.workers.dev";
const REQUEST_MODEL = "openai/gpt-5.6-luna";
const REASONING_EFFORT = "medium";

// The customer's access token.
//
// TWO stores, because they fail differently. localStorage is per machine and
// per client, so a customer had to paste the token on their desktop, their
// laptop and again in Outlook on the web. The optional store below is backed
// by the MAILBOX, so one paste follows them everywhere.
//
// Keeping a credential in the mailbox was avoided at first, on the grounds
// that other add-ins in that mailbox could read it. That reasoning is weaker
// now the token is registered to one mailbox: anything with enough access to
// read it is already inside the only account it can spend. Weighed against
// making the customer paste a 51-character string three times, the mailbox
// wins -- and localStorage remains as the fallback where roaming is absent.
const TOKEN_KEY = "bqbase_token";

// Supplied by the host at startup (see taskpane.js). Null on any client that
// cannot offer one, which simply leaves the localStorage behaviour intact.
let tokenStore = null;

function setTokenStore(store) {
  tokenStore = store;
}

function getToken() {
  // The mailbox is authoritative: it is the thing the subscription belongs
  // to, so a token found there beats a stale one left on this machine.
  if (tokenStore) {
    try {
      const roamed = tokenStore.load();
      if (roamed) return roamed;
    } catch (err) {
      // fall through to the local copy
    }
  }
  try {
    return localStorage.getItem(TOKEN_KEY) || "";
  } catch (err) {
    return ""; // storage blocked -- treated as "no token", not as an error
  }
}

// Written to BOTH, so a host that later loses one still has the other, and
// so a customer who has already pasted the token on this machine gets it
// carried up to the mailbox without doing anything.
function setToken(value) {
  let stored = false;
  try {
    localStorage.setItem(TOKEN_KEY, value);
    stored = true;
  } catch (err) {
    // no local copy; the mailbox may still take it
  }
  if (tokenStore) {
    try {
      tokenStore.save(value);
      stored = true;
    } catch (err) {
      // no roaming copy; the local one may have worked
    }
  }
  return stored;
}

// The signed-in mailbox, which is what the subscription is tied to. Set once
// by taskpane.js at startup and sent with every call, so the same token used
// on a desktop, a laptop and Outlook on the web is one account -- and so a
// token forwarded to a colleague is refused in their mailbox.
//
// NOT read from Office.context here on purpose: this file is the transport
// layer and knows nothing about the host, which is what let the upstream
// provider change without touching it.
let mailboxAddress = "";

function setMailbox(address) {
  mailboxAddress = String(address || "").trim();
}

// Every request that carries the token carries the mailbox with it, so the
// two are never checked apart.
function authHeaders() {
  const headers = {};
  const token = getToken();
  if (token) headers["X-BQBase-Token"] = token;
  if (mailboxAddress) headers["X-BQBase-Mailbox"] = mailboxAddress;
  return headers;
}

// Asks the Worker what this token is worth. Returns {ok, balance} or
// {ok:false, error} -- never throws, same contract as everything else here.
async function fetchBalance() {
  const token = getToken();
  if (!token) return { ok: false, error: "No access token set." };
  let response;
  try {
    response = await fetch(PROXY_URL + "/balance", {
      method: "GET",
      headers: authHeaders(),
    });
  } catch (err) {
    return { ok: false, error: "Could not reach the assistant service." };
  }
  const bodyText = await _safeReadText(response);
  if (!response.ok) {
    return { ok: false, error: _serverMessage(bodyText) || "That token was not accepted." };
  }
  try {
    return { ok: true, balance: JSON.parse(bodyText) };
  } catch (err) {
    return { ok: false, error: "Unreadable reply from the assistant service." };
  }
}

// purpose is what the call is CHARGED as: "review" is free, "reply" is one
// coin, "attachment" is priced from the files. The Worker decides the price;
// this only declares the kind.
async function sendTurnStreaming(systemPrompt, history, onDelta, purpose) {
  const messages = [{ role: "system", content: systemPrompt }].concat(history);
  let response;
  try {
    const headers = Object.assign({ "Content-Type": "application/json" }, authHeaders());
    response = await fetch(PROXY_URL, {
      method: "POST",
      headers: headers,
      body: JSON.stringify({
        model: REQUEST_MODEL,
        messages: messages,
        stream: true,
        reasoning_effort: REASONING_EFFORT,
        bqbase_purpose: purpose || "reply",
      }),
    });
  } catch (err) {
    return { ok: false, error: "Network error contacting the proxy: " + err };
  }

  if (!response.ok) {
    const bodyText = await _safeReadText(response);
    if (response.status === 401) {
      return { ok: false, error: "The proxy's API key was rejected (401)." };
    }
    if (response.status === 403) {
      // Also how a suspended account arrives, and the Worker's wording is
      // far more useful than anything guessable from the status alone.
      return { ok: false, error: _serverMessage(bodyText) || "The proxy refused this request (403)." };
    }
    if (response.status === 402) {
      // The daily spending cap. Shown verbatim because only the Worker knows
      // the figure and when it resets -- and because "wait a moment", the
      // 429 wording below, would be a lie here.
      return { ok: false, error: _serverMessage(bodyText) || "Daily spending cap reached." };
    }
    if (response.status === 429) {
      return { ok: false, error: "Rate limited (429). Wait a moment and try again." };
    }
    return { ok: false, error: "API error (HTTP " + response.status + "): " + bodyText };
  }

  // The coin figures ride back on headers, so the pane can update its counter
  // without a second round trip to /balance. Absent for the unmetered paths,
  // which is why every caller must tolerate nulls here.
  const coins = {
    charged: _headerNumber(response, "X-BQBase-Coins-Charged"),
    left: _headerNumber(response, "X-BQBase-Coins-Left"),
    resets: response.headers.get("X-BQBase-Resets"),
  };
  const result = await _readSse(response, onDelta);
  result.coins = coins;
  return result;
}

function _headerNumber(response, name) {
  const raw = response.headers.get(name);
  if (raw === null || raw === "") return null;
  const value = Number(raw);
  return Number.isFinite(value) ? value : null;
}

// Parses an OpenAI-compatible Server-Sent-Events stream: lines
// starting "data: {json}", terminated by a literal "data: [DONE]" line. Each
// JSON chunk's choices[0].delta.content is the incremental text, same field
// path -- hand-parsed here since there is no SDK in a browser task pane.
async function _readSse(response, onDelta) {
  const reader = response.body.getReader();
  const decoder = new TextDecoder("utf-8");
  let buffer = "";
  let textParts = [];
  let finishReason = null;

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n");
      buffer = lines.pop(); // last (possibly incomplete) line stays buffered
      for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line.startsWith("data:")) continue;
        const payload = line.slice("data:".length).trim();
        if (payload === "[DONE]") continue;
        let parsed;
        try {
          parsed = JSON.parse(payload);
        } catch (err) {
          continue; // a stray/partial line -- skip rather than abort the whole stream
        }
        const choice = parsed.choices && parsed.choices[0];
        if (!choice) continue;
        if (choice.finish_reason) finishReason = choice.finish_reason;
        const deltaText = choice.delta && choice.delta.content;
        if (deltaText) {
          textParts.push(deltaText);
          onDelta(deltaText);
        }
      }
    }
  } catch (err) {
    return { ok: false, error: "Error while streaming: " + err };
  }

  const text = textParts.join("").trim();

  // content_filter -- a successful HTTP response that is nonetheless a
  // refusal; checked before trusting the accumulated text.
  if (finishReason === "content_filter") {
    return { ok: false, refusalCategory: "content_filter" };
  }
  if (!text) {
    return { ok: false, error: "Model finished (" + finishReason + ") without producing any text." };
  }
  return { ok: true, text: text };
}

// Pulls the human-readable message out of the Worker's error envelope,
// {error:{message}}, so a server-authored explanation reaches the user
// instead of a raw JSON blob.
function _serverMessage(bodyText) {
  try {
    return JSON.parse(bodyText).error.message;
  } catch (err) {
    return null;
  }
}

async function _safeReadText(response) {
  try {
    return await response.text();
  } catch (err) {
    return "(could not read error body)";
  }
}
