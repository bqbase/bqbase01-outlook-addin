/*
============================================================================
 providers.js -- OpenRouter streaming chat call
============================================================================

 Streams a chat turn: zero or more text deltas via the onDelta callback,
 then exactly one terminal result. OpenRouter with openai/gpt-5.6-luna is
 the only provider and the only model, and neither is user-configurable.

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
// Deployed from ../worker; change this if the Worker is renamed.
const PROXY_URL = "https://bqbase-openrouter-proxy.bqbase.workers.dev";
const OPENROUTER_MODEL = "openai/gpt-5.6-luna";
const OPENROUTER_EFFORT = "medium";

async function sendTurnStreaming(systemPrompt, history, onDelta) {
  const messages = [{ role: "system", content: systemPrompt }].concat(history);
  let response;
  try {
    response = await fetch(PROXY_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: OPENROUTER_MODEL,
        messages: messages,
        stream: true,
        reasoning_effort: OPENROUTER_EFFORT,
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
      return { ok: false, error: "The proxy refused this request (403). Check its allowed origins and model." };
    }
    if (response.status === 429) {
      return { ok: false, error: "Rate limited (429). Wait a moment and try again." };
    }
    return { ok: false, error: "API error (HTTP " + response.status + "): " + bodyText };
  }

  return _readSse(response, onDelta);
}

// Parses OpenRouter's OpenAI-compatible Server-Sent-Events stream: lines
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

async function _safeReadText(response) {
  try {
    return await response.text();
  } catch (err) {
    return "(could not read error body)";
  }
}
