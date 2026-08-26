/*
============================================================================
 providers.js -- OpenRouter streaming chat call
============================================================================

 Mirrors V4's openrouter_client.py in spirit (same generator/callback shape:
 zero or more text deltas, then exactly one terminal result). OpenRouter with
 openai/gpt-5.6-luna is the only provider and the only model -- the key,
 model and reasoning effort are all baked in below, not user-configurable.

 An earlier revision supported OpenAI and Anthropic as well, with each user
 supplying their own key via a Settings screen. That was removed on request
 (2026-08-25): one provider, one model, one key, no settings.

 SECURITY, stated plainly rather than buried: OPENROUTER_API_KEY below is a
 real, live key shipped inside a task pane served from a PUBLIC GitHub Pages
 site. Anyone who views the page source can read it and spend against this
 account. This was an explicit, informed decision by the project owner, not
 an oversight -- rotate the key at openrouter.ai if it is ever abused.

 sendTurnStreaming(systemPrompt, history, onDelta) -> Promise<{ok, text,
 error, refusalCategory}>
   history: [{role: "user"|"assistant", content: string}], NOT including the
            system prompt (that's passed separately, same split V4's
            build_system_prompt/history kept)
   onDelta: called with each incremental text chunk as it streams in
            (display-only, same contract as V4's StreamDelta -- the final
            resolved text is authoritative, not the concatenation of deltas,
            though in practice they should match)

 Never throws -- every failure path resolves {ok: false, error: "..."} so
 taskpane.js's caller doesn't need a try/catch around this by contract,
 mirroring openrouter_client.py's "yield ClaudeTurnResult(error=...)" pattern
 that never lets an exception escape to main_app.py.
============================================================================
*/

// Baked-in OpenRouter credentials and call settings. Matches V4's
// config.ini values (model openai/gpt-5.6-luna, effort medium).
const OPENROUTER_API_KEY =
  "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";
const OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions";
const OPENROUTER_MODEL = "openai/gpt-5.6-luna";
const OPENROUTER_EFFORT = "medium";

async function sendTurnStreaming(systemPrompt, history, onDelta) {
  const messages = [{ role: "system", content: systemPrompt }].concat(history);
  let response;
  try {
    response = await fetch(OPENROUTER_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + OPENROUTER_API_KEY,
      },
      body: JSON.stringify({
        model: OPENROUTER_MODEL,
        messages: messages,
        stream: true,
        reasoning_effort: OPENROUTER_EFFORT,
      }),
    });
  } catch (err) {
    return { ok: false, error: "Network error contacting openrouter.ai: " + err };
  }

  if (!response.ok) {
    const bodyText = await _safeReadText(response);
    if (response.status === 401) {
      return { ok: false, error: "OpenRouter rejected the built-in API key (401)." };
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
// path openrouter_client.py reads via the openai SDK's typed
// chunk.choices[0].delta.content -- this is that same wire format,
// hand-parsed here since there is no SDK in a browser task pane.
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

  // content_filter -- same successful-but-refused case openrouter_client.py
  // checks for before trusting the accumulated text.
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
