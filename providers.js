/*
============================================================================
 providers.js -- per-provider streaming chat calls
============================================================================

 Mirrors V4's openrouter_client.py in spirit (same generator/callback shape:
 zero or more text deltas, then exactly one terminal result), but supports
 three providers instead of one, since this add-in is "bring your own key" --
 each user picks their own provider/model/key in Settings (see settings.js),
 not a fixed one baked into the app the way V4's config.ini was.

 All three providers were confirmed LIVE via real CORS preflight checks
 (OPTIONS requests, not just documentation) before this was written:
   - OpenRouter: Access-Control-Allow-Origin: *  (no special header needed)
   - OpenAI:     Access-Control-Allow-Origin echoes the request Origin;
                 allows content-type + authorization headers
   - Anthropic:  Access-Control-Allow-Origin: *, requires the
                 anthropic-dangerous-direct-browser-access: true header
                 (Anthropic's own documented flag for exactly this
                 bring-your-own-key browser pattern)

 sendTurnStreaming(settings, systemPrompt, history, onDelta) -> Promise<{ok,
 text, error, refusalCategory}>
   settings: {provider, apiKey, model, effort} from settings.js
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

const PROVIDERS = {
  openrouter: {
    label: "OpenRouter",
    // Reasonable default; user can override in Settings. openai/gpt-5.6-luna
    // is what V4 uses server-side -- kept as the suggested default here too,
    // not hardcoded/forced, since OpenRouter hosts many models and the user
    // may prefer a different one.
    defaultModel: "openai/gpt-5.6-luna",
    supportsEffort: true,
  },
  openai: {
    label: "OpenAI",
    defaultModel: "gpt-5.1",
    supportsEffort: false,
  },
  anthropic: {
    label: "Anthropic",
    defaultModel: "claude-opus-5",
    supportsEffort: false,
  },
};

async function sendTurnStreaming(settings, systemPrompt, history, onDelta) {
  if (!settings || !settings.apiKey) {
    return { ok: false, error: "No API key configured. Open Settings and enter your API key." };
  }
  const provider = settings.provider;
  if (provider === "openrouter") {
    return _streamOpenAiCompatible(
      "https://openrouter.ai/api/v1/chat/completions",
      settings,
      systemPrompt,
      history,
      onDelta,
      { reasoning_effort: settings.effort || "medium" }
    );
  }
  if (provider === "openai") {
    return _streamOpenAiCompatible(
      "https://api.openai.com/v1/chat/completions",
      settings,
      systemPrompt,
      history,
      onDelta,
      {}
    );
  }
  if (provider === "anthropic") {
    return _streamAnthropic(settings, systemPrompt, history, onDelta);
  }
  return { ok: false, error: "Unknown provider: " + provider };
}

// -- OpenRouter / OpenAI share the same Chat Completions request/response
//    shape (OpenRouter is an OpenAI-compatible proxy) -- one implementation
//    for both, parameterized by URL and extra per-provider body fields,
//    mirroring how openrouter_client.py itself is just the openai SDK
//    pointed at a different base_url.
async function _streamOpenAiCompatible(url, settings, systemPrompt, history, onDelta, extraBody) {
  const messages = [{ role: "system", content: systemPrompt }].concat(history);
  let response;
  try {
    response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + settings.apiKey,
      },
      body: JSON.stringify(Object.assign({
        model: settings.model,
        messages: messages,
        stream: true,
      }, extraBody)),
    });
  } catch (err) {
    return { ok: false, error: "Network error contacting " + url + ": " + err };
  }

  if (!response.ok) {
    const bodyText = await _safeReadText(response);
    if (response.status === 401) {
      return { ok: false, error: "API key rejected (401). Check your key in Settings." };
    }
    if (response.status === 429) {
      return { ok: false, error: "Rate limited (429). Wait a moment and try again." };
    }
    return { ok: false, error: "API error (HTTP " + response.status + "): " + bodyText };
  }

  return _readOpenAiCompatibleSse(response, onDelta);
}

// Parses an OpenAI-compatible Server-Sent-Events stream: lines starting
// "data: {json}", terminated by a literal "data: [DONE]" line. Each JSON
// chunk's choices[0].delta.content is the incremental text, same field path
// openrouter_client.py reads via the openai SDK's typed chunk.choices[0]
// .delta.content -- this is that same wire format, hand-parsed here since
// there is no SDK in a browser task pane.
async function _readOpenAiCompatibleSse(response, onDelta) {
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

// Anthropic's Messages API has a different wire shape from OpenAI-compatible
// APIs (SSE event TYPES like "content_block_delta", not a flat delta.content
// per chunk) -- kept as its own function rather than shoehorned into the
// shared one above, same reasoning V4 had for isolating provider-specific
// code (see the master prompt's own "isolate provider-specific code"
// guidance).
async function _streamAnthropic(settings, systemPrompt, history, onDelta) {
  let response;
  try {
    response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": settings.apiKey,
        "anthropic-version": "2023-06-01",
        // Anthropic's own documented flag enabling exactly this
        // bring-your-own-key, direct-from-browser pattern -- confirmed live
        // via a real CORS preflight before this was written, not assumed
        // from documentation alone.
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({
        model: settings.model,
        max_tokens: 8000,
        system: systemPrompt,
        messages: history,
        stream: true,
      }),
    });
  } catch (err) {
    return { ok: false, error: "Network error contacting api.anthropic.com: " + err };
  }

  if (!response.ok) {
    const bodyText = await _safeReadText(response);
    if (response.status === 401) {
      return { ok: false, error: "API key rejected (401). Check your key in Settings." };
    }
    if (response.status === 429) {
      return { ok: false, error: "Rate limited (429). Wait a moment and try again." };
    }
    return { ok: false, error: "API error (HTTP " + response.status + "): " + bodyText };
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder("utf-8");
  let buffer = "";
  let textParts = [];
  let stopReason = null;

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n");
      buffer = lines.pop();
      for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line.startsWith("data:")) continue;
        const payload = line.slice("data:".length).trim();
        if (!payload) continue;
        let parsed;
        try {
          parsed = JSON.parse(payload);
        } catch (err) {
          continue;
        }
        if (parsed.type === "content_block_delta" && parsed.delta && parsed.delta.type === "text_delta") {
          textParts.push(parsed.delta.text);
          onDelta(parsed.delta.text);
        } else if (parsed.type === "message_delta" && parsed.delta && parsed.delta.stop_reason) {
          stopReason = parsed.delta.stop_reason;
        }
      }
    }
  } catch (err) {
    return { ok: false, error: "Error while streaming: " + err };
  }

  const text = textParts.join("").trim();

  // Opus 5's refusal-as-200 behavior (documented in V4's own provider notes)
  // has no direct Anthropic-streaming-SSE equivalent surfaced here yet --
  // stop_reason "refusal" IS a real documented value, checked the same way.
  if (stopReason === "refusal") {
    return { ok: false, refusalCategory: "refusal" };
  }
  if (!text) {
    return { ok: false, error: "Model finished (" + stopReason + ") without producing any text." };
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
