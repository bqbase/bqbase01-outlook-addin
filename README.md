# BQBase-01 — Outlook Add-in

Hosted assets for the BQBase-01 Outlook Add-in (task pane + icons), served via
GitHub Pages so Outlook (desktop, new, and web) can load them over a real
HTTPS URL — `localhost` URLs are rejected by Office Add-in manifest validation
for icons, and don't survive outside this one dev machine anyway.

This is a client-only Office Add-in: it reads the current email via Office.js,
lets the user chat with an LLM about it, and can insert a drafted reply into
the compose body. No backend server — requests go from the user's own
browser/Outlook client straight to OpenRouter. Nothing passes through any
server this project operates.

The provider, model and API key are fixed in `providers.js` (OpenRouter,
`openai/gpt-5.6-luna`); there is no Settings screen and nothing for the user
to configure. **The API key is therefore visible in the source of this
publicly hosted page** — a deliberate choice by the project owner, superseding
the earlier bring-your-own-key design. Rotate the key at openrouter.ai if it
is ever abused.

See `D:\TerraSync\Claude\Project\Email_Assistant_V5\STATE.md` for full project
history/decisions — this repo holds only what must be publicly hosted, not the
project's working notes.
