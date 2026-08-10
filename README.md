# BQBase-01 — Outlook Add-in

Hosted assets for the BQBase-01 Outlook Add-in (task pane + icons), served via
GitHub Pages so Outlook (desktop, new, and web) can load them over a real
HTTPS URL — `localhost` URLs are rejected by Office Add-in manifest validation
for icons, and don't survive outside this one dev machine anyway.

This is a client-only Office Add-in: it reads the current email via Office.js,
lets the user chat with an LLM they've configured with their OWN API key
(OpenRouter, OpenAI, or Anthropic — chosen and entered by the user in the
add-in's own Settings screen), and can insert a drafted reply into the compose
body. No backend server — each user's key is stored via Office.js's
`RoamingSettings` (their own Microsoft account's roaming add-in settings) and
sent directly from their own browser/Outlook client straight to their chosen
provider's API. Nothing passes through any server this project operates.

See `D:\TerraSync\AI\Project\Email_Assisant_V5\STATE.md` for full project
history/decisions — this repo holds only what must be publicly hosted, not the
project's working notes.
