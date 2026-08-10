/*
============================================================================
 settings.js -- per-user provider/model/API key settings
============================================================================

 "Bring your own key" storage via Office.js RoamingSettings -- syncs across
 the user's own devices via their Microsoft account, survives cache clears,
 no backend of ours involved at all (see providers.js's header and the
 project's outlook_addin/README.md for why there is no backend).

 SECURITY NOTE, recorded here deliberately, not glossed over: Microsoft's own
 RoamingSettings documentation explicitly states it "shouldn't be used to
 store sensitive information, such as user credentials or security tokens" --
 it is accessible via Exchange Web Services/Extended MAPI, not encrypted at
 rest the way a password manager would be. This project uses it anyway for
 the API key, as an explicit, informed tradeoff (chosen over the alternative,
 browser-only localStorage, for the "same settings on every device" benefit) --
 not an oversight. The key is the USER'S OWN key, in THEIR OWN mailbox's
 roaming data, never sent anywhere except directly to their chosen provider's
 API (see providers.js) -- never to any server this project operates, since
 there isn't one. The Settings UI (see taskpane.js) says this plainly so a
 novice user isn't surprised by where their key lives.
============================================================================
*/

const SETTINGS_KEYS = {
  PROVIDER: "bqbase01_provider",
  API_KEY: "bqbase01_apiKey",
  MODEL: "bqbase01_model",
  EFFORT: "bqbase01_effort",
};

const DEFAULT_PROVIDER = "openrouter";

function loadSettings() {
  const rs = Office.context.roamingSettings;
  const provider = rs.get(SETTINGS_KEYS.PROVIDER) || DEFAULT_PROVIDER;
  const apiKey = rs.get(SETTINGS_KEYS.API_KEY) || "";
  const model = rs.get(SETTINGS_KEYS.MODEL) || PROVIDERS[provider].defaultModel;
  const effort = rs.get(SETTINGS_KEYS.EFFORT) || "medium";
  return { provider, apiKey, model, effort };
}

// saveAsync is genuinely async and can fail (see Office.js docs -- e.g. the
// 32KB roaming-settings cap, though a key+model/effort string is nowhere
// close to that). Returns a Promise so taskpane.js can await it and show a
// clear error rather than silently losing a settings change.
function saveSettings(settings) {
  const rs = Office.context.roamingSettings;
  rs.set(SETTINGS_KEYS.PROVIDER, settings.provider);
  rs.set(SETTINGS_KEYS.API_KEY, settings.apiKey);
  rs.set(SETTINGS_KEYS.MODEL, settings.model);
  rs.set(SETTINGS_KEYS.EFFORT, settings.effort);
  return new Promise((resolve, reject) => {
    rs.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error ? result.error.message : "Unknown error saving settings"));
      }
    });
  });
}

function hasValidSettings(settings) {
  return Boolean(settings && settings.apiKey && settings.provider && settings.model);
}
