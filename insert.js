/*
============================================================================
 insert.js -- inserting a drafted reply into the compose body
============================================================================

 DO NOT USE body.setAsync HERE. It REPLACES THE ENTIRE BODY (per Microsoft's
 own docs: "the entire body of the conversation thread is replaced" unless
 bodyMode: HostConfig is used, and even then a whole reply's body, not just
 an insertion point) -- on a reply that silently destroys the user's quoted
 original thread. This module uses setSelectedDataAsync instead: insert AT
 THE CURSOR, replacing only a text SELECTION if one exists. Never
 setAsync/prependAsync's whole-body semantics.

 Known limitation, accepted rather than solved: Office.js gives a task pane
 no bookmark or range identity to hold onto across calls, so clicking Insert
 twice inserts the draft twice, at wherever the cursor happens to be each
 time -- there is no "replace my own previous insertion" behavior here.
 Documented in the UI (see taskpane.js).
============================================================================
*/

// insertDraft(item, rawDraftText) -> Promise<{ok, message}>
// rawDraftText is the model's raw reply text -- may or may not follow the
// SUBJECT:/body/<<<END EMAIL>>> contract (see mail-context.js's parseDraft).
async function insertDraft(item, rawDraftText) {
  const { proposedSubject, body } = parseDraft(rawDraftText);

  if (proposedSubject) {
    const subjectResult = await _setSubject(item, proposedSubject);
    if (!subjectResult.ok) {
      // Non-fatal -- still attempt the body insert even if the subject
      // update failed: partial success beats total failure.
      // eslint-disable-next-line no-console
      console.warn("insertDraft: subject update failed: " + subjectResult.message);
    }
  }

  const bodyType = await _getBodyType(item);
  const coercionType = bodyType === Office.CoercionType.Html ? Office.CoercionType.Html : Office.CoercionType.Text;
  // Body text from the model is always plain text (see mail-context.js's
  // system prompt -- never asked for HTML), so when the compose item is
  // HTML-formatted, newlines need to become <br> or the draft renders as one
  // unbroken paragraph -- confirmed against the documented coercionType
  // table (HTML item + text data must still declare coercionType: Text
  // unless reshaped as real HTML first).
  const dataToInsert = coercionType === Office.CoercionType.Html
    ? body.split("\n").map(_escapeHtml).join("<br>")
    : body;

  return new Promise((resolve) => {
    item.body.setSelectedDataAsync(
      dataToInsert,
      { coercionType: coercionType },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve({ ok: true, message: "Inserted into the email." });
        } else {
          resolve({ ok: false, message: "Insert failed: " + result.error.message });
        }
      }
    );
  });
}

function _getBodyType(item) {
  return new Promise((resolve) => {
    item.body.getTypeAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : Office.CoercionType.Text);
    });
  });
}

function _setSubject(item, subjectText) {
  return new Promise((resolve) => {
    if (typeof item.subject === "string") {
      // Office.js has no synchronous subject setter even in read-shaped
      // items encountered so far in this add-in's compose-only usage, but
      // guard defensively rather than assume -- item.subject is a string in
      // read mode, and this module is compose-only in practice.
      resolve({ ok: false, message: "Subject is read-only in this context." });
      return;
    }
    item.subject.setAsync(subjectText, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve({ ok: true });
      } else {
        resolve({ ok: false, message: result.error.message });
      }
    });
  });
}

function _escapeHtml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
