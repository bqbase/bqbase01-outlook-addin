/*
============================================================================
 attachments.js -- drag-and-drop file attachments for the chat input
============================================================================

 Classifies a dropped file, builds the matching content block or extracted
 text, and caps size/count. DOCX/XLSX parsing has no lightweight browser
 equivalent without pulling in a real parsing library, so this supports
 images, PDF, and plain text only. Unsupported types get a clear per-file
 error -- never silently dropped.

 Uses the plain HTML5 Drag and Drop API, NOT Office.js's DragAndDropEvent --
 confirmed via Microsoft's own drag-drop-items.md doc: dropping an OS
 desktop file onto a task pane is explicitly "Not supported" via the
 Office.js event and "Supported" via HTML Drag and Drop, and HTML Drag and
 Drop is also what classic Outlook on Windows and Mac require anyway. One
 implementation, works everywhere this add-in targets.
============================================================================
*/

const IMAGE_MEDIA_TYPES = {
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".png": "image/png",
  ".gif": "image/gif",
  ".webp": "image/webp",
};
const PDF_EXTENSIONS = new Set([".pdf"]);
const TEXT_EXTENSIONS = new Set([".txt"]);
const UNSUPPORTED_BUT_KNOWN = new Set([".docx", ".xlsx", ".msg"]); // clear error, not silent drop

const MAX_ATTACHMENT_BYTES = 20 * 1024 * 1024; // 20MB per file
const MAX_ATTACHMENTS_PER_MESSAGE = 10;
// No truncation ceiling. A .txt attachment is sent whole; the only bound
// is the Worker's 30MB body cap. The old 200,000-char limit cut documents
// off mid-sentence and told the model so, which quietly degraded exactly
// the attachment review this assistant exists for -- and it was never a
// cost question: 200,000 chars is ~50,000 tokens, about one cent.

function classifyFile(filename) {
  const ext = _extOf(filename);
  if (ext in IMAGE_MEDIA_TYPES) return "image";
  if (PDF_EXTENSIONS.has(ext)) return "pdf";
  if (TEXT_EXTENSIONS.has(ext)) return "txt";
  if (UNSUPPORTED_BUT_KNOWN.has(ext)) return "unsupported-known";
  return "unknown";
}

function _extOf(filename) {
  const idx = filename.lastIndexOf(".");
  return idx === -1 ? "" : filename.slice(idx).toLowerCase();
}

// processFiles(fileList) -> Promise<Attachment[]>
// Attachment: {filename, category, contentBlock, extractedText, error}
// -- exactly one of contentBlock/extractedText/error is set.
async function processFiles(fileList) {
  const files = Array.from(fileList).slice(0, MAX_ATTACHMENTS_PER_MESSAGE);
  const results = [];
  for (const file of files) {
    results.push(await _processOneFile(file));
  }
  return results;
}

async function _processOneFile(file) {
  const category = classifyFile(file.name);

  if (category === "unknown") {
    return {
      filename: file.name, category,
      error: "Unsupported file type. Supported: images (jpg/png/gif/webp), PDF, and plain text (.txt).",
    };
  }
  if (category === "unsupported-known") {
    return {
      filename: file.name, category,
      error: "DOCX/XLSX/.msg attachments aren't supported yet in this add-in (known, tracked gap -- see STATE.md).",
    };
  }
  if (file.size > MAX_ATTACHMENT_BYTES) {
    return {
      filename: file.name, category,
      error: "File is too large (" + (file.size / (1024 * 1024)).toFixed(1) + " MB, max " + (MAX_ATTACHMENT_BYTES / (1024 * 1024)) + " MB).",
    };
  }

  try {
    if (category === "image") {
      const dataUri = await _readAsDataUrl(file);
      return { filename: file.name, category, bytes: file.size, contentBlock: { type: "image_url", image_url: { url: dataUri } } };
    }
    if (category === "pdf") {
      const dataUri = await _readAsDataUrl(file);
      return { filename: file.name, category, bytes: file.size, contentBlock: { type: "file", file: { filename: file.name, file_data: dataUri } } };
    }
    if (category === "txt") {
      const text = await _readAsText(file);
      return { filename: file.name, category, bytes: file.size, extractedText: text };
    }
  } catch (err) {
    return { filename: file.name, category, error: "Could not read this file: " + err };
  }
}

function _readAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function _readAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsText(file);
  });
}

// ---- attachments already on the message ---------------------------------
//
// Files the USER drops go through processFiles above. These are the files
// already on the open message, which is a different path in every respect:
// the metadata arrives from Office rather than from a File object, and the
// bytes have to be asked for one at a time.
//
// What this does NOT do is reach the attachments of the message a REPLY was
// created from. Outlook drops attachments on Reply and Reply All (Forward
// keeps them), and Office.js has no way to open a different message --
// getAttachmentContentAsync only ever reads the item the pane is open on.
// V4 did it through COM (mail_context.find_original_message, walking
// PR_IN_REPLY_TO_ID to an Items.Find over the right store); nothing in
// Office.js corresponds to any of those three steps. Reaching it needs
// Microsoft Graph and tenant consent -- a separate decision, not a gap here.

// getAttachmentsAsync and getAttachmentContentAsync are newer than the 1.5
// this add-in's manifest asks for. Detected at runtime rather than raising
// the manifest minimum, which would lock out hosts that run everything else
// here perfectly well. Microsoft's own pages disagree about the number
// (1.1 on classic Windows, 1.8 on the web) so the CALL is what gets tested,
// not the version.
function attachmentApiAvailable(item) {
  try {
    if (Array.isArray(item.attachments)) return true;   // read mode needs no API
    return typeof item.getAttachmentsAsync === "function" &&
           typeof item.getAttachmentContentAsync === "function";
  } catch (err) {
    return false;
  }
}

// Lists what is attached to the open message WITHOUT fetching any bytes.
// Read mode exposes item.attachments as a plain array; compose mode has no
// such property at all and requires getAttachmentsAsync -- the same
// read/compose shape difference this add-in already handles for item.subject
// and item.body.
function listItemAttachments(item) {
  return new Promise((resolve) => {
    if (Array.isArray(item.attachments)) {
      resolve(item.attachments);
      return;
    }
    if (typeof item.getAttachmentsAsync !== "function") {
      resolve([]);
      return;
    }
    try {
      item.getAttachmentsAsync((result) => {
        const ok = result && result.status === Office.AsyncResultStatus.Succeeded;
        resolve(ok && result.value ? result.value : []);
      });
    } catch (err) {
      resolve([]);
    }
  });
}

// Turns Office's attachment metadata into the same shape processFiles()
// produces, so the chips, the estimate and the message builder all treat
// both sources identically.
//
// Nothing is fetched here on purpose: the price has to appear the moment the
// pane opens, and a round trip per file would make the customer wait to find
// out what something costs.
function describeItemAttachments(details) {
  const out = [];
  for (const att of details || []) {
    // Inline images are signature logos and embedded decoration. Charging
    // 2 coins to "review" a corporate footer would be indefensible, and
    // every message from a corporate sender carries several.
    if (att.isInline) continue;
    const bytes = att.size || 0;
    // An attached EMAIL is not a file -- there are no bytes to fetch, and
    // .msg is unsupported anyway, so it gets the same clear error.
    const isItem = att.attachmentType === "item";
    const category = isItem ? "unsupported-known" : classifyFile(att.name);
    const entry = { filename: att.name, category, bytes: bytes,
                    source: "item", attachmentId: att.id };
    if (category === "unknown") {
      entry.error = "Unsupported file type. Supported: images (jpg/png/gif/webp), PDF, and plain text (.txt).";
    } else if (category === "unsupported-known") {
      entry.error = isItem
        ? "This is an attached email, which this add-in cannot read yet."
        : "DOCX/XLSX/.msg attachments aren't supported yet in this add-in (known, tracked gap -- see STATE.md).";
    } else if (bytes > MAX_ATTACHMENT_BYTES) {
      entry.error = "File is too large (" + (bytes / (1024 * 1024)).toFixed(1) +
        " MB, max " + (MAX_ATTACHMENT_BYTES / (1024 * 1024)) + " MB).";
    }
    out.push(entry);
  }
  return out;
}

// Fetches one attachment's bytes and fills in the contentBlock or
// extractedText the message builder needs. Mutates and resolves the entry.
//
// Failure sets entry.error rather than throwing, so one unreadable file
// costs its own coins and nothing else -- Office refuses to release some
// attachments (protected messages, and files above a host-specific size),
// and that refusal arrives here AFTER the price was quoted.
function fetchItemAttachment(item, entry) {
  return new Promise((resolve) => {
    if (entry.error || entry.source !== "item") {
      resolve(entry);
      return;
    }
    try {
      item.getAttachmentContentAsync(entry.attachmentId, (result) => {
        const ok = result && result.status === Office.AsyncResultStatus.Succeeded;
        if (!ok || !result.value) {
          entry.error = "Outlook would not release this attachment" +
            (result && result.error ? " (" + result.error.message + ")" : "") + ".";
          resolve(entry);
          return;
        }
        const format = result.value.format;
        const content = result.value.content;
        if (format !== "base64") {
          entry.error = "This attachment came back as " + format +
            ", which this add-in cannot read.";
          resolve(entry);
          return;
        }
        try {
          if (entry.category === "txt") {
            entry.extractedText = _decodeBase64Text(content);
          } else if (entry.category === "image") {
            entry.contentBlock = { type: "image_url", image_url: { url: _dataUri(entry, content) } };
          } else if (entry.category === "pdf") {
            entry.contentBlock = { type: "file", file: { filename: entry.filename, file_data: _dataUri(entry, content) } };
          }
        } catch (err) {
          entry.error = "Could not read this attachment: " + err;
        }
        resolve(entry);
      });
    } catch (err) {
      entry.error = "Could not read this attachment: " + err;
      resolve(entry);
    }
  });
}

function _dataUri(entry, base64) {
  const mime = IMAGE_MEDIA_TYPES[_extOf(entry.filename)] ||
    (entry.category === "pdf" ? "application/pdf" : "application/octet-stream");
  return "data:" + mime + ";base64," + base64;
}

// atob yields one character per BYTE, so a UTF-8 file with any accented
// character decodes to mojibake unless the bytes are reassembled and decoded
// as UTF-8 explicitly.
function _decodeBase64Text(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new TextDecoder("utf-8").decode(bytes);
}

// ---- coin estimate -------------------------------------------------------
//
// MUST match the Worker's attachmentCoins() exactly. The number shown next to
// the Review button is the number the customer is charged -- that is the
// promise -- so if these two formulas ever disagree, the customer sees one
// price and pays another. Change both together, or not at all.
//
// The Worker works from the base64 payload it receives and this works from
// the raw file, which is why the ratio appears here and not there: base64
// carries 3 bytes per 4 characters, so the Worker's length * 0.75 recovers
// the original size that these numbers are already in.
const COINS_PER_IMAGE = 2;
const COINS_PER_MB_PDF = 3;
const TEXT_CHARS_PER_COIN = 20000;

function estimateCoins(attachments) {
  let fileBytes = 0;
  let images = 0;
  let textChars = 0;
  for (const att of attachments || []) {
    if (att.error) continue; // never charge for a file that will not be sent
    if (att.category === "image") images += 1;
    else if (att.category === "pdf") fileBytes += att.bytes || 0;
    else if (att.category === "txt") {
      // A dropped file has been read already, so its exact character count is
      // known. One still ON the message has not -- quoting it would mean
      // fetching every file just to price it, and the price has to appear the
      // moment the pane opens. Its BYTE count stands in, which for UTF-8 is
      // always >= the character count the Worker charges on. So this can only
      // ever quote HIGH, never low: the customer is never charged more than
      // the number they pressed the button on.
      textChars += att.extractedText !== undefined
        ? att.extractedText.length
        : (att.bytes || 0);
    } else if (att.extractedText !== undefined) {
      textChars += att.extractedText.length;
    }
  }
  const coins =
    Math.ceil((fileBytes / (1024 * 1024)) * COINS_PER_MB_PDF) +
    images * COINS_PER_IMAGE +
    Math.ceil(textChars / TEXT_CHARS_PER_COIN);
  return Math.max(1, coins);
}

// buildMessageContent(userText, attachments) -> string | array
// Returns a plain string when nothing is attached, else a content-block
// array: text-extracted content first, then native image/file blocks, then
// the user's own text last -- supporting material before the instruction
// that references it.
function buildMessageContent(userText, attachments) {
  const ok = (attachments || []).filter((a) => !a.error);
  if (ok.length === 0) return userText;

  const blocks = [];
  for (const a of ok) {
    if (a.extractedText !== undefined) {
      blocks.push({ type: "text", text: "[Attached file: " + a.filename + "]\n" + a.extractedText });
    }
  }
  for (const a of ok) {
    if (a.contentBlock !== undefined) {
      blocks.push(a.contentBlock);
    }
  }
  blocks.push({ type: "text", text: userText.trim() || "(see attached file(s))" });
  return blocks;
}
