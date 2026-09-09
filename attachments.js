/*
============================================================================
 attachments.js -- file attachments, dropped in and already on the message
============================================================================

 Classifies a file, builds the matching content block or extracted text, and
 caps size and count. TWO sources: files the user drops or browses to, and
 the ones already on the open message (see "attachments already on the
 message" below).

 Supported: images (jpg/png/gif/webp), PDF, plain text (.txt/.csv/.md/.log),
 and Word/Excel/PowerPoint via ooxml.js, which reads them without a library
 because they are ZIP archives of XML. The LEGACY binary .doc/.xls/.ppt are
 not ZIPs and cannot be read in a browser; they, .rtf and .msg get a clear
 per-file error -- never silently dropped.

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
const TEXT_EXTENSIONS = new Set([".txt", ".csv", ".md", ".log"]);
// Word, Excel and PowerPoint, read by ooxml.js. Their extension maps to the
// category name so classifyFile can return it directly.
const OFFICE_EXTENSIONS = { ".docx": "docx", ".xlsx": "xlsx", ".pptx": "pptx" };
// The LEGACY binary formats, which are not ZIPs and cannot be read in a
// browser. A clear per-file error, never a silent drop.
const UNSUPPORTED_BUT_KNOWN = new Set([".doc", ".xls", ".ppt", ".msg", ".rtf"]);
// Every category that becomes extracted TEXT rather than a native block.
// These are billed on their text, so they must be extracted BEFORE the
// customer is quoted -- see resolveTextAttachments.
const TEXT_PRODUCING = new Set(["txt", "docx", "xlsx", "pptx"]);

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
  if (ext in OFFICE_EXTENSIONS) return OFFICE_EXTENSIONS[ext];
  if (TEXT_EXTENSIONS.has(ext)) return "txt";
  if (UNSUPPORTED_BUT_KNOWN.has(ext)) return "unsupported-known";
  return "unknown";
}

function _extOf(filename) {
  const idx = filename.lastIndexOf(".");
  return idx === -1 ? "" : filename.slice(idx).toLowerCase();
}

// Names the newer format the user can save as, rather than just refusing.
// The legacy Office formats are binary compound documents, not ZIPs, so
// nothing in the browser can open them -- but "save it as .docx" is a fix
// the customer can actually carry out.
function _legacyFormatError(filename) {
  const ext = _extOf(filename);
  const newer = { ".doc": ".docx", ".xls": ".xlsx", ".ppt": ".pptx" }[ext];
  if (newer) {
    return "Old " + ext + " format isn't readable here. Save it as " + newer + " and attach that.";
  }
  if (ext === ".rtf") return "RTF isn't supported. Save it as .docx or paste the text instead.";
  return "Attached emails (.msg) aren't supported yet.";
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
      error: "Unsupported file type. Supported: Word, Excel, PowerPoint, PDF, images (jpg/png/gif/webp), and text (.txt/.csv/.md/.log).",
    };
  }
  if (category === "unsupported-known") {
    return {
      filename: file.name, category,
      error: _legacyFormatError(file.name),
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
    if (TEXT_PRODUCING.has(category)) {                    // .docx / .xlsx / .pptx
      const buffer = await file.arrayBuffer();
      const text = await extractOfficeText(new Uint8Array(buffer), category);
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
    // getAttachmentContentAsync is required either way -- listing the files
    // is no use without being able to read them. Checking only
    // item.attachments in read mode would show a priced Review button that
    // fails for every file the moment it is pressed.
    if (typeof item.getAttachmentContentAsync !== "function") return false;
    if (Array.isArray(item.attachments)) return true;   // read mode lists them directly
    return typeof item.getAttachmentsAsync === "function";
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
      entry.error = "Unsupported file type. Supported: Word, Excel, PowerPoint, PDF, images (jpg/png/gif/webp), and text (.txt/.csv/.md/.log).";
    } else if (category === "unsupported-known") {
      entry.error = isItem
        ? "This is an attached email, which this add-in cannot read yet."
        : _legacyFormatError(att.name);
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
    // Already resolved: text-producing attachments are fetched when the pane
    // opens so their price can be exact, and must not be fetched twice when
    // Review is pressed.
    if (entry.error || entry.source !== "item" ||
        entry.extractedText !== undefined || entry.contentBlock !== undefined) {
      resolve(entry);
      return;
    }
    try {
      item.getAttachmentContentAsync(entry.attachmentId, async (result) => {
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
          } else if (TEXT_PRODUCING.has(entry.category)) {
            entry.extractedText = await extractOfficeText(_base64ToBytes(content), entry.category);
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

function _base64ToBytes(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

// atob yields one character per BYTE, so a UTF-8 file with any accented
// character decodes to mojibake unless the bytes are reassembled and decoded
// as UTF-8 explicitly.
function _decodeBase64Text(base64) {
  return new TextDecoder("utf-8").decode(_base64ToBytes(base64));
}

// Fetches and extracts every attachment that becomes TEXT, so the price shown
// is the real one rather than a guess from a compressed file's size.
//
// This is the cost of quoting honestly for Office documents: a .docx gives no
// usable clue to how much text is inside -- it is a ZIP, so a small file can
// hold a lot of words -- and pricing from the compressed size could quote
// BELOW what the Worker then charges. PDFs and images are left alone: they
// are priced from size and count, which needs nothing fetched.
async function resolveTextAttachments(item, entries) {
  for (const entry of entries || []) {
    if (entry.error || entry.source !== "item") continue;
    if (!TEXT_PRODUCING.has(entry.category)) continue;
    await fetchItemAttachment(item, entry);
    // A document that yielded no text cannot be priced from its size -- see
    // the fallback note in estimateCoins -- so it is refused here rather than
    // quoted at zero and then charged for whatever the Worker makes of it.
    if (!entry.error && entry.category !== "txt" && entry.extractedText === undefined) {
      entry.error = "No readable text could be extracted from this document.";
    }
  }
  return entries;
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

// Marks a text block as coming from an ATTACHMENT rather than from the
// customer's own typing. The Worker bills only blocks carrying it, so the
// instruction the pane sends alongside the files ("Review the attached
// document...") is not charged for. Both sides must agree on this string.
const ATTACHED_TEXT_MARKER = "[Attached file: ";

function estimateCoins(attachments) {
  let fileBytes = 0;
  let images = 0;
  let textChars = 0;
  for (const att of attachments || []) {
    if (att.error) continue; // never charge for a file that will not be sent
    if (att.category === "image") {
      images += 1;
    } else if (att.category === "pdf") {
      fileBytes += att.bytes || 0;
    } else {
      // Everything else -- .txt, .csv, and the three Office formats -- is
      // sent as extracted TEXT, and is billed on the string as it will
      // actually be sent, marker prefix and filename included. Counting the
      // bare text would quote low by the length of that prefix.
      const prefix = ATTACHED_TEXT_MARKER + att.filename + "]\n";
      // extractedText is normally present: resolveTextAttachments reads these
      // when the pane opens precisely so the quote is exact, and anything it
      // could not read carries an error and was skipped above.
      //
      // The byte fallback is for PLAIN TEXT ONLY. A UTF-8 .txt is at least as
      // many bytes as characters, so bytes can only quote HIGH. That does NOT
      // hold for the Office formats: they are ZIP archives, so a small file
      // can hold far more text than its size suggests and bytes would quote
      // LOW -- the one direction that is not allowed. An Office file with no
      // text extracted is refused by resolveTextAttachments instead.
      const fallback = att.category === "txt" ? (att.bytes || 0) : 0;
      textChars += prefix.length + (att.extractedText !== undefined
        ? att.extractedText.length
        : fallback);
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
      blocks.push({ type: "text", text: ATTACHED_TEXT_MARKER + a.filename + "]\n" + a.extractedText });
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
