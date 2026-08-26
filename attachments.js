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
const MAX_EXTRACTED_TEXT_CHARS = 200000; // truncation ceiling

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
      return { filename: file.name, category, contentBlock: { type: "image_url", image_url: { url: dataUri } } };
    }
    if (category === "pdf") {
      const dataUri = await _readAsDataUrl(file);
      return { filename: file.name, category, contentBlock: { type: "file", file: { filename: file.name, file_data: dataUri } } };
    }
    if (category === "txt") {
      let text = await _readAsText(file);
      if (text.length > MAX_EXTRACTED_TEXT_CHARS) {
        text = text.slice(0, MAX_EXTRACTED_TEXT_CHARS) + "\n...[truncated -- file was longer than " + MAX_EXTRACTED_TEXT_CHARS + " characters]";
      }
      return { filename: file.name, category, extractedText: text };
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
