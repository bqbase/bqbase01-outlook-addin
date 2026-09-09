/*
============================================================================
 ooxml.js -- text out of .docx / .xlsx / .pptx, with no library
============================================================================

 Word, Excel and PowerPoint files are ZIP archives of XML. That makes them
 readable in the pane without pulling in a parsing library: read the ZIP
 directory, inflate the parts that hold text, strip the tags.

 Inflate comes from DecompressionStream("deflate-raw"), which is built into
 the browser engine. There is no bundled decompressor here and there should
 not be -- everything in outlook_addin/ is served from a public GitHub Pages
 site and every added kilobyte is downloaded on every pane open.

 EXTRACTS TEXT ONLY. A chart or picture inside a document is a separate image
 part in the archive and comes through as nothing. That is a deliberate
 limit, not an oversight: pulling embedded images out and sending them would
 silently multiply what a picture-heavy document costs to review.

 The legacy binary formats -- .doc, .xls, .ppt -- are NOT ZIPs and cannot be
 read this way. They get a clear per-file error from attachments.js.
============================================================================
*/

// ---- ZIP ----------------------------------------------------------------
//
// Only the parts of the format this needs: find the central directory, walk
// its entries, and inflate the few whose names matter. Entries are read from
// the CENTRAL directory rather than by scanning for local headers, because a
// local header may declare sizes of zero and defer them to a trailing data
// descriptor -- which is exactly what several producers of these files do.

const _ZIP_EOCD = 0x06054b50;
const _ZIP_CENTRAL = 0x02014b50;

// An attachment is data from a stranger, so nothing about the archive is
// trusted: not its sizes, not its offsets, not its entry count.
//
// The size caps matter most. MAX_ATTACHMENT_BYTES bounds the file at 20MB
// COMPRESSED, which bounds the decompressed size not at all -- a deflate
// stream of zeros expands about a thousandfold, so a 1MB attachment can
// become a gigabyte and take the pane down with it. Inflation is therefore
// aborted mid-stream at the cap rather than measured afterwards, which would
// mean allocating the whole thing first.
const MAX_INFLATED_BYTES = 64 * 1024 * 1024;   // one part, e.g. a big sheet's XML
const MAX_TOTAL_INFLATED = 128 * 1024 * 1024;  // everything read from one archive
const MAX_ENTRIES = 5000;                      // a real .pptx has dozens
// Past this the file is REFUSED, not truncated. A silently shortened
// contract is worse than a clear refusal: the review would read as complete
// while missing whatever came after the cut.
const MAX_EXTRACTED_CHARS = 2 * 1000 * 1000;

function _findEocd(view) {
  // The end-of-central-directory record sits at the very end unless the file
  // carries a trailing comment, so it is searched for backwards. 22 is its
  // fixed size; the comment may add up to 65535.
  const min = Math.max(0, view.byteLength - 22 - 65535);
  for (let i = view.byteLength - 22; i >= min; i--) {
    if (view.getUint32(i, true) === _ZIP_EOCD) return i;
  }
  return -1;
}

// Returns a Map of entry name -> {offset, compressedSize, method}. Nothing is
// inflated here: a .docx carries dozens of parts and only two or three are
// ever wanted.
function _readDirectory(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = _findEocd(view);
  if (eocd === -1) throw new Error("not a ZIP archive (no end-of-directory record)");
  const count = Math.min(view.getUint16(eocd + 10, true), MAX_ENTRIES);
  let pointer = view.getUint32(eocd + 16, true);
  const entries = new Map();
  for (let i = 0; i < count; i++) {
    // Every read below is bounds-checked BEFORE it happens. A DataView throws
    // on an out-of-range read, which would be caught upstream and reported as
    // an unreadable file -- but a truncated or hostile archive should be
    // named as such rather than surfacing as a stray RangeError.
    if (pointer < 0 || pointer + 46 > bytes.length) break;
    if (view.getUint32(pointer, true) !== _ZIP_CENTRAL) break;
    const method = view.getUint16(pointer + 10, true);
    const compressedSize = view.getUint32(pointer + 20, true);
    const nameLength = view.getUint16(pointer + 28, true);
    const extraLength = view.getUint16(pointer + 30, true);
    const commentLength = view.getUint16(pointer + 32, true);
    const localOffset = view.getUint32(pointer + 42, true);
    if (pointer + 46 + nameLength > bytes.length) break;
    const name = new TextDecoder("utf-8").decode(
      bytes.subarray(pointer + 46, pointer + 46 + nameLength));
    // An entry pointing outside the file is either corruption or an attempt
    // to make the reader read something else. Skipped, not fatal: the rest of
    // the archive may still hold the part that is wanted.
    if (localOffset >= 0 && localOffset + 30 <= bytes.length) {
      entries.set(name, { localOffset, compressedSize, method });
    }
    pointer += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

// Inflates one entry, ABORTING as soon as the output passes the cap rather
// than after allocating all of it. The counter is shared across an archive so
// many small bombs cost no more than one big one.
async function _inflate(raw, budget) {
  const stream = new Blob([raw]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  const reader = stream.getReader();
  const chunks = [];
  let total = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.length;
      budget.used += value.length;
      if (total > MAX_INFLATED_BYTES || budget.used > MAX_TOTAL_INFLATED) {
        await reader.cancel();
        throw new Error("this file expands to far more than it claims and was not read");
      }
      chunks.push(value);
    }
  } finally {
    reader.releaseLock();
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

async function _readEntry(bytes, entry, budget) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // The local header repeats the name and extra fields, and its extra length
  // often DIFFERS from the central directory's -- so the data offset has to
  // be computed from the local header, never from the central copy.
  const nameLength = view.getUint16(entry.localOffset + 26, true);
  const extraLength = view.getUint16(entry.localOffset + 28, true);
  const start = entry.localOffset + 30 + nameLength + extraLength;
  // Clamped rather than trusted: compressedSize comes from the archive, so a
  // hostile one can claim an entry runs past the end of the file.
  const end = Math.min(start + entry.compressedSize, bytes.length);
  if (start >= bytes.length || end <= start) return new Uint8Array(0);
  const raw = bytes.subarray(start, end);
  if (entry.method === 0) {                                // stored, not compressed
    budget.used += raw.length;
    if (budget.used > MAX_TOTAL_INFLATED) throw new Error("archive is too large to read");
    return raw;
  }
  if (entry.method !== 8) throw new Error("unsupported ZIP compression method " + entry.method);
  return _inflate(raw, budget);
}

async function _readText(bytes, entries, name, budget) {
  const entry = entries.get(name);
  if (!entry) return "";
  return new TextDecoder("utf-8").decode(await _readEntry(bytes, entry, budget));
}

// ---- XML ----------------------------------------------------------------
//
// Deliberately regex-based rather than DOMParser. These parts are large --
// a long spreadsheet's sheet XML runs to megabytes -- and building a full
// DOM to then throw it away costs time and memory the pane does not need to
// spend for a flat text dump.

function _decodeEntities(text) {
  return text
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&amp;/g, "&");                               // last, or it double-decodes
}

// Turns a fragment of OOXML into readable text: the given tags become line
// breaks, everything else is dropped, entities are decoded.
function _xmlToText(xml, breakTags) {
  let out = xml;
  for (const tag of breakTags) {
    out = out.replace(new RegExp("</" + tag + ">", "g"), "\n");
  }
  out = out.replace(/<w:tab\/>/g, "\t").replace(/<w:br\/>/g, "\n");
  out = out.replace(/<[^>]+>/g, "");                       // every remaining tag
  return _decodeEntities(out).replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}

// ---- the three formats --------------------------------------------------

async function _extractDocx(bytes, entries, budget) {
  const xml = await _readText(bytes, entries, "word/document.xml", budget);
  if (!xml) throw new Error("no word/document.xml -- not a Word file");
  return _xmlToText(xml, ["w:p"]);
}

async function _extractPptx(bytes, entries, budget) {
  // Slide parts are named slide1.xml, slide2.xml ... and sort WRONGLY as
  // text once past nine, so they are ordered by their number.
  const slides = [...entries.keys()]
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));
  const parts = [];
  for (let i = 0; i < slides.length; i++) {
    const text = _xmlToText(await _readText(bytes, entries, slides[i], budget), ["a:p"]);
    if (text) parts.push("--- Slide " + (i + 1) + " ---\n" + text);
  }
  if (parts.length === 0) throw new Error("no slides found -- not a PowerPoint file");
  return parts.join("\n\n");
}

async function _extractXlsx(bytes, entries, budget) {
  // Cell text is stored ONCE in a shared table and referenced by index, so
  // that table has to be read before any sheet makes sense.
  const sharedXml = await _readText(bytes, entries, "xl/sharedStrings.xml", budget);
  const shared = [];
  for (const si of sharedXml.match(/<si>[\s\S]*?<\/si>/g) || []) {
    shared.push(_decodeEntities(si.replace(/<[^>]+>/g, "")));
  }

  // Sheet NAMES live in workbook.xml while the CONTENT lives in numbered
  // files. Pairing them properly means resolving relationship ids; the names
  // are only a label here, so they are matched in document order and fall
  // back to "Sheet N" whenever the counts disagree.
  const workbook = await _readText(bytes, entries, "xl/workbook.xml", budget);
  const names = [...workbook.matchAll(/<sheet[^>]*name="([^"]*)"/g)].map((m) => _decodeEntities(m[1]));

  const sheets = [...entries.keys()]
    .filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));

  const parts = [];
  for (let i = 0; i < sheets.length; i++) {
    const xml = await _readText(bytes, entries, sheets[i], budget);
    const rows = [];
    for (const row of xml.match(/<row[\s\S]*?<\/row>/g) || []) {
      const cells = [];
      for (const cell of row.match(/<c[ >][\s\S]*?(?:<\/c>|\/>)/g) || []) {
        const type = (cell.match(/\st="([^"]*)"/) || [])[1];
        const value = (cell.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
        if (type === "s") {
          cells.push(shared[Number(value)] || "");
        } else if (type === "inlineStr") {
          cells.push(_decodeEntities((cell.match(/<t[^>]*>([\s\S]*?)<\/t>/) || [])[1] || ""));
        } else {
          cells.push(value === undefined ? "" : _decodeEntities(value));
        }
      }
      // A row of nothing but empty cells carries no information and would
      // otherwise pad a big sheet with blank lines the model has to read.
      if (cells.some((c) => c !== "")) rows.push(cells.join("\t"));
    }
    if (rows.length) {
      parts.push("--- " + (names[i] || "Sheet " + (i + 1)) + " ---\n" + rows.join("\n"));
    }
  }
  if (parts.length === 0) throw new Error("no sheet data found -- not an Excel file");
  return parts.join("\n\n");
}

// ---- entry point --------------------------------------------------------

// True when this browser can inflate at all. Feature-detected because the
// whole approach rests on it, and a missing DecompressionStream should read
// as "this host cannot open Word files" rather than as a crash.
function ooxmlSupported() {
  return typeof DecompressionStream === "function" && typeof Blob === "function";
}

// extractOfficeText(bytes, category) -> Promise<string>
// category is "docx" | "xlsx" | "pptx". Throws with a readable reason, which
// attachments.js turns into a per-file error.
async function extractOfficeText(bytes, category) {
  if (!ooxmlSupported()) {
    throw new Error("this version of Outlook cannot open Office files in the pane");
  }
  const entries = _readDirectory(bytes);
  // One shared budget for the whole archive, so many small bombs cost no more
  // than one big one.
  const budget = { used: 0 };
  let text;
  if (category === "docx") text = await _extractDocx(bytes, entries, budget);
  else if (category === "xlsx") text = await _extractXlsx(bytes, entries, budget);
  else if (category === "pptx") text = await _extractPptx(bytes, entries, budget);
  else throw new Error("unsupported Office format: " + category);

  // REFUSED rather than truncated. A shortened contract would be reviewed as
  // though it were complete, and the reader would have no way to tell -- the
  // same reason the .txt truncation limit was removed. It also keeps the
  // request under the Worker's 30MB body cap, which extracted text could
  // otherwise blow through AFTER the customer had agreed a price.
  if (text.length > MAX_EXTRACTED_CHARS) {
    throw new Error("this document holds about " +
      Math.round(text.length / 1800) + " pages of text, too much to review in one go" +
      " -- send the relevant section instead");
  }
  return text;
}
