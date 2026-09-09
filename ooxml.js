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
// One wording, thrown from three places, so the customer sees the same
// sentence wherever the limit is reached.
const TOO_MUCH_TEXT = "this document holds too much text to review in one go" +
  " -- send the relevant section instead";

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

// Finds <open ...>...<closer> blocks by SCANNING, never by a lazy regex.
//
// /<row[\s\S]*?<\/row>/g looks equivalent and is quadratic: every opening
// token that never closes makes the engine scan to the end of the string
// before failing and advancing to the next one. Measured 2026-09-09 on a
// 576-byte .xlsx holding 100,000 unclosed "<row" tokens: 5.9 seconds, rising
// fourfold per doubling. The byte caps do not help -- one repeated 4-byte
// token compresses to almost nothing.
//
// Here each match advances the cursor past its closer, so the regions
// searched never overlap and the whole pass is linear. A missing closer ends
// the scan: nothing after it can form a complete block either.
// Finds <name ...>...</name> and <name ... /> elements by parsing the
// OPENING TAG, then looking for exactly one closer.
//
// The first version searched for several possible closers and kept whichever
// came first, which was wrong twice over:
//
//   WRONG ANSWERS. "/>" matched the end of any self-closing CHILD, so
//   <c r="B3"><f t="shared" si="0"/><v>30</v></c> -- what Excel writes for
//   every filled-down formula cell after the first -- was cut at the <f/>
//   and its value silently vanished. A whole formula column read as blank
//   while the customer paid for a review that looked complete. Rich inline
//   strings (<b/> inside <is>) lost their text the same way.
//
//   QUADRATIC. A closer that never occurs is searched for from every
//   element, scanning to the end each time. Measured on "<c />" repeated:
//   32,000 cells took 2,467ms, quadrupling per doubling. A 96KB .xlsx
//   inflating to the permitted 64MB works out at days.
//
// Reading the opening tag settles both: self-closing is decided by the tag
// itself rather than by a "/>" that might belong to a child, and there is
// only ever one closer to look for.
function _elements(text, name) {
  const open = "<" + name;
  const close = "</" + name + ">";
  const out = [];
  let pos = 0;
  while (pos < text.length) {
    const start = text.indexOf(open, pos);
    if (start === -1) break;
    // "<c" must not match "<cols>" or "<color>". Any XML whitespace counts,
    // not just a space: Excel and third-party writers both emit tags broken
    // across lines.
    const after = text[start + open.length];
    if (after !== undefined && !/[\s>/]/.test(after)) {
      pos = start + open.length;
      continue;
    }
    const tagEnd = text.indexOf(">", start);
    if (tagEnd === -1) break;                              // unterminated tag ends the scan
    if (text[tagEnd - 1] === "/") {                        // <c r="A1"/> -- no children
      out.push(text.slice(start, tagEnd + 1));
      pos = tagEnd + 1;
      continue;
    }
    const end = text.indexOf(close, tagEnd + 1);
    if (end === -1) break;                                 // unclosed: none can follow either
    out.push(text.slice(start, end + close.length));
    pos = end + close.length;
  }
  return out;
}

// Removes whole <name>...</name> subtrees, contents included, in one linear
// pass. Same scanning shape as _elements, so it cannot backtrack.
function _removeElements(text, name) {
  const open = "<" + name;
  const close = "</" + name + ">";
  const parts = [];
  let pos = 0;
  while (pos < text.length) {
    const start = text.indexOf(open, pos);
    if (start === -1) break;
    // "<w:del" must not match "<w:delText", so the next character decides.
    const after = text[start + open.length];
    if (after !== undefined && !/[\s>/]/.test(after)) {
      parts.push(text.slice(pos, start + open.length));
      pos = start + open.length;
      continue;
    }
    const tagEnd = text.indexOf(">", start);
    if (tagEnd === -1) break;
    parts.push(text.slice(pos, start));                    // everything before it stays
    if (text[tagEnd - 1] === "/") {                        // self-closing: drop the tag
      pos = tagEnd + 1;
      continue;
    }
    const end = text.indexOf(close, tagEnd + 1);
    if (end === -1) {                                      // unclosed: drop the tag only
      pos = tagEnd + 1;
      continue;
    }
    pos = end + close.length;                              // drop the whole subtree
  }
  parts.push(text.slice(pos));
  return parts.join("");
}

// Strips XML tags in one linear pass. /<[^>]+>/g has the same quadratic
// shape as the lazy scans above when the input carries many "<" and no ">",
// so tag removal is done by hand too. An unterminated tag ends the document
// rather than being treated as text.
function _stripTags(text) {
  const parts = [];
  let i = 0;
  while (i < text.length) {
    const lt = text.indexOf("<", i);
    if (lt === -1) {
      parts.push(text.slice(i));
      break;
    }
    parts.push(text.slice(i, lt));
    const gt = text.indexOf(">", lt + 1);
    if (gt === -1) break;
    i = gt + 1;
  }
  return parts.join("");
}

function _decodeEntities(text) {
  return text
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    // fromCodePoint THROWS RangeError above 0x10FFFF, and the number comes
    // straight out of the attachment -- "&#99999999;" would kill the whole
    // file rather than one character. Out-of-range entities are left as
    // written instead.
    .replace(/&#(\d+);/g, (whole, d) => _codePoint(Number(d), whole))
    .replace(/&#x([0-9a-fA-F]+);/g, (whole, h) => _codePoint(parseInt(h, 16), whole))
    .replace(/&amp;/g, "&");                               // last, or it double-decodes
}

function _codePoint(value, original) {
  if (!Number.isFinite(value) || value < 0 || value > 0x10ffff) return original;
  return String.fromCodePoint(value);
}

// Reads the text of ONE paragraph by collecting only the DESIGNATED text
// elements -- <w:t> in Word, <a:t> in PowerPoint -- plus the tab and break
// markers, in document order.
//
// This is an ALLOWLIST, and that is the whole point. Stripping tags and
// keeping whatever text remained was a blacklist against a specification
// with hundreds of elements, and each audit round found another that had to
// be excluded: tracked deletions (w:delText), field codes (w:instrText),
// moved text, animation internals (p:attrName -> "style.visibility",
// "ppt_x"), phonetic guides. Every one of those was text the application
// does NOT show, handed to the model as document content. Naming the two
// elements that genuinely carry prose ends that search: anything not on the
// list is excluded because it was never included.
//
// Linear -- indexOf from a cursor that only moves forward.
function _textRuns(fragment, textTag, breakTags) {
  const out = [];
  const closeTag = "</" + textTag + ">";
  let pos = 0;
  while (pos < fragment.length) {
    const lt = fragment.indexOf("<", pos);
    if (lt === -1) break;
    const gt = fragment.indexOf(">", lt);
    if (gt === -1) break;
    const inside = fragment.slice(lt + 1, gt);
    const name = inside.split(/[\s/>]/)[0];
    if (name === textTag && inside[inside.length - 1] !== "/") {
      const close = fragment.indexOf(closeTag, gt);
      if (close === -1) break;
      out.push(_decodeEntities(fragment.slice(gt + 1, close)));
      pos = close + closeTag.length;
      continue;
    }
    if (breakTags.tab && name === breakTags.tab) out.push("\t");
    if (breakTags.br && name === breakTags.br) out.push("\n");
    pos = gt + 1;
  }
  return out.join("");
}

// Every paragraph in a part, as lines. Paragraphs with no text drop out.
function _paragraphs(xml, paraTag, textTag, breakTags) {
  const lines = [];
  for (const para of _elements(xml, paraTag)) {
    const text = _textRuns(para, textTag, breakTags).trimEnd();
    if (text) lines.push(text);
  }
  return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

// Turns a fragment of OOXML into readable text: the given tags become line
// breaks, everything else is dropped, entities are decoded.
function _xmlToText(xml, breakTags) {
  let out = xml;
  for (const tag of breakTags) {
    out = out.split("</" + tag + ">").join("\n");          // literal, so no regex build
  }
  out = out.split("<w:tab/>").join("\t").split("<w:br/>").join("\n");
  out = _stripTags(out);                                   // every remaining tag
  // Trailing whitespace is trimmed line by line with trimEnd(), NOT with
  // /[ \t]+\n/g. That pattern backtracks across every start position in a run
  // of spaces not followed by a newline: measured at 4.5 seconds for a
  // 296-byte .docx, quadrupling per doubling, which at the size the inflate
  // cap allows never returns. trimEnd is linear.
  return _decodeEntities(out)
    .split("\n").map((line) => line.trimEnd()).join("\n")
    .replace(/\n{3,}/g, "\n\n")                            // fixed char, cannot backtrack
    .trim();
}

// ---- the three formats --------------------------------------------------

async function _extractDocx(bytes, entries, budget) {
  const xml = await _readText(bytes, entries, "word/document.xml", budget);
  if (!xml) throw new Error("no word/document.xml -- not a Word file");
  // TRACKED DELETIONS ARE NOT DOCUMENT TEXT. Word stores removed text in
  // <w:delText> inside <w:del>, as a sibling of the <w:ins> that replaced
  // it, and stripping tags indiscriminately ran the two together:
  // "Payment due within 30 days" edited to 90 came out as "Payment due
  // within 3090 days", and a struck-out clause came out as binding text.
  // Word's default Simple Markup view HIDES deletions, so the sender sees
  // the clean document while the model is fed one that never existed --
  // which for a contract is the reason this feature exists.
  //
  // <w:instrText> goes too: field instruction codes (HYPERLINK, MERGEFIELD)
  // are machinery, not prose, and read as noise in the middle of a sentence.
  //
  // A tracked MOVE is not a delete plus an insert: Word writes a matched
  // <w:moveFrom>/<w:moveTo> pair, and the moveFrom source survives a w:del
  // sweep untouched. Left in, a clause dragged from one section to another
  // appears TWICE -- once where it no longer belongs -- while Word shows it
  // once. <w:moveTo> is kept, exactly as <w:ins> is: that is where the text
  // now lives.
  // <w:del> and <w:moveFrom> are still removed as SUBTREES, because both can
  // hold ordinary <w:t> runs and the allowlist alone would keep them.
  // <mc:Fallback> too: Word writes every text box twice, once as a modern
  // <mc:Choice> and once as a VML fallback for old readers, and it renders
  // only the Choice -- so keeping both handed the model each clause twice
  // AND charged the customer for the second copy.
  //
  // w:instrText no longer needs a sweep: it is not <w:t>, so the allowlist
  // never picks it up. The list below is now only what the allowlist cannot
  // settle by itself.
  let cleaned = _removeElements(xml, "w:del");
  cleaned = _removeElements(cleaned, "w:moveFrom");
  cleaned = _removeElements(cleaned, "mc:Fallback");
  return _paragraphs(cleaned, "w:p", "w:t", { tab: "w:tab", br: "w:br" });
}

async function _extractPptx(bytes, entries, budget) {
  // Slide parts are named slide1.xml, slide2.xml ... and sort WRONGLY as
  // text once past nine, so they are ordered by their number.
  const slides = [...entries.keys()]
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));
  const parts = [];
  let produced = 0;
  for (let i = 0; i < slides.length; i++) {
    // <a:t> only. The whole slide part used to be scanned, which swept in
    // <p:timing> -- so any entrance animation leaked "style.visibility" and
    // "ppt_x" into the slide's text as though the deck said it.
    const text = _paragraphs(await _readText(bytes, entries, slides[i], budget),
                             "a:p", "a:t", { br: "a:br" });
    if (text) {
      produced += text.length;                             // checked as it grows, as in _extractXlsx
      if (produced > MAX_EXTRACTED_CHARS) throw new Error(TOO_MUCH_TEXT);
      parts.push("--- Slide " + (i + 1) + " ---\n" + text);
    }
  }
  if (parts.length === 0) throw new Error("no slides found -- not a PowerPoint file");
  return parts.join("\n\n");
}

async function _extractXlsx(bytes, entries, budget) {
  // Cell text is stored ONCE in a shared table and referenced by index, so
  // that table has to be read before any sheet makes sense.
  const sharedXml = await _readText(bytes, entries, "xl/sharedStrings.xml", budget);
  const shared = [];
  for (const si of _elements(sharedXml, "si")) {
    // <rPh> holds the phonetic reading Excel stores beside a CJK value and
    // does not display. Keeping it appended the furigana straight onto the
    // word with no separator, inventing a term the sheet never contained.
    shared.push(_textRuns(_removeElements(si, "rPh"), "t", {}));
  }

  // Sheet NAMES live in workbook.xml while the CONTENT lives in numbered
  // files. Pairing them properly means resolving relationship ids; the names
  // are only a label here, so they are matched in document order and fall
  // back to "Sheet N" whenever the counts disagree.
  const workbook = await _readText(bytes, entries, "xl/workbook.xml", budget);
  // Each <sheet/> tag is isolated first, then its name read from that short
  // string. Running /<sheet[^>]*name="..."/g over the whole part was
  // quadratic on a workbook.xml full of "<sheet" with no ">" -- measured at
  // 2.3 seconds for a 616-byte .xlsx. A regex bounded to one tag cannot run
  // away, however hostile the tag is.
  const names = _elements(workbook, "sheet")
    .map((tag) => (tag.match(/\sname="([^"]*)"/) || [])[1])
    .map((name) => (name === undefined ? "" : _decodeEntities(name)));

  const sheets = [...entries.keys()]
    .filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));

  const parts = [];
  let produced = 0;                                        // chars emitted so far
  for (let i = 0; i < sheets.length; i++) {
    const xml = await _readText(bytes, entries, sheets[i], budget);
    const rows = [];
    for (const row of _elements(xml, "row")) {
      const cells = [];
      // "<c" also begins <cols> and <color>, so only a space or ">" counts.
      for (const cell of _elements(row, "c")) {
        const type = (cell.match(/\st="([^"]*)"/) || [])[1];
        const held = _elements(cell, "v");
        const value = held.length ? _stripTags(held[0]) : undefined;
        if (type === "s") {
          // An EMPTY <v> is an empty cell, not index 0 -- Number("") is 0,
          // so a blank shared-string cell used to print whatever text
          // happened to sit first in the shared table. An index that is
          // absent, NaN or out of range is likewise an empty cell rather
          // than "undefined".
          const index = value === undefined || value === "" ? -1 : Number(value);
          cells.push(Number.isInteger(index) && index >= 0 ? (shared[index] || "") : "");
        } else if (type === "inlineStr") {
          cells.push(_textRuns(_removeElements(cell, "rPh"), "t", {}));
        } else {
          cells.push(value === undefined ? "" : _decodeEntities(value));
        }
      }
      // A row of nothing but empty cells carries no information and would
      // otherwise pad a big sheet with blank lines the model has to read.
      if (cells.some((c) => c !== "")) {
        // Measured from the PARTS, before joining them. The previous version
        // checked line.length -- after cells.join() had already allocated the
        // whole amplified row, which is the one allocation the check exists
        // to prevent. Pushing shared-string references costs nothing; the
        // blow-up happens entirely inside the join.
        let width = cells.length;                          // the tab separators
        for (const cell of cells) width += cell.length;
        if (produced + width > MAX_EXTRACTED_CHARS) throw new Error(TOO_MUCH_TEXT);
        const line = cells.join("\t");
        // Checked AS IT GROWS, not once at the end. A sheet is a list of
        // REFERENCES into the shared-string table, so a small part can point
        // at one long string thousands of times and amplify far past what the
        // inflate cap allows -- and the final check cannot help if the string
        // it is meant to measure has already exhausted memory.
        produced += width + 1;
        rows.push(line);
      }
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
  if (text.length > MAX_EXTRACTED_CHARS) throw new Error(TOO_MUCH_TEXT);
  return text;
}
