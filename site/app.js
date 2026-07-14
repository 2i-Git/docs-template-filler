/*
 * Docs Filler — client-side core.
 *
 * A faithful JavaScript port of docs_filler/core.py that runs entirely in the
 * browser (no server, no upload). The same file also loads under Node with
 * `require`, which is how the parity test drives it against the Python output.
 *
 * Placeholder convention: anywhere in the template, `{{Column Header}}` is
 * replaced by that column's value for the current row. Matching ignores
 * capitalisation and extra spaces — identical to the Python tool.
 *
 * Dependencies (vendored, no CDN): JSZip (read/write the .docx/.zip) and
 * SheetJS/XLSX (parse the .xlsx).
 */
(function (global, factory) {
  if (typeof module === "object" && module.exports) {
    // Node (parity test): pull the vendored UMD builds.
    module.exports = factory(
      require("./vendor/jszip.min.js"),
      require("./vendor/xlsx.full.min.js")
    );
  } else {
    global.DocsFiller = factory(global.JSZip, global.XLSX);
  }
})(typeof self !== "undefined" ? self : this, function (JSZip, XLSX) {
  "use strict";

  // Placeholders look like {{ Header Name }} — spaces inside the braces are ok.
  // (Built fresh per use because we rely on the global flag / lastIndex.)
  function placeholderRegex() {
    return /\{\{\s*(.+?)\s*\}\}/g;
  }

  /** Raised for user-fixable problems (bad spreadsheet, unknown column, …). */
  class DocsFillerError extends Error {
    constructor(message) {
      super(message);
      this.name = "DocsFillerError";
    }
  }

  /** Lower-case and collapse whitespace, for tolerant header matching. */
  function normalise(text) {
    return String(text).trim().replace(/\s+/g, " ").toLowerCase();
  }

  // ---------------------------------------------------------------------------
  // Value formatting  (mirror core.format_value)
  // ---------------------------------------------------------------------------

  function groupInteger(n) {
    const neg = n < 0;
    const s = Math.abs(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return (neg ? "-" : "") + s;
  }

  function groupFixed2(value) {
    // Mirror Python f"{round(value, 2):,.2f}". For values that aren't on an
    // exact half-cent tie (all real currency amounts computed to 2dp), toFixed
    // and Python's round-half-to-even agree.
    const neg = value < 0;
    const [intPart, frac] = Math.abs(value).toFixed(2).split(".");
    const grouped = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return (neg ? "-" : "") + grouped + "." + frac;
  }

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function formatValue(value) {
    if (value === null || value === undefined) return "";
    if (typeof value === "boolean") return value ? "Yes" : "No";
    if (typeof value === "number") {
      if (!isFinite(value)) return String(value);
      if (Number.isInteger(value)) return groupInteger(value);
      return groupFixed2(value);
    }
    if (value instanceof Date) {
      // openpyxl returns datetime/date; str() gives "YYYY-MM-DD[ HH:MM:SS]".
      // Date-formatted cells (no time) render without the time component.
      const d =
        pad2(value.getFullYear()).padStart(4, "0") +
        "-" + pad2(value.getMonth() + 1) +
        "-" + pad2(value.getDate());
      const hasTime =
        value.getHours() || value.getMinutes() || value.getSeconds();
      if (!hasTime) return d;
      return d + " " + pad2(value.getHours()) + ":" +
        pad2(value.getMinutes()) + ":" + pad2(value.getSeconds());
    }
    return String(value);
  }

  // ---------------------------------------------------------------------------
  // Spreadsheet reading  (mirror core.read_records)
  // ---------------------------------------------------------------------------

  function readRecords(xlsxBytes, sheet) {
    let wb;
    try {
      wb = XLSX.read(toUint8(xlsxBytes), { type: "array", cellDates: true });
    } catch (err) {
      throw new DocsFillerError(
        "Could not read the spreadsheet. Please upload a valid .xlsx file."
      );
    }

    if (sheet && !Object.prototype.hasOwnProperty.call(wb.Sheets, sheet)) {
      throw new DocsFillerError(
        "Worksheet '" + sheet + "' was not found in the spreadsheet."
      );
    }
    const wsName = sheet || wb.SheetNames[0];
    const ws = wb.Sheets[wsName];

    const grid = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      raw: true,
      blankrows: false,
      defval: null,
    });

    // Keep rows with at least one non-empty cell (matches the Python filter).
    const rows = grid.filter((r) =>
      r.some((c) => c !== null && c !== undefined && String(c).trim() !== "")
    );
    if (!rows.length) {
      throw new DocsFillerError("The spreadsheet is empty.");
    }

    const headerRow = rows[0];
    const headers = [];
    const indexByKey = {};
    headerRow.forEach((cell, idx) => {
      if (cell === null || cell === undefined || String(cell).trim() === "") {
        return;
      }
      const name = String(cell).trim();
      headers.push(name);
      indexByKey[normalise(name)] = idx;
    });
    if (!headers.length) {
      throw new DocsFillerError(
        "Could not find any column headers in the first row of the spreadsheet."
      );
    }

    const records = rows.slice(1).map((row) => {
      const record = {};
      for (const key in indexByKey) {
        const idx = indexByKey[key];
        record[key] = idx < row.length ? row[idx] : null;
      }
      return record;
    });
    if (!records.length) {
      throw new DocsFillerError(
        "The spreadsheet has headers but no data rows."
      );
    }

    return { headers: headers, records: records };
  }

  // ---------------------------------------------------------------------------
  // DOCX replacement (run-aware, so text formatting is preserved)
  // ---------------------------------------------------------------------------

  function xmlUnescape(s) {
    return s
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, "&");
  }

  function xmlEscape(s) {
    return s
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  /**
   * Replace all {{...}} in one paragraph while keeping run formatting.
   *
   * A placeholder may be split across several <w:t> nodes (python-docx "runs"),
   * so we work on the joined text and write the result back across the original
   * nodes — the same right-to-left, offset-preserving algorithm as
   * core.fill_paragraph.
   */
  function fillParagraph(paragraphXml, resolve, missing) {
    const tRe = /<w:t\b([^>]*)>([\s\S]*?)<\/w:t>/g;
    const nodes = [];
    let m;
    while ((m = tRe.exec(paragraphXml)) !== null) {
      nodes.push({
        start: m.index,
        end: tRe.lastIndex,
        attrs: m[1],
        inner: m[2],
      });
    }
    if (!nodes.length) return paragraphXml;

    const texts = nodes.map((n) => xmlUnescape(n.inner));
    const full = texts.join("");

    const phRe = placeholderRegex();
    const matches = [];
    let mm;
    while ((mm = phRe.exec(full)) !== null) {
      matches.push({ start: mm.index, end: mm.index + mm[0].length, name: mm[1] });
      if (mm.index === phRe.lastIndex) phRe.lastIndex++; // guard (shouldn't happen)
    }
    if (!matches.length) return paragraphXml;

    const bounds = [];
    let pos = 0;
    for (const t of texts) {
      bounds.push([pos, pos + t.length]);
      pos += t.length;
    }
    const total = pos;

    function locate(p) {
      for (let i = 0; i < bounds.length; i++) {
        const s = bounds[i][0];
        const e = bounds[i][1];
        if (s <= p && p < e) return [i, p - s];
      }
      const last = bounds.length - 1;
      return [last, bounds[last][1] - bounds[last][0]];
    }

    let changed = false;
    // Right-to-left so earlier offsets remain valid as we edit.
    for (let k = matches.length - 1; k >= 0; k--) {
      const mt = matches[k];
      const res = resolve(mt.name);
      if (!res.found) {
        missing.add(mt.name.trim());
        continue;
      }
      changed = true;
      const replacement = res.value;
      const startLoc = locate(mt.start);
      const si = startLoc[0];
      const so = startLoc[1];
      let ei, eo;
      if (mt.end >= total) {
        ei = texts.length - 1;
        eo = texts[texts.length - 1].length;
      } else {
        const endLoc = locate(mt.end);
        ei = endLoc[0];
        eo = endLoc[1];
      }
      if (si === ei) {
        texts[si] = texts[si].slice(0, so) + replacement + texts[si].slice(eo);
      } else {
        texts[si] = texts[si].slice(0, so) + replacement;
        for (let j = si + 1; j < ei; j++) texts[j] = "";
        texts[ei] = texts[ei].slice(eo);
      }
    }
    if (!changed) return paragraphXml;

    // Write the edited texts back into the original <w:t> nodes, preserving
    // whitespace (add xml:space="preserve" if the node lacks it).
    let out = "";
    let cursor = 0;
    for (let i = 0; i < nodes.length; i++) {
      const n = nodes[i];
      out += paragraphXml.slice(cursor, n.start);
      let attrs = n.attrs;
      if (!/xml:space\s*=/.test(attrs)) attrs += ' xml:space="preserve"';
      out += "<w:t" + attrs + ">" + xmlEscape(texts[i]) + "</w:t>";
      cursor = n.end;
    }
    out += paragraphXml.slice(cursor);
    return out;
  }

  /** Fill every paragraph (<w:p>) in one XML part. */
  function fillPart(xml, resolve, missing) {
    return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, function (p) {
      return fillParagraph(p, resolve, missing);
    });
  }

  // The document body plus every header/footer — mirrors core.iter_paragraphs
  // (body + tables + headers/footers). Tables live inside these parts already.
  function isFillablePart(path) {
    return (
      path === "word/document.xml" ||
      /^word\/header\d*\.xml$/.test(path) ||
      /^word\/footer\d*\.xml$/.test(path)
    );
  }

  async function fillDocument(templateBytes, record, missing) {
    const zip = await JSZip.loadAsync(toUint8(templateBytes));

    function resolve(placeholderName) {
      const key = normalise(placeholderName);
      if (Object.prototype.hasOwnProperty.call(record, key)) {
        return { found: true, value: formatValue(record[key]) };
      }
      return { found: false, value: "" };
    }

    const paths = Object.keys(zip.files).filter(isFillablePart);
    for (const path of paths) {
      const xml = await zip.file(path).async("string");
      zip.file(path, fillPart(xml, resolve, missing));
    }
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  // ---------------------------------------------------------------------------
  // Orchestration  (mirror core.build_documents / generate_zip)
  // ---------------------------------------------------------------------------

  function safeFilename(text, fallback) {
    // Python: re.sub(r"[^\w\-. ]+", "_", str(text).strip()).strip(". ")
    // \w is Unicode in Python, so keep letters/digits/underscore of any script.
    const cleaned = String(text)
      .trim()
      .replace(/[^\p{L}\p{N}_\-. ]+/gu, "_")
      .replace(/^[. ]+|[. ]+$/g, "");
    return cleaned || fallback;
  }

  function stem(filename) {
    const base = String(filename || "").split(/[\\/]/).pop();
    return base.replace(/\.[^.]+$/, "") || base;
  }

  async function buildDocuments(templateBytes, xlsxBytes, opts) {
    opts = opts || {};
    const nameColumn = opts.name_column || null;
    const sheet = opts.sheet || null;
    const baseName = opts.base_name || "document";
    const progress = opts.progress || null;

    const read = readRecords(xlsxBytes, sheet);
    const headers = read.headers;
    const records = read.records;

    // Which column drives each output filename.
    let nameKey;
    if (nameColumn) {
      nameKey = normalise(nameColumn);
      const known = new Set(headers.map(normalise));
      if (!known.has(nameKey)) {
        throw new DocsFillerError(
          "Column '" + nameColumn + "' was not found. Available columns: " +
          headers.join(", ")
        );
      }
    } else {
      nameKey = normalise(headers[0]);
    }

    const total = records.length;
    const documents = [];
    const allMissing = new Set();
    const seenNames = {};

    for (let i = 0; i < records.length; i++) {
      const record = records[i];
      const missing = new Set();
      const bytes = await fillDocument(templateBytes, record, missing);
      missing.forEach((x) => allMissing.add(x));

      const base = safeFilename(formatValue(record[nameKey]), "row_" + (i + 1));
      let filename = baseName + " - " + base;
      const count = (seenNames[filename] || 0) + 1;
      seenNames[filename] = count;
      if (count > 1) filename = filename + " (" + count + ")";

      documents.push({ filename: filename + ".docx", bytes: bytes });

      if (progress) progress(i + 1, total);
    }

    return { documents: documents, missing: Array.from(allMissing).sort() };
  }

  async function generateZip(templateBytes, xlsxBytes, opts) {
    const built = await buildDocuments(templateBytes, xlsxBytes, opts);
    const zip = new JSZip();
    for (const doc of built.documents) {
      zip.file(doc.filename, doc.bytes);
    }
    const zipBytes = await zip.generateAsync({
      type: "uint8array",
      compression: "DEFLATE",
    });
    return { zipBytes: zipBytes, missing: built.missing };
  }

  // ---------------------------------------------------------------------------

  function toUint8(data) {
    if (data instanceof Uint8Array) return data;
    if (typeof ArrayBuffer !== "undefined" && data instanceof ArrayBuffer) {
      return new Uint8Array(data);
    }
    // Node Buffer is already a Uint8Array subclass; anything else, try to wrap.
    return new Uint8Array(data);
  }

  return {
    DocsFillerError: DocsFillerError,
    normalise: normalise,
    formatValue: formatValue,
    readRecords: readRecords,
    fillPart: fillPart,
    buildDocuments: buildDocuments,
    generateZip: generateZip,
  };
});
