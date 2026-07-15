/*
 * Unit tests for the Docs Filler browser core (site/app.js).
 *
 * Uses only the Node built-in test runner and the vendored libs the site
 * already ships — there is no install step. Run with:
 *
 *   node --test tests/
 */
const test = require("node:test");
const assert = require("node:assert");

const JSZip = require("../site/vendor/jszip.min.js");
const XLSX = require("../site/vendor/xlsx.full.min.js");
const DocsFiller = require("../site/app.js");

// ---------------------------------------------------------------------------
// Helpers for building fixtures in memory
// ---------------------------------------------------------------------------

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** One paragraph, one run. */
function para(text) {
  return "<w:p><w:r><w:t>" + text + "</w:t></w:r></w:p>";
}

/** One paragraph whose text is split across several runs (as Word often does). */
function splitPara(parts) {
  return (
    "<w:p>" +
    parts.map((t) => "<w:r><w:t>" + t + "</w:t></w:r>").join("") +
    "</w:p>"
  );
}

function documentXml(body) {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="' + W_NS + '"><w:body>' + body + "</w:body></w:document>"
  );
}

/** Build a minimal but structurally valid .docx as a Buffer. */
async function makeDocx(body, extraParts) {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
  );
  zip.file("word/document.xml", documentXml(body));
  for (const path in extraParts || {}) {
    zip.file(path, documentXml(extraParts[path]));
  }
  return zip.generateAsync({ type: "nodebuffer" });
}

/** Build an .xlsx from an array-of-arrays (first row = headers). */
function makeXlsx(rows, sheetName) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet(rows),
    sheetName || "Sheet1"
  );
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

/** Mirrors the resolve() closure inside fillDocument, for fillPart tests. */
function resolverFor(record) {
  const norm = {};
  for (const k in record) norm[DocsFiller.normalise(k)] = record[k];
  return (name) => {
    const key = DocsFiller.normalise(name);
    return Object.prototype.hasOwnProperty.call(norm, key)
      ? { found: true, value: DocsFiller.formatValue(norm[key]) }
      : { found: false, value: "" };
  };
}

/** Concatenated visible text of a filled XML part. */
function textOf(xml) {
  const out = [];
  const re = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = re.exec(xml)) !== null) out.push(m[1]);
  return out
    .join("")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

async function docTextInZip(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  return textOf(await zip.file("word/document.xml").async("string"));
}

// ---------------------------------------------------------------------------
// normalise
// ---------------------------------------------------------------------------

test("normalise lower-cases, trims, and collapses whitespace", () => {
  assert.equal(DocsFiller.normalise("First Name"), "first name");
  assert.equal(DocsFiller.normalise("  First   Name  "), "first name");
  assert.equal(DocsFiller.normalise("FIRST\tNAME"), "first name");
});

// ---------------------------------------------------------------------------
// formatValue
// ---------------------------------------------------------------------------

test("formatValue renders blanks, booleans, and strings", () => {
  assert.equal(DocsFiller.formatValue(null), "");
  assert.equal(DocsFiller.formatValue(undefined), "");
  assert.equal(DocsFiller.formatValue(true), "Yes");
  assert.equal(DocsFiller.formatValue(false), "No");
  assert.equal(DocsFiller.formatValue("Jane"), "Jane");
});

test("formatValue groups integers with thousands separators", () => {
  assert.equal(DocsFiller.formatValue(0), "0");
  assert.equal(DocsFiller.formatValue(42), "42");
  assert.equal(DocsFiller.formatValue(1234), "1,234");
  assert.equal(DocsFiller.formatValue(1234567), "1,234,567");
  assert.equal(DocsFiller.formatValue(-9876), "-9,876");
});

test("formatValue renders non-integers to 2dp with grouping", () => {
  assert.equal(DocsFiller.formatValue(1234.5), "1,234.50");
  assert.equal(DocsFiller.formatValue(0.5), "0.50");
  assert.equal(DocsFiller.formatValue(-1234.567), "-1,234.57");
});

test("formatValue renders dates, with time only when present", () => {
  assert.equal(DocsFiller.formatValue(new Date(2026, 8, 1)), "2026-09-01");
  assert.equal(
    DocsFiller.formatValue(new Date(2026, 8, 1, 9, 30, 5)),
    "2026-09-01 09:30:05"
  );
});

// ---------------------------------------------------------------------------
// readRecords
// ---------------------------------------------------------------------------

test("readRecords returns headers and normalised-key records", () => {
  const { headers, records } = DocsFiller.readRecords(
    makeXlsx([
      ["First Name", "Job Title"],
      ["Jane", "Engineer"],
      ["Sam", "Designer"],
    ])
  );
  assert.deepEqual(headers, ["First Name", "Job Title"]);
  assert.equal(records.length, 2);
  assert.equal(records[0]["first name"], "Jane");
  assert.equal(records[1]["job title"], "Designer");
});

test("readRecords reads a named worksheet and rejects an unknown one", () => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    wb, XLSX.utils.aoa_to_sheet([["A"], ["first"]]), "One"
  );
  XLSX.utils.book_append_sheet(
    wb, XLSX.utils.aoa_to_sheet([["B"], ["second"]]), "Two"
  );
  const bytes = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  assert.deepEqual(DocsFiller.readRecords(bytes, "Two").headers, ["B"]);
  // No sheet given -> first sheet.
  assert.deepEqual(DocsFiller.readRecords(bytes).headers, ["A"]);
  assert.throws(
    () => DocsFiller.readRecords(bytes, "Nope"),
    /Worksheet 'Nope' was not found/
  );
});

test("readRecords rejects an empty sheet and a header-only sheet", () => {
  assert.throws(
    () => DocsFiller.readRecords(makeXlsx([])),
    /spreadsheet is empty/
  );
  assert.throws(
    () => DocsFiller.readRecords(makeXlsx([["First Name"]])),
    /headers but no data rows/
  );
});

test("readRecords rejects a file that is not a spreadsheet", () => {
  assert.throws(
    () => DocsFiller.readRecords(Buffer.from("this is not an xlsx")),
    DocsFiller.DocsFillerError
  );
});

// ---------------------------------------------------------------------------
// fillPart — replacement mechanics
// ---------------------------------------------------------------------------

test("fillPart replaces a placeholder in a single run", () => {
  const xml = documentXml(para("Dear {{First Name}}, welcome."));
  const out = DocsFiller.fillPart(xml, resolverFor({ "First Name": "Jane" }), new Set());
  assert.equal(textOf(out), "Dear Jane, welcome.");
});

test("fillPart matches placeholders case- and space-insensitively", () => {
  const xml = documentXml(
    para("{{first name}} / {{FIRST NAME}} / {{  First   Name  }}")
  );
  const out = DocsFiller.fillPart(xml, resolverFor({ "First Name": "Jane" }), new Set());
  assert.equal(textOf(out), "Jane / Jane / Jane");
});

test("fillPart joins a placeholder split across runs", () => {
  const xml = documentXml(splitPara(["Dear {{First ", "Name}}", "!"]));
  const out = DocsFiller.fillPart(xml, resolverFor({ "First Name": "Jane" }), new Set());
  assert.equal(textOf(out), "Dear Jane!");
});

test("fillPart replaces several placeholders in one paragraph", () => {
  const xml = documentXml(para("{{A}}-{{B}}-{{C}}"));
  const out = DocsFiller.fillPart(
    xml, resolverFor({ A: "1", B: "2", C: "3" }), new Set()
  );
  assert.equal(textOf(out), "1-2-3");
});

test("fillPart preserves run formatting and adds xml:space", () => {
  const xml = documentXml(
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>{{First Name}}</w:t></w:r></w:p>'
  );
  const out = DocsFiller.fillPart(xml, resolverFor({ "First Name": "Jane" }), new Set());
  assert.match(out, /<w:rPr><w:b\/><\/w:rPr>/, "bold run properties survive");
  assert.match(out, /xml:space="preserve"/);
  assert.equal(textOf(out), "Jane");
});

test("fillPart XML-escapes replacement values", () => {
  const xml = documentXml(para("{{Company}}"));
  const out = DocsFiller.fillPart(
    xml, resolverFor({ Company: 'Smith & Sons <"Ltd">' }), new Set()
  );
  assert.match(out, /Smith &amp; Sons &lt;/, "raw & and < are escaped in the XML");
  assert.equal(textOf(out), 'Smith & Sons <"Ltd">');
});

test("fillPart leaves an unmatched placeholder alone and reports it", () => {
  const xml = documentXml(para("Dear {{First Name}}, role: {{Job Title}}."));
  const missing = new Set();
  const out = DocsFiller.fillPart(xml, resolverFor({ "First Name": "Jane" }), missing);
  assert.equal(textOf(out), "Dear Jane, role: {{Job Title}}.");
  assert.deepEqual([...missing], ["Job Title"]);
});

test("fillPart writes an empty cell as blank", () => {
  const xml = documentXml(para("Notes: {{Notes}}."));
  const out = DocsFiller.fillPart(xml, resolverFor({ Notes: null }), new Set());
  assert.equal(textOf(out), "Notes: .");
});

// ---------------------------------------------------------------------------
// buildDocuments — orchestration
// ---------------------------------------------------------------------------

test("buildDocuments produces one document per data row", async () => {
  const tpl = await makeDocx(para("Dear {{First Name}}."));
  const xlsx = makeXlsx([["First Name"], ["Jane"], ["Sam"]]);
  const { documents } = await DocsFiller.buildDocuments(tpl, xlsx, {
    base_name: "offer",
  });

  assert.equal(documents.length, 2);
  assert.deepEqual(
    documents.map((d) => d.filename),
    ["offer - Jane.docx", "offer - Sam.docx"]
  );
  assert.equal(await docTextInZip(documents[0].bytes), "Dear Jane.");
  assert.equal(await docTextInZip(documents[1].bytes), "Dear Sam.");
});

test("buildDocuments fills headers and footers too", async () => {
  const tpl = await makeDocx(para("Body: {{First Name}}"), {
    "word/header1.xml": para("Header: {{First Name}}"),
    "word/footer1.xml": para("Footer: {{First Name}}"),
  });
  const xlsx = makeXlsx([["First Name"], ["Jane"]]);
  const { documents, leftovers } = await DocsFiller.buildDocuments(tpl, xlsx, {});

  const zip = await JSZip.loadAsync(documents[0].bytes);
  assert.equal(textOf(await zip.file("word/header1.xml").async("string")), "Header: Jane");
  assert.equal(textOf(await zip.file("word/footer1.xml").async("string")), "Footer: Jane");
  assert.deepEqual(leftovers, []);
});

test("buildDocuments names files by name_column when given", async () => {
  const tpl = await makeDocx(para("{{First Name}}"));
  const xlsx = makeXlsx([
    ["First Name", "Employee Id"],
    ["Jane", "E-100"],
  ]);
  const { documents } = await DocsFiller.buildDocuments(tpl, xlsx, {
    base_name: "offer",
    name_column: "employee id",
  });
  assert.deepEqual(documents.map((d) => d.filename), ["offer - E-100.docx"]);
});

test("buildDocuments rejects an unknown name_column", async () => {
  const tpl = await makeDocx(para("{{First Name}}"));
  const xlsx = makeXlsx([["First Name"], ["Jane"]]);
  await assert.rejects(
    DocsFiller.buildDocuments(tpl, xlsx, { name_column: "Nope" }),
    /Column 'Nope' was not found. Available columns: First Name/
  );
});

test("buildDocuments disambiguates duplicate filenames", async () => {
  const tpl = await makeDocx(para("{{First Name}}"));
  const xlsx = makeXlsx([["First Name"], ["Jane"], ["Jane"], ["Jane"]]);
  const { documents } = await DocsFiller.buildDocuments(tpl, xlsx, {
    base_name: "offer",
  });
  assert.deepEqual(
    documents.map((d) => d.filename),
    ["offer - Jane.docx", "offer - Jane (2).docx", "offer - Jane (3).docx"]
  );
});

test("buildDocuments sanitises unsafe filename characters", async () => {
  const tpl = await makeDocx(para("{{Name}}"));
  const xlsx = makeXlsx([["Name"], ["a/b:c*d"]]);
  const { documents } = await DocsFiller.buildDocuments(tpl, xlsx, {
    base_name: "doc",
  });
  assert.deepEqual(documents.map((d) => d.filename), ["doc - a_b_c_d.docx"]);
});

// ---------------------------------------------------------------------------
// Leftover / unreplaced-variable detection
// ---------------------------------------------------------------------------

test("a fully-filled template reports no leftovers", async () => {
  const tpl = await makeDocx(para("Dear {{First Name}}, role: {{Job Title}}."));
  const xlsx = makeXlsx([
    ["First Name", "Job Title"],
    ["Jane", "Engineer"],
  ]);
  const { missing, leftovers, leftoverFiles } = await DocsFiller.buildDocuments(
    tpl, xlsx, {}
  );
  assert.deepEqual(missing, []);
  assert.deepEqual(leftovers, []);
  assert.deepEqual(leftoverFiles, []);
});

test("an unmatched column is reported as both missing and leftover", async () => {
  const tpl = await makeDocx(para("Dear {{First Name}}, role: {{Job Title}}."));
  const xlsx = makeXlsx([["First Name"], ["Jane"]]);
  const { missing, leftovers, leftoverFiles } = await DocsFiller.buildDocuments(
    tpl, xlsx, { base_name: "offer" }
  );
  assert.deepEqual(missing, ["Job Title"]);
  assert.deepEqual(leftovers, ["{{Job Title}}"]);
  assert.deepEqual(leftoverFiles, ["offer - Jane.docx"]);
});

test("a dangling brace is caught as a leftover even though it is not missing", async () => {
  // The fill regex never matches this, so `missing` cannot see it — this is
  // exactly the gap the leftover scan exists to close.
  const tpl = await makeDocx(para("Salary: {{Salary broken"));
  const xlsx = makeXlsx([["First Name"], ["Jane"]]);
  const { missing, leftovers } = await DocsFiller.buildDocuments(tpl, xlsx, {});

  assert.deepEqual(missing, [], "no placeholder was matched, so nothing is missing");
  assert.equal(leftovers.length, 1);
  assert.match(leftovers[0], /\{\{Salary broken/);
});

test("a stray closing brace is caught as a leftover", async () => {
  const tpl = await makeDocx(para("Ends badly }} here"));
  const xlsx = makeXlsx([["First Name"], ["Jane"]]);
  const { leftovers } = await DocsFiller.buildDocuments(tpl, xlsx, {});
  assert.equal(leftovers.length, 1);
  assert.match(leftovers[0], /\}\}/);
});

test("leftovers are detected in headers and footers", async () => {
  const tpl = await makeDocx(para("Body ok"), {
    "word/header1.xml": para("{{Unknown Header Field}}"),
  });
  const xlsx = makeXlsx([["First Name"], ["Jane"]]);
  const { leftovers } = await DocsFiller.buildDocuments(tpl, xlsx, {});
  assert.deepEqual(leftovers, ["{{Unknown Header Field}}"]);
});

test("leftovers are deduplicated across rows but list every affected file", async () => {
  const tpl = await makeDocx(para("Dear {{First Name}}, role: {{Job Title}}."));
  const xlsx = makeXlsx([["First Name"], ["Jane"], ["Sam"]]);
  const { leftovers, leftoverFiles } = await DocsFiller.buildDocuments(tpl, xlsx, {
    base_name: "offer",
  });
  assert.deepEqual(leftovers, ["{{Job Title}}"], "reported once, not per row");
  assert.deepEqual(leftoverFiles, ["offer - Jane.docx", "offer - Sam.docx"]);
});

test("a leftover split across runs is still detected", async () => {
  const tpl = await makeDocx(splitPara(["{{Job ", "Title}}"]));
  const xlsx = makeXlsx([["First Name"], ["Jane"]]);
  const { leftovers } = await DocsFiller.buildDocuments(tpl, xlsx, {});
  assert.deepEqual(leftovers, ["{{Job Title}}"]);
});

test("braces coming from spreadsheet data are flagged (known limitation)", async () => {
  const tpl = await makeDocx(para("Note: {{Notes}}"));
  const xlsx = makeXlsx([["Notes"], ["literally {{not a placeholder}}"]]);
  const { missing, leftovers } = await DocsFiller.buildDocuments(tpl, xlsx, {});

  assert.deepEqual(missing, []);
  // Known limitation: the scan reads the filled output, so braces coming from
  // spreadsheet data are flagged. Warning the user is the safe failure here.
  assert.deepEqual(leftovers, ["{{not a placeholder}}"]);
});

// ---------------------------------------------------------------------------
// generateZip
// ---------------------------------------------------------------------------

test("generateZip bundles every document and passes warnings through", async () => {
  const tpl = await makeDocx(para("Dear {{First Name}}, role: {{Job Title}}."));
  const xlsx = makeXlsx([["First Name"], ["Jane"], ["Sam"]]);
  const { zipBytes, missing, leftovers, leftoverFiles } =
    await DocsFiller.generateZip(tpl, xlsx, { base_name: "offer" });

  const zip = await JSZip.loadAsync(zipBytes);
  assert.deepEqual(
    Object.keys(zip.files).sort(),
    ["offer - Jane.docx", "offer - Sam.docx"]
  );
  assert.deepEqual(missing, ["Job Title"]);
  assert.deepEqual(leftovers, ["{{Job Title}}"]);
  assert.equal(leftoverFiles.length, 2);
});

test("generateZip reports progress once per document, in order", async () => {
  const tpl = await makeDocx(para("{{First Name}}"));
  const xlsx = makeXlsx([["First Name"], ["Jane"], ["Sam"], ["Ada"]]);
  const calls = [];
  await DocsFiller.generateZip(tpl, xlsx, {
    progress: (done, total) => calls.push([done, total]),
  });
  assert.deepEqual(calls, [[1, 3], [2, 3], [3, 3]]);
});
