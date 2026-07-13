#!/usr/bin/env python3
"""Docs Filler — fill a Word (.docx) template from a spreadsheet.

It replaces every ``{{Placeholder}}`` in the template with the value from the
matching spreadsheet column, and writes one finished document per row.

The placeholder text and the spreadsheet column header must be the same word(s).
Matching ignores capitalisation and extra spaces, so ``{{New CTC}}`` matches a
column headed ``New CTC``, ``new ctc`` or ``New  CTC``.

Usage (from the project folder):
    python fill_docs.py --template T.docx --data D.xlsx
    python fill_docs.py --template T.docx --data D.xlsx --outdir out
    python fill_docs.py --template T.docx --data D.xlsx --name-column "Employee Name"

Run ``python fill_docs.py --help`` for all options.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

try:
    import openpyxl
    from docx import Document
except ImportError as exc:  # friendly message for non-technical users
    missing = getattr(exc, "name", "a required library")
    sys.exit(
        f"Missing required library ({missing}).\n"
        "Please run this once:  pip install -r requirements.txt"
    )

# Placeholders look like {{ Header Name }} — spaces inside the braces are ok.
PLACEHOLDER_RE = re.compile(r"\{\{\s*(.+?)\s*\}\}")


def normalise(text: str) -> str:
    """Lower-case and collapse whitespace, for tolerant header matching."""
    return re.sub(r"\s+", " ", str(text).strip()).lower()


# ---------------------------------------------------------------------------
# Spreadsheet reading
# ---------------------------------------------------------------------------

def read_records(data_path: Path, sheet: str | None):
    """Return (headers, records).

    headers: list of original header strings, in column order.
    records: list of dicts mapping normalised header -> raw cell value.
    """
    wb = openpyxl.load_workbook(data_path, data_only=True)
    ws = wb[sheet] if sheet else wb.active

    rows = [
        list(r) for r in ws.iter_rows(values_only=True)
        if any(c is not None and str(c).strip() != "" for c in r)
    ]
    if not rows:
        sys.exit(f"No data found in '{data_path}'.")

    header_row = rows[0]
    headers, index_by_key = [], {}
    for idx, cell in enumerate(header_row):
        if cell is None or str(cell).strip() == "":
            continue
        name = str(cell).strip()
        headers.append(name)
        index_by_key[normalise(name)] = idx

    records = []
    for row in rows[1:]:
        record = {
            key: (row[idx] if idx < len(row) else None)
            for key, idx in index_by_key.items()
        }
        records.append(record)

    return headers, records


# ---------------------------------------------------------------------------
# Value formatting
# ---------------------------------------------------------------------------

def format_value(value) -> str:
    """Turn a spreadsheet cell into a human-friendly string."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, int):
        return f"{value:,}"
    if isinstance(value, float):
        if value.is_integer():
            return f"{int(value):,}"
        return f"{round(value, 2):,.2f}"
    return str(value)


# ---------------------------------------------------------------------------
# DOCX replacement (run-aware, so text formatting is preserved)
# ---------------------------------------------------------------------------

def iter_paragraphs(document):
    """Yield every paragraph in the document, including tables/headers/footers."""
    def walk(container):
        for para in container.paragraphs:
            yield para
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    yield from walk(cell)

    yield from walk(document)
    for section in document.sections:
        for hf in (section.header, section.footer):
            yield from walk(hf)


def fill_paragraph(paragraph, resolve, missing: set) -> None:
    """Replace all {{...}} in one paragraph while keeping run formatting.

    A placeholder may be split across several runs, so we work on the joined
    text and then write the result back across the original runs.
    """
    runs = paragraph.runs
    if not runs:
        return
    texts = [r.text or "" for r in runs]
    full = "".join(texts)

    matches = list(PLACEHOLDER_RE.finditer(full))
    if not matches:
        return

    bounds, pos = [], 0
    for t in texts:
        bounds.append((pos, pos + len(t)))
        pos += len(t)
    total = pos

    def locate(p: int):
        for i, (s, e) in enumerate(bounds):
            if s <= p < e:
                return i, p - s
        return len(bounds) - 1, bounds[-1][1] - bounds[-1][0]

    # Right-to-left so earlier offsets remain valid as we edit.
    for m in reversed(matches):
        found, replacement = resolve(m.group(1))
        if not found:
            missing.add(m.group(1).strip())
            continue
        si, so = locate(m.start())
        if m.end() >= total:
            ei, eo = len(texts) - 1, len(texts[-1])
        else:
            ei, eo = locate(m.end())
        if si == ei:
            texts[si] = texts[si][:so] + replacement + texts[si][eo:]
        else:
            texts[si] = texts[si][:so] + replacement
            for k in range(si + 1, ei):
                texts[k] = ""
            texts[ei] = texts[ei][eo:]

    for run, new_text in zip(runs, texts):
        run.text = new_text


def fill_document(template_path: Path, record: dict, missing: set) -> Document:
    doc = Document(str(template_path))

    def resolve(placeholder_name: str):
        key = normalise(placeholder_name)
        if key in record:
            return True, format_value(record[key])
        return False, ""

    for paragraph in iter_paragraphs(doc):
        fill_paragraph(paragraph, resolve, missing)
    return doc


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def safe_filename(text: str, fallback: str) -> str:
    cleaned = re.sub(r"[^\w\-. ]+", "_", str(text).strip()).strip(". ")
    return cleaned or fallback


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Fill a .docx template from a spreadsheet — one file per row.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--template", required=True,
                        help="Path to the .docx template")
    parser.add_argument("--data", required=True,
                        help="Path to the .xlsx spreadsheet")
    parser.add_argument("--outdir", default="output",
                        help="Folder where finished documents are written")
    parser.add_argument("--sheet", default=None,
                        help="Worksheet name (default: the first sheet)")
    parser.add_argument("--name-column", default=None,
                        help="Column header used to name each output file "
                             "(default: the first column)")
    args = parser.parse_args()

    template_path = Path(args.template)
    data_path = Path(args.data)
    outdir = Path(args.outdir)

    if not template_path.exists():
        sys.exit(f"Template not found: {template_path}")
    if not data_path.exists():
        sys.exit(f"Spreadsheet not found: {data_path}")

    headers, records = read_records(data_path, args.sheet)
    if not headers:
        sys.exit("Could not find any column headers in the first row of the spreadsheet.")
    if not records:
        sys.exit("The spreadsheet has headers but no data rows.")

    # Which column drives the output filename.
    if args.name_column:
        key = normalise(args.name_column)
        if key not in {normalise(h) for h in headers}:
            sys.exit(f"--name-column '{args.name_column}' not found. "
                     f"Available columns: {', '.join(headers)}")
        name_key = key
    else:
        name_key = normalise(headers[0])

    outdir.mkdir(parents=True, exist_ok=True)
    template_stem = template_path.stem

    all_missing: set = set()
    seen_names: dict = {}
    generated = []
    for i, record in enumerate(records, start=1):
        missing: set = set()
        doc = fill_document(template_path, record, missing)
        all_missing |= missing

        base = safe_filename(format_value(record.get(name_key)), f"row_{i}")
        filename = f"{template_stem} - {base}"
        count = seen_names.get(filename, 0) + 1
        seen_names[filename] = count
        if count > 1:
            filename = f"{filename} ({count})"

        out_path = outdir / f"{filename}.docx"
        doc.save(str(out_path))
        generated.append(out_path)
        print(f"  ✓ {out_path.name}")

    if all_missing:
        names = ", ".join(f"{{{{{m}}}}}" for m in sorted(all_missing))
        print("\nWarning: these template placeholders have no matching spreadsheet "
              f"column and were left unchanged:\n  {names}")

    print(f"\nDone. {len(generated)} document(s) written to '{outdir}/'.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
