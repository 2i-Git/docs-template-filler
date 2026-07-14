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

The actual filling logic lives in ``docs_filler/core.py`` and is shared with the
web app, so both behave identically.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    from docs_filler.core import DocsFillerError, build_documents
except ImportError as exc:  # friendly message for non-technical users
    missing = getattr(exc, "name", "a required library")
    sys.exit(
        f"Missing required library ({missing}).\n"
        "Please run this once:  pip install -r requirements.txt"
    )


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

    template_bytes = template_path.read_bytes()
    xlsx_bytes = data_path.read_bytes()

    def show_progress(done: int, total: int) -> None:
        print(f"  ✓ {done}/{total}")

    try:
        documents, missing = build_documents(
            template_bytes, xlsx_bytes,
            name_column=args.name_column,
            sheet=args.sheet,
            base_name=template_path.stem,
            progress=show_progress,
        )
    except DocsFillerError as exc:
        sys.exit(str(exc))

    outdir.mkdir(parents=True, exist_ok=True)
    for filename, data in documents:
        (outdir / filename).write_bytes(data)

    if missing:
        names = ", ".join(f"{{{{{m}}}}}" for m in missing)
        print("\nWarning: these template placeholders have no matching spreadsheet "
              f"column and were left unchanged:\n  {names}")

    print(f"\nDone. {len(documents)} document(s) written to '{outdir}/'.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
