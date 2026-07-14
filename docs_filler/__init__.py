"""Docs Filler — shared logic for filling a .docx template from a spreadsheet.

Both the command-line tool (``fill_docs.py``) and the web app (``webapp``)
import from :mod:`docs_filler.core`, so the replacement behaviour is identical
everywhere.
"""

from .core import (
    DocsFillerError,
    build_documents,
    generate_zip,
    read_records,
)

__all__ = [
    "DocsFillerError",
    "build_documents",
    "generate_zip",
    "read_records",
]
