"""Unit tests for docs_filler.core."""

from __future__ import annotations

import zipfile
from io import BytesIO

import pytest

from docs_filler.core import (
    DocsFillerError,
    build_documents,
    format_value,
    generate_zip,
    normalise,
    read_records,
)
from tests.conftest import build_xlsx_bytes, read_docx_text


# --- small pure helpers ----------------------------------------------------

def test_normalise_collapses_case_and_space():
    assert normalise("  New   CTC ") == "new ctc"
    assert normalise("HRA") == "hra"


@pytest.mark.parametrize("value,expected", [
    (None, ""),
    (True, "Yes"),
    (False, "No"),
    (3210000, "3,210,000"),
    (1605000.0, "1,605,000"),          # whole float -> no decimals
    (1234.5, "1,234.50"),              # non-whole -> 2 dp
    (864236.5384615385, "864,236.54"),  # rounded to 2 dp
    ("Head of Delivery", "Head of Delivery"),
])
def test_format_value(value, expected):
    assert format_value(value) == expected


# --- read_records ----------------------------------------------------------

def test_read_records_headers_and_rows(xlsx_bytes):
    headers, records = read_records(BytesIO(xlsx_bytes))
    assert headers == ["Employee Name", "First Name", "Salary"]
    assert len(records) == 2
    assert records[0][normalise("Employee Name")] == "Alice Smith"
    assert records[1][normalise("Salary")] == 1234.5


def test_read_records_rejects_non_xlsx():
    with pytest.raises(DocsFillerError):
        read_records(BytesIO(b"this is not a spreadsheet"))


def test_read_records_rejects_headers_only():
    only_headers = build_xlsx_bytes(rows=[], headers=["A", "B"])
    with pytest.raises(DocsFillerError):
        read_records(BytesIO(only_headers))


def test_read_records_unknown_sheet(xlsx_bytes):
    with pytest.raises(DocsFillerError):
        read_records(BytesIO(xlsx_bytes), sheet="NoSuchSheet")


# --- build_documents -------------------------------------------------------

def test_build_documents_one_per_row(template_bytes, xlsx_bytes):
    docs, missing = build_documents(template_bytes, xlsx_bytes, base_name="Offer")
    assert len(docs) == 2
    # Filenames default to the first column (Employee Name).
    names = [name for name, _ in docs]
    assert names == ["Offer - Alice Smith.docx", "Offer - Bob Jones.docx"]


def test_build_documents_substitutes_including_split_runs(template_bytes, xlsx_bytes):
    docs, _ = build_documents(template_bytes, xlsx_bytes, base_name="Offer")
    text = read_docx_text(docs[0][1])
    assert "Employee: Alice Smith" in text
    assert "Dear Alice," in text          # placeholder split across 3 runs
    assert "Salary: 50,000" in text       # split runs + number formatting
    assert "Confirm: Alice Smith" in text  # case/space tolerant match
    # Known placeholders are fully resolved (only the deliberately-unknown one remains).
    for resolved in ("{{Employee Name}}", "{{First Name}}", "{{Salary}}", "{{  employee   name  }}"):
        assert resolved not in text


def test_build_documents_reports_missing_placeholder(template_bytes, xlsx_bytes):
    docs, missing = build_documents(template_bytes, xlsx_bytes)
    assert "Unknown Field" in missing
    # The unknown placeholder is left untouched in the output.
    assert "{{Unknown Field}}" in read_docx_text(docs[0][1])


def test_build_documents_progress_callback(template_bytes, xlsx_bytes):
    seen = []
    build_documents(template_bytes, xlsx_bytes, progress=lambda d, t: seen.append((d, t)))
    assert seen == [(1, 2), (2, 2)]


def test_build_documents_name_column(template_bytes, xlsx_bytes):
    docs, _ = build_documents(template_bytes, xlsx_bytes,
                              base_name="Offer", name_column="First Name")
    assert [name for name, _ in docs] == ["Offer - Alice.docx", "Offer - Bob.docx"]


def test_build_documents_unknown_name_column(template_bytes, xlsx_bytes):
    with pytest.raises(DocsFillerError):
        build_documents(template_bytes, xlsx_bytes, name_column="Nope")


def test_build_documents_dedupes_colliding_filenames(template_bytes):
    xlsx = build_xlsx_bytes(rows=[
        ["Alice Smith", "Alice", 1],
        ["Alice Smith", "Alice", 2],
    ])
    docs, _ = build_documents(template_bytes, xlsx, base_name="Offer")
    names = [name for name, _ in docs]
    assert names == ["Offer - Alice Smith.docx", "Offer - Alice Smith (2).docx"]


def test_build_documents_blank_name_falls_back_to_row(template_bytes):
    xlsx = build_xlsx_bytes(rows=[[None, "Alice", 1]])
    docs, _ = build_documents(template_bytes, xlsx, base_name="Offer")
    assert docs[0][0] == "Offer - row_1.docx"


# --- generate_zip ----------------------------------------------------------

def test_generate_zip_contains_one_docx_per_row(template_bytes, xlsx_bytes):
    zip_bytes, missing = generate_zip(template_bytes, xlsx_bytes, base_name="Offer")
    zf = zipfile.ZipFile(BytesIO(zip_bytes))
    assert zf.namelist() == ["Offer - Alice Smith.docx", "Offer - Bob Jones.docx"]
    # Each entry is a real, non-empty docx (zip container starts with 'PK').
    for name in zf.namelist():
        assert zf.read(name)[:2] == b"PK"
    assert "Unknown Field" in missing
