"""Smoke tests for the web app using FastAPI's TestClient.

These drive the real HTTP endpoints, including the background-thread job
processing, so they cover the app end-to-end without a server or a browser.
"""

from __future__ import annotations

import time
import zipfile
from io import BytesIO

import pytest
from fastapi.testclient import TestClient

from webapp.main import app
from tests.conftest import build_template_bytes, build_xlsx_bytes


@pytest.fixture
def client():
    return TestClient(app)


def test_healthz(client):
    assert client.get("/healthz").json() == {"status": "ok"}


def test_me_empty_when_no_auth_header(client):
    # Locally (no Easy Auth in front) there is no signed-in user.
    assert client.get("/api/me").json() == {"name": ""}


def _wait_for_download(client, job_id, timeout=10.0):
    """Poll the download endpoint until the background job is ready."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        res = client.get(f"/api/jobs/{job_id}/download")
        if res.status_code == 200:
            return res
        assert res.status_code in (409,), res.status_code  # 409 = not ready yet
        time.sleep(0.1)
    raise AssertionError("job did not finish in time")


def test_full_job_flow(client):
    files = {
        "template": ("t.docx", build_template_bytes(),
                     "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
        "data": ("d.xlsx", build_xlsx_bytes(),
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    }
    res = client.post("/api/jobs", files=files)
    assert res.status_code == 200
    job_id = res.json()["job_id"]

    download = _wait_for_download(client, job_id)
    assert download.headers["content-type"] == "application/zip"
    # The zip is named after the uploaded template ("t.docx" -> "t.zip").
    assert 't.zip' in download.headers["content-disposition"]
    zf = zipfile.ZipFile(BytesIO(download.content))
    assert len(zf.namelist()) == 2

    # One-time download: the job is dropped from memory afterwards.
    assert client.get(f"/api/jobs/{job_id}/download").status_code == 404


def test_rejects_empty_file(client):
    files = {
        "template": ("t.docx", b"", "application/octet-stream"),
        "data": ("d.xlsx", build_xlsx_bytes(), "application/octet-stream"),
    }
    assert client.post("/api/jobs", files=files).status_code == 400


def test_bad_spreadsheet_surfaces_error(client):
    files = {
        "template": ("t.docx", build_template_bytes(), "application/octet-stream"),
        "data": ("d.xlsx", b"not a spreadsheet", "application/octet-stream"),
    }
    job_id = client.post("/api/jobs", files=files).json()["job_id"]

    # The job should end in error; download reflects it as a 400.
    deadline = time.monotonic() + 10.0
    while time.monotonic() < deadline:
        res = client.get(f"/api/jobs/{job_id}/download")
        if res.status_code != 409:
            break
        time.sleep(0.1)
    assert res.status_code == 400
