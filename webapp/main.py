"""Docs Filler web app.

A tiny FastAPI front-end over :mod:`docs_filler.core`. HR users upload a
``.docx`` template and an ``.xlsx`` spreadsheet, watch a progress bar, and
download a single ``.zip`` of the finished documents.

Privacy by design: uploads and results live **only in memory** for the life of
a job. Nothing is written to disk and no storage service is attached. Jobs are
dropped as soon as they are downloaded, and expire automatically otherwise.

Authentication is handled *outside* this app by Azure Container Apps "Easy
Auth" (Microsoft Entra ID), so there is no auth code here — we only read the
signed-in user's name from the header Easy Auth injects.
"""

from __future__ import annotations

import asyncio
import json
import threading
import time
import uuid
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from docs_filler.core import DocsFillerError, generate_zip

# --- limits & housekeeping -------------------------------------------------
MAX_UPLOAD_BYTES = 25 * 1024 * 1024      # 25 MB per file is plenty for these docs
JOB_TTL_SECONDS = 15 * 60                # forget finished/abandoned jobs after 15 min
STATIC_DIR = Path(__file__).parent / "static"

app = FastAPI(title="Docs Filler")

# Serve static assets (the 2i logo, etc.) from the static directory.
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@dataclass
class Job:
    """One processing request, held in memory only."""
    status: str = "processing"           # processing | done | error
    done: int = 0
    total: Optional[int] = None
    zip_bytes: Optional[bytes] = None
    missing: list = field(default_factory=list)
    error: Optional[str] = None
    created_at: float = field(default_factory=time.monotonic)


_jobs: dict[str, Job] = {}
_lock = threading.Lock()


def _sweep_expired() -> None:
    now = time.monotonic()
    with _lock:
        for jid in [j for j, job in _jobs.items()
                    if now - job.created_at > JOB_TTL_SECONDS]:
            _jobs.pop(jid, None)


def _run_job(job_id: str, template_bytes: bytes, xlsx_bytes: bytes,
             base_name: str, name_column: Optional[str], sheet: Optional[str]) -> None:
    """Do the actual filling (runs in a background thread)."""
    def progress(done: int, total: int) -> None:
        with _lock:
            job = _jobs.get(job_id)
            if job:
                job.done, job.total = done, total

    try:
        zip_bytes, missing = generate_zip(
            template_bytes, xlsx_bytes,
            name_column=name_column or None,
            sheet=sheet or None,
            base_name=base_name,
            progress=progress,
        )
        with _lock:
            job = _jobs.get(job_id)
            if job:
                job.zip_bytes, job.missing, job.status = zip_bytes, missing, "done"
    except DocsFillerError as exc:
        _fail(job_id, str(exc))
    except Exception:  # noqa: BLE001 - don't leak internals to the user
        _fail(job_id, "Something went wrong while creating the documents. "
                      "Please check the template and spreadsheet and try again.")


def _fail(job_id: str, message: str) -> None:
    with _lock:
        job = _jobs.get(job_id)
        if job:
            job.status, job.error = "error", message


# --- routes ----------------------------------------------------------------

@app.get("/")
async def index():
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/me")
async def me(request: Request):
    """Name of the signed-in user, from Easy Auth headers (empty when local)."""
    name = (request.headers.get("X-MS-CLIENT-PRINCIPAL-NAME")
            or request.headers.get("X-MS-CLIENT-PRINCIPAL-ID")
            or "")
    return {"name": name}


@app.post("/api/jobs")
async def create_job(
    template: UploadFile = File(...),
    data: UploadFile = File(...),
    name_column: str = Form(""),
    sheet: str = Form(""),
):
    _sweep_expired()

    template_bytes = await template.read()
    xlsx_bytes = await data.read()
    for blob, label in ((template_bytes, "template"), (xlsx_bytes, "spreadsheet")):
        if not blob:
            raise HTTPException(400, f"The {label} file is empty.")
        if len(blob) > MAX_UPLOAD_BYTES:
            raise HTTPException(413, f"The {label} file is too large "
                                     f"(limit {MAX_UPLOAD_BYTES // (1024 * 1024)} MB).")

    base_name = Path(template.filename or "document").stem or "document"

    job_id = uuid.uuid4().hex
    with _lock:
        _jobs[job_id] = Job()

    threading.Thread(
        target=_run_job,
        args=(job_id, template_bytes, xlsx_bytes, base_name, name_column, sheet),
        daemon=True,
    ).start()

    return {"job_id": job_id}


@app.get("/api/jobs/{job_id}/events")
async def job_events(job_id: str):
    """Server-Sent Events stream of progress for one job."""
    with _lock:
        if job_id not in _jobs:
            raise HTTPException(404, "Job not found.")

    async def stream():
        last = None
        while True:
            with _lock:
                job = _jobs.get(job_id)
            if job is None:
                break
            payload = {"status": job.status, "done": job.done, "total": job.total}
            if job.status == "done":
                payload["missing"] = job.missing
            if job.status == "error":
                payload["error"] = job.error

            snapshot = json.dumps(payload)
            if snapshot != last:
                yield f"data: {snapshot}\n\n"
                last = snapshot
            if job.status in ("done", "error"):
                break
            await asyncio.sleep(0.2)

    return StreamingResponse(
        stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/api/jobs/{job_id}/download")
async def download(job_id: str):
    with _lock:
        job = _jobs.get(job_id)
    if job is None:
        raise HTTPException(404, "Job not found or already downloaded.")
    if job.status == "error":
        raise HTTPException(400, job.error or "The job failed.")
    if job.status != "done" or job.zip_bytes is None:
        raise HTTPException(409, "The documents are not ready yet.")

    zip_bytes = job.zip_bytes
    with _lock:
        _jobs.pop(job_id, None)  # one-time download; drop from memory immediately

    return StreamingResponse(
        BytesIO(zip_bytes),
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="documents.zip"'},
    )


@app.get("/healthz")
async def healthz():
    return JSONResponse({"status": "ok"})
